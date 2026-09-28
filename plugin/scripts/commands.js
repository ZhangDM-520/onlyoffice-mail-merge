/*
 * Editor-page seam for the Mail Merge plugin: command bodies that run inside
 * the editor page via Asc.plugin.callCommand, plus `run`, the one way the
 * plugin frame dispatches them.
 *
 * Same discipline as the latex-math plugin (see
 * /home/zhangdm/Projects/onlyoffice-latex-math/plugin/scripts/commands.js):
 *   - callCommand stringifies the function -> commands cannot close over anything;
 *     makeCommand composes self-contained bodies from a shared PRELUDE
 *   - every command returns a JSON string ({ok:true,...} | {ok:false,error})
 *   - Asc.scope carries the payload; commands are serialized (one in flight)
 *   - 30 s backstop with a named error taxonomy
 *
 * Command registry (payload -> result):
 *   getDocumentText            {}        -> {ok, paragraphs: [string]}
 *   wrapFields                 {tokens}  -> {ok, wrapped:[{name,count}], errors:[]}
 *       Scans paragraphs for {{Name}} tokens; splits runs at token boundaries
 *       and wraps each token run in a real MERGEFIELD displayed as «Name»
 *       (mirrors ApiParagraph/ApiRun.WrapInMailMergeField; see SDK
 *       ParaField(AscWord.fieldtype_MERGEFIELD, [name]) construction).
 *   loadMergeData              {data}    -> {ok, count}
 *       Api.LoadMailMergeData(data); data[0] = field names.
 *   getMergeCount              {}        -> {ok, count}   (Api.GetMailMergeReceptionsCount)
 *   snapshotTemplate           {}        -> {ok}
 *       Api.GetMailMergeTemplateDocContent() kept on the editor-page side
 *       (attach to the resolved Api object; complex objects cannot cross
 *       the callCommand boundary).
 *   mergeRange                 {start,end} -> {ok}
 *       Api.MailMerge(start,end) — replaces the open document with merged output.
 *   restoreTemplate            {}        -> {ok}
 *       Api.ReplaceDocumentContent(snapshot).
 *   replacePlain               {plan}    -> {ok, replaced, errors}
 *       Sequential SearchAndReplace over [{searchString,replaceString,matchCase}];
 *       `replaced` counts OCCURRENCES (the engine replaces all matches per call);
 *       `errors` lists entries skipped (e.g. 'unsafe-value' for data containing
 *       the caret-marker). Replacement values are caret-safe (see body).
 *
 * Dispatch contract: `run(name, payload, callback)` - the callback receives
 * exactly one argument, the parsed result shaped {ok:true,...} | {ok:false,error}.
 * Seam-level failures arrive through the same shape, with the taxonomy names
 * `timeout`, `unparsable-command-result`, `clobbered` (internal only),
 * `callCommand-unavailable`, `callCommand-threw: ...`. An optional fourth
 * argument overrides the 30 s backstop (how the timeout paths are tested).
 */
(function (root, factory) {
	if (typeof module === "object" && module.exports) {
		module.exports = factory(root);
	} else {
		root.OnlyOfficeMailMergeCommands = factory(root);
	}
})(typeof self !== "undefined" ? self : this, function (root) {
	"use strict";

	// Helpers shared by every editor-page command. The prelude is evaluated
	// *inside the editor page*, where ONLYOFFICE 7.1+ shadow-stubs `window`,
	// `document` and `alert` as empty objects/no-ops (safePluginEval binds them
	// as function parameters) - nothing below may rely on them behaving like
	// the real DOM, and `root.*` here belongs to the plugin iframe instead.
	var PRELUDE = [
		"function resolveApi() {",
		"	if (typeof Api !== 'undefined' && Api) return Api;",
		"	return null;",
		"}",
		"function scopeOf(scopeArg) {",
		"	if (scopeArg) return scopeArg;",
		"	try { if (typeof scope !== 'undefined' && scope) return scope; } catch (e) {}",
		"	return {};",
		"}",
		"function errorText(e) {",
		"	return (e && e.message) ? String(e.message) : String(e);",
		"}",
		// The logic document (CDocument) behind the resolved Api: the ApiDocument
		// holds it as `.Document`; a build that hands back the logic document
		// directly is accepted as-is.
		"function logicDocumentOf(A) {",
		"	var doc = null;",
		"	try { doc = A.GetDocument(); } catch (e) { return null; }",
		"	if (!doc) return null;",
		"	return doc.Document ? doc.Document : doc;",
		"}",
		// Whole-document paragraph list. `GetAllParagraphs` when the Api surface
		// offers it (latex-math's seam relies on it), otherwise the body's own
		// `Content` array (CDocument.Content).
		"function allParagraphsOf(A) {",
		"	var doc = null;",
		"	try { doc = A.GetDocument(); } catch (e) { return null; }",
		"	if (!doc) return null;",
		"	if (typeof doc.GetAllParagraphs === 'function') {",
		"		try {",
		"			var ps = doc.GetAllParagraphs();",
		"			if (ps && typeof ps.length === 'number') return ps;",
		"		} catch (e) {}",
		"	}",
		"	var ld = doc.Document ? doc.Document : doc;",
		"	if (ld && ld.Content && typeof ld.Content.length === 'number') return ld.Content;",
		"	return null;",
		"}",
		// ApiParagraph carries its internal paragraph as `.Paragraph`; paragraphs
		// reached through CDocument.Content are already internal.
		"function paraObjectOf(p) {",
		"	return (p && p.Paragraph) ? p.Paragraph : p;",
		"}",
		"function paraTextOf(p) {",
		"	if (!p || typeof p.GetText !== 'function') return '';",
		"	try {",
		"		var t = p.GetText({ Numbering: false });",
		"		return typeof t === 'string' ? t : '';",
		"	} catch (e) {",
		"		try {",
		"			var t2 = p.GetText();",
		"			return typeof t2 === 'string' ? t2 : '';",
		"		} catch (e2) { return ''; }",
		"	}",
		"}",
		// Rendered text of one paragraph content element, used as the element's
		// character cost when mapping token offsets onto run boundaries.
		"function elementText(item) {",
		"	if (!item) return '';",
		"	if (typeof item.GetText === 'function') {",
		"		try {",
		"			var t = item.GetText();",
		"			if (typeof t === 'string') return t;",
		"		} catch (e) {}",
		"	}",
		"	if (item.Content && typeof item.Content.length === 'number') {",
		"		var s = '';",
		"		for (var i = 0; i < item.Content.length; i++) s += elementText(item.Content[i]);",
		"		return s;",
		"	}",
		"	return '';",
		"}",
		// A run is what can be split at a character offset (ParaRun exposes
		// Split2/AddText; see sdk-all.js ParaRun.prototype.Split2).
		"function isRunElement(item) {",
		"	return !!(item && typeof item.AddText === 'function' && typeof item.Split2 === 'function' && typeof item.Copy === 'function');",
		"}",
		// Token offsets come from the pure module
		// root.OnlyOfficeMailMergeFieldMap.wrapPlan(text): [{token,name,start,end}].
		// The module is loaded alongside this one via the same UMD pattern (its
		// global is read here at command time). When the editor page offers no
		// field map - the module lives in the plugin frame, which callCommand
		// cannot see - the contract is served by a self-contained scanner below.
		// Both agree on: `name` is the trimmed interior of {{...}}.
		"function fallbackWrapPlan(text) {",
		"	var out = [];",
		"	if (typeof text !== 'string' || !text) return out;",
		"	var re = /\\{\\{\\s*([^{}]+?)\\s*\\}\\}/g;",
		"	var m;",
		"	while ((m = re.exec(text)) !== null) {",
		"		var raw = String(m[1]);",
		"		var name = raw.replace(/^\\s+|\\s+$/g, '');",
		"		out.push({ token: m[0], name: name, start: m.index, end: m.index + m[0].length });",
		"	}",
		"	return out;",
		"}",
		"function wrapPlanFor(text) {",
		"	var fm = null;",
		"	try { if (typeof OnlyOfficeMailMergeFieldMap !== 'undefined' && OnlyOfficeMailMergeFieldMap) fm = OnlyOfficeMailMergeFieldMap; } catch (e) {}",
		"	if (fm && typeof fm.wrapPlan === 'function') {",
		"		try {",
		"			var plan = fm.wrapPlan(text);",
		"			if (plan && typeof plan.length === 'number') return plan;",
		"		} catch (e) { /* fall through to the built-in scanner */ }",
		"	}",
		"	return fallbackWrapPlan(text);",
		"}",
		// The MERGEFIELD display content: the field NAME (so the field renders
		// «Name»), carrying the formatting of the span's first run. Mirrors the
		// SDK's WrapInMailMergeField quote runs ("«" left, "»" right around the
		// field content) except the content is retyped from {{Name}} to Name -
		// the display contract of this plugin.
		"function nameRunFor(firstRun, name) {",
		"	var inner = new ParaRun();",
		"	try {",
		"		if (firstRun && firstRun.Pr && typeof firstRun.Pr.Copy === 'function' && typeof inner.SetPr === 'function') {",
		"			inner.SetPr(firstRun.Pr.Copy(true));",
		"		}",
		"	} catch (e) { /* default formatting is an acceptable degradation */ }",
		"	inner.AddText(name);",
		"	return inner;",
		"}",
		// Splits runs so `start` and `end` fall on element boundaries, then
		// reports the element index range covering [start, end). Returns null
		// when a boundary lies inside an element that cannot be split (a field,
		// a drawing, ...) - the caller reports that span and moves on.
		"function splitAtBoundaries(paraObj, start, end) {",
		"	var content = paraObj.Content;",
		"	if (!content || typeof content.length !== 'number') return null;",
		"	if (typeof paraObj.Add_ToContent !== 'function' || typeof paraObj.Remove_FromContent !== 'function') return null;",
		"	var boundaries = [start, end];",
		"	for (var b = 0; b < boundaries.length; b++) {",
		"		var offset = boundaries[b];",
		"		var pos = 0;",
		"		for (var i = 0; i < content.length; i++) {",
		"			var item = content[i];",
		"			var len = elementText(item).length;",
		"			if (offset > pos && offset < pos + len) {",
		"				if (!isRunElement(item)) return null;",
		"				var right = null;",
		"				try { right = item.Split2(offset - pos); } catch (e) { return null; }",
		"				if (!right) return null;",
		"				paraObj.Add_ToContent(i + 1, right);",
		"				break;",
		"			}",
		"			pos += len;",
		"		}",
		"	}",
		"	var pos2 = 0, from = -1, to = -1;",
		"	for (var j = 0; j < content.length; j++) {",
		"		var l = elementText(content[j]).length;",
		"		var e2 = pos2 + l;",
		"		if (l === 0) {",
		"			if (pos2 > start && pos2 < end) { if (from === -1) from = j; to = j + 1; }",
		"		} else if (e2 > start && pos2 < end) {",
		"			if (from === -1) from = j;",
		"			to = j + 1;",
		"		}",
		"		pos2 = e2;",
		"	}",
		"	if (from === -1) return null;",
		"	return { from: from, to: to };",
		"}",
		// Low-level run surgery: wrap one paragraph's token spans in MERGEFIELDs.
		// Processing spans right-to-left keeps the not-yet-processed spans'
		// element indices valid after each replacement. The replacement itself
		// mirrors ApiRun.WrapInMailMergeField: build ParaField, swap it into the
		// parent content array at the span's place, Register_Field.
		"function wrapSpansInParagraph(paraObj, spans, ld) {",
		"	var result = { wrapped: [], errors: [] };",
		"	var ordered = spans.slice().sort(function (a, b) { return b.start - a.start; });",
		"	for (var s = 0; s < ordered.length; s++) {",
		"		var span = ordered[s];",
		"		var range = null;",
		"		try { range = splitAtBoundaries(paraObj, span.start, span.end); } catch (e) { range = null; }",
		"		if (!range) {",
		"			result.errors.push({ name: span.name, start: span.start, end: span.end, error: 'run-split-unavailable' });",
		"			continue;",
		"		}",
		"		var content = paraObj.Content;",
		"		var items = content.slice(range.from, range.to);",
		"		var allRuns = items.length > 0;",
		"		for (var i = 0; i < items.length; i++) {",
		"			if (!isRunElement(items[i])) { allRuns = false; break; }",
		"		}",
		"		if (!allRuns) {",
		"			result.errors.push({ name: span.name, start: span.start, end: span.end, error: 'unsupported-element-in-span' });",
		"			continue;",
		"		}",
		"		var oField = null;",
		"		try {",
		"			oField = new ParaField(AscWord.fieldtype_MERGEFIELD, [span.name], []);",
		"			var leftQuote = new ParaRun();",
		"			var rightQuote = new ParaRun();",
		"			leftQuote.AddText('\\u00AB');",
		"			rightQuote.AddText('\\u00BB');",
		"			oField.Add_ToContent(0, leftQuote);",
		"			oField.Add_ToContent(1, nameRunFor(items[0], span.name));",
		"			oField.Add_ToContent(2, rightQuote);",
		"			paraObj.Remove_FromContent(range.from, range.to - range.from);",
		"			paraObj.Add_ToContent(range.from, oField);",
		"		} catch (e) {",
		"			result.errors.push({ name: span.name, start: span.start, end: span.end, error: 'field-insert-failed: ' + errorText(e) });",
		"			continue;",
		"		}",
		"		try { if (ld && typeof ld.Register_Field === 'function') ld.Register_Field(oField); } catch (e) { /* bookkeeping only */ }",
		"		result.wrapped.push({ name: span.name, count: 1 });",
		"	}",
		"	result.wrapped.reverse();",
		"	return result;",
		"}",
		// Whether the low-level surgery primitives exist in this build at all.
		"function canOperate(paraObj) {",
		"	try {",
		"		return !!(paraObj && paraObj.Content && typeof paraObj.Content.length === 'number' &&",
		"			typeof paraObj.Add_ToContent === 'function' && typeof paraObj.Remove_FromContent === 'function' &&",
		"			typeof ParaRun === 'function' && typeof ParaField === 'function' &&",
		"			typeof AscWord !== 'undefined' && AscWord);",
		"	} catch (e) { return false; }",
		"}",
		// Degradation path (see header): without run surgery only whole
		// single-token paragraphs can be wrapped, via the high-level
		// ApiParagraph.WrapInMailMergeField. That SDK call names the field after
		// the paragraph text, so the name is repaired afterwards when - and only
		// when - the field object stays reachable.
		"function renameFieldBestEffort(paraObj, name) {",
		"	try {",
		"		var content = paraObj && paraObj.Content;",
		"		if (!content) return;",
		"		for (var i = 0; i < content.length; i++) {",
		"			var item = content[i];",
		"			if (item && item.Arguments && item.FieldType !== undefined) {",
		"				item.Arguments = [name];",
		"				return;",
		"			}",
		"		}",
		"	} catch (e) { /* best effort */ }",
		"}",
		"function isWholeParagraphToken(text, span) {",
		"	var trimmed = String(text).replace(/^\\s+|\\s+$/g, '');",
		"	return trimmed === span.token && span.start >= 0 && span.end <= String(text).length;",
		"}"
	].join("\n");

	function makeCommand(bodySource) {
		// eslint-disable-next-line no-new-func
		return new Function("return function (scopeArg) {\n" + PRELUDE + "\n" + bodySource + "\n};")();
	}

	// The editor answers callCommand asynchronously; without a backstop a
	// dropped callback would leave the UI waiting forever. `run`'s fourth
	// argument overrides it, which is how the timeout paths are tested without
	// waiting 30 s.
	var COMMAND_TIMEOUT_MS = 30000;

	// Every command answers with a JSON string. A string that does not parse is
	// named (`unparsable-command-result`) and kept raw for the report; a
	// callback that carries nothing at all, or no callback, is the "no
	// response" this seam reports and the taxonomy names `timeout`.
	function parseCommandResult(result) {
		if (typeof result === "string" && result !== "") {
			try {
				return JSON.parse(result);
			} catch (e) {
				return { error: "unparsable-command-result", raw: result };
			}
		}
		if (result && typeof result === "object") {
			return result;
		}
		return null;
	}

	// The callback contract is absolute: {ok:true,...} | {ok:false,error}.
	// Everything the seam or a misbehaving command body produces is folded into
	// that shape before it reaches the caller.
	function normalizeResult(parsed) {
		if (!parsed) {
			return { ok: false, error: "timeout" };
		}
		if (typeof parsed.ok === "boolean") {
			return parsed;
		}
		if (parsed.error) {
			var out = { ok: false, error: parsed.error };
			if (parsed.raw !== undefined) {
				out.raw = parsed.raw;
			}
			return out;
		}
		return { ok: false, error: "unparsable-command-result", raw: JSON.stringify(parsed) };
	}

	var GET_DOCUMENT_TEXT_BODY = [
		"var A = resolveApi();",
		"if (!A) return JSON.stringify({ ok: false, error: 'api-unavailable' });",
		"var paragraphs = allParagraphsOf(A);",
		"if (!paragraphs) return JSON.stringify({ ok: false, error: 'no-document' });",
		"var out = [];",
		"for (var i = 0; i < paragraphs.length; i++) {",
		"	out.push(paraTextOf(paragraphs[i]));",
		"}",
		"return JSON.stringify({ ok: true, paragraphs: out });"
	].join("\n");

	var WRAP_FIELDS_BODY = [
		"var S = scopeOf(scopeArg);",
		"var A = resolveApi();",
		"if (!A) return JSON.stringify({ ok: false, error: 'api-unavailable' });",
		"var paragraphs = allParagraphsOf(A);",
		"if (!paragraphs) return JSON.stringify({ ok: false, error: 'no-document' });",
		"var ld = logicDocumentOf(A);",
		// `tokens` restricts wrapping to the listed names; absent means every
		// token in the document. Matching is case-insensitive on the name.
		"var wanted = null;",
		"if (S.tokens && typeof S.tokens.length === 'number') {",
		"	wanted = {};",
		"	for (var w = 0; w < S.tokens.length; w++) {",
		"		var wn = S.tokens[w];",
		"		if (wn && typeof wn === 'object') wn = wn.name;",
		"		if (typeof wn === 'string') wanted[wn.toLowerCase()] = true;",
		"	}",
		"}",
		"var wrappedMap = {};",
		"var wrappedOrder = [];",
		"function addWrapped(name) {",
		"	if (wrappedMap[name] === undefined) { wrappedMap[name] = 0; wrappedOrder.push(name); }",
		"	wrappedMap[name]++;",
		"}",
		"var errors = [];",
		"var degraded = false;",
		"for (var pi = 0; pi < paragraphs.length; pi++) {",
		"	var p = paragraphs[pi];",
		"	var paraObj = paraObjectOf(p);",
		"	var text = paraTextOf(p);",
		"	var plan = [];",
		"	try { plan = wrapPlanFor(text) || []; } catch (e) { errors.push({ paragraph: pi, error: 'wrap-plan-failed: ' + errorText(e) }); continue; }",
		"	var spans = [];",
		"	for (var k = 0; k < plan.length; k++) {",
		"		var span = plan[k];",
		"		if (!span || typeof span.start !== 'number' || typeof span.end !== 'number' || span.end <= span.start) continue;",
		"		var name = typeof span.name === 'string' ? span.name.replace(/^\\s+|\\s+$/g, '') : '';",
		"		if (!name) continue;",
		"		if (wanted && !wanted[name.toLowerCase()]) continue;",
		"		spans.push({ name: name, token: typeof span.token === 'string' ? span.token : '', start: span.start, end: span.end });",
		"	}",
		"	if (!spans.length) continue;",
		"	if (canOperate(paraObj)) {",
		"		var r = wrapSpansInParagraph(paraObj, spans, ld);",
		"		for (var rw = 0; rw < r.wrapped.length; rw++) addWrapped(r.wrapped[rw].name);",
		"		for (var re2 = 0; re2 < r.errors.length; re2++) {",
		"			var err = r.errors[re2];",
		"			err.paragraph = pi;",
		"			errors.push(err);",
		"		}",
		"	} else {",
		"		degraded = true;",
		"		for (var si = 0; si < spans.length; si++) {",
		"			var dspan = spans[si];",
		"			if (spans.length === 1 && isWholeParagraphToken(text, dspan) && p && typeof p.WrapInMailMergeField === 'function') {",
		"				try {",
		"					p.WrapInMailMergeField();",
		"					renameFieldBestEffort(paraObj, dspan.name);",
		"					addWrapped(dspan.name);",
		"				} catch (e) {",
		"					errors.push({ name: dspan.name, paragraph: pi, error: 'wrap-failed: ' + errorText(e) });",
		"				}",
		"			} else {",
		"				errors.push({ name: dspan.name, paragraph: pi, error: 'inline-wrap-unavailable' });",
		"			}",
		"		}",
		"	}",
		"}",
		"var wrapped = [];",
		"for (var wo = 0; wo < wrappedOrder.length; wo++) {",
		"	wrapped.push({ name: wrappedOrder[wo], count: wrappedMap[wrappedOrder[wo]] });",
		"}",
		"return JSON.stringify({ ok: true, wrapped: wrapped, errors: errors, degraded: degraded });"
	].join("\n");

	var LOAD_MERGE_DATA_BODY = [
		"var S = scopeOf(scopeArg);",
		"var A = resolveApi();",
		"if (!A) return JSON.stringify({ ok: false, error: 'api-unavailable' });",
		"if (typeof A.LoadMailMergeData !== 'function') return JSON.stringify({ ok: false, error: 'load-unavailable' });",
		"var data = S.data;",
		"if (!data || typeof data.length !== 'number' || data.length === 0) return JSON.stringify({ ok: false, error: 'merge-data-empty' });",
		"var loaded = false;",
		"try { loaded = A.LoadMailMergeData(data); } catch (e) { return JSON.stringify({ ok: false, error: 'load-failed: ' + errorText(e) }); }",
		"if (!loaded) return JSON.stringify({ ok: false, error: 'merge-data-rejected' });",
		"var count = Math.max(0, data.length - 1);",
		"try {",
		"	if (typeof A.GetMailMergeReceptionsCount === 'function') count = A.GetMailMergeReceptionsCount();",
		"} catch (e) { /* keep the row count */ }",
		"return JSON.stringify({ ok: true, count: count });"
	].join("\n");

	var GET_MERGE_COUNT_BODY = [
		"var A = resolveApi();",
		"if (!A) return JSON.stringify({ ok: false, error: 'api-unavailable' });",
		"if (typeof A.GetMailMergeReceptionsCount !== 'function') return JSON.stringify({ ok: false, error: 'count-unavailable' });",
		"try {",
		"	return JSON.stringify({ ok: true, count: A.GetMailMergeReceptionsCount() });",
		"} catch (e) {",
		"	return JSON.stringify({ ok: false, error: 'count-failed: ' + errorText(e) });",
		"}"
	].join("\n");

	var SNAPSHOT_TEMPLATE_BODY = [
		"var A = resolveApi();",
		"if (!A) return JSON.stringify({ ok: false, error: 'api-unavailable' });",
		"if (typeof A.GetMailMergeTemplateDocContent !== 'function') return JSON.stringify({ ok: false, error: 'snapshot-unavailable' });",
		"try {",
		"	var snapshot = A.GetMailMergeTemplateDocContent();",
		"	if (!snapshot) return JSON.stringify({ ok: false, error: 'snapshot-unavailable' });",
		// Complex objects cannot cross the callCommand boundary - the snapshot
		// stays editor-page-side, attached to the resolved Api object.
		"	A.__mmSnapshot = snapshot;",
		"	return JSON.stringify({ ok: true });",
		"} catch (e) {",
		"	return JSON.stringify({ ok: false, error: 'snapshot-failed: ' + errorText(e) });",
		"}"
	].join("\n");

	var MERGE_RANGE_BODY = [
		"var S = scopeOf(scopeArg);",
		"var A = resolveApi();",
		"if (!A) return JSON.stringify({ ok: false, error: 'api-unavailable' });",
		"if (typeof A.MailMerge !== 'function') return JSON.stringify({ ok: false, error: 'merge-unavailable' });",
		"var start = S.start;",
		"var end = S.end;",
		"if (typeof start !== 'number' || typeof end !== 'number') return JSON.stringify({ ok: false, error: 'merge-range-invalid' });",
		"try {",
		"	var merged = A.MailMerge(start, end);",
		"	if (!merged) return JSON.stringify({ ok: false, error: 'merge-failed' });",
		"	return JSON.stringify({ ok: true });",
		"} catch (e) {",
		"	return JSON.stringify({ ok: false, error: 'merge-failed: ' + errorText(e) });",
		"}"
	].join("\n");

	var RESTORE_TEMPLATE_BODY = [
		"var A = resolveApi();",
		"if (!A) return JSON.stringify({ ok: false, error: 'api-unavailable' });",
		"var snapshot = A.__mmSnapshot;",
		"if (!snapshot) return JSON.stringify({ ok: false, error: 'no-snapshot' });",
		"if (typeof A.ReplaceDocumentContent !== 'function') return JSON.stringify({ ok: false, error: 'restore-unavailable' });",
		"try {",
		"	var replaced = A.ReplaceDocumentContent(snapshot);",
		"	if (!replaced) return JSON.stringify({ ok: false, error: 'restore-failed' });",
		"	return JSON.stringify({ ok: true });",
		"} catch (e) {",
		"	return JSON.stringify({ ok: false, error: 'restore-failed: ' + errorText(e) });",
		"}"
	].join("\n");

	var REPLACE_PLAIN_BODY = [
		"var S = scopeOf(scopeArg);",
		"var A = resolveApi();",
		"if (!A) return JSON.stringify({ ok: false, error: 'api-unavailable' });",
		"var doc = null;",
		"try { doc = A.GetDocument(); } catch (e) { doc = null; }",
		// The builder surface puts SearchAndReplace on the document; accept it
		// on Api too (api.executeMethod equivalent).
		"var target = null;",
		"if (doc && typeof doc.SearchAndReplace === 'function') target = doc;",
		"else if (typeof A.SearchAndReplace === 'function') target = A;",
		"if (!target) return JSON.stringify({ ok: false, error: 'replace-unavailable' });",
		"var plan = S.plan;",
		"if (!plan || typeof plan.length !== 'number') return JSON.stringify({ ok: false, error: 'no-plan' });",
		// CSearchPatternEngine interprets ^t/^p/^l/^n/^m/^~ in replacement
		// strings (inserting field codes) and offers no escape - but a lone
		// trailing '^' inserts a literal caret. So every caret in the data
		// travels as a private-use marker and is demoted back in one final
		// pass: a value like "a^p" survives as text instead of a paragraph
		// mark. A value that already carries the marker is refused outright.
		"var CARET_MARKER = '\\uE0FF';",
		"function demoteCarets() {",
		"	try { target.SearchAndReplace({ searchString: CARET_MARKER, replaceString: '^', matchCase: true }); } catch (e) { /* leave markers rather than throw */ }",
		"}",
		"function countOccurrences(needle, matchCase) {",
		"	var paragraphs = allParagraphsOf(A);",
		"	if (!paragraphs) return 0;",
		"	var total = 0;",
		"	var hay = matchCase === false ? needle.toLowerCase() : needle;",
		"	for (var ci = 0; ci < paragraphs.length; ci++) {",
		"		var text = paraTextOf(paragraphs[ci]) || '';",
		"		if (matchCase === false) text = text.toLowerCase();",
		"		var from = 0;",
		"		for (;;) {",
		"			var at = text.indexOf(hay, from);",
		"			if (at < 0) break;",
		"			total++;",
		"			from = at + hay.length;",
		"		}",
		"	}",
		"	return total;",
		"}",
		"var replaced = 0;",
		"var errors = [];",
		"var demoteNeeded = false;",
		"for (var i = 0; i < plan.length; i++) {",
		"	var entry = plan[i];",
		"	if (!entry || typeof entry.searchString !== 'string' || entry.searchString === '') continue;",
		"	var value = typeof entry.replaceString === 'string' ? entry.replaceString : '';",
		"	if (value.indexOf(CARET_MARKER) >= 0) {",
		"		errors.push({ searchString: entry.searchString, error: 'unsafe-value' });",
		"		continue;",
		"	}",
		"	if (value.indexOf('^') >= 0) {",
		"		value = value.split('^').join(CARET_MARKER);",
		"		demoteNeeded = true;",
		"	}",
		"	var matchCase = entry.matchCase === undefined ? true : !!entry.matchCase;",
		"	var occurrences = countOccurrences(entry.searchString, matchCase);",
		"	var props = { searchString: entry.searchString, replaceString: value, matchCase: matchCase };",
		"	try {",
		"		if (target.SearchAndReplace(props)) replaced += (occurrences > 0 ? occurrences : 1);",
		"	} catch (e) {",
		"		if (demoteNeeded) demoteCarets();",
		"		return JSON.stringify({ ok: false, error: 'replace-failed: ' + errorText(e), replaced: replaced, errors: errors });",
		"	}",
		"}",
		"if (demoteNeeded) demoteCarets();",
		"return JSON.stringify({ ok: true, replaced: replaced, errors: errors });"
	].join("\n");

	var RUNNERS = {
		getDocumentText: makeCommand(GET_DOCUMENT_TEXT_BODY),
		wrapFields: makeCommand(WRAP_FIELDS_BODY),
		loadMergeData: makeCommand(LOAD_MERGE_DATA_BODY),
		getMergeCount: makeCommand(GET_MERGE_COUNT_BODY),
		snapshotTemplate: makeCommand(SNAPSHOT_TEMPLATE_BODY),
		mergeRange: makeCommand(MERGE_RANGE_BODY),
		restoreTemplate: makeCommand(RESTORE_TEMPLATE_BODY),
		replacePlain: makeCommand(REPLACE_PLAIN_BODY)
	};

	// Runs are serialised: one command is in flight and later runs queue behind
	// it. The payload rides ONE shared slot (`Asc.scope`), so two commands in
	// flight at once swap payloads - the same live pathology latex-math
	// measured. Queueing keeps every answer paired with the payload its own run
	// put in the slot. The alternative is fail-fast supersede (a new run
	// displaces the in-flight one, which settles `clobbered`): a fresh click
	// would start immediately instead of waiting, but the displaced caller gets
	// an error and the displaced command may still read the superseding run's
	// payload. Queueing chooses "every caller keeps its own answer" over "the
	// newest run starts now" - a run queued behind a hung one waits out that
	// run's timeout before its own round trip begins.
	var inFlight = null;
	var waiting = [];

	/**
	 * Run one command in the editor page. `callback` receives exactly one
	 * argument: the parsed result shaped {ok:true,...} | {ok:false,error},
	 * where `error` is either a command-level reason (e.g. `no-snapshot`) or a
	 * taxonomy name (`timeout`, `unparsable-command-result`,
	 * `callCommand-unavailable`, `callCommand-threw: ...`). `timeoutMs` overrides
	 * the backstop for tests. Unknown names throw: that is a caller-side
	 * programming error, not a seam failure.
	 */
	function run(name, payload, callback, timeoutMs) {
		var command = RUNNERS[name];
		if (!command) {
			throw new Error("unknown editor command: " + name);
		}
		if (typeof callback !== "function") {
			throw new Error("run(" + name + ") needs a callback");
		}
		waiting.push({
			command: command,
			payload: payload || {},
			callback: callback,
			timeoutMs: typeof timeoutMs === "number" ? timeoutMs : COMMAND_TIMEOUT_MS,
			settled: false,
			timer: null
		});
		pump();
	}

	function pump() {
		if (inFlight || waiting.length === 0) {
			return;
		}
		var entry = waiting.shift();
		inFlight = entry;
		dispatch(entry);
	}

	function dispatch(entry) {
		// The UMD root is the plugin frame's `window` (or `self`) wherever this
		// module is loaded; the host bridge lives there.
		var plugin = root.Asc && root.Asc.plugin;
		if (!plugin || typeof plugin.callCommand !== "function") {
			finish(entry, { ok: false, error: "callCommand-unavailable" });
			return;
		}
		try {
			// The generated command wrapper reads the payload from Asc.scope.
			root.Asc.scope = entry.payload;
			// Arm the backstop before dispatching: a host that answers
			// synchronously would otherwise leave the timer orphaned.
			entry.timer = root.setTimeout(function () {
				finish(entry, { ok: false, error: "timeout" });
			}, entry.timeoutMs);
			plugin.callCommand(entry.command, false, true, function (value) {
				if (entry.settled) {
					// A duplicated or late callback settles as `clobbered` and is
					// discarded: the guard in `finish` is what drops it, so a late
					// answer can neither revive a finished run nor leak into the
					// run that owns the slot now. Under the queueing policy above
					// no caller ever observes `clobbered`; it is the displaced
					// run's answer under the fail-fast alternative.
					finish(entry, { ok: false, error: "clobbered" });
					return;
				}
				finish(entry, normalizeResult(parseCommandResult(value)));
			});
		} catch (e) {
			finish(entry, { ok: false, error: "callCommand-threw: " + (e && e.message) });
		}
	}

	function finish(entry, result) {
		if (entry.settled) {
			return;
		}
		entry.settled = true;
		if (entry.timer !== null) {
			root.clearTimeout(entry.timer);
			entry.timer = null;
		}
		inFlight = null;
		try {
			entry.callback(result);
		} finally {
			// A throwing caller callback must not wedge the queue behind it.
			pump();
		}
	}

	return {
		run: run,
		// The compiled commands stay exported for tests that drive one body
		// directly. Plugin code goes through `run`.
		commands: RUNNERS
	};
});
