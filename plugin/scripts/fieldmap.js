/*
 * Placeholder/field logic for the Mail Merge plugin — pure functions only.
 *
 * Token syntax: {{Name}} — braces with optional inner spaces, name is the
 * trimmed interior ("{{ first name }}" -> "first name"). Field names are
 * matched case-insensitively against data-source headers.
 *
 * Contract (all pure, no DOM, no editor calls):
 *   TOKEN_RE                        -> RegExp (global, for {{...}} tokens)
 *   findTokens(text)                -> string[]  unique, first-appearance order
 *   findTokensInTexts(texts)        -> string[]  across many paragraphs
 *   normalizeName(s)                -> string    trim + collapse inner spaces
 *   matchFields(tokens, headers)    -> { matched: [{token, header}],
 *                                        unmatched: [token],
 *                                        unusedHeaders: [header] }
 *   plainReplacePlan(text, record, tokens?) -> [{searchString, replaceString}]
 *       record := { header: value } (values already strings)
 *       one entry per distinct token occurrence pattern; missing values -> ""
 *   wrapPlan(text)                  -> [{token, name, start, end}]
 *       character offsets of each token within ONE paragraph's text, in order;
 *       `name` is the merge-field name; used by the editor-page run surgery
 *   fieldDisplay(name)              -> "«Name»"
 */
(function (root, factory) {
	if (typeof module === "object" && module.exports) {
		module.exports = factory(root);
	} else {
		root.OnlyOfficeMailMergeFieldMap = factory(root);
	}
})(typeof self !== "undefined" ? self : this, function (root) {
	"use strict";

	var TOKEN_RE = /\{\{\s*([^{}]+?)\s*\}\}/g;
	// Same pattern, anchored: distinguishes a bare token from a bare name.
	var TOKEN_FULL_RE = new RegExp("^" + TOKEN_RE.source + "$");

	function normalizeName(s) {
		if (s === null || s === undefined) {
			return "";
		}
		return String(s).trim().replace(/\s+/g, " ");
	}

	/*
	 * Visits every token occurrence as {raw, name, start}. TOKEN_RE is a global
	 * regex, i.e. stateful; each scan gets a private clone so scans can never
	 * leak `lastIndex` into each other or into callers holding TOKEN_RE.
	 */
	function eachToken(text, visit) {
		if (typeof text !== "string") {
			return;
		}
		var re = new RegExp(TOKEN_RE.source, "g");
		var match;
		while ((match = re.exec(text)) !== null) {
			var name = normalizeName(match[1]);
			if (name === "") {
				continue; // "{{ }}" names no field, so it is not a token
			}
			visit(match[0], name, match.index);
		}
	}

	// A token is the raw "{{...}}" text as it appears; its field name is the
	// interior. A bare name is accepted too, so header-shaped strings can flow
	// through the same matching path.
	function tokenName(token) {
		var raw = String(token === null || token === undefined ? "" : token);
		var match = TOKEN_FULL_RE.exec(raw);
		return normalizeName(match ? match[1] : raw);
	}

	function findTokens(text) {
		var tokens = [];
		var seen = Object.create(null);
		eachToken(text, function (raw) {
			if (!Object.prototype.hasOwnProperty.call(seen, raw)) {
				seen[raw] = true;
				tokens.push(raw);
			}
		});
		return tokens;
	}

	function findTokensInTexts(texts) {
		var tokens = [];
		var seen = Object.create(null);
		var list = Array.isArray(texts) ? texts : [];
		for (var i = 0; i < list.length; i++) {
			eachToken(list[i], function (raw) {
				if (!Object.prototype.hasOwnProperty.call(seen, raw)) {
					seen[raw] = true;
					tokens.push(raw);
				}
			});
		}
		return tokens;
	}

	// Field names meet headers on normalized, case-insensitive terms; the
	// first header wins when two headers normalize to the same key.
	function matchFields(tokens, headers) {
		var list = Array.isArray(tokens) ? tokens : [];
		var cols = Array.isArray(headers) ? headers : [];
		var byName = Object.create(null);
		for (var c = 0; c < cols.length; c++) {
			var key = normalizeName(cols[c]).toLowerCase();
			if (key !== "" && !Object.prototype.hasOwnProperty.call(byName, key)) {
				byName[key] = c;
			}
		}
		var matched = [];
		var unmatched = [];
		var used = Object.create(null);
		for (var t = 0; t < list.length; t++) {
			var nameKey = tokenName(list[t]).toLowerCase();
			if (nameKey !== "" && Object.prototype.hasOwnProperty.call(byName, nameKey)) {
				matched.push({ token: list[t], header: cols[byName[nameKey]] });
				used[byName[nameKey]] = true;
			} else {
				unmatched.push(list[t]);
			}
		}
		var unusedHeaders = [];
		for (var h = 0; h < cols.length; h++) {
			if (!Object.prototype.hasOwnProperty.call(used, h)) {
				unusedHeaders.push(cols[h]);
			}
		}
		return { matched: matched, unmatched: unmatched, unusedHeaders: unusedHeaders };
	}

	// Record keys are data-source headers, so they meet token names the same
	// way matchFields matches them. A token without a value becomes "".
	function valueFor(record, name) {
		var key = name.toLowerCase();
		var keys = Object.keys(record);
		for (var i = 0; i < keys.length; i++) {
			if (normalizeName(keys[i]).toLowerCase() === key) {
				var value = record[keys[i]];
				return value === null || value === undefined ? "" : String(value);
			}
		}
		return "";
	}

	/*
	 * One entry per distinct raw token form ("{{name}}" and "{{ name }}" are two
	 * occurrence patterns and both need their own searchString to be found
	 * again). `tokens` restricts the plan; each entry either names one raw form
	 * (as findTokens reports them) or one bare field name (every form of that
	 * field), so both caller styles restrict exactly what they say.
	 */
	function plainReplacePlan(text, record, tokens) {
		var restrict = null;
		if (Array.isArray(tokens)) {
			restrict = { forms: Object.create(null), names: Object.create(null) };
			for (var i = 0; i < tokens.length; i++) {
				var entry = String(tokens[i]);
				if (TOKEN_FULL_RE.test(entry)) {
					restrict.forms[entry] = true;
				} else {
					restrict.names[normalizeName(entry).toLowerCase()] = true;
				}
			}
		}
		var source = record && typeof record === "object" ? record : {};
		var plan = [];
		var seen = Object.create(null);
		eachToken(text, function (raw, name) {
			if (Object.prototype.hasOwnProperty.call(seen, raw)) {
				return;
			}
			if (restrict && !Object.prototype.hasOwnProperty.call(restrict.forms, raw)
					&& !Object.prototype.hasOwnProperty.call(restrict.names, name.toLowerCase())) {
				return;
			}
			seen[raw] = true;
			plan.push({ searchString: raw, replaceString: valueFor(source, name) });
		});
		return plan;
	}

	function wrapPlan(text) {
		var entries = [];
		eachToken(text, function (raw, name, start) {
			entries.push({ token: raw, name: name, start: start, end: start + raw.length });
		});
		return entries;
	}

	function fieldDisplay(name) {
		return "«" + (name === null || name === undefined ? "" : String(name)) + "»";
	}

	return {
		TOKEN_RE: TOKEN_RE,
		findTokens: findTokens,
		findTokensInTexts: findTokensInTexts,
		normalizeName: normalizeName,
		matchFields: matchFields,
		plainReplacePlan: plainReplacePlan,
		wrapPlan: wrapPlan,
		fieldDisplay: fieldDisplay
	};
});
