/*
 * ONLYOFFICE Mail Merge plugin - the wizard window's composition root.
 *
 * A `type: "window"` variation (640x560, buttons Cancel/Merge) mounting a three
 * step wizard into <div id="mm-app">:
 *
 *   1. Data source  pick a CSV/XLSX/JSON file, parse it (dataparse.js), preview
 *                   the first 10 rows, and scan the document for {{Field}} tokens
 *   2. Fields       token vs header mapping with auto-match status (fieldmap.js);
 *                   matched tokens can be wrapped into real MERGEFIELDs, the
 *                   rest fall back to plain text replacement
 *   3. Output       one combined document or one file per recipient, a record
 *                   range ("1-10,15") and the output format (docx/pdf)
 *
 * Architecture (composition root - no UMD here, that is what the modules do):
 *   dataparse.js   String in -> String[][] (row 0 = headers)
 *   fieldmap.js    token scanning / matching / plain-replace plans (pure)
 *   commands.js    the `run` seam into the editor page (callCommand bodies)
 *   this file      wizard state, DOM, and the merge pipeline's sequencing
 *
 * Pipeline (sequential `run` calls - the seam serializes them and a second call
 * in flight would swap the shared Asc.scope payload):
 *
 *   loadMergeData(data) -> wrapFields(matched) [if wrapping] -> snapshotTemplate ->
 *   combined:   mergeRange(0, n-1) -> replacePlain(plan) [if plain tokens] ->
 *               GetFileToDownload(format)
 *   per-record: for each record: mergeRange(i, i) -> replacePlain(plan for i) ->
 *               GetFileToDownload(format) -> restoreTemplate
 *
 * There is no create-new-document API, so per-record output mutates the open
 * document record by record and restores the snapshot after every save. Errors
 * surface in the status area and the window is NEVER closed on error; Cancel
 * (button id 0) and the X (id -1) close via CloseWindow once the pipeline has
 * unwound. The window id comes from the page URL (`?windowID=...`, the same
 * source the SDK's own plugins.js `q()` reads) because `Asc.plugin.windowID` is
 * only published from an XHR callback and may be missing early.
 */
(function (window, document) {
	"use strict";

	var DataParse = window.OnlyOfficeMailMergeDataParse;
	var FieldMap = window.OnlyOfficeMailMergeFieldMap;
	var Commands = window.OnlyOfficeMailMergeCommands;

	if (!DataParse || !FieldMap || !Commands) {
		// Loading order problem: fail loudly instead of silently doing nothing.
		console.error("[mail-merge] dataparse.js / fieldmap.js / commands.js missing");
		return;
	}

	var SETTINGS_KEY = "onlyoffice-mail-merge.settings";
	var SETTINGS_VERSION = 1;
	var PREVIEW_ROWS = 10;
	var STEP_TITLES = ["Data source", "Fields", "Output"];
	// The mapping table needs more room than the default variation size; the
	// host honours ResizeWindow on a window-type variation.
	var STEP_SIZES = [[640, 560], [640, 660], [640, 560]];

	/* ------------------------------------------------------------------ *
	 * Wizard state
	 * ------------------------------------------------------------------ */

	var state = {
		step: 0,
		source: { name: "", data: null, warnings: [] },
		paragraphs: [],
		tokens: [],
		scanned: false,
		mapping: { matched: [], unmatched: [], unusedHeaders: [] },
		plainUnmatched: {},
		wrapMatched: true,
		mode: "combined",
		range: "",
		format: "docx",
		running: false,
		cancelled: false,
		closeAfterRun: false
	};

	/* ------------------------------------------------------------------ *
	 * Small helpers
	 * ------------------------------------------------------------------ */

	function tr(text) {
		try {
			if (window.Asc && window.Asc.plugin && typeof window.Asc.plugin.tr === "function") {
				return window.Asc.plugin.tr(text);
			}
		} catch (e) {
			/* fall through */
		}
		return text;
	}

	function el(tag, className, text) {
		var node = document.createElement(tag);
		if (className) {
			node.className = className;
		}
		if (text !== undefined && text !== null) {
			node.textContent = String(text);
		}
		return node;
	}

	function clear(node) {
		while (node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function isDarkTheme(theme) {
		var label = "";
		if (theme) {
			label = String(theme.type || theme.name || theme || "");
		}
		return label.toLowerCase().indexOf("dark") !== -1;
	}

	/* ------------------------------------------------------------------ *
	 * Settings (last-used wizard choices; versioned so a schema change
	 * degrades to defaults instead of half-loading)
	 * ------------------------------------------------------------------ */

	function defaultSettings() {
		return {
			version: SETTINGS_VERSION,
			wrapMatched: true,
			mode: "combined",
			range: "",
			format: "docx"
		};
	}

	function loadSettings(storage) {
		var defaults = defaultSettings();
		if (!storage) {
			return defaults;
		}
		try {
			var raw = storage.getItem(SETTINGS_KEY);
			if (!raw) {
				return defaults;
			}
			var parsed = JSON.parse(raw);
			if (!parsed || parsed.version !== SETTINGS_VERSION) {
				return defaults;
			}
			return {
				version: SETTINGS_VERSION,
				wrapMatched: parsed.wrapMatched !== false,
				mode: parsed.mode === "per-record" ? "per-record" : "combined",
				range: typeof parsed.range === "string" ? parsed.range : "",
				format: parsed.format === "pdf" ? "pdf" : "docx"
			};
		} catch (e) {
			return defaults;
		}
	}

	function saveSettings(storage, settings) {
		if (!storage) {
			return;
		}
		try {
			storage.setItem(
				SETTINGS_KEY,
				JSON.stringify({
					version: SETTINGS_VERSION,
					wrapMatched: settings.wrapMatched !== false,
					mode: settings.mode === "per-record" ? "per-record" : "combined",
					range: settings.range || "",
					format: settings.format === "pdf" ? "pdf" : "docx"
				})
			);
		} catch (e) {
			/* storage may be full or blocked; the wizard still works */
		}
	}

	function storage() {
		try {
			return window.localStorage || null;
		} catch (e) {
			return null;
		}
	}

	function currentSettings() {
		return {
			version: SETTINGS_VERSION,
			wrapMatched: state.wrapMatched,
			mode: state.mode,
			range: state.range,
			format: state.format
		};
	}

	/* ------------------------------------------------------------------ *
	 * Record range: "1-10,15" (1-based recipient numbers) -> 0-based indices
	 * ------------------------------------------------------------------ */

	function parseRange(text, count) {
		var trimmed = String(text === undefined || text === null ? "" : text).trim();
		var i;
		if (!trimmed) {
			var all = [];
			for (i = 0; i < count; i++) {
				all.push(i);
			}
			return { ok: true, indices: all };
		}
		if (!(count > 0)) {
			return { ok: false, error: tr("The data source has no recipient rows.") };
		}
		var indices = [];
		var seen = {};
		var parts = trimmed.split(",");
		for (i = 0; i < parts.length; i++) {
			var part = parts[i].trim();
			if (!part) {
				return { ok: false, error: tr("The record range has an empty segment.") };
			}
			var match = /^(\d+)(?:\s*-\s*(\d+))?$/.exec(part);
			if (!match) {
				return {
					ok: false,
					error: tr('Cannot read the record range segment "' + part + '". Use numbers like 1-10,15.')
				};
			}
			var from = parseInt(match[1], 10);
			var to = match[2] !== undefined ? parseInt(match[2], 10) : from;
			if (from < 1 || to < 1) {
				return { ok: false, error: tr("Record numbers start at 1.") };
			}
			if (to < from) {
				return { ok: false, error: tr('The record range segment "' + part + '" runs backwards.') };
			}
			if (to > count) {
				return {
					ok: false,
					error: tr('The record range segment "' + part + '" is beyond the last recipient (' + count + ").")
				};
			}
			for (var n = from; n <= to; n++) {
				if (!seen[n]) {
					seen[n] = true;
					indices.push(n - 1);
				}
			}
		}
		indices.sort(function (a, b) {
			return a - b;
		});
		return { ok: true, indices: indices };
	}

	function isContiguous(indices) {
		for (var i = 1; i < indices.length; i++) {
			if (indices[i] !== indices[0] + i) {
				return false;
			}
		}
		return true;
	}

	/* ------------------------------------------------------------------ *
	 * Editor bridge - one `run` in flight at a time
	 * ------------------------------------------------------------------ */

	function exec(name, payload) {
		return new Promise(function (resolve, reject) {
			var settled = false;
			try {
				Commands.run(name, payload || {}, function (result) {
					if (settled) {
						return;
					}
					settled = true;
					resolve(result);
				});
			} catch (e) {
				if (!settled) {
					settled = true;
					reject(e);
				}
			}
		});
	}

	/**
	 * The save seam. `GetFileToDownload` opens the OS save dialog on Desktop
	 * Editors per call and answers with a URL - or the literal string "error".
	 */
	function downloadFile(format) {
		return new Promise(function (resolve) {
			try {
				window.Asc.plugin.executeMethod("GetFileToDownload", [format], function (result) {
					if (!result || result === "error") {
						resolve({ ok: false, error: tr("the save dialog was cancelled or failed") });
						return;
					}
					resolve({ ok: true, url: result });
				});
			} catch (e) {
				resolve({ ok: false, error: String((e && e.message) || e) });
			}
		});
	}

	/* ------------------------------------------------------------------ *
	 * Merge pipeline
	 *
	 * Sequencing policy lives here (which command, in which order, and when the
	 * window may close); the mapping policy is fieldmap.js's and the command
	 * bodies are commands.js's.
	 * ------------------------------------------------------------------ */

	function cancelledError() {
		var error = new Error(tr("Merge cancelled."));
		error.cancelled = true;
		return error;
	}

	async function stage(context, label, name, payload) {
		if (context.isCancelled()) {
			throw cancelledError();
		}
		context.onProgress(label);
		var result = await context.exec(name, payload);
		if (!result || result.ok !== true) {
			throw new Error(label + ": " + ((result && result.error) || tr("the command failed")));
		}
		return result;
	}

	/**
	 * The record object plainReplacePlan looks values up in: keyed by header
	 * (the contract's shape) and mirrored under each token's own name, because
	 * the plan is generated per token. `rowIndex < 0` builds the combined-mode
	 * record, where one document holds every record's copy of the template and
	 * therefore no single value can be right - plain tokens are blanked.
	 */
	function plainRecord(data, rowIndex, matched, plainTokens) {
		var headers = data[0];
		var row = rowIndex >= 0 ? data[rowIndex] || [] : [];
		var record = {};
		var c;
		for (c = 0; c < headers.length; c++) {
			record[headers[c]] = rowIndex >= 0 && row[c] !== undefined && row[c] !== null ? String(row[c]) : "";
		}
		matched.forEach(function (entry) {
			var value = record[entry.header];
			value = value === undefined || value === null ? "" : value;
			record[entry.token] = value;
			try {
				if (typeof FieldMap.normalizeName === "function") {
					record[FieldMap.normalizeName(entry.token)] = value;
				}
			} catch (e) {
				/* a name the mapper cannot normalise is keyed literally above */
			}
		});
		plainTokens.forEach(function (token) {
			if (record[token] === undefined) {
				record[token] = "";
			}
		});
		return record;
	}

	/**
	 * The pipeline. `options`:
	 *   data            String[][] (row 0 = headers)
	 *   paragraphs      template paragraphs (the plain plan's text)
	 *   matched         [{token, header}] from FieldMap.matchFields
	 *   unmatched       [token]
	 *   allTokens       [token] every token in the template
	 *   wrapMatched     wrap matched tokens into real merge fields
	 *   plainUnmatched  {token: bool} which unmatched tokens to plain-replace
	 *   mode            "combined" | "per-record"
	 *   format          "docx" | "pdf"
	 *   rangeText       "1-10,15" or "" for all
	 *   exec / download / isCancelled / onProgress / fieldMap  seams (tests)
	 */
	async function runMergePipeline(options) {
		var fieldMap = options.fieldMap || FieldMap;
		var runExec = options.exec || exec;
		var download = options.download || downloadFile;
		var isCancelled = options.isCancelled || function () {
			return false;
		};
		var onProgress = options.onProgress || function () {};
		var context = { exec: runExec, download: download, isCancelled: isCancelled, onProgress: onProgress };

		function fail(error) {
			return { ok: false, error: error };
		}

		var data = options.data;
		if (!data || data.length < 2 || !data[0] || !data[0].length) {
			return fail(tr("The data source has no recipient rows."));
		}
		var recipients = data.length - 1;

		var range = parseRange(options.rangeText, recipients);
		if (!range.ok) {
			return fail(range.error);
		}
		var indices = range.indices;
		if (!indices.length) {
			return fail(tr("The record range selects no recipients."));
		}
		if (options.mode !== "per-record" && !isContiguous(indices)) {
			// MailMerge(start, end) takes one contiguous record span, so a gap
			// like "1-10,15" can only be produced as separate files.
			return fail(tr("A combined document needs a contiguous record range (like 1-10); pick one file per recipient to leave gaps."));
		}

		var matched = options.matched || [];
		var unmatched = options.unmatched || [];
		var allTokens = options.allTokens || [];
		var plainUnmatched = options.plainUnmatched || {};
		var plainTokens = options.wrapMatched
			? unmatched.filter(function (token) {
				return plainUnmatched[token] !== false;
			})
			: allTokens.slice();
		var templateText = (options.paragraphs || []).join("\n");

		// The load payload is the parse contract's canonical shape; fall back to
		// the grid itself if toLoadMailMergeData is unavailable.
		var loadData = data;
		try {
			if (typeof DataParse.toLoadMailMergeData === "function") {
				loadData = DataParse.toLoadMailMergeData({ ok: true, data: data, warnings: [] }) || data;
			}
		} catch (e) {
			loadData = data;
		}

		var merged = false;
		try {
			await stage(context, tr("Loading merge data"), "loadMergeData", { data: loadData });
			if (options.wrapMatched && matched.length) {
				await stage(context, tr("Wrapping merge fields"), "wrapFields", {
					tokens: matched.map(function (entry) {
						return entry.token;
					})
				});
			}
			await stage(context, tr("Saving the template snapshot"), "snapshotTemplate", {});

			if (options.mode === "per-record") {
				for (var i = 0; i < indices.length; i++) {
					var row = indices[i];
					await stage(
						context,
						tr("Merging record") + " " + (i + 1) + "/" + indices.length,
						"mergeRange",
						{ start: row, end: row }
					);
					merged = true;
					if (plainTokens.length) {
						var plan = fieldMap.plainReplacePlan(
							templateText,
							plainRecord(data, row, matched, plainTokens),
							plainTokens
						);
						if (plan && plan.length) {
							await stage(context, tr("Replacing placeholders"), "replacePlain", { plan: plan });
						}
					}
					var saved = await download(options.format);
					if (!saved || saved.ok !== true) {
						throw new Error(
							tr("Saving record") + " " + (i + 1) + ": " + ((saved && saved.error) || tr("save failed"))
						);
					}
					await stage(context, tr("Restoring the template"), "restoreTemplate", {});
					merged = false;
				}
				return { ok: true, records: indices.length };
			}

			await stage(context, tr("Merging"), "mergeRange", {
				start: indices[0],
				end: indices[indices.length - 1]
			});
			merged = true;
			if (plainTokens.length) {
				// One combined document holds every record's copy of the template,
				// so a plain token cannot vary per copy - see plainRecord().
				var combinedPlan = fieldMap.plainReplacePlan(
					templateText,
					plainRecord(data, -1, matched, plainTokens),
					plainTokens
				);
				if (combinedPlan && combinedPlan.length) {
					await stage(context, tr("Replacing placeholders"), "replacePlain", { plan: combinedPlan });
				}
			}
			var one = await download(options.format);
			if (!one || one.ok !== true) {
				throw new Error(tr("Saving the document") + ": " + ((one && one.error) || tr("save failed")));
			}
			return { ok: true, records: indices.length };
		} catch (error) {
			if (merged) {
				// Never strand the user's template as merged output.
				try {
					await context.exec("restoreTemplate", {});
				} catch (e) {
					/* best effort - the error below is what the user sees */
				}
			}
			if (error && error.cancelled) {
				return { ok: false, cancelled: true, error: error.message };
			}
			return fail(String((error && error.message) || error));
		}
	}

	/* ------------------------------------------------------------------ *
	 * Data source
	 * ------------------------------------------------------------------ */

	function loadSourceText(name, text) {
		var result;
		try {
			result = /\.json$/i.test(name) ? DataParse.parseJson(text) : DataParse.parseCsv(text, window.Papa);
		} catch (e) {
			result = { ok: false, error: String((e && e.message) || e) };
		}
		return applyParseResult(name, result);
	}

	function parseWorkbook(name, arrayBuffer) {
		var XLSXLib = window.XLSX;
		if (!XLSXLib) {
			return applyParseResult(name, { ok: false, error: tr("the spreadsheet reader (XLSX) is not loaded") });
		}
		try {
			var workbook = XLSXLib.read(new Uint8Array(arrayBuffer), { type: "array", cellDates: true });
			var sheet = workbook.Sheets[workbook.SheetNames[0]];
			var grid = XLSXLib.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
			return applyParseResult(name, DataParse.parseGrid(grid));
		} catch (e) {
			return applyParseResult(name, { ok: false, error: String((e && e.message) || e) });
		}
	}

	function applyParseResult(name, result) {
		if (!result || result.ok !== true) {
			statusText(
				tr("Could not parse") + " " + name + ": " + ((result && result.error) || tr("unknown error")),
				"error"
			);
			return false;
		}
		state.source = { name: name, data: result.data, warnings: result.warnings || [] };
		recomputeMapping();
		statusText(
			tr("Loaded") +
				" " +
				(result.data.length - 1) +
				" " +
				tr("recipients") +
				", " +
				result.data[0].length +
				" " +
				tr("fields") +
				".",
			"ok"
		);
		renderAll();
		return true;
	}

	function readFile(file) {
		var name = (file && file.name) || "data";
		try {
			var reader = new window.FileReader();
			if (/\.xlsx?$/i.test(name)) {
				reader.onload = function () {
					parseWorkbook(name, reader.result);
				};
				reader.readAsArrayBuffer(file);
			} else {
				reader.onload = function () {
					loadSourceText(name, String(reader.result));
				};
				reader.readAsText(file);
			}
			reader.onerror = function () {
				statusText(tr("Could not read the file."), "error");
			};
		} catch (e) {
			statusText(tr("Could not read the file") + ": " + String((e && e.message) || e), "error");
		}
	}

	/* ------------------------------------------------------------------ *
	 * Document scan + mapping
	 * ------------------------------------------------------------------ */

	async function scanDocument() {
		statusText(tr("Scanning the document for {{Field}} tokens..."));
		var result;
		try {
			result = await exec("getDocumentText", {});
		} catch (e) {
			statusText(tr("Could not read the document") + ": " + String((e && e.message) || e), "error");
			return false;
		}
		if (!result || result.ok !== true) {
			statusText(tr("Could not read the document") + ": " + ((result && result.error) || tr("the command failed")), "error");
			return false;
		}
		state.paragraphs = result.paragraphs || [];
		try {
			state.tokens = typeof FieldMap.findTokensInTexts === "function"
				? FieldMap.findTokensInTexts(state.paragraphs)
				: FieldMap.findTokens(state.paragraphs.join("\n"));
		} catch (e) {
			statusText(tr("Could not scan the document") + ": " + String((e && e.message) || e), "error");
			return false;
		}
		state.scanned = true;
		recomputeMapping();
		statusText(tr("Found") + " " + state.tokens.length + " " + tr("{{Field}} tokens") + ".", "ok");
		renderAll();
		return true;
	}

	function recomputeMapping() {
		var headers = state.source.data ? state.source.data[0] : [];
		if (state.tokens.length && headers.length) {
			state.mapping = FieldMap.matchFields(state.tokens, headers);
		} else {
			state.mapping = { matched: [], unmatched: state.tokens.slice(), unusedHeaders: headers.slice() };
		}
		state.mapping.unmatched.forEach(function (token) {
			if (state.plainUnmatched[token] === undefined) {
				state.plainUnmatched[token] = true;
			}
		});
	}

	function previewRows(data, limit) {
		return data.slice(0, limit + 1);
	}

	/* ------------------------------------------------------------------ *
	 * Host window plumbing
	 * ------------------------------------------------------------------ */

	function getWindowId() {
		try {
			var match = /[?&]windowID=([^&]+)/.exec(window.location.search || "");
			if (match) {
				return decodeURIComponent(match[1]);
			}
		} catch (e) {
			/* fall through to the SDK-published value */
		}
		try {
			return (window.Asc && window.Asc.plugin && window.Asc.plugin.windowID) || undefined;
		} catch (e) {
			return undefined;
		}
	}

	function closeWindow() {
		var id = getWindowId();
		try {
			window.Asc.plugin.executeMethod("CloseWindow", id !== undefined ? [id] : []);
		} catch (e) {
			console.error("[mail-merge] CloseWindow failed", e);
		}
	}

	function resizeForStep(step) {
		var size = STEP_SIZES[step] || STEP_SIZES[0];
		try {
			if (window.Asc && window.Asc.plugin && typeof window.Asc.plugin.resizeWindow === "function") {
				window.Asc.plugin.resizeWindow(size[0], size[1]);
			}
		} catch (e) {
			/* cosmetic - the wizard still works at the default size */
		}
	}

	/* ------------------------------------------------------------------ *
	 * Merge entry points
	 * ------------------------------------------------------------------ */

	async function startMerge() {
		if (state.running) {
			return;
		}
		if (!state.source.data) {
			statusText(tr("Choose a data source first."), "error");
			return;
		}
		// The re-entrancy latch is set BEFORE the first await: a rapid
		// double-click on Merge arrives while the scan below is still pending,
		// and without the early latch both clicks run interleaved pipelines
		// over the one document.
		state.running = true;
		state.cancelled = false;
		renderAll();
		try {
			if (!state.scanned) {
				var scanned = await scanDocument();
				if (!scanned) {
					return;
				}
			}
			saveSettings(storage(), currentSettings());

			var outcome = await runMergePipeline({
			data: state.source.data,
			paragraphs: state.paragraphs,
			matched: state.mapping.matched,
			unmatched: state.mapping.unmatched,
			allTokens: state.tokens,
			wrapMatched: state.wrapMatched,
			plainUnmatched: state.plainUnmatched,
			mode: state.mode,
			format: state.format,
			rangeText: state.range,
			isCancelled: function () {
				return state.cancelled;
			},
			onProgress: function (label) {
				statusText(label);
			}
		});

			if (outcome.ok) {
				statusText(tr("Merge complete."), "ok");
				// The window closes only after a successful run - errors keep it open.
				closeWindow();
			} else if (outcome.cancelled) {
				statusText(outcome.error);
				if (state.closeAfterRun) {
					closeWindow();
				}
			} else {
				statusText(outcome.error, "error");
			}
		} finally {
			// Reset the latch whatever happened - a throwing pipeline seam must
			// not leave the wizard stuck on "running".
			state.running = false;
			renderAll();
		}
	}

	/**
	 * The dialog footer arrives here as ids into config.json's buttons array:
	 * 0 = Cancel, 1 = Merge, -1 = the X. MANDATORY hook - without it the host's
	 * injected router throws on the X and nothing closes at all.
	 */
	function handleButton(id) {
		var numeric = typeof id === "string" ? parseInt(id, 10) : id;
		if (isNaN(numeric)) {
			return;
		}
		if (state.running) {
			// Cancel while merging: stop at the next stage boundary, then close.
			if (numeric === -1 || numeric === 0) {
				state.cancelled = true;
				state.closeAfterRun = true;
				statusText(tr("Cancelling..."));
			}
			return;
		}
		if (numeric === -1 || numeric === 0) {
			closeWindow();
			return;
		}
		if (numeric === 1) {
			return startMerge();
		}
		return undefined;
	}

	/* ------------------------------------------------------------------ *
	 * Wizard DOM
	 * ------------------------------------------------------------------ */

	var view = {
		wrap: null,
		chips: [],
		panels: [],
		body: null,
		summary: null,
		preview: null,
		tokens: null,
		mapBody: null,
		mapNote: null,
		wrapToggle: null,
		combinedInput: null,
		perRecordInput: null,
		modeNote: null,
		outputWarn: null,
		rangeInput: null,
		formatSelect: null,
		backBtn: null,
		nextBtn: null,
		status: null,
		progress: null
	};

	function statusText(text, kind) {
		if (!view.status) {
			return;
		}
		view.status.textContent = text || "";
		view.status.className = "mm-status" + (kind ? " mm-" + kind : "");
	}

	function buildChip(index) {
		var chip = el("li", "mm-chip");
		chip.appendChild(el("span", "mm-chip-num", String(index + 1)));
		chip.appendChild(document.createTextNode(STEP_TITLES[index]));
		chip.addEventListener("click", function () {
			gotoStep(index);
		});
		return chip;
	}

	function buildSourceStep(parent) {
		parent.appendChild(el("h2", null, tr("Data source")));

		var pick = el("div", "mm-row");
		var input = el("input", "mm-file");
		input.setAttribute("type", "file");
		input.setAttribute("accept", ".csv,.xlsx,.xls,.json");
		input.addEventListener("change", function (event) {
			var target = (event && event.target) || input;
			if (target.files && target.files[0]) {
				readFile(target.files[0]);
			}
		});
		pick.appendChild(input);
		pick.appendChild(el("p", "mm-hint", tr("CSV, XLSX or JSON. Row 1 holds the field names, each later row one recipient.")));
		parent.appendChild(pick);

		view.summary = el("div", "mm-summary", tr("No data source loaded yet."));
		parent.appendChild(view.summary);

		var previewWrap = el("div", "mm-row");
		previewWrap.appendChild(el("p", "mm-hint", tr("Preview (first 10 recipients):")));
		view.preview = el("div", "mm-scroll");
		previewWrap.appendChild(view.preview);
		parent.appendChild(previewWrap);

		var scanRow = el("div", "mm-row");
		var scanBtn = el("button", "mm-btn", tr('Scan document for {{Fields}}'));
		scanBtn.setAttribute("type", "button");
		scanBtn.addEventListener("click", function () {
			scanDocument();
		});
		scanRow.appendChild(scanBtn);
		parent.appendChild(scanRow);

		view.tokens = el("div", "mm-row");
		parent.appendChild(view.tokens);
	}

	function buildFieldsStep(parent) {
		parent.appendChild(el("h2", null, tr("Fields")));

		var wrapRow = el("div", "mm-row");
		var label = el("label", "mm-check");
		view.wrapToggle = el("input", null);
		view.wrapToggle.setAttribute("type", "checkbox");
		view.wrapToggle.checked = state.wrapMatched;
		view.wrapToggle.addEventListener("change", function () {
			state.wrapMatched = !!view.wrapToggle.checked;
			renderOutputWarnings();
		});
		label.appendChild(view.wrapToggle);
		label.appendChild(document.createTextNode(tr("Wrap matched tokens into real merge fields")));
		wrapRow.appendChild(label);
		parent.appendChild(wrapRow);

		view.mapNote = el("p", "mm-hint", "");
		parent.appendChild(view.mapNote);

		var table = el("table", "mm-table");
		var head = el("thead", null);
		var headRow = el("tr", null);
		[tr("{{Token}}"), tr("Data field"), tr("Status"), tr("Fallback")].forEach(function (title) {
			headRow.appendChild(el("th", null, title));
		});
		head.appendChild(headRow);
		table.appendChild(head);
		view.mapBody = el("tbody", null);
		table.appendChild(view.mapBody);

		var scroll = el("div", "mm-scroll");
		scroll.appendChild(table);
		parent.appendChild(scroll);
	}

	function buildOutputStep(parent) {
		parent.appendChild(el("h2", null, tr("Output")));

		var modeRow = el("div", "mm-row");
		var combinedLabel = el("label", "mm-radio");
		view.combinedInput = el("input", null);
		view.combinedInput.setAttribute("type", "radio");
		view.combinedInput.setAttribute("name", "mm-mode");
		view.combinedInput.checked = state.mode === "combined";
		view.combinedInput.addEventListener("change", function () {
			state.mode = "combined";
			renderOutputWarnings();
		});
		combinedLabel.appendChild(view.combinedInput);
		combinedLabel.appendChild(document.createTextNode(tr("Single combined document (default)")));

		var perRecordLabel = el("label", "mm-radio");
		view.perRecordInput = el("input", null);
		view.perRecordInput.setAttribute("type", "radio");
		view.perRecordInput.setAttribute("name", "mm-mode");
		view.perRecordInput.checked = state.mode === "per-record";
		view.perRecordInput.addEventListener("change", function () {
			state.mode = "per-record";
			renderOutputWarnings();
		});
		perRecordLabel.appendChild(view.perRecordInput);
		perRecordLabel.appendChild(document.createTextNode(tr("One file per recipient")));

		modeRow.appendChild(combinedLabel);
		modeRow.appendChild(perRecordLabel);
		parent.appendChild(modeRow);

		view.modeNote = el("p", "mm-hint", "");
		parent.appendChild(view.modeNote);

		var rangeRow = el("div", "mm-row");
		rangeRow.appendChild(el("span", null, tr("Records: ")));
		view.rangeInput = el("input", "mm-range");
		view.rangeInput.setAttribute("type", "text");
		view.rangeInput.setAttribute("placeholder", tr("all (e.g. 1-10,15)"));
		view.rangeInput.value = state.range;
		view.rangeInput.addEventListener("change", function () {
			state.range = String(view.rangeInput.value || "").trim();
		});
		rangeRow.appendChild(view.rangeInput);
		rangeRow.appendChild(el("p", "mm-hint", tr("Empty = every recipient. Numbers are 1-based row numbers below the header.")));
		parent.appendChild(rangeRow);

		var formatRow = el("div", "mm-row");
		formatRow.appendChild(el("span", null, tr("Format: ")));
		view.formatSelect = el("select", "mm-select");
		["docx", "pdf"].forEach(function (format) {
			var option = el("option", null, format);
			option.value = format;
			view.formatSelect.appendChild(option);
		});
		view.formatSelect.value = state.format;
		view.formatSelect.addEventListener("change", function () {
			state.format = view.formatSelect.value === "pdf" ? "pdf" : "docx";
		});
		formatRow.appendChild(view.formatSelect);
		parent.appendChild(formatRow);

		view.outputWarn = el("div", "mm-warning", "");
		parent.appendChild(view.outputWarn);
	}

	function buildFooter(parent) {
		var nav = el("div", "mm-nav");
		view.backBtn = el("button", "mm-btn", tr("Back"));
		view.backBtn.setAttribute("type", "button");
		view.backBtn.addEventListener("click", function () {
			gotoStep(state.step - 1);
		});
		nav.appendChild(view.backBtn);

		var right = el("div", "mm-nav-right");
		view.nextBtn = el("button", "mm-btn mm-primary", tr("Next"));
		view.nextBtn.setAttribute("type", "button");
		view.nextBtn.addEventListener("click", onNext);
		right.appendChild(view.nextBtn);
		nav.appendChild(right);
		parent.appendChild(nav);

		view.status = el("div", "mm-status", "");
		parent.appendChild(view.status);

		view.progress = el("div", "mm-progress");
		view.progress.appendChild(el("div", "mm-progress-bar"));
		parent.appendChild(view.progress);
	}

	async function onNext() {
		if (state.step === 0) {
			if (!state.source.data) {
				statusText(tr("Choose a data source first."), "error");
				return;
			}
			if (!state.scanned) {
				var scanned = await scanDocument();
				if (!scanned) {
					return;
				}
			}
		}
		gotoStep(state.step + 1);
	}

	function gotoStep(index) {
		if (index < 0 || index >= STEP_TITLES.length || isNaN(index)) {
			return;
		}
		// Steps beyond the data source are unreachable without data.
		if (index > 0 && !state.source.data) {
			statusText(tr("Choose a data source first."), "error");
			return;
		}
		if (state.running) {
			return;
		}
		state.step = index;
		renderAll();
		resizeForStep(index);
	}

	function renderSteps() {
		view.chips.forEach(function (chip, index) {
			chip.className =
				"mm-chip" + (index === state.step ? " mm-active" : "") + (index < state.step ? " mm-done" : "");
		});
		view.panels.forEach(function (panel, index) {
			panel.className = "mm-step" + (index === state.step ? " mm-current" : "");
		});
	}

	function renderSource() {
		var data = state.source.data;
		if (!data) {
			view.summary.textContent = tr("No data source loaded yet.");
			clear(view.preview);
			clear(view.tokens);
			return;
		}
		view.summary.textContent =
			data.length - 1 + " " + tr("recipients") + ", " + data[0].length + " " + tr("fields") +
			(state.source.name ? " (" + state.source.name + ")" : "");

		clear(view.preview);
		var table = el("table", "mm-table");
		previewRows(data, PREVIEW_ROWS).forEach(function (row, rowIndex) {
			var trEl = el("tr", null);
			row.forEach(function (cell) {
				trEl.appendChild(el(rowIndex === 0 ? "th" : "td", null, cell));
			});
			table.appendChild(trEl);
		});
		view.preview.appendChild(table);

		clear(view.tokens);
		if (!state.scanned) {
			view.tokens.appendChild(el("p", "mm-hint", tr("The document has not been scanned yet.")));
			return;
		}
		if (!state.tokens.length) {
			view.tokens.appendChild(el("p", "mm-hint", tr("No {{Field}} tokens found in the document.")));
			return;
		}
		view.tokens.appendChild(el("p", "mm-hint", tr("Tokens found in the document:")));
		var list = el("div", "mm-scroll");
		var tokenTable = el("table", "mm-table");
		var body = el("tbody", null);
		state.tokens.forEach(function (token) {
			var row = el("tr", null);
			row.appendChild(el("td", "mm-token", "{{" + token + "}}"));
			body.appendChild(row);
		});
		tokenTable.appendChild(body);
		list.appendChild(tokenTable);
		view.tokens.appendChild(list);
	}

	function renderMapping() {
		clear(view.mapBody);
		var matchedByToken = {};
		state.mapping.matched.forEach(function (entry) {
			matchedByToken[entry.token] = entry.header;
		});

		if (!state.tokens.length) {
			view.mapNote.textContent = state.scanned
				? tr("No {{Field}} tokens found in the document.")
				: tr("Scan the document to list its {{Field}} tokens.");
		} else if (!state.source.data) {
			view.mapNote.textContent = tr("Load a data source to match tokens against its fields.");
		} else {
			view.mapNote.textContent =
				state.mapping.matched.length +
				" " +
				tr("matched") +
				", " +
				state.mapping.unmatched.length +
				" " +
				tr("unmatched") +
				". " +
				tr("Unmatched tokens are replaced as plain text (empty value if the data has no such field).");
		}

		state.tokens.forEach(function (token) {
			var row = el("tr", null);
			row.appendChild(el("td", "mm-token", "{{" + token + "}}"));

			var header = matchedByToken[token];
			var fieldCell = el("td", null, header ? header : tr("(none)"));
			row.appendChild(fieldCell);

			var statusCell = el("td", null);
			if (header) {
				statusCell.appendChild(el("span", "mm-badge mm-ok", tr("matched")));
			} else {
				statusCell.appendChild(el("span", "mm-badge mm-miss", tr("unmatched")));
			}
			row.appendChild(statusCell);

			var fallbackCell = el("td", null);
			if (!header) {
				var label = el("label", "mm-check");
				var checkbox = el("input", null);
				checkbox.setAttribute("type", "checkbox");
				checkbox.checked = state.plainUnmatched[token] !== false;
				checkbox.addEventListener("change", function () {
					state.plainUnmatched[token] = !!checkbox.checked;
					// The combined-mode warning depends on which fallbacks stay on.
					renderOutputWarnings();
				});
				label.appendChild(checkbox);
				label.appendChild(document.createTextNode(tr("plain replace")));
				fallbackCell.appendChild(label);
			} else if (!state.wrapMatched) {
				fallbackCell.appendChild(document.createTextNode(tr("plain replace")));
			} else {
				fallbackCell.appendChild(document.createTextNode(tr("merge field")));
			}
			row.appendChild(fallbackCell);

			view.mapBody.appendChild(row);
		});
	}

	function renderOutputWarnings() {
		if (view.modeNote) {
			view.modeNote.textContent =
				state.mode === "per-record"
					? tr("Opens one save dialog per recipient.")
					: tr("All selected recipients are merged into one document.");
		}
		if (!view.outputWarn) {
			return;
		}
		var warnings = [];
		// The same plain-token set the pipeline will plain-replace: with
		// wrapping on, the unmatched tokens still on "plain replace"; with
		// wrapping off, every token. In a combined document each record's copy
		// of the template shares one replacement, so these tokens cannot vary
		// per recipient and are blanked - warn instead of producing silent
		// blanks (the known limitation of plain replacement in combined mode).
		var plainTokens = state.wrapMatched
			? state.mapping.unmatched.filter(function (token) {
				return state.plainUnmatched[token] !== false;
			})
			: state.tokens.slice();
		if (state.mode === "combined" && plainTokens.length) {
			warnings.push(
				tr("Plain replacement in a combined document cannot vary per recipient - the plain-replaced tokens will be blank. Wrap the tokens or use one file per recipient.")
			);
		}
		view.outputWarn.textContent = warnings.join(" ");
		view.outputWarn.style.display = warnings.length ? "" : "none";
	}

	function renderNav() {
		view.backBtn.style.display = state.step === 0 ? "none" : "";
		view.nextBtn.style.display = state.step === STEP_TITLES.length - 1 ? "none" : "";
		view.nextBtn.disabled = state.running;
		view.backBtn.disabled = state.running;
		if (view.progress) {
			view.progress.className = "mm-progress" + (state.running ? " mm-running" : "");
		}
	}

	function renderAll() {
		if (!view.wrap) {
			return;
		}
		renderSteps();
		renderSource();
		renderMapping();
		renderOutputWarnings();
		renderNav();
	}

	function mount(root) {
		if (!root) {
			return;
		}
		clear(root);
		view.wrap = el("div", "mm-wrap");

		var stepsBar = el("ol", "mm-steps");
		view.chips = [];
		view.panels = [];
		for (var i = 0; i < STEP_TITLES.length; i++) {
			view.chips.push(buildChip(i));
			stepsBar.appendChild(view.chips[i]);
		}
		view.wrap.appendChild(stepsBar);

		view.body = el("div", "mm-body");
		var builders = [buildSourceStep, buildFieldsStep, buildOutputStep];
		for (var b = 0; b < builders.length; b++) {
			var panel = el("section", "mm-step");
			builders[b](panel);
			view.panels.push(panel);
			view.body.appendChild(panel);
		}
		view.wrap.appendChild(view.body);

		buildFooter(view.wrap);
		root.appendChild(view.wrap);
		renderAll();
	}

	/* ------------------------------------------------------------------ *
	 * Host hooks
	 * ------------------------------------------------------------------ */

	function applyTheme(theme) {
		var dark = isDarkTheme(theme);
		var targets = [document.documentElement, document.body];
		targets.forEach(function (node) {
			if (!node || !node.classList) {
				return;
			}
			["theme-dark", "theme-type-dark"].forEach(function (name) {
				node.classList.remove(name);
			});
			["theme-light", "theme-type-light"].forEach(function (name) {
				node.classList.remove(name);
			});
			var names = dark ? ["theme-dark", "theme-type-dark"] : ["theme-light", "theme-type-light"];
			names.forEach(function (name) {
				node.classList.add(name);
			});
		});
	}

	function installHostHooks() {
		if (!window.Asc || !window.Asc.plugin) {
			return;
		}
		window.Asc.plugin.init = function () {
			var settings = loadSettings(storage());
			state.wrapMatched = settings.wrapMatched;
			state.mode = settings.mode;
			state.range = settings.range;
			state.format = settings.format;
			if (view.wrapToggle) {
				view.wrapToggle.checked = state.wrapMatched;
			}
			if (view.rangeInput) {
				view.rangeInput.value = state.range;
			}
			if (view.formatSelect) {
				view.formatSelect.value = state.format;
			}
			if (view.combinedInput) {
				view.combinedInput.checked = state.mode === "combined";
				view.perRecordInput.checked = state.mode === "per-record";
			}
			renderAll();
		};

		// MANDATORY: variation buttons arrive as ids; -1 is the X.
		window.Asc.plugin.button = function (id, windowId) {
			handleButton(id, windowId);
		};

		window.Asc.plugin.onThemeChanged = function (theme) {
			applyTheme(theme);
			if (typeof window.Asc.plugin.onThemeChangedBase === "function") {
				window.Asc.plugin.onThemeChangedBase(theme);
			}
		};

		// The earliest point at which Asc.plugin.tr() returns real translations.
		window.Asc.plugin.onTranslate = function () {
			renderAll();
		};
	}

	// Public surface: the dev console and the tests drive the same seams the
	// wizard does - nothing here is test-only behaviour.
	window.OnlyOfficeMailMergeUi = {
		state: state,
		parseRange: parseRange,
		isContiguous: isContiguous,
		defaultSettings: defaultSettings,
		loadSettings: loadSettings,
		saveSettings: saveSettings,
		currentSettings: currentSettings,
		runMergePipeline: runMergePipeline,
		plainRecord: plainRecord,
		previewRows: previewRows,
		loadSourceText: loadSourceText,
		parseWorkbook: parseWorkbook,
		scanDocument: scanDocument,
		startMerge: startMerge,
		handleButton: handleButton,
		gotoStep: gotoStep,
		getWindowId: getWindowId,
		closeWindow: closeWindow,
		mount: mount,
		applyTheme: applyTheme
	};

	installHostHooks();
	mount(document.getElementById("mm-app"));
})(window, document);
