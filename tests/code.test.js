/*
 * Wizard + pipeline tests for plugin/scripts/code.js.
 *
 * code.js is the composition root: it sequences the command seam (commands.js),
 * the parser (dataparse.js) and the mapper (fieldmap.js) behind a three step
 * DOM wizard. The three modules belong to other agents and are still stubs, so
 * this harness injects faithful fakes at their UMD boundaries and loads the real
 * code.js into a minimal fake window/document - the same seam the reference
 * project's plugin-harness.js uses (only far smaller: this plugin renders DOM,
 * it does not register ribbon menus).
 *
 * What is covered here: parseRange, settings persistence, wizard mounting and
 * step gating, the data summary/preview, the document scan, the host button
 * routing (including "close only after success"), and the exact command
 * sequencing of the combined and per-record pipelines with error and cancel
 * paths.
 */
"use strict";

const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");

const ROOT = path.join(__dirname, "..");
const CODE_SOURCE = fs.readFileSync(path.join(ROOT, "plugin", "scripts", "code.js"), "utf8");

/* ------------------------------------------------------------------ *
 * Minimal fake DOM - only what code.js touches
 * ------------------------------------------------------------------ */

function El(tag) {
	this.tagName = String(tag).toUpperCase();
	this.children = [];
	this.parentNode = null;
	this.attributes = {};
	this.style = {};
	this.className = "";
	this.value = "";
	this.checked = false;
	this.disabled = false;
	this.listeners = {};
	this._text = "";
	const self = this;
	this.classList = {
		add: function (name) {
			if (!self.classList.contains(name)) {
				self.className = (self.className ? self.className + " " : "") + name;
			}
		},
		remove: function (name) {
			self.className = self.className
				.split(/\s+/)
				.filter(function (piece) {
					return piece && piece !== name;
				})
				.join(" ");
		},
		contains: function (name) {
			return self.className.split(/\s+/).indexOf(name) !== -1;
		},
		toggle: function (name) {
			if (self.classList.contains(name)) {
				self.classList.remove(name);
			} else {
				self.classList.add(name);
			}
		}
	};
}

Object.defineProperty(El.prototype, "textContent", {
	get: function () {
		return (
			this._text +
			this.children
				.map(function (child) {
					return child.nodeType === 3 ? child.data : child.textContent;
				})
				.join("")
		);
	},
	set: function (value) {
		this._text = String(value);
		this.children.forEach(function (child) {
			child.parentNode = null;
		});
		this.children = [];
	}
});

Object.defineProperty(El.prototype, "firstChild", {
	get: function () {
		return this.children.length ? this.children[0] : null;
	}
});

El.prototype.appendChild = function (child) {
	child.parentNode = this;
	this.children.push(child);
	return child;
};

El.prototype.removeChild = function (child) {
	const index = this.children.indexOf(child);
	if (index !== -1) {
		this.children.splice(index, 1);
	}
	child.parentNode = null;
	return child;
};

El.prototype.addEventListener = function (type, handler) {
	(this.listeners[type] = this.listeners[type] || []).push(handler);
};

El.prototype.setAttribute = function (name, value) {
	this.attributes[name] = String(value);
	if (name === "class") {
		this.className = String(value);
	}
};

El.prototype.getAttribute = function (name) {
	return this.attributes[name];
};

/** Fire a DOM event the way a user interaction would. */
El.prototype.dispatch = function (type, event) {
	const self = this;
	(this.listeners[type] || []).slice().forEach(function (handler) {
		handler.call(self, event || { target: self });
	});
};

function findAll(node, predicate) {
	const found = [];
	(function walk(current) {
		if (current.nodeType === 3) {
			return;
		}
		if (predicate(current)) {
			found.push(current);
		}
		current.children.forEach(walk);
	})(node);
	return found;
}

function byClass(root, name) {
	return findAll(root, function (node) {
		return node.classList.contains(name);
	});
}

/* ------------------------------------------------------------------ *
 * Harness: fake window + fake contract modules + the real code.js
 * ------------------------------------------------------------------ */

function makeStorage() {
	const data = {};
	return {
		getItem: function (key) {
			return Object.prototype.hasOwnProperty.call(data, key) ? data[key] : null;
		},
		setItem: function (key, value) {
			data[key] = String(value);
		},
		removeItem: function (key) {
			delete data[key];
		},
		dump: function () {
			return data;
		}
	};
}

const DEFAULT_DATA = [
	["Name", "City", "Team"],
	["Ann", "Rome", "red"],
	["Bob", "Oslo", "blue"],
	["Cid", "Kyoto", "red"]
];

function makePlugin(log, options) {
	return {
		windowID: "sdk-id",
		tr: function (text) {
			return text;
		},
		resizeWindow: function (width, height) {
			log.push({ kind: "method", name: "resizeWindow", args: [width, height] });
		},
		executeMethod: function (name, args, callback) {
			log.push({ kind: "method", name: name, args: args });
			if (name === "GetFileToDownload") {
				queueMicrotask(function () {
					callback(options.downloadResult === undefined ? "blob://saved" : options.downloadResult);
				});
			} else if (typeof callback === "function") {
				queueMicrotask(function () {
					callback("ok");
				});
			}
		}
	};
}

function boot(options) {
	options = options || {};
	const log = []; // every seam call, in order: {kind:"run"|"method", name, payload|args}
	const storage = makeStorage();

	const DataParse = {
		parseCsv: function () {
			return { ok: true, data: options.data || DEFAULT_DATA, warnings: [] };
		},
		parseJson: function () {
			return { ok: true, data: options.data || DEFAULT_DATA, warnings: [] };
		},
		parseGrid: function (grid) {
			return { ok: true, data: grid, warnings: [] };
		},
		toLoadMailMergeData: function (result) {
			return result.data;
		}
	};

	const FieldMap = {
		TOKEN_RE: /\{\{\s*([^{}]+?)\s*\}\}/g,
		normalizeName: function (name) {
			return String(name).trim().replace(/\s+/g, " ");
		},
		findTokens: function () {
			return (options.tokens || ["Name", "Missing"]).slice();
		},
		findTokensInTexts: function () {
			return (options.tokens || ["Name", "Missing"]).slice();
		},
		matchFields: function (tokens, headers) {
			const matched = [];
			const unmatched = [];
			tokens.forEach(function (token) {
				const hit = headers.filter(function (header) {
					return header.toLowerCase() === token.toLowerCase();
				})[0];
				if (hit) {
					matched.push({ token: token, header: hit });
				} else {
					unmatched.push(token);
				}
			});
			return {
				matched: matched,
				unmatched: unmatched,
				unusedHeaders: headers.filter(function (header) {
					return !matched.some(function (entry) {
						return entry.header === header;
					});
				})
			};
		},
		plainReplacePlan: function (text, record, tokens) {
			return (tokens || []).map(function (token) {
				return {
					searchString: "{{" + token + "}}",
					replaceString: record[token] === undefined || record[token] === null ? "" : String(record[token])
				};
			});
		}
	};

	const Commands = {
		run: function (name, payload, callback) {
			log.push({ kind: "run", name: name, payload: payload });
			let result;
			if (options.failCommand === name) {
				result = { ok: false, error: "boom " + name };
			} else if (name === "getDocumentText") {
				result = { ok: true, paragraphs: options.paragraphs || ["Dear {{Name}}, cc {{Missing}}."] };
			} else {
				result = { ok: true };
			}
			queueMicrotask(function () {
				callback(result);
			});
		}
	};

	// The load-order tests drive the document the way a real parse does:
	// code.js evaluates before #mm-app exists (deferApp) while readyState is
	// still "loading", and boot only runs when DOMContentLoaded fires.
	const root = new El("div");
	root.setAttribute("id", "mm-app");

	let appReady = !options.deferApp;
	const docListeners = {};

	const document = {
		createElement: function (tag) {
			return new El(tag);
		},
		createTextNode: function (text) {
			return { nodeType: 3, data: String(text), textContent: String(text) };
		},
		getElementById: function (id) {
			return id === "mm-app" && appReady ? root : null;
		},
		documentElement: new El("html"),
		body: new El("body"),
		readyState: options.readyState || "complete",
		addEventListener: function (type, handler) {
			(docListeners[type] = docListeners[type] || []).push(handler);
		}
	};

	const win = {
		location: { search: options.search !== undefined ? options.search : "?windowID=win-42" },
		localStorage: storage,
		Papa: undefined,
		XLSX: undefined,
		OnlyOfficeMailMergeDataParse: DataParse,
		OnlyOfficeMailMergeFieldMap: FieldMap,
		OnlyOfficeMailMergeCommands: Commands
	};
	if (!options.noAsc) {
		win.Asc = { plugin: makePlugin(log, options) };
	}

	// Run the real code.js in THIS realm (a Function wrapper supplying the two
	// free variables its IIFE is called with). A vm context would give every
	// promise/array its own prototype and break deepStrictEqual across realms.
	new Function("window", "document", CODE_SOURCE)(win, document);
	assert.ok(win.OnlyOfficeMailMergeUi, "code.js published its UI surface");
	return {
		win: win,
		doc: document,
		root: root,
		log: log,
		storage: storage,
		ui: win.OnlyOfficeMailMergeUi,
		names: function () {
			return log.map(function (entry) {
				return entry.name;
			});
		},
		fireDocument: function (type) {
			(docListeners[type] || []).slice().forEach(function (handler) {
				handler.call(document, { type: type });
			});
		},
		setAppReady: function () {
			appReady = true;
		},
		installFakeHost: function () {
			win.Asc = { plugin: makePlugin(log, options) };
			return win.Asc.plugin;
		}
	};
}

/* ------------------------------------------------------------------ *
 * parseRange
 * ------------------------------------------------------------------ */

test("parseRange: an empty range selects every recipient", () => {
	const { ui } = boot();
	assert.deepStrictEqual(ui.parseRange("", 3), { ok: true, indices: [0, 1, 2] });
	assert.deepStrictEqual(ui.parseRange("   ", 2), { ok: true, indices: [0, 1] });
});

test("parseRange: 1-based segments become sorted 0-based indices", () => {
	const { ui } = boot();
	assert.deepStrictEqual(ui.parseRange("1-3,5", 6), { ok: true, indices: [0, 1, 2, 4] });
	assert.deepStrictEqual(ui.parseRange("15, 2 - 4 , 2", 20), { ok: true, indices: [1, 2, 3, 14] });
	assert.deepStrictEqual(ui.parseRange("2", 2), { ok: true, indices: [1] });
});

test("parseRange: bad input fails with a readable error", () => {
	const { ui } = boot();
	["0", "abc", "5-2", "1-2,", "99", "1;2"].forEach(function (text) {
		const result = ui.parseRange(text, 5);
		assert.strictEqual(result.ok, false, text + " must be rejected");
		assert.ok(result.error, text + " names the problem");
	});
});

test("isContiguous: a gapped selection is not one MailMerge span", () => {
	const { ui } = boot();
	assert.strictEqual(ui.isContiguous([0, 1, 2]), true);
	assert.strictEqual(ui.isContiguous([0, 2]), false);
	assert.strictEqual(ui.isContiguous([3]), true);
});

/* ------------------------------------------------------------------ *
 * Settings
 * ------------------------------------------------------------------ */

test("settings: defaults, round-trip and version migration", () => {
	const { ui, storage } = boot();
	const defaults = ui.defaultSettings();
	assert.strictEqual(defaults.wrapMatched, true, "wrapping is on by default");
	assert.strictEqual(defaults.mode, "combined", "combined output is the default");
	assert.strictEqual(defaults.format, "docx");

	const stored = { version: 1, wrapMatched: false, mode: "per-record", range: "2-3", format: "pdf" };
	ui.saveSettings(storage, stored);
	assert.deepStrictEqual(ui.loadSettings(storage), {
		version: 1,
		wrapMatched: false,
		mode: "per-record",
		range: "2-3",
		format: "pdf"
	});

	storage.setItem("onlyoffice-mail-merge.settings", JSON.stringify({ version: 99, mode: "per-record" }));
	assert.deepStrictEqual(ui.loadSettings(storage), defaults, "an unknown schema degrades to defaults");
	storage.setItem("onlyoffice-mail-merge.settings", "not json");
	assert.deepStrictEqual(ui.loadSettings(storage), defaults, "corrupt storage degrades to defaults");
});

/* ------------------------------------------------------------------ *
 * Wizard mounting and steps
 * ------------------------------------------------------------------ */

test("the wizard mounts three steps, a file picker and the nav buttons", () => {
	const { root } = boot();
	assert.strictEqual(byClass(root, "mm-chip").length, 3, "one chip per step");
	assert.strictEqual(byClass(root, "mm-step").length, 3, "one panel per step");
	assert.ok(byClass(root, "mm-current").length === 1, "exactly one step is current");

	const fileInput = byClass(root, "mm-file")[0];
	assert.ok(fileInput, "the data source step has a file picker");
	assert.strictEqual(fileInput.getAttribute("type"), "file");
	assert.match(fileInput.getAttribute("accept"), /\.csv/);
	assert.match(fileInput.getAttribute("accept"), /\.xlsx/);
	assert.match(fileInput.getAttribute("accept"), /\.json/);

	assert.ok(byClass(root, "mm-status")[0], "a status area exists");
	assert.ok(byClass(root, "mm-progress")[0], "a progress bar exists");
});

test("gotoStep: the later steps are unreachable without a data source", () => {
	const h = boot();
	h.ui.gotoStep(1);
	assert.strictEqual(h.ui.state.step, 0, "step 2 stays locked");
	assert.match(byClass(h.root, "mm-status")[0].textContent, /data source/i);

	h.ui.loadSourceText("recipients.csv", "ignored");
	h.ui.gotoStep(2);
	assert.strictEqual(h.ui.state.step, 2, "with data loaded every step opens");
	assert.ok(h.log.some(function (entry) {
		return entry.name === "resizeWindow";
	}), "the window is resized per step");
});

test("loadSourceText: parses, summarizes and previews the first 10 rows", () => {
	const rows = [["Name"]];
	for (let i = 1; i <= 12; i++) {
		rows.push(["r" + i]);
	}
	const h = boot({ data: rows });
	const applied = h.ui.loadSourceText("recipients.csv", "raw");
	assert.strictEqual(applied, true);
	assert.strictEqual(h.ui.state.source.data.length, 13);

	const summary = byClass(h.root, "mm-summary")[0].textContent;
	assert.match(summary, /12 recipients/);
	assert.match(summary, /1 fields/);
	assert.match(summary, /recipients\.csv/);

	const previewRows = byClass(h.root, "mm-scroll")[0].children[0].children;
	assert.strictEqual(previewRows.length, 11, "header + first 10 recipients only");

	assert.strictEqual(h.ui.previewRows(rows, 10).length, 11);
});

test("scanDocument: lists template tokens and computes the mapping", async () => {
	const h = boot();
	h.ui.loadSourceText("recipients.csv", "raw");
	const scanned = await h.ui.scanDocument();
	assert.strictEqual(scanned, true);
	assert.deepStrictEqual(h.ui.state.tokens, ["Name", "Missing"]);
	assert.deepStrictEqual(h.ui.state.mapping.matched, [{ token: "Name", header: "Name" }]);
	assert.deepStrictEqual(h.ui.state.mapping.unmatched, ["Missing"]);
	assert.ok(h.log.some(function (entry) {
		return entry.name === "getDocumentText";
	}));

	const mapRows = findAll(h.root, function (node) {
		return node.tagName === "TR" && node.children.length === 4 && node.children[0].tagName === "TD";
	});
	assert.strictEqual(mapRows.length, 2, "one mapping row per token");
	const badges = byClass(h.root, "mm-badge");
	assert.strictEqual(badges.length, 2);
	assert.ok(badges[0].classList.contains("mm-ok"));
	assert.ok(badges[1].classList.contains("mm-miss"));
});

/* ------------------------------------------------------------------ *
 * Host window plumbing
 * ------------------------------------------------------------------ */

test("handleButton: -1 and 0 close the window with the URL's windowID", () => {
	const h = boot();
	h.ui.handleButton(-1);
	h.ui.handleButton(0);
	const closes = h.log.filter(function (entry) {
		return entry.name === "CloseWindow";
	});
	assert.strictEqual(closes.length, 2, "both ids close");
	assert.deepStrictEqual(closes[0].args, ["win-42"], "the id comes from the page URL, before the SDK publishes one");
});

test("getWindowId: falls back to the SDK value when the URL has none", () => {
	const h = boot({ search: "" });
	assert.strictEqual(h.ui.getWindowId(), "sdk-id");
});

test("close: the windowId hint that arrives with the click wins over the URL and SDK ids", () => {
	const h = boot(); // URL ?windowID=win-42, SDK windowID sdk-id
	// The host dispatches the id as the button hook's second argument
	// (`Asc.plugin.button(k, g.buttonWindowId)`); the direct seam is the same
	// path hostButton forwards to.
	h.win.Asc.plugin.button(0, "win-9");
	h.ui.handleButton(-1, "win-9");
	const closes = h.log.filter(function (entry) {
		return entry.name === "CloseWindow";
	});
	assert.strictEqual(closes.length, 2, "both ids close");
	assert.deepStrictEqual(closes[0].args, ["win-9"], "the hint beats the page URL's windowID");
	assert.deepStrictEqual(closes[1].args, ["win-9"], "the hint beats the SDK-published windowID");
});

test("close: no id anywhere falls through to bare CloseWindow and executeCommand('close')", () => {
	const calls = [];
	const swallow = function () {};

	const h = boot({ search: "" });
	delete h.win.Asc.plugin.windowID;
	h.win.Asc.plugin.executeMethod = function (name, args) {
		calls.push(["executeMethod", name, args]);
		throw new Error("host refused the bare close");
	};
	h.win.Asc.plugin.executeCommand = function (command, param) {
		calls.push(["executeCommand", command, param]);
	};
	const original = console.error;
	console.error = swallow;
	try {
		h.ui.handleButton(-1);
	} finally {
		console.error = original;
	}
	assert.deepStrictEqual(
		calls,
		[
			["executeMethod", "CloseWindow", []],
			["executeCommand", "close", ""]
		],
		"bare CloseWindow first, then the shim's own default close"
	);

	// executeMethod missing entirely must reach the same final fallback.
	const bare = boot({ search: "" });
	delete bare.win.Asc.plugin.windowID;
	delete bare.win.Asc.plugin.executeMethod;
	const bareCalls = [];
	bare.win.Asc.plugin.executeCommand = function (command, param) {
		bareCalls.push([command, param]);
	};
	console.error = swallow;
	try {
		bare.ui.handleButton(-1);
	} finally {
		console.error = original;
	}
	assert.deepStrictEqual(bareCalls, [["close", ""]], "an unavailable executeMethod still closes via executeCommand");
});

test("close: a silently succeeding bare CloseWindow still reaches executeCommand('close')", () => {
	// Real-host behaviour: pluginMethod_CloseWindow accepts anything and
	// no-ops on ids it does not know - it never throws. The old chain
	// returned on that "success" and the main window could never close.
	const calls = [];
	const h = boot({ search: "" });
	delete h.win.Asc.plugin.windowID;
	h.win.Asc.plugin.executeMethod = function (name, args) {
		calls.push(["executeMethod", name, args]);
	};
	h.win.Asc.plugin.executeCommand = function (command, param) {
		calls.push(["executeCommand", command, param]);
	};
	h.ui.handleButton(-1);
	assert.deepStrictEqual(
		calls,
		[
			["executeMethod", "CloseWindow", []],
			["executeCommand", "close", ""]
		],
		"executeCommand is unconditional - CloseWindow success means nothing for the main window"
	);
});

/* ------------------------------------------------------------------ *
 * Pipeline sequencing
 * ------------------------------------------------------------------ */

test("Merge runs the combined pipeline in order and closes only after success", async () => {
	const h = boot();
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.scanDocument();
	const outcome = await h.ui.handleButton(1);
	assert.deepStrictEqual(outcome, undefined, "handleButton returns the run promise only for tests to await");

	// Everything the pipeline does, in order - the seam is serialized.
	assert.deepStrictEqual(h.names(), [
		"getDocumentText", // the scan
		"loadMergeData",
		"wrapFields",
		"snapshotTemplate",
		"mergeRange",
		"replacePlain",
		"GetFileToDownload",
		"CloseWindow"
	]);

	const merge = h.log.filter(function (entry) {
		return entry.name === "mergeRange";
	})[0];
	assert.deepStrictEqual(merge.payload, { start: 0, end: 2 }, "one span over every recipient");

	const wrap = h.log.filter(function (entry) {
		return entry.name === "wrapFields";
	})[0];
	assert.deepStrictEqual(wrap.payload, { tokens: ["Name"] }, "only the matched token becomes a merge field");

	const plain = h.log.filter(function (entry) {
		return entry.name === "replacePlain";
	})[0];
	assert.deepStrictEqual(
		plain.payload.plan.map(function (step) {
			return step.searchString;
		}),
		["{{Missing}}"],
		"the unmatched token is plain-replaced"
	);

	const download = h.log.filter(function (entry) {
		return entry.name === "GetFileToDownload";
	})[0];
	assert.deepStrictEqual(download.args, ["docx"], "the default format");
});

test("per-record mode saves each record and restores the template between them", async () => {
	const h = boot();
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.scanDocument();
	h.ui.state.mode = "per-record";
	await h.ui.handleButton(1);

	assert.deepStrictEqual(h.names(), [
		"getDocumentText",
		"loadMergeData",
		"wrapFields",
		"snapshotTemplate",
		"mergeRange",
		"replacePlain",
		"GetFileToDownload",
		"restoreTemplate",
		"mergeRange",
		"replacePlain",
		"GetFileToDownload",
		"restoreTemplate",
		"mergeRange",
		"replacePlain",
		"GetFileToDownload",
		"restoreTemplate",
		"CloseWindow"
	]);

	const merges = h.log.filter(function (entry) {
		return entry.name === "mergeRange";
	});
	assert.deepStrictEqual(merges.map(function (entry) {
		return entry.payload;
	}), [{ start: 0, end: 0 }, { start: 1, end: 1 }, { start: 2, end: 2 }]);

	const saves = h.log.filter(function (entry) {
		return entry.name === "GetFileToDownload";
	});
	assert.strictEqual(saves.length, 3, "one save dialog per recipient");
});

test("the per-record plain plan carries each record's values", async () => {
	const h = boot();
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.scanDocument();
	h.ui.state.mode = "per-record";
	await h.ui.handleButton(1);

	const plans = h.log.filter(function (entry) {
		return entry.name === "replacePlain";
	}).map(function (entry) {
		return entry.payload.plan[0].replaceString;
	});
	assert.deepStrictEqual(plans, ["", "", ""], "the unmatched token has no value to fill");
});

test("a failing command stops the pipeline and never closes the window", async () => {
	const h = boot({ failCommand: "snapshotTemplate" });
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.scanDocument();
	await h.ui.handleButton(1);

	assert.deepStrictEqual(h.names(), ["getDocumentText", "loadMergeData", "wrapFields", "snapshotTemplate"]);
	assert.ok(!h.names().includes("CloseWindow"), "an error keeps the window open");
	assert.ok(!h.names().includes("GetFileToDownload"), "nothing is saved after a failure");

	const status = byClass(h.root, "mm-status")[0];
	assert.match(status.textContent, /snapshot/i, "the failing step is named");
	assert.ok(status.classList.contains("mm-error"));
	assert.strictEqual(h.ui.state.running, false, "the wizard is usable again");
});

test("a failed save restores the template and keeps the window open", async () => {
	const h = boot({ downloadResult: "error" });
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.scanDocument();
	await h.ui.handleButton(1);

	assert.ok(h.names().includes("restoreTemplate"), "the merged body is rolled back");
	assert.ok(!h.names().includes("CloseWindow"));
	assert.match(byClass(h.root, "mm-status")[0].textContent, /sav/i);
});

test("cancel aborts at the next stage boundary after unwinding", async () => {
	const h = boot();
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.scanDocument();
	h.ui.state.mode = "per-record";

	let runCalls = 0;
	const outcome = await h.ui.runMergePipeline({
		data: h.ui.state.source.data,
		paragraphs: h.ui.state.paragraphs,
		matched: h.ui.state.mapping.matched,
		unmatched: h.ui.state.mapping.unmatched,
		allTokens: h.ui.state.tokens,
		wrapMatched: true,
		mode: "per-record",
		format: "docx",
		rangeText: "",
		exec: function (name, payload) {
			runCalls++;
			h.log.push({ kind: "run", name: name, payload: payload });
			return new Promise(function (resolve) {
				queueMicrotask(function () {
					resolve({ ok: true });
				});
			});
		},
		download: function () {
			return Promise.resolve({ ok: true, url: "blob://saved" });
		},
		isCancelled: function () {
			return runCalls >= 6;
		},
		onProgress: function () {}
	});

	assert.strictEqual(outcome.ok, false);
	assert.strictEqual(outcome.cancelled, true);
	assert.strictEqual(h.names().filter(function (name) {
		return name === "mergeRange";
	}).length, 1, "only the record in flight was merged");
	assert.ok(h.names().includes("restoreTemplate"), "the template is back");
	assert.strictEqual(runCalls, 6, "the boundary check is per stage");
});

test("cancel before any command runs touches nothing", async () => {
	const h = boot();
	const outcome = await h.ui.runMergePipeline({
		data: DEFAULT_DATA,
		paragraphs: [],
		matched: [],
		unmatched: [],
		allTokens: [],
		wrapMatched: true,
		mode: "combined",
		format: "docx",
		rangeText: "",
		isCancelled: function () {
			return true;
		}
	});
	assert.strictEqual(outcome.cancelled, true);
	assert.deepStrictEqual(h.log, [], "no editor command was issued");
});

test("handleButton(0) while running flags a cancel and closes after the unwind", () => {
	const h = boot();
	h.ui.state.running = true;
	h.ui.handleButton(0);
	assert.strictEqual(h.ui.state.cancelled, true);
	assert.strictEqual(h.ui.state.closeAfterRun, true);
	assert.ok(!h.names().includes("CloseWindow"), "the window closes only once the pipeline has unwound");
});

test("a cancel while running closes later with the windowId captured from the click", async () => {
	const h = boot();
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.scanDocument();
	const run = h.ui.handleButton(1);
	h.ui.handleButton(0, "win-9");
	assert.strictEqual(h.ui.state.cancelled, true);
	assert.strictEqual(h.ui.state.closeAfterRun, true);
	assert.strictEqual(h.ui.state.closeWindowId, "win-9", "the cancel click's hint is captured");
	assert.ok(!h.names().includes("CloseWindow"), "the window closes only once the pipeline has unwound");

	await run;
	const closes = h.log.filter(function (entry) {
		return entry.name === "CloseWindow";
	});
	assert.strictEqual(closes.length, 1, "the unwind closes exactly once");
	assert.deepStrictEqual(closes[0].args, ["win-9"], "the deferred close uses the captured hint, not a re-derived id");
});

test("wrapping off replaces every token as plain text", async () => {
	const h = boot();
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.scanDocument();
	h.ui.state.wrapMatched = false;
	await h.ui.handleButton(1);

	assert.ok(!h.names().includes("wrapFields"), "nothing is wrapped");
	const plain = h.log.filter(function (entry) {
		return entry.name === "replacePlain";
	})[0];
	assert.deepStrictEqual(
		plain.payload.plan.map(function (step) {
			return step.searchString;
		}),
		["{{Name}}", "{{Missing}}"],
		"matched and unmatched tokens are both plain-replaced"
	);
	assert.match(byClass(h.root, "mm-warning")[0].textContent, /combined/i, "the output step warns about the loss");
});

test("a gapped range is refused in combined mode before any command runs", async () => {
	const h = boot();
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.scanDocument();
	const outcome = await h.ui.runMergePipeline({
		data: h.ui.state.source.data,
		paragraphs: h.ui.state.paragraphs,
		matched: h.ui.state.mapping.matched,
		unmatched: h.ui.state.mapping.unmatched,
		allTokens: h.ui.state.tokens,
		wrapMatched: true,
		mode: "combined",
		format: "docx",
		rangeText: "1,3"
	});
	assert.strictEqual(outcome.ok, false);
	assert.match(outcome.error, /contiguous/i);
	assert.deepStrictEqual(h.names(), ["getDocumentText"], "the scan happened, but the refusal precedes every pipeline command");
});

test("plainRecord: values are keyed by header and mirrored under the token", () => {
	const { ui } = boot();
	const record = ui.plainRecord(DEFAULT_DATA, 1, [{ token: "Name", header: "Name" }], ["Missing"]);
	assert.strictEqual(record.Name, "Ann");
	assert.strictEqual(record.City, "Rome");
	assert.strictEqual(record.Missing, "", "an unmatched token blanks out");

	const combined = ui.plainRecord(DEFAULT_DATA, -1, [{ token: "Name", header: "Name" }], []);
	assert.strictEqual(combined.Name, "", "a combined record cannot carry one recipient's values");
});

test("merge only closes the window once the save dialog has answered", async () => {
	const h = boot();
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.scanDocument();
	await h.ui.handleButton(1);
	const names = h.names();
	assert.ok(
		names.indexOf("GetFileToDownload") < names.indexOf("CloseWindow"),
		"the save dialog comes before CloseWindow"
	);
});

/* ------------------------------------------------------------------ *
 * Edge cases: degenerate data, gapped ranges, cancel/error unwinding,
 * re-entrancy and UI resilience
 * ------------------------------------------------------------------ */

test("zero recipients is refused before any merge command runs", async () => {
	// A header-only source yields data.length === 1 - loadMergeData would
	// report count 0, so the pipeline refuses before it loads anything.
	const h = boot({ data: [["Name"]] });
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.handleButton(1);

	assert.deepStrictEqual(h.names(), ["getDocumentText"], "not even loadMergeData is reached");
	assert.match(byClass(h.root, "mm-status")[0].textContent, /no recipient rows/i);
	assert.strictEqual(h.ui.state.running, false, "the wizard is usable again");
});

test("exactly one recipient merges a single {start:0,end:0} record", async () => {
	const h = boot({ data: [["Name"], ["Ann"]] });
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.handleButton(1);

	const merges = h.log.filter(function (entry) {
		return entry.name === "mergeRange";
	});
	assert.deepStrictEqual(merges.map(function (entry) {
		return entry.payload;
	}), [{ start: 0, end: 0 }]);
	assert.strictEqual(h.log.filter(function (entry) {
		return entry.name === "GetFileToDownload";
	}).length, 1, "one save dialog for the one recipient");
	assert.ok(h.names().includes("CloseWindow"));
});

test('a gapped range "1-3,7" merges exactly those records, one file each', async () => {
	const rows = [["Name"]];
	for (let i = 1; i <= 7; i++) {
		rows.push(["r" + i]);
	}
	const h = boot({ data: rows });
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.scanDocument();
	h.ui.state.mode = "per-record";
	h.ui.state.range = "1-3,7";
	await h.ui.handleButton(1);

	const merges = h.log.filter(function (entry) {
		return entry.name === "mergeRange";
	});
	assert.deepStrictEqual(merges.map(function (entry) {
		return entry.payload;
	}), [{ start: 0, end: 0 }, { start: 1, end: 1 }, { start: 2, end: 2 }, { start: 6, end: 6 }]);
	assert.strictEqual(h.log.filter(function (entry) {
		return entry.name === "GetFileToDownload";
	}).length, 4, "one save dialog per selected recipient");
	assert.ok(h.names().includes("CloseWindow"));
});

test('the same gapped "1-3,7" range is refused in combined mode', async () => {
	const rows = [["Name"]];
	for (let i = 1; i <= 7; i++) {
		rows.push(["r" + i]);
	}
	const h = boot({ data: rows });
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.scanDocument();
	const outcome = await h.ui.runMergePipeline({
		data: h.ui.state.source.data,
		paragraphs: h.ui.state.paragraphs,
		matched: h.ui.state.mapping.matched,
		unmatched: h.ui.state.mapping.unmatched,
		allTokens: h.ui.state.tokens,
		wrapMatched: true,
		mode: "combined",
		format: "docx",
		rangeText: "1-3,7"
	});
	assert.strictEqual(outcome.ok, false);
	assert.match(outcome.error, /contiguous/i);
	assert.deepStrictEqual(h.names(), ["getDocumentText"], "the refusal precedes every pipeline command");
});

test("an error mid per-record loop restores the template and stops", async () => {
	const h = boot({ failCommand: "replacePlain" });
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.scanDocument();
	h.ui.state.mode = "per-record";
	await h.ui.handleButton(1);

	assert.strictEqual(h.log.filter(function (entry) {
		return entry.name === "mergeRange";
	}).length, 1, "only the failing record was merged");
	assert.ok(h.names().includes("restoreTemplate"), "the template is back");
	assert.ok(!h.names().includes("CloseWindow"), "an error keeps the window open");
	assert.ok(!h.names().includes("GetFileToDownload"), "nothing is saved after a failure");
	assert.match(byClass(h.root, "mm-status")[0].textContent, /replacePlain/);
	assert.strictEqual(h.ui.state.running, false, "the wizard is usable again");
});

test("a cancel mid combined run unwinds: restore, no save dialog", async () => {
	const h = boot();
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.scanDocument();

	let runCalls = 0;
	let downloadCalls = 0;
	const outcome = await h.ui.runMergePipeline({
		data: h.ui.state.source.data,
		paragraphs: h.ui.state.paragraphs,
		matched: h.ui.state.mapping.matched,
		unmatched: h.ui.state.mapping.unmatched,
		allTokens: h.ui.state.tokens,
		wrapMatched: true,
		mode: "combined",
		format: "docx",
		rangeText: "",
		exec: function (name, payload) {
			runCalls++;
			h.log.push({ kind: "run", name: name, payload: payload });
			return new Promise(function (resolve) {
				queueMicrotask(function () {
					resolve({ ok: true });
				});
			});
		},
		download: function () {
			downloadCalls++;
			return Promise.resolve({ ok: true, url: "blob://saved" });
		},
		isCancelled: function () {
			return runCalls >= 4;
		},
		onProgress: function () {}
	});

	assert.strictEqual(outcome.ok, false);
	assert.strictEqual(outcome.cancelled, true);
	assert.strictEqual(h.log.filter(function (entry) {
		return entry.name === "mergeRange";
	}).length, 1, "the combined merge happened before the cancel landed");
	assert.ok(h.names().includes("restoreTemplate"), "the template is back");
	assert.strictEqual(downloadCalls, 0, "no save dialog is opened after a cancel");
});

test("a rapid double-click on Merge runs exactly one pipeline", async () => {
	// The latch must be set before the first await (the document scan):
	// otherwise the second click starts its own interleaved run.
	const h = boot();
	h.ui.loadSourceText("recipients.csv", "raw");
	h.ui.state.mode = "per-record";
	const first = h.ui.handleButton(1);
	h.ui.handleButton(1);
	await first;

	assert.strictEqual(h.log.filter(function (entry) {
		return entry.name === "getDocumentText";
	}).length, 1, "the document is scanned once");
	assert.strictEqual(h.log.filter(function (entry) {
		return entry.name === "loadMergeData";
	}).length, 1, "one pipeline, one load");
	assert.strictEqual(h.log.filter(function (entry) {
		return entry.name === "mergeRange";
	}).length, 3, "still one record per recipient");
	assert.strictEqual(h.log.filter(function (entry) {
		return entry.name === "CloseWindow";
	}).length, 1, "the window closes once");
	assert.strictEqual(h.ui.state.running, false, "the latch is released");
});

test("combined mode warns about plain-replaced tokens even with wrapping on", async () => {
	const h = boot();
	h.ui.loadSourceText("recipients.csv", "raw");
	await h.ui.scanDocument();

	const warning = byClass(h.root, "mm-warning")[0];
	assert.notStrictEqual(warning.style.display, "none", "the warning is visible");
	assert.match(warning.textContent, /combined/i);
	assert.match(warning.textContent, /blank/i);

	// Turning off the fallback for the unmatched token removes the only
	// plain-replaced token, and the warning goes with it.
	const fallback = findAll(h.root, function (node) {
		return node.tagName === "INPUT" &&
			node.getAttribute("type") === "checkbox" &&
			node.parentNode &&
			node.parentNode.classList.contains("mm-check") &&
			/plain replace/.test(node.parentNode.textContent);
	})[0];
	assert.ok(fallback, "the unmatched token has a plain-replace checkbox");
	fallback.checked = false;
	fallback.dispatch("change");
	assert.strictEqual(warning.style.display, "none", "no plain tokens, no warning");

	// Per-record mode never blanks, so it never warns.
	fallback.checked = true;
	fallback.dispatch("change");
	assert.notStrictEqual(warning.style.display, "none");
	h.ui.state.mode = "per-record";
	h.ui.gotoStep(2);
	assert.strictEqual(warning.style.display, "none");
});

test("very long names stay in the table cells: truncation is pinned in CSS", () => {
	// The DOM cannot elide a 500-char header; the ellipsis is styles.css's job,
	// so pin the contract instead of pretending the fake DOM lays out text.
	const css = fs.readFileSync(path.join(ROOT, "plugin", "styles.css"), "utf8");
	const rule = /\.mm-table th,\s*\.mm-table td\s*\{([^}]*)\}/.exec(css);
	assert.ok(rule, "the shared th/td rule exists");
	assert.match(rule[1], /max-width:\s*\d/);
	assert.match(rule[1], /overflow:\s*hidden/);
	assert.match(rule[1], /text-overflow:\s*ellipsis/);
	assert.match(rule[1], /white-space:\s*nowrap/);
});

test("a throwing resizeWindow does not block navigation", () => {
	const h = boot();
	h.ui.loadSourceText("recipients.csv", "raw");
	h.win.Asc.plugin.resizeWindow = function () {
		throw new Error("host refused the resize");
	};
	h.ui.gotoStep(2);
	assert.strictEqual(h.ui.state.step, 2, "the step still changes");
});

/* ------------------------------------------------------------------ *
 * Boot sequence - the real load order
 * ------------------------------------------------------------------ */

test("boot: code.js parsed from <head> mounts the wizard once the body exists", () => {
	// index.html loads every script in <head>, so code.js evaluates while
	// #mm-app does not exist yet; the body arrives afterwards and
	// DOMContentLoaded is what actually boots the wizard.
	const h = boot({ readyState: "loading", deferApp: true });
	assert.strictEqual(byClass(h.root, "mm-wrap").length, 0, "nothing mounts at parse time");

	h.setAppReady();
	h.fireDocument("DOMContentLoaded");
	assert.strictEqual(byClass(h.root, "mm-wrap").length, 1, "the wizard mounts on DOMContentLoaded");
	assert.strictEqual(byClass(h.root, "mm-chip").length, 3);
	assert.strictEqual(byClass(h.root, "mm-step").length, 3);
	assert.strictEqual(typeof h.win.Asc.plugin.button, "function", "the button hook is installed");
	assert.strictEqual(typeof h.win.Asc.plugin.init, "function", "the init hook is installed");
});

test("boot: an already-parsed document mounts immediately without any event", () => {
	const h = boot({ readyState: "interactive" });
	assert.strictEqual(byClass(h.root, "mm-wrap").length, 1, "the wizard is mounted right after code.js runs");
});

test("boot: window.Asc arriving after code.js still gets the button/init hooks", async () => {
	const h = boot({ readyState: "loading", deferApp: true, noAsc: true, search: "" });
	h.setAppReady();
	h.fireDocument("DOMContentLoaded");
	assert.strictEqual(byClass(h.root, "mm-wrap").length, 1, "the wizard mounts without the host");
	assert.ok(!h.win.Asc, "the host shim is not there yet");

	const plugin = h.installFakeHost();
	await new Promise(function (resolve) {
		setTimeout(resolve, 150);
	});
	assert.strictEqual(typeof plugin.button, "function", "the button hook landed after the host did");
	assert.strictEqual(typeof plugin.init, "function", "the init hook landed after the host did");

	plugin.button(-1);
	const closes = h.log.filter(function (entry) {
		return entry.name === "CloseWindow";
	});
	assert.strictEqual(closes.length, 1, "the X routes to a CloseWindow call");
	assert.deepStrictEqual(closes[0].args, ["sdk-id"], "with no URL id the SDK-published id closes the window");
});

test("closeWindow: a failing CloseWindow walks the id sources and ends bare", () => {
	const h = boot();
	const calls = [];
	h.win.Asc.plugin.executeMethod = function (name, args) {
		calls.push(args);
		if (args.length && args[0] === "win-42") {
			throw new Error("host refused the URL id");
		}
	};
	const original = console.error;
	console.error = function () {};
	try {
		h.ui.closeWindow();
	} finally {
		console.error = original;
	}
	assert.deepStrictEqual(calls, [["win-42"], ["sdk-id"]], "the URL id first, then the SDK id");

	const bare = boot({ search: "" });
	delete bare.win.Asc.plugin.windowID;
	const bareCalls = [];
	bare.win.Asc.plugin.executeMethod = function (name, args) {
		bareCalls.push(args);
	};
	bare.ui.closeWindow();
	assert.deepStrictEqual(bareCalls, [[]], "no id anywhere: a bare CloseWindow still goes out");
});
