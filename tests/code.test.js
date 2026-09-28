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

	const root = new El("div");
	root.setAttribute("id", "mm-app");

	const document = {
		createElement: function (tag) {
			return new El(tag);
		},
		createTextNode: function (text) {
			return { nodeType: 3, data: String(text), textContent: String(text) };
		},
		getElementById: function (id) {
			return id === "mm-app" ? root : null;
		},
		documentElement: new El("html"),
		body: new El("body")
	};

	const win = {
		location: { search: options.search !== undefined ? options.search : "?windowID=win-42" },
		localStorage: storage,
		Papa: undefined,
		XLSX: undefined,
		OnlyOfficeMailMergeDataParse: DataParse,
		OnlyOfficeMailMergeFieldMap: FieldMap,
		OnlyOfficeMailMergeCommands: Commands,
		Asc: {
			plugin: {
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
			}
		}
	};

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
