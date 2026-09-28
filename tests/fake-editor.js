/*
 * Minimal ONLYOFFICE editor + plugin-frame simulator for the Mail Merge seam.
 *
 * It models just enough of the real contract to drive plugin/scripts/commands.js
 * end to end outside the editor:
 *   - a document body of paragraphs of RUNS plus a fields registry
 *     (ParaRun.Split2/AddText/GetText/Copy/SetPr, ParaField(FieldType, Arguments)
 *      + Add_ToContent, paragraph Add_ToContent/Remove_FromContent/AddToParagraph -
 *      the exact primitives the wrapFields surgery uses, mirroring
 *      ApiParagraph/ApiRun.WrapInMailMergeField from sdk-all.js)
 *   - Api.LoadMailMergeData(String[][]) -> MailMergeMap (row 0 = headers),
 *     Api.GetMailMergeReceptionsCount, Api.MailMerge(start, end) body
 *     replacement, Api.GetMailMergeTemplateDocContent / Api.ReplaceDocumentContent
 *   - ApiDocument.SearchAndReplace({searchString, replaceString, matchCase})
 *   - a fake plugin frame: commands.js loaded as a plain script into a vm
 *     realm whose globals are `Api`, `ParaRun`, `ParaField`, `AscWord` and
 *     `OnlyOfficeMailMergeFieldMap` - the stringified command bodies resolve
 *     them via the PRELUDE exactly as the editor page does. `window`,
 *     `document` and `alert` are deliberately ABSENT (blocked since
 *     ONLYOFFICE 7.1 inside callCommand).
 *
 * Host contract mirror: Desktop Editors 9.4.0.130-1,
 * editors/sdkjs/word/sdk-all.js (WrapInMailMergeField, MailMerge APIs) and
 * editors/sdkjs-plugins/{pluginBase.js,v1/plugins.js} (callCommand seam).
 */
"use strict";

var fs = require("node:fs");
var path = require("node:path");
var vm = require("node:vm");

var COMMANDS_SOURCE = fs.readFileSync(
	path.join(__dirname, "..", "plugin", "scripts", "commands.js"),
	"utf8"
);

// MIRROR(sdk-all.js AscWord.fieldtype_MERGEFIELD): 0x0001.
var AscWord = { fieldtype_MERGEFIELD: 0x0001 };

// ---------------------------------------------------------------------------
// Internal document model (the Paragraph/ParaRun/ParaField shapes the surgery
// touches). Everything carries the method names the SDK code paths use.
// ---------------------------------------------------------------------------

var prCounter = 0;
// Each formatting object carries an id and remembers its copy's origin, so
// tests can prove the wrapped field inherited the span run's formatting.
function makePr() {
	var pr = {
		id: ++prCounter,
		Copy: function () {
			var copy = makePr();
			copy.copiedFrom = pr.id;
			return copy;
		}
	};
	return pr;
}

function Run(text, pr) {
	this.text = typeof text === "string" ? text : "";
	this.Pr = pr || makePr();
}
Run.prototype.GetText = function () {
	return this.text;
};
Run.prototype.AddText = function (text) {
	this.text += String(text);
	return true;
};
// MIRROR(ParaRun.prototype.Split2, sdk-all.js): the left run keeps [0, pos),
// the returned right run takes the rest, formatting copied.
Run.prototype.Split2 = function (pos) {
	if (typeof pos !== "number" || pos < 0 || pos > this.text.length) {
		return null;
	}
	var right = new Run(this.text.slice(pos), this.Pr.Copy());
	this.text = this.text.slice(0, pos);
	return right;
};
Run.prototype.Copy = function () {
	return new Run(this.text, this.Pr.Copy());
};
Run.prototype.SetPr = function (pr) {
	this.Pr = pr;
	return true;
};

// MIRROR(sdk-all.js function ParaField(FieldType, Arguments, Switches)): the
// merge-field name lives in Arguments[0].
function Field(fieldType, args, switches) {
	this.FieldType = fieldType;
	this.Arguments = args || [];
	this.Switches = switches || [];
	this.Content = [];
}
Field.prototype.Add_ToContent = function (index, item) {
	this.Content.splice(index, 0, item);
	return true;
};
Field.prototype.GetText = function () {
	return this.Content.map(elementText).join("");
};
Field.prototype.Copy = function () {
	var copy = new Field(this.FieldType, this.Arguments.slice(), this.Switches.slice());
	for (var i = 0; i < this.Content.length; i++) {
		copy.Add_ToContent(i, this.Content[i].Copy());
	}
	return copy;
};

function elementText(item) {
	if (!item) {
		return "";
	}
	if (typeof item.GetText === "function") {
		var t = item.GetText();
		if (typeof t === "string") {
			return t;
		}
	}
	if (item.Content && typeof item.Content.length === "number") {
		return item.Content.map(elementText).join("");
	}
	return "";
}

function Para(elements) {
	this.Content = (elements || []).slice();
}
Para.prototype.GetText = function () {
	// The fake has no numbering and no paragraph mark: element texts, joined.
	return this.Content.map(elementText).join("");
};
Para.prototype.Add_ToContent = function (index, item) {
	this.Content.splice(index, 0, item);
	return true;
};
Para.prototype.Remove_FromContent = function (index, count) {
	this.Content.splice(index, count);
	return true;
};
Para.prototype.AddToParagraph = function (item) {
	this.Content.push(item);
	return true;
};
Para.prototype.Copy = function () {
	return new Para(
		this.Content.map(function (item) {
			return item.Copy();
		})
	);
};

function ApiParagraph(para, owner) {
	this.Paragraph = para;
	this.owner = owner;
}
ApiParagraph.prototype.GetText = function () {
	return this.Paragraph.GetText();
};
ApiParagraph.prototype.GetElementsCount = function () {
	return this.Paragraph.Content.length;
};
ApiParagraph.prototype.GetElement = function (pos) {
	return this.Paragraph.Content[pos] || null;
};
ApiParagraph.prototype.RemoveAllElements = function () {
	this.Paragraph.Content = [];
	return true;
};
// MIRROR(ApiParagraph.prototype.WrapInMailMergeField, sdk-all.js): the field is
// named after the paragraph text, quoted runs wrap copies of the content, the
// paragraph is emptied and the field registered + appended.
ApiParagraph.prototype.WrapInMailMergeField = function () {
	var fieldName = this.GetText();
	var oField = new Field(AscWord.fieldtype_MERGEFIELD, [fieldName], []);
	var leftQuote = new Run("\u00AB");
	var rightQuote = new Run("\u00BB");
	oField.Add_ToContent(0, leftQuote);
	for (var nElement = 0; nElement < this.Paragraph.Content.length; nElement++) {
		oField.Add_ToContent(nElement + 1, this.Paragraph.Content[nElement].Copy());
	}
	oField.Add_ToContent(oField.Content.length, rightQuote);
	this.RemoveAllElements();
	this.owner.Register_Field(oField);
	this.Paragraph.AddToParagraph(oField);
	return true;
};

// ---------------------------------------------------------------------------
// The simulated editor.
// ---------------------------------------------------------------------------

function escapeRegExp(text) {
	return String(text).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function replaceAll(text, searchString, replaceString, matchCase) {
	if (matchCase === false) {
		// MIRROR(ApiDocument.SearchAndReplace): the replacement is a plain
		// string, so "$&"-shaped values must stay literal - a replacer function
		// keeps String.replace from reading them as replacement patterns.
		return String(text).replace(new RegExp(escapeRegExp(searchString), "gi"), function () {
			return replaceString;
		});
	}
	return String(text).split(searchString).join(replaceString);
}

/**
 * @param {object} [options]
 *   paragraphs: array of paragraph specs. A string becomes a single run;
 *   an array of element specs builds runs and fields:
 *     { text: "..." }               -> a run
 *     { field: "Name" }             -> a MERGEFIELD displaying «Name»
 *     { field: "Name", text: "..." }-> a MERGEFIELD with custom display content
 *   noGetAllParagraphs: hide ApiDocument.GetAllParagraphs (forces the
 *     CDocument.Content paragraph path in the seam)
 *   failLoad: LoadMailMergeData answers false
 *   omit: names of Api/ApiDocument methods to delete (e.g. ["SearchAndReplace"])
 */
function createEditor(options) {
	options = options || {};

	function buildParagraph(spec) {
		if (typeof spec === "string") {
			return new Para([new Run(spec)]);
		}
		var elements = (spec || []).map(function (elementSpec) {
			if (elementSpec && elementSpec.field !== undefined) {
				var field = new Field(AscWord.fieldtype_MERGEFIELD, [elementSpec.field], []);
				if (typeof elementSpec.text === "string") {
					field.Add_ToContent(0, new Run(elementSpec.text));
				} else {
					field.Add_ToContent(0, new Run("\u00AB"));
					field.Add_ToContent(1, new Run(elementSpec.field));
					field.Add_ToContent(2, new Run("\u00BB"));
				}
				return field;
			}
			return new Run(elementSpec && elementSpec.text !== undefined ? elementSpec.text : "");
		});
		return new Para(elements);
	}

	var body = (options.paragraphs || []).map(buildParagraph);

	var state = {
		fields: [],
		mailMergeData: null
	};

	var logicDocument = {
		Content: body,
		MailMergeMap: null,
		Register_Field: function (field) {
			state.fields.push(field);
			return true;
		},
		Get_MailMergeReceptionsCount: function () {
			return logicDocument.MailMergeMap ? logicDocument.MailMergeMap.length : 0;
		}
	};

	function valueFor(record, name) {
		if (!record) {
			return "";
		}
		if (Object.prototype.hasOwnProperty.call(record, name)) {
			return record[name] === undefined || record[name] === null ? "" : String(record[name]);
		}
		var lower = String(name).toLowerCase();
		for (var key in record) {
			if (Object.prototype.hasOwnProperty.call(record, key) && String(key).toLowerCase() === lower) {
				return record[key] === undefined || record[key] === null ? "" : String(record[key]);
			}
		}
		return "";
	}

	function mergePara(para, record) {
		var copy = para.Copy();
		for (var i = 0; i < copy.Content.length; i++) {
			var item = copy.Content[i];
			if (item instanceof Field) {
				// A merge result carries the record value where the template
				// carried the MERGEFIELD.
				copy.Content[i] = new Run(valueFor(record, item.Arguments[0]));
			}
		}
		return copy;
	}

	var apiDocument = {
		Document: logicDocument,
		SearchAndReplace: function (props) {
			if (!props || typeof props.searchString !== "string" || props.searchString === "") {
				return false;
			}
			var replaceString = typeof props.replaceString === "string" ? props.replaceString : "";
			var matchCase = props.matchCase === undefined ? true : !!props.matchCase;
			var replacedAny = false;
			function walk(item) {
				if (item instanceof Run) {
					var next = replaceAll(item.text, props.searchString, replaceString, matchCase);
					if (next !== item.text) {
						item.text = next;
						replacedAny = true;
					}
					return;
				}
				if (item && item.Content) {
					item.Content.forEach(walk);
				}
			}
			body.forEach(function (para) {
				para.Content.forEach(walk);
			});
			return replacedAny;
		}
	};
	if (!options.noGetAllParagraphs) {
		apiDocument.GetAllParagraphs = function () {
			return body.map(function (para) {
				return new ApiParagraph(para, logicDocument);
			});
		};
	}

	var api = {
		GetDocument: function () {
			return apiDocument;
		},
		LoadMailMergeData: function (data) {
			if (!data || data.length === 0) {
				return false;
			}
			if (options.failLoad) {
				return false;
			}
			state.mailMergeData = data;
			var headers = data[0] || [];
			var rows = [];
			for (var r = 1; r < data.length; r++) {
				var record = {};
				for (var c = 0; c < headers.length; c++) {
					record[headers[c]] = data[r][c];
				}
				rows.push(record);
			}
			logicDocument.MailMergeMap = rows;
			return true;
		},
		GetMailMergeReceptionsCount: function () {
			return logicDocument.Get_MailMergeReceptionsCount();
		},
		GetMailMergeTemplateDocContent: function () {
			// A complex object: it must never cross the callCommand boundary.
			return {
				Document: {
					Content: body.map(function (para) {
						return para.Copy();
					})
				}
			};
		},
		MailMerge: function (startIndex, endIndex) {
			var map = logicDocument.MailMergeMap;
			if (!map || map.length === 0) {
				return false;
			}
			var start = Math.max(0, Math.floor(startIndex));
			var end = Math.min(map.length - 1, Math.floor(endIndex));
			if (start > end) {
				return false;
			}
			var template = body.map(function (para) {
				return para.Copy();
			});
			var merged = [];
			for (var i = start; i <= end; i++) {
				template.forEach(function (para) {
					merged.push(mergePara(para, map[i]));
				});
			}
			body.length = 0;
			Array.prototype.push.apply(body, merged);
			return true;
		},
		ReplaceDocumentContent: function (documentContent) {
			if (!documentContent || !documentContent.Document || !documentContent.Document.Content) {
				return false;
			}
			body.length = 0;
			documentContent.Document.Content.forEach(function (para) {
				body.push(para.Copy());
			});
			return true;
		}
	};

	(options.omit || []).forEach(function (name) {
		if (Object.prototype.hasOwnProperty.call(api, name)) {
			delete api[name];
		}
		if (Object.prototype.hasOwnProperty.call(apiDocument, name)) {
			delete apiDocument[name];
		}
	});

	return {
		api: api,
		apiDocument: apiDocument,
		logicDocument: logicDocument,
		state: state,
		// Paragraph texts, joined - the whole document as the UI sees it.
		text: function () {
			return body
				.map(function (para) {
					return para.GetText();
				})
				.join("\n");
		},
		// JSON-serialisable body structure (run texts, field names/displays) -
		// what the round-trip test compares after restoreTemplate.
		structure: function () {
			return body.map(function (para) {
				return para.Content.map(function (item) {
					if (item instanceof Field) {
						return {
							kind: "field",
							name: item.Arguments[0],
							display: item.GetText()
						};
					}
					return { kind: "run", text: item.text };
				});
			});
		}
	};
}

// ---------------------------------------------------------------------------
// The pure field-map module (plugin/scripts/fieldmap.js), loaded the way the
// tests load it. It belongs to another workstream and may still be a stub, so
// its wrapPlan is probed first and a reference implementation stands in when
// the module cannot answer - the seam must be testable either way. The
// reference mirrors the documented contract: [{token, name, start, end}] with
// `name` the trimmed interior of {{...}}.
// ---------------------------------------------------------------------------

function referenceWrapPlan(text) {
	var out = [];
	if (typeof text !== "string" || !text) {
		return out;
	}
	var re = /\{\{\s*([^{}]+?)\s*\}\}/g;
	var m;
	while ((m = re.exec(text)) !== null) {
		var name = String(m[1]).replace(/^\s+|\s+$/g, "");
		out.push({ token: m[0], name: name, start: m.index, end: m.index + m[0].length });
	}
	return out;
}

function loadFieldMap() {
	var moduleMap = null;
	try {
		// eslint-disable-next-line global-require
		moduleMap = require(path.join(__dirname, "..", "plugin", "scripts", "fieldmap.js"));
	} catch (e) {
		moduleMap = null;
	}
	if (moduleMap && typeof moduleMap.wrapPlan === "function") {
		try {
			var probe = moduleMap.wrapPlan("a{{Name}}b");
			if (
				Array.isArray(probe) &&
				probe.length === 1 &&
				probe[0] &&
				probe[0].name === "Name" &&
				probe[0].start === 1 &&
				probe[0].end === 9
			) {
				return moduleMap;
			}
		} catch (e) {
			// stub or non-conforming - fall through to the reference.
		}
	}
	return { wrapPlan: referenceWrapPlan };
}

// ---------------------------------------------------------------------------
// The fake plugin frame: commands.js as a plain script of a vm realm, the host
// bridge at Asc.plugin.callCommand, payloads on the one shared Asc.scope slot.
// ---------------------------------------------------------------------------

/**
 * @param {object} editor result of `createEditor`
 * @param {object} [options]
 *   commandDelay: answer N ms late (async host)
 *   commandNeverAnswers: never answer (dropped callback)
 *   commandRaw: answer with this raw string instead of the command's JSON
 *   commandOverlaps: hook fired once at dispatch time (before the answer)
 *   omitLowLevelGlobals: drop ParaRun/ParaField/AscWord from the realm -
 *     models a build without low-level run surgery (wrapFields must degrade)
 *   fieldMap: the OnlyOfficeMailMergeFieldMap global; defaults to
 *     `loadFieldMap()`, `null` leaves it absent (prelude fallback scanner)
 *   noApi: do not expose the `Api` global at all
 *   omit: names of Api/ApiDocument methods to remove from the simulated host
 *     (same knob as createEditor's, applied at frame construction)
 */
function createFrame(editor, options) {
	options = options || {};

	(options.omit || []).forEach(function (name) {
		if (Object.prototype.hasOwnProperty.call(editor.api, name)) {
			delete editor.api[name];
		}
		if (Object.prototype.hasOwnProperty.call(editor.apiDocument, name)) {
			delete editor.apiDocument[name];
		}
	});

	var harness = {
		executedCommands: [],
		timers: [],
		errors: []
	};

	var context = {};
	var overlapped = false;

	function setTimer(handler, delay) {
		var record = { delay: delay, cleared: false, fired: false, handle: null };
		harness.timers.push(record);
		record.handle = setTimeout(function () {
			record.fired = true;
			handler();
		}, delay);
		if (record.handle && typeof record.handle.unref === "function") {
			record.handle.unref();
		}
		return record.handle;
	}
	function clearTimer(handle) {
		harness.timers.forEach(function (record) {
			if (record.handle === handle) {
				record.cleared = true;
			}
		});
		clearTimeout(handle);
	}

	var Asc = {
		scope: {},
		plugin: {
			// FAITHFUL(async seam): the real host stringifies the command and
			// evaluates it in the editor page some time later, reading the shared
			// Asc.scope slot at *evaluation* time - which is what lets overlapping
			// runs swap payloads. This double defers the evaluation the same way
			// and records the scope it saw, so a queue regression is visible.
			callCommand: function (commandFn, isClose, isCalc, callback) {
				var record = { isClose: isClose, isCalc: isCalc, scopeAtEval: null };
				harness.executedCommands.push(record);
				if (options.commandNeverAnswers) {
					return;
				}
				if (typeof options.commandOverlaps === "function" && !overlapped) {
					overlapped = true;
					options.commandOverlaps();
				}
				function deliver() {
					if (options.commandRaw !== undefined) {
						if (typeof callback === "function") {
							callback(options.commandRaw);
						}
						return;
					}
					record.scopeAtEval = Asc.scope;
					context.__scope = Asc.scope;
					var wrapper =
						"(function () { var Asc = {}; Asc.scope = __scope; var scope = Asc.scope; return (" +
						commandFn.toString() +
						")(); })()";
					var result;
					try {
						result = vm.runInContext(wrapper, context, { filename: "command.js" });
					} catch (error) {
						harness.errors.push(error);
						result = JSON.stringify({ ok: false, error: "threw: " + error.message });
					}
					if (typeof callback === "function") {
						callback(result);
					}
				}
				if (typeof options.commandDelay === "number") {
					setTimer(deliver, options.commandDelay);
				} else {
					Promise.resolve().then(deliver);
				}
			}
		}
	};

	context.Asc = Asc;
	if (!options.noApi) {
		context.Api = editor.api;
	}
	if (!options.omitLowLevelGlobals) {
		context.ParaRun = Run;
		context.ParaField = Field;
		context.AscWord = AscWord;
	}
	var fieldMap = options.fieldMap === undefined ? loadFieldMap() : options.fieldMap;
	if (fieldMap) {
		context.OnlyOfficeMailMergeFieldMap = fieldMap;
	}
	context.setTimeout = setTimer;
	context.clearTimeout = clearTimer;
	context.self = context;
	// No `window`, `document` or `alert` here on purpose: ONLYOFFICE 7.1+
	// blocks them inside callCommand, and the prelude must survive that.

	vm.createContext(context);
	vm.runInContext(COMMANDS_SOURCE, context, { filename: "commands.js" });

	return {
		// The fake plugin-frame realm (latex-math tests call this `window`).
		window: context,
		Asc: Asc,
		harness: harness,
		runCommand: function (name, payload, timeoutMs) {
			return new Promise(function (resolve) {
				context.OnlyOfficeMailMergeCommands.run(name, payload, resolve, timeoutMs);
			});
		}
	};
}

module.exports = {
	createEditor: createEditor,
	createFrame: createFrame,
	loadFieldMap: loadFieldMap,
	referenceWrapPlan: referenceWrapPlan,
	Run: Run,
	Field: Field,
	Para: Para,
	AscWord: AscWord
};
