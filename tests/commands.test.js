/*
 * Contract tests for the Mail Merge editor-page command seam
 * (plugin/scripts/commands.js) driven through the fake editor/frame
 * (tests/fake-editor.js).
 *
 * Run: node --test tests/commands.test.js
 */
"use strict";

var test = require("node:test");
var assert = require("node:assert/strict");
var fake = require("./fake-editor.js");

var createEditor = fake.createEditor;
var createFrame = fake.createFrame;

// Results come back as objects of the vm realm that hosts commands.js;
// JSON round-tripping normalises them for deepStrictEqual.
function plain(value) {
	return JSON.parse(JSON.stringify(value));
}

function wait(ms) {
	return new Promise(function (resolve) {
		setTimeout(resolve, ms);
	});
}

test("the fake frame realm exposes the seam and blocks window/document/alert", async function () {
	var frame = createFrame(createEditor({ paragraphs: ["x"] }));
	var context = frame.window;
	assert.equal(typeof context.OnlyOfficeMailMergeCommands, "object");
	assert.equal(typeof context.OnlyOfficeMailMergeCommands.run, "function");
	["getDocumentText", "wrapFields", "loadMergeData", "getMergeCount", "snapshotTemplate",
		"mergeRange", "restoreTemplate", "replacePlain"].forEach(function (name) {
		assert.equal(typeof context.OnlyOfficeMailMergeCommands.commands[name], "function", name);
	});
	assert.equal(typeof context.window, "undefined");
	assert.equal(typeof context.document, "undefined");
	assert.equal(typeof context.alert, "undefined");
});

test("getDocumentText returns every paragraph's text", async function () {
	var editor = createEditor({
		paragraphs: ["Dear {{Name}},", [{ text: "To: " }, { field: "City" }], "plain tail"]
	});
	var frame = createFrame(editor);
	var result = plain(await frame.runCommand("getDocumentText", {}));
	assert.deepEqual(result, {
		ok: true,
		paragraphs: ["Dear {{Name}},", "To: \u00ABCity\u00BB", "plain tail"]
	});
});

test("getDocumentText falls back to CDocument.Content and reports api/no-document", async function () {
	var editor = createEditor({
		paragraphs: ["one", "two"],
		noGetAllParagraphs: true
	});
	var result = plain(await createFrame(editor).runCommand("getDocumentText", {}));
	assert.deepEqual(result, { ok: true, paragraphs: ["one", "two"] });

	var noApi = plain(await createFrame(createEditor({}), { noApi: true }).runCommand("getDocumentText", {}));
	assert.deepEqual(noApi, { ok: false, error: "api-unavailable" });

	var noDoc = plain(await createFrame(createEditor({}), { omit: ["GetDocument"] }).runCommand("getDocumentText", {}));
	assert.deepEqual(noDoc, { ok: false, error: "no-document" });
});

test("loadMergeData loads rows and getMergeCount agrees with the receptions count", async function () {
	var frame = createFrame(createEditor({}));
	var loaded = plain(await frame.runCommand("loadMergeData", {
		data: [["Name", "City"], ["Ada", "Paris"], ["Bob", "Rome"]]
	}));
	assert.deepEqual(loaded, { ok: true, count: 2 });
	var count = plain(await frame.runCommand("getMergeCount", {}));
	assert.deepEqual(count, { ok: true, count: 2 });
});

test("loadMergeData error paths: empty, rejected, unavailable", async function () {
	var empty = plain(await createFrame(createEditor({})).runCommand("loadMergeData", {}));
	assert.deepEqual(empty, { ok: false, error: "merge-data-empty" });

	var rejected = plain(await createFrame(createEditor({ failLoad: true })).runCommand("loadMergeData", {
		data: [["Name"], ["Ada"]]
	}));
	assert.deepEqual(rejected, { ok: false, error: "merge-data-rejected" });

	var unavailable = plain(await createFrame(createEditor({}), { omit: ["LoadMailMergeData"] })
		.runCommand("loadMergeData", { data: [["Name"], ["Ada"]] }));
	assert.deepEqual(unavailable, { ok: false, error: "load-unavailable" });
});

test("getMergeCount reports count-unavailable and api-unavailable", async function () {
	var unavailable = plain(await createFrame(createEditor({}), { omit: ["GetMailMergeReceptionsCount"] })
		.runCommand("getMergeCount", {}));
	assert.deepEqual(unavailable, { ok: false, error: "count-unavailable" });

	var noApi = plain(await createFrame(createEditor({}), { noApi: true }).runCommand("getMergeCount", {}));
	assert.deepEqual(noApi, { ok: false, error: "api-unavailable" });
});

test("wrapFields wraps an inline token and preserves the plain runs around it", async function () {
	var editor = createEditor({
		paragraphs: ["Dear {{Name}},", [{ text: "Hi " }, { text: "{{Name}}" }, { text: "!" }]]
	});
	// The explicit-runs paragraph's token run, to prove formatting inheritance.
	var tokenRun = editor.logicDocument.Content[1].Content[1];
	var tokenPrId = tokenRun.Pr.id;

	var frame = createFrame(editor);
	var result = plain(await frame.runCommand("wrapFields", {}));
	assert.deepEqual(result, {
		ok: true,
		wrapped: [{ name: "Name", count: 2 }],
		errors: [],
		degraded: false
	});

	// The string paragraph went through the split-at-boundaries path: only the
	// token became a field, the runs before and after are untouched plain text.
	assert.deepEqual(editor.structure(), [
		[
			{ kind: "run", text: "Dear " },
			{ kind: "field", name: "Name", display: "\u00ABName\u00BB" },
			{ kind: "run", text: "," }
		],
		[
			{ kind: "run", text: "Hi " },
			{ kind: "field", name: "Name", display: "\u00ABName\u00BB" },
			{ kind: "run", text: "!" }
		]
	]);

	// The field carries the MERGEFIELD name and is registered with the document.
	assert.equal(editor.state.fields.length, 2);
	assert.deepEqual(editor.state.fields.map(function (f) { return f.Arguments[0]; }), ["Name", "Name"]);

	// The name run inherited the token run's formatting (a copy of its Pr).
	var nameRun = editor.logicDocument.Content[1].Content[1].Content[1];
	assert.equal(nameRun.Pr.copiedFrom, tokenPrId);

	// getDocumentText now shows the rendered display.
	var text = plain(await frame.runCommand("getDocumentText", {}));
	assert.deepEqual(text, {
		ok: true,
		paragraphs: ["Dear \u00ABName\u00BB,", "Hi \u00ABName\u00BB!"]
	});
});

test("wrapFields handles several tokens and repeated names", async function () {
	var editor = createEditor({
		paragraphs: ["{{A}} and {{B}}", "{{A}}"]
	});
	var result = plain(await createFrame(editor).runCommand("wrapFields", {}));
	assert.deepEqual(result, {
		ok: true,
		wrapped: [{ name: "A", count: 2 }, { name: "B", count: 1 }],
		errors: [],
		degraded: false
	});
	assert.deepEqual(editor.structure(), [
		[
			{ kind: "field", name: "A", display: "\u00ABA\u00BB" },
			{ kind: "run", text: " and " },
			{ kind: "field", name: "B", display: "\u00ABB\u00BB" }
		],
		[{ kind: "field", name: "A", display: "\u00ABA\u00BB" }]
	]);
});

test("wrapFields honors the tokens filter case-insensitively", async function () {
	var editor = createEditor({ paragraphs: ["{{Name}} {{City}}"] });
	var result = plain(await createFrame(editor).runCommand("wrapFields", {
		tokens: [{ name: "name" }]
	}));
	assert.deepEqual(result, {
		ok: true,
		wrapped: [{ name: "Name", count: 1 }],
		errors: [],
		degraded: false
	});
	assert.deepEqual(editor.structure(), [[
		{ kind: "field", name: "Name", display: "\u00ABName\u00BB" },
		{ kind: "run", text: " {{City}}" }
	]]);
});

test("wrapFields reports run-split-unavailable when a token boundary lands inside a field", async function () {
	// Paragraph text "{{a}}by": the token [0,5) ends inside the field's display.
	var editor = createEditor({
		paragraphs: [[{ text: "{{" }, { field: "X", text: "a}}b" }, { text: "y" }]]
	});
	var result = plain(await createFrame(editor).runCommand("wrapFields", {}));
	assert.equal(result.ok, true);
	assert.deepEqual(result.wrapped, []);
	assert.equal(result.degraded, false);
	assert.deepEqual(result.errors, [{
		name: "a", start: 0, end: 5, error: "run-split-unavailable", paragraph: 0
	}]);
});

test("wrapFields reports unsupported-element-in-span when a field sits inside the token", async function () {
	// Paragraph text "{{Name}}": run "{{Na" + field "me}}" cover the whole token.
	var editor = createEditor({
		paragraphs: [[{ text: "{{Na" }, { field: "X", text: "me}}" }]]
	});
	var result = plain(await createFrame(editor).runCommand("wrapFields", {}));
	assert.equal(result.ok, true);
	assert.deepEqual(result.wrapped, []);
	assert.deepEqual(result.errors, [{
		name: "Name", start: 0, end: 8, error: "unsupported-element-in-span", paragraph: 0
	}]);
});

test("wrapFields degrades to whole-paragraph wrapping when run surgery is unavailable", async function () {
	var editor = createEditor({
		paragraphs: ["{{Name}}", "Dear {{City}},"]
	});
	var frame = createFrame(editor, { omitLowLevelGlobals: true });
	var result = plain(await frame.runCommand("wrapFields", {}));
	assert.deepEqual(result, {
		ok: true,
		wrapped: [{ name: "Name", count: 1 }],
		errors: [{ name: "City", paragraph: 1, error: "inline-wrap-unavailable" }],
		degraded: true
	});
	// Whole-token paragraph went through ApiParagraph.WrapInMailMergeField
	// (field named after the paragraph text) and the name was repaired to Name.
	assert.deepEqual(editor.structure(), [
		[{ kind: "field", name: "Name", display: "\u00AB{{Name}}\u00BB" }],
		// The inline paragraph is left untouched: no surgery, no splitting.
		[{ kind: "run", text: "Dear {{City}}," }]
	]);
	assert.deepEqual(editor.state.fields.map(function (f) { return f.Arguments[0]; }), ["Name"]);
});

test("wrapFields serves token offsets from the prelude fallback when no field map is present", async function () {
	var editor = createEditor({ paragraphs: ["Dear {{Name}},"] });
	var result = plain(await createFrame(editor, { fieldMap: null }).runCommand("wrapFields", {}));
	assert.deepEqual(result, {
		ok: true,
		wrapped: [{ name: "Name", count: 1 }],
		errors: [],
		degraded: false
	});
	assert.deepEqual(editor.structure(), [[
		{ kind: "run", text: "Dear " },
		{ kind: "field", name: "Name", display: "\u00ABName\u00BB" },
		{ kind: "run", text: "," }
	]]);
});

test("wrapFields reports api-unavailable and no-document", async function () {
	var noApi = plain(await createFrame(createEditor({ paragraphs: ["{{N}}"] }), { noApi: true })
		.runCommand("wrapFields", {}));
	assert.deepEqual(noApi, { ok: false, error: "api-unavailable" });

	var noDoc = plain(await createFrame(createEditor({ paragraphs: ["{{N}}"] }), { omit: ["GetDocument"] })
		.runCommand("wrapFields", {}));
	assert.deepEqual(noDoc, { ok: false, error: "no-document" });
});

test("snapshot -> mergeRange -> restoreTemplate round trip restores the body exactly", async function () {
	var editor = createEditor({ paragraphs: ["Dear {{Name}},"] });
	var frame = createFrame(editor);
	assert.deepEqual(plain(await frame.runCommand("wrapFields", {})), {
		ok: true, wrapped: [{ name: "Name", count: 1 }], errors: [], degraded: false
	});
	assert.deepEqual(plain(await frame.runCommand("loadMergeData", {
		data: [["Name"], ["Ada"], ["Bob"]]
	})), { ok: true, count: 2 });

	var before = editor.structure();
	assert.deepEqual(plain(await frame.runCommand("snapshotTemplate", {})), { ok: true });
	// The complex snapshot never crosses the boundary: it stays on the Api object.
	assert.ok(editor.api.__mmSnapshot);
	assert.deepEqual(editor.api.__mmSnapshot.Document.Content.length, 1);

	assert.deepEqual(plain(await frame.runCommand("mergeRange", { start: 0, end: 1 })), { ok: true });
	// MailMerge replaced the body with one merged paragraph per record; the
	// MERGEFIELD element became a run holding the record value.
	assert.deepEqual(editor.structure(), [
		[
			{ kind: "run", text: "Dear " },
			{ kind: "run", text: "Ada" },
			{ kind: "run", text: "," }
		],
		[
			{ kind: "run", text: "Dear " },
			{ kind: "run", text: "Bob" },
			{ kind: "run", text: "," }
		]
	]);
	assert.equal(editor.text(), "Dear Ada,\nDear Bob,");

	assert.deepEqual(plain(await frame.runCommand("restoreTemplate", {})), { ok: true });
	assert.deepEqual(editor.structure(), before);
	assert.deepEqual(editor.text(), "Dear \u00ABName\u00BB,");
});

test("restoreTemplate without a snapshot reports no-snapshot", async function () {
	var frame = createFrame(createEditor({ paragraphs: ["x"] }));
	assert.deepEqual(plain(await frame.runCommand("restoreTemplate", {})), { ok: false, error: "no-snapshot" });
});

test("mergeRange error paths: invalid range, refused merge, missing API", async function () {
	var invalid = plain(await createFrame(createEditor({})).runCommand("mergeRange", { start: "0", end: 1 }));
	assert.deepEqual(invalid, { ok: false, error: "merge-range-invalid" });

	// No merge data loaded: the editor refuses and the seam names it merge-failed.
	var refused = plain(await createFrame(createEditor({ paragraphs: ["x"] })).runCommand("mergeRange", { start: 0, end: 0 }));
	assert.deepEqual(refused, { ok: false, error: "merge-failed" });

	var unavailable = plain(await createFrame(createEditor({}), { omit: ["MailMerge"] })
		.runCommand("mergeRange", { start: 0, end: 0 }));
	assert.deepEqual(unavailable, { ok: false, error: "merge-unavailable" });

	var noApi = plain(await createFrame(createEditor({}), { noApi: true }).runCommand("mergeRange", { start: 0, end: 0 }));
	assert.deepEqual(noApi, { ok: false, error: "api-unavailable" });
});

test("replacePlain replaces sequentially and counts the replacements", async function () {
	// No snapshot involved at all: replacePlain never reads __mmSnapshot.
	var editor = createEditor({ paragraphs: ["Hi {{Name}}, welcome to {{City}}."], fields: [] });
	var frame = createFrame(editor);
	var result = plain(await frame.runCommand("replacePlain", {
		plan: [
			{ searchString: "{{Name}}", replaceString: "Ada" },
			{ searchString: "{{City}}", replaceString: "Paris" },
			{ searchString: "{{Nope}}", replaceString: "x" }
		]
	}));
	assert.deepEqual(result, { ok: true, replaced: 2 });
	assert.equal(editor.text(), "Hi Ada, welcome to Paris.");
	assert.equal(editor.api.__mmSnapshot, undefined);

	// matchCase defaults to true; false matches any case.
	var cased = createEditor({ paragraphs: ["Colour color"] });
	var casedResult = plain(await createFrame(cased).runCommand("replacePlain", {
		plan: [
			{ searchString: "color", replaceString: "X" },
			{ searchString: "Colour", replaceString: "Y", matchCase: false }
		]
	}));
	assert.deepEqual(casedResult, { ok: true, replaced: 2 });
	assert.equal(cased.text(), "Y X");
});

test("replacePlain error paths: no plan, missing API, no Api", async function () {
	var noPlan = plain(await createFrame(createEditor({ paragraphs: ["x"] })).runCommand("replacePlain", {}));
	assert.deepEqual(noPlan, { ok: false, error: "no-plan" });

	var unavailable = plain(await createFrame(createEditor({ paragraphs: ["x"] }), { omit: ["SearchAndReplace"] })
		.runCommand("replacePlain", { plan: [{ searchString: "x", replaceString: "y" }] }));
	assert.deepEqual(unavailable, { ok: false, error: "replace-unavailable" });

	var noApi = plain(await createFrame(createEditor({ paragraphs: ["x"] }), { noApi: true })
		.runCommand("replacePlain", { plan: [{ searchString: "x", replaceString: "y" }] }));
	assert.deepEqual(noApi, { ok: false, error: "api-unavailable" });
});

test("the dispatch queue serializes overlapping runs over the one Asc.scope slot", async function () {
	var frame = createFrame(createEditor({}), { commandDelay: 20 });
	var first = frame.runCommand("loadMergeData", { data: [["Name"], ["Ada"]] }, 5000);
	var second = frame.runCommand("loadMergeData", { data: [["Name"], ["Ada"], ["Bob"]] }, 5000);

	// One command in flight; the second run queues instead of clobbering.
	assert.equal(frame.harness.executedCommands.length, 1);

	var results = await Promise.all([first, second]);
	assert.deepEqual(plain(results[0]), { ok: true, count: 1 });
	assert.deepEqual(plain(results[1]), { ok: true, count: 2 });

	// Each command evaluated against the payload its own run parked in the slot.
	assert.equal(frame.harness.executedCommands[0].scopeAtEval.data.length, 2);
	assert.equal(frame.harness.executedCommands[1].scopeAtEval.data.length, 3);
});

test("a hung command times out and its late answer is discarded", async function () {
	// commandDelay 60 > timeoutMs 25: the backstop wins.
	var frame = createFrame(createEditor({}), { commandDelay: 60 });
	var hung = plain(await frame.runCommand("getMergeCount", {}, 25));
	assert.deepEqual(hung, { ok: false, error: "timeout" });

	// The late answer arrives after the entry settled; it must neither revive
	// the timed-out run nor leak into whatever owns the slot next.
	await wait(80);
	var next = plain(await frame.runCommand("getMergeCount", {}, 200));
	assert.deepEqual(next, { ok: true, count: 0 });
});

test("a command that never answers times out through the same taxonomy", async function () {
	var frame = createFrame(createEditor({}), { commandNeverAnswers: true });
	var result = plain(await frame.runCommand("getMergeCount", {}, 25));
	assert.deepEqual(result, { ok: false, error: "timeout" });
});

test("an unparsable command answer reports unparsable-command-result with the raw text", async function () {
	var frame = createFrame(createEditor({}), { commandRaw: "not-json" });
	var result = plain(await frame.runCommand("getMergeCount", {}, 5000));
	assert.deepEqual(result, { ok: false, error: "unparsable-command-result", raw: "not-json" });
});

test("a missing callCommand bridge reports callCommand-unavailable", async function () {
	var frame = createFrame(createEditor({}));
	frame.window.Asc.plugin.callCommand = undefined;
	var result = plain(await frame.runCommand("getMergeCount", {}, 5000));
	assert.deepEqual(result, { ok: false, error: "callCommand-unavailable" });
});

test("a throwing callCommand bridge reports callCommand-threw", async function () {
	var frame = createFrame(createEditor({}));
	frame.window.Asc.plugin.callCommand = function () {
		throw new Error("bridge boom");
	};
	var result = plain(await frame.runCommand("getMergeCount", {}, 5000));
	assert.deepEqual(result, { ok: false, error: "callCommand-threw: bridge boom" });
});

test("run rejects unknown command names and missing callbacks", async function () {
	var frame = createFrame(createEditor({}));
	assert.throws(function () {
		frame.window.OnlyOfficeMailMergeCommands.run("nope", {}, function () {});
	}, /unknown editor command/);
	assert.throws(function () {
		frame.window.OnlyOfficeMailMergeCommands.run("getMergeCount", {});
	}, /needs a callback/);
});

test("a synchronous host answer clears the armed backstop instead of orphaning it", async function () {
	var frame = createFrame(createEditor({}));
	frame.window.Asc.plugin.callCommand = function (commandFn, isClose, isCalc, callback) {
		callback('{"ok":true,"count":7}');
	};
	var result = plain(await frame.runCommand("getMergeCount", {}, 30));
	assert.deepEqual(result, { ok: true, count: 7 });

	var timers = frame.harness.timers;
	assert.equal(timers.length, 1);
	assert.equal(timers[0].delay, 30);
	assert.equal(timers[0].cleared, true);
	assert.equal(timers[0].fired, false);
	await wait(50);
	assert.equal(timers[0].fired, false);
});
