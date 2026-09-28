"use strict";

/*
 * fieldmap.js owns the {{...}} placeholder vocabulary. A "token" in this API is
 * the raw "{{...}}" text exactly as it appears in the paragraph (that is what
 * wrapPlan labels `token`, with `name` alongside it for the field name inside).
 * These tests pin raw-form identity, dedupe order, offsets, and the
 * case-insensitive name matching against data-source headers.
 */

const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");

const fieldmap = require(path.join(__dirname, "..", "plugin", "scripts", "fieldmap.js"));

test("findTokens reports raw forms deduped in first-appearance order", () => {
	assert.deepEqual(fieldmap.findTokens("{{b}} x {{a}} y {{b}}"), ["{{b}}", "{{a}}"]);
});

test("findTokens keeps spaces, dots and dashes inside token names", () => {
	assert.deepEqual(
		fieldmap.findTokens("Dear {{ first name }}, ref {{addr.line-1}} / {{tax-id}}"),
		["{{ first name }}", "{{addr.line-1}}", "{{tax-id}}"]
	);
	assert.equal(fieldmap.wrapPlan("{{ first name }}")[0].name, "first name");
});

test("findTokens treats distinct raw forms as distinct tokens", () => {
	// Each form needs its own searchString later, so "{{ name }}" and
	// "{{name}}" cannot collapse into one entry.
	assert.deepEqual(
		fieldmap.findTokens("{{ name }} and {{name}} and {{ name }}"),
		["{{ name }}", "{{name}}"]
	);
});

test("findTokens finds adjacent tokens and ignores text without tokens", () => {
	assert.deepEqual(fieldmap.findTokens("{{a}}{{b}}"), ["{{a}}", "{{b}}"]);
	assert.deepEqual(fieldmap.findTokens("plain text and {{}} stray"), []);
	assert.deepEqual(fieldmap.findTokens("{{ }}"), [], "braces with no field name are not tokens");
});

test("findTokensInTexts dedupes across paragraphs in order", () => {
	assert.deepEqual(
		fieldmap.findTokensInTexts(["{{a}} {{b}}", "{{b}} {{c}}", "{{a}}"]),
		["{{a}}", "{{b}}", "{{c}}"]
	);
	assert.deepEqual(fieldmap.findTokensInTexts("not an array"), []);
});

test("normalizeName trims and collapses inner spaces", () => {
	assert.equal(fieldmap.normalizeName("  first \t name  "), "first name");
	assert.equal(fieldmap.normalizeName(null), "");
	assert.equal(fieldmap.normalizeName(undefined), "");
});

test("matchFields matches names case-insensitively and reports gaps", () => {
	const result = fieldmap.matchFields(
		["{{name}}", "{{ first  name }}", "{{missing}}"],
		["Name", "First Name", "Age"]
	);
	assert.deepEqual(result, {
		matched: [
			{ token: "{{name}}", header: "Name" },
			{ token: "{{ first  name }}", header: "First Name" }
		],
		unmatched: ["{{missing}}"],
		unusedHeaders: ["Age"]
	});
});

test("matchFields tolerates bare names and matches nothing without headers", () => {
	assert.deepEqual(fieldmap.matchFields(["name"], ["Name"]).matched, [
		{ token: "name", header: "Name" }
	]);
	assert.deepEqual(fieldmap.matchFields(["{{a}}"], []), {
		matched: [],
		unmatched: ["{{a}}"],
		unusedHeaders: []
	});
});

test("plainReplacePlan builds one entry per raw form with exact search strings", () => {
	const plan = fieldmap.plainReplacePlan(
		"{{ name }} and {{name}} and {{ name }} again",
		{ Name: "Ann" }
	);
	assert.deepEqual(plan, [
		{ searchString: "{{ name }}", replaceString: "Ann" },
		{ searchString: "{{name}}", replaceString: "Ann" }
	]);
});

test("plainReplacePlan fills missing values with empty strings", () => {
	assert.deepEqual(fieldmap.plainReplacePlan("Hi {{a}}", {}), [
		{ searchString: "{{a}}", replaceString: "" }
	]);
	assert.deepEqual(fieldmap.plainReplacePlan("Hi {{a}}", null), [
		{ searchString: "{{a}}", replaceString: "" }
	]);
});

test("plainReplacePlan restricts to the given tokens, by raw form or by name", () => {
	const text = "{{a}} {{b}} {{ a }} {{ b }}";
	const record = { a: "1", b: "2" };
	// A raw form restricts to that exact form; a bare name restricts to every
	// form of the field.
	assert.deepEqual(fieldmap.plainReplacePlan(text, record, ["{{a}}"]), [
		{ searchString: "{{a}}", replaceString: "1" }
	]);
	assert.deepEqual(fieldmap.plainReplacePlan(text, record, ["b"]), [
		{ searchString: "{{b}}", replaceString: "2" },
		{ searchString: "{{ b }}", replaceString: "2" }
	]);
	assert.deepEqual(fieldmap.plainReplacePlan(text, record, []), []);
});

test("wrapPlan gives exclusive character offsets in order of occurrence", () => {
	const text = "Hi {{ first name }}, {{last}}!";
	assert.deepEqual(fieldmap.wrapPlan(text), [
		{ token: "{{ first name }}", name: "first name", start: 3, end: 19 },
		{ token: "{{last}}", name: "last", start: 21, end: 29 }
	]);
});

test("wrapPlan lists every occurrence of a repeated token", () => {
	assert.deepEqual(fieldmap.wrapPlan("{{a}} {{a}}{{b}}"), [
		{ token: "{{a}}", name: "a", start: 0, end: 5 },
		{ token: "{{a}}", name: "a", start: 6, end: 11 },
		{ token: "{{b}}", name: "b", start: 11, end: 16 }
	]);
	assert.deepEqual(fieldmap.wrapPlan("no tokens here"), []);
});

test("fieldDisplay renders a merge-field name in guillemets", () => {
	assert.equal(fieldmap.fieldDisplay("Name"), "«Name»");
	assert.equal(fieldmap.fieldDisplay(""), "«»");
});
