"use strict";

/*
 * dataparse.js owns the one canonical shape Api.LoadMailMergeData consumes:
 * row 0 = merge-field names, later rows = recipient values, every cell a
 * string. These tests pin the normalization rules (BOM, header validity,
 * ragged rows, dropped blank rows) and every documented failure mode.
 */

const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");

const dataparse = require(path.join(__dirname, "..", "plugin", "scripts", "dataparse.js"));
const Papa = require(path.join(__dirname, "..", "plugin", "vendor", "papaparse.min.js"));

function allStrings(data) {
	return data.every((row) => row.every((cell) => typeof cell === "string"));
}

test("parseCsv keeps quoted commas, quotes and newlines inside cells", () => {
	const text = "\uFEFFName,Notes\r\nAnn,\"a, b,\"\"c\"\"\nsecond line\"\r\n";
	const result = dataparse.parseCsv(text, Papa);
	assert.equal(result.ok, true);
	assert.deepEqual(result.data, [
		["Name", "Notes"],
		["Ann", "a, b,\"c\"\nsecond line"]
	]);
	// The BOM must not leak into the first header, CRLF must not leak into cells.
	assert.ok(allStrings(result.data));
	assert.deepEqual(result.warnings, []);
});

test("parseCsv trims header cells but keeps data cells as-is", () => {
	const result = dataparse.parseCsv(" Name  , Age \nBob,30", Papa);
	assert.deepEqual(result.data, [["Name", "Age"], ["Bob", "30"]]);
});

test("parseCsv pads ragged rows and drops cells past the header width", () => {
	const result = dataparse.parseCsv("a,b\n1\n1,2,3\n", Papa);
	assert.deepEqual(result.data, [["a", "b"], ["1", ""], ["1", "2"]]);
});

test("parseCsv drops fully empty rows anywhere in the body", () => {
	const result = dataparse.parseCsv("a,b\n\n1,2\n,,\n3,4\n", Papa);
	assert.deepEqual(result.data, [["a", "b"], ["1", "2"], ["3", "4"]]);
});

test("parseCsv rejects duplicate headers", () => {
	const result = dataparse.parseCsv("a,b,a\n1,2,3", Papa);
	assert.equal(result.ok, false);
	assert.match(result.error, /duplicate header "a"/);
});

test("parseCsv rejects empty headers with the offending column", () => {
	const result = dataparse.parseCsv("a,,c\n1,2,3", Papa);
	assert.equal(result.ok, false);
	assert.match(result.error, /empty header at column 2/);
});

test("parseCsv rejects a source with zero data rows", () => {
	for (const text of ["a,b", "a,b\n", "a,b\n,,\n"]) {
		const result = dataparse.parseCsv(text, Papa);
		assert.equal(result.ok, false, JSON.stringify(text));
		assert.equal(result.error, "no data rows");
	}
});

test("parseCsv without a Papa parser fails with papa-unavailable", () => {
	// No injection and no root.Papa in the node module scope.
	assert.deepEqual(dataparse.parseCsv("a,b\n1,2"), {
		ok: false,
		error: "papa-unavailable"
	});
});

test("parseCsv accepts an injected parse function and a Papa-shaped object", () => {
	const asFunction = dataparse.parseCsv("a,b\n1,2", (input) => Papa.parse(input));
	assert.deepEqual(asFunction.data, [["a", "b"], ["1", "2"]]);
	const asObject = dataparse.parseCsv("a,b\n1,2", { parse: (input) => Papa.parse(input) });
	assert.deepEqual(asObject.data, [["a", "b"], ["1", "2"]]);
});

test("parseJson takes the union of object keys in first-appearance order", () => {
	const result = dataparse.parseJson('[{"a":1,"b":2},{"b":3,"c":4}]');
	assert.equal(result.ok, true);
	assert.deepEqual(result.data, [
		["a", "b", "c"],
		["1", "2", ""],
		["", "3", "4"]
	]);
});

test("parseJson pads ragged arrays of arrays and truncates past the headers", () => {
	const result = dataparse.parseJson('[["a","b"],["1"],["2","3","4"]]');
	assert.deepEqual(result.data, [["a", "b"], ["1", ""], ["2", "3"]]);
});

test("parseJson strips a BOM and rejects invalid JSON descriptively", () => {
	const bad = dataparse.parseJson("\uFEFF{not json");
	assert.equal(bad.ok, false);
	assert.match(bad.error, /invalid JSON/);
});

test("parseJson rejects non-array and empty documents", () => {
	assert.match(dataparse.parseJson('{"a":1}').error, /expected an array/);
	assert.equal(dataparse.parseJson("[]").error, "no data rows");
});

test("parseJson rejects mixed row shapes", () => {
	assert.match(dataparse.parseJson('[{"a":1},[2]]').error, /row 2: expected an object/);
	assert.match(dataparse.parseJson('[["a"],{"b":2}]').error, /expected an array of cells/);
});

test("parseJson applies the header rules to object keys", () => {
	assert.match(dataparse.parseJson('[{"a":1," a ":2}]').error, /duplicate header "a"/);
	assert.match(dataparse.parseJson('[{"":1}]').error, /empty header at column 1/);
});

test("parseGrid stringifies cells: Dates as YYYY-MM-DD, numbers plain, blanks as empty", () => {
	const result = dataparse.parseGrid([
		["when", "n", "blank"],
		[new Date(Date.UTC(2024, 2, 5)), 1.5, null],
		[new Date("2024-12-31T23:59:59Z"), 0, undefined]
	]);
	assert.deepEqual(result.data, [
		["when", "n", "blank"],
		["2024-03-05", "1.5", ""],
		["2024-12-31", "0", ""]
	]);
	assert.ok(allStrings(result.data));
});

test("parseGrid pads ragged rows and drops fully empty rows", () => {
	const result = dataparse.parseGrid([
		["a", "b"],
		[1],
		[null, "", undefined],
		[2, 3]
	]);
	assert.deepEqual(result.data, [["a", "b"], ["1", ""], ["2", "3"]]);
});

test("parseGrid honours opts.width as the forced column count", () => {
	const wide = dataparse.parseGrid([["a", "b", "c"], [1, 2, 3]], { width: 2 });
	assert.deepEqual(wide.data, [["a", "b"], ["1", "2"]]);
	// Width past the header row would need header names that do not exist.
	const tooWide = dataparse.parseGrid([["a", "b"], [1, 2]], { width: 3 });
	assert.equal(tooWide.ok, false);
	assert.match(tooWide.error, /empty header at column 3/);
});

test("parseGrid rejects malformed grids and header-only grids", () => {
	assert.match(dataparse.parseGrid("nope").error, /invalid grid/);
	assert.match(dataparse.parseGrid([[1], "nope"]).error, /invalid grid row 2/);
	assert.equal(dataparse.parseGrid([]).error, "no data rows");
	assert.equal(dataparse.parseGrid([["a"], []]).error, "no data rows");
	assert.equal(dataparse.parseGrid([[""], [1]]).error, "empty header at column 1");
});

test("toLoadMailMergeData returns ok data unchanged and throws on failures", () => {
	const result = dataparse.parseGrid([["a"], ["1"]]);
	assert.equal(dataparse.toLoadMailMergeData(result), result.data);
	assert.throws(() => dataparse.toLoadMailMergeData({ ok: false, error: "no data rows" }), {
		message: "no data rows"
	});
	assert.throws(() => dataparse.toLoadMailMergeData(undefined), /invalid parse result/);
});
