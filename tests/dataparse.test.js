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
	// Leading blank rows are skipped (see below), so the empty-header rule is
	// pinned with a header row that is unambiguously a header row.
	assert.equal(dataparse.parseGrid([["", "x"], [1, 2]]).error, "empty header at column 1");
});

test("toLoadMailMergeData returns ok data unchanged and throws on failures", () => {
	const result = dataparse.parseGrid([["a"], ["1"]]);
	assert.equal(dataparse.toLoadMailMergeData(result), result.data);
	assert.throws(() => dataparse.toLoadMailMergeData({ ok: false, error: "no data rows" }), {
		message: "no data rows"
	});
	assert.throws(() => dataparse.toLoadMailMergeData(undefined), /invalid parse result/);
});

/* ------------------------------------------------------------------ *
 * Edge cases beyond the happy path
 * ------------------------------------------------------------------ */

test("duplicate headers are judged case-insensitively and by normalized form", () => {
	// Token matching is case-insensitive, so "Name" and "name" would both back
	// {{name}} and the winner would depend on the consumer.
	const cased = dataparse.parseCsv("Name,name\nAda,Bob", Papa);
	assert.equal(cased.ok, false);
	assert.match(cased.error, /duplicate header "name"/);
	const normalized = dataparse.parseCsv("First  Name,first name\nA,B", Papa);
	assert.equal(normalized.ok, false);
	assert.match(normalized.error, /duplicate header "first name"/);
	const grid = dataparse.parseGrid([["Name", "NAME"], ["Ada", "Bob"]]);
	assert.equal(grid.ok, false);
	assert.match(grid.error, /duplicate header "NAME"/);
	// Distinct non-colliding headers still load.
	assert.equal(dataparse.parseCsv("Name,Name2\nAda,Bob", Papa).ok, true);
});

test("parseCsv loads a file whose blank leading rows sit before the header row", () => {
	const text = "\uFEFF\r\n\r\nName,City\r\nAda,Paris\r\n";
	const result = dataparse.parseCsv(text, Papa);
	assert.deepEqual(result.data, [["Name", "City"], ["Ada", "Paris"]]);
	// The same shape as an all-blank grid: not "empty header", but no rows.
	assert.equal(dataparse.parseCsv("\n,,\n", Papa).error, "no data rows");
});

test("parseCsv keeps CRLF inside quoted cells alongside the BOM", () => {
	const text = "\uFEFFName,Notes\r\nAnn,\"line1\r\nline2\"\r\n";
	const result = dataparse.parseCsv(text, Papa);
	assert.equal(result.ok, true);
	assert.deepEqual(result.data, [["Name", "Notes"], ["Ann", "line1\r\nline2"]]);
});

test("parseGrid skips blank leading rows before the header row", () => {
	const result = dataparse.parseGrid([[], [""], [null, null], ["Name", "City"], ["Ada", "Paris"]]);
	assert.equal(result.ok, true);
	assert.deepEqual(result.data, [["Name", "City"], ["Ada", "Paris"]]);
	// A blank row before a lone header is dropped like any other blank row:
	// the grid is header-only and has no data rows.
	assert.equal(dataparse.parseGrid([[""], [1]]).error, "no data rows");
});

test("parseGrid accepts an XLSX grid whose used range starts with blank rows", async () => {
	// Round-trip through the vendored reader exactly like code.js parseWorkbook,
	// so the test pins the real sheet_to_json shape of a sheet's blank rows.
	const XLSX = require(path.join(__dirname, "..", "plugin", "vendor", "xlsx.full.min.js"));
	const sheet = XLSX.utils.aoa_to_sheet([[], [""], ["Name", "City"], ["Ada", "Paris"]]);
	const book = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(book, sheet, "S");
	const buffer = XLSX.write(book, { type: "array", bookType: "xlsx" });
	const workbook = XLSX.read(new Uint8Array(buffer), { type: "array", cellDates: true });
	const grid = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {
		header: 1,
		raw: false,
		defval: ""
	});
	const result = dataparse.parseGrid(grid);
	assert.equal(result.ok, true);
	assert.deepEqual(result.data, [["Name", "City"], ["Ada", "Paris"]]);
});

test("parseJson keeps every object record, even one whose values are all empty", () => {
	// One object is one recipient: dropping all-blank records would silently
	// lose recipients (and turn a one-record file into "no data rows").
	const allBlank = dataparse.parseJson('[{"Name": null}]');
	assert.equal(allBlank.ok, true);
	assert.deepEqual(allBlank.data, [["Name"], [""]]);
	const mixed = dataparse.parseJson('[{}, {"Name": "Ada"}]');
	assert.equal(mixed.ok, true);
	assert.deepEqual(mixed.data, [["Name"], [""], ["Ada"]]);
});

test("parseJson converts non-string values deterministically", () => {
	const result = dataparse.parseJson(
		'[{"n": 1.5, "nil": null, "flag": false, "obj": {"x": 1}, "arr": [1, 2], "s": " keep "}]'
	);
	assert.equal(result.ok, true);
	assert.deepEqual(result.data, [
		["n", "nil", "flag", "obj", "arr", "s"],
		["1.5", "", "false", '{"x":1}', "[1,2]", " keep "]
	]);
	assert.ok(allStrings(result.data), "every cell is a string, never [object Object]");
});

test("parseJson arrays-of-arrays also skip blank leading rows", () => {
	const result = dataparse.parseJson('[[], ["a"], ["1"]]');
	assert.equal(result.ok, true);
	assert.deepEqual(result.data, [["a"], ["1"]]);
	assert.equal(dataparse.parseJson("[[], []]").error, "no data rows");
});

test("one recipient loads; zero recipients are refused with no data rows", () => {
	for (const result of [
		dataparse.parseCsv("Name\nAda", Papa),
		dataparse.parseJson('[{"Name": "Ada"}]'),
		dataparse.parseGrid([["Name"], ["Ada"]])
	]) {
		assert.equal(result.ok, true);
		assert.deepEqual(result.data, [["Name"], ["Ada"]]);
	}
	assert.equal(dataparse.parseGrid([["Name"]]).error, "no data rows");
});

test("parsing 2000 rows stays linear in row count", () => {
	const lines = ["Name,City"];
	for (let i = 0; i < 2000; i++) {
		lines.push("recipient " + i + ",city " + i);
	}
	const started = Date.now();
	const result = dataparse.parseCsv(lines.join("\r\n") + "\r\n", Papa);
	const elapsed = Date.now() - started;
	assert.equal(result.ok, true);
	assert.equal(result.data.length, 2001);
	assert.deepEqual(result.data[1], ["recipient 0", "city 0"]);
	assert.deepEqual(result.data[2000], ["recipient 1999", "city 1999"]);
	// A quadratic parse would blow far past this even on a slow machine; the
	// bound is a regression tripwire, not a benchmark.
	assert.ok(elapsed < 3000, "parsing 2000 rows took " + elapsed + "ms");
});
