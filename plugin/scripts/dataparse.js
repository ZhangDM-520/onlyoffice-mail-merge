/*
 * Data-source parsing for the Mail Merge plugin: CSV, XLSX-derived grids and
 * JSON documents in, ONE canonical shape out — `String[][]` where row 0 holds
 * the merge-field names and every later row is one recipient's values. That is
 * exactly the shape `Api.LoadMailMergeData` consumes.
 *
 * Contract (all functions pure, no DOM, no editor calls):
 *   parseCsv(text, papa?)             -> result
 *   parseJson(text)                   -> result
 *   parseGrid(grid, opts?)            -> result   (XLSX sheets arrive as a grid)
 *   toLoadMailMergeData(okResult)     -> String[][]
 *
 *   result := { ok: true,  data: String[][] , warnings: string[] }
 *           | { ok: false, error: string }
 *
 * Rules:
 *   - strip a UTF-8 BOM; trim header cells; headers must be non-empty and unique
 *     (duplicate/empty headers -> { ok:false, error:'...' }); uniqueness is
 *     judged the way tokens meet headers - normalized and case-insensitively,
 *     so "Name" and "name" cannot both back {{name}}
 *   - the header row is the first row that carries anything: fully empty rows
 *     before it are dropped (XLSX used ranges and CSV exports start with them)
 *   - at least one data row is required
 *   - fully empty rows are dropped; ragged rows are padded with ""; rows built
 *     from JSON OBJECTS are kept even when every value is empty - one object is
 *     one recipient, and a recipient with blank values is still a recipient
 *   - every cell is a string; callers convert types
 *   - parseCsv uses the injected Papa parser, falling back to root.Papa
 *   - parseGrid treats a leading row as headers; opts.width forces column count
 */
(function (root, factory) {
	if (typeof module === "object" && module.exports) {
		module.exports = factory(root);
	} else {
		root.OnlyOfficeMailMergeDataParse = factory(root);
	}
})(typeof self !== "undefined" ? self : this, function (root) {
	"use strict";

	function fail(error) {
		return { ok: false, error: error };
	}

	function stripBom(text) {
		return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
	}

	// XLSX hands cells over as raw JS values while the merge table is
	// strings-only. A Date cell means the calendar date it holds, so it goes
	// through UTC to keep the printed day stable across timezones. JSON values
	// may be nested objects/arrays: String() would print "[object Object]", so
	// they keep their serialized form instead of silent garbage.
	function cellText(value) {
		if (value === null || value === undefined) {
			return "";
		}
		if (value instanceof Date) {
			return isNaN(value.getTime()) ? "" : value.toISOString().slice(0, 10);
		}
		if (typeof value === "object") {
			try {
				var json = JSON.stringify(value);
				return typeof json === "string" ? json : "";
			} catch (error) {
				return String(value);
			}
		}
		return String(value);
	}

	function isBlankRow(row) {
		for (var i = 0; i < row.length; i++) {
			if (cellText(row[i]) !== "") {
				return false;
			}
		}
		return true;
	}

	// Index of the first row that carries anything (any cell). Leading fully
	// empty rows are not headers - XLSX used ranges and CSV exports start with
	// them - so the header row is the first row that says something. Rows that
	// are not arrays stop the scan: they cannot be judged and the caller's own
	// validation reports them.
	function firstContentRow(rows) {
		for (var i = 0; i < rows.length; i++) {
			if (!Array.isArray(rows[i]) || !isBlankRow(rows[i])) {
				return i;
			}
		}
		return -1;
	}

	/*
	 * The one place a source turns into the canonical table. `width` forces the
	 * rectangular shape LoadMailMergeData needs: ragged rows are padded with "",
	 * and cells beyond the width are dropped - they have no header to merge
	 * into, so keeping them would desynchronize rows from field names.
	 * `keepBlankRows` skips the empty-row drop: JSON OBJECT rows are recipients,
	 * and one whose every value is empty is still a recipient (see the header).
	 */
	function build(headerCells, dataRows, width, keepBlankRows) {
		if (width === 0) {
			return fail("no headers");
		}
		var headers = [];
		for (var c = 0; c < width; c++) {
			var header = cellText(headerCells[c]);
			if (c === 0) {
				header = stripBom(header);
			}
			headers.push(header.trim());
		}
		var seen = Object.create(null);
		for (var h = 0; h < headers.length; h++) {
			if (headers[h] === "") {
				return fail("empty header at column " + (h + 1));
			}
			// Uniqueness follows the token-matching rules (normalized,
			// case-insensitive): "Name" and "name" would both back {{name}},
			// and which one won would depend on the consumer.
			var key = headers[h].replace(/\s+/g, " ").toLowerCase();
			if (Object.prototype.hasOwnProperty.call(seen, key)) {
				return fail('duplicate header "' + headers[h] + '"');
			}
			seen[key] = true;
		}
		var rows = [];
		for (var r = 0; r < dataRows.length; r++) {
			var raw = dataRows[r];
			if (!Array.isArray(raw)) {
				return fail("row " + (r + 2) + ": expected an array of cells");
			}
			if (!keepBlankRows && isBlankRow(raw)) {
				continue;
			}
			var row = [];
			for (var k = 0; k < width; k++) {
				row.push(k < raw.length ? cellText(raw[k]) : "");
			}
			rows.push(row);
		}
		if (rows.length === 0) {
			return fail("no data rows");
		}
		return { ok: true, data: [headers].concat(rows), warnings: [] };
	}

	function parseCsv(text, papa) {
		if (typeof text !== "string") {
			return fail("invalid input: expected a string");
		}
		var parser = papa || root.Papa;
		var parse = typeof parser === "function" ? parser
			: parser && typeof parser.parse === "function" ? function (input) {
				return parser.parse(input);
			} : null;
		if (!parse) {
			return fail("papa-unavailable");
		}
		var parsed = parse(stripBom(text));
		var rows = Array.isArray(parsed) ? parsed
			: parsed && Array.isArray(parsed.data) ? parsed.data : null;
		if (!rows) {
			return fail("csv parse failed");
		}
		if (rows.length === 0) {
			return fail("no data rows");
		}
		var start = firstContentRow(rows);
		if (start === -1) {
			return fail("no data rows");
		}
		if (!Array.isArray(rows[start])) {
			return fail("row " + (start + 1) + ": expected an array of cells");
		}
		return build(rows[start], rows.slice(start + 1), rows[start].length);
	}

	function parseJson(text) {
		if (typeof text !== "string") {
			return fail("invalid input: expected a string");
		}
		var doc;
		try {
			doc = JSON.parse(stripBom(text));
		} catch (error) {
			return fail("invalid JSON: " + error.message);
		}
		if (!Array.isArray(doc)) {
			return fail("unsupported JSON: expected an array");
		}
		if (doc.length === 0) {
			return fail("no data rows");
		}
		if (Array.isArray(doc[0])) {
			var start = firstContentRow(doc);
			if (start === -1) {
				return fail("no data rows");
			}
			if (!Array.isArray(doc[start])) {
				return fail("row " + (start + 1) + ": expected an array of cells");
			}
			return build(doc[start], doc.slice(start + 1), doc[start].length);
		}
		if (doc[0] !== null && typeof doc[0] === "object") {
			// Headers are the union of the records' keys in first-appearance
			// order, so a key that only shows up late still gets its column.
			var headers = [];
			var seen = Object.create(null);
			for (var r = 0; r < doc.length; r++) {
				var record = doc[r];
				if (record === null || typeof record !== "object" || Array.isArray(record)) {
					return fail("unsupported JSON row " + (r + 1) + ": expected an object");
				}
				var keys = Object.keys(record);
				for (var k = 0; k < keys.length; k++) {
					if (!Object.prototype.hasOwnProperty.call(seen, keys[k])) {
						seen[keys[k]] = true;
						headers.push(keys[k]);
					}
				}
			}
			var rows = [];
			for (var i = 0; i < doc.length; i++) {
				var cells = [];
				for (var h = 0; h < headers.length; h++) {
					cells.push(Object.prototype.hasOwnProperty.call(doc[i], headers[h])
						? doc[i][headers[h]] : "");
				}
				rows.push(cells);
			}
			// Every object is a recipient, even one whose every value is empty.
			return build(headers, rows, headers.length, true);
		}
		return fail("unsupported JSON row 1: expected an object or an array");
	}

	function parseGrid(grid, opts) {
		if (!Array.isArray(grid)) {
			return fail("invalid grid: expected an array of rows");
		}
		if (grid.length === 0) {
			return fail("no data rows");
		}
		for (var i = 0; i < grid.length; i++) {
			if (!Array.isArray(grid[i])) {
				return fail("invalid grid row " + (i + 1) + ": expected an array of cells");
			}
		}
		// A sheet's used range often starts with fully empty rows before the
		// header row; the header is the first row that carries anything.
		var start = firstContentRow(grid);
		if (start === -1) {
			return fail("no data rows");
		}
		var width = grid[start].length;
		if (opts && typeof opts.width === "number" && opts.width > 0) {
			width = Math.floor(opts.width);
		}
		return build(grid[start], grid.slice(start + 1), width);
	}

	function toLoadMailMergeData(result) {
		if (!result || result.ok !== true || !Array.isArray(result.data)) {
			throw new Error(result && result.error ? result.error : "invalid parse result");
		}
		return result.data;
	}

	return {
		parseCsv: parseCsv,
		parseJson: parseJson,
		parseGrid: parseGrid,
		toLoadMailMergeData: toLoadMailMergeData
	};
});
