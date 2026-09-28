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
 *     (duplicate/empty headers -> { ok:false, error:'...' })
 *   - row 0 is headers; at least one data row is required
 *   - fully empty rows are dropped; ragged rows are padded with ""
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
	// through UTC to keep the printed day stable across timezones.
	function cellText(value) {
		if (value === null || value === undefined) {
			return "";
		}
		if (value instanceof Date) {
			return isNaN(value.getTime()) ? "" : value.toISOString().slice(0, 10);
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

	/*
	 * The one place a source turns into the canonical table. `width` forces the
	 * rectangular shape LoadMailMergeData needs: ragged rows are padded with "",
	 * and cells beyond the width are dropped - they have no header to merge
	 * into, so keeping them would desynchronize rows from field names.
	 */
	function build(headerCells, dataRows, width) {
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
			if (Object.prototype.hasOwnProperty.call(seen, headers[h])) {
				return fail('duplicate header "' + headers[h] + '"');
			}
			seen[headers[h]] = true;
		}
		var rows = [];
		for (var r = 0; r < dataRows.length; r++) {
			var raw = dataRows[r];
			if (!Array.isArray(raw)) {
				return fail("row " + (r + 2) + ": expected an array of cells");
			}
			if (isBlankRow(raw)) {
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
		return build(rows[0], rows.slice(1), rows[0].length);
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
			return build(doc[0], doc.slice(1), doc[0].length);
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
			return build(headers, rows, headers.length);
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
		var width = grid[0].length;
		if (opts && typeof opts.width === "number" && opts.width > 0) {
			width = Math.floor(opts.width);
		}
		return build(grid[0], grid.slice(1), width);
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
