/*
 * The icon regression test.
 *
 * The defect this exists for: ONLYOFFICE expands `%scale%(default)` into five
 * scales (100/125/150/175/200 %) and then picks the one nearest the display
 * scale, with **no fallback** to the 100 % file. A plugin shipping only 100 %
 * and 200 % renders as *nothing* on a 1.25 display, and nothing anywhere
 * reports an error. So: every icon the plugin asks for must exist at every
 * scale, at the right pixel size, and actually have ink in it.
 *
 * Deliberately independent of the icon generator: it reads what config.json
 * declares and checks the files, so it also holds for a machine that has
 * neither Pillow nor the noctalia font.
 */
"use strict";

const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");
const { execFileSync } = require("node:child_process");

const png = require("./png.js");

const ROOT = path.join(__dirname, "..");
const THEMES = ["light", "dark"];
const SCALES = [
	{ key: "100%", suffix: "" },
	{ key: "125%", suffix: "@1.25x" },
	{ key: "150%", suffix: "@1.5x" },
	{ key: "175%", suffix: "@1.75x" },
	{ key: "200%", suffix: "@2x" }
];

// Base pixel size, matching the shipped plugins: entry and store icons are 28 px
// at 100 % (measured from Send's entry icon and AI's big icons).
const ENTRY_BASE = 28;
const MIN_VISIBLE_FRACTION = 0.05;
// The tile is config.json's store background; the mark is white ink.
const TILE = [0x2a, 0x5d, 0xb0, 255];

function config() {
	return JSON.parse(fs.readFileSync(path.join(ROOT, "plugin", "config.json"), "utf8"));
}

/** Every icon file the plugin declares: entry (per theme) + store. */
function declaredFiles() {
	const files = [];
	const variation = config().variations[0];

	assert.strictEqual(
		typeof variation.icons,
		"string",
		"the entry icon is declared with placeholders, so every scale can be resolved"
	);
	THEMES.forEach(function (theme) {
		SCALES.forEach(function (scale) {
			files.push({
				file: path.join(
					ROOT,
					"plugin",
					variation.icons.replace("%theme-type%(light|dark)", theme).replace("%scale%(default)", scale.suffix)
				),
				scale: scale,
				source: "config.json icons"
			});
		});
	});

	const storeIcons = variation.store && variation.store.icons;
	assert.ok(storeIcons, "config.json declares store icons");
	THEMES.forEach(function (theme) {
		assert.strictEqual(
			storeIcons[theme],
			"resources/store/icons",
			"both themes share the store icon directory"
		);
	});
	SCALES.forEach(function (scale) {
		files.push({
			file: path.join(ROOT, "plugin", "resources", "store", "icons", "icon" + scale.suffix + ".png"),
			scale: scale,
			source: "config.json store.icons"
		});
	});

	return files;
}

test("every declared icon resolves at every scale, at the right size, with ink in it", () => {
	const files = declaredFiles();
	assert.strictEqual(files.length, 15, "2 themes x 5 scales + 5 store icons");

	files.forEach(function (entry) {
		const relative = path.relative(ROOT, entry.file);

		assert.ok(
			fs.existsSync(entry.file),
			"missing " +
				relative +
				" - the host asks for the scale nearest the display one and has no fallback, so this renders as no icon at all (" +
				entry.source +
				")"
		);

		const image = png.readPng(entry.file);
		const expected = Math.round((ENTRY_BASE * parseFloat(entry.scale.key)) / 100);
		assert.strictEqual(image.width, expected, relative + " width");
		assert.strictEqual(image.height, expected, relative + " height");

		const floor = Math.floor(expected * expected * MIN_VISIBLE_FRACTION);
		const visible = png.visiblePixels(image);
		assert.ok(visible >= floor, relative + " is blank (" + visible + " visible pixels)");
	});
});

test("the icon is the mail-merge mark on the store-blue tile, not an empty square", () => {
	const file = path.join(ROOT, "plugin", "resources", "light", "icon@1.25x.png");
	const image = png.readPng(file);
	const middle = Math.floor(image.height / 2);

	// The left edge's midpoint is tile and always outside the glyph.
	const edge = png.pixel(image, 1, middle);
	assert.deepStrictEqual(edge, TILE, "the edge is the solid tile: " + edge);

	// The mark is the only thing on the tile in white ink.
	let white = 0;
	for (let y = 0; y < image.height; y++) {
		for (let x = 0; x < image.width; x++) {
			const pixel = png.pixel(image, x, y);
			const near = pixel[0] > 210 && pixel[1] > 210 && pixel[2] > 210;
			if (pixel[3] > 200 && near) {
				white++;
			}
		}
	}
	assert.ok(white >= 10, "the mark is drawn on the tile (" + white + " white pixels)");
	assert.notDeepStrictEqual(
		png.pixel(image, Math.floor(image.width / 2), middle),
		TILE,
		"the centre of the tile is the glyph, not more tile"
	);
});

test("the generator's manifest covers the icon slot config.json asks for", () => {
	let listing;
	try {
		listing = execFileSync("python3", [path.join("tools", "make-icons.py"), "--list"], {
			cwd: ROOT,
			encoding: "utf8"
		});
	} catch (error) {
		// The icon files themselves are checked above without Python.
		return;
	}

	const slots = listing
		.split("\n")
		.map(function (line) {
			return /^(\S+)\s+(\S+)$/.exec(line);
		})
		.filter(Boolean)
		.map(function (match) {
			return match[1];
		});
	assert.ok(slots.indexOf("icon") !== -1, "the manifest lists the entry/store slot: " + slots.join(", "));

	// config.json's entry pattern resolves to the `icon` slot at every scale.
	assert.match(config().variations[0].icons, /\(light\|dark\)\/icon%scale%\(default\)\.png$/);
});

test("the generator's own check passes and leaves no orphans", () => {
	let result;
	try {
		result = execFileSync("python3", [path.join("tools", "make-icons.py"), "--check"], {
			cwd: ROOT,
			encoding: "utf8"
		});
	} catch (error) {
		if (/Pillow is required/.test(String(error.stdout) + String(error.stderr) + String(error.message))) {
			// Without Pillow the committed PNGs are still checked above.
			return;
		}
		assert.fail("make-icons.py --check failed:\n" + (error.stdout || error.message));
	}
	assert.match(result, /files ok/, result.trim());

	// Every PNG under resources must belong to the manifest, or a rename would
	// leave an unused file behind that nothing regenerates.
	const onDisk = [];
	(function walk(directory) {
		fs.readdirSync(directory, { withFileTypes: true }).forEach(function (entry) {
			const full = path.join(directory, entry.name);
			if (entry.isDirectory()) {
				walk(full);
			} else if (entry.name.endsWith(".png")) {
				onDisk.push(full);
			}
		});
	})(path.join(ROOT, "plugin", "resources"));

	assert.ok(onDisk.length >= 15, "the plugin ships a full icon set: " + onDisk.length + " files");
	onDisk.forEach(function (file) {
		const name = path.basename(file).replace(/@[\d.]+x/, "").replace(/\.png$/, "");
		assert.strictEqual(name, "icon", path.relative(ROOT, file) + " is not in the generator manifest");
	});
});
