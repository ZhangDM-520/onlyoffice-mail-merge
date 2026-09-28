/*
 * End-to-end host load-order test: the plugin's REAL index.html script chain,
 * executed the way ONLYOFFICE Desktop Editors actually loads a window-variation
 * plugin. This is the regression test for the "wizard renders EMPTY" bug, where
 * plugin/scripts/code.js auto-mounted at script-parse time - but index.html
 * places every <script> in <head> and <div id="mm-app"> in <body>, so at parse
 * time document.getElementById("mm-app") is still null and mount() silently
 * no-ops. The CONTRACT under test: the wizard boots on DOM-ready, so after
 * <body> has parsed and DOMContentLoaded has fired, #mm-app must contain the
 * wizard DOM. This test is EXPECTED to be RED against a parse-time code.js.
 *
 * Browser parse-order model (requirement 2):
 *   1. <head> is parsed: every <script src> runs NOW, in document order, while
 *      document.body is still null and #mm-app does not exist (asserted).
 *   2. Only afterwards is <body> parsed: <div id="mm-app"> appears.
 *   3. DOMContentLoaded fires (then window "load", as in a real browser); timer
 *      ticks are flushed so a DOM-ready boot using setTimeout(0) also lands.
 *   4. Only then do the assertions run.
 *
 * Real vs stubbed:
 *   REAL     - plugin/index.html (parsed for the script chain), and every
 *              script in it executed for real inside node:vm, in order:
 *                ../v1/plugins.js      (see "host shim" below)
 *                ../v1/plugins-ui.js   (see "host shim" below)
 *                vendor/papaparse.min.js  (real, no stub - it runs fine in vm)
 *                vendor/xlsx.full.min.js  (real, no stub - it runs fine in vm)
 *                scripts/dataparse.js, fieldmap.js, commands.js, code.js
 *   HOST SHIM - the REAL ONLYOFFICE plugins.js/plugins-ui.js are loaded
 *              read-only at test time (they are AGPL, same as this repo; a
 *              tests/fixtures/ copy was the preferred option but this session's
 *              write guard only permits this file, so the sanctioned
 *              read-only + fallback route is used):
 *                1. tests/fixtures/host-v1/{plugins.js,plugins-ui.js} if a
 *                   vendored copy is added later (checked first),
 *                2. else /opt/onlyoffice/desktopeditors/editors/sdkjs-plugins/
 *                   v1/ (present on the dev host; exercised here),
 *                3. else a MINIMAL inline replica of the parse-time surface of
 *                   plugins.js (Asc.plugin bootstrap, tr(), Asc.scope) - this
 *                   fallback is declared in the header and surfaced via
 *                   test diagnostics; plugins-ui.js (PerfectScrollbar only,
 *                   unused by our code) degrades to a no-op comment script.
 *   STUBBED  - a thin hand-rolled DOM (no jsdom, zero new deps; the wizard
 *              needs createElement/createTextNode/appendChild/classList/
 *              addEventListener/getElementById and little else), the
 *              host-injected plugin API (executeMethod/callCommand/resizeWindow
 *              recorders - the real host defines these by eval-ing its
 *              "plugin_init" payload into the frame, see plugins.js's message
 *              handler), XMLHttpRequest (serves the real plugin/config.json to
 *              the real plugins.js onload path), localStorage, and
 *              parent.postMessage (recorded). window/document/addEventListener
 *              and friends are provided, as plugins.js expects.
 *
 * Style matches the existing tests: node:test + node:assert/strict + node:vm.
 */
"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.join(__dirname, "..");
const PLUGIN_DIR = path.join(ROOT, "plugin");

// host shim candidates, in preference order (see header).
const HOST_V1_DIRS = [
	path.join(__dirname, "fixtures", "host-v1"),
	"/opt/onlyoffice/desktopeditors/editors/sdkjs-plugins/v1"
];

// The load chain that plugin/index.html must declare, in document order.
const EXPECTED_SCRIPTS = [
	"../v1/plugins.js",
	"../v1/plugins-ui.js",
	"vendor/papaparse.min.js",
	"vendor/xlsx.full.min.js",
	"scripts/dataparse.js",
	"scripts/fieldmap.js",
	"scripts/commands.js",
	"scripts/code.js"
];

const WINDOW_ID = "w-42";
const LOCATION_SEARCH =
	"?windowID=" + WINDOW_ID + "&guid=asc.2A0D08A5-D356-4057-A366-8A2AA579B4D7";

// FALLBACK ONLY (see header): the parse-time surface of the real plugins.js -
// Asc.plugin bootstrap + identity tr() + Asc.scope + supportOrigins. Used only
// when neither a fixture copy nor the installed Desktop Editors SDK is present.
const MINIMAL_HOST_BOOTSTRAP = [
	"/* FALLBACK: minimal replica of the parse-time surface of ONLYOFFICE's",
	"   sdkjs-plugins/v1/plugins.js (real shim unavailable on this machine). */",
	"(function (window) {",
	"\twindow.Asc = window.Asc || {};",
	"\twindow.Asc.plugin = window.Asc.plugin || {};",
	"\twindow.Asc.plugin.tr_init = false;",
	"\twindow.Asc.plugin.tr = function (text) { return text; };",
	"\twindow.Asc.scope = window.Asc.scope || {};",
	"\twindow.Asc.supportOrigins = window.Asc.supportOrigins || {};",
	"})(window);"
].join("\n");

/* ------------------------------------------------------------------ *
 * Thin DOM - only what plugins.js and plugin/scripts/** touch
 * ------------------------------------------------------------------ */

function TextNode(data, ownerDocument) {
	this.nodeType = 3;
	this.data = String(data);
	this.parentNode = null;
	this.ownerDocument = ownerDocument;
}

function El(tag, ownerDocument) {
	this.nodeType = 1;
	this.tagName = String(tag).toUpperCase();
	this.ownerDocument = ownerDocument;
	this.parentNode = null;
	this.children = [];
	this.attributes = {};
	this.style = {};
	this.className = "";
	this.id = "";
	this.value = "";
	this.checked = false;
	this.disabled = false;
	this.files = null;
	this.listeners = {};
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

Object.defineProperty(El.prototype, "firstChild", {
	get: function () {
		return this.children.length ? this.children[0] : null;
	}
});

Object.defineProperty(El.prototype, "textContent", {
	get: function () {
		return this.children
			.map(function (child) {
				return child.nodeType === 3 ? child.data : child.textContent;
			})
			.join("");
	},
	set: function (text) {
		const node = new TextNode(text, this.ownerDocument);
		node.parentNode = this;
		this.children = [node];
	}
});

El.prototype.appendChild = function (child) {
	if (child.parentNode) {
		child.parentNode.removeChild(child);
	}
	child.parentNode = this;
	this.children.push(child);
	return child;
};

El.prototype.removeChild = function (child) {
	const index = this.children.indexOf(child);
	if (index === -1) {
		throw new Error("removeChild: not a child");
	}
	this.children.splice(index, 1);
	child.parentNode = null;
	return child;
};

El.prototype.setAttribute = function (name, value) {
	this.attributes[name] = String(value);
	if (name === "id") {
		this.id = String(value);
	}
	if (name === "class") {
		this.className = String(value);
	}
};

El.prototype.getAttribute = function (name) {
	return Object.prototype.hasOwnProperty.call(this.attributes, name)
		? this.attributes[name]
		: null;
};

El.prototype.addEventListener = function (type, fn) {
	if (typeof fn !== "function") {
		return;
	}
	this.listeners[type] = (this.listeners[type] || []).concat(fn);
};

El.prototype.removeEventListener = function (type, fn) {
	this.listeners[type] = (this.listeners[type] || []).filter(function (item) {
		return item !== fn;
	});
};

El.prototype.dispatchEvent = function (event) {
	event.target = event.target || this;
	(this.listeners[event.type] || []).slice().forEach(function (fn) {
		fn.call(this, event);
	}, this);
	const handler = this["on" + event.type];
	if (typeof handler === "function") {
		handler.call(this, event);
	}
	return true;
};

function findDescendant(root, predicate) {
	for (const child of root.children || []) {
		if (child.nodeType !== 1) {
			continue;
		}
		if (predicate(child)) {
			return child;
		}
		const hit = findDescendant(child, predicate);
		if (hit) {
			return hit;
		}
	}
	return null;
}

function makeEventTargetMethods(target) {
	target.listeners = {};
	target.addEventListener = function (type, fn) {
		if (typeof fn !== "function") {
			return;
		}
		target.listeners[type] = (target.listeners[type] || []).concat(fn);
	};
	target.removeEventListener = function (type, fn) {
		target.listeners[type] = (target.listeners[type] || []).filter(function (item) {
			return item !== fn;
		});
	};
	target.dispatchEvent = function (event) {
		event.target = event.target || target;
		(target.listeners[event.type] || []).slice().forEach(function (fn) {
			fn.call(target, event);
		});
		const handler = target["on" + event.type];
		if (typeof handler === "function") {
			handler.call(target, event);
		}
		return true;
	};
}

/* ------------------------------------------------------------------ *
 * Sandbox: window/document + the host contract plugins.js expects
 * ------------------------------------------------------------------ */

function makeSandbox(record) {
	const doc = {};
	makeEventTargetMethods(doc);
	// Browser parse state while <head> scripts run: <body> is NOT parsed yet.
	doc.readyState = "loading";
	doc.documentElement = new El("html", doc);
	doc.body = null;
	doc.createElement = function (tag) {
		return new El(tag, doc);
	};
	doc.createTextNode = function (data) {
		return new TextNode(data, doc);
	};
	doc.getElementById = function (id) {
		return findDescendant(doc.documentElement, function (el) {
			return el.id === id || el.attributes.id === id;
		});
	};

	const sb = {};
	makeEventTargetMethods(sb);
	sb.window = sb;
	sb.self = sb;
	sb.globalThis = sb;
	sb.document = doc;
	sb.console = console;
	sb.setTimeout = setTimeout;
	sb.clearTimeout = clearTimeout;
	sb.queueMicrotask = queueMicrotask;
	sb.requestAnimationFrame = function (fn) {
		return setTimeout(fn, 0);
	};
	sb.navigator = { userAgent: "ONLYOFFICE DesktopEditors host-integration.test.js" };
	sb.location = {
		search: LOCATION_SEARCH,
		href: "file://" + path.join(PLUGIN_DIR, "index.html") + LOCATION_SEARCH
	};
	sb.origin = "null";
	sb.crypto = {
		getRandomValues: function (array) {
			for (let i = 0; i < array.length; i++) {
				array[i] = Math.floor(Math.random() * 65536);
			}
			return array;
		}
	};
	// ONLYOFFICE's plugins.js onload XHR-fetches ./config.json and merges it
	// into Asc.plugin; serve the real plugin/config.json.
	sb.XMLHttpRequest = function FakeXHR() {
		this.method = null;
		this.url = null;
		this.status = 0;
		this.readyState = 0;
		this.response = null;
		this.responseType = "";
		this.onload = null;
		this.onerror = null;
		record.xhr.push(this);
	};
	sb.XMLHttpRequest.prototype.open = function (method, url) {
		this.method = method;
		this.url = url;
		this.readyState = 1;
	};
	sb.XMLHttpRequest.prototype.send = function () {
		this.status = 200;
		this.readyState = 4;
		this.response = JSON.parse(fs.readFileSync(path.join(PLUGIN_DIR, "config.json"), "utf8"));
		if (typeof this.onload === "function") {
			this.onload();
		}
	};
	sb.parent = {
		postMessage: function (message) {
			record.postMessage.push(message);
		}
	};
	const storage = new Map();
	sb.localStorage = {
		getItem: function (key) {
			return storage.has(String(key)) ? storage.get(String(key)) : null;
		},
		setItem: function (key, value) {
			storage.set(String(key), String(value));
		},
		removeItem: function (key) {
			storage.delete(String(key));
		},
		clear: function () {
			storage.clear();
		}
	};
	return { sb, doc };
}

// The real host defines executeMethod/callCommand/resizeWindow by eval-ing its
// "plugin_init" payload into the frame (see the message handler in plugins.js).
// Modelled here as the same injection landing before DOMContentLoaded, with
// every call recorded so tests can assert on them.
function installHostApi(sb, record) {
	assert.ok(sb.Asc && sb.Asc.plugin, "plugins.js must have bootstrapped Asc.plugin before the host API is injected");
	sb.Asc.plugin.executeMethod = function (name, args, callback) {
		record.executeMethod.push({ name: String(name), args: args === undefined ? [] : args });
		if (typeof callback === "function") {
			callback();
		}
	};
	sb.Asc.plugin.callCommand = function (command, isClose, isCalc, callback) {
		record.callCommand.push({ command: String(command), isClose: !!isClose, isCalc: !!isCalc });
		if (typeof callback === "function") {
			callback();
		}
	};
	sb.Asc.plugin.resizeWindow = function (width, height) {
		record.resizeWindow.push([width, height]);
	};
}

/* ------------------------------------------------------------------ *
 * Script chain resolution and the boot sequence
 * ------------------------------------------------------------------ */

function loadHostScript(fileName) {
	for (const dir of HOST_V1_DIRS) {
		const full = path.join(dir, fileName);
		if (fs.existsSync(full)) {
			return { code: fs.readFileSync(full, "utf8"), origin: full };
		}
	}
	return null;
}

function resolveScriptChain() {
	const html = fs.readFileSync(path.join(PLUGIN_DIR, "index.html"), "utf8");
	const srcs = [];
	const re = /<script\b[^>]*\bsrc="([^"]+)"[^>]*>/g;
	let match;
	while ((match = re.exec(html)) !== null) {
		srcs.push(match[1]);
	}

	const chain = srcs.map(function (src) {
		if (src.indexOf("../v1/") === 0) {
			const fileName = src.slice("../v1/".length);
			const host = loadHostScript(fileName);
			if (host) {
				return { src: src, code: host.code, filename: host.origin, origin: host.origin };
			}
			if (fileName === "plugins.js") {
				// Declared fallback (see header): minimal Asc.plugin bootstrap.
				return {
					src: src,
					code: MINIMAL_HOST_BOOTSTRAP,
					filename: "inline:minimal-plugins.js-bootstrap",
					origin: "minimal-inline-fallback"
				};
			}
			// plugins-ui.js only defines PerfectScrollbar, which nothing here
			// uses; degrade to a no-op so the chain stays intact.
			return {
				src: src,
				code: "/* fallback: " + fileName + " unavailable; only defines PerfectScrollbar */",
				filename: "inline:noop-" + fileName,
				origin: "noop-inline-fallback"
			};
		}
		const full = path.join(PLUGIN_DIR, src);
		return { src: src, code: fs.readFileSync(full, "utf8"), filename: full, origin: full };
	});
	return chain;
}

function flushTimers() {
	// One macrotask + one microtask checkpoint: enough for any DOM-ready boot
	// that defers through setTimeout(0)/queueMicrotask.
	return new Promise(function (resolve) {
		setTimeout(function () {
			setImmediate(resolve);
		}, 0);
	});
}

async function runBoot() {
	const record = {
		executed: [],
		executeMethod: [],
		callCommand: [],
		resizeWindow: [],
		postMessage: [],
		xhr: []
	};
	const chain = resolveScriptChain();
	const { sb, doc } = makeSandbox(record);
	const ctx = vm.createContext(sb);

	// 1. <head> parse: run the whole chain in document order, before <body>.
	for (const item of chain) {
		record.executed.push({
			src: item.src,
			origin: item.origin,
			bodyNull: doc.body === null,
			mmAppMissing: doc.getElementById("mm-app") === null
		});
		vm.runInContext(item.code, ctx, { filename: item.filename });
	}

	// 2. The host's plugin_init payload arrives (executeMethod et al).
	installHostApi(sb, record);

	// 3. The parser reaches <body>: only NOW does #mm-app exist.
	const body = new El("body", doc);
	doc.documentElement.appendChild(body);
	doc.body = body;
	const app = doc.createElement("div");
	app.setAttribute("id", "mm-app");
	body.appendChild(app);

	// 4. DOM-ready window: DOMContentLoaded first, then window "load".
	doc.readyState = "interactive";
	doc.dispatchEvent({ type: "DOMContentLoaded", target: doc });
	sb.dispatchEvent({ type: "DOMContentLoaded", target: sb });
	doc.readyState = "complete";
	sb.dispatchEvent({ type: "load", target: sb });
	doc.dispatchEvent({ type: "load", target: doc });

	await flushTimers();
	await flushTimers();

	return { sb, doc, record, chain };
}

let bootPromise = null;
function bootOnce() {
	if (!bootPromise) {
		bootPromise = runBoot();
	}
	return bootPromise;
}

function hasClass(el, name) {
	return el.classList.contains(name);
}

/* ------------------------------------------------------------------ *
 * Assertion groups
 * ------------------------------------------------------------------ */

test("load chain: index.html scripts run in document order, before <body> exists", async (t) => {
	const { sb, doc, record, chain } = await bootOnce();

	// index.html must declare exactly the expected chain, in order.
	const srcs = chain.map(function (item) {
		return item.src;
	});
	assert.deepEqual(srcs, EXPECTED_SCRIPTS, "plugin/index.html script chain changed");

	// Every script executed exactly once, in document order.
	assert.deepEqual(
		record.executed.map(function (item) {
			return item.src;
		}),
		EXPECTED_SCRIPTS,
		"scripts must execute exactly once, in document order"
	);

	// The browser parse-order model: scripts run while <head> is parsing, so
	// neither <body> nor #mm-app exists yet (this is what made parse-time
	// mount() silently no-op).
	for (const item of record.executed) {
		assert.equal(item.bodyNull, true, item.src + " executed after <body> parsed - harness parse-order broken");
		assert.equal(item.mmAppMissing, true, item.src + " executed with #mm-app present - harness parse-order broken");
	}

	// Side effects prove each layer really ran: the host bootstrap (real
	// plugins.js or the declared fallback), the vendor libs (REAL, not
	// stubbed), and our three UMD modules.
	assert.ok(sb.Asc && sb.Asc.plugin, "plugins.js must bootstrap window.Asc.plugin");
	assert.equal(typeof sb.Asc.plugin.tr, "function", "plugins.js must provide Asc.plugin.tr");
	assert.ok(sb.Papa !== undefined, "vendor/papaparse.min.js must define window.Papa");
	assert.ok(sb.XLSX !== undefined, "vendor/xlsx.full.min.js must define window.XLSX");
	assert.ok(sb.OnlyOfficeMailMergeDataParse, "scripts/dataparse.js must define OnlyOfficeMailMergeDataParse");
	assert.ok(sb.OnlyOfficeMailMergeFieldMap, "scripts/fieldmap.js must define OnlyOfficeMailMergeFieldMap");
	assert.ok(sb.OnlyOfficeMailMergeCommands, "scripts/commands.js must define OnlyOfficeMailMergeCommands");
	assert.ok(sb.OnlyOfficeMailMergeUi, "scripts/code.js must define OnlyOfficeMailMergeUi");

	t.diagnostic("host shim: " + record.executed[0].origin);
	assert.equal(doc.readyState, "complete");
});

test("DOM ready: #mm-app contains the wizard DOM (.mm-wrap) - the empty-window regression", async () => {
	const { doc } = await bootOnce();

	const app = doc.getElementById("mm-app");
	assert.ok(app, "#mm-app must exist once <body> has parsed");

	// THE regression assertion: with parse-time mounting, mount(null) no-ops
	// and #mm-app stays empty forever.
	const wrap = findDescendant(app, function (el) {
		return hasClass(el, "mm-wrap");
	});
	assert.ok(
		wrap,
		'after DOMContentLoaded, #mm-app must contain the wizard root ".mm-wrap" - ' +
			"an empty #mm-app is exactly the empty-wizard bug this test exists for"
	);
	assert.ok(wrap.children.length > 0, "the wizard root must contain rendered children");
});

test("DOM ready: window.Asc.plugin.init and .button are functions", async () => {
	const { sb } = await bootOnce();

	assert.ok(sb.Asc && sb.Asc.plugin, "Asc.plugin must exist");
	assert.equal(
		typeof sb.Asc.plugin.init,
		"function",
		"window.Asc.plugin.init must be a function after DOM-ready (host entry point)"
	);
	assert.equal(
		typeof sb.Asc.plugin.button,
		"function",
		"window.Asc.plugin.button must be a function after DOM-ready (host entry point)"
	);
});

test("host button(-1) routes to Asc.plugin.executeMethod('CloseWindow')", async (t) => {
	const { sb, record } = await bootOnce();

	// The host drives the two mandatory hooks, exactly as Desktop Editors does.
	sb.Asc.plugin.init();
	assert.equal(record.executeMethod.length, 0, "init() alone must not issue executeMethod calls");

	const before = record.executeMethod.length;
	sb.Asc.plugin.button(-1, WINDOW_ID);
	const calls = record.executeMethod.slice(before);

	const closeCalls = calls.filter(function (call) {
		return call.name === "CloseWindow";
	});
	assert.equal(closeCalls.length, 1, "button(-1, wid) must attempt exactly one CloseWindow executeMethod");

	// The window id comes from the page URL (?windowID=..., same source the
	// SDK's plugins.js reads) and must reach the host as the method argument.
	const args = closeCalls[0].args;
	assert.equal(args.length, 1, "CloseWindow must carry exactly one argument (the window id)");
	assert.equal(args[0], WINDOW_ID, "CloseWindow must carry the window id from the page URL");

	t.diagnostic("executeMethod calls: " + JSON.stringify(calls.map(function (c) { return c.name; })));
});
