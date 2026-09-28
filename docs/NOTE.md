# Session notes — mail merge plugin (2026-09-28)

Durable facts: `docs/MEMORY.md`. User contract: `README.md`. This file keeps dated
findings, decisions and post-mortems only (map: README.md "Docs map").

## SDK engine evidence (verified on host, onlyoffice-git 9.4.0.130)

Behaviour mined to `docs/MEMORY.md` facts 1-2. Re-check: grep `Api.MailMerge` in
`/opt/onlyoffice/desktopeditors/editors/sdkjs/word/sdk-all.js` (JSDoc
`@typeofeditors ["CDE"]`).

## Host install findings (2026-09-28)

- This host's per-user sdkjs-plugins root shipped 3 zero-byte `v1/` stubs
  (`plugins.js`, `plugins-ui.js`, `plugins.css`): a per-user-only plugin enumerates
  (inotify-proven) but `Asc.plugin` is undefined and the UI is silently inert.
  latex-math escaped this only because its `/opt` copy shadows the user copy. Fixed
  2026-09-28 by copying the `/opt` v1 trio into the user root. Policy home:
  latex-math README § Install (`docs/MEMORY.md` fact 3).
- Batch document edits into one history point (`CreateNewHistoryPoint`) per merge.

## Search-and-replace semantics (learned from sdk-all.js, 2026-09-28)

- Caret-marker scheme and `unsafe-value` refusal: `docs/MEMORY.md` fact 9;
  decision: `docs/adr/0001`.

## Decisions (2026-09-28)

- Caret-safe plain replacement: `docs/adr/0001`.
- Both placeholder styles (wrap matched, plain-replace unmatched) + combined default:
  `docs/adr/0002`.
- Per-record output mutates the open document: `docs/adr/0003`.
- CSV + XLSX + JSON data sources, parsed in-iframe (PapaParse / SheetJS vendored).
- Out of scope: email output, marketplace packaging, Document Builder batch mode.

## Wizard UI (code.js / styles.css / icons) — 2026-09-28

- `plugin/scripts/code.js` is the composition root: 3-step wizard in `#mm-app`,
  sequential `run()` pipeline (sequence: its header).
- Close semantics: `docs/MEMORY.md` fact 8.
- Combined-mode plain-replacement limitation found here (plain tokens come out blank):
  README.md "Limitations".

## Empty plugin window regression — 2026-09-28, second session

- Symptom: plugin registered but the window rendered empty — "no event handlers, can't
  close, no function at all".
- Root cause: `index.html` loads its scripts in `<head>`, `#mm-app` sits in `<body>`, and
  `code.js` ran `installHostHooks(); mount(document.getElementById("mm-app"))` at
  script-parse time; `mount()` no-ops on a null root, so the wizard was never built.
  Compounding: `installHostHooks()` silently bailed when `window.Asc.plugin` was not yet
  defined and was never retried, so variation buttons and the X/close route were never
  wired. The existing tests masked it — their harness mounted into a pre-built root.
- Fix: boot on DOM-ready (`whenDocumentReady` + mount-once guard + late-Asc retry
  polling), double-install-safe host hooks, hardened `closeWindow` fallback walk
  (resolution order: `closeWindow` comment in code.js).
- Test evidence: `tests/code.test.js` regressions (real head-before-body load order,
  already-parsed document, late Asc, close fallback walk) — 3 of 4 red on old code;
  `tests/host-integration.test.js` runs the real script chain incl. the `/opt` `v1/plugins.js`
  shim and real vendored libs in `node:vm` — red at the `.mm-wrap` group on buggy code.
- `v1/plugins.js` shim findings: it copies the URL `?windowID=` into
  `Asc.plugin.windowID` only late, via its `window.onload` config.json XHR callback;
  `executeMethod` routes per-window only once `windowID` is set — hence the URL-param
  first branch in `closeWindow`.

## Docs restructure — 2026-09-28, third session

- The doc set carried a live contradiction on close semantics (four homes, four stories)
  and facts restated 2-5x across README/MEMORY/NOTE/code headers; the journal also held
  stale test/icon counts. Restructured to one fact one home: durable
  facts in `docs/MEMORY.md` (facts 8-10 added/rewritten), user contract in `README.md`
  ("Limitations"), module contracts in script headers, decisions in `docs/adr/0001-0003`,
  vocabulary in `CONTEXT.md`, and the docs map + check commands in README. Counts live
  nowhere — the check runners live in README.md "Docs map".
- Pitfall: a documented `sed` range check whose pattern text appears in its own command
  re-opens its range and leaks the whole tail of the file into the "extract" (the check
  matches its own pointer text). The recipes-sync check in README therefore spells its
  patterns `<![-]-` — a character-class trick so the pattern cannot match its own text.
