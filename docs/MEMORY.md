# Durable memory — ONLYOFFICE plugin development (this machine)

## Facts that took work to learn (re-verified 2026-09-28)

1. **ONLYOFFICE Desktop Editors ships a full mail-merge engine** in
   `editors/sdkjs/word/sdk-all.js` (`Api.LoadMailMergeData`, `Api.MailMerge(start,end)`,
   `Api.GetMailMergeReceptionsCount`, `Api.GetMailMergeTemplateDocContent`,
   `Api.ReplaceDocumentContent`, `WrapInMailMergeField`) — plugin/macro-reachable, but
   with NO UI in desktop. `Api.MailMerge` replaces the open document body with the merged
   result for the given record range; it does not open new documents.
2. **No plugin API creates or opens a new document.** Per-record file output must loop:
   merge(i,i) -> `GetFileToDownload` (desktop opens an OS save dialog per call) ->
   restore template snapshot -> next.
3. **Plugin install policy has one home:** the latex-math README § Install
   (https://github.com/ZhangDM-520/onlyoffice-latex-math#install) — both roots, braced-GUID
   directory, delete-dest upgrade, cache clear, and the real-`v1/`-loader rule. This
   plugin's GUID recipe: README.md "Install". Manifest truth (`asc.{GUID}`):
   `plugin/config.json`. Host evidence (this machine, `v1/` stubs): docs/NOTE.md.
4. `callCommand` bodies are stringified (no closures); `Asc.scope` is one shared slot
   (serialize dispatch); returns must be JSON strings; `window`/`document`/`alert` are
   blocked inside callCommand since 7.1.
5. Host here: `onlyoffice-git` (AUR, version 9.4.0.130 at time of writing) — git build,
   so SDK surface can drift; re-grep `sdk-all.js` before relying on an Api method.
6. **Browser parse order is part of the contract.** `index.html` loads its scripts in
   `<head>` while the mount root (`#mm-app`) is in `<body>` — anything that mounts at
   script-parse time sees `null` and no-ops (empty plugin window). Boot on DOM-ready and
   mount at most once; re-try host-hook installation too, because `Asc.plugin` can be
   defined late (the `v1/plugins.js` shim fills `windowID` only in its onload config.json
   XHR callback). Test harnesses must model the real load order, not a pre-built root —
   a harness that builds the DOM first hides exactly this class of bug.
7. Installed plugin inventory lives in both roots; the LaTeX math plugin
   (`~/Projects/onlyoffice-latex-math`, GUID `{5B4C1A72-...}`) is the reference sample for
   manifest layout, callCommand seam, icon slots (5 scales x 2 themes, no fallback) and
   test harness (`node --test`, fake-editor double).
8. **Closing a plugin window (verified in the 9.4.0 host, 2026-09-28).** The MAIN
   window-type variation gets NO windowId with its clicks (web-apps calls
   `asc_pluginButtonClick(id, guid)` 2-arg). `executeMethod("CloseWindow", [id])` only
   closes windows registered via `ShowWindow` — it silently no-ops on any other id
   (including `[]`/undefined) and never throws, so never treat a "successful" call as
   evidence the window closed. The ONLY working close for the main window is
   `Asc.plugin.executeCommand("close", "")`. The window X/ESC routes to
   `Asc.plugin.button(-1, …)`; once the plugin defines that hook, the host never
   auto-closes. Check: grep `pluginMethod_CloseWindow` / `asc_pluginButtonClick` in the
   host's `editors/sdkjs/word/sdk-all.js` + web-apps shim (audit 9.4.0, 2026-09-28);
   `node --test tests/code.test.js` `close:` cases pin the fallback chain. The id
   resolution order this plugin walks: `plugin/scripts/code.js` `closeWindow` comment.
9. **`SearchAndReplace` caret semantics (learned from sdk-all.js, 2026-09-28).** The
   replacement string runs through `CSearchPatternEngine.Set`: `^t ^l ^p ^n ^m ^~ ^?
   ^# ^$` become field codes and there is NO `^^` escape (a lone trailing `^` is the
   only literal caret); one call replaces ALL occurrences. Plain-replace values
   therefore route carets through a private-use marker (`\uE0FF`) demoted in one final
   pass, and values already containing the marker are refused (`unsafe-value`).
   Decision rationale: docs/adr/0001.
10. **Iframe file input works in plugin windows** — `<input type=file>` + FileReader
   (OCR/Clipdrop plugins prove it) and heavy vendored libraries are fine (Tesseract,
   highlight.js precedent); the vendored PapaParse/SheetJS choice rests on this
   (api.onlyoffice.com/docs/plugins, checked 2026-09-28).

## Rejected approaches

The verdict table (no desktop mail-merge UI, no marketplace plugin, engine ships in
every build) has one home: README.md "Why a plugin". Evidence dates: catalog repo +
marketplace A-Z checked 2026-09; upstream issues re-checked 2026-09-28.
