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
3. Plugin install: system `editors/sdkjs-plugins/` + user
   `~/.local/share/onlyoffice/desktopeditors/sdkjs-plugins/`, merged by GUID+version;
   directory name = braced GUID; `config.json` guid = `asc.{GUID}`. Upgrade = remove dest
   first (`cp -a` only adds). Clear `data/cache/{Cache,Code Cache}` after file replacement.
   **Per-user root must have a real `v1/` loader sibling** — frames resolve `../v1/plugins.js`
   from their own root; this host's user root had zero-byte stubs (plugins enumerate but
   `Asc.plugin` is undefined → silently inert UI). Keep the `/opt` v1 trio copied into the
   user root, or install plugins system-wide.
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
8. **Closing a plugin window (verified in the 9.4.0 host, 2026-09-28).**
   `executeMethod("CloseWindow", [id])` only closes windows registered via `ShowWindow` —
   it silently no-ops on any other id (including `[]`/undefined) and never throws. The MAIN
   window-type variation gets NO windowId with its clicks (web-apps calls
   `asc_pluginButtonClick(id, guid)` 2-arg), so the ONLY working close is
   `Asc.plugin.executeCommand("close", "")`. The window X/ESC routes to
   `Asc.plugin.button(-1, …)`; once the plugin defines that hook, the host never
   auto-closes. Never treat a "successful" `CloseWindow` as evidence the window closed.

## Rejected approaches

- Driving the online-only Mail Merge UI from desktop — no UI exists to drive.
- Waiting for upstream: DesktopEditors#733 / tracker 54752 open since 2021, no ETA.
- Existing marketplace plugin: none (checked catalog repo + marketplace A-Z, 2026-09).
