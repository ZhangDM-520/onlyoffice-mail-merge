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
6. Installed plugin inventory lives in both roots; the LaTeX math plugin
   (`~/Projects/onlyoffice-latex-math`, GUID `{5B4C1A72-...}`) is the reference sample for
   manifest layout, callCommand seam, icon slots (5 scales x 2 themes, no fallback) and
   test harness (`node --test`, fake-editor double).

## Rejected approaches

- Driving the online-only Mail Merge UI from desktop — no UI exists to drive.
- Waiting for upstream: DesktopEditors#733 / tracker 54752 open since 2021, no ETA.
- Existing marketplace plugin: none (checked catalog repo + marketplace A-Z, 2026-09).
