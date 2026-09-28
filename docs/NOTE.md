# Session notes — mail merge plugin (2026-09-28)

## Why this repo exists

Goal: mail merge in ONLYOFFICE Desktop Editors (template + recipient spreadsheet ->
personalized documents). Verdict from research:

- Desktop Editors has NO built-in mail merge ("available in the online version only";
  DesktopEditors#733 open since 2021, internal tracker 54752, no ETA).
- NO marketplace/catalog plugin does it (55 catalog plugins checked; nearest is *Curly*,
  form-based `{tags}` fill; *Send* emails the current doc via the desktop mail client).
- BUT the merge ENGINE ships in every build's public plugin API.

## SDK engine evidence (verified on host, onlyoffice-git 9.4.0.130)

`/opt/onlyoffice/desktopeditors/editors/sdkjs/word/sdk-all.js`, JSDoc `@typeofeditors ["CDE"]`:

| Method | Behaviour |
| :--- | :--- |
| `Api.LoadMailMergeData(String[][])` | row 0 = field names, rows 1..n = values -> `editor.asc_StartMailMergeByList(data)` |
| `Api.GetMailMergeReceptionsCount()` | record count |
| `Api.GetMailMergeTemplateDocContent()` | ApiDocumentContent snapshot (complex; cannot cross callCommand) |
| `Api.MailMerge(start, end)` | `Get_MailMergedDocument(start,end)` then `ReplaceDocumentContent` — replaces the OPEN body with merged output for the record range |
| `Api.ReplaceDocumentContent(content)` | replaces body from a snapshot |
| `ApiParagraph/ApiRun.WrapInMailMergeField()` | wraps text into `ParaField(AscWord.fieldtype_MERGEFIELD, [name])` with «» display |

Desktop UI (web-apps) exposes none of it — no Mailings tab, nothing in locales. The
100-recipient cap is the online UI's, not the engine's.

## Plugin API facts used (api.onlyoffice.com/docs/plugins)

- `Asc.plugin.callCommand(func, isClose, isCalc, cb)` — functions are stringified and lose
  closures; `Asc.scope` is the single payload slot; returns must be JSON-serializable.
  Since 7.1 `window`/`document`/`alert` are blocked inside callCommand code.
- Variation `type: "window"` + `size` + `buttons`; `Asc.plugin.button(id, windowId)`
  mandatory for close (-1 = close/x); `CloseWindow`/`ResizeWindow` via `executeMethod`.
- `<input type=file>` + FileReader work in the iframe (OCR/Clipdrop plugins prove it);
  heavy vendored libs fine (Tesseract, highlight.js precedent).
- `Asc.plugin.executeMethod("GetFileToDownload", ["docx"], cb)` — on desktop the OS save
  dialog opens per call. NO create-new-document API exists (Document-API method list);
  per-record output therefore = merge -> save -> `ReplaceDocumentContent(snapshot)` -> next.

## Pitfalls carried over from onlyoffice-latex-math

- Install dir must be the braced GUID (`asc.` prefix only in config.json); `cp -a plugin/.`,
  delete dest on upgrade; clear editor `data/cache/{Cache,Code Cache}` after replacing files.
- One callCommand in flight (shared Asc.scope swaps payloads); serialize the queue.
- Return JSON strings only — `Asc.checkReturnCommand` drops complex objects.
- Batch document edits into one history point (`CreateNewHistoryPoint`) per merge.

## Decisions

- Both placeholder styles: wrap `{{Name}}` (and bare selection) into real MERGEFIELD,
  plain `SearchAndReplace` fallback for unmatched tokens.
- Both outputs; default single combined document (per-record = N save dialogs).
- CSV + XLSX + JSON data sources, parsed in-iframe (PapaParse / SheetJS vendored).
- Out of scope: email output, marketplace packaging, Document Builder batch mode.

## Wizard UI (code.js / styles.css / icons) — 2026-09-28

- `plugin/scripts/code.js` is the composition root: 3-step wizard in `#mm-app`, sequential
  `run()` pipeline (loadMergeData -> wrapFields(matched) -> snapshotTemplate -> mergeRange ->
  replacePlain -> GetFileToDownload; per-record loops merge(i,i) -> save -> restoreTemplate).
  Cancel/X = button ids 0/-1 -> `CloseWindow` with the `?windowID=` URL param (SDK publishes
  `Asc.plugin.windowID` late, from an XHR callback); the window closes only after a successful run.
- Known limitation: plain-replace values cannot vary across copies of a *combined* document
  (`MailMerge` replicates the open body), so combined + wrap-off blanks matched tokens - the
  Output step warns and points to per-record mode. Gapped ranges ("1,3") are refused in combined
  mode (one `MailMerge(start,end)` span must be contiguous).
- Icons: `tools/make-icons.py` renders Tabler `mail-fast` (noctalia font, U+F069) on the #2A5DB0
  store tile at 5 scales x (light|dark) + store set; `--check` = "16 files ok".
- Tests: `tests/code.test.js` (fake DOM + fake contract seams, 23 tests) and
  `tests/icons.test.js` (PNG pixel checks, 4 tests); `node --test tests/` green.
