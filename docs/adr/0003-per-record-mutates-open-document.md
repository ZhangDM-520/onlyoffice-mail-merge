# ADR 0003 — Per-record output mutates the open document

* **Status:** adopted, 2026-09-28.
* **Scope:** the per-record output loop in `plugin/scripts/code.js` `runMergePipeline`
  and the `snapshotTemplate` / `restoreTemplate` pair in `plugin/scripts/commands.js`.

## Context

No plugin API creates or opens a new document (`docs/MEMORY.md` fact 2), while
`Api.MailMerge(start,end)` replaces the open document body with the merged result for
the record range (`docs/MEMORY.md` fact 1). Producing one file per record therefore has
no clean workspace to merge into.

## Decision

Per-record output runs `merge(i, i)` → `GetFileToDownload` (the desktop opens one OS
save dialog per call) → restore the template snapshot → next record. The pipeline never
closes the window on error, and the snapshot restore is what puts the user's template
back after every save.

## Considered options

* **Document Builder batch mode** — out of scope: a different runtime with no plugin
  seam, and the wizard cannot drive it (rejected again 2026-09-28, see `docs/NOTE.md`).
* **Refuse per-record output entirely** — loses the one feature the online product also
  lacks (README.md "Why a plugin").
* **Wait for a new-document API** — does not exist in the plugin surface.

## Consequences

* The open document visibly mutates during a per-record run (restored after each save).
* Each record costs one save dialog — the wizard warns and points to combined mode for
  large runs (README.md "Usage").
* The snapshot must survive the whole run on the editor-page side: complex
  `ApiDocumentContent` objects cannot cross the `callCommand` boundary
  (`docs/MEMORY.md` fact 4).
