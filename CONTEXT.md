# Context

Mail merge plugin for ONLYOFFICE Desktop Editors. This file is the repo's vocabulary and
nothing else: one canonical term per concept, each entry saying what the term *is* and
pointing at the code symbol that defines it. Architecture lives in README.md "Development"
and the module headers; decisions live in `docs/adr/`.

## Language

- **token**: the literal `{{Name}}` placeholder text in the template document.
  _Avoid_: placeholder, tag. Home: `plugin/scripts/fieldmap.js` `TOKEN_RE`.
- **field name**: the normalized interior of a token — the case-insensitive key matched
  against data headers. _Avoid_: bare "field" for this. Home: `fieldmap.js`
  `normalizeName` / `matchFields`.
- **merge field**: a real Word `MERGEFIELD` (`ParaField(fieldtype_MERGEFIELD, [name])`)
  displayed as «Name», produced by wrapping a token. _Avoid_: field, token. Home:
  `plugin/scripts/commands.js` `wrapFields`.
- **header**: the row-0 cell naming one data column; matched against field names.
  _Avoid_: "field name" on the data side. Home: `plugin/scripts/dataparse.js` header.
- **record**: one data row (rows 1..n) — one recipient's values. _Avoid_: row, entry.
  Home: `dataparse.js` `toLoadMailMergeData`.
- **recipient**: the person one record produces output for (UI copy). _Avoid_: user.
  Home: `plugin/scripts/code.js` wizard strings.
- **data grid**: the one canonical parse shape `String[][]` (row 0 = headers) that
  `Api.LoadMailMergeData` consumes. Home: `dataparse.js` header.
- **wrap**: converting a token into a merge field via run surgery. Home: `fieldmap.js`
  `wrapPlan` + `commands.js` `wrapFields`.
- **plain replace**: substituting a token's text through `SearchAndReplace`, used for
  unmatched tokens. _Avoid_: text replacement. Home: `fieldmap.js` `plainReplacePlan`,
  `commands.js` `replacePlain`.
- **caret-marker**: the `\uE0FF` stand-in that carries a value's `^` through
  `SearchAndReplace` (which has no `^^` escape). Home: `commands.js` `CARET_MARKER`;
  decision: `docs/adr/0001`.
- **unsafe-value**: a data value already containing the caret-marker; refused, never
  merged. Home: `commands.js` `replacePlain` error taxonomy.
- **combined / per-record**: the two output modes — one document holding every selected
  record's copy, or one file per record. Home: `code.js` `state.mode`.
- **template snapshot**: the editor-side `ApiDocumentContent` capture that
  `restoreTemplate` puts back after every per-record save. Home: `commands.js`
  `snapshotTemplate` / `restoreTemplate`.
