# Mail merge — data-driven merge for ONLYOFFICE Desktop Editors

`https://github.com/ZhangDM-520/onlyoffice-mail-merge`

An ONLYOFFICE **Document Editor** plugin that merges a template document with recipients
from a **CSV / XLSX / JSON** data source: it wraps `{{Field}}` tokens into real `MERGEFIELD`
fields (displayed as «Field») and generates either **one combined document** or **one file
per recipient**.

## Why a plugin

ONLYOFFICE has mail merge, but not where it is needed:

| Capability | Where it lives |
| :--- | :--- |
| Mail merge UI (Collaboration tab) | **online ONLYOFFICE Docs only** — help: "this option is available in the online version only" |
| Mail merge in Desktop Editors | **not exposed** — [DesktopEditors#733](https://github.com/ONLYOFFICE/DesktopEditors/issues/733) open since 2021, no ETA |
| Mail-merge plugin in the marketplace | **none** exists (nearest: *Curly*, form-based `{tags}` fill) |
| Merge engine (`Api.LoadMailMergeData`, `Api.MailMerge`, `WrapInMailMergeField`) | **in every build**, including Desktop — just unreachable from the UI |
| One output file per recipient | even the online product lacks it ([DocumentServer#2218](https://github.com/ONLYOFFICE/DocumentServer/issues/2218)) |

The plugin is the thin layer that exposes that engine already shipped in
`editors/sdkjs/word/sdk-all.js` (behaviour: `docs/MEMORY.md` facts 1-2).

## Usage

1. Open a template document containing `{{Name}}`, `{{City}}` … tokens (or select text and
   let the plugin wrap it).
2. Plugins tab → **Mail Merge** → pick your CSV/XLSX/JSON data source (first row = field
   names, one row per recipient; shape contract: `plugin/scripts/dataparse.js`).
3. Review the field mapping (matched tokens become real «Name» merge fields; unmatched
   tokens fall back to plain text replacement).
4. Choose output: single combined document (default) or one file per recipient, a record
   range, and docx/pdf. **Merge** runs the pipeline.

Per-recipient mode opens the standard save dialog once per record (there is no
create-new-document API in the plugin surface) — use the combined mode for large runs.

## Limitations

- Combined output blanks plain-replaced tokens: one combined document holds every
  record's copy of the template, so a plain replacement cannot vary per recipient.
  With wrapping off that includes matched tokens. The Output step warns; wrap the
  tokens or use one file per recipient.
- Gapped record ranges ("1,3") are refused in combined mode — one merge span must be
  contiguous.
- The 100-recipient cap belongs to the online mail-merge UI; the engine has none.

## Install

The plugin GUID is `{2A0D08A5-D356-4057-A366-8A2AA579B4D7}`.

<!-- install-recipe:start -->
```bash
G='{2A0D08A5-D356-4057-A366-8A2AA579B4D7}'

# system-wide (needs root)
sudo rm -rf "/opt/onlyoffice/desktopeditors/editors/sdkjs-plugins/$G"
sudo cp -a plugin/. "/opt/onlyoffice/desktopeditors/editors/sdkjs-plugins/$G/"

# per-user
rm -rf "$HOME/.local/share/onlyoffice/desktopeditors/sdkjs-plugins/$G"
cp -a plugin/. "$HOME/.local/share/onlyoffice/desktopeditors/sdkjs-plugins/$G/"

# after replacing files: quit the editor, then clear its renderer cache
rm -rf ~/.local/share/onlyoffice/desktopeditors/data/cache/{Cache,Code\ Cache}
```
<!-- install-recipe:end -->

Host install policy has one home: the
[latex-math README § Install](https://github.com/ZhangDM-520/onlyoffice-latex-math#install) —
braced-directory rules, `cp -a` semantics, delete-dest upgrades, the real-`v1/`-loader rule
and verification. This section is only this plugin's recipe. Check the two recipes stay in
sync (run from `~/Projects`):

```bash
# The "<![-]-" spelling keeps the patterns from matching their own text.
diff <(sed -n '/<![-]- install-recipe/,/<![-]- install-recipe:end -->/p' onlyoffice-latex-math/README.md | sed 's/{5B4C1A72-[^}]*}/{GUID}/g') \
     <(sed -n '/<![-]- install-recipe/,/<![-]- install-recipe:end -->/p' onlyoffice-mail-merge/README.md | sed 's/{2A0D08A5-[^}]*}/{GUID}/g')
```

## Development

```
plugin/
  config.json          manifest (window variation, GUID, icons)
  index.html           iframe page: vendor libs + scripts
  scripts/dataparse.js pure: CSV/XLSX-grid/JSON -> String[][] (grid shape: its header)
  scripts/fieldmap.js  pure: {{Token}} scanning, header matching, replace plans
  scripts/commands.js  editor-page seam (callCommand bodies + serialized dispatch)
  scripts/code.js      wizard UI + orchestration
  vendor/              PapaParse + SheetJS (offline, no CDN)
tests/                 node --test suites (pure modules, command seam, icons)
tools/make-icons.py    generates theme/scale icon PNGs
```

Editor-page commands follow one discipline (`docs/MEMORY.md` fact 4), shared with
[onlyoffice-latex-math](https://github.com/ZhangDM-520/onlyoffice-latex-math).
SDK mail-merge engine evidence: `docs/MEMORY.md` facts 1-2.

## Docs map

One fact, one home: each fact is stated in exactly one place below; every other mention
is a pointer. Code comments state *why*, never *what*.

| Home | Owns |
| :--- | :--- |
| `CONTEXT.md` | vocabulary — one canonical term per concept |
| `docs/adr/NNNN-*.md` | one decision each (what was chosen and why) |
| `docs/MEMORY.md` | durable host/SDK facts + how to re-verify them |
| `docs/NOTE.md` | session journal — dated findings only |
| `plugin/scripts/*.js` header | that module's contract |
| `tests/*.test.js` header | that suite's run line |
| `plugin/config.json` | manifest truth (GUID, variation, locales) |

How to check:

```bash
node --test tests/                          # one suite: node --test tests/<name>.test.js
python3 tools/make-icons.py --check
```
