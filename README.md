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

The plugin is the thin layer that exposes the engine already shipped in
`editors/sdkjs/word/sdk-all.js`: parse the data source in the iframe, load it with
`Api.LoadMailMergeData`, and drive `Api.MailMerge` record by record.

## Usage

1. Open a template document containing `{{Name}}`, `{{City}}` … tokens (or select text and
   let the plugin wrap it).
2. Plugins tab → **Mail Merge** → pick your CSV/XLSX/JSON data source (first row = field
   names, one row per recipient).
3. Review the field mapping (matched tokens become real «Name» merge fields; unmatched
   tokens fall back to plain text replacement).
4. Choose output: single combined document (default) or one file per recipient, a record
   range, and docx/pdf. **Merge** runs the pipeline.

Per-recipient mode opens the standard save dialog once per record (there is no
create-new-document API in the plugin surface) — use the combined mode for large runs.

## Install

The plugin GUID is `{2A0D08A5-D356-4057-A366-8A2AA579B4D7}`. **Keep the braces** — the
directory the app enumerates is the braced one.

```bash
G='{2A0D08A5-D356-4057-A366-8A2AA579B4D7}'

# per-user (merged with the system root by GUID + version)
rm -rf "$HOME/.local/share/onlyoffice/desktopeditors/sdkjs-plugins/$G"
cp -a plugin/. "$HOME/.local/share/onlyoffice/desktopeditors/sdkjs-plugins/$G/"

# or system-wide (needs root)
sudo rm -rf "/opt/onlyoffice/desktopeditors/editors/sdkjs-plugins/$G"
sudo cp -a plugin/. "/opt/onlyoffice/desktopeditors/editors/sdkjs-plugins/$G/"
```

`cp -a plugin/. <dest>` — *not* `cp -r plugin <dest>`. Remove the destination first on
upgrade (`rsync -a --delete plugin/ <dest>/` also works). After replacing files, quit the
editor and clear its renderer cache:

```bash
rm -rf ~/.local/share/onlyoffice/desktopeditors/data/cache/{Cache,Code\ Cache}
```

**Per-user installs need a real `v1/` loader.** Plugin frames resolve `../v1/plugins.js`
against the sdkjs-plugins root *they live in*. On some installs the per-user root
(`~/.local/share/.../sdkjs-plugins/v1/`) ships only zero-byte stubs, so a per-user-only
plugin enumerates but its API is dead. Fix it once by copying the real loaders from the
system root (or install the plugin system-wide like the bundled plugins):

```bash
cp -p /opt/onlyoffice/desktopeditors/editors/sdkjs-plugins/v1/{plugins.js,plugins-ui.js,plugins.css} \
      ~/.local/share/onlyoffice/desktopeditors/sdkjs-plugins/v1/
```

## Development

```
plugin/
  config.json          manifest (window variation, GUID, icons)
  index.html           iframe page: vendor libs + scripts
  scripts/dataparse.js pure: CSV/XLSX-grid/JSON -> String[][] (row 0 = field names)
  scripts/fieldmap.js  pure: {{Token}} scanning, header matching, replace plans
  scripts/commands.js  editor-page seam (callCommand bodies + serialized dispatch)
  scripts/code.js      wizard UI + orchestration
  vendor/              PapaParse + SheetJS (offline, no CDN)
tests/                 node --test suites (pure modules, command seam, icons)
tools/make-icons.py    generates theme/scale icon PNGs
```

Tests: `node --test tests/`. Icons: `python3 tools/make-icons.py --check`.

The command seam follows the same discipline as
[onlyoffice-latex-math](https://github.com/ZhangDM-520/onlyoffice-latex-math): commands are
stringified into the editor page, return JSON strings only, run one-at-a-time over
`Asc.scope`, and time out against a 30 s backstop. See `docs/NOTE.md` for the SDK
mail-merge engine evidence.
