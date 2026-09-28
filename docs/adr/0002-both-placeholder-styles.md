# ADR 0002 — Both placeholder styles, combined default

* **Status:** adopted, 2026-09-28.
* **Scope:** the token-handling strategy (`wrap` vs `plain replace`) and the default
  output mode of the wizard (`plugin/scripts/code.js`, `plugin/scripts/fieldmap.js`).

## Context

Templates carry `{{Name}}` tokens matched against data headers; some match, some do not.
Wrapping a matched token into a real merge field gives «Name» display and per-record
values through the mail-merge engine; plain replace handles everything else — including
bare-selection wrapping, where the user has no token syntax at all.

## Decision

Matched tokens are wrapped into real merge fields; unmatched tokens fall back to plain
replace (togglable per token in the Fields step). Neither path alone covers the
template space. Default output is one combined document; per-record output is opt-in,
because it opens one save dialog per record (`docs/MEMORY.md` fact 2).

## Considered options

* **Wrap only** — unmatched tokens would remain as literal `{{…}}` text in the output.
* **Plain replace only** — no merge fields at all, and plain values cannot vary across
  the copies of a combined document (see below), so large merges would be impossible
  without per-record mode.
* **Per-record default** — one OS save dialog per recipient punishes the common case
  (one combined document).

## Consequences

* Combined output blanks plain-replaced tokens: one combined document holds every
  record's copy of the template, so a plain replacement cannot vary per recipient; with
  wrapping off that includes matched tokens. The Output step warns (user contract:
  README.md "Limitations").
* Gapped record ranges ("1,3") are refused in combined mode — one
  `Api.MailMerge(start,end)` span must be contiguous.
* Record-range semantics and the wrap-off warning are code contracts; user-visible
  phrasing lives in README.md "Limitations" only.
