# ADR 0001 — Caret-safe plain replacement

* **Status:** adopted, 2026-09-28.
* **Scope:** caret handling in `plugin/scripts/commands.js` `replacePlain` (the
  `CARET_MARKER` scheme) and everything the rest of the repo may say about it.

## Context

`ApiDocument.SearchAndReplace` runs its replacement string through
`CSearchPatternEngine.Set`, where `^t ^l ^p ^n ^m ^~ ^? ^# ^$` become field codes. There
is **no** `^^` escape; the only way to insert a literal caret is a lone trailing `^`.
Data values may legitimately contain carets (`a^p` must survive as text, not become a
paragraph mark), and one call replaces all occurrences — so the value text cannot pass
through the engine unchanged. Behaviour facts: `docs/MEMORY.md` fact 9.

## Decision

Every caret in a data value travels through the replacement as the private-use marker
`\uE0FF` and is demoted back to `^` in one final `SearchAndReplace` pass after all real
replacements. A value that already contains the marker is refused outright
(`unsafe-value`) instead of being merged.

## Considered options

* **Escape the carets** — impossible: the pattern engine offers no escape sequence.
* **Strip carets from values** — silent data loss in user data; unacceptable.
* **Use the trailing-`^` literal-caret rule** — a literal caret is produced only at the
  very end of the replacement string; mid-string carets are unrepresentable this way.
* **Chosen: marker + demote pass** — lossless for every position of `^`, at the cost of
  one extra engine call and a theoretical refusal for values containing `\uE0FF`.

## Consequences

* Values with carets survive as text; the refusal taxonomy grows `unsafe-value`.
* The demotion pass is one extra `SearchAndReplace` call per `replacePlain` run.
* A value containing the private-use marker can never be merged — by design, visible as
  an error entry rather than corrupted output.
