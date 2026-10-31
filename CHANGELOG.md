# Changelog

## [5.0] — Current

### Changed

- **Structured phone/email data model** — phones and emails are stored
  internally as `{'value': ..., 'type': ...}` dicts instead of pre-built
  display strings (CONTEXT.md → D1). Effects:
  - **JSON / XML** — `phones`/`emails` are arrays of `{"value", "type"}`
    objects
  - **CSV** — `Phone{i}`/`Email{i}` columns gain paired `Phone{i}Type`/
    `Email{i}Type` columns
  - **VCF** — `format_vcf()` maps fields directly instead of re-parsing
    display strings
  - **SQL / SQLite, text, Markdown, HTML, DOCX** — unchanged in appearance
  - Text and VCF output verified byte-identical to v4; `--stats` counts
    identical on the 72-block test file
- **README restructured** — the three reference-heavy sections (Output
  Formats, Recognised Properties, Troubleshooting) are now collapsible,
  and a terminal-preview example and project icon were added

### Added

- **`--convert-to <version>`** — converts any input to a single merged,
  deduplicated `.vcf` at version `3.0` or `4.0` (`2.1` is not offered as a
  target — CONTEXT.md → D10). Mutually exclusive with `--merge`, `--split`,
  `--format` and `--preview`. Default output name gets a `_v30`/`_v40`
  suffix.
  - 4.0 output uses `tel:`/`mailto:` URIs for TEL/EMAIL; 3.0 does not
  - Yearless dates (`--MMDD`) are written as-is for 4.0. 3.0 has no valid
    form for them, so the field is skipped and counted in the terminal
    summary
  - `--merge` and `--convert-to` share one implementation
    (`_write_merged_vcf()`)
- **Nine utility flags:**
  - **`--dedup-only`** — duplicate report only; writes no file
  - **`--validate`** — per-contact validation report, grouped by contact;
    writes no file
  - **`--quiet`** — suppresses all non-error terminal output
  - **`--verbose`** — prints each skipped field as it is found and lists
    every duplicate pair by name
  - **`--dry-run`** — runs the full pipeline and prints the summary, but
    writes no file
  - **`--diff a.vcf b.vcf`** (D11) — field-level comparison; contacts are
    matched by name; reports added, removed and changed contacts
  - **`--update base.vcf new.vcf`** (D12) — merges `new.vcf` into
    `base.vcf`; `new.vcf` wins on identity fields, and fields missing from
    it are kept from `base.vcf`. Writes a `.vcf`
  - **`--template <file>`** (D13) — `{token}` placeholder output
    (`{name}`, `{phones}`, `{emails}`, ...). No loops or conditionals;
    unknown tokens are left literal so typos are visible. Implies
    `--format template`
  - **`--watch <folder>`** (D14) — foreground polling loop (5s) that
    converts new or modified `.vcf` files until `Ctrl+C`. No new
    dependencies
- `--quiet`, `--verbose` and `--dry-run` apply to the default conversion
  pipeline only; combining them with `--merge`, `--split` or
  `--convert-to` is an error
- New test files `test_diff_before.vcf` / `test_diff_after.vcf` covering
  every `--diff`/`--update` case (changed, unchanged, added, removed)

### Fixed

- **`--sort-by birthday` sorted alphabetically by month name** ("April"
  before "January"). Birthdays and anniversaries now also store a
  chronological raw key at parse time, used for sorting in both the normal
  and `--preview` pipelines (D16)
- **`format_ics()` could emit invalid events** for impossible dates such as
  April 31. These are now skipped and counted in the terminal summary
- **`format_vcf()` wrote `BDAY:April 12, 1985`** (the display string) in
  `--merge`/`--split` output instead of valid vCard syntax
  (`BDAY:19850412`). Present since v4
- **`--update` dropped the base file's birthday** when the new file had none
- **`--diff` and `--update` continued past a missing file** and reported a
  misleading result; either file missing is now a clear error
- **`--watch --format pdf`/`docx`** now checks for the optional dependency
  at startup, not after the first file arrives
- **`--quiet`/`--verbose`** were silent no-ops with `--merge`/`--split`/
  `--convert-to`; the combination is now an error
- **Misleading terminal wording** — the summary said "exported" even for
  `--stats` and `--dry-run`, and "review output file" even when no file
  was written. It now reads "kept" for `--stats` and "would be exported"
  for `--dry-run` (D15)
- `--diff`/`--update` summaries now include a `Malformed:` line when either
  file had malformed contacts
- `--help` now lists all 12 `--format` values and every output that
  `--no-photos` affects (JSON, HTML, XML, PDF, DOCX)
- Docstrings corrected to match behavior: `_load_contacts()` (callers and
  `None` contract), `run_paged_preview()` (the `d` detail control) and
  `identity_key()` (IMPP is intentionally not an identity field)
- Unused variables and unnecessary `f""` prefixes removed

---

## [4.0]

Renamed from vCard2text to vCardTra. Multi-format output and new field
parsing.

### Added

- Project renamed: `vCard2text.py` → `vCardTra.py`; all internal references,
  help text, and output headers updated
- `--format <name>` flag — selects output format; default `text` (unchanged
  behavior). Ten formats implemented: `text`, `json`, `csv`, `markdown`/`md`,
  `html`, `ics`, `sql`, `sqlite`, `xml`, `pdf`, `docx`
- **JSON** (`format_json()`, `contact_to_json_dict()`) — envelope structure
  `{"meta": {...}, "contacts": [...]}`; `meta` mirrors the text/VCF header
  rows (exported time, source(s), counts, versions, sort/filter/select state)
  as native types and snake_case keys. Photos included as base64 by default
- **CSV** (`format_csv()`) — wide table via stdlib `csv`; numbered columns
  for multi-value fields (`Phone1`, `Phone2`, ... `Email1`, ...), column
  count set by whichever contact in the run has the most values for that
  field. Photos skipped (no sane CSV representation)
- **Markdown** (`format_markdown()`) — single `.md` file, one `##` section
  per contact, bullet-list fields; for Notion/Obsidian/GitHub wikis
- **HTML** (`format_html()`) — single self-contained file, no external
  CSS/JS, photos embedded inline as base64 `data:` URIs. Also the shared
  rendering base for PDF output
- **ICS** (`format_ics()`, `parse_display_date()`) — birthdays/anniversaries
  as yearly-recurring `VEVENT`s. Reverse-parses `format_date()`'s
  human-readable output back into (year, month, day) since raw values
  aren't retained after parsing (see Design Decisions); only dates with
  day-level precision become events — month/year-only birthdays are skipped
  (reported to the terminal, not silently dropped). No-year dates use `1900`
  as a conventional placeholder and say so in the event summary
- **SQL** (`format_sql()`) — single wide `contacts` table, plain `.sql` text:
  one `CREATE TABLE IF NOT EXISTS` + one `INSERT` per contact
- **SQLite** (`write_sqlite()`) — same schema as SQL, written directly to a
  `.db` file via stdlib `sqlite3` — no server, no extra dependency
- **XML** (`format_xml()`) — same meta+contacts shape as JSON, built via
  manual string assembly with `xml.sax.saxutils.escape` for safe escaping
- **PDF** (`write_pdf()`) — renders the same HTML as `--format html` through
  `weasyprint` (optional dependency, lazy-imported; clear install-hint error
  if missing, checked upfront in `main()` before any file reading starts)
- **DOCX** (`write_docx()`) — one heading + field list per contact via
  `python-docx` (optional dependency, same lazy-import pattern), photo
  inserted inline as an image when present
- `--browser` flag — with `--preview`, opens a one-shot HTML snapshot in the
  system default browser (`run_browser_preview()`, stdlib `webbrowser` +
  `tempfile`) instead of the interactive terminal TUI/paged view. Reuses
  `format_html()` — the same renderer `--format html` writes to disk
- New fields parsed (previously silently dropped):
  - `TZ` → `timezone`, `GEO` → `geo`, `ROLE` → `role` — single values, empty-guarded
  - `IMPP` → `impp` list, multiple values, field-level deduped; scheme
    (`xmpp:`, `skype:`, `sip:`, etc.) kept since it identifies the protocol
  - `PHOTO` / `LOGO` — decoded from Base64 into `_photo_data`/`_logo_data`
    (bytes) + `_photo_type`/`_logo_type`; malformed Base64 reported in
    `[Skipped Fields]` instead of crashing
  - `extract_media_type()` — extracts image type (JPEG, PNG, ...) from
    PHOTO/LOGO field parameters for both v3.0+ `TYPE=` and v2.1 inline styles
- `--no-photos` flag — skips PHOTO/LOGO Base64 decoding entirely; also
  excludes them from JSON/HTML/XML/PDF/DOCX output
- `DISPLAYABLE_FIELDS` gains `role`, `timezone`, `geo`, `impp`
- PHOTO/LOGO removed from `SILENT_FIELDS` (now get real handling); SOUND/KEY
  remain silent — no practical use in any current or planned output format
- `FORMAT_DEPENDENCIES` — maps `pdf`/`docx` to their required package name;
  `main()` checks upfront via `importlib.util.find_spec()` so a missing
  optional dependency fails fast with an install hint, before any file I/O

### Changed

- `format_contact()` (text output) intentionally **unchanged** — new fields
  (role, timezone, geo, impp, photo, logo) appear only in JSON/CSV/Markdown/
  HTML/SQL/SQLite/XML/PDF/DOCX, not in text. Text output is byte-for-byte
  identical to v3
- `convert()`'s write step refactored into a format dispatcher — text
  rendering extracted into `_render_text()` (behavior unchanged) so it's
  called the same way as every other formatter, instead of being a special
  case; `sqlite`/`pdf`/`docx` write the file themselves (binary / library-
  owned save) rather than returning a string for `write_text()`

### Fixed (v4 deep-audit pass)

- `merge_contacts()` wasn't rescuing `role`/`timezone`/`geo`/`impp` from an
  exact duplicate — real data loss on merge, confirmed with a live test
  before fixing. Extended using the same pattern as the existing
  title/nickname/anniversary/gender/categories/label rescue list. Photo/logo
  deliberately left out — see Known limitations
- `run_paged_preview()`: removed a dead `detail = False` variable (assigned,
  never read) — pre-existing since v3
- `run_textual_preview()`: removed a dead `fuzzy_idx` computation shadowed by
  the actually-used `fuzzy_set` — pre-existing since v3
- `--format json --preview` was silently accepted with `--format` completely
  ignored — added to the existing mutual-exclusion checks alongside
  `--merge`/`--split`, with a clear error message
- **PDF crash on corrupted-but-plausible image data**: a photo with a valid
  image header but a truncated/corrupted data stream could crash the
  *entire* multi-contact PDF export via weasyprint, losing every contact,
  not just the one with the bad photo — reproduced live before fixing.
  `_validate_image_bytes()` (lazy Pillow import, already a weasyprint
  dependency) now checks each photo/logo before rendering; a contact with
  invalid image data gets rendered without that image rather than aborting
  the whole document. DOCX already handled this safely per-image via its
  existing try/except — only PDF needed the fix

### Known limitations (documented, not fixed)

- PHOTO/LOGO are decoded only when sent as inline `ENCODING=BASE64` data —
  vCard 4.0's `PHOTO:data:image/jpeg;base64,...` data-URI style and external
  `PHOTO:https://...` URI references are not yet extracted (silently ignored,
  same as pre-v4 behavior)
- `--no-photos` only wired into the normal `convert()` path; `--merge` and
  `--split` still decode PHOTO/LOGO unconditionally (harmless — `format_vcf()`
  doesn't use the decoded bytes — but wastes some time/memory on large files)
- `merge_contacts()` doesn't rescue photo/logo from a removed duplicate —
  deliberately, since "which photo wins" is a real design decision, not a
  mechanical extension like the other rescued fields
- CSV/SQL/SQLite/XML keep the D1 presentation-string format for phones/emails
  (`"Phone: +1-555-0101 (Work)"`) rather than structured columns — same
  deferred-to-v5 reasoning as JSON
- `--filter has=...` was not extended to cover the new v4 fields (role,
  timezone, geo, impp) — out of scope for this round, existing filter
  behavior unchanged

---

## [3.0]

### Added

- `--preview` — interactive preview mode; runs full pipeline then opens preview instead of writing a file
  - `run_textual_preview()` — two-panel Textual TUI; left list panel with `●` selected and `?` fuzzy markers; right detail panel; `/` live search; `s` cycles 4 sort modes; `space`/`a`/`n` selection; `e` exports to `preview_export.txt`; stats bar at bottom
  - `run_paged_preview()` — pure stdlib paged fallback; 8 per page; Enter/n/b/d/`/`/q controls
  - `run_preview()` — dispatcher; tries textual, catches `ImportError`, falls back silently
  - `--preview` is mutually exclusive with `--merge`, `--split`, `--stats`
- `--sort-by <field>` — sort contacts by `name`, `org`, `birthday`, or `created`; implies sorting, no need to also pass `--sort`
- `--reverse` — reverse sort order; unknowns always sort last regardless of direction; requires `--sort` or `--sort-by`
- `--limit <n>` — export only the first N contacts after all sorting and filtering
- `--filter <cond>` — filter contacts by field condition; repeatable, all conditions must match (AND logic)
  - Supported: `name=`, `org=`, `category=`, `has=phone/email/birthday/address/note/url`
  - Uses `partition('=')` so values containing `=` are handled correctly
  - All conditions validated before any filtering — fails fast with a clear message
- `--select <expr>` — keep only matched contacts; supports integers, ranges (`1-10`), `last-N`, and fnmatch name wildcards (`John*`); token types mix freely
- `--exclude <expr>` — remove matched contacts; same syntax as `--select`
- `--select` and `--exclude` are mutually exclusive — clean error if both given
- `--merge` flag — merge one or more VCF files into a single clean vCard 3.0 output file
  - Full duplicate detection runs before writing — exact duplicates merged, fuzzy duplicates warned
  - Smart default output name derived from input filenames
- `--split` flag — split contacts into individual per-contact `.vcf` files
  - Deduplication runs before splitting — output set is always clean
  - Zero-padded numbered filenames: `contact_01.vcf`, `contact_02.vcf`, etc.
  - Default output directory: `split_<input_stem>/`
- `--merge` and `--split` are mutually exclusive — clean error if both given
- `--encoding <codec>` flag — force a specific file encoding, bypassing auto-detection
- `format_vcf()` — new VCF 3.0 formatter; reverse-parses phone/email presentation strings, uses raw ISO timestamps, restores `X-` prefix on custom fields
- `read_file()` — smart encoding detection: charset-normalizer (optional) → `utf-8-sig` → `utf-8` → `utf-16` → `windows-1252` → `latin-1` → `latin-1+replace`
- `apply_filters()` — filter helper with full validation before execution
- `apply_selection()` — selection helper with clean range parsing via `re.match(r'^\d+-\d+$')`
- `_load_contacts()` — shared file-reading helper used by `convert_merge()` and `convert_split()`
- All `convert*` functions now propagate exit code 1 on failure
- Sort, Filter, Select, Exclude, Limit rows added to file header and terminal summary

### Changed

- `--sort` label in summary changed from `A to Z` to `name A → Z` (or `name Z → A` with `--reverse`)
- `-o` now accepts a directory path when used with `--split`
- Help text updated with all new flags and examples

---

## [2.0]

Complete rewrite and overhaul from v1.0 (original GitHub release).

### Fixed

- Lowercase property names not parsed — `.upper()` normalization added
- Group prefix not stripped — `item1.TEL` now correctly parsed as `TEL`
- Windows `\r\n` line endings not normalized before processing
- Quoted-Printable decoding missing for vCard 2.1 inline style (`NOTE;QUOTED-PRINTABLE:`)
- `--MMDD` birthday format (vCard 4.0 year-omitted) incorrectly rejected
- Multiple `ORG` values silently overwritten — now collected as a list
- Multiple `URL` values silently overwritten — now collected as a list
- Multiple `NOTE` values silently overwritten — now collected as a list
- `REV` field treated as unknown custom field — now shown as `Revised:`
- `key_part` not stripped before parsing
- `ADR` escape sequences (`\,`, `\n`, `\;`) not unescaped in address values
- `CATEGORIES` output raw — now cleaned and comma-space padded
- Phone and email lines were indented inconsistently — all fields now flush left
- Type label casing inconsistent — now always Title Case (`Work/Voice/Pref`)
- `item1.X-LABEL` incorrectly landed in custom fields — now routes to `Label:`
- `is_base64()` dead code removed
- `STANDARD_FIELDS` constant defined but never used — now wired into custom field logic
- Two near-identical converter functions merged into one
- Birthday and anniversary stored as raw vCard strings — now formatted (e.g. `April 12, 1985`)
- `--stats` mode ran a full sort step unnecessarily — skipped in stats-only mode now
- Exact duplicate removal discarded unique data from the removed contact — now merged
- `TITLE`, `NICKNAME`, `LABEL`, `GENDER` empty values stored — now guarded with `if value:`
- `NOTE` empty value stored empty string in notes list — now guarded
- `CATEGORIES` with only whitespace or commas stored empty string — now guarded
- `CREATED` and `REVISED`/`REV` empty values stored — now guarded
- Custom `X-` fields with empty values stored — now guarded
- `is_newer()` compared formatted display strings instead of raw ISO timestamps — now uses `_revised_raw`
- `merge_contacts()` note dedup missed same content without `[merged]` prefix — fixed
- All-empty-field contacts incorrectly counted as valid — caught by `has_displayable_content()`
- `-o` with no filename silently fell back to default output name — now errors clearly
- `merge_contacts()` revised comparison used formatted strings — now uses `_revised_raw`
- `Files` row shown twice in terminal output (once at top, once in summary table) — removed from table
- `file(s)`, `contact(s)`, `field(s)` bracket notation in terminal — proper singular/plural throughout
- `merge_contacts()` docstring said it merged `created` but implementation did not — now implemented
- Duplicate `--version` check in `main()` — dead code removed
- Single mode contact header missing trailing colon — now consistent (`⭐ Contact 1:`)
- `_created_raw` comment incorrectly referenced `is_newer()` — corrected to `merge_contacts`
- `exact_removed` comment said index was "approximate" — indices are exact (fuzzy pass never removes)
- 6-digit date format (`YYYYMM`) validator and formatter now consistent — formats as `April 1985`
- Slash-form date validator used loose `'/' in value` check — now requires exactly 3 parts

### Added

- Two-tier duplicate contact detection:
  - Exact duplicates (all identity fields match) — merged and removed, reported in warnings
  - Fuzzy duplicates (same name + shared phone or email) — kept, warned only
- Smart merge on exact duplicates — unique data rescued from removed contact:
  - `title`, `nickname`, `anniversary`, `gender`, `categories`, `label`, `urls` — taken from duplicate if missing in base
  - `notes` — appended with `[merged]` prefix; skipped if same content already exists
  - Custom `X-` fields — merged, base wins on conflict
  - `revised` — newer `_revised_raw` wins
  - `created` — earlier `_created_raw` wins (oldest origin date is the true one)
- REV-aware base selection — contact with newer `REV` timestamp becomes primary; older merged into it
- Warning output now shows which contact a duplicate merged into: `- John Smith → merged into Contact 3`
- Field-level deduplication — identical values within the same contact removed
- `--sort` flag — sort contacts A to Z, unknowns last
- `--stats` flag — print summary statistics only, no output file written
- `--version` flag — print `vCard2text 2.0` and exit
- `-o` with no filename now errors cleanly: `Error: -o flag requires a filename.`
- Auto-rename output if file already exists (`contacts_1.txt`, `contacts_2.txt`, ...)
- Smart output filename when no `-o` given (single / two / three-or-more files)
- Human-readable output file header with all relevant rows; missing rows suppressed
- `Files:` row in output file header for multi-file runs
- `⭐ Contact N:` decoration on each contact header
- Contact count line: `10 found, 1 malformed, 1 duplicate removed, 8 exported`
- `[Conversion Warnings]` section at bottom of output file
- Per-contact `[Skipped Fields]` section with `⚠` warnings for invalid fields
- `format_date` — human-readable dates for birthday/anniversary
- `format_datetime` — human-readable datetimes with time component
- Slash-form date support (`4/12/1985` → `April 12, 1985`; 2-digit years handled)
- 6-digit date support (`YYYYMM` → `April 1985`)
- `has_displayable_content()` — all-empty-field vCards counted as malformed
- `DISPLAYABLE_FIELDS` constant
- `_revised_raw` and `_created_raw` stored alongside formatted values for correct ISO comparison
- `safe_output_path()` — auto-rename logic
- `unescape()` — centralised vCard escape sequence handling
- `is_newer()` — REV-aware timestamp comparison using raw ISO strings
- `merge_contacts()` — field merge logic with correct revised/created handling
- Error handling on output file write — clean message on permission denied / disk full
- File encoding detection: UTF-8 with Latin-1 fallback
- vCard version detection per contact, aggregated in header
- `identity_key` tuple positions documented in code comments
- Proper singular/plural throughout all terminal output

### Changed

- `X-` vendor field prefix stripped from display name (`X-SKYPE` → `* Skype:`)
- Custom fields shown with `*` prefix
- `detect_encoding` handles both v2.1 and v3.0+ encoding styles
- `extract_type` returns consistent Title Case output
- `unfold` also normalizes `\r\n` before unfolding
- Section dividers unified via `DIVIDER`, `DIVIDER_HVY`, `DIVIDER_CTX` constants

---

## [1.0] — Original GitHub Release

Initial release. Basic vCard to text conversion.
Single file input, UTF-8/Latin-1 encoding support, wildcard file patterns.
Supports vCard 2.1, 3.0, 4.0 (partial).
