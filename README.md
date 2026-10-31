<p align="center">
  <img src="icon.png" width="120" alt="vCardTra">
</p>

# vCardTra

A Python script that transforms vCard (.vcf) files into text, JSON, CSV, Markdown, HTML, ICS, SQL, SQLite, XML, PDF, or DOCX. Works with exports from iPhone, Android, Google Contacts, Apple Contacts, Outlook, WhatsApp, and most other apps that export contacts as `.vcf`.

> Renamed from vCard2text at v4.0 — same script, broader output support.

---

## Features

- Parse single or multiple vCard files
- Support for vCard versions 2.1, 3.0, and 4.0
- Extract all common contact fields (name, phone, email, address, birthday, and more)
- Output as text, JSON, CSV, Markdown, HTML, ICS (calendar), SQL, SQLite, XML, PDF, or DOCX
- Handle multiple contacts per file
- Wildcard pattern support (`*.vcf`)
- Detect and auto-merge exact duplicate contacts
- Flag possible duplicate contacts for your review
- Sort contacts A to Z, by org, birthday, or creation date
- Filter contacts by field, value, or presence
- Select or exclude specific contacts by number, range, or name pattern
- Interactive preview — TUI with live search and selective export, or a one-shot browser view
- Merge multiple VCF files into one clean VCF
- Split contacts into individual VCF files
- Clean, readable text output with a summary header
- Smart file encoding detection with optional manual override
- Quoted-Printable decoding (common in older Android/Nokia exports)
- Track source file for each contact when merging multiple files
- Never modifies your original `.vcf` files

---

## Requirements

- Python 3.8 or higher
- No required external dependencies (uses only standard library)

**Optional enhancements** — install for a better experience, but the script works without them:
- `pip install textual` — enables the full two-panel TUI for `--preview` (paged fallback is used if not installed)
- `pip install charset-normalizer` — enables automatic detection of CJK encodings like Shift-JIS and GB2312 (sequence fallback is used if not installed)
- `pip install weasyprint` — required only for `--format pdf`. Not needed for anything else
- `pip install python-docx` — required only for `--format docx`. Not needed for anything else

---

## Installation

1. Download `vCardTra.py`
2. No pip install needed

Optionally make it executable on Unix/macOS:

```bash
chmod +x vCardTra.py
```

---

## Usage

```bash
python vCardTra.py <file.vcf> [options]
```

### Options

| Option | Description |
|---|---|
| `<file.vcf>` | Input vCard file(s). Supports multiple files and wildcard patterns |
| `-o <output>` | Custom output filename or directory (optional) |
| `--format <name>` | Output format: `text` (default), `json`, `csv`, `markdown`/`md`, `html`, `ics`, `sql`, `sqlite`, `xml`, `pdf`, `docx` |
| `--no-photos` | Skip decoding PHOTO/LOGO fields — faster on large files, omits them from formats that would embed them |
| `--sort` | Sort contacts A to Z by name |
| `--sort-by <field>` | Sort by field: `name`, `org`, `birthday`, `created` |
| `--reverse` | Reverse sort order (requires `--sort` or `--sort-by`) |
| `--limit <n>` | Export only the first N contacts after sorting and filtering |
| `--filter <cond>` | Filter contacts — repeatable, all conditions must match (AND logic) |
| `--select <expr>` | Keep only matched contacts (index, range, last-N, or name wildcard) |
| `--exclude <expr>` | Remove matched contacts (same syntax as `--select`) |
| `--stats` | Print summary statistics only, no output file written |
| `--preview` | Interactive preview — textual TUI if installed, paged terminal fallback otherwise |
| `--browser` | With `--preview`: one-shot HTML view in your browser instead of the terminal UI |
| `--merge` | Merge all input VCFs into one clean `.vcf` file |
| `--split` | Split contacts into individual `.vcf` files, one per contact |
| `--convert-to <version>` | Convert to a target vCard version — `3.0` or `4.0` — and write one merged `.vcf` |
| `--dedup-only` | Scan for duplicates and print a report — no output file written |
| `--validate` | Deep per-contact validation report — no output file written |
| `--diff <file1> <file2>` | Compare two exports: added/removed/changed contacts, field-level detail |
| `--update <base> <new>` | Absorb `new.vcf` into `base.vcf` — `new.vcf` wins conflicts, non-identity fields rescued from `base.vcf` |
| `--template <file>` | Render output using your own `{token}` placeholder template |
| `--watch <folder>` | Poll a folder and auto-convert new/changed `.vcf` files until `Ctrl+C` |
| `--quiet` | Suppress terminal output except errors |
| `--dry-run` | Run the full pipeline and print the summary, but write nothing |
| `--verbose` | Print every field-level parsing/dedup decision as it happens |
| `--encoding <codec>` | Force a specific file encoding (e.g. `shift-jis`, `gb2312`) |
| `--version` | Show version and exit |
| `-h, --help` | Show help |

### Examples

Convert a single file (creates `contacts.txt`):
```bash
python vCardTra.py contacts.vcf
```

Export as JSON instead of text:
```bash
python vCardTra.py contacts.vcf --format json
```

Export to a spreadsheet-ready CSV:
```bash
python vCardTra.py contacts.vcf --format csv
```

Export as a calendar of birthdays and anniversaries:
```bash
python vCardTra.py contacts.vcf --format ics
```

Generate a printable PDF or a Word document:
```bash
python vCardTra.py contacts.vcf --format pdf
python vCardTra.py contacts.vcf --format docx
```

JSON without embedded photo data (faster, smaller file):
```bash
python vCardTra.py contacts.vcf --format json --no-photos
```

Preview contacts in your browser instead of the terminal:
```bash
python vCardTra.py contacts.vcf --preview --browser
```

Sort alphabetically with a custom output name:
```bash
python vCardTra.py contacts.vcf --sort -o my_contacts.txt
```

Sort by birthday, reverse order:
```bash
python vCardTra.py contacts.vcf --sort-by birthday --reverse
```

Export only the first 10 contacts after sorting:
```bash
python vCardTra.py contacts.vcf --sort --limit 10
```

Filter to contacts with a phone number at a specific org:
```bash
python vCardTra.py contacts.vcf --filter org=Acme --filter has=phone
```

Select contacts by number, range, or name wildcard:
```bash
python vCardTra.py contacts.vcf --select 1-10,15
python vCardTra.py contacts.vcf --sort --select "John*"
python vCardTra.py contacts.vcf --select last-20
```

Exclude specific contacts:
```bash
python vCardTra.py contacts.vcf --exclude 1,5,9
```

Preview contacts interactively:
```bash
python vCardTra.py contacts.vcf --preview
python vCardTra.py contacts.vcf --sort --filter has=phone --preview
```

Quick stats without writing a file:
```bash
python vCardTra.py contacts.vcf --stats
```

Merge multiple files into one clean VCF:
```bash
python vCardTra.py iphone.vcf google.vcf --merge -o all_contacts.vcf
```

Split into individual per-contact files:
```bash
python vCardTra.py contacts.vcf --split -o my_split/
```

Convert to a different vCard version:
```bash
python vCardTra.py old_contacts.vcf --convert-to 4.0 -o modern.vcf
```

Force a specific file encoding:
```bash
python vCardTra.py contacts.vcf --encoding shift-jis
```

Merge all `.vcf` files in a folder using wildcard:
```bash
python vCardTra.py "*.vcf" -o all_contacts.txt
```

Script and files in different folders:
```bash
python /tools/vCardTra.py /data/contacts.vcf -o /output/result.txt
```

> **Windows users:** Always wrap wildcard patterns in quotes: `"*.vcf"`

---

## Output

### Terminal

When you run the script, the terminal shows a summary of what happened:

```
vCardTra — 1 file to process

  Reading contacts.vcf ... 10 contacts, 1 malformed

  ────────────────────────────────────────────────────────────
  Contacts    :  10 found, 1 malformed, 1 duplicate removed, 8 exported
  vCard       :  2.1, 3.0
  Skipped     :  3 fields across 2 contacts
  Warnings    :  1 duplicate warning — review output file
  Sorted      :  name A → Z
  Output      :  contacts.txt
  ────────────────────────────────────────────────────────────
```

Lines only appear when relevant. A clean run shows only Contacts and Output.

### File header

Every output file starts with a summary:

```
Generated by vCardTra
────────────────────────────────────────────────────────────
  Exported          :  May 29, 2026, 14:32
  Source            :  contacts.vcf
  Contacts          :  10 found, 1 malformed, 1 duplicate removed, 8 exported
  vCard versions    :  2.1, 3.0
  Skipped fields    :  3 fields across 2 contacts
  Duplicate warnings:  1 (review in warnings section)
  Sorted            :  name A → Z
────────────────────────────────────────────────────────────
```

For multi-file runs, a `Files` row is also shown. Lines only appear when relevant — a clean file shows only Exported, Source, and Contacts.

### Contact format

```
⭐ Contact 1:
------------------------------------------------------------
Name: John Smith
Organization: Acme Corp
Title: Software Engineer
Phone: +1-555-0123 (Mobile)
Phone: +1-555-0199 (Work)
Email: john.smith@acme.com (Work/Internet)
Address: 123 Main St, San Francisco, CA, 94105, USA (Work)
Website: https://johnsmith.com
Birthday: April 12, 1985
Note: Met at conference 2023.
Nickname: Johnny
Anniversary: June 15, 2010
Gender: M
Categories: Work, VIP
Revised: January 1, 2024, 12:00 UTC
* Skype: john.smith.skype
```

Custom or vendor-specific fields (e.g. `X-SKYPE`) appear with a `*` prefix, with the `X-` stripped.

### Output file behaviour

- When no `-o` is given, the output name is derived from the input:
  - Single file: `contacts.vcf` → `contacts.txt`
  - Two files: `work.vcf + personal.vcf` → `work_personal.txt`
  - Three or more files: `work.vcf + ...` → `work_and_2_more.txt`
- If the output file already exists, it is never overwritten. The script saves as `contacts_1.txt`, `contacts_2.txt`, etc. automatically.

### --merge output

Produces a single clean `.vcf` file in vCard 3.0 format with all contacts deduplicated. Default output name:
- Single input: `contacts_merged.vcf`
- Two inputs: `work_personal.vcf`
- Three or more: `work_and_2_more.vcf`

### --split output

Splits contacts into individual `.vcf` files written into a directory. Files are zero-padded and numbered: `contact_01.vcf`, `contact_02.vcf`, etc. Deduplication runs before splitting so the output set is clean. Default output directory: `split_<input_stem>/` (e.g. `split_contacts/`).

### --convert-to output

Converts input at any vCard version to a target version and writes one clean, deduplicated `.vcf` file — like `--merge`, but at your chosen version instead of always 3.0. Supported targets: `3.0` and `4.0`.

```bash
python vCardTra.py contacts.vcf --convert-to 4.0 -o converted.vcf
```

Default output name follows the `--merge` pattern with a version suffix: `contacts_v40.vcf` (single input), `work_personal_v40.vcf` (two inputs), `work_and_2_more_v40.vcf` (three or more).

What actually changes between the two targets:
- **TEL/EMAIL** — a `4.0` target adds `tel:`/`mailto:` URI prefixes (`TEL:tel:+1-555-0101`); a `3.0` target doesn't
- **BDAY/ANNIVERSARY** — a birthday with no year (e.g. one originally parsed from a 4.0-source `--MMDD` value) converts to `--0704`-style syntax for a `4.0` target, since 4.0 supports year-less dates. A `3.0` target has no valid way to represent a yearless date, so that field is skipped for that contact — reported in the terminal summary (`Dates: N fields skipped...`) rather than silently dropped or written incorrectly

`2.1` is not offered as a `--convert-to` target. A correct 2.1 writer needs its own serialization (bare type params, `QUOTED-PRINTABLE` encoding for non-ASCII text), and a subtle encoding bug would garble accented names silently on exactly the old devices that need 2.1. 2.1 files are still fully supported as *input*.

### Utility flags

Terminal-report flags — no output file is written, these just print to the terminal:

```bash
python vCardTra.py contacts.vcf --dedup-only     # duplicate scan only
python vCardTra.py contacts.vcf --validate       # per-contact field validation report
python vCardTra.py contacts.vcf --dry-run        # run the full pipeline, print the summary, write nothing
python vCardTra.py old.vcf new.vcf --diff        # added/removed/changed contacts, field-level detail
```

`--diff` matches contacts between the two files by name (not the identity fields used elsewhere for duplicate detection) — a changed phone number for the same person is reported as one *changed* contact, not one removed plus a different one added.

**`--update <base.vcf> <new.vcf>`** absorbs `new.vcf` into `base.vcf` and writes a merged `.vcf` file:

```bash
python vCardTra.py --update contacts_old.vcf contacts_export.vcf -o updated.vcf
```

`new.vcf` wins unconditionally on phones/emails/addresses/organizations/birthday for any contact matched by name in both files — it's treated as the fresher, authoritative export. Anything `new.vcf` is missing for a matched contact (title, notes, custom fields, etc.) is rescued from `base.vcf` rather than lost. Contacts only in `base.vcf` are kept as-is; contacts only in `new.vcf` are added.

**`--template <file>`** renders your own flat text format using `{token}` placeholders instead of the built-in text/Markdown/etc. layouts:

```bash
python vCardTra.py contacts.vcf --template myformat.txt -o out.txt
```

Example template file:
```
Name: {name}
Company: {organizations}
Phone: {phones}
Email: {emails}
---
```
Recognised tokens: `name`, `organizations`, `title`, `phones`, `emails`, `addresses`, `urls`, `birthday`, `notes`, `nickname`, `label`, `anniversary`, `gender`, `categories`, `created`, `revised`, `source`. Multi-value fields (phones, emails, organizations, addresses, urls, notes) are joined with `, ` into a single token — there's no loop syntax for iterating them individually. A typo'd token (e.g. `{name}` misspelled as `{nam}`) is left in the output as a literal `{nam}` rather than silently becoming blank, so mistakes are easy to spot. Sort/filter/select/limit all work normally with `--template`, the same as any other `--format`.

**`--watch <folder>`** polls a folder every 5 seconds and auto-converts any new or modified `.vcf` file it finds, using the normal conversion pipeline with default settings (add `--format` to change the output format). Runs in the foreground until you press `Ctrl+C` — it's not a background service.

```bash
python vCardTra.py --watch ./incoming
python vCardTra.py --watch ./incoming --format json
```

**`--quiet`**, **`--verbose`**: opposite ends of the terminal-output dial. `--quiet` suppresses everything except errors; `--verbose` prints each field-level parsing decision (skipped fields, as they're found) plus every individual duplicate merge/warning instead of just the summary counts. Both currently apply to the default conversion pipeline only — combining either with `--merge`, `--split`, or `--convert-to` is an error rather than a silent no-op.

<details>
<summary><strong>Output Formats</strong> — click to expand full format-by-format reference (json, csv, markdown, html, ics, sql, sqlite, xml, pdf, docx)</summary>

Choose the output format with `--format`. All formats share the same parsing, deduplication, sorting, filtering, and selection — only the final write step changes. `--format` errors clearly and lists what's available if you typo a format name.

#### json

Writes one `.json` file for the whole run — a metadata envelope plus the contact array:

```json
{
  "meta": {
    "generated_by": "vCardTra",
    "version": "4.0",
    "exported": "August 09, 2026, 17:14",
    "source": "contacts.vcf",
    "contacts_found": 10,
    "contacts_exported": 8,
    "vcard_versions": ["2.1", "3.0"]
  },
  "contacts": [
    {
      "name": "John Smith",
      "organizations": ["Acme Corporation"],
      "phones": [{"value": "+1-555-0101", "type": "Work"}],
      "emails": [{"value": "john@acme.com", "type": "Work"}],
      "role": "Senior Engineer",
      "timezone": "America/New_York",
      "impp": ["xmpp:john@example.com"]
    }
  ]
}
```

`meta` carries the same information as the text/VCF header rows (exported time, source, counts, versions, sort/filter/select state) as native types instead of display strings. Keys with empty or missing values are omitted from each contact for cleanliness.

Phone and email values are structured objects (`{"value": "+1-555-0101", "type": "Work"}`), with `type` as `null` when there's no TYPE parameter in the source vCard.

If a contact has a decoded PHOTO or LOGO, it's included as base64:
```json
"photo": { "type": "JPEG", "data_base64": "/9j/4AAQ..." }
```
Use `--no-photos` to omit these and shrink the file.

#### csv

A wide table, one row per contact. Multi-value fields get numbered columns sized to whichever contact in the run has the most values for that field: `Address1`, `Address2`, etc. for addresses/URLs. Phones and emails additionally get a paired type column — `Phone1`, `Phone1Type`, `Phone2`, `Phone2Type`, ... `Email1`, `Email1Type`, ... — so the type is independently sortable/filterable in a spreadsheet rather than baked into the value cell. Photos aren't included (no sane CSV representation).

#### markdown / md

A single `.md` file — one `##` heading per contact, fields as a bullet list. Designed to drop straight into Notion, Obsidian, or a GitHub wiki.

#### html

A single self-contained `.html` file — no external CSS/JS. Photos are embedded inline as base64 `data:` images, so the file works standalone (open it, email it, whatever) without needing the original `.vcf`. This is also the layout PDF export builds on.

#### ics

A calendar file of birthdays and anniversaries as yearly-recurring events — import into any calendar app. Only dates with a specific day can become an event; a birthday stored as just a year (or year+month, no day) can't anchor a calendar date and is skipped, with a note printed to the terminal. Birthdays with no year use a placeholder year internally (the event still repeats every year regardless) and say "(year unknown)" in the event title.

#### sql

A plain `.sql` text file: one `CREATE TABLE IF NOT EXISTS contacts` statement followed by one `INSERT` per contact. Importable into any SQL database.

#### sqlite

Same schema as `sql`, written directly to a `.db` file — no separate import step, no server. Query it immediately with any SQLite client.

#### xml

The same information as `json`, structured as XML instead.

#### pdf

*Requires `pip install weasyprint`.* Renders the same layout as `--format html`, converted to a PDF you can print or share. A corrupted embedded photo is safely skipped (with the rest of the document rendering normally) rather than failing the whole export.

#### docx

*Requires `pip install python-docx`.* A Word document — one heading and field list per contact, photo inserted inline as an image when present and valid.

`template` (a twelfth, user-defined format) is covered separately under [Utility flags](#utility-flags) above, since it needs its own `--template <file>` argument rather than just `--format template`.

</details>

### --preview

Launches an interactive preview instead of writing a file. All sorting, filtering, and selection flags work exactly as in normal mode — contacts are processed first, then the preview opens on the result.

If the `textual` library is installed (`pip install textual`), a full two-panel TUI opens:
- Left panel: numbered contact list with `●` on selected and `?` on fuzzy duplicate pairs
- Right panel: full contact detail, updates as you move through the list
- `/` — live search by name, org, or phone
- `space` — select / deselect contact
- `a` / `n` — select all visible / clear all
- `s` — cycle sort modes (original, name A→Z, name Z→A, org A→Z)
- `e` — export selected contacts (or all visible if none selected) to a `.txt` file
- `q` — quit

If `textual` is not installed, a paged terminal fallback runs automatically — no error, no crash. 8 contacts per page; Enter/n = next page, `b` = back, `d` = detail view, `/` = search, `q` = quit.

```
vCardTra — 1 file to process

  Reading test_contacts.vcf ... 69 contacts, 3 malformed
vCardTra — Preview  (61 contacts)

  #    Name              Org               Phone
  ---- ----------------- ----------------- -----------
  1    John Smith        Acme Corporation  +1-555-0101
  2    Partial Paula     Paula Inc
  3    jane doe          Freelancer        +1-555-0301
  4    Mixed Case Mike   Mike Co           +1-555-0401
  5    Group Guy                           +1-555-0501
  6    Multi Value Mary  Company A         +1-555-0601
  7    Dedup Dana                          +1-555-0701
  8    Custom Carl                         +1-555-0802

  Page 1/8  |  Enter/n=next  b=back  d=detail  /search  q=quit
  >
```

**`--preview --browser`** skips both terminal modes and opens a one-shot HTML snapshot in your system's default browser instead — the same layout as `--format html`. It's a quick look, not an interactive session: no live search/select/export inside the browser tab. `--browser` on its own (without `--preview`) is an error, since it's a display-mode switch for preview, not a standalone flag.

### Filtering

Use `--filter` to export only contacts matching a condition. Repeatable — all conditions must match (AND logic):

| Condition | Matches contacts where… |
|---|---|
| `name=John` | name contains "John" (case-insensitive) |
| `org=Acme` | any organization contains "Acme" |
| `category=Work` | categories contains "Work" |
| `has=phone` | has at least one phone number |
| `has=email` | has at least one email address |
| `has=birthday` | has a birthday set |
| `has=address` | has at least one address |
| `has=note` | has at least one note |
| `has=url` | has at least one URL |

### Selection and exclusion

Use `--select` to keep only specific contacts and `--exclude` to remove them. Both accept the same expression syntax:

| Token | Meaning |
|---|---|
| `5` | Contact number 5 |
| `1-10` | Contacts 1 through 10 (inclusive) |
| `1-10,15,20-25` | Mixed ranges and individuals |
| `last-10` | The last 10 contacts |
| `John*` | All contacts whose name matches the wildcard |

Token types can be mixed: `--select "1-5,John*,last-2"`

Numbers refer to contacts in their final order after dedup and sort. `--select` and `--exclude` cannot be used together.

### Warnings

If a field value fails validation, it is reported at the bottom of that contact:

```
[Skipped Fields]
  ⚠ TEL: Too short: 123
  ⚠ EMAIL: Invalid email: notavalidemail
  ⚠ BDAY: Unrecognized format: NOTADATE
```

Any conversion-level issues appear at the very bottom of the file:

```
============================================================
[Conversion Warnings]
  ⚠ 1 malformed vCard block skipped (no parseable content)
  ⚠ 1 exact duplicate removed (newer revision kept where available):
      - John Smith → merged into Contact 3
  ⚠ Duplicate warning: Contact 2 and Contact 4 (Jane Doe)
============================================================
```

---

<details>
<summary><strong>Recognised Properties</strong> — click to expand the full vCard field reference</summary>

| Property | Type | Description |
|---|---|---|
| FN / N | Single | Full name. Falls back to N (name components) if FN is missing |
| ORG | Multiple | Company or organization name |
| TITLE | Single | Job title or position |
| TEL | Multiple | Phone numbers with type labels (Work, Home, Mobile, Cell, etc.) |
| EMAIL | Multiple | Email addresses with type labels |
| ADR | Multiple | Physical addresses with type labels |
| URL | Multiple | Websites |
| BDAY | Single | Birthday, formatted to human-readable (e.g. April 12, 1985) |
| NOTE | Multiple | Notes or comments |
| NICKNAME | Single | Nickname |
| LABEL | Single | Pre-formatted mailing address (vCard 2.1/3.0) |
| ANNIVERSARY | Single | Anniversary date, formatted same as birthday |
| GENDER | Single | Gender |
| CATEGORIES | Single | Comma-separated category list |
| CREATED | Single | Record creation datetime |
| REV | Single | Last modified datetime |
| TZ | Single | Time zone. Shown in every format except text/VCF |
| GEO | Single | Geographic coordinates (`lat;lon`). Shown in every format except text/VCF; HTML also renders a Google Maps link |
| ROLE | Single | Role or function (distinct from TITLE). Shown in every format except text/VCF |
| IMPP | Multiple | Instant messaging handles, scheme kept (e.g. `xmpp:john@example.com`). Shown in every format except text/VCF |
| PHOTO / LOGO | Single | Decoded from inline `ENCODING=BASE64` data into raw bytes; embedded in JSON/HTML/XML/PDF/DOCX (`--no-photos` to skip). Not available in CSV/SQL/SQLite. Data-URI and external-URI photo references are not yet extracted |
| X-* (vendor extensions) | Multiple | App-specific fields like Skype, WhatsApp, etc. Shown with a `*` prefix, `X-` removed from the label |

**TYPE parameter support:** Phone and email types are automatically detected and labelled (Work, Home, Mobile, Cell, Voice, Internet, Pref, etc.) for both vCard 2.1 and 3.0+ formats.

Binary fields with no practical use in any output format (SOUND, KEY) are silently skipped.

</details>

---

## Duplicate Handling

**Exact duplicates** are merged and removed automatically. Two contacts are considered exact duplicates when all of these match: name, phone numbers, email addresses, physical addresses, organizations, and birthday. Metadata fields like REV are excluded from this comparison intentionally.

When an exact duplicate is found, the script does not simply discard it. Instead, unique data from the duplicate is rescued and merged into the kept contact:

| Field | Merge behaviour |
|---|---|
| Title, Nickname, Anniversary, Gender, Categories, Label, Role, Timezone, Geo, URLs, IMPP | Taken from duplicate if missing in the base contact (URLs and IMPP merge unique values in) |
| Notes | Appended with `[merged]` prefix so you know the source |
| Custom X- fields | Merged into base, base wins on conflict |
| Revised timestamp | Newer value kept |
| Created timestamp | Earlier value kept (oldest origin date is the true one) |
| Photo, Logo | **Not** rescued — if the base contact has no photo but the removed duplicate did, it's lost. A deliberate gap, not an oversight — "which photo wins" is a real design decision |

REV-aware base selection: if both contacts have a REV timestamp, the newer one becomes the primary contact. The older one is merged into it. If neither has REV, the first occurrence in the file is kept as base.

**Fuzzy duplicates** — same name plus at least one shared phone or email — are flagged with a warning but kept. The script does not auto-remove these since two different people can share a name. You decide.

---

## What Gets Validated

| Field | Validation |
|---|---|
| Phone | Minimum 5 digits |
| Email | Must contain exactly one `@`, valid domain with `.` |
| Birthday | Must be a recognisable date format |

Fields that fail validation are skipped and reported. The rest of the contact is still exported normally.

---

<details>
<summary><strong>Troubleshooting</strong> — click to expand common issues and fixes</summary>

**"No files found matching '*.vcf'"**
Use quotes around wildcard patterns: `"*.vcf"`. Without quotes, some shells expand the pattern before Python sees it. You can also specify the full path: `"/path/to/files/*.vcf"`

**Output looks garbled (strange characters)**
The script automatically tries several encodings in sequence (UTF-8, UTF-16, Windows-1252, Latin-1). If your file uses a CJK encoding like Shift-JIS or GB2312, use `--encoding shift-jis` or `--encoding gb2312` to force the correct one. Installing the `charset-normalizer` package (`pip install charset-normalizer`) enables automatic detection of these encodings.

**Contact shows "Name: (Unknown)"**
The vCard has an empty or missing FN and N field. The contact is still exported with whatever other fields are present.

**Fewer contacts than expected**
Check the `[Conversion Warnings]` section at the bottom of the output file. It lists how many blocks were malformed or removed as exact duplicates.

**"Error: No .vcf files found"**
Verify the files have the `.vcf` extension. Use `ls *.vcf` (Unix/macOS) or `dir *.vcf` (Windows) to confirm files exist in the current directory.

</details>

---

## Common Use Cases

- 📱 Back up iPhone or Android contacts in readable format
- 📧 Convert email client contacts to text
- 🔍 Search contacts with `grep` or `find` without a contacts app
- 📝 Create contact lists for documentation
- 💾 Archive contacts before switching phones or apps
- 🔄 Merge exports from multiple sources and clean up duplicates
- 📊 Generate reports from contact data
- 🔐 Keep an offline readable backup of important contacts

---

## Platform Compatibility

Works on Windows, macOS, and Linux — anywhere Python 3.8+ runs.

---

## Roadmap

v5 is released: structured phone/email data model, `--convert-to` version converter, and nine utility flags (`--dedup-only`, `--validate`, `--diff`, `--update`, `--template`, `--watch`, `--quiet`, `--verbose`, `--dry-run`). See [CHANGELOG.md](CHANGELOG.md) for full history.

---

## License

Free to use, modify, and distribute.

## Contributing

Found a bug or want to add a feature? Contributions are welcome.

See [CHANGELOG.md](CHANGELOG.md) for full version history.

---

**Program:** vCardTra
**Version:** 5.0
**Last Updated:** October 2026