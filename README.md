# BGV Report Generator

Fills the "Address Verification Check" Word template with data from an
Excel sheet and produces one finished PDF per case.

## What was wrong with the old system

The old generator rebuilt the report's layout when inserting data, which is
why the table drifted right and cut off data. This version never rebuilds
anything: it copies the original `.docx`, and swaps only the text inside
the `<Placeholder>` tokens directly in the file's internal XML. Every
table width, margin, font, and style byte outside those placeholders is
left completely untouched — so the output layout is guaranteed to match
the approved template, because it *is* the same file. The last two pages
(disclaimer) are never touched at all, for the same reason.

## Usage

```bash
python3 generate_reports.py <input.xlsx> <output_dir> [--sheet SheetName]
```

Example:
```bash
python3 generate_reports.py Address.xlsx output
```

This produces one PDF per row in `output/`, named
`<CaseReferenceNumber>_<ApplicantName>_<CheckName>.pdf`.

## Input sheet columns expected (Address check)

`S.No.`, `Case Reference Number`, `Check Name`, `Applicant Name`, `Address`,
`Date of Birth`, `Country`, `Receiving date`, `Closure date`, `Status`,
`Relevant Court`

`Check Name` must be `Address` (case-insensitive) to match the current
template. Rows with any other/blank Check Name are skipped with a clear
message rather than failing the whole batch.

## Field mapping (Address template)

| Template placeholder | Source |
|---|---|
| `<Applicant Name>` | `Applicant Name` column |
| `<Date of Birth>` | `Date of Birth` column |
| `<Address>` | `Address` + `Country` columns, combined |
| `<Case Reference Number>` | `Case Reference Number` column |
| `<Relevant Court>` (shown next to "Source") | `Relevant Court` column |
| `<Closure Date>` | `Closure date` column |

**Note:** `Severity`, the `Results` (address-match) answer, and `Remarks`
are **not** placeholder tokens in the supplied template — they're static
analyst-written text, since they reflect judgment calls made during the
actual verification, not raw case metadata. Right now every generated
report will carry the same defaults present in the template (`Green` /
`Yes` / the sample remark). **If you want these automated per case too**,
add `Severity`, `Address Match`, and `Remarks` columns to the Excel sheet
and turn the corresponding text in the Word template into `<Severity>`,
`<Address Match>`, and `<Remarks>` tokens — then wire them into
`build_address_fields()` in `generate_reports.py` the same way the other
six fields work. Happy to do this now if you'd like.

## Robustness features

- **No layout rebuilding** — fixes the table-shift/data-cut bug at the root.
- **Long text (long names/addresses)** wraps and grows the row instead of
  clipping, verified with a stress test.
- **Special characters** (`&`, `<`, `>`, quotes) in data are XML-escaped
  before insertion, so they can't corrupt the output file or get dropped.
- **Blank/missing data** produces a warning attached to that row's result
  (visible in the JSON report and console output) instead of a silently
  broken PDF.
- **Missing required columns** in the input sheet fail that row with a
  clear message instead of an unreadable Python traceback.
- **Unsupported/blank Check Name** is skipped with a clear message instead
  of crashing the whole batch — one bad row doesn't block the rest.
- **Extensible by design** — add a new check type (Employment, Education,
  Criminal, …) by dropping a new template into `templates/` and adding an
  entry to `CHECK_TYPE_REGISTRY` in `generate_reports.py`. No other code
  changes needed.

## Files

```
generate_reports.py          the tool
templates/
  Address_report_format.docx the approved Word template (unmodified)
```

## Requirements

Python 3 with `openpyxl`, and LibreOffice (`soffice`) available for the
docx→PDF conversion step (already set up in this environment via
`/mnt/skills/public/docx/scripts/office/soffice.py`; in a different
environment, point `SOFFICE_SCRIPT` in the script at a `soffice` install
or call `soffice --headless --convert-to pdf` directly).
