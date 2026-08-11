#!/usr/bin/env python3
"""
BGV Report Generator — Core Engine  (v2.1)
============================================
Changes v2.1:
  - Human-readable filenames: "{First}_{Last}_{CheckType}_Report.pdf"
  - PDF bytes returned in report for in-memory serving (no persistent output dir)
  - LibreOffice CLI fallback when win32com unavailable (e.g. Linux/Vercel)
  - Temp files cleaned up immediately after PDF bytes are captured
"""

import sys, os, re, shutil, zipfile, datetime, argparse, json
import threading, tempfile
from pathlib import Path

import openpyxl

SCRIPT_DIR    = Path(__file__).parent.resolve()
TEMPLATES_DIR = SCRIPT_DIR / "templates"

# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def fmt_date(value, fallback="N/A"):
    if value is None or str(value).strip() == "":
        return fallback
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime("%d-%b-%Y")
    val_str = str(value).strip()
    for fmt in ("%m/%d/%y", "%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.datetime.strptime(val_str, fmt).strftime("%d-%b-%Y")
        except ValueError:
            continue
    return val_str


def clean(value, fallback="N/A"):
    if value is None or str(value).strip() in ("", "none", "None"):
        return fallback
    return str(value).strip()


def safe_int_str(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def make_report_filename(applicant_name, check_type):
    """
    Produce a short, human-readable filename.
    'Cherrie May Quitua Requelman', 'Address'  →  'Cherrie_Requelman_Address_Report'
    """
    parts = str(applicant_name or "Unknown").strip().split()
    if len(parts) >= 2:
        short = f"{parts[0]}_{parts[-1]}"
    elif parts:
        short = parts[0]
    else:
        short = "Unknown"

    check_clean = re.sub(r"[^A-Za-z0-9]+", "_", str(check_type or "Report")).strip("_")
    name = f"{short}_{check_clean}_Report"
    name = re.sub(r"[^A-Za-z0-9_\-]+", "_", name)
    return re.sub(r"_+", "_", name).strip("_")


# ---------------------------------------------------------------------------
# Field builders
# ---------------------------------------------------------------------------

def build_address_fields(row, warnings):
    address = clean(row.get("Address"), fallback="")
    country = clean(row.get("Country"), fallback="")
    if country and country.lower() not in address.lower():
        full_address = f"{address}, {country}" if address else country
    else:
        full_address = address or country

    if not full_address or full_address == "N/A":
        warnings.append("Address and Country are both blank.")
        full_address = "N/A"

    case_ref = row.get("Case Reference Number")
    if case_ref is None or str(case_ref).strip() == "":
        warnings.append("Case Reference Number is blank.")
        case_ref_str = "N/A"
    else:
        case_ref_str = safe_int_str(case_ref)

    fields = {
        "Applicant Name":        clean(row.get("Applicant Name")),
        "Date of Birth":         fmt_date(row.get("Date of Birth")),
        "Address":               full_address,
        "Case Reference Number": case_ref_str,
        "Relevant Court":        clean(row.get("Relevant Court")),
        "Closure Date":          fmt_date(row.get("Closure date")),
    }
    return fields


# ---------------------------------------------------------------------------
# Check-type registry
# ---------------------------------------------------------------------------

CHECK_TYPE_REGISTRY = {
    "address": {
        "template":          TEMPLATES_DIR / "Address_report_format.docx",
        "build_fields":      build_address_fields,
        "required_columns":  [
            "Applicant Name", "Date of Birth", "Address", "Country",
            "Case Reference Number", "Relevant Court", "Closure date",
        ],
    },
}


# ---------------------------------------------------------------------------
# XML-level placeholder substitution
# ---------------------------------------------------------------------------

def xml_escape(value):
    return (
        str(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def fill_docx_template(template_path, fields, out_docx_path, warnings):
    work_dir = Path(tempfile.mkdtemp(prefix="bgv_unpack_"))
    try:
        with zipfile.ZipFile(template_path, "r") as z:
            z.extractall(work_dir)

        doc_xml_path = work_dir / "word" / "document.xml"
        xml_text = doc_xml_path.read_text(encoding="utf-8")

        tokens_in_doc = set(re.findall(r"&lt;([^&<>]+)&gt;", xml_text))

        for key, value in fields.items():
            token = f"&lt;{key}&gt;"
            replacement = xml_escape(value)
            if xml_text.count(token) == 0:
                warnings.append(f"Placeholder <{key}> not found in template.")
                continue
            xml_text = xml_text.replace(token, replacement)
            tokens_in_doc.discard(key)

        if tokens_in_doc:
            warnings.append(
                "Unfilled placeholders: " + ", ".join(sorted(f"<{t}>" for t in tokens_in_doc))
            )

        doc_xml_path.write_text(xml_text, encoding="utf-8")

        out_docx_path = Path(out_docx_path)
        if out_docx_path.exists():
            out_docx_path.unlink()

        with zipfile.ZipFile(out_docx_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for file in work_dir.rglob("*"):
                if file.is_file():
                    zf.write(file, file.relative_to(work_dir))
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# PDF conversion — Word COM (primary) + LibreOffice CLI (fallback)
# ---------------------------------------------------------------------------

_word_lock = threading.Lock()


def convert_to_pdf_word_com(docx_path, out_dir):
    """Use Microsoft Word's ExportAsFixedFormat (Windows only, best quality)."""
    import win32com.client, pythoncom

    docx_path = Path(docx_path).resolve()
    out_dir   = Path(out_dir).resolve()
    pdf_path  = out_dir / (docx_path.stem + ".pdf")

    pythoncom.CoInitialize()
    word = None
    try:
        with _word_lock:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            doc = word.Documents.Open(str(docx_path))
            try:
                doc.ExportAsFixedFormat(
                    OutputFileName=str(pdf_path),
                    ExportFormat=17,
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    Range=0,
                    Item=0,
                    IncludeDocProps=True,
                    KeepIRM=True,
                    CreateBookmarks=0,
                    DocStructureTags=True,
                    BitmapMissingFonts=True,
                    UseISO19005_1=False,
                )
            finally:
                doc.Close(False)
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    if not pdf_path.exists():
        raise RuntimeError(f"Word COM produced no PDF for {docx_path.name}")
    return pdf_path


def convert_to_pdf_libreoffice(docx_path, out_dir):
    """LibreOffice CLI fallback — works on Linux/macOS servers."""
    import subprocess
    docx_path = Path(docx_path).resolve()
    out_dir   = Path(out_dir).resolve()
    soffice_candidates = ["soffice", "libreoffice"]
    cmd = None
    for candidate in soffice_candidates:
        try:
            subprocess.run([candidate, "--version"], capture_output=True, timeout=10, check=True)
            cmd = candidate
            break
        except Exception:
            continue
    if not cmd:
        raise RuntimeError("Neither 'soffice' nor 'libreoffice' found on PATH.")
    result = subprocess.run(
        [cmd, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx_path)],
        capture_output=True, text=True, timeout=120,
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice failed: {result.stderr}")
    pdf_path = out_dir / (docx_path.stem + ".pdf")
    if not pdf_path.exists():
        raise RuntimeError("LibreOffice produced no PDF")
    return pdf_path


def convert_to_pdf(docx_path, out_dir):
    """Try win32com first; fall back to LibreOffice CLI."""
    try:
        return convert_to_pdf_word_com(docx_path, out_dir)
    except ImportError:
        pass  # win32com not available (Linux/Mac/Vercel)
    except Exception as e:
        import logging
        logging.warning(f"win32com PDF conversion failed ({e}), trying LibreOffice…")
    return convert_to_pdf_libreoffice(docx_path, out_dir)


# ---------------------------------------------------------------------------
# Excel reading
# ---------------------------------------------------------------------------

def read_rows(xlsx_path, sheet_name=None):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError(f"Sheet '{ws.title}' is empty.")
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    data_rows = []
    for raw in rows[1:]:
        if raw is None or all(v is None for v in raw):
            continue
        data_rows.append(dict(zip(headers, raw)))
    return headers, data_rows


# ---------------------------------------------------------------------------
# Main generation pipeline
# ---------------------------------------------------------------------------

def generate_all(xlsx_path, sheet_name=None, template_override=None,
                 progress_callback=None):
    """
    Generate one PDF per data row.

    Returns a dict: {
      "generated": [{"filename": str, "pdf_bytes": bytes, "applicant": str,
                     "check_name": str, "row": int, "warnings": list}],
      "errors":    [str]
    }
    PDF bytes are held in memory; no output directory is written by this function.
    """
    work_dir = Path(tempfile.mkdtemp(prefix="bgv_work_"))
    try:
        headers, rows = read_rows(xlsx_path, sheet_name)
        total  = len(rows)
        report = {"generated": [], "errors": []}

        for i, row in enumerate(rows, start=1):
            row_warnings = []
            s_no = row.get("S.No.", i)
            check_name_raw = row.get("Check Name", "")
            check_name = (check_name_raw or "").strip().lower()

            if check_name not in CHECK_TYPE_REGISTRY:
                report["errors"].append(
                    f"Row {i} (S.No. {s_no}): unknown Check Name '{check_name_raw}'. "
                    f"Supported: {list(CHECK_TYPE_REGISTRY.keys())}. Skipped."
                )
                if progress_callback:
                    progress_callback(i, total, f"Skipped row {i}: unknown check type")
                continue

            config = CHECK_TYPE_REGISTRY[check_name]
            template_path = Path(template_override) if template_override else config["template"]

            missing_cols = [c for c in config["required_columns"] if c not in headers]
            if missing_cols:
                report["errors"].append(
                    f"Row {i} (S.No. {s_no}): missing column(s) {missing_cols}. Skipped."
                )
                if progress_callback:
                    progress_callback(i, total, f"Skipped row {i}: missing columns")
                continue

            try:
                fields    = config["build_fields"](row, row_warnings)
                applicant = row.get("Applicant Name", f"Row{i}") or f"Row{i}"
                base_name = make_report_filename(applicant, check_name_raw)
                pdf_name  = f"{base_name}.pdf"

                docx_out = work_dir / f"{base_name}.docx"

                if progress_callback:
                    progress_callback(i - 1, total, f"Filling template for {applicant}…")

                fill_docx_template(template_path, fields, docx_out, row_warnings)

                if progress_callback:
                    progress_callback(i - 1, total, f"Converting to PDF for {applicant}…")

                pdf_path = convert_to_pdf(docx_out, work_dir)

                # Read bytes then delete temp files immediately
                pdf_bytes = pdf_path.read_bytes()
                pdf_path.unlink(missing_ok=True)
                docx_out.unlink(missing_ok=True)

                report["generated"].append({
                    "row":       i,
                    "s_no":      s_no,
                    "applicant": str(applicant),
                    "check_name": check_name_raw,
                    "filename":  pdf_name,
                    "pdf_bytes": pdf_bytes,
                    "warnings":  row_warnings,
                })

                if progress_callback:
                    progress_callback(i, total, f"Done: {applicant}")

            except Exception as e:
                report["errors"].append(f"Row {i} (S.No. {s_no}): {type(e).__name__}: {e}")
                if progress_callback:
                    progress_callback(i, total, f"Error on row {i}: {e}")

    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

    return report


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Generate BGV PDF reports from an Excel sheet.")
    parser.add_argument("xlsx_path",   help="Path to the input .xlsx file")
    parser.add_argument("output_dir",  help="Directory to write generated PDFs into")
    parser.add_argument("--sheet",     default=None)
    parser.add_argument("--template",  default=None)
    args = parser.parse_args()

    report = generate_all(
        args.xlsx_path,
        sheet_name=args.sheet,
        template_override=args.template,
    )

    out = Path(args.output_dir)
    out.mkdir(parents=True, exist_ok=True)
    for g in report["generated"]:
        (out / g["filename"]).write_bytes(g["pdf_bytes"])

    print(json.dumps(
        {k: [{kk: vv for kk, vv in r.items() if kk != "pdf_bytes"} for r in v]
         if isinstance(v, list) else v for k, v in report.items()},
        indent=2, default=str
    ))
    print(f"\nGenerated: {len(report['generated'])}   Errors: {len(report['errors'])}")
    for g in report["generated"]:
        flag = " (with warnings)" if g["warnings"] else ""
        print(f"  OK  row {g['row']:>3}  {g['applicant']:<35}  →  {out / g['filename']}{flag}")
    for e in report["errors"]:
        print(f"  FAIL {e}")
    if report["errors"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
