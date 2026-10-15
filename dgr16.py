"""
CEA NPP DGR16 - "Capacity (MW) Utilization For Power Generation" daily
report: downloader + monthly panel builder.
"""

import re
import os
import time
from datetime import date
from pathlib import Path

import pandas as pd
import pdfplumber
import requests

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from config import (
    NPP_HEADERS, DGR_METADATA_FILE,
    financial_year, expected_days,
    to_num,
    dgr_month_folder, dgr_panel_path,
)


NPP_BASE_URL      = "https://npp.gov.in/public-reports/cea/daily/dgr"
REPORT_CODE   = "dgr16"
DGR_DELAY_SECONDS = 1
EXTENSIONS_TO_TRY = ("pdf", "xls")   # pdf is current live format; xls kept as fallback for older dates

TARGET_ROWS = [
    "Coal", "Lignite", "Gas Liquid", "Diesel",
    "Total  THERMAL", "Total  NUCLEAR", "Total  HYDRO", "Total",
]

# Final panel column schema (superset - PDF format has more columns than
# the old xls format did; old-format rows simply leave the extra ones None)
COLUMN_NPP_HEADERS = [
    "Date", "Particulars",
    "Monitored Capacity (MW)",
    "Capacity under Long Outage (MW)",
    "New Capacity under Stabilization (MW)",
    "Outage - PM (MW)", "Outage - FO (MW)", "Outage - Others (MW)", "Outage - Total (MW)",
    "Capacity Outage as % of Monitored",
    "Capacity on Line - Gross (MW)", "Capacity on Line - Net (MW)",
    "Sum of Max Power Generated (MW)",
    "Max Power Output as % of Capacity on Line (Gross)",
    "Max Power Output as % of Monitored Capacity",
]

# Old xls format column indices (0-based), confirmed against real 2026 files
XLS_COL_MAP = {
    "particulars": 1, "monitored_cap": 3, "long_outage": 5,
    "outage_pm": 7, "outage_fo": 8, "outage_others": 9, "outage_total": 10,
    "cap_on_line": 12, "net_gen": 14, "pct_cap_on_line": 15,
}

NUM_RE = re.compile(r"^-?[\d,]+\.?\d*$")
ROW_Y_TOLERANCE = 3  # points; words within this vertical gap are treated as the same line


# --- GENERIC HELPERS ---

def build_report_url(report_date: date, report_code: str, ext: str) -> str:
    folder_part = report_date.strftime("%d-%m-%Y")
    file_part = report_date.strftime("%Y-%m-%d")
    return f"{NPP_BASE_URL}/{folder_part}/{report_code}-{file_part}.{ext}"


def verify_xls_readable(path: Path) -> bool:
    try:
        pd.read_excel(path, engine="xlrd", header=None, nrows=5)
        return True
    except Exception:
        return False


def verify_pdf_readable(path: Path) -> bool:
    try:
        data = path.read_bytes()
        if not data.startswith(b"%PDF"):
            return False
        with pdfplumber.open(path) as pdf:
            if not pdf.pages:
                return False
            text = pdf.pages[0].extract_text() or ""
            return "Particulars" in text and "Capacity" in text
    except Exception:
        return False


def verify_readable(path: Path, ext: str) -> bool:
    return verify_pdf_readable(path) if ext == "pdf" else verify_xls_readable(path)


# --- DOWNLOADER ---

def fetch_dgr16_month(cal_year: int, cal_month: int, report_code: str = REPORT_CODE,
                       verify_parse: bool = True, extensions=EXTENSIONS_TO_TRY):
    """
    Download every daily DGR16 report for one calendar month. Tries each
    extension in `extensions` order per day (pdf first, xls fallback) and
    keeps whichever one actually verifies as real content.
    """
    month_name = date(cal_year, cal_month, 1).strftime("%B")
    fy = financial_year(cal_year, cal_month)
    days_in_month = expected_days(cal_year, cal_month)
    save_folder = dgr_month_folder(cal_year, cal_month)
    save_folder.mkdir(parents=True, exist_ok=True)

    print(f"\n{'='*60}")
    print(f"  Fetching {report_code.upper()} (Capacity Utilization): {month_name} {cal_year}  |  FY {fy}")
    print(f"  Expected days in month: {days_in_month}")
    print(f"  Save folder: {save_folder}")
    print(f"{'='*60}")

    downloaded = skipped = failed = not_yet_published = 0
    saved_records = []

    for day_num in range(1, days_in_month + 1):
        report_date = date(cal_year, cal_month, day_num)
        date_str = report_date.strftime("%Y-%m-%d")
        display_date = report_date.strftime("%d-%b-%Y")

        # already have it under any known extension?
        existing = None
        for ext in extensions:
            p = save_folder / f"{report_code}-{date_str}.{ext}"
            if p.exists():
                existing = (p, ext)
                break

        record = {"date": date_str, "display_date": display_date}

        if existing:
            save_path, ext = existing
            print(f"  [SKIP] {save_path.name}")
            skipped += 1
            record.update(local_file=str(save_path), status="skipped_existing")
            saved_records.append(record)
            continue

        outcome = None
        for ext in extensions:
            filename = f"{report_code}-{date_str}.{ext}"
            save_path = save_folder / filename
            file_url = build_report_url(report_date, report_code, ext)

            print(f"  [DOWN] {filename}", end="")
            try:
                r = requests.get(file_url, headers=NPP_HEADERS, verify=False, timeout=60)
            except Exception as e:
                print(f"  -> ERR {e}")
                continue

            if r.status_code == 404:
                print(f"  -> 404 (.{ext} not found)")
                continue
            if r.status_code != 200:
                print(f"  -> HTTP {r.status_code}")
                continue

            save_path.write_bytes(r.content)
            if verify_parse and not verify_readable(save_path, ext):
                save_path.unlink(missing_ok=True)
                print(f"  -> .{ext} returned but not valid/parseable - trying next format")
                continue

            print("  -> OK")
            outcome = (save_path, ext, "downloaded")
            break

        if outcome is None:
            not_yet_published += 1
            record["status"] = "not_yet_published_or_unavailable"
        else:
            save_path, ext, status = outcome
            downloaded += 1
            record.update(local_file=str(save_path), status=status, extension=ext,
                           download_url=build_report_url(report_date, report_code, ext))

        saved_records.append(record)
        time.sleep(DGR_DELAY_SECONDS)

    total_on_disk = downloaded + skipped
    print(f"\n  Downloaded         : {downloaded}")
    print(f"  Skipped (existing) : {skipped}")
    print(f"  Not yet published  : {not_yet_published}")
    print(f"  Total on disk      : {total_on_disk}  |  Expected: {days_in_month}")

    if total_on_disk < days_in_month:
        print(f"  WARNING: {days_in_month - total_on_disk} day(s) missing.")
    elif total_on_disk == days_in_month:
        print(f"  OK: All {days_in_month} daily reports accounted for.")

    if saved_records:
        df_new = pd.DataFrame(saved_records)
        if DGR_METADATA_FILE.exists():
            df_existing = pd.read_csv(DGR_METADATA_FILE)
            df_combined = pd.concat([df_existing, df_new], ignore_index=True).drop_duplicates(
                subset=["date", "local_file"], keep="last"
            )
        else:
            df_combined = df_new
        df_combined.to_csv(DGR_METADATA_FILE, index=False)

    return saved_records


def fetch_dgr16_month_range(start_year, start_month, end_year, end_month,
                             report_code: str = REPORT_CODE, verify_parse: bool = True):
    cal_year, cal_month = start_year, start_month
    all_results = []
    while (cal_year, cal_month) <= (end_year, end_month):
        all_results.extend(fetch_dgr16_month(cal_year, cal_month, report_code, verify_parse))
        cal_year, cal_month = (cal_year + 1, 1) if cal_month == 12 else (cal_year, cal_month + 1)
    return all_results


# --- PARSERS ---

def safe_val(df, row_idx, col_idx):
    try:
        v = df.iloc[row_idx, col_idx]
        return None if pd.isna(v) else v
    except Exception:
        return None


def extract_date_from_filename(fname: str):
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", fname)
    if not m:
        return None
    y, mo, d = m.group(1), int(m.group(2)), int(m.group(3))
    month_abbr = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    return f"{d:02d}-{month_abbr[mo - 1]}-{y}"


def _empty_record(date_str, target):
    rec = {h: None for h in COLUMN_NPP_HEADERS}
    rec["Date"] = date_str
    rec["Particulars"] = target
    return rec


def parse_dgr16_xls(filepath):
    """Old xls format - fixed column indices (no blank-cell issue here)."""
    df = pd.read_excel(filepath, engine="xlrd", header=None)
    date_str = extract_date_from_filename(os.path.basename(filepath))

    particulars_lookup = {}
    for i in range(len(df)):
        raw = safe_val(df, i, XLS_COL_MAP["particulars"])
        if raw is not None:
            particulars_lookup[str(raw).strip()] = i

    rows = []
    for target in TARGET_ROWS:
        row_idx = particulars_lookup.get(target.strip())
        if row_idx is None:
            rows.append(_empty_record(date_str, target))
            continue

        monitored = safe_val(df, row_idx, XLS_COL_MAP["monitored_cap"])
        cap_on_line = safe_val(df, row_idx, XLS_COL_MAP["cap_on_line"])
        net_gen = safe_val(df, row_idx, XLS_COL_MAP["net_gen"])
        outage_total = safe_val(df, row_idx, XLS_COL_MAP["outage_total"])

        rec = _empty_record(date_str, target)
        rec["Monitored Capacity (MW)"] = monitored
        rec["Capacity under Long Outage (MW)"] = safe_val(df, row_idx, XLS_COL_MAP["long_outage"])
        rec["Outage - PM (MW)"] = safe_val(df, row_idx, XLS_COL_MAP["outage_pm"])
        rec["Outage - FO (MW)"] = safe_val(df, row_idx, XLS_COL_MAP["outage_fo"])
        rec["Outage - Others (MW)"] = safe_val(df, row_idx, XLS_COL_MAP["outage_others"])
        rec["Outage - Total (MW)"] = outage_total
        rec["Capacity on Line - Gross (MW)"] = cap_on_line
        rec["Sum of Max Power Generated (MW)"] = net_gen
        rec["Max Power Output as % of Monitored Capacity"] = safe_val(df, row_idx, XLS_COL_MAP["pct_cap_on_line"])
        try:
            if monitored:
                rec["Capacity Outage as % of Monitored"] = round(outage_total / monitored * 100, 2)
            if cap_on_line:
                rec["Max Power Output as % of Capacity on Line (Gross)"] = round(net_gen / cap_on_line * 100, 2)
        except (TypeError, ZeroDivisionError):
            pass
        rows.append(rec)

    return rows


def parse_dgr16_pdf(filepath):
    """
    New pdf format. Reads word-level x/y coordinates via pdfplumber and
    buckets numbers into column bands (using the Coal row - which has no
    blank cells - as the reference), rather than assuming a fixed
    left-to-right token sequence. This avoids misalignment on rows like
    Nuclear, which can have a blank "New Capacity under Stabilization" cell.
    """
    date_str = extract_date_from_filename(os.path.basename(filepath))
    target_word_lists = [re.sub(r"\s+", " ", t).strip().split() for t in TARGET_ROWS]

    ## Only checking page 1
    with pdfplumber.open(filepath) as pdf:
        words = pdf.pages[0].extract_words(keep_blank_chars=False)

    # group words into visual lines by y-position
    lines = []
    current_line, line_top = [], None
    for w in sorted(words, key=lambda w: w["top"]):
        if line_top is None or (w["top"] - line_top) <= ROW_Y_TOLERANCE:
            current_line.append(w)
            if line_top is None:
                line_top = w["top"]
        else:
            lines.append(sorted(current_line, key=lambda w: w["x0"]))
            current_line, line_top = [w], w["top"]
    if current_line:
        lines.append(sorted(current_line, key=lambda w: w["x0"]))

    def matches_label(line_words, target_words):
        texts = [w["text"] for w in line_words]
        if len(texts) < len(target_words) or texts[:len(target_words)] != target_words:
            return False
        if len(texts) > len(target_words) and not NUM_RE.match(texts[len(target_words)]):
            return False  # e.g. "Total" must not match "Total THERMAL"
        return True

    row_lines = {}
    for target, target_words in zip(TARGET_ROWS, target_word_lists):
        for lw in lines:
            if matches_label(lw, target_words):
                row_lines[target] = lw
                break

    # reference row for column x-bands: prefer Coal (no blank cells)
    reference = None
    for candidate in ("Coal", "Total"):
        lw = row_lines.get(candidate)
        if lw:
            nums = [w for w in lw if NUM_RE.match(w["text"])]
            if len(nums) >= 12:
                reference = nums
                break
    if reference is None:
        for lw in row_lines.values():
            nums = [w for w in lw if NUM_RE.match(w["text"])]
            if reference is None or len(nums) > len(reference):
                reference = nums

    if not reference:
        raise ValueError(f"Could not locate a reference data row in {filepath} - PDF layout may differ from expected.")

    ref_x = sorted(w["x0"] for w in reference)
    bounds = [-1e9] + [(ref_x[i] + ref_x[i + 1]) / 2 for i in range(len(ref_x) - 1)] + [1e9]
    n_cols = len(ref_x)
    field_names = [h for h in COLUMN_NPP_HEADERS if h not in ("Date", "Particulars")][:n_cols]

    def bucket(x0):
        for i in range(len(bounds) - 1):
            if bounds[i] <= x0 < bounds[i + 1]:
                return i
        return n_cols - 1

    rows = []
    for target in TARGET_ROWS:
        rec = _empty_record(date_str, target)
        lw = row_lines.get(target)
        if lw is None:
            rows.append(rec)
            continue

        slots = [None] * n_cols
        for w in lw:
            if NUM_RE.match(w["text"]):
                idx = bucket(w["x0"])
                if 0 <= idx < n_cols:
                    slots[idx] = to_num(w["text"])
        for fname, val in zip(field_names, slots):
            rec[fname] = val
        rows.append(rec)

    # self-check: PM + FO + Others should equal Total - warns (not fails) on mismatch
    for rec in rows:
        pm, fo, others, total = (rec.get("Outage - PM (MW)"), rec.get("Outage - FO (MW)"),
                                  rec.get("Outage - Others (MW)"), rec.get("Outage - Total (MW)"))
        if None not in (pm, fo, others, total) and abs((pm + fo + others) - total) > 1:
            print(f"  WARNING: {date_str} {rec['Particulars']}: PM+FO+Others ({pm+fo+others:.2f}) "
                  f"!= Total ({total:.2f}) - check for column misalignment.")

    return rows


def parse_dgr16_file(filepath: Path):
    ext = filepath.suffix.lower().lstrip(".")
    if ext == "pdf":
        return parse_dgr16_pdf(filepath)
    if ext == "xls":
        return parse_dgr16_xls(filepath)
    raise ValueError(f"Unsupported file extension: {ext}")


# --- PANEL BUILDER ---

def build_capacity_panel(cal_year: int, cal_month: int, report_code: str = REPORT_CODE):
    folder = dgr_month_folder(cal_year, cal_month)
    output_file = dgr_panel_path(cal_year, cal_month)

    if not folder.exists():
        print(f"No folder found at {folder} - run fetch_dgr16_month({cal_year}, {cal_month}) first.")
        return None

    report_files = sorted(
        f for f in os.listdir(folder)
        if re.match(rf"{re.escape(report_code)}-\d{{4}}-\d{{2}}-\d{{2}}\.(pdf|xls)$", f, re.IGNORECASE)
    )
    if not report_files:
        print(f"No {report_code}-YYYY-MM-DD.(pdf|xls) files found in {folder}.")
        return None

    print(f"Found {len(report_files)} file(s) in {folder}. Processing...")

    all_rows = []
    for fname in report_files:
        try:
            all_rows.extend(parse_dgr16_file(folder / fname))
            print(f"  OK: {fname}")
        except Exception as e:
            print(f"  ERROR: {fname}: {e}")

    if not all_rows:
        print("No data extracted. Exiting.")
        return None

    days_in_month = expected_days(cal_year, cal_month)
    if len(report_files) < days_in_month:
        print(f"  WARNING: Only {len(report_files)}/{days_in_month} days present in this panel.")

    wb = Workbook()
    ws = wb.active
    ws.title = "Panel Data"

    header_font   = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    header_fill   = PatternFill("solid", start_color="1F4E79")
    date_fill     = PatternFill("solid", start_color="D6E4F0")
    subtotal_fill = PatternFill("solid", start_color="EBF3FB")
    total_fill    = PatternFill("solid", start_color="BDD7EE")
    center_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align    = Alignment(horizontal="left", vertical="center")
    right_align   = Alignment(horizontal="right", vertical="center")
    thin          = Side(style="thin", color="AAAAAA")
    thin_border   = Border(left=thin, right=thin, top=thin, bottom=thin)
    pct_rows      = {"Total  THERMAL", "Total  NUCLEAR", "Total  HYDRO", "Total"}

    ws.row_dimensions[1].height = 36
    for col_idx, header in enumerate(COLUMN_NPP_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font, cell.fill, cell.alignment, cell.border = header_font, header_fill, center_align, thin_border

    row_num = 2
    for record in all_rows:
        particulars = record.get("Particulars", "")
        if particulars == "Total":
            fill, bold = total_fill, True
        elif particulars in pct_rows:
            fill, bold = subtotal_fill, True
        else:
            fill, bold = None, False

        for col_idx, header in enumerate(COLUMN_NPP_HEADERS, start=1):
            value = record.get(header)
            cell = ws.cell(row=row_num, column=col_idx, value=value)
            cell.border = thin_border
            cell.font = Font(name="Arial", size=9, bold=bold)
            if fill:
                cell.fill = fill
            elif header == "Date":
                cell.fill = date_fill

            if header == "Particulars":
                cell.alignment = left_align
            elif header == "Date":
                cell.alignment = center_align
            else:
                cell.alignment = right_align
                if value is not None:
                    cell.number_format = "0.00" if "%" in header else "#,##0.00"
        row_num += 1

    col_widths = [14, 20, 22, 28, 26, 16, 16, 18, 18, 22, 22, 22, 26, 30, 30]
    for i, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"

    wb.save(output_file)
    print(f"\nPanel data saved to:\n   {output_file}")
    print(f"   Rows: {row_num - 2} data rows  |  Files processed: {len(report_files)}")
    return output_file


# --- ORCHESTRATOR ---

def download_and_build_month(cal_year: int, cal_month: int, report_code: str = REPORT_CODE,
                              verify_parse: bool = True):
    fetch_dgr16_month(cal_year, cal_month, report_code, verify_parse)
    return build_capacity_panel(cal_year, cal_month, report_code)
