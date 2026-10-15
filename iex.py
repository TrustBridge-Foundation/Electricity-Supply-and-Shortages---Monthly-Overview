"""
IEX market data - DAM / RTM / HP-DAM, daily and 15-minute-block
granularity. Downloads from IEX website and writes monthly Excel files.
"""

import time
from datetime import date
from io import StringIO
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

from config import (
    BASE_FOLDER,
    financial_year, expected_days, month_folder_name,
)


# --- Market configuration ---
# Everything that differs between DAM / RTM / HP-DAM lives here. Adding a
# new IEX market in future only requires adding one entry to this dict.

MARKETS = {
    "DAM": {
        "label": "Day-Ahead Market",
        "base_url": "https://www.iexindia.com/market-data/day-ahead-market/market-snapshot",
        "output_folder": BASE_FOLDER / "IEX_DAM_REPORTS",
        "has_session_id": False,   # expected 15-min table shape, for validation
    },
    "RTM": {
        "label": "Real-Time Market",
        "base_url": "https://www.iexindia.com/market-data/real-time-market/market-snapshot",
        "output_folder": BASE_FOLDER / "IEX_RTM_REPORTS",
        "has_session_id": True,
    },
    "HPDAM": {
        "label": "High-Price Day-Ahead Market",
        "base_url": "https://www.iexindia.com/market-data/high-price-day-ahead-market/market-snapshot",
        "output_folder": BASE_FOLDER / "IEX_HPDAM_REPORTS",
        "has_session_id": False,
    },
}

INTERVAL_PARAM = {
    "daily": "DAILY",
    "15min": "ONE_FOURTH_HOUR",
}

IEX_HEADERS = {
    "User-Agent": "Mozilla/5.0",
}

IEX_DELAY_SECONDS = 1  # delay between day-by-day requests (15-Min)


# --- generic helpers ---

def _validate_market(market: str) -> dict:
    if market not in MARKETS:
        raise ValueError(f"Unknown market '{market}'. Must be one of {list(MARKETS)}.")
    return MARKETS[market]


def get_paths(market: str):
    """Return (excel_folder, meta_daily, meta_15min) for a market."""
    cfg = _validate_market(market)
    output_folder = cfg["output_folder"]
    excel_folder = output_folder / "Excel"
    excel_folder.mkdir(parents=True, exist_ok=True)
    prefix = f"iex_{market.lower()}"
    meta_daily = output_folder / f"{prefix}_daily_metadata.csv"
    meta_15min = output_folder / f"{prefix}_15min_metadata.csv"
    return excel_folder, meta_daily, meta_15min


def _append_metadata(meta_file: Path, meta_row: pd.DataFrame, dedup_keys):
    if meta_file.exists():
        df_existing = pd.read_csv(meta_file)
        df_combined = pd.concat([df_existing, meta_row], ignore_index=True).drop_duplicates(
            subset=dedup_keys, keep="last"
        )
    else:
        df_combined = meta_row
    df_combined.to_csv(meta_file, index=False)


# --- generic HTTP fetch ---

def fetch_snapshot_html(base_url: str, interval_key: str, from_date: str, to_date: str) -> str:
    """
    GET a market-snapshot page for any of the three markets. Dates must be
    DD-MM-YYYY strings (matches the site's own format). Returns raw HTML.
    """
    params = {
        "interval": INTERVAL_PARAM[interval_key],
        "dp": "SELECT_RANGE",
        "showGraph": "false",
        "fromDate": from_date,
        "toDate": to_date,
    }
    resp = requests.get(base_url, params=params, headers=IEX_HEADERS, timeout=60)
    resp.raise_for_status()
    return resp.text


# --- generic parsers ---

def parse_daily_tables(html: str):
    """
    Parse the two HTML tables on a Daily-interval page: the daily data table
    (Date, Purchase Bid, Sell Bid, MCV, Final Scheduled Volume, MCP,
    Weighted MCP) and the Summary table (Total/Max/Min/Avg). Identical
    structure across DAM / RTM / HP-DAM. Returns (df_daily, df_summary).
    """
    tables = pd.read_html(StringIO(html))

    df_daily = None
    df_summary = None
    for t in tables:
        first_col = str(t.columns[0]).strip().lower()
        if first_col == "date" or first_col.startswith("date "):
            df_daily = t
        elif first_col == "summary" or first_col.startswith("summary "):
            df_summary = t

    if df_daily is None:
        raise ValueError(
            "Could not find the daily data table in the response. "
            "The page may have returned a 'Today' fallback (check params) "
            "or IEX changed the page layout."
        )

    return df_daily, df_summary


def parse_15min_table(html: str) -> pd.DataFrame:
    """
    Parse a 15-Min-Block-interval data table using BeautifulSoup.

    Auto-detects the table shape from its own <thead> rather than hard-coding
    it per market:
      - DAM / HP-DAM shape: 2 independently-collapsing leading columns
        (Date rowspan=96, Hour rowspan=4).
      - RTM shape: 3 independently-collapsing leading columns (Date
        rowspan=96, Hour rowspan=4, Session ID rowspan=2 - Session ID
        counts 1-48 continuously across the day, independent of Hour).

    Detection is by comparing the number of <td> cells actually present in
    each <tr> to the header count, so it self-adapts to whichever shape the
    page actually returns:
      2-column shape: missing=0 -> Date+Hour present (new day)
                       missing=1 -> only Hour present (new hour)
                       missing=2 -> both collapsed (continuing same hour)
      3-column shape: missing=0 -> Date+Hour+Session present (new day)
                       missing=1 -> Hour+Session present (new hour)
                       missing=2 -> only Session present (new session, same hour)
                       missing=3 -> all collapsed (continuing same session)

    pandas.read_html() is NOT used - it corrupts/shifts columns once a
    rowspan is present.
    """
    soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")

    if not tables:
        raise ValueError(
            "No tables found in the response. The page may have returned a "
            "'Today' fallback (check params) or IEX changed the page layout."
        )

    table = tables[0]
    thead = table.find("thead")
    if thead is None:
        raise ValueError("Expected table has no <thead> - page layout may have changed.")

    headers_found = [th.get_text(strip=True) for th in thead.find_all("th")]
    n_cols = len(headers_found)

    if not headers_found or headers_found[0] != "Date" or len(headers_found) < 2 or headers_found[1] != "Hour":
        raise ValueError(
            f"Unexpected table headers: {headers_found}. "
            "The page may have returned a 'Today' fallback (check params) "
            "or IEX changed the page layout."
        )

    has_session_id = len(headers_found) > 2 and headers_found[2] == "Session ID"

    tbody = table.find("tbody")
    if tbody is None:
        raise ValueError("Expected table has no <tbody> - page layout may have changed.")

    rows_out = []
    current_date = None
    current_hour = None
    current_session = None

    for tr in tbody.find_all("tr"):
        cells = tr.find_all("td")
        if not cells:
            continue

        missing = n_cols - len(cells)
        idx = 0

        if has_session_id:
            if missing == 0:
                current_date = cells[idx].get_text(strip=True); idx += 1
                current_hour = cells[idx].get_text(strip=True); idx += 1
                current_session = cells[idx].get_text(strip=True); idx += 1
            elif missing == 1:
                current_hour = cells[idx].get_text(strip=True); idx += 1
                current_session = cells[idx].get_text(strip=True); idx += 1
            elif missing == 2:
                current_session = cells[idx].get_text(strip=True); idx += 1
            elif missing == 3:
                pass
            else:
                raise ValueError(
                    f"Unexpected cell count: {len(cells)} cells vs {n_cols} header "
                    f"columns (missing={missing}). Page layout may have changed."
                )
            row_values = [current_date, current_hour, current_session] + [
                c.get_text(strip=True) for c in cells[idx:]
            ]
        else:
            if missing == 0:
                current_date = cells[idx].get_text(strip=True); idx += 1
                current_hour = cells[idx].get_text(strip=True); idx += 1
            elif missing == 1:
                current_hour = cells[idx].get_text(strip=True); idx += 1
            elif missing == 2:
                pass
            else:
                raise ValueError(
                    f"Unexpected cell count: {len(cells)} cells vs {n_cols} header "
                    f"columns (missing={missing}). Page layout may have changed."
                )
            row_values = [current_date, current_hour] + [c.get_text(strip=True) for c in cells[idx:]]

        rows_out.append(row_values)

    df = pd.DataFrame(rows_out, columns=headers_found)

    skip_cols = {"Date", "Hour", "Time Block"}
    if has_session_id:
        skip_cols.add("Session ID")
    for col in df.columns:
        if col in skip_cols:
            continue
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


# --- DAILY: fetch a full month in one request ---

def fetch_market_month_daily(market: str, cal_year: int, cal_month: int):
    """
    Download Daily-interval snapshot data for one market/month, save as
    Excel (Daily + Summary sheets), and append to the running metadata CSV.
    """
    cfg = _validate_market(market)
    excel_folder, meta_daily, _ = get_paths(market)

    month_name = date(cal_year, cal_month, 1).strftime("%B")
    month_padded = str(cal_month).zfill(2)
    fy = financial_year(cal_year, cal_month)
    days_in_month = expected_days(cal_year, cal_month)

    from_date = f"01-{month_padded}-{cal_year}"
    to_date = f"{str(days_in_month).zfill(2)}-{month_padded}-{cal_year}"

    save_folder = iex_month_dir(market, cal_year, cal_month)
    save_folder.mkdir(parents=True, exist_ok=True)

    save_path = iex_file_path(market, "Daily", cal_year, cal_month)
    excel_filename = save_path.name

    print(f"\n{'='*60}")
    print(f"  Fetching IEX {market} ({cfg['label']}) - Daily interval: {month_name} {cal_year}  |  FY {fy}")
    print(f"  Date range: {from_date} to {to_date}")
    print(f"  Expected days in month: {days_in_month}")
    print(f"  Save folder: {save_folder}")
    print(f"{'='*60}")

    if save_path.exists():
        print(f"  [SKIP] {excel_filename} already exists")
        return None

    html = fetch_snapshot_html(cfg["base_url"], "daily", from_date, to_date)
    df_daily, df_summary = parse_daily_tables(html)

    print(f"  Rows returned: {len(df_daily)}")

    if len(df_daily) < days_in_month:
        missing = days_in_month - len(df_daily)
        print(f"  WARNING: {missing} day(s) missing - "
              f"data may not yet be published, or request may have partially failed.")
    elif len(df_daily) == days_in_month:
        print(f"  OK: All {days_in_month} daily rows accounted for.")
    else:
        print("  NOTE: More rows than expected days - investigate before trusting this file.")

    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
        df_daily.to_excel(writer, sheet_name="Daily", index=False)
        if df_summary is not None:
            df_summary.to_excel(writer, sheet_name="Summary", index=False)

    print(f"  [SAVED] {save_path}")

    meta_row = pd.DataFrame([{
        "year": cal_year,
        "month": cal_month,
        "month_name": month_name,
        "fy": fy,
        "from_date": from_date,
        "to_date": to_date,
        "rows_returned": len(df_daily),
        "expected_days": days_in_month,
        "local_file": str(save_path),
        "fetched_at": pd.Timestamp.now().isoformat(timespec="seconds"),
    }])
    _append_metadata(meta_daily, meta_row, dedup_keys=["year", "month"])

    return df_daily, df_summary


# --- 15-MIN: fetch day-by-day, concatenate ---

def fetch_market_month_15min(market: str, cal_year: int, cal_month: int):
    """
    Download 15-Min-Block-interval snapshot data for one market/month by
    looping ONE DAY AT A TIME (pagination cap: 96 rows per request
    regardless of requested date span), save as Excel, and append to the
    running metadata CSV.
    """
    cfg = _validate_market(market)
    excel_folder, _, meta_15min = get_paths(market)

    month_name = date(cal_year, cal_month, 1).strftime("%B")
    month_padded = str(cal_month).zfill(2)
    fy = financial_year(cal_year, cal_month)
    days_in_month = expected_days(cal_year, cal_month)
    expected_rows = days_in_month * 96

    save_folder = iex_month_dir(market, cal_year, cal_month)
    save_folder.mkdir(parents=True, exist_ok=True)

    save_path = iex_file_path(market, "15Min", cal_year, cal_month)
    excel_filename = save_path.name

    print(f"\n{'='*60}")
    print(f"  Fetching IEX {market} ({cfg['label']}) - 15-Min-Block interval: {month_name} {cal_year}  |  FY {fy}")
    print(f"  Looping day-by-day: 01-{month_padded}-{cal_year} to "
          f"{str(days_in_month).zfill(2)}-{month_padded}-{cal_year}")
    print(f"  Expected rows: {expected_rows}  ({days_in_month} days x 96 blocks)")
    print(f"  Save folder: {save_folder}")
    print(f"{'='*60}")

    if save_path.exists():
        print(f"  [SKIP] {excel_filename} already exists")
        return None

    daily_frames = []
    failed_days = []

    for day_num in range(1, days_in_month + 1):
        day_str = f"{str(day_num).zfill(2)}-{month_padded}-{cal_year}"
        try:
            html = fetch_snapshot_html(cfg["base_url"], "15min", day_str, day_str)
            df_day = parse_15min_table(html)

            # Validate the table shape matched what this market is expected to have.
            got_session_id = "Session ID" in df_day.columns
            if got_session_id != cfg["has_session_id"]:
                raise ValueError(
                    f"'Session ID' column presence ({got_session_id}) does not match "
                    f"what is expected for {market} ({cfg['has_session_id']}). "
                    "Check that the URL/market mapping is correct."
                )

            returned_dates = df_day["Date"].unique()
            if len(returned_dates) != 1 or returned_dates[0] != day_str:
                print(f"  WARNING: {day_str}: got unexpected date(s) {list(returned_dates)} - skipping this day's data")
                failed_days.append(day_str)
                continue

            if len(df_day) != 96:
                print(f"  WARNING: {day_str}: got {len(df_day)} rows instead of 96")
            else:
                print(f"  [OK] {day_str}: 96 rows")

            daily_frames.append(df_day)

        except Exception as e:
            print(f"  ERROR: {day_str}: request/parse failed - {e}")
            failed_days.append(day_str)

        time.sleep(IEX_DELAY_SECONDS)

    if not daily_frames:
        raise ValueError(f"No data could be fetched for {market} {month_name} {cal_year} - every day failed.")

    df_15min = pd.concat(daily_frames, ignore_index=True)

    print(f"\n  Rows returned: {len(df_15min)}")

    if failed_days:
        print(f"  WARNING: {len(failed_days)} day(s) failed entirely: {failed_days}")

    if len(df_15min) < expected_rows:
        missing = expected_rows - len(df_15min)
        print(f"  WARNING: {missing} row(s) missing overall - "
              f"data may not yet be published for some days, or some requests failed.")
    elif len(df_15min) == expected_rows:
        print(f"  OK: All {expected_rows} 15-min rows accounted for.")
    else:
        print("  NOTE: More rows than expected - investigate before trusting this file.")

    per_day_counts = df_15min["Date"].value_counts()
    bad_days = per_day_counts[per_day_counts != 96]
    if not bad_days.empty:
        print("  WARNING: these dates do not have exactly 96 rows:")
        print(bad_days.to_string())

    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
        df_15min.to_excel(writer, sheet_name="15Min", index=False)

    print(f"  [SAVED] {save_path}")

    meta_row = pd.DataFrame([{
        "year": cal_year,
        "month": cal_month,
        "month_name": month_name,
        "fy": fy,
        "rows_returned": len(df_15min),
        "expected_rows": expected_rows,
        "failed_days": ";".join(failed_days) if failed_days else "",
        "local_file": str(save_path),
        "fetched_at": pd.Timestamp.now().isoformat(timespec="seconds"),
    }])
    _append_metadata(meta_15min, meta_row, dedup_keys=["year", "month"])

    return df_15min


# --- orchestrators ---

_INTERVAL_FUNCS = {
    "daily": fetch_market_month_daily,
    "15min": fetch_market_month_15min,
}


def fetch_market_month(market: str, cal_year: int, cal_month: int,
                        intervals=("daily", "15min")):
    """
    Fetch one market's data for one month, across the requested intervals
    (any subset of "daily", "15min"). Returns a dict keyed by interval name.
    """
    _validate_market(market)
    results = {}
    for interval in intervals:
        if interval not in _INTERVAL_FUNCS:
            raise ValueError(f"Unknown interval '{interval}'. Must be one of {list(_INTERVAL_FUNCS)}.")
        results[interval] = _INTERVAL_FUNCS[interval](market, cal_year, cal_month)
    return results


def fetch_all_markets_month(cal_year: int, cal_month: int,
                             markets=("DAM", "RTM", "HPDAM"),
                             intervals=("daily", "15min")):
    """
    Fetch all requested markets' data for one month, across the requested
    intervals. Returns a dict keyed by market name, each holding a dict
    keyed by interval name.
    """
    all_results = {}
    for market in markets:
        cfg = _validate_market(market)
        print(f"\n{'#'*70}")
        print(f"#  MARKET: {market} - {cfg['label']}")
        print(f"{'#'*70}")
        all_results[market] = fetch_market_month(market, cal_year, cal_month, intervals)
    return all_results


def fetch_month_range(start_year: int, start_month: int, end_year: int, end_month: int,
                       markets=("DAM", "RTM", "HPDAM"),
                       intervals=("daily", "15min")):
    """
    Convenience wrapper to fetch every month from (start_year, start_month)
    to (end_year, end_month) inclusive, e.g. a full financial year:
        fetch_month_range(2026, 4, 2027, 3)   # April 2026 .. March 2027
    """
    cal_year, cal_month = start_year, start_month
    while (cal_year, cal_month) <= (end_year, end_month):
        fetch_all_markets_month(cal_year, cal_month, markets=markets, intervals=intervals)
        if cal_month == 12:
            cal_year += 1
            cal_month = 1
        else:
            cal_month += 1


# --- Path builders (depend on MARKETS dict above) ---

def iex_month_dir(market: str, cal_year: int, cal_month: int) -> Path:
    return (MARKETS[market]["output_folder"] / "Excel"
            / financial_year(cal_year, cal_month) / month_folder_name(cal_year, cal_month))


def iex_file_path(market: str, interval_label: str, cal_year: int, cal_month: int) -> Path:
    """interval_label is 'Daily' or '15Min' - matches the on-disk filenames."""
    folder = month_folder_name(cal_year, cal_month)
    return iex_month_dir(market, cal_year, cal_month) / f"IEX_{market}_{interval_label}_{folder}.xlsx"
