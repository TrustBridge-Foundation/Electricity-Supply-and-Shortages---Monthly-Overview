"""
Shared configuration: folder roots, path builders, and small helpers.

Every data folder, output folder and filename convention used by the
downloaders and plotters is derived from BASE_FOLDER + the path builders
defined here.
"""

import calendar
import re
from datetime import date
from pathlib import Path

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- Roots (relative - resolved against the working directory) ---

BASE_FOLDER   = Path(".")
PSP_ROOT      = BASE_FOLDER / "GRID_PSP_REPORTS"
AS_ROOT       = BASE_FOLDER / "GRID_AS_SCUC_REPORTS"
DGR_ROOT      = BASE_FOLDER / "NPP_DGR16_REPORTS"
VRE_ROOT      = BASE_FOLDER / "GRID_VRE_REPORTS"
GRAPHS_FOLDER = BASE_FOLDER / "Monthly_Graphs"

PSP_EXCEL_FOLDER  = PSP_ROOT / "Excel"
PSP_METADATA_FILE = PSP_ROOT / "grid_psp_excel_metadata.csv"


AS_PDF_FOLDER    = AS_ROOT / "PDF"
AS_PANEL_FOLDER  = AS_ROOT / "Panels"
AS_METADATA_FILE = AS_ROOT / "grid_as_scuc_metadata.csv"
AS_CACHE_FILE    = AS_ROOT / "_detected_report_type.json"

DGR_EXCEL_FOLDER  = DGR_ROOT / "Excel"
DGR_PANEL_FOLDER  = DGR_ROOT / "Panels"
DGR_METADATA_FILE = DGR_ROOT / "npp_dgr16_metadata.csv"

VRE_PDF_FOLDER    = VRE_ROOT / "PDF"
VRE_PANEL_FOLDER  = VRE_ROOT / "Panels"
VRE_METADATA_FILE = VRE_ROOT / "grid_vre_metadata.csv"

for _d in (PSP_EXCEL_FOLDER, AS_PDF_FOLDER, AS_PANEL_FOLDER,
           DGR_EXCEL_FOLDER, DGR_PANEL_FOLDER, VRE_PDF_FOLDER, VRE_PANEL_FOLDER,
           GRAPHS_FOLDER):
    _d.mkdir(parents=True, exist_ok=True)
    
CDN_BASE = "https://webcdn.grid-india.in/"

GRID_HEADERS = {"User-Agent": "Mozilla/5.0", "Content-Type": "application/json"}
NPP_HEADERS  = {"User-Agent": "Mozilla/5.0"}


# --- Shared helpers ---

def financial_year(cal_year: int, cal_month: int) -> str:
    """April-March FY: month 4-12 -> FY starts cal_year, month 1-3 -> cal_year-1."""
    if cal_month >= 4:
        return f"{cal_year}-{str(cal_year + 1)[-2:]}"
    return f"{cal_year - 1}-{str(cal_year)[-2:]}"


def expected_days(cal_year: int, cal_month: int) -> int:
    """Days in a calendar month (handles leap years)."""
    return calendar.monthrange(cal_year, cal_month)[1]


def month_folder_name(cal_year: int, cal_month: int) -> str:
    """Shared folder-name convention, e.g. '2026-07_July'."""
    return f"{cal_year}-{cal_month:02d}_{date(cal_year, cal_month, 1).strftime('%B')}"


def clean_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()


def to_num(val, default=None):
    """Safely convert a cell value to float; return default if not possible."""
    if val is None:
        return default
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return default


# --- Path builders ---

def psp_month_dir(cal_year: int, cal_month: int) -> Path:
    return PSP_EXCEL_FOLDER / financial_year(cal_year, cal_month) / month_folder_name(cal_year, cal_month)

def psp_timeseries_path(cal_year: int, cal_month: int) -> Path:
    folder = psp_month_dir(cal_year, cal_month)
    return folder / f"monthly_timeseries_{folder.name}.xlsx"

def as_month_folder(cal_year: int, cal_month: int) -> Path:
    return AS_PDF_FOLDER / financial_year(cal_year, cal_month) / month_folder_name(cal_year, cal_month)

def as_panel_path(cal_year: int, cal_month: int) -> Path:
    month_name = date(cal_year, cal_month, 1).strftime("%B")
    return AS_PANEL_FOLDER / f"AS_SCUC_monthly_{month_name.lower()}{cal_year}.xlsx"

def dgr_month_folder(cal_year: int, cal_month: int) -> Path:
    return DGR_EXCEL_FOLDER / financial_year(cal_year, cal_month) / month_folder_name(cal_year, cal_month)

def dgr_panel_path(cal_year: int, cal_month: int) -> Path:
    month_name = date(cal_year, cal_month, 1).strftime("%B")
    return DGR_PANEL_FOLDER / f"capacity_utilization_panel_{month_name.lower()}{cal_year}.xlsx"

def vre_month_folder(cal_year: int, cal_month: int) -> Path:
    return VRE_PDF_FOLDER / financial_year(cal_year, cal_month) / month_folder_name(cal_year, cal_month)

def vre_panel_path(cal_year: int, cal_month: int) -> Path:
    month_name = date(cal_year, cal_month, 1).strftime("%B")
    return VRE_PANEL_FOLDER / f"VRE_TRAS_Down_{month_name.lower()}{cal_year}.xlsx"

def graphs_month_folder(cal_year: int, cal_month: int) -> Path:
    folder = GRAPHS_FOLDER / month_folder_name(cal_year, cal_month)
    folder.mkdir(parents=True, exist_ok=True)
    return folder
