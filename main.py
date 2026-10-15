"""
Monthly power-sector overview: entry point.

Usage:
    1. Set CAL_YEAR and CAL_MONTH at the bottom of this file.
    2. python main.py

DOWNLOAD = True   downloads and consolidates the month's data first (opens Chrome for the
                  VRE reports), then draws the charts.
DOWNLOAD = False  skips every download and redraws the charts from the files already on disk.
"""

import os

# set working directory to wherever this script lives
os.chdir(os.path.dirname(os.path.abspath(__file__)))

from config import (
    psp_month_dir, psp_timeseries_path,
    as_panel_path, dgr_panel_path, vre_panel_path,
    graphs_month_folder,
)

from psp import fetch_psp_month, build_monthly_timeseries
from as_scuc import download_and_build_as_scuc_month
from iex import fetch_all_markets_month, iex_file_path
from dgr16 import download_and_build_month
from vre import fetch_vre_month, build_vre_month_excel
from vre_charts import plot_vre_curtailment
from charts import (
    plot_energy_met, plot_peak_demand_and_shortage,
    plot_sourcewise_gen, plot_generation_outages,
    plot_capacity_utilization,
    plot_dam_volumes_and_ceiling_rate, plot_rtm_volumes_and_ceiling_rate,
    plot_hpdam_volumes_and_ceiling_rate,
    plot_sras_tras_up, plot_sras_tras_down,
)
from psp_15min import build_extreme_days_15min, extreme_days_15min_path
from psp_15min_charts import plot_extreme_days_15min
from hourly_charts import run_hourly_analysis



def run_month(cal_year: int, cal_month: int):
    """Download data, build panels, and generate all charts for one month."""

    # --- Step 1: download + consolidate ---
    if DOWNLOAD:
        fetch_psp_month(cal_year, cal_month)
        build_monthly_timeseries(psp_month_dir(cal_year, cal_month))

        download_and_build_as_scuc_month(cal_year, cal_month)

        fetch_all_markets_month(cal_year, cal_month)

        download_and_build_month(cal_year, cal_month)

        fetch_vre_month(cal_year, cal_month)     # opens Chrome; leave the window alone

    build_vre_month_excel(cal_year, cal_month)   # parses the downloaded VRE PDFs (no browser)

    # 15-min TimeSeries data for the peak / lowest requirement days of the month
    # (needs the PSP downloads + monthly PSP workbook above to exist already)
    build_extreme_days_15min(cal_year, cal_month)

    # --- Step 2: resolve paths ---
    PSP_PATH      = psp_timeseries_path(cal_year, cal_month)
    AS_PATH       = as_panel_path(cal_year, cal_month)
    CAPACITY_PATH = dgr_panel_path(cal_year, cal_month)
    VRE_PATH      = vre_panel_path(cal_year, cal_month)

    DAM_DAILY     = iex_file_path("DAM",   "Daily", cal_year, cal_month)
    RTM_DAILY     = iex_file_path("RTM",   "Daily", cal_year, cal_month)
    HPDAM_DAILY   = iex_file_path("HPDAM", "Daily", cal_year, cal_month)
    DAM_15MIN     = iex_file_path("DAM",   "15Min", cal_year, cal_month)
    RTM_15MIN     = iex_file_path("RTM",   "15Min", cal_year, cal_month)
    HPDAM_15MIN   = iex_file_path("HPDAM", "15Min", cal_year, cal_month)

    for label, p in [("PSP", PSP_PATH), ("AS/SCUC", AS_PATH),
                     ("DGR16", CAPACITY_PATH), ("VRE/TRAS", VRE_PATH),
                     ("DAM daily", DAM_DAILY), ("RTM daily", RTM_DAILY),
                     ("HP-DAM daily", HPDAM_DAILY),
                     ("DAM 15-min", DAM_15MIN), ("RTM 15-min", RTM_15MIN),
                     ("HP-DAM 15-min", HPDAM_15MIN)]:
        print(f"{'OK ' if p.exists() else 'MISSING'}  {label:<15} {p}")

    # --- Step 3: charts ---
    plot_energy_met(PSP_PATH)
    plot_peak_demand_and_shortage(PSP_PATH)
    plot_sourcewise_gen(PSP_PATH)
    plot_generation_outages(PSP_PATH)

    plot_capacity_utilization(CAPACITY_PATH)

    plot_dam_volumes_and_ceiling_rate(DAM_DAILY, DAM_15MIN)
    plot_rtm_volumes_and_ceiling_rate(RTM_DAILY, RTM_15MIN)
    plot_hpdam_volumes_and_ceiling_rate(HPDAM_DAILY, HPDAM_15MIN)

    plot_sras_tras_up(AS_PATH)
    plot_sras_tras_down(AS_PATH)

    if VRE_PATH.exists():
        plot_vre_curtailment(VRE_PATH)
    else:
        print("Skipping VRE curtailment chart: no VRE workbook for this month yet.")

    # --- Step 4: analysis of the highest / lowest requirement days ---
    # Same two days in both: picked automatically from the monthly PSP file.
    plot_extreme_days_15min(extreme_days_15min_path(cal_year, cal_month))   # PSP 15-min charts + table
    run_hourly_analysis(                                                    # IEX bids/MCV/price charts
        cal_year, cal_month, PSP_PATH,
        iex_15min_paths={"DAM": DAM_15MIN, "RTM": RTM_15MIN, "HPDAM": HPDAM_15MIN},
    )

    print(f"\nAll charts written to {graphs_month_folder(cal_year, cal_month)}")


# --- SET THE MONTH HERE ---
CAL_YEAR  = 2026
CAL_MONTH = 9

# True: download + build everything, then plot.  False: plot from files already on disk.
DOWNLOAD  = True

if __name__ == "__main__":
    run_month(CAL_YEAR, CAL_MONTH)
