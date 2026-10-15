# Monthly Indian Power-Sector Overview

Automated pipeline that downloads a month of Indian power-sector data from public sources, consolidates it into monthly panels, and generates 15 graphs covering demand, shortages, generation mix, capacity utilization, ancillary services, wind and solar curtailment, IEX market outcomes, and a 15-minute comparison of the highest- and lowest-demand days of the month.

## Charts

### Monthly overview (one point per day of the given month)

- Demand required, met and shortages
- Solar and non-solar peak demand, with shortages during solar and non-solar peak hours
- Source-wise generation
- Generation outages
- Capacity utilisation of thermal, nuclear and hydro plants
- SRAS and TRAS
  - Up procurement
  - Down procurement
- Solar and wind curtailment and TRAS Down
- IEX volumes bid and traded, and number of times prices hit the ceiling
  - DAM
  - RTM
  - HP-DAM

### Highest- and lowest-requirement days (15-minute resolution)

The two days are chosen automatically: the day of highest and the day of lowest all-India electricity requirement in the month, where requirement = energy met + energy shortage (MU).

- Source-wise generation by 15-minute block, one panel per day
- IEX bids, market clearing volume and price (MCP) by 15-minute block, one chart each for DAM, RTM and HP-DAM

A summary workbook also compares the two days: peak and minimum demand with their times, load factor, net-load ramps, solar and wind energy, wind + solar share, thermal range, storage use, and grid frequency statistics.

## Data sources

| Source | What it provides |
|---|---|
| Grid-India daily Power Supply Position (PSP) report | All-India energy met/shortage, peak demand, source-wise generation, outages; the `TimeSeries` sheet gives 15-minute SCADA data |
| Grid-India daily AS/SCUC report | Ancillary services (SRAS/TRAS up and down regulation) |
| Grid-India daily VRE (NLDC REMC) report | Solar and wind curtailment and TRAS Down (Tables 8 and 9) |
| Indian Energy Exchange | DAM, RTM, and HP-DAM prices and volumes (daily and 15-minute) |
| National Power Portal Daily Generation Report | Daily capacity utilization by fuel type |

## File structure

| File | Description |
|---|---|
| `main.py` | Set the month/year here and run this |
| `config.py` | Folder roots, path builders, shared helpers |
| `psp.py` | Grid-India Power Supply Position reports: download + monthly timeseries builder |
| `as_scuc.py` | Ancillary Services: PDF download + monthly panel builder |
| `iex.py` | IEX DAM / RTM / HP-DAM data (daily + 15-min) |
| `dgr16.py` | NPP capacity utilization: download + panel |
| `vre.py` | Grid-India VRE reports: PDF download (Selenium) + curtailment / TRAS Down panel builder |
| `psp_15min.py` | Picks the highest- and lowest-requirement days and extracts their 15-minute data from the daily PSP files |
| `charts.py` | Monthly overview plotting functions |
| `vre_charts.py` | Solar and wind curtailment vs TRAS Down chart |
| `psp_15min_charts.py` | 15-minute analysis of the two extreme days: summary table and source-wise generation chart |
| `hourly_charts.py` | IEX bids / clearing volume / price charts for the two extreme days; also holds the chart styling shared with `psp_15min_charts.py` |
| `requirements.txt` | Python dependencies |

## Usage

Install dependencies:

```
python -m pip install -r requirements.txt
```

The VRE download drives Chrome through Selenium, so Google Chrome must be installed.

Open `main.py` and set the year and month number at the bottom of the file, for example:

```python
CAL_YEAR  = 2026
CAL_MONTH = 4
DOWNLOAD  = True
```

Run:

```
python main.py
```

With `DOWNLOAD = True` the run first downloads and consolidates the month's data, then draws the charts. The VRE step opens a Chrome window; leave it alone until it finishes. Once a month's data is on disk, set `DOWNLOAD = False` to redraw the charts without downloading anything again.

Before plotting, `main.py` prints one line per input file (`OK` or `MISSING`). A missing file makes the chart that needs it fail, except for the VRE chart, which is skipped with a message.

## Output

All data and charts are written relative to the script's directory:

```
GRID_PSP_REPORTS/        raw + consolidated PSP files
  Extreme_Days_15Min/    15-minute data for the highest- and lowest-requirement days
GRID_AS_SCUC_REPORTS/    raw + consolidated AS files
GRID_VRE_REPORTS/        raw + consolidated VRE (curtailment / TRAS Down) files
NPP_DGR16_REPORTS/       raw + consolidated capacity utilization files
IEX_DAM_REPORTS/         DAM price and volume data
IEX_RTM_REPORTS/         RTM price and volume data
IEX_HPDAM_REPORTS/       HP-DAM price and volume data
Monthly_Graphs/          output charts and summary tables, one subfolder per month
```

## Notes on definitions

- **Requirement** is energy met plus energy shortage, in MU (1 MU = 1 GWh).
- **Net load** is demand met minus wind minus solar. Grid-India's demand met includes storage load (pumping and battery charging), so net load does too.
- **Price ceilings** used for the ceiling hit-rate: Rs 10/kWh for DAM and RTM, Rs 20/kWh for HP-DAM. To note, price ceilings on DAM and RTM before April 2023 was Rs 12/kWh.
- 15-minute PSP values are instantaneous SCADA readings. Energy integrated from them can differ slightly from the energy reported in the PSP summary.
- All data are taken as published. Grid-India SCADA data can contain telemetry errors, and some reports are published with a delay, so recent days may be missing.

## Acknowledgement

We thank Aryan Anand and Akshat Shukla for their contributions during the initial stages of developing this code.
