"""
Daily solar + wind curtailment (MU) and TRAS Down (MU) from the NLDC REMC workbook
written by vre.py (sheet 'Daily Data'), plotted in one panel in the same style as
the monthly overview charts. The chart is saved to the month's Monthly_Graphs folder.

Each day has two stacked bars side by side:
    left  (solid)   : Solar + Wind Curtailment
    right (hatched) : Solar + Wind TRAS Down
Colour = source (red solar, blue wind); fill pattern = quantity.

Usage:
    from vre_charts import plot_vre_curtailment
    plot_vre_curtailment(vre_panel_path(2026, 9))

Expected columns: 'Report for Date', 'Wind Curtailment (MU)',
'Solar Curtailment (MU)', 'TRAS Down Wind (MU)', 'TRAS Down Solar (MU)'
(and optionally 'Status').
"""

from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd

from config import graphs_month_folder

# --- Style constants (same as the overview charts) ---
FS_TITLE, FS_LABEL, FS_TICK, FS_LEG = 20, 17, 15, 16
C_GRID_MAJ = '#ebebeb'
C_GRID_MIN = '#f4f4f4'
C_TEXT     = '#000000'
C_TICK     = '#4d4d4d'
C_SOLAR, C_WIND = '#c44f6c', '#0085AC'

VALUE_COLS = ['Solar Curtailment (MU)', 'Wind Curtailment (MU)',
              'TRAS Down Solar (MU)',   'TRAS Down Wind (MU)']

plt.rcParams['hatch.linewidth'] = 1.2

plt.rcParams.update({
    'font.family':     'sans-serif',
    'font.sans-serif': ['Arial', 'Helvetica', 'DejaVu Sans'],
    'axes.unicode_minus': False,
})


def style_axes(ax):
    """theme_minimal(): no spines, no ticks, light major + minor grid."""
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.set_facecolor('white')
    ax.xaxis.set_minor_locator(mticker.AutoMinorLocator(2))
    ax.yaxis.set_minor_locator(mticker.AutoMinorLocator(2))
    ax.tick_params(axis='both', which='both', length=0)
    ax.tick_params(axis='both', which='major', labelsize=FS_TICK,
                   labelcolor=C_TICK, pad=6)
    ax.grid(which='major', color=C_GRID_MAJ, lw=1.0)
    ax.grid(which='minor', color=C_GRID_MIN, lw=0.5)
    ax.set_axisbelow(True)


def day_xticks(ax, total_days, target=6):
    """About `target` evenly spaced day labels, always including the last day."""
    step = max(1, round(total_days / target))
    tick_days = list(range(1, total_days + 1, step))
    if tick_days[-1] != total_days:
        tick_days.append(total_days)
    ax.xaxis.set_major_locator(mticker.FixedLocator(tick_days))
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'{int(v)}'))


def fmt_auto(vmax):
    """Tick formatter whose decimals depend on the axis magnitude."""
    d = 0 if vmax >= 20 else (1 if vmax >= 2 else 2)
    return mticker.FuncFormatter(lambda x, _: f'{x:,.{d}f}')


def read_vre(file_path):
    """Read the workbook; handle datetime or Excel-serial dates; keep Status == OK."""
    df = pd.read_excel(file_path)
    df = df.rename(columns=lambda c: str(c).strip())
    if 'Status' in df.columns:
        df = df[df['Status'].astype(str).str.strip().str.upper() == 'OK']

    raw = df['Report for Date']
    if pd.api.types.is_numeric_dtype(raw):
        df['Date'] = pd.to_datetime(raw, unit='D', origin='1899-12-30')
    else:
        df['Date'] = pd.to_datetime(raw, dayfirst=True, errors='coerce')

    for col in VALUE_COLS:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna(subset=['Date']).sort_values('Date').reset_index(drop=True)
    if df.empty:
        raise ValueError(f"No valid (Status == OK) daily rows in {file_path}.")
    return df


def plot_vre_curtailment(file_path, output_dir=None):
    """output_dir defaults to the month's Monthly_Graphs folder."""
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"{file_path} not found - run build_vre_month_excel first.")
    df = read_vre(file_path)

    month_name = df['Date'].dt.strftime('%B').iloc[0]
    year       = int(df['Date'].dt.year.iloc[0])
    month_num  = int(df['Date'].dt.month.iloc[0])
    # Plot only up to the last day with data (not the full calendar month)
    total_days = int(df['Date'].dt.day.max())

    # Reindex to every calendar day so days without a report show as a gap, not a zero
    daily = df.set_index(df['Date'].dt.day)[VALUE_COLS].reindex(range(1, total_days + 1))
    days       = daily.index.values.astype(float)
    solar_c    = daily['Solar Curtailment (MU)'].values.astype(float)
    wind_c     = daily['Wind Curtailment (MU)'].values.astype(float)
    solar_t    = daily['TRAS Down Solar (MU)'].values.astype(float)
    wind_t     = daily['TRAS Down Wind (MU)'].values.astype(float)
    total_c    = solar_c + wind_c
    total_t    = solar_t + wind_t

    missing = [int(d) for d, t in zip(days, total_c) if np.isnan(t)]
    if missing:
        print(f"No report for day(s) {missing}; left blank.")
    print(f"Curtailed:  {np.nansum(total_c):,.1f} MU "
          f"(solar {np.nansum(solar_c):,.1f}, wind {np.nansum(wind_c):,.1f})")
    print(f"TRAS Down:  {np.nansum(total_t):,.1f} MU "
          f"(solar {np.nansum(solar_t):,.1f}, wind {np.nansum(wind_t):,.1f})")

    fig, ax = plt.subplots(figsize=(17, 8), facecolor='white')
    fig.subplots_adjust(left=0.075, right=0.985, top=0.83, bottom=0.12)

    w = 0.38
    x_c, x_t = days - w / 2, days + w / 2

    # Curtailment: solid fill
    ax.bar(x_c, solar_c, width=w, color=C_SOLAR, alpha=0.9, zorder=3,
           label='Solar Curtailment')
    ax.bar(x_c, wind_c, width=w, bottom=np.nan_to_num(solar_c), color=C_WIND,
           alpha=0.9, zorder=3, label='Wind Curtailment')

    # TRAS Down: same colours, lighter fill + hatching
    ax.bar(x_t, solar_t, width=w, facecolor=C_SOLAR, alpha=0.45, edgecolor=C_SOLAR,
           hatch='///', lw=0, zorder=3, label='Solar TRAS Down')
    ax.bar(x_t, wind_t, width=w, bottom=np.nan_to_num(solar_t), facecolor=C_WIND,
           alpha=0.45, edgecolor=C_WIND, hatch='///', lw=0, zorder=3,
           label='Wind TRAS Down')

    ymax = np.nanmax(np.concatenate([total_c, total_t])) \
        if np.isfinite(np.concatenate([total_c, total_t])).any() else 0
    ax.set_ylim(0, ymax * 1.12 if ymax > 0 else 1)
    ax.set_xlim(0.5, total_days + 0.5)
    ax.yaxis.set_major_formatter(fmt_auto(ymax))
    day_xticks(ax, total_days)

    fig.text(0.005, 0.985,
             f'Daily Solar and Wind Curtailment vs TRAS Down, All India: {month_name} {year}',
             ha='left', va='top', fontsize=FS_TITLE, color=C_TEXT)
    ax.set_ylabel('Energy (GWh)', fontsize=FS_LABEL, color=C_TEXT, labelpad=10)
    ax.set_xlabel(f'Day of {month_name}', fontsize=FS_LABEL, color=C_TEXT, labelpad=10)
    style_axes(ax)
    ax.legend(loc='lower center', bbox_to_anchor=(0.5, 1.02), ncol=4,
              fontsize=FS_TICK, frameon=False, handlelength=2.0,
              columnspacing=1.6, handletextpad=0.6)

    output_dir = Path(output_dir) if output_dir else graphs_month_folder(year, month_num)
    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_dir / f"chart_vre_curtailment_tras_down_{month_name}_{year}.png"
    fig.savefig(out_path, dpi=300, bbox_inches='tight', facecolor='white')
    plt.show()
    plt.close(fig)
    print(f"Chart saved: {out_path}")
    return out_path


if __name__ == "__main__":
    import sys
    from config import vre_panel_path
    from datetime import date
    if len(sys.argv) > 1:
        plot_vre_curtailment(Path(sys.argv[1]), Path(sys.argv[2]) if len(sys.argv) > 2 else None)
    else:                                   # default: current month's panel
        today = date.today()
        plot_vre_curtailment(vre_panel_path(today.year, today.month))
