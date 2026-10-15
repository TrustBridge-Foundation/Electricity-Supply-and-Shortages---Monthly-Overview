"""
IEX bids, market clearing volume and price (MCP) by 15-minute block for the two
extreme days of a month, plus the chart styling shared with psp_15min_charts.py.

The two days come from psp_15min.select_extreme_days: the highest and the lowest
all-India requirement (energy met + energy shortage, MU) in the month.

Output, one chart per IEX market, in the month's Monthly_Graphs folder:
  chart_hourly_<dam|rtm|hpdam>_bids_mcv_price_<Month>_<Year>.png
Input is the 15Min workbook written by iex.py. Volume is on the left axis (GW) and
price (MCP) on the right axis in Rs/kWh.
"""

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
from matplotlib.lines import Line2D

from config import graphs_month_folder
from charts import CEILING_DAM_RTM, CEILING_HPDAM
from psp_15min import select_extreme_days as find_extreme_days


# --- Settings ---

# (IEX market key used in iex.py, display name, price ceiling in Rs/MWh)
IEX_PLOT_MARKETS = [
    ("DAM",   "DAM",    CEILING_DAM_RTM),
    ("RTM",   "RTM",    CEILING_DAM_RTM),
    ("HPDAM", "HP-DAM", CEILING_HPDAM),
]

RECON_TOL_MW = 1.0   # tolerance: stacked sources vs the file's TOTAL GENERATION

# --- Style ---

C_GRID_MAJ = '#ebebeb'
C_GRID_MIN = '#f4f4f4'
C_TEXT     = '#000000'
C_TICK     = '#4d4d4d'


@dataclass(frozen=True)
class Style:
    title: int
    label: int
    tick: int
    leg: int
    lw: float


STYLE_GEN    = Style(title=20, label=17, tick=15, leg=20, lw=2.0)
STYLE_IEX    = Style(title=24, label=21, tick=19, leg=21, lw=2.4)

plt.rcParams.update({
    'font.family':        'sans-serif',
    'font.sans-serif':    ['Arial', 'Helvetica', 'DejaVu Sans'],
    'axes.unicode_minus': False,
})


# Source-wise generation: (label, header prefix, colour), bottom -> top of stack
SOURCES = [
    ('Nuclear', 'NUCLEAR',   '#8e44ad'),
    ('Hydro',   'HYDRO',     '#0085AC'),
    ('Gas',     'GAS',       '#6b3864'),
    ('Thermal', 'THERMAL',   '#7f8c8d'),
    ('Others',  'OTHERS',    '#bdc3c7'),
    ('Storage (PSP & BESS)', 'STORAGE (', '#e67e22'),
    ('Wind',    'WIND',      '#16a085'),
    ('Solar',   'SOLAR',     '#f1c40f'),
]
CHARGING_COLOR = '#e67e22'
C_REF          = '#222222'

# IEX colours
C_PURCHASE = '#0085AC'
C_SELL     = '#96b980'
C_MCV      = '#c44f6c'
C_PRICE    = '#6b3864'


# --- Small helpers ---

def style_axes(ax, st: Style):
    """theme_minimal(): no spines, no ticks, light major + minor grid."""
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.set_facecolor('white')
    ax.xaxis.set_minor_locator(mticker.AutoMinorLocator(2))
    ax.yaxis.set_minor_locator(mticker.AutoMinorLocator(2))
    ax.tick_params(axis='both', which='both', length=0)
    ax.tick_params(axis='both', which='major', labelsize=st.tick,
                   labelcolor=C_TICK, pad=6)
    ax.grid(which='major', color=C_GRID_MAJ, lw=1.0)
    ax.grid(which='minor', color=C_GRID_MIN, lw=0.5)
    ax.set_axisbelow(True)


def set_labels(ax, st: Style, ylabel=None, xlabel=None):
    if ylabel is not None:
        ax.set_ylabel(ylabel, fontsize=st.label, color=C_TEXT, labelpad=10)
    if xlabel is not None:
        ax.set_xlabel(xlabel, fontsize=st.label, color=C_TEXT, labelpad=10)


def add_title(fig, text, st: Style):
    fig.text(0.005, 0.985, text, ha='left', va='top', fontsize=st.title, color=C_TEXT)


def add_fig_legend(fig, handles, labels, ncol, y, st: Style):
    leg = fig.legend(handles, labels, loc='lower center', bbox_to_anchor=(0.5, y),
                     ncol=ncol, fontsize=st.leg, frameon=False, handlelength=2.6,
                     columnspacing=1.8, handletextpad=0.7, labelspacing=0.5)
    for ln in leg.get_lines():
        ln.set_linewidth(st.lw + 1.2)
    return leg


def hour_axis(ax, step=2, last=24):
    ax.set_xlim(0, 24)
    ax.xaxis.set_major_locator(mticker.FixedLocator(range(0, last + 1, step)))
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda h, _: f'{int(h):02d}:00'))


def save_hourly_chart(fig, stem, cal_year, cal_month):
    month_name = date(cal_year, cal_month, 1).strftime('%B')
    out_path = graphs_month_folder(cal_year, cal_month) / f"chart_{stem}_{month_name}_{cal_year}.png"
    fig.savefig(str(out_path), dpi=300, bbox_inches='tight', facecolor='white')
    plt.show()
    plt.close(fig)
    print(f"Chart saved: {out_path}")
    return str(out_path)


# --- Stacked generation panel (used by psp_15min_charts.py) ---

def _draw_gen_panel(ax, hours, series, show_ylabel, st):
    labels = [s[0] for s in SOURCES]
    colors = [s[2] for s in SOURCES]
    data = np.array([series[l] for l in labels]) / 1000                  # MW -> GW
    total_line = data.sum(axis=0)
    charging = -np.abs(series['Storage charging']) / 1000
    net_load = (series['Demand met'] - series['Wind'] - series['Solar']) / 1000

    ax.stackplot(hours, data, labels=labels, colors=colors, alpha=0.9, zorder=2)
    ax.fill_between(hours, charging, 0, facecolor=CHARGING_COLOR, alpha=0.45,
                    hatch='///', edgecolor='white', linewidth=0,
                    label='Storage charging (PSP & BESS)', zorder=2)
    ax.axhline(0, color='#555555', lw=0.8, zorder=3)

    cumsum = np.zeros(len(hours))
    for row in data:
        cumsum += row
        ax.plot(hours, cumsum, color='white', lw=0.8, alpha=0.6, zorder=3)

    ax.plot(hours, net_load, color=C_REF, lw=st.lw, solid_capstyle='round',
            alpha=0.85, label='Net Load (Demand - Wind - Solar)', zorder=4)
    ax.plot(hours, total_line, color=C_REF, lw=2.2, ls='--', alpha=0.6,
            label='Total Generation', zorder=4)

    hour_axis(ax)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:,.0f}'))
    set_labels(ax, st, 'Generation (GW)' if show_ylabel else None, 'Hour of day')
    style_axes(ax, st)


# --- IEX bids vs clearing volume vs price (from the 15Min workbooks) ---

def _find_col(columns, *starts, exclude=()):
    for c in columns:
        norm = " ".join(str(c).split()).upper()
        if any(norm.startswith(s.upper()) for s in starts) and not any(e in norm for e in exclude):
            return c
    return None


def _block_start_hour(value) -> float:
    start = str(value).split("-")[0].strip()
    t = datetime.strptime(start, "%H:%M").time()
    return t.hour + t.minute / 60


def load_iex_15min(path):
    """Read the whole monthly 15Min sheet once; add a parsed '_date' column."""
    df = pd.read_excel(path, sheet_name="15Min")
    df = df.rename(columns=lambda c: " ".join(str(c).split()))
    parsed = pd.to_datetime(df["Date"], format="%d-%m-%Y", errors="coerce")
    if parsed.isna().all():
        parsed = pd.to_datetime(df["Date"], errors="coerce")
    df["_date"] = parsed.dt.date
    return df


def iex_day(df, day: date):
    """(hours, purchase, sell, mcv, mcp) in MW / Rs-per-MWh for one day.
    MCP is all-NaN if the column is missing or empty (no trading that day)."""
    sub = df[df["_date"] == day]
    if sub.empty:
        raise ValueError(f"No 15-min rows for {day:%d-%m-%Y}.")
    tb  = _find_col(sub.columns, "TIME BLOCK")
    pur = _find_col(sub.columns, "PURCHASE BID")
    sel = _find_col(sub.columns, "SELL BID")
    mcv = _find_col(sub.columns, "MCV")
    mcp = _find_col(sub.columns, "MCP", exclude=("WEIGHTED",))
    if None in (tb, pur, sel, mcv):
        raise ValueError(f"Expected columns not found. Available: {list(sub.columns)}")

    num = lambda col: pd.to_numeric(sub[col], errors="coerce")
    hours = sub[tb].map(_block_start_hour).to_numpy(float)
    out = (hours,
           num(pur).fillna(0).to_numpy(float),
           num(sel).fillna(0).to_numpy(float),
           num(mcv).fillna(0).to_numpy(float),
           num(mcp).to_numpy(float) if mcp is not None else np.full(len(sub), np.nan))
    order = np.argsort(out[0], kind="stable")
    return tuple(a[order] for a in out)


def set_volume_axis(ax, vmax):
    """Round ticks; decimals matched to the tick step (no repeated '0, 0, 1, 1')."""
    ax.set_ylim(0, vmax)
    locator = mticker.MaxNLocator(nbins=6, steps=[1, 2, 2.5, 5, 10])
    ax.yaxis.set_major_locator(locator)
    ticks = locator.tick_values(0, vmax)
    step = ticks[1] - ticks[0]
    decimals = 0
    while decimals < 3 and abs(step * 10**decimals - round(step * 10**decimals)) > 1e-9:
        decimals += 1
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:,.{decimals}f}'))


def _draw_iex_panel(ax, d, title, price_ceiling, show_ylabel, st):
    hours, purchase, sell, mcv, mcp = d
    ax.fill_between(hours, 0, mcv / 1000, color=C_MCV, alpha=0.18, zorder=1)
    ax.plot(hours, purchase / 1000, color=C_PURCHASE, lw=st.lw, label='Purchase Bid', zorder=4)
    ax.plot(hours, sell / 1000,     color=C_SELL,     lw=st.lw, label='Sell Bid',     zorder=3)
    ax.plot(hours, mcv / 1000,      color=C_MCV,      lw=st.lw, label='Market Clearing Volume', zorder=4)

    ax.set_title(title, fontsize=st.label, color=C_TEXT, pad=10)
    hour_axis(ax, step=4, last=24)
    set_labels(ax, st, 'Volume (GW)' if show_ylabel else None, 'Hour of day')
    style_axes(ax, st)

    ax2 = ax.twinx()
    if not np.all(np.isnan(mcp)):
        ax2.plot(hours, mcp / 1000, color=C_PRICE, lw=st.lw, ls=':', alpha=0.9)   # Rs/MWh -> Rs/kWh
        ax2.axhline(price_ceiling / 1000, color=C_PRICE, lw=0.9, ls=':', alpha=0.5)
    set_volume_axis(ax2, price_ceiling / 1000 * 1.15)
    for spine in ax2.spines.values():
        spine.set_visible(False)
    ax2.grid(False)
    ax2.tick_params(axis='both', which='both', length=0)
    ax2.tick_params(axis='y', which='major', labelsize=st.tick, labelcolor=C_PRICE, pad=6)
    return ax2


def plot_iex_market_two_days(market_name, price_ceiling, df15, peak, low,
                             stem, cal_year, cal_month):
    st = STYLE_IEX
    d1, d2 = iex_day(df15, peak.day), iex_day(df15, low.day)

    fig, axes = plt.subplots(1, 2, figsize=(17, 9.5), facecolor='white', sharey=True)
    fig.subplots_adjust(left=0.075, right=0.91, top=0.77, bottom=0.12, wspace=0.40)
    _draw_iex_panel(axes[0], d1, peak.panel_title, price_ceiling, True, st)
    ax2_right = _draw_iex_panel(axes[1], d2, low.panel_title, price_ceiling, False, st)

    vol_max = max(a[1:4].max() for a in (np.array(d1), np.array(d2))) / 1000
    set_volume_axis(axes[0], vol_max * 1.08)
    ax2_right.set_ylabel('Price, MCP (Rs/kWh)', fontsize=st.label, labelpad=14, color=C_PRICE)

    add_title(fig, f'{market_name}: Bids, Clearing Volume and Price, '
                   f'{peak.label} vs {low.label}', st)
    price_proxy = Line2D([0], [0], color=C_PRICE, lw=st.lw, ls=':')
    handles, labels = axes[0].get_legend_handles_labels()
    add_fig_legend(fig, handles + [price_proxy], labels + ['Price (MCP)'], ncol=4, y=0.85, st=st)
    return save_hourly_chart(fig, stem, cal_year, cal_month)


# --- Orchestrator ---

def run_hourly_analysis(cal_year, cal_month, psp_path, iex_15min_paths=None):
    """
    psp_path         monthly PSP time-series workbook (used to pick the two days)
    iex_15min_paths  {"DAM": Path, "RTM": Path, "HPDAM": Path} of the 15Min workbooks.
                     Missing files are skipped.
    Each market is attempted independently, so one failure does not stop the rest.
    """
    print(f"\n{'='*60}\n  IEX 15-minute charts: extreme-requirement days\n{'='*60}")
    peak, low = find_extreme_days(psp_path, cal_year, cal_month)

    for key, name, ceiling in IEX_PLOT_MARKETS:
        path = (iex_15min_paths or {}).get(key)
        if not path or not Path(path).exists():
            print(f"  {name} 15-min workbook not found; skipping.")
            continue
        try:
            df15 = load_iex_15min(path)
            plot_iex_market_two_days(name, ceiling, df15, peak, low,
                                     f"hourly_{key.lower()}_bids_mcv_price",
                                     cal_year, cal_month)
        except Exception as e:
            print(f"  ERROR: {name} chart failed: {e}")

    return peak, low
