"""
Plotting functions for the monthly power-sector overview.

Style: ggplot2 theme_minimal() look
  - no spines or tick marks
  - light-grey major and minor gridlines on both axes
  - regular-weight title, axis titles and tick labels
  - title left-aligned to the figure edge, horizontal legend centred above the panel
  - optional direct end-of-line labels instead of a legend (LABEL_MODE = 'direct')

The Daily-interval workbook (volumes) and the 15Min workbook (ceiling hit-rate) are
separate files, so each plot_*_volumes_and_ceiling_rate function takes both paths.
"""

import calendar

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
from matplotlib.gridspec import GridSpec

from config import graphs_month_folder


# --- Shared style constants ---

LW_MAIN  = 2.0    # main series line width
MARKER_SIZE = 4.5  # dot at each data point (points)
LW_MEAN  = 1.5    # dashed mean / reference lines
FS_TITLE = 20
FS_LABEL = 17
FS_TICK  = 15
FS_LEG   = 16

C_GRID_MAJ = '#ebebeb'   # major gridlines  (ggplot grey92)
C_GRID_MIN = '#f4f4f4'   # minor gridlines  (lighter, thinner)
C_TEXT     = '#000000'   # title / axis titles
C_TICK     = '#4d4d4d'   # tick labels      (ggplot grey30)

# 'legend' -> horizontal legend above the panel (like the CE figure)
# 'direct' -> coloured labels at the end of each line (like the state figure);
#             applied only to charts made purely of lines
LABEL_MODE = 'legend'

plt.rcParams.update({
    'font.family':     'sans-serif',
    'font.sans-serif': ['Arial', 'Helvetica', 'DejaVu Sans'],
    'axes.unicode_minus': False,
})


# --- Shared helpers ---

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


def set_labels(ax, ylabel=None, xlabel=None):
    if ylabel is not None:
        ax.set_ylabel(ylabel, fontsize=FS_LABEL, color=C_TEXT, labelpad=10)
    if xlabel is not None:
        ax.set_xlabel(xlabel, fontsize=FS_LABEL, color=C_TEXT, labelpad=10)


def add_title(fig, text):
    """Regular-weight title flush with the left edge of the figure."""
    fig.text(0.005, 0.985, text, ha='left', va='top',
             fontsize=FS_TITLE, color=C_TEXT)


def day_xticks(ax, total_days, target=6):
    """About `target` evenly spaced day labels, always including the last day."""
    step = max(1, round(total_days / target))
    tick_days = list(range(1, total_days + 1, step))
    if tick_days[-1] != total_days:
        tick_days.append(total_days)
    ax.xaxis.set_major_locator(mticker.FixedLocator(tick_days))
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'{int(v)}'))


def direct_labels(ax, pad_frac=0.22, min_gap_frac=0.055):
    """
    Coloured end-of-line labels (state-wise figure style). Widens the right
    side of the panel and nudges labels apart so they never overlap.
    """
    ax.figure.subplots_adjust(top=0.90)   # no legend row needed above the panel
    x0, x1 = ax.get_xlim()
    span = x1 - x0
    ax.set_xlim(x0, x1 + span * pad_frac)
    y0, y1 = ax.get_ylim()
    gap = (y1 - y0) * min_gap_frac

    items = [(l.get_ydata()[-1], l.get_label(), l.get_color())
             for l in ax.get_lines()
             if not l.get_label().startswith('_') and l.get_linestyle() == '-']
    items.sort(key=lambda t: t[0])

    placed = []
    for y, _, _ in items:
        placed.append(y if not placed else max(y, placed[-1] + gap))
    for (y, lab, col), y_pos in zip(items, placed):
        ax.text(x1 + span * 0.015, y_pos, lab, color=col, fontsize=FS_LEG,
                va='center', ha='left', clip_on=False)


def add_legend(ax, ncol, y=1.02, direct=False, fontsize=FS_LEG):
    """Top-centred, frameless legend with thick line keys (or direct labels)."""
    if direct and LABEL_MODE == 'direct':
        direct_labels(ax)
        return None
    leg = ax.legend(loc='lower center', bbox_to_anchor=(0.5, y), ncol=ncol,
                    fontsize=fontsize, frameon=False, handlelength=2.0,
                    columnspacing=1.6, handletextpad=0.6)
    for ln in leg.get_lines():
        ln.set_linewidth(LW_MAIN + 0.6)
    return leg


def month_info(dates):
    d = pd.Series(dates)
    month_name = d.dt.strftime('%B').iloc[0]
    year       = int(d.dt.year.iloc[0])
    month_num  = int(d.dt.month.iloc[0])
    total_days = calendar.monthrange(year, month_num)[1]
    return d.dt.day.tolist(), month_name, year, month_num, total_days


def single_fig(figsize=(17, 8)):
    fig, ax = plt.subplots(figsize=figsize, facecolor='white')
    fig.subplots_adjust(left=0.075, right=0.985, top=0.83, bottom=0.12)
    return fig, ax


def two_panel_fig(figsize=(17, 10), height_ratios=(2, 1)):
    fig = plt.figure(figsize=figsize, facecolor='white')
    gs = GridSpec(2, 1, height_ratios=list(height_ratios), hspace=0.30, figure=fig,
                  left=0.075, right=0.985, top=0.84, bottom=0.08)
    ax1 = fig.add_subplot(gs[0])
    ax2 = fig.add_subplot(gs[1], sharex=ax1)
    return fig, ax1, ax2


def save_chart(fig, stem, month_name, year, month_num):
    month_folder = graphs_month_folder(year, month_num)
    out_path = month_folder / f"chart_{stem}_{month_name}_{year}.png"
    fig.savefig(str(out_path), dpi=300, bbox_inches='tight', facecolor='white')
    plt.show()
    plt.close(fig)
    print(f"Chart saved: {out_path}")
    return str(out_path)


def read_psp(file_path):
    df = pd.read_excel(file_path)
    df['Date'] = pd.to_datetime(df['Date'], format='%d-%m-%Y')
    return df.sort_values('Date').reset_index(drop=True)


fmt_int = mticker.FuncFormatter(lambda x, _: f'{x:,.0f}')


def fmt_auto(vmax):
    """Tick formatter whose decimals depend on the axis magnitude."""
    d = 0 if vmax >= 20 else (1 if vmax >= 2 else 2)
    return mticker.FuncFormatter(lambda x, _: f'{x:,.{d}f}')


# --- Graph 1: daily energy required vs met (PSP) ---

def plot_energy_met(file_path):

    df = read_psp(file_path)
    days, month_name, year, month_num, total_days = month_info(df['Date'])

    met      = df['A | Energy Met All-India (MU)'].values.astype(float)
    shortage = df['A | Energy Shortage All-India (MU)'].values.astype(float)
    required = met + shortage

    C_REQ, C_MET, C_FILL = '#c44f6c', '#0085AC', '#E79D88'

    fig, ax = single_fig()
    ax.fill_between(days, required, met, where=(required > met), alpha=0.55,
                    color=C_FILL, interpolate=True, label='Shortage (GWh)', zorder=2)
    ax.plot(days, required, color=C_REQ, lw=LW_MAIN, solid_capstyle='round', marker='o', ms=MARKER_SIZE,
            zorder=4, label='Requirement (Met + Shortage)')
    ax.plot(days, met, color=C_MET, lw=LW_MAIN, solid_capstyle='round', marker='o', ms=MARKER_SIZE,
            zorder=4, label='Energy Met')

    ymin, ymax = min(met.min(), required.min()), max(met.max(), required.max())
    pad = (ymax - ymin) * 0.12
    ax.set_xlim(0.5, total_days + 0.5)
    ax.set_ylim(ymin - pad, ymax + pad)
    ax.yaxis.set_major_formatter(fmt_int)
    day_xticks(ax, total_days)

    add_title(fig, f'All-India Energy Requirement vs Energy Met: {month_name} {year}')
    set_labels(ax, 'Energy (GWh)', f'Day of {month_name}')
    style_axes(ax)
    add_legend(ax, ncol=3)
    return save_chart(fig, 'energy_met', month_name, year, month_num)


# --- Graph 2 (merged): peak demand + shortage, solar vs non-solar hours (PSP) ---

def plot_peak_demand_and_shortage(file_path):

    df = read_psp(file_path)
    days, month_name, year, month_num, total_days = month_info(df['Date'])

    MW_PER_GW = 1000.0
    solar    = df['I | Solar-hr Max Demand Met (MW)'].values.astype(float) / MW_PER_GW
    nonsolar = df['I | Non-Solar-hr Max Demand Met (MW)'].values.astype(float) / MW_PER_GW
    mean_solar, mean_nonsolar = solar.mean(), nonsolar.mean()
    sol_short = df['I | Solar-hr Shortage (MW)'].values.astype(float) / MW_PER_GW
    nos_short = df['I | Non-Solar-hr Shortage (MW)'].values.astype(float) / MW_PER_GW

    C_SOLAR, C_NONSOLAR = '#0085AC', '#c44f6c'

    fig, ax1, ax2 = two_panel_fig()

    # Top panel: peak demand lines
    ax1.plot(days, solar, color=C_SOLAR, lw=LW_MAIN, solid_capstyle='round', marker='o', ms=MARKER_SIZE,
             zorder=4, label='Solar-Hour Peak Demand')
    ax1.plot(days, nonsolar, color=C_NONSOLAR, lw=LW_MAIN, solid_capstyle='round', marker='o', ms=MARKER_SIZE,
             zorder=4, label='Non-Solar-Hour Peak Demand')
    ax1.axhline(mean_solar, color=C_SOLAR, lw=LW_MEAN, ls='--', alpha=0.7,
                label=f'Mean Solar ({mean_solar:,.1f} GW)')
    ax1.axhline(mean_nonsolar, color=C_NONSOLAR, lw=LW_MEAN, ls='--', alpha=0.7,
                label=f'Mean Non-Solar ({mean_nonsolar:,.1f} GW)')

    ymin, ymax = min(solar.min(), nonsolar.min()), max(solar.max(), nonsolar.max())
    pad = (ymax - ymin) * 0.12
    ax1.set_xlim(0.5, total_days + 0.5)
    ax1.set_ylim(ymin - pad, ymax + pad)
    ax1.yaxis.set_major_formatter(fmt_auto(ymax))
    set_labels(ax1, 'Peak Demand Met (GW)')
    style_axes(ax1)
    plt.setp(ax1.get_xticklabels(), visible=False)
    add_legend(ax1, ncol=4, y=1.03)

    # Bottom panel: grouped bars of shortage
    days_arr, w = np.array(days), 0.38
    ax2.bar(days_arr - w / 2, sol_short, width=w, color=C_SOLAR, alpha=0.9,
            label='Solar-Hour Shortage', zorder=3)
    ax2.bar(days_arr + w / 2, nos_short, width=w, color=C_NONSOLAR, alpha=0.9,
            label='Non-Solar-Hour Shortage', zorder=3)
    short_max = max(sol_short.max(), nos_short.max())
    if short_max > 0:
        ax2.set_ylim(0, short_max * 1.12)
    ax2.yaxis.set_major_formatter(fmt_auto(short_max))
    day_xticks(ax2, total_days)
    set_labels(ax2, 'Shortage at Peak (GW)', f'Day of {month_name}')
    style_axes(ax2)
    add_legend(ax2, ncol=2, y=1.05)

    add_title(fig, f'All-India Peak Demand and Shortage, Solar vs Non-Solar Hours: {month_name} {year}')
    return save_chart(fig, 'peak_demand_shortage', month_name, year, month_num)


# --- Graph 3: source-wise generation, stacked (PSP) ---

def plot_sourcewise_gen(file_path):

    df = read_psp(file_path)
    days, month_name, year, month_num, total_days = month_info(df['Date'])

    sources = [
        ('Coal',             'G | Coal All-India (MU)',               '#4e3b2b'),
        ('Lignite',          'G | Lignite All-India (MU)',            '#a0785a'),
        ('Nuclear',          'G | Nuclear All-India (MU)',            '#A86093'),
        ('Gas / Naphtha',    'G | Gas/Naphtha/Diesel All-India (MU)', '#6b3864'),
        ('Hydro',            'G | Hydro All-India (MU)',              '#0085AC'),
        ('RES (Wind+Solar)', 'G | RES All-India (MU)',                '#96b980'),
    ]
    labels = [s[0] for s in sources]
    colors = [s[2] for s in sources]
    data   = np.array([df[s[1]].values.astype(float) for s in sources])
    total_line = data.sum(axis=0)

    fig, ax = single_fig()
    ax.stackplot(days, data, labels=labels, colors=colors, alpha=0.9, zorder=2)

    cumsum = np.zeros(len(days))
    for row in data:
        cumsum += row
        ax.plot(days, cumsum, color='white', lw=0.8, alpha=0.6, zorder=3)
    ax.plot(days, total_line, color='#222222', lw=2.2, ls='--', alpha=0.6,
            label='Total Generation', zorder=4)

    ax.set_xlim(0.5, total_days + 0.5)
    ax.set_ylim(0, total_line.max() * 1.08)
    ax.yaxis.set_major_formatter(fmt_int)
    day_xticks(ax, total_days)

    add_title(fig, f'Daily Source-wise Electricity Generation, All India: {month_name} {year}')
    set_labels(ax, 'Generation (GWh)', f'Day of {month_name}')
    style_axes(ax)
    add_legend(ax, ncol=len(sources) + 1, fontsize=FS_TICK)
    return save_chart(fig, 'sourcewise_gen', month_name, year, month_num)


# --- Graph 4: generation outages (PSP) ---

def plot_generation_outages(file_path):

    df = read_psp(file_path)
    days, month_name, year, month_num, total_days = month_info(df['Date'])

    grand_total = df['F | Total Outage All-India (MW)'].astype(float).values / 1000
    central     = df['F | Central Sector Outage All-India (MW)'].astype(float).values / 1000
    state       = df['F | State Sector Outage All-India (MW)'].astype(float).values / 1000

    C_TOTAL, C_CENTRAL, C_STATE = '#0085AC', '#c44f6c', '#96b980'

    fig, ax = single_fig()
    ax.plot(days, grand_total, color=C_TOTAL, lw=LW_MAIN, solid_capstyle='round', marker='o', ms=MARKER_SIZE,
            label='Grand Total Outage', zorder=4)
    ax.plot(days, central, color=C_CENTRAL, lw=LW_MAIN, solid_capstyle='round', marker='o', ms=MARKER_SIZE,
            label='Central Sector Outage', zorder=4)
    ax.plot(days, state, color=C_STATE, lw=LW_MAIN, solid_capstyle='round', marker='o', ms=MARKER_SIZE,
            label='State Sector Outage', zorder=4)

    ymax = max(grand_total.max(), central.max(), state.max())
    ymin = min(grand_total.min(), central.min(), state.min())
    pad  = (ymax - ymin) * 0.12
    ax.set_xlim(0.5, total_days + 0.5)
    ax.set_ylim(max(0, ymin - pad), ymax + pad)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:,.1f}'))
    day_xticks(ax, total_days)

    add_title(fig, f'Daily Generation Outages, All India: {month_name} {year}')
    set_labels(ax, 'Outage (GW)', f'Day of {month_name}')
    style_axes(ax)
    add_legend(ax, ncol=3, direct=True)
    return save_chart(fig, 'generation_outages', month_name, year, month_num)


# --- Graph 5: capacity utilization by fuel (DGR16 panel) ---

def plot_capacity_utilization(file_path):

    df = pd.read_excel(file_path)
    df['Date'] = pd.to_datetime(df['Date'], dayfirst=True)

    fuel_map = {'Total  THERMAL': 'Thermal', 'Total  NUCLEAR': 'Nuclear', 'Total  HYDRO': 'Hydro'}
    df = df[df['Particulars'].isin(fuel_map.keys())].copy()
    df['fuel'] = df['Particulars'].map(fuel_map)
    df = df.sort_values('Date').reset_index(drop=True)

    _, month_name, year, month_num, total_days = month_info(df['Date'])

    col = 'Max Power Output as % of Capacity on Line (Gross)'
    fuels = ['Thermal', 'Nuclear', 'Hydro']
    series = {}
    for fuel in fuels:
        sub = df[df['fuel'] == fuel].copy()
        sub['day'] = sub['Date'].dt.day
        sub = sub.sort_values('day')
        series[fuel] = {'days': sub['day'].tolist(), 'vals': sub[col].values.astype(float)}

    C = {'Thermal': '#bea69f', 'Nuclear': '#A86093', 'Hydro': '#0085AC'}

    fig, ax = single_fig()
    for fuel in fuels:
        d = series[fuel]
        ax.plot(d['days'], d['vals'], color=C[fuel], lw=LW_MAIN,
                solid_capstyle='round', marker='o', ms=MARKER_SIZE, zorder=4, label=fuel)

    all_vals = np.concatenate([series[f]['vals'] for f in fuels])
    ymin = max(all_vals.min() - 3, 60)
    ymax = all_vals.max() + 2
    ymin_r, ymax_r = np.floor(ymin / 5) * 5, np.ceil(ymax / 5) * 5

    ax.set_ylim(ymin_r, ymax_r)
    ax.set_xlim(0.5, total_days + 0.5)
    ax.set_yticks(np.arange(ymin_r, ymax_r + 1, 5))
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:.0f}%'))
    day_xticks(ax, total_days)

    add_title(fig, f'All-India Capacity Utilization, Thermal vs Nuclear vs Hydro: {month_name} {year}')
    set_labels(ax, 'Max power output as % of capacity on line', f'Day of {month_name}')
    style_axes(ax)
    add_legend(ax, ncol=3, direct=True)
    return save_chart(fig, 'capacity_utilization', month_name, year, month_num)


# --- Price-ceiling constants & 15-min MCP loader ---

CEILING_DAM_RTM = 10000.0   # CERC ceiling: Rs 10/kWh = Rs 10,000/MWh
CEILING_HPDAM   = 20000.0   # CERC HP-DAM ceiling: Rs 20/kWh = Rs 20,000/MWh
C_CEILING       = '#E79D88'  # single bar colour for the hit-rate panel


def load_mcp(path):
    """
    Read the automated downloader's 15-min sheet. The downloader writes this
    sheet as "15Min", with a clean single header row scraped straight from
    IEX's own table - no metadata rows to skip and no Total/Max/Min/Avg
    block mixed in (that only exists in the separate "Summary" sheet on
    Daily-interval files).
    """
    df = pd.read_excel(path, sheet_name="15Min")
    df = df.rename(columns=lambda c: str(c).strip())

    mcp_col = next(
        (c for c in df.columns if "MCP" in str(c).upper() and "WEIGHTED" not in str(c).upper()),
        None,
    )
    if mcp_col is None:
        raise KeyError(
            f"No MCP column found in {path.name}. Available columns: {list(df.columns)}"
        )
    df["MCP"] = pd.to_numeric(df[mcp_col], errors="coerce")

    parsed_date = pd.to_datetime(df["Date"], format="%d-%m-%Y", errors="coerce")
    if parsed_date.isna().all():
        parsed_date = pd.to_datetime(df["Date"], errors="coerce")
    df["Day"] = parsed_date.dt.day
    df = df.dropna(subset=["Day"])
    return df



def daily_ceiling_hit_rate(mcp_df, total_days, ceiling):
    """
    Number of 15-min blocks at or above `ceiling` for each day of the
    month (1..total_days).
    """
    hits_per_day = (mcp_df["MCP"] >= ceiling).groupby(mcp_df["Day"]).sum()

    full_days = pd.Series(0.0, index=range(1, total_days + 1))
    full_days.update(hits_per_day)
    return full_days.sort_index()


# --- Shared body for the three "exchange volumes + ceiling hit-rate" charts ---

def _plot_volumes_and_ceiling_rate(days, purchase, sell, mcv, hit_rate, total_days,
                                   market_label, month_name, year, month_num,
                                   out_stub):

    C_PURCHASE, C_SELL, C_MCV = '#0085AC', '#96b980', '#c44f6c'

    # Taller figure and a larger bottom share so the hit-rate panel stays readable
    fig, ax1, ax2 = two_panel_fig(figsize=(17, 12), height_ratios=(1.5, 1))

    # Top panel: exchange volumes
    ax1.plot(days, purchase, color=C_PURCHASE, lw=LW_MAIN, solid_capstyle='round', marker='o', ms=MARKER_SIZE,
             label='Purchase Bid (GWh)', zorder=5)
    ax1.plot(days, sell, color=C_SELL, lw=LW_MAIN, solid_capstyle='round', marker='o', ms=MARKER_SIZE,
             label='Sell Bid (GWh)', zorder=5)
    ax1.plot(days, mcv, color=C_MCV, lw=LW_MAIN, solid_capstyle='round', marker='o', ms=MARKER_SIZE,
             label='MCV (GWh)', zorder=5)

    left_max = max(purchase.max(), sell.max(), mcv.max())
    ax1.set_xlim(0.5, total_days + 0.5)
    ax1.set_ylim(0, left_max * 1.15)
    ax1.yaxis.set_major_locator(mticker.MaxNLocator(nbins=6, steps=[1, 2, 2.5, 5, 10]))
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(
        lambda x, _: f'{x:,.1f}' if left_max < 20 else f'{x:,.0f}'))
    set_labels(ax1, 'Daily volume aggregated\nacross trading blocks (GWh)')
    style_axes(ax1)
    plt.setp(ax1.get_xticklabels(), visible=False)
    add_legend(ax1, ncol=3, y=1.03)

    # Bottom panel: % of 15-min blocks at the ceiling, one flat colour
    peak_rate = hit_rate.max()
    ax2.bar(hit_rate.index, hit_rate.values, width=0.7, color=C_CEILING, zorder=3)
    if peak_rate > 0:
        ax2.set_ylim(0, max(peak_rate * 1.15, 10))   # at least 0-10% so tiny rates don't look huge
    else:
        ax2.set_ylim(0, 10)
        ax2.text(0.5, 0.5, f'No 15-min block hit the price ceiling in {month_name} {year}',
                 transform=ax2.transAxes, ha='center', va='center',
                 fontsize=FS_TICK, color='#888888', fontstyle='italic', zorder=7)

    
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:.0f}'))
    day_xticks(ax2, total_days)
    set_labels(ax2, 'No. of ceiling-hit\nblocks', f'Day of {month_name}')
    style_axes(ax2)

    add_title(fig, f'IEX {market_label} Exchange Volumes and Price-Ceiling Hit-Rate: {month_name} {year}')
    return save_chart(fig, out_stub, month_name, year, month_num)


def _load_daily_volumes(path):
    """Shared loader for the Daily-interval sheet used by all three markets."""
    df = pd.read_excel(path, sheet_name="Daily")
    df = df.rename(columns=lambda c: str(c).strip())
    df['Date'] = pd.to_datetime(df['Date'], format='%d-%m-%Y', errors='coerce')
    df = df.dropna(subset=['Date']).copy()
    for col in ['Purchase Bid (MWh)', 'Sell Bid (MWh)', 'MCV (MWh)']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    return df.sort_values('Date').reset_index(drop=True)


def _market_chart(daily_path, min15_path, ceiling, market_label, out_stub):
    """Common driver for DAM / RTM / HP-DAM."""
    dfv = _load_daily_volumes(daily_path)
    days, month_name, year, month_num, total_days = month_info(dfv['Date'])

    MWH_PER_GWH = 1000.0
    purchase = dfv['Purchase Bid (MWh)'].astype(float).values / MWH_PER_GWH
    sell     = dfv['Sell Bid (MWh)'].astype(float).values / MWH_PER_GWH
    mcv      = dfv['MCV (MWh)'].astype(float).values / MWH_PER_GWH

    mcp_df   = load_mcp(min15_path)
    hit_rate = daily_ceiling_hit_rate(mcp_df, total_days, ceiling)
    print(f"  {market_label}: ceiling {ceiling:,.0f} Rs/MWh | MCP "
          f"{mcp_df['MCP'].min():,.0f}-{mcp_df['MCP'].max():,.0f} "
          f"({len(mcp_df)} blocks, peak day hit-rate {hit_rate.max():.0f}%)")

    return _plot_volumes_and_ceiling_rate(
        days, purchase, sell, mcv, hit_rate, total_days,
        market_label=market_label, month_name=month_name, year=year,
        month_num=month_num, out_stub=out_stub,
    )


# --- Graphs 6-8: DAM / RTM / HP-DAM volumes + price-ceiling hit-rate ---

def plot_dam_volumes_and_ceiling_rate(dam_daily_path, dam_15min_path):
    return _market_chart(dam_daily_path, dam_15min_path, CEILING_DAM_RTM,
                         'DAM', 'dam_volumes_ceiling_rate')


def plot_rtm_volumes_and_ceiling_rate(rtm_daily_path, rtm_15min_path):
    return _market_chart(rtm_daily_path, rtm_15min_path, CEILING_DAM_RTM,
                         'RTM', 'rtm_volumes_ceiling_rate')


def plot_hpdam_volumes_and_ceiling_rate(hpdam_daily_path, hpdam_15min_path):
    return _market_chart(hpdam_daily_path, hpdam_15min_path, CEILING_HPDAM,
                         'HP-DAM', 'hpdam_volumes_ceiling_rate')


# --- Graph 9: SRAS Up & TRAS Up (AS/SCUC) ---

def plot_sras_tras_up(as_path):

    as_df = pd.read_excel(as_path, sheet_name="AS_SCUC_Monthly")
    as_df = as_df.rename(columns=lambda c: str(c).strip())
    as_df['Date'] = pd.to_datetime(as_df['Date'], errors='coerce')
    as_df = as_df.dropna(subset=['Date']).copy()
    for col in ['SRAS Up', 'TRAS Up']:
        as_df[col] = pd.to_numeric(as_df[col], errors='coerce')
    as_df = as_df.sort_values('Date').reset_index(drop=True)

    days, month_name, year, month_num, total_days = month_info(as_df['Date'])

    MWH_PER_MU = 1000.0
    sras_up = as_df['SRAS Up'].astype(float).values / MWH_PER_MU
    tras_up = as_df['TRAS Up'].astype(float).values / MWH_PER_MU

    C_SRAS, C_TRAS = '#0085AC', '#c44f6c'

    fig, ax = single_fig()
    ax.plot(days, sras_up, color=C_SRAS, lw=LW_MAIN, solid_capstyle='round', marker='o', ms=MARKER_SIZE,
            label='SRAS Up', zorder=5)
    ax.plot(days, tras_up, color=C_TRAS, lw=LW_MAIN, solid_capstyle='round', marker='o', ms=MARKER_SIZE,
            label='TRAS Up', zorder=5)

    ymax = np.nanmax([np.nanmax(sras_up), np.nanmax(tras_up)])
    ax.set_ylim(0, ymax * 1.12)
    ax.set_xlim(0.5, total_days + 0.5)
    ax.yaxis.set_major_locator(mticker.MaxNLocator(nbins=6, steps=[1, 2, 2.5, 5, 10]))
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(
        lambda x, _: f'{x:,.1f}' if ymax < 20 else f'{x:,.0f}'))
    day_xticks(ax, total_days)

    add_title(fig, f'SRAS Up & TRAS Up: {month_name} {year}')
    set_labels(ax, 'Volume Up (MU)', f'Day of {month_name}')
    style_axes(ax)
    add_legend(ax, ncol=2, direct=True)
    return save_chart(fig, 'sras_tras_up', month_name, year, month_num)


# --- Graph 10: SRAS Down & TRAS Down (AS/SCUC) ---
# The report signs down-regulation negative. Plotted here as absolute values so
# both series sit above zero; the y-axis names the direction instead of the sign.

def plot_sras_tras_down(as_path):

    as_df = pd.read_excel(as_path, sheet_name="AS_SCUC_Monthly")
    as_df = as_df.rename(columns=lambda c: str(c).strip())
    as_df['Date'] = pd.to_datetime(as_df['Date'], errors='coerce')
    as_df = as_df.dropna(subset=['Date']).copy()
    for col in ['SRAS Down', 'TRAS Down']:
        as_df[col] = pd.to_numeric(as_df[col], errors='coerce')
    as_df = as_df.sort_values('Date').reset_index(drop=True)

    days, month_name, year, month_num, total_days = month_info(as_df['Date'])

    MWH_PER_MU = 1000.0
    sras_down = np.abs(as_df['SRAS Down'].astype(float).values) / MWH_PER_MU
    tras_down = np.abs(as_df['TRAS Down'].astype(float).values) / MWH_PER_MU

    C_SRAS, C_TRAS = '#0085AC', '#c44f6c'

    fig, ax = single_fig()
    ax.plot(days, sras_down, color=C_SRAS, lw=LW_MAIN, solid_capstyle='round', marker='o', ms=MARKER_SIZE,
            label='SRAS Down', zorder=5)
    ax.plot(days, tras_down, color=C_TRAS, lw=LW_MAIN, solid_capstyle='round', marker='o', ms=MARKER_SIZE,
            label='TRAS Down', zorder=5)

    ymax = max(sras_down.max(), tras_down.max())
    ax.set_ylim(0, ymax * 1.12)
    ax.set_xlim(0.5, total_days + 0.5)
    ax.yaxis.set_major_locator(mticker.MaxNLocator(nbins=6, steps=[1, 2, 2.5, 5, 10]))
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(
        lambda x, _: f'{x:,.1f}' if ymax < 20 else f'{x:,.0f}'))
    day_xticks(ax, total_days)

    add_title(fig, f'SRAS Down & TRAS Down: {month_name} {year}')
    set_labels(ax, 'Volume Down (MU)', f'Day of {month_name}')
    style_axes(ax)
    add_legend(ax, ncol=2, direct=True)
    return save_chart(fig, 'sras_tras_down', month_name, year, month_num)