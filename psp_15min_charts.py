"""
15-minute analysis and charts for the peak-requirement and lowest-requirement
days of a month. Input is the workbook produced by psp_15min.py.

Chart (saved to the month's Monthly_Graphs folder):
  chart_psp15_sourcewise_gen_<Month>_<Year>.png stacked generation, storage charging,
                                                net load; one panel per day

Table:
  summary_psp15_<Month>_<Year>.xlsx             day-by-day comparison metrics

Definitions (all from the SCADA TimeSeries sheet, so all are instantaneous
15-minute values on the transmission system):
  Net load      NL_t = DemandMet_t - Wind_t - Solar_t        (the file's NET DEMAND MET)
  Energy        E = sum_t P_t * 0.25 h / 1000                (MW -> MU)
  k-hour ramp   max_t (NL_{t+4k} - NL_t)  (up)  and  min_t (...)  (down), within the day
  VRE share     s_t = (Wind_t + Solar_t) / DemandMet_t
Note that, as defined by Grid-India, DEMAND MET includes storage load
(pumping + BESS charging).

Usage:
    from psp_15min_charts import plot_extreme_days_15min
    plot_extreme_days_15min(extreme_days_15min_path(2026, 9))
"""

from datetime import date
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from config import graphs_month_folder
from psp_15min import DayPick
from hourly_charts import (
    STYLE_GEN, SOURCES, C_TEXT, RECON_TOL_MW,
    add_title, add_fig_legend, save_hourly_chart, _draw_gen_panel,
)

FREQ_BAND = (49.90, 50.05)     # Hz, normal operating band assumed in the summary table
BLOCK_H   = 0.25               # hours per 15-min block


# --- Load ---

def load_extreme_days(path):
    """Return (picks, {kind: day DataFrame}) from the psp_15min.py workbook."""
    sel = pd.read_excel(path, sheet_name="Day_Selection")
    data = pd.read_excel(path, sheet_name="15Min_Data")
    sel["Date"] = pd.to_datetime(sel["Date"]).dt.date
    data["Date"] = pd.to_datetime(data["Date"]).dt.date

    picks, days = {}, {}
    for _, r in sel.iterrows():
        kind = str(r["Day Type"]).lower()
        picks[kind] = DayPick(r["Date"], float(r["Value (MU)"]), kind)
        d = data[data["Date"] == r["Date"]].sort_values("Block").reset_index(drop=True)
        if d.empty:
            raise ValueError(f"No 15-min rows for {r['Date']} in {path}")
        days[kind] = d
    if set(picks) != {"peak", "lowest"}:
        raise ValueError(f"Expected one 'peak' and one 'lowest' day, found {list(picks)}")
    return picks, days


def _gen_series(d):
    """Day DataFrame -> the dict layout expected by the stacked-generation panel."""
    g = lambda c: d[c].to_numpy(float)
    s = {
        'Nuclear': g('Nuclear (MW)'), 'Hydro': g('Hydro (MW)'), 'Gas': g('Gas (MW)'),
        'Thermal': g('Thermal (MW)'), 'Others': g('Others (MW)'),
        'Storage (PSP & BESS)': np.nan_to_num(g('Storage Discharge (MW)')),
        'Wind': g('Wind (MW)'), 'Solar': g('Solar (MW)'),
        'Storage charging': np.nan_to_num(g('Storage Demand (MW)')),
        'Demand met': g('Demand Met (MW)'),
        'Total (file)': g('Total Generation (MW)'),
    }
    gap = np.abs(sum(s[l] for l, _, _ in SOURCES) - s['Total (file)']).max()
    if gap > RECON_TOL_MW:
        print(f"  [warning] {d['Date'].iloc[0]}: stacked sources differ from the file's "
              f"TOTAL GENERATION by up to {gap:,.0f} MW.")
    return s


# --- Analysis ---

def _clock(hour: float) -> str:
    m = int(round(hour * 60))
    return f"{m // 60:02d}:{m % 60:02d}"


def day_metrics(d: pd.DataFrame, requirement_mu: float):
    """Ordered list of (metric, value) for one day. Powers in GW, energies in MU."""
    h = d['Hour'].to_numpy(float)
    g = lambda c: d[c].to_numpy(float)
    dm, nl = g('Demand Met (MW)'), g('Net Demand Met (MW)')
    wind, solar = g('Wind (MW)'), g('Solar (MW)')
    thermal, freq = g('Thermal (MW)'), g('Frequency (Hz)')
    chg, dis = g('Storage Demand (MW)'), g('Storage Discharge (MW)')

    to_mu = lambda mw: float(np.nansum(mw) * BLOCK_H / 1000)
    ramp = lambda x, k: (np.nanmax(x[k:] - x[:-k]), np.nanmin(x[k:] - x[:-k]))
    i = lambda f, x: int(f(x))

    e_dm, e_vre = to_mu(dm), to_mu(wind + solar)
    up1, dn1 = ramp(nl, 4)
    up3, _   = ramp(nl, 12)
    in_band = np.mean((freq >= FREQ_BAND[0]) & (freq <= FREQ_BAND[1])) * 100

    return [
        ("Requirement, Met + Shortage, PSP (MU)",          requirement_mu),
        ("Energy met, integrated from 15-min data (MU)",   e_dm),
        ("Peak demand met (GW)",                           dm.max() / 1000),
        ("  time of peak demand",                          _clock(h[i(np.nanargmax, dm)])),
        ("Minimum demand met (GW)",                        dm.min() / 1000),
        ("  time of minimum demand",                       _clock(h[i(np.nanargmin, dm)])),
        ("Load factor, mean / peak demand (%)",            np.nanmean(dm) / np.nanmax(dm) * 100),
        ("Peak net load (GW)",                             nl.max() / 1000),
        ("  time of peak net load",                        _clock(h[i(np.nanargmax, nl)])),
        ("Minimum net load (GW)",                          nl.min() / 1000),
        ("  time of minimum net load",                     _clock(h[i(np.nanargmin, nl)])),
        ("Max 1-hour net-load up-ramp (GW)",               up1 / 1000),
        ("Max 1-hour net-load down-ramp (GW)",             dn1 / 1000),
        ("Max 3-hour net-load up-ramp (GW)",               up3 / 1000),
        ("Peak solar (GW)",                                solar.max() / 1000),
        ("  time of peak solar",                           _clock(h[i(np.nanargmax, solar)])),
        ("Solar energy (MU)",                              to_mu(solar)),
        ("Wind energy (MU)",                               to_mu(wind)),
        ("Wind + solar share of energy met (%)",           e_vre / e_dm * 100),
        ("Max instantaneous wind + solar share (%)",       np.nanmax((wind + solar) / dm) * 100),
        ("Thermal, minimum (GW)",                          thermal.min() / 1000),
        ("Thermal, maximum (GW)",                          thermal.max() / 1000),
        ("Storage charging energy (MU)",                  "n/a" if np.isnan(chg).all() else to_mu(chg)),
        ("Storage discharge energy (MU)",                 "n/a" if np.isnan(dis).all() else to_mu(dis)),
        ("Frequency, minimum (Hz)",                        freq.min()),
        ("Frequency, maximum (Hz)",                        freq.max()),
        (f"Blocks inside {FREQ_BAND[0]:.2f}-{FREQ_BAND[1]:.2f} Hz (%)", in_band),
    ]


def build_summary(picks, days):
    m_peak = day_metrics(days['peak'],   picks['peak'].value_mu)
    m_low  = day_metrics(days['lowest'], picks['lowest'].value_mu)
    rows = []
    for (name, v1), (_, v2) in zip(m_peak, m_low):
        num = isinstance(v1, (int, float)) and isinstance(v2, (int, float))
        rows.append({
            "Metric": name,
            f"Peak day ({picks['peak'].label})": v1,
            f"Lowest day ({picks['lowest'].label})": v2,
            "Peak - Lowest": (v1 - v2) if num else "",
        })
    return pd.DataFrame(rows)


# --- Charts ---

def plot_sourcewise_gen(picks, days, y, m):
    st = STYLE_GEN
    s1, s2 = _gen_series(days['peak']), _gen_series(days['lowest'])
    stack_total = lambda s: np.array([s[l] for l, _, _ in SOURCES]).sum(axis=0)
    ymax = max(stack_total(s1).max(), stack_total(s2).max()) / 1000
    chg_max = max(np.abs(s1['Storage charging']).max(), np.abs(s2['Storage charging']).max()) / 1000
    ymin = -chg_max * 1.15 if chg_max > 0 else 0

    fig, axes = plt.subplots(1, 2, figsize=(24, 12), facecolor='white', sharey=True)
    fig.subplots_adjust(left=0.05, right=0.99, top=0.78, bottom=0.11, wspace=0.06)
    _draw_gen_panel(axes[0], days['peak']['Hour'].to_numpy(float), s1, True, st)
    _draw_gen_panel(axes[1], days['lowest']['Hour'].to_numpy(float), s2, False, st)
    for ax, kind in zip(axes, ('peak', 'lowest')):
        ax.set_title(picks[kind].panel_title, fontsize=st.label, color=C_TEXT, pad=10)
    axes[0].set_ylim(ymin, ymax * 1.08)

    add_title(fig, f"15-minute Source-wise Electricity Generation, All India: "
                   f"{picks['peak'].label} vs {picks['lowest'].label}", st)
    handles, labels = axes[0].get_legend_handles_labels()
    has_storage = any(days[k][c].notna().any() for k in days
                      for c in ('Storage Demand (MW)', 'Storage Discharge (MW)'))
    if not has_storage:        # layouts without storage columns (e.g. April 2026)
        keep = [i for i, lab in enumerate(labels) if not lab.startswith('Storage')]
        handles, labels = [handles[i] for i in keep], [labels[i] for i in keep]
    add_fig_legend(fig, handles, labels, ncol=6, y=0.855, st=st)
    return save_hourly_chart(fig, 'psp15_sourcewise_gen', y, m)


# --- Orchestrator ---

def plot_extreme_days_15min(path):
    """Load the psp_15min.py workbook, print + save the summary table, draw all charts."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"{path} not found - run build_extreme_days_15min first.")

    picks, days = load_extreme_days(path)
    y, m = picks['peak'].day.year, picks['peak'].day.month
    month_name = date(y, m, 1).strftime('%B')

    summary = build_summary(picks, days)
    with pd.option_context('display.width', 200, 'display.max_colwidth', 60,
                           'display.float_format', '{:,.2f}'.format):
        print(f"\n{'='*60}\n  15-minute comparison: {picks['peak'].label} (peak) vs "
              f"{picks['lowest'].label} (lowest)\n{'='*60}")
        print(summary.to_string(index=False))

    out_xlsx = graphs_month_folder(y, m) / f"summary_psp15_{month_name}_{y}.xlsx"
    summary.to_excel(out_xlsx, index=False)
    print(f"Summary saved: {out_xlsx}")

    plot_sourcewise_gen(picks, days, y, m)
    return summary
