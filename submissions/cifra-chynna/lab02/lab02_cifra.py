"""
NM-LAB-08252026  |  Laboratory Activity 02 - Fitting a curve to the dam
lab02_cifra.py

Everything the dashboard shows is computed here. lab02_cifra.html only
displays the numbers this script writes into it (Rule 0: Python computes,
HTML displays - no fitting, regression or statistics happens in the browser).

Pipeline:
  1. Load the stage log, put time on a float-hours axis.
  2. Forward/backward/central finite differences -> dh/dt, d2h/dt2.
  3. Fit a continuous h(t) with Levenberg-Marquardt (scipy.optimize.curve_fit).
  4. Full statistical evaluation of the fit (SSE, SST, R^2, s, SE/t/p per
     parameter) and a residual reading (drift, runs, spread, logger resolution).
  5. Integrate the fitted curve with scipy.integrate.quad, cross-check with
     np.trapezoid on the raw log.
  6. Render lab02_cifra.html: one self-contained file, vanilla-JS charts,
     no CDN dependency, no build step.
"""

import json
import os
import numpy as np
import pandas as pd
from scipy.optimize import curve_fit
from scipy import stats
from scipy.integrate import quad

# --------------------------------------------------------------------------
# 0. Paths
# --------------------------------------------------------------------------
HERE = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(HERE, "Data 01(Sensor Log).csv")
HTML_PATH = os.path.join(HERE, "lab02_cifra.html")
LOGGER_RESOLUTION_M = 0.01  # the logger rounds to the nearest centimetre

# --------------------------------------------------------------------------
# 1. Load the stage log and build a clean numeric time axis
# --------------------------------------------------------------------------
# The logger sheet has three banner rows before the header. The Date/Time
# split columns contain one corrupted row (reading 10, where the logger's
# own export briefly wrote the literal header strings "Date"/"Time" into
# the data cells) - the combined Timestamp column and the Depth column are
# intact throughout, so those are the two columns this script actually
# trusts.
raw = pd.read_csv(CSV_PATH, skiprows=3)
raw["ts"] = pd.to_datetime(raw["Timestamp"])
raw = raw.sort_values("ts").reset_index(drop=True)

corrupted_rows = raw.index[raw["Date"].astype(str) == "Date"].tolist()

t0_stamp = raw["ts"].iloc[0]
t = ((raw["ts"] - t0_stamp).dt.total_seconds() / 3600.0).to_numpy()  # hours
h = raw["Depth (m)"].to_numpy(dtype=float)
n = len(t)
timestamps_iso = raw["ts"].dt.strftime("%Y-%m-%d %H:%M").tolist()

step_h = float(np.median(np.diff(t)))  # 0.25 h, kept honest by construction

# --------------------------------------------------------------------------
# 2. Finite differences: forward/backward at the ends, central in between
# --------------------------------------------------------------------------
dh = np.empty(n)
dh[0] = (h[1] - h[0]) / (t[1] - t[0])           # forward at the start
dh[-1] = (h[-1] - h[-2]) / (t[-1] - t[-2])       # backward at the end
dh[1:-1] = (h[2:] - h[:-2]) / (t[2:] - t[:-2])   # central everywhere else

d2h = np.empty(n)
d2h[0] = (dh[1] - dh[0]) / (t[1] - t[0])
d2h[-1] = (dh[-1] - dh[-2]) / (t[-1] - t[-2])
d2h[1:-1] = (dh[2:] - dh[:-2]) / (t[2:] - t[:-2])

i_max_dhdt = int(np.argmax(dh))
t_max_dhdt = float(t[i_max_dhdt])
val_max_dhdt = float(dh[i_max_dhdt])

i_d2h_max = int(np.argmax(d2h))
i_d2h_min = int(np.argmin(d2h))
second_derivative_note = (
    f"d2h/dt2 climbs to its sharpest positive value (+{d2h[i_d2h_max]:.3f} m/h^2) "
    f"at t = {t[i_d2h_max]:.2f} h, just before the fastest filling rate at "
    f"t = {t_max_dhdt:.2f} h, then swings to its most negative value "
    f"({d2h[i_d2h_min]:.3f} m/h^2) at t = {t[i_d2h_min]:.2f} h a few hours later - "
    "the inflow was switched on and throttled off abruptly rather than "
    "ramping and easing smoothly, and the sign change alone (at the inflection "
    "of h) would have said nothing about how abrupt either edge was."
)

# --------------------------------------------------------------------------
# 3. Choose and fit a model
# --------------------------------------------------------------------------
# The raw log does not settle toward a single ceiling: it holds a slow, noisy
# baseline near 14.2-14.5 m, rises sharply to a peak near 21.1 m over about
# ten hours, then eases back down to a lower plateau near 19.5 m and holds
# there. That rise-then-partial-drawdown-then-plateau shape has two
# inflections of opposite sense, which neither a single four-parameter
# logistic nor a Gompertz can produce (both are monotonic). A sum of two
# logistics, one with a positive amplitude for the fill and one with a
# negative amplitude for the subsequent drawdown/spillway release, matches
# the shape with parameters that still mean something physically: a fill
# time, a fill rate, a release time and a release rate.
MODEL_NAME = "Sum of two logistics"
MODEL_ASSUMPTIONS = (
    "Assumes one filling pulse into the reservoir followed by one release or "
    "drawdown pulse, each with its own timing and rate, settling onto a "
    "second, lower plateau rather than a single ceiling."
)
MODEL_EQUATION = "h(t) = c + a1 / (1 + exp(-k1 (t - t01))) + a2 / (1 + exp(-k2 (t - t02)))"
PARAM_NAMES = ["c", "a1", "k1", "t01", "a2", "k2", "t02"]
PARAM_LABELS = {
    "c": "baseline level, c (m)",
    "a1": "fill amplitude, a1 (m)",
    "k1": "fill rate, k1 (1/h)",
    "t01": "fill midpoint time, t01 (h)",
    "a2": "release amplitude, a2 (m)",
    "k2": "release rate, k2 (1/h)",
    "t02": "release midpoint time, t02 (h)",
}


def sigmoid(x, k, x0):
    return 1.0 / (1.0 + np.exp(-k * (x - x0)))


def model(x, c, a1, k1, t01, a2, k2, t02):
    return c + a1 * sigmoid(x, k1, t01) + a2 * sigmoid(x, k2, t02)


# Initial guess read off the raw record itself (not left at the default of
# ones - Levenberg-Marquardt takes no bounds, so p0 has to carry everything
# a bound would otherwise express):
#   c0        - median of the first 20 readings, before the fill starts
#   peak/tp   - the observed maximum and when it occurs
#   final     - median of the last 20 readings, the settled plateau
#   t01       - time the record first crosses halfway between c0 and the peak
#   t02       - time the record first falls back to halfway between the peak
#               and the final plateau
c0 = float(np.median(h[:20]))
peak_val = float(np.max(h))
i_peak = int(np.argmax(h))
final_level = float(np.median(h[-20:]))

half_rise = (c0 + peak_val) / 2.0
i_t01 = int(np.argmax(h >= half_rise))
t01_guess = float(t[i_t01])

half_fall = (peak_val + final_level) / 2.0
post_peak = h[i_peak:]
i_t02_rel = int(np.argmax(post_peak <= half_fall))
t02_guess = float(t[i_peak + i_t02_rel])

p0 = [c0, peak_val - c0, 1.0, t01_guess, final_level - peak_val, 0.3, t02_guess]
p0_rationale = (
    f"c0={c0:.2f} m (median of the first 20 readings, before the fill), "
    f"a1={peak_val - c0:.2f} m (observed peak {peak_val:.2f} m minus c0), "
    f"k1=1.0 /h (rise completes in roughly 8-10 h on the raw plot), "
    f"t01={t01_guess:.2f} h (first crossing of the rise's halfway level), "
    f"a2={final_level - peak_val:.2f} m (settled plateau {final_level:.2f} m "
    f"minus the peak), k2=0.3 /h (the drawdown is visibly slower than the fill), "
    f"t02={t02_guess:.2f} h (first crossing of the drawdown's halfway level)."
)

popt, pcov = curve_fit(model, t, h, p0=p0, method="lm", maxfev=20000)
resid = h - model(t, *popt)

p = len(popt)
dof = n - p
sse = float(np.sum(resid ** 2))
sst = float(np.sum((h - h.mean()) ** 2))
r2 = 1.0 - sse / sst
s_est = float(np.sqrt(sse / dof))

pcov_has_inf = bool(np.any(np.isinf(pcov)) or np.any(np.isnan(pcov)))
se = np.sqrt(np.diag(pcov))
tstat = popt / se
pval = 2.0 * (1.0 - stats.t.cdf(np.abs(tstat), dof))

params_table = [
    {
        "name": PARAM_NAMES[i],
        "label": PARAM_LABELS[PARAM_NAMES[i]],
        "value": float(popt[i]),
        "se": float(se[i]),
        "t": float(tstat[i]),
        "p": float(pval[i]),
    }
    for i in range(p)
]

# --------------------------------------------------------------------------
# 4. Read the residuals, don't just plot them
# --------------------------------------------------------------------------
half = n // 2
mean_resid_first = float(np.mean(resid[:half]))
mean_resid_second = float(np.mean(resid[half:]))
drift_mag = abs(mean_resid_first - mean_resid_second)
drifts = drift_mag > 0.25 * float(np.std(resid))
drift_note = (
    f"The mean residual is {mean_resid_first:+.3f} m over the first half of the "
    f"record and {mean_resid_second:+.3f} m over the second half"
    + (
        " - a real drift, so some of the shape is still unexplained."
        if drifts
        else " - close enough to call the scatter centred on zero throughout."
    )
)

signs = np.sign(resid)
signs[signs == 0] = 1
run_len = 1
max_run = 1
run_start = 0
max_run_start_t = float(t[0])
for i in range(1, n):
    if signs[i] == signs[i - 1]:
        run_len += 1
        if run_len > max_run:
            max_run = run_len
            max_run_start_t = float(t[i - run_len + 1])
    else:
        run_len = 1
runs_note = (
    f"The longest unbroken run of same-signed residuals is {max_run} points "
    f"(out of {n}), starting near t = {max_run_start_t:.2f} h - a run that long "
    "means the model's shape is missing something systematic there (most "
    "likely a small diurnal wobble riding on the fill/release trend), not "
    "that the data is simply noisy."
)

fitted_vals = model(t, *popt)
spread_corr = float(np.corrcoef(np.abs(resid), fitted_vals)[0, 1])
spread_note = (
    f"Correlation of |residual| with the fitted level is {spread_corr:+.2f}: "
    + (
        "the scatter widens as the level rises."
        if spread_corr > 0.15
        else "the scatter narrows as the level rises."
        if spread_corr < -0.15
        else "the scatter is roughly flat across the range of levels."
    )
)

max_abs_resid = float(np.max(np.abs(resid)))
resid_ratio = max_abs_resid / LOGGER_RESOLUTION_M
resid_resolution_note = (
    f"The largest residual is {max_abs_resid:.3f} m, about {resid_ratio:.1f}x the "
    f"logger's own {LOGGER_RESOLUTION_M:.2f} m resolution - bigger than rounding "
    "alone can explain, consistent with the unmodelled wobble the long "
    "residual run already pointed to."
)

fit_verdict = (
    f"R^2 = {r2:.4f} on top of a {max_run}-point residual run is a fit that "
    "explains the fill/release trend very well but is not a perfect account of "
    "the record: a small, systematic wobble survives the two-logistic model."
)

# --------------------------------------------------------------------------
# 5. Integrate the fitted level
# --------------------------------------------------------------------------
area_value, area_abserr = quad(lambda x: model(x, *popt), t[0], t[-1])
area_value = float(area_value)
area_abserr = float(area_abserr)
trapezoid_value = float(np.trapezoid(h, t))
area_gap = area_value - trapezoid_value
area_gap_note = (
    f"quad integrates the smooth fitted curve while trapezoid sums straight "
    f"segments between the noisy raw points, so the {area_gap:+.3f} m*h gap is "
    "the area of the residual wobble itself, not an error in either method."
)
area_meaning_note = (
    "This is meters times hours - a measure of how high and how long the "
    "reservoir stood above the zero datum over the day, useful for comparing "
    "one day's stage record to another. It is not a volume: turning it into "
    "one would need the reservoir's stage-to-area or stage-to-volume curve, "
    "which this log does not contain."
)

# --------------------------------------------------------------------------
# 6. Dense curve for plotting, and package everything for the dashboard
# --------------------------------------------------------------------------
t_dense = np.linspace(t[0], t[-1], 600)
h_dense = model(t_dense, *popt)

data_notes = []
if corrupted_rows:
    data_notes.append(
        f"Reading {int(raw.loc[corrupted_rows[0], 'Reading'])}'s split Date/Time "
        "columns were corrupted in the source export (literal header text in "
        "the data cells); its Timestamp and Depth values were intact and used "
        "as logged."
    )

payload = {
    "meta": {
        "course": "Numerical Methods, BES6-M",
        "dataset": "Reservoir stage log, 15-minute sampling",
        "n": n,
        "step_h": step_h,
        "t_start": timestamps_iso[0],
        "t_end": timestamps_iso[-1],
        "data_notes": data_notes,
    },
    "raw": {"t": t.tolist(), "h": h.tolist(), "ts": timestamps_iso},
    "deriv": {
        "t": t.tolist(),
        "dh": dh.tolist(),
        "d2h": d2h.tolist(),
        "max_dhdt": {"t": t_max_dhdt, "value": val_max_dhdt},
        "second_derivative_note": second_derivative_note,
    },
    "model": {
        "name": MODEL_NAME,
        "assumptions": MODEL_ASSUMPTIONS,
        "equation": MODEL_EQUATION,
        "p0": p0,
        "p0_rationale": p0_rationale,
        "params": params_table,
        "pcov_has_inf": pcov_has_inf,
        "t_fit": t_dense.tolist(),
        "h_fit": h_dense.tolist(),
    },
    "residuals": {
        "t": t.tolist(),
        "e": resid.tolist(),
        "fitted": fitted_vals.tolist(),
        "drift_note": drift_note,
        "runs_note": runs_note,
        "spread_note": spread_note,
        "resolution_note": resid_resolution_note,
        "verdict": fit_verdict,
        "max_run": max_run,
        "max_abs_resid": max_abs_resid,
        "logger_resolution": LOGGER_RESOLUTION_M,
    },
    "stats": {
        "n": n,
        "p": p,
        "dof": dof,
        "sse": sse,
        "sst": sst,
        "r2": r2,
        "s": s_est,
    },
    "area": {
        "t0": float(t[0]),
        "tn": float(t[-1]),
        "value": area_value,
        "abserr": area_abserr,
        "units": "m*h",
        "trapezoid": trapezoid_value,
        "gap": area_gap,
        "gap_note": area_gap_note,
        "meaning_note": area_meaning_note,
    },
}

print("Fit converged. R^2 =", round(r2, 4), " s =", round(s_est, 4), "m  n =", n, " p =", p)
print("Max dh/dt =", round(val_max_dhdt, 3), "m/h at t =", round(t_max_dhdt, 2), "h")
print("Area (fitted, quad) =", round(area_value, 3), "+/-", round(area_abserr, 6), "m*h")
print("Area (raw, trapezoid) =", round(trapezoid_value, 3), "m*h  gap =", round(area_gap, 4))
if data_notes:
    print("Data note:", data_notes[0])

# --------------------------------------------------------------------------
# 7. Render the dashboard
# --------------------------------------------------------------------------
DATA_JSON = json.dumps(payload)

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Fitting a curve to the dam &mdash; NM-LAB-08252026</title>
<style>
  :root{
    --navy:#12345a;
    --navy-2:#1d527f;
    --blue:#2d72ad;
    --blue-bright:#5aa6dc;
    --sky:#9bc9e9;
    --sky-soft:#dceefa;
    --page-bg:#eaf5fc;
    --white:#ffffff;
    --panel:#ffffff;
    --panel-soft:#f5faff;
    --ink:#203447;
    --ink-soft:#62788c;
    --ink-faint:#8ca0b0;
    --gray:#6b7885;
    --gray-soft:#eef2f5;
    --border:#d2e2ee;
    --border-strong:#b9d1e3;
    --shadow:0 8px 22px rgba(18,52,90,.10), 0 2px 5px rgba(18,52,90,.06);
    --shadow-deep:0 16px 36px rgba(18,52,90,.15), 0 4px 9px rgba(18,52,90,.08);
    --inset:inset 0 1px 0 rgba(255,255,255,.9);
    --mono:'IBM Plex Mono','SFMono-Regular',Consolas,monospace;
    --serif:'Iowan Old Style','Palatino Linotype',Palatino,Georgia,serif;
    --sans:'Inter','Helvetica Neue',Arial,sans-serif;
  }

  *{box-sizing:border-box}
  html{scroll-behavior:smooth}
  body{
    margin:0;
    color:var(--ink);
    font-family:var(--sans);
    -webkit-font-smoothing:antialiased;
    background-color:var(--page-bg);
    background-image:
      radial-gradient(circle at 12% 12%, rgba(90,166,220,.15) 0 2px, transparent 2.5px),
      radial-gradient(circle at 88% 24%, rgba(45,114,173,.09) 0 1.5px, transparent 2px),
      linear-gradient(rgba(255,255,255,.45) 1px, transparent 1px),
      linear-gradient(90deg, rgba(255,255,255,.45) 1px, transparent 1px),
      radial-gradient(900px 500px at 10% -5%, #fff 0%, transparent 65%),
      linear-gradient(135deg,#e8f4fc 0%,#f4f9fd 48%,#e5f1fa 100%);
    background-size:24px 24px,34px 34px,48px 48px,48px 48px,auto,auto;
    background-attachment:fixed;
  }

  body::before{
    content:"";
    position:fixed;
    inset:0;
    pointer-events:none;
    background:
      linear-gradient(120deg,transparent 0 42%,rgba(90,166,220,.035) 42.2% 42.5%,transparent 42.7%),
      linear-gradient(300deg,transparent 0 58%,rgba(18,52,90,.025) 58.2% 58.5%,transparent 58.7%);
    z-index:-1;
  }

  header.top{
    position:relative;
    overflow:hidden;
    color:#eef7ff;
    padding:36px 5vw 30px;
    background:
      radial-gradient(600px 180px at 85% 0%,rgba(90,166,220,.28),transparent 70%),
      linear-gradient(125deg,#0d2847 0%,#123b63 42%,#1d527f 72%,#2d72ad 100%);
    box-shadow:0 14px 35px rgba(18,52,90,.25);
    border-bottom:1px solid rgba(255,255,255,.12);
  }

  header.top::before{
    content:"";
    position:absolute;
    width:340px;height:340px;
    right:-120px;top:-190px;
    border:1px solid rgba(255,255,255,.12);
    border-radius:50%;
    box-shadow:0 0 0 30px rgba(255,255,255,.025),0 0 0 60px rgba(255,255,255,.018);
  }

  header.top::after{
    content:"";
    position:absolute;
    left:0;right:0;bottom:0;height:4px;
    background:linear-gradient(90deg,transparent,var(--sky),var(--blue-bright),var(--sky),transparent);
  }

  header.top .eyebrow{
    font-family:var(--mono);
    font-size:11px;
    letter-spacing:.14em;
    text-transform:uppercase;
    color:#b8d8ee;
    display:flex;
    justify-content:space-between;
    flex-wrap:wrap;
    gap:8px;
  }

  header.top h1{
    font-family:var(--serif);
    font-weight:700;
    font-size:clamp(30px,4vw,46px);
    margin:9px 0 5px;
    letter-spacing:.2px;
    text-shadow:0 3px 12px rgba(0,0,0,.2);
  }

  header.top .sub{
    font-family:var(--mono);
    font-size:12px;
    color:#c9e1f3;
    letter-spacing:.03em;
  }

  .meta-strip{
    display:flex;
    flex-wrap:wrap;
    gap:12px;
    margin-top:20px;
  }

  .meta-strip > div{
    min-width:135px;
    font-size:12px;
    color:#dcecf9;
    font-family:var(--mono);
    background:linear-gradient(145deg,rgba(255,255,255,.12),rgba(255,255,255,.035));
    border:1px solid rgba(255,255,255,.16);
    border-radius:10px;
    padding:9px 14px;
    box-shadow:inset 0 1px 0 rgba(255,255,255,.12),0 7px 16px rgba(0,0,0,.10);
    backdrop-filter:blur(4px);
  }

  .meta-strip b{
    color:#fff;
    display:block;
    font-size:11px;
    margin-bottom:4px;
    font-family:var(--sans);
    font-weight:700;
    letter-spacing:.05em;
    text-transform:uppercase;
  }

  main{
    padding:0 5vw 70px;
    max-width:1240px;
    margin:0 auto;
  }

  section.panel{
    position:relative;
    overflow:hidden;
    background:
      linear-gradient(180deg,rgba(255,255,255,.98),rgba(245,250,255,.98));
    border:1px solid var(--border);
    border-radius:16px;
    margin-top:28px;
    padding:30px;
    box-shadow:var(--shadow),var(--inset);
    transition:box-shadow .18s ease,transform .18s ease;
  }

  section.panel:hover{
    transform:translateY(-1px);
    box-shadow:var(--shadow-deep),var(--inset);
  }

  section.panel::before{
    content:"";
    position:absolute;
    top:0;left:0;right:0;height:4px;
    background:linear-gradient(90deg,var(--navy),var(--blue),var(--blue-bright),var(--sky-soft));
  }

  section.panel::after{
    content:"";
    position:absolute;
    width:160px;height:160px;
    right:-90px;bottom:-100px;
    border-radius:50%;
    background:radial-gradient(circle,rgba(155,201,233,.20),transparent 68%);
    pointer-events:none;
  }

  section.panel h2{
    font-family:var(--serif);
    font-weight:700;
    font-size:24px;
    margin:0 0 7px;
    color:var(--navy);
  }

  section.panel p.lede{
    color:var(--ink-soft);
    font-size:13.5px;
    margin:0 0 19px;
    max-width:75ch;
    line-height:1.6;
  }

  .tabs{
    position:sticky;
    top:0;
    z-index:20;
    display:flex;
    gap:5px;
    margin-top:28px;
    border:1px solid var(--border);
    border-radius:13px;
    background:rgba(255,255,255,.94);
    padding:6px;
    box-shadow:0 8px 22px rgba(18,52,90,.11),inset 0 1px 0 #fff;
    backdrop-filter:blur(10px);
  }

  .tab-btn{
    flex:1;
    font-family:var(--mono);
    font-size:11px;
    letter-spacing:.05em;
    text-transform:uppercase;
    background:linear-gradient(180deg,#fff,#f3f8fc);
    border:1px solid transparent;
    border-radius:9px;
    padding:13px 18px;
    cursor:pointer;
    color:var(--ink-soft);
    font-weight:600;
    transition:background .15s,color .15s,box-shadow .15s,transform .15s;
  }

  .tab-btn:hover{
    color:var(--navy);
    background:linear-gradient(180deg,#fff,var(--sky-soft));
  }

  .tab-btn:active{transform:translateY(1px)}

  .tab-btn.active{
    color:#fff;
    background:linear-gradient(145deg,var(--navy),var(--blue));
    border-color:#1d527f;
    box-shadow:0 5px 12px rgba(18,52,90,.22),inset 0 1px 0 rgba(255,255,255,.18);
  }

  .tab-panel{display:none}
  .tab-panel.active{display:block}

  .chart-wrap{
    position:relative;
    width:100%;
    margin-top:11px;
    border:1px solid var(--border);
    border-radius:12px;
    background:
      linear-gradient(rgba(45,114,173,.035) 1px,transparent 1px),
      linear-gradient(90deg,rgba(45,114,173,.035) 1px,transparent 1px),
      linear-gradient(180deg,#fff,#f4f9fd);
    background-size:32px 32px,32px 32px,auto;
    padding:10px 10px 2px;
    box-shadow:inset 0 1px 0 #fff,0 5px 14px rgba(18,52,90,.07);
  }

  svg.chart{
    width:100%;
    height:auto;
    display:block;
    font-family:var(--mono);
  }

  .grid-line{stroke:#dce8f2;stroke-width:1}
  .axis-line{stroke:#8194a4;stroke-width:1}
  .axis-label{fill:#62788c;font-size:10.5px}
  .curve{
    fill:none;
    stroke-width:2;
    filter:drop-shadow(0 2px 2px rgba(18,52,90,.18));
  }

  .zero-line{stroke:#788692;stroke-width:1;stroke-dasharray:4 3}
  .marker-line{stroke:#2d72ad;stroke-width:1.4;stroke-dasharray:2 3}

  .tooltip{
    position:absolute;
    pointer-events:none;
    background:linear-gradient(145deg,#12345a,#1d527f);
    color:#eef7ff;
    font-family:var(--mono);
    font-size:11px;
    padding:8px 11px;
    border-radius:8px;
    border:1px solid #5aa6dc;
    transform:translate(-50%,-115%);
    white-space:nowrap;
    box-shadow:0 10px 24px rgba(18,52,90,.35),inset 0 1px 0 rgba(255,255,255,.15);
    opacity:0;
    transition:opacity .08s;
    z-index:5;
  }

  .hover-dot{
    fill:#12345a;
    stroke:#fff;
    stroke-width:1.5;
    opacity:0;
    filter:drop-shadow(0 2px 3px rgba(18,52,90,.4));
  }

  .legend{
    display:flex;
    gap:18px;
    flex-wrap:wrap;
    font-family:var(--mono);
    font-size:11px;
    margin-top:9px;
    color:var(--ink-soft);
  }

  .legend span{display:inline-flex;align-items:center;gap:6px}
  .swatch{width:14px;height:3px;display:inline-block;border-radius:2px}

  .grid2{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:22px;
    margin-top:8px;
  }

  @media(max-width:840px){
    .grid2{grid-template-columns:1fr}
    .tabs{position:static}
  }

  @media(max-width:640px){
    section.panel{padding:22px 18px;border-radius:13px}
    main{padding-left:3vw;padding-right:3vw}
    .tab-btn{padding:11px 8px;font-size:9px}
    .meta-strip > div{flex:1}
  }

  table.data-tbl{
    width:100%;
    border-collapse:separate;
    border-spacing:0;
    margin-top:16px;
    font-size:13px;
    border:1px solid var(--border);
    border-radius:12px;
    overflow:hidden;
    box-shadow:var(--shadow);
  }

  table.data-tbl caption{
    text-align:left;
    font-family:var(--mono);
    font-size:10.5px;
    text-transform:uppercase;
    letter-spacing:.06em;
    color:var(--ink-faint);
    margin-bottom:8px;
  }

  table.data-tbl th{
    text-align:left;
    font-family:var(--mono);
    font-size:10.5px;
    text-transform:uppercase;
    letter-spacing:.04em;
    color:var(--navy);
    font-weight:700;
    background:linear-gradient(180deg,#e0effa,#f7fbfe);
    border-bottom:2px solid var(--blue-bright);
    padding:11px 12px;
  }

  table.data-tbl td{
    padding:9px 12px;
    border-bottom:1px solid var(--border);
    font-family:var(--mono);
    color:var(--ink);
  }

  table.data-tbl tbody tr:nth-child(even){background:#f7fbfe}
  table.data-tbl tbody tr:hover{background:var(--sky-soft)}
  table.data-tbl td.sig{color:var(--blue);font-weight:700}
  table.data-tbl td.nosig{color:var(--gray)}

  .stat-grid{
    display:grid;
    grid-template-columns:repeat(auto-fit,minmax(155px,1fr));
    gap:15px;
    margin-top:17px;
  }

  .stat-card{
    position:relative;
    overflow:hidden;
    border:1px solid var(--border);
    border-radius:12px;
    padding:16px 17px;
    background:
      radial-gradient(100px 70px at 100% 0%,rgba(155,201,233,.24),transparent 75%),
      linear-gradient(145deg,#fff 0%,#f5faff 72%,#e8f4fc 100%);
    box-shadow:var(--shadow),var(--inset);
    transition:box-shadow .15s ease,transform .15s ease;
  }

  .stat-card::before{
    content:"";
    position:absolute;
    top:0;left:0;right:0;height:2px;
    background:linear-gradient(90deg,var(--blue),var(--sky));
  }

  .stat-card:hover{
    transform:translateY(-2px);
    box-shadow:var(--shadow-deep),var(--inset);
  }

  .stat-card .k{
    font-family:var(--mono);
    font-size:10px;
    text-transform:uppercase;
    letter-spacing:.06em;
    color:var(--ink-faint);
    font-weight:700;
  }

  .stat-card .v{
    font-family:var(--serif);
    font-weight:700;
    font-size:26px;
    color:var(--navy);
    margin-top:6px;
  }

  .stat-card .v small{
    font-family:var(--sans);
    font-weight:400;
    font-size:12px;
    color:var(--ink-soft);
  }

  .note{
    background:linear-gradient(135deg,#e0f0fb,#fff);
    border:1px solid var(--border);
    border-left:4px solid var(--blue);
    border-radius:9px;
    padding:12px 16px;
    font-size:13.5px;
    line-height:1.58;
    margin-top:13px;
    color:#254661;
    box-shadow:var(--shadow);
  }

  .note.warn{
    background:linear-gradient(135deg,#eef2f5,#fff);
    border-left-color:var(--gray);
    color:#394651;
    font-weight:500;
  }

  .eq{
    font-family:var(--mono);
    background:
      linear-gradient(145deg,#102f50,#1d527f);
    color:#dcefff;
    padding:12px 16px;
    display:inline-block;
    font-size:13px;
    margin-top:10px;
    border-radius:9px;
    border:1px solid #5aa6dc;
    box-shadow:0 7px 16px rgba(18,52,90,.20),inset 0 1px 0 rgba(255,255,255,.12);
  }

  .verdict{
    margin-top:18px;
    font-size:14px;
    line-height:1.58;
    padding:15px 17px;
    background:linear-gradient(135deg,#dceefa,#fff);
    border:1px solid var(--blue-bright);
    border-radius:9px;
    color:#173a5e;
    box-shadow:var(--shadow);
  }

  footer{
    padding:30px 5vw 45px;
    font-family:var(--mono);
    font-size:10.5px;
    color:var(--ink-faint);
    text-align:center;
    border-top:1px solid var(--border);
    margin-top:24px;
    background:
      linear-gradient(rgba(45,114,173,.025) 1px,transparent 1px),
      linear-gradient(90deg,rgba(45,114,173,.025) 1px,transparent 1px),
      linear-gradient(180deg,#fff,#edf6fc);
    background-size:30px 30px,30px 30px,auto;
  }
</style>
</head>
<body>

<header class="top">
  <div class="eyebrow"><span>NM-LAB-08252026 &middot; Laboratory Activity 02</span><span id="meta-course"></span></div>
  <h1>Fitting a curve to the dam</h1>
  <div class="sub">Levenberg&ndash;Marquardt, residuals, and the area under the level</div>
  <div class="meta-strip">
    <div><b>Dataset</b><span id="meta-dataset"></span></div>
    <div><b>Readings</b><span id="meta-n"></span></div>
    <div><b>Sampling</b><span id="meta-step"></span></div>
    <div><b>Window</b><span id="meta-window"></span></div>
  </div>
</header>

<main>

  <section class="panel header-chart">
    <h2>Stage log</h2>
    <p class="lede">Raw logged depth against time, every fifteen minutes, unfiltered.</p>
    <div class="chart-wrap" id="chart-raw"></div>
    <div id="data-notes"></div>
  </section>

  <div class="tabs">
    <button class="tab-btn active" data-tab="t1">01 &middot; Derivatives</button>
    <button class="tab-btn" data-tab="t2">02 &middot; Fit &amp; statistics</button>
    <button class="tab-btn" data-tab="t3">03 &middot; Area under the curve</button>
  </div>

  <!-- TAB 1 : DERIVATIVES -->
  <div class="tab-panel active" id="t1">
    <section class="panel">
      <h2>First derivative &mdash; dh/dt</h2>
      <p class="lede">Central differences in the interior, forward/backward at the two ends. Peak marked in blue.</p>
      <div class="chart-wrap" id="chart-dh"></div>
      <div class="stat-grid" id="dhdt-stats"></div>
    </section>
    <section class="panel">
      <h2>Second derivative &mdash; d&sup2;h/dt&sup2;</h2>
      <p class="lede">What it says about the inflow, not just where it changes sign.</p>
      <div class="chart-wrap" id="chart-d2h"></div>
      <div class="note" id="d2h-note"></div>
    </section>
  </div>

  <!-- TAB 2 : FIT & STATS -->
  <div class="tab-panel" id="t2">
    <section class="panel">
      <h2>Model</h2>
      <p class="lede" id="model-assumptions"></p>
      <div class="eq" id="model-equation"></div>
      <div class="note" id="p0-note"></div>
    </section>

    <section class="panel">
      <h2>Fitted curve over the raw log</h2>
      <div class="chart-wrap" id="chart-fit"></div>
      <div class="legend">
        <span><i class="swatch" style="background:var(--ink-soft)"></i>raw readings</span>
        <span><i class="swatch" style="background:var(--blue-bright)"></i>fitted h(t)</span>
      </div>
    </section>

    <section class="panel">
      <h2>Fitted parameters</h2>
      <table class="data-tbl" id="params-tbl">
        <caption>Standard error, t statistic and two-tailed p-value per parameter (df in footer)</caption>
        <thead><tr><th>Parameter</th><th>Value</th><th>SE</th><th>t</th><th>p</th></tr></thead>
        <tbody></tbody>
      </table>
      <div class="note" id="pcov-note" style="display:none"></div>
    </section>

    <section class="panel">
      <h2>Goodness of fit</h2>
      <div class="stat-grid" id="fit-stats"></div>
    </section>

    <section class="panel">
      <h2>Residuals</h2>
      <p class="lede">Plotted against time and against the fitted value. Read before trusted.</p>
      <div class="grid2">
        <div>
          <div class="chart-wrap" id="chart-resid-t"></div>
          <div class="legend"><span>residual vs time</span></div>
        </div>
        <div>
          <div class="chart-wrap" id="chart-resid-fit"></div>
          <div class="legend"><span>residual vs fitted h</span></div>
        </div>
      </div>
      <div class="note" id="resid-drift"></div>
      <div class="note" id="resid-runs"></div>
      <div class="note" id="resid-spread"></div>
      <div class="note warn" id="resid-resolution"></div>
      <div class="verdict" id="fit-verdict"></div>
    </section>
  </div>

  <!-- TAB 3 : AREA -->
  <div class="tab-panel" id="t3">
    <section class="panel">
      <h2>Area under the fitted level</h2>
      <p class="lede">Fitted curve, shaded to the time axis, between the first and last logged times.</p>
      <div class="chart-wrap" id="chart-area"></div>
      <div class="stat-grid" id="area-stats"></div>
      <div class="note" id="area-gap-note"></div>
      <div class="note" id="area-meaning-note"></div>
    </section>
  </div>

</main>

<footer>lab02_cifra.py &middot; scipy.optimize.curve_fit (method=lm) &middot; scipy.integrate.quad &middot; no fitting performed in this page</footer>

<script id="lab-data" type="application/json">__DATA_JSON__</script>
<script>
(function(){
  const DATA = JSON.parse(document.getElementById('lab-data').textContent);

  // ---------- meta strip ----------
  document.getElementById('meta-course').textContent = DATA.meta.course;
  document.getElementById('meta-dataset').textContent = DATA.meta.dataset;
  document.getElementById('meta-n').textContent = DATA.meta.n + ' readings';
  document.getElementById('meta-step').textContent = (DATA.meta.step_h*60).toFixed(0) + ' min';
  document.getElementById('meta-window').textContent = DATA.meta.t_start + ' \u2192 ' + DATA.meta.t_end;
  if (DATA.meta.data_notes && DATA.meta.data_notes.length){
    const d = document.createElement('div');
    d.className = 'note';
    d.style.marginTop = '14px';
    d.textContent = 'Data QA: ' + DATA.meta.data_notes.join(' ');
    document.getElementById('data-notes').appendChild(d);
  }

  // ---------- tabs ----------
  document.querySelectorAll('.tab-btn').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.tab).classList.add('active');
    });
  });

  // ==================================================================
  // Minimal SVG line/scatter chart with hover tooltip. No dependencies.
  // ==================================================================
  function niceExtent(arr){
    let lo = Math.min(...arr), hi = Math.max(...arr);
    if (lo === hi){ lo -= 1; hi += 1; }
    const pad = (hi-lo)*0.08;
    return [lo-pad, hi+pad];
  }

  function drawChart(containerId, opts){
    const container = document.getElementById(containerId);
    const W = 900, H = opts.height || 260;
    const M = {top:16, right:20, bottom:32, left:54};
    const innerW = W - M.left - M.right, innerH = H - M.top - M.bottom;

    let allX = [], allY = [];
    opts.series.forEach(s=>{ s.points.forEach(pt=>{ allX.push(pt[0]); allY.push(pt[1]); }); });
    if (opts.extraY) allY = allY.concat(opts.extraY);
    const [xlo,xhi] = niceExtent(allX);
    const [ylo,yhi] = niceExtent(allY);

    const xs = x => M.left + (x-xlo)/(xhi-xlo)*innerW;
    const ys = y => M.top + innerH - (y-ylo)/(yhi-ylo)*innerH;

    const svgns = 'http://www.w3.org/2000/svg';
    const svg = document.createElementNS(svgns,'svg');
    svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
    svg.setAttribute('class','chart');
    svg.setAttribute('preserveAspectRatio','xMidYMid meet');

    // gridlines
    const nGridY = 5;
    for(let i=0;i<=nGridY;i++){
      const gy = M.top + innerH*i/nGridY;
      const line = document.createElementNS(svgns,'line');
      line.setAttribute('x1',M.left); line.setAttribute('x2',M.left+innerW);
      line.setAttribute('y1',gy); line.setAttribute('y2',gy);
      line.setAttribute('class','grid-line');
      svg.appendChild(line);
      const val = yhi - (yhi-ylo)*i/nGridY;
      const lbl = document.createElementNS(svgns,'text');
      lbl.setAttribute('x', M.left-8); lbl.setAttribute('y', gy+3);
      lbl.setAttribute('text-anchor','end'); lbl.setAttribute('class','axis-label');
      lbl.textContent = val.toFixed(opts.yDecimals!==undefined?opts.yDecimals:2);
      svg.appendChild(lbl);
    }
    const nGridX = 6;
    for(let i=0;i<=nGridX;i++){
      const gx = M.left + innerW*i/nGridX;
      const val = xlo + (xhi-xlo)*i/nGridX;
      const lbl = document.createElementNS(svgns,'text');
      lbl.setAttribute('x', gx); lbl.setAttribute('y', M.top+innerH+18);
      lbl.setAttribute('text-anchor','middle'); lbl.setAttribute('class','axis-label');
      lbl.textContent = val.toFixed(0) + 'h';
      svg.appendChild(lbl);
    }

    // zero line
    if (opts.zeroLine && ylo < 0 && yhi > 0){
      const zl = document.createElementNS(svgns,'line');
      zl.setAttribute('x1',M.left); zl.setAttribute('x2',M.left+innerW);
      zl.setAttribute('y1',ys(0)); zl.setAttribute('y2',ys(0));
      zl.setAttribute('class','zero-line');
      svg.appendChild(zl);
    }

    // axes
    const axX = document.createElementNS(svgns,'line');
    axX.setAttribute('x1',M.left); axX.setAttribute('x2',M.left+innerW);
    axX.setAttribute('y1',M.top+innerH); axX.setAttribute('y2',M.top+innerH);
    axX.setAttribute('class','axis-line'); svg.appendChild(axX);
    const axY = document.createElementNS(svgns,'line');
    axY.setAttribute('x1',M.left); axY.setAttribute('x2',M.left);
    axY.setAttribute('y1',M.top); axY.setAttribute('y2',M.top+innerH);
    axY.setAttribute('class','axis-line'); svg.appendChild(axY);

    // fill area (single series only)
    if (opts.fillSeriesIndex !== undefined){
      const s = opts.series[opts.fillSeriesIndex];
      let d = `M ${xs(s.points[0][0])} ${ys(0)} `;
      s.points.forEach(pt=> d += `L ${xs(pt[0])} ${ys(pt[1])} `);
      d += `L ${xs(s.points[s.points.length-1][0])} ${ys(0)} Z`;
      const path = document.createElementNS(svgns,'path');
      path.setAttribute('d', d);
      path.setAttribute('fill', s.color);
      path.setAttribute('opacity', '0.16');
      path.setAttribute('stroke','none');
      svg.appendChild(path);
    }

    // series
    opts.series.forEach(s=>{
      if (s.type === 'scatter'){
        s.points.forEach(pt=>{
          const c = document.createElementNS(svgns,'circle');
          c.setAttribute('cx', xs(pt[0])); c.setAttribute('cy', ys(pt[1]));
          c.setAttribute('r', s.r || 1.6);
          c.setAttribute('fill', s.color);
          c.setAttribute('opacity', s.opacity!==undefined?s.opacity:0.75);
          svg.appendChild(c);
        });
      } else {
        let d = '';
        s.points.forEach((pt,i)=>{ d += (i===0?'M ':'L ') + xs(pt[0]) + ' ' + ys(pt[1]) + ' '; });
        const path = document.createElementNS(svgns,'path');
        path.setAttribute('d', d);
        path.setAttribute('class','curve');
        path.setAttribute('stroke', s.color);
        path.setAttribute('stroke-width', s.width || 2);
        svg.appendChild(path);
      }
    });

    // marker vertical line
    if (opts.markerX !== undefined){
      const mx = xs(opts.markerX);
      const ml = document.createElementNS(svgns,'line');
      ml.setAttribute('x1',mx); ml.setAttribute('x2',mx);
      ml.setAttribute('y1',M.top); ml.setAttribute('y2',M.top+innerH);
      ml.setAttribute('class','marker-line');
      svg.appendChild(ml);
      if (opts.markerLabel){
        const lbl = document.createElementNS(svgns,'text');
        lbl.setAttribute('x', mx+4); lbl.setAttribute('y', M.top+12);
        lbl.setAttribute('class','axis-label'); lbl.setAttribute('fill','var(--blue)');
        lbl.textContent = opts.markerLabel;
        svg.appendChild(lbl);
      }
    }

    // hover interaction: nearest point on the primary (first) series
    const hoverDot = document.createElementNS(svgns,'circle');
    hoverDot.setAttribute('r', 4); hoverDot.setAttribute('class','hover-dot');
    svg.appendChild(hoverDot);
    const hitRect = document.createElementNS(svgns,'rect');
    hitRect.setAttribute('x',M.left); hitRect.setAttribute('y',M.top);
    hitRect.setAttribute('width',innerW); hitRect.setAttribute('height',innerH);
    hitRect.setAttribute('fill','transparent');
    svg.appendChild(hitRect);

    container.style.position = 'relative';
    container.innerHTML = '';
    container.appendChild(svg);
    const tip = document.createElement('div');
    tip.className = 'tooltip';
    container.appendChild(tip);

    const primary = opts.series[opts.hoverSeriesIndex !== undefined ? opts.hoverSeriesIndex : 0].points;
    hitRect.addEventListener('mousemove', (e)=>{
      const rect = svg.getBoundingClientRect();
      const scaleX = W / rect.width;
      const mxPix = (e.clientX - rect.left) * scaleX;
      const dataX = xlo + (mxPix - M.left) / innerW * (xhi-xlo);
      let best = primary[0], bestD = Infinity;
      for (const pt of primary){
        const dd = Math.abs(pt[0]-dataX);
        if (dd < bestD){ bestD = dd; best = pt; }
      }
      hoverDot.setAttribute('cx', xs(best[0]));
      hoverDot.setAttribute('cy', ys(best[1]));
      hoverDot.style.opacity = 1;
      tip.style.opacity = 1;
      const pctX = (xs(best[0]) / W) * 100;
      const pctY = (ys(best[1]) / H) * 100;
      tip.style.left = pctX + '%';
      tip.style.top = pctY + '%';
      tip.textContent = (opts.tooltipFmt ? opts.tooltipFmt(best) : `t=${best[0].toFixed(2)}h, y=${best[1].toFixed(3)}`);
    });
    hitRect.addEventListener('mouseleave', ()=>{ hoverDot.style.opacity = 0; tip.style.opacity = 0; });
  }

  // ---------- header: raw stage log ----------
  const rawPts = DATA.raw.t.map((tv,i)=>[tv, DATA.raw.h[i]]);
  drawChart('chart-raw', {
    height: 260,
    series: [{points: rawPts, type:'line', color:'#1e4368', width:1.6}],
    yDecimals: 1,
    tooltipFmt: p => `t=${p[0].toFixed(2)}h  h=${p[1].toFixed(2)} m`
  });

  // ---------- derivatives ----------
  const dhPts = DATA.deriv.t.map((tv,i)=>[tv, DATA.deriv.dh[i]]);
  drawChart('chart-dh', {
    series: [{points: dhPts, type:'line', color:'#2f6fa8', width:1.8}],
    zeroLine: true,
    markerX: DATA.deriv.max_dhdt.t,
    markerLabel: `max dh/dt = ${DATA.deriv.max_dhdt.value.toFixed(3)} m/h`,
    tooltipFmt: p => `t=${p[0].toFixed(2)}h  dh/dt=${p[1].toFixed(3)} m/h`
  });
  document.getElementById('dhdt-stats').innerHTML = `
    <div class="stat-card"><div class="k">Time of max dh/dt</div><div class="v">${DATA.deriv.max_dhdt.t.toFixed(2)}<small> h</small></div></div>
    <div class="stat-card"><div class="k">Max dh/dt</div><div class="v">${DATA.deriv.max_dhdt.value.toFixed(3)}<small> m/h</small></div></div>`;

  const d2hPts = DATA.deriv.t.map((tv,i)=>[tv, DATA.deriv.d2h[i]]);
  drawChart('chart-d2h', {
    series: [{points: d2hPts, type:'line', color:'#64707c', width:1.6}],
    zeroLine: true,
    tooltipFmt: p => `t=${p[0].toFixed(2)}h  d\u00b2h/dt\u00b2=${p[1].toFixed(3)} m/h\u00b2`
  });
  document.getElementById('d2h-note').textContent = DATA.deriv.second_derivative_note;

  // ---------- model / fit ----------
  document.getElementById('model-assumptions').textContent = DATA.model.name + '. ' + DATA.model.assumptions;
  document.getElementById('model-equation').textContent = DATA.model.equation;
  document.getElementById('p0-note').textContent = 'Initial guess p0, read off the raw log: ' + DATA.model.p0_rationale;

  const fitPts = DATA.model.t_fit.map((tv,i)=>[tv, DATA.model.h_fit[i]]);
  drawChart('chart-fit', {
    height: 280,
    series: [
      {points: rawPts, type:'scatter', color:'#8598a8', r:1.6, opacity:0.6},
      {points: fitPts, type:'line', color:'#4a90c9', width:2.2}
    ],
    hoverSeriesIndex: 1,
    yDecimals: 1,
    tooltipFmt: p => `t=${p[0].toFixed(2)}h  \u0125=${p[1].toFixed(3)} m`
  });

  const paramsBody = document.querySelector('#params-tbl tbody');
  DATA.model.params.forEach(pr=>{
    const tr = document.createElement('tr');
    const sigClass = pr.p < 0.05 ? 'sig' : 'nosig';
    tr.innerHTML = `<td>${pr.label}</td><td>${pr.value.toFixed(4)}</td><td>${pr.se.toFixed(4)}</td>
                     <td>${pr.t.toFixed(2)}</td><td class="${sigClass}">${pr.p < 0.0001 ? '<0.0001' : pr.p.toFixed(4)}</td>`;
    paramsBody.appendChild(tr);
  });
  if (DATA.model.pcov_has_inf){
    const el = document.getElementById('pcov-note');
    el.style.display = 'block';
    el.textContent = 'Warning: the covariance matrix contains inf/nan \u2014 the model is carrying more parameters than the data supports, or the fit did not converge cleanly.';
  }

  document.getElementById('fit-stats').innerHTML = `
    <div class="stat-card"><div class="k">SSE</div><div class="v">${DATA.stats.sse.toFixed(3)}<small> m\u00b2</small></div></div>
    <div class="stat-card"><div class="k">SST</div><div class="v">${DATA.stats.sst.toFixed(1)}<small> m\u00b2</small></div></div>
    <div class="stat-card"><div class="k">R\u00b2</div><div class="v">${DATA.stats.r2.toFixed(4)}</div></div>
    <div class="stat-card"><div class="k">s (n\u2212p=${DATA.stats.dof})</div><div class="v">${DATA.stats.s.toFixed(4)}<small> m</small></div></div>
    <div class="stat-card"><div class="k">n / p</div><div class="v">${DATA.stats.n}<small> / ${DATA.stats.p}</small></div></div>`;

  // ---------- residuals ----------
  const residTPts = DATA.residuals.t.map((tv,i)=>[tv, DATA.residuals.e[i]]);
  drawChart('chart-resid-t', {
    series: [{points: residTPts, type:'scatter', color:'#16324f', r:1.8}],
    zeroLine: true,
    tooltipFmt: p => `t=${p[0].toFixed(2)}h  e=${p[1].toFixed(3)} m`
  });
  const residFPts = DATA.residuals.fitted.map((fv,i)=>[fv, DATA.residuals.e[i]]);
  drawChart('chart-resid-fit', {
    series: [{points: residFPts, type:'scatter', color:'#16324f', r:1.8}],
    zeroLine: true,
    yDecimals: 2,
    tooltipFmt: p => `\u0125=${p[0].toFixed(2)}m  e=${p[1].toFixed(3)} m`
  });
  document.getElementById('resid-drift').textContent = DATA.residuals.drift_note;
  document.getElementById('resid-runs').textContent = DATA.residuals.runs_note;
  document.getElementById('resid-spread').textContent = DATA.residuals.spread_note;
  document.getElementById('resid-resolution').textContent = DATA.residuals.resolution_note;
  document.getElementById('fit-verdict').textContent = DATA.residuals.verdict;

  // ---------- area ----------
  drawChart('chart-area', {
    height: 280,
    series: [{points: fitPts, type:'line', color:'#4a90c9', width:2}],
    fillSeriesIndex: 0,
    yDecimals: 1,
    tooltipFmt: p => `t=${p[0].toFixed(2)}h  \u0125=${p[1].toFixed(3)} m`
  });
  document.getElementById('area-stats').innerHTML = `
    <div class="stat-card"><div class="k">Area (quad on fit)</div><div class="v">${DATA.area.value.toFixed(2)}<small> m\u00b7h</small></div></div>
    <div class="stat-card"><div class="k">quad abs. error</div><div class="v">${DATA.area.abserr.toExponential(2)}</div></div>
    <div class="stat-card"><div class="k">Trapezoid (raw)</div><div class="v">${DATA.area.trapezoid.toFixed(2)}<small> m\u00b7h</small></div></div>
    <div class="stat-card"><div class="k">Gap</div><div class="v">${DATA.area.gap.toFixed(3)}<small> m\u00b7h</small></div></div>
    <div class="stat-card"><div class="k">Limits</div><div class="v" style="font-size:16px">${DATA.area.t0.toFixed(2)}\u2013${DATA.area.tn.toFixed(2)}<small> h</small></div></div>`;
  document.getElementById('area-gap-note').textContent = DATA.area.gap_note;
  document.getElementById('area-meaning-note').textContent = DATA.area.meaning_note;

})();
</script>
</body>
</html>
"""

html_out = HTML_TEMPLATE.replace("__DATA_JSON__", DATA_JSON)
with open(HTML_PATH, "w", encoding="utf-8") as f:
    f.write(html_out)

print("Wrote", HTML_PATH)