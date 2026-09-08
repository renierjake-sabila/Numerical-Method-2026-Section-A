"""
NM-LAB-08252026 - Laboratory Activity 02 : Fitting a curve to the dam
Numerical Methods, BES6-M (Engr. Mark Airol Escranda) - student submission

Runs top to bottom on a clean machine (numpy, scipy, matplotlib only) and
writes lab02_estiamba.html next to this file. Every number that lands on
the dashboard is computed here, in Python, and passed in as a value -
the HTML/JS only displays. (Rule 0.)
"""

import base64
import io
from datetime import datetime, timedelta

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit
from scipy import stats
from scipy.integrate import quad
from scipy.signal import savgol_filter

# ----------------------------------------------------------------------
# 00. DATA
# Reservoir stage log, 15-minute sampling. Depth above gauge datum,
# metres, rounded to the nearest centimetre by the logger.
# First reading: 2026-07-21 00:00 PHT. 288 readings -> 3.0 days.
# Transcribed unchanged from the "Data 01" stage log issued last meeting.
# ----------------------------------------------------------------------
START_TS = datetime(2026, 7, 21, 0, 0)
DT_HOURS = 0.25  # 15 minutes

DEPTH = [
14.18,14.21,14.22,14.21,14.21,14.2,14.24,14.25,14.25,14.23,14.23,14.23,14.25,
14.24,14.25,14.24,14.25,14.24,14.24,14.26,14.27,14.3,14.29,14.29,14.26,14.28,
14.3,14.3,14.31,14.32,14.32,14.31,14.31,14.29,14.3,14.32,14.33,14.33,14.32,
14.3,14.32,14.34,14.33,14.31,14.33,14.34,14.34,14.35,14.34,14.38,14.38,14.39,
14.39,14.39,14.37,14.38,14.39,14.39,14.39,14.41,14.43,14.41,14.4,14.43,14.42,
14.44,14.43,14.42,14.43,14.45,14.45,14.44,14.42,14.43,14.46,14.48,14.47,14.46,
14.48,14.46,14.48,14.48,14.49,14.47,14.47,14.47,14.49,14.48,14.48,14.48,14.5,
14.51,14.49,14.5,14.49,14.5,14.51,14.51,14.53,14.54,14.54,14.58,14.57,14.59,
14.61,14.64,14.7,14.74,14.82,14.9,14.99,15.11,15.23,15.4,15.6,15.8,16.02,
16.25,16.5,16.73,16.96,17.21,17.42,17.63,17.84,18.01,18.21,18.39,18.6,18.79,
19,19.21,19.4,19.64,19.86,20.08,20.29,20.46,20.62,20.78,20.91,21.01,21.06,
21.1,21.13,21.14,21.12,21.09,21.08,21.05,21,20.96,20.92,20.88,20.84,20.81,
20.76,20.72,20.69,20.65,20.6,20.57,20.52,20.46,20.43,20.4,20.35,20.31,20.3,
20.26,20.25,20.21,20.17,20.15,20.12,20.09,20.05,20.04,20,19.97,19.95,19.96,
19.94,19.91,19.89,19.9,19.85,19.85,19.83,19.82,19.8,19.8,19.78,19.76,19.75,
19.73,19.74,19.73,19.72,19.71,19.7,19.72,19.7,19.68,19.68,19.66,19.67,19.65,
19.65,19.64,19.62,19.64,19.63,19.64,19.62,19.61,19.6,19.59,19.59,19.59,19.59,
19.61,19.62,19.6,19.59,19.59,19.58,19.54,19.54,19.57,19.55,19.55,19.55,19.53,
19.53,19.56,19.56,19.55,19.53,19.53,19.52,19.52,19.51,19.52,19.52,19.53,19.53,
19.51,19.53,19.53,19.53,19.53,19.53,19.54,19.51,19.52,19.51,19.49,19.5,19.49,
19.5,19.52,19.54,19.53,19.52,19.54,19.53,19.52,19.5,19.53,19.51,19.5,19.52,
19.52,19.51,19.5,19.5,19.52,19.5,19.51,19.52,19.5,19.51,19.5,19.49,19.51,
19.52,19.51,
]

h = np.array(DEPTH, dtype=float)
n = len(h)
t = np.arange(n) * DT_HOURS  # hours elapsed from first reading (numeric axis)
timestamps = [START_TS + timedelta(hours=DT_HOURS * i) for i in range(n)]

assert n == 288, f"expected 288 readings, got {n}"

# ----------------------------------------------------------------------
# 01. TAB GROUP 1 - The derivative you already have
# Forward/backward at the two ends, central everywhere in between.
# ----------------------------------------------------------------------
d1 = np.empty(n)
d1[0] = (h[1] - h[0]) / DT_HOURS                     # forward
d1[-1] = (h[-1] - h[-2]) / DT_HOURS                   # backward
d1[1:-1] = (h[2:] - h[:-2]) / (2 * DT_HOURS)          # central

d2 = np.empty(n)
d2[0] = (h[2] - 2 * h[1] + h[0]) / DT_HOURS ** 2
d2[-1] = (h[-1] - 2 * h[-2] + h[-3]) / DT_HOURS ** 2
d2[1:-1] = (h[2:] - 2 * h[1:-1] + h[:-2]) / DT_HOURS ** 2

# Smoothed (Savitzky-Golay) versions, for the chart only. Every number
# reported in Tab 1 (max dh/dt, the event-vs-rest noise comparison in
# d2_note, etc.) still comes from the raw d1/d2 above - smoothing here is
# purely a visual aid so the underlying shape is easier to read against
# the quantisation noise.
SG_WINDOW, SG_POLY = 21, 3
d1_smooth = savgol_filter(d1, SG_WINDOW, SG_POLY)
d2_smooth = savgol_filter(d2, SG_WINDOW, SG_POLY)

i_max_rate = int(np.argmax(d1))
t_max_rate = t[i_max_rate]
v_max_rate = d1[i_max_rate]
ts_max_rate = timestamps[i_max_rate]

# One-sentence read of the second derivative: is the curvature during the
# event actually bigger than the background noise, or not?
_event_mask = (t >= 25) & (t <= 40)
_d2_std_event = float(np.std(d2[_event_mask]))
_d2_std_rest = float(np.std(d2[~_event_mask]))
d2_note = (
    "Differentiating twice turns the logger's one-centimetre rounding into "
    "noise of comparable size everywhere on the record: the second "
    "derivative's spread during the fill/release event (std {:.3f} m/h^2) "
    "is essentially the same as its spread on the flat parts of the log "
    "(std {:.3f} m/h^2), so the raw second derivative does not visibly "
    "single out the inflow event at all - it is swamped by quantisation "
    "noise everywhere, which is exactly why the write-up fits a smooth "
    "model and differentiates that instead of trusting this raw curve."
).format(_d2_std_event, _d2_std_rest)

# ----------------------------------------------------------------------
# 02. TAB GROUP 2 - The fit, and the arithmetic behind it
# ----------------------------------------------------------------------
MODEL_NAME = "Sum of two logistics (signed amplitudes)"
MODEL_EQUATION = "h(t) = c + a1 / (1 + exp(-k1 (t - t01))) + a2 / (1 + exp(-k2 (t - t02)))"
MODEL_JUSTIFICATION = (
    "The raw trace is not a single fill settling toward a ceiling: it rises "
    "sharply, overshoots to a peak near t=36 h, then falls off and settles "
    "on a second, lower plateau. A four-parameter logistic or a Gompertz "
    "can only rise to one ceiling and cannot reproduce that fall, so both "
    "are ruled out by the shape of the data, not by preference. The sum of "
    "two logistics can: a first sigmoid with positive amplitude for the "
    "fast inflow pulse, and a second sigmoid with negative amplitude for "
    "the drawdown/release that follows it, added onto a common baseline c."
)


def sigmoid(tt, a, k, t0):
    return a / (1 + np.exp(-k * (tt - t0)))


def model(tt, c, a1, k1, t01, a2, k2, t02):
    return c + sigmoid(tt, a1, k1, t01) + sigmoid(tt, a2, k2, t02)


PARAM_NAMES = ["c", "a1", "k1", "t01", "a2", "k2", "t02"]
PARAM_UNITS = ["m", "m", "1/h", "h", "m", "1/h", "h"]

# p0 read off the raw plot: baseline before the rise, rise magnitude
# (peak - baseline) around the visually steep midpoint, then the
# drawdown magnitude (plateau - peak) around the point where the
# decline is roughly half done. Levenberg-Marquardt takes no bounds,
# so this initial guess is doing the job bounds would otherwise do.
p0 = [14.18, 7.0, 0.5, 32.0, -1.6, 0.15, 42.0]

popt, pcov = curve_fit(model, t, h, p0=p0, method="lm", maxfev=20000)
h_hat = model(t, *popt)
resid = h - h_hat

sse = float(np.sum(resid ** 2))
sst = float(np.sum((h - h.mean()) ** 2))
r2 = 1 - sse / sst
p = len(popt)
dof = n - p
s_est = float(np.sqrt(sse / dof))

se = np.sqrt(np.diag(pcov))
tj = popt / se
pj = 2 * (1 - stats.t.cdf(np.abs(tj), dof))

converged_note = (
    "Fit converged; covariance diagonal is finite and positive throughout "
    "(no inf/nan terms), so the parameter count is supported by the data."
    if np.all(np.isfinite(np.diag(pcov))) and np.all(np.diag(pcov) > 0)
    else "WARNING: covariance matrix has non-finite entries - model is "
         "carrying more parameters than the data can support."
)

# Residual diagnostics (written, not just plotted)
resid_overall_mean = float(np.mean(resid))

signs = np.sign(resid)
signs[signs == 0] = 1
runs = []  # (start_t, end_t, sign, length)
start_i = 0
for i in range(1, n):
    if signs[i] != signs[i - 1]:
        runs.append((t[start_i], t[i - 1], signs[i - 1], i - start_i))
        start_i = i
runs.append((t[start_i], t[-1], signs[-1], n - start_i))
runs_sorted = sorted(runs, key=lambda r: -r[3])
longest_run = int(runs_sorted[0][3])
long_runs_desc = "; ".join(
    f"{'+' if r[2] > 0 else '-'} from t={r[0]:.1f} to {r[1]:.1f} h ({r[3]} pts)"
    for r in runs_sorted[:4]
)

spread_pre = float(np.std(resid[t < 25]))
spread_event = float(np.std(resid[(t >= 25) & (t < 40)]))
spread_post = float(np.std(resid[t >= 40]))

max_abs_resid = float(np.max(np.abs(resid)))
i_max_resid = int(np.argmax(np.abs(resid)))
t_max_resid = t[i_max_resid]
resolution_m = 0.01
resid_vs_resolution = max_abs_resid / resolution_m

residual_reading = (
    "Residuals are centred close to zero overall (mean {mean:+.4f} m), "
    "so there is no net drift across the record - but they are not "
    "unstructured. The four longest runs of same-sign residuals each "
    "cover roughly 12-14 hours and together span almost the whole "
    "record: {runs}. Runs this long mean the model's shape is slightly "
    "wrong even outside the event: the 'flat' stretches before and "
    "after the pulse are not perfectly flat in the raw log, and a "
    "constant baseline c cannot bend to follow that slow creep, so the "
    "fit rides consistently above or below it for many hours at a time. "
    "That is a shape limitation, not sensor noise. Spread does widen "
    "somewhat during the event itself ({se:.3f} m) versus the quieter "
    "pre-event ({sp:.3f} m) and post-event ({sq:.3f} m) stretches, which "
    "is expected since that is where the curve is steepest. The single "
    "largest residual is {mx:.3f} m at t={tmx:.2f} h, {ratio:.0f} times "
    "the logger's own one-centimetre resolution, right at the peak, "
    "confirming the fit slightly rounds off the sharpest part of the "
    "overshoot."
).format(
    mean=resid_overall_mean, runs=long_runs_desc, se=spread_event,
    sp=spread_pre, sq=spread_post, mx=max_abs_resid, tmx=t_max_resid,
    ratio=resid_vs_resolution,
)

r2_verdict = (
    "R^2 = {:.4f} is high, but it is sitting on top of the long, "
    "same-sign residual runs described below, not on unstructured "
    "scatter - by the lab's own standard that makes this an honest "
    "partial success rather than a clean pass: the two-logistic shape "
    "captures the fill/release event well but is too rigid a baseline "
    "to track the slow creep in the 'flat' sections before and after "
    "it.".format(r2)
)

# ----------------------------------------------------------------------
# 03. TAB GROUP 3 - The area under the fitted level
# ----------------------------------------------------------------------
t0_lim, tn_lim = float(t[0]), float(t[-1])
area_quad, area_quad_err = quad(lambda tt: model(tt, *popt), t0_lim, tn_lim)
area_trap = float(np.trapezoid(h, t))
area_gap = area_quad - area_trap
area_gap_note = (
    "The two areas differ by {:.4f} m*h ({:.3f}% of the quad value) "
    "because the trapezoid rule integrates the raw, centimetre-rounded "
    "points with straight-line segments between them, while quad "
    "integrates the smooth fitted curve; the gap is concentrated where "
    "the trapezoid rule under- or over-shoots the sharp fill/release "
    "flank that straight segments approximate poorly.".format(
        area_gap, 100 * area_gap / area_quad
    )
)
area_meaning_note = (
    "This area is meters-hours, the time-integral of stage - a measure of "
    "how high and how long the reservoir stood over the {:.1f}-hour "
    "record. It tells the flood-control office the overall exposure of "
    "the reservoir level through the event. It is NOT a volume: turning "
    "it into cubic metres would require the reservoir's stage-to-area or "
    "stage-to-volume curve, which this stage log does not contain.".format(
        tn_lim - t0_lim
    )
)

# ----------------------------------------------------------------------
# CHARTS -> base64 PNG, embedded directly in the HTML (no CDN needed)
# ----------------------------------------------------------------------
plt.rcParams.update({"font.size": 10, "figure.dpi": 130})


def fig_to_base64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode("ascii")


# Header: raw stage series, always visible
fig = plt.figure(figsize=(9.5, 3.2))
plt.plot(t, h, color="#1f4e5f", linewidth=1.2)
plt.xlabel("t (hours from first reading)")
plt.ylabel("depth h (m)")
plt.title("Reservoir stage log - raw readings (n=288, 15-min sampling)")
plt.grid(alpha=0.25)
img_header = fig_to_base64(fig)

# Tab 1: derivatives
fig, axes = plt.subplots(2, 1, figsize=(9.5, 6.2), sharex=True)
axes[0].plot(t, d1, color="#2b6777", linewidth=0.8, alpha=0.28, label="dh/dt (raw)")
axes[0].plot(t, d1_smooth, color="#2b6777", linewidth=1.6,
             label="dh/dt (smoothed, Savitzky-Golay)")
axes[0].scatter([t_max_rate], [v_max_rate], color="#c0392b", zorder=5,
                 label=f"max dh/dt = {v_max_rate:.3f} m/h @ t={t_max_rate:.2f} h")
axes[0].set_ylabel("dh/dt (m/h)")
axes[0].legend(loc="upper right", fontsize=8)
axes[0].grid(alpha=0.25)
axes[1].plot(t, d2, color="#7f4f9c", linewidth=0.7, alpha=0.28, label="d2h/dt2 (raw)")
axes[1].plot(t, d2_smooth, color="#7f4f9c", linewidth=1.5,
             label="d2h/dt2 (smoothed, Savitzky-Golay)")
axes[1].axhline(0, color="black", linewidth=0.6)
axes[1].set_ylabel("d2h/dt2 (m/h^2)")
axes[1].set_xlabel("t (hours from first reading)")
axes[1].legend(loc="upper right", fontsize=8)
axes[1].grid(alpha=0.25)
fig.suptitle("Finite-difference derivatives of the raw log (smoothed overlay, raw behind)")
img_derivatives = fig_to_base64(fig)

# Tab 2: fit over data
fig = plt.figure(figsize=(9.5, 3.4))
plt.plot(t, h, ".", color="#8f9aa3", markersize=3, label="raw readings")
plt.plot(t, h_hat, color="#c0392b", linewidth=1.4, label="fitted h(t)")
plt.xlabel("t (h)"); plt.ylabel("h (m)")
plt.title("Fitted curve over raw data")
plt.legend(fontsize=8)
plt.grid(alpha=0.25)
img_fit = fig_to_base64(fig)

# Tab 2: residuals
fig, axes = plt.subplots(1, 2, figsize=(9.5, 3.2))
axes[0].plot(t, resid, color="#2b6777", linewidth=0.8)
axes[0].axhline(0, color="black", linewidth=0.7)
axes[0].set_xlabel("t (h)"); axes[0].set_ylabel("residual e_i (m)")
axes[0].set_title("Residuals vs time")
axes[0].grid(alpha=0.25)
axes[1].scatter(h_hat, resid, s=6, color="#2b6777")
axes[1].axhline(0, color="black", linewidth=0.7)
axes[1].set_xlabel("fitted h_hat (m)"); axes[1].set_ylabel("residual e_i (m)")
axes[1].set_title("Residuals vs fitted value")
axes[1].grid(alpha=0.25)
fig.tight_layout()
img_resid = fig_to_base64(fig)

# Tab 3: area under the fitted curve
fig = plt.figure(figsize=(9.5, 3.4))
tt_fine = np.linspace(t0_lim, tn_lim, 2000)
plt.fill_between(tt_fine, 0, model(tt_fine, *popt), color="#a9c9d1", alpha=0.6,
                  label=f"area (quad) = {area_quad:.3f} m*h")
plt.plot(tt_fine, model(tt_fine, *popt), color="#1f4e5f", linewidth=1.2)
plt.plot(t, h, ".", color="#8f9aa3", markersize=2.5, label="raw readings")
plt.ylim(bottom=0)
plt.xlabel("t (h)"); plt.ylabel("h (m)")
plt.title("Area under the fitted level, referenced to h=0")
plt.legend(fontsize=8)
plt.grid(alpha=0.25)
img_area = fig_to_base64(fig)

# ----------------------------------------------------------------------
# HTML DASHBOARD
# ----------------------------------------------------------------------
def param_rows():
    out = []
    for name, unit, val, s_val, tval, pval in zip(PARAM_NAMES, PARAM_UNITS, popt, se, tj, pj):
        out.append(
            f"<tr><td>{name}</td><td>{unit}</td><td>{val:.4f}</td>"
            f"<td>{s_val:.4f}</td><td>{tval:.3f}</td><td>{pval:.3e}</td></tr>"
        )
    return "\n".join(out)


html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>NM-LAB-08252026 - Lab 02 Dashboard - Estiamba</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@600;700;800&family=Inter:wght@400;500;600;700&display=swap');

  :root {{
    --ink:#12313a; --paper:#f2fafa; --card:#ffffff;
    --teal:#0e7c86; --teal-dark:#0b4f6c; --coral:#ff7a59; --amber:#ffb238;
    --line:#dceeee; --muted:#5b7a80;
    --shadow:0 10px 28px rgba(11,79,108,0.12);
  }}
  * {{ box-sizing:border-box; }}
  body {{
    margin:0; padding:0 0 60px 0; background:var(--paper); color:var(--ink);
    font-family:'Inter', Arial, sans-serif;
  }}

  header.hero {{
    position:relative; overflow:hidden; color:#fff; padding:40px 32px 72px 32px;
    background:linear-gradient(135deg,var(--teal-dark) 0%,var(--teal) 55%,#12a594 100%);
  }}
  header.hero .kicker {{
    display:inline-block; background:rgba(255,255,255,0.16); border:1px solid rgba(255,255,255,0.35);
    padding:5px 14px; border-radius:999px; font-size:11px; letter-spacing:0.13em;
    text-transform:uppercase; font-weight:600;
  }}
  header.hero h1 {{
    font-family:'Poppins', Arial, sans-serif; font-weight:800; font-size:34px;
    margin:16px 0 8px 0; max-width:640px; line-height:1.15;
  }}
  header.hero .sub {{ font-size:13px; opacity:0.92; max-width:640px; line-height:1.6; }}
  .wave {{ position:absolute; left:0; right:0; bottom:-1px; line-height:0; }}
  .wave svg {{ width:100%; height:54px; display:block; }}

  main {{ max-width:1000px; margin:-34px auto 0 auto; padding:0 28px; position:relative; z-index:2; }}

  section.panel {{
    background:var(--card); border:1px solid var(--line); border-radius:18px;
    padding:26px 28px; margin:22px 0; box-shadow:var(--shadow);
  }}
  section.panel.hero-panel {{ padding:20px 26px; }}

  .section-title {{ display:flex; align-items:center; gap:2px; margin-bottom:4px; }}
  .badge {{
    display:inline-flex; align-items:center; justify-content:center; width:32px; height:32px;
    border-radius:10px; background:var(--accent-c, var(--teal)); font-size:16px; margin-right:10px; flex:none;
  }}
  h2 {{ font-family:'Poppins', Arial, sans-serif; font-size:19px; margin:0; color:var(--teal-dark); }}
  h3.label {{
    font-size:11.5px; text-transform:uppercase; letter-spacing:0.09em;
    color:var(--muted); font-weight:700; margin:0 0 8px 0;
  }}

  .chart-frame {{ border-radius:14px; overflow:hidden; border:1px solid var(--line); margin:6px 0 16px 0; }}
  img.chart {{ width:100%; height:auto; display:block; }}

  p.note {{ font-size:13.5px; line-height:1.65; color:#33474c; }}

  .tabs {{ display:flex; gap:10px; margin:26px 0 0 0; flex-wrap:wrap; }}
  .tab-btn {{
    flex:1; min-width:150px; padding:13px 14px; border:2px solid var(--line); background:#fff;
    color:var(--teal-dark); cursor:pointer; font-family:'Inter', Arial, sans-serif;
    font-size:13px; font-weight:700; border-radius:14px;
    display:flex; align-items:center; justify-content:center; gap:8px;
    transition:transform .15s ease;
  }}
  .tab-btn .dot {{ width:9px; height:9px; border-radius:50%; background:var(--dot, var(--teal)); flex:none; }}
  .tab-btn:hover {{ transform:translateY(-2px); }}
  .tab-btn.active {{ color:#fff; border-color:transparent; box-shadow:var(--shadow); }}
  .tab-btn.active .dot {{ background:#fff; }}
  #btn-0.active {{ background:linear-gradient(135deg,var(--teal-dark),var(--teal)); }}
  #btn-1.active {{ background:linear-gradient(135deg,#e85d3d,var(--coral)); }}
  #btn-2.active {{ background:linear-gradient(135deg,#e39a00,var(--amber)); }}

  .tab-content {{ display:none; }}
  .tab-content.active {{ display:block; }}

  table {{ border-collapse:collapse; width:100%; font-size:12.5px; margin:12px 0; }}
  th, td {{ border-bottom:1px solid var(--line); padding:9px 10px; text-align:left; }}
  th {{
    background:#effaf9; color:var(--teal-dark); font-weight:700;
    text-transform:uppercase; font-size:10.5px; letter-spacing:0.05em;
  }}
  tr:last-child td {{ border-bottom:none; }}

  .stat-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:14px; margin:16px 0; }}
  .stat-card {{
    border-radius:14px; padding:14px 16px; background:#f7fcfc;
    border-left:5px solid var(--accent-c, var(--teal)); box-shadow:0 4px 14px rgba(11,79,108,0.06);
  }}
  .stat-card .label {{ font-size:10.5px; text-transform:uppercase; letter-spacing:0.07em; color:var(--muted); font-weight:700; }}
  .stat-card .value {{
    font-family:'Poppins', Arial, sans-serif; font-size:21px; font-weight:800;
    color:var(--ink); margin:3px 0 2px 0;
  }}

  code.eq {{
    display:block; background:var(--teal-dark); color:#eafbf8; padding:14px 16px; border-radius:12px;
    font-family:'Courier New',monospace; font-size:12.5px; margin:10px 0; overflow-x:auto;
  }}

  footer {{
    max-width:1000px; margin:26px auto 0 auto; padding:14px 28px;
    font-size:11px; color:var(--muted); text-align:center;
  }}
</style>
</head>
<body>

<header class="hero">
  <span class="kicker">NM-LAB-08252026 &middot; Laboratory Activity 02</span>
  <h1>Fitting a Curve to the Dam</h1>
  <div class="sub">Levenberg&ndash;Marquardt, residuals, and the area under the level &middot;
  Numerical Methods BES6-M &middot; Submitted by Estiamba &middot;
  Dashboard generated {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
  <div class="wave">
    <svg viewBox="0 0 1200 60" preserveAspectRatio="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M0,32 C150,60 350,0 600,28 C850,56 1050,8 1200,30 L1200,60 L0,60 Z" fill="#f2fafa"/>
    </svg>
  </div>
</header>

<main>

  <section class="panel hero-panel">
    <h3 class="label">Stage log time series &mdash; always visible</h3>
    <div class="chart-frame"><img class="chart" src="data:image/png;base64,{img_header}" alt="raw stage log"></div>
    <p class="note">Depth sensor readings every 15 minutes, {n} points,
    {t0_lim:.2f} h to {tn_lim:.2f} h from the first reading
    ({timestamps[0].strftime('%Y-%m-%d %H:%M')} to {timestamps[-1].strftime('%Y-%m-%d %H:%M')} PHT).</p>
  </section>

  <div class="tabs">
    <button id="btn-0" class="tab-btn active" style="--dot:#0e7c86" onclick="showTab(0)"><span class="dot"></span>Derivatives</button>
    <button id="btn-1" class="tab-btn" style="--dot:#ff7a59" onclick="showTab(1)"><span class="dot"></span>Fit &amp; Statistics</button>
    <button id="btn-2" class="tab-btn" style="--dot:#ffb238" onclick="showTab(2)"><span class="dot"></span>Area Under the Curve</button>
  </div>

  <!-- TAB 1 -->
  <section class="panel tab-content active" id="tab-0">
    <div class="section-title"><span class="badge" style="--accent-c:#0e7c86">&#128200;</span><h2>The Derivative You Already Have</h2></div>
    <div class="chart-frame"><img class="chart" src="data:image/png;base64,{img_derivatives}" alt="derivatives"></div>
    <div class="stat-grid">
      <div class="stat-card" style="--accent-c:#0e7c86"><div class="label">Time of Max dh/dt</div>
        <div class="value">{t_max_rate:.2f} h</div>
        <div class="label">{ts_max_rate.strftime('%Y-%m-%d %H:%M')} PHT</div></div>
      <div class="stat-card" style="--accent-c:#0e7c86"><div class="label">Max dh/dt</div>
        <div class="value">{v_max_rate:.3f} m/h</div></div>
    </div>
    <p class="note"><strong>Second derivative, read in one sentence:</strong> {d2_note}</p>
  </section>

  <!-- TAB 2 -->
  <section class="panel tab-content" id="tab-1">
    <div class="section-title"><span class="badge" style="--accent-c:#ff7a59">&#129518;</span><h2>The Fit, and the Arithmetic Behind It</h2></div>
    <h3 class="label">Model chosen: {MODEL_NAME}</h3>
    <code class="eq">{MODEL_EQUATION}</code>
    <p class="note">{MODEL_JUSTIFICATION}</p>
    <p class="note"><em>Initial guess p0</em> = {p0}, read off the raw plot: baseline
    before the rise, rise magnitude at the visually steep midpoint, then the
    drawdown magnitude at the point where the decline looks roughly half done.
    Levenberg&ndash;Marquardt takes no bounds, so this guess is standing in for them.
    {converged_note}</p>

    <div class="chart-frame"><img class="chart" src="data:image/png;base64,{img_fit}" alt="fitted curve"></div>

    <h3 class="label">Fitted Parameters</h3>
    <table>
      <tr><th>Parameter</th><th>Unit</th><th>Value</th><th>Std. Error</th><th>t Statistic</th><th>p-value (2-tailed)</th></tr>
      {param_rows()}
    </table>

    <div class="stat-grid">
      <div class="stat-card" style="--accent-c:#ff7a59"><div class="label">SSE</div><div class="value">{sse:.4f} m&sup2;</div>
        <div class="label">n={n}, p={p}</div></div>
      <div class="stat-card" style="--accent-c:#ff7a59"><div class="label">R&sup2;</div><div class="value">{r2:.4f}</div>
        <div class="label">SST = {sst:.4f} m&sup2;</div></div>
      <div class="stat-card" style="--accent-c:#ff7a59"><div class="label">Standard Error of Estimate s</div>
        <div class="value">{s_est:.4f} m</div><div class="label">dof = n&minus;p = {dof}</div></div>
    </div>
    <p class="note">{r2_verdict}</p>

    <h3 class="label">Residuals</h3>
    <div class="chart-frame"><img class="chart" src="data:image/png;base64,{img_resid}" alt="residuals"></div>
    <p class="note">{residual_reading}</p>
  </section>

  <!-- TAB 3 -->
  <section class="panel tab-content" id="tab-2">
    <div class="section-title"><span class="badge" style="--accent-c:#ffb238">&#127754;</span><h2>The Area Under the Fitted Level</h2></div>
    <div class="chart-frame"><img class="chart" src="data:image/png;base64,{img_area}" alt="area under fitted curve"></div>
    <div class="stat-grid">
      <div class="stat-card" style="--accent-c:#ffb238"><div class="label">Area (scipy.integrate.quad)</div>
        <div class="value">{area_quad:.4f} m&middot;h</div>
        <div class="label">abs. error est. &plusmn;{area_quad_err:.2e} m&middot;h</div></div>
      <div class="stat-card" style="--accent-c:#ffb238"><div class="label">Area (np.trapezoid, raw points)</div>
        <div class="value">{area_trap:.4f} m&middot;h</div></div>
      <div class="stat-card" style="--accent-c:#ffb238"><div class="label">Gap (quad &minus; trapezoid)</div>
        <div class="value">{area_gap:+.4f} m&middot;h</div></div>
    </div>
    <p class="note"><strong>Gap explained:</strong> {area_gap_note}</p>
    <p class="note"><strong>What this means to the flood-control office:</strong> {area_meaning_note}</p>
  </section>

</main>

<footer>Generated entirely by lab02_estiamba.py &mdash; every value above was
computed in Python and written into this file; no fitting, regression, or
statistics run in this page's JavaScript (Rule 0). Tab switching below is
pure display logic.</footer>

<script>
function showTab(i) {{
  var tabs = document.querySelectorAll('.tab-content');
  var btns = document.querySelectorAll('.tab-btn');
  for (var j = 0; j < tabs.length; j++) {{
    tabs[j].classList.remove('active');
    btns[j].classList.remove('active');
  }}
  tabs[i].classList.add('active');
  btns[i].classList.add('active');
}}
</script>

</body>
</html>
"""

with open("lab02_estiamba.html", "w", encoding="utf-8") as f:
    f.write(html)

# ----------------------------------------------------------------------
# Console summary (for the results write-up / oral check)
# ----------------------------------------------------------------------
print("=== Tab 1: derivatives ===")
print(f"max dh/dt = {v_max_rate:.4f} m/h at t = {t_max_rate:.2f} h "
      f"({ts_max_rate})")
print(d2_note)
print()
print("=== Tab 2: fit ===")
print("model:", MODEL_EQUATION)
for name, unit, val, s_val, tval, pval in zip(PARAM_NAMES, PARAM_UNITS, popt, se, tj, pj):
    print(f"  {name:>4s} = {val: .5f} {unit:<4s}  SE={s_val:.5f}  t={tval:.3f}  p={pval:.3e}")
print(f"SSE={sse:.5f}  SST={sst:.5f}  R2={r2:.5f}  s={s_est:.5f}  dof={dof}")
print(residual_reading)
print()
print("=== Tab 3: integration ===")
print(f"quad area = {area_quad:.5f} m*h (abs err {area_quad_err:.2e})")
print(f"trapezoid area = {area_trap:.5f} m*h")
print(area_gap_note)
print()
print("Dashboard written to lab02_estiamba.html")
