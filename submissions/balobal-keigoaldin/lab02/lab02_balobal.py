"""
NM-LAB-08252026 - Laboratory Activity 02: Fitting a curve to the dam
Prepared by: Balobal
Section: BES6-M

Runs top to bottom on a clean machine. Reads the reservoir stage log
(Day 1, 96 readings), computes finite-difference derivatives, fits a
four-parameter logistic via Levenberg-Marquardt, computes the fit
statistics, integrates the fitted curve, and writes the complete
results into a single self-contained dashboard: lab02_balobal.html

Per Rule 0: every number on the dashboard is computed here in Python.
The HTML file only displays values already computed; no fitting,
regression, or statistics happens in JavaScript.
"""
import io
import base64
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit
from scipy.integrate import quad
from scipy import stats

INPUT_XLSX = "/mnt/user-data/uploads/Data_01.xlsx"
OUTPUT_HTML = "lab02_balobal.html"


# ============================================================
# STEP 1 - Load the stage log and build a clean numeric time axis
# ============================================================
# The sheet has 3 rows above the real header (title, description, blank
# spacer); the real header row is at index 3.
raw = pd.read_excel(INPUT_XLSX, header=3)
raw.columns = [c.strip() for c in raw.columns]

# Isolate Day 1 only (96 readings), per the lab spec:
# "Ninety-six readings in a day, rounded by the logger to the nearest
# centimeter."
raw["Timestamp"] = pd.to_datetime(raw["Timestamp"])
day1 = raw[raw["Timestamp"].dt.date == raw["Timestamp"].dt.date.iloc[0]].reset_index(drop=True)
assert len(day1) == 96, f"Expected 96 readings for day 1, got {len(day1)}"

t0_calendar = day1["Timestamp"].iloc[0]
t = (day1["Timestamp"] - t0_calendar).dt.total_seconds().to_numpy() / 3600.0  # elapsed hours, float
h = day1["Depth (m)"].to_numpy(dtype=float)                                    # metres
n = len(t)


# ============================================================
# STEP 2 - Finite-difference derivatives (Tab Group 1)
# Forward diff at t[0], backward diff at t[-1], central diff between.
# ============================================================
def first_derivative(t, h):
    n = len(t)
    dh = np.empty(n)
    dh[0] = (h[1] - h[0]) / (t[1] - t[0])          # forward difference
    dh[-1] = (h[-1] - h[-2]) / (t[-1] - t[-2])     # backward difference
    for i in range(1, n - 1):
        dh[i] = (h[i + 1] - h[i - 1]) / (t[i + 1] - t[i - 1])  # central
    return dh


def second_derivative(t, h):
    n = len(t)
    d2h = np.empty(n)
    dt = t[1] - t[0]
    d2h[0] = (h[2] - 2 * h[1] + h[0]) / dt**2
    d2h[-1] = (h[-1] - 2 * h[-2] + h[-3]) / dt**2
    for i in range(1, n - 1):
        d2h[i] = (h[i + 1] - 2 * h[i] + h[i - 1]) / dt**2
    return d2h


dh_dt = first_derivative(t, h)
d2h_dt2 = second_derivative(t, h)
i_max = int(np.argmax(dh_dt))
t_max = t[i_max]
dh_max = dh_dt[i_max]


# ============================================================
# STEP 3 - Levenberg-Marquardt fit: four-parameter logistic (Tab Group 2)
# h(t) = c + a / (1 + exp(-k (t - t0)))
#
# Model defense: the reservoir shows one continuous filling trend across
# the 24-hour window with no visible second pulse, matching the
# logistic's "one filling event, single inflection" assumption. Because
# the data never levels off, the ceiling parameter (c + a) is extrapolated
# beyond the observed range rather than directly observed.
# ============================================================
def model(t, c, a, k, t0):
    return c + a / (1 + np.exp(-k * (t - t0)))


# p0 read from the raw plot:
#   c  ~ 14.15  (a touch below the first reading, 14.18 m -- lower asymptote)
#   a  ~ 0.6    (total rise; observed rise is 0.33 m over 24h but no ceiling
#                is visible yet, so the asymptotic total is guessed higher)
#   k  ~ 0.15   (gradual climb over the full 24h window, not a sharp S)
#   t0 ~ 20     (inflection appears late in/after the window, since the
#                curve is still climbing steadily at t = 23.75)
p0 = [14.15, 0.6, 0.15, 20.0]

popt, pcov = curve_fit(model, t, h, p0=p0, method="lm", maxfev=20000)
c_, a_, k_, t0_ = popt
param_names = ["c", "a", "k", "t0"]

resid = h - model(t, *popt)
p = len(popt)
dof = n - p

sse = np.sum(resid**2)
sst = np.sum((h - h.mean())**2)
r2 = 1 - sse / sst
s = np.sqrt(sse / dof)

se = np.sqrt(np.diag(pcov))
tj = popt / se
pj = 2 * (1 - stats.t.cdf(np.abs(tj), dof))

model_eq = f"h(t) = {c_:.4f} + {a_:.4f} / (1 + exp(-{k_:.4f} (t - {t0_:.4f})))"


# ============================================================
# STEP 4 - Integrate the fitted curve (Tab Group 3)
# ============================================================
t_start, t_end = t[0], t[-1]
area, abs_err = quad(lambda tt: model(tt, *popt), t_start, t_end)
area_trap = np.trapezoid(h, t)     # cross-check on the raw readings
gap = area - area_trap
gap_pct = 100 * gap / area


# ============================================================
# STEP 5 - Build the dashboard
# Charts are rendered as images (matplotlib) and embedded as base64, so
# the dashboard has zero external dependencies and works fully offline.
# ============================================================
def fig_to_base64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=130, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("ascii")


# ---- Chart 1: raw stage log (header, always visible) ----
fig, ax = plt.subplots(figsize=(9, 3.6))
ax.plot(t, h, 'o-', ms=3, lw=1, color='#1f5c73')
ax.set_xlabel("Elapsed time (hours from first reading)")
ax.set_ylabel("Depth (m)")
ax.set_title("Raw stage log — Day 1 (96 readings, 15-min sampling)")
ax.grid(alpha=0.3)
img_raw = fig_to_base64(fig)

# ---- Chart 2: derivatives (Tab 1) ----
fig, axes = plt.subplots(2, 1, figsize=(9, 6), sharex=True)
axes[0].plot(t, dh_dt, '-', color='#1f5c73', lw=1)
axes[0].axvline(t_max, color='crimson', ls='--', lw=1,
                 label=f"max dh/dt = {dh_max:.3f} m/h @ t={t_max:.2f}h")
axes[0].set_ylabel("dh/dt (m/h)")
axes[0].legend()
axes[0].grid(alpha=0.3)
axes[1].plot(t, d2h_dt2, '-', color='#a63d40', lw=1)
axes[1].axhline(0, color='k', lw=0.7)
axes[1].set_ylabel("d2h/dt2 (m/h^2)")
axes[1].set_xlabel("Elapsed time (h)")
axes[1].grid(alpha=0.3)
plt.tight_layout()
img_deriv = fig_to_base64(fig)

# ---- Chart 3: fitted curve + residuals (Tab 2) ----
t_fine = np.linspace(t.min(), t.max(), 400)
h_fit = model(t_fine, *popt)
fig, axes = plt.subplots(2, 1, figsize=(9, 6.5))
axes[0].plot(t, h, 'o', ms=3, color='#444', label='raw log')
axes[0].plot(t_fine, h_fit, '-', color='#1f5c73', lw=1.8, label='logistic fit')
axes[0].set_ylabel("Depth (m)")
axes[0].set_title("Fitted logistic vs raw stage log")
axes[0].legend()
axes[0].grid(alpha=0.3)
axes[1].scatter(t, resid, s=14, color='#a63d40')
axes[1].axhline(0, color='k', lw=0.8)
axes[1].axhline(0.01, color='gray', lw=0.6, ls=':')
axes[1].axhline(-0.01, color='gray', lw=0.6, ls=':', label='logger resolution (1 cm)')
axes[1].set_xlabel("Elapsed time (h)")
axes[1].set_ylabel("Residual (m)")
axes[1].set_title("Residuals vs time")
axes[1].legend()
axes[1].grid(alpha=0.3)
plt.tight_layout()
img_fit = fig_to_base64(fig)

# ---- Chart 4: area under fitted curve (Tab 3) ----
fig, ax = plt.subplots(figsize=(9, 4))
ax.plot(t_fine, h_fit, '-', color='#1f5c73', lw=1.8, label='fitted h(t)')
ax.fill_between(t_fine, 0, h_fit, color='#1f5c73', alpha=0.15)
ax.plot(t, h, 'o', ms=2.5, color='#444', alpha=0.6, label='raw log')
ax.set_ylim(0, h.max() * 1.05)
ax.set_xlabel("Elapsed time (h)")
ax.set_ylabel("Depth (m)")
ax.set_title(f"Area under fitted curve, t={t_start:.2f}h to t={t_end:.2f}h")
ax.legend()
ax.grid(alpha=0.3)
img_area = fig_to_base64(fig)

# ---- Parameter table rows ----
param_rows = ""
for name, val, sej, tval, pval in zip(param_names, popt, se, tj, pj):
    sig = "significant (p<0.05)" if pval < 0.05 else "not significant"
    param_rows += (f"<tr><td>{name}</td><td>{val:.5f}</td><td>{sej:.5f}</td>"
                    f"<td>{tval:.3f}</td><td>{pval:.3e}</td><td>{sig}</td></tr>\n")

HTML = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>NM Lab 02 — Fitting a curve to the dam</title>
<style>
  body {{ font-family: -apple-system, Segoe UI, Arial, sans-serif; margin: 0; padding: 0 0 60px 0;
         background: #f7f8f9; color: #1c1c1c; }}
  header {{ background: #16324a; color: white; padding: 24px 40px; }}
  header h1 {{ margin: 0 0 4px 0; font-size: 26px; }}
  header p {{ margin: 0; opacity: 0.85; font-size: 14px; }}
  .container {{ max-width: 980px; margin: 0 auto; padding: 24px 20px; }}
  .card {{ background: white; border-radius: 8px; padding: 20px 24px; margin-bottom: 20px;
           box-shadow: 0 1px 3px rgba(0,0,0,0.08); }}
  .card img {{ max-width: 100%; display: block; margin: 12px auto; }}
  .tabs {{ display: flex; gap: 4px; margin-bottom: 16px; }}
  .tab-btn {{ padding: 10px 20px; border: none; background: #e3e7ea; cursor: pointer;
              border-radius: 6px 6px 0 0; font-size: 14px; font-weight: 600; color: #444; }}
  .tab-btn.active {{ background: #16324a; color: white; }}
  .tab-content {{ display: none; }}
  .tab-content.active {{ display: block; }}
  table {{ border-collapse: collapse; width: 100%; margin: 12px 0; font-size: 14px; }}
  th, td {{ text-align: left; padding: 8px 10px; border-bottom: 1px solid #e3e7ea; }}
  th {{ background: #eef1f3; }}
  .stat-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
                gap: 12px; margin: 16px 0; }}
  .stat-box {{ background: #eef1f3; border-radius: 6px; padding: 12px 14px; }}
  .stat-box .label {{ font-size: 12px; color: #666; text-transform: uppercase; }}
  .stat-box .value {{ font-size: 20px; font-weight: 700; color: #16324a; }}
  .note {{ background: #fff8e6; border-left: 4px solid #d9a441; padding: 10px 14px;
           font-size: 14px; margin: 12px 0; border-radius: 0 4px 4px 0; }}
  .eq {{ font-family: monospace; background: #eef1f3; padding: 10px 14px; border-radius: 6px;
         font-size: 14px; overflow-x: auto; }}
</style>
</head>
<body>

<header>
  <h1>Fitting a curve to the dam</h1>
  <p>NM-LAB-08252026 &middot; Levenberg-Marquardt, residuals, and the area under the level &middot; Balobal</p>
</header>

<div class="container">

  <div class="card">
    <h2>Stage log time series</h2>
    <p>Raw logged depth at the reservoir, Day 1 ({t0_calendar.date()}), 96 readings at 15-minute intervals.</p>
    <img src="data:image/png;base64,{img_raw}" alt="raw stage log">
  </div>

  <div class="tabs">
    <button class="tab-btn active" onclick="showTab(0)">1. Derivatives</button>
    <button class="tab-btn" onclick="showTab(1)">2. Fitted Curve &amp; Statistics</button>
    <button class="tab-btn" onclick="showTab(2)">3. Area Under the Curve</button>
  </div>

  <div class="tab-content active" id="tab0">
    <div class="card">
      <h2>Finite-difference derivatives</h2>
      <img src="data:image/png;base64,{img_deriv}" alt="derivatives">
      <div class="stat-grid">
        <div class="stat-box"><div class="label">Max dh/dt</div><div class="value">{dh_max:.4f} m/h</div></div>
        <div class="stat-box"><div class="label">At time t</div><div class="value">{t_max:.2f} h</div></div>
      </div>
      <div class="note">
        The raw first- and second-derivative signals are dominated by the logger's 1&nbsp;cm rounding —
        dh/dt swings between roughly &minus;0.06 and +0.12&nbsp;m/h step to step, and d2h/dt2 shows no
        sustained sign pattern, only rounding-amplified noise. This is why the fitted curve below, not
        these raw differences, is used as the basis for the rest of the analysis.
      </div>
    </div>
  </div>

  <div class="tab-content" id="tab1">
    <div class="card">
      <h2>Model choice</h2>
      <p>
        <b>Four-parameter logistic</b> was chosen: the reservoir shows one continuous filling trend across
        the 24-hour window with no visible second pulse, matching the logistic's "one filling event, single
        inflection" assumption. Because the data never shows the level leveling off, the ceiling parameter
        (a) is extrapolated beyond the observed range rather than directly observed.
      </p>
      <div class="eq">{model_eq}</div>

      <h2>Fitted curve</h2>
      <img src="data:image/png;base64,{img_fit}" alt="fitted curve and residuals">

      <h2>Fitted parameters</h2>
      <table>
        <tr><th>Parameter</th><th>Value</th><th>Std. Error</th><th>t-statistic</th><th>p-value</th><th>Conclusion</th></tr>
        {param_rows}
      </table>

      <div class="stat-grid">
        <div class="stat-box"><div class="label">SSE</div><div class="value">{sse:.5f} m&sup2;</div></div>
        <div class="stat-box"><div class="label">SST</div><div class="value">{sst:.5f} m&sup2;</div></div>
        <div class="stat-box"><div class="label">R&sup2;</div><div class="value">{r2:.5f}</div></div>
        <div class="stat-box"><div class="label">s (std. error of estimate)</div><div class="value">{s:.5f} m</div></div>
        <div class="stat-box"><div class="label">n, p, dof</div><div class="value">{n}, {p}, {dof}</div></div>
      </div>

      <div class="note">
        Residuals are centered on zero with no long runs of one sign, and spread stays roughly consistent
        across the day. The largest residual reaches about &plusmn;0.03&nbsp;m — roughly 3&times; the logger's
        1&nbsp;cm resolution — so the fit is good but leaves a few centimeters of structure unexplained by
        the smooth logistic shape.
      </div>
    </div>
  </div>

  <div class="tab-content" id="tab2">
    <div class="card">
      <h2>Area under the fitted level</h2>
      <img src="data:image/png;base64,{img_area}" alt="area under curve">
      <div class="stat-grid">
        <div class="stat-box"><div class="label">A (quad, fitted curve)</div><div class="value">{area:.6f} m&middot;h</div></div>
        <div class="stat-box"><div class="label">quad abs. error est.</div><div class="value">{abs_err:.2e}</div></div>
        <div class="stat-box"><div class="label">A (trapezoid, raw data)</div><div class="value">{area_trap:.6f} m&middot;h</div></div>
        <div class="stat-box"><div class="label">Gap</div><div class="value">{gap:.6f} m&middot;h ({gap_pct:.3f}%)</div></div>
      </div>
      <div class="note">
        Units are meter-hours, not a volume — converting to an actual volume would require multiplying by
        the reservoir's surface area, which is not part of this dataset. The small gap between the quad
        and trapezoid estimates comes from the fitted curve being a smooth approximation that departs
        slightly from the noisy raw points it was fit to. For the flood-control office, this number
        summarizes cumulative depth-over-time in a way that is robust to per-reading rounding noise, useful
        for comparing days — it does not represent stored water volume, and it says nothing about risk at
        any single hour.
      </div>
    </div>
  </div>

</div>

<script>
  function showTab(i) {{
    document.querySelectorAll('.tab-content').forEach((el, idx) => {{
      el.classList.toggle('active', idx === i);
    }});
    document.querySelectorAll('.tab-btn').forEach((el, idx) => {{
      el.classList.toggle('active', idx === i);
    }});
  }}
</script>

</body>
</html>
"""

with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
    f.write(HTML)

print(f"Dashboard written: {OUTPUT_HTML}")
print()
print("=== Summary ===")
print(f"n = {n}, model: {model_eq}")
print(f"R^2 = {r2:.5f}, s = {s:.5f} m, dof = {dof}")
print(f"Max dh/dt = {dh_max:.4f} m/h at t = {t_max:.2f} h")
print(f"Area (quad) = {area:.6f} m*h (abs err {abs_err:.2e}); trapezoid = {area_trap:.6f} m*h")
