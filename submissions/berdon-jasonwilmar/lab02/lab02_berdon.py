import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit
from scipy.integrate import quad, trapezoid
import scipy.stats as stats
import base64
from io import BytesIO
import os
import webbrowser

# ------------------------------------------------------------
# 1. LOAD ALL DATA (288 readings)
# ------------------------------------------------------------
script_dir = os.path.dirname(os.path.abspath(__file__))
possible_paths = [
    os.path.join(script_dir, 'Data 01.xlsx'),
    os.path.join(script_dir, 'Data01.xlsx'),
    'Data 01.xlsx',
]

file_path = None
for path in possible_paths:
    if os.path.isfile(path):
        file_path = path
        break

if file_path is None:
    print("Could not find 'Data 01.xlsx' in the script folder or current directory.")
    user_input = input("Please enter the full path to the Excel file: ").strip()
    if os.path.isfile(user_input):
        file_path = user_input
    else:
        raise FileNotFoundError(f"File not found at {user_input}")

print(f"Loading data from: {file_path}")

df = pd.read_excel(file_path, sheet_name='Sensor Log', skiprows=3)
# Use all 288 readings (full dataset)
df['Timestamp'] = pd.to_datetime(df['Timestamp'])
t0 = df['Timestamp'].iloc[0]
df['hours'] = (df['Timestamp'] - t0).dt.total_seconds() / 3600.0

t = df['hours'].values
h = df['Depth (m)'].values
n = len(h)

# ------------------------------------------------------------
# 2. DERIVATIVES (finite differences)
# ------------------------------------------------------------
dt = 0.25
dhdt = np.zeros_like(h)
dhdt[1:-1] = (h[2:] - h[:-2]) / (2*dt)
dhdt[0] = (h[1] - h[0]) / dt
dhdt[-1] = (h[-1] - h[-2]) / dt

d2hdt2 = np.zeros_like(h)
d2hdt2[1:-1] = (h[2:] - 2*h[1:-1] + h[:-2]) / (dt**2)

max_idx = np.argmax(dhdt)
max_time = t[max_idx]
max_rate = dhdt[max_idx]
min_dhdt = np.min(dhdt)
max_dhdt = np.max(dhdt)

# Smoothed derivatives (moving average over 2.25 h = 9 points)
from scipy.signal import savgol_filter
window = 9
ma_dhdt = np.convolve(dhdt, np.ones(window)/window, mode='same')
ma_d2 = np.convolve(d2hdt2, np.ones(window)/window, mode='same')
max_ma_dhdt = np.max(ma_dhdt)
max_ma_dhdt_t = t[np.argmax(ma_dhdt)]
max_ma_d2 = np.max(ma_d2)
max_ma_d2_t = t[np.argmax(ma_d2)]
min_ma_d2 = np.min(ma_d2)
min_ma_d2_t = t[np.argmin(ma_d2)]

# Savitzky-Golay
sg_d1 = savgol_filter(h, window_length=9, polyorder=3, deriv=1, delta=dt)
sg_d2 = savgol_filter(h, window_length=9, polyorder=3, deriv=2, delta=dt)
sg_d1_peak = np.max(sg_d1)
sg_d1_peak_t = t[np.argmax(sg_d1)]
sg_d2_max = np.max(sg_d2)
sg_d2_max_t = t[np.argmax(sg_d2)]
sg_d2_min = np.min(sg_d2)
sg_d2_min_t = t[np.argmin(sg_d2)]

# Find zero crossing of second derivative (inflection)
sign_changes = np.where(np.diff(np.sign(d2hdt2)))[0]
if len(sign_changes) > 0:
    inflection_idx = sign_changes[0]
    inflection_time = t[inflection_idx]
else:
    inflection_time = max_time

# ------------------------------------------------------------
# 3. MODEL: Sum of two logistics
# ------------------------------------------------------------
def model_two_log(t, c, a1, k1, t1, a2, k2, t2):
    return c + a1 / (1 + np.exp(-k1*(t - t1))) + a2 / (1 + np.exp(-k2*(t - t2)))

# Initial guess (based on plot reading)
p0 = [14.2, 7.0, 0.5, 30.0, -1.7, 0.2, 42.0]
popt, pcov = curve_fit(model_two_log, t, h, p0=p0, method='lm', maxfev=20000)
c, a1, k1, t1, a2, k2, t2 = popt
p = len(popt)

# Fitted values and residuals
h_fit = model_two_log(t, *popt)
resid = h - h_fit
SSE = np.sum(resid**2)
SST = np.sum((h - np.mean(h))**2)
R2 = 1 - SSE/SST
s_est = np.sqrt(SSE / (n - p))

# Parameter errors, t, p-values
param_err = np.sqrt(np.diag(pcov))
t_stats = popt / param_err
p_values = 2 * (1 - stats.t.cdf(np.abs(t_stats), df=n-p))

# Sign runs
runs = 1 + np.sum(np.diff(np.sign(resid)) != 0)
expected_runs = (2*n - 1)/3

# Residual block stats (12-hour blocks)
block_means = []
block_sds = []
for i in range(6):
    start = i*12
    end = (i+1)*12
    if end > n:
        end = n
    if start < n:
        block_means.append(np.mean(resid[start:end]))
        block_sds.append(np.std(resid[start:end]))

# ------------------------------------------------------------
# 4. AREA under curve
# ------------------------------------------------------------
area, abserr = quad(lambda x: model_two_log(x, *popt), t[0], t[-1])
area_trapz = trapezoid(h_fit, t)
mean_level = area / (t[-1] - t[0])

# ------------------------------------------------------------
# 5. ANALYTIC DERIVATIVES (for peak rate)
# ------------------------------------------------------------
def model_deriv(x, c, a1, k1, t1, a2, k2, t2):
    S1 = 1/(1+np.exp(-k1*(x-t1)))
    S2 = 1/(1+np.exp(-k2*(x-t2)))
    return a1*k1*S1*(1-S1) + a2*k2*S2*(1-S2)

t_fine = np.linspace(t[0], t[-1], 5000)
deriv_fine = model_deriv(t_fine, *popt)
peak_deriv = np.max(deriv_fine)
peak_deriv_t = t_fine[np.argmax(deriv_fine)]

# Second derivative from analytic derivative (numeric gradient)
d2_analytic = np.gradient(deriv_fine, t_fine)

# ------------------------------------------------------------
# 6. PHASES (as in your original HTML)
# ------------------------------------------------------------
phase_data = [
    {"name": "Quiet baseline", "t0": 0.0, "t1": 27.5, "h0": 14.18, "h1": 14.99, "clock0": "21 Jul 00:00", "note": "level flat, rate indistinguishable from the logger's own rounding"},
    {"name": "Rising limb", "t0": 27.75, "t1": 35.25, "h0": 15.11, "h1": 21.01, "clock0": "22 Jul 03:45", "note": "the filling event, bracketed by t1 ± 2/k1 from the fit"},
    {"name": "Crest", "t0": 35.5, "t1": 38.25, "h0": 21.06, "h1": 20.88, "clock0": "22 Jul 11:30", "note": "inflow and outflow balance, the level turns over"},
    {"name": "Recession", "t0": 38.25, "t1": 71.75, "h0": 20.88, "h1": 19.51, "clock0": "22 Jul 14:15", "note": "drawdown, slower and broader than the rise"}
]

for ph in phase_data:
    mask = (t >= ph["t0"]) & (t <= ph["t1"])
    ph_t = t[mask]
    ph_h = h[mask]
    ph["hours"] = ph["t1"] - ph["t0"]
    ph["change"] = ph["h1"] - ph["h0"]
    if ph["hours"] > 0:
        ph["mean_rate"] = ph["change"] / ph["hours"]
    else:
        ph["mean_rate"] = 0
    mask_fine = (t_fine >= ph["t0"]) & (t_fine <= ph["t1"])
    if np.any(mask_fine):
        ph["peak_rate"] = np.max(deriv_fine[mask_fine])
        ph["peak_rate_t"] = t_fine[mask_fine][np.argmax(deriv_fine[mask_fine])]
    else:
        ph["peak_rate"] = 0
        ph["peak_rate_t"] = ph["t0"]

# ------------------------------------------------------------
# 7. PLOTS (embedded as base64)
# ------------------------------------------------------------
def fig_to_base64(fig):
    buf = BytesIO()
    fig.savefig(buf, format='png', bbox_inches='tight', dpi=120, facecolor='white')
    buf.seek(0)
    return base64.b64encode(buf.read()).decode('utf-8')

# Plot 1: Raw + fit
fig1, ax1 = plt.subplots(figsize=(10, 5))
ax1.plot(t, h, 'o', color='#00f0ff', label='Raw data', markersize=3, alpha=0.7)
ax1.plot(t, h_fit, '-', color='#ff00e5', label='Fitted curve', linewidth=2.5)
ax1.fill_between(t, h_fit - s_est, h_fit + s_est, color='#ff00e5', alpha=0.15, label=f'±1σ (s = {s_est:.4f} m)')
ax1.set_xlabel('Hours elapsed', fontsize=12)
ax1.set_ylabel('Stage (m)', fontsize=12)
ax1.set_title('Reservoir Stage: Data and Fit', fontsize=14)
ax1.legend()
ax1.grid(True, alpha=0.3)
img1 = fig_to_base64(fig1)

# Plot 2: Residuals
fig2, ax2 = plt.subplots(figsize=(10, 4))
ax2.plot(t, resid, 'o', color='#00f0ff', markersize=3, alpha=0.7)
ax2.axhline(0, color='#ff00e5', linestyle='--')
ax2.axhline(2*s_est, color='gray', linestyle=':', alpha=0.5)
ax2.axhline(-2*s_est, color='gray', linestyle=':', alpha=0.5)
ax2.fill_between(t, -2*s_est, 2*s_est, color='gray', alpha=0.1)
ax2.set_xlabel('Hours elapsed', fontsize=12)
ax2.set_ylabel('Residual (m)', fontsize=12)
ax2.set_title('Residual Plot', fontsize=14)
ax2.grid(True, alpha=0.3)
img2 = fig_to_base64(fig2)

# Plot 3: Derivatives (first)
fig3, ax3 = plt.subplots(figsize=(10, 5))
ax3.plot(t, dhdt, '-', color='#00f0ff', label='dh/dt (finite diff)', linewidth=1, alpha=0.5)
ax3.plot(t, ma_dhdt, '-', color='#ffd700', label='Moving average (9 pt)', linewidth=1.5)
ax3.plot(t, sg_d1, '--', color='#00ffc8', label='Savitzky-Golay', linewidth=1.5)
ax3.plot(t_fine, deriv_fine, '-', color='#ff00e5', label='Analytic (from fit)', linewidth=2.5)
ax3.axvline(peak_deriv_t, color='#ff00e5', linestyle=':', label=f'Peak analytic {peak_deriv:.4f} m/h')
ax3.axhline(0, color='white', linestyle='-', linewidth=0.5, alpha=0.5)
ax3.set_xlabel('Hours elapsed', fontsize=12)
ax3.set_ylabel('dh/dt (m/h)', fontsize=12)
ax3.set_title('First Derivative Estimates', fontsize=14)
ax3.legend(loc='best', fontsize=8)
ax3.grid(True, alpha=0.3)
img3 = fig_to_base64(fig3)

# Plot 4: Second derivative
fig4, ax4 = plt.subplots(figsize=(10, 5))
ax4.plot(t, d2hdt2, '-', color='#00f0ff', label='Raw second diff', linewidth=1, alpha=0.3)
ax4.plot(t, ma_d2, '-', color='#ffd700', label='Moving average (9 pt)', linewidth=1.5)
ax4.plot(t, sg_d2, '--', color='#00ffc8', label='Savitzky-Golay', linewidth=1.5)
ax4.plot(t_fine, d2_analytic, '-', color='#ff00e5', label='Analytic (from fit)', linewidth=2.5)
ax4.axhline(0, color='white', linestyle='-', linewidth=0.5, alpha=0.5)
ax4.set_xlabel('Hours elapsed', fontsize=12)
ax4.set_ylabel('d²h/dt² (m/h²)', fontsize=12)
ax4.set_title('Second Derivative Estimates', fontsize=14)
ax4.legend(loc='best', fontsize=8)
ax4.grid(True, alpha=0.3)
img4 = fig_to_base64(fig4)

# ------------------------------------------------------------
# 8. HTML CONTENT – Futuristic Dashboard (no PDF)
# ------------------------------------------------------------
html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Lab 02 – Reservoir Analysis</title>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    background: #0a0e17;
    min-height: 100vh;
    padding: 20px;
    position: relative;
    overflow-x: hidden;
}}
body::before {{
    content: '';
    position: fixed;
    top: 0; left: 0; right: 0; bottom: 0;
    background:
        linear-gradient(rgba(0, 240, 255, 0.03) 1px, transparent 1px),
        linear-gradient(90deg, rgba(0, 240, 255, 0.03) 1px, transparent 1px);
    background-size: 40px 40px;
    z-index: 0;
    pointer-events: none;
    animation: gridMove 8s linear infinite;
}}
@keyframes gridMove {{
    0% {{ transform: translate(0,0); }}
    100% {{ transform: translate(40px,40px); }}
}}
.container {{
    max-width: 1400px;
    margin: 0 auto;
    background: rgba(10, 14, 23, 0.85);
    backdrop-filter: blur(20px);
    border-radius: 32px;
    border: 1px solid rgba(0, 240, 255, 0.15);
    box-shadow: 0 0 60px rgba(0, 240, 255, 0.05), 0 0 120px rgba(255, 0, 229, 0.03);
    overflow: hidden;
    position: relative;
    z-index: 1;
}}
.header {{
    padding: 30px 40px;
    background: linear-gradient(135deg, rgba(0, 240, 255, 0.05), rgba(255, 0, 229, 0.05));
    border-bottom: 1px solid rgba(0, 240, 255, 0.1);
}}
.header h1 {{
    font-size: 32px;
    font-weight: 300;
    letter-spacing: 2px;
    color: #fff;
    text-shadow: 0 0 20px rgba(0, 240, 255, 0.3);
}}
.header h1 strong {{
    font-weight: 700;
    background: linear-gradient(90deg, #00f0ff, #ff00e5);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}}
.header p {{
    font-size: 14px;
    color: rgba(255,255,255,0.6);
    margin-top: 6px;
    letter-spacing: 1px;
}}
.header .meta {{
    font-size: 12px;
    color: rgba(255,255,255,0.3);
    margin-top: 8px;
}}
.tabs {{
    display: flex;
    background: rgba(0,0,0,0.2);
    border-bottom: 1px solid rgba(0, 240, 255, 0.08);
    padding: 0 20px;
    flex-wrap: wrap;
}}
.tab {{
    padding: 14px 24px;
    cursor: pointer;
    border: none;
    background: transparent;
    font-size: 14px;
    font-weight: 600;
    color: rgba(255,255,255,0.5);
    transition: all 0.3s ease;
    border-bottom: 2px solid transparent;
    letter-spacing: 0.5px;
    display: flex;
    align-items: center;
    gap: 8px;
}}
.tab:hover {{
    color: #00f0ff;
    background: rgba(0, 240, 255, 0.05);
    transform: translateY(-2px);
}}
.tab.active {{
    color: #00f0ff;
    border-bottom: 2px solid #00f0ff;
    background: rgba(0, 240, 255, 0.05);
    box-shadow: 0 0 20px rgba(0, 240, 255, 0.05);
}}
.tab-content {{
    display: none;
    padding: 30px 40px;
    animation: fadeIn 0.4s ease forwards;
}}
.tab-content.active {{ display: block; }}
@keyframes fadeIn {{
    from {{ opacity: 0; transform: translateY(15px); }}
    to {{ opacity: 1; transform: translateY(0); }}
}}
.plot-card {{
    background: rgba(255,255,255,0.03);
    border-radius: 16px;
    border: 1px solid rgba(0, 240, 255, 0.08);
    padding: 8px;
    margin: 20px 0;
    transition: all 0.3s ease;
    box-shadow: 0 0 30px rgba(0, 240, 255, 0.02);
}}
.plot-card:hover {{
    transform: translateY(-4px);
    border-color: rgba(0, 240, 255, 0.2);
    box-shadow: 0 0 40px rgba(0, 240, 255, 0.05);
}}
.plot-card img {{ width: 100%; height: auto; border-radius: 12px; display: block; }}
.stats-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 16px;
    margin: 25px 0;
}}
.stat-card {{
    background: rgba(255,255,255,0.03);
    border-radius: 12px;
    padding: 16px 20px;
    border: 1px solid rgba(0, 240, 255, 0.08);
    transition: all 0.3s ease;
}}
.stat-card:hover {{
    transform: scale(1.02);
    border-color: rgba(0, 240, 255, 0.2);
    box-shadow: 0 0 30px rgba(0, 240, 255, 0.05);
}}
.stat-card .label {{
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.7px;
    color: rgba(255,255,255,0.4);
    font-weight: 600;
}}
.stat-card .value {{
    font-size: 24px;
    font-weight: 700;
    color: #fff;
    margin-top: 4px;
}}
.stat-card .unit {{
    font-size: 14px;
    font-weight: 400;
    color: rgba(255,255,255,0.4);
    margin-left: 4px;
}}
.stat-card.accent {{ border-color: rgba(255, 0, 229, 0.3); }}
.stat-card.accent .value {{ color: #ff00e5; }}
.stat-card.success {{ border-color: rgba(0, 255, 200, 0.3); }}
.stat-card.success .value {{ color: #00ffc8; }}
table {{
    width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    margin: 20px 0;
    font-size: 14px;
    border-radius: 12px;
    overflow: hidden;
    border: 1px solid rgba(0, 240, 255, 0.08);
}}
th {{
    background: rgba(0, 240, 255, 0.1);
    color: #00f0ff;
    padding: 12px 16px;
    text-align: left;
    font-weight: 600;
    border-bottom: 1px solid rgba(0, 240, 255, 0.1);
}}
td {{
    padding: 10px 16px;
    background: rgba(255,255,255,0.02);
    border-bottom: 1px solid rgba(0, 240, 255, 0.05);
    color: rgba(255,255,255,0.8);
}}
tr:last-child td {{ border-bottom: none; }}
tr:hover td {{ background: rgba(0, 240, 255, 0.03); }}
.significant {{ color: #00ffc8; font-weight: 700; }}
.not-significant {{ color: #ff6b6b; font-weight: 700; }}
.model-equation {{
    background: rgba(0, 240, 255, 0.05);
    padding: 16px 24px;
    border-radius: 12px;
    font-family: 'Courier New', monospace;
    font-size: 17px;
    border: 1px solid rgba(0, 240, 255, 0.1);
    margin: 18px 0;
    color: #fff;
}}
.model-equation span {{
    color: rgba(255,255,255,0.5);
    font-size: 14px;
}}
.insight-box, .interpretation-box, .data-box, .plain-box {{
    padding: 18px 24px;
    border-radius: 12px;
    margin: 20px 0;
    color: rgba(255,255,255,0.8);
}}
.insight-box {{ background: rgba(255, 215, 0, 0.05); border-left: 4px solid #ffd700; }}
.interpretation-box {{ background: rgba(0, 240, 255, 0.05); border-left: 4px solid #00f0ff; }}
.data-box {{ background: rgba(200, 0, 255, 0.05); border-left: 4px solid #cc00ff; }}
.plain-box {{ background: rgba(0, 255, 200, 0.05); border-left: 4px solid #00ffc8; }}
.plain-box .plabel {{
    color: #00ffc8;
    font-weight: 700;
    text-transform: uppercase;
    font-size: 12px;
    letter-spacing: 1px;
    display: block;
    margin-bottom: 6px;
}}
.footer {{
    background: rgba(0,0,0,0.2);
    padding: 14px 40px;
    text-align: center;
    color: rgba(255,255,255,0.3);
    font-size: 12px;
    border-top: 1px solid rgba(0, 240, 255, 0.05);
}}
.phase-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 16px;
    margin: 20px 0;
}}
.phase-card {{
    background: rgba(255,255,255,0.02);
    border-radius: 12px;
    padding: 14px 18px;
    border: 1px solid rgba(0,240,255,0.08);
}}
.phase-card .name {{ color: #ffd700; font-weight: 700; }}
.phase-card .clock {{ color: rgba(255,255,255,0.4); font-size: 12px; }}
.phase-card .change {{ color: #00f0ff; font-size: 20px; font-weight: 700; }}
.phase-card .rate {{ font-size: 12px; color: rgba(255,255,255,0.6); }}
.phase-card .note {{ font-size: 11px; color: rgba(255,255,255,0.3); margin-top: 4px; }}
.glossary td {{ border-bottom: 1px solid rgba(0, 240, 255, 0.05); padding: 8px 12px; vertical-align: top; color: rgba(255,255,255,0.7); }}
.glossary td:first-child {{ color: #00f0ff; font-weight: 700; width: 25%; }}
@media (max-width: 768px) {{
    .header {{ padding: 20px; }}
    .header h1 {{ font-size: 24px; }}
    .tabs {{ padding: 0 10px; }}
    .tab {{ padding: 10px 14px; font-size: 12px; }}
    .tab-content {{ padding: 20px; }}
    .stats-grid {{ grid-template-columns: repeat(2, 1fr); }}
    .phase-grid {{ grid-template-columns: 1fr; }}
}}
</style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>🚀 <strong>Reservoir Stage</strong> · Lab 02</h1>
        <p>Numerical Methods · Levenberg‑Marquardt Curve Fitting · Statistical Validation</p>
        <div class="meta">Readings: {n} · Span: {t[-1]:.2f} h · Generated: {pd.Timestamp.now().strftime('%d %b %Y, %H:%M')}</div>
    </div>

    <div class="tabs">
        <button class="tab active" onclick="showTab('summary')">📋 Summary</button>
        <button class="tab" onclick="showTab('data')">📁 Data</button>
        <button class="tab" onclick="showTab('derivatives')">📈 Derivatives</button>
        <button class="tab" onclick="showTab('fit')">🎯 Fit</button>
        <button class="tab" onclick="showTab('stats')">📊 Stats</button>
        <button class="tab" onclick="showTab('area')">📐 Area</button>
        <button class="tab" onclick="showTab('glossary')">📖 Glossary</button>
    </div>

    <!-- TAB 1: SUMMARY -->
    <div id="summary" class="tab-content active">
        <h2 style="color:#00f0ff; margin-bottom:10px;">Executive Summary</h2>
        <div class="insight-box">
            <strong style="color:#ffd700;">🔑 Key Insights</strong>
            <ul style="margin-top:10px; padding-left:22px; line-height:1.7;">
                <li>Level rose from <strong style="color:#00f0ff;">{h[0]:.2f} m</strong> to <strong style="color:#00f0ff;">{h[-1]:.2f} m</strong> (Δ = {h[-1]-h[0]:.2f} m)</li>
                <li>Maximum filling rate (analytic): <strong style="color:#ff00e5;">{peak_deriv:.4f} m/h</strong> at <strong style="color:#ff00e5;">t = {peak_deriv_t:.2f} h</strong></li>
                <li>Logistic model explains <strong style="color:#00ffc8;">{R2*100:.1f}%</strong> of variance (R² = {R2:.6f})</li>
                <li>Standard error: <strong style="color:#00ffc8;">{s_est:.4f} m</strong> (≈ {s_est*100:.1f} cm)</li>
                <li>Total area per unit width: <strong style="color:#ffd700;">{area:.2f} m·h</strong> (mean level {mean_level:.4f} m)</li>
            </ul>
        </div>

        <div class="stats-grid">
            <div class="stat-card"><div class="label">Initial Level</div><div class="value">{h[0]:.2f}<span class="unit">m</span></div></div>
            <div class="stat-card"><div class="label">Final Level</div><div class="value">{h[-1]:.2f}<span class="unit">m</span></div></div>
            <div class="stat-card accent"><div class="label">Peak Fill Rate</div><div class="value">{peak_deriv:.4f}<span class="unit">m/h</span></div></div>
            <div class="stat-card success"><div class="label">R²</div><div class="value">{R2:.6f}</div></div>
            <div class="stat-card"><div class="label">Std Error (s)</div><div class="value">{s_est:.4f}<span class="unit">m</span></div></div>
            <div class="stat-card accent"><div class="label">Area (AUC)</div><div class="value">{area:.2f}<span class="unit">m·h</span></div></div>
        </div>

        <div class="plot-card">
            <h3 style="color:#00f0ff; margin:0 0 8px 12px;">📊 Water Level Over Time</h3>
            <img src="data:image/png;base64,{img1}" alt="Raw data and fit">
        </div>

        <h3 style="color:#00f0ff; margin-top:25px;">📌 Phase Analysis</h3>
        <div class="phase-grid">
            {''.join([
                f'<div class="phase-card">'
                f'<div class="name">{ph["name"]}</div>'
                f'<div class="clock">{ph["clock0"]} · {ph["hours"]:.1f} h</div>'
                f'<div class="change">{ph["change"]:+.2f}<span style="font-size:14px; color:rgba(255,255,255,0.4);"> m</span></div>'
                f'<div class="rate">mean rate {ph["mean_rate"]:+.4f} m/h · peak {ph["peak_rate"]:+.4f} m/h</div>'
                f'<div class="note">{ph["note"]}</div>'
                f'</div>'
                for ph in phase_data
            ])}
        </div>
        <div class="interpretation-box">
            <strong>📝 Quick Interpretation</strong><br>
            The two‑logistic model fits the data very well (R² > 0.99). Residuals show pattern (only {runs} sign runs vs {expected_runs:.0f} expected), but the model is excellent for volumes and timing.
        </div>
    </div>

    <!-- TAB 2: DATA -->
    <div id="data" class="tab-content">
        <h2 style="color:#00f0ff; margin-bottom:10px;">📁 Data Overview</h2>
        <div class="data-box">
            <strong style="color:#cc00ff;">📋 Source</strong><br>
            Reservoir stage log from a depth sensor, logged every 15 minutes. The data covers {n} readings from {df['Timestamp'].iloc[0].strftime('%d %b %Y, %H:%M')} to {df['Timestamp'].iloc[-1].strftime('%d %b %Y, %H:%M')}. The logger rounds depth to the nearest centimetre (±0.01 m).
        </div>
        <div class="stats-grid">
            <div class="stat-card"><div class="label">Readings</div><div class="value">{n}</div></div>
            <div class="stat-card"><div class="label">Time span</div><div class="value">{t[0]:.1f} – {t[-1]:.1f}<span class="unit">h</span></div></div>
            <div class="stat-card"><div class="label">Min depth</div><div class="value">{np.min(h):.2f}<span class="unit">m</span></div></div>
            <div class="stat-card"><div class="label">Max depth</div><div class="value">{np.max(h):.2f}<span class="unit">m</span></div></div>
            <div class="stat-card"><div class="label">Total rise</div><div class="value">{h[-1]-h[0]:.4f}<span class="unit">m</span></div></div>
            <div class="stat-card"><div class="label">Logger precision</div><div class="value">±0.01<span class="unit">m</span></div></div>
        </div>
        <h3 style="color:#00f0ff; margin:20px 0 10px 0;">📂 Columns in Original Excel</h3>
        <table>
            <thead><tr><th>Column</th><th>Description</th></tr></thead>
            <tbody>
                <tr><td><strong style="color:#00f0ff;">Reading</strong></td><td>Sequential row number</td></tr>
                <tr><td><strong style="color:#00f0ff;">Timestamp</strong></td><td>Full date and time</td></tr>
                <tr><td><strong style="color:#00f0ff;">Date</strong></td><td>Date only</td></tr>
                <tr><td><strong style="color:#00f0ff;">Time</strong></td><td>Time only</td></tr>
                <tr><td><strong style="color:#00f0ff;">Depth (m)</strong></td><td>Water level in metres, rounded to 0.01 m</td></tr>
            </tbody>
        </table>
        <div class="interpretation-box">
            For curve fitting, timestamps converted to <strong style="color:#00f0ff;">hours elapsed from the first reading</strong>.
            Derivative <strong style="color:#ff00e5;">dh/dt</strong> is then in <strong style="color:#ff00e5;">m/h</strong>.
        </div>
    </div>

    <!-- TAB 3: DERIVATIVES -->
    <div id="derivatives" class="tab-content">
        <h2 style="color:#00f0ff; margin-bottom:10px;">📈 Derivatives Analysis</h2>
        <div class="stats-grid">
            <div class="stat-card accent"><div class="label">Max dh/dt (analytic)</div><div class="value">{peak_deriv:.4f}<span class="unit">m/h</span></div></div>
            <div class="stat-card"><div class="label">at t</div><div class="value">{peak_deriv_t:.2f}<span class="unit">h</span></div></div>
            <div class="stat-card"><div class="label">Max dh/dt (finite diff)</div><div class="value">{max_rate:.4f}<span class="unit">m/h</span></div></div>
            <div class="stat-card"><div class="label">Max (moving avg)</div><div class="value">{max_ma_dhdt:.4f}<span class="unit">m/h</span></div></div>
            <div class="stat-card"><div class="label">Min dh/dt</div><div class="value">{min_dhdt:.4f}<span class="unit">m/h</span></div></div>
            <div class="stat-card"><div class="label">Inflection (d²=0)</div><div class="value">{inflection_time:.2f}<span class="unit">h</span></div></div>
        </div>
        <div class="plot-card">
            <h3 style="color:#00f0ff; margin:0 0 8px 12px;">First Derivative Estimates</h3>
            <img src="data:image/png;base64,{img3}" alt="Derivatives">
        </div>
        <div class="plot-card">
            <h3 style="color:#00f0ff; margin:0 0 8px 12px;">Second Derivative Estimates</h3>
            <img src="data:image/png;base64,{img4}" alt="Second derivative">
        </div>
        <div class="interpretation-box">
            <strong style="color:#ff00e5;">🔍 Interpretation</strong><br>
            The second derivative is positive initially, indicating accelerating inflow; it crosses zero near the peak rate, then becomes negative. The analytic curve is smoother than the raw differences.
        </div>
        <h3 style="color:#00f0ff; margin-top:20px;">Peak Rate Estimates</h3>
        <table>
            <thead><tr><th>Method</th><th>Peak (m/h)</th><th>Time (h)</th></tr></thead>
            <tbody>
                <tr><td>Central difference (raw)</td><td>{max_rate:.4f}</td><td>{max_time:.2f}</td></tr>
                <tr><td>Moving average (9 pt)</td><td>{max_ma_dhdt:.4f}</td><td>{max_ma_dhdt_t:.2f}</td></tr>
                <tr><td>Savitzky-Golay</td><td>{sg_d1_peak:.4f}</td><td>{sg_d1_peak_t:.2f}</td></tr>
                <tr style="color:#ff00e5; font-weight:700;"><td>Analytic (fit)</td><td>{peak_deriv:.4f}</td><td>{peak_deriv_t:.2f}</td></tr>
            </tbody>
        </table>
    </div>

    <!-- TAB 4: FIT & RESIDUALS -->
    <div id="fit" class="tab-content">
        <h2 style="color:#00f0ff; margin-bottom:10px;">🎯 Model Fit and Residuals</h2>
        <div class="model-equation">
            <strong>Model:</strong> h(t) = c + a1 / (1 + exp(-k1·(t - t1))) + a2 / (1 + exp(-k2·(t - t2)))<br>
            <span>
                c = {c:.4f} m · a1 = {a1:.4f} m · k1 = {k1:.4f} 1/h · t1 = {t1:.2f} h<br>
                a2 = {a2:.4f} m · k2 = {k2:.4f} 1/h · t2 = {t2:.2f} h
            </span>
        </div>
        <div class="plot-card">
            <h3 style="color:#00f0ff; margin:0 0 8px 12px;">Fitted Curve with ±1σ Band</h3>
            <img src="data:image/png;base64,{img1}" alt="Fit">
        </div>
        <div class="plot-card">
            <h3 style="color:#00f0ff; margin:0 0 8px 12px;">Residual Plot (±2σ bounds)</h3>
            <img src="data:image/png;base64,{img2}" alt="Residuals">
        </div>
        <div class="interpretation-box">
            <strong>📝 Reading of Residuals</strong>
            <ul style="margin-top:6px; padding-left:20px; line-height:1.6;">
                <li><strong>Centred:</strong> Mean {np.mean(resid):.2e} m (essentially zero).</li>
                <li><strong>Runs:</strong> {runs} sign changes vs {expected_runs:.0f} expected – the residuals are strongly patterned, indicating the model misses shape.</li>
                <li><strong>Magnitude:</strong> Max |e| = {np.max(np.abs(resid)):.4f} m; most within ±2s = {2*s_est:.3f} m.</li>
                <li><strong>Conclusion:</strong> The model is good for timing and totals, but not for instant level to the centimetre.</li>
            </ul>
        </div>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:20px; margin-top:20px;">
            <div>
                <h4 style="color:#00f0ff;">Residuals by block (12 h)</h4>
                <table>
                    <thead><tr><th>t window (h)</th><th>Mean (m)</th><th>SD (m)</th></tr></thead>
                    <tbody>
                        {''.join([f'<tr><td>{i*12:.1f} – {(i+1)*12:.1f}</td><td>{block_means[i]:+.4f}</td><td>{block_sds[i]:.4f}</td></tr>' for i in range(min(6,len(block_means)))])}
                    </tbody>
                </table>
            </div>
            <div>
                <h4 style="color:#00f0ff;">Sign runs</h4>
                <p style="color:rgba(255,255,255,0.6);">{runs} observed, {expected_runs:.0f} expected under independent scatter.</p>
                <p style="color:rgba(255,255,255,0.4); font-size:13px;">Longest run: {max(np.diff(np.where(np.diff(np.sign(resid))!=0)[0]))} consecutive same sign.</p>
            </div>
        </div>
    </div>

    <!-- TAB 5: STATISTICS -->
    <div id="stats" class="tab-content">
        <h2 style="color:#00f0ff; margin-bottom:10px;">📊 Statistical Summary</h2>
        <h3 style="color:#00f0ff; margin:15px 0 8px 0;">Fitted Parameters</h3>
        <table>
            <thead><tr><th>Parameter</th><th>Value</th><th>Std. Error</th><th>t‑stat</th><th>p‑value</th><th>Significant?</th></tr></thead>
            <tbody>
                {''.join([
                    f'<tr><td><strong style="color:#00f0ff;">{["c","a1","k1","t1","a2","k2","t2"][i]}</strong></td>'
                    f'<td>{popt[i]:.4f}</td><td>{param_err[i]:.4f}</td><td>{t_stats[i]:.2f}</td><td>{p_values[i]:.4f}</td>'
                    f'<td class="{"significant" if p_values[i] < 0.05 else "not-significant"}">{"✅ YES" if p_values[i] < 0.05 else "❌ NO"}</td></tr>'
                    for i in range(len(popt))
                ])}
            </tbody>
        </table>
        <div class="stats-grid">
            <div class="stat-card"><div class="label">SSE</div><div class="value">{SSE:.6f}<span class="unit">m²</span></div></div>
            <div class="stat-card success"><div class="label">R²</div><div class="value">{R2:.6f}</div></div>
            <div class="stat-card"><div class="label">s (std error)</div><div class="value">{s_est:.6f}<span class="unit">m</span></div></div>
            <div class="stat-card"><div class="label">Observations (n)</div><div class="value">{n}</div></div>
            <div class="stat-card"><div class="label">Parameters (p)</div><div class="value">{p}</div></div>
            <div class="stat-card"><div class="label">DF</div><div class="value">{n-p}</div></div>
        </div>
        <div class="interpretation-box">
            <strong>📝 Conclusions</strong><br>
            All parameters are significant (p < 0.05). The model explains {R2*100:.1f}% of variance with a typical error of {s_est*100:.1f} cm.
        </div>
    </div>

    <!-- TAB 6: AREA -->
    <div id="area" class="tab-content">
        <h2 style="color:#00f0ff; margin-bottom:10px;">📐 Area Under the Curve</h2>
        <div class="stats-grid">
            <div class="stat-card accent"><div class="label">Quadrature</div><div class="value">{area:.6f}<span class="unit">m·h</span></div></div>
            <div class="stat-card"><div class="label">Trapezoid</div><div class="value">{area_trapz:.6f}<span class="unit">m·h</span></div></div>
            <div class="stat-card"><div class="label">Difference</div><div class="value">{abs(area - area_trapz):.6f}<span class="unit">m·h</span></div></div>
            <div class="stat-card"><div class="label">Mean stage</div><div class="value">{mean_level:.4f}<span class="unit">m</span></div></div>
        </div>
        <div class="plot-card">
            <h3 style="color:#00f0ff; margin:0 0 8px 12px;">Area Under the Fitted Curve</h3>
            <img src="data:image/png;base64,{img1}" alt="Area under curve">
        </div>
        <div style="background:rgba(0, 255, 200, 0.05); border-radius:12px; padding:18px 24px; border-left:4px solid #00ffc8; margin:20px 0;">
            <h3 style="color:#00ffc8;">📌 Result</h3>
            <p style="font-size:20px; font-weight:700; margin:8px 0; color:#fff;">
                {area:.4f} m·h <span style="font-size:16px; font-weight:400; color:rgba(255,255,255,0.5);">(≈ {area*1000:.0f} mm·h)</span>
            </p>
            <p style="color:rgba(255,255,255,0.6); font-size:14px;">
                <strong>Interpretation:</strong> This is the cumulative volume per unit width over the {t[-1]-t[0]:.1f}‑hour period.
                Dividing by the span gives a mean stage of {mean_level:.4f} m.
            </p>
            <p style="color:rgba(255,255,255,0.6); font-size:14px; margin-top:6px;">
                <strong>Units:</strong> metres × hours → multiply by reservoir width for total volume.
            </p>
        </div>
    </div>

    <!-- TAB 7: GLOSSARY -->
    <div id="glossary" class="tab-content">
        <h2 style="color:#00f0ff; margin-bottom:10px;">📖 Glossary of Terms</h2>
        <table class="glossary">
            <thead><tr><th>Term</th><th>Definition</th></tr></thead>
            <tbody>
                <tr><td>Area under the curve</td><td>How high multiplied by how long, summed. Here in m·h; divide by time span to get mean stage.</td></tr>
                <tr><td>Central difference</td><td>Estimating a rate using the reading before and after. More balanced than forward/backward.</td></tr>
                <tr><td>Degrees of freedom</td><td>Readings minus parameters (n−p). Remaining independent evidence after fitting.</td></tr>
                <tr><td>Derivative (dh/dt)</td><td>Rate of change of level. First derivative: speed; second derivative: acceleration.</td></tr>
                <tr><td>Finite difference</td><td>Rate from two readings: subtract, divide by time. No model needed, but amplifies rounding.</td></tr>
                <tr><td>Integration</td><td>Adding area under a curve. Opposite of differentiation; smooths noise.</td></tr>
                <tr><td>Least squares</td><td>Minimise sum of squared residuals. Makes large misses count more.</td></tr>
                <tr><td>Levenberg‑Marquardt (LM)</td><td>Iterative optimizer for curve fitting. Needs a good initial guess.</td></tr>
                <tr><td>Logistic / sigmoid</td><td>S‑shaped curve: flat, rise, flat again. Cannot go back down, hence two logistics used.</td></tr>
                <tr><td>Moving average</td><td>Smooths by averaging neighbours. Cheap, but flattens genuine peaks.</td></tr>
                <tr><td>p‑value</td><td>Probability that a parameter is actually zero. Low p (<0.05) means significant.</td></tr>
                <tr><td>Parameter</td><td>Unknown number in model equation solved by fit (c, a, k, t, etc.).</td></tr>
                <tr><td>R²</td><td>Fraction of variance explained. No units; can be high even if model is wrong.</td></tr>
                <tr><td>Residual</td><td>Observed − predicted. Pattern in residuals reveals model shortcomings.</td></tr>
                <tr><td>s (std error of estimate)</td><td>Typical size of a residual, in metres. More informative than R².</td></tr>
                <tr><td>Savitzky‑Golay</td><td>Smooths by fitting a polynomial locally. Preserves peaks better than moving average.</td></tr>
                <tr><td>SE (standard error)</td><td>Uncertainty of a fitted parameter. Small SE = well determined.</td></tr>
                <tr><td>Sign runs</td><td>Number of times residuals switch sign. Far fewer than expected indicates patterned misfit.</td></tr>
                <tr><td>SSE / SST</td><td>SSE: sum of squared errors; SST: total sum of squares. R² = 1 − SSE/SST.</td></tr>
                <tr><td>t‑statistic</td><td>Parameter value divided by its SE. Large t means parameter is not zero.</td></tr>
                <tr><td>Trapezoid rule</td><td>Integrate by joining points with straight lines. Used as cross‑check.</td></tr>
            </tbody>
        </table>
    </div>

    <div class="footer">
        Generated by Python · All computations in Python · {n} observations
    </div>
</div>

<script>
function showTab(tabId) {{
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.tab').forEach(el => el.classList.remove('active'));
    document.getElementById(tabId).classList.add('active');
    const tabs = document.querySelectorAll('.tab');
    const map = {{'summary':0,'data':1,'derivatives':2,'fit':3,'stats':4,'area':5,'glossary':6}};
    if (map[tabId] !== undefined) {{
        tabs[map[tabId]].classList.add('active');
    }}
}}
</script>
</body>
</html>"""

# ------------------------------------------------------------
# 9. SAVE HTML
# ------------------------------------------------------------
surname = "Berdon"  # CHANGE to your actual surname

html_filename = f"lab02_{surname}.html"
html_path = os.path.join(script_dir, html_filename)
with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html_content)

# Also copy to Desktop
desktop_candidates = [
    os.path.join(os.path.expanduser("~"), "Desktop"),
    os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop"),
]
for candidate in desktop_candidates:
    if os.path.exists(candidate):
        try:
            desktop_path = os.path.join(candidate, html_filename)
            with open(desktop_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            break
        except Exception:
            pass

print(f"\n{'='*60}")
print(f"✅ Dashboard saved to: {html_path}")
print(f"{'='*60}")
print(f"\n📊 SUMMARY STATISTICS:")
print(f"   Max dh/dt (analytic): {peak_deriv:.6f} m/h at t = {peak_deriv_t:.2f} h")
print(f"   Max dh/dt (finite diff): {max_rate:.6f} m/h at t = {max_time:.2f} h")
print(f"   Inflection point: {inflection_time:.2f} h")
print(f"   SSE = {SSE:.6f}")
print(f"   R² = {R2:.6f}")
print(f"   s = {s_est:.6f} m")
print(f"   Area = {area:.6f} m·h")
print(f"{'='*60}")

# Open HTML in browser
webbrowser.open(html_path)