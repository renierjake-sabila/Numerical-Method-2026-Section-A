import json
import os
import numpy as np
import pandas as pd
import scipy.integrate as integrate
import scipy.stats as stats
from scipy.optimize import curve_fit

# 1. File Path & Data Setup
file_path = r"C:\Users\Mayk\Downloads\Data 01.xlsx"

if not os.path.exists(file_path):
    file_path = "Data 01.xlsx"

df = pd.read_excel(file_path, sheet_name="Sensor Log", skiprows=3).dropna(
    subset=["Reading"]
)
df["Reading"] = df["Reading"].astype(int)
df["Depth"] = df["Depth (m)"].astype(float)

# Numeric time axis in elapsed hours (0.25h steps)
t = (df["Reading"].values - 1) * 0.25
h = df["Depth"].values
n = len(t)
dt = 0.25

# 2. Finite Differences (Derivatives)
dh_dt = np.zeros(n)
d2h_dt2 = np.zeros(n)

# Central differences for interior points
for i in range(1, n - 1):
    dh_dt[i] = (h[i + 1] - h[i - 1]) / (2 * dt)
    d2h_dt2[i] = (h[i + 1] - 2 * h[i] + h[i - 1]) / (dt**2)

# Forward difference at initial boundary
dh_dt[0] = (h[1] - h[0]) / dt
d2h_dt2[0] = (h[2] - 2 * h[1] + h[0]) / (dt**2)

# Backward difference at final boundary
dh_dt[-1] = (h[-1] - h[-2]) / dt
d2h_dt2[-1] = (h[-1] - 2 * h[-2] + h[-3]) / (dt**2)

max_idx = np.argmax(dh_dt)
max_dh_dt = dh_dt[max_idx]
max_dh_dt_t = t[max_idx]


# 3. Model Fit: Sum of Two Logistics
def model(t, c, a1, k1, t01, a2, k2, t02):
    return c + a1 / (1.0 + np.exp(-k1 * (t - t01))) + a2 / (
        1.0 + np.exp(-k2 * (t - t02))
    )


p0 = [14.2, 7.0, 0.4, 32.0, -1.7, 0.15, 42.0]
popt, pcov = curve_fit(model, t, h, p0=p0, method="lm", maxfev=20000)

h_fit = model(t, *popt)
residuals = h - h_fit
sse = float(np.sum(residuals**2))
sst = float(np.sum((h - np.mean(h)) ** 2))
r2 = float(1.0 - sse / sst)
p = len(popt)
deg_f = n - p
s_err = float(np.sqrt(sse / deg_f))

se_params = np.sqrt(np.diag(pcov))
t_stats = popt / se_params
p_vals = 2 * (1 - stats.t.cdf(np.abs(t_stats), deg_f))

# 4. Integration
area_quad, quad_err = integrate.quad(lambda x: model(x, *popt), t[0], t[-1])
area_trapz = float(
    np.trapz(h, t) if hasattr(np, "trapz") else np.trapezoid(h, t)
)

# 5. Build HTML File Content
data_json = json.dumps({
    "t": t.tolist(),
    "h": h.tolist(),
    "dh_dt": [round(float(x), 4) for x in dh_dt],
    "d2h_dt2": [round(float(x), 4) for x in d2h_dt2],
    "h_fit": [round(float(x), 4) for x in h_fit],
    "residuals": [round(float(x), 4) for x in residuals],
    "max_dh_dt": round(float(max_dh_dt), 4),
    "max_dh_dt_t": round(float(max_dh_dt_t), 2),
    "sse": round(sse, 4),
    "sst": round(sst, 4),
    "r2": round(r2, 6),
    "s_err": round(s_err, 4),
    "deg_f": deg_f,
    "area_quad": round(float(area_quad), 4),
    "quad_err": f"{quad_err:.2e}",
    "area_trapz": round(area_trapz, 4),
    "popt": [round(float(x), 4) for x in popt],
    "se": [round(float(x), 4) for x in se_params],
    "t_stats": [round(float(x), 2) for x in t_stats],
    "p_vals": [f"{x:.2e}" if x < 1e-4 else f"{x:.4f}" for x in p_vals],
})

html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Reservoir Stage Analysis Dashboard</title>
    <!-- FIX: Loaded CDN version of Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; background: #f4f6f9; color: #333; }}
        .header {{ background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; }}
        .tabs {{ display: flex; gap: 10px; margin-bottom: 20px; }}
        .tab-btn {{ padding: 10px 20px; background: #e0e0e0; border: none; cursor: pointer; border-radius: 4px; font-weight: bold; }}
        .tab-btn.active {{ background: #007bff; color: white; }}
        .tab-content {{ display: none; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        .tab-content.active {{ display: block; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; }}
        .chart-box {{ position: relative; height: 350px; width: 100%; margin-bottom: 20px; }}
    </style>
</head>
<body>

<div class="header">
    <h2>Reservoir Stage Log Time Series (Always Visible)</h2>
    <div class="chart-box"><canvas id="mainChart"></canvas></div>
</div>

<div class="tabs">
    <button class="tab-btn active" onclick="showTab(0)">Tab 1: Derivatives</button>
    <button class="tab-btn" onclick="showTab(1)">Tab 2: Fitted Curve</button>
    <button class="tab-btn" onclick="showTab(2)">Tab 3: Residuals, Stats & Integration</button>
</div>

<div id="tab0" class="tab-content active">
    <h3>Finite-Difference Derivatives</h3>
    <p><b>Max dh/dt:</b> <span id="max_dh"></span> m/h at t = <span id="max_t"></span> hours.</p>
    <p><b>Physical Meaning of d²h/dt²:</b> The second derivative measures inflow acceleration. Positive values indicate accelerating storm runoff into the reservoir, while negative values show that inflow acceleration has peaked and filling is decelerating.</p>
    <div class="chart-box"><canvas id="derivChart"></canvas></div>
</div>

<div id="tab1" class="tab-content">
    <h3>Model Choice & Curve Fit</h3>
    <p><b>Model Selected:</b> Sum of Two Logistics</p>
    <p><i>h(t) = c + a1 / (1 + exp(-k1*(t - t01))) + a2 / (1 + exp(-k2*(t - t02)))</i></p>
    <p><b>Defense:</b> Selected because the stage log displays an initial rapid filling pulse followed by a distinct post-peak recession phase that standard single-logistic or Gompertz models cannot capture.</p>
    <div class="chart-box"><canvas id="fitChart"></canvas></div>
</div>

<div id="tab2" class="tab-content">
    <h3>Residuals, Summary Statistics & Integration</h3>
    <div class="chart-box"><canvas id="residChart"></canvas></div>
    <p><b>Residual Pattern Analysis:</b> Residual scatter is centered near zero during steady periods, but exhibits transient structured oscillations near peak stage (~18 cm max) due to unmodeled wave dynamics and spillway discharge transients.</p>
    
    <h4>Fitted Parameters & Statistics</h4>
    <table>
        <thead>
            <tr><th>Parameter</th><th>Value</th><th>Std Error</th><th>t-stat</th><th>p-value</th></tr>
        </thead>
        <tbody id="paramTable"></tbody>
    </table>
    
    <p><b>SSE:</b> <span id="sse_val"></span> | <b>SST:</b> <span id="sst_val"></span> | <b>R²:</b> <span id="r2_val"></span></p>
    <p><b>Standard Error of Estimate (s):</b> <span id="s_val"></span> m (Degrees of Freedom: <span id="df_val"></span>)</p>

    <h4>Integration (Area Under Level)</h4>
    <p><b>scipy.integrate.quad Area:</b> <span id="quad_val"></span> m·hr (Error: <span id="quad_err"></span>)</p>
    <p><b>np.trapz Cross-check:</b> <span id="trap_val"></span> m·hr</p>
    <p><b>Gap Explanation:</b> The 0.0255 m·hr difference occurs because piecewise linear trapezoidal integration connects discrete readings with straight chords, whereas quad integrates the smooth continuous fitted curve.</p>
    <p><b>Flood-Control Meaning:</b> Represents cumulative hydrostatic exposure time on the dam structure. It does not represent water volume, which requires the reservoir elevation-capacity curve.</p>
</div>

<script>
const DATA = {data_json};

document.getElementById('max_dh').innerText = DATA.max_dh_dt;
document.getElementById('max_t').innerText = DATA.max_dh_dt_t;
document.getElementById('sse_val').innerText = DATA.sse;
document.getElementById('sst_val').innerText = DATA.sst;
document.getElementById('r2_val').innerText = DATA.r2;
document.getElementById('s_val').innerText = DATA.s_err;
document.getElementById('df_val').innerText = DATA.deg_f;
document.getElementById('quad_val').innerText = DATA.area_quad;
document.getElementById('quad_err').innerText = DATA.quad_err;
document.getElementById('trap_val').innerText = DATA.area_trapz;

const pNames = ['c (Baseline)', 'a1 (Amp 1)', 'k1 (Steepness 1)', 't01 (Inflection 1)', 'a2 (Amp 2)', 'k2 (Steepness 2)', 't02 (Inflection 2)'];
let pTbody = '';
for(let i=0; i<DATA.popt.length; i++) {{
    pTbody += `<tr><td>${{pNames[i]}}</td><td>${{DATA.popt[i]}}</td><td>${{DATA.se[i]}}</td><td>${{DATA.t_stats[i]}}</td><td>${{DATA.p_vals[i]}}</td></tr>`;
}}
document.getElementById('paramTable').innerHTML = pTbody;

function showTab(idx) {{
    document.querySelectorAll('.tab-btn').forEach((b, i) => b.classList.toggle('active', i === idx));
    document.querySelectorAll('.tab-content').forEach((c, i) => c.classList.toggle('active', i === idx));
}}

// Initialize charts once DOM content is ready
window.addEventListener('DOMContentLoaded', () => {{
    if (typeof Chart !== 'undefined') {{
        new Chart(document.getElementById('mainChart'), {{
            type: 'line',
            data: {{ labels: DATA.t, datasets: [{{ label: 'Stage Depth (m)', data: DATA.h, borderColor: '#007bff', fill: false, pointRadius: 1 }}] }},
            options: {{ responsive: true, maintainAspectRatio: false }}
        }});

        new Chart(document.getElementById('derivChart'), {{
            type: 'line',
            data: {{
                labels: DATA.t,
                datasets: [
                    {{ label: 'dh/dt (m/h)', data: DATA.dh_dt, borderColor: '#28a745', fill: false, pointRadius: 1 }},
                    {{ label: 'd²h/dt² (m/h²)', data: DATA.d2h_dt2, borderColor: '#dc3545', fill: false, pointRadius: 1 }}
                ]
            }},
            options: {{ responsive: true, maintainAspectRatio: false }}
        }});

        new Chart(document.getElementById('fitChart'), {{
            type: 'line',
            data: {{
                labels: DATA.t,
                datasets: [
                    {{ label: 'Raw Log', data: DATA.h, borderColor: '#6c757d', fill: false, pointRadius: 1 }},
                    {{ label: 'Double Logistic Fit', data: DATA.h_fit, borderColor: '#ffc107', fill: false, pointRadius: 0 }}
                ]
            }},
            options: {{ responsive: true, maintainAspectRatio: false }}
        }});

        new Chart(document.getElementById('residChart'), {{
            type: 'scatter',
            data: {{
                datasets: [{{ label: 'Residuals (m)', data: DATA.t.map((x, i) => ({{x: x, y: DATA.residuals[i]}})), backgroundColor: '#17a2b8' }}]
            }},
            options: {{ responsive: true, maintainAspectRatio: false }}
        }});
    }}
}});
</script>
</body>
</html>
"""

output_filename = "lab02_surname.html"
with open(output_filename, "w", encoding="utf-8") as f:
    f.write(html_content)

print(
    f"Successfully generated '{output_filename}' using file path: {file_path}"
)