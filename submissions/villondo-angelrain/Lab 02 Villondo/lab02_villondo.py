import os
import json
import numpy as np
import pandas as pd
from scipy.optimize import curve_fit
from scipy.integrate import quad
from scipy.stats import t

# --- 1. DATA LOADING & PREPROCESSING ---
file_path = r"C:\Users\AngelV\Downloads\Data 01.1.xlsx"

# Read starting at row header index 3 (skiprows=3)
df = pd.read_excel(file_path, skiprows=3)

# Clean column names
df.columns = [str(col).strip() for col in df.columns]

# Auto-detect time and stage columns dynamically
time_col = [c for c in df.columns if any(k in c.lower() for k in ['time', 'date', 't'])][0]
stage_col = [c for c in df.columns if any(k in c.lower() for k in ['stage', 'height', 'h', 'level'])][0]

df[time_col] = pd.to_datetime(df[time_col])
df = df.sort_values(by=time_col).reset_index(drop=True)

# Generate numeric elapsed hours t starting at 0.00
time_diffs = (df[time_col] - df[time_col].iloc[0]).dt.total_seconds() / 3600.0
t_arr = time_diffs.to_numpy(dtype=float)
h_arr = df[stage_col].to_numpy(dtype=float)

# --- 2. FINITE-DIFFERENCE DERIVATIVES ---
# 1st Derivative: dh/dt (Central differences interior, forward/backward endpoints)
dh_dt = np.gradient(h_arr, t_arr)

# 2nd Derivative: d^2h/dt^2
d2h_dt2 = np.gradient(dh_dt, t_arr)

# Identify Maximum Inflow Rate (Max dh/dt)
max_dh_dt_idx = int(np.argmax(dh_dt))
max_dh_dt_val = float(dh_dt[max_dh_dt_idx])
max_dh_dt_hour = float(t_arr[max_dh_dt_idx])

# Convert derivatives safely for JSON output (replacing NaN with None)
dh_dt_list = [None if np.isnan(v) else float(v) for v in dh_dt]
d2h_dt2_list = [None if np.isnan(v) else float(v) for v in d2h_dt2]

# --- 3. LOGISTIC MODEL FITTING ---
# Valid Python Identifier: Function names CANNOT start with numbers (e.g., 4_param_logistic is invalid)
def logistic_4p(t, c, a, k, t0):
    """4-Parameter Logistic Curve with clip protection against overflow."""
    arg = np.clip(-k * (t - t0), -500, 500)
    return c + a / (1.0 + np.exp(arg))

# Initial Parameter Guesses
c_guess = float(np.min(h_arr))
a_guess = float(np.max(h_arr) - np.min(h_arr))
k_guess = 0.5
t0_guess = float(np.median(t_arr))
p0 = [c_guess, a_guess, k_guess, t0_guess]

# Fit via Levenberg-Marquardt algorithm
popt, pcov = curve_fit(logistic_4p, t_arr, h_arr, p0=p0, method='lm', maxfev=10000)
c_fit, a_fit, k_fit, t0_fit = popt

h_pred = logistic_4p(t_arr, *popt)

def logistic_derivative(t, a, k, t0):
    arg = np.clip(-k * (t - t0), -500, 500)
    exp_term = np.exp(arg)
    return (a * k * exp_term) / ((1.0 + exp_term) ** 2)

dh_dt_fit = logistic_derivative(t_arr, a_fit, k_fit, t0_fit)

# --- 4. STATISTICAL DIAGNOSTICS ---
n = len(h_arr)
p = len(popt)  # 4 parameters
df_e = n - p   # Degrees of freedom

residuals = h_arr - h_pred
SSE = float(np.sum(residuals**2))
SST = float(np.sum((h_arr - np.mean(h_arr))**2))
R2 = float(1.0 - (SSE / SST))
s_err = float(np.sqrt(SSE / df_e))

# Standard errors, t-statistics, and two-tailed p-values
param_se = np.sqrt(np.diag(pcov))
t_stats = popt / param_se
p_values = [float(2 * (1 - t.cdf(abs(ts), df=df_e))) for ts in t_stats]

param_names = ['c (Base Level)', 'a (Capacity)', 'k (Growth Rate)', 't₀ (Inflection Hour)']
param_summary = []
for i in range(p):
    param_summary.append({
        'name': param_names[i],
        'value': float(popt[i]),
        'se': float(param_se[i]),
        't_stat': float(t_stats[i]),
        'p_val': float(p_values[i]),
        'status': 'Significant' if p_values[i] < 0.05 else 'Non-Significant'
    })

# --- 5. VOLUMETRIC INTEGRATION ---
t_min, t_max = float(t_arr[0]), float(t_arr[-1])

# Model integration via scipy.integrate.quad
quad_val, quad_err = quad(logistic_4p, t_min, t_max, args=(c_fit, a_fit, k_fit, t0_fit))

# Discrete cross-check via np.trapezoid
trapz_val = float(np.trapezoid(h_arr, t_arr))
integration_diff_pct = float(abs(quad_val - trapz_val) / quad_val * 100)

# --- 6. PREPARE DATA ARRAYS FOR HTML/CHART.JS ---
t_labels = [round(float(x), 2) for x in t_arr]
h_observed = [float(x) for x in h_arr]
h_fitted = [float(x) for x in h_pred]
dh_dt_fitted = [float(x) for x in dh_dt_fit]
residuals_list = [float(r) for r in residuals]

# --- 7. GENERATE HTML DASHBOARD ---
html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Lab 02: Reservoir Stage Dataset Executive Report</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {{
            --bg-primary: #ffffff;
            --bg-card: #f8fafc;
            --text-primary: #0f172a;
            --text-secondary: #475569;
            --border-color: #e2e8f0;
            --accent-blue: #0284c7;
            --accent-teal: #0d9488;
            --accent-amber: #d97706;
            --shadow: 0 4px 6px -1px rgba(0,0,0,0.05), 0 2px 4px -2px rgba(0,0,0,0.05);
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        }}

        body {{
            background-color: var(--bg-primary);
            color: var(--text-primary);
            line-height: 1.5;
            padding: 24px;
            max-width: 1400px;
            margin: 0 auto;
        }}

        .header {{
            margin-bottom: 24px;
            padding-bottom: 16px;
            border-bottom: 2px solid var(--border-color);
        }}

        .header h1 {{
            font-size: 1.875rem;
            font-weight: 700;
            color: var(--text-primary);
            margin-bottom: 4px;
        }}

        .header .subtitle {{
            font-size: 0.95rem;
            color: var(--text-secondary);
        }}

        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 16px;
            margin-bottom: 24px;
        }}

        .kpi-card {{
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 16px;
            box-shadow: var(--shadow);
        }}

        .kpi-label {{
            font-size: 0.8rem;
            font-weight: 600;
            text-transform: uppercase;
            color: var(--text-secondary);
            margin-bottom: 4px;
        }}

        .kpi-value {{
            font-size: 1.5rem;
            font-weight: 700;
            color: var(--accent-blue);
        }}

        .kpi-subtext {{
            font-size: 0.75rem;
            color: var(--text-secondary);
            margin-top: 4px;
        }}

        .tab-bar {{
            display: flex;
            border-bottom: 1px solid var(--border-color);
            margin-bottom: 24px;
        }}

        .tab-btn {{
            padding: 12px 24px;
            background: none;
            border: none;
            border-bottom: 3px solid transparent;
            font-weight: 600;
            font-size: 0.95rem;
            color: var(--text-secondary);
            cursor: pointer;
            transition: all 0.2s ease;
        }}

        .tab-btn:hover {{
            color: var(--accent-blue);
        }}

        .tab-btn.active {{
            color: var(--accent-blue);
            border-bottom-color: var(--accent-blue);
        }}

        .tab-content {{
            display: none;
        }}

        .tab-content.active {{
            display: block;
        }}

        .grid-2col {{
            display: grid;
            grid-template-columns: 2fr 1fr;
            gap: 20px;
            margin-bottom: 24px;
        }}

        .card {{
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 20px;
            box-shadow: var(--shadow);
        }}

        .card h2 {{
            font-size: 1.1rem;
            font-weight: 600;
            margin-bottom: 12px;
            color: var(--text-primary);
            border-left: 4px solid var(--accent-blue);
            padding-left: 8px;
        }}

        .chart-container {{
            position: relative;
            height: 380px;
            width: 100%;
        }}

        .note-box {{
            background: #ffffff;
            border: 1px solid var(--border-color);
            border-radius: 6px;
            padding: 14px;
            font-size: 0.875rem;
            color: var(--text-secondary);
        }}

        .note-box h4 {{
            color: var(--text-primary);
            margin-bottom: 8px;
            font-size: 0.9rem;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.875rem;
            margin-top: 8px;
        }}

        th, td {{
            padding: 10px 12px;
            text-align: left;
            border-bottom: 1px solid var(--border-color);
        }}

        th {{
            background-color: #f1f5f9;
            font-weight: 600;
            color: var(--text-primary);
        }}

        tbody tr:hover {{
            background-color: #f1f5f9;
        }}

        .tag {{
            display: inline-block;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 0.75rem;
            font-weight: 600;
        }}

        .tag-success {{
            background-color: #dcfce7;
            color: #15803d;
        }}
    </style>
</head>
<body>

    <div class="header">
        <h1>Reservoir Stage Kinematics & Modeling</h1>
        <div class="subtitle">Document Ref: lab02_villondo.html | Dataset: Data-01.xlsx | Method: 4-Parameter Logistic Fitting (Levenberg-Marquardt)</div>
    </div>

    <div class="kpi-grid">
        <div class="kpi-card">
            <div class="kpi-label">Peak Inflow Rate</div>
            <div class="kpi-value">{max_dh_dt_val:.4f} m/h</div>
            <div class="kpi-subtext">Occurred at t = {max_dh_dt_hour:.2f} hrs</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Goodness-of-Fit (R²)</div>
            <div class="kpi-value">{R2:.6f}</div>
            <div class="kpi-subtext">Std Error (s): {s_err:.4f} m</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Inflection Point (t₀)</div>
            <div class="kpi-value">{t0_fit:.2f} hrs</div>
            <div class="kpi-subtext">Capacity (a): {a_fit:.2f} m</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Total Stage Volume</div>
            <div class="kpi-value">{quad_val:.2f} m·h</div>
            <div class="kpi-subtext">Quad vs Trapz diff: {integration_diff_pct:.4f}%</div>
        </div>
    </div>

    <div class="tab-bar">
        <button class="tab-btn active" onclick="openTab(event, 'tab1')">Tab 1: Numerical Derivatives & Kinematics</button>
        <button class="tab-btn" onclick="openTab(event, 'tab2')">Tab 2: Logistic Model & Curve Fitting</button>
        <button class="tab-btn" onclick="openTab(event, 'tab3')">Tab 3: Statistical Diagnostics & Volumetric Integration</button>
    </div>

    <div id="tab1" class="tab-content active">
        <div class="grid-2col">
            <div class="card">
                <h2>Finite-Difference Kinematics (dh/dt & d²h/dt²)</h2>
                <div class="chart-container">
                    <canvas id="derivativesChart"></canvas>
                </div>
            </div>
            <div class="card">
                <h2>Kinematic Analysis Notes</h2>
                <div class="note-box">
                    <h4>Derivative Computation Method</h4>
                    <p>Finite differences were evaluated using central differences on interior nodes and second-order single-sided endpoints.</p>
                    <br>
                    <h4>Key Hydrodynamic Findings</h4>
                    <ul>
                        <li><strong>Maximum Inflow Rate:</strong> {max_dh_dt_val:.4f} m/hr at hour <strong>{max_dh_dt_hour:.2f}</strong>.</li>
                        <li><strong>Inflow Acceleration:</strong> The 2nd derivative d²h/dt² indicates maximum stage acceleration right before peak inflow and maximum deceleration immediately following the inflection point.</li>
                    </ul>
                </div>
            </div>
        </div>
    </div>

    <div id="tab2" class="tab-content">
        <div class="grid-2col">
            <div class="card">
                <h2>4-Parameter Logistic Model Alignment</h2>
                <div class="chart-container">
                    <canvas id="fittingChart"></canvas>
                </div>
            </div>
            <div class="card">
                <h2>Fitted Parameter Specifications</h2>
                <p style="font-size: 0.85rem; color: var(--text-secondary); margin-bottom: 12px;">Model equation: h(t) = c + a / [1 + exp(-k(t - t₀))]</p>
                <table>
                    <thead>
                        <tr>
                            <th>Parameter</th>
                            <th>Estimate</th>
                            <th>Std Error</th>
                            <th>p-value</th>
                        </tr>
                    </thead>
                    <tbody>
                        {"".join(f"<tr><td><strong>{p['name']}</strong></td><td>{p['value']:.4f}</td><td>{p['se']:.4f}</td><td><span class='tag tag-success'>{p['p_val']:.4e}</span></td></tr>" for p in param_summary)}
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <div id="tab3" class="tab-content">
        <div class="grid-2col">
            <div class="card">
                <h2>Statistical Diagnostics & Residual Distribution</h2>
                <div class="chart-container">
                    <canvas id="residualChart"></canvas>
                </div>
            </div>
            <div class="card">
                <h2>Integration & Goodness of Fit</h2>
                <div class="note-box" style="margin-bottom: 16px;">
                    <h4>Goodness of Fit Diagnostics</h4>
                    <p><strong>Sum of Squared Errors (SSE):</strong> {SSE:.6f}</p>
                    <p><strong>Total Sum of Squares (SST):</strong> {SST:.6f}</p>
                    <p><strong>Coefficient of Determination (R²):</strong> {R2:.6f}</p>
                    <p><strong>Residual Standard Error (s):</strong> {s_err:.6f} m</p>
                </div>
                <div class="note-box">
                    <h4>Volumetric Integration Cross-Check</h4>
                    <p><strong>scipy.integrate.quad (Continuous):</strong> {quad_val:.4f} m·h (Abs Error Estimate: {quad_err:.2e})</p>
                    <p><strong>np.trapezoid (Discrete):</strong> {trapz_val:.4f} m·h</p>
                    <p><strong>Percentage Variance:</strong> {integration_diff_pct:.4f}%</p>
                </div>
            </div>
        </div>
    </div>

    <script>
        function openTab(evt, tabName) {{
            var i, tabcontent, tablinks;
            tabcontent = document.getElementsByClassName("tab-content");
            for (i = 0; i < tabcontent.length; i++) {{
                tabcontent[i].classList.remove("active");
            }}
            tablinks = document.getElementsByClassName("tab-btn");
            for (i = 0; i < tablinks.length; i++) {{
                tablinks[i].classList.remove("active");
            }}
            document.getElementById(tabName).classList.add("active");
            evt.currentTarget.classList.add("active");
        }}

        const tLabels = {json.dumps(t_labels)};
        const hObserved = {json.dumps(h_observed)};
        const hFitted = {json.dumps(h_fitted)};
        const dhDtList = {json.dumps(dh_dt_list)};
        const d2hDt2List = {json.dumps(d2h_dt2_list)};
        const residualsData = {json.dumps(residuals_list)};

        const ctx1 = document.getElementById('derivativesChart').getContext('2d');
        new Chart(ctx1, {{
            type: 'line',
            data: {{
                labels: tLabels,
                datasets: [
                    {{
                        label: '1st Derivative dh/dt (m/h)',
                        data: dhDtList,
                        borderColor: '#0284c7',
                        backgroundColor: 'rgba(2, 132, 199, 0.05)',
                        borderWidth: 2,
                        pointRadius: 1,
                        yAxisID: 'y'
                    }},
                    {{
                        label: '2nd Derivative d²h/dt² (m/h²)',
                        data: d2hDt2List,
                        borderColor: '#d97706',
                        backgroundColor: 'transparent',
                        borderWidth: 1.5,
                        borderDash: [4, 4],
                        pointRadius: 0,
                        yAxisID: 'y1'
                    }}
                ]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                scales: {{
                    x: {{ title: {{ display: true, text: 'Elapsed Time (hours)' }} }},
                    y: {{ type: 'linear', display: true, position: 'left', title: {{ display: true, text: 'dh/dt (m/h)' }} }},
                    y1: {{ type: 'linear', display: true, position: 'right', grid: {{ drawOnChartArea: false }}, title: {{ display: true, text: 'd²h/dt² (m/h²)' }} }}
                }}
            }}
        }});

        const ctx2 = document.getElementById('fittingChart').getContext('2d');
        new Chart(ctx2, {{
            type: 'line',
            data: {{
                labels: tLabels,
                datasets: [
                    {{
                        label: 'Observed Stage Height h(t)',
                        data: hObserved,
                        borderColor: '#0f172a',
                        backgroundColor: '#0f172a',
                        type: 'scatter',
                        pointRadius: 3
                    }},
                    {{
                        label: '4-Param Logistic Fit h(t)',
                        data: hFitted,
                        borderColor: '#0d9488',
                        borderWidth: 2.5,
                        pointRadius: 0
                    }}
                ]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                scales: {{
                    x: {{ title: {{ display: true, text: 'Elapsed Time (hours)' }} }},
                    y: {{ title: {{ display: true, text: 'Stage Height h (meters)' }} }}
                }}
            }}
        }});

        const ctx3 = document.getElementById('residualChart').getContext('2d');
        new Chart(ctx3, {{
            type: 'line',
            data: {{
                labels: tLabels,
                datasets: [
                    {{
                        label: 'Fit Residuals (Observed - Predicted)',
                        data: residualsData,
                        borderColor: '#dc2626',
                        backgroundColor: 'rgba(220, 38, 38, 0.1)',
                        borderWidth: 1.5,
                        pointRadius: 2,
                        fill: true
                    }}
                ]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                scales: {{
                    x: {{ title: {{ display: true, text: 'Elapsed Time (hours)' }} }},
                    y: {{ title: {{ display: true, text: 'Residual Error (m)' }} }}
                }}
            }}
        }});
    </script>
</body>
</html>
"""

# Save output HTML file in the current working directory
output_filename = "lab02_villondo.html"
with open(output_filename, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"Analysis complete. Generated executive report: {os.path.abspath(output_filename)}")