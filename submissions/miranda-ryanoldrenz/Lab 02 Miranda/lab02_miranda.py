import json
import os
import re
import numpy as np
import pandas as pd
from scipy.integrate import quad
from scipy.optimize import curve_fit
from scipy.stats import t

# ==========================================
# 1. FILE PATHS & DATA INGESTION
# ==========================================
FILE_PATH = r"C:\Users\Lenovo\Downloads\Data 01.xlsx"
OUTPUT_HTML = "lab02_miranda.html"

# Load dataset starting at row index 3 (0-indexed, corresponding to row 4 in Excel)
df = pd.read_excel(FILE_PATH, header=3)
df.columns = [str(col).strip() for col in df.columns]

# Locate relevant columns for Time and Stage height h(t)
time_col = [col for col in df.columns if "time" in col.lower() or "date" in col.lower() or "t" in col.lower()][0]
stage_col = [col for col in df.columns if "stage" in col.lower() or "height" in col.lower() or "h" in col.lower() or "level" in col.lower()][0]

df = df.dropna(subset=[time_col, stage_col]).reset_index(drop=True)

# Parse time and calculate elapsed hours t
df[time_col] = pd.to_datetime(df[time_col])
t_start = df[time_col].iloc[0]
t_hours = (df[time_col] - t_start).dt.total_seconds() / 3600.0

t_arr = t_hours.to_numpy(dtype=float)
h_arr = df[stage_col].to_numpy(dtype=float)
n_pts = len(t_arr)

# ==========================================
# 2. NUMERICAL DERIVATIVES (FINITE DIFFERENCE)
# ==========================================
# Central differences for interior points, forward/backward for boundaries
dh_dt = np.gradient(h_arr, t_arr)
d2h_dt2 = np.gradient(dh_dt, t_arr)

# Maximum inflow rate (max 1st derivative)
max_inflow_idx = np.argmax(dh_dt)
max_inflow_rate = float(dh_dt[max_inflow_idx])
max_inflow_hour = float(t_arr[max_inflow_idx])

# Helper function to sanitize float arrays for valid JSON formatting (NaN -> None)
def sanitize_array(arr):
    return [None if np.isnan(val) or np.isinf(val) else round(float(val), 6) for val in arr]

t_list = sanitize_array(t_arr)
h_list = sanitize_array(h_arr)
dh_list = sanitize_array(dh_dt)
d2h_list = sanitize_array(d2h_dt2)

# ==========================================
# 3. LOGISTIC CURVE FITTING
# ==========================================
# 4-parameter logistic function: h(t) = c + a / (1 + exp(-k * (t - t0)))
def logistic_4p(t_val, a, k, t0, c):
    return c + a / (1.0 + np.exp(-k * (t_val - t0)))

# Initial parameter guesses based on data bounds
c_guess = float(np.min(h_arr))
a_guess = float(np.max(h_arr) - np.min(h_arr))
t0_guess = float(np.median(t_arr))
k_guess = 0.5

p0 = [a_guess, k_guess, t0_guess, c_guess]

# Fit model using Levenberg-Marquardt algorithm
popt, pcov = curve_fit(logistic_4p, t_arr, h_arr, p0=p0, method='lm', maxfev=10000)
a_fit, k_fit, t0_fit, c_fit = popt

h_pred = logistic_4p(t_arr, *popt)
h_pred_list = sanitize_array(h_pred)

# Analytical derivative functions from fitted logistic parameters
def logistic_dh_dt(t_val):
    exp_term = np.exp(-k_fit * (t_val - t0_fit))
    return (a_fit * k_fit * exp_term) / ((1.0 + exp_term) ** 2)

def logistic_d2h_dt2(t_val):
    exp_term = np.exp(-k_fit * (t_val - t0_fit))
    return (a_fit * (k_fit**2) * exp_term * (exp_term - 1.0)) / ((1.0 + exp_term) ** 3)

dh_model_list = sanitize_array(logistic_dh_dt(t_arr))
d2h_model_list = sanitize_array(logistic_d2h_dt2(t_arr))

# ==========================================
# 4. STATISTICAL DIAGNOSTICS
# ==========================================
residuals = h_arr - h_pred
SSE = float(np.sum(residuals**2))
SST = float(np.sum((h_arr - np.mean(h_arr))**2))
R2 = float(1.0 - (SSE / SST))

p_num = len(popt)  # 4 parameters
dof = n_pts - p_num
s_err = float(np.sqrt(SSE / dof)) if dof > 0 else 0.0

# Standard errors, t-statistics, and two-tailed p-values for parameters
param_names = ["Scale Parameter (a)", "Growth Rate (k)", "Inflection Point (t₀)", "Baseline Offset (c)"]
param_units = ["m", "hr⁻¹", "hr", "m"]

perr = np.sqrt(np.diag(pcov))
t_stats = popt / perr
p_values = 2.0 * (1.0 - t.cdf(np.abs(t_stats), df=dof))

param_stats = []
for idx in range(p_num):
    param_stats.append({
        "name": param_names[idx],
        "value": round(float(popt[idx]), 6),
        "unit": param_units[idx],
        "stderr": round(float(perr[idx]), 6),
        "t_stat": round(float(t_stats[idx]), 4),
        "p_val": "< 0.0001" if p_values[idx] < 0.0001 else f"{p_values[idx]:.4f}",
        "status": "Significant" if p_values[idx] < 0.05 else "Non-Significant"
    })

# ==========================================
# 5. VOLUMETRIC INTEGRATION
# ==========================================
t_min, t_max = float(t_arr[0]), float(t_arr[-1])

# Integration via scipy.integrate.quad
quad_val, quad_err = quad(logistic_4p, t_min, t_max, args=(a_fit, k_fit, t0_fit, c_fit))

# Integration cross-check via np.trapezoid (or np.trapz for compatibility)
if hasattr(np, 'trapezoid'):
    trapz_val = float(np.trapezoid(h_pred, t_arr))
    trapz_raw = float(np.trapezoid(h_arr, t_arr))
else:
    trapz_val = float(np.trapz(h_pred, t_arr))
    trapz_raw = float(np.trapz(h_arr, t_arr))

abs_diff = abs(quad_val - trapz_val)
rel_diff_pct = (abs_diff / quad_val) * 100.0

# Generate dense smooth curve for visualization (200 points)
t_dense = np.linspace(t_min, t_max, 200)
h_dense = logistic_4p(t_dense, *popt)
t_dense_list = sanitize_array(t_dense)
h_dense_list = sanitize_array(h_dense)

# Dynamic text metrics for injection
kpi_total_duration = round(t_max - t_min, 2)
kpi_initial_stage = round(float(h_arr[0]), 3)
kpi_final_stage = round(float(h_arr[-1]), 3)

# ==========================================
# 6. HTML REPORT GENERATION
# ==========================================
html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Lab 02: Reservoir Stage Kinematics & Logistic Modeling</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Chart.js CDN -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body {{
            background-color: #ffffff;
            color: #0f172a;
            font-family: 'Inter', system-ui, -apple-system, sans-serif;
        }}
        .tab-btn.active {{
            border-bottom-color: #0284c7;
            color: #0284c7;
            font-weight: 600;
        }}
    </style>
</head>
<body class="bg-white min-h-screen pb-12">

    <!-- Top Executive Header Block -->
    <header class="bg-slate-50 border-b border-slate-200 py-6 px-8 mb-8">
        <div class="max-w-7xl mx-auto flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
            <div>
                <span class="inline-block px-3 py-1 text-xs font-semibold uppercase tracking-wider text-sky-700 bg-sky-100 rounded-full mb-2">
                    Executive Engineering Report
                </span>
                <h1 class="text-2xl font-bold text-slate-900">Lab 02: Reservoir Stage Dataset Analysis</h1>
                <p class="text-sm text-slate-500 mt-1">
                    Student: <span class="font-medium text-slate-700">Miranda</span> &nbsp;|&nbsp; 
                    Dataset: <span class="font-medium text-slate-700">Data-01.xlsx</span> &nbsp;|&nbsp; 
                    Algorithm: <span class="font-medium text-slate-700">Levenberg-Marquardt (scipy.optimize)</span>
                </p>
            </div>
            <div class="text-right">
                <span class="text-xs text-slate-400 block">Report Generation</span>
                <span class="text-sm font-semibold text-slate-700">Automated Pipeline</span>
            </div>
        </div>
    </header>

    <main class="max-w-7xl mx-auto px-6">

        <!-- 4-Column Highlight KPIs Box -->
        <section class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
            <div class="bg-slate-50 border border-slate-200 rounded-xl p-5 shadow-sm">
                <span class="text-xs font-semibold text-slate-500 uppercase tracking-wide">Total Duration</span>
                <div class="text-2xl font-bold text-slate-900 mt-1">{kpi_total_duration} <span class="text-base font-normal text-slate-600">hrs</span></div>
                <p class="text-xs text-slate-400 mt-1">Start: {t_min:.2f}h | End: {t_max:.2f}h</p>
            </div>

            <div class="bg-slate-50 border border-slate-200 rounded-xl p-5 shadow-sm">
                <span class="text-xs font-semibold text-slate-500 uppercase tracking-wide">Max Inflow Rate (dh/dt)</span>
                <div class="text-2xl font-bold text-sky-600 mt-1">{max_inflow_rate:.4f} <span class="text-base font-normal text-slate-600">m/hr</span></div>
                <p class="text-xs text-slate-400 mt-1">Observed at hour <span class="font-medium text-slate-700">{max_inflow_hour:.2f} h</span></p>
            </div>

            <div class="bg-slate-50 border border-slate-200 rounded-xl p-5 shadow-sm">
                <span class="text-xs font-semibold text-slate-500 uppercase tracking-wide">Goodness of Fit (R²)</span>
                <div class="text-2xl font-bold text-teal-600 mt-1">{R2:.6f}</div>
                <p class="text-xs text-slate-400 mt-1">Std Error (s): {s_err:.4f} m</p>
            </div>

            <div class="bg-slate-50 border border-slate-200 rounded-xl p-5 shadow-sm">
                <span class="text-xs font-semibold text-slate-500 uppercase tracking-wide">Cumulative Integration</span>
                <div class="text-2xl font-bold text-amber-600 mt-1">{quad_val:.3f} <span class="text-base font-normal text-slate-600">m·hr</span></div>
                <p class="text-xs text-slate-400 mt-1">scipy.quad (Diff: {rel_diff_pct:.4f}%)</p>
            </div>
        </section>

        <!-- Interactive Tab Navigation Bar -->
        <div class="border-b border-slate-200 mb-6">
            <nav class="flex gap-8" id="tabNav">
                <button onclick="switchTab('tab1')" id="btn-tab1" class="tab-btn active pb-3 text-sm font-medium text-slate-500 border-b-2 border-transparent hover:text-slate-700 transition-colors">
                    Tab 1: Numerical Derivatives & Kinematics
                </button>
                <button onclick="switchTab('tab2')" id="btn-tab2" class="tab-btn pb-3 text-sm font-medium text-slate-500 border-b-2 border-transparent hover:text-slate-700 transition-colors">
                    Tab 2: Logistic Model & Curve Fitting
                </button>
                <button onclick="switchTab('tab3')" id="btn-tab3" class="tab-btn pb-3 text-sm font-medium text-slate-500 border-b-2 border-transparent hover:text-slate-700 transition-colors">
                    Tab 3: Statistical Diagnostics & Volumetric Integration
                </button>
            </nav>
        </div>

        <!-- ================= TAB 1 CONTENT ================= -->
        <div id="tab1" class="tab-content">
            <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <!-- Left 2 Cols: Charts -->
                <div class="lg:col-span-2 space-y-6">
                    <div class="bg-slate-50 border border-slate-200 rounded-xl p-5 shadow-sm">
                        <h2 class="text-md font-semibold text-slate-800 mb-1">First Derivative: Velocity (dh/dt)</h2>
                        <p class="text-xs text-slate-500 mb-4">Calculated via finite-difference numerical scheme.</p>
                        <div class="h-64">
                            <canvas id="chart1stDeriv"></canvas>
                        </div>
                    </div>

                    <div class="bg-slate-50 border border-slate-200 rounded-xl p-5 shadow-sm">
                        <h2 class="text-md font-semibold text-slate-800 mb-1">Second Derivative: Acceleration (d²h/dt²)</h2>
                        <p class="text-xs text-slate-500 mb-4">Reflects inflection transitions in reservoir stage kinetics.</p>
                        <div class="h-64">
                            <canvas id="chart2ndDeriv"></canvas>
                        </div>
                    </div>
                </div>

                <!-- Right Col: Engineering Notes & Summary -->
                <div class="space-y-6">
                    <div class="bg-slate-50 border border-slate-200 rounded-xl p-5 shadow-sm">
                        <h3 class="text-md font-semibold text-slate-900 mb-3 border-b border-slate-200 pb-2">Kinematic Insights</h3>
                        <div class="space-y-4 text-sm text-slate-600">
                            <div>
                                <span class="text-xs font-semibold uppercase text-slate-400 block">Peak Inflow Peak</span>
                                <p class="text-slate-800 font-medium">Rate: <span class="text-sky-600 font-bold">{max_inflow_rate:.4f} m/hr</span></p>
                                <p class="text-slate-800 font-medium">Timestamp: <span class="text-slate-900 font-bold">{max_inflow_hour:.2f} hours</span></p>
                            </div>
                            <hr class="border-slate-200">
                            <div>
                                <span class="text-xs font-semibold uppercase text-slate-400 block">Engineering Context</span>
                                <p class="mt-1 leading-relaxed text-xs text-slate-600">
                                    The first derivative represents the instantaneous rate of stage change (inflow velocity). The second derivative indicates the rate of change of inflow velocity, highlighting the point where stage acceleration transitions to deceleration.
                                </p>
                            </div>
                            <div class="bg-amber-50 border-l-4 border-amber-500 p-3 text-xs text-amber-800 rounded">
                                <strong>Inflection Point:</strong> Peak inflow corresponds precisely to zero crossing in the second derivative ($d^2h/dt^2 = 0$).
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- ================= TAB 2 CONTENT ================= -->
        <div id="tab2" class="tab-content hidden">
            <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <!-- Left 2 Cols: Main Fitting Chart -->
                <div class="lg:col-span-2">
                    <div class="bg-slate-50 border border-slate-200 rounded-xl p-5 shadow-sm">
                        <div class="flex justify-between items-center mb-2">
                            <div>
                                <h2 class="text-md font-semibold text-slate-800">Logistic Model vs. Measured Stage</h2>
                                <p class="text-xs text-slate-500">Nonlinear regression using 4-parameter logistic curve.</p>
                            </div>
                            <span class="text-xs font-mono bg-slate-200 text-slate-700 px-2 py-1 rounded">R² = {R2:.6f}</span>
                        </div>
                        <div class="h-[460px]">
                            <canvas id="chartLogisticFit"></canvas>
                        </div>
                    </div>
                </div>

                <!-- Right Col: Parameter Table -->
                <div>
                    <div class="bg-slate-50 border border-slate-200 rounded-xl p-5 shadow-sm">
                        <h3 class="text-md font-semibold text-slate-900 mb-3 border-b border-slate-200 pb-2">Fitted Parameters</h3>
                        <p class="text-xs text-slate-500 mb-4">Form: $h(t) = c + \frac{{a}}{{1 + e^{{-k(t - t_0)}}}}$</p>
                        
                        <div class="overflow-x-auto">
                            <table class="w-full text-left text-xs text-slate-600">
                                <thead class="bg-slate-100 text-slate-700 font-semibold border-b border-slate-200">
                                    <tr>
                                        <th class="p-2">Parameter</th>
                                        <th class="p-2">Estimate</th>
                                        <th class="p-2">Std Error</th>
                                    </tr>
                                </thead>
                                <tbody class="divide-y divide-slate-200">
                                    {"".join([f'''
                                    <tr class="hover:bg-slate-100/50 transition-colors">
                                        <td class="p-2 font-medium text-slate-800">{p["name"]}</td>
                                        <td class="p-2 font-mono text-sky-700 font-semibold">{p["value"]} {p["unit"]}</td>
                                        <td class="p-2 font-mono text-slate-500">±{p["stderr"]}</td>
                                    </tr>
                                    ''' for p in param_stats])}
                                </tbody>
                            </table>
                        </div>

                        <div class="mt-4 p-3 bg-sky-50 border border-sky-100 rounded-lg text-xs text-sky-800">
                            <strong>Physical Meaning:</strong>
                            <ul class="list-disc list-inside mt-1 space-y-1 text-slate-600">
                                <li><strong>a:</strong> Total stage rise capacity ({a_fit:.3f} m)</li>
                                <li><strong>k:</strong> Steepness / growth constant ({k_fit:.3f} hr⁻¹)</li>
                                <li><strong>t₀:</strong> Midpoint time ({t0_fit:.3f} hr)</li>
                                <li><strong>c:</strong> Baseline stage elevation ({c_fit:.3f} m)</li>
                            </ul>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- ================= TAB 3 CONTENT ================= -->
        <div id="tab3" class="tab-content hidden">
            <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                <!-- Left Card: Full Statistical Parameter Diagnostics -->
                <div class="bg-slate-50 border border-slate-200 rounded-xl p-5 shadow-sm">
                    <h3 class="text-md font-semibold text-slate-900 mb-3 border-b border-slate-200 pb-2">Statistical Significance & Diagnostics</h3>
                    <div class="overflow-x-auto">
                        <table class="w-full text-left text-xs text-slate-600">
                            <thead class="bg-slate-100 text-slate-700 font-semibold border-b border-slate-200">
                                <tr>
                                    <th class="p-2">Parameter</th>
                                    <th class="p-2">t-Statistic</th>
                                    <th class="p-2">p-Value</th>
                                    <th class="p-2">Status</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-slate-200">
                                {"".join([f'''
                                <tr class="hover:bg-slate-100/50 transition-colors">
                                    <td class="p-2 font-medium text-slate-800">{p["name"]}</td>
                                    <td class="p-2 font-mono">{p["t_stat"]}</td>
                                    <td class="p-2 font-mono">{p["p_val"]}</td>
                                    <td class="p-2">
                                        <span class="px-2 py-0.5 text-[10px] font-semibold rounded-full bg-emerald-100 text-emerald-700">
                                            {p["status"]}
                                        </span>
                                    </td>
                                </tr>
                                ''' for p in param_stats])}
                            </tbody>
                        </table>
                    </div>

                    <div class="mt-6 space-y-2 text-xs text-slate-600">
                        <div class="flex justify-between py-1 border-b border-slate-200">
                            <span>Sum of Squared Errors (SSE):</span>
                            <span class="font-mono font-semibold text-slate-800">{SSE:.6f}</span>
                        </div>
                        <div class="flex justify-between py-1 border-b border-slate-200">
                            <span>Total Sum of Squares (SST):</span>
                            <span class="font-mono font-semibold text-slate-800">{SST:.6f}</span>
                        </div>
                        <div class="flex justify-between py-1 border-b border-slate-200">
                            <span>Standard Error of Estimate (s):</span>
                            <span class="font-mono font-semibold text-slate-800">{s_err:.6f} m</span>
                        </div>
                        <div class="flex justify-between py-1">
                            <span>Degrees of Freedom (dof):</span>
                            <span class="font-mono font-semibold text-slate-800">{dof}</span>
                        </div>
                    </div>
                </div>

                <!-- Right Card: Volumetric Integration & Cross-Check -->
                <div class="bg-slate-50 border border-slate-200 rounded-xl p-5 shadow-sm">
                    <h3 class="text-md font-semibold text-slate-900 mb-3 border-b border-slate-200 pb-2">Volumetric Integration Comparison</h3>
                    <p class="text-xs text-slate-500 mb-4">Evaluation of cumulative stage-time area under curve $\int_{{t_0}}^{{t_f}} h(t) dt$.</p>
                    
                    <div class="space-y-4">
                        <div class="p-4 bg-white border border-slate-200 rounded-lg flex justify-between items-center">
                            <div>
                                <span class="text-xs text-slate-400 font-semibold block uppercase">Adaptive Quadrature (scipy.integrate.quad)</span>
                                <span class="text-lg font-bold text-sky-600">{quad_val:.6f} <span class="text-xs text-slate-500 font-normal">m·hr</span></span>
                            </div>
                            <span class="text-xs bg-sky-100 text-sky-800 px-2 py-1 rounded">Model Integral</span>
                        </div>

                        <div class="p-4 bg-white border border-slate-200 rounded-lg flex justify-between items-center">
                            <div>
                                <span class="text-xs text-slate-400 font-semibold block uppercase">Trapezoidal Method (np.trapezoid)</span>
                                <span class="text-lg font-bold text-teal-600">{trapz_val:.6f} <span class="text-xs text-slate-500 font-normal">m·hr</span></span>
                            </div>
                            <span class="text-xs bg-teal-100 text-teal-800 px-2 py-1 rounded">Discrete Check</span>
                        </div>

                        <div class="p-4 bg-slate-100 border border-slate-200 rounded-lg">
                            <div class="flex justify-between text-xs font-semibold text-slate-700 mb-1">
                                <span>Absolute Difference:</span>
                                <span class="font-mono">{abs_diff:.6e} m·hr</span>
                            </div>
                            <div class="flex justify-between text-xs font-semibold text-slate-700">
                                <span>Relative Error:</span>
                                <span class="font-mono text-emerald-600">{rel_diff_pct:.6f}%</span>
                            </div>
                        </div>

                        <div class="text-xs text-slate-500 leading-relaxed bg-white p-3 border border-slate-200 rounded-lg">
                            <strong>Note on Accuracy:</strong> The high degree of agreement ({rel_diff_pct:.4f}%) between continuous adaptive Gaussian quadrature and discrete trapezoidal integration verifies numerical stability across the analysis domain.
                        </div>
                    </div>
                </div>
            </div>
        </div>

    </main>

    <!-- Embedded Data & Chart Script -->
    <script>
        // Data arrays exported cleanly from Python
        const tData = {json.dumps(t_list)};
        const hData = {json.dumps(h_list)};
        const dhData = {json.dumps(dh_list)};
        const d2hData = {json.dumps(d2h_list)};
        const hPredData = {json.dumps(h_pred_list)};
        const tDense = {json.dumps(t_dense_list)};
        const hDense = {json.dumps(h_dense_list)};

        // Chart instances
        let c1, c2, c3;

        // Common Chart Options
        const commonOptions = {{
            responsive: true,
            maintainAspectRatio: false,
            plugins: {{
                legend: {{ position: 'top', labels: {{ font: {{ size: 11 }} }} }}
            }},
            scales: {{
                x: {{
                    title: {{ display: true, text: 'Elapsed Time (hours)', font: {{ size: 11 }} }},
                    grid: {{ color: '#f1f5f9' }}
                }},
                y: {{
                    grid: {{ color: '#f1f5f9' }}
                }}
            }}
        }};

        // Initialize Charts
        window.addEventListener('DOMContentLoaded', () => {{
            // 1st Derivative Chart
            const ctx1 = document.getElementById('chart1stDeriv').getContext('2d');
            c1 = new Chart(ctx1, {{
                type: 'line',
                data: {{
                    labels: tData,
                    datasets: [{{
                        label: 'dh/dt (m/hr) Finite Diff',
                        data: dhData,
                        borderColor: '#0284c7',
                        backgroundColor: 'rgba(2, 132, 199, 0.08)',
                        fill: true,
                        tension: 0.3,
                        pointRadius: 2
                    }}]
                }},
                options: commonOptions
            }});

            // 2nd Derivative Chart
            const ctx2 = document.getElementById('chart2ndDeriv').getContext('2d');
            c2 = new Chart(ctx2, {{
                type: 'line',
                data: {{
                    labels: tData,
                    datasets: [{{
                        label: 'd²h/dt² (m/hr²) Finite Diff',
                        data: d2hData,
                        borderColor: '#d97706',
                        backgroundColor: 'rgba(217, 119, 6, 0.08)',
                        fill: true,
                        tension: 0.3,
                        pointRadius: 2
                    }}]
                }},
                options: commonOptions
            }});

            // Logistic Model Fit Chart
            const ctx3 = document.getElementById('chartLogisticFit').getContext('2d');
            c3 = new Chart(ctx3, {{
                type: 'scatter',
                data: {{
                    datasets: [
                        {{
                            label: 'Measured Stage h(t)',
                            data: tData.map((t, i) => ({{ x: t, y: hData[i] }})),
                            backgroundColor: '#0f172a',
                            pointRadius: 4
                        }},
                        {{
                            type: 'line',
                            label: 'Logistic Model h(t)',
                            data: tDense.map((t, i) => ({{ x: t, y: hDense[i] }})),
                            borderColor: '#0d9488',
                            borderWidth: 2,
                            pointRadius: 0,
                            fill: false
                        }}
                    ]
                }},
                options: {{
                    ...commonOptions,
                    scales: {{
                        x: {{ title: {{ display: true, text: 'Elapsed Time (hours)' }} }},
                        y: {{ title: {{ display: true, text: 'Stage Height h (meters)' }} }}
                    }}
                }}
            }});
        }});

        // Tab Switching Logic
        function switchTab(tabId) {{
            document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
            document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));

            document.getElementById(tabId).classList.remove('hidden');
            document.getElementById('btn-' + tabId).classList.add('active');
        }}
    </script>
</body>
</html>
"""

# Save output to HTML file
with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"Analysis complete. Dashboard exported successfully to '{OUTPUT_HTML}'.")