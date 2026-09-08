"""Cubic water-level analysis and self-contained HTML dashboard.

Usage (PowerShell):
    python water_level_dashboard.py --input "C:\\Users\\stewa\\Downloads\\Data 01.xlsx"

The reader intentionally uses only the Python standard library for .xlsx files;
SciPy and NumPy are the only analysis dependencies.
"""
from __future__ import annotations

import argparse
import html
import json
import math
import re
from datetime import datetime, timedelta
from pathlib import Path
from statistics import median
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import numpy as np
from scipy import integrate, optimize, stats


NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def excel_column(reference: str) -> int:
    """Return zero-based column number from an Excel reference such as E17."""
    letters = re.match(r"[A-Z]+", reference).group(0)
    n = 0
    for letter in letters:
        n = n * 26 + ord(letter) - 64
    return n - 1


def read_xlsx_table(path: Path) -> tuple[list[str], list[list[object]]]:
    """Read the first worksheet as a sparse table, without pandas/openpyxl."""
    with ZipFile(path) as book:
        shared = []
        if "xl/sharedStrings.xml" in book.namelist():
            root = ET.fromstring(book.read("xl/sharedStrings.xml"))
            shared = ["".join(x.text or "" for x in item.iter(NS + "t")) for item in root]
        worksheet = ET.fromstring(book.read("xl/worksheets/sheet1.xml"))
        rows = []
        for row in worksheet.findall(".//" + NS + "sheetData/" + NS + "row"):
            cells = {}
            for cell in row.findall(NS + "c"):
                value = cell.find(NS + "v")
                if value is None:
                    text = ""
                else:
                    text = value.text or ""
                    if cell.attrib.get("t") == "s":
                        text = shared[int(text)]
                cells[excel_column(cell.attrib["r"])] = text
            if cells:
                rows.append([cells.get(i, "") for i in range(max(cells) + 1)])
    header_i = next(i for i, row in enumerate(rows) if any(str(v).strip().lower() == "depth (m)" for v in row))
    header = [str(x).strip() for x in rows[header_i]]
    return header, rows[header_i + 1 :]


def load_observations(path: Path):
    header, rows = read_xlsx_table(path)
    normalized = [x.lower() for x in header]
    depth_i = normalized.index("depth (m)")
    time_i = normalized.index("timestamp") if "timestamp" in normalized else normalized.index("time")
    timestamps, depth = [], []
    for row in rows:
        if max(depth_i, time_i) >= len(row) or row[depth_i] == "" or row[time_i] == "":
            continue
        timestamps.append(float(row[time_i]))
        depth.append(float(row[depth_i]))
    if len(depth) < 5:
        raise ValueError("Need at least five valid timestamp/depth records.")
    timestamps = np.asarray(timestamps)
    depth = np.asarray(depth)
    order = np.argsort(timestamps)
    timestamps, depth = timestamps[order], depth[order]
    hours = (timestamps - timestamps[0]) * 24.0
    return timestamps, hours, depth


def cubic(t, a, b, c, d):
    return ((a * t + b) * t + c) * t + d


def numerical_derivatives(t, h):
    forward = np.full_like(h, np.nan, dtype=float)
    backward = np.full_like(h, np.nan, dtype=float)
    central = np.full_like(h, np.nan, dtype=float)
    second = np.full_like(h, np.nan, dtype=float)
    dt = np.diff(t)
    forward[:-1] = np.diff(h) / dt
    backward[1:] = np.diff(h) / dt
    central[1:-1] = (h[2:] - h[:-2]) / (t[2:] - t[:-2])
    left, right = dt[:-1], dt[1:]
    second[1:-1] = 2 * ((h[2:] - h[1:-1]) / right - (h[1:-1] - h[:-2]) / left) / (left + right)
    return forward, central, backward, second


def fmt(x, places=5):
    return f"{x:.{places}g}" if math.isfinite(float(x)) else "—"


def points(values, x_min, x_max, y_min, y_max, left=58, top=30, width=682, height=270):
    result = []
    for x, y in values:
        if not (np.isfinite(x) and np.isfinite(y)):
            continue
        sx = left + (x - x_min) / (x_max - x_min) * width if x_max != x_min else left + width / 2
        sy = top + height - (y - y_min) / (y_max - y_min) * height if y_max != y_min else top + height / 2
        result.append(f"{sx:.2f},{sy:.2f}")
    return " ".join(result)


def svg_chart(title, series, x_label, y_label, marker=None, zero=False, area=False):
    """Create a compact inline SVG chart from a list of (x, y, color, label, dash)."""
    all_x = np.concatenate([np.asarray(s[0], float)[np.isfinite(s[0]) & np.isfinite(s[1])] for s in series])
    all_y = np.concatenate([np.asarray(s[1], float)[np.isfinite(s[0]) & np.isfinite(s[1])] for s in series])
    if zero:
        all_y = np.append(all_y, 0.0)
    x0, x1 = float(all_x.min()), float(all_x.max())
    y0, y1 = float(all_y.min()), float(all_y.max())
    dx, dy = x1 - x0, y1 - y0
    x0 -= dx * 0.03 if dx else 1; x1 += dx * 0.03 if dx else 1
    y0 -= dy * 0.12 if dy else 1; y1 += dy * 0.12 if dy else 1
    left, top, w, h = 58, 30, 682, 270
    def sx(x): return left + (x-x0)/(x1-x0)*w
    def sy(y): return top+h-(y-y0)/(y1-y0)*h
    grid = []
    for i in range(6):
        x = x0 + i*(x1-x0)/5; y = y0 + i*(y1-y0)/5
        grid.append(f'<line x1="{sx(x):.1f}" y1="{top}" x2="{sx(x):.1f}" y2="{top+h}" class="grid"/><text x="{sx(x):.1f}" y="{top+h+22}" text-anchor="middle">{fmt(x,4)}</text>')
        grid.append(f'<line x1="{left}" y1="{sy(y):.1f}" x2="{left+w}" y2="{sy(y):.1f}" class="grid"/><text x="{left-8}" y="{sy(y)+4:.1f}" text-anchor="end">{fmt(y,4)}</text>')
    body = []
    if zero:
        body.append(f'<line x1="{left}" y1="{sy(0):.1f}" x2="{left+w}" y2="{sy(0):.1f}" class="zero"/>')
    if area:
        xs, ys = series[0][0], series[0][1]
        poly = points(list(zip(xs, ys)), x0, x1, y0, y1, left, top, w, h)
        fill = f"{left:.2f},{sy(y0):.2f} {poly} {left+w:.2f},{sy(y0):.2f}"
        body.append(f'<polygon points="{fill}" class="area"/>')
    for xs, ys, color, label, dash in series:
        poly = points(list(zip(xs, ys)), x0, x1, y0, y1, left, top, w, h)
        if dash == "scatter":
            for x, y in zip(xs, ys):
                if np.isfinite(x) and np.isfinite(y):
                    body.append(f'<circle cx="{sx(x):.2f}" cy="{sy(y):.2f}" r="2.2" fill="{color}" fill-opacity=".78"/>')
        else:
            body.append(f'<polyline points="{poly}" fill="none" stroke="{color}" stroke-width="2" {"stroke-dasharray=\"6 4\"" if dash else ""}/>')
    if marker:
        mx, my, label = marker
        body.append(f'<circle cx="{sx(mx):.1f}" cy="{sy(my):.1f}" r="5" class="marker"/><text x="{sx(mx)+8:.1f}" y="{sy(my)-10:.1f}" class="marklabel">{html.escape(label)}</text>')
    legend = ''.join(f'<span><i style="background:{c}"></i>{html.escape(lbl)}</span>' for _,_,c,lbl,_ in series)
    return f'''<section class="chart"><h3>{html.escape(title)}</h3><svg viewBox="0 0 800 350" role="img" aria-label="{html.escape(title)}"><style>.grid{{stroke:#dce5ed;stroke-width:1}}text{{fill:#526477;font:11px system-ui}}.zero{{stroke:#e05d5d;stroke-width:1.5;stroke-dasharray:5 4}}.marker{{fill:#f6b93b;stroke:#8c5d00;stroke-width:2}}.marklabel{{font-weight:700;fill:#704700}}.area{{fill:#3478c933}}</style>{''.join(grid)}{''.join(body)}<text x="400" y="342" text-anchor="middle">{html.escape(x_label)}</text><text transform="translate(14 165) rotate(-90)" text-anchor="middle">{html.escape(y_label)}</text></svg><div class="legend">{legend}</div></section>'''


def run_analysis(source: Path, output: Path):
    serial, t, h = load_observations(source)
    p0 = np.polyfit(t, h, 3)
    params, covariance = optimize.curve_fit(cubic, t, h, p0=p0, method="lm", maxfev=20000)
    fitted = cubic(t, *params)
    residual = h - fitted
    n, p = len(h), len(params)
    dof = n - p
    sse = float(np.sum(residual**2)); sst = float(np.sum((h - h.mean())**2))
    r2 = 1 - sse/sst
    see = math.sqrt(sse/dof)
    se = np.sqrt(np.diag(covariance))
    t_stats = params/se
    p_values = 2*stats.t.sf(np.abs(t_stats), dof)
    fd, cd, bd, second = numerical_derivatives(t, h)
    # Candidate extrema of the fitted first derivative inside the observation window.
    a, b, c, d = params
    candidates = [float(t[0]), float(t[-1])]
    if abs(a) > 1e-14:
        vertex = -b/(3*a)
        if t[0] <= vertex <= t[-1]: candidates.append(float(vertex))
    rates = [3*a*x*x + 2*b*x + c for x in candidates]
    max_i = int(np.argmax(rates)); max_t, max_rate = candidates[max_i], rates[max_i]
    dense_t = np.linspace(t[0], t[-1], 700); dense_fit = cubic(dense_t, *params)
    model_d1 = 3*a*dense_t**2 + 2*b*dense_t + c
    model_d2 = 6*a*dense_t + 2*b
    area, quad_error = integrate.quad(lambda x: cubic(x, *params), float(t[0]), float(t[-1]))
    trap = float(np.trapezoid(fitted, t))
    largest = float(np.max(np.abs(residual)))
    max_residual_excess = largest - 0.01
    res_slope = float(np.polyfit(t, residual, 1)[0])
    signs = np.where(residual >= 0, 1, -1)
    longest, current = 1, 1
    for i in range(1, n):
        current = current + 1 if signs[i] == signs[i-1] else 1
        longest = max(longest, current)
    corr_abs_fitted = float(np.corrcoef(np.abs(residual), fitted)[0, 1])
    low, high = np.abs(residual)[fitted <= median(fitted)], np.abs(residual)[fitted > median(fitted)]
    spread_ratio = float(np.std(high, ddof=1) / np.std(low, ddof=1))
    centered = abs(float(residual.mean())) < 1e-10
    gradual = abs(res_slope) > 0.0005  # metres/hour; stated alongside its measured value
    long_runs = longest >= 5
    spreading = corr_abs_fitted > 0.25 and spread_ratio > 1.2
    start = datetime(1899, 12, 30) + timedelta(days=float(serial[0]))
    end = datetime(1899, 12, 30) + timedelta(days=float(serial[-1]))
    equation = f"h(t) = {a:.8g}t³ {b:+.8g}t² {c:+.8g}t {d:+.8g}"
    interpretation = [
        f"Residuals are {'centered' if centered else 'not centered'} around zero (mean {residual.mean():+.3g} m).",
        f"They {'show' if gradual else 'do not show'} a material gradual increase/decrease; residual trend = {res_slope:+.4g} m/h.",
        f"The longest same-sign run is {longest} readings, so {'there are' if long_runs else 'there are no long'} groups of one residual sign.",
        f"The residual spread {'does' if spreading else 'does not clearly'} become larger as level rises (|residual|-fitted correlation {corr_abs_fitted:+.3f}; high/low SD ratio {spread_ratio:.2f}).",
        f"Largest |residual| is {largest:.4f} m ({largest*100:.2f} cm), {'above' if largest > .01 else 'within'} the 1 cm logger accuracy by {abs(max_residual_excess):.4f} m ({abs(max_residual_excess)*100:.2f} cm).",
    ]
    rows = "".join(f"<tr><td>{name}</td><td>{value:.8g}</td><td>{err:.4g}</td><td>{ts:.4g}</td><td>{pv:.4g}</td></tr>" for name,value,err,ts,pv in zip(["a (m/h³)","b (m/h²)","c (m/h)","d (m)"],params,se,t_stats,p_values))
    raw = svg_chart("Raw reservoir level", [(t,h,"#1a73b8","Logger readings",False)], "Elapsed time (h)", "Water level (m)")
    derivatives = svg_chart("First derivative", [(t,fd,"#e07a26","Forward difference",False),(t,cd,"#147d64","Central difference",False),(t,bd,"#7c4dff","Backward difference",False),(dense_t,model_d1,"#1d3557","Fitted cubic dh/dt",True)], "Elapsed time (h)", "dh/dt (m/h)", (max_t,max_rate,f"Maximum: {max_rate:.4g} m/h at {max_t:.3g} h"))
    derivatives += svg_chart("Second derivative", [(t,second,"#b13c6e","Central second difference",False),(dense_t,model_d2,"#1d3557","Fitted cubic d²h/dt²",True)], "Elapsed time (h)", "d²h/dt² (m/h²)")
    fitting = svg_chart("Cubic Levenberg–Marquardt fit", [(t,h,"#1a73b8","Raw readings",False),(dense_t,dense_fit,"#e05d35","Fitted cubic",False)], "Elapsed time (h)", "Water level (m)")
    residual_charts = svg_chart("Residuals vs. time", [(t,residual,"#6c4fb4","Observed − fitted","scatter")], "Elapsed time (h)", "Residual (m)", zero=True)
    residual_charts += svg_chart("Residuals vs. fitted water level", [(fitted,residual,"#6c4fb4","Observed − fitted","scatter")], "Fitted water level (m)", "Residual (m)", zero=True)
    area_chart = svg_chart("Area under fitted curve", [(dense_t,dense_fit,"#e05d35","Fitted cubic",False)], "Elapsed time (h)", "Water level (m)", area=True)
    css = """
    :root{--ink:#122b3b;--muted:#607584;--line:#dce8ed;--navy:#123952;--blue:#187bb5;--aqua:#2cb9af;--gold:#f2b84b;--bg:#f3f8f9;--paper:#fff}*{box-sizing:border-box}body{margin:0;background:radial-gradient(circle at 92% 0,#d9f2f1 0,transparent 25rem),var(--bg);color:var(--ink);font:15px/1.5 Inter,ui-sans-serif,system-ui,-apple-system,"Segoe UI",sans-serif}.hero{position:relative;overflow:hidden;background:linear-gradient(115deg,#0a3048 0%,#155b78 58%,#168c9a 100%);color:#fff;padding:42px max(24px,calc((100% - 1220px)/2))}.hero:before,.hero:after{content:"";position:absolute;border:1px solid #ffffff30;border-radius:50%;pointer-events:none}.hero:before{width:420px;height:420px;right:-140px;top:-270px}.hero:after{width:260px;height:260px;right:90px;bottom:-190px}.hero h1,.hero p{position:relative;z-index:1}.hero h1{margin:0 0 7px;font-size:clamp(1.8rem,4vw,2.6rem);letter-spacing:-.035em}.hero p{margin:0;max-width:790px;opacity:.88}.wrap{max-width:1220px;margin:26px auto 42px;padding:0 20px}.tabs{display:flex;gap:8px;flex-wrap:wrap;margin:22px 0}.tabs button{background:#fff;border:1px solid #dbe7eb;border-radius:999px;padding:10px 17px;color:var(--navy);font-weight:750;letter-spacing:.01em;cursor:pointer;box-shadow:0 2px 5px #193d5210;transition:.18s ease}.tabs button:hover{transform:translateY(-1px);border-color:#92c8d1}.tabs button.active{background:var(--navy);border-color:var(--navy);color:#fff;box-shadow:0 5px 12px #12395233}.tab{display:none}.tab.active{display:block;animation:tab-enter .42s cubic-bezier(.22,.8,.28,1) both}.tab.active>.chart,.tab.active>.card,.tab.active>.grid2,.tab.active>.metrics,.tab.active>.callout{animation:content-rise .48s cubic-bezier(.22,.8,.28,1) both}@keyframes tab-enter{0%{opacity:0;transform:translateY(12px) scale(.992);filter:blur(3px)}100%{opacity:1;transform:translateY(0) scale(1);filter:blur(0)}}@keyframes content-rise{0%{opacity:0;transform:translateY(8px)}100%{opacity:1;transform:translateY(0)}}@media(prefers-reduced-motion:reduce){.tab.active,.tab.active>*{animation:none!important}.tabs button{transition:none}}.grid2{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px}.chart,.card{background:linear-gradient(145deg,#fff,#fbfdfd);border:1px solid var(--line);border-radius:16px;padding:17px;box-shadow:0 8px 22px #183b550d}.chart{position:relative;overflow:hidden}.chart:before{content:"";display:block;width:34px;height:3px;border-radius:2px;background:linear-gradient(90deg,var(--aqua),var(--blue));margin-bottom:10px}.chart h3{font-size:1rem;letter-spacing:-.01em;margin:0 0 2px}.chart svg{width:100%;height:auto;display:block}.legend{font-size:.82rem;color:var(--muted);display:flex;gap:13px;flex-wrap:wrap;padding:2px 4px}.legend i{display:inline-block;width:17px;height:3px;border-radius:3px;vertical-align:middle;margin-right:5px}.metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(175px,1fr));gap:12px}.metric{position:relative;overflow:hidden;background:linear-gradient(145deg,#fff,#f8fcfc);border:1px solid var(--line);border-radius:13px;padding:13px 15px;box-shadow:0 5px 15px #183b550b}.metric:before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:linear-gradient(var(--aqua),var(--blue))}.metric b{display:block;font-size:1.18rem;letter-spacing:-.025em;color:var(--navy)}.metric span{font-size:.82rem;color:var(--muted)}table{border-collapse:separate;border-spacing:0;width:100%;background:#fff;border:1px solid var(--line);border-radius:11px;overflow:hidden}th,td{padding:10px;text-align:left;border-bottom:1px solid var(--line)}tbody tr:last-child td{border-bottom:0}tbody tr:nth-child(even){background:#f7fafb}th{background:linear-gradient(90deg,#e4f3f2,#edf5f8);color:var(--navy);font-size:.84rem;text-transform:uppercase;letter-spacing:.035em}code{background:#edf5f5;color:#155b78;padding:3px 6px;border-radius:5px}.callout{background:linear-gradient(100deg,#fff9e8,#fffdf7);border:1px solid #f1d997;border-left:4px solid var(--gold);border-radius:10px;padding:13px 15px;margin:18px 0;box-shadow:0 4px 10px #6e591308}.interpret h3{margin:0;color:var(--navy)}.interpret li{margin:8px 0}.fit-stats{margin-top:18px}@media(max-width:760px){.grid2{grid-template-columns:1fr}.wrap{padding:0 10px}.hero{padding:29px 16px}.chart,.card{padding:13px}th,td{padding:7px;font-size:.8rem}}
    """
    html_doc = f"""<!doctype html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Water-level model dashboard</title><style>{css}</style></head><body>
    <header class='hero'><h1>Reservoir water-level model</h1><p>Cubic model fitted with Levenberg–Marquardt · {start:%d %b %Y %H:%M}–{end:%d %b %Y %H:%M} PHT · {n} readings at 15-minute intervals</p></header>
    <main class='wrap'><div class='metrics'><div class='metric'><span>Maximum fitted dh/dt</span><b>{max_rate:.5f} m/h</b><span>at {max_t:.3f} elapsed h ({(start+timedelta(hours=max_t)):%d %b %H:%M} PHT)</span></div><div class='metric'><span>Model fit</span><b>R² = {r2:.6f}</b><span>SSE = {sse:.6g} m²</span></div><div class='metric'><span>Area under fitted curve</span><b>{area:.5f} m·h</b><span>quad error ±{quad_error:.2e} m·h</span></div><div class='metric'><span>Largest residual</span><b>{largest:.4f} m</b><span>{'exceeds' if largest>.01 else 'within'} 1 cm accuracy</span></div></div>
    <div class='tabs'><button class='active' onclick="showTab('overview',this)">Main graph</button><button onclick="showTab('derivatives',this)">1 – Derivatives</button><button onclick="showTab('fit',this)">2 – Fitted curve & residuals</button><button onclick="showTab('area',this)">3 – Area</button></div>
    <div id='overview' class='tab active'>{raw}<div class='callout'>Model equation (t in elapsed hours; h in metres): <code>{equation}</code></div></div>
    <div id='derivatives' class='tab'><div class='grid2'>{derivatives}</div><div class='callout'>Finite differences are calculated directly from logger readings. The highlighted maximum is the maximum of the fitted cubic derivative over the observed time range.</div></div>
    <div id='fit' class='tab'><div class='grid2'>{fitting}{residual_charts}</div><div class='card interpret'><h3>Residual interpretation</h3><ul>{''.join(f'<li>{html.escape(x)}</li>' for x in interpretation)}</ul></div><div class='card fit-stats'><h3>Fitting statistics</h3><table><thead><tr><th>Parameter</th><th>Estimate</th><th>Standard error</th><th>t-statistic</th><th>p-value</th></tr></thead><tbody>{rows}</tbody></table></div><div class='metrics fit-stats'><div class='metric'><span>SSE</span><b>{sse:.7g}</b><span>m²</span></div><div class='metric'><span>SST</span><b>{sst:.7g}</b><span>m²</span></div><div class='metric'><span>R²</span><b>{r2:.7f}</b><span>coefficient of determination</span></div><div class='metric'><span>Standard error of estimate</span><b>{see:.7g} m</b><span>√(SSE / df)</span></div><div class='metric'><span>Degrees of freedom</span><b>{dof}</b><span>{n} observations − 4 parameters</span></div></div><div class='callout'>Method: SciPy <code>curve_fit(..., method='lm')</code>, a Levenberg–Marquardt fit of the cubic equation.</div></div>
    <div id='area' class='tab'>{area_chart}<div class='metrics'><div class='metric'><span>Adaptive quadrature (quad)</span><b>{area:.7f} m·h</b><span>Error estimate: ±{quad_error:.3e} m·h</span></div><div class='metric'><span>Trapezoidal cross-check</span><b>{trap:.7f} m·h</b><span>Difference from quad: {abs(trap-area):.3e} m·h</span></div><div class='metric'><span>Integration interval</span><b>0.000–{t[-1]:.3f} h</b><span>Area units: metres × hours = m·h</span></div></div></div>
    </main><script>function showTab(id,button){{document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));document.querySelectorAll('.tabs button').forEach(x=>x.classList.remove('active'));document.getElementById(id).classList.add('active');button.classList.add('active')}}</script></body></html>"""
    output.write_text(html_doc, encoding="utf-8")
    summary = {"source":str(source),"records":n,"time_start_pht":start.isoformat(sep=" "),"time_end_pht":end.isoformat(sep=" "),"equation":equation,"parameters":params.tolist(),"standard_errors":se.tolist(),"t_statistics":t_stats.tolist(),"p_values":p_values.tolist(),"sse":sse,"sst":sst,"r_squared":r2,"standard_error_estimate":see,"degrees_freedom":dof,"maximum_dh_dt_m_per_h":max_rate,"maximum_time_elapsed_hours":max_t,"maximum_time_pht":(start+timedelta(hours=max_t)).isoformat(sep=" "),"area_m_h":area,"quad_error_m_h":quad_error,"trapezoidal_m_h":trap,"largest_abs_residual_m":largest,"logger_accuracy_m":0.01,"residual_interpretation":interpretation}
    output.with_suffix(".json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print(f"Dashboard: {output.resolve()}")
    print(f"Maximum dh/dt: {max_rate:.8f} m/h at {max_t:.6f} h ({(start+timedelta(hours=max_t)):%Y-%m-%d %H:%M:%S} PHT)")
    print(f"Area: {area:.8f} m*h; quad error: {quad_error:.3e} m*h; trapezoid: {trap:.8f} m*h")
    print("\n".join(interpretation))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Cubic reservoir water-level dashboard")
    parser.add_argument("--input", type=Path, default=Path(r"C:\Users\stewa\Downloads\Data 01.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("water_level_dashboard.html"))
    args = parser.parse_args()
    run_analysis(args.input, args.output)
