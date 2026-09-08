"""
NM-LAB-08252026 - Laboratory Activity 02: Fitting a curve to the dam
Levenberg-Marquardt, residuals, and the area under the level

Rule 0: Python computes, HTML displays. Every number on the dashboard is
produced here and written into the HTML as plain values / embedded images.
No fitting, regression, or statistics happens in JavaScript.

HOW TO USE THIS FILE
---------------------
1. Set DATA_PATH below to your stage-log spreadsheet.
2. Fill in the WRITTEN NOTES section with YOUR OWN reading of the
   evidence (residual pattern, second-derivative interpretation, model
   justification, area interpretation). These are your analysis, not
   something a script can generate for you honestly -- the lab asks for
   YOUR judgement here.
3. Run: python lab02_yourname.py
   It writes lab02_yourname.html next to this file.
4. Rename this file (and the .py/.html pair) to lab02_<YourSurname>.py
   / .html before you commit, per the submission spec.
"""

import base64
import io
import json
import re
import warnings
from datetime import datetime

import numpy as np
import pandas as pd
from scipy.optimize import curve_fit
from scipy.integrate import quad
from scipy.stats import t as student_t

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors as rl_colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable,
)

# ============================================================
# COLOR THEME -- primary is #E89EB8; the rest are chosen to sit
# comfortably around it (a deeper rose for contrast/headers, a
# teal-green complement for accents/notes, warm neutrals for text).
# Used consistently across the HTML dashboard, the matplotlib plots,
# and the results PDF so all three deliverables match.
# ============================================================
THEME = {
    "primary": "#E89EB8",       # main theme color (given)
    "primary_dark": "#C9749B",  # deeper rose, header / strong accents
    "primary_darker": "#7A3B57",  # deep plum, headings / body text accent
    "complement": "#4FA88A",    # teal-green, complementary accent (notes, 2nd series)
    "complement_dark": "#357A63",
    "bg": "#FDF3F6",            # page background, faint pink tint
    "card_bg": "#FFFFFF",
    "text": "#3A2A30",          # near-black with a warm plum cast
    "muted": "#8A6F78",
    "grid": "#EBD3DC",
    "raw_dot": "#B98A97",       # muted rose-gray for raw data points
}

# ============================================================
# 0. CONFIGURATION -- edit these
# ============================================================

DATA_PATH = r"C:\Users\Camille\Downloads\Data 01.xlsx"
OUTPUT_HTML = "lab02_dacillo.html"
OUTPUT_PDF = "lab02_dacillo_results.pdf"

# Which candidate model to report on the dashboard as the chosen fit.
# The script fits ALL candidates and prints a comparison table to the
# console so you can make (and defend) an informed choice. Options:
# "logistic", "gompertz", "two_logistic", "cubic"
CHOSEN_MODEL = "logistic"

# ---- WRITTEN NOTES: replace these placeholder strings with your own
# ---- analysis before you submit. Keep them short and specific, the
# ---- way the activity sheet asks.
NOTES = {
    "model_justification": (
        "TODO: two sentences defending your model choice -- what it "
        "assumes about the reservoir and why that matches (or doesn't "
        "match) the shape of the raw plot."
    ),
    "second_derivative": (
        "TODO: one sentence on what d2h/dt2 tells you about the inflow "
        "(not just where it changes sign -- e.g. whether inflow was "
        "still accelerating, already easing off, etc.)."
    ),
    "residual_drift": "TODO: is the scatter centred on zero throughout, or does it drift?",
    "residual_runs": "TODO: are there long runs of the same sign? What does that say about the model shape?",
    "residual_spread": "TODO: does the spread widen as the level rises?",
    "residual_vs_resolution": "TODO: is the largest residual bigger than the logger's 1 cm resolution, and by how much?",
    "area_gap": "TODO: one sentence explaining the gap between the quad integral and the trapezoid cross-check.",
    "area_meaning": "TODO: what the area number means to the flood-control office, and what it does NOT mean (it is meter-hours, not a volume).",
}

# ============================================================
# 1. LOAD DATA
# ============================================================

def _cell_is_datetime(val):
    if pd.isna(val):
        return False
    if isinstance(val, (pd.Timestamp, datetime)):
        return True
    try:
        parsed = pd.to_datetime(str(val), errors="raise")
        return True
    except Exception:
        return False


def _cell_is_numeric(val):
    if pd.isna(val):
        return False
    if isinstance(val, (pd.Timestamp, datetime)):
        return False
    try:
        float(val)
        return True
    except Exception:
        return False


def _find_header_and_data_rows(raw):
    """Scan the first rows of a header-less read for the first row that
    actually looks like data (a parseable timestamp + a parseable number
    somewhere in the row). Handles title rows, merged banners, and blank
    spacer rows sitting above the real table. Returns (header_row_idx or
    None, data_start_idx) using the same 0-based row numbering pandas
    uses for `header=`."""
    n_rows = len(raw)
    scan_limit = min(20, n_rows)
    data_start = None
    for i in range(scan_limit):
        row = raw.iloc[i]
        dt_hits = sum(_cell_is_datetime(v) for v in row)
        num_hits = sum(_cell_is_numeric(v) for v in row)
        if dt_hits >= 1 and num_hits >= 1:
            data_start = i
            break
    if data_start is None:
        return None, None

    header_row = None
    if data_start > 0:
        prev = raw.iloc[data_start - 1]
        non_null = [v for v in prev if not pd.isna(v)]
        if len(non_null) >= 2:
            looks_like_header = all(
                not _cell_is_numeric(v) and not _cell_is_datetime(v) for v in non_null
            )
            if looks_like_header:
                header_row = data_start - 1
    return header_row, data_start


def _detect_columns_by_content(df):
    """Column-name-agnostic fallback: classify each column by what its
    values actually parse as, rather than trusting header text."""
    time_col = None
    for col in df.columns:
        parsed = pd.to_datetime(df[col], errors="coerce")
        if parsed.notna().sum() >= max(3, int(0.8 * len(df))):
            time_col = col
            break

    numeric_candidates = []
    for col in df.columns:
        if col == time_col:
            continue
        series = pd.to_numeric(df[col], errors="coerce")
        if series.notna().sum() >= max(3, int(0.8 * len(df))):
            numeric_candidates.append((col, series.dropna().to_numpy()))

    level_col = None
    for col, vals in numeric_candidates:
        n = len(vals)
        is_index_like = np.allclose(vals, np.arange(1, n + 1)) or np.allclose(
            vals, np.arange(0, n)
        )
        if not is_index_like:
            level_col = col
            break
    if level_col is None and numeric_candidates:
        level_col = numeric_candidates[0][0]

    return time_col, level_col


def load_stage_log(path):
    """Read the stage log and return (t_hours, h_meters, raw_timestamps).

    Handles title/banner rows above the real table (common in spreadsheets
    exported with a merged heading), auto-detects the timestamp and
    water-level columns by content (not just header text), and falls back
    to manual overrides below if it still can't figure it out.
    """
    TIME_COL = None   # e.g. "Timestamp" -- set manually if auto-detect fails
    LEVEL_COL = None  # e.g. "Stage (m)"

    raw = pd.read_excel(path, header=None)
    header_row, data_start = _find_header_and_data_rows(raw)

    if data_start is None:
        preview = raw.head(10).to_string()
        raise ValueError(
            "Could not find a row that looks like actual data (a timestamp "
            "plus a number) anywhere in the first 20 rows. Here is a preview "
            f"of the top of the file so you can see what's going on:\n\n{preview}\n\n"
            "Set TIME_COL / LEVEL_COL manually inside load_stage_log(), or "
            "adjust the sheet."
        )

    if header_row is not None:
        df = pd.read_excel(path, header=header_row)
        df = df.iloc[: len(raw) - data_start].reset_index(drop=True)
    else:
        df = raw.iloc[data_start:].reset_index(drop=True)
        df.columns = [f"col{i}" for i in range(df.shape[1])]

    df = df.dropna(how="all")

    time_col = TIME_COL
    level_col = LEVEL_COL

    # name-based pass first (only useful if we found real header labels)
    if header_row is not None:
        if time_col is None:
            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    time_col = col
                    break
            if time_col is None:
                for col in df.columns:
                    if re.search(r"time|date|stamp", str(col), re.I):
                        time_col = col
                        break
        if level_col is None:
            for col in df.columns:
                if col == time_col:
                    continue
                if re.search(r"level|stage|depth|height|^h$", str(col), re.I):
                    if pd.api.types.is_numeric_dtype(df[col]) or pd.to_numeric(
                        df[col], errors="coerce"
                    ).notna().sum() >= max(3, int(0.8 * len(df))):
                        level_col = col
                        break

    # content-based fallback / confirmation
    if time_col is None or level_col is None:
        ct, cl = _detect_columns_by_content(df)
        time_col = time_col or ct
        level_col = level_col or cl

    if time_col is None or level_col is None:
        raise ValueError(
            f"Could not auto-detect columns. Columns seen: {list(df.columns)}. "
            "Set TIME_COL and LEVEL_COL manually inside load_stage_log()."
        )

    df[time_col] = pd.to_datetime(df[time_col])
    df[level_col] = pd.to_numeric(df[level_col], errors="coerce")
    df = df.dropna(subset=[time_col, level_col])
    df = df.sort_values(time_col).reset_index(drop=True)

    t0 = df[time_col].iloc[0]
    t_hours = (df[time_col] - t0).dt.total_seconds().to_numpy() / 3600.0
    h = df[level_col].to_numpy(dtype=float)

    print(f"[load] header row: {header_row}  data starts at file row: {data_start}")
    print(f"[load] time column: '{time_col}'  level column: '{level_col}'")
    print(f"[load] n = {len(h)} readings, span = {t_hours[-1]:.2f} h")

    return t_hours, h, df[time_col]


# ============================================================
# 2. FINITE-DIFFERENCE DERIVATIVES
# ============================================================

def finite_diff_first(t, h):
    n = len(t)
    d = np.empty(n)
    d[0] = (h[1] - h[0]) / (t[1] - t[0])                    # forward
    d[-1] = (h[-1] - h[-2]) / (t[-1] - t[-2])                # backward
    d[1:-1] = (h[2:] - h[:-2]) / (t[2:] - t[:-2])            # central
    return d


def finite_diff_second(t, h):
    n = len(t)
    d2 = np.empty(n)
    dt0 = t[1] - t[0]
    dt_end = t[-1] - t[-2]
    d2[0] = (h[2] - 2 * h[1] + h[0]) / dt0 ** 2
    d2[-1] = (h[-1] - 2 * h[-2] + h[-3]) / dt_end ** 2
    for i in range(1, n - 1):
        dt_f = t[i + 1] - t[i]
        dt_b = t[i] - t[i - 1]
        d2[i] = 2 * (h[i + 1] * dt_b - h[i] * (dt_f + dt_b) + h[i - 1] * dt_f) / (
            dt_f * dt_b * (dt_f + dt_b)
        )
    return d2


# ============================================================
# 3. CANDIDATE MODELS
# ============================================================

def model_logistic(t, c, a, k, t0):
    return c + a / (1 + np.exp(-k * (t - t0)))


def model_gompertz(t, c, a, k, t0):
    return c + a * np.exp(-np.exp(-k * (t - t0)))


def model_two_logistic(t, c, a1, k1, t01, a2, k2, t02):
    return (
        c
        + a1 / (1 + np.exp(-k1 * (t - t01)))
        + a2 / (1 + np.exp(-k2 * (t - t02)))
    )


def model_cubic(t, a, b, c, d):
    return a * t ** 3 + b * t ** 2 + c * t + d


MODELS = {
    "logistic": {
        "func": model_logistic,
        "params": ["c", "a", "k", "t0"],
        "equation": "h(t) = c + a / (1 + exp(-k (t - t0)))",
        "assumption": "One filling event, a single inflection, level settling toward a ceiling.",
    },
    "gompertz": {
        "func": model_gompertz,
        "params": ["c", "a", "k", "t0"],
        "equation": "h(t) = c + a * exp(-exp(-k (t - t0)))",
        "assumption": "Rises sharply, eases off slowly -- asymmetric filling.",
    },
    "two_logistic": {
        "func": model_two_logistic,
        "params": ["c", "a1", "k1", "t01", "a2", "k2", "t02"],
        "equation": "h(t) = c + a1*S1(t) + a2*S2(t)",
        "assumption": "Two pulses of inflow (e.g. a second rainfall band).",
    },
    "cubic": {
        "func": model_cubic,
        "params": ["a", "b", "c", "d"],
        "equation": "h(t) = a t^3 + b t^2 + c t + d",
        "assumption": "Nothing physical -- baseline to beat.",
    },
}


def estimate_p0(name, t, h):
    """Legitimate, data-driven initial guesses (curve_fit/LM takes no bounds,
    so everything you would express as a bound has to live in p0 instead)."""
    hmin, hmax = float(np.min(h)), float(np.max(h))
    span = hmax - hmin if hmax > hmin else 1.0
    mid = hmin + span / 2
    t_mid = t[np.argmin(np.abs(h - mid))]
    t_range = t[-1] - t[0] if t[-1] > t[0] else 1.0

    if name == "logistic":
        return [hmin, span, 4.0 / max(t_range, 1e-6), t_mid]
    if name == "gompertz":
        return [hmin, span, 4.0 / max(t_range, 1e-6), t_mid]
    if name == "two_logistic":
        t01 = t[0] + t_range / 3
        t02 = t[0] + 2 * t_range / 3
        return [hmin, span / 2, 4.0 / max(t_range, 1e-6), t01,
                span / 2, 4.0 / max(t_range, 1e-6), t02]
    if name == "cubic":
        # a good, legitimate initial guess for a polynomial is a polyfit
        coeffs = np.polyfit(t, h, 3)
        return list(coeffs)
    raise ValueError(name)


def fit_all_candidates(t, h):
    results = {}
    for name, spec in MODELS.items():
        p0 = estimate_p0(name, t, h)
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                popt, pcov = curve_fit(
                    spec["func"], t, h, p0=p0, method="lm", maxfev=20000
                )
            resid = h - spec["func"](t, *popt)
            sse = float(np.sum(resid ** 2))
            sst = float(np.sum((h - h.mean()) ** 2))
            r2 = 1 - sse / sst
            converged = np.all(np.isfinite(pcov))
            results[name] = dict(popt=popt, pcov=pcov, sse=sse, r2=r2,
                                  converged=converged, p0=p0)
        except Exception as exc:
            results[name] = dict(error=str(exc), p0=p0)
    return results


# ============================================================
# 4. STATISTICS FOR THE CHOSEN MODEL
# ============================================================

def compute_stats(model_func, popt, pcov, t, h):
    n = len(h)
    p = len(popt)
    dof = n - p

    resid = h - model_func(t, *popt)
    sse = float(np.sum(resid ** 2))
    sst = float(np.sum((h - h.mean()) ** 2))
    r2 = 1 - sse / sst
    s = float(np.sqrt(sse / dof))

    if np.all(np.isfinite(pcov)):
        se = np.sqrt(np.diag(pcov))
    else:
        se = np.full(p, np.nan)

    with np.errstate(divide="ignore", invalid="ignore"):
        tj = popt / se
        pj = 2 * (1 - student_t.cdf(np.abs(tj), dof))

    return dict(
        n=n, p=p, dof=dof, resid=resid, sse=sse, sst=sst, r2=r2, s=s,
        se=se, tj=tj, pj=pj,
    )


# ============================================================
# 5. INTEGRATION
# ============================================================

def integrate_fit(model_func, popt, t0, tn):
    value, abserr = quad(lambda x: model_func(x, *popt), t0, tn)
    return value, abserr


# ============================================================
# 6. PLOTTING (matplotlib -> base64 PNG, embedded directly in the HTML;
#    no charting library, no CDN, no JS computation -- Rule 0 compliant)
# ============================================================

def fig_to_base64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=140, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("ascii")


def _style_axes(*axes):
    for ax in axes:
        ax.grid(alpha=0.5, color=THEME["grid"], lw=0.8)
        ax.set_facecolor("#FFFFFF")
        for spine in ax.spines.values():
            spine.set_color(THEME["grid"])
        ax.tick_params(colors=THEME["text"], labelsize=8)
        ax.title.set_color(THEME["primary_darker"])
        ax.xaxis.label.set_color(THEME["text"])
        ax.yaxis.label.set_color(THEME["text"])


def plot_raw(t, h):
    fig, ax = plt.subplots(figsize=(9, 3))
    ax.plot(t, h, ".", ms=4, color=THEME["primary_dark"])
    ax.set_xlabel("Time (h)")
    ax.set_ylabel("Level (m)")
    ax.set_title("Raw stage log")
    _style_axes(ax)
    return fig_to_base64(fig)


def plot_derivatives(t, dh, d2h, t_max_dhdt):
    fig, axes = plt.subplots(2, 1, figsize=(9, 6), sharex=True)
    axes[0].plot(t, dh, color=THEME["primary_dark"])
    axes[0].axvline(t_max_dhdt, color=THEME["complement_dark"], ls="--", lw=1.2,
                     label=f"max dh/dt at t={t_max_dhdt:.2f} h")
    axes[0].set_ylabel("dh/dt (m/h)")
    axes[0].legend(fontsize=8)
    axes[0].set_title("First finite-difference derivative")

    axes[1].plot(t, d2h, color=THEME["complement_dark"])
    axes[1].axhline(0, color=THEME["muted"], lw=0.7)
    axes[1].set_ylabel("d2h/dt2 (m/h2)")
    axes[1].set_xlabel("Time (h)")
    axes[1].set_title("Second finite-difference derivative")

    _style_axes(*axes)
    fig.tight_layout()
    return fig_to_base64(fig)


def plot_fit(t, h, model_func, popt):
    fig, ax = plt.subplots(figsize=(9, 3.5))
    ax.plot(t, h, ".", ms=4, color=THEME["raw_dot"], label="raw log")
    tt = np.linspace(t[0], t[-1], 400)
    ax.plot(tt, model_func(tt, *popt), color=THEME["complement_dark"], lw=2, label="fitted h(t)")
    ax.set_xlabel("Time (h)")
    ax.set_ylabel("Level (m)")
    ax.legend(fontsize=8)
    ax.set_title("Fitted curve over raw data")
    _style_axes(ax)
    return fig_to_base64(fig)


def plot_residuals(t, h_hat, resid):
    fig, axes = plt.subplots(1, 2, figsize=(9, 3.2))
    axes[0].axhline(0, color=THEME["muted"], lw=0.8)
    axes[0].plot(t, resid, ".", ms=4, color=THEME["primary_dark"])
    axes[0].set_xlabel("Time (h)")
    axes[0].set_ylabel("Residual (m)")
    axes[0].set_title("Residuals vs time")

    axes[1].axhline(0, color=THEME["muted"], lw=0.8)
    axes[1].plot(h_hat, resid, ".", ms=4, color=THEME["complement_dark"])
    axes[1].set_xlabel("Fitted h (m)")
    axes[1].set_ylabel("Residual (m)")
    axes[1].set_title("Residuals vs fitted value")

    _style_axes(*axes)
    fig.tight_layout()
    return fig_to_base64(fig)


def plot_area(t, h, model_func, popt, t0, tn):
    fig, ax = plt.subplots(figsize=(9, 3.5))
    ax.plot(t, h, ".", ms=4, color=THEME["raw_dot"], label="raw log")
    tt = np.linspace(t0, tn, 400)
    yy = model_func(tt, *popt)
    ax.plot(tt, yy, color=THEME["primary_dark"], lw=2, label="fitted h(t)")
    ax.fill_between(tt, 0, yy, color=THEME["primary"], alpha=0.35)
    ax.set_ylim(bottom=0)
    ax.set_xlabel("Time (h)")
    ax.set_ylabel("Level (m)")
    ax.legend(fontsize=8)
    ax.set_title("Area under the fitted level")
    _style_axes(ax)
    return fig_to_base64(fig)


# ============================================================
# 7. RESULTS PDF (one page: parameters, SE/t/p, SSE, R2, s, area, residual reading)
# ============================================================

def build_results_pdf(
    out_path, model_name, spec, popt, stats, area_value, area_abserr,
    area_trapz, area_diff, notes,
):
    primary_dark = rl_colors.HexColor(THEME["primary_dark"])
    primary_darker = rl_colors.HexColor(THEME["primary_darker"])
    complement_dark = rl_colors.HexColor(THEME["complement_dark"])
    header_fill = rl_colors.HexColor("#FBE4EC")
    note_fill = rl_colors.HexColor("#EAF6F1")
    grid_color = rl_colors.HexColor(THEME["grid"])

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleStyle", parent=styles["Title"], textColor=primary_darker,
        fontSize=17, spaceAfter=2,
    )
    sub_style = ParagraphStyle(
        "SubStyle", parent=styles["Normal"], textColor=rl_colors.HexColor(THEME["muted"]),
        fontSize=9, spaceAfter=10,
    )
    h2_style = ParagraphStyle(
        "H2Style", parent=styles["Heading2"], textColor=primary_darker,
        fontSize=11.5, spaceBefore=10, spaceAfter=4,
    )
    body_style = ParagraphStyle(
        "BodyStyle", parent=styles["Normal"], fontSize=8.8, leading=12,
        textColor=rl_colors.HexColor(THEME["text"]),
    )
    note_style = ParagraphStyle(
        "NoteStyle", parent=body_style, backColor=note_fill,
        borderPadding=(5, 6, 5, 6), leftIndent=2,
    )

    story = []
    story.append(Paragraph("Fitting a curve to the dam &mdash; Results", title_style))
    story.append(Paragraph(
        f"NM-LAB-08252026 &middot; Numerical Methods BES6-M &middot; "
        f"generated {datetime.now().strftime('%Y-%m-%d %H:%M')}", sub_style
    ))
    story.append(HRFlowable(width="100%", thickness=1, color=grid_color, spaceAfter=8))

    story.append(Paragraph(
        f"Chosen model: <b>{model_name}</b> &nbsp;&mdash;&nbsp; {spec['equation']}", body_style
    ))
    story.append(Spacer(1, 8))

    # --- parameter table ---
    story.append(Paragraph("Fitted parameters", h2_style))
    param_data = [["Parameter", "Value", "Std. Error", "t statistic", "p-value"]]
    for name, val, se, tj, pj in zip(spec["params"], popt, stats["se"], stats["tj"], stats["pj"]):
        param_data.append([name, f"{val:.4f}", f"{se:.4f}", f"{tj:.3f}", f"{pj:.4g}"])
    param_table = Table(param_data, hAlign="LEFT", colWidths=[80, 90, 90, 90, 90])
    param_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_fill),
        ("TEXTCOLOR", (0, 0), (-1, 0), primary_darker),
        ("FONTSIZE", (0, 0), (-1, -1), 8.8),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, grid_color),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(param_table)
    story.append(Spacer(1, 8))

    # --- fit statistics ---
    story.append(Paragraph("Fit statistics", h2_style))
    fit_data = [
        ["n", "p", "dof", "SSE (m2)", "SST (m2)", "R2", "s (m)"],
        [
            str(stats["n"]), str(stats["p"]), str(stats["dof"]),
            f"{stats['sse']:.5f}", f"{stats['sst']:.5f}",
            f"{stats['r2']:.5f}", f"{stats['s']:.4f}",
        ],
    ]
    fit_table = Table(fit_data, hAlign="LEFT", colWidths=[45, 35, 40, 68, 68, 68, 60])
    fit_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_fill),
        ("TEXTCOLOR", (0, 0), (-1, 0), primary_darker),
        ("FONTSIZE", (0, 0), (-1, -1), 8.8),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, grid_color),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(fit_table)
    story.append(Spacer(1, 8))

    # --- area under the curve ---
    story.append(Paragraph("Area under the fitted level", h2_style))
    area_data = [
        ["Quad integral A", "Abs. error (quad)", "Trapezoid cross-check", "Difference"],
        [
            f"{area_value:.5f} m&middot;h", f"{area_abserr:.2e} m&middot;h",
            f"{area_trapz:.5f} m&middot;h", f"{area_diff:.5f} m&middot;h",
        ],
    ]
    area_data = [[Paragraph(c, body_style) for c in row] for row in area_data]
    area_table = Table(area_data, hAlign="LEFT", colWidths=[130, 120, 130, 100])
    area_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_fill),
        ("GRID", (0, 0), (-1, -1), 0.5, grid_color),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(area_table)
    story.append(Paragraph(
        "Units: h in meters, t in hours &rarr; area is in meter-hours, not a volume.",
        ParagraphStyle("small", parent=body_style, fontSize=7.8,
                        textColor=rl_colors.HexColor(THEME["muted"])),
    ))
    story.append(Spacer(1, 8))

    # --- residual reading ---
    story.append(Paragraph("Reading of the residuals", h2_style))
    residual_text = (
        f"<b>Centred / drift:</b> {notes['residual_drift']}<br/>"
        f"<b>Runs:</b> {notes['residual_runs']}<br/>"
        f"<b>Spread:</b> {notes['residual_spread']}<br/>"
        f"<b>Vs. logger resolution (1 cm):</b> {notes['residual_vs_resolution']}"
    )
    story.append(Paragraph(residual_text, note_style))

    doc = SimpleDocTemplate(
        out_path, pagesize=letter,
        topMargin=0.6 * inch, bottomMargin=0.6 * inch,
        leftMargin=0.7 * inch, rightMargin=0.7 * inch,
    )
    doc.build(story)


# ============================================================
# 8. HTML DASHBOARD
# ============================================================

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>NM-LAB-08252026 - Dashboard</title>
<style>
  :root {{
    --primary: #E89EB8;
    --primary-dark: #C9749B;
    --primary-darker: #7A3B57;
    --complement: #4FA88A;
    --complement-dark: #357A63;
    --bg: #FDF3F6;
    --card-bg: #FFFFFF;
    --text: #3A2A30;
    --muted: #8A6F78;
    --grid: #EBD3DC;
  }}
  body {{ font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif;
         margin: 0; background: var(--bg); color: var(--text); }}
  header {{ background: linear-gradient(135deg, var(--primary-darker), var(--primary-dark));
            color: #fff; padding: 24px 32px; }}
  header h1 {{ margin: 0 0 4px 0; font-size: 22px; }}
  header p {{ margin: 0; color: #F6DCE6; font-size: 13px; }}
  .wrap {{ max-width: 980px; margin: 0 auto; padding: 24px; }}
  .card {{ background: var(--card-bg); border-radius: 8px; padding: 20px 24px;
           margin-bottom: 20px; box-shadow: 0 1px 3px rgba(122,59,87,0.12);
           border: 1px solid var(--grid); }}
  .card h2 {{ margin-top: 0; font-size: 16px; color: var(--primary-darker); }}
  img {{ max-width: 100%; display: block; margin: 8px 0; border-radius: 6px; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 13px; margin-top: 8px; }}
  th, td {{ text-align: left; padding: 6px 10px; border-bottom: 1px solid var(--grid); }}
  th {{ background: #FBE4EC; color: var(--primary-darker); }}
  .tabs {{ display: flex; gap: 6px; margin-bottom: 12px; }}
  .tab-btn {{ padding: 8px 16px; border: none; border-radius: 6px 6px 0 0;
              background: #F3D9E2; color: var(--primary-darker); cursor: pointer; font-size: 13px; }}
  .tab-btn.active {{ background: var(--primary-dark); color: #fff; }}
  .tab-panel {{ display: none; }}
  .tab-panel.active {{ display: block; }}
  .stat-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(160px,1fr));
                gap: 10px; margin-top: 10px; }}
  .stat-box {{ background: #FDF1F5; border-radius: 6px; padding: 10px 12px;
               border: 1px solid var(--grid); }}
  .stat-box .label {{ font-size: 11px; color: var(--muted); text-transform: uppercase; }}
  .stat-box .value {{ font-size: 18px; font-weight: 600; color: var(--primary-darker); }}
  .note {{ background: #EAF6F1; border-left: 3px solid var(--complement); padding: 8px 12px;
           font-size: 13px; margin-top: 10px; color: var(--text); }}
  .note b {{ color: var(--complement-dark); }}
  code {{ background: #FBE4EC; color: var(--primary-darker); padding: 2px 5px; border-radius: 4px; }}
</style>
</head>
<body>
<header>
  <h1>Fitting a curve to the dam</h1>
  <p>NM-LAB-08252026 &middot; Numerical Methods BES6-M &middot; generated {generated_at}</p>
</header>
<div class="wrap">

  <div class="card">
    <h2>Stage log time series ({n} readings, {span:.2f} h span)</h2>
    <img src="data:image/png;base64,{img_raw}" alt="raw stage log">
  </div>

  <div class="tabs">
    <button class="tab-btn active" onclick="showTab('t1', this)">Derivatives</button>
    <button class="tab-btn" onclick="showTab('t2', this)">Fit &amp; statistics</button>
    <button class="tab-btn" onclick="showTab('t3', this)">Area under the curve</button>
  </div>

  <div id="t1" class="tab-panel active card">
    <h2>Finite-difference derivatives</h2>
    <img src="data:image/png;base64,{img_deriv}" alt="derivatives">
    <div class="stat-grid">
      <div class="stat-box"><div class="label">Max dh/dt</div>
        <div class="value">{max_dhdt:.4f} m/h</div></div>
      <div class="stat-box"><div class="label">At time</div>
        <div class="value">{t_max_dhdt:.2f} h</div></div>
    </div>
    <div class="note"><b>Second derivative reading:</b> {note_second_derivative}</div>
  </div>

  <div id="t2" class="tab-panel card">
    <h2>Chosen model: {model_name}</h2>
    <p><code>{equation}</code><br>Assumption: {assumption}</p>
    <div class="note"><b>Model justification:</b> {note_model_justification}</div>
    <img src="data:image/png;base64,{img_fit}" alt="fitted curve">

    <h2>Fitted parameters</h2>
    <table>
      <tr><th>Parameter</th><th>Value</th><th>Std. Error</th><th>t statistic</th><th>p-value</th></tr>
      {param_rows}
    </table>

    <div class="stat-grid">
      <div class="stat-box"><div class="label">n, p, dof</div>
        <div class="value">{n} / {p} / {dof}</div></div>
      <div class="stat-box"><div class="label">SSE</div><div class="value">{sse:.5f} m2</div></div>
      <div class="stat-box"><div class="label">SST</div><div class="value">{sst:.5f} m2</div></div>
      <div class="stat-box"><div class="label">R2</div><div class="value">{r2:.5f}</div></div>
      <div class="stat-box"><div class="label">Std. error of estimate s</div>
        <div class="value">{s:.4f} m</div></div>
    </div>

    <h2>Residuals</h2>
    <img src="data:image/png;base64,{img_resid}" alt="residuals">
    <div class="note"><b>Centred / drift:</b> {note_residual_drift}</div>
    <div class="note"><b>Runs:</b> {note_residual_runs}</div>
    <div class="note"><b>Spread:</b> {note_residual_spread}</div>
    <div class="note"><b>Vs. logger resolution (1 cm):</b> {note_residual_vs_resolution}</div>
  </div>

  <div id="t3" class="tab-panel card">
    <h2>Area under the fitted level</h2>
    <img src="data:image/png;base64,{img_area}" alt="area under the curve">
    <div class="stat-grid">
      <div class="stat-box"><div class="label">Quad integral A</div>
        <div class="value">{area_value:.5f} m&middot;h</div></div>
      <div class="stat-box"><div class="label">Absolute error (quad)</div>
        <div class="value">{area_abserr:.2e} m&middot;h</div></div>
      <div class="stat-box"><div class="label">Trapezoid cross-check</div>
        <div class="value">{area_trapz:.5f} m&middot;h</div></div>
      <div class="stat-box"><div class="label">Difference</div>
        <div class="value">{area_diff:.5f} m&middot;h</div></div>
    </div>
    <div class="note"><b>Gap explanation:</b> {note_area_gap}</div>
    <div class="note"><b>What this means to the flood-control office:</b> {note_area_meaning}</div>
  </div>

</div>
<script>
  // UI-only: tab switching. No fitting, regression, or statistics here.
  function showTab(id, btn) {{
    document.querySelectorAll('.tab-panel').forEach(function(p) {{ p.classList.remove('active'); }});
    document.querySelectorAll('.tab-btn').forEach(function(b) {{ b.classList.remove('active'); }});
    document.getElementById(id).classList.add('active');
    btn.classList.add('active');
  }}
</script>
</body>
</html>
"""


def build_dashboard(path):
    t, h, timestamps = load_stage_log(path)

    # --- derivatives ---
    dh = finite_diff_first(t, h)
    d2h = finite_diff_second(t, h)
    imax = int(np.argmax(dh))
    max_dhdt = float(dh[imax])
    t_max_dhdt = float(t[imax])

    # --- fit all candidates, report comparison to console ---
    all_fits = fit_all_candidates(t, h)
    print("\n[model comparison]")
    print(f"{'model':14s} {'converged':10s} {'SSE':>12s} {'R2':>10s}")
    for name, res in all_fits.items():
        if "error" in res:
            print(f"{name:14s} {'FAILED':10s} {'--':>12s} {'--':>10s}  ({res['error']})")
        else:
            print(f"{name:14s} {str(res['converged']):10s} {res['sse']:12.5f} {res['r2']:10.5f}")

    if CHOSEN_MODEL not in all_fits or "error" in all_fits[CHOSEN_MODEL]:
        raise RuntimeError(
            f"CHOSEN_MODEL='{CHOSEN_MODEL}' did not fit successfully. "
            "Pick a different model or check p0."
        )

    spec = MODELS[CHOSEN_MODEL]
    model_func = spec["func"]
    fit = all_fits[CHOSEN_MODEL]
    popt, pcov = fit["popt"], fit["pcov"]

    stats = compute_stats(model_func, popt, pcov, t, h)
    h_hat = model_func(t, *popt)

    # --- integration ---
    area_value, area_abserr = integrate_fit(model_func, popt, t[0], t[-1])
    area_trapz = float(np.trapezoid(h, t))
    area_diff = area_value - area_trapz

    # --- plots ---
    img_raw = plot_raw(t, h)
    img_deriv = plot_derivatives(t, dh, d2h, t_max_dhdt)
    img_fit = plot_fit(t, h, model_func, popt)
    img_resid = plot_residuals(t, h_hat, stats["resid"])
    img_area = plot_area(t, h, model_func, popt, t[0], t[-1])

    param_rows = ""
    for name, val, se, tj, pj in zip(
        spec["params"], popt, stats["se"], stats["tj"], stats["pj"]
    ):
        param_rows += (
            f"<tr><td>{name}</td><td>{val:.4f}</td><td>{se:.4f}</td>"
            f"<td>{tj:.3f}</td><td>{pj:.4g}</td></tr>\n"
        )

    html = HTML_TEMPLATE.format(
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
        n=stats["n"], span=t[-1] - t[0],
        img_raw=img_raw,
        img_deriv=img_deriv,
        max_dhdt=max_dhdt, t_max_dhdt=t_max_dhdt,
        note_second_derivative=NOTES["second_derivative"],
        model_name=CHOSEN_MODEL, equation=spec["equation"], assumption=spec["assumption"],
        note_model_justification=NOTES["model_justification"],
        img_fit=img_fit,
        param_rows=param_rows,
        p=stats["p"], dof=stats["dof"],
        sse=stats["sse"], sst=stats["sst"], r2=stats["r2"], s=stats["s"],
        img_resid=img_resid,
        note_residual_drift=NOTES["residual_drift"],
        note_residual_runs=NOTES["residual_runs"],
        note_residual_spread=NOTES["residual_spread"],
        note_residual_vs_resolution=NOTES["residual_vs_resolution"],
        img_area=img_area,
        area_value=area_value, area_abserr=area_abserr,
        area_trapz=area_trapz, area_diff=area_diff,
        note_area_gap=NOTES["area_gap"],
        note_area_meaning=NOTES["area_meaning"],
    )

    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    build_results_pdf(
        OUTPUT_PDF, CHOSEN_MODEL, spec, popt, stats,
        area_value, area_abserr, area_trapz, area_diff, NOTES,
    )

    print(f"\n[done] wrote {OUTPUT_HTML}")
    print(f"[done] wrote {OUTPUT_PDF}")
    print(f"[summary] chosen model = {CHOSEN_MODEL}")
    for name, val, se, tj, pj in zip(spec["params"], popt, stats["se"], stats["tj"], stats["pj"]):
        print(f"  {name:>4s} = {val:.4f}  (SE={se:.4f}, t={tj:.3f}, p={pj:.4g})")
    print(f"  SSE={stats['sse']:.5f}  R2={stats['r2']:.5f}  s={stats['s']:.4f}  dof={stats['dof']}")
    print(f"  max dh/dt = {max_dhdt:.4f} m/h at t = {t_max_dhdt:.2f} h")
    print(f"  Area (quad) = {area_value:.5f} +/- {area_abserr:.2e} m*h")
    print(f"  Area (trapezoid) = {area_trapz:.5f} m*h  (diff = {area_diff:.5f})")


if __name__ == "__main__":
    build_dashboard(DATA_PATH)