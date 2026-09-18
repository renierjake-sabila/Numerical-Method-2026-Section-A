"""
Laboratory Exercise 1
---------------------
Implements (1 + 1/n)^n for compounding periods ranging from yearly down to
every nanosecond, and visualizes the convergence to Euler's number e with
a Matplotlib histogram (gradient bar chart).

Color palette (required):
    #606c38  (olive)
    #283618  (dark green)
    #dda15e  (tan)
    #bc6c25  (brown)
    #fefae0  (cream)
    white
    green
"""

import math
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
from matplotlib.patches import Rectangle

PALETTE = {
    "olive": "#606c38",
    "dark_green": "#283618",
    "tan": "#dda15e",
    "brown": "#bc6c25",
    "cream": "#fefae0",
    "white": "#ffffff",
    "green": "green",
}

plt.rcParams["font.family"] = "DejaVu Sans"

# ----------------------------------------------------------------------
# Gradient bar helper — draws a smooth cream -> palette-color gradient
# fill inside each bar by stacking many thin slices (correctly handles
# both linear AND log-scaled axes), with a soft shadow underneath for
# depth and a crisp outline on top for a clean, vibrant, poster finish.
# ----------------------------------------------------------------------
def gradient_bar(ax, xcenter, top, width, cmap, bottom=0.0, n_seg=90, log=False,
                  edgecolor="#283618", linewidth=1.4, zorder=3, shadow=True):
    if top <= bottom:
        return
    if shadow:
        ax.add_patch(Rectangle(
            (xcenter - width / 2 + width * 0.06, bottom),
            width, (top - bottom) * 1.01,
            facecolor="#283618", alpha=0.10, zorder=zorder - 1, linewidth=0))
    if log:
        edges = np.logspace(np.log10(bottom), np.log10(top), n_seg + 1)
    else:
        edges = np.linspace(bottom, top, n_seg + 1)
    for i in range(n_seg):
        y0, y1 = edges[i], edges[i + 1]
        ax.add_patch(Rectangle((xcenter - width / 2, y0), width, (y1 - y0) * 1.01,
                                facecolor=cmap(i / max(n_seg - 1, 1)),
                                edgecolor="none", zorder=zorder))
    ax.add_patch(Rectangle((xcenter - width / 2, bottom), width, top - bottom,
                            fill=False, edgecolor=edgecolor, linewidth=linewidth,
                            zorder=zorder + 2))


def style_axes(ax, title, ylabel, bg=PALETTE["white"]):
    ax.set_facecolor(bg)
    ax.set_title(title, color=PALETTE["dark_green"], fontsize=14, fontweight="bold", pad=14)
    ax.set_ylabel(ylabel, color=PALETTE["dark_green"], fontsize=11, fontweight="medium")
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    for side in ("left", "bottom"):
        ax.spines[side].set_color(PALETTE["olive"])
        ax.spines[side].set_linewidth(1.2)
    ax.tick_params(colors=PALETTE["dark_green"], labelsize=9)
    ax.grid(axis="y", color=PALETTE["tan"], alpha=0.35, linewidth=0.9, zorder=0)
    ax.set_axisbelow(True)


# Bright warm gradients built from the required palette
cmap_value = mcolors.LinearSegmentedColormap.from_list(
    "value_grad", ["#fefae0", "#dda15e", "#606c38"])
cmap_error = mcolors.LinearSegmentedColormap.from_list(
    "error_grad", ["#fefae0", "#dda15e", "#bc6c25"])

# ----------------------------------------------------------------------
# 1. Build the table of compounding periods -> n -> (1 + 1/n)^n
# ----------------------------------------------------------------------
SECONDS_PER_YEAR = 365 * 24 * 60 * 60  # 31,536,000

periods = [
    ("yearly",              1),
    ("twice a year",        2),
    ("quarterly",           4),
    ("monthly",             12),
    ("weekly",              52),
    ("daily",               365),
    ("hourly",              365 * 24),
    ("every minute",        365 * 24 * 60),
    ("every second",        SECONDS_PER_YEAR),
    ("every millisecond",   SECONDS_PER_YEAR * 1_000),
    ("every microsecond",   SECONDS_PER_YEAR * 1_000_000),
    ("every nanosecond",    SECONDS_PER_YEAR * 1_000_000_000),
]


def compound_value(n: int) -> float:
    """
    Numerically stable computation of (1 + 1/n)^n via
    exp(n * log1p(1/n)), accurate even for n ~ 3.15e16.
    """
    return math.exp(n * math.log1p(1.0 / n))


labels, n_values, results = [], [], []

print(f"{'How often':<20}{'n':>20}{'(1 + 1/n)^n':>20}")
print("-" * 60)
for label, n in periods:
    value = compound_value(n)
    labels.append(label)
    n_values.append(n)
    results.append(value)
    print(f"{label:<20}{n:>20,}{value:>20.10f}")

e_true = math.e
print("-" * 60)
print(f"{'Euler number e':<20}{'':>20}{e_true:>20.10f}")

# ----------------------------------------------------------------------
# 2. Draw the histogram (gradient bar chart) in Matplotlib
# ----------------------------------------------------------------------
fig, axes = plt.subplots(1, 2, figsize=(16, 7.5))
fig.patch.set_facecolor(PALETTE["cream"])

# --- Left panel: value per compounding period -----------------------
ax = axes[0]
style_axes(ax, "Value per Compounding Period", "(1 + 1/n)ⁿ")
x = np.arange(len(labels))
for xi, value in zip(x, results):
    gradient_bar(ax, xi, value, width=0.68, cmap=cmap_value)

ax.axhline(e_true, color=PALETTE["dark_green"], linestyle="--", linewidth=1.8,
           zorder=5, label=f"e = {e_true:.6f}")
for xi, value in zip(x, results):
    ax.text(xi, value + 0.06, f"{value:.4f}", ha="center", va="bottom",
            fontsize=7.5, rotation=90, color=PALETTE["dark_green"], fontweight="medium", zorder=6)

ax.set_xticks(x)
ax.set_xticklabels(labels, rotation=40, ha="right", fontsize=8.5)
ax.set_ylim(0, 3.25)
ax.set_xlim(-0.6, len(labels) - 0.4)
ax.legend(facecolor=PALETTE["white"], edgecolor=PALETTE["tan"], fontsize=9, loc="lower right")

# --- Right panel: error shrinking as n grows -------------------------
ax2 = axes[1]
style_axes(ax2, "Error vs. e (log scale)", "|(1 + 1/n)ⁿ − e|")
errors = [abs(v - e_true) for v in results]
for xi, err in zip(x, errors):
    gradient_bar(ax2, xi, err, width=0.68, cmap=cmap_error, bottom=1e-16, log=True)

ax2.set_yscale("log")
ax2.set_xticks(x)
ax2.set_xticklabels(labels, rotation=40, ha="right", fontsize=8.5)
ax2.set_xlim(-0.6, len(labels) - 0.4)
ax2.grid(axis="y", which="both", color=PALETTE["tan"], alpha=0.30, linewidth=0.7, zorder=0)

fig.suptitle("Exercise 1 — Convergence of (1 + 1/n)ⁿ to e",
             color=PALETTE["dark_green"], fontsize=17, fontweight="bold", y=0.995)
fig.subplots_adjust(left=0.06, right=0.98, top=0.87, bottom=0.22, wspace=0.22)

output_path = "exercise1_histogram.png"
fig.savefig(output_path, dpi=220, facecolor=fig.get_facecolor())
print(f"\nSaved histogram to {output_path}")
