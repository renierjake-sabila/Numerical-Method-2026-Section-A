"""
Laboratory Exercise 3
---------------------
Implements the Taylor / Maclaurin series:

    e^x = sum_{n=0}^{infinity} x^n / n!

evaluated at x = 1 (so the series converges to e), summing up to
10,000 terms. Visualizes the partial sums with a Matplotlib gradient
histogram plus a gradient-stroke log-log accuracy plot showing how many
correct decimal digits are gained as more terms are added.

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
from matplotlib.collections import LineCollection

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
# Gradient bar helper (see Exercise 1 for full explanation)
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


def style_axes(ax, title, xlabel, ylabel, bg=PALETTE["white"]):
    ax.set_facecolor(bg)
    ax.set_title(title, color=PALETTE["dark_green"], fontsize=14, fontweight="bold", pad=14)
    ax.set_xlabel(xlabel, color=PALETTE["dark_green"], fontsize=10.5)
    ax.set_ylabel(ylabel, color=PALETTE["dark_green"], fontsize=10.5)
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    for side in ("left", "bottom"):
        ax.spines[side].set_color(PALETTE["olive"])
        ax.spines[side].set_linewidth(1.2)
    ax.tick_params(colors=PALETTE["dark_green"], labelsize=9)
    ax.set_axisbelow(True)


cmap_value = mcolors.LinearSegmentedColormap.from_list(
    "value_grad", ["#fefae0", "#dda15e", "#606c38"])
cmap_line = mcolors.LinearSegmentedColormap.from_list(
    "line_grad", ["#bc6c25", "#dda15e", "green", "#283618"])

# ----------------------------------------------------------------------
# 1. Compute the Taylor series partial sums for e^x at x = 1
# ----------------------------------------------------------------------
X = 1.0
N_TERMS = 10_000
checkpoints = [1, 3, 5, 10, 15, 20, 25, 30, 50, 100, 1000, N_TERMS]

term = 1.0
partial_sum = 0.0
checkpoint_sums = {}

for n in range(N_TERMS):
    partial_sum += term
    if (n + 1) in checkpoints:
        checkpoint_sums[n + 1] = partial_sum
    term *= X / (n + 1)

e_true = math.e

print(f"{'N terms':>10}{'partial sum S_N':>20}{'|S_N - e|':>18}")
print("-" * 48)
for n in checkpoints:
    s = checkpoint_sums[n]
    print(f"{n:>10}{s:>20.15f}{abs(s - e_true):>18.2e}")
print("-" * 48)
print(f"{'e (true)':>10}{e_true:>20.15f}")

# ----------------------------------------------------------------------
# 2. Draw the histogram (gradient bar chart) in Matplotlib
# ----------------------------------------------------------------------
labels = [str(n) for n in checkpoints]
values = [checkpoint_sums[n] for n in checkpoints]
errors = [abs(v - e_true) for v in values]

fig, axes = plt.subplots(1, 2, figsize=(16, 7.5))
fig.patch.set_facecolor(PALETTE["cream"])

# --- Left panel: histogram of partial sums ---------------------------
ax = axes[0]
style_axes(ax, "Histogram of Partial Sums S_N", "Number of terms N", "S_N = Σ xⁿ/n!  (x = 1)")
x = np.arange(len(labels))
for xi, val in zip(x, values):
    gradient_bar(ax, xi, val, width=0.68, cmap=cmap_value)

ax.axhline(e_true, color=PALETTE["dark_green"], linestyle="--", linewidth=1.8, zorder=5,
           label=f"e = {e_true:.6f}")
for xi, val in zip(x, values):
    ax.text(xi, val + 0.06, f"{val:.4f}", ha="center", va="bottom", fontsize=7.5,
            rotation=90, color=PALETTE["dark_green"], zorder=6)

ax.set_xticks(x)
ax.set_xticklabels(labels, rotation=40, ha="right", fontsize=9)
ax.set_ylim(0, 3.25)
ax.set_xlim(-0.6, len(labels) - 0.4)
ax.grid(axis="y", color=PALETTE["tan"], alpha=0.35, linewidth=0.9, zorder=0)
ax.legend(facecolor=PALETTE["white"], edgecolor=PALETTE["tan"], fontsize=9.5, loc="lower right")

# --- Right panel: convergence / accuracy with gradient stroke --------
ax2 = axes[1]
style_axes(ax2, "Convergence: |S_N − e| (log-log)", "Number of terms N (log scale)",
           "|S_N − e|  (log scale)")
plot_errors = [max(e, 1e-17) for e in errors]

pts = np.array([checkpoints, plot_errors]).T
segments = np.concatenate([pts[:-1, None, :], pts[1:, None, :]], axis=1)
lc = LineCollection(segments, cmap=cmap_line, linewidth=3.2, zorder=4)
lc.set_array(np.linspace(0, 1, len(segments)))
ax2.add_collection(lc)

# glowing markers: soft halo + solid center
ax2.scatter(checkpoints, plot_errors, s=110, color=PALETTE["dark_green"], alpha=0.12, zorder=3)
sc = ax2.scatter(checkpoints, plot_errors, s=42, c=np.linspace(0, 1, len(checkpoints)),
                  cmap=cmap_line, edgecolor=PALETTE["dark_green"], linewidth=1.0, zorder=5)

ax2.set_xscale("log")
ax2.set_yscale("log")
ax2.set_xlim(0.7, N_TERMS * 1.5)
ax2.grid(True, which="both", color=PALETTE["tan"], alpha=0.30, linewidth=0.7, zorder=0)

fig.suptitle(f"Exercise 3 — e^x = Σ xⁿ/n!  (x = 1, summed to N = {N_TERMS:,} terms)",
             color=PALETTE["dark_green"], fontsize=17, fontweight="bold", y=0.995)
fig.subplots_adjust(left=0.06, right=0.98, top=0.87, bottom=0.13, wspace=0.22)

output_path = "exercise3_histogram.png"
fig.savefig(output_path, dpi=220, facecolor=fig.get_facecolor())
print(f"\nSaved histogram to {output_path}")
