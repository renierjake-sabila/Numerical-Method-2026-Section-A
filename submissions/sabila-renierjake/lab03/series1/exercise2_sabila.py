"""
Laboratory Exercise 2
---------------------
Implements the difference quotient (a^h - 1) / h for three bases
(a = 2, a = e = 2.71828..., a = 3) as h shrinks from 0.1 down to a
tolerance of 1e-6, and visualizes the results with a Matplotlib
histogram (grouped gradient bar chart).

As h -> 0, (a^h - 1)/h -> ln(a). This is the definition-of-derivative
argument for why d/dx[a^x] = ln(a) * a^x, and why a = e is special
(ln(e) = 1).

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
# Gradient bar helper (see Exercise 1 for full explanation)
# ----------------------------------------------------------------------
def gradient_bar(ax, xcenter, top, width, cmap, bottom=0.0, n_seg=90, log=False,
                  edgecolor="#283618", linewidth=1.3, zorder=3, shadow=True):
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


# Bright gradients — each series rises from cream into its accent color
cmap_a2 = mcolors.LinearSegmentedColormap.from_list("a2_grad", ["#fefae0", "#dda15e", "#606c38"])
cmap_ae = mcolors.LinearSegmentedColormap.from_list("ae_grad", ["#fefae0", "#bfe08a", "green"])
cmap_a3 = mcolors.LinearSegmentedColormap.from_list("a3_grad", ["#fefae0", "#dda15e", "#bc6c25"])

# ----------------------------------------------------------------------
# 1. Build the table for h = 0.1, 0.01, ..., down to tolerance 1e-6
# ----------------------------------------------------------------------
TOLERANCE = 1e-6

h_values = []
h = 0.1
while h >= TOLERANCE:
    h_values.append(round(h, 10))
    h /= 10
if not math.isclose(h_values[-1], TOLERANCE, rel_tol=1e-9):
    h_values.append(TOLERANCE)

bases = {
    "a = 2": 2.0,
    "a = 2.71828... (e)": math.e,
    "a = 3": 3.0,
}


def difference_quotient(a: float, h: float) -> float:
    """(a^h - 1) / h, computed via expm1/log for stability at small h."""
    return math.expm1(h * math.log(a)) / h


table = {name: [] for name in bases}
print(f"{'h':>12}" + "".join(f"{name:>22}" for name in bases))
print("-" * (12 + 22 * len(bases)))
for h in h_values:
    row = []
    for name, a in bases.items():
        val = difference_quotient(a, h)
        table[name].append(val)
        row.append(val)
    print(f"{h:>12.7f}" + "".join(f"{v:>22.6f}" for v in row))

print("-" * (12 + 22 * len(bases)))
print("settles at" + " " * 2 + "".join(f"{math.log(a):>22.6f}" for a in bases.values()))
print(f"\nTolerance = {TOLERANCE}")

# ----------------------------------------------------------------------
# 2. Draw the histogram (grouped gradient bar chart) in Matplotlib
# ----------------------------------------------------------------------
fig, ax = plt.subplots(figsize=(14, 7.5))
fig.patch.set_facecolor(PALETTE["cream"])
style_axes(ax, "Exercise 2 — (aʰ − 1)/h Settling to ln(a)   |   Tolerance = 1e-6", "(aʰ − 1) / h")

x = np.arange(len(h_values))
width = 0.24
cmaps = [cmap_a2, cmap_ae, cmap_a3]
line_colors = [PALETTE["olive"], PALETTE["green"], PALETTE["brown"]]

for i, (name, values) in enumerate(table.items()):
    offset = (i - 1) * width
    for xi, val in zip(x, values):
        gradient_bar(ax, xi + offset, val, width=width * 0.92, cmap=cmaps[i], shadow=False)

# dashed limit lines + legend proxies
from matplotlib.lines import Line2D
legend_handles = []
for i, (name, a) in enumerate(bases.items()):
    ax.axhline(math.log(a), color=line_colors[i], linestyle="--", linewidth=1.6, alpha=0.85, zorder=5)
    legend_handles.append(Line2D([0], [0], color=line_colors[i], lw=6,
                                  label=f"{name}  (ln a = {math.log(a):.4f})"))

ax.set_xticks(x)
ax.set_xticklabels([f"h = {h:g}" for h in h_values], fontsize=10)
ax.set_ylim(0, 1.25)
ax.set_xlim(-0.55, len(h_values) - 0.45)
ax.legend(handles=legend_handles, facecolor=PALETTE["white"], edgecolor=PALETTE["tan"],
          fontsize=9.5, loc="upper left", framealpha=0.95)

fig.tight_layout()
output_path = "exercise2_histogram.png"
fig.savefig(output_path, dpi=220, facecolor=fig.get_facecolor())
print(f"\nSaved histogram to {output_path}")
