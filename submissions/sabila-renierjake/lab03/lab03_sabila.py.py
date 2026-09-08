"""
Lab 03: Real-World Data Linear Regression
Numerical Methods, Section 3D

Dataset  : Philippines total population, 2005-2022
Source   : World Bank, World Development Indicators (series SP.POP.TOTL)
           https://databank.worldbank.org/source/world-development-indicators
x        : Year
y        : Total population (persons)

This script fits y = a0 + a1*x by least squares using the formulas
covered in class (no np.polyfit / no library regression call), then
reports Sr (SSE), r^2, the standard error of the estimate s(y/x), and
produces a fitted-line plot and a residual plot.
"""

import numpy as np
import matplotlib.pyplot as plt

# ---------------------------------------------------------------
# 1. Data  (x = year, y = population in persons)
# ---------------------------------------------------------------
year = np.array([2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013,
                  2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022])

population = np.array([88015962, 89508986, 91075184, 92699095, 94384250,
                        96337125, 98248614, 100175512, 102076336, 103767130,
                        105312992, 106735719, 108119693, 109465287, 110804683,
                        112081264, 113100950, 113964338])

n = len(year)

# ---------------------------------------------------------------
# 2. Least-squares formulas:
#    a1 = (n*Sum(xy) - Sum(x)*Sum(y)) / (n*Sum(x^2) - (Sum(x))^2)
#    a0 = ybar - a1*xbar
# ---------------------------------------------------------------
sum_x = np.sum(year)
sum_y = np.sum(population)
sum_xy = np.sum(year * population)
sum_x2 = np.sum(year ** 2)

a1 = (n * sum_xy - sum_x * sum_y) / (n * sum_x2 - sum_x ** 2)
a0 = (sum_y / n) - a1 * (sum_x / n)

print("Regression equation: y = {:.1f} + {:.1f} * x".format(a0, a1))
print("Slope     a1 = {:.4f} persons/year".format(a1))
print("Intercept a0 = {:.4f} persons".format(a0))

# ---------------------------------------------------------------
# 3. Goodness of fit: Sr (SSE), St, r^2, standard error s(y/x)
# ---------------------------------------------------------------
y_pred = a0 + a1 * year
y_bar = np.mean(population)

St = np.sum((population - y_bar) ** 2)          # total sum of squares
Sr = np.sum((population - y_pred) ** 2)         # sum of squares of residuals (SSE)
r2 = (St - Sr) / St
syx = np.sqrt(Sr / (n - 2))                     # standard error of the estimate

print("St  (total sum of squares)      = {:.4e}".format(St))
print("Sr  (SSE, sum of squared resid.)= {:.4e}".format(Sr))
print("r^2                              = {:.6f}".format(r2))
print("s(y/x) (standard error)          = {:.2f} persons".format(syx))

# ---------------------------------------------------------------
# 4. Prediction for an x not in the dataset
# ---------------------------------------------------------------
x_new = 2027
y_new = a0 + a1 * x_new
print("\nPrediction: at x = {}, y = {:.0f} persons".format(x_new, y_new))

# ---------------------------------------------------------------
# 5. Plots (green / white color theme)
# ---------------------------------------------------------------
GREEN_DARK = "#1b5e20"
GREEN_MED  = "#43a047"
GREEN_SOFT = "#a5d6a7"
GREEN_BG   = "#f1f8f2"

plt.rcParams.update({
    "font.size": 11,
    "axes.edgecolor": GREEN_DARK,
    "axes.labelcolor": GREEN_DARK,
    "xtick.color": GREEN_DARK,
    "ytick.color": GREEN_DARK,
    "text.color": GREEN_DARK,
})

# --- Fitted line plot ---
fig, ax = plt.subplots(figsize=(7, 4.5), facecolor="white")
ax.set_facecolor(GREEN_BG)
ax.scatter(year, population / 1e6, color=GREEN_DARK, s=45, zorder=3,
           label="Observed data")
x_line = np.linspace(year.min(), x_new, 100)
ax.plot(x_line, (a0 + a1 * x_line) / 1e6, color=GREEN_MED, linewidth=2.2,
        zorder=2, label="Fitted line: y = a0 + a1x")
ax.scatter([x_new], [y_new / 1e6], color="white", edgecolor=GREEN_DARK,
           s=70, zorder=4, marker="D", label="Prediction (2027)")
ax.set_xlabel("Year")
ax.set_ylabel("Population (millions)")
ax.set_title("Philippines Population vs. Year — Least-Squares Fit",
              color=GREEN_DARK, fontweight="bold")
ax.grid(True, color=GREEN_SOFT, linewidth=0.6)
ax.legend(facecolor="white", edgecolor=GREEN_SOFT)
fig.tight_layout()
fig.savefig("fit_plot.png", dpi=200, facecolor="white")
plt.close(fig)

# --- Residual plot ---
residuals = population - y_pred
fig, ax = plt.subplots(figsize=(7, 4), facecolor="white")
ax.set_facecolor(GREEN_BG)
ax.axhline(0, color=GREEN_DARK, linewidth=1.4)
ax.scatter(year, residuals, color=GREEN_MED, s=45, edgecolor=GREEN_DARK,
           zorder=3)
ax.set_xlabel("Year")
ax.set_ylabel("Residual (persons)")
ax.set_title("Residual Plot", color=GREEN_DARK, fontweight="bold")
ax.grid(True, color=GREEN_SOFT, linewidth=0.6)
fig.tight_layout()
fig.savefig("residual_plot.png", dpi=200, facecolor="white")
plt.close(fig)

print("\nSaved fit_plot.png and residual_plot.png")
