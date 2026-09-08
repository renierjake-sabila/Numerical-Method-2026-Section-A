"""
lab02_golloy.py
Numerical Methods — Laboratory Activity 02

This script reproduces the calculations for the supplied reservoir-stage log
and writes lab02_golloy.html.  The raw readings are embedded unchanged from
Time_Series_Reservoir_Stage.xlsx so the four-file submission remains runnable
without adding the workbook to the submission folder.
"""

from pathlib import Path
import json
import numpy as np
from scipy.optimize import curve_fit, minimize_scalar
from scipy import stats, integrate

# The supplied workbook contains 288 readings at 15-minute intervals.
# They are copied here unchanged so this submission can run as a four-file set.
RAW_DATA = [
  {
    "timestamp": "2026-07-21 00:00:00",
    "depth_m": 14.18
  },
  {
    "timestamp": "2026-07-21 00:15:00",
    "depth_m": 14.21
  },
  {
    "timestamp": "2026-07-21 00:30:00",
    "depth_m": 14.22
  },
  {
    "timestamp": "2026-07-21 00:45:00",
    "depth_m": 14.21
  },
  {
    "timestamp": "2026-07-21 01:00:00",
    "depth_m": 14.21
  },
  {
    "timestamp": "2026-07-21 01:15:00",
    "depth_m": 14.2
  },
  {
    "timestamp": "2026-07-21 01:30:00",
    "depth_m": 14.24
  },
  {
    "timestamp": "2026-07-21 01:45:00",
    "depth_m": 14.25
  },
  {
    "timestamp": "2026-07-21 02:00:00",
    "depth_m": 14.25
  },
  {
    "timestamp": "2026-07-21 02:15:00",
    "depth_m": 14.23
  },
  {
    "timestamp": "2026-07-21 02:30:00",
    "depth_m": 14.23
  },
  {
    "timestamp": "2026-07-21 02:45:00",
    "depth_m": 14.23
  },
  {
    "timestamp": "2026-07-21 03:00:00",
    "depth_m": 14.25
  },
  {
    "timestamp": "2026-07-21 03:15:00",
    "depth_m": 14.24
  },
  {
    "timestamp": "2026-07-21 03:30:00",
    "depth_m": 14.25
  },
  {
    "timestamp": "2026-07-21 03:45:00",
    "depth_m": 14.24
  },
  {
    "timestamp": "2026-07-21 04:00:00",
    "depth_m": 14.25
  },
  {
    "timestamp": "2026-07-21 04:15:00",
    "depth_m": 14.24
  },
  {
    "timestamp": "2026-07-21 04:30:00",
    "depth_m": 14.24
  },
  {
    "timestamp": "2026-07-21 04:45:00",
    "depth_m": 14.26
  },
  {
    "timestamp": "2026-07-21 05:00:00",
    "depth_m": 14.27
  },
  {
    "timestamp": "2026-07-21 05:15:00",
    "depth_m": 14.3
  },
  {
    "timestamp": "2026-07-21 05:30:00",
    "depth_m": 14.29
  },
  {
    "timestamp": "2026-07-21 05:45:00",
    "depth_m": 14.29
  },
  {
    "timestamp": "2026-07-21 06:00:00",
    "depth_m": 14.26
  },
  {
    "timestamp": "2026-07-21 06:15:00",
    "depth_m": 14.28
  },
  {
    "timestamp": "2026-07-21 06:30:00",
    "depth_m": 14.3
  },
  {
    "timestamp": "2026-07-21 06:45:00",
    "depth_m": 14.3
  },
  {
    "timestamp": "2026-07-21 07:00:00",
    "depth_m": 14.31
  },
  {
    "timestamp": "2026-07-21 07:15:00",
    "depth_m": 14.32
  },
  {
    "timestamp": "2026-07-21 07:30:00",
    "depth_m": 14.32
  },
  {
    "timestamp": "2026-07-21 07:45:00",
    "depth_m": 14.31
  },
  {
    "timestamp": "2026-07-21 08:00:00",
    "depth_m": 14.31
  },
  {
    "timestamp": "2026-07-21 08:15:00",
    "depth_m": 14.29
  },
  {
    "timestamp": "2026-07-21 08:30:00",
    "depth_m": 14.3
  },
  {
    "timestamp": "2026-07-21 08:45:00",
    "depth_m": 14.32
  },
  {
    "timestamp": "2026-07-21 09:00:00",
    "depth_m": 14.33
  },
  {
    "timestamp": "2026-07-21 09:15:00",
    "depth_m": 14.33
  },
  {
    "timestamp": "2026-07-21 09:30:00",
    "depth_m": 14.32
  },
  {
    "timestamp": "2026-07-21 09:45:00",
    "depth_m": 14.3
  },
  {
    "timestamp": "2026-07-21 10:00:00",
    "depth_m": 14.32
  },
  {
    "timestamp": "2026-07-21 10:15:00",
    "depth_m": 14.34
  },
  {
    "timestamp": "2026-07-21 10:30:00",
    "depth_m": 14.33
  },
  {
    "timestamp": "2026-07-21 10:45:00",
    "depth_m": 14.31
  },
  {
    "timestamp": "2026-07-21 11:00:00",
    "depth_m": 14.33
  },
  {
    "timestamp": "2026-07-21 11:15:00",
    "depth_m": 14.34
  },
  {
    "timestamp": "2026-07-21 11:30:00",
    "depth_m": 14.34
  },
  {
    "timestamp": "2026-07-21 11:45:00",
    "depth_m": 14.35
  },
  {
    "timestamp": "2026-07-21 12:00:00",
    "depth_m": 14.34
  },
  {
    "timestamp": "2026-07-21 12:15:00",
    "depth_m": 14.38
  },
  {
    "timestamp": "2026-07-21 12:30:00",
    "depth_m": 14.38
  },
  {
    "timestamp": "2026-07-21 12:45:00",
    "depth_m": 14.39
  },
  {
    "timestamp": "2026-07-21 13:00:00",
    "depth_m": 14.39
  },
  {
    "timestamp": "2026-07-21 13:15:00",
    "depth_m": 14.39
  },
  {
    "timestamp": "2026-07-21 13:30:00",
    "depth_m": 14.37
  },
  {
    "timestamp": "2026-07-21 13:45:00",
    "depth_m": 14.38
  },
  {
    "timestamp": "2026-07-21 14:00:00",
    "depth_m": 14.39
  },
  {
    "timestamp": "2026-07-21 14:15:00",
    "depth_m": 14.39
  },
  {
    "timestamp": "2026-07-21 14:30:00",
    "depth_m": 14.39
  },
  {
    "timestamp": "2026-07-21 14:45:00",
    "depth_m": 14.41
  },
  {
    "timestamp": "2026-07-21 15:00:00",
    "depth_m": 14.43
  },
  {
    "timestamp": "2026-07-21 15:15:00",
    "depth_m": 14.41
  },
  {
    "timestamp": "2026-07-21 15:30:00",
    "depth_m": 14.4
  },
  {
    "timestamp": "2026-07-21 15:45:00",
    "depth_m": 14.43
  },
  {
    "timestamp": "2026-07-21 16:00:00",
    "depth_m": 14.42
  },
  {
    "timestamp": "2026-07-21 16:15:00",
    "depth_m": 14.44
  },
  {
    "timestamp": "2026-07-21 16:30:00",
    "depth_m": 14.43
  },
  {
    "timestamp": "2026-07-21 16:45:00",
    "depth_m": 14.42
  },
  {
    "timestamp": "2026-07-21 17:00:00",
    "depth_m": 14.43
  },
  {
    "timestamp": "2026-07-21 17:15:00",
    "depth_m": 14.45
  },
  {
    "timestamp": "2026-07-21 17:30:00",
    "depth_m": 14.45
  },
  {
    "timestamp": "2026-07-21 17:45:00",
    "depth_m": 14.44
  },
  {
    "timestamp": "2026-07-21 18:00:00",
    "depth_m": 14.42
  },
  {
    "timestamp": "2026-07-21 18:15:00",
    "depth_m": 14.43
  },
  {
    "timestamp": "2026-07-21 18:30:00",
    "depth_m": 14.46
  },
  {
    "timestamp": "2026-07-21 18:45:00",
    "depth_m": 14.48
  },
  {
    "timestamp": "2026-07-21 19:00:00",
    "depth_m": 14.47
  },
  {
    "timestamp": "2026-07-21 19:15:00",
    "depth_m": 14.46
  },
  {
    "timestamp": "2026-07-21 19:30:00",
    "depth_m": 14.48
  },
  {
    "timestamp": "2026-07-21 19:45:00",
    "depth_m": 14.46
  },
  {
    "timestamp": "2026-07-21 20:00:00",
    "depth_m": 14.48
  },
  {
    "timestamp": "2026-07-21 20:15:00",
    "depth_m": 14.48
  },
  {
    "timestamp": "2026-07-21 20:30:00",
    "depth_m": 14.49
  },
  {
    "timestamp": "2026-07-21 20:45:00",
    "depth_m": 14.47
  },
  {
    "timestamp": "2026-07-21 21:00:00",
    "depth_m": 14.47
  },
  {
    "timestamp": "2026-07-21 21:15:00",
    "depth_m": 14.47
  },
  {
    "timestamp": "2026-07-21 21:30:00",
    "depth_m": 14.49
  },
  {
    "timestamp": "2026-07-21 21:45:00",
    "depth_m": 14.48
  },
  {
    "timestamp": "2026-07-21 22:00:00",
    "depth_m": 14.48
  },
  {
    "timestamp": "2026-07-21 22:15:00",
    "depth_m": 14.48
  },
  {
    "timestamp": "2026-07-21 22:30:00",
    "depth_m": 14.5
  },
  {
    "timestamp": "2026-07-21 22:45:00",
    "depth_m": 14.51
  },
  {
    "timestamp": "2026-07-21 23:00:00",
    "depth_m": 14.49
  },
  {
    "timestamp": "2026-07-21 23:15:00",
    "depth_m": 14.5
  },
  {
    "timestamp": "2026-07-21 23:30:00",
    "depth_m": 14.49
  },
  {
    "timestamp": "2026-07-21 23:45:00",
    "depth_m": 14.5
  },
  {
    "timestamp": "2026-07-22 00:00:00",
    "depth_m": 14.51
  },
  {
    "timestamp": "2026-07-22 00:15:00",
    "depth_m": 14.51
  },
  {
    "timestamp": "2026-07-22 00:30:00",
    "depth_m": 14.53
  },
  {
    "timestamp": "2026-07-22 00:45:00",
    "depth_m": 14.54
  },
  {
    "timestamp": "2026-07-22 01:00:00",
    "depth_m": 14.54
  },
  {
    "timestamp": "2026-07-22 01:15:00",
    "depth_m": 14.58
  },
  {
    "timestamp": "2026-07-22 01:30:00",
    "depth_m": 14.57
  },
  {
    "timestamp": "2026-07-22 01:45:00",
    "depth_m": 14.59
  },
  {
    "timestamp": "2026-07-22 02:00:00",
    "depth_m": 14.61
  },
  {
    "timestamp": "2026-07-22 02:15:00",
    "depth_m": 14.64
  },
  {
    "timestamp": "2026-07-22 02:30:00",
    "depth_m": 14.7
  },
  {
    "timestamp": "2026-07-22 02:45:00",
    "depth_m": 14.74
  },
  {
    "timestamp": "2026-07-22 03:00:00",
    "depth_m": 14.82
  },
  {
    "timestamp": "2026-07-22 03:15:00",
    "depth_m": 14.9
  },
  {
    "timestamp": "2026-07-22 03:30:00",
    "depth_m": 14.99
  },
  {
    "timestamp": "2026-07-22 03:45:00",
    "depth_m": 15.11
  },
  {
    "timestamp": "2026-07-22 04:00:00",
    "depth_m": 15.23
  },
  {
    "timestamp": "2026-07-22 04:15:00",
    "depth_m": 15.4
  },
  {
    "timestamp": "2026-07-22 04:30:00",
    "depth_m": 15.6
  },
  {
    "timestamp": "2026-07-22 04:45:00",
    "depth_m": 15.8
  },
  {
    "timestamp": "2026-07-22 05:00:00",
    "depth_m": 16.02
  },
  {
    "timestamp": "2026-07-22 05:15:00",
    "depth_m": 16.25
  },
  {
    "timestamp": "2026-07-22 05:30:00",
    "depth_m": 16.5
  },
  {
    "timestamp": "2026-07-22 05:45:00",
    "depth_m": 16.73
  },
  {
    "timestamp": "2026-07-22 06:00:00",
    "depth_m": 16.96
  },
  {
    "timestamp": "2026-07-22 06:15:00",
    "depth_m": 17.21
  },
  {
    "timestamp": "2026-07-22 06:30:00",
    "depth_m": 17.42
  },
  {
    "timestamp": "2026-07-22 06:45:00",
    "depth_m": 17.63
  },
  {
    "timestamp": "2026-07-22 07:00:00",
    "depth_m": 17.84
  },
  {
    "timestamp": "2026-07-22 07:15:00",
    "depth_m": 18.01
  },
  {
    "timestamp": "2026-07-22 07:30:00",
    "depth_m": 18.21
  },
  {
    "timestamp": "2026-07-22 07:45:00",
    "depth_m": 18.39
  },
  {
    "timestamp": "2026-07-22 08:00:00",
    "depth_m": 18.6
  },
  {
    "timestamp": "2026-07-22 08:15:00",
    "depth_m": 18.79
  },
  {
    "timestamp": "2026-07-22 08:30:00",
    "depth_m": 19.0
  },
  {
    "timestamp": "2026-07-22 08:45:00",
    "depth_m": 19.21
  },
  {
    "timestamp": "2026-07-22 09:00:00",
    "depth_m": 19.4
  },
  {
    "timestamp": "2026-07-22 09:15:00",
    "depth_m": 19.64
  },
  {
    "timestamp": "2026-07-22 09:30:00",
    "depth_m": 19.86
  },
  {
    "timestamp": "2026-07-22 09:45:00",
    "depth_m": 20.08
  },
  {
    "timestamp": "2026-07-22 10:00:00",
    "depth_m": 20.29
  },
  {
    "timestamp": "2026-07-22 10:15:00",
    "depth_m": 20.46
  },
  {
    "timestamp": "2026-07-22 10:30:00",
    "depth_m": 20.62
  },
  {
    "timestamp": "2026-07-22 10:45:00",
    "depth_m": 20.78
  },
  {
    "timestamp": "2026-07-22 11:00:00",
    "depth_m": 20.91
  },
  {
    "timestamp": "2026-07-22 11:15:00",
    "depth_m": 21.01
  },
  {
    "timestamp": "2026-07-22 11:30:00",
    "depth_m": 21.06
  },
  {
    "timestamp": "2026-07-22 11:45:00",
    "depth_m": 21.1
  },
  {
    "timestamp": "2026-07-22 12:00:00",
    "depth_m": 21.13
  },
  {
    "timestamp": "2026-07-22 12:15:00",
    "depth_m": 21.14
  },
  {
    "timestamp": "2026-07-22 12:30:00",
    "depth_m": 21.12
  },
  {
    "timestamp": "2026-07-22 12:45:00",
    "depth_m": 21.09
  },
  {
    "timestamp": "2026-07-22 13:00:00",
    "depth_m": 21.08
  },
  {
    "timestamp": "2026-07-22 13:15:00",
    "depth_m": 21.05
  },
  {
    "timestamp": "2026-07-22 13:30:00",
    "depth_m": 21.0
  },
  {
    "timestamp": "2026-07-22 13:45:00",
    "depth_m": 20.96
  },
  {
    "timestamp": "2026-07-22 14:00:00",
    "depth_m": 20.92
  },
  {
    "timestamp": "2026-07-22 14:15:00",
    "depth_m": 20.88
  },
  {
    "timestamp": "2026-07-22 14:30:00",
    "depth_m": 20.84
  },
  {
    "timestamp": "2026-07-22 14:45:00",
    "depth_m": 20.81
  },
  {
    "timestamp": "2026-07-22 15:00:00",
    "depth_m": 20.76
  },
  {
    "timestamp": "2026-07-22 15:15:00",
    "depth_m": 20.72
  },
  {
    "timestamp": "2026-07-22 15:30:00",
    "depth_m": 20.69
  },
  {
    "timestamp": "2026-07-22 15:45:00",
    "depth_m": 20.65
  },
  {
    "timestamp": "2026-07-22 16:00:00",
    "depth_m": 20.6
  },
  {
    "timestamp": "2026-07-22 16:15:00",
    "depth_m": 20.57
  },
  {
    "timestamp": "2026-07-22 16:30:00",
    "depth_m": 20.52
  },
  {
    "timestamp": "2026-07-22 16:45:00",
    "depth_m": 20.46
  },
  {
    "timestamp": "2026-07-22 17:00:00",
    "depth_m": 20.43
  },
  {
    "timestamp": "2026-07-22 17:15:00",
    "depth_m": 20.4
  },
  {
    "timestamp": "2026-07-22 17:30:00",
    "depth_m": 20.35
  },
  {
    "timestamp": "2026-07-22 17:45:00",
    "depth_m": 20.31
  },
  {
    "timestamp": "2026-07-22 18:00:00",
    "depth_m": 20.3
  },
  {
    "timestamp": "2026-07-22 18:15:00",
    "depth_m": 20.26
  },
  {
    "timestamp": "2026-07-22 18:30:00",
    "depth_m": 20.25
  },
  {
    "timestamp": "2026-07-22 18:45:00",
    "depth_m": 20.21
  },
  {
    "timestamp": "2026-07-22 19:00:00",
    "depth_m": 20.17
  },
  {
    "timestamp": "2026-07-22 19:15:00",
    "depth_m": 20.15
  },
  {
    "timestamp": "2026-07-22 19:30:00",
    "depth_m": 20.12
  },
  {
    "timestamp": "2026-07-22 19:45:00",
    "depth_m": 20.09
  },
  {
    "timestamp": "2026-07-22 20:00:00",
    "depth_m": 20.05
  },
  {
    "timestamp": "2026-07-22 20:15:00",
    "depth_m": 20.04
  },
  {
    "timestamp": "2026-07-22 20:30:00",
    "depth_m": 20.0
  },
  {
    "timestamp": "2026-07-22 20:45:00",
    "depth_m": 19.97
  },
  {
    "timestamp": "2026-07-22 21:00:00",
    "depth_m": 19.95
  },
  {
    "timestamp": "2026-07-22 21:15:00",
    "depth_m": 19.96
  },
  {
    "timestamp": "2026-07-22 21:30:00",
    "depth_m": 19.94
  },
  {
    "timestamp": "2026-07-22 21:45:00",
    "depth_m": 19.91
  },
  {
    "timestamp": "2026-07-22 22:00:00",
    "depth_m": 19.89
  },
  {
    "timestamp": "2026-07-22 22:15:00",
    "depth_m": 19.9
  },
  {
    "timestamp": "2026-07-22 22:30:00",
    "depth_m": 19.85
  },
  {
    "timestamp": "2026-07-22 22:45:00",
    "depth_m": 19.85
  },
  {
    "timestamp": "2026-07-22 23:00:00",
    "depth_m": 19.83
  },
  {
    "timestamp": "2026-07-22 23:15:00",
    "depth_m": 19.82
  },
  {
    "timestamp": "2026-07-22 23:30:00",
    "depth_m": 19.8
  },
  {
    "timestamp": "2026-07-22 23:45:00",
    "depth_m": 19.8
  },
  {
    "timestamp": "2026-07-23 00:00:00",
    "depth_m": 19.78
  },
  {
    "timestamp": "2026-07-23 00:15:00",
    "depth_m": 19.76
  },
  {
    "timestamp": "2026-07-23 00:30:00",
    "depth_m": 19.75
  },
  {
    "timestamp": "2026-07-23 00:45:00",
    "depth_m": 19.73
  },
  {
    "timestamp": "2026-07-23 01:00:00",
    "depth_m": 19.74
  },
  {
    "timestamp": "2026-07-23 01:15:00",
    "depth_m": 19.73
  },
  {
    "timestamp": "2026-07-23 01:30:00",
    "depth_m": 19.72
  },
  {
    "timestamp": "2026-07-23 01:45:00",
    "depth_m": 19.71
  },
  {
    "timestamp": "2026-07-23 02:00:00",
    "depth_m": 19.7
  },
  {
    "timestamp": "2026-07-23 02:15:00",
    "depth_m": 19.72
  },
  {
    "timestamp": "2026-07-23 02:30:00",
    "depth_m": 19.7
  },
  {
    "timestamp": "2026-07-23 02:45:00",
    "depth_m": 19.68
  },
  {
    "timestamp": "2026-07-23 03:00:00",
    "depth_m": 19.68
  },
  {
    "timestamp": "2026-07-23 03:15:00",
    "depth_m": 19.66
  },
  {
    "timestamp": "2026-07-23 03:30:00",
    "depth_m": 19.67
  },
  {
    "timestamp": "2026-07-23 03:45:00",
    "depth_m": 19.65
  },
  {
    "timestamp": "2026-07-23 04:00:00",
    "depth_m": 19.65
  },
  {
    "timestamp": "2026-07-23 04:15:00",
    "depth_m": 19.64
  },
  {
    "timestamp": "2026-07-23 04:30:00",
    "depth_m": 19.62
  },
  {
    "timestamp": "2026-07-23 04:45:00",
    "depth_m": 19.64
  },
  {
    "timestamp": "2026-07-23 05:00:00",
    "depth_m": 19.63
  },
  {
    "timestamp": "2026-07-23 05:15:00",
    "depth_m": 19.64
  },
  {
    "timestamp": "2026-07-23 05:30:00",
    "depth_m": 19.62
  },
  {
    "timestamp": "2026-07-23 05:45:00",
    "depth_m": 19.61
  },
  {
    "timestamp": "2026-07-23 06:00:00",
    "depth_m": 19.6
  },
  {
    "timestamp": "2026-07-23 06:15:00",
    "depth_m": 19.59
  },
  {
    "timestamp": "2026-07-23 06:30:00",
    "depth_m": 19.59
  },
  {
    "timestamp": "2026-07-23 06:45:00",
    "depth_m": 19.59
  },
  {
    "timestamp": "2026-07-23 07:00:00",
    "depth_m": 19.59
  },
  {
    "timestamp": "2026-07-23 07:15:00",
    "depth_m": 19.61
  },
  {
    "timestamp": "2026-07-23 07:30:00",
    "depth_m": 19.62
  },
  {
    "timestamp": "2026-07-23 07:45:00",
    "depth_m": 19.6
  },
  {
    "timestamp": "2026-07-23 08:00:00",
    "depth_m": 19.59
  },
  {
    "timestamp": "2026-07-23 08:15:00",
    "depth_m": 19.59
  },
  {
    "timestamp": "2026-07-23 08:30:00",
    "depth_m": 19.58
  },
  {
    "timestamp": "2026-07-23 08:45:00",
    "depth_m": 19.54
  },
  {
    "timestamp": "2026-07-23 09:00:00",
    "depth_m": 19.54
  },
  {
    "timestamp": "2026-07-23 09:15:00",
    "depth_m": 19.57
  },
  {
    "timestamp": "2026-07-23 09:30:00",
    "depth_m": 19.55
  },
  {
    "timestamp": "2026-07-23 09:45:00",
    "depth_m": 19.55
  },
  {
    "timestamp": "2026-07-23 10:00:00",
    "depth_m": 19.55
  },
  {
    "timestamp": "2026-07-23 10:15:00",
    "depth_m": 19.53
  },
  {
    "timestamp": "2026-07-23 10:30:00",
    "depth_m": 19.53
  },
  {
    "timestamp": "2026-07-23 10:45:00",
    "depth_m": 19.56
  },
  {
    "timestamp": "2026-07-23 11:00:00",
    "depth_m": 19.56
  },
  {
    "timestamp": "2026-07-23 11:15:00",
    "depth_m": 19.55
  },
  {
    "timestamp": "2026-07-23 11:30:00",
    "depth_m": 19.53
  },
  {
    "timestamp": "2026-07-23 11:45:00",
    "depth_m": 19.53
  },
  {
    "timestamp": "2026-07-23 12:00:00",
    "depth_m": 19.52
  },
  {
    "timestamp": "2026-07-23 12:15:00",
    "depth_m": 19.52
  },
  {
    "timestamp": "2026-07-23 12:30:00",
    "depth_m": 19.51
  },
  {
    "timestamp": "2026-07-23 12:45:00",
    "depth_m": 19.52
  },
  {
    "timestamp": "2026-07-23 13:00:00",
    "depth_m": 19.52
  },
  {
    "timestamp": "2026-07-23 13:15:00",
    "depth_m": 19.53
  },
  {
    "timestamp": "2026-07-23 13:30:00",
    "depth_m": 19.53
  },
  {
    "timestamp": "2026-07-23 13:45:00",
    "depth_m": 19.51
  },
  {
    "timestamp": "2026-07-23 14:00:00",
    "depth_m": 19.53
  },
  {
    "timestamp": "2026-07-23 14:15:00",
    "depth_m": 19.53
  },
  {
    "timestamp": "2026-07-23 14:30:00",
    "depth_m": 19.53
  },
  {
    "timestamp": "2026-07-23 14:45:00",
    "depth_m": 19.53
  },
  {
    "timestamp": "2026-07-23 15:00:00",
    "depth_m": 19.53
  },
  {
    "timestamp": "2026-07-23 15:15:00",
    "depth_m": 19.54
  },
  {
    "timestamp": "2026-07-23 15:30:00",
    "depth_m": 19.51
  },
  {
    "timestamp": "2026-07-23 15:45:00",
    "depth_m": 19.52
  },
  {
    "timestamp": "2026-07-23 16:00:00",
    "depth_m": 19.51
  },
  {
    "timestamp": "2026-07-23 16:15:00",
    "depth_m": 19.49
  },
  {
    "timestamp": "2026-07-23 16:30:00",
    "depth_m": 19.5
  },
  {
    "timestamp": "2026-07-23 16:45:00",
    "depth_m": 19.49
  },
  {
    "timestamp": "2026-07-23 17:00:00",
    "depth_m": 19.5
  },
  {
    "timestamp": "2026-07-23 17:15:00",
    "depth_m": 19.52
  },
  {
    "timestamp": "2026-07-23 17:30:00",
    "depth_m": 19.54
  },
  {
    "timestamp": "2026-07-23 17:45:00",
    "depth_m": 19.53
  },
  {
    "timestamp": "2026-07-23 18:00:00",
    "depth_m": 19.52
  },
  {
    "timestamp": "2026-07-23 18:15:00",
    "depth_m": 19.54
  },
  {
    "timestamp": "2026-07-23 18:30:00",
    "depth_m": 19.53
  },
  {
    "timestamp": "2026-07-23 18:45:00",
    "depth_m": 19.52
  },
  {
    "timestamp": "2026-07-23 19:00:00",
    "depth_m": 19.5
  },
  {
    "timestamp": "2026-07-23 19:15:00",
    "depth_m": 19.53
  },
  {
    "timestamp": "2026-07-23 19:30:00",
    "depth_m": 19.51
  },
  {
    "timestamp": "2026-07-23 19:45:00",
    "depth_m": 19.5
  },
  {
    "timestamp": "2026-07-23 20:00:00",
    "depth_m": 19.52
  },
  {
    "timestamp": "2026-07-23 20:15:00",
    "depth_m": 19.52
  },
  {
    "timestamp": "2026-07-23 20:30:00",
    "depth_m": 19.51
  },
  {
    "timestamp": "2026-07-23 20:45:00",
    "depth_m": 19.5
  },
  {
    "timestamp": "2026-07-23 21:00:00",
    "depth_m": 19.5
  },
  {
    "timestamp": "2026-07-23 21:15:00",
    "depth_m": 19.52
  },
  {
    "timestamp": "2026-07-23 21:30:00",
    "depth_m": 19.5
  },
  {
    "timestamp": "2026-07-23 21:45:00",
    "depth_m": 19.51
  },
  {
    "timestamp": "2026-07-23 22:00:00",
    "depth_m": 19.52
  },
  {
    "timestamp": "2026-07-23 22:15:00",
    "depth_m": 19.5
  },
  {
    "timestamp": "2026-07-23 22:30:00",
    "depth_m": 19.51
  },
  {
    "timestamp": "2026-07-23 22:45:00",
    "depth_m": 19.5
  },
  {
    "timestamp": "2026-07-23 23:00:00",
    "depth_m": 19.49
  },
  {
    "timestamp": "2026-07-23 23:15:00",
    "depth_m": 19.51
  },
  {
    "timestamp": "2026-07-23 23:30:00",
    "depth_m": 19.52
  },
  {
    "timestamp": "2026-07-23 23:45:00",
    "depth_m": 19.51
  }
]

def double_logistic(t, c, a1, k1, t01, a2, k2, t02):
    """Continuous model: a rising logistic plus a logistic recession."""
    z1 = np.clip(-k1 * (t - t01), -700, 700)
    z2 = np.clip(-k2 * (t - t02), -700, 700)
    return c + a1 / (1 + np.exp(z1)) + a2 / (1 + np.exp(z2))

def first_second_differences(h, dt):
    """Forward/backward differences at the ends; central differences inside."""
    first = np.empty_like(h)
    second = np.empty_like(h)

    first[0] = (h[1] - h[0]) / dt
    first[-1] = (h[-1] - h[-2]) / dt
    first[1:-1] = (h[2:] - h[:-2]) / (2 * dt)

    second[0] = (h[2] - 2*h[1] + h[0]) / dt**2
    second[-1] = (h[-1] - 2*h[-2] + h[-3]) / dt**2
    second[1:-1] = (h[2:] - 2*h[1:-1] + h[:-2]) / dt**2
    return first, second

def make_path(x, y, x0, x1, y0, y1, left, top, width, height):
    """Convert already-computed data to SVG coordinates for display only."""
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    sx = left + (x - x0) / (x1 - x0) * width
    sy = top + (y1 - y) / (y1 - y0) * height
    return " ".join(f"{xx:.1f},{yy:.1f}" for xx, yy in zip(sx, sy))

def build_svg(x, series, labels, zero_line=None, fill_series=None, points=False):
    """Build a static SVG. No fitting/statistics/arithmetic are performed in HTML."""
    x = np.asarray(x, dtype=float)
    ys = [np.asarray(v, dtype=float) for v in series]
    W, H = 980, 330
    left, right, top, bottom = 70, 20, 20, 48
    pw, ph = W-left-right, H-top-bottom
    xmin, xmax = float(x.min()), float(x.max())
    ymin = min(float(v.min()) for v in ys)
    ymax = max(float(v.max()) for v in ys)
    if zero_line is not None:
        ymin, ymax = min(ymin, zero_line), max(ymax, zero_line)
    pad = max(ymax-ymin, 1e-9) * 0.08
    ymin -= pad
    ymax += pad

    def X(v):
        return left + (v-xmin)/(xmax-xmin)*pw
    def Y(v):
        return top + (ymax-v)/(ymax-ymin)*ph

    out = [f'<svg viewBox="0 0 {W} {H}" role="img">']
    for frac in np.linspace(0, 1, 5):
        yy = top + frac*ph
        val = ymax - frac*(ymax-ymin)
        out.append(f'<line x1="{left}" y1="{yy:.1f}" x2="{W-right}" y2="{yy:.1f}" class="grid"/>')
        out.append(f'<text x="{left-8}" y="{yy+4:.1f}" text-anchor="end" class="axis">{val:.2f}</text>')
    for frac in np.linspace(0, 1, 7):
        xx = left + frac*pw
        val = xmin + frac*(xmax-xmin)
        out.append(f'<line x1="{xx:.1f}" y1="{top}" x2="{xx:.1f}" y2="{H-bottom}" class="grid"/>')
        out.append(f'<text x="{xx:.1f}" y="{H-bottom+20}" text-anchor="middle" class="axis">{val:.0f}</text>')
    if zero_line is not None:
        yy = Y(zero_line)
        out.append(f'<line x1="{left}" y1="{yy:.1f}" x2="{W-right}" y2="{yy:.1f}" class="zero"/>')
    if fill_series is not None:
        pts = [f"{X(x[0]):.1f},{Y(0):.1f}"]
        pts += [f"{X(xx):.1f},{Y(yy):.1f}" for xx, yy in zip(x, fill_series)]
        pts += [f"{X(x[-1]):.1f},{Y(0):.1f}"]
        out.append('<polygon points="' + " ".join(pts) + '" class="area"/>')
    for j, values in enumerate(ys):
        out.append(f'<polyline points="{make_path(x, values, xmin, xmax, ymin, ymax, left, top, pw, ph)}" class="line l{j}"/>')
        if points:
            for xx, yy in zip(x, values):
                out.append(f'<circle cx="{X(xx):.1f}" cy="{Y(yy):.1f}" r="1.8" class="pt p{j}"/>')
    out.append(f'<line x1="{left}" y1="{top}" x2="{left}" y2="{H-bottom}" class="axisline"/>')
    out.append(f'<line x1="{left}" y1="{H-bottom}" x2="{W-right}" y2="{H-bottom}" class="axisline"/>')
    lx = left
    for j, label in enumerate(labels):
        out.append(f'<line x1="{lx}" y1="12" x2="{lx+28}" y2="12" class="line l{j}"/>')
        out.append(f'<text x="{lx+34}" y="16" class="legend">{label}</text>')
        lx += 210
    out.append('</svg>')
    return "\n".join(out)

def main():
    out_file = Path(__file__).with_name("lab02_golloy.html")

    timestamps = np.array([r["timestamp"] for r in RAW_DATA], dtype="datetime64[s]")
    h = np.array([r["depth_m"] for r in RAW_DATA], dtype=float)

    # Time origin: first reading. Each logger step is 0.25 hour.
    t = (timestamps - timestamps[0]).astype("timedelta64[s]").astype(float) / 3600.0
    dt = 0.25
    n = len(h)

    # Group 1: finite differences.
    first, second = first_second_differences(h, dt)

    # Group 2: Levenberg-Marquardt fit.
    # p0 was chosen from the raw plot: baseline ~14 m, rapid rise ~8 m,
    # steepest rise near 31 h, and a later recession beginning near 38 h.
    p0 = [14.0, 8.0, 0.5, 31.0, -4.0, 0.3, 38.0]
    popt, pcov = curve_fit(
        double_logistic, t, h, p0=p0, method="lm", maxfev=20000
    )

    fitted = double_logistic(t, *popt)
    residuals = h - fitted

    # Statistics, in the order specified by the activity.
    p = len(popt)
    dof = n - p
    sse = np.sum(residuals**2)
    sst = np.sum((h - h.mean())**2)
    r2 = 1 - sse/sst
    s = np.sqrt(sse/dof)
    se = np.sqrt(np.diag(pcov))
    t_stats = popt/se
    p_values = 2 * stats.t.sf(np.abs(t_stats), dof)

    # Fitted derivative and its maximum.
    grid = np.linspace(t[0], t[-1], 1201)
    fitted_grid = double_logistic(grid, *popt)
    c, a1, k1, t01, a2, k2, t02 = popt
    e1 = np.exp(np.clip(-k1*(grid-t01), -700, 700))
    e2 = np.exp(np.clip(-k2*(grid-t02), -700, 700))
    derivative = a1*k1*e1/(1+e1)**2 + a2*k2*e2/(1+e2)**2
    second_model = (
        a1*k1**2*e1*(e1-1)/(1+e1)**3
        + a2*k2**2*e2*(e2-1)/(1+e2)**3
    )

    optimum = minimize_scalar(
        lambda x: -(
            a1*k1*np.exp(np.clip(-k1*(x-t01), -700, 700))
            /(1+np.exp(np.clip(-k1*(x-t01), -700, 700)))**2
            + a2*k2*np.exp(np.clip(-k2*(x-t02), -700, 700))
            /(1+np.exp(np.clip(-k2*(x-t02), -700, 700)))**2
        ),
        bounds=(t[0], t[-1]), method="bounded"
    )
    t_max = optimum.x
    max_dhdt = -optimum.fun

    # Group 3: area under fitted curve and raw trapezoid cross-check.
    area, quad_error = integrate.quad(
        lambda x: float(double_logistic(np.array([x]), *popt)[0]),
        float(t[0]), float(t[-1])
    )
    trapezoid = np.trapezoid(h, t)

    # Residual evidence.
    max_abs_resid = np.max(np.abs(residuals))
    max_resid_index = int(np.argmax(np.abs(residuals)))

    # Static SVG charts are generated by Python; the HTML only displays them.
    raw_svg = build_svg(t, [h], ["Raw stage h(t)"], points=True)
    fit_svg = build_svg(t, [h, fitted], ["Raw stage", "Double-logistic fit"], points=True)
    residual_svg = build_svg(t, [residuals], ["Residual e_i"], zero_line=0, points=True)
    derivative_svg = build_svg(t, [first, second], ["First derivative dh/dt", "Second derivative d²h/dt²"])
    area_svg = build_svg(grid, [fitted_grid], ["Fitted h(t)"], fill_series=fitted_grid)

    # The complete dashboard is assembled here. JavaScript only switches tabs.
    html = build_dashboard_html(
        timestamps, h, t, raw_svg, fit_svg, residual_svg, derivative_svg, area_svg,
        popt, se, t_stats, p_values, n, p, dof, sse, sst, r2, s,
        t_max, max_dhdt, area, quad_error, trapezoid, max_abs_resid,
        max_resid_index, residuals
    )
    out_file.write_text(html, encoding="utf-8")
    print(f"Dashboard written to {out_file}")

def build_dashboard_html(
    timestamps, h, t, raw_svg, fit_svg, residual_svg, derivative_svg, area_svg,
    popt, se, t_stats, p_values, n, p, dof, sse, sst, r2, s,
    t_max, max_dhdt, area, quad_error, trapezoid, max_abs_resid,
    max_resid_index, residuals
):
    # This function contains display text only; numerical results are passed in.
    # (The generated dashboard itself contains no fitting/regression/statistics code.)
    start = str(timestamps[0]).replace("T", " ")[:16]
    end = str(timestamps[-1]).replace("T", " ")[:16]
    start_date = timestamps[0]
    max_time = start_date + np.timedelta64(int(round(t_max*3600)), "s")

    segment_means = [float(residuals[i:i+96].mean()) for i in range(0, n, 96)]

    # Longest residual sign run.
    signs = np.sign(residuals)
    longest = 1
    current = 1
    for i in range(1, n):
        if signs[i] == signs[i-1]:
            current += 1
            longest = max(longest, current)
        else:
            current = 1

    names = ["c", "a₁", "k₁", "t₀₁", "a₂", "k₂", "t₀₂"]
    rows = ""
    for name, value, err, ts, pv in zip(names, popt, se, t_stats, p_values):
        p_text = "< 1×10⁻¹⁵" if pv < 1e-15 else f"{pv:.3g}"
        rows += (
            f"<tr><td>{name}</td><td>{value:.4f}</td><td>{err:.4f}</td>"
            f"<td>{ts:.3f}</td><td>{p_text}</td></tr>"
        )

    model_equation = (
        "h(t) = c + a₁/(1 + exp(-k₁(t − t₀₁))) "
        "+ a₂/(1 + exp(-k₂(t − t₀₂)))"
    )
    residual_text = (
        f"Residual means for the three 96-reading blocks are "
        f"{segment_means[0]:.4f}, {segment_means[1]:.4f}, and "
        f"{segment_means[2]:.4f} m. The spread does not show a clear "
        f"increase with stage, but the longest same-sign run is {longest} points, "
        f"so the very high R² is not by itself evidence of a perfect shape match. "
        f"The largest absolute residual is {max_abs_resid:.4f} m, which is "
        f"{max_abs_resid-0.01:.4f} m above the logger's 0.01 m resolution."
    )

    html = """<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Lab 02 — Fitting a Curve to the Dam</title>
<style>
body{font-family:system-ui,-apple-system,Segoe UI,sans-serif;margin:0;background:#f5f6f8;color:#17202a}
header{padding:28px 36px;background:#18212b;color:white} h1{margin:0 0 6px;font-size:28px} header p{margin:4px 0;opacity:.88}
main{max-width:1150px;margin:22px auto;padding:0 18px} .panel{background:white;border:1px solid #dfe4ea;border-radius:14px;padding:20px;margin-bottom:18px;box-shadow:0 2px 8px rgba(0,0,0,.04)}
.tabs{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px} button{border:1px solid #cbd3dc;background:#fff;border-radius:9px;padding:9px 14px;cursor:pointer} button.active{background:#18212b;color:#fff}
.tab{display:none} .tab.active{display:block} .note{background:#fff7df;border-left:4px solid #c99700;padding:12px 14px;border-radius:7px;margin-bottom:18px}
.grid{stroke:#e7ebef;stroke-width:1} .axisline{stroke:#5d6874;stroke-width:1.2} .zero{stroke:#a04b4b;stroke-width:1.2;stroke-dasharray:6 5}
.line{fill:none;stroke-width:2.5} .l0{stroke:#2f5d8c} .l1{stroke:#b05b3b} .l2{stroke:#4b7a52} .pt{opacity:.55} .p0{fill:#2f5d8c} .p1{fill:#b05b3b} .area{fill:#8aa6c1;opacity:.25;stroke:none}
.axis,.legend{font-size:11px;fill:#4d5965} .legend{font-size:12px} table{border-collapse:collapse;width:100%} th,td{border-bottom:1px solid #e6e9ed;padding:9px;text-align:right} th:first-child,td:first-child{text-align:left} .metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px} .metric{padding:12px;background:#f7f8fa;border-radius:10px} .metric b{display:block;font-size:20px;margin-top:4px} .small{font-size:13px;color:#5b6670} code{background:#f0f2f4;padding:2px 5px;border-radius:4px}
</style></head>
<body>
<header><h1>Numerical Methods — Laboratory Activity 02</h1>
<p>Fitting a curve to the dam: Levenberg–Marquardt, residuals, and area under the level</p>
<p>Reservoir stage log • @@P0@@ readings • 15-minute sampling • @@P1@@ to @@P2@@</p></header>
<main>
<div class="panel"><div class="note"><b>Data audit:</b> The activity sheet states 96 readings in one day, but the supplied workbook contains @@P0@@ readings spanning 72 hours. This dashboard uses the supplied workbook unchanged rather than silently discarding two-thirds of the observations.</div>
<div class="metrics">
<div class="metric">Readings<b>@@P0@@</b><span class="small">n</span></div>
<div class="metric">Parameters<b>@@P3@@</b><span class="small">p = @@P3@@, df = @@P4@@</span></div>
<div class="metric">R²<b>@@P5@@</b><span class="small">SSE = @@P6@@ m²</span></div>
<div class="metric">Estimate error<b>@@P7@@ m</b><span class="small">standard error of estimate</span></div>
</div></div>
<div class="panel"><h2>Stage log time series</h2><p class="small">Raw logged level against elapsed time in hours from the first reading. Always visible.</p>@@P8@@</div>
<div class="panel"><div class="tabs">
<button class="active" onclick="showTab('deriv',this)">Tab 1 · Derivatives</button>
<button onclick="showTab('fit',this)">Tab 2 · Fitted curve</button>
<button onclick="showTab('area',this)">Tab 3 · Area</button>
</div>
<section id="deriv" class="tab active"><h2>Finite-difference derivatives</h2>@@P9@@
<p><b>Maximum raw dh/dt:</b> {np.max((h[1:]-h[:-1])/0.25):.4f} m/h. The second derivative is noisy because it differentiates rounded measurements; its main sign change around the fitted peak rate indicates that the filling rate first accelerates and then decelerates.</p></section>
<section id="fit" class="tab"><h2>Fitted curve</h2><p><b>Model:</b> @@P10@@</p>
<p>The double-logistic model was selected because the supplied record contains a rapid rise followed by recession/settling. A single rising logistic or Gompertz cannot represent that full shape; the second logistic term is therefore given a negative amplitude.</p>
<p><b>Initial guess p₀:</b> [14, 8, 0.5, 31, −4, 0.3, 38]. These values come from approximate baseline, rise size, steepest-rise timing, and later recession timing visible in the raw plot.</p>
@@P11@@
<h3>Parameter statistics</h3><table><thead><tr><th>Parameter</th><th>Estimate</th><th>SE</th><th>t</th><th>Two-tailed p</th></tr></thead><tbody>@@P12@@</tbody></table>
<div class="metrics" style="margin-top:12px">
<div class="metric">SSE<b>{sse:.6f}</b><span class="small">m²</span></div>
<div class="metric">SST<b>@@P13@@</b><span class="small">m²</span></div>
<div class="metric">R²<b>@@P5@@</b><span class="small">1 − SSE/SST</span></div>
<div class="metric">s<b>@@P14@@ m</b><span class="small">df = @@P4@@</span></div>
</div>
<h3>Residual reading</h3><p>@@P16@@</p>@@P17@@</section>
<section id="area" class="tab"><h2>Area under the fitted curve</h2>@@P18@@
<div class="metrics">
<div class="metric">quad area<b>@@P19@@</b><span class="small">meter-hours</span></div>
<div class="metric">quad error<b>@@P20@@</b><span class="small">absolute error estimate</span></div>
<div class="metric">Trapezoid check<b>@@P21@@</b><span class="small">meter-hours</span></div>
<div class="metric">Difference<b>@@P22@@</b><span class="small">quad − trapezoid</span></div>
</div>
<p>The area is in meter-hours, not volume. It represents accumulated reservoir level over the 72-hour interval. The trapezoid check differs because it integrates the discrete raw readings rather than the smooth fitted function.</p></section>
</div>
<div class="panel"><h2>Key interpretation</h2>
<p>The fitted curve reaches maximum <b>dh/dt = @@P23@@ m/h</b> at @@P24@@ hours after the first reading (approximately @@P25@@).</p>
<p>@@P16@@</p></div>
</main>
<script>
function showTab(id,btn){{{{
  document.querySelectorAll('.tab').forEach(x=>{{{x.classList.remove('active');}}});
  document.getElementById(id).classList.add('active');
  document.querySelectorAll('button').forEach(x=>x.classList.remove('active'));
  btn.classList.add('active');
}}}}
</script></body></html>"""
    html = html.replace("@@P0@@", str(n))
    html = html.replace("@@P1@@", str(start))
    html = html.replace("@@P2@@", str(end))
    html = html.replace("@@P3@@", str(p))
    html = html.replace("@@P4@@", str(dof))
    html = html.replace("@@P5@@", str(r2))
    html = html.replace("@@P6@@", str(sse))
    html = html.replace("@@P7@@", str(s))
    html = html.replace("@@P8@@", str(raw_svg))
    html = html.replace("@@P9@@", str(derivative_svg))
    html = html.replace("@@P10@@", str(model_equation))
    html = html.replace("@@P11@@", str(fit_svg))
    html = html.replace("@@P12@@", str(rows))
    html = html.replace("@@P13@@", str(sst))
    html = html.replace("@@P14@@", str(s))
    html = html.replace("@@P15@@", str(dof))
    html = html.replace("@@P16@@", str(residual_text))
    html = html.replace("@@P17@@", str(residual_svg))
    html = html.replace("@@P18@@", str(area_svg))
    html = html.replace("@@P19@@", str(area))
    html = html.replace("@@P20@@", str(quad_error))
    html = html.replace("@@P21@@", str(trapezoid))
    html = html.replace("@@P22@@", str(area-trapezoid))
    html = html.replace("@@P23@@", str(max_dhdt))
    html = html.replace("@@P24@@", str(t_max))
    html = html.replace("@@P25@@", str(str(max_time).replace("T"," ")[:16]))
    return html

if __name__ == "__main__":
    main()
