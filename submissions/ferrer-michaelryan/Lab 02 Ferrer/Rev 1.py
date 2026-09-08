import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
from matplotlib.lines import Line2D

# ---------------------------------------------------------
# 1. STRUCTURAL GEOMETRY & DATA (Y = Vertical)
# ---------------------------------------------------------
nodes_dict = {
    1: (0, 0, 0),
    2: (6, 0, 0),
    3: (6, 0, 6),
    4: (0, 0, 6),
    5: (0, 6, 0),
    6: (6, 6, 0),
    7: (6, 6, 6),
    8: (0, 6, 6)
}

# Member tuples: (ID, NodeA, NodeB, Type, Beta Angle, StartPinned, EndPinned)
members_info = [
    (1,  1, 2, 'Beam',   0, True, True),     # Base beam 1-2 (pinned at both ends)
    (2,  2, 3, 'Beam',   0, False, False),   # Base beam 2-3
    (3,  3, 4, 'Beam',   0, False, False),   # Base beam 3-4
    (4,  4, 1, 'Beam',   0, True, True),     # Base beam 4-1 (pinned at both ends)
    (5,  5, 6, 'Beam',   0, False, False),   # Roof beam 5-6
    (6,  6, 7, 'Beam',   0, False, False),   # Roof beam 6-7
    (7,  7, 8, 'Beam',   0, False, False),   # Roof beam 7-8
    (8,  8, 5, 'Beam',   0, False, False),   # Roof beam 8-5
    (9,  1, 5, 'Column', 90, True, False),   # Column 1-5 (pinned base at node 1)
    (10, 2, 6, 'Column', 90, True, False),   # Column 2-6 (pinned base at node 2)
    (11, 3, 7, 'Column', 90, True, False),   # Column 3-7 (pinned base at node 3)
    (12, 4, 8, 'Column', 90, True, False),   # Column 4-8 (pinned base at node 4)
]

# ---------------------------------------------------------
# 2. LOCAL AXIS ORIENTATION COMPUTATION
# ---------------------------------------------------------
def get_member_local_axes(pA, pB, beta_deg):
    pA, pB = np.array(pA, dtype=float), np.array(pB, dtype=float)
    v = pB - pA
    L = np.linalg.norm(v)
    local_x = v / L
    
    # Vertical member case (along global Y)
    if np.isclose(local_x[0], 0) and np.isclose(local_x[2], 0):
        beta = np.radians(beta_deg)
        ref_z = np.array([0, 0, 1.0]) # Parallel to global Z
        ref_y = np.cross(ref_z, local_x)
        ref_y = ref_y / np.linalg.norm(ref_y)
        
        local_y = ref_y * np.cos(beta) + ref_z * np.sin(beta)
        local_z = np.cross(local_x, local_y)
    else:
        # Horizontal beam case
        global_Y = np.array([0, 1.0, 0])
        beta = np.radians(beta_deg)
        
        z_temp = np.cross(local_x, global_Y)
        z_temp = z_temp / np.linalg.norm(z_temp)
        y_temp = np.cross(z_temp, local_x)
        
        local_y = y_temp * np.cos(beta) + z_temp * np.sin(beta)
        local_z = np.cross(local_x, local_y)
        
    return local_x, local_y, local_z

# ---------------------------------------------------------
# 3. PLOT INITIALIZATION
# ---------------------------------------------------------
fig = plt.figure(figsize=(13, 9.5))
ax = fig.add_subplot(111, projection='3d')

# Scatter nodes (Mapping data coords to display: X_plot=X, Y_plot=Z, Z_plot=Y)
for n, coord in nodes_dict.items():
    x, y, z = coord
    color = 'darkgreen' if n in [1, 2, 3, 4] else 'red'
    ax.scatter(x, z, y, color=color, s=70, zorder=5)
        
    dof_start = (n - 1) * 6 + 1
    dof_end = dof_start + 5
    ax.text(x + 0.1, z + 0.1, y + 0.25, f'N{n}\nDOF {dof_start}-{dof_end}', 
            fontsize=8, fontweight='bold', color='black')

# Plot origin green star
ax.scatter(0, 0, 0, color='green', marker='*', s=150, zorder=6)

# Draw members, local coordinate triads, and pinned symbols
arrow_len = 0.9

for mem_id, nA, nB, mtype, beta, start_pinned, end_pinned in members_info:
    pA, pB = nodes_dict[nA], nodes_dict[nB]
    color = '#1f77b4' if mtype == 'Beam' else '#2ca02c'
    
    # Draw member framing line
    ax.plot3D([pA[0], pB[0]], [pA[2], pB[2]], [pA[1], pB[1]], color=color, linewidth=2.5)
    
    # Draw Pinned Symbols (small hollow/white circles with dark borders) at member ends if pinned
    if start_pinned:
        # Offset slightly from node toward mid-span for clarity
        pt_pin = np.array(pA) * 0.85 + np.array(pB) * 0.15
        ax.scatter(pt_pin[0], pt_pin[2], pt_pin[1], color='white', edgecolor='black', s=45, marker='o', zorder=6)
    
    if end_pinned:
        pt_pin = np.array(pB) * 0.85 + np.array(pA) * 0.15
        ax.scatter(pt_pin[0], pt_pin[2], pt_pin[1], color='white', edgecolor='black', s=45, marker='o', zorder=6)

    # Draw local triad at mid-span
    p_mid = (np.array(pA) + np.array(pB)) / 2.0
    lx, ly, lz = get_member_local_axes(pA, pB, beta)
    
    # Local x (Red arrow)
    ax.quiver(p_mid[0], p_mid[2], p_mid[1], lx[0], lx[2], lx[1], 
              length=arrow_len, color='red', arrow_length_ratio=0.3, linewidth=1.5)
    # Local y (Green arrow)
    ax.quiver(p_mid[0], p_mid[2], p_mid[1], ly[0], ly[2], ly[1], 
              length=arrow_len, color='lime', arrow_length_ratio=0.3, linewidth=1.5)
    # Local z (Purple arrow)
    ax.quiver(p_mid[0], p_mid[2], p_mid[1], lz[0], lz[2], lz[1], 
              length=arrow_len, color='purple', arrow_length_ratio=0.3, linewidth=1.5)
              
    # Label member ID & Beta angle
    beta_txt = f' ($\\beta={beta}^\\circ$)' if beta != 0 else ''
    ax.text(p_mid[0] + 0.1, p_mid[2] + 0.1, p_mid[1] + 0.2, f'M{mem_id}{beta_txt}', 
            fontsize=8, fontweight='bold', color='navy' if mtype=='Beam' else 'darkgreen')

# Draw 3D Pinned Support Pyramids at Base Nodes (1 to 4)
for n in [1, 2, 3, 4]:
    x, y, z = nodes_dict[n]
    h, w = 0.8, 0.5
    apex = [x, z, y]
    base = [
        [x - w, z - w, y - h],
        [x + w, z - w, y - h],
        [x + w, z + w, y - h],
        [x - w, z + w, y - h]
    ]
    faces = [
        [apex, base[0], base[1]],
        [apex, base[1], base[2]],
        [apex, base[2], base[3]],
        [apex, base[3], base[0]],
        [base[0], base[1], base[2], base[3]]
    ]
    pyr = Poly3DCollection(faces, facecolors='dimgray', edgecolors='black', alpha=0.9, linewidths=0.8)
    ax.add_collection3d(pyr)

# ---------------------------------------------------------
# 4. HEADERS, LEGEND & MODEL DATA PANEL
# ---------------------------------------------------------
plt.suptitle('6m x 6m x 6m Cube - Structural Model, Rev. 1\nPinned supports, member end releases/pinned indicators, local axes shown', 
             fontsize=12, fontweight='bold', y=0.94)

legend_elements = [
    Line2D([0], [0], color='#1f77b4', lw=2.5, label='Beam'),
    Line2D([0], [0], color='#2ca02c', lw=2.5, label='Column'),
    Line2D([0], [0], marker='o', color='w', markerfacecolor='white', markeredgecolor='black', markersize=7, label='Pinned member connection'),
    Line2D([0], [0], marker='o', color='w', markerfacecolor='darkgreen', markersize=8, label='Supported node (pinned)'),
    Line2D([0], [0], marker='o', color='w', markerfacecolor='red', markersize=8, label='Free node'),
    Line2D([0], [0], marker='*', color='w', markerfacecolor='green', markersize=12, label='Origin (0, 0, 0)'),
    Line2D([0], [0], color='red', lw=2, label='Local x axis'),
    Line2D([0], [0], color='lime', lw=2, label='Local y axis'),
    Line2D([0], [0], color='purple', lw=2, label='Local z axis'),
]
ax.legend(handles=legend_elements, loc='upper left', fontsize=8, framealpha=0.9)

# Right-side Model Information Box
text_box_content = (
    "MODEL DATA - REV. 1\n\n"
    "Geometry\n"
    "  Cube edge           6.0 m\n"
    "  Nodes               8\n"
    "  Members             12\n"
    "  Vertical axis       global Y\n\n"
    "Supports & Releases\n"
    "  Support Type        Pinned (Nodes 1-4)\n"
    "  Restrained DOF      UX, UY, UZ\n"
    "  Member Pins         M1, M4 (Ends)\n"
    "                      M9-M12 (Base bases)\n\n"
    "Degrees of freedom\n"
    "  DOF per node        6\n"
    "  Total DOF           48\n"
    "  Restrained DOF      12\n"
    "  Active DOF          36\n"
    "  Numbering           (node - 1) * 6 + 1..6\n\n"
    "Beta angles\n"
    "  Base Beam           0 deg\n"
    "  Roof Beam           0 deg\n"
    "  Column              90 deg\n\n"
    "Local axes\n"
    "  local x  start node i to end node j\n"
    "  local y  in the vertical plane, upward\n"
    "  local z  completes right-handed set"
)

fig.text(0.72, 0.45, text_box_content, fontsize=7.5, family='monospace',
         bbox=dict(boxstyle='round,pad=0.8', facecolor='#eef2fa', edgecolor='#5577aa', alpha=0.95))

# Plot configuration
ax.set_xlabel('X (m) - lateral', fontsize=10, fontweight='bold', labelpad=10)
ax.set_ylabel('Z (m) - lateral', fontsize=10, fontweight='bold', labelpad=10)
ax.set_zlabel('Y (m) - vertical', fontsize=10, fontweight='bold', labelpad=10)

ax.set_xlim([-1, 7])
ax.set_ylim([-1, 7])
ax.set_zlim([-1, 7])
ax.grid(True)
ax.view_init(elev=20, azim=-60)

plt.subplots_adjust(left=0.05, right=0.70, top=0.88, bottom=0.08)
plt.show()