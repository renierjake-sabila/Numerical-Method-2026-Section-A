import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
from matplotlib.lines import Line2D
import openpyxl

# -------------------------------------------------------------------------
# 1. DEFINE NODE COORDINATES (Y as Vertical Axis)
# -------------------------------------------------------------------------
nodes_data = {
    'Node ID': [1, 2, 3, 4, 5, 6, 7, 8],
    'X (m)': [0.0, 6.0, 6.0, 0.0, 0.0, 6.0, 6.0, 0.0],
    'Y (m)': [0.0, 0.0, 0.0, 0.0, 6.0, 6.0, 6.0, 6.0],  # Y is vertical (up)
    'Z (m)': [0.0, 0.0, 6.0, 6.0, 0.0, 0.0, 6.0, 6.0],
    'Support': ['Pinned', 'Pinned', 'Pinned', 'Pinned', 'Free', 'Free', 'Free', 'Free'],
    'DOF Start': [1, 7, 13, 19, 25, 31, 37, 43],
    'DOF End': [6, 12, 18, 24, 30, 36, 42, 48]
}
df_nodes = pd.DataFrame(nodes_data)

# -------------------------------------------------------------------------
# 2. DEFINE MEMBER CONNECTIVITY, TYPES, & BETA ANGLES
# -------------------------------------------------------------------------
members_data = {
    'Member ID': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],
    'Start Node': [1, 2, 3, 4, 5, 6, 7, 8, 1, 2, 3, 4],
    'End Node': [2, 3, 4, 1, 6, 7, 8, 5, 5, 6, 7, 8],
    'Type': ['Base Beam', 'Base Beam', 'Base Beam', 'Base Beam', 
             'Roof Beam', 'Roof Beam', 'Roof Beam', 'Roof Beam', 
             'Column', 'Column', 'Column', 'Column'],
    'Length (m)': [6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0],
    'Beta Angle (deg)': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 90.0, 90.0, 90.0, 90.0],
    'MZ Released': [True, False, True, False, True, False, True, False, False, False, False, False]
}
df_members = pd.DataFrame(members_data)

# -------------------------------------------------------------------------
# 3. SETUP FIGURE WITH GRIDSPEC (Separate Panel for Model Data & 3D Model)
# -------------------------------------------------------------------------
fig = plt.figure(figsize=(18, 11), dpi=100)
gs = fig.add_gridspec(1, 2, width_ratios=[2.3, 1], wspace=0.05)

ax = fig.add_subplot(gs[0, 0], projection='3d')
ax_text = fig.add_subplot(gs[0, 1])
ax_text.axis('off')

coords = df_nodes[['X (m)', 'Y (m)', 'Z (m)']].values

# Define Cube Faces (Y is vertical)
faces_idx = [
    [0, 1, 2, 3], # Bottom face (Y=0)
    [4, 5, 6, 7], # Top face (Y=6)
    [0, 1, 5, 4], # Front face (Z=0)
    [3, 2, 6, 7], # Back face (Z=6)
    [0, 3, 7, 4], # Left face (X=0)
    [1, 2, 6, 5]  # Right face (X=6)
]
face_vertices = [[coords[idx] for idx in face] for face in faces_idx]
poly3d = Poly3DCollection(face_vertices, facecolors='#3498db', linewidths=1.2, edgecolors='#2c3e50', alpha=0.12)
ax.add_collection3d(poly3d)

# Plot Member Lines & End Releases [MZ]
for _, row in df_members.iterrows():
    n_start = df_nodes[df_nodes['Node ID'] == row['Start Node']].values[0]
    n_end = df_nodes[df_nodes['Node ID'] == row['End Node']].values[0]
    
    x_s, y_s, z_s = n_start[1], n_start[2], n_start[3]
    x_e, y_e, z_e = n_end[1], n_end[2], n_end[3]
    
    color = '#2980b9' if 'Beam' in row['Type'] else '#27ae60'
    lw = 2.5 if 'Column' in row['Type'] else 2.0
    ax.plot([x_s, x_e], [y_s, y_e], [z_s, z_e], color=color, linewidth=lw)

    # Midpoint annotation for members and beta angles / releases
    mid_x, mid_y, mid_z = (x_s + x_e) / 2, (y_s + y_e) / 2, (z_s + z_e) / 2
    if row['Beta Angle (deg)'] != 0:
        label = f"M{int(row['Member ID'])} (\u03b2={int(row['Beta Angle (deg)'])}°)"
    elif row['MZ Released']:
        label = f"M{int(row['Member ID'])} [MZ]"
    else:
        label = f"M{int(row['Member ID'])}"
    
    ax.text(mid_x, mid_y + 0.18, mid_z, label, fontsize=8, color='#1b4f72', fontweight='bold')

    # If MZ released, plot hollow white circles near ends
    if row['MZ Released']:
        frac1 = 0.2
        frac2 = 0.8
        rx1, ry1, rz1 = x_s + frac1*(x_e - x_s), y_s + frac1*(y_e - y_s), z_s + frac1*(z_e - z_s)
        rx2, ry2, rz2 = x_s + frac2*(x_e - x_s), y_s + frac2*(y_e - y_s), z_s + frac2*(z_e - z_s)
        ax.scatter([rx1, rx2], [ry1, ry2], [rz1, rz2], color='white', edgecolor='black', s=35, depthshade=False, zorder=6)

# Helper function to draw 3D Pyramidal Support Bases at Pinned Nodes
def draw_support_pyramid(ax, x, y, z, size=0.8):
    h = size
    w = size / 2
    base_pts = np.array([
        [x - w, y - h, z - w],
        [x + w, y - h, z - w],
        [x + w, y - h, z + w],
        [x - w, y - h, z + w]
    ])
    apex = np.array([x, y, z])
    
    faces = [
        [base_pts[0], base_pts[1], apex],
        [base_pts[1], base_pts[2], apex],
        [base_pts[2], base_pts[3], apex],
        [base_pts[3], base_pts[0], apex],
        [base_pts[0], base_pts[1], base_pts[2], base_pts[3]]
    ]
    ax.add_collection3d(Poly3DCollection(faces, facecolors='#7f8c8d', edgecolors='#2c3e50', alpha=0.85))

# Plot Nodes & Supports
for _, row in df_nodes.iterrows():
    x, y, z = row['X (m)'], row['Y (m)'], row['Z (m)']
    if row['Support'] == 'Pinned':
        draw_support_pyramid(ax, x, y, z, size=0.9)
        ax.scatter([x], [y], [z], color='#27ae60', s=70, depthshade=False, zorder=5)
    else:
        ax.scatter([x], [y], [z], color='#e74c3c', s=55, depthshade=False, zorder=5)
    
    ax.text(x + 0.2, y + 0.35, z + 0.2, f"N{int(row['Node ID'])}\nDOF {int(row['DOF Start'])}-{int(row['DOF End'])}", fontsize=8, fontweight='bold', color='#111111')

# Draw Global Coordinate Axis Triad at Origin (-1.2, -1.2, -1.2)
ox, oy, oz = -1.2, -1.2, -1.2
ax.quiver(ox, oy, oz, 1.5, 0, 0, color='#d35400', arrow_length_ratio=0.2, linewidth=2.5) # X
ax.quiver(ox, oy, oz, 0, 1.5, 0, color='#2980b9', arrow_length_ratio=0.2, linewidth=2.5) # Y (vertical)
ax.quiver(ox, oy, oz, 0, 0, 1.5, color='#d35400', arrow_length_ratio=0.2, linewidth=2.5) # Z

ax.text(ox + 1.7, oy, oz, 'X', color='#d35400', fontweight='bold', fontsize=11)
ax.text(ox, oy + 1.7, oz, 'Y', color='#2980b9', fontweight='bold', fontsize=11)
ax.text(ox, oy, oz + 1.7, 'Z', color='#d35400', fontweight='bold', fontsize=11)
ax.text(ox, oy - 0.4, oz, 'GLOBAL', color='#333333', fontweight='bold', fontsize=9)

# Plot Formatting & Labels
ax.set_xlabel('X (m) - lateral', fontsize=10, fontweight='bold', labelpad=8)
ax.set_ylabel('Y (m) - vertical', fontsize=10, fontweight='bold', labelpad=8)
ax.set_zlabel('Z (m) - lateral', fontsize=10, fontweight='bold', labelpad=8)
ax.set_title('6m x 6m x 6m Cube - Structural Model, Rev. 1\nPinned supports at nodes 1-4, global and local axes, beta angles, and MZ end releases', fontsize=11, fontweight='bold', pad=15)

ax.set_xlim([-2, 8])
ax.set_ylim([-2, 8])
ax.set_zlim([-2, 8])
ax.set_box_aspect([1, 1, 1])
ax.view_init(elev=22, azim=-52)

# Add Legends explicitly on the top-left of the 3D plot
legend_elements = [
    Line2D([0], [0], color='#2980b9', lw=2.5, label='Beam'),
    Line2D([0], [0], color='#27ae60', lw=2.5, label='Column'),
    Line2D([0], [0], marker='o', color='w', markerfacecolor='#27ae60', markersize=8, label='Supported node (pinned)'),
    Line2D([0], [0], marker='o', color='w', markerfacecolor='#e74c3c', markersize=8, label='Free node'),
    Line2D([0], [0], marker='o', color='w', markerfacecolor='white', markeredgecolor='black', markersize=7, label='Pinned member end (MZ released)'),
    Line2D([0], [0], marker='*', color='w', markerfacecolor='#27ae60', markersize=10, label='Origin (0, 0, 0)'),
    Line2D([0], [0], color='#d35400', lw=2, label='Global X axis'),
    Line2D([0], [0], color='#2980b9', lw=2, label='Global Y axis'),
    Line2D([0], [0], color='#d35400', lw=2, label='Global Z axis'),
]
ax.legend(handles=legend_elements, loc='upper left', fontsize=8, framealpha=0.9, facecolor='#f8f9f9', edgecolor='#bdc3c7')

# Add Model Data Panel cleanly on the right subplot (`ax_text`)
model_text = (
    "MODEL DATA - REV. 1\n\n"
    "Geometry\n"
    "  Cube edge        6.0 m\n"
    "  Nodes            8\n"
    "  Members          12\n\n"
    "Global axes\n"
    "  X                lateral\n"
    "  Y                vertical (up)\n"
    "  Z                lateral\n\n"
    "Supports\n"
    "  Type             pinned\n"
    "  Nodes            1, 2, 3, 4\n"
    "  Restrained       UX, UY, UZ\n"
    "  Released         RX, RY, RZ\n\n"
    "Nodal degrees of freedom\n"
    "  DOF per node     6\n"
    "  Total DOF        48\n"
    "  Restrained DOF   12\n"
    "  Active DOF       36\n\n"
    "Number and releases\n"
    "  Pinned members   1, 3, 5, 7\n"
    "  Pattern          Pinned i and j\n"
    "  Component        MZ, moment about local z\n"
    "  Released end DOF 8\n"
    "  Symbol           hollow circle at the end\n\n"
    "Beta angles\n"
    "  Base Beam        0 deg\n"
    "  Roof Beam        0 deg\n"
    "  Column           90 deg\n\n"
    "Local axes\n"
    "  local x          start node i to end node j\n"
    "  local y          in the vertical plane, upward\n"
    "  local z          completes right-handed set"
)
ax_text.text(0.05, 0.98, model_text, transform=ax_text.transAxes, fontsize=8.5, verticalalignment='top', 
             bbox=dict(boxstyle='round,pad=0.8', facecolor='#f4f6f7', alpha=0.95, edgecolor='#bdc3c7'), fontfamily='monospace')

plt.tight_layout()
plt.show()