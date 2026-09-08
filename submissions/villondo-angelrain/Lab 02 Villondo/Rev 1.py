import time
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from mpl_toolkits.mplot3d.art3d import Poly3DCollection

# ==============================================================================
# 1. GEOMETRY & STRUCTURAL DEFINITIONS
# ==============================================================================
# Nodes Definition
nodes_data = {
    'Node': [1, 2, 3, 4, 5, 6, 7, 8],
    'X (m)': [0.0, 6.0, 6.0, 0.0, 0.0, 6.0, 6.0, 0.0],
    'Y (m)': [0.0, 0.0, 0.0, 0.0, 6.0, 6.0, 6.0, 6.0],
    'Z (m)': [0.0, 0.0, 6.0, 6.0, 0.0, 0.0, 6.0, 6.0],
    'Support': ['Pinned', 'Pinned', 'Pinned', 'Pinned', 'Free', 'Free', 'Free', 'Free']
}
df_nodes = pd.DataFrame(nodes_data)

# Members Definition: (ID, Node i, Node j, Type, Beta, Release i, Release j)
# 'Release-Mz' adds a moment pin release at that specific end
raw_members = [
    (1, 1, 2, 'Base Beam', 0.0, 'None', 'Release-Mz'),
    (2, 2, 3, 'Base Beam', 0.0, 'None', 'Release-Mz'),
    (3, 3, 4, 'Base Beam', 0.0, 'None', 'Release-Mz'),
    (4, 4, 1, 'Base Beam', 0.0, 'None', 'Release-Mz'),
    (5, 5, 6, 'Roof Beam', 0.0, 'None', 'Release-Mz'),
    (6, 6, 7, 'Roof Beam', 0.0, 'None', 'Release-Mz'),
    (7, 7, 8, 'Roof Beam', 0.0, 'None', 'Release-Mz'),
    (8, 8, 5, 'Roof Beam', 0.0, 'None', 'Release-Mz'),
    (9, 1, 5, 'Column', 90.0, 'None', 'None'),
    (10, 2, 6, 'Column', 90.0, 'None', 'None'),
    (11, 3, 7, 'Column', 90.0, 'None', 'None'),
    (12, 4, 8, 'Column', 90.0, 'None', 'None'),
]

def compute_local_axes(p1, p2, beta_deg):
    v = p2 - p1
    L = np.linalg.norm(v)
    x_loc = v / L
    
    # Check if vertical (parallel to global Y)
    if np.isclose(abs(x_loc[1]), 1.0):
        ref = np.array([0.0, 0.0, 1.0])
        z_temp = ref
        y_temp = np.cross(z_temp, x_loc)
    else:
        ref = np.array([0.0, 1.0, 0.0])
        z_temp = np.cross(x_loc, ref)
        z_temp /= np.linalg.norm(z_temp)
        y_temp = np.cross(z_temp, x_loc)
        
    beta_rad = np.radians(beta_deg)
    y_loc = y_temp * np.cos(beta_rad) + z_temp * np.sin(beta_rad)
    z_loc = -y_temp * np.sin(beta_rad) + z_temp * np.cos(beta_rad)
    
    return float(L), x_loc, y_loc, z_loc

members_list = []
local_axes_list = []

for m_id, ni, nj, mtype, beta, rel_i, rel_j in raw_members:
    p1 = df_nodes[df_nodes['Node'] == ni][['X (m)', 'Y (m)', 'Z (m)']].values[0]
    p2 = df_nodes[df_nodes['Node'] == nj][['X (m)', 'Y (m)', 'Z (m)']].values[0]
    
    length, x_loc, y_loc, z_loc = compute_local_axes(p1, p2, beta)
    
    members_list.append({
        'Member': m_id,
        'Node i (Start)': ni,
        'Node j (End)': nj,
        'Type': mtype,
        'Length (m)': length,
        'Beta (deg)': beta,
        'Release Node i': rel_i,
        'Release Node j': rel_j
    })
    
    local_axes_list.append({
        'Member': m_id,
        'Node i': ni,
        'Node j': nj,
        'Type': mtype,
        'Length (m)': length,
        'Beta (deg)': int(beta),
        'Release i': rel_i,
        'Release j': rel_j,
        'local x - X': round(x_loc[0], 3),
        'local x - Y': round(x_loc[1], 3),
        'local x - Z': round(x_loc[2], 3),
        'local y - X': round(y_loc[0], 3),
        'local y - Y': round(y_loc[1], 3),
        'local y - Z': round(y_loc[2], 3),
        'local z - X': round(z_loc[0], 3),
        'local z - Y': round(z_loc[1], 3),
        'local z - Z': round(z_loc[2], 3),
    })

df_members = pd.DataFrame(members_list)
df_local_axes = pd.DataFrame(local_axes_list)

# Supports Data
supports_list = []
for _, row in df_nodes.iterrows():
    is_pinned = row['Support'] == 'Pinned'
    supports_list.append({
        'Node': row['Node'],
        'X (m)': row['X (m)'],
        'Y (m)': row['Y (m)'],
        'Z (m)': row['Z (m)'],
        'Support Type': row['Support'],
        'UX': 'Restrained' if is_pinned else 'Active',
        'UY': 'Restrained' if is_pinned else 'Active',
        'UZ': 'Restrained' if is_pinned else 'Active',
        'RX': 'Active',
        'RY': 'Active',
        'RZ': 'Active',
        'Restraint Code': '111000' if is_pinned else '000000'
    })
df_supports = pd.DataFrame(supports_list)

# Node DOF Data
node_dof_list = []
for _, row in df_nodes.iterrows():
    nid = int(row['Node'])
    start_dof = (nid - 1) * 6 + 1
    end_dof = nid * 6
    is_pinned = row['Support'] == 'Pinned'
    node_dof_list.append({
        'Node': nid,
        'Support': row['Support'],
        'Global DOF Range': f"{start_dof} - {end_dof}",
        'UX': 'Restrained' if is_pinned else 'Active',
        'UY': 'Restrained' if is_pinned else 'Active',
        'UZ': 'Restrained' if is_pinned else 'Active',
        'RX': 'Active',
        'RY': 'Active',
        'RZ': 'Active',
        'Active DOF': 3 if is_pinned else 6
    })
df_node_dof = pd.DataFrame(node_dof_list)

# DOF Numbering Data
dof_num_list = []
for _, row in df_nodes.iterrows():
    nid = int(row['Node'])
    b_dof = (nid - 1) * 6
    dof_num_list.append({
        'Node': nid,
        'DOF_UX': b_dof + 1,
        'DOF_UY': b_dof + 2,
        'DOF_UZ': b_dof + 3,
        'DOF_RX': b_dof + 4,
        'DOF_RY': b_dof + 5,
        'DOF_RZ': b_dof + 6,
    })
df_dof_num = pd.DataFrame(dof_num_list)

# Model Summary Data
summary_data = {
    'Item': [
        'Revision', 'Cube edge length (m)', 'Number of nodes', 'Number of members',
        'Supported nodes', 'Support type', 'DOF per node', 'Total DOF',
        'Restrained DOF', 'Active DOF (equations)', 'Global vertical axis',
        'Global lateral axes', 'Beta angle, base beams (deg)',
        'Beta angle, roof beams (deg)', 'Beta angle, columns (deg)'
    ],
    'Value': [
        'Rev. 1', 6, 8, 12, 4, 'Pinned', 6, 48, 12, 36, 'Y', 'X and Z', 0, 0, 90
    ]
}
df_summary = pd.DataFrame(summary_data)

# ==============================================================================
# 2. GENERATE EXCEL WORKBOOK (WITH PERMISSION ERROR FALLBACK)
# ==============================================================================
excel_filename = 'cube_structural_model_Rev1.xlsx'

def write_excel_data(filename):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='Model Summary', index=False)
        df_nodes.to_excel(writer, sheet_name='Nodes', index=False)
        df_members.to_excel(writer, sheet_name='Member Incidences', index=False)
        df_supports.to_excel(writer, sheet_name='Supports', index=False)
        df_local_axes.to_excel(writer, sheet_name='Local Axes', index=False)
        df_node_dof.to_excel(writer, sheet_name='Node DOF', index=False)
        df_dof_num.to_excel(writer, sheet_name='DOF Numbering', index=False)

try:
    write_excel_data(excel_filename)
except PermissionError:
    excel_filename = f'cube_structural_model_Rev1_{int(time.time())}.xlsx'
    print(f"\n[WARNING] Primary Excel file is currently open! Exporting to fallback file: {excel_filename}\n")
    write_excel_data(excel_filename)

# Apply Styling
wb = openpyxl.load_workbook(excel_filename)
header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
center_align = Alignment(horizontal="center", vertical="center")
left_align = Alignment(horizontal="left", vertical="center")
thin_border = Border(
    left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
)

restrained_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
active_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

for sheetname in wb.sheetnames:
    ws = wb[sheetname]
    ws.views.sheetView[0].showGridLines = True
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = center_align
            cell.border = thin_border
            
            if str(cell.value) == 'Restrained':
                cell.fill = restrained_fill
            elif str(cell.value) == 'Active':
                cell.fill = active_fill

    if sheetname == 'Model Summary':
        for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
            for cell in row:
                cell.alignment = left_align

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

wb.save(excel_filename)

# ==============================================================================
# 3. GENERATE 3D STRUCTURAL DIAGRAM
# ==============================================================================
fig = plt.figure(figsize=(16, 11), facecolor='white')
ax = fig.add_subplot(111, projection='3d')

plt.suptitle("6m x 6m x 6m Cube - Structural Model, Rev. 1\nPinned supports at nodes 1-4, member local axes, beta angles, and beam end releases shown",
             fontsize=13, fontweight='bold', y=0.95)

# Plot Members & Member Releases
for _, row in df_local_axes.iterrows():
    ni, nj = int(row['Node i']), int(row['Node j'])
    mtype = row['Type']
    p1 = df_nodes[df_nodes['Node'] == ni][['X (m)', 'Y (m)', 'Z (m)']].values[0]
    p2 = df_nodes[df_nodes['Node'] == nj][['X (m)', 'Y (m)', 'Z (m)']].values[0]
    
    # Member line
    color = 'royalblue' if 'Beam' in mtype else 'seagreen'
    ax.plot([p1[0], p2[0]], [p1[2], p2[2]], [p1[1], p2[1]], color=color, linewidth=2.5, zorder=2)
    
    # Local coordinate triad
    mid = (p1 + p2) / 2.0
    x_vec = np.array([row['local x - X'], row['local x - Y'], row['local x - Z']])
    y_vec = np.array([row['local y - X'], row['local y - Y'], row['local y - Z']])
    z_vec = np.array([row['local z - X'], row['local z - Y'], row['local z - Z']])
    
    scale = 0.7
    ax.quiver(mid[0], mid[2], mid[1], x_vec[0], x_vec[2], x_vec[1], length=scale, color='crimson', linewidth=1.5, arrow_length_ratio=0.25)
    ax.quiver(mid[0], mid[2], mid[1], y_vec[0], y_vec[2], y_vec[1], length=scale, color='limegreen', linewidth=1.5, arrow_length_ratio=0.25)
    ax.quiver(mid[0], mid[2], mid[1], z_vec[0], z_vec[2], z_vec[1], length=scale, color='mediumpurple', linewidth=1.5, arrow_length_ratio=0.25)
    
    # Member Labels
    m_label = f"M{int(row['Member'])}"
    beta_val = int(row['Beta (deg)'])
    if beta_val != 0:
        m_label += f" ($\\beta={beta_val}^\\circ$)"
    ax.text(mid[0], mid[2], mid[1] + 0.3, m_label, color='navy', fontsize=8, fontweight='bold', ha='center')

    # Draw Beam Pin Symbols at member releases (10% offset from end joints)
    if 'Release' in str(row['Release i']):
        pin_p1 = p1 + 0.10 * (p2 - p1)
        ax.scatter(pin_p1[0], pin_p1[2], pin_p1[1], color='orange', s=55, marker='o', edgecolors='black', zorder=6)
        
    if 'Release' in str(row['Release j']):
        pin_p2 = p1 + 0.90 * (p2 - p1)
        ax.scatter(pin_p2[0], pin_p2[2], pin_p2[1], color='orange', s=55, marker='o', edgecolors='black', zorder=6)

# Plot Nodes & Labels
for _, row in df_nodes.iterrows():
    nid = int(row['Node'])
    x, y, z = row['X (m)'], row['Y (m)'], row['Z (m)']
    
    ax.scatter(x, z, y, color='crimson', s=60, zorder=5)
    
    dof_start = (nid - 1) * 6 + 1
    dof_end = nid * 6
    ax.text(x, z, y + 0.35, f"  N{nid}\n  DOF {dof_start}-{dof_end}", fontsize=8, fontweight='bold', color='black')

# Plot Pinned Support Pyramids (Nodes 1-4)
def draw_pinned_pyramid(ax, x, z, y, h=0.7, w=0.5):
    apex = np.array([x, z, y])
    b1 = np.array([x - w, z - w, y - h])
    b2 = np.array([x + w, z - w, y - h])
    b3 = np.array([x + w, z + w, y - h])
    b4 = np.array([x - w, z + w, y - h])
    
    faces = [[apex, b1, b2], [apex, b2, b3], [apex, b3, b4], [apex, b4, b1], [b1, b2, b3, b4]]
    pyramid = Poly3DCollection(faces, facecolors='dimgray', edgecolors='black', alpha=0.85, zorder=4)
    ax.add_collection3d(pyramid)

for _, row in df_nodes[df_nodes['Support'] == 'Pinned'].iterrows():
    draw_pinned_pyramid(ax, row['X (m)'], row['Z (m)'], row['Y (m)'])

# Custom Legend
legend_elements = [
    plt.Line2D([0], [0], color='royalblue', lw=2, label='Beam'),
    plt.Line2D([0], [0], color='seagreen', lw=2, label='Column'),
    plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='crimson', markersize=8, label='Supported node (pinned)'),
    plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='orange', markeredgecolor='black', markersize=7, label='Member end release (pin)'),
    plt.Line2D([0], [0], color='crimson', lw=1.5, label='Local x axis'),
    plt.Line2D([0], [0], color='limegreen', lw=1.5, label='Local y axis'),
    plt.Line2D([0], [0], color='mediumpurple', lw=1.5, label='Local z axis'),
]
ax.legend(handles=legend_elements, loc='upper left', bbox_to_anchor=(-0.1, 1.0), frameon=True, fontsize=8)

# Model Data Side Panel
panel_text = """MODEL DATA - REV. 1

Geometry
  Cube edge         6.0 m
  Nodes             8
  Members           12
  Vertical axis     global Y

Supports
  Type              pinned
  Nodes             1, 2, 3, 4
  Restrained        UX, UY, UZ
  Released          RX, RY, RZ

Degrees of freedom
  DOF per node      6
  Total DOF         48
  Restrained DOF    12
  Active DOF        36
  Numbering         (node - 1) * 6 + 1..6

Beta angles
  Base Beam         0 deg
  Roof Beam         0 deg
  Column            90 deg

Local axes
  local x   start node i to end node j
  local y   in the vertical plane, upward
  local z   completes the right-handed set
  Vertical members follow the limiting
  case: local z parallel to global Z."""

fig.text(0.74, 0.50, panel_text, fontsize=8, family='monospace',
         verticalalignment='center',
         bbox=dict(boxstyle='round,pad=0.8', facecolor='#F0F4F8', edgecolor='#1B365D', alpha=0.9))

# Axis Formatting
ax.set_xlabel('X (m) - lateral', fontsize=9, fontweight='bold', labelpad=8)
ax.set_ylabel('Z (m) - lateral', fontsize=9, fontweight='bold', labelpad=8)
ax.set_zlabel('Y (m) - vertical', fontsize=9, fontweight='bold', labelpad=8)

ax.set_xlim([-1, 7])
ax.set_ylim([-1, 7])
ax.set_zlim([-1, 7])

ax.view_init(elev=18, azim=-60)
ax.grid(True, linestyle=':', alpha=0.6)

plt.subplots_adjust(left=0.05, right=0.72, top=0.90, bottom=0.05)
plt.show()
