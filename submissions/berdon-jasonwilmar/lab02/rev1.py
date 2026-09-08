"""
structural_solver_rev1_excel_tables.py
3D frame solver – Excel output now matches the requested tables.
"""
import pandas as pd
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
import numpy as np
from math import radians, cos, sin
import os
import matplotlib.gridspec as gridspec

# Try to import openpyxl
try:
    import openpyxl
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

# ============================================================================
# 1. MODEL DEFINITION (edit these to change the model)
# ============================================================================

nodes = {
    1: (0, 0, 0),
    2: (6, 0, 0),
    3: (6, 0, 6),
    4: (0, 0, 6),
    5: (0, 6, 0),
    6: (6, 6, 0),
    7: (6, 6, 6),
    8: (0, 6, 6)
}

supports = {1: 'pinned', 2: 'pinned', 3: 'pinned', 4: 'pinned'}

members_data = [
    (1, 2, 0, [], []),
    (2, 3, 0, [], []),
    (3, 4, 0, [], []),
    (4, 1, 0, [], []),
    (5, 6, 0, [], []),
    (6, 7, 0, [], []),
    (7, 8, 0, [], []),
    (8, 5, 0, [], []),
    (1, 5, 90, ['rz'], []),
    (2, 6, 90, ['rz'], []),
    (3, 7, 90, ['rz'], []),
    (4, 8, 90, ['rz'], [])
]

loads = {
    5: (0, -100, 0, 0, 0, 0),
    6: (0, -100, 0, 0, 0, 0),
    7: (0, -100, 0, 0, 0, 0),
    8: (0, -100, 0, 0, 0, 0)
}

E = 200e9
G = 80e9
A = 0.01
I_y = 1e-6
I_z = 1e-6
J = 1e-6

# ============================================================================
# 2. DOF NUMBERING
# ============================================================================

node_ids = list(nodes.keys())
num_nodes = len(node_ids)
node_id_to_idx = {nid: i for i, nid in enumerate(node_ids)}

fixed_dofs = set()
for nid, stype in supports.items():
    if stype == 'pinned':
        idx = node_id_to_idx[nid]
        for dof in range(3):  # UX, UY, UZ
            fixed_dofs.add(idx * 6 + dof)

total_dofs = num_nodes * 6
free_dofs = [d for d in range(total_dofs) if d not in fixed_dofs]
num_free = len(free_dofs)
dof_to_eq = [-1] * total_dofs
for eq, dof in enumerate(free_dofs):
    dof_to_eq[dof] = eq

node_dofs = {}
for nid in node_ids:
    idx = node_id_to_idx[nid]
    node_dofs[nid] = [idx*6 + d for d in range(6)]

# ============================================================================
# 3. ELEMENT STIFFNESS
# ============================================================================

def beam_stiffness_3d(E, G, A, I_y, I_z, J, L, beta_rad, releases_start, releases_end):
    k_local = np.zeros((12, 12))
    k_local[0,0] = k_local[6,6] = E*A/L
    k_local[0,6] = k_local[6,0] = -E*A/L
    k_local[3,3] = k_local[9,9] = G*J/L
    k_local[3,9] = k_local[9,3] = -G*J/L
    k_local[1,1] = 12*E*I_z/L**3
    k_local[1,5] = 6*E*I_z/L**2
    k_local[1,7] = -12*E*I_z/L**3
    k_local[1,11] = 6*E*I_z/L**2
    k_local[5,1] = k_local[1,5]
    k_local[5,5] = 4*E*I_z/L
    k_local[5,7] = -6*E*I_z/L**2
    k_local[5,11] = 2*E*I_z/L
    k_local[7,1] = k_local[1,7]
    k_local[7,5] = k_local[5,7]
    k_local[7,7] = 12*E*I_z/L**3
    k_local[7,11] = -6*E*I_z/L**2
    k_local[11,1] = k_local[1,11]
    k_local[11,5] = k_local[5,11]
    k_local[11,7] = k_local[7,11]
    k_local[11,11] = 4*E*I_z/L
    k_local[2,2] = 12*E*I_y/L**3
    k_local[2,4] = -6*E*I_y/L**2
    k_local[2,8] = -12*E*I_y/L**3
    k_local[2,10] = -6*E*I_y/L**2
    k_local[4,2] = k_local[2,4]
    k_local[4,4] = 4*E*I_y/L
    k_local[4,8] = 6*E*I_y/L**2
    k_local[4,10] = 2*E*I_y/L
    k_local[8,2] = k_local[2,8]
    k_local[8,4] = k_local[4,8]
    k_local[8,8] = 12*E*I_y/L**3
    k_local[8,10] = 6*E*I_y/L**2
    k_local[10,2] = k_local[2,10]
    k_local[10,4] = k_local[4,10]
    k_local[10,8] = k_local[8,10]
    k_local[10,10] = 4*E*I_y/L

    dof_map = {'ux':0, 'uy':1, 'uz':2, 'rx':3, 'ry':4, 'rz':5}
    released_indices = []
    for rel in releases_start:
        if rel in dof_map:
            released_indices.append(dof_map[rel])
    for rel in releases_end:
        if rel in dof_map:
            released_indices.append(dof_map[rel] + 6)
    small = 1e-6 * E * A / L
    for idx in released_indices:
        k_local[idx, :] = 0.0
        k_local[:, idx] = 0.0
        k_local[idx, idx] = small
    return k_local

# ============================================================================
# 4. ASSEMBLY & SOLUTION
# ============================================================================

members = [(n1, n2, beta, rel_s, rel_e) for (n1, n2, beta, rel_s, rel_e) in members_data]

K_global = np.zeros((num_free, num_free))
F_global = np.zeros(num_free)

for nid, (fx, fy, fz, mx, my, mz) in loads.items():
    idx = node_id_to_idx[nid]
    for dof, val in enumerate([fx, fy, fz, mx, my, mz]):
        glob_dof = idx * 6 + dof
        eq = dof_to_eq[glob_dof]
        if eq != -1:
            F_global[eq] += val

for n1, n2, beta_deg, rel_s, rel_e in members:
    p1 = np.array(nodes[n1])
    p2 = np.array(nodes[n2])
    L = np.linalg.norm(p2 - p1)
    if L < 1e-12:
        continue
    ex = (p2 - p1) / L
    if abs(ex[1]) < 0.9:
        ref = np.array([0, 1, 0])
    else:
        ref = np.array([0, 0, 1])
    ey_unnorm = np.cross(ex, ref)
    if np.linalg.norm(ey_unnorm) < 1e-12:
        ey_unnorm = np.array([0, 0, 1])
    ey = ey_unnorm / np.linalg.norm(ey_unnorm)
    beta_rad = radians(beta_deg)
    v = ey
    k = ex
    cosb = cos(beta_rad)
    sinb = sin(beta_rad)
    ey_rot = v * cosb + np.cross(k, v) * sinb + k * (np.dot(k, v)) * (1 - cosb)
    ey = ey_rot / np.linalg.norm(ey_rot)
    ez = np.cross(ex, ey)
    T = np.array([ex, ey, ez]).T
    Tmat = np.zeros((12, 12))
    for i in range(4):
        Tmat[i*3:(i+1)*3, i*3:(i+1)*3] = T

    k_local = beam_stiffness_3d(E, G, A, I_y, I_z, J, L, beta_rad, rel_s, rel_e)
    k_global = Tmat.T @ k_local @ Tmat

    idx1 = node_id_to_idx[n1]
    idx2 = node_id_to_idx[n2]
    dofs1 = [idx1*6 + d for d in range(6)]
    dofs2 = [idx2*6 + d for d in range(6)]
    all_dofs = dofs1 + dofs2
    eqs = [dof_to_eq[d] for d in all_dofs]
    for i, eq_i in enumerate(eqs):
        if eq_i == -1:
            continue
        for j, eq_j in enumerate(eqs):
            if eq_j == -1:
                continue
            K_global[eq_i, eq_j] += k_global[i, j]

U_free = np.linalg.solve(K_global, F_global)
U_total = np.zeros(total_dofs)
for eq, dof in enumerate(free_dofs):
    U_total[dof] = U_free[eq]

# ============================================================================
# 5. MEMBER FORCES (for Excel output)
# ============================================================================

member_results = []
for n1, n2, beta_deg, rel_s, rel_e in members:
    p1 = np.array(nodes[n1])
    p2 = np.array(nodes[n2])
    L = np.linalg.norm(p2 - p1)
    if L < 1e-12:
        continue
    ex = (p2 - p1) / L
    if abs(ex[1]) < 0.9:
        ref = np.array([0, 1, 0])
    else:
        ref = np.array([0, 0, 1])
    ey_unnorm = np.cross(ex, ref)
    if np.linalg.norm(ey_unnorm) < 1e-12:
        ey_unnorm = np.array([0, 0, 1])
    ey = ey_unnorm / np.linalg.norm(ey_unnorm)
    beta_rad = radians(beta_deg)
    v = ey
    k = ex
    cosb = cos(beta_rad)
    sinb = sin(beta_rad)
    ey_rot = v * cosb + np.cross(k, v) * sinb + k * (np.dot(k, v)) * (1 - cosb)
    ey = ey_rot / np.linalg.norm(ey_rot)
    ez = np.cross(ex, ey)
    T = np.array([ex, ey, ez]).T
    Tmat = np.zeros((12, 12))
    for i in range(4):
        Tmat[i*3:(i+1)*3, i*3:(i+1)*3] = T

    idx1 = node_id_to_idx[n1]
    idx2 = node_id_to_idx[n2]
    dofs1 = [idx1*6 + d for d in range(6)]
    dofs2 = [idx2*6 + d for d in range(6)]
    all_dofs = dofs1 + dofs2
    u_global = np.array([U_total[d] for d in all_dofs])
    u_local = Tmat @ u_global
    k_local = beam_stiffness_3d(E, G, A, I_y, I_z, J, L, beta_rad, rel_s, rel_e)
    forces_local = k_local @ u_local
    Fi = forces_local[0:6]
    Fj = forces_local[6:12]
    member_results.append({
        'member': (n1, n2),
        'L': L,
        'beta': beta_deg,
        'Fi': Fi,
        'Fj': Fj,
        'u_local': u_local
    })

# ============================================================================
# 6. EXCEL OUTPUT – with all required tables
# ============================================================================

# Determine script directory for saving
script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
excel_filename = 'structural_solver_rev1_tables.xlsx'
excel_path = os.path.join(script_dir, excel_filename)

if not OPENPYXL_OK:
    print("❌ 'openpyxl' is not installed. Please install: pip install openpyxl")
else:
    try:
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            # ---- Sheet 1: Model Data ----
            model_data = {
                'Item': ['Revision', 'Cube edge length (m)', 'Number of nodes', 'Number of members',
                         'Supported nodes', 'Support type', 'DOF per node', 'Total DOF',
                         'Restrained DOF', 'Active DOF (equations)',
                         'Global vertical axis', 'Global lateral axes',
                         'Beta angle, base beams (deg)', 'Beta angle, roof beams (deg)',
                         'Beta angle, columns (deg)'],
                'Value': ['Rev. 1', 6.0, num_nodes, len(members_data),
                          len(supports), 'Pinned', 6, total_dofs,
                          len(fixed_dofs), num_free,
                          'Y', 'X and Z',
                          0, 0, 90]
            }
            model_df = pd.DataFrame(model_data)
            model_df.to_excel(writer, sheet_name='Model Data', index=False)

            # ---- Sheet 2: Node Data ----
            node_rows = []
            for nid in node_ids:
                x, y, z = nodes[nid]
                support = supports.get(nid, 'Free')
                # Support type string
                support_type = 'Pinned' if support == 'pinned' else 'Free'
                # Global DOF range
                start_dof = (nid - 1) * 6 + 1
                end_dof = nid * 6
                dof_range = f"{start_dof} - {end_dof}"
                # Status per DOF
                statuses = []
                for dof in range(6):
                    glob_dof = (nid - 1) * 6 + dof
                    if glob_dof in fixed_dofs:
                        statuses.append('Restrained')
                    else:
                        statuses.append('Active')
                # Restraint code: 1 for restrained, 0 for active
                code = ''.join(['1' if s == 'Restrained' else '0' for s in statuses])
                active_count = statuses.count('Active')
                node_rows.append([
                    nid, x, y, z, support_type, dof_range,
                    statuses[0], statuses[1], statuses[2],
                    statuses[3], statuses[4], statuses[5],
                    code, active_count
                ])
            node_df = pd.DataFrame(node_rows,
                                   columns=['Node', 'X (m)', 'Y (m)', 'Z (m)', 'Support Type',
                                            'Global DOF Range', 'UX', 'UY', 'UZ', 'RX', 'RY', 'RZ',
                                            'Restraint Code', 'Active DOF'])
            node_df.to_excel(writer, sheet_name='Node Data', index=False)

            # ---- Sheet 3: Member Data ----
            member_rows = []
            for i, (n1, n2, beta, rel_s, rel_e) in enumerate(members, start=1):
                # Determine member type
                y1 = nodes[n1][1]
                y2 = nodes[n2][1]
                if y1 == 0 and y2 == 0:
                    mtype = 'Base Beam'
                elif y1 == 6 and y2 == 6:
                    mtype = 'Roof Beam'
                else:
                    mtype = 'Column'
                L = np.linalg.norm(np.array(nodes[n2]) - np.array(nodes[n1]))
                member_rows.append([i, n1, n2, mtype, round(L, 3), beta])
            member_df = pd.DataFrame(member_rows,
                                     columns=['Member', 'Node i (Start)', 'Node j (End)',
                                              'Type', 'Length (m)', 'Beta (deg)'])
            member_df.to_excel(writer, sheet_name='Member Data', index=False)

            # ---- Sheet 4: DOF Numbering ----
            dof_rows = []
            for nid in node_ids:
                dofs = node_dofs[nid]  # [UX, UY, UZ, RX, RY, RZ] as global DOF numbers
                dof_rows.append([nid] + [d+1 for d in dofs])  # convert 0‑based to 1‑based
            dof_df = pd.DataFrame(dof_rows,
                                  columns=['Node', 'DOF_UX', 'DOF_UY', 'DOF_UZ',
                                           'DOF_RX', 'DOF_RY', 'DOF_RZ'])
            dof_df.to_excel(writer, sheet_name='DOF Numbering', index=False)

        print(f"✅ Excel file successfully saved at:\n   {excel_path}")

    except PermissionError:
        print(f"❌ Permission denied – cannot write to {excel_path}. Check folder permissions.")
    except Exception as e:
        print(f"❌ Failed to save Excel: {e}")

# ============================================================================
# 7. GENERATE DYNAMIC MODEL DATA TABLE FOR THE PLOT (unchanged, but we keep it)
# ============================================================================

def generate_model_data(nodes, supports, members_data, num_nodes, total_dofs, num_free, fixed_dofs):
    pinned = [str(nid) for nid, stype in supports.items() if stype == 'pinned']
    lines = []
    lines.append("MODEL DATA - REV. 1")
    lines.append("")
    lines.append("Geometry")
    lines.append(f"    Cube edge    6.0 m")
    lines.append(f"    Nodes    {num_nodes}")
    lines.append(f"    Members    {len(members_data)}")
    lines.append("    Vertical axis   global Y")
    lines.append("")
    lines.append("Supports")
    lines.append(f"    Type    pinned")
    lines.append(f"    Nodes    {', '.join(pinned)}")
    lines.append("    Restrained    UX, UY, UZ")
    lines.append("    Released    RX, RY, RZ")
    lines.append("")
    lines.append("Degrees of freedom")
    lines.append(f"    DOF per node    6")
    lines.append(f"    Total DOF    {total_dofs}")
    lines.append(f"    Restrained DOF    {len(fixed_dofs)}")
    lines.append(f"    Active DOF    {num_free}")
    lines.append("    Numbering    (node - 1) * 6 + 1..6")
    lines.append("")
    lines.append("Beta angles")
    lines.append(f"    Base Beam    0 deg")
    lines.append(f"    Roof Beam    0 deg")
    lines.append(f"    Column    90 deg")
    lines.append("")
    lines.append("Local axes")
    lines.append("    local x    start node i to end node j")
    lines.append("    local y    in the vertical plane, upward")
    lines.append("    local z    completes the right-handed set")
    lines.append("    Vertical members follow the limiting")
    lines.append("    case: local z parallel to global Z.")
    return "\n".join(lines)

model_data_text = generate_model_data(nodes, supports, members_data, num_nodes,
                                      total_dofs, num_free, fixed_dofs)

# ============================================================================
# 8. PLOTTING – TABLE ON THE RIGHT (same as before)
# ============================================================================

fig = plt.figure(figsize=(16, 10))
gs = gridspec.GridSpec(1, 2, width_ratios=[2, 1])

# ---- Left: 3D Cube ----
ax = fig.add_subplot(gs[0], projection='3d')

for n1, n2, beta, rel_s, rel_e in members:
    p1 = np.array(nodes[n1])
    p2 = np.array(nodes[n2])
    ax.plot([p1[0], p2[0]], [p1[1], p2[1]], [p1[2], p2[2]],
            color='blue', linewidth=2.5, zorder=5)

xs = [nodes[n][0] for n in node_ids]
ys = [nodes[n][1] for n in node_ids]
zs = [nodes[n][2] for n in node_ids]
ax.scatter(xs, ys, zs, color='red', s=150, zorder=10, edgecolors='black')

for nid in node_ids:
    if nid == 8:  # hide node 8 to match diagram
        continue
    x, y, z = nodes[nid]
    ax.text(x, y, z, f'  {nid}', color='black', fontsize=12, fontweight='bold')

for nid, stype in supports.items():
    if stype == 'pinned':
        x, y, z = nodes[nid]
        size = 0.3
        tri = np.array([[0, size, 0], [size, -size, 0], [-size, -size, 0], [0, size, 0]])
        tri += np.array([x, y, z])
        ax.plot(tri[:,0], tri[:,1], tri[:,2], color='green', linewidth=2)

# Local axes on beam 1-2
n1, n2 = 1, 2
p1 = np.array(nodes[n1])
p2 = np.array(nodes[n2])
mid = (p1 + p2) / 2
ex = (p2 - p1) / np.linalg.norm(p2 - p1)
ref = np.array([0, 1, 0])
ey_unnorm = np.cross(ex, ref)
ey = ey_unnorm / np.linalg.norm(ey_unnorm)
beta_rad = 0
v = ey
k = ex
cosb = cos(beta_rad)
sinb = sin(beta_rad)
ey_rot = v * cosb + np.cross(k, v) * sinb + k * (np.dot(k, v)) * (1 - cosb)
ey = ey_rot / np.linalg.norm(ey_rot)
ez = np.cross(ex, ey)
scale = 0.8
ax.quiver(mid[0], mid[1], mid[2], ex[0]*scale, ex[1]*scale, ex[2]*scale,
          color='cyan', arrow_length_ratio=0.2, linewidth=2)
ax.text(mid[0]+ex[0]*scale*1.1, mid[1]+ex[1]*scale*1.1, mid[2]+ex[2]*scale*1.1,
        'Local x', color='cyan', fontsize=10, fontweight='bold')
ax.quiver(mid[0], mid[1], mid[2], ey[0]*scale, ey[1]*scale, ey[2]*scale,
          color='magenta', arrow_length_ratio=0.2, linewidth=2)
ax.text(mid[0]+ey[0]*scale*1.1, mid[1]+ey[1]*scale*1.1, mid[2]+ey[2]*scale*1.1,
        'Local y', color='magenta', fontsize=10, fontweight='bold')

# Annotations
col_pos = np.array(nodes[1]) + np.array([-0.8, 3, 0])
ax.text(col_pos[0], col_pos[1], col_pos[2], 'Column', color='black', fontsize=12, fontweight='bold', ha='center')
beam_pos = np.array(nodes[1]) + np.array([3, 0.5, 0])
ax.text(beam_pos[0], beam_pos[1], beam_pos[2], 'Beam', color='black', fontsize=12, fontweight='bold', ha='center')
sup_pos = np.array(nodes[1]) + np.array([0, -1.2, 0])
ax.text(sup_pos[0], sup_pos[1], sup_pos[2], 'Supposition of Movement', color='green', fontsize=10, fontweight='bold', ha='center')
free_pos = np.array(nodes[5]) + np.array([0, 0.8, 0])
ax.text(free_pos[0], free_pos[1], free_pos[2], 'Freebody', color='red', fontsize=10, fontweight='bold', ha='center')

# Global axes
arrow_length = 7.5
ax.quiver(0, 0, 0, arrow_length, 0, 0, color='black', arrow_length_ratio=0.1, linewidth=2.5)
ax.text(arrow_length, 0, 0, ' X', color='black', fontsize=16, fontweight='bold')
ax.quiver(0, 0, 0, 0, arrow_length, 0, color='black', arrow_length_ratio=0.1, linewidth=2.5)
ax.text(0, arrow_length, 0, ' Y', color='black', fontsize=16, fontweight='bold')
ax.quiver(0, 0, 0, 0, 0, arrow_length, color='black', arrow_length_ratio=0.1, linewidth=2.5)
ax.text(0, 0, arrow_length, ' Z', color='black', fontsize=16, fontweight='bold')

ax.set_xlim([-1, 7.5])
ax.set_ylim([-1, 7.5])
ax.set_zlim([-1, 7.5])
ax.set_xlabel('X (m)')
ax.set_ylabel('Y (m)')
ax.set_zlabel('Z (m)')
ax.set_box_aspect([1, 1, 1])
ax.set_title('6m × 6m × 6m Cube – Structural Model, Rev. 1', fontsize=14)

# ---- Right: Model Data table ----
ax_text = fig.add_subplot(gs[1])
ax_text.axis('off')
ax_text.set_xlim(0, 1)
ax_text.set_ylim(0, 1)
ax_text.text(0.05, 0.95, model_data_text, transform=ax_text.transAxes,
             fontsize=10, verticalalignment='top', family='monospace', linespacing=1.5)

plt.tight_layout()
plt.show()

# ============================================================================
# 9. ANALYSIS SUMMARY
# ============================================================================
print("\n--- Analysis Summary ---")
print(f"Nodes: {num_nodes}, Total DOF: {total_dofs}, Free DOF: {num_free}")
print("Displacements (free DOFs):")
for eq, dof in enumerate(free_dofs):
    print(f"  DOF {dof}: {U_free[eq]:.6e} m/rad")
print("Member forces – see Excel for details.")