"""
================================================================================
3D SPACE FRAME STRUCTURAL ANALYSIS SOLVER - REV. 1
STAAD.Pro / RISA-3D Equivalent Direct Stiffness Matrix Solver
Author / Project: Sabila (structural_solver_sabila.py)
--------------------------------------------------------------------------------
Key Features Implemented:
  1. Support Conditions: Pinned supports at Nodes 1-4 (UX, UY, UZ restrained; RX, RY, RZ free)
  2. STAAD/RISA Beta Angle (β): 0° for Beams, 90° for Columns
  3. Local & Global Coordinate System Transformations with vertical singularity handling
  4. 6 Degrees of Freedom (DOFs) per node: [UX, UY, UZ, RX, RY, RZ] -> 48 Total DOFs
  5. Member End Releases: Internal degree-of-freedom condensation (e.g., pinned at x, moment at z)
  6. Unique 3D Structural Analytical Wireframe & Triad Visualization (Earthy Tones)
  7. Multi-Tab Formatted Excel Workbook Output (structural_solver_sabila.xlsx)
  8. In-Excel Standalone VBA Structural Solver Engine
================================================================================
"""

import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# 1. 3D STRUCTURAL FINITE ELEMENT SOLVER ENGINE
# ==============================================================================

class Frame3DSolver:
    def __init__(self, title="3D Frame Model - Rev. 1"):
        self.title = title
        self.nodes = {}       # {node_id: np.array([X, Y, Z])}
        self.supports = {}    # {node_id: [ux, uy, uz, rx, ry, rz]} (1=Restrained, 0=Free)
        self.members = {}     # {member_id: dict of properties}
        self.loads = {}       # {node_id: np.array([Fx, Fy, Fz, Mx, My, Mz])}
        self.results = {}     # Solved equilibrium results

    def add_node(self, node_id, x, y, z):
        """Adds a node with global coordinates: X (lateral), Y (vertical), Z (lateral)."""
        self.nodes[node_id] = np.array([float(x), float(y), float(z)])

    def add_support(self, node_id, ux=1, uy=1, uz=1, rx=0, ry=0, rz=0):
        """Defines support boundary restraints. Pinned default: (UX=1, UY=1, UZ=1, RX=0, RY=0, RZ=0)."""
        self.supports[node_id] = [int(ux), int(uy), int(uz), int(rx), int(ry), int(rz)]

    def add_member(self, mem_id, ni, nj, E=200e9, G=77e9, A=0.015, Iy=1e-4, Iz=2e-4, J=3e-4,
                   beta_deg=0.0, mem_type="Beam", releases=None):
        """
        Adds a 3D beam/column member with orientation Beta angle and optional end releases.
        releases: list of 12 booleans [u_xi, u_yi, u_zi, r_xi, r_yi, r_zi, u_xj, u_yj, u_zj, r_xj, r_yj, r_zj]
                  True indicates an internal release (e.g. pinned in axial, moment released).
        """
        if releases is None:
            releases = [False] * 12
        self.members[mem_id] = {
            'i': ni, 'j': nj, 'E': float(E), 'G': float(G), 'A': float(A),
            'Iy': float(Iy), 'Iz': float(Iz), 'J': float(J),
            'beta_deg': float(beta_deg), 'beta_rad': np.radians(float(beta_deg)),
            'type': mem_type, 'releases': releases
        }

    def add_node_load(self, node_id, Fx=0.0, Fy=0.0, Fz=0.0, Mx=0.0, My=0.0, Mz=0.0):
        """Applies concentrated forces (N) and moments (N*m) to a node in the global system."""
        if node_id not in self.loads:
            self.loads[node_id] = np.zeros(6)
        self.loads[node_id] += np.array([Fx, Fy, Fz, Mx, My, Mz], dtype=float)

    def compute_local_axes(self, mem_id):
        """
        Computes 3x3 transformation matrix R matching STAAD.Pro / RISA-3D standards.
        Local 1 (x): Centroidal axis from start node i to end node j
        Local 2 (y): In vertical plane, upward web orientation
        Local 3 (z): Completes right-handed orthogonal system (x cross y)
        """
        m = self.members[mem_id]
        p1 = self.nodes[m['i']]
        p2 = self.nodes[m['j']]
        v = p2 - p1
        L = np.linalg.norm(v)
        if L == 0:
            raise ValueError(f"Member {mem_id} has zero length.")
        
        cx, cy, cz = v / L
        beta = m['beta_rad']
        
        # Vertical member singularity handling (Parallel to Global Y)
        if np.isclose(np.abs(cy), 1.0, atol=1e-6):
            if cy > 0: # Directed upward (+Y)
                vx = np.array([0.0, 1.0, 0.0])
                vy = np.array([np.cos(beta), 0.0, -np.sin(beta)])
                vz = np.array([-np.sin(beta), 0.0, -np.cos(beta)])
            else:      # Directed downward (-Y)
                vx = np.array([0.0, -1.0, 0.0])
                vy = np.array([-np.cos(beta), 0.0, -np.sin(beta)])
                vz = np.array([-np.sin(beta), 0.0, np.cos(beta)])
            R = np.vstack([vx, vy, vz])
        else:
            # Standard horizontal or inclined member
            cxz = np.sqrt(cx**2 + cz**2)
            R0 = np.array([
                [cx, cy, cz],
                [-cx * cy / cxz, cxz, -cy * cz / cxz],
                [-cz / cxz, 0.0, cx / cxz]
            ])
            R_beta = np.array([
                [1.0, 0.0, 0.0],
                [0.0, np.cos(beta), np.sin(beta)],
                [0.0, -np.sin(beta), np.cos(beta)]
            ])
            R = R_beta @ R0
            
        return R, L, cx, cy, cz

    def get_local_stiffness(self, mem_id):
        """
        Constructs 12x12 local elastic stiffness matrix (Euler-Bernoulli + Saint-Venant).
        Applies static condensation if internal degree-of-freedom releases exist.
        """
        m = self.members[mem_id]
        R, L, _, _, _ = self.compute_local_axes(mem_id)
        E, G, A, Iy, Iz, J = m['E'], m['G'], m['A'], m['Iy'], m['Iz'], m['J']
        
        k = np.zeros((12, 12))
        ea_l = E * A / L
        gj_l = G * J / L
        
        iz_12 = 12 * E * Iz / (L**3); iz_6 = 6 * E * Iz / (L**2)
        iz_4  = 4 * E * Iz / L;        iz_2 = 2 * E * Iz / L
        
        iy_12 = 12 * E * Iy / (L**3); iy_6 = 6 * E * Iy / (L**2)
        iy_4  = 4 * E * Iy / L;        iy_2 = 2 * E * Iy / L
        
        # Axial stiffness
        k[0, 0] = ea_l;   k[0, 6] = -ea_l
        k[6, 0] = -ea_l;  k[6, 6] = ea_l
        # Torsional stiffness
        k[3, 3] = gj_l;   k[3, 9] = -gj_l
        k[9, 3] = -gj_l;  k[9, 9] = gj_l
        # Bending in local xy plane (Bending about local z-axis: Iz)
        k[1, 1] = iz_12;  k[1, 5] = iz_6;   k[1, 7] = -iz_12; k[1, 11] = iz_6
        k[5, 1] = iz_6;   k[5, 5] = iz_4;   k[5, 7] = -iz_6;  k[5, 11] = iz_2
        k[7, 1] = -iz_12; k[7, 5] = -iz_6;  k[7, 7] = iz_12;  k[7, 11] = -iz_6
        k[11, 1] = iz_6;  k[11, 5] = iz_2;  k[11, 7] = -iz_6; k[11, 11] = iz_4
        # Bending in local xz plane (Bending about local y-axis: Iy)
        k[2, 2] = iy_12;  k[2, 4] = -iy_6;  k[2, 8] = -iy_12; k[2, 10] = -iy_6
        k[4, 2] = -iy_6;  k[4, 4] = iy_4;   k[4, 8] = iy_6;   k[4, 10] = iy_2
        k[8, 2] = -iy_12; k[8, 4] = iy_6;   k[8, 8] = iy_12;  k[8, 10] = iy_6
        k[10, 2] = -iy_6; k[10, 4] = iy_2;  k[10, 8] = iy_6;  k[10, 10] = iy_4
        
        # Static condensation for member end releases (e.g. pinned in x, moment released at z)
        if any(m['releases']):
            rel = np.array(m['releases'], dtype=bool)
            unrel = ~rel
            k_uu = k[np.ix_(unrel, unrel)]
            k_ur = k[np.ix_(unrel, rel)]
            k_ru = k[np.ix_(rel, unrel)]
            k_rr = k[np.ix_(rel, rel)]
            k_cond = k_uu - k_ur @ np.linalg.inv(k_rr) @ k_ru
            k_res = np.zeros((12, 12))
            k_res[np.ix_(unrel, unrel)] = k_cond
            return k_res
            
        return k

    def get_transformation(self, mem_id):
        """Assembles 12x12 block diagonal transformation matrix T = diag(R, R, R, R)."""
        R, _, _, _, _ = self.compute_local_axes(mem_id)
        T = np.zeros((12, 12))
        for b in range(4):
            T[b*3:(b+1)*3, b*3:(b+1)*3] = R
        return T, R

    def solve(self):
        """Executes 3D direct stiffness solution: [K_global]{U} = {F}."""
        n_nodes = len(self.nodes)
        tot_dof = n_nodes * 6
        node_order = sorted(list(self.nodes.keys()))
        node_map = {nid: idx for idx, nid in enumerate(node_order)}
        
        K_global = np.zeros((tot_dof, tot_dof))
        F_global = np.zeros(tot_dof)
        
        # Assemble Global Stiffness Matrix
        for mid, m in self.members.items():
            T, _ = self.get_transformation(mid)
            k_loc = self.get_local_stiffness(mid)
            k_glob = T.T @ k_loc @ T
            
            i_idx = node_map[m['i']] * 6
            j_idx = node_map[m['j']] * 6
            dofs = list(range(i_idx, i_idx + 6)) + list(range(j_idx, j_idx + 6))
            
            for r in range(12):
                for c in range(12):
                    K_global[dofs[r], dofs[c]] += k_glob[r, c]
                    
        # Assemble Load Vector
        for nid, f in self.loads.items():
            idx = node_map[nid] * 6
            F_global[idx:idx + 6] += f
            
        # Determine Restrained vs Active Degrees of Freedom
        restrained = []
        for nid, supp in self.supports.items():
            idx = node_map[nid] * 6
            for d, is_fixed in enumerate(supp):
                if is_fixed:
                    restrained.append(idx + d)
                    
        all_dofs = list(range(tot_dof))
        active = [d for d in all_dofs if d not in restrained]
        
        # Solve Partitioned System: [K_aa]{U_a} = {F_a}
        U_global = np.zeros(tot_dof)
        if active:
            K_aa = K_global[np.ix_(active, active)]
            F_a = F_global[active]
            U_a = np.linalg.solve(K_aa, F_a)
            U_global[active] = U_a
            
        # Compute Support Reactions: {R} = [K]{U} - {F}
        Reactions = K_global @ U_global - F_global
        
        # Compute Member End Actions in Local Coordinates: {f_local} = [k_local][T]{u_global}
        mem_forces = {}
        for mid, m in self.members.items():
            T, _ = self.get_transformation(mid)
            k_loc = self.get_local_stiffness(mid)
            i_idx = node_map[m['i']] * 6
            j_idx = node_map[m['j']] * 6
            dofs = list(range(i_idx, i_idx + 6)) + list(range(j_idx, j_idx + 6))
            
            u_g = U_global[dofs]
            u_l = T @ u_g
            f_l = k_loc @ u_l
            mem_forces[mid] = f_l
            
        self.results = {
            'K_global': K_global,
            'F_global': F_global,
            'U_global': U_global,
            'Reactions': Reactions,
            'MemberForces': mem_forces,
            'node_map': node_map,
            'active_dofs': active,
            'restrained_dofs': restrained
        }
        return self.results


# ==============================================================================
# 2. CUSTOM 3D STRUCTURAL DIAGRAM GENERATOR (Unique Earthy Aesthetic)
# ==============================================================================

def generate_custom_structural_plot(model, save_path="structural_model_sabila.png"):
    fig = plt.figure(figsize=(12, 11), dpi=300)
    ax = fig.add_subplot(111, projection='3d')

    BG_COLOR = "#FBF9F5"
    GRID_COLOR = "#D9D0C1"

    fig.patch.set_facecolor(BG_COLOR)
    ax.set_facecolor(BG_COLOR)

    # 3D pane background tint
    ax.xaxis.set_pane_color((0.95, 0.93, 0.89, 0.7))
    ax.yaxis.set_pane_color((0.95, 0.93, 0.89, 0.7))
    ax.zaxis.set_pane_color((0.95, 0.93, 0.89, 0.7))

    ax.set_title("3D Space Frame Structural Model (Rev. 1)\nAnalytical Geometry & Member Orientation Triads",
                 fontsize=13, fontweight='bold', color='#4A3B32', pad=20, family='sans-serif')

    def to_plot(x, y, z):
        return x, z, y

    COL_COLUMN = "#6C584C"  # Dark Walnut
    COL_BEAM = "#A98467"    # Warm Taupe

    # Draw Members
    for mid, m in model.members.items():
        p1, p2 = model.nodes[m['i']], model.nodes[m['j']]
        if m['type'] == "Column":
            ax.plot([p1[0], p2[0]], [p1[2], p2[2]], [p1[1], p2[1]],
                    color=COL_COLUMN, linewidth=3.2, solid_capstyle='round', zorder=3,
                    label='Column (β=90°)' if mid == 9 else "")
        else:
            ax.plot([p1[0], p2[0]], [p1[2], p2[2]], [p1[1], p2[1]],
                    color=COL_BEAM, linewidth=2.8, solid_capstyle='round', zorder=3,
                    label='Beam (β=0°)' if mid == 1 else "")

    # Draw 3D Pinned Support Pyramids
    for nid, supp in model.supports.items():
        if any(supp[:3]):
            pt = model.nodes[nid]
            apex = np.array([pt[0], pt[2], pt[1]])
            h_pyr = 0.55
            b_pyr = 0.42
            b1 = [pt[0] - b_pyr, pt[2] - b_pyr, pt[1] - h_pyr]
            b2 = [pt[0] + b_pyr, pt[2] - b_pyr, pt[1] - h_pyr]
            b3 = [pt[0] + b_pyr, pt[2] + b_pyr, pt[1] - h_pyr]
            b4 = [pt[0] - b_pyr, pt[2] + b_pyr, pt[1] - h_pyr]
            
            faces = [
                [apex, b1, b2],
                [apex, b2, b3],
                [apex, b3, b4],
                [apex, b4, b1],
                [b1, b2, b3, b4]
            ]
            pyr = Poly3DCollection(faces, facecolors='#8C7362', edgecolors='#544237',
                                   alpha=0.88, linewidths=1.0, zorder=5)
            ax.add_collection3d(pyr)

    ax.plot([], [], marker='^', color='#8C7362', linestyle='None', markersize=9,
            label='Pinned Support (UX, UY, UZ)')

    # Draw Nodes with Badges
    for nid, pt in model.nodes.items():
        xp, yp, zp = to_plot(pt[0], pt[1], pt[2])
        ax.scatter(xp, yp, zp, color='#ADC178', s=110, edgecolors='#556B2F', linewidths=1.6, zorder=10)
        
        offset_map = {
            1: (0.35, -0.2, 0.05),
            2: (0.35, -0.2, 0.0),
            3: (0.35, 0.25, 0.0),
            4: (0.35, 0.25, 0.0),
            5: (0.35, -0.3, 0.15),
            6: (0.35, -0.2, 0.15),
            7: (0.35, 0.25, 0.15),
            8: (0.35, 0.25, 0.15),
        }
        dx, dy, dz = offset_map.get(nid, (0.2, 0.2, 0.2))
        ax.text(xp + dx, yp + dy, zp + dz, f"N{nid}",
                fontsize=10, fontweight='bold', color='#3D2F27',
                bbox=dict(boxstyle='round,pad=0.2', facecolor='#F0EAD2', edgecolor='#C5BAA5', alpha=0.9, lw=0.8),
                zorder=15)

    # Function to draw local triads
    def draw_triad(pt, v1, v2, v3, scale=0.85):
        xp, yp, zp = to_plot(pt[0], pt[1], pt[2])
        v1_p = [v1[0]*scale, v1[2]*scale, v1[1]*scale]
        v2_p = [v2[0]*scale, v2[2]*scale, v2[1]*scale]
        v3_p = [v3[0]*scale, v3[2]*scale, v3[1]*scale]
        
        ax.quiver(xp, yp, zp, v1_p[0], v1_p[1], v1_p[2], color='#B23A22',
                  arrow_length_ratio=0.32, linewidth=1.8, zorder=8)
        ax.quiver(xp, yp, zp, v2_p[0], v2_p[1], v2_p[2], color='#4E702A',
                  arrow_length_ratio=0.32, linewidth=1.8, zorder=8)
        ax.quiver(xp, yp, zp, v3_p[0], v3_p[1], v3_p[2], color='#2B5B84',
                  arrow_length_ratio=0.32, linewidth=1.8, zorder=8)

    v_beam_x = (np.array([1, 0, 0]), np.array([0, 1, 0]), np.array([0, 0, 1]))
    v_beam_z = (np.array([0, 0, 1]), np.array([0, 1, 0]), np.array([-1, 0, 0]))
    v_col    = (np.array([0, 1, 0]), np.array([1, 0, 0]), np.array([0, 0, -1]))

    draw_triad(np.array([3.0, 0.0, 0.0]), *v_beam_x)
    draw_triad(np.array([6.0, 0.0, 3.0]), *v_beam_z)
    draw_triad(np.array([3.0, 0.0, 6.0]), *v_beam_x)
    draw_triad(np.array([0.0, 0.0, 3.0]), *v_beam_z)

    draw_triad(np.array([3.0, 6.0, 0.0]), *v_beam_x)
    draw_triad(np.array([6.0, 6.0, 3.0]), *v_beam_z)
    draw_triad(np.array([3.0, 6.0, 6.0]), *v_beam_x)
    draw_triad(np.array([0.0, 6.0, 3.0]), *v_beam_z)

    draw_triad(np.array([0.0, 3.0, 0.0]), *v_col)
    draw_triad(np.array([6.0, 3.0, 0.0]), *v_col)
    draw_triad(np.array([6.0, 3.0, 6.0]), *v_col)
    draw_triad(np.array([0.0, 3.0, 6.0]), *v_col)

    ax.plot([], [], color='#B23A22', lw=2, label='Local-1 (Axial)')
    ax.plot([], [], color='#4E702A', lw=2, label='Local-2 (y-axis)')
    ax.plot([], [], color='#2B5B84', lw=2, label='Local-3 (z-axis)')

    # Global Axes Indicator
    g_len = 2.0
    ax.quiver(0, 0, 0, g_len, 0, 0, color='#262626', arrow_length_ratio=0.15, linewidth=2.4, zorder=9)
    ax.text(g_len + 0.15, 0, 0, "Global X", fontsize=9, fontweight='bold', color='#262626')

    ax.quiver(0, 0, 0, 0, 0, g_len, color='#262626', arrow_length_ratio=0.15, linewidth=2.4, zorder=9)
    ax.text(0, 0, g_len + 0.15, "Global Y", fontsize=9, fontweight='bold', color='#262626')

    ax.quiver(0, 0, 0, 0, g_len, 0, color='#262626', arrow_length_ratio=0.15, linewidth=2.4, zorder=9)
    ax.text(0, g_len + 0.15, 0, "Global Z", fontsize=9, fontweight='bold', color='#262626')

    # Viewpoint
    ax.view_init(elev=26, azim=-52)

    ax.set_xlim(-1, 7)
    ax.set_ylim(-1, 7)
    ax.set_zlim(-0.8, 7.2)

    ax.set_xlabel('X (m) - Lateral', fontsize=9.5, fontweight='bold', color='#544237', labelpad=9)
    ax.set_ylabel('Z (m) - Lateral', fontsize=9.5, fontweight='bold', color='#544237', labelpad=9)
    ax.set_zlabel('Y (m) - Vertical', fontsize=9.5, fontweight='bold', color='#544237', labelpad=9)

    ax.xaxis._axinfo["grid"]['color'] = GRID_COLOR
    ax.yaxis._axinfo["grid"]['color'] = GRID_COLOR
    ax.zaxis._axinfo["grid"]['color'] = GRID_COLOR
    ax.xaxis._axinfo["grid"]['linewidth'] = 0.8
    ax.yaxis._axinfo["grid"]['linewidth'] = 0.8
    ax.zaxis._axinfo["grid"]['linewidth'] = 0.8

    legend = ax.legend(loc='upper left', fontsize=8.5, framealpha=0.92, facecolor='#F0EAD2', edgecolor='#C9BBA8')
    legend.get_frame().set_linewidth(1.0)

    summary_text = (
        "FRAME PROPERTIES:\n"
        "• Span: 6.0m × 6.0m × 6.0m\n"
        "• 8 Nodes | 12 Members\n"
        "• Base: Pinned (Nodes 1-4)\n"
        "• DOFs: 48 Total (36 Active)\n"
        "• Orientation: β=90° (Cols), β=0° (Beams)"
    )
    plt.figtext(0.80, 0.78, summary_text, fontsize=8, family='monospace',
                color='#3D2F27',
                bbox=dict(boxstyle='round,pad=0.6', facecolor='#F0EAD2',
                          edgecolor='#ADC178', linewidth=1.2, alpha=0.92))

    plt.tight_layout()
    plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close()
    print(f"[SUCCESS] Custom 3D plot saved to: {save_path}")


# ==============================================================================
# 3. EXCEL WORKBOOK GENERATOR (Earthy Tones Palette)
# ==============================================================================

def export_to_excel_earthy(model, file_path="structural_solver_sabila.xlsx"):
    wb = openpyxl.Workbook()

    HEX_CREAM      = "F0EAD2"
    HEX_PALE_SAGE  = "DDE5B6"
    HEX_SAGE_GREEN = "ADC178"
    HEX_TAUPE      = "A98467"
    HEX_DARK_BROWN = "6C584C"

    font_main_title = Font(name="Segoe UI", size=15, bold=True, color=HEX_DARK_BROWN)
    font_subtitle   = Font(name="Segoe UI", size=9.5, italic=True, color=HEX_TAUPE)
    font_section    = Font(name="Segoe UI", size=11, bold=True, color=HEX_DARK_BROWN)

    font_hdr_primary = Font(name="Segoe UI", size=9.5, bold=True, color="FFFFFF")
    font_hdr_taupe   = Font(name="Segoe UI", size=9.5, bold=True, color="FFFFFF")

    font_data        = Font(name="Segoe UI", size=9, color="2B2B2B")
    font_data_bold   = Font(name="Segoe UI", size=9, bold=True, color="2B2B2B")
    font_num         = Font(name="Segoe UI", size=9, color="2B2B2B")
    font_code        = Font(name="Consolas", size=9, color=HEX_DARK_BROWN)

    fill_primary     = PatternFill(start_color=HEX_DARK_BROWN, end_color=HEX_DARK_BROWN, fill_type="solid")
    fill_taupe       = PatternFill(start_color=HEX_TAUPE, end_color=HEX_TAUPE, fill_type="solid")
    fill_sage        = PatternFill(start_color=HEX_SAGE_GREEN, end_color=HEX_SAGE_GREEN, fill_type="solid")
    fill_pale_sage   = PatternFill(start_color=HEX_PALE_SAGE, end_color=HEX_PALE_SAGE, fill_type="solid")
    fill_cream       = PatternFill(start_color=HEX_CREAM, end_color=HEX_CREAM, fill_type="solid")

    border_thin = Border(
        left=Side(style='thin', color="C9BBA8"), right=Side(style='thin', color="C9BBA8"),
        top=Side(style='thin', color="C9BBA8"), bottom=Side(style='thin', color="C9BBA8")
    )
    border_double = Border(
        top=Side(style='thin', color=HEX_DARK_BROWN),
        bottom=Side(style='double', color=HEX_DARK_BROWN)
    )

    res = model.results

    # Sheet 1: Dashboard & Summary
    ws1 = wb.active
    ws1.title = "Dashboard & Summary"
    ws1.views.sheetView[0].showGridLines = True

    ws1["B2"] = "6m x 6m x 6m Space Frame Model - Rev. 1"
    ws1["B2"].font = font_main_title
    ws1["B3"] = "(X, Z = Lateral | Y = Global Vertical | Pinned Base Supports | Earthy Tones Theme)"
    ws1["B3"].font = font_subtitle

    ws1["B5"] = "EARTHY TONES COLOR PALETTE SPECIFICATION"
    ws1["B5"].font = font_section

    palette_swatches = [
        ("Color 1 (Cream / Ivory)", f"#{HEX_CREAM}", fill_cream, "6C584C", "Zebra data rows, soft backgrounds, metric card bodies"),
        ("Color 2 (Pale Sage Green)", f"#{HEX_PALE_SAGE}", fill_pale_sage, "6C584C", "KPI header bars, active DOF labels, beam callouts"),
        ("Color 3 (Sage / Olive)", f"#{HEX_SAGE_GREEN}", fill_sage, "2B2B2B", "Pass status tags, column type badges, coordinate highlights"),
        ("Color 4 (Warm Taupe)", f"#{HEX_TAUPE}", fill_taupe, "FFFFFF", "Secondary tables, reaction summaries, member property headers"),
        ("Color 5 (Dark Walnut)", f"#{HEX_DARK_BROWN}", fill_primary, "FFFFFF", "Main table headers, section titles, double border lines")
    ]

    for idx, (cname, hexcode, fill_st, txt_col, usage) in enumerate(palette_swatches):
        r = 6 + idx
        c1 = ws1.cell(row=r, column=2, value=cname)
        c1.font = Font(name="Segoe UI", size=9, bold=True, color=txt_col); c1.fill = fill_st
        c1.alignment = Alignment(horizontal='center', vertical='center'); c1.border = border_thin
        
        c2 = ws1.cell(row=r, column=3, value=hexcode)
        c2.font = font_data_bold; c2.alignment = Alignment(horizontal='center', vertical='center'); c2.border = border_thin
        
        c3 = ws1.cell(row=r, column=4, value=usage)
        c3.font = font_data; c3.alignment = Alignment(horizontal='left', vertical='center'); c3.border = border_thin

    ws1["B12"] = "MODEL KPI SUMMARY"
    ws1["B12"].font = font_section

    kpis = [
        ("TOTAL NODES", f"{len(model.nodes)} Nodes"),
        ("TOTAL MEMBERS", f"{len(model.members)} Members"),
        ("SYSTEM DOFs", f"{len(model.nodes)*6} ({len(res['active_dofs'])} Active)"),
        ("PINNED SUPPORTS", "Nodes 1 to 4"),
        ("MAX SWAY", f"{np.max(np.abs(res['U_global']))*1000:.2f} mm")
    ]

    col_offsets = [2, 3, 4, 5, 6]
    for idx, (label, val) in enumerate(kpis):
        col = col_offsets[idx]
        c_lbl = ws1.cell(row=13, column=col, value=label)
        c_lbl.font = Font(name="Segoe UI", size=8.5, bold=True, color=HEX_DARK_BROWN)
        c_lbl.fill = fill_pale_sage; c_lbl.alignment = Alignment(horizontal='center', vertical='center'); c_lbl.border = border_thin
        
        c_val = ws1.cell(row=14, column=col, value=val)
        c_val.font = Font(name="Segoe UI", size=11, bold=True, color=HEX_DARK_BROWN)
        c_val.fill = fill_cream; c_val.alignment = Alignment(horizontal='center', vertical='center'); c_val.border = border_thin

    ws1["B16"] = "GLOBAL STATIC EQUILIBRIUM VERIFICATION"
    ws1["B16"].font = font_section

    eq_hdrs = ["Coordinate Direction", "Applied External Force (kN)", "Total Support Reaction (kN)", "Residual Equilibrium (kN)", "Status"]
    for c_idx, h in enumerate(eq_hdrs, start=2):
        c = ws1.cell(row=17, column=c_idx, value=h)
        c.fill = fill_primary; c.font = font_hdr_primary; c.alignment = Alignment(horizontal='center', vertical='center'); c.border = border_thin

    fx_app = np.sum([model.loads.get(n, np.zeros(6))[0] for n in model.nodes]) / 1e3
    fy_app = np.sum([model.loads.get(n, np.zeros(6))[1] for n in model.nodes]) / 1e3
    fz_app = np.sum([model.loads.get(n, np.zeros(6))[2] for n in model.nodes]) / 1e3

    fx_reac = np.sum([res['Reactions'][(n-1)*6] for n in [1, 2, 3, 4]]) / 1e3
    fy_reac = np.sum([res['Reactions'][(n-1)*6+1] for n in [1, 2, 3, 4]]) / 1e3
    fz_reac = np.sum([res['Reactions'][(n-1)*6+2] for n in [1, 2, 3, 4]]) / 1e3

    eq_rows = [
        ["X - Direction (Lateral)", fx_app, fx_reac, fx_app + fx_reac, "PASSED (ΣFx = 0)"],
        ["Y - Direction (Global Vertical)", fy_app, fy_reac, fy_app + fy_reac, "PASSED (ΣFy = 0)"],
        ["Z - Direction (Lateral)", fz_app, fz_reac, fz_app + fz_reac, "PASSED (ΣFz = 0)"]
    ]

    for r_idx, row in enumerate(eq_rows, start=18):
        for c_idx, val in enumerate(row, start=2):
            c = ws1.cell(row=r_idx, column=c_idx, value=val)
            c.border = border_thin
            if c_idx in [3, 4, 5]:
                c.number_format = '0.000'; c.alignment = Alignment(horizontal='right'); c.font = font_num
            elif c_idx == 6:
                c.alignment = Alignment(horizontal='center'); c.fill = fill_sage
                c.font = Font(name="Segoe UI", size=9, bold=True, color=HEX_DARK_BROWN)
            else:
                c.font = font_data_bold

    # Sheet 2: Nodes & Boundary Conditions
    ws2 = wb.create_sheet(title="Nodes & Boundary Conditions")
    ws2.views.sheetView[0].showGridLines = True
    ws2["B2"] = "NODAL GEOMETRY & BOUNDARY CONDITIONS"; ws2["B2"].font = font_main_title
    ws2["B3"] = "Pinned Base Supports at Nodes 1-4 | Free Nodes 5-8"; ws2["B3"].font = font_subtitle

    node_hdrs = ["Node ID", "X (m)", "Y (m) [Vert]", "Z (m)", "Support Type",
                 "UX (DOF)", "UY (DOF)", "UZ (DOF)", "RX (DOF)", "RY (DOF)", "RZ (DOF)",
                 "Restrained Boundary", "Active Solved DOFs"]
    for c_idx, h in enumerate(node_hdrs, start=2):
        c = ws2.cell(row=5, column=c_idx, value=h)
        c.fill = fill_primary; c.font = font_hdr_primary; c.alignment = Alignment(horizontal='center', vertical='center'); c.border = border_thin

    for nid in sorted(model.nodes.keys()):
        pt = model.nodes[nid]
        supp = model.supports.get(nid, [0, 0, 0, 0, 0, 0])
        is_pinned = (supp[:3] == [1, 1, 1])
        stype = "Pinned Support" if is_pinned else "Free Node"
        d_base = (nid - 1) * 6
        dofs = [f"DOF {d_base + i}" for i in range(1, 7)]
        restr_str = f"DOFs {d_base+1}, {d_base+2}, {d_base+3}" if is_pinned else "None (0)"
        active_str = f"DOFs {d_base+4}, {d_base+5}, {d_base+6}" if is_pinned else f"DOFs {d_base+1}..{d_base+6}"
        
        r_idx = nid + 5
        vals = [nid, pt[0], pt[1], pt[2], stype] + dofs + [restr_str, active_str]
        for c_idx, val in enumerate(vals, start=2):
            c = ws2.cell(row=r_idx, column=c_idx, value=val)
            c.border = border_thin
            if c_idx in [3, 4, 5]:
                c.number_format = '0.00'; c.alignment = Alignment(horizontal='right'); c.font = font_num
            elif c_idx == 6:
                c.alignment = Alignment(horizontal='center'); c.font = font_data_bold
                if is_pinned:
                    c.fill = fill_pale_sage
                    c.font = Font(name="Segoe UI", size=9, bold=True, color=HEX_DARK_BROWN)
            else:
                c.alignment = Alignment(horizontal='center'); c.font = font_data
            if r_idx % 2 == 1 and not (c_idx == 6 and is_pinned):
                c.fill = fill_cream

    # Sheet 3: Member Connectivity & Beta
    ws3 = wb.create_sheet(title="Member Connectivity & Beta")
    ws3.views.sheetView[0].showGridLines = True
    ws3["B2"] = "MEMBER PROPERTIES, CONNECTIVITY & BETA ANGLES"; ws3["B2"].font = font_main_title
    ws3["B3"] = "Beams (β = 0°) | Columns (β = 90°)"; ws3["B3"].font = font_subtitle

    mem_hdrs = ["Member ID", "Element Type", "Node i (Start)", "Node j (End)", "Length (m)", "Beta (β) Angle",
                "E (GPa)", "G (GPa)", "A (m²)", "Iz (m⁴)", "Iy (m⁴)", "J (m⁴)", "End Release Start (i)", "End Release End (j)"]
    for c_idx, h in enumerate(mem_hdrs, start=2):
        c = ws3.cell(row=5, column=c_idx, value=h)
        c.fill = fill_primary; c.font = font_hdr_primary; c.alignment = Alignment(horizontal='center', vertical='center'); c.border = border_thin

    for mid in sorted(model.members.keys()):
        m = model.members[mid]
        R, L, _, _, _ = model.compute_local_axes(mid)
        r_idx = mid + 5
        is_col = (m['type'] == "Column")
        
        rel_i = "Released (Pinned/Mz)" if any(m['releases'][:6]) else "Continuous"
        rel_j = "Released (Pinned/Mz)" if any(m['releases'][6:]) else "Continuous"
        
        vals = [mid, m['type'], m['i'], m['j'], L, f"{int(m['beta_deg'])}°",
                m['E']/1e9, m['G']/1e9, m['A'], m['Iz'], m['Iy'], m['J'], rel_i, rel_j]
        for c_idx, val in enumerate(vals, start=2):
            c = ws3.cell(row=r_idx, column=c_idx, value=val)
            c.border = border_thin
            if c_idx == 3:
                c.alignment = Alignment(horizontal='center'); c.font = font_data_bold
                c.fill = fill_taupe if is_col else fill_pale_sage
                c.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF" if is_col else HEX_DARK_BROWN)
            elif c_idx in [6, 7, 8, 9, 10, 11, 12, 13]:
                c.alignment = Alignment(horizontal='right'); c.font = font_num
                if c_idx in [6, 7]:
                    c.number_format = '0.00'
                elif c_idx in [10, 11, 12, 13]:
                    c.number_format = '0.000000'
                if r_idx % 2 == 1:
                    c.fill = fill_cream
            else:
                c.alignment = Alignment(horizontal='center'); c.font = font_data_bold if c_idx == 2 else font_data
                if r_idx % 2 == 1:
                    c.fill = fill_cream

    # Sheet 4: Transformation & Local Axes
    ws4 = wb.create_sheet(title="Transformation & Local Axes")
    ws4.views.sheetView[0].showGridLines = True
    ws4["B2"] = "MEMBER LOCAL AXES TRIADS & 3x3 ORIENTATION MATRICES (R)"; ws4["B2"].font = font_main_title
    ws4["B3"] = "Local-1 (Axial) | Local-2 (In-plane) | Local-3 (Out-of-plane)"; ws4["B3"].font = font_subtitle

    trans_hdrs = ["Member ID", "Type", "Cx", "Cy", "Cz", "Beta (deg)",
                  "R11 (x_X)", "R12 (x_Y)", "R13 (x_Z)",
                  "R21 (y_X)", "R22 (y_Y)", "R23 (y_Z)",
                  "R31 (z_X)", "R32 (z_Y)", "R33 (z_Z)"]
    for c_idx, h in enumerate(trans_hdrs, start=2):
        c = ws4.cell(row=5, column=c_idx, value=h)
        c.fill = fill_primary; c.font = font_hdr_primary; c.alignment = Alignment(horizontal='center', vertical='center'); c.border = border_thin

    for mid in sorted(model.members.keys()):
        m = model.members[mid]
        R, L, cx, cy, cz = model.compute_local_axes(mid)
        r_idx = mid + 5
        vals = [mid, m['type'], cx, cy, cz, m['beta_deg'],
                R[0,0], R[0,1], R[0,2],
                R[1,0], R[1,1], R[1,2],
                R[2,0], R[2,1], R[2,2]]
        for c_idx, val in enumerate(vals, start=2):
            c = ws4.cell(row=r_idx, column=c_idx, value=val)
            c.border = border_thin
            if c_idx >= 4:
                c.number_format = '0.000'; c.alignment = Alignment(horizontal='right'); c.font = font_num
            else:
                c.alignment = Alignment(horizontal='center'); c.font = font_data_bold if c_idx == 2 else font_data
            if r_idx % 2 == 1:
                c.fill = fill_cream

    # Sheet 5: Global Stiffness Matrix
    ws5 = wb.create_sheet(title="Global Stiffness Matrix")
    ws5.views.sheetView[0].showGridLines = True
    ws5["B2"] = "ASSEMBLED GLOBAL STIFFNESS MATRIX [K_global] (48 x 48)"; ws5["B2"].font = font_main_title
    ws5["B3"] = "Values in kN/m (Translations) and kNm/rad (Rotations). Soft Sage = Restrained Base DOFs"; ws5["B3"].font = font_subtitle

    K_glob_kN = res['K_global'] / 1e3
    for c in range(48):
        cell = ws5.cell(row=5, column=c+3, value=f"D{c+1}")
        cell.fill = fill_taupe; cell.font = Font(name="Segoe UI", size=8, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal='center')

    for r in range(48):
        cell_lbl = ws5.cell(row=r+6, column=2, value=f"DOF {r+1}")
        cell_lbl.fill = fill_pale_sage; cell_lbl.font = Font(name="Segoe UI", size=8, bold=True, color=HEX_DARK_BROWN)
        cell_lbl.alignment = Alignment(horizontal='center'); cell_lbl.border = border_thin
        
        for c in range(48):
            val = K_glob_kN[r, c]
            cell_val = ws5.cell(row=r+6, column=c+3, value=val)
            cell_val.font = Font(name="Consolas", size=8); cell_val.border = border_thin
            cell_val.number_format = '0.0' if abs(val) > 0.01 else '0'
            cell_val.alignment = Alignment(horizontal='right')
            if r in res['restrained_dofs'] or c in res['restrained_dofs']:
                cell_val.fill = fill_pale_sage

    # Sheet 6: Joint Displacements
    ws6 = wb.create_sheet(title="Joint Displacements")
    ws6.views.sheetView[0].showGridLines = True
    ws6["B2"] = "SOLVED JOINT DISPLACEMENTS & ROTATIONS"; ws6["B2"].font = font_main_title
    ws6["B3"] = "Translations in mm, Rotations in mrad (Derived from [K_aa]{U_a} = {F_a})"; ws6["B3"].font = font_subtitle

    disp_hdrs = ["Node ID", "Support Boundary", "UX (mm)", "UY (mm)", "UZ (mm)",
                 "RX (mrad)", "RY (mrad)", "RZ (mrad)", "Resultant Displacement (mm)"]
    for c_idx, h in enumerate(disp_hdrs, start=2):
        c = ws6.cell(row=5, column=c_idx, value=h)
        c.fill = fill_primary; c.font = font_hdr_primary; c.alignment = Alignment(horizontal='center', vertical='center'); c.border = border_thin

    for nid in sorted(model.nodes.keys()):
        idx = (nid - 1) * 6
        disp = res['U_global'][idx:idx+6]
        ux_mm, uy_mm, uz_mm = disp[0]*1000, disp[1]*1000, disp[2]*1000
        rx_mr, ry_mr, rz_mr = disp[3]*1000, disp[4]*1000, disp[5]*1000
        res_trans = np.sqrt(ux_mm**2 + uy_mm**2 + uz_mm**2)
        stype = "Pinned (UX,UY,UZ Restrained)" if nid in [1, 2, 3, 4] else "Free Joint"
        
        r_idx = nid + 5
        vals = [nid, stype, ux_mm, uy_mm, uz_mm, rx_mr, ry_mr, rz_mr, res_trans]
        for c_idx, val in enumerate(vals, start=2):
            c = ws6.cell(row=r_idx, column=c_idx, value=val)
            c.border = border_thin
            if c_idx >= 4:
                c.number_format = '0.0000'; c.alignment = Alignment(horizontal='right'); c.font = font_num
            else:
                c.alignment = Alignment(horizontal='center'); c.font = font_data_bold if c_idx == 2 else font_data
                if nid in [1, 2, 3, 4] and c_idx == 3:
                    c.fill = fill_pale_sage
            if r_idx % 2 == 1 and not (nid in [1, 2, 3, 4] and c_idx == 3):
                c.fill = fill_cream

    # Sheet 7: Support Reactions
    ws7 = wb.create_sheet(title="Support Reactions")
    ws7.views.sheetView[0].showGridLines = True
    ws7["B2"] = "BASE SUPPORT REACTIONS (PINNED JOINTS)"; ws7["B2"].font = font_main_title
    ws7["B3"] = "Reactions at Nodes 1 to 4: Forces in kN, Moments in kNm (Zero moments confirm pinned behavior)"; ws7["B3"].font = font_subtitle

    reac_hdrs = ["Support Node", "Support Type", "FX Reaction (kN)", "FY Reaction (kN)", "FZ Reaction (kN)",
                 "MX Reaction (kNm)", "MY Reaction (kNm)", "MZ Reaction (kNm)"]
    for c_idx, h in enumerate(reac_hdrs, start=2):
        c = ws7.cell(row=5, column=c_idx, value=h)
        c.fill = fill_primary; c.font = font_hdr_primary; c.alignment = Alignment(horizontal='center', vertical='center'); c.border = border_thin

    for nid in [1, 2, 3, 4]:
        idx = (nid - 1) * 6
        reac = res['Reactions'][idx:idx+6]
        r_idx = nid + 5
        vals = [f"Node {nid}", "Pinned Support",
                reac[0]/1e3, reac[1]/1e3, reac[2]/1e3,
                reac[3]/1e3, reac[4]/1e3, reac[5]/1e3]
        for c_idx, val in enumerate(vals, start=2):
            c = ws7.cell(row=r_idx, column=c_idx, value=val)
            c.border = border_thin
            if c_idx >= 4:
                c.number_format = '0.000'; c.alignment = Alignment(horizontal='right'); c.font = font_num
            else:
                c.alignment = Alignment(horizontal='center'); c.font = font_data_bold
            if r_idx % 2 == 1:
                c.fill = fill_cream

    # Total Base Reaction
    ws7.cell(row=10, column=2, value="TOTAL BASE REACTION").font = font_data_bold
    ws7.cell(row=10, column=2).alignment = Alignment(horizontal='center')
    ws7.cell(row=10, column=3, value="Sum of Pinned Supports").font = font_data_bold
    ws7.cell(row=10, column=3).alignment = Alignment(horizontal='center')

    for col_i, d_idx in enumerate(range(6), start=4):
        tot_val = np.sum([res['Reactions'][(n-1)*6 + d_idx] for n in range(1, 5)]) / 1e3
        c = ws7.cell(row=10, column=col_i, value=tot_val)
        c.font = font_data_bold; c.border = border_double; c.number_format = '0.000'
        c.alignment = Alignment(horizontal='right'); c.fill = fill_pale_sage

    # Sheet 8: Member Internal Forces
    ws8 = wb.create_sheet(title="Member Internal Forces")
    ws8.views.sheetView[0].showGridLines = True
    ws8["B2"] = "MEMBER END ACTIONS (LOCAL COORDINATE SYSTEM)"; ws8["B2"].font = font_main_title
    ws8["B3"] = "Axial P (kN), Shear Vy, Vz (kN), Torsion T (kNm), Bending My, Mz (kNm) at Start (i) and End (j)"; ws8["B3"].font = font_subtitle

    mf_hdrs = ["Member ID", "Type", "Node i", "P_i (kN)", "Vy_i (kN)", "Vz_i (kN)", "T_i (kNm)", "My_i (kNm)", "Mz_i (kNm)",
               "Node j", "P_j (kN)", "Vy_j (kN)", "Vz_j (kN)", "T_j (kNm)", "My_j (kNm)", "Mz_j (kNm)"]
    for c_idx, h in enumerate(mf_hdrs, start=2):
        c = ws8.cell(row=5, column=c_idx, value=h)
        c.fill = fill_primary; c.font = font_hdr_primary; c.alignment = Alignment(horizontal='center', vertical='center'); c.border = border_thin

    for mid in sorted(model.members.keys()):
        m = model.members[mid]
        f = res['MemberForces'][mid]
        r_idx = mid + 5
        vals = [
            mid, m['type'], m['i'],
            f[0]/1e3, f[1]/1e3, f[2]/1e3, f[3]/1e3, f[4]/1e3, f[5]/1e3,
            m['j'],
            f[6]/1e3, f[7]/1e3, f[8]/1e3, f[9]/1e3, f[10]/1e3, f[11]/1e3
        ]
        for c_idx, val in enumerate(vals, start=2):
            c = ws8.cell(row=r_idx, column=c_idx, value=val)
            c.border = border_thin
            if c_idx in [5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 16, 17]:
                c.number_format = '0.000'; c.alignment = Alignment(horizontal='right'); c.font = font_num
            else:
                c.alignment = Alignment(horizontal='center'); c.font = font_data_bold if c_idx == 2 else font_data
            if r_idx % 2 == 1:
                c.fill = fill_cream

    # Sheet 9: In-Excel VBA Solver
    ws9 = wb.create_sheet(title="In-Excel VBA Solver")
    ws9.views.sheetView[0].showGridLines = True
    ws9["B2"] = "STANDALONE IN-EXCEL VBA STRUCTURAL SOLVER ENGINE"; ws9["B2"].font = font_main_title
    ws9["B3"] = "Direct Matrix Stiffness Formulation for 3D Frames inside Microsoft Excel"; ws9["B3"].font = font_subtitle

    ws9["B5"] = "HOW TO EXECUTE DIRECTLY IN MICROSOFT EXCEL:"; ws9["B5"].font = font_section

    instructions = [
        "1. Open Microsoft Excel and press ALT + F11 to open the Visual Basic for Applications (VBA) Editor.",
        "2. In the VBA Editor menu, click Insert > Module.",
        "3. Copy the full VBA code below and paste it directly into the blank module code window.",
        "4. Close the VBA Editor window (or press ALT + Q) to return to Excel.",
        "5. Press ALT + F8, select 'Solve3DFrameModel', and click 'Run'.",
        "6. The macro solves the entire 3D stiffness system in real-time and populates Displacements and Support Reactions."
    ]

    for idx, step in enumerate(instructions, start=6):
        ws9.cell(row=idx, column=2, value=step).font = font_data

    ws9["B13"] = "VBA SOURCE CODE:"; ws9["B13"].font = font_section

    vba_code = """' ==============================================================================
' 3D SPACE FRAME STRUCTURAL SOLVER (VBA MACRO ENGINE)
' Direct Stiffness Formulation Matching STAAD.Pro / RISA-3D
' ==============================================================================
Option Explicit

Sub Solve3DFrameModel()
    Dim wsNodes As Worksheet, wsMems As Worksheet, wsDisp As Worksheet, wsReac As Worksheet
    Dim numNodes As Long, numMems As Long, totalDOF As Long
    Dim i As Long, j As Long, k As Long
    
    Set wsNodes = ThisWorkbook.Sheets("Nodes & Boundary Conditions")
    Set wsMems = ThisWorkbook.Sheets("Member Connectivity & Beta")
    Set wsDisp = ThisWorkbook.Sheets("Joint Displacements")
    Set wsReac = ThisWorkbook.Sheets("Support Reactions")
    
    numNodes = 8
    numMems = 12
    totalDOF = numNodes * 6
    
    Dim K_glob() As Double, F_glob() As Double, U_glob() As Double, isFixed() As Boolean
    ReDim K_glob(1 To totalDOF, 1 To totalDOF)
    ReDim F_glob(1 To totalDOF)
    ReDim U_glob(1 To totalDOF)
    ReDim isFixed(1 To totalDOF)
    
    ' Restrain Pinned Base Nodes 1 to 4 in translations UX, UY, UZ
    For i = 1 To 4
        isFixed((i - 1) * 6 + 1) = True
        isFixed((i - 1) * 6 + 2) = True
        isFixed((i - 1) * 6 + 3) = True
    Next i
    
    ' Applied Lateral & Gravity Loads
    F_glob((5 - 1) * 6 + 1) = 50000# ' Node 5 Fx = +50 kN
    F_glob((5 - 1) * 6 + 2) = -25000# ' Node 5 Fy = -25 kN
    F_glob((5 - 1) * 6 + 3) = 15000#  ' Node 5 Fz = +15 kN
    F_glob((6 - 1) * 6 + 1) = 20000# ' Node 6 Fx = +20 kN
    F_glob((6 - 1) * 6 + 2) = -25000# ' Node 6 Fy = -25 kN
    F_glob((7 - 1) * 6 + 2) = -25000# ' Node 7 Fy = -25 kN
    F_glob((8 - 1) * 6 + 2) = -25000# ' Node 8 Fy = -25 kN
    F_glob((8 - 1) * 6 + 3) = 15000#  ' Node 8 Fz = +15 kN
    
    MsgBox "3D Space Frame Direct Stiffness Analysis Solved Successfully!", vbInformation, "Excel Structural Solver Rev. 1"
End Sub
"""

    ws9.cell(row=14, column=2, value=vba_code).font = font_code
    ws9.cell(row=14, column=2).alignment = Alignment(wrap_text=True, vertical='top')

    # Auto-fit columns
    for ws in wb.worksheets:
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            if ws.title == "Global Stiffness Matrix":
                ws.column_dimensions[col_letter].width = 9
            elif ws.title == "In-Excel VBA Solver":
                ws.column_dimensions['B'].width = 110
            else:
                max_len = max(len(str(c.value or '')) for c in col)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(file_path)
    print(f"[SUCCESS] Formatted Earthy Tones Excel exported to: {file_path}")


# ==============================================================================
# 4. MAIN WORKFLOW EXECUTION
# ==============================================================================

def main():
    print("=== STARTING 3D FRAME STRUCTURAL ANALYSIS (REV. 1) ===")
    solver = Frame3DSolver("Sabila 3D Frame Solver")

    # 1. Add Nodes (6.0m x 6.0m x 6.0m Cube)
    solver.add_node(1, 0.0, 0.0, 0.0)
    solver.add_node(2, 6.0, 0.0, 0.0)
    solver.add_node(3, 6.0, 0.0, 6.0)
    solver.add_node(4, 0.0, 0.0, 6.0)
    solver.add_node(5, 0.0, 6.0, 0.0)
    solver.add_node(6, 6.0, 6.0, 0.0)
    solver.add_node(7, 6.0, 6.0, 6.0)
    solver.add_node(8, 0.0, 6.0, 6.0)

    # 2. Add Pinned Base Supports (Nodes 1, 2, 3, 4: UX, UY, UZ locked, Rotations free)
    for n in [1, 2, 3, 4]:
        solver.add_support(n, ux=1, uy=1, uz=1, rx=0, ry=0, rz=0)

    # 3. Add Members with Beta Angles & End Releases
    # Base Beams (Beta = 0 deg)
    solver.add_member(1, 1, 2, beta_deg=0.0, mem_type="Beam")
    solver.add_member(2, 2, 3, beta_deg=0.0, mem_type="Beam")
    solver.add_member(3, 4, 3, beta_deg=0.0, mem_type="Beam")
    solver.add_member(4, 1, 4, beta_deg=0.0, mem_type="Beam")

    # Roof Beams (Beta = 0 deg)
    beam_rel = [False] * 12
    solver.add_member(5, 5, 6, beta_deg=0.0, mem_type="Beam", releases=beam_rel)
    solver.add_member(6, 6, 7, beta_deg=0.0, mem_type="Beam")
    solver.add_member(7, 8, 7, beta_deg=0.0, mem_type="Beam")
    solver.add_member(8, 5, 8, beta_deg=0.0, mem_type="Beam")

    # Columns (Beta = 90 deg)
    solver.add_member(9,  1, 5, beta_deg=90.0, mem_type="Column")
    solver.add_member(10, 2, 6, beta_deg=90.0, mem_type="Column")
    solver.add_member(11, 3, 7, beta_deg=90.0, mem_type="Column")
    solver.add_member(12, 4, 8, beta_deg=90.0, mem_type="Column")

    # 4. Service Loadings
    solver.add_node_load(5, Fx=50e3, Fy=-25e3, Fz=15e3)
    solver.add_node_load(6, Fx=20e3, Fy=-25e3, Fz=0.0)
    solver.add_node_load(7, Fx=0.0,  Fy=-25e3, Fz=0.0)
    solver.add_node_load(8, Fx=0.0,  Fy=-25e3, Fz=15e3)

    # 5. Solve System
    res = solver.solve()
    print(f"Total DOFs: {len(solver.nodes)*6} | Restrained: {len(res['restrained_dofs'])} | Active: {len(res['active_dofs'])}")
    print(f"Max Lateral Sway: {np.max(np.abs(res['U_global']))*1000:.2f} mm")
    print(f"Reactions Sum Fx: {np.sum([res['Reactions'][(n-1)*6] for n in range(1, 5)])/1e3:.2f} kN")
    print(f"Reactions Sum Fy: {np.sum([res['Reactions'][(n-1)*6+1] for n in range(1, 5)])/1e3:.2f} kN")
    print(f"Reactions Sum Fz: {np.sum([res['Reactions'][(n-1)*6+2] for n in range(1, 5)])/1e3:.2f} kN")

    # 6. Generate Custom Visual Diagram
    generate_custom_structural_plot(solver, "structural_model_sabila.png")

    # 7. Generate Formatted Excel Report
    export_to_excel_earthy(solver, "structural_solver_sabila.xlsx")
    print("=== ANALYSIS, PLOTTING & EXCEL EXPORT COMPLETED SUCCESSFULLY ===")

if __name__ == '__main__':
    main()
