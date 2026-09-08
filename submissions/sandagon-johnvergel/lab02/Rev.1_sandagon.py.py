"""Rev. 1 - Navigable 3-D structural model of a 6 m cube frame.

Global coordinate system: X and Z are lateral; Y is vertical.  Each node has
six global degrees of freedom: UX, UY, UZ, RX, RY, and RZ.  Nodes 1--4 have
pinned supports (translations restrained, rotations free).

Beam end connections are pinned by releasing global RZ at both member ends.
The magenta hollow circles in the diagram identify those pinned ends.

Run: python cube_nodes_6m_Rev_1.py
This opens the navigable model and writes structural_model_Rev_1.xlsx.
"""

from __future__ import annotations

import math
from pathlib import Path
from typing import Iterable
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d.art3d import Poly3DCollection


REVISION = "Rev. 1"
SIDE_LENGTH_M = 6.0
GLOBAL_DOFS = ("UX", "UY", "UZ", "RX", "RY", "RZ")

# Node: (x, y, z). Nodes 1--4 form the base; nodes 5--8 form the roof.
NODES = {
    1: (0.0, 0.0, 0.0), 2: (6.0, 0.0, 0.0),
    3: (6.0, 0.0, 6.0), 4: (0.0, 0.0, 6.0),
    5: (0.0, 6.0, 0.0), 6: (6.0, 6.0, 0.0),
    7: (6.0, 6.0, 6.0), 8: (0.0, 6.0, 6.0),
}

# Pinned supports restrain all translations and allow all rotations.
SUPPORTS = {
    node: {"UX": True, "UY": True, "UZ": True,
           "RX": False, "RY": False, "RZ": False}
    for node in (1, 2, 3, 4)
}

# A release value of True means that the member-end DOF is released.
# Beam pin connections are represented by released RZ (moment about global Z).
MEMBERS = [
    {"id": 1, "i": 1, "j": 2, "type": "Base beam", "beta_deg": 0.0, "pinned": True},
    {"id": 2, "i": 2, "j": 3, "type": "Base beam", "beta_deg": 0.0, "pinned": True},
    {"id": 3, "i": 3, "j": 4, "type": "Base beam", "beta_deg": 0.0, "pinned": True},
    {"id": 4, "i": 4, "j": 1, "type": "Base beam", "beta_deg": 0.0, "pinned": True},
    {"id": 5, "i": 5, "j": 6, "type": "Roof beam", "beta_deg": 0.0, "pinned": True},
    {"id": 6, "i": 6, "j": 7, "type": "Roof beam", "beta_deg": 0.0, "pinned": True},
    {"id": 7, "i": 7, "j": 8, "type": "Roof beam", "beta_deg": 0.0, "pinned": True},
    {"id": 8, "i": 8, "j": 5, "type": "Roof beam", "beta_deg": 0.0, "pinned": True},
    {"id": 9, "i": 1, "j": 5, "type": "Column", "beta_deg": 90.0, "pinned": False},
    {"id": 10, "i": 2, "j": 6, "type": "Column", "beta_deg": 90.0, "pinned": False},
    {"id": 11, "i": 3, "j": 7, "type": "Column", "beta_deg": 90.0, "pinned": False},
    {"id": 12, "i": 4, "j": 8, "type": "Column", "beta_deg": 90.0, "pinned": False},
]


def vector_subtract(a: tuple[float, float, float], b: tuple[float, float, float]):
    return tuple(a_i - b_i for a_i, b_i in zip(a, b))


def dot(a: Iterable[float], b: Iterable[float]) -> float:
    return sum(a_i * b_i for a_i, b_i in zip(a, b))


def cross(a: tuple[float, float, float], b: tuple[float, float, float]):
    return (a[1] * b[2] - a[2] * b[1],
            a[2] * b[0] - a[0] * b[2],
            a[0] * b[1] - a[1] * b[0])


def normalize(v: tuple[float, float, float]):
    magnitude = math.sqrt(dot(v, v))
    if magnitude == 0:
        raise ValueError("A zero-length member does not have local axes.")
    return tuple(value / magnitude for value in v)


def local_axes(member: dict):
    """Return local x, y, z unit vectors after applying beta about local x."""
    local_x = normalize(vector_subtract(NODES[member["j"]], NODES[member["i"]]))
    global_y = (0.0, 1.0, 0.0)
    # Keep local y close to vertical; for a vertical member use global Z instead.
    projection = tuple(global_y[index] - dot(global_y, local_x) * local_x[index]
                       for index in range(3))
    if math.sqrt(dot(projection, projection)) < 1e-9:
        projection = (0.0, 0.0, 1.0)
    local_y_0 = normalize(projection)
    local_z_0 = normalize(cross(local_x, local_y_0))
    beta = math.radians(member["beta_deg"])
    local_y = tuple(math.cos(beta) * local_y_0[index] + math.sin(beta) * local_z_0[index]
                    for index in range(3))
    local_z = tuple(-math.sin(beta) * local_y_0[index] + math.cos(beta) * local_z_0[index]
                    for index in range(3))
    return local_x, local_y, local_z


def member_releases(member: dict):
    released = {dof: False for dof in GLOBAL_DOFS}
    if member["pinned"]:
        released["RZ"] = True
    return released


def node_dof_state(node: int):
    restraints = SUPPORTS.get(node, {})
    return {dof: "Restrained" if restraints.get(dof, False) else "Free" for dof in GLOBAL_DOFS}


def plot_structure():
    """Display the model in the RISA/STAAD-style layout shown in the brief."""
    fig = plt.figure(figsize=(17, 12))
    ax = fig.add_axes((0.05, 0.08, 0.66, 0.74), projection="3d")
    beam_color, column_color = "#2454d8", "#08775d"

    for member in MEMBERS:
        start, end = NODES[member["i"]], NODES[member["j"]]
        member_color = column_color if member["type"] == "Column" else beam_color
        ax.plot([start[0], end[0]], [start[2], end[2]], [start[1], end[1]],
                color=member_color, linewidth=2.7,
                label="Column" if member["id"] == 9 else "Beam" if member["id"] == 1 else None)

        midpoint = tuple((a + b) / 2 for a, b in zip(start, end))
        local_x, local_y, local_z = local_axes(member)
        for axis, color, label in zip((local_x, local_y, local_z),
                                      ("#dd2929", "#22aa31", "#9361d1"),
                                      ("Local x axis", "Local y axis", "Local z axis")):
            ax.quiver(midpoint[0], midpoint[2], midpoint[1], axis[0], axis[2], axis[1],
                      length=0.65, normalize=True, color=color, arrow_length_ratio=0.24,
                      linewidth=1.5, label=label if member["id"] == 1 else None)
        beta_label = f" (β={member['beta_deg']:.0f}°)" if member["beta_deg"] else ""
        ax.text(midpoint[0], midpoint[2], midpoint[1] + 0.10,
                f"M{member['id']}{beta_label}", color="#123a8c", fontsize=8, weight="bold")

    for node, (x, y, z) in NODES.items():
        supported = node in SUPPORTS
        color = "#c40000" if supported else "#ff5959"
        ax.scatter(x, z, y, color=color, edgecolors="black", linewidths=0.7, s=76,
                   depthshade=False, label="Supported node (pinned)" if node == 1 else "Free node" if node == 5 else None)
        dof_start = (node - 1) * 6 + 1
        ax.text(x, z, y + 0.20, f"N{node}\nDOF {dof_start}-{dof_start + 5}",
                fontsize=8.5, weight="bold", color="#111111")
        if supported:
            draw_pinned_support(ax, x, y, z)

    ax.scatter(0, 0, 0, marker="*", color="#12952c", edgecolors="black", s=115,
               depthshade=False, label="Origin (0, 0, 0)")
    ax.set_title("6m x 6m x 6m Cube - Structural Model, Rev. 1\n"
                 "Pinned supports at nodes 1-4, member local axes and beta angles shown",
                 fontsize=14, weight="bold", pad=30)
    ax.set_xlabel("X (m) - lateral", fontsize=11, weight="bold", labelpad=10)
    ax.set_ylabel("Z (m) - lateral", fontsize=11, weight="bold", labelpad=10)
    ax.set_zlabel("Y (m) - vertical", fontsize=11, weight="bold", labelpad=10)
    ax.set_xlim(-1, 7)
    ax.set_ylim(-1, 7)
    ax.set_zlim(-1, 7)
    ax.set_box_aspect((1, 1, 1))
    ax.view_init(elev=17, azim=-58)
    ax.legend(loc="upper left", bbox_to_anchor=(-0.12, 0.96), fontsize=9, framealpha=0.94)
    ax.grid(True)
    add_model_data_panel(fig)
    plt.show()


def draw_pinned_support(ax, x: float, y: float, z: float):
    """Draw a grey pyramid support directly below a supported physical node."""
    base_y, half_width = y - 0.55, 0.33
    apex = (x, z, y)
    base = [(x - half_width, z - half_width, base_y), (x + half_width, z - half_width, base_y),
            (x + half_width, z + half_width, base_y), (x - half_width, z + half_width, base_y)]
    faces = [[apex, base[0], base[1]], [apex, base[1], base[2]],
             [apex, base[2], base[3]], [apex, base[3], base[0]], base]
    ax.add_collection3d(Poly3DCollection(faces, facecolors="#555555", edgecolors="#222222",
                                         linewidths=0.8, alpha=0.9))


def add_model_data_panel(fig):
    """Add the right-hand model summary, matching the requested plot style."""
    summary = """MODEL DATA - REV. 1

Geometry
    Cube edge                 6.0 m
    Nodes                     8
    Members                   12
    Vertical axis             global Y

Supports
    Type                      pinned
    Nodes                     1, 2, 3, 4
    Restrained                UX, UY, UZ
    Released                  RX, RY, RZ

Degrees of freedom
    DOF per node              6
    Total DOF                 48
    Restrained DOF            12
    Active DOF                36
    Numbering                 (node - 1) x 6 + 1...6

Beta angles
    Base Beam                 0 deg
    Roof Beam                 0 deg
    Column                    90 deg

Local axes
    local x   start node i to end node j
    local y   in the vertical plane, when possible
    local z   completes the right-handed set
    Vertical members follow the limiting special
    case: local z parallel to global Z."""
    fig.text(0.75, 0.82, summary, va="top", ha="left", fontsize=9.5, family="monospace",
             bbox={"boxstyle": "round,pad=0.8", "facecolor": "#f5f6fc", "edgecolor": "#153c77", "linewidth": 1.25})


def column_name(index: int) -> str:
    """Convert a zero-based column index to an Excel column name."""
    result = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def worksheet_xml(rows: list[list[object]]) -> str:
    xml_rows = []
    for row_number, row in enumerate(rows, 1):
        cells = []
        for column, value in enumerate(row):
            reference = f"{column_name(column)}{row_number}"
            style_number = 1 if row_number == 1 else 0
            if row_number > 1 and value == "Active":
                style_number = 2
            elif row_number > 1 and value == "Restrained":
                style_number = 3
            elif isinstance(value, float):
                style_number = 4
            style = f' s="{style_number}"' if style_number else ""
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                cells.append(f'<c r="{reference}"{style}><v>{value}</v></c>')
            else:
                cells.append(f'<c r="{reference}"{style} t="inlineStr"><is><t>{escape(str(value))}</t></is></c>')
        xml_rows.append(f'<row r="{row_number}">{"".join(cells)}</row>')
    widths = ''.join(f'<col min="{number}" max="{number}" width="18" customWidth="1"/>'
                     for number in range(1, len(rows[0]) + 1))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<cols>{widths}</cols><sheetData>{"".join(xml_rows)}</sheetData></worksheet>')


def write_excel_output(filename: Path):
    """Write the Rev. 1 model input and detailed DOF schedule to a .xlsx workbook."""
    summary_rows = [["Item", "Value"], ["Revision", REVISION], ["Cube edge length (m)", SIDE_LENGTH_M],
                    ["Number of nodes", 8], ["Number of members", 12], ["Supported nodes", 4],
                    ["Support type", "Pinned"], ["DOF per node", 6], ["Total DOF", 48],
                    ["Restrained DOF", 12], ["Active DOF (equations)", 36], ["Global vertical axis", "Y"],
                    ["Global lateral axes", "X and Z"], ["Beta angle, base beams (deg)", 0],
                    ["Beta angle, roof beams (deg)", 0], ["Beta angle, columns (deg)", 90]]
    node_rows = [["Node", "X (m)", "Y (m)", "Z (m)", "Support"]]
    incidence_rows = [["Member", "Node i (Start)", "Node j (End)", "Type", "Length (m)", "Beta (deg)"]]
    support_rows = [["Node", "X (m)", "Y (m)", "Z (m)", "Support Type", *GLOBAL_DOFS, "Restraint Code"]]
    axis_rows = [["Member", "Node i", "Node j", "Type", "Length (m)", "Beta (deg)", "local x - X", "local x - Y", "local x - Z", "local y - X", "local y - Y", "local y - Z", "local z - X", "local z - Y", "local z - Z"]]
    node_dof_rows = [["Node", "Support", "Global DOF Range", *GLOBAL_DOFS, "Active DOF"]]
    for node, (x, y, z) in NODES.items():
        states = node_dof_state(node)
        node_rows.append([node, x, y, z, "Pinned" if node in SUPPORTS else "Free"])
        support_rows.append([node, x, y, z, "Pinned" if node in SUPPORTS else "Free", *(states[dof] for dof in GLOBAL_DOFS), "111000" if node in SUPPORTS else "000000"])
        start = (node - 1) * 6 + 1
        node_dof_rows.append([node, "Pinned" if node in SUPPORTS else "Free", f"{start} - {start + 5}", *(states[dof] for dof in GLOBAL_DOFS), sum(state == "Free" for state in states.values())])
    for member in MEMBERS:
        axes = local_axes(member)
        delta = vector_subtract(NODES[member["j"]], NODES[member["i"]])
        length = math.sqrt(dot(delta, delta))
        incidence_rows.append([member["id"], member["i"], member["j"], member["type"], length, member["beta_deg"]])
        axis_rows.append([member["id"], member["i"], member["j"], member["type"], length, member["beta_deg"], *(value for axis in axes for value in axis)])
    descriptions = {"UX": "Translation X", "UY": "Translation Y", "UZ": "Translation Z", "RX": "Rotation X", "RY": "Rotation Y", "RZ": "Rotation Z"}
    numbering_rows = [["Node", "Local DOF", "DOF", "Description", "Global DOF No.", "Status", "Equation No."]]
    equation = 0
    for node in NODES:
        for local_number, dof in enumerate(GLOBAL_DOFS, 1):
            status = "Restrained" if node_dof_state(node)[dof] == "Restrained" else "Active"
            if status == "Active":
                equation += 1
            numbering_rows.append([node, local_number, dof, descriptions[dof], (node - 1) * 6 + local_number, status, equation if status == "Active" else "-"])
    sheets = [("Model Summary", summary_rows), ("Nodes", node_rows), ("Member Incidences", incidence_rows), ("Supports", support_rows), ("Local Axes", axis_rows), ("Node DOF", node_dof_rows), ("DOF Numbering", numbering_rows)]
    with ZipFile(filename, "w", ZIP_DEFLATED) as workbook:
        overrides = ''.join(f'<Override PartName="/xl/worksheets/sheet{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for n in range(1, len(sheets) + 1))
        workbook.writestr("[Content_Types].xml", f'''<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>{overrides}</Types>''')
        workbook.writestr("_rels/.rels", '''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>''')
        sheet_xml = ''.join(f'<sheet name="{escape(name)}" sheetId="{n}" r:id="rId{n}"/>' for n, (name, _) in enumerate(sheets, 1))
        workbook.writestr("xl/workbook.xml", f'''<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{sheet_xml}</sheets></workbook>''')
        relationships = ''.join(f'<Relationship Id="rId{n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{n}.xml"/>' for n in range(1, len(sheets) + 1))
        workbook.writestr("xl/_rels/workbook.xml.rels", f'''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relationships}<Relationship Id="rId{len(sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>''')
        workbook.writestr("xl/styles.xml", '''<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF17365D"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="5"><xf xfId="0"/><xf xfId="0" fontId="1" fillId="1" applyFont="1" applyFill="1"/><xf xfId="0" fillId="2" applyFill="1"/><xf xfId="0" fillId="3" applyFill="1"/><xf xfId="0" numFmtId="164" applyNumberFormat="1"/></cellXfs></styleSheet>''')
        for number, (_, rows) in enumerate(sheets, 1):
            workbook.writestr(f"xl/worksheets/sheet{number}.xml", worksheet_xml(rows))


if __name__ == "__main__":
    output_file = Path(__file__).with_name("structural_model_Rev_1_detailed.xlsx")
    write_excel_output(output_file)
    print(f"Wrote Excel model data: {output_file}")
    plot_structure()
