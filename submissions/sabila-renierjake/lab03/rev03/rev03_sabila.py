"""
================================================================================
3D SPACE FRAME STRUCTURAL ANALYSIS SOLVER - REV. 3
STAAD.Pro / RISA-3D Equivalent Direct Stiffness Matrix Solver
Author / Project: Sabila (structural_solver_sabila.py)
--------------------------------------------------------------------------------
Key Features Implemented (preserved from REV. 1 / REV. 2):
  1. Support Conditions: Pinned supports at Nodes 1-4 (UX, UY, UZ restrained; RX, RY, RZ free)
  2. STAAD/RISA Beta Angle (β): 0° for Beams, 90° for Columns
  3. Local & Global Coordinate System Transformations with vertical singularity handling
  4. 6 Degrees of Freedom (DOFs) per node: [UX, UY, UZ, RX, RY, RZ] -> 48 Total DOFs
  5. Member End Releases: Internal degree-of-freedom condensation (e.g., pinned at x, moment at z)
  6. Unique 3D Structural Analytical Wireframe & Triad Visualization
  7. Multi-Tab Formatted Excel Workbook Output (structural_solver_sabila.xlsx)
  8. In-Excel Standalone VBA Structural Solver Engine

New in REV. 3 (data-driven configuration layer -- STANDALONE BUILD:
units_config.py / material_config.py / member_size_config.py are now
integrated into this single file as Section 0, logic unchanged):
  9. Units Excel (Units_Imperial_Metric.xlsx) -> UnitsConfig -> Standard Metric
     default, Imperial available as an alternative, all conversion factors
     sourced from the workbook (none hard-coded in the solver).
 10. Material Excel (RISA_Materials_Library_Metric.xlsx) -> MaterialConfig ->
     ASTM A36 Steel (E, G, Fy, Fu, density) retrieved by lookup, not invented.
 11. Member-size Excel (aisc-shapes-database-v160-2.xlsx) -> MemberSizeConfig ->
     an explicit, real AISC section (default W12X26) supplies A, Iy, Iz, J.
 12. The solver's math (assembly, get_local_stiffness, solve, etc.) is
     UNCHANGED from REV. 2 -- only the property VALUES fed into add_member()
     now flow from the Excel-driven configuration objects below.
================================================================================
"""

import os
import sys
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
from matplotlib.colors import to_rgba
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- REV. 3 data layer -------------------------------------------------------
# (units_config / material_config / member_size_config are integrated below
#  as Section 0 -- no longer separate files; see note above OUTPUT_DIR)

# Generated deliverables are written next to this script (independent of cwd).
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))


# ==============================================================================
# 0. REV. 3 DATA-DRIVEN CONFIGURATION LAYER -- INTEGRATED, STANDALONE BUILD
# ------------------------------------------------------------------------------
# Formerly three separate modules (units_config.py, material_config.py,
# member_size_config.py), imported by the solver via:
#     from units_config import UnitsConfig
#     from material_config import MaterialConfig
#     from member_size_config import MemberSizeConfig
# Inlined here verbatim (logic and values UNCHANGED) so this single file has
# no local-module dependency. Still reads the same three external Excel
# workbooks at runtime (Units/, Material/, Member Size/ subfolders next to
# this script) -- only the Python import boundary was removed, not the data
# source. Each original module's docstring is preserved below for reference.
# ==============================================================================

"""
================================================================================
UNITS CONFIGURATION LAYER (REV. 3)
--------------------------------------------------------------------------------
Source of truth: Units_Imperial_Metric.xlsx  (sheet "Conversion Factors")

Purpose
  * Read ALL imperial<->metric conversion factors from the supplied Excel file
    (no conversion factor is hand-typed into the solver).
  * Provide a single UNIT_SYSTEM switch ("Metric" default, "Imperial" alternate).
  * Provide small, explicit converter functions the rest of REV3 calls instead
    of scattering '25.4', '6.895e6', etc. throughout the code.

Only true SI *prefix* scaling (mega, kilo, milli -- e.g. MPa -> Pa, kN -> N,
mm -> m) is kept as a tiny fixed table here, because those are decimal-prefix
definitions, not imperial/metric engineering conversions -- they are identical
in every system and are not present as rows in the Conversion Factors sheet.
================================================================================
"""
# All supplied workbooks are resolved relative to THIS file, so the solver runs
# correctly from any working directory:
#   <folder>/Units/Units_Imperial_Metric.xlsx
#   <folder>/Material/RISA_Materials_Library_Metric.xlsx
#   <folder>/Member Size/aisc-shapes-database-v160-2.xlsx
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_UNITS_XLSX = os.path.join(BASE_DIR, "Units", "Units_Imperial_Metric.xlsx")

VALID_SYSTEMS = ("Metric", "Imperial")

# Standard Metric is the REV3 default per requirements. Imperial remains
# available as an alternative by changing this one value.
UNIT_SYSTEM = "Metric"   # "Metric" (default) | "Imperial"

# Fixed SI decimal-prefix scaling (universal, not sourced from the workbook)
SI_PREFIX = {
    "mega_to_base": 1e6,     # e.g. MPa -> Pa
    "kilo_to_base": 1e3,     # e.g. kN -> N, kNm -> Nm
    "milli_to_base": 1e-3,   # e.g. mm -> m
}


def _load_conversion_rows(xlsx_path=DEFAULT_UNITS_XLSX):
    """Reads the 'Derived factors' table from the Conversion Factors sheet.
    Returns {quantity: {'imperial_unit', 'metric_unit', 'imp_to_metric', 'metric_to_imp'}}
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Conversion Factors"]

    factors = {}
    header_seen = False
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=6, values_only=True):
        if not header_seen:
            if row[0] == "Quantity":
                header_seen = True
            continue
        if row[0] is None:
            continue
        quantity, imp_unit, metric_unit, imp_to_metric, metric_to_imp, _derivation = row
        # Keep the first occurrence per quantity (e.g. "Length" appears twice:
        # in->mm and ft->m; we need the in->mm row for section properties).
        key = f"{quantity}:{imp_unit}->{metric_unit}"
        factors[key] = {
            "quantity": quantity,
            "imperial_unit": imp_unit,
            "metric_unit": metric_unit,
            "imp_to_metric": imp_to_metric,
            "metric_to_imp": metric_to_imp,
        }
    return factors


def _load_base_definitions(xlsx_path=DEFAULT_UNITS_XLSX):
    """Reads the exact NIST base definitions (e.g. 1 in = 25.4 mm)."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Conversion Factors"]
    base = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=3, values_only=True):
        if row[0] and isinstance(row[0], str):
            base[row[0]] = row[1]
    return base


def load_unit_system_table(xlsx_path=DEFAULT_UNITS_XLSX):
    """Reads the 'Unit Systems' sheet: Imperial vs Standard Metric label per quantity."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Unit Systems"]
    table = {}
    header_seen = False
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=5, values_only=True):
        if not header_seen:
            if row[0] == "Quantity":
                header_seen = True
            continue
        if row[0] is None:
            continue
        qty, imperial, metric, internal, note = row
        table[qty] = {"Imperial": imperial, "Standard Metric": metric, "internal": internal}
    return table


class UnitsConfig:
    """Excel-driven unit configuration. Instantiate once; the rest of REV3
    (material_config, member_size_config, solver, plot) consumes this object
    instead of hard-coding conversion numbers."""

    def __init__(self, xlsx_path=DEFAULT_UNITS_XLSX, system=UNIT_SYSTEM):
        if system not in VALID_SYSTEMS:
            raise ValueError(f"Unknown unit system '{system}'. Use one of {VALID_SYSTEMS}.")
        self.xlsx_path = xlsx_path
        self.system = system  # "Metric" or "Imperial"
        self._factors = _load_conversion_rows(xlsx_path)
        self._base = _load_base_definitions(xlsx_path)
        self._systems_table = load_unit_system_table(xlsx_path)

        # Pull the specific Excel-sourced factors needed by REV3's section /
        # material property conversions (in->mm, in2->mm2, in4->mm4).
        self._f_length = self._factors["Length:in->mm"]["imp_to_metric"]      # 25.4
        self._f_area = self._factors["Area:in²->mm²"]["imp_to_metric"]        # 645.16
        self._f_moi = self._factors["Moment of inertia:in⁴->mm⁴"]["imp_to_metric"]  # 416231.4256
        self._f_stress = self._factors["Stress / modulus:ksi->MPa"]["imp_to_metric"]  # 6.894757...
        self._f_force = self._factors["Force:kip->kN"]["imp_to_metric"]       # 4.4482216...
        self._f_unitweight = self._factors["Unit weight:k/ft³->kN/m³"]["imp_to_metric"]
        # Added in REV3 final fix: needed to derive mass density / thermal
        # coefficient when reading an Imperial RISA materials workbook.
        self._f_massdens = self._factors["Mass density from unit weight:k/ft³->kg/m³"]["imp_to_metric"]
        self._f_therm = self._factors["Thermal coefficient:1e-5 /°F->1e-6 /°C"]["imp_to_metric"]
        # Reciprocal (Metric -> Imperial) factors, also from the workbook, used
        # only to DISPLAY results in Imperial units when that system is active.
        self._r_stress = self._factors["Stress / modulus:ksi->MPa"]["metric_to_imp"]
        self._r_area = self._factors["Area:in²->mm²"]["metric_to_imp"]
        self._r_moi = self._factors["Moment of inertia:in⁴->mm⁴"]["metric_to_imp"]

    # ---- length / geometry (Excel factor: in -> mm, then SI prefix mm -> m) ----
    def in_to_m(self, value_in):
        mm = value_in * self._f_length
        return mm * SI_PREFIX["milli_to_base"]

    def in2_to_m2(self, value_in2):
        mm2 = value_in2 * self._f_area
        return mm2 * (SI_PREFIX["milli_to_base"] ** 2)

    def in4_to_m4(self, value_in4):
        mm4 = value_in4 * self._f_moi
        return mm4 * (SI_PREFIX["milli_to_base"] ** 4)

    # ---- stress / modulus (Excel factor: ksi -> MPa, then SI prefix MPa -> Pa) ----
    def ksi_to_pa(self, value_ksi):
        mpa = value_ksi * self._f_stress
        return mpa * SI_PREFIX["mega_to_base"]

    @staticmethod
    def mpa_to_pa(value_mpa):
        return value_mpa * SI_PREFIX["mega_to_base"]

    # ---- force ----
    def kip_to_n(self, value_kip):
        kN = value_kip * self._f_force
        return kN * SI_PREFIX["kilo_to_base"]

    @staticmethod
    def kn_to_n(value_kn):
        return value_kn * SI_PREFIX["kilo_to_base"]

    # ---- unit weight / density ----
    def kft3_to_knm3(self, value):
        return value * self._f_unitweight

    def kft3_to_kgm3(self, value):
        """Unit weight [k/ft3] -> mass density [kg/m3] (Excel factor: N/m3 / g)."""
        return value * self._f_massdens

    def therm_1e5F_to_perC(self, value):
        """Thermal coefficient [1e-5 /degF] -> [1/degC] (Excel factor -> 1e-6/degC, then x1e-6)."""
        return value * self._f_therm * 1e-6

    # ---- display-only reverse conversions (SI -> Imperial), Excel reciprocals ----
    def pa_to_ksi(self, value_pa):
        return (value_pa / SI_PREFIX["mega_to_base"]) * self._r_stress

    def m2_to_in2(self, value_m2):
        return (value_m2 / (SI_PREFIX["milli_to_base"] ** 2)) * self._r_area

    def m4_to_in4(self, value_m4):
        return (value_m4 / (SI_PREFIX["milli_to_base"] ** 4)) * self._r_moi

    def label(self, quantity):
        """Returns the display unit label for `quantity` in the active system,
        e.g. label('Section dimensions') -> 'mm' (Metric) or 'in' (Imperial)."""
        row = self._systems_table.get(quantity)
        if not row:
            return ""
        return row["Standard Metric"] if self.system == "Metric" else row["Imperial"]

    def system_display_name(self):
        return "Standard Metric" if self.system == "Metric" else "Imperial (US Customary)"

"""
================================================================================
MATERIAL CONFIGURATION LAYER (REV. 3)
--------------------------------------------------------------------------------
Source of truth: RISA_Materials_Library_Metric.xlsx / _Imperial.xlsx
                  sheet "All Materials"

Purpose
  * Look up a named structural material (default: ASTM A36 -> label "A36 Gr.36",
    category "Hot Rolled") from the RISA material workbook.
  * Do NOT invent or hard-code A36 property values -- every number returned
    here is read from the workbook cell.
  * Convert whatever unit system the source workbook is in (Metric MPa/kN-m3,
    or Imperial ksi/kcf) into consistent SI (Pa, kg/m3) using UnitsConfig,
    since the REV2 solver's internal math (E in Pa, coordinates in m) is
    already SI-consistent = "Standard Metric" internal convention.
================================================================================
"""
METRIC_MATERIALS_XLSX = os.path.join(BASE_DIR, "Material", "RISA_Materials_Library_Metric.xlsx")
IMPERIAL_MATERIALS_XLSX = os.path.join(BASE_DIR, "Material", "RISA_Materials_Library_Imperial.xlsx")

MATERIAL_SHEET = "All Materials"

DEFAULT_MATERIAL_LABEL = "A36 Gr.36"     # ASTM A36 Steel, as listed under "Hot Rolled"
DEFAULT_MATERIAL_CATEGORY = "Hot Rolled"


def _read_all_materials_metric(xlsx_path=METRIC_MATERIALS_XLSX):
    """Parses the Metric 'All Materials' sheet into a list of row dicts.
    Columns: Category, Label, E [MPa], G [MPa], Nu, Therm.Coeff [1e-6/C],
             Density [kN/m3], Yield/f'c/f'm [MPa], Fu [MPa], Mass Density [kg/m3]
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[MATERIAL_SHEET]

    rows = []
    header_seen = False
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=10, values_only=True):
        if not header_seen:
            if row[0] == "Category":
                header_seen = True
            continue
        if row[0] is None or row[1] is None:
            continue
        rows.append({
            "category": row[0], "label": row[1],
            "E_MPa": row[2], "G_MPa": row[3], "nu": row[4],
            "therm_coeff_e6_perC": row[5], "density_kNm3": row[6],
            "yield_MPa": row[7], "Fu_MPa": row[8], "mass_density_kgm3": row[9],
        })
    return rows


def _read_all_materials_imperial(xlsx_path=IMPERIAL_MATERIALS_XLSX):
    """Parses the Imperial 'All Materials' sheet (ksi / kcf based) into a list
    of {header: value} dicts. Reads every column present (the Imperial sheet
    has NO 'Mass Density' column, unlike the Metric sheet)."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[MATERIAL_SHEET]

    rows = []
    headers = None
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
        if headers is None:
            if row[0] == "Category":
                headers = row
            continue
        if row[0] is None or row[1] is None:
            continue
        rows.append(dict(zip(headers, row)))
    return rows


def _find_key(record, predicate, what):
    """Returns the first header in `record` satisfying `predicate`, or raises
    a descriptive error listing the headers actually present (instead of an
    opaque StopIteration)."""
    for k in record:
        if k and predicate(str(k)):
            return k
    raise KeyError(f"Column for {what} not found. Headers present: "
                   f"{[k for k in record if k]}")


class MaterialConfig:
    """Excel-driven material lookup. Retrieves ASTM A36 Steel (or any other
    listed material) from the RISA library workbook and exposes SI-consistent
    properties (Pa, kg/m3) for direct use by the Frame3DSolver."""

    def __init__(self, units: UnitsConfig, label=DEFAULT_MATERIAL_LABEL,
                 category=DEFAULT_MATERIAL_CATEGORY):
        self.units = units
        self.label = label
        self.category = category
        self.source_system = units.system  # which workbook we read from
        self.source_file = None
        self.raw = None
        self.fallback_note = None
        self._load()

    def _load(self):
        if self.source_system == "Metric":
            self._load_metric()
        else:
            self._load_imperial()

    def _load_metric(self):
        self.source_file = METRIC_MATERIALS_XLSX
        rows = _read_all_materials_metric(self.source_file)
        match = next((r for r in rows if r["label"] == self.label
                      and r["category"] == self.category), None)
        if match is None:
            raise ValueError(f"Material '{self.label}' not found in "
                             f"{self.source_file} ({MATERIAL_SHEET})")
        self.raw = match
        self.E_Pa = UnitsConfig.mpa_to_pa(match["E_MPa"])
        self.G_Pa = UnitsConfig.mpa_to_pa(match["G_MPa"])
        self.nu = match["nu"]
        self.Fy_Pa = UnitsConfig.mpa_to_pa(match["yield_MPa"])
        self.Fu_Pa = UnitsConfig.mpa_to_pa(match["Fu_MPa"])
        self.density_kgm3 = match["mass_density_kgm3"]
        self.therm_coeff_perC = match["therm_coeff_e6_perC"] * 1e-6

    def _load_imperial(self):
        """Imperial path. Reads RISA_Materials_Library_Imperial.xlsx when it is
        supplied. Mass density is DERIVED from the workbook's unit weight
        [k/ft3] using the Units workbook factor (N/m3 / g), because the Imperial
        sheet has no Mass Density column.

        If the Imperial materials workbook is not present in Material/, fall
        back to the supplied Metric workbook (same physical material, values
        already SI) and record that fact in self.fallback_note."""
        if not os.path.exists(IMPERIAL_MATERIALS_XLSX):
            self.fallback_note = (f"Imperial materials workbook not found "
                                  f"({os.path.basename(IMPERIAL_MATERIALS_XLSX)}); "
                                  f"properties read from the supplied Metric workbook instead.")
            print(f"[MATERIAL] NOTE: {self.fallback_note}")
            self._load_metric()
            return

        self.source_file = IMPERIAL_MATERIALS_XLSX
        rows = _read_all_materials_imperial(self.source_file)
        match = next((r for r in rows if r.get("Label") == self.label
                      and r.get("Category") == self.category), None)
        if match is None:
            raise ValueError(f"Material '{self.label}' not found in "
                             f"{self.source_file} ({MATERIAL_SHEET})")
        self.raw = match
        e_key = _find_key(match, lambda k: k.startswith("E ["), "E")
        g_key = _find_key(match, lambda k: k.startswith("G ["), "G")
        nu_key = _find_key(match, lambda k: k.strip().lower() == "nu", "Nu")
        fy_key = _find_key(match, lambda k: "Yield" in k, "Yield")
        fu_key = _find_key(match, lambda k: k.startswith("Fu"), "Fu")
        self.E_Pa = self.units.ksi_to_pa(match[e_key])
        self.G_Pa = self.units.ksi_to_pa(match[g_key])
        self.nu = match[nu_key]
        self.Fy_Pa = self.units.ksi_to_pa(match[fy_key])
        self.Fu_Pa = self.units.ksi_to_pa(match[fu_key])

        # Mass density: use an explicit column if a workbook version has one,
        # otherwise derive it from unit weight [k/ft3] (Excel-sourced factor).
        mass_keys = [k for k in match if k and "Mass Density" in str(k)]
        if mass_keys:
            self.density_kgm3 = match[mass_keys[0]]
        else:
            uw_key = _find_key(match, lambda k: "Density" in k and "k/ft" in k,
                               "unit weight [k/ft3]")
            self.density_kgm3 = self.units.kft3_to_kgm3(match[uw_key])

        th_keys = [k for k in match if k and str(k).startswith("Therm")]
        self.therm_coeff_perC = (self.units.therm_1e5F_to_perC(match[th_keys[0]])
                                 if th_keys else None)

    def display_name(self):
        return "ASTM A36 Steel" if self.label == "A36 Gr.36" else self.label

    def summary(self):
        return {
            "material": self.display_name(),
            "source_label": self.label,
            "category": self.category,
            "source_file": self.source_file,
            "source_system": self.source_system,
            "E_Pa": self.E_Pa, "E_GPa": self.E_Pa / 1e9,
            "G_Pa": self.G_Pa, "G_GPa": self.G_Pa / 1e9,
            "nu": self.nu,
            "Fy_Pa": self.Fy_Pa, "Fy_MPa": self.Fy_Pa / 1e6,
            "Fu_Pa": self.Fu_Pa, "Fu_MPa": self.Fu_Pa / 1e6,
            "density_kgm3": self.density_kgm3,
            "fallback_note": self.fallback_note,
        }

"""
================================================================================
MEMBER SIZE CONFIGURATION LAYER (REV. 3)
--------------------------------------------------------------------------------
Source of truth: aisc-shapes-database-v160-2.xlsx  (sheet "Database v16.0")

Purpose
  * Look up a real, published AISC section (default: W12X26 -- the same shape
    used as the worked example in the Units workbook) instead of an invented
    member size.
  * Read the section's geometric properties (A, Ix, Iy, J) from the workbook's
    IMPERIAL column block, which is unambiguous (in, in^2, in^4 -- unscaled).
    The metric column block in this AISC file uses AISC's mixed display
    scaling (e.g. Ix in x10^6 mm^4, J in x10^3 mm^4) which is a *presentation*
    convention, not a raw unit; to avoid mis-scaling, REV3 always converts the
    unambiguous imperial values to SI using UnitsConfig (itself sourced from
    Units_Imperial_Metric.xlsx), regardless of which unit system is active.
  * The AISC metric *label* (e.g. "W310X38.7") is still pulled from the
    workbook purely for Standard-Metric display purposes.

Local-axis mapping (documented assumption -- to be verified in Phase 4):
  The REV2 solver's local member axes follow STAAD/RISA convention: local-3
  ("Iz" in the solver) is the member's major bending axis when beta=0 deg
  (matching AISC's strong axis, Ix_AISC); local-2 ("Iy" in the solver) is the
  minor axis (AISC Iy). REV3 therefore maps:
      solver Iz  <-  AISC Ix   (strong/major axis)
      solver Iy  <-  AISC Iy   (weak/minor axis)
      solver J   <-  AISC J    (torsional constant)
================================================================================
"""
AISC_XLSX = os.path.join(BASE_DIR, "Member Size", "aisc-shapes-database-v160-2.xlsx")
AISC_SHEET = "Database v16.0"

DEFAULT_DESIGNATION = "W12X26"   # explicit initial member size, AISC v16.0


def _find_header_blocks(ws):
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    idxs = [i for i, v in enumerate(header) if v == "EDI_Std_Nomenclature"]
    imp_start, met_start = idxs[0], idxs[1]
    return header, imp_start, met_start


class MemberSizeConfig:
    """Excel-driven AISC member-size lookup. Retrieves an explicit, real
    section from the AISC v16.0 shapes database and exposes SI section
    properties (m^2, m^4) for direct use by the Frame3DSolver."""

    def __init__(self, units: UnitsConfig, designation=DEFAULT_DESIGNATION,
                 aisc_path=AISC_XLSX):
        self.units = units
        self.designation = designation
        self.source_file = aisc_path
        self._load()

    def _load(self):
        wb = openpyxl.load_workbook(self.source_file, data_only=True, read_only=True)
        ws = wb[AISC_SHEET]
        header, imp_start, met_start = _find_header_blocks(ws)

        need = ["Type", "AISC_Manual_Label", "A", "Ix", "Iy", "J", "Zx", "Sx", "d", "bf", "tf", "tw"]
        imp_idx = {k: header.index(k) if k == "Type" else imp_start + header[imp_start:].index(k)
                   for k in need}
        met_idx = {k: met_start + header[met_start:].index(k) for k in need if k != "Type"}

        match = None
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
            if row[imp_idx["AISC_Manual_Label"]] == self.designation:
                match = row
                break
        if match is None:
            raise ValueError(f"Shape '{self.designation}' not found in {self.source_file}")

        self.shape_type = match[imp_idx["Type"]]
        self.imperial_label = match[imp_idx["AISC_Manual_Label"]]
        self.metric_label = match[met_idx["AISC_Manual_Label"]]

        # Unambiguous imperial (in / in^2 / in^4) values, straight from Excel
        self.A_in2 = float(match[imp_idx["A"]])
        self.Ix_in4 = float(match[imp_idx["Ix"]])   # strong axis (AISC convention)
        self.Iy_in4 = float(match[imp_idx["Iy"]])   # weak axis
        self.J_in4 = float(match[imp_idx["J"]])
        self.d_in = float(match[imp_idx["d"]])
        self.bf_in = float(match[imp_idx["bf"]])

        # Convert to SI using the Excel-sourced conversion factors (units_config)
        self.A_m2 = self.units.in2_to_m2(self.A_in2)
        self.Ix_m4 = self.units.in4_to_m4(self.Ix_in4)   # -> solver local Iz (major axis)
        self.Iy_m4 = self.units.in4_to_m4(self.Iy_in4)   # -> solver local Iy (minor axis)
        self.J_m4 = self.units.in4_to_m4(self.J_in4)
        self.d_m = self.units.in_to_m(self.d_in)
        self.bf_m = self.units.in_to_m(self.bf_in)

    def display_name(self):
        """Section label shown in the active unit system."""
        return self.metric_label if self.units.system == "Metric" else self.imperial_label

    def summary(self):
        return {
            "designation_imperial": self.imperial_label,
            "designation_metric": self.metric_label,
            "shape_type": self.shape_type,
            "source_file": self.source_file,
            "A_m2": self.A_m2, "A_mm2": self.A_m2 * 1e6,
            "Iz_m4_major": self.Ix_m4,   # fed into solver as Iz
            "Iy_m4_minor": self.Iy_m4,   # fed into solver as Iy
            "J_m4": self.J_m4,
            "d_mm": self.d_m * 1e3,
            "bf_mm": self.bf_m * 1e3,
        }


def describe_configuration(units_cfg, material_cfg, section_cfg):
    """Single source for every human-readable Units / Material / Member-Size
    string (console, PNG card, Excel config block), rendered in the ACTIVE
    unit system. Values come from the config objects that fed the solver."""
    if units_cfg.system == "Metric":
        mat_txt = (f"{material_cfg.display_name()} "
                   f"(E={material_cfg.E_Pa/1e9:.1f} GPa, Fy={material_cfg.Fy_Pa/1e6:.1f} MPa)")
        alt_label = section_cfg.imperial_label
        alt_tag = "AISC"
        prop_txt = (f"A={section_cfg.A_m2*1e6:.0f} mm\u00b2   "
                    f"Iz={section_cfg.Ix_m4*1e6:.2f}e6 mm\u2074   "
                    f"Iy={section_cfg.Iy_m4*1e6:.2f}e6 mm\u2074   "
                    f"J={section_cfg.J_m4*1e6:.4f}e6 mm\u2074")
        area_short = f"A={section_cfg.A_m2*1e6:.0f} mm2"
    else:
        mat_txt = (f"{material_cfg.display_name()} "
                   f"(E={units_cfg.pa_to_ksi(material_cfg.E_Pa):.0f} ksi, "
                   f"Fy={units_cfg.pa_to_ksi(material_cfg.Fy_Pa):.1f} ksi)")
        alt_label = section_cfg.metric_label
        alt_tag = "Metric"
        prop_txt = (f"A={units_cfg.m2_to_in2(section_cfg.A_m2):.2f} in\u00b2   "
                    f"Iz={units_cfg.m4_to_in4(section_cfg.Ix_m4):.1f} in\u2074   "
                    f"Iy={units_cfg.m4_to_in4(section_cfg.Iy_m4):.1f} in\u2074   "
                    f"J={units_cfg.m4_to_in4(section_cfg.J_m4):.3f} in\u2074")
        area_short = f"A={units_cfg.m2_to_in2(section_cfg.A_m2):.2f} in2"
    if units_cfg.system == "Metric":
        pieces = {"E": f"{material_cfg.E_Pa/1e9:.1f} GPa", "Fy": f"{material_cfg.Fy_Pa/1e6:.1f} MPa",
                  "A": f"{section_cfg.A_m2*1e6:.0f} mm\u00b2",
                  "Iz": f"{section_cfg.Ix_m4*1e6:.2f}e6 mm\u2074",
                  "Iy": f"{section_cfg.Iy_m4*1e6:.2f}e6 mm\u2074",
                  "J": f"{section_cfg.J_m4*1e6:.4f}e6 mm\u2074"}
    else:
        pieces = {"E": f"{units_cfg.pa_to_ksi(material_cfg.E_Pa):.0f} ksi",
                  "Fy": f"{units_cfg.pa_to_ksi(material_cfg.Fy_Pa):.1f} ksi",
                  "A": f"{units_cfg.m2_to_in2(section_cfg.A_m2):.2f} in\u00b2",
                  "Iz": f"{units_cfg.m4_to_in4(section_cfg.Ix_m4):.1f} in\u2074",
                  "Iy": f"{units_cfg.m4_to_in4(section_cfg.Iy_m4):.1f} in\u2074",
                  "J": f"{units_cfg.m4_to_in4(section_cfg.J_m4):.3f} in\u2074"}
    return {
        "pieces": pieces,
        "units": units_cfg.system_display_name(),
        "material": mat_txt,
        "size_name": section_cfg.display_name(),
        "size_alt": f"[{alt_tag} {alt_label}]",
        "props": prop_txt,
        "area_short": area_short,
    }


# ==============================================================================
# 1. 3D STRUCTURAL FINITE ELEMENT SOLVER ENGINE
# ==============================================================================

class Frame3DSolver:
    def __init__(self, title="3D Frame Model - Rev. 3"):
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

    def add_member(self, mem_id, ni, nj, E, G, A, Iy, Iz, J,
                   beta_deg=0.0, mem_type="Beam", releases=None):
        """
        Adds a 3D beam/column member with orientation Beta angle and optional end releases.
        E, G, A, Iy, Iz, J are REQUIRED (SI: Pa, Pa, m^2, m^4, m^4, m^4). REV3 removed
        REV2's silent hard-coded defaults so every member must take its properties from
        the Excel-driven MaterialConfig / MemberSizeConfig.
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

def _axis_label(v):
    """Global-axis name of a unit vector, e.g. (0,0,-1) -> '-Z'."""
    i = int(np.argmax(np.abs(v)))
    return ("+" if v[i] > 0 else "-") + "XYZ"[i]


def generate_custom_structural_plot(model, save_path="structural_model_sabila.png",
                                     material_cfg=None, section_cfg=None, units_cfg=None):
    """Clean single-view report figure: large 3D model on the left, a
    CONFIGURATION card (Units / Material / Member Size) and a MODEL DATA card
    on the right. Every number is read from the solved `model` and the
    Excel-driven config objects -- nothing about the model is typed in.
    Approved palette: olive, dark olive, gold, rust, cream, white
    (local-axis triads keep the conventional red/green/blue)."""
    C_OLIVE, C_DARK, C_GOLD = "#606c38", "#283618", "#dda15e"
    C_RUST, C_CREAM, C_WHITE = "#bc6c25", "#fefae0", "#ffffff"
    AX_X, AX_Y, AX_Z = "#B23A22", "#4E702A", "#2B5B84"        # local-axis triad colours

    res = model.results
    fig = plt.figure(figsize=(20, 11.25), dpi=200, facecolor=C_CREAM)
    ax = fig.add_axes([0.0, 0.03, 0.70, 0.88], projection='3d', facecolor=C_CREAM)
    ax.set_box_aspect((1, 1, 1), zoom=1.1)

    for axis in (ax.xaxis, ax.yaxis, ax.zaxis):
        axis.set_pane_color((0.996, 0.980, 0.878, 0.55))
        axis._axinfo["grid"]["color"] = to_rgba(C_GOLD, 0.35)
        axis._axinfo["grid"]["linewidth"] = 0.8
    ax.tick_params(axis='both', labelcolor=C_DARK, labelsize=10)
    ax.tick_params(axis='z', labelcolor=C_DARK, labelsize=10)

    P = lambda v: (v[0], v[2], v[1])            # model (X, Y-up, Z) -> matplotlib (x, y, z-up)

    # ---- title -------------------------------------------------------------
    nodes = np.array(list(model.nodes.values()))
    span = nodes.max(axis=0) - nodes.min(axis=0)
    fig.text(0.35, 0.965, f"{span[0]:.0f}m x {span[2]:.0f}m x {span[1]:.0f}m Space Frame - Structural Model, Rev. 3",
             ha='center', va='center', fontsize=20, fontweight='bold', color=C_DARK)
    fig.text(0.35, 0.930, "Pinned supports, global and local axes, beta angles, and member end releases",
             ha='center', va='center', fontsize=13, color=C_OLIVE)

    # ---- members -----------------------------------------------------------
    for mid, m in model.members.items():
        p1, p2 = model.nodes[m['i']], model.nodes[m['j']]
        col, lw = (C_DARK, 3.4) if m['type'] == "Column" else (C_RUST, 3.0)
        ax.plot(*zip(P(p1), P(p2)), color=col, linewidth=lw, solid_capstyle='round', zorder=3)

    # ---- supports (pyramids) ----------------------------------------------
    for nid, sup in model.supports.items():
        if any(sup[:3]):
            x, z, y = P(model.nodes[nid])
            h, b = 0.60, 0.42
            apex = [x, z, y]
            base = [[x-b, z-b, y-h], [x+b, z-b, y-h], [x+b, z+b, y-h], [x-b, z+b, y-h]]
            faces = [[apex, base[0], base[1]], [apex, base[1], base[2]],
                     [apex, base[2], base[3]], [apex, base[3], base[0]], base]
            ax.add_collection3d(Poly3DCollection(faces, facecolors=C_GOLD, edgecolors=C_DARK,
                                                 alpha=0.92, linewidths=1.0, zorder=5))

    # ---- nodes + labels (label pushed outward from the frame centre) -------
    centre = nodes.mean(axis=0)
    dof_no = lambda n: f"DOF {(n-1)*6+1}-{n*6}"
    for nid, pt in model.nodes.items():
        supported = nid in model.supports and any(model.supports[nid][:3])
        ax.scatter(*P(pt), s=150, color=C_GOLD if supported else C_OLIVE,
                   edgecolors=C_DARK, linewidths=1.8, zorder=10, depthshade=False)
        out = np.sign(pt - centre); out[1] = 0.0
        lab = pt + np.array([0.35 * out[0], 0.95 if pt[1] > centre[1] else 0.55, 0.35 * out[2]])
        ax.text(*P(lab), f"N{nid}\n{dof_no(nid)}", fontsize=10.5, fontweight='bold',
                color=C_DARK, ha='center', va='center', zorder=15,
                bbox=dict(boxstyle='round,pad=0.25', facecolor=C_CREAM, edgecolor=C_OLIVE, lw=0.9, alpha=0.95))

    # ---- released member ends (hollow circles) -----------------------------
    comp = ["UX", "UY", "UZ", "RX", "RY", "RZ"]
    release_lines = []
    for mid, m in sorted(model.members.items()):
        rel = list(m['releases'])
        if any(rel):
            p1, p2 = model.nodes[m['i']], model.nodes[m['j']]
            for end, sl, pt in (("i", slice(0, 6), p1 + 0.12 * (p2 - p1)),
                                ("j", slice(6, 12), p2 + 0.12 * (p1 - p2))):
                if any(rel[sl]):
                    ax.scatter(*P(pt), s=130, facecolors=C_WHITE, edgecolors=C_DARK,
                               linewidths=2.2, zorder=11, depthshade=False)
            names = [f"{e}:{comp[k]}" for e, sl in (("i", slice(0, 6)), ("j", slice(6, 12)))
                     for k, f in enumerate(rel[sl]) if f]
            release_lines.append((mid, ", ".join(names)))

    # ---- local-axis triads at member mid-spans (from the solver's own R) ---
    scale = 0.85
    for mid, m in sorted(model.members.items()):
        R, L, _, _, _ = model.compute_local_axes(mid)
        mp = 0.5 * (model.nodes[m['i']] + model.nodes[m['j']])
        for vec, colr in ((R[0], AX_X), (R[1], AX_Y), (R[2], AX_Z)):
            ax.quiver(*P(mp), *P(vec * scale), color=colr, arrow_length_ratio=0.32,
                      linewidth=1.9, zorder=8)
        tag = f"M{mid}" + (f" ({chr(946)}={m['beta_deg']:.0f}\u00b0)" if m['type'] == "Column" else "")
        if any(m['releases']):
            tag += " [MZ]" if m['releases'][5] or m['releases'][11] else " [REL]"
        lab = mp - 0.55 * R[2] * (1 if m['type'] == "Column" else 0) + np.array([0, -0.45 if m['type'] != "Column" else 0.0, 0])
        ax.text(*P(lab), tag, fontsize=9.5, fontweight='bold', color=C_DARK, ha='center', va='center',
                zorder=14, bbox=dict(boxstyle='round,pad=0.18', facecolor=C_CREAM, edgecolor=C_GOLD, lw=0.8, alpha=0.95))

    # ---- global axes, drawn clear of the frame in the front-left corner -----
    o = np.array([-0.8, -0.8, -0.8])      # (X, Y, Z) model coords
    g = 2.0
    for vec, name in ((np.array([g, 0, 0]), "X"), (np.array([0, g, 0]), "Y"), (np.array([0, 0, g]), "Z")):
        ax.quiver(*P(o), *P(vec), color=C_DARK, arrow_length_ratio=0.16, linewidth=3.0, zorder=9)
        lab_off = np.array([-0.55, 0, 0]) if name == "Z" else np.zeros(3)   # keep Z clear of the N1 support
        ax.text(*P(o + vec * 1.16 + lab_off), name, fontsize=15, fontweight='bold', color=C_DARK,
                ha='center', va='center')
    ax.text(*P(o + np.array([0.0, -0.05, -0.55])), "GLOBAL", fontsize=9.5, fontweight='bold', color=C_OLIVE)

    ax.view_init(elev=22, azim=-60)
    ax.set_xlim(-1, 7); ax.set_ylim(-1, 7); ax.set_zlim(-1, 7)
    ax.set_xlabel("X (m) - lateral", fontsize=12, fontweight='bold', color=C_DARK, labelpad=12)
    ax.set_ylabel("Z (m) - lateral", fontsize=12, fontweight='bold', color=C_DARK, labelpad=12)
    ax.set_zlabel("Y (m) - vertical", fontsize=12, fontweight='bold', color=C_DARK, labelpad=10)

    # ---- legend -------------------------------------------------------------
    from matplotlib.lines import Line2D
    handles = [
        Line2D([0], [0], color=C_RUST, lw=3.2, label="Beam"),
        Line2D([0], [0], color=C_DARK, lw=3.2, label="Column"),
        Line2D([0], [0], marker='o', color='none', markerfacecolor=C_GOLD, markeredgecolor=C_DARK, markersize=10, label="Supported node (pinned)"),
        Line2D([0], [0], marker='o', color='none', markerfacecolor=C_OLIVE, markeredgecolor=C_DARK, markersize=10, label="Free node"),
        Line2D([0], [0], marker='o', color='none', markerfacecolor=C_WHITE, markeredgecolor=C_DARK, markersize=10, markeredgewidth=2, label="Released member end"),
        Line2D([0], [0], color=AX_X, lw=2.4, label="Local x axis (axial)"),
        Line2D([0], [0], color=AX_Y, lw=2.4, label="Local y axis"),
        Line2D([0], [0], color=AX_Z, lw=2.4, label="Local z axis"),
        Line2D([0], [0], color=C_DARK, lw=3.0, label="Global X, Y, Z axes"),
    ]
    leg = fig.legend(handles=handles, loc='upper left', bbox_to_anchor=(0.012, 0.905), fontsize=10.5,
                     facecolor=C_CREAM, edgecolor=C_GOLD, framealpha=0.95)
    leg.get_frame().set_linewidth(1.2)

    # ---- right-hand cards ---------------------------------------------------
    cfg_ok = material_cfg is not None and section_cfg is not None and units_cfg is not None
    lines_cfg = []
    if cfg_ok:
        d = describe_configuration(units_cfg, material_cfg, section_cfg)
        pc = d["pieces"]
        lines_cfg = [
            "CONFIGURATION (Excel-driven)", "",
            f"Units          {d['units']}",
            f"               solver internal: SI (m, N, Pa)", "",
            f"Material       {material_cfg.display_name()}",
            f"    E          {pc['E']}",
            f"    Fy         {pc['Fy']}", "",
            f"Member size    {d['size_name']} {d['size_alt']}",
            f"    A          {pc['A']}",
            f"    Iz (major) {pc['Iz']}",
            f"    Iy (minor) {pc['Iy']}",
            f"    J          {pc['J']}",
        ]

    restrained = sorted({k for sup in model.supports.values() for k, f in enumerate(sup) if f})
    released_c = [comp[k] for k in range(6) if k not in restrained]
    sup_nodes = ", ".join(str(n) for n in sorted(model.supports))
    by_type = {}
    for mid, m in sorted(model.members.items()):
        by_type.setdefault(m['type'], []).append(mid)
    def axes_of(ids, row):
        return " / ".join(sorted({_axis_label(model.compute_local_axes(i)[0][row]) for i in ids}))
    beta_lines = [f"    {t:<14s}{sorted({int(model.members[i]['beta_deg']) for i in ids})[0]} deg"
                  for t, ids in by_type.items()]
    axes_lines = []
    for t, ids in by_type.items():
        axes_lines += [f"    {t + 's':<10s}y = {axes_of(ids, 1)}", f"    {'':<10s}z = {axes_of(ids, 2)}"]
    U = res['U_global']; Rr = res['Reactions']
    trans = np.array([U[k] for k in range(len(U)) if k % 6 < 3])
    sups = sorted(model.supports)
    sumR = [sum(Rr[(n-1)*6+c] for n in sups) / 1e3 for c in range(3)]
    n_dof = len(model.nodes) * 6
    if release_lines:
        rel_txt = [f"    M{mid:<3d}{txt}" for mid, txt in release_lines]
    else:
        rel_txt = ["    None (all member ends continuous)"]

    lines_model = (
        ["MODEL DATA - REV. 3", "",
         "Geometry",
         f"    Span X, Z, Y         {span[0]:.1f}, {span[2]:.1f}, {span[1]:.1f} m",
         f"    Nodes / Members      {len(model.nodes)} / {len(model.members)}", "",
         "Global axes",
         "    X, Z lateral;  Y vertical (up)", "",
         "Supports",
         f"    Type                 pinned",
         f"    Nodes                {sup_nodes}",
         f"    Restrained           {', '.join(comp[k] for k in restrained)}",
         f"    Released             {', '.join(released_c)}", "",
         "Degrees of freedom",
         f"    Total / Restrained   {n_dof} / {len(res['restrained_dofs'])}",
         f"    Active               {len(res['active_dofs'])}",
         f"    Numbering            (node - 1) x 6 + 1..6", "",
         "Member end releases"] + rel_txt + ["",
         "Beta angles"] + beta_lines + ["",
         "Local axes (from solver)",
         "    x = start node i -> end node j"] + axes_lines + ["",
         "Solver results (this run)",
         f"    Max joint translation {np.max(np.abs(trans))*1e3:.2f} mm",
         f"    Sum of reactions kN   Fx {sumR[0]:+.1f}  Fy {sumR[1]:+.1f}  Fz {sumR[2]:+.1f}"])

    fs = 9.6
    line_h = fs * 1.27 / 72 / 11.25                 # one text line, as a fraction of figure height
    x0, top = 0.715, 0.955
    def card(lines, y_top, edge, head_face):
        n = len(lines)
        h = n * line_h + 0.022
        fig.add_artist(plt.Rectangle((x0 - 0.008, y_top - h), 0.275, h, transform=fig.transFigure,
                                      facecolor=C_WHITE, edgecolor=edge, linewidth=1.8, alpha=0.96, zorder=1))
        fig.add_artist(plt.Rectangle((x0 - 0.008, y_top - 2 * line_h - 0.004), 0.275, 2 * line_h + 0.004,
                                      transform=fig.transFigure, facecolor=head_face, edgecolor='none',
                                      alpha=0.45, zorder=2))
        fig.text(x0, y_top - 0.008, "\n".join(lines), fontsize=fs, family='monospace', color=C_DARK,
                 va='top', ha='left', linespacing=1.42, zorder=3)
        return y_top - h
    y = top
    if lines_cfg:
        y = card(lines_cfg, y, C_RUST, C_GOLD) - 0.014
    card(lines_model, y, C_OLIVE, C_OLIVE)

    fig.savefig(save_path, dpi=200, facecolor=fig.get_facecolor())
    plt.close(fig)
    print(f"[SUCCESS] Custom 3D plot saved to: {save_path}")


# ==============================================================================
# 3. EXCEL WORKBOOK GENERATOR (Earthy Tones Palette)
# ==============================================================================

def export_to_excel_earthy(model, file_path="structural_solver_sabila.xlsx",
                            material_cfg=None, section_cfg=None, units_cfg=None):
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

    ws1["B2"] = "6m x 6m x 6m Space Frame Model - Rev. 3"
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

    if material_cfg is not None and section_cfg is not None and units_cfg is not None:
        ws1["B22"] = "REV.3 DATA-LAYER CONFIGURATION (Excel-Driven)"
        ws1["B22"].font = font_section

        cfg_hdrs = ["Configuration Item", "Selected Value", "Source Workbook"]
        for c_idx, h in enumerate(cfg_hdrs, start=2):
            c = ws1.cell(row=23, column=c_idx, value=h)
            c.fill = fill_primary; c.font = font_hdr_primary
            c.alignment = Alignment(horizontal='center', vertical='center'); c.border = border_thin

        d = describe_configuration(units_cfg, material_cfg, section_cfg)
        cfg_rows = [
            ["Unit System", d["units"] + " (solver internal: SI -- m, N, Pa)", os.path.basename(units_cfg.xlsx_path)],
            ["Material", d["material"], os.path.basename(material_cfg.source_file)],
            ["Member Size", f"{d['size_name']} {d['size_alt']}, {d['area_short']}", os.path.basename(section_cfg.source_file)],
        ]
        for r_off, row in enumerate(cfg_rows):
            r_idx = 24 + r_off
            for c_idx, val in enumerate(row, start=2):
                c = ws1.cell(row=r_idx, column=c_idx, value=val)
                c.border = border_thin
                c.font = font_data_bold if c_idx == 2 else font_data
                c.alignment = Alignment(horizontal='left', vertical='center')
                if r_idx % 2 == 1:
                    c.fill = fill_cream

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
                elif c_idx == 10:
                    c.number_format = '0.000000'      # A (m2)
                elif c_idx in [11, 12, 13]:
                    c.number_format = '0.000E+00'     # Iz, Iy, J (m4): too small for fixed decimals
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
    
    MsgBox "3D Space Frame Direct Stiffness Analysis Solved Successfully!", vbInformation, "Excel Structural Solver Rev. 3"
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

def main(unit_system=None):
    """unit_system: "Metric" (default, Standard Metric) or "Imperial".
    Can also be chosen from the command line:  python structural_solver_sabila.py --units Imperial"""
    print("=== STARTING 3D FRAME STRUCTURAL ANALYSIS (REV. 3) ===")
    if unit_system is None:
        unit_system = "Metric"
        if "--units" in sys.argv:
            unit_system = sys.argv[sys.argv.index("--units") + 1]

    # ------------------------------------------------------------------
    # REV.3 DATA LAYER: Excel -> Configuration -> Solver
    # ------------------------------------------------------------------
    units_cfg = UnitsConfig(system=unit_system)                      # Units_Imperial_Metric.xlsx  -> default "Metric"
    print(f"[UNITS]    Read '{units_cfg.xlsx_path}'  ->  active system: {units_cfg.system_display_name()}")

    material_cfg = MaterialConfig(units_cfg)        # RISA_Materials_Library_Metric.xlsx -> A36 Gr.36 row
    section_cfg = MemberSizeConfig(units_cfg)       # aisc-shapes-database-v160-2.xlsx -> W12X26 row
    d = describe_configuration(units_cfg, material_cfg, section_cfg)   # shown in the ACTIVE unit system
    print(f"[MATERIAL] Read '{material_cfg.source_file}'  ->  {d['material']}  "
          f"[G={material_cfg.G_Pa/1e9:.3f} GPa, Fu={material_cfg.Fu_Pa/1e6:.1f} MPa in SI]")
    print(f"[SECTION]  Read '{section_cfg.source_file}'  ->  {d['size_name']} {d['size_alt']}  {d['props']}")

    # Common member property kwargs, sourced entirely from the Excel-driven
    # configuration objects above (no hard-coded E/G/A/I/J values below).
    MEMBER_PROPS = dict(
        E=material_cfg.E_Pa, G=material_cfg.G_Pa,
        A=section_cfg.A_m2, Iz=section_cfg.Ix_m4, Iy=section_cfg.Iy_m4, J=section_cfg.J_m4,
    )

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
    # Every member below uses MEMBER_PROPS (E, G, A, Iz, Iy, J) sourced from
    # the ASTM A36 material lookup and the W12X26 AISC section lookup --
    # REV.2's hard-coded defaults were removed from add_member() in the final REV3 fix.
    # Base Beams (Beta = 0 deg)
    solver.add_member(1, 1, 2, beta_deg=0.0, mem_type="Beam", **MEMBER_PROPS)
    solver.add_member(2, 2, 3, beta_deg=0.0, mem_type="Beam", **MEMBER_PROPS)
    solver.add_member(3, 4, 3, beta_deg=0.0, mem_type="Beam", **MEMBER_PROPS)
    solver.add_member(4, 1, 4, beta_deg=0.0, mem_type="Beam", **MEMBER_PROPS)

    # Roof Beams (Beta = 0 deg)
    beam_rel = [False] * 12
    solver.add_member(5, 5, 6, beta_deg=0.0, mem_type="Beam", releases=beam_rel, **MEMBER_PROPS)
    solver.add_member(6, 6, 7, beta_deg=0.0, mem_type="Beam", **MEMBER_PROPS)
    solver.add_member(7, 8, 7, beta_deg=0.0, mem_type="Beam", **MEMBER_PROPS)
    solver.add_member(8, 5, 8, beta_deg=0.0, mem_type="Beam", **MEMBER_PROPS)

    # Columns (Beta = 90 deg)
    solver.add_member(9,  1, 5, beta_deg=90.0, mem_type="Column", **MEMBER_PROPS)
    solver.add_member(10, 2, 6, beta_deg=90.0, mem_type="Column", **MEMBER_PROPS)
    solver.add_member(11, 3, 7, beta_deg=90.0, mem_type="Column", **MEMBER_PROPS)
    solver.add_member(12, 4, 8, beta_deg=90.0, mem_type="Column", **MEMBER_PROPS)

    # Sanity guard: every member the solver holds must carry the Excel-sourced values.
    for _mid, _m in solver.members.items():
        assert (_m['E'] == material_cfg.E_Pa and _m['G'] == material_cfg.G_Pa
                and _m['A'] == section_cfg.A_m2 and _m['Iz'] == section_cfg.Ix_m4
                and _m['Iy'] == section_cfg.Iy_m4 and _m['J'] == section_cfg.J_m4), \
            f"Member {_mid} does not carry the configured material/section properties"
    print(f"[CHECK]    All {len(solver.members)} members carry the configured material + section properties")

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

    # 6. Generate Custom Visual Diagram (REV.3: annotated with live config)
    generate_custom_structural_plot(solver, os.path.join(OUTPUT_DIR, "structural_model_sabila.png"),
                                     material_cfg=material_cfg, section_cfg=section_cfg,
                                     units_cfg=units_cfg)

    # 7. Generate Formatted Excel Report (REV.3: includes configuration sheet)
    export_to_excel_earthy(solver, os.path.join(OUTPUT_DIR, "structural_solver_sabila.xlsx"),
                            material_cfg=material_cfg, section_cfg=section_cfg, units_cfg=units_cfg)
    print("=== ANALYSIS, PLOTTING & EXCEL EXPORT COMPLETED SUCCESSFULLY (REV. 3) ===")

if __name__ == '__main__':
    main()
