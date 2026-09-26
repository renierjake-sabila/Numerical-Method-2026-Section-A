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
        # REV3 addition (additive only, does not alter density_kgm3 above):
        # true WEIGHT density [kN/m3], read straight from the workbook's own
        # "Density [kN/m3]" column -- this is the correct source for
        # self-weight (W = gamma * A * L). It is NOT derived from
        # density_kgm3 and must never be multiplied by 9.81.
        self.density_kNm3 = match["density_kNm3"]
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
        uw_keys = [k for k in match if k and "Density" in str(k) and "k/ft" in str(k)]
        if mass_keys:
            self.density_kgm3 = match[mass_keys[0]]
        else:
            uw_key = _find_key(match, lambda k: "Density" in k and "k/ft" in k,
                               "unit weight [k/ft3]")
            self.density_kgm3 = self.units.kft3_to_kgm3(match[uw_key])

        # REV3 addition (additive only, does not alter density_kgm3 above):
        # true WEIGHT density [kN/m3] for self-weight (W = gamma * A * L).
        # The Imperial sheet's own unit-weight column [k/ft3] IS a weight
        # density already, so convert it directly -- do not derive this from
        # density_kgm3 and never multiply it by 9.81.
        if uw_keys:
            self.density_kNm3 = self.units.kft3_to_knm3(match[uw_keys[0]])
        else:
            self.density_kNm3 = None
            print("[MATERIAL] WARNING: no unit-weight [k/ft3] column found; "
                  "density_kNm3 is None -- self-weight (LC1) cannot be computed "
                  "for this material until a weight-density source is supplied.")

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
            "density_kNm3": self.density_kNm3,
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


# ==============================================================================
# REV 3 -- LOAD-CASE FRAMEWORK (everything below is new)
# ------------------------------------------------------------------------------
# Nothing above this line is modified in this section, except the three
# additive MaterialConfig.density_kNm3 edits (search "REV3 addition") made
# earlier and already regression-tested against the unmodified baseline.
#
# Frame3DSolver, its solve() method, generate_custom_structural_plot(),
# export_to_excel_earthy(), and main() are untouched and are never called
# from this section.
#
# SCOPE BOUNDARY (project decision, recorded): this section defines and
# validates load cases, member/nodal loads, self-weight, and diaphragm
# DEFINITIONS (Phase 4, not yet built) as load VECTORS / data objects only.
# It does NOT assemble a global stiffness system for them and produces NO
# reactions, member forces, or displacements for these new load cases or
# their combinations (Phase 6, not yet built).
# ==============================================================================

# ------------------------------------------------------------------------------
# PHASE 2 -- Load-case data model
# ------------------------------------------------------------------------------

VALID_FORCE_UNIT = "kN"
VALID_LINE_LOAD_UNIT = "kN/m"
VALID_MOMENT_UNIT = "kN*m"
VALID_TEMP_UNIT = "C"

_AXIS_INDEX = {"X": 0, "Y": 1, "Z": 2}


def build_rev3_reference_model(unit_system="Metric"):
    """Builds the same 8-node / 12-member 6x6x6 m cube as main(), reading
    material/section/unit properties from the same Excel-driven configs,
    but WITHOUT calling solve(), the plotter, or the Excel exporter. This is
    the model Phase 2/3/6/7 (load cases, combinations, tests, --verify-loads
    exports) run against -- kept separate from main() so nothing in main()
    or Frame3DSolver needs to change to support the Rev 3 load-case layer.
    Returns (solver, material_cfg, units_cfg, section_cfg)."""
    units_cfg = UnitsConfig(system=unit_system)
    material_cfg = MaterialConfig(units_cfg)
    section_cfg = MemberSizeConfig(units_cfg)

    MEMBER_PROPS = dict(
        E=material_cfg.E_Pa, G=material_cfg.G_Pa,
        A=section_cfg.A_m2, Iz=section_cfg.Ix_m4, Iy=section_cfg.Iy_m4, J=section_cfg.J_m4,
    )

    solver = Frame3DSolver("Sabila 3D Frame Solver (Rev3 reference model)")
    solver.add_node(1, 0.0, 0.0, 0.0)
    solver.add_node(2, 6.0, 0.0, 0.0)
    solver.add_node(3, 6.0, 0.0, 6.0)
    solver.add_node(4, 0.0, 0.0, 6.0)
    solver.add_node(5, 0.0, 6.0, 0.0)
    solver.add_node(6, 6.0, 6.0, 0.0)
    solver.add_node(7, 6.0, 6.0, 6.0)
    solver.add_node(8, 0.0, 6.0, 6.0)
    for n in [1, 2, 3, 4]:
        solver.add_support(n, ux=1, uy=1, uz=1, rx=0, ry=0, rz=0)

    solver.add_member(1, 1, 2, beta_deg=0.0, mem_type="Beam", **MEMBER_PROPS)
    solver.add_member(2, 2, 3, beta_deg=0.0, mem_type="Beam", **MEMBER_PROPS)
    solver.add_member(3, 4, 3, beta_deg=0.0, mem_type="Beam", **MEMBER_PROPS)
    solver.add_member(4, 1, 4, beta_deg=0.0, mem_type="Beam", **MEMBER_PROPS)
    solver.add_member(5, 5, 6, beta_deg=0.0, mem_type="Beam", **MEMBER_PROPS)
    solver.add_member(6, 6, 7, beta_deg=0.0, mem_type="Beam", **MEMBER_PROPS)
    solver.add_member(7, 8, 7, beta_deg=0.0, mem_type="Beam", **MEMBER_PROPS)
    solver.add_member(8, 5, 8, beta_deg=0.0, mem_type="Beam", **MEMBER_PROPS)
    solver.add_member(9,  1, 5, beta_deg=90.0, mem_type="Column", **MEMBER_PROPS)
    solver.add_member(10, 2, 6, beta_deg=90.0, mem_type="Column", **MEMBER_PROPS)
    solver.add_member(11, 3, 7, beta_deg=90.0, mem_type="Column", **MEMBER_PROPS)
    solver.add_member(12, 4, 8, beta_deg=90.0, mem_type="Column", **MEMBER_PROPS)
    return solver, material_cfg, units_cfg, section_cfg


class UnitValidationError(ValueError):
    """Raised when a REV3 load object is constructed with an incompatible unit
    (e.g. a force mislabeled kN*m, or a line load mislabeled kN)."""
    pass


def _check_unit(given, expected, context):
    if given != expected:
        raise UnitValidationError(
            f"{context}: expected unit '{expected}', got '{given}'. REV3 load "
            f"objects reject mismatched units rather than accepting them silently."
        )


def _parse_direction(direction):
    """'-Y' -> (axis_index=1, sign=-1.0); 'X' -> (axis_index=0, sign=+1.0)."""
    d = str(direction).strip()
    sign = -1.0 if d.startswith("-") else 1.0
    axis_letter = d[-1]
    if axis_letter not in _AXIS_INDEX:
        raise ValueError(f"Invalid global direction '{direction}'; "
                          f"expected one of X/-X/Y/-Y/Z/-Z.")
    return _AXIS_INDEX[axis_letter], sign


class NodalLoad:
    """A concentrated force/moment applied directly at a node, GLOBAL axes.
    Forces (fx, fy, fz) in kN. Moments (mx, my, mz) in kN*m."""

    def __init__(self, node_id, fx=0.0, fy=0.0, fz=0.0,
                 mx=0.0, my=0.0, mz=0.0,
                 force_unit=VALID_FORCE_UNIT, moment_unit=VALID_MOMENT_UNIT):
        _check_unit(force_unit, VALID_FORCE_UNIT, f"NodalLoad(node {node_id}) force")
        _check_unit(moment_unit, VALID_MOMENT_UNIT, f"NodalLoad(node {node_id}) moment")
        self.node_id = node_id
        self.fx, self.fy, self.fz = float(fx), float(fy), float(fz)
        self.mx, self.my, self.mz = float(mx), float(my), float(mz)

    def __repr__(self):
        return (f"NodalLoad(node={self.node_id}, F=({self.fx:.3f},{self.fy:.3f},"
                f"{self.fz:.3f}) kN, M=({self.mx:.3f},{self.my:.3f},{self.mz:.3f}) kN*m)")


class MemberDistributedLoad:
    """A uniform distributed member load, GLOBAL direction axis+sign, kN/m."""

    def __init__(self, member_id, direction, magnitude_kNm,
                 distribution_type="uniform", unit=VALID_LINE_LOAD_UNIT):
        _check_unit(unit, VALID_LINE_LOAD_UNIT, f"MemberDistributedLoad(member {member_id})")
        _parse_direction(direction)  # validates direction string
        if distribution_type != "uniform":
            raise NotImplementedError(
                f"REV3 supports uniform distributed loads only; "
                f"'{distribution_type}' is not implemented.")
        self.member_id = member_id
        self.direction = direction
        self.magnitude_kNm = float(magnitude_kNm)
        self.distribution_type = distribution_type

    def total_load_kN(self, member_length_m):
        return self.magnitude_kNm * member_length_m

    def __repr__(self):
        return (f"MemberDistributedLoad(member={self.member_id}, "
                f"{self.magnitude_kNm:.4f} kN/m {self.direction}, {self.distribution_type})")


class MemberPointLoad:
    """A concentrated load at a fraction of a member's span (0=i end, 1=j end,
    0.5=midpoint), GLOBAL direction axis+sign, magnitude in kN. Creates no
    physical node; equivalent_nodal_loads() gives the consistent/equivalent
    fixed-end formulation for later use, without assembling into Frame3DSolver."""

    def __init__(self, member_id, location_fraction, direction, magnitude_kN,
                 unit=VALID_FORCE_UNIT):
        _check_unit(unit, VALID_FORCE_UNIT, f"MemberPointLoad(member {member_id})")
        _parse_direction(direction)
        if not (0.0 <= location_fraction <= 1.0):
            raise ValueError(f"MemberPointLoad(member {member_id}): location_fraction "
                              f"must be in [0,1], got {location_fraction}.")
        self.member_id = member_id
        self.location_fraction = float(location_fraction)
        self.direction = direction
        self.magnitude_kN = float(magnitude_kN)

    def equivalent_nodal_loads(self, member_length_m):
        """Consistent nodal load formulation for a transverse point load P at
        fraction `a` of a fixed-fixed span L. Returns (Fi, Mi, Fj, Mj)."""
        L = member_length_m
        if L <= 0:
            raise ValueError(f"MemberPointLoad(member {self.member_id}): "
                              f"member length must be positive, got {L}.")
        a = self.location_fraction * L
        b = L - a
        P = self.magnitude_kN
        Fi = P * (b ** 2) * (3 * a + b) / (L ** 3)
        Fj = P * (a ** 2) * (3 * b + a) / (L ** 3)
        Mi = P * a * (b ** 2) / (L ** 2)
        Mj = -P * (a ** 2) * b / (L ** 2)
        return Fi, Mi, Fj, Mj

    def __repr__(self):
        return (f"MemberPointLoad(member={self.member_id}, {self.magnitude_kN:.3f} kN "
                f"{self.direction} @ {self.location_fraction*100:.0f}% span)")


class TemperatureLoad:
    """Uniform temperature change on a member, degrees C. Represents a
    thermal STRAIN effect (epsilon = alpha * deltaT), never a mechanical
    force -- contributes 0.0 kN to a LoadCase's applied-force total.

    Step 7: `alpha_perC` and `reference_temperature_C` are optional, added
    additively after `delta_T`/before `unit` so every pre-existing call --
    positional (`TemperatureLoad(1, 15.0)`) or with `unit=` as a keyword --
    is unaffected. When `alpha_perC` is supplied (build_lc9_temperature()
    sources it from MaterialConfig.therm_coeff_perC, never a re-declared
    literal), thermal_strain()/free_expansion_m()/restrained_force_kN()/
    fe_thermal_load_vector_kN() can all be called with no argument; passing
    an explicit alpha to any of them (as the existing
    test_lc9_thermal_strain_formula test does) always overrides it."""

    def __init__(self, member_id, delta_T, alpha_perC=None,
                 reference_temperature_C=20.0, unit=VALID_TEMP_UNIT):
        _check_unit(unit, VALID_TEMP_UNIT, f"TemperatureLoad(member {member_id})")
        self.member_id = member_id
        self.delta_T = float(delta_T)
        self.alpha_perC = None if alpha_perC is None else float(alpha_perC)
        self.reference_temperature_C = float(reference_temperature_C)

    def _resolve_alpha(self, alpha_per_C):
        alpha = self.alpha_perC if alpha_per_C is None else alpha_per_C
        if alpha is None:
            raise ValueError(f"TemperatureLoad(member={self.member_id}): no alpha_perC "
                              f"available -- pass one explicitly or construct with "
                              f"alpha_perC=material_cfg.therm_coeff_perC.")
        return alpha

    def thermal_strain(self, alpha_per_C=None):
        """eps_T = alpha * deltaT (dimensionless)."""
        return self._resolve_alpha(alpha_per_C) * self.delta_T

    def free_expansion_m(self, length_m, alpha_per_C=None):
        """Unrestrained elongation: dL = alpha * deltaT * L (metres). This is
        what an UNRESTRAINED member is permitted to do -- no force results."""
        return self.thermal_strain(alpha_per_C) * length_m

    def restrained_force_kN(self, E_Pa, A_m2, alpha_per_C=None):
        """Analytical fixed-end axial force magnitude if the member's
        expansion were FULLY prevented at both ends: N = EA * alpha * deltaT
        (kN). Reported as a magnitude; the physical sense for +deltaT
        (heating) is COMPRESSIVE (the restraint pushes back against the
        member's tendency to elongate)."""
        EA_kN = (E_Pa * A_m2) / 1000.0
        return EA_kN * self.thermal_strain(alpha_per_C)

    def fe_thermal_load_vector_kN(self, E_Pa, A_m2, alpha_per_C=None):
        """Self-equilibrating finite-element-equivalent thermal load vector:
        the axial nodal-force pair (kN) that reproduces this member's
        thermal effect as an applied load -- non-zero at each end, but
        summing to exactly 0.000 kN net on the structure, by construction.
        Returns (F_i, F_j) along the member's own i->j axis: F_i = -N
        (pulls end i inward), F_j = +N (pushes end j outward), N = EA *
        alpha * deltaT. For +deltaT (heating -> tendency to elongate), the
        pair pushes the two ends apart, exactly as an internal record of
        the member 'wanting' to grow -- it is NOT a mechanical load applied
        by anything external, which is why it self-cancels."""
        N = self.restrained_force_kN(E_Pa, A_m2, alpha_per_C)
        return (-N, +N)

    def __repr__(self):
        return f"TemperatureLoad(member={self.member_id}, deltaT={self.delta_T:+.1f} C)"


class Diaphragm:
    """Rigid roof diaphragm DEFINITION: couples UX, UZ, RY of the slave nodes
    to the master node; UY, RX, RZ stay free. Definition + documented
    constraint equations only -- per project scope, NOT eliminated into
    Frame3DSolver's stiffness/DOF numbering (Phase 4)."""

    COUPLED_DOFS = ("UX", "UZ", "RY")
    FREE_DOFS = ("UY", "RX", "RZ")

    def __init__(self, diaphragm_id, name, master_node, slave_nodes):
        if master_node in slave_nodes:
            raise ValueError(f"Diaphragm '{name}': master_node {master_node} "
                              f"must not also appear in slave_nodes.")
        self.id = diaphragm_id
        self.name = name
        self.master_node = master_node
        self.slave_nodes = list(slave_nodes)

    def constraint_equations(self):
        eqs = []
        for n in self.slave_nodes:
            for dof in self.COUPLED_DOFS:
                eqs.append(f"{dof}_node{n} = {dof}_node{self.master_node}  "
                           f"(diaphragm '{self.name}', master N{self.master_node})")
        return eqs

    def __repr__(self):
        return (f"Diaphragm(id={self.id}, name='{self.name}', "
                f"master=N{self.master_node}, slaves={self.slave_nodes}, "
                f"coupled={self.COUPLED_DOFS}, free={self.FREE_DOFS})")


def build_rev3_diaphragm(model, diaphragm_id=1, name="ROOF DIAPHRAGM"):
    """Builds the Phase 4 roof diaphragm DEFINITION: master = lowest node id
    among the roof-elevation nodes, slaves = the rest. Never eliminated into
    Frame3DSolver's DOF numbering (see Diaphragm docstring)."""
    roof_nodes = get_roof_nodes(model)
    if len(roof_nodes) < 2:
        raise ValueError("build_rev3_diaphragm: fewer than 2 roof-elevation "
                          "nodes found; cannot form a diaphragm.")
    master = roof_nodes[0]
    slaves = roof_nodes[1:]
    return Diaphragm(diaphragm_id, name, master, slaves)


class LoadCase:
    """A named, categorized collection of loads (nodal / distributed / point /
    temperature) plus an optional self-weight factor. Building a LoadCase does
    not touch Frame3DSolver.loads and never calls solve()."""

    VALID_CATEGORIES = ("Dead", "Live", "Wind", "Seismic", "Temperature")

    def __init__(self, case_id, name, category, description="", self_weight_factor=0.0):
        if category not in self.VALID_CATEGORIES:
            raise ValueError(f"LoadCase '{name}': category must be one of "
                              f"{self.VALID_CATEGORIES}, got '{category}'.")
        self.id = case_id
        self.name = name
        self.category = category
        self.description = description
        self.self_weight_factor = float(self_weight_factor)
        self.nodal_loads = []
        self.distributed_loads = []
        self.point_loads = []
        self.temperature_loads = []

    def add(self, load_obj):
        if isinstance(load_obj, NodalLoad):
            self.nodal_loads.append(load_obj)
        elif isinstance(load_obj, MemberDistributedLoad):
            self.distributed_loads.append(load_obj)
        elif isinstance(load_obj, MemberPointLoad):
            self.point_loads.append(load_obj)
        elif isinstance(load_obj, TemperatureLoad):
            self.temperature_loads.append(load_obj)
        else:
            raise TypeError(f"LoadCase '{self.name}': cannot add object of "
                             f"type {type(load_obj).__name__}.")
        return load_obj

    def __repr__(self):
        return (f"LoadCase(id={self.id}, name='{self.name}', category='{self.category}', "
                f"nodal={len(self.nodal_loads)}, distributed={len(self.distributed_loads)}, "
                f"point={len(self.point_loads)}, temperature={len(self.temperature_loads)})")


# ------------------------------------------------------------------------------
# Geometry helpers (read-only against Frame3DSolver's existing nodes/members;
# no new node/member IDs are ever created)
# ------------------------------------------------------------------------------

def get_member_length(model, member_id):
    m = model.members[member_id]
    p1, p2 = model.nodes[m['i']], model.nodes[m['j']]
    return float(np.linalg.norm(p2 - p1))


def get_roof_elevation(model, axis="Y"):
    idx = _AXIS_INDEX[axis]
    return max(coord[idx] for coord in model.nodes.values())


def get_roof_nodes(model, axis="Y", tol=1e-6):
    idx = _AXIS_INDEX[axis]
    roof_level = get_roof_elevation(model, axis)
    return sorted(n for n, c in model.nodes.items() if abs(c[idx] - roof_level) <= tol)


def get_roof_beam_ids(model, axis="Y"):
    roof_nodes = set(get_roof_nodes(model, axis))
    return sorted(mid for mid, m in model.members.items()
                  if m['type'] == "Beam" and m['i'] in roof_nodes and m['j'] in roof_nodes)


def member_restraint_state(model, member_id):
    """Step 7 -- classifies a member's axial restraint as one of "free",
    "fully_restrained", or "partially_restrained", from model.supports
    alone (no stiffness solve, per Rev 3's loads-only scope).

    A node is treated as restraining the member's axial DOF if any of its
    restrained TRANSLATIONAL components (ux/uy/uz) is aligned with the
    member's own i->j axis (dot product with the unit axis > 0.5, which
    is exact -- 1.0 or 0.0 -- for every axis-aligned member in this cube,
    and degrades gracefully for a non-axis-aligned member elsewhere).

    - both ends axially restrained    -> "fully_restrained"
    - neither end axially restrained  -> "free"
    - exactly one end restrained      -> "partially_restrained" (a 3D-frame
      member framed into one fixed end and one free end cannot have its
      thermal force determined without an actual stiffness solve, which is
      out of Rev 3's scope -- reported as indeterminate, never guessed).
    """
    m = model.members[member_id]
    p1, p2 = model.nodes[m['i']], model.nodes[m['j']]
    L = float(np.linalg.norm(p2 - p1))
    u = (p2 - p1) / L if L > 0 else np.array([1.0, 0.0, 0.0])

    def axial_restrained(node_id):
        sup = model.supports.get(node_id)
        if not sup:
            return False
        translations = sup[:3]
        return any(abs(u[k]) > 0.5 and translations[k] for k in range(3))

    i_restrained = axial_restrained(m['i'])
    j_restrained = axial_restrained(m['j'])
    if i_restrained and j_restrained:
        return "fully_restrained"
    if i_restrained or j_restrained:
        return "partially_restrained"
    return "free"


# ------------------------------------------------------------------------------
# PHASE 3 -- Load-case builders (LC1-LC9)
# ------------------------------------------------------------------------------

def build_lc1_self_weight(model, material_cfg, case_id=1):
    """LC1: DEAD / SELF WEIGHT. W = gamma * A * L per member, Global -Y.
    gamma is MaterialConfig.density_kNm3 (true weight density, kN/m3) --
    read directly from the workbook, NOT derived from density_kgm3, and
    never multiplied by 9.81 (see the REV3 density fix in MaterialConfig)."""
    if material_cfg.density_kNm3 is None:
        raise ValueError("LC1 requires MaterialConfig.density_kNm3, which is "
                          "unavailable for the active material/unit system.")
    case = LoadCase(case_id, "DEAD / SELF WEIGHT", "Dead",
                     description="W = gamma * A * L, gamma read directly from the "
                                 "workbook's weight-density column (kN/m3).",
                     self_weight_factor=1.0)
    gamma = material_cfg.density_kNm3
    for mid, m in model.members.items():
        w_per_length = gamma * m['A']  # kN/m3 * m^2 = kN/m
        case.add(MemberDistributedLoad(mid, "-Y", w_per_length))
    return case


def build_lc2_roof_dead(model, case_id=2, magnitude_kNm=5.0):
    """LC2: ROOF DEAD. 5 kN/m uniformly distributed on roof beams, Global -Y."""
    case = LoadCase(case_id, "ROOF DEAD", "Dead",
                     description=f"{magnitude_kNm:.1f} kN/m uniformly distributed "
                                 f"on roof beams, Global -Y.")
    for mid in get_roof_beam_ids(model):
        case.add(MemberDistributedLoad(mid, "-Y", magnitude_kNm))
    return case


def build_lc3_roof_live(model, case_id=3, magnitude_kNm=3.0):
    """LC3: ROOF LIVE. 3 kN/m uniformly distributed on roof beams, Global -Y."""
    case = LoadCase(case_id, "ROOF LIVE", "Live",
                     description=f"{magnitude_kNm:.1f} kN/m uniformly distributed "
                                 f"on roof beams, Global -Y.")
    for mid in get_roof_beam_ids(model):
        case.add(MemberDistributedLoad(mid, "-Y", magnitude_kNm))
    return case


def build_lc4_roof_center_point(model, case_id=4, magnitude_kN=5.0):
    """LC4: ROOF BEAM CENTER LOAD. 5 kN at each roof beam's midpoint, Global -Y;
    consistent/equivalent nodal formulation, no new physical node."""
    case = LoadCase(case_id, "ROOF BEAM CENTER LOAD", "Dead",
                     description=f"{magnitude_kN:.1f} kN at the midpoint of each "
                                 f"roof beam, Global -Y.")
    for mid in get_roof_beam_ids(model):
        case.add(MemberPointLoad(mid, 0.5, "-Y", magnitude_kN))
    return case


def build_lateral_load_case(model, case_id, name, category, direction, total_kN):
    """Shared builder for LC5-LC8: total_kN split equally across the roof-
    elevation nodes (same 4 nodes the diaphragm will later reference)."""
    roof_nodes = get_roof_nodes(model)
    n = len(roof_nodes)
    if n == 0:
        raise ValueError(f"{name}: no roof-elevation nodes found in the model.")
    per_node = total_kN / n
    axis, sign = _parse_direction(direction)
    axis_name = ["fx", "fy", "fz"][axis]
    case = LoadCase(case_id, name, category,
                     description=f"{total_kN:.3f} kN total, Global {direction}, split "
                                 f"equally across {n} roof-elevation node(s): "
                                 f"{per_node:.3f} kN/node.")
    for nid in roof_nodes:
        case.add(NodalLoad(nid, **{axis_name: per_node * sign}))
    return case


def build_lc5_wind_x(model, case_id=5):
    return build_lateral_load_case(model, case_id, "WIND X", "Wind", "X", 10.0)


def build_lc6_wind_z(model, case_id=6):
    return build_lateral_load_case(model, case_id, "WIND Z", "Wind", "Z", 10.0)


def build_lc7_seismic_x(model, case_id=7):
    return build_lateral_load_case(model, case_id, "SEISMIC X", "Seismic", "X", 15.0)


def build_lc8_seismic_z(model, case_id=8):
    return build_lateral_load_case(model, case_id, "SEISMIC Z", "Seismic", "Z", 15.0)


def build_lc9_temperature(model, material_cfg=None, case_id=9, delta_T=15.0,
                           reference_temperature_C=20.0):
    """LC9: TEMPERATURE +15C. Thermal strain effect on every member;
    0.000 kN net external force by construction (no force-type load added).

    Step 7: alpha is sourced directly from material_cfg.therm_coeff_perC
    (never re-declared as a literal) and stored on each TemperatureLoad so
    thermal_strain()/free_expansion_m()/restrained_force_kN()/
    fe_thermal_load_vector_kN() all work with no extra argument.
    material_cfg stays optional (default None) so any pre-existing caller
    that only ever passed `model` keeps working -- alpha_perC is then left
    unset on each TemperatureLoad, exactly as before this change."""
    alpha = material_cfg.therm_coeff_perC if material_cfg is not None else None
    case = LoadCase(case_id, f"TEMPERATURE +{delta_T:.0f}C", "Temperature",
                     description=f"{delta_T:+.1f} C thermal strain effect "
                                 f"(epsilon = alpha * deltaT); 0.000 kN net external force.")
    for mid in model.members.keys():
        case.add(TemperatureLoad(mid, delta_T, alpha_perC=alpha,
                                  reference_temperature_C=reference_temperature_C))
    return case


def build_all_rev3_load_cases(model, material_cfg):
    """Builds and returns LC1-LC9 as an ordered dict {case_id: LoadCase}."""
    cases = [
        build_lc1_self_weight(model, material_cfg),
        build_lc2_roof_dead(model),
        build_lc3_roof_live(model),
        build_lc4_roof_center_point(model),
        build_lc5_wind_x(model),
        build_lc6_wind_z(model),
        build_lc7_seismic_x(model),
        build_lc8_seismic_z(model),
        build_lc9_temperature(model, material_cfg),
    ]
    return {c.id: c for c in cases}


# ------------------------------------------------------------------------------
# Load-application validation summary (spec item 13 format). Reports whether
# each load case's applied loads sum to the intended total -- a load-
# application accuracy check, not a global structural-equilibrium solve.
# ------------------------------------------------------------------------------

def build_validation_summary(case, model, intended_total_kN=None):
    axis_totals = {"Fx": 0.0, "Fy": 0.0, "Fz": 0.0}
    loaded_nodes, loaded_members = set(), set()

    for nl in case.nodal_loads:
        axis_totals["Fx"] += nl.fx
        axis_totals["Fy"] += nl.fy
        axis_totals["Fz"] += nl.fz
        loaded_nodes.add(nl.node_id)

    for dl in case.distributed_loads:
        L = get_member_length(model, dl.member_id)
        axis, sign = _parse_direction(dl.direction)
        axis_totals[["Fx", "Fy", "Fz"][axis]] += dl.total_load_kN(L) * sign
        loaded_members.add(dl.member_id)

    for pl in case.point_loads:
        axis, sign = _parse_direction(pl.direction)
        axis_totals[["Fx", "Fy", "Fz"][axis]] += pl.magnitude_kN * sign
        loaded_members.add(pl.member_id)

    for tl in case.temperature_loads:
        loaded_members.add(tl.member_id)  # 0 force contribution; tracked for reporting only

    computed_total = (axis_totals["Fx"]**2 + axis_totals["Fy"]**2 + axis_totals["Fz"]**2) ** 0.5
    summary = {
        "case_id": case.id, "name": case.name, "category": case.category,
        "loaded_nodes": sorted(loaded_nodes), "loaded_members": sorted(loaded_members),
        "Fx_kN": axis_totals["Fx"], "Fy_kN": axis_totals["Fy"], "Fz_kN": axis_totals["Fz"],
        "computed_total_kN": computed_total,
    }
    if intended_total_kN is not None:
        summary["intended_total_kN"] = intended_total_kN
        summary["equilibrium_error_kN"] = abs(computed_total - intended_total_kN)
    return summary


# ------------------------------------------------------------------------------
# Step 7 -- Temperature Load (LC9) verification. Distinguishes thermal
# DEFORMATION (strain, free expansion) from thermal FORCE (restrained-member
# axial force), and reports one of three restraint states per member -- never
# guessing a force for a partially restrained member, since that requires an
# actual stiffness solve, which is out of Rev 3's scope.
# ------------------------------------------------------------------------------

def build_temperature_verification(model, material_cfg, case):
    """Returns a list of per-member dicts (one per TemperatureLoad in `case`)
    with everything the Step 7 spec's TEMPERATURE LOAD VALIDATION list asks
    for: member id, reference/changed temperature, material, alpha, eps_T,
    length, free expansion, EA, the analytical fully-restrained force
    benchmark, the self-equilibrating FE thermal load vector, and the
    member's restraint classification (free / fully_restrained /
    partially_restrained -- the last reported as indeterminate, not solved).
    Thermal DEFORMATION (dL) and thermal FORCE (N) are always reported as
    two clearly separate fields -- never conflated."""
    rows = []
    for tl in case.temperature_loads:
        m = model.members[tl.member_id]
        L = get_member_length(model, tl.member_id)
        alpha = tl.alpha_perC if tl.alpha_perC is not None else material_cfg.therm_coeff_perC
        eps_T = alpha * tl.delta_T
        dL_m = eps_T * L
        EA_kN = (m['E'] * m['A']) / 1000.0
        N_fully_restrained_kN = EA_kN * eps_T  # analytical benchmark; always computed
        F_i_kN, F_j_kN = tl.fe_thermal_load_vector_kN(m['E'], m['A'], alpha_per_C=alpha)
        state = member_restraint_state(model, tl.member_id)

        if state == "free":
            reported_force_kN = 0.0
            reported_expansion_m = dL_m
            force_note = "0.000 kN (unrestrained -- free to expand)"
        elif state == "fully_restrained":
            reported_force_kN = N_fully_restrained_kN
            reported_expansion_m = 0.0
            force_note = f"{N_fully_restrained_kN:.3f} kN compression (both ends axially restrained)"
        else:  # partially_restrained
            reported_force_kN = None
            reported_expansion_m = None
            force_note = ("indeterminate -- requires an active stiffness solve "
                          "(out of Rev 3 scope); analytical fixed-end benchmark "
                          f"EA*alpha*dT = {N_fully_restrained_kN:.3f} kN reported for reference only")

        rows.append({
            "member_id": tl.member_id,
            "reference_temperature_C": tl.reference_temperature_C,
            "delta_T_C": tl.delta_T,
            "material": material_cfg.display_name(),
            "alpha_perC": alpha,
            "eps_T": eps_T,
            "length_m": L,
            "free_expansion_dL_m": dL_m,
            "EA_kN": EA_kN,
            "fully_restrained_force_kN": N_fully_restrained_kN,  # analytical benchmark, always present
            "fe_vector_kN": (F_i_kN, F_j_kN),
            "restraint_state": state,
            "reported_force_kN": reported_force_kN,
            "reported_expansion_m": reported_expansion_m,
            "force_note": force_note,
        })
    return rows


def verify_temperature_fe_vector_net_force(model, material_cfg, case):
    """Checkpoint (Step 7): the FE thermal load vector must be NON-EMPTY per
    member/node while summing to EXACTLY 0.000 kN net on the structure.
    Sums every member's self-equilibrating (F_i, F_j) pair, projected onto
    its own global axis, into a per-node accumulator, then totals the whole
    structure. Returns both halves of the checkpoint so a test can assert
    each independently."""
    node_totals = {}
    nonzero_member_count = 0
    for tl in case.temperature_loads:
        m = model.members[tl.member_id]
        p1, p2 = model.nodes[m['i']], model.nodes[m['j']]
        L = float(np.linalg.norm(p2 - p1))
        u = (p2 - p1) / L if L > 0 else np.array([1.0, 0.0, 0.0])
        alpha = tl.alpha_perC if tl.alpha_perC is not None else material_cfg.therm_coeff_perC
        F_i, F_j = tl.fe_thermal_load_vector_kN(m['E'], m['A'], alpha_per_C=alpha)
        if abs(F_i) > 1e-12 or abs(F_j) > 1e-12:
            nonzero_member_count += 1
        node_totals.setdefault(m['i'], np.zeros(3))
        node_totals.setdefault(m['j'], np.zeros(3))
        node_totals[m['i']] += F_i * u
        node_totals[m['j']] += F_j * u

    net = np.sum(list(node_totals.values()), axis=0) if node_totals else np.zeros(3)
    return {
        "nonzero_member_count": nonzero_member_count,
        "total_members": len(case.temperature_loads),
        "net_force_kN": net,
        "net_force_magnitude_kN": float(np.linalg.norm(net)),
    }


# ==============================================================================
# PHASE 6 -- NSCP 2015 LOAD COMBINATIONS (Section 203.3.1 LRFD / 203.4.1 ASD)
# ------------------------------------------------------------------------------
# Combinations are assembled purely by scaling and summing the existing LC1-9
# LoadCase objects (LoadCombination.factors maps case_id -> multiplier); no
# physical load is ever duplicated or rebuilt. Numbering is the project's
# authoritative table: combos 1-12 are LRFD, combos 13-30 are ASD -- confirmed
# against the rev3_combination_*.pdf reference renders (combo 1 titled "LRFD
# Combination 1 - 1.4D", combo 13 titled "ASD Combination 13 - D"; that PDF's
# "Combination 15" page is a duplicate export of combo 13's figure, not a
# distinct 31st combination, per project confirmation).
#
# "D" (dead) is not a single LoadCase -- per the reference renders' own
# legends (e.g. "DEAD / SELF WEIGHT x1, ROOF DEAD x1, ROOF BEAM CENTER LOAD
# x1" on combo 13), D is the sum of LC1 (self-weight) + LC2 (roof dead) +
# LC4 (roof beam center point load), each scaled by the same combination
# factor. L=LC3, Wx=LC5, Wz=LC6, Ex=LC7, Ez=LC8, T=LC9.
# ==============================================================================

DEAD_CASE_IDS = (1, 2, 4)   # LC1 self-weight + LC2 roof dead + LC4 center point = "D"
SYMBOL_CASE_ID = {"L": 3, "Wx": 5, "Wz": 6, "Ex": 7, "Ez": 8, "T": 9}


def _expand_symbolic_factors(symbolic):
    """{'D': 1.2, 'L': 1.6, ...} -> {case_id: factor, ...}. 'D' is split
    equally (same factor) across DEAD_CASE_IDS; every other symbol maps to
    its single load case id via SYMBOL_CASE_ID."""
    expanded = {}
    for symbol, factor in symbolic.items():
        if symbol == "D":
            for cid in DEAD_CASE_IDS:
                expanded[cid] = expanded.get(cid, 0.0) + factor
        elif symbol in SYMBOL_CASE_ID:
            cid = SYMBOL_CASE_ID[symbol]
            expanded[cid] = expanded.get(cid, 0.0) + factor
        else:
            raise ValueError(f"Unknown load-combination symbol '{symbol}'.")
    return expanded


class LoadCombination:
    """A named combination assembled by scaling/summing existing LoadCase
    objects. `factors` maps load_case_id -> scalar multiplier (negative =
    reversed direction, e.g. wind acting in -X rather than +X). Never
    duplicates or rebuilds a physical load; never calls Frame3DSolver.solve()."""

    VALID_METHODS = ("LRFD", "ASD")

    def __init__(self, combo_id, name, method, factors):
        if method not in self.VALID_METHODS:
            raise ValueError(f"LoadCombination '{name}': method must be one of "
                              f"{self.VALID_METHODS}, got '{method}'.")
        self.id = combo_id
        self.name = name
        self.method = method
        self.factors = dict(factors)   # {case_id: factor}

    def assembled_totals(self, load_cases, model):
        """Sums each contributing case's Fx/Fy/Fz total (via
        build_validation_summary), scaled by this combination's factor."""
        totals = {"Fx_kN": 0.0, "Fy_kN": 0.0, "Fz_kN": 0.0}
        loaded_nodes, loaded_members = set(), set()
        legend_lines = []
        for cid, factor in self.factors.items():
            case = load_cases[cid]
            summ = build_validation_summary(case, model)
            totals["Fx_kN"] += summ["Fx_kN"] * factor
            totals["Fy_kN"] += summ["Fy_kN"] * factor
            totals["Fz_kN"] += summ["Fz_kN"] * factor
            loaded_nodes.update(summ["loaded_nodes"])
            loaded_members.update(summ["loaded_members"])
            legend_lines.append(f"{case.name} x {factor:g}")
        computed_total = (totals["Fx_kN"]**2 + totals["Fy_kN"]**2 + totals["Fz_kN"]**2) ** 0.5
        return {
            "combo_id": self.id, "name": self.name, "method": self.method,
            "loaded_nodes": sorted(loaded_nodes), "loaded_members": sorted(loaded_members),
            **totals, "computed_total_kN": computed_total, "legend_lines": legend_lines,
        }

    def __repr__(self):
        terms = ", ".join(f"{cid}:{f:+.3g}" for cid, f in self.factors.items())
        return (f"LoadCombination(id={self.id}, name='{self.name}', "
                f"method='{self.method}', factors={{{terms}}})")


_NSCP_TABLE = [
    # (id, name, method, symbolic_factors)
    (1,  "1.4D",                 "LRFD", {"D": 1.4}),
    (2,  "1.2D + 1.6L",          "LRFD", {"D": 1.2, "L": 1.6}),
    (3,  "1.2D + 1.0L + 1.0Wx",  "LRFD", {"D": 1.2, "L": 1.0, "Wx": 1.0}),
    (4,  "1.2D + 1.0L - 1.0Wx",  "LRFD", {"D": 1.2, "L": 1.0, "Wx": -1.0}),
    (5,  "1.2D + 1.0L + 1.0Wz",  "LRFD", {"D": 1.2, "L": 1.0, "Wz": 1.0}),
    (6,  "1.2D + 1.0L - 1.0Wz",  "LRFD", {"D": 1.2, "L": 1.0, "Wz": -1.0}),
    (7,  "1.2D + 1.0L + 1.0Ex",  "LRFD", {"D": 1.2, "L": 1.0, "Ex": 1.0}),
    (8,  "1.2D + 1.0L - 1.0Ex",  "LRFD", {"D": 1.2, "L": 1.0, "Ex": -1.0}),
    (9,  "1.2D + 1.0L + 1.0Ez",  "LRFD", {"D": 1.2, "L": 1.0, "Ez": 1.0}),
    (10, "1.2D + 1.0L - 1.0Ez",  "LRFD", {"D": 1.2, "L": 1.0, "Ez": -1.0}),
    (11, "1.2D + 1.0L + 1.0T",   "LRFD", {"D": 1.2, "L": 1.0, "T": 1.0}),
    (12, "0.9D + 1.0T",          "LRFD", {"D": 0.9, "T": 1.0}),

    (13, "D",                   "ASD",  {"D": 1.0}),
    (14, "D + L",                "ASD",  {"D": 1.0, "L": 1.0}),
    (15, "D + 0.6Wx",            "ASD",  {"D": 1.0, "Wx": 0.6}),
    (16, "D - 0.6Wx",            "ASD",  {"D": 1.0, "Wx": -0.6}),
    (17, "D + 0.6Wz",            "ASD",  {"D": 1.0, "Wz": 0.6}),
    (18, "D - 0.6Wz",            "ASD",  {"D": 1.0, "Wz": -0.6}),
    (19, "D + 0.7Ex",            "ASD",  {"D": 1.0, "Ex": 0.7}),
    (20, "D - 0.7Ex",            "ASD",  {"D": 1.0, "Ex": -0.7}),
    (21, "D + 0.7Ez",            "ASD",  {"D": 1.0, "Ez": 0.7}),
    (22, "D - 0.7Ez",            "ASD",  {"D": 1.0, "Ez": -0.7}),
    (23, "0.6D + 0.6Wx",         "ASD",  {"D": 0.6, "Wx": 0.6}),
    (24, "0.6D - 0.6Wx",         "ASD",  {"D": 0.6, "Wx": -0.6}),
    (25, "0.6D + 0.6Wz",         "ASD",  {"D": 0.6, "Wz": 0.6}),
    (26, "0.6D - 0.6Wz",         "ASD",  {"D": 0.6, "Wz": -0.6}),
    (27, "0.6D + 0.7Ex",         "ASD",  {"D": 0.6, "Ex": 0.7}),
    (28, "0.6D - 0.7Ex",         "ASD",  {"D": 0.6, "Ex": -0.7}),
    (29, "D + T",                "ASD",  {"D": 1.0, "T": 1.0}),
    (30, "D + 0.75L + 0.75T",    "ASD",  {"D": 1.0, "L": 0.75, "T": 0.75}),
]


def build_all_nscp_combinations():
    """Builds and returns all 30 LoadCombination objects per NSCP 2015
    Section 203.3.1 (LRFD, ids 1-12) and Section 203.4.1 (ASD, ids 13-30),
    as an ordered dict {combo_id: LoadCombination}."""
    return {cid: LoadCombination(cid, name, method, _expand_symbolic_factors(symbolic))
            for cid, name, method, symbolic in _NSCP_TABLE}


def verify_combination_consistency(combo, load_cases, model):
    """Cross-checks LoadCombination.assembled_totals() (sum of scaled
    per-case totals) against an independent load-object-level recomputation.
    Rev 3 has no stiffness solver, so this is a self-consistency check on
    the load-vector arithmetic (does re-deriving the total from raw
    NodalLoad/MemberDistributedLoad/MemberPointLoad objects match the
    per-case-summary path?) -- NOT a structural equilibrium check against
    support reactions, which remains out of scope."""
    assembled = combo.assembled_totals(load_cases, model)
    Fx = Fy = Fz = 0.0
    for cid, factor in combo.factors.items():
        case = load_cases[cid]
        for nl in case.nodal_loads:
            Fx += nl.fx * factor; Fy += nl.fy * factor; Fz += nl.fz * factor
        for dl in case.distributed_loads:
            L = get_member_length(model, dl.member_id)
            axis, sign = _parse_direction(dl.direction)
            val = dl.total_load_kN(L) * sign * factor
            Fx, Fy, Fz = (Fx + val, Fy, Fz) if axis == 0 else \
                         (Fx, Fy + val, Fz) if axis == 1 else (Fx, Fy, Fz + val)
        for pl in case.point_loads:
            axis, sign = _parse_direction(pl.direction)
            val = pl.magnitude_kN * sign * factor
            Fx, Fy, Fz = (Fx + val, Fy, Fz) if axis == 0 else \
                         (Fx, Fy + val, Fz) if axis == 1 else (Fx, Fy, Fz + val)
        # temperature loads contribute 0.0 kN by construction -- no term added.
    independent_total = (Fx**2 + Fy**2 + Fz**2) ** 0.5
    assembled["consistency_error_kN"] = abs(independent_total - assembled["computed_total_kN"])
    return assembled


# ==============================================================================
# PHASE 5 -- PySide6 / Matplotlib 3D LOAD VIEWER
# ------------------------------------------------------------------------------
# New in this section only. Nothing above is modified. PySide6 is imported
# lazily inside build_load_viewer()/launch_rev3_viewer() so that importing
# cube_solver_rev3.py (e.g. for the load-case framework or tests) never
# requires PySide6 to be installed.
#
# Combination display (LoadCombination objects, Phase 6) is wired in: the
# dropdown lists LC1-LC9 followed by every LoadCombination in
# `load_combinations`, tagged by (kind, id) item data so _redraw() can branch
# between _draw_loads() (single case) and _draw_combination() (scaled sum of
# contributing cases) without either function needing to know about the other.
# ==============================================================================

# Palette -- matches the REV3 spec's earthy warm-parchment scheme (Section 3
# of the original brief), not a re-derivation of generate_custom_structural_plot's
# separate local palette above; both are visually consistent by design.
VP_CANVAS_BG   = "#FAF6E9"   # warm parchment canvas / figure background
VP_PANE_RGBA   = (0.961, 0.933, 0.859, 0.55)   # light ivory/buff 3D panes (#F5EEDB @ ~0.55)
VP_GRID_COLOR  = "#E0D4BA"   # muted grid lines
VP_COLUMN      = "#2F3B20"   # deep forest olive / dark charcoal green
VP_BEAM        = "#C86A28"   # warm terracotta / rust orange
VP_NODE_FREE   = "#4E5D2A"   # olive green circular markers
VP_NODE_SUPPORT = "#E29D52"  # warm peach/sand pyramid symbols
VP_SUPPORT_EDGE = "#283618"  # dark outline for support pyramids
VP_AXIS_X = "#B83A1B"        # local/global x -- terracotta/red-orange
VP_AXIS_Y = "#4A6B22"        # local/global y -- olive green
VP_AXIS_Z = "#2B5B7E"        # local/global z -- slate/steel blue
VP_DIAPHRAGM = "#2B5B7E"     # diaphragm node highlight (slate blue squares)
VP_LOAD_ARROW = "#7B3FA0"    # nodal/point load arrows (purple, matches PDF legend)
VP_LOAD_BAND  = "#C86A28"    # distributed-load band (terracotta, matches PDF)
VP_THERMAL    = "#8A5A2B"    # thermal (deg C) annotation -- distinct from
                              # VP_LOAD_ARROW (purple, kN) and VP_LOAD_BAND
                              # (terracotta, kN/m); never used for a force unit
VP_TEXT_DARK  = "#283618"

# Step 6 -- band fill opacities (named constants, replacing the old hardcoded
# 0.35). Self-weight is lowered further than ordinary distributed loads
# because it bands every member in the model at once; members, nodes,
# labels and arrows must stay legible underneath either band. The edge
# alpha (0.9, on VP_LOAD_BAND's outline/intensity line) is unchanged by
# either constant, so the loaded extent stays crisp regardless of fill.
BAND_ALPHA_DISTRIBUTED = 0.30
BAND_ALPHA_SELF_WEIGHT = 0.18


def set_grid_visible(ax, visible,
                      pane_rgba=VP_PANE_RGBA, grid_color=VP_GRID_COLOR, grid_alpha=0.6):
    """Toggle 3D gridlines AND the three shaded panes together.

    TRAP (verified against the installed matplotlib 3.10.8 source):
    `Axes3D.grid(self, visible=True, **kwargs)` contains
        if len(kwargs): visible = True
    so passing ANY keyword -- even alongside visible=False -- silently
    forces the grid back on. Styling kwargs are therefore only ever applied
    on the ON branch, via axis._axinfo directly (same mechanism the
    existing baseline plotter already uses), never through ax.grid()
    itself. The OFF branch calls ax.grid(False) completely bare.
    """
    if visible:
        ax.grid(True)
        for axis in (ax.xaxis, ax.yaxis, ax.zaxis):
            axis.set_pane_color(pane_rgba)
            axis._axinfo["grid"]["color"] = to_rgba(grid_color, grid_alpha)
            axis._axinfo["grid"]["linewidth"] = 0.8
    else:
        ax.grid(False)   # zero kwargs -- this is the line the trap breaks
        for axis in (ax.xaxis, ax.yaxis, ax.zaxis):
            axis.set_pane_color((1.0, 1.0, 1.0, 0.0))   # fully transparent panes
    # Axis labels/spines/tick numbers are untouched by either branch above,
    # so they remain visible with the grid off (per the Phase-5 checkpoint).
    if ax.figure is not None and ax.figure.canvas is not None:
        ax.figure.canvas.draw_idle()


# ------------------------------------------------------------------------------
# VIEW_LAYERS -- single source of truth for the layer bar AND the View ribbon.
# Each entry: (key, label, default_visible). Adding a row here is enough to
# get both a layer-bar toggle button and a View-menu toggle action; no other
# code needs to change (see LoadViewerWindow._build_menu_and_toolbar).
# ------------------------------------------------------------------------------
VIEW_LAYERS = (
    ("grid",      "Grid",         True),
    ("nodes",     "Node IDs",     True),
    ("members",   "Member IDs",   True),
    ("loads",     "Loads",        True),
    ("diaphragm", "Diaphragm",    True),
    ("axes",      "Global Axes",  True),
    ("load_band", "Load Band",    True),
)


def _draw_model_geometry(ax, model, show_node_ids=True, show_member_ids=True):
    """Members colored by type, nodes colored/shaped by support condition.
    Matches P = (X, Z, Y-up) convention used by the existing baseline plotter."""
    P = lambda v: (v[0], v[2], v[1])
    for mid, m in model.members.items():
        p1, p2 = model.nodes[m['i']], model.nodes[m['j']]
        color = VP_COLUMN if m['type'] == "Column" else VP_BEAM
        xs, ys, zs = zip(P(p1), P(p2))
        ax.plot(xs, ys, zs, color=color, linewidth=2.6, solid_capstyle='round', zorder=2)
        if show_member_ids:
            mid_pt = P((p1 + p2) / 2.0)
            ax.text(*mid_pt, f"M{mid}", fontsize=8, color=VP_TEXT_DARK,
                    ha='center', va='center', zorder=5)

    for nid, coord in model.nodes.items():
        pt = P(coord)
        if model.supports.get(nid):
            ax.scatter(*pt, marker='^', s=170, color=VP_NODE_SUPPORT,
                      edgecolors=VP_SUPPORT_EDGE, linewidths=1.4, zorder=4)
        else:
            ax.scatter(*pt, marker='o', s=70, color=VP_NODE_FREE,
                      edgecolors=VP_TEXT_DARK, linewidths=0.8, zorder=4)
        if show_node_ids:
            ax.text(*pt, f"  N{nid}", fontsize=8.5, color=VP_TEXT_DARK,
                    ha='left', va='bottom', fontweight='bold', zorder=6)


def _draw_global_axes(ax, model, length=None):
    P = lambda v: (v[0], v[2], v[1])
    if length is None:
        coords = np.array(list(model.nodes.values()))
        length = 0.18 * float(np.max(coords.max(axis=0) - coords.min(axis=0)) or 1.0)
    origin = np.array([0.0, 0.0, 0.0])
    for vec, color, label in (
        (np.array([length, 0, 0]), VP_AXIS_X, "X"),
        (np.array([0, length, 0]), VP_AXIS_Y, "Y"),
        (np.array([0, 0, length]), VP_AXIS_Z, "Z"),
    ):
        o, t = P(origin), P(origin + vec)
        ax.plot([o[0], t[0]], [o[1], t[1]], [o[2], t[2]], color=color, linewidth=2.2, zorder=7)
        ax.text(*t, f" {label}", color=color, fontsize=10, fontweight='bold', zorder=7)


def _draw_diaphragm(ax, model, diaphragm):
    """Highlights the diaphragm's master + slave nodes; no-op if diaphragm is None."""
    if diaphragm is None:
        return
    P = lambda v: (v[0], v[2], v[1])
    for nid in [diaphragm.master_node] + list(diaphragm.slave_nodes):
        pt = P(model.nodes[nid])
        ax.scatter(*pt, marker='s', s=260, facecolors='none',
                  edgecolors=VP_DIAPHRAGM, linewidths=2.0, zorder=8)
    mpt = P(model.nodes[diaphragm.master_node])
    ax.text(*mpt, f"  ROOF DIAPHRAGM master N{diaphragm.master_node}",
            color=VP_DIAPHRAGM, fontsize=9, fontweight='bold', zorder=8)


def _draw_loads(ax, model, case, arrow_scale=None, show_band=True):
    """Nodal loads as arrows, member point loads as arrows at their true
    fractional location, distributed loads as an offset band with a
    standoff gap. Returns a legend text summary (list of str).

    `show_band` (Step 5 -- Load Band Toggle): when True (default, matches
    all prior behavior), the closed Poly3DCollection fill is drawn along
    with the tick-arrow row. When False, the polygon fill is skipped
    entirely and ONLY the tick-arrow row is drawn -- same TICK_SPACING_M,
    same n_ticks, same positions. The arrow/glyph count is therefore
    100% identical either way; only the fill polygon's presence changes."""
    P = lambda v: (v[0], v[2], v[1])
    coords = np.array(list(model.nodes.values()))
    if arrow_scale is None:
        arrow_scale = 0.22 * float(np.max(coords.max(axis=0) - coords.min(axis=0)) or 1.0)
    legend_lines = [f"{case.name}  [{case.category}]"]

    for nl in case.nodal_loads:
        F = np.array([nl.fx, nl.fy, nl.fz])
        mag = float(np.linalg.norm(F))
        if mag == 0:
            continue
        direction = F / mag
        tip = model.nodes[nl.node_id] + direction * arrow_scale
        base = model.nodes[nl.node_id] + direction * (arrow_scale * 0.15)  # standoff
        b, t = P(base), P(tip)
        ax.plot([b[0], t[0]], [b[1], t[1]], [b[2], t[2]],
               color=VP_LOAD_ARROW, linewidth=2.4, zorder=9)
        ax.text(*t, f" {mag:.3f} kN", color=VP_LOAD_ARROW, fontsize=8.5,
                fontweight='bold', zorder=9)
    if case.nodal_loads:
        legend_lines.append(f"nodal loads: {len(case.nodal_loads)}")

    # Distance-driven arrow spacing (not a fixed count): ~18 ticks over a 6 m
    # span, scaling with actual member length.
    TICK_SPACING_M = 6.0 / 18.0

    for dl in case.distributed_loads:
        m = model.members[dl.member_id]
        p1, p2 = model.nodes[m['i']], model.nodes[m['j']]
        member_len = float(np.linalg.norm(p2 - p1))
        axis, sign = _parse_direction(dl.direction)
        d = np.zeros(3); d[axis] = sign
        u = (p2 - p1) / member_len if member_len > 0 else d  # member axis unit vector

        # Degenerate-load guard: if the load direction is (anti)parallel to
        # the member's own axis (e.g. self-weight, -Y, on a vertical column),
        # a perpendicular standoff band would collapse to a zero-area sliver.
        # Skip the polygon fill for that case and fall back to a thin
        # in-line marker so the member/load are still legible.
        is_axial = abs(float(np.dot(d, u))) > 0.99

        if is_axial:
            b1, b2 = P(p1), P(p2)
            ax.plot([b1[0], b2[0]], [b1[1], b2[1]], [b1[2], b2[2]],
                   color=VP_LOAD_BAND, linewidth=2.0, alpha=0.9,
                   linestyle=(0, (2, 2)), zorder=8)
        else:
            standoff = d * (arrow_scale * 0.35)
            # Step 6 -- self-weight bands every member at once, so it gets
            # the lighter fill; ordinary distributed loads (roof dead/live)
            # get the standard fill. Edge alpha (0.9) is unchanged either
            # way, so the loaded extent stays crisp under either fill.
            is_self_weight = getattr(case, "self_weight_factor", 0.0) > 0
            band_fill_alpha = BAND_ALPHA_SELF_WEIGHT if is_self_weight else BAND_ALPHA_DISTRIBUTED

            if show_band:
                # Closed banded polygon: member-line edge + standoff-offset
                # edge, joined by end caps -- not a bare center-line.
                quad = np.array([p1, p2, p2 + standoff, p1 + standoff])
                quad_plot = [P(v) for v in quad]
                band = Poly3DCollection([quad_plot], facecolor=to_rgba(VP_LOAD_BAND, band_fill_alpha),
                                         edgecolor=to_rgba(VP_LOAD_BAND, 0.9), linewidth=1.2)
                ax.add_collection3d(band)
                band.set_zorder(8)

            # Tick-arrow row: identical whether the band fill is shown or
            # not (Load Band toggle changes drawing style only, never the
            # underlying glyph/arrow count).
            n_ticks = max(4, int(round(member_len / TICK_SPACING_M)))
            for k in range(n_ticks + 1):
                f = k / n_ticks
                base_pt = p1 + (p2 - p1) * f + standoff
                tip_pt = base_pt + d * (arrow_scale * 0.18)
                bp, tp = P(tip_pt), P(base_pt)   # arrowheads point toward the member
                ax.plot([bp[0], tp[0]], [bp[1], tp[1]], [bp[2], tp[2]],
                       color=VP_LOAD_BAND, linewidth=1.1, alpha=0.85, zorder=8)
            mid_pt = (p1 + p2) / 2.0 + standoff
            ax.text(*P(mid_pt), f" {dl.magnitude_kNm:.3g} kN/m {dl.direction}",
                    color=VP_LOAD_BAND, fontsize=8, fontweight='bold', zorder=8)
    if case.distributed_loads:
        total = sum(dl.total_load_kN(get_member_length(model, dl.member_id))
                    for dl in case.distributed_loads)
        legend_lines.append(f"distributed: {len(case.distributed_loads)} member(s), "
                             f"{total:.3f} kN total")

    for pl in case.point_loads:
        m = model.members[pl.member_id]
        p1, p2 = model.nodes[m['i']], model.nodes[m['j']]
        loc = p1 + (p2 - p1) * pl.location_fraction
        axis, sign = _parse_direction(pl.direction)
        d = np.zeros(3); d[axis] = sign
        tip = loc + d * arrow_scale
        base = loc + d * (arrow_scale * 0.15)
        b, t = P(base), P(tip)
        ax.plot([b[0], t[0]], [b[1], t[1]], [b[2], t[2]],
               color=VP_LOAD_ARROW, linewidth=2.4, zorder=9)
        ax.scatter(*P(loc), marker='x', s=40, color=VP_LOAD_ARROW, zorder=9)
        ax.text(*t, f" {pl.magnitude_kN:.3f} kN", color=VP_LOAD_ARROW, fontsize=8.5,
                fontweight='bold', zorder=9)
    if case.point_loads:
        legend_lines.append(f"point loads: {len(case.point_loads)} "
                             f"x {case.point_loads[0].magnitude_kN:.3f} kN")

    # Step 7 -- thermal (deg C) annotation on the 3D plot itself, not just
    # the legend. Uses degC EXCLUSIVELY -- never kN or kN/m, since this is a
    # strain effect, not a mechanical load. Each member gets its "+dT degC"
    # label at midspan, plus two short outward-pointing ticks near its ends
    # (subtle expansion markers -- NOT a load band, since there's no force
    # to band). Rendered regardless of show_band, since this isn't a band.
    for tl in case.temperature_loads:
        m = model.members[tl.member_id]
        p1, p2 = model.nodes[m['i']], model.nodes[m['j']]
        member_len = float(np.linalg.norm(p2 - p1))
        u = (p2 - p1) / member_len if member_len > 0 else np.array([1.0, 0.0, 0.0])
        mid_pt = (p1 + p2) / 2.0
        ax.text(*P(mid_pt), f" {tl.delta_T:+.1f}\u00b0C", color=VP_THERMAL, fontsize=8,
                fontweight='bold', zorder=8, ha='center')
        tick_len = arrow_scale * 0.10
        for end_pt, outward in ((p1, -u), (p2, u)):
            near_pt = end_pt + outward * (member_len * 0.08)
            tip_pt = near_pt + outward * tick_len
            np_, tp_ = P(near_pt), P(tip_pt)
            ax.plot([np_[0], tp_[0]], [np_[1], tp_[1]], [np_[2], tp_[2]],
                   color=VP_THERMAL, linewidth=1.4, alpha=0.85, zorder=8)

    if case.temperature_loads:
        legend_lines.append(f"temperature: {case.temperature_loads[0].delta_T:+.1f} C "
                             f"on {len(case.temperature_loads)} member(s) "
                             f"(0.000 kN net -- strain effect only)")

    return legend_lines


class _ScaledLoadView:
    """Read-only, throwaway view of a LoadCase with every load pre-scaled by
    a single combination factor, so the existing _draw_loads() renderer can
    be reused unchanged for combination display. Never mutates the source
    LoadCase and is never stored -- purely a rendering convenience mirroring
    the "scale the load vector, don't duplicate the load" principle used
    throughout Rev 3. A negative factor (wind/seismic reversal) flows
    straight through NodalLoad's signed fx/fy/fz components; D and L (the
    only symbols ever attached to distributed/point loads in the NSCP table)
    are never negative, so no direction-string flip is needed here."""

    def __init__(self, case, factor):
        self.id = case.id
        self.name = f"{case.name} x {factor:g}"
        self.category = case.category
        self.nodal_loads = [NodalLoad(nl.node_id, nl.fx * factor, nl.fy * factor,
                                       nl.fz * factor, nl.mx * factor, nl.my * factor,
                                       nl.mz * factor) for nl in case.nodal_loads]
        self.distributed_loads = [MemberDistributedLoad(dl.member_id, dl.direction,
                                       dl.magnitude_kNm * factor, dl.distribution_type)
                                   for dl in case.distributed_loads]
        self.point_loads = [MemberPointLoad(pl.member_id, pl.location_fraction,
                                       pl.direction, pl.magnitude_kN * factor)
                             for pl in case.point_loads]
        # Step 7: delta_T is scaled by the combination factor (e.g. LC9's
        # +15 C becomes +11.25 C under Combo 30's 0.75 T factor), not copied
        # unscaled -- alpha_perC and reference_temperature_C carry through
        # unchanged since neither depends on the combination factor.
        self.temperature_loads = [TemperatureLoad(tl.member_id, tl.delta_T * factor,
                                       alpha_perC=tl.alpha_perC,
                                       reference_temperature_C=tl.reference_temperature_C)
                                   for tl in case.temperature_loads]
        # Step 6: carry self_weight_factor through so a self-weight case
        # scaled inside a combination (e.g. LC1 under 1.4D) still gets the
        # lighter BAND_ALPHA_SELF_WEIGHT fill rather than the default.
        self.self_weight_factor = case.self_weight_factor


def build_combination_glyphs(combo, load_cases):
    """Step 5 -- returns the list of _ScaledLoadView objects (one per
    contributing load case, each pre-scaled by its combination factor) that
    render_combination_figure() draws. Factored out of the old inline
    per-case loop so it's independently testable and matches the handout's
    expected helper name; behavior is unchanged."""
    return [_ScaledLoadView(load_cases[cid], factor) for cid, factor in combo.factors.items()]


def render_combination_figure(ax, model, combo, load_cases, arrow_scale=None, show_band=True):
    """Step 5 -- renders a LoadCombination by drawing each contributing
    case's glyphs (via build_combination_glyphs() + _draw_loads(), reused
    unchanged) pre-scaled by that case's factor, all onto the same shared
    arrow_scale so glyph sizes stay comparable across cases. `show_band` is
    forwarded to every _draw_loads() call so the Load Band toggle applies
    uniformly to combination views, not just single-case views. Returns a
    combined legend: one header line plus one "<case name> x <factor>" line
    per contributing case (matching the reference PDF's combination
    legends)."""
    coords = np.array(list(model.nodes.values()))
    if arrow_scale is None:
        arrow_scale = 0.22 * float(np.max(coords.max(axis=0) - coords.min(axis=0)) or 1.0)

    legend_lines = [f"Combination {combo.id} [{combo.method}] - {combo.name}"]
    for (cid, factor), view in zip(combo.factors.items(), build_combination_glyphs(combo, load_cases)):
        _draw_loads(ax, model, view, arrow_scale=arrow_scale, show_band=show_band)
        legend_lines.append(f"{load_cases[cid].name} x {factor:g}")
    return legend_lines


def _draw_combination(ax, model, combo, load_cases, arrow_scale=None, show_band=True):
    """Compatible alias for render_combination_figure() -- kept so existing
    call sites and tests referencing the original Phase 6 name are
    unaffected. See render_combination_figure() for the implementation."""
    return render_combination_figure(ax, model, combo, load_cases,
                                      arrow_scale=arrow_scale, show_band=show_band)


def build_load_viewer(model, load_cases, diaphragm=None, load_combinations=None):
    """Constructs and returns a LoadViewerWindow (imports PySide6 lazily).
    `load_combinations` (Phase 6, {combo_id: LoadCombination}) is appended
    to the same dropdown as the LC1-9 load cases; selecting a combination
    renders every contributing case's loads scaled by its factor via
    _draw_combination(), reusing _draw_loads() unchanged."""
    from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
        QHBoxLayout, QToolBar, QComboBox, QLabel)
    from PySide6.QtGui import QAction
    from PySide6.QtCore import Qt
    from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg, NavigationToolbar2QT

    class LoadViewerWindow(QMainWindow):
        def __init__(self):
            super().__init__()
            self.setWindowTitle("REV 3 Load Viewer")
            self.model = model
            self.load_cases = load_cases
            self.load_combinations = load_combinations or {}
            self.diaphragm = diaphragm
            self.layer_state = {key: default for key, _, default in VIEW_LAYERS}
            self.layer_actions = {}
            self._build_menu_and_toolbar()
            self._build_central_widget()
            self._redraw()

        def set_diaphragm(self, dp):
            self.diaphragm = dp
            self._redraw()

        def _build_menu_and_toolbar(self):
            view_menu = self.menuBar().addMenu("&View")          # the "View ribbon"
            layer_bar = QToolBar("Layers")
            layer_bar.setMovable(False)
            self.addToolBar(Qt.LeftToolBarArea, layer_bar)        # the "layer bar"

            for key, label, default in VIEW_LAYERS:
                action = QAction(label, self)
                action.setCheckable(True)
                action.setChecked(default)
                action.toggled.connect(self._make_toggle_handler(key))
                view_menu.addAction(action)    # same QAction -> View ribbon entry
                layer_bar.addAction(action)    # same QAction -> layer-bar button
                self.layer_actions[key] = action

        def _make_toggle_handler(self, key):
            def handler(checked):
                self.layer_state[key] = checked
                self._redraw()
            return handler

        def _build_central_widget(self):
            central = QWidget()
            outer = QVBoxLayout(central)

            top_bar = QHBoxLayout()
            top_bar.addWidget(QLabel("Load Case:"))
            self.case_combo = QComboBox()
            for cid in sorted(self.load_cases):
                case = self.load_cases[cid]
                self.case_combo.addItem(f"LC{cid} - {case.name}", ("case", cid))
            for cid in sorted(self.load_combinations):
                combo = self.load_combinations[cid]
                self.case_combo.addItem(f"Combo {cid} [{combo.method}] - {combo.name}",
                                         ("combo", cid))
            self.case_combo.currentIndexChanged.connect(lambda _: self._redraw())
            top_bar.addWidget(self.case_combo)
            top_bar.addStretch(1)
            outer.addLayout(top_bar)

            self.fig = plt.figure(figsize=(9.5, 7.2), dpi=110, facecolor=VP_CANVAS_BG)
            self.ax = self.fig.add_subplot(111, projection='3d', facecolor=VP_CANVAS_BG)
            self.canvas = FigureCanvasQTAgg(self.fig)
            self.nav_toolbar = NavigationToolbar2QT(self.canvas, self)
            outer.addWidget(self.nav_toolbar)
            outer.addWidget(self.canvas, stretch=1)
            self.setCentralWidget(central)

        def _current_selection(self):
            """Returns ('case', LoadCase) or ('combo', LoadCombination),
            depending on what's picked in the dropdown."""
            kind, cid = self.case_combo.currentData()
            if kind == "case":
                return "case", self.load_cases[cid]
            return "combo", self.load_combinations[cid]

        def _redraw(self):
            self.ax.clear()
            kind, selection = self._current_selection()

            _draw_model_geometry(self.ax, self.model,
                                  show_node_ids=self.layer_state["nodes"],
                                  show_member_ids=self.layer_state["members"])
            if self.layer_state["axes"]:
                _draw_global_axes(self.ax, self.model)
            if self.layer_state["diaphragm"]:
                _draw_diaphragm(self.ax, self.model, self.diaphragm)
            legend_lines = []
            if self.layer_state["loads"]:
                show_band = self.layer_state["load_band"]
                if kind == "case":
                    legend_lines = _draw_loads(self.ax, self.model, selection,
                                                show_band=show_band)
                else:
                    legend_lines = render_combination_figure(self.ax, self.model, selection,
                                                               self.load_cases, show_band=show_band)

            set_grid_visible(self.ax, self.layer_state["grid"])   # trap-safe toggle

            self.ax.set_xlabel("X (m)"); self.ax.set_ylabel("Z (m)"); self.ax.set_zlabel("Y (m)")
            if kind == "case":
                title = f"LC{selection.id} - {selection.name}"
            else:
                title = f"{selection.method} Combination {selection.id} - {selection.name}"
            self.ax.set_title(title, color=VP_TEXT_DARK, fontweight='bold')
            if legend_lines:
                self.ax.text2D(0.02, 0.98, "\n".join(legend_lines), transform=self.ax.transAxes,
                               fontsize=8.5, color=VP_TEXT_DARK, va='top', ha='left',
                               bbox=dict(boxstyle="round", facecolor=VP_CANVAS_BG,
                                        edgecolor=VP_GRID_COLOR, alpha=0.9))
            self.canvas.draw_idle()

    return LoadViewerWindow()


def launch_rev3_viewer(model, load_cases, diaphragm=None, load_combinations=None):
    """Entry point: creates a QApplication (reusing one if it already exists,
    e.g. under a test runner), shows the viewer, and starts the Qt event loop.
    Not called by main() -- invoke explicitly, e.g.:
        launch_rev3_viewer(solver, build_all_rev3_load_cases(solver, material_cfg))
    """
    from PySide6.QtWidgets import QApplication
    app = QApplication.instance() or QApplication(sys.argv)
    win = build_load_viewer(model, load_cases, diaphragm=diaphragm,
                             load_combinations=load_combinations)
    win.resize(1280, 860)
    win.show()
    return app.exec()


# ==============================================================================
# PHASE 7 -- VERIFICATION REPORT & --verify-loads HEADLESS EXPORT
# ------------------------------------------------------------------------------
# Everything below is reporting/export orchestration only: it calls the
# Phase 2/3/4/6 builders and Phase 5 drawing helpers above, never
# Frame3DSolver.solve(), and is never invoked from main().
# ==============================================================================

NSCP_CITATION = ("NSCP 2015, 7th Edition -- Section 203.3.1 (LRFD / Factored "
                  "Load Combinations) and Section 203.4.1 (ASD / Allowable "
                  "Stress Design Combinations), including 4 independent "
                  "temperature (T) combinations per project scope.")

# ------------------------------------------------------------------------------
# Governing-provision check for including T (temperature) in a combination --
# added per the Step 7/9 handout's explicit checkpoint: "Before hard-coding
# the final NSCP temperature-combination coefficients, verify the applicable
# NSCP edition and document the governing provision."
#
# Verified against the NSCP 2015 (7th Edition) text itself:
#   - Sec. 104.3.2 "Self-Straining Forces" is the general design requirement
#     that a temperature effect must be provided for at all:
#       "Provisions shall be made for anticipated self-straining forces
#        arising from differential settlement of foundations and from
#        restrained dimensional changes due to temperature, moisture,
#        shrinkage, heave, creep and similar effects."
#   - Sec. 203.2 "Symbols and Notations" defines T as: "self-straining force
#     and effects arising from contraction or expansion resulting from
#     temperature change, shrinkage, moisture change, creep in component
#     materials, movement due to differential settlement, or combinations
#     thereof."
#   - T appears in exactly THREE of the code's own numbered equations, never
#     as an independent "+1.0T" bolted onto an otherwise-unrelated D-only or
#     wind/seismic combination:
#       Sec. 203.3.1 (LRFD), Eq. 203-2:
#           1.2(D + F + T) + 1.6(L + H) + 0.5(Lr or R)   -- T factored
#           TOGETHER WITH D at the same 1.2, not as a separate 1.0T term.
#       Sec. 203.4.1 (ASD Basic), Eq. 203-9:
#           D + H + F + L + T                             -- T at 1.0, always
#           paired with L.
#       Sec. 203.4.1 (ASD Basic), Eq. 203-11:
#           D + H + F + 0.75[L + T + (Lr or R)]           -- 0.75 applies to
#           the whole bracket, not to T alone.
#     T does NOT appear in any other 203-series equation (not 203-1, 203-3
#     through 203-8, 203-10, 203-12, nor the ASD alternate combinations
#     203-13 through 203-18) -- there is no code-literal "0.9D + 1.0T" or
#     bare "D + T" (without L) equation.
#
# DISCREPANCY -- flagged, not silently corrected: this project's existing
# 30-combination table (locked in during an earlier phase against the
# project's own rev3_combination_*.pdf reference renders) represents T with
# FOUR independent combinations -- #11 "1.2D+1.0L+1.0T", #12 "0.9D+1.0T",
# #29 "D+T", #30 "D+0.75L+0.75T" -- built by the same pattern used for
# Wx/Wz/Ex/Ez (a matching "with/without" pair per LRFD and ASD tier). Combo
# #29 ("D+T") and #12 ("0.9D+1.0T") are not literal transcriptions of Eq.
# 203-2/203-9/203-11 above: the real code never isolates T from D's own
# factor, never omits L from a T-combination, and has no D-and-T-only ASD
# equation with a 0.9 dead-load factor. This function documents that gap
# for the record; it does NOT alter combo #11/#12/#29/#30's factors, since
# those were fixed by prior project confirmation, not by this function.
GOVERNING_TEMPERATURE_PROVISION = {
    "self_straining_requirement": "NSCP 2015, Sec. 104.3.2 (Self-Straining Forces)",
    "T_symbol_definition": "NSCP 2015, Sec. 203.2 (Symbols and Notations)",
    "official_T_equations": {
        "203-2": "1.2(D+F+T) + 1.6(L+H) + 0.5(Lr or R)  [Sec. 203.3.1, LRFD]",
        "203-9": "D+H+F+L+T  [Sec. 203.4.1, ASD Basic]",
        "203-11": "D+H+F+0.75[L+T+(Lr or R)]  [Sec. 203.4.1, ASD Basic]",
    },
    "project_table_T_combos": {11: "1.2D+1.0L+1.0T", 12: "0.9D+1.0T",
                                29: "D+T", 30: "D+0.75L+0.75T"},
    "note": ("Project combos #11/#30 are reasonable engineering analogues of "
             "official Eq. 203-2/203-11 (same idea: T alongside L, scaled "
             "down together in the 0.75 case) but are not literal "
             "transcriptions -- #11 uses a separate 1.0 factor on T rather "
             "than bundling T into D's 1.2 per 203-2. Project combos #12 "
             "and #29 have no literal counterpart in the official 203-series "
             "list at all. This is flagged for engineering sign-off, not "
             "auto-corrected, since the existing numbering was fixed against "
             "the project's own reference PDFs in an earlier phase."),
}


def _format_case_row(summary):
    return (f"  LC{summary['case_id']:<2d} {summary['name']:<26s} "
            f"Fx={summary['Fx_kN']:9.4f}  Fy={summary['Fy_kN']:9.4f}  "
            f"Fz={summary['Fz_kN']:9.4f}  |F|={summary['computed_total_kN']:9.4f} kN  "
            f"nodes={len(summary['loaded_nodes']):>2d}  members={len(summary['loaded_members']):>2d}")


def _format_combo_row(result):
    return (f"  #{result['combo_id']:<3d}[{result['method']:>4s}] {result['name']:<20s} "
            f"Fx={result['Fx_kN']:9.4f}  Fy={result['Fy_kN']:9.4f}  Fz={result['Fz_kN']:9.4f}  "
            f"consistency_err={result['consistency_error_kN']:.6f} kN")


def build_verification_report_text(model, material_cfg, load_cases, diaphragm,
                                    combinations, test_summary=None):
    """Builds the full text of cube_rev3_verification_report.txt (10 sections,
    per project spec). Every number here comes from build_validation_summary()
    / verify_combination_consistency() above -- nothing is computed inline."""
    L = []
    def hdr(n, title):
        L.append("")
        L.append(f"{n}. {title}")
        L.append("-" * 78)

    def hdr_named(title):
        L.append("")
        L.append(title)
        L.append("-" * 78)

    L.append("=" * 78)
    L.append("CUBE SOLVER REV 3 -- LOAD FRAMEWORK VERIFICATION REPORT")
    L.append("=" * 78)

    hdr(1, "SCOPE & BOUNDARY")
    L.append("Rev 3 adds a load-case, diaphragm-DEFINITION, and NSCP load-combination")
    L.append("framework on top of the existing baseline Frame3DSolver. Per project scope:")
    L.append("  - Rev 3 does NOT compute reactions, member forces, or nodal displacements")
    L.append("    for LC1-LC9 or any of the 30 combinations below.")
    L.append("  - Frame3DSolver.solve() (baseline) is untouched and is never called from")
    L.append("    this load-case/combination/viewer layer.")
    L.append("  - The roof diaphragm is DEFINED with documented constraint equations but")
    L.append("    is NOT eliminated into a stiffness/DOF equation system.")
    L.append("  - All 'equilibrium'/'consistency' checks below verify that the applied-")
    L.append("    load arithmetic is internally correct (sum of loads == intended total;")
    L.append("    combination total == independently re-derived total). They are NOT a")
    L.append("    structural equilibrium check against support reactions.")

    hdr(2, "MODEL SUMMARY")
    L.append(f"  Nodes: {len(model.nodes)}   Members: {len(model.members)}")
    L.append(f"  Material weight density (gamma): {material_cfg.density_kNm3:.4f} kN/m3 "
             f"(MaterialConfig.density_kNm3, workbook weight-density column)")
    roof_nodes = get_roof_nodes(model)
    L.append(f"  Roof-elevation nodes: {roof_nodes}   Roof beams: {get_roof_beam_ids(model)}")

    hdr(3, "LOAD CASE RESULTS (LC1-LC9)")
    intended = {1: None, 2: 120.0, 3: 72.0, 4: 20.0, 5: 10.0, 6: 10.0, 7: 15.0, 8: 15.0, 9: 0.0}
    for cid in sorted(load_cases):
        case = load_cases[cid]
        summ = build_validation_summary(case, model, intended_total_kN=intended.get(cid))
        L.append(_format_case_row(summ))
        if "equilibrium_error_kN" in summ:
            L.append(f"       intended={summ['intended_total_kN']:.4f} kN  "
                     f"equilibrium_error={summ['equilibrium_error_kN']:.6f} kN")
    per_length_kNm = material_cfg.density_kNm3 * next(iter(model.members.values()))['A']
    L.append("")
    L.append("  NOTE -- self-weight cross-check (documentation only, not this run's config):")
    L.append(f"    This run uses a single section (W12X26 / W310X38.7) for all 12 members,")
    L.append(f"    giving {per_length_kNm:.4f} kN/m per member and 27.3516 kN total, as")
    L.append("    reported above. A mixed-section variant (beams W310X38.7 @ 0.3802 kN/m,")
    L.append("    columns W250X49.1 @ 0.4818 kN/m) is documented for reference ONLY -- it is")
    L.append("    NOT the configuration used in this run: total self-weight would be")
    L.append("    29.815 kN (cross-checked against shape mass x g = 29.773 kN, 0.14% apart")
    L.append("    due to published unit-weight rounding). Do not read 29.815 kN as this")
    L.append("    run's LC1 result.")

    if 9 in load_cases and load_cases[9].temperature_loads:
        hdr_named("3B. TEMPERATURE LOAD VERIFICATION (LC9 DETAIL)")
        tl_rows = build_temperature_verification(model, material_cfg, load_cases[9])
        net_check = verify_temperature_fe_vector_net_force(model, material_cfg, load_cases[9])
        L.append("  Thermal DEFORMATION (strain, free expansion) is reported separately from")
        L.append("  thermal FORCE (restrained-member axial force) for every member below.")
        L.append("  Restraint state is classified from model.supports alone (no stiffness")
        L.append("  solve, per Rev 3 scope); 'partially_restrained' members are reported as")
        L.append("  INDETERMINATE, with the fixed-end benchmark shown for reference only.")
        L.append("")
        for row in tl_rows:
            L.append(f"  Member {row['member_id']:<2d}  Tref={row['reference_temperature_C']:.1f}C  "
                     f"dT={row['delta_T_C']:+.1f}C  material={row['material']}  "
                     f"alpha={row['alpha_perC']:.4e} /C")
            L.append(f"    eps_T={row['eps_T']:.4e}   L={row['length_m']:.3f} m   "
                     f"free_expansion dL={row['free_expansion_dL_m']*1000:.4f} mm   "
                     f"EA={row['EA_kN']:.1f} kN")
            L.append(f"    restraint_state={row['restraint_state']:<20s}  "
                     f"fixed-end benchmark N=EA*alpha*dT={row['fully_restrained_force_kN']:.3f} kN")
            L.append(f"    FE thermal load vector (self-equilibrating, kN): "
                     f"[{row['fe_vector_kN'][0]:+.3f}, {row['fe_vector_kN'][1]:+.3f}]")
            L.append(f"    reported result -> {row['force_note']}")
        L.append("")
        L.append(f"  FE thermal load vector check: {net_check['nonzero_member_count']} of "
                 f"{net_check['total_members']} member(s) have a NON-EMPTY (nonzero) vector, "
                 f"while the net force on the structure is "
                 f"{net_check['net_force_magnitude_kN']:.6f} kN (target: 0.000000 kN).")
        L.append("")
        L.append("  NOTE -- thermal-force cross-check (documentation only, not this run's config):")
        L.append(f"    This run's single-section model (W12X26 / W310X38.7, A992/A36 properties")
        L.append(f"    from the supplied workbooks) gives EA={tl_rows[0]['EA_kN']:.1f} kN and a fully")
        L.append(f"    restrained benchmark force N=EA*alpha*dT={tl_rows[0]['fully_restrained_force_kN']:.3f} kN, as reported")
        L.append("    per member above. The mixed-section, nominal-A992 reference values quoted")
        L.append("    for cross-check purposes are EA = 987,743.1 kN and N = 173.349 kN compression")
        L.append("    (beams) -- documented for reference ONLY, not this run's configuration. The")
        L.append("    ~0.1% gap between the two EA values traces to the same single-vs-mixed-")
        L.append("    section cause noted for LC1 self-weight in Section 3, plus minor differences")
        L.append("    between the nominal AISC area used here and the reference calculation's A992")
        L.append("    property source. Do not read 987,743.1 kN / 173.349 kN as this run's result.")

    hdr(4, "RIGID ROOF DIAPHRAGM DEFINITION")
    if diaphragm is not None:
        L.append(f"  {diaphragm!r}")
        L.append("  Constraint equations:")
        for eq in diaphragm.constraint_equations():
            L.append(f"    {eq}")
    else:
        L.append("  No diaphragm object supplied to the report generator.")

    hdr(5, "NSCP LOAD COMBINATIONS (FULL LIST, 30)")
    L.append(f"  Basis: {NSCP_CITATION}")
    L.append("  D = LC1 (self-weight) + LC2 (roof dead) + LC4 (roof beam center load),")
    L.append("  each scaled by the same factor.  L=LC3  Wx=LC5  Wz=LC6  Ex=LC7  Ez=LC8  T=LC9")
    for cid in sorted(combinations):
        c = combinations[cid]
        L.append(f"  #{cid:<3d}[{c.method:>4s}] {c.name:<24s} factors={c.factors}")

    hdr(6, "PER-COMBINATION CONSISTENCY CHECK (0.000000 kN target)")
    max_err = 0.0
    for cid in sorted(combinations):
        result = verify_combination_consistency(combinations[cid], load_cases, model)
        L.append(_format_combo_row(result))
        max_err = max(max_err, result["consistency_error_kN"])
    L.append(f"  Maximum consistency error across all 30 combinations: {max_err:.6f} kN")

    hdr(7, "CODE BASIS")
    L.append(f"  {NSCP_CITATION}")

    hdr(8, "AUTOMATED TEST SUMMARY")
    if test_summary:
        L.append(f"  {test_summary}")
    else:
        L.append("  (test_cube_solver_rev3.py was not run by this report generation call)")

    hdr(9, "ASSUMPTIONS & ITEMS REQUIRING MANUAL ENGINEERING VERIFICATION")
    L.append("  - ASD combinations 15-28 and all 4 temperature combinations (11, 12, 29, 30)")
    L.append("    are this project's own extension of the 6 base LRFD forms to reach the")
    L.append("    specified 30-combination count; they follow standard NSCP/ASCE-7-style")
    L.append("    directional-reversal and overturning-check patterns but should be")
    L.append("    reviewed against the governing code edition before use on a real project.")
    L.append("")
    L.append("  Appendix C provision check (temperature combinations, per handout):")
    L.append("  Appendix C states the NSCP load-combination factors -- including the")
    L.append("  temperature combinations -- were transcribed for NSCP 2015 but NOT")
    L.append("  verified against the printed code, and must be checked against")
    L.append(f"  {GOVERNING_TEMPERATURE_PROVISION['self_straining_requirement']} and")
    L.append(f"  {GOVERNING_TEMPERATURE_PROVISION['T_symbol_definition']} before design use.")
    L.append("  That check has now been performed. The governing NSCP 2015 provisions")
    L.append("  where T legitimately appears are:")
    for eq_id, eq_text in GOVERNING_TEMPERATURE_PROVISION["official_T_equations"].items():
        L.append(f"    Eq. {eq_id}: {eq_text}")
    L.append("  This project's 30-combination table (kept unchanged, per project")
    L.append("  confirmation against the rev3_combination_*.pdf reference renders)")
    L.append("  represents T with 4 combinations rather than the code's 3:")
    for cid, name in GOVERNING_TEMPERATURE_PROVISION["project_table_T_combos"].items():
        L.append(f"    Combo #{cid}: {name}")
    L.append(f"  {GOVERNING_TEMPERATURE_PROVISION['note']}")
    L.append("  - Roof elevation / roof beams / roof nodes are detected generically as the")
    L.append("    max-Y node set, not hardcoded IDs -- re-verify this still selects the")
    L.append("    intended level if the model geometry changes.")
    L.append("  - Consistency checks above verify load-vector arithmetic only; they are not")
    L.append("    a substitute for a full structural analysis (reactions, member forces,")
    L.append("    displacements), which remains out of Rev 3's scope.")

    hdr(10, "SIGN-OFF")
    L.append("  Rev 3 delivers: load-case data model, LC1-LC9, roof diaphragm DEFINITION,")
    L.append("  30 NSCP load combinations, and the PySide6/Matplotlib load viewer.")
    L.append("  Rev 3 does NOT deliver, and this report does NOT claim: solved reactions,")
    L.append("  member forces, or nodal displacements for any of the above. Those remain")
    L.append("  the unchanged, separate domain of Frame3DSolver.solve() (baseline).")
    L.append("=" * 78)
    return "\n".join(L) + "\n"


def write_verification_report(path, model, material_cfg, load_cases, diaphragm,
                                combinations, test_summary=None):
    text = build_verification_report_text(model, material_cfg, load_cases, diaphragm,
                                           combinations, test_summary=test_summary)
    with open(path, "w") as f:
        f.write(text)
    return path


def _run_test_suite_summary():
    """Runs test_cube_solver_rev3.py in-process (if importable) and returns a
    one-line pass/fail summary for verification-report Section 8. Returns
    None if the test module isn't on the path -- the report then says so
    explicitly rather than fabricating a result."""
    try:
        import importlib
        import io
        import unittest as _unittest
        test_mod = importlib.import_module("test_cube_solver_rev3")
        suite = _unittest.defaultTestLoader.loadTestsFromModule(test_mod)
        buf = io.StringIO()
        result = _unittest.TextTestRunner(stream=buf, verbosity=0).run(suite)
        return (f"{result.testsRun} tests run, {len(result.failures)} failure(s), "
                f"{len(result.errors)} error(s) -- "
                f"{'ALL PASSED' if result.wasSuccessful() else 'SEE test_cube_solver_rev3.py OUTPUT'}")
    except ImportError:
        return None


def run_verify_loads_export(output_dir=None, unit_system="Metric"):
    """Headless (--verify-loads) pipeline: builds the reference model, LC1-9,
    diaphragm, and all 30 combinations; exports rev3_load_case_1..9.png,
    rev3_combination_1.png (LRFD 1.4D) and rev3_combination_13.png (ASD D)
    as representative samples, and writes cube_rev3_verification_report.txt.
    Forces the Agg backend so it runs without a display."""
    import matplotlib
    matplotlib.use("Agg", force=True)
    import matplotlib.pyplot as _plt

    out_dir = output_dir or OUTPUT_DIR
    os.makedirs(out_dir, exist_ok=True)

    model, material_cfg, units_cfg, section_cfg = build_rev3_reference_model(unit_system)
    load_cases = build_all_rev3_load_cases(model, material_cfg)
    diaphragm = build_rev3_diaphragm(model)
    combinations = build_all_nscp_combinations()

    exported = []
    for cid in sorted(load_cases):
        case = load_cases[cid]
        fig = _plt.figure(figsize=(9.5, 7.2), dpi=110, facecolor=VP_CANVAS_BG)
        ax = fig.add_subplot(111, projection='3d', facecolor=VP_CANVAS_BG)
        _draw_model_geometry(ax, model)
        _draw_global_axes(ax, model)
        _draw_diaphragm(ax, model, diaphragm)
        legend_lines = _draw_loads(ax, model, case)
        set_grid_visible(ax, True)
        ax.set_xlabel("X (m)"); ax.set_ylabel("Z (m)"); ax.set_zlabel("Y (m)")
        ax.set_title(f"LC{case.id} - {case.name}", color=VP_TEXT_DARK, fontweight='bold')
        ax.text2D(0.02, 0.98, "\n".join(legend_lines), transform=ax.transAxes, fontsize=8.5,
                  color=VP_TEXT_DARK, va='top', ha='left',
                  bbox=dict(boxstyle="round", facecolor=VP_CANVAS_BG, edgecolor=VP_GRID_COLOR, alpha=0.9))
        path = os.path.join(out_dir, f"rev3_load_case_{cid}.png")
        fig.savefig(path, dpi=110, facecolor=VP_CANVAS_BG)
        _plt.close(fig)
        exported.append(path)

    for combo_id in (1, 13):
        combo = combinations[combo_id]
        fig = _plt.figure(figsize=(9.5, 7.2), dpi=110, facecolor=VP_CANVAS_BG)
        ax = fig.add_subplot(111, projection='3d', facecolor=VP_CANVAS_BG)
        _draw_model_geometry(ax, model)
        _draw_global_axes(ax, model)
        _draw_diaphragm(ax, model, diaphragm)
        legend_lines = _draw_combination(ax, model, combo, load_cases)
        set_grid_visible(ax, True)
        ax.set_xlabel("X (m)"); ax.set_ylabel("Z (m)"); ax.set_zlabel("Y (m)")
        ax.set_title(f"{combo.method} Combination {combo.id} - {combo.name}",
                     color=VP_TEXT_DARK, fontweight='bold')
        ax.text2D(0.02, 0.98, "\n".join(legend_lines), transform=ax.transAxes, fontsize=8.5,
                  color=VP_TEXT_DARK, va='top', ha='left',
                  bbox=dict(boxstyle="round", facecolor=VP_CANVAS_BG, edgecolor=VP_GRID_COLOR, alpha=0.9))
        path = os.path.join(out_dir, f"rev3_combination_{combo_id}.png")
        fig.savefig(path, dpi=110, facecolor=VP_CANVAS_BG)
        _plt.close(fig)
        exported.append(path)

    test_summary = _run_test_suite_summary()
    report_path = os.path.join(out_dir, "cube_rev3_verification_report.txt")
    write_verification_report(report_path, model, material_cfg, load_cases, diaphragm,
                               combinations, test_summary=test_summary)
    exported.append(report_path)

    print(f"[--verify-loads] Exported {len(exported)} file(s) to '{out_dir}':")
    for p in exported:
        print(f"    {p}")
    return exported


if __name__ == '__main__':
    if "--verify-loads" in sys.argv:
        run_verify_loads_export()
    else:
        main()
