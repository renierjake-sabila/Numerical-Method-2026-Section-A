"""
test_cube_solver_rev3.py -- Rev 3 automated test suite (standard unittest,
no pytest dependency). Targets ONLY the Rev 3 load-case / diaphragm /
combination / viewer-helper layer added on top of the untouched baseline
Frame3DSolver -- solve(), get_local_stiffness(), export_to_excel_earthy(),
and main() are never called or asserted against here.

Run with:  python -m unittest test_cube_solver_rev3 -v
"""
import math
import os
import tempfile
import unittest

import matplotlib
matplotlib.use("Agg")  # headless -- no display required
import matplotlib.pyplot as plt

import cube_solver_rev3 as rev3


def _build_fixture():
    """Shared fixture: reference model + LC1-9 + diaphragm + all 30 combos."""
    model, material_cfg, units_cfg, section_cfg = rev3.build_rev3_reference_model()
    load_cases = rev3.build_all_rev3_load_cases(model, material_cfg)
    diaphragm = rev3.build_rev3_diaphragm(model)
    combinations = rev3.build_all_nscp_combinations()
    return model, material_cfg, load_cases, diaphragm, combinations


# ==============================================================================
# 1. UNIT VALIDATION (UnitValidationError / direction parsing)
# ==============================================================================
class TestUnitValidation(unittest.TestCase):

    def test_nodal_load_rejects_wrong_force_unit(self):
        with self.assertRaises(rev3.UnitValidationError):
            rev3.NodalLoad(1, fx=1.0, force_unit="kN*m")

    def test_nodal_load_rejects_wrong_moment_unit(self):
        with self.assertRaises(rev3.UnitValidationError):
            rev3.NodalLoad(1, mx=1.0, moment_unit="kN")

    def test_nodal_load_accepts_correct_units(self):
        nl = rev3.NodalLoad(1, fx=1.0, mx=1.0)
        self.assertEqual(nl.fx, 1.0)

    def test_distributed_load_rejects_force_unit(self):
        with self.assertRaises(rev3.UnitValidationError):
            rev3.MemberDistributedLoad(1, "-Y", 5.0, unit="kN")

    def test_distributed_load_accepts_line_load_unit(self):
        dl = rev3.MemberDistributedLoad(1, "-Y", 5.0, unit="kN/m")
        self.assertEqual(dl.magnitude_kNm, 5.0)

    def test_distributed_load_rejects_non_uniform_distribution(self):
        with self.assertRaises(NotImplementedError):
            rev3.MemberDistributedLoad(1, "-Y", 5.0, distribution_type="trapezoidal")

    def test_point_load_rejects_line_load_unit(self):
        with self.assertRaises(rev3.UnitValidationError):
            rev3.MemberPointLoad(1, 0.5, "-Y", 5.0, unit="kN/m")

    def test_point_load_rejects_out_of_range_fraction(self):
        with self.assertRaises(ValueError):
            rev3.MemberPointLoad(1, 1.5, "-Y", 5.0)

    def test_point_load_accepts_boundary_fractions(self):
        rev3.MemberPointLoad(1, 0.0, "-Y", 5.0)
        rev3.MemberPointLoad(1, 1.0, "-Y", 5.0)

    def test_temperature_load_rejects_wrong_unit(self):
        with self.assertRaises(rev3.UnitValidationError):
            rev3.TemperatureLoad(1, 15.0, unit="F")

    def test_temperature_load_accepts_celsius(self):
        tl = rev3.TemperatureLoad(1, 15.0)
        self.assertEqual(tl.delta_T, 15.0)

    def test_parse_direction_valid(self):
        axis, sign = rev3._parse_direction("-Y")
        self.assertEqual((axis, sign), (1, -1.0))

    def test_parse_direction_invalid_letter(self):
        with self.assertRaises(ValueError):
            rev3._parse_direction("Q")

    def test_parse_direction_positive_x(self):
        axis, sign = rev3._parse_direction("X")
        self.assertEqual((axis, sign), (0, 1.0))


# ==============================================================================
# 2. LOAD CASE CONSTRUCTION
# ==============================================================================
class TestLoadCaseConstruction(unittest.TestCase):

    def test_rejects_invalid_category(self):
        with self.assertRaises(ValueError):
            rev3.LoadCase(1, "BAD", "Snow")

    def test_accepts_valid_category(self):
        case = rev3.LoadCase(1, "OK", "Dead")
        self.assertEqual(case.category, "Dead")

    def test_add_nodal_load_routes_correctly(self):
        case = rev3.LoadCase(1, "OK", "Dead")
        case.add(rev3.NodalLoad(1, fx=1.0))
        self.assertEqual(len(case.nodal_loads), 1)

    def test_add_distributed_load_routes_correctly(self):
        case = rev3.LoadCase(1, "OK", "Dead")
        case.add(rev3.MemberDistributedLoad(1, "-Y", 5.0))
        self.assertEqual(len(case.distributed_loads), 1)

    def test_add_point_load_routes_correctly(self):
        case = rev3.LoadCase(1, "OK", "Dead")
        case.add(rev3.MemberPointLoad(1, 0.5, "-Y", 5.0))
        self.assertEqual(len(case.point_loads), 1)

    def test_add_temperature_load_routes_correctly(self):
        case = rev3.LoadCase(1, "OK", "Temperature")
        case.add(rev3.TemperatureLoad(1, 15.0))
        self.assertEqual(len(case.temperature_loads), 1)

    def test_add_rejects_unknown_type(self):
        case = rev3.LoadCase(1, "OK", "Dead")
        with self.assertRaises(TypeError):
            case.add(object())


# ==============================================================================
# 3. GEOMETRY HELPERS
# ==============================================================================
class TestGeometryHelpers(unittest.TestCase):

    def setUp(self):
        self.model, self.material_cfg, _, _ = rev3.build_rev3_reference_model()

    def test_roof_elevation_is_6m(self):
        self.assertAlmostEqual(rev3.get_roof_elevation(self.model), 6.0)

    def test_roof_nodes_are_5_6_7_8(self):
        self.assertEqual(rev3.get_roof_nodes(self.model), [5, 6, 7, 8])

    def test_roof_beams_are_5_6_7_8(self):
        self.assertEqual(rev3.get_roof_beam_ids(self.model), [5, 6, 7, 8])

    def test_member_length_of_base_beam_is_6m(self):
        self.assertAlmostEqual(rev3.get_member_length(self.model, 1), 6.0)

    def test_member_length_of_column_is_6m(self):
        self.assertAlmostEqual(rev3.get_member_length(self.model, 9), 6.0)


# ==============================================================================
# 4. LOAD CASES LC1-LC9 (correctness against build_validation_summary)
# ==============================================================================
class TestLoadCases(unittest.TestCase):

    def setUp(self):
        (self.model, self.material_cfg, self.load_cases,
         self.diaphragm, self.combinations) = _build_fixture()

    def test_lc1_self_weight_is_negative_global_y(self):
        s = rev3.build_validation_summary(self.load_cases[1], self.model)
        self.assertLess(s["Fy_kN"], 0.0)
        self.assertAlmostEqual(s["Fx_kN"], 0.0)
        self.assertAlmostEqual(s["Fz_kN"], 0.0)

    def test_lc1_self_weight_touches_all_12_members(self):
        s = rev3.build_validation_summary(self.load_cases[1], self.model)
        self.assertEqual(len(s["loaded_members"]), 12)

    def test_lc1_self_weight_uses_density_kNm3_not_kgm3(self):
        gamma = self.material_cfg.density_kNm3
        self.assertLess(gamma, 1000.0)  # true weight density (~77), not mass density (~7850)

    def test_lc2_roof_dead_is_5_kNm_on_4_roof_beams(self):
        case = self.load_cases[2]
        self.assertEqual(len(case.distributed_loads), 4)
        for dl in case.distributed_loads:
            self.assertAlmostEqual(dl.magnitude_kNm, 5.0)
            self.assertEqual(dl.direction, "-Y")

    def test_lc2_roof_dead_total_is_120_kN(self):
        s = rev3.build_validation_summary(self.load_cases[2], self.model, intended_total_kN=120.0)
        self.assertAlmostEqual(s["equilibrium_error_kN"], 0.0)

    def test_lc3_roof_live_is_3_kNm_on_4_roof_beams(self):
        case = self.load_cases[3]
        self.assertEqual(len(case.distributed_loads), 4)
        for dl in case.distributed_loads:
            self.assertAlmostEqual(dl.magnitude_kNm, 3.0)

    def test_lc3_roof_live_total_is_72_kN(self):
        s = rev3.build_validation_summary(self.load_cases[3], self.model, intended_total_kN=72.0)
        self.assertAlmostEqual(s["equilibrium_error_kN"], 0.0)

    def test_lc4_center_point_is_at_midpoint(self):
        for pl in self.load_cases[4].point_loads:
            self.assertAlmostEqual(pl.location_fraction, 0.5)

    def test_lc4_center_point_is_5kN_downward(self):
        for pl in self.load_cases[4].point_loads:
            self.assertAlmostEqual(pl.magnitude_kN, 5.0)
            self.assertEqual(pl.direction, "-Y")

    def test_lc4_total_is_20_kN(self):
        s = rev3.build_validation_summary(self.load_cases[4], self.model, intended_total_kN=20.0)
        self.assertAlmostEqual(s["equilibrium_error_kN"], 0.0)

    def test_lc5_wind_x_is_4x2p5_equals_10(self):
        s = rev3.build_validation_summary(self.load_cases[5], self.model, intended_total_kN=10.0)
        self.assertEqual(len(s["loaded_nodes"]), 4)
        self.assertAlmostEqual(s["equilibrium_error_kN"], 0.0)
        self.assertAlmostEqual(s["Fx_kN"], 10.0)

    def test_lc6_wind_z_is_4x2p5_equals_10(self):
        s = rev3.build_validation_summary(self.load_cases[6], self.model, intended_total_kN=10.0)
        self.assertAlmostEqual(s["Fz_kN"], 10.0)
        self.assertAlmostEqual(s["equilibrium_error_kN"], 0.0)

    def test_lc7_seismic_x_is_4x3p75_equals_15(self):
        s = rev3.build_validation_summary(self.load_cases[7], self.model, intended_total_kN=15.0)
        self.assertAlmostEqual(s["Fx_kN"], 15.0)
        self.assertAlmostEqual(s["equilibrium_error_kN"], 0.0)

    def test_lc8_seismic_z_is_4x3p75_equals_15(self):
        s = rev3.build_validation_summary(self.load_cases[8], self.model, intended_total_kN=15.0)
        self.assertAlmostEqual(s["Fz_kN"], 15.0)
        self.assertAlmostEqual(s["equilibrium_error_kN"], 0.0)

    def test_lc5_through_lc8_use_the_same_4_roof_nodes(self):
        for cid in (5, 6, 7, 8):
            nodes = {nl.node_id for nl in self.load_cases[cid].nodal_loads}
            self.assertEqual(nodes, {5, 6, 7, 8})

    def test_lc9_temperature_net_force_is_zero(self):
        s = rev3.build_validation_summary(self.load_cases[9], self.model, intended_total_kN=0.0)
        self.assertAlmostEqual(s["computed_total_kN"], 0.0)
        self.assertAlmostEqual(s["equilibrium_error_kN"], 0.0)

    def test_lc9_temperature_touches_all_12_members(self):
        self.assertEqual(len(self.load_cases[9].temperature_loads), 12)

    def test_lc9_thermal_strain_formula(self):
        tl = self.load_cases[9].temperature_loads[0]
        self.assertAlmostEqual(tl.thermal_strain(1.2e-5), 1.2e-5 * 15.0)

    def test_all_nine_load_cases_present(self):
        self.assertEqual(set(self.load_cases.keys()), set(range(1, 10)))


# ==============================================================================
# 5. DIAPHRAGM
# ==============================================================================
class TestDiaphragm(unittest.TestCase):

    def setUp(self):
        self.model, _, _, _ = rev3.build_rev3_reference_model()

    def test_master_is_node_5(self):
        dp = rev3.build_rev3_diaphragm(self.model)
        self.assertEqual(dp.master_node, 5)

    def test_slaves_are_6_7_8(self):
        dp = rev3.build_rev3_diaphragm(self.model)
        self.assertEqual(dp.slave_nodes, [6, 7, 8])

    def test_coupled_dofs_are_ux_uz_ry(self):
        self.assertEqual(rev3.Diaphragm.COUPLED_DOFS, ("UX", "UZ", "RY"))

    def test_free_dofs_are_uy_rx_rz(self):
        self.assertEqual(rev3.Diaphragm.FREE_DOFS, ("UY", "RX", "RZ"))

    def test_master_in_slaves_raises(self):
        with self.assertRaises(ValueError):
            rev3.Diaphragm(1, "BAD", master_node=5, slave_nodes=[5, 6, 7])

    def test_constraint_equation_count_is_9(self):
        dp = rev3.build_rev3_diaphragm(self.model)
        self.assertEqual(len(dp.constraint_equations()), 9)  # 3 slaves x 3 coupled DOFs

    def test_constraint_equations_reference_master(self):
        dp = rev3.build_rev3_diaphragm(self.model)
        for eq in dp.constraint_equations():
            self.assertIn("node5", eq)

    def test_diaphragm_repr_contains_master_and_slaves(self):
        dp = rev3.build_rev3_diaphragm(self.model)
        r = repr(dp)
        self.assertIn("N5", r)
        self.assertIn("[6, 7, 8]", r)


# ==============================================================================
# 6. NSCP LOAD COMBINATIONS -- all 30, factor expansion + consistency
# ==============================================================================
class TestNSCPCombinations(unittest.TestCase):

    def setUp(self):
        (self.model, self.material_cfg, self.load_cases,
         self.diaphragm, self.combinations) = _build_fixture()

    def test_exactly_30_combinations_built(self):
        self.assertEqual(len(self.combinations), 30)

    def test_combinations_1_to_12_are_lrfd(self):
        for cid in range(1, 13):
            self.assertEqual(self.combinations[cid].method, "LRFD")

    def test_combinations_13_to_30_are_asd(self):
        for cid in range(13, 31):
            self.assertEqual(self.combinations[cid].method, "ASD")

    def test_combo_1_is_1p4D(self):
        self.assertEqual(self.combinations[1].factors, {1: 1.4, 2: 1.4, 4: 1.4})

    def test_combo_13_is_canonical_D(self):
        self.assertEqual(self.combinations[13].factors, {1: 1.0, 2: 1.0, 4: 1.0})

    def _assert_combo(self, combo_id, expected_fy=None, expected_fx=None, expected_fz=None):
        combo = self.combinations[combo_id]
        result = rev3.verify_combination_consistency(combo, self.load_cases, self.model)
        self.assertAlmostEqual(result["consistency_error_kN"], 0.0, places=6)
        if expected_fy is not None:
            self.assertAlmostEqual(result["Fy_kN"], expected_fy, places=3)
        if expected_fx is not None:
            self.assertAlmostEqual(result["Fx_kN"], expected_fx, places=3)
        if expected_fz is not None:
            self.assertAlmostEqual(result["Fz_kN"], expected_fz, places=3)
        return result

    D = None  # placeholder, replaced in setUp via self._D

    def setUp_D(self):
        s1 = rev3.build_validation_summary(self.load_cases[1], self.model)
        s2 = rev3.build_validation_summary(self.load_cases[2], self.model)
        s4 = rev3.build_validation_summary(self.load_cases[4], self.model)
        return s1["Fy_kN"] + s2["Fy_kN"] + s4["Fy_kN"]  # negative kN

    def test_combo_01(self):
        D = self.setUp_D()
        self._assert_combo(1, expected_fy=1.4 * D)

    def test_combo_02(self):
        D = self.setUp_D()
        L = rev3.build_validation_summary(self.load_cases[3], self.model)["Fy_kN"]
        self._assert_combo(2, expected_fy=1.2 * D + 1.6 * L)

    def test_combo_03(self):
        self._assert_combo(3, expected_fx=10.0)

    def test_combo_04(self):
        self._assert_combo(4, expected_fx=-10.0)

    def test_combo_05(self):
        self._assert_combo(5, expected_fz=10.0)

    def test_combo_06(self):
        self._assert_combo(6, expected_fz=-10.0)

    def test_combo_07(self):
        self._assert_combo(7, expected_fx=15.0)

    def test_combo_08(self):
        self._assert_combo(8, expected_fx=-15.0)

    def test_combo_09(self):
        self._assert_combo(9, expected_fz=15.0)

    def test_combo_10(self):
        self._assert_combo(10, expected_fz=-15.0)

    def test_combo_11(self):
        D = self.setUp_D()
        L = rev3.build_validation_summary(self.load_cases[3], self.model)["Fy_kN"]
        self._assert_combo(11, expected_fy=1.2 * D + 1.0 * L)

    def test_combo_12(self):
        D = self.setUp_D()
        self._assert_combo(12, expected_fy=0.9 * D)

    def test_combo_13(self):
        D = self.setUp_D()
        self._assert_combo(13, expected_fy=D)

    def test_combo_14(self):
        D = self.setUp_D()
        L = rev3.build_validation_summary(self.load_cases[3], self.model)["Fy_kN"]
        self._assert_combo(14, expected_fy=D + L)

    def test_combo_15(self):
        self._assert_combo(15, expected_fx=6.0)

    def test_combo_16(self):
        self._assert_combo(16, expected_fx=-6.0)

    def test_combo_17(self):
        self._assert_combo(17, expected_fz=6.0)

    def test_combo_18(self):
        self._assert_combo(18, expected_fz=-6.0)

    def test_combo_19(self):
        self._assert_combo(19, expected_fx=10.5)

    def test_combo_20(self):
        self._assert_combo(20, expected_fx=-10.5)

    def test_combo_21(self):
        self._assert_combo(21, expected_fz=10.5)

    def test_combo_22(self):
        self._assert_combo(22, expected_fz=-10.5)

    def test_combo_23(self):
        D = self.setUp_D()
        self._assert_combo(23, expected_fy=0.6 * D, expected_fx=6.0)

    def test_combo_24(self):
        D = self.setUp_D()
        self._assert_combo(24, expected_fy=0.6 * D, expected_fx=-6.0)

    def test_combo_25(self):
        D = self.setUp_D()
        self._assert_combo(25, expected_fy=0.6 * D, expected_fz=6.0)

    def test_combo_26(self):
        D = self.setUp_D()
        self._assert_combo(26, expected_fy=0.6 * D, expected_fz=-6.0)

    def test_combo_27(self):
        D = self.setUp_D()
        self._assert_combo(27, expected_fy=0.6 * D, expected_fx=10.5)

    def test_combo_28(self):
        D = self.setUp_D()
        self._assert_combo(28, expected_fy=0.6 * D, expected_fx=-10.5)

    def test_combo_29(self):
        D = self.setUp_D()
        self._assert_combo(29, expected_fy=D)  # T contributes 0.0 kN

    def test_combo_30(self):
        D = self.setUp_D()
        L = rev3.build_validation_summary(self.load_cases[3], self.model)["Fy_kN"]
        self._assert_combo(30, expected_fy=D + 0.75 * L)  # 0.75T contributes 0.0 kN

    def test_all_30_combinations_have_zero_consistency_error(self):
        for cid in range(1, 31):
            result = rev3.verify_combination_consistency(
                self.combinations[cid], self.load_cases, self.model)
            self.assertAlmostEqual(result["consistency_error_kN"], 0.0, places=6,
                                    msg=f"combo {cid} failed consistency check")

    def test_expand_symbolic_factors_rejects_unknown_symbol(self):
        with self.assertRaises(ValueError):
            rev3._expand_symbolic_factors({"S": 1.0})

    def test_load_combination_rejects_bad_method(self):
        with self.assertRaises(ValueError):
            rev3.LoadCombination(99, "BAD", "WSD", {1: 1.0})


# ==============================================================================
# 7. VIEWER-LAYER RENDERING HELPERS (matplotlib only -- no PySide6 required)
# ==============================================================================
class TestViewerRenderingHelpers(unittest.TestCase):

    def setUp(self):
        (self.model, self.material_cfg, self.load_cases,
         self.diaphragm, self.combinations) = _build_fixture()
        self.fig = plt.figure()
        self.ax = self.fig.add_subplot(111, projection='3d')

    def tearDown(self):
        plt.close(self.fig)

    def test_view_layers_has_seven_entries(self):
        # 6 original layers + Step 5's "load_band" toggle.
        self.assertEqual(len(rev3.VIEW_LAYERS), 7)

    def test_view_layers_keys_are_unique(self):
        keys = [k for k, _, _ in rev3.VIEW_LAYERS]
        self.assertEqual(len(keys), len(set(keys)))

    def test_grid_is_first_view_layer(self):
        self.assertEqual(rev3.VIEW_LAYERS[0][0], "grid")

    def test_set_grid_visible_off_leaves_no_kwargs_trap(self):
        rev3.set_grid_visible(self.ax, False)
        self.assertFalse(self.ax._draw_grid)

    def test_set_grid_visible_on_sets_draw_grid_true(self):
        rev3.set_grid_visible(self.ax, True)
        self.assertTrue(self.ax._draw_grid)

    def test_draw_loads_returns_legend_for_lc1(self):
        legend = rev3._draw_loads(self.ax, self.model, self.load_cases[1])
        self.assertTrue(any("DEAD" in line for line in legend))

    def test_draw_loads_self_weight_skips_polygon_on_columns(self):
        # LC1 (self-weight, -Y) on 4 columns + 8 beams: only beams (8 of them:
        # 4 base + 4 roof) get a Poly3DCollection band; the 4 vertical columns
        # must NOT, per the degenerate-axial-load guard.
        n_before = len(self.ax.collections)
        rev3._draw_loads(self.ax, self.model, self.load_cases[1])
        n_polys = len(self.ax.collections) - n_before
        self.assertEqual(n_polys, 8)  # 8 beams get bands, 4 columns are skipped

    def test_draw_loads_roof_dead_bands_all_4_roof_beams(self):
        n_before = len(self.ax.collections)
        rev3._draw_loads(self.ax, self.model, self.load_cases[2])
        self.assertEqual(len(self.ax.collections) - n_before, 4)

    def test_draw_combination_returns_header_and_per_case_lines(self):
        legend = rev3._draw_combination(self.ax, self.model, self.combinations[1], self.load_cases)
        self.assertEqual(legend[0], "Combination 1 [LRFD] - 1.4D")
        self.assertEqual(len(legend), 1 + len(self.combinations[1].factors))

    def test_scaled_load_view_scales_nodal_load(self):
        view = rev3._ScaledLoadView(self.load_cases[5], 0.6)
        self.assertAlmostEqual(view.nodal_loads[0].fx, self.load_cases[5].nodal_loads[0].fx * 0.6)

    def test_scaled_load_view_scales_distributed_load(self):
        view = rev3._ScaledLoadView(self.load_cases[2], 1.4)
        self.assertAlmostEqual(view.distributed_loads[0].magnitude_kNm, 7.0)

    def test_scaled_load_view_preserves_temperature_loads(self):
        view = rev3._ScaledLoadView(self.load_cases[9], 0.75)
        self.assertEqual(len(view.temperature_loads), len(self.load_cases[9].temperature_loads))

    def test_scaled_load_view_carries_self_weight_factor(self):
        view = rev3._ScaledLoadView(self.load_cases[1], 1.4)
        self.assertEqual(view.self_weight_factor, self.load_cases[1].self_weight_factor)

    # -- Step 5: Load Band Toggle --------------------------------------------
    def test_load_band_toggle_preserves_glyph_count(self):
        # LC2 (roof dead) bands 4 beams. The tick-arrow row (ax.lines) must
        # be identical whether the band fill is shown or not; only the
        # Poly3DCollection fill (ax.collections) should differ.
        lines_before = len(self.ax.lines)
        collections_before = len(self.ax.collections)
        rev3._draw_loads(self.ax, self.model, self.load_cases[2], show_band=True)
        lines_with_band = len(self.ax.lines) - lines_before
        collections_with_band = len(self.ax.collections) - collections_before
        self.assertEqual(collections_with_band, 4)  # 4 beams get a polygon band

        self.fig2 = plt.figure()
        ax2 = self.fig2.add_subplot(111, projection='3d')
        try:
            lines_before2 = len(ax2.lines)
            collections_before2 = len(ax2.collections)
            rev3._draw_loads(ax2, self.model, self.load_cases[2], show_band=False)
            lines_without_band = len(ax2.lines) - lines_before2
            collections_without_band = len(ax2.collections) - collections_before2
            self.assertEqual(collections_without_band, 0)  # no polygon fill
            self.assertEqual(lines_with_band, lines_without_band)  # same arrow/tick count
        finally:
            plt.close(self.fig2)

    def test_build_combination_glyphs_returns_one_view_per_factor(self):
        glyphs = rev3.build_combination_glyphs(self.combinations[1], self.load_cases)
        self.assertEqual(len(glyphs), len(self.combinations[1].factors))
        self.assertTrue(all(isinstance(g, rev3._ScaledLoadView) for g in glyphs))

    def test_render_combination_figure_matches_draw_combination_alias(self):
        legend_a = rev3.render_combination_figure(self.ax, self.model, self.combinations[1],
                                                    self.load_cases)
        self.fig3 = plt.figure()
        ax3 = self.fig3.add_subplot(111, projection='3d')
        try:
            legend_b = rev3._draw_combination(ax3, self.model, self.combinations[1],
                                               self.load_cases)
            self.assertEqual(legend_a, legend_b)
        finally:
            plt.close(self.fig3)

    # -- Step 6: Lighter Band Fill Opacity -----------------------------------
    def test_band_alpha_constants_exist(self):
        self.assertAlmostEqual(rev3.BAND_ALPHA_DISTRIBUTED, 0.30)
        self.assertAlmostEqual(rev3.BAND_ALPHA_SELF_WEIGHT, 0.18)

    def test_self_weight_band_uses_lighter_fill_alpha(self):
        rev3._draw_loads(self.ax, self.model, self.load_cases[1])  # LC1
        polys = [c for c in self.ax.collections
                 if isinstance(c, rev3.Poly3DCollection)]
        self.assertTrue(polys)
        for poly in polys:
            fc = poly.get_facecolor()[0]
            self.assertAlmostEqual(fc[3], rev3.BAND_ALPHA_SELF_WEIGHT, places=3)
            ec = poly.get_edgecolor()[0]
            self.assertAlmostEqual(ec[3], 0.9, places=3)

    def test_distributed_load_band_uses_standard_fill_alpha(self):
        rev3._draw_loads(self.ax, self.model, self.load_cases[2])  # LC2 roof dead
        polys = [c for c in self.ax.collections
                 if isinstance(c, rev3.Poly3DCollection)]
        self.assertTrue(polys)
        for poly in polys:
            fc = poly.get_facecolor()[0]
            self.assertAlmostEqual(fc[3], rev3.BAND_ALPHA_DISTRIBUTED, places=3)
            ec = poly.get_edgecolor()[0]
            self.assertAlmostEqual(ec[3], 0.9, places=3)


# ==============================================================================
# 8. VERIFICATION REPORT
# ==============================================================================
class TestVerificationReport(unittest.TestCase):

    def setUp(self):
        (self.model, self.material_cfg, self.load_cases,
         self.diaphragm, self.combinations) = _build_fixture()

    def test_report_contains_all_10_section_headers(self):
        text = rev3.build_verification_report_text(
            self.model, self.material_cfg, self.load_cases, self.diaphragm, self.combinations)
        for n, title in [
            (1, "SCOPE & BOUNDARY"), (2, "MODEL SUMMARY"), (3, "LOAD CASE RESULTS"),
            (4, "RIGID ROOF DIAPHRAGM"), (5, "NSCP LOAD COMBINATIONS"),
            (6, "PER-COMBINATION CONSISTENCY"), (7, "CODE BASIS"),
            (8, "AUTOMATED TEST SUMMARY"), (9, "ASSUMPTIONS"), (10, "SIGN-OFF"),
        ]:
            self.assertIn(f"{n}. {title}", text)

    def test_report_states_no_solver_boundary(self):
        text = rev3.build_verification_report_text(
            self.model, self.material_cfg, self.load_cases, self.diaphragm, self.combinations)
        self.assertIn("does NOT compute", text)

    def test_report_lists_all_30_combinations(self):
        text = rev3.build_verification_report_text(
            self.model, self.material_cfg, self.load_cases, self.diaphragm, self.combinations)
        for cid in range(1, 31):
            self.assertIn(f"#{cid}", text.replace(f"#{cid} ", f"#{cid}"))
        # every combo name appears at least once
        for cid, name, _, _ in rev3._NSCP_TABLE:
            self.assertIn(name, text)

    def test_report_cites_nscp_2015(self):
        text = rev3.build_verification_report_text(
            self.model, self.material_cfg, self.load_cases, self.diaphragm, self.combinations)
        self.assertIn("NSCP 2015", text)

    def test_write_verification_report_creates_file(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "report.txt")
            rev3.write_verification_report(path, self.model, self.material_cfg,
                                            self.load_cases, self.diaphragm, self.combinations)
            self.assertTrue(os.path.exists(path))
            self.assertGreater(os.path.getsize(path), 0)

    def test_report_contains_temperature_verification_section(self):
        text = rev3.build_verification_report_text(
            self.model, self.material_cfg, self.load_cases, self.diaphragm, self.combinations)
        self.assertIn("3B. TEMPERATURE LOAD VERIFICATION (LC9 DETAIL)", text)
        self.assertIn("restraint_state=fully_restrained", text)
        self.assertIn("restraint_state=free", text)
        self.assertIn("restraint_state=partially_restrained", text)
        self.assertIn("indeterminate", text)


# ==============================================================================
# 9. STEP 7 -- TEMPERATURE LOAD (LC9): 13 required tests + supporting helpers
# ==============================================================================
class TestTemperatureLoadStep7(unittest.TestCase):

    def setUp(self):
        (self.model, self.material_cfg, self.load_cases,
         self.diaphragm, self.combinations) = _build_fixture()
        self.lc9 = self.load_cases[9]

    # 1. Temperature load case exists.
    def test_01_temperature_load_case_exists(self):
        self.assertIn(9, self.load_cases)
        self.assertEqual(self.load_cases[9].category, "Temperature")

    # 2. Temperature change equals +15 degC.
    def test_02_temperature_change_is_plus_15C(self):
        for tl in self.lc9.temperature_loads:
            self.assertAlmostEqual(tl.delta_T, 15.0)

    # 3. Material coefficient of thermal expansion is correctly retrieved.
    def test_03_alpha_retrieved_from_material_config(self):
        self.assertAlmostEqual(self.material_cfg.therm_coeff_perC, 11.7e-6)
        for tl in self.lc9.temperature_loads:
            self.assertAlmostEqual(tl.alpha_perC, self.material_cfg.therm_coeff_perC)

    # 4. Thermal strain calculation is correct.
    def test_04_thermal_strain_calculation(self):
        tl = self.lc9.temperature_loads[0]
        expected = self.material_cfg.therm_coeff_perC * 15.0
        self.assertAlmostEqual(tl.thermal_strain(), expected)
        self.assertAlmostEqual(expected, 1.7550e-04, places=8)

    # 5. Free thermal expansion calculation is correct.
    def test_05_free_expansion_calculation(self):
        tl = self.lc9.temperature_loads[0]  # member 1, L = 6 m
        L = rev3.get_member_length(self.model, 1)
        expected_dL = self.material_cfg.therm_coeff_perC * 15.0 * L
        self.assertAlmostEqual(tl.free_expansion_m(L), expected_dL)
        self.assertAlmostEqual(expected_dL * 1000, 1.0530, places=3)

    # 6. Restrained-member thermal force calculation is correct.
    def test_06_restrained_member_thermal_force(self):
        tl = self.lc9.temperature_loads[0]  # member 1: base beam, fully restrained
        m = self.model.members[1]
        EA_kN = (m['E'] * m['A']) / 1000.0
        expected_N = EA_kN * self.material_cfg.therm_coeff_perC * 15.0
        self.assertAlmostEqual(tl.restrained_force_kN(m['E'], m['A']), expected_N)
        rows = rev3.build_temperature_verification(self.model, self.material_cfg, self.lc9)
        row1 = next(r for r in rows if r["member_id"] == 1)
        self.assertEqual(row1["restraint_state"], "fully_restrained")
        self.assertAlmostEqual(row1["reported_force_kN"], expected_N)
        self.assertEqual(row1["reported_expansion_m"], 0.0)

    # 7. Partially restrained behavior is handled consistently (reported as
    # indeterminate, never a guessed force), matching the solver formulation
    # (no active stiffness solve exists in Rev 3 to actually determine it).
    def test_07_partially_restrained_reported_as_indeterminate(self):
        rows = rev3.build_temperature_verification(self.model, self.material_cfg, self.lc9)
        row9 = next(r for r in rows if r["member_id"] == 9)  # a column
        self.assertEqual(row9["restraint_state"], "partially_restrained")
        self.assertIsNone(row9["reported_force_kN"])
        self.assertIsNone(row9["reported_expansion_m"])
        self.assertIn("indeterminate", row9["force_note"])
        # the analytical fixed-end benchmark is still computed for reference
        self.assertGreater(row9["fully_restrained_force_kN"], 0.0)

    # 8. Temperature loads use correct units.
    def test_08_temperature_load_rejects_wrong_unit(self):
        with self.assertRaises(rev3.UnitValidationError):
            rev3.TemperatureLoad(1, 15.0, unit="F")
        tl = rev3.TemperatureLoad(1, 15.0, unit="C")
        self.assertEqual(tl.delta_T, 15.0)

    # 9. Temperature is not accidentally interpreted as a mechanical force.
    def test_09_temperature_case_contributes_zero_force(self):
        summ = rev3.build_validation_summary(self.lc9, self.model, intended_total_kN=0.0)
        self.assertAlmostEqual(summ["Fx_kN"], 0.0)
        self.assertAlmostEqual(summ["Fy_kN"], 0.0)
        self.assertAlmostEqual(summ["Fz_kN"], 0.0)
        self.assertAlmostEqual(summ["equilibrium_error_kN"], 0.0)
        # ...while the FE thermal load vector is explicitly NON-empty per
        # member -- both halves of the checkpoint, not just the net-zero half.
        net_check = rev3.verify_temperature_fe_vector_net_force(
            self.model, self.material_cfg, self.lc9)
        self.assertEqual(net_check["nonzero_member_count"], 12)
        self.assertEqual(net_check["total_members"], 12)
        self.assertAlmostEqual(net_check["net_force_magnitude_kN"], 0.0, places=6)

    # 10. Load Case Viewer correctly displays +15 degC.
    def test_10_viewer_displays_plus_15_degC_not_kN(self):
        fig = plt.figure()
        ax = fig.add_subplot(111, projection='3d')
        try:
            legend = rev3._draw_loads(ax, self.model, self.lc9)
            self.assertTrue(any("+15.0\u00b0C" in line or "+15.0 C" in line for line in legend))
            # The on-plot member annotations (the actual "Load Case Viewer"
            # display, not the legend's clarifying "0.000 kN net" footnote)
            # must show the temperature value in degC ONLY -- never mislabel
            # +15 as a force/line-load unit.
            texts = [t.get_text() for t in ax.texts]
            thermal_texts = [t for t in texts if "\u00b0C" in t]
            self.assertTrue(thermal_texts)
            self.assertTrue(all("kN" not in t for t in thermal_texts))
            self.assertEqual(len(thermal_texts), len(self.lc9.temperature_loads))
        finally:
            plt.close(fig)

    # 11. Load Combination Viewer correctly includes temperature when
    # applicable (Combos 11, 12, 29, 30 -- never in a combo without T).
    def test_11_combinations_include_temperature_only_where_applicable(self):
        thermal_combo_ids = {cid for cid, c in self.combinations.items() if 9 in c.factors}
        self.assertEqual(thermal_combo_ids, {11, 12, 29, 30})
        # factor scaling: Combo 30 uses 0.75T, others use 1.0T
        self.assertAlmostEqual(self.combinations[30].factors[9], 0.75)
        self.assertAlmostEqual(self.combinations[11].factors[9], 1.0)
        self.assertAlmostEqual(self.combinations[12].factors[9], 1.0)
        self.assertAlmostEqual(self.combinations[29].factors[9], 1.0)
        # a non-thermal combo (e.g. #2, 1.2D+1.6L) must NOT carry T
        self.assertNotIn(9, self.combinations[2].factors)

    # 12. Viewer and solver use the same temperature-load definition (the
    # _ScaledLoadView used for combination rendering scales delta_T by the
    # same factor the combination-assembly/report path uses).
    def test_12_scaled_view_and_report_use_same_temperature_definition(self):
        view = rev3._ScaledLoadView(self.lc9, 0.75)
        self.assertAlmostEqual(view.temperature_loads[0].delta_T, 15.0 * 0.75)
        self.assertAlmostEqual(view.temperature_loads[0].delta_T, 11.25)
        self.assertEqual(view.temperature_loads[0].alpha_perC,
                          self.lc9.temperature_loads[0].alpha_perC)
        # Combo 30's factor (0.75) matches what render_combination_figure()
        # actually draws -- same authoritative source (combo.factors), no
        # separate hardcoded value anywhere else.
        self.assertAlmostEqual(self.combinations[30].factors[9], 0.75)

    # 13. Temperature results pass numerical verification against an
    # analytical benchmark (hand-derived, independent of the code path).
    def test_13_analytical_benchmark_cross_check(self):
        alpha = 11.7e-6
        dT = 15.0
        eps_expected = alpha * dT
        self.assertAlmostEqual(eps_expected, 1.7550e-04, places=8)

        L = 6.0
        dL_expected_mm = eps_expected * L * 1000
        self.assertAlmostEqual(dL_expected_mm, 1.0530, places=3)

        m1 = self.model.members[1]
        EA_kN_expected = (m1['E'] * m1['A']) / 1000.0
        N_expected = EA_kN_expected * eps_expected

        rows = rev3.build_temperature_verification(self.model, self.material_cfg, self.lc9)
        row1 = next(r for r in rows if r["member_id"] == 1)
        self.assertAlmostEqual(row1["eps_T"], eps_expected, places=8)
        self.assertAlmostEqual(row1["free_expansion_dL_m"] * 1000, dL_expected_mm, places=3)
        self.assertAlmostEqual(row1["fully_restrained_force_kN"], N_expected, places=3)
        # order-of-magnitude cross-check against the handout's approximate
        # benchmark (173.349 kN) -- real workbook-derived E/A governs, but
        # must agree within a small tolerance, not just "close enough to eyeball"
        self.assertAlmostEqual(row1["fully_restrained_force_kN"], 173.349, delta=1.0)

    # -- Supporting coverage for the new helpers themselves --------------
    def test_member_restraint_state_classifies_all_12_members_correctly(self):
        expected = {1: "fully_restrained", 2: "fully_restrained",
                    3: "fully_restrained", 4: "fully_restrained",
                    5: "free", 6: "free", 7: "free", 8: "free",
                    9: "partially_restrained", 10: "partially_restrained",
                    11: "partially_restrained", 12: "partially_restrained"}
        for mid, state in expected.items():
            self.assertEqual(rev3.member_restraint_state(self.model, mid), state,
                              f"member {mid} restraint mismatch")

    def test_fe_thermal_load_vector_is_self_equilibrating(self):
        tl = self.lc9.temperature_loads[0]
        m = self.model.members[1]
        F_i, F_j = tl.fe_thermal_load_vector_kN(m['E'], m['A'])
        self.assertAlmostEqual(F_i, -F_j)
        self.assertNotAlmostEqual(F_i, 0.0)
        self.assertNotAlmostEqual(F_j, 0.0)

    def test_temperature_load_backward_compatible_constructor(self):
        # The exact pre-Step-7 call signature must still work unchanged.
        tl = rev3.TemperatureLoad(1, 15.0)
        self.assertIsNone(tl.alpha_perC)
        self.assertAlmostEqual(tl.thermal_strain(1.2e-5), 1.2e-5 * 15.0)
        with self.assertRaises(ValueError):
            tl.thermal_strain()  # no alpha available anywhere -> explicit error


if __name__ == "__main__":
    unittest.main(verbosity=2)
