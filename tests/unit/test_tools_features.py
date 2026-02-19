"""Dispatch tests for tools/features.py composite tools."""

from unittest.mock import MagicMock

import pytest

from solidedge_mcp.tools.features import (
    add_body,
    create_bend,
    create_blend,
    create_bounded_surface,
    create_chamfer,
    create_contour_flange,
    create_dimple,
    create_draft_angle,
    create_drawn_cutout,
    create_extrude,
    create_extruded_cutout,
    create_extruded_surface,
    create_flange,
    create_helix,
    create_helix_cutout,
    create_hole,
    create_loft,
    create_lofted_cutout,
    create_lofted_flange,
    create_lofted_surface,
    create_louver,
    create_mirror,
    create_normal_cutout,
    create_pattern,
    create_primitive,
    create_primitive_cutout,
    create_ref_plane,
    create_reinforcement,
    create_revolve,
    create_revolved_cutout,
    create_revolved_surface,
    create_round,
    create_sheet_metal_base,
    create_slot,
    create_split,
    create_stamped,
    create_surface_mark,
    create_sweep,
    create_swept_cutout,
    create_swept_surface,
    create_thin_wall,
    create_thread,
    create_web_network,
    delete_topology,
    face_operation,
    manage_feature,
    sheet_metal_misc,
    simplify,
    thicken,
)


_FEATURE_SUBMODULES = [
    "solidedge_mcp.tools.features._extrude",
    "solidedge_mcp.tools.features._revolve",
    "solidedge_mcp.tools.features._cutout",
    "solidedge_mcp.tools.features._loft_sweep",
    "solidedge_mcp.tools.features._surfaces",
    "solidedge_mcp.tools.features._primitives",
    "solidedge_mcp.tools.features._ref_planes",
    "solidedge_mcp.tools.features._rounds_chamfers",
    "solidedge_mcp.tools.features._holes",
    "solidedge_mcp.tools.features._sheet_metal",
    "solidedge_mcp.tools.features._misc",
]


@pytest.fixture
def mock_mgr(monkeypatch):
    mgr = MagicMock()
    for mod in _FEATURE_SUBMODULES:
        monkeypatch.setattr(f"{mod}.feature_manager", mgr)
    return mgr


# === create_extrude ===

class TestCreateExtrude:
    @pytest.mark.parametrize("disc, method", [
        ("finite", "create_extrude"),
        ("infinite", "create_extrude_infinite"),
        ("through_next", "create_extrude_through_next"),
        ("from_to", "create_extrude_from_to"),
        ("thin_wall", "create_extrude_thin_wall"),
        ("symmetric", "create_extrude_symmetric"),
        ("through_next_v2", "create_extrude_through_next_v2"),
        ("from_to_v2", "create_extrude_from_to_v2"),
        ("by_keypoint", "create_extrude_by_keypoint"),
        ("from_to_single", "create_extrude_from_to_single"),
        ("through_next_single", "create_extrude_through_next_single"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_extrude(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_finite_passes_args(self, mock_mgr):
        mock_mgr.create_extrude.return_value = {"status": "ok"}
        create_extrude(method="finite", distance=0.05, direction="Reverse")
        mock_mgr.create_extrude.assert_called_once_with(0.05, "Reverse")

    def test_unknown(self, mock_mgr):
        result = create_extrude(method="bogus")
        assert "error" in result


# === create_revolve ===

class TestCreateRevolve:
    @pytest.mark.parametrize("disc, method", [
        ("full", "create_revolve"),
        ("finite", "create_revolve_finite"),
        ("sync", "create_revolve_sync"),
        ("finite_sync", "create_revolve_finite_sync"),
        ("thin_wall", "create_revolve_thin_wall"),
        ("by_keypoint", "create_revolve_by_keypoint"),
        ("full_360", "create_revolve_full"),
        ("by_keypoint_sync", "create_revolve_by_keypoint_sync"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_revolve(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_revolve(method="bogus")
        assert "error" in result


# === create_extruded_cutout ===

class TestCreateExtrudedCutout:
    @pytest.mark.parametrize("disc, method", [
        ("finite", "create_extruded_cutout"),
        ("through_all", "create_extruded_cutout_through_all"),
        ("through_next", "create_extruded_cutout_through_next"),
        ("from_to", "create_extruded_cutout_from_to"),
        ("from_to_v2", "create_extruded_cutout_from_to_v2"),
        ("by_keypoint", "create_extruded_cutout_by_keypoint"),
        ("through_next_single", "create_extruded_cutout_through_next_single"),
        ("multi_body", "create_extruded_cutout_multi_body"),
        ("from_to_multi_body", "create_extruded_cutout_from_to_multi_body"),
        ("through_all_multi_body", "create_extruded_cutout_through_all_multi_body"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_extruded_cutout(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_extruded_cutout(method="bogus")
        assert "error" in result


# === create_revolved_cutout ===

class TestCreateRevolvedCutout:
    @pytest.mark.parametrize("disc, method", [
        ("finite", "create_revolved_cutout"),
        ("sync", "create_revolved_cutout_sync"),
        ("by_keypoint", "create_revolved_cutout_by_keypoint"),
        ("multi_body", "create_revolved_cutout_multi_body"),
        ("full", "create_revolved_cutout_full"),
        ("full_sync", "create_revolved_cutout_full_sync"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_revolved_cutout(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_revolved_cutout(method="bogus")
        assert "error" in result


# === create_normal_cutout ===

class TestCreateNormalCutout:
    @pytest.mark.parametrize("disc, method", [
        ("finite", "create_normal_cutout"),
        ("through_all", "create_normal_cutout_through_all"),
        ("from_to", "create_normal_cutout_from_to"),
        ("through_next", "create_normal_cutout_through_next"),
        ("by_keypoint", "create_normal_cutout_by_keypoint"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_normal_cutout(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_normal_cutout(method="bogus")
        assert "error" in result


# === create_lofted_cutout ===

class TestCreateLoftedCutout:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "create_lofted_cutout"),
        ("full", "create_lofted_cutout_full"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_lofted_cutout(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_lofted_cutout(method="bogus")
        assert "error" in result


# === create_swept_cutout ===

class TestCreateSweptCutout:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "create_swept_cutout"),
        ("multi_body", "create_swept_cutout_multi_body"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_swept_cutout(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_swept_cutout(method="bogus")
        assert "error" in result


# === create_helix ===

class TestCreateHelix:
    @pytest.mark.parametrize("disc, method", [
        ("finite", "create_helix"),
        ("sync", "create_helix_sync"),
        ("thin_wall", "create_helix_thin_wall"),
        ("sync_thin_wall", "create_helix_sync_thin_wall"),
        ("from_to", "create_helix_from_to"),
        ("from_to_thin_wall", "create_helix_from_to_thin_wall"),
        ("from_to_sync", "create_helix_from_to_sync"),
        ("from_to_sync_thin_wall", "create_helix_from_to_sync_thin_wall"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_helix(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_helix(method="bogus")
        assert "error" in result


# === create_helix_cutout ===

class TestCreateHelixCutout:
    @pytest.mark.parametrize("disc, method", [
        ("finite", "create_helix_cutout"),
        ("sync", "create_helix_cutout_sync"),
        ("from_to", "create_helix_cutout_from_to"),
        ("from_to_sync", "create_helix_cutout_from_to_sync"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_helix_cutout(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_helix_cutout(method="bogus")
        assert "error" in result


# === create_loft ===

class TestCreateLoft:
    @pytest.mark.parametrize("disc, method", [
        ("solid", "create_loft"),
        ("thin_wall", "create_loft_thin_wall"),
        ("with_guides", "create_loft_with_guides"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_loft(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_loft(method="bogus")
        assert "error" in result


# === create_sweep ===

class TestCreateSweep:
    @pytest.mark.parametrize("disc, method", [
        ("solid", "create_sweep"),
        ("thin_wall", "create_sweep_thin_wall"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_sweep(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_sweep(method="bogus")
        assert "error" in result


# === create_extruded_surface ===

class TestCreateExtrudedSurface:
    @pytest.mark.parametrize("disc, method", [
        ("finite", "create_extruded_surface"),
        ("from_to", "create_extruded_surface_from_to"),
        ("by_keypoint", "create_extruded_surface_by_keypoint"),
        ("by_curves", "create_extruded_surface_by_curves"),
        ("full", "create_extruded_surface_full"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_extruded_surface(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_extruded_surface(method="bogus")
        assert "error" in result


# === create_revolved_surface ===

class TestCreateRevolvedSurface:
    @pytest.mark.parametrize("disc, method", [
        ("finite", "create_revolved_surface"),
        ("sync", "create_revolved_surface_sync"),
        ("by_keypoint", "create_revolved_surface_by_keypoint"),
        ("full", "create_revolved_surface_full"),
        ("full_sync", "create_revolved_surface_full_sync"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_revolved_surface(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_revolved_surface(method="bogus")
        assert "error" in result


# === create_lofted_surface ===

class TestCreateLoftedSurface:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "create_lofted_surface"),
        ("v2", "create_lofted_surface_v2"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_lofted_surface(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_lofted_surface(method="bogus")
        assert "error" in result


# === create_swept_surface ===

class TestCreateSweptSurface:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "create_swept_surface"),
        ("ex", "create_swept_surface_ex"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_swept_surface(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_swept_surface(method="bogus")
        assert "error" in result


# === thicken ===

class TestThicken:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "thicken_surface"),
        ("sync", "create_thicken_sync"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = thicken(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = thicken(method="bogus")
        assert "error" in result


# === create_primitive ===

class TestCreatePrimitive:
    @pytest.mark.parametrize("disc, method", [
        ("box_two_points", "create_box_by_two_points"),
        ("box_center", "create_box_by_center"),
        ("box_three_points", "create_box_by_three_points"),
        ("cylinder", "create_cylinder"),
        ("sphere", "create_sphere"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_primitive(shape=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_primitive(shape="bogus")
        assert "error" in result


# === create_primitive_cutout ===

class TestCreatePrimitiveCutout:
    @pytest.mark.parametrize("disc, method", [
        ("box", "create_box_cutout_by_two_points"),
        ("cylinder", "create_cylinder_cutout"),
        ("sphere", "create_sphere_cutout"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_primitive_cutout(shape=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_primitive_cutout(shape="bogus")
        assert "error" in result


# === create_hole ===

class TestCreateHole:
    @pytest.mark.parametrize("disc, method", [
        ("finite", "create_hole"),
        ("through_all", "create_hole_through_all"),
        ("from_to", "create_hole_from_to"),
        ("through_next", "create_hole_through_next"),
        ("sync", "create_hole_sync"),
        ("finite_ex", "create_hole_finite_ex"),
        ("from_to_ex", "create_hole_from_to_ex"),
        ("through_next_ex", "create_hole_through_next_ex"),
        ("through_all_ex", "create_hole_through_all_ex"),
        ("sync_ex", "create_hole_sync_ex"),
        ("multi_body", "create_hole_multi_body"),
        ("sync_multi_body", "create_hole_sync_multi_body"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_hole(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_hole(method="bogus")
        assert "error" in result


# === create_round ===

class TestCreateRound:
    @pytest.mark.parametrize("disc, method", [
        ("all_edges", "create_round"),
        ("on_face", "create_round_on_face"),
        ("variable", "create_variable_round"),
        ("blend", "create_round_blend"),
        ("surface_blend", "create_round_surface_blend"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_round(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_round(method="bogus")
        assert "error" in result


# === create_chamfer ===

class TestCreateChamfer:
    @pytest.mark.parametrize("disc, method", [
        ("equal", "create_chamfer"),
        ("on_face", "create_chamfer_on_face"),
        ("unequal", "create_chamfer_unequal"),
        ("unequal_on_face", "create_chamfer_unequal_on_face"),
        ("angle", "create_chamfer_angle"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_chamfer(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_chamfer(method="bogus")
        assert "error" in result


# === create_blend ===

class TestCreateBlend:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "create_blend"),
        ("variable", "create_blend_variable"),
        ("surface", "create_blend_surface"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_blend(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_blend(method="bogus")
        assert "error" in result


# === delete_topology ===

class TestDeleteTopology:
    @pytest.mark.parametrize("disc, method", [
        ("hole", "create_delete_hole"),
        ("hole_by_face", "delete_hole_by_face"),
        ("blend", "create_delete_blend"),
        ("faces", "delete_faces"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = delete_topology(type=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = delete_topology(type="bogus")
        assert "error" in result


# === create_ref_plane ===

class TestCreateRefPlane:
    @pytest.mark.parametrize("disc, method", [
        ("offset", "create_ref_plane_by_offset"),
        ("angle", "create_ref_plane_by_angle"),
        ("three_points", "create_ref_plane_by_3_points"),
        ("midplane", "create_ref_plane_midplane"),
        ("normal_to_curve", "create_ref_plane_normal_to_curve"),
        ("normal_at_distance", "create_ref_plane_normal_at_distance"),
        ("normal_at_arc_ratio", "create_ref_plane_normal_at_arc_ratio"),
        ("normal_at_distance_along", "create_ref_plane_normal_at_distance_along"),
        ("parallel_by_tangent", "create_ref_plane_parallel_by_tangent"),
        ("normal_at_keypoint", "create_ref_plane_normal_at_keypoint"),
        ("tangent_cylinder_angle", "create_ref_plane_tangent_cylinder_angle"),
        ("tangent_cylinder_keypoint", "create_ref_plane_tangent_cylinder_keypoint"),
        ("tangent_surface_keypoint", "create_ref_plane_tangent_surface_keypoint"),
        ("normal_at_distance_v2", "create_ref_plane_normal_at_distance_v2"),
        ("normal_at_arc_ratio_v2", "create_ref_plane_normal_at_arc_ratio_v2"),
        ("normal_at_distance_along_v2", "create_ref_plane_normal_at_distance_along_v2"),
        ("tangent_parallel", "create_ref_plane_tangent_parallel"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_ref_plane(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_offset_passes_args(self, mock_mgr):
        mock_mgr.create_ref_plane_by_offset.return_value = {"status": "ok"}
        create_ref_plane(
            method="offset", parent_plane_index=2, distance=0.01, normal_side="Reverse",
        )
        mock_mgr.create_ref_plane_by_offset.assert_called_once_with(2, 0.01, "Reverse")

    def test_unknown(self, mock_mgr):
        result = create_ref_plane(method="bogus")
        assert "error" in result


# === create_flange ===

class TestCreateFlange:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "create_flange"),
        ("by_match_face", "create_flange_by_match_face"),
        ("sync", "create_flange_sync"),
        ("by_face", "create_flange_by_face"),
        ("with_bend_calc", "create_flange_with_bend_calc"),
        ("sync_with_bend_calc", "create_flange_sync_with_bend_calc"),
        ("match_face_with_bend", "create_flange_match_face_with_bend"),
        ("by_face_with_bend", "create_flange_by_face_with_bend"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_flange(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_by_match_face_defaults_radius(self, mock_mgr):
        """inside_radius=None should default to 0.001."""
        mock_mgr.create_flange_by_match_face.return_value = {"status": "ok"}
        create_flange(method="by_match_face", inside_radius=None)
        args = mock_mgr.create_flange_by_match_face.call_args[0]
        assert args[4] == 0.001  # inside_radius or 0.001

    def test_sync_defaults_radius(self, mock_mgr):
        """inside_radius=None should default to 0.001."""
        mock_mgr.create_flange_sync.return_value = {"status": "ok"}
        create_flange(method="sync", inside_radius=None)
        args = mock_mgr.create_flange_sync.call_args[0]
        assert args[3] == 0.001  # inside_radius or 0.001

    def test_match_face_with_bend_defaults_radius(self, mock_mgr):
        mock_mgr.create_flange_match_face_with_bend.return_value = {"status": "ok"}
        create_flange(method="match_face_with_bend", inside_radius=None)
        args = mock_mgr.create_flange_match_face_with_bend.call_args[0]
        assert args[4] == 0.001  # inside_radius or 0.001

    def test_unknown(self, mock_mgr):
        result = create_flange(method="bogus")
        assert "error" in result


# === create_contour_flange ===

class TestCreateContourFlange:
    @pytest.mark.parametrize("disc, method", [
        ("ex", "create_contour_flange_ex"),
        ("sync", "create_contour_flange_sync"),
        ("sync_with_bend", "create_contour_flange_sync_with_bend"),
        ("v3", "create_contour_flange_v3"),
        ("sync_ex", "create_contour_flange_sync_ex"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_contour_flange(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_contour_flange(method="bogus")
        assert "error" in result


# === create_sheet_metal_base ===

class TestCreateSheetMetalBase:
    @pytest.mark.parametrize("disc, method", [
        ("flange", "create_base_flange"),
        ("tab", "create_base_tab"),
        ("contour_advanced", "create_base_contour_flange_advanced"),
        ("tab_multi_profile", "create_base_tab_multi_profile"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_sheet_metal_base(type=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_flange_none_width_defaults_zero(self, mock_mgr):
        """width=None should default to 0.0."""
        mock_mgr.create_base_flange.return_value = {"status": "ok"}
        create_sheet_metal_base(type="flange", width=None, thickness=0.001)
        mock_mgr.create_base_flange.assert_called_once_with(0.0, 0.001, None)

    def test_contour_advanced_none_bend_radius_defaults(self, mock_mgr):
        """bend_radius=None should default to 0.001."""
        mock_mgr.create_base_contour_flange_advanced.return_value = {"status": "ok"}
        create_sheet_metal_base(type="contour_advanced", bend_radius=None, thickness=0.002)
        mock_mgr.create_base_contour_flange_advanced.assert_called_once_with(
            0.002, 0.001, "Default"
        )

    def test_unknown(self, mock_mgr):
        result = create_sheet_metal_base(type="bogus")
        assert "error" in result


# === create_lofted_flange ===

class TestCreateLoftedFlange:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "create_lofted_flange"),
        ("advanced", "create_lofted_flange_advanced"),
        ("ex", "create_lofted_flange_ex"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_lofted_flange(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_lofted_flange(method="bogus")
        assert "error" in result


# === create_bend ===

class TestCreateBend:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "create_bend"),
        ("with_calc", "create_bend_with_calc"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_bend(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_bend(method="bogus")
        assert "error" in result


# === create_slot ===

class TestCreateSlot:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "create_slot"),
        ("ex", "create_slot_ex"),
        ("sync", "create_slot_sync"),
        ("multi_body", "create_slot_multi_body"),
        ("sync_multi_body", "create_slot_sync_multi_body"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_slot(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_slot(method="bogus")
        assert "error" in result


# === create_thread ===

class TestCreateThread:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "create_thread"),
        ("physical", "create_thread_ex"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_thread(method=disc, face_index=2)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_zero_diameter_becomes_none(self, mock_mgr):
        mock_mgr.create_thread.return_value = {"status": "ok"}
        create_thread(method="basic", face_index=1, thread_diameter=0.0, thread_depth=0.0)
        mock_mgr.create_thread.assert_called_once_with(1, thread_diameter=None, thread_depth=None)

    def test_positive_diameter_passed(self, mock_mgr):
        mock_mgr.create_thread.return_value = {"status": "ok"}
        create_thread(method="basic", face_index=1, thread_diameter=0.01, thread_depth=0.02)
        mock_mgr.create_thread.assert_called_once_with(
            1, thread_diameter=0.01, thread_depth=0.02
        )

    def test_unknown(self, mock_mgr):
        result = create_thread(method="bogus")
        assert "error" in result


# === create_drawn_cutout ===

class TestCreateDrawnCutout:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "create_drawn_cutout"),
        ("ex", "create_drawn_cutout_ex"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_drawn_cutout(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_drawn_cutout(method="bogus")
        assert "error" in result


# === create_dimple ===

class TestCreateDimple:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "create_dimple"),
        ("ex", "create_dimple_ex"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_dimple(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_dimple(method="bogus")
        assert "error" in result


# === create_louver ===

class TestCreateLouver:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "create_louver"),
        ("sync", "create_louver_sync"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_louver(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_louver(method="bogus")
        assert "error" in result


# === create_pattern ===

class TestCreatePattern:
    @pytest.mark.parametrize("disc, method", [
        ("rectangular_ex", "create_pattern_rectangular_ex"),
        ("rectangular", "create_pattern_rectangular"),
        ("circular", "create_pattern_circular"),
        ("circular_ex", "create_pattern_circular_ex"),
        ("duplicate", "create_pattern_duplicate"),
        ("by_fill", "create_pattern_by_fill"),
        ("by_table", "create_pattern_by_table"),
        ("by_table_sync", "create_pattern_by_table_sync"),
        ("by_fill_ex", "create_pattern_by_fill_ex"),
        ("by_curve_ex", "create_pattern_by_curve_ex"),
        ("user_defined", "create_user_defined_pattern"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_pattern(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_pattern(method="bogus")
        assert "error" in result


# === create_mirror ===

class TestCreateMirror:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "create_mirror"),
        ("sync_ex", "create_mirror_sync_ex"),
        ("save_as_part", "save_as_mirror_part"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_mirror(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_save_as_part_zero_plane_defaults_to_3(self, mock_mgr):
        """mirror_plane_index=0 should default to 3."""
        mock_mgr.save_as_mirror_part.return_value = {"status": "ok"}
        create_mirror(method="save_as_part", new_file_name="mirror.par", mirror_plane_index=0)
        mock_mgr.save_as_mirror_part.assert_called_once_with("mirror.par", 3, True)

    def test_save_as_part_nonzero_plane_passed(self, mock_mgr):
        mock_mgr.save_as_mirror_part.return_value = {"status": "ok"}
        create_mirror(
            method="save_as_part",
            new_file_name="m.par",
            mirror_plane_index=2,
            link_to_original=False,
        )
        mock_mgr.save_as_mirror_part.assert_called_once_with("m.par", 2, False)

    def test_unknown(self, mock_mgr):
        result = create_mirror(method="bogus")
        assert "error" in result


# === create_thin_wall ===

class TestCreateThinWall:
    @pytest.mark.parametrize("disc", [
        "basic",
        "with_open_faces",
    ])
    def test_dispatch(self, mock_mgr, disc):
        mock_mgr.create_shell.return_value = {"status": "ok"}
        result = create_thin_wall(method=disc)
        mock_mgr.create_shell.assert_called_once()
        assert result == {"status": "ok"}

    def test_basic_no_open_faces(self, mock_mgr):
        mock_mgr.create_shell.return_value = {"status": "ok"}
        create_thin_wall(method="basic", thickness=0.002)
        mock_mgr.create_shell.assert_called_once_with(0.002)

    def test_with_open_faces_passes_list(self, mock_mgr):
        mock_mgr.create_shell.return_value = {"status": "ok"}
        create_thin_wall(method="with_open_faces", thickness=0.002, open_face_indices=[0, 1])
        mock_mgr.create_shell.assert_called_once_with(0.002, [0, 1])

    def test_unknown(self, mock_mgr):
        result = create_thin_wall(method="bogus")
        assert "error" in result


# === face_operation ===

class TestFaceOperation:
    @pytest.mark.parametrize("disc, method", [
        ("rotate_by_points", "create_face_rotate_by_points"),
        ("rotate_by_edge", "create_face_rotate_by_edge"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = face_operation(type=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = face_operation(type="bogus")
        assert "error" in result


# === add_body ===

class TestAddBody:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "add_body"),
        ("by_mesh", "add_body_by_mesh"),
        ("feature", "add_body_feature"),
        ("construction", "add_by_construction"),
        ("by_tag", "add_body_by_tag"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = add_body(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = add_body(method="bogus")
        assert "error" in result


# === simplify ===

class TestSimplify:
    @pytest.mark.parametrize("disc, method", [
        ("auto", "auto_simplify"),
        ("enclosure", "simplify_enclosure"),
        ("duplicate", "simplify_duplicate"),
        ("local_enclosure", "local_simplify_enclosure"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = simplify(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = simplify(method="bogus")
        assert "error" in result


# === manage_feature ===

class TestManageFeature:
    @pytest.mark.parametrize("disc, method", [
        ("delete", "delete_feature"),
        ("suppress", "feature_suppress"),
        ("unsuppress", "feature_unsuppress"),
        ("reorder", "feature_reorder"),
        ("rename", "feature_rename"),
        ("convert", "convert_feature_type"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = manage_feature(action=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = manage_feature(action="bogus")
        assert "error" in result


# === sheet_metal_misc ===

class TestSheetMetalMisc:
    @pytest.mark.parametrize("disc, method", [
        ("hem", "create_hem"),
        ("jog", "create_jog"),
        ("close_corner", "create_close_corner"),
        ("multi_edge_flange", "create_multi_edge_flange"),
        ("convert", "convert_part_to_sheet_metal"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = sheet_metal_misc(action=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_multi_edge_flange_defaults_empty_edges(self, mock_mgr):
        mock_mgr.create_multi_edge_flange.return_value = {"status": "ok"}
        sheet_metal_misc(action="multi_edge_flange")
        args = mock_mgr.create_multi_edge_flange.call_args[0]
        assert args[1] == []  # edge_indices or []

    def test_unknown(self, mock_mgr):
        result = sheet_metal_misc(action="bogus")
        assert "error" in result


# === create_stamped ===

class TestCreateStamped:
    @pytest.mark.parametrize("disc, method", [
        ("bead", "create_bead"),
        ("gusset", "create_gusset"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_stamped(type=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_stamped(type="bogus")
        assert "error" in result


# === create_surface_mark ===

class TestCreateSurfaceMark:
    @pytest.mark.parametrize("disc, method", [
        ("emboss", "create_emboss"),
        ("etch", "create_etch"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_surface_mark(type=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_emboss_defaults_empty_faces(self, mock_mgr):
        mock_mgr.create_emboss.return_value = {"status": "ok"}
        create_surface_mark(type="emboss")
        mock_mgr.create_emboss.assert_called_once_with([], 0.001, 0.0, False, True)

    def test_unknown(self, mock_mgr):
        result = create_surface_mark(type="bogus")
        assert "error" in result


# === create_reinforcement ===

class TestCreateReinforcement:
    @pytest.mark.parametrize("disc, method", [
        ("rib", "create_rib"),
        ("lip", "create_lip"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_reinforcement(type=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_rib_passes_args(self, mock_mgr):
        mock_mgr.create_rib.return_value = {"status": "ok"}
        create_reinforcement(type="rib", thickness=0.005, direction="Reverse")
        mock_mgr.create_rib.assert_called_once_with(0.005, "Reverse")

    def test_unknown(self, mock_mgr):
        result = create_reinforcement(type="bogus")
        assert "error" in result


# === Standalone tools ===

class TestStandaloneFeatures:
    def test_create_web_network(self, mock_mgr):
        mock_mgr.create_web_network.return_value = {"status": "ok"}
        result = create_web_network()
        mock_mgr.create_web_network.assert_called_once()
        assert result == {"status": "ok"}

    def test_create_split(self, mock_mgr):
        mock_mgr.create_split.return_value = {"status": "ok"}
        result = create_split()
        mock_mgr.create_split.assert_called_once()
        assert result == {"status": "ok"}

    def test_create_draft_angle(self, mock_mgr):
        mock_mgr.create_draft_angle.return_value = {"status": "ok"}
        result = create_draft_angle(face_index=0, angle=5.0, plane_index=2)
        mock_mgr.create_draft_angle.assert_called_once_with(0, 5.0, 2)
        assert result == {"status": "ok"}

    def test_create_bounded_surface(self, mock_mgr):
        mock_mgr.create_bounded_surface.return_value = {"status": "ok"}
        result = create_bounded_surface(want_end_caps=False, periodic=True)
        mock_mgr.create_bounded_surface.assert_called_once_with(False, True)
        assert result == {"status": "ok"}
