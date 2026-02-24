"""Dispatch tests for tools/assembly.py composite tools."""

from unittest.mock import MagicMock

import pytest

from solidedge_mcp.tools.assembly import (
    add_assembly_component,
    add_assembly_constraint,
    add_assembly_relation,
    assembly_feature,
    manage_component,
    manage_relation,
    query_component,
    rotate_component,
    set_component_appearance,
    set_component_orientation,
    structural_frame,
    transform_component,
    virtual_component,
    wiring,
)


@pytest.fixture
def mock_mgr(monkeypatch):
    mgr = MagicMock()
    monkeypatch.setattr("solidedge_mcp.tools.assembly.assembly_manager", mgr)
    # Bypass path validation so tests with dummy file paths still pass
    monkeypatch.setattr(
        "solidedge_mcp.tools.assembly.validate_path",
        lambda p, **kw: (p, None),
    )
    return mgr


# === add_assembly_component ===

class TestAddAssemblyComponent:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "add_component"),
        ("with_transform", "add_component_with_transform"),
        ("family", "add_family_member"),
        ("family_with_transform", "add_family_with_transform"),
        ("family_with_matrix", "add_family_with_matrix"),
        ("by_template", "add_by_template"),
        ("adjustable", "add_adjustable_part"),
        ("tube", "add_tube"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = add_assembly_component(method=disc, file_path="p.par")
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_family_with_matrix_defaults_empty(self, mock_mgr):
        mock_mgr.add_family_with_matrix.return_value = {"status": "ok"}
        add_assembly_component(method="family_with_matrix", file_path="p.par")
        mock_mgr.add_family_with_matrix.assert_called_once_with("p.par", "", [])

    def test_tube_defaults_empty(self, mock_mgr):
        mock_mgr.add_tube.return_value = {"status": "ok"}
        add_assembly_component(method="tube", file_path="tube.par")
        mock_mgr.add_tube.assert_called_once_with([], "tube.par")

    def test_unknown(self, mock_mgr):
        result = add_assembly_component(method="bogus")
        assert "error" in result


# === manage_component ===

class TestManageComponent:
    @pytest.mark.parametrize("disc, method", [
        ("delete", "delete_component"),
        ("replace", "replace_component"),
        ("suppress", "suppress_component"),
        ("reorder", "reorder_occurrence"),
        ("make_writable", "make_writable"),
        ("swap_family", "swap_family_member"),
        ("ground", "ground_component"),
        ("pattern", "pattern_component"),
        ("mirror", "mirror_component"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = manage_component(action=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = manage_component(action="bogus")
        assert "error" in result


# === query_component ===

class TestQueryComponent:
    @pytest.mark.parametrize("disc, method", [
        ("list", "list_components"),
        ("info", "get_component_info"),
        ("bounding_box", "get_occurrence_bounding_box"),
        ("bom", "get_bom"),
        ("structured_bom", "get_structured_bom"),
        ("tree", "get_document_tree"),
        ("transform", "get_component_transform"),
        ("count", "get_occurrence_count"),
        ("is_subassembly", "is_subassembly"),
        ("display_name", "get_component_display_name"),
        ("document", "get_occurrence_document"),
        ("sub_occurrences", "get_sub_occurrences"),
        ("bodies", "get_occurrence_bodies"),
        ("style", "get_occurrence_style"),
        ("is_tube", "is_tube"),
        ("adjustable_part", "get_adjustable_part"),
        ("face_style", "get_face_style"),
        ("occurrence", "get_occurrence"),
        ("interference", "check_interference"),
        ("tube", "get_tube"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        # interference needs a nonzero component_index to avoid None transform
        if disc == "interference":
            result = query_component(property=disc, component_index=1)
        elif disc == "occurrence":
            result = query_component(property=disc, internal_id=42)
        else:
            result = query_component(property=disc, component_index=1)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_interference_zero_index_passes_none(self, mock_mgr):
        mock_mgr.check_interference.return_value = {"status": "ok"}
        query_component(property="interference", component_index=0)
        mock_mgr.check_interference.assert_called_once_with(None)

    def test_interference_nonzero_index_passes_value(self, mock_mgr):
        mock_mgr.check_interference.return_value = {"status": "ok"}
        query_component(property="interference", component_index=5)
        mock_mgr.check_interference.assert_called_once_with(5)

    def test_unknown(self, mock_mgr):
        result = query_component(property="bogus")
        assert "error" in result


# === set_component_appearance ===

class TestSetComponentAppearance:
    @pytest.mark.parametrize("disc, method", [
        ("visibility", "set_component_visibility"),
        ("color", "set_component_color"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = set_component_appearance(property=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_color_passes_rgb(self, mock_mgr):
        mock_mgr.set_component_color.return_value = {"status": "ok"}
        set_component_appearance(property="color", component_index=1, red=255, green=0, blue=128)
        mock_mgr.set_component_color.assert_called_once_with(1, 255, 0, 128)

    def test_unknown(self, mock_mgr):
        result = set_component_appearance(property="bogus")
        assert "error" in result


# === transform_component ===

class TestTransformComponent:
    @pytest.mark.parametrize("disc, method", [
        ("update_position", "update_component_position"),
        ("set_origin", "set_component_origin"),
        ("put_origin", "put_origin"),
        ("move", "occurrence_move"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = transform_component(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_update_position_passes_xyz(self, mock_mgr):
        mock_mgr.update_component_position.return_value = {"status": "ok"}
        transform_component(method="update_position", component_index=2, x=0.1, y=0.2, z=0.3)
        mock_mgr.update_component_position.assert_called_once_with(2, 0.1, 0.2, 0.3)

    def test_unknown(self, mock_mgr):
        result = transform_component(method="bogus")
        assert "error" in result


# === set_component_orientation ===

class TestSetComponentOrientation:
    @pytest.mark.parametrize("disc, method", [
        ("set_transform", "set_component_transform"),
        ("put_euler", "put_transform_euler"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = set_component_orientation(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_set_transform_passes_args(self, mock_mgr):
        mock_mgr.set_component_transform.return_value = {"status": "ok"}
        set_component_orientation(
            method="set_transform", component_index=1,
            origin_x=0.1, origin_y=0.2, origin_z=0.3,
            angle_x=10, angle_y=20, angle_z=30,
        )
        mock_mgr.set_component_transform.assert_called_once_with(
            1, 0.1, 0.2, 0.3, 10, 20, 30,
        )

    def test_put_euler_passes_args(self, mock_mgr):
        mock_mgr.put_transform_euler.return_value = {"status": "ok"}
        set_component_orientation(
            method="put_euler", component_index=2,
            x=0.1, y=0.2, z=0.3, rx=45, ry=90, rz=180,
        )
        mock_mgr.put_transform_euler.assert_called_once_with(
            2, 0.1, 0.2, 0.3, 45, 90, 180,
        )

    def test_unknown(self, mock_mgr):
        result = set_component_orientation(method="bogus")
        assert "error" in result


# === rotate_component ===

class TestRotateComponent:
    def test_dispatch(self, mock_mgr):
        mock_mgr.occurrence_rotate.return_value = {"status": "ok"}
        result = rotate_component(
            component_index=1,
            axis_x1=0, axis_y1=0, axis_z1=0,
            axis_x2=0, axis_y2=0, axis_z2=1,
            angle=90,
        )
        mock_mgr.occurrence_rotate.assert_called_once_with(
            1, 0, 0, 0, 0, 0, 1, 90,
        )
        assert result == {"status": "ok"}


# === add_assembly_constraint ===

class TestAddAssemblyConstraint:
    @pytest.mark.parametrize("disc, method", [
        ("mate", "create_mate"),
        ("align", "add_align_constraint"),
        ("planar_align", "add_planar_align_constraint"),
        ("axial_align", "add_axial_align_constraint"),
        ("angle", "add_angle_constraint"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = add_assembly_constraint(type=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = add_assembly_constraint(type="bogus")
        assert "error" in result


# === add_assembly_relation ===

class TestAddAssemblyRelation:
    @pytest.mark.parametrize("disc, method", [
        ("planar", "add_planar_relation"),
        ("axial", "add_axial_relation"),
        ("angular", "add_angular_relation"),
        ("point", "add_point_relation"),
        ("tangent", "add_tangent_relation"),
        ("gear", "add_gear_relation"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = add_assembly_relation(type=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = add_assembly_relation(type="bogus")
        assert "error" in result


# === manage_relation ===

class TestManageRelation:
    @pytest.mark.parametrize("disc, method", [
        ("list", "get_assembly_relations"),
        ("info", "get_relation_info"),
        ("delete", "delete_relation"),
        ("get_offset", "get_relation_offset"),
        ("set_offset", "set_relation_offset"),
        ("get_angle", "get_relation_angle"),
        ("set_angle", "set_relation_angle"),
        ("get_normals", "get_normals_aligned"),
        ("set_normals", "set_normals_aligned"),
        ("suppress", "suppress_relation"),
        ("unsuppress", "unsuppress_relation"),
        ("get_geometry", "get_relation_geometry"),
        ("get_gear_ratio", "get_gear_ratio"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = manage_relation(action=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = manage_relation(action="bogus")
        assert "error" in result


# === assembly_feature ===

class TestAssemblyFeature:
    @pytest.mark.parametrize("disc, method", [
        ("extruded_cutout", "create_assembly_extruded_cutout"),
        ("revolved_cutout", "create_assembly_revolved_cutout"),
        ("hole", "create_assembly_hole"),
        ("extruded_protrusion", "create_assembly_extruded_protrusion"),
        ("revolved_protrusion", "create_assembly_revolved_protrusion"),
        ("mirror", "create_assembly_mirror"),
        ("pattern", "create_assembly_pattern"),
        ("swept_protrusion", "create_assembly_swept_protrusion"),
        ("recompute", "recompute_assembly_features"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = assembly_feature(type=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_extruded_cutout_defaults_empty_scope(self, mock_mgr):
        mock_mgr.create_assembly_extruded_cutout.return_value = {"status": "ok"}
        assembly_feature(type="extruded_cutout")
        args = mock_mgr.create_assembly_extruded_cutout.call_args[0]
        assert args[0] == []  # scope_parts or []

    def test_mirror_defaults_empty_features(self, mock_mgr):
        mock_mgr.create_assembly_mirror.return_value = {"status": "ok"}
        assembly_feature(type="mirror")
        args = mock_mgr.create_assembly_mirror.call_args[0]
        assert args[0] == []  # feature_indices or []

    def test_unknown(self, mock_mgr):
        result = assembly_feature(type="bogus")
        assert "error" in result


# === virtual_component ===

class TestVirtualComponent:
    @pytest.mark.parametrize("disc, method", [
        ("new", "add_virtual_component"),
        ("predefined", "add_virtual_component_predefined"),
        ("bidm", "add_virtual_component_bidm"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = virtual_component(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = virtual_component(method="bogus")
        assert "error" in result


# === structural_frame ===

class TestStructuralFrame:
    @pytest.mark.parametrize("disc, method", [
        ("basic", "add_structural_frame"),
        ("by_orientation", "add_structural_frame_by_orientation"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = structural_frame(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_basic_defaults_empty_paths(self, mock_mgr):
        mock_mgr.add_structural_frame.return_value = {"status": "ok"}
        structural_frame(method="basic", part_filename="frame.par")
        mock_mgr.add_structural_frame.assert_called_once_with("frame.par", [])

    def test_unknown(self, mock_mgr):
        result = structural_frame(method="bogus")
        assert "error" in result


# === wiring ===

class TestWiring:
    @pytest.mark.parametrize("disc, method", [
        ("wire", "add_wire"),
        ("cable", "add_cable"),
        ("bundle", "add_bundle"),
        ("splice", "add_splice"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = wiring(type=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_wire_defaults_empty_lists(self, mock_mgr):
        mock_mgr.add_wire.return_value = {"status": "ok"}
        wiring(type="wire")
        mock_mgr.add_wire.assert_called_once_with([], [], "")

    def test_splice_defaults_empty_conductors(self, mock_mgr):
        mock_mgr.add_splice.return_value = {"status": "ok"}
        wiring(type="splice", x=0.1, y=0.2, z=0.3)
        mock_mgr.add_splice.assert_called_once_with(0.1, 0.2, 0.3, [], "")

    def test_unknown(self, mock_mgr):
        result = wiring(type="bogus")
        assert "error" in result
