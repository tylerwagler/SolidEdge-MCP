"""Dispatch tests for tools/query.py composite tools."""

from unittest.mock import MagicMock

import pytest

from solidedge_mcp.tools.query import (
    edit_feature_extent,
    manage_feature_tree,
    manage_layer,
    manage_material,
    manage_property,
    manage_variable,
    measure,
    query_body,
    query_bspline,
    query_edge,
    query_face,
    recompute,
    select_set,
    set_appearance,
)


@pytest.fixture
def mock_mgr(monkeypatch):
    mgr = MagicMock()
    monkeypatch.setattr("solidedge_mcp.tools.query.query_manager", mgr)
    return mgr


# === measure ===

class TestMeasure:
    @pytest.mark.parametrize("disc, method", [
        ("distance", "measure_distance"),
        ("angle", "measure_angle"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = measure(type=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_distance_passes_coords(self, mock_mgr):
        mock_mgr.measure_distance.return_value = {"distance": 1.0}
        measure(type="distance", x1=0.0, y1=0.0, z1=0.0, x2=1.0, y2=0.0, z2=0.0)
        mock_mgr.measure_distance.assert_called_once_with(0.0, 0.0, 0.0, 1.0, 0.0, 0.0)

    def test_unknown(self, mock_mgr):
        result = measure(type="bogus")
        assert "error" in result


# === manage_variable ===

class TestManageVariable:
    @pytest.mark.parametrize("disc, method, kwargs", [
        ("set", "set_variable", {"value": 0.0}),
        ("add", "add_variable", {"formula": "0"}),
        ("query", "query_variables", {}),
        ("rename", "rename_variable", {"new_name": "x"}),
        ("translate", "translate_variable", {}),
        ("copy_clipboard", "copy_variable_to_clipboard", {}),
        ("add_from_clipboard", "add_variable_from_clipboard", {}),
        ("set_formula", "set_variable_formula", {"formula": "0"}),
    ])
    def test_dispatch(self, mock_mgr, disc, method, kwargs):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = manage_variable(action=disc, **kwargs)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_set_passes_name_value(self, mock_mgr):
        mock_mgr.set_variable.return_value = {"status": "ok"}
        manage_variable(action="set", name="Width", value=0.1)
        mock_mgr.set_variable.assert_called_once_with("Width", 0.1)

    def test_unknown(self, mock_mgr):
        result = manage_variable(action="bogus")
        assert "error" in result


# === manage_property ===

class TestManageProperty:
    @pytest.mark.parametrize("disc, method", [
        ("set_document", "set_document_property"),
        ("set_custom", "set_custom_property"),
        ("delete_custom", "delete_custom_property"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = manage_property(action=disc, name="Title", value="Test")
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = manage_property(action="bogus")
        assert "error" in result


# === manage_material ===

class TestManageMaterial:
    @pytest.mark.parametrize("disc, method", [
        ("set", "set_material"),
        ("set_density", "set_material_density"),
        ("set_by_name", "set_material_by_name"),
        ("get_library", "get_material_library"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = manage_material(action=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = manage_material(action="bogus")
        assert "error" in result


# === set_appearance ===

class TestSetAppearance:
    @pytest.mark.parametrize("disc, method", [
        ("body_color", "set_body_color"),
        ("face_color", "set_face_color"),
        ("opacity", "set_body_opacity"),
        ("reflectivity", "set_body_reflectivity"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = set_appearance(target=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_body_color_passes_rgb(self, mock_mgr):
        mock_mgr.set_body_color.return_value = {"status": "ok"}
        set_appearance(target="body_color", red=255, green=128, blue=0)
        mock_mgr.set_body_color.assert_called_once_with(255, 128, 0)

    def test_face_color_passes_args(self, mock_mgr):
        mock_mgr.set_face_color.return_value = {"status": "ok"}
        set_appearance(target="face_color", face_index=2, red=100, green=200, blue=50)
        mock_mgr.set_face_color.assert_called_once_with(2, 100, 200, 50)

    def test_unknown(self, mock_mgr):
        result = set_appearance(target="bogus")
        assert "error" in result


# === manage_layer ===

class TestManageLayer:
    @pytest.mark.parametrize("disc, method", [
        ("add", "add_layer"),
        ("activate", "activate_layer"),
        ("set_properties", "set_layer_properties"),
        ("delete", "delete_layer"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = manage_layer(action=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = manage_layer(action="bogus")
        assert "error" in result


# === select_set ===

class TestSelectSet:
    @pytest.mark.parametrize("disc, method", [
        ("clear", "clear_select_set"),
        ("add", "select_add"),
        ("remove", "select_remove"),
        ("all", "select_all"),
        ("copy", "select_copy"),
        ("cut", "select_cut"),
        ("delete", "select_delete"),
        ("suspend_display", "select_suspend_display"),
        ("resume_display", "select_resume_display"),
        ("refresh_display", "select_refresh_display"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = select_set(action=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_add_passes_args(self, mock_mgr):
        mock_mgr.select_add.return_value = {"status": "ok"}
        select_set(action="add", object_type="Face", index=3)
        mock_mgr.select_add.assert_called_once_with("Face", 3)

    def test_unknown(self, mock_mgr):
        result = select_set(action="bogus")
        assert "error" in result


# === edit_feature_extent ===

class TestEditFeatureExtent:
    @pytest.mark.parametrize("disc, method", [
        ("get_direction1", "get_direction1_extent"),
        ("set_direction1", "set_direction1_extent"),
        ("get_direction2", "get_direction2_extent"),
        ("set_direction2", "set_direction2_extent"),
        ("get_thin_wall", "get_thin_wall_options"),
        ("set_thin_wall", "set_thin_wall_options"),
        ("get_from_face", "get_from_face_offset"),
        ("set_from_face", "set_from_face_offset"),
        ("get_body_array", "get_body_array"),
        ("set_body_array", "set_body_array"),
        ("get_to_face", "get_to_face_offset"),
        ("set_to_face", "set_to_face_offset"),
        ("get_direction1_treatment", "get_direction1_treatment"),
        ("apply_direction1_treatment", "apply_direction1_treatment"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = edit_feature_extent(property=disc, feature_name="Extrude1")
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_set_direction1_passes_args(self, mock_mgr):
        mock_mgr.set_direction1_extent.return_value = {"status": "ok"}
        edit_feature_extent(
            property="set_direction1",
            feature_name="Extrude1",
            extent_type=13,
            distance=0.05,
        )
        mock_mgr.set_direction1_extent.assert_called_once_with("Extrude1", 13, 0.05)

    def test_unknown(self, mock_mgr):
        result = edit_feature_extent(property="bogus")
        assert "error" in result


# === manage_feature_tree ===

class TestManageFeatureTree:
    @pytest.mark.parametrize("disc, method", [
        ("rename", "rename_feature"),
        ("suppress", "suppress_feature"),
        ("unsuppress", "unsuppress_feature"),
        ("set_mode", "set_modeling_mode"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = manage_feature_tree(action=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_rename_passes_args(self, mock_mgr):
        mock_mgr.rename_feature.return_value = {"status": "ok"}
        manage_feature_tree(action="rename", feature_name="Extrude1", new_name="Base")
        mock_mgr.rename_feature.assert_called_once_with("Extrude1", "Base")

    def test_unknown(self, mock_mgr):
        result = manage_feature_tree(action="bogus")
        assert "error" in result


# === query_edge ===

class TestQueryEdge:
    @pytest.mark.parametrize("disc, method", [
        ("endpoints", "get_edge_endpoints"),
        ("length", "get_edge_length"),
        ("tangent", "get_edge_tangent"),
        ("geometry", "get_edge_geometry"),
        ("curvature", "get_edge_curvature"),
        ("vertex", "get_vertex_point"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = query_edge(property=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_tangent_passes_param(self, mock_mgr):
        mock_mgr.get_edge_tangent.return_value = {"status": "ok"}
        query_edge(property="tangent", face_index=1, edge_index=2, param=0.75)
        mock_mgr.get_edge_tangent.assert_called_once_with(1, 2, 0.75)

    def test_unknown(self, mock_mgr):
        result = query_edge(property="bogus")
        assert "error" in result


# === query_face ===

class TestQueryFace:
    @pytest.mark.parametrize("disc, method", [
        ("normal", "get_face_normal"),
        ("geometry", "get_face_geometry"),
        ("loops", "get_face_loops"),
        ("curvature", "get_face_curvature"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = query_face(property=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_normal_passes_uv(self, mock_mgr):
        mock_mgr.get_face_normal.return_value = {"status": "ok"}
        query_face(property="normal", face_index=1, u=0.3, v=0.7)
        mock_mgr.get_face_normal.assert_called_once_with(1, 0.3, 0.7)

    def test_unknown(self, mock_mgr):
        result = query_face(property="bogus")
        assert "error" in result


# === query_body ===

class TestQueryBody:
    @pytest.mark.parametrize("disc, method", [
        ("extreme_point", "get_body_extreme_point"),
        ("faces_by_ray", "get_faces_by_ray"),
        ("shells", "get_body_shells"),
        ("vertices", "get_body_vertices"),
        ("shell_info", "get_shell_info"),
        ("point_inside", "is_point_inside_body"),
        ("user_physical_properties", "get_user_physical_properties"),
        ("facet_data", "get_body_facet_data"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = query_body(property=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = query_body(property="bogus")
        assert "error" in result


# === query_bspline ===

class TestQueryBspline:
    @pytest.mark.parametrize("disc, method", [
        ("curve", "get_bspline_curve_info"),
        ("surface", "get_bspline_surface_info"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = query_bspline(type=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_curve_passes_indices(self, mock_mgr):
        mock_mgr.get_bspline_curve_info.return_value = {"status": "ok"}
        query_bspline(type="curve", face_index=1, edge_index=2)
        mock_mgr.get_bspline_curve_info.assert_called_once_with(1, 2)

    def test_unknown(self, mock_mgr):
        result = query_bspline(type="bogus")
        assert "error" in result


# === recompute ===

class TestRecompute:
    @pytest.mark.parametrize("disc, method", [
        ("model", "recompute"),
        ("document", "recompute_document"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = recompute(scope=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = recompute(scope="bogus")
        assert "error" in result
