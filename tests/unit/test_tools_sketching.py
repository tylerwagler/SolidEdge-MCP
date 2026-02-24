"""Dispatch tests for tools/sketching.py composite tools."""

from unittest.mock import MagicMock

import pytest

from solidedge_mcp.tools.sketching import (
    draw,
    manage_sketch,
    sketch_advanced_modify,
    sketch_constraint,
    sketch_modify,
    sketch_project,
)


@pytest.fixture
def mock_mgr(monkeypatch):
    mgr = MagicMock()
    monkeypatch.setattr("solidedge_mcp.tools.sketching.sketch_manager", mgr)
    return mgr


# === manage_sketch ===

class TestManageSketch:
    @pytest.mark.parametrize("disc, method", [
        ("create", "create_sketch"),
        ("close", "close_sketch"),
        ("create_on_plane", "create_sketch_on_plane_index"),
        ("set_axis", "set_axis_of_revolution"),
        ("set_visibility", "hide_profile"),
        ("get_geometry", "get_ordered_geometry"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = manage_sketch(action=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_create_passes_plane(self, mock_mgr):
        mock_mgr.create_sketch.return_value = {"status": "ok"}
        manage_sketch(action="create", plane="Front")
        mock_mgr.create_sketch.assert_called_once_with("Front")

    def test_set_axis_passes_coords(self, mock_mgr):
        mock_mgr.set_axis_of_revolution.return_value = {"status": "ok"}
        manage_sketch(action="set_axis", x1=0.0, y1=0.0, x2=0.1, y2=0.0)
        mock_mgr.set_axis_of_revolution.assert_called_once_with(0.0, 0.0, 0.1, 0.0)

    def test_unknown(self, mock_mgr):
        result = manage_sketch(action="bogus")
        assert "error" in result


# === draw ===

class TestDraw:
    @pytest.mark.parametrize("disc, method", [
        ("line", "draw_line"),
        ("circle", "draw_circle"),
        ("rectangle", "draw_rectangle"),
        ("arc", "draw_arc"),
        ("polygon", "draw_polygon"),
        ("ellipse", "draw_ellipse"),
        ("spline", "draw_spline"),
        ("arc_3pt", "draw_arc_by_3_points"),
        ("circle_2pt", "draw_circle_by_2_points"),
        ("circle_3pt", "draw_circle_by_3_points"),
        ("point", "draw_point"),
        ("construction_line", "draw_construction_line"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = draw(shape=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_line_passes_coords(self, mock_mgr):
        mock_mgr.draw_line.return_value = {"status": "ok"}
        draw(shape="line", x1=0.1, y1=0.2, x2=0.3, y2=0.4)
        mock_mgr.draw_line.assert_called_once_with(0.1, 0.2, 0.3, 0.4)

    def test_circle_passes_params(self, mock_mgr):
        mock_mgr.draw_circle.return_value = {"status": "ok"}
        draw(shape="circle", center_x=0.1, center_y=0.2, radius=0.05)
        mock_mgr.draw_circle.assert_called_once_with(0.1, 0.2, 0.05)

    def test_spline_defaults_empty_list(self, mock_mgr):
        mock_mgr.draw_spline.return_value = {"status": "ok"}
        draw(shape="spline")
        mock_mgr.draw_spline.assert_called_once_with([])

    def test_spline_passes_points(self, mock_mgr):
        mock_mgr.draw_spline.return_value = {"status": "ok"}
        pts = [[0, 0], [1, 1]]
        draw(shape="spline", points=pts)
        mock_mgr.draw_spline.assert_called_once_with(pts)

    def test_unknown(self, mock_mgr):
        result = draw(shape="bogus")
        assert "error" in result


# === sketch_modify ===

class TestSketchModify:
    @pytest.mark.parametrize("disc, method", [
        ("fillet", "sketch_fillet"),
        ("chamfer", "sketch_chamfer"),
        ("offset", "sketch_offset"),
        ("rotate", "sketch_rotate"),
        ("scale", "sketch_scale"),
        ("mirror", "sketch_mirror"),
        ("paste", "sketch_paste"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = sketch_modify(action=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_fillet_passes_radius(self, mock_mgr):
        mock_mgr.sketch_fillet.return_value = {"status": "ok"}
        sketch_modify(action="fillet", radius=0.005)
        mock_mgr.sketch_fillet.assert_called_once_with(0.005)

    def test_rotate_passes_args(self, mock_mgr):
        mock_mgr.sketch_rotate.return_value = {"status": "ok"}
        sketch_modify(action="rotate", center_x=0.1, center_y=0.2, angle_degrees=45.0)
        mock_mgr.sketch_rotate.assert_called_once_with(0.1, 0.2, 45.0)

    def test_unknown(self, mock_mgr):
        result = sketch_modify(action="bogus")
        assert "error" in result


# === sketch_advanced_modify ===

class TestSketchAdvancedModify:
    @pytest.mark.parametrize("disc, method", [
        ("mirror_spline", "mirror_spline"),
        ("offset_2d", "offset_sketch_2d"),
        ("clean", "clean_sketch_geometry"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = sketch_advanced_modify(action=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_mirror_spline_passes_args(self, mock_mgr):
        mock_mgr.mirror_spline.return_value = {"status": "ok"}
        sketch_advanced_modify(
            action="mirror_spline",
            axis_x1=0.0, axis_y1=0.0, axis_x2=1.0, axis_y2=0.0, copy=False,
        )
        mock_mgr.mirror_spline.assert_called_once_with(0.0, 0.0, 1.0, 0.0, False)

    def test_offset_2d_passes_args(self, mock_mgr):
        mock_mgr.offset_sketch_2d.return_value = {"status": "ok"}
        sketch_advanced_modify(
            action="offset_2d",
            offset_side_x=1.0, offset_side_y=0.0, offset_distance=0.01,
        )
        mock_mgr.offset_sketch_2d.assert_called_once_with(1.0, 0.0, 0.01)

    def test_clean_passes_args(self, mock_mgr):
        mock_mgr.clean_sketch_geometry.return_value = {"status": "ok"}
        sketch_advanced_modify(
            action="clean",
            clean_points=False, clean_splines=True,
            clean_identical=False, clean_small=True, small_tolerance=0.01,
        )
        mock_mgr.clean_sketch_geometry.assert_called_once_with(
            False, True, False, True, 0.01,
        )

    def test_unknown(self, mock_mgr):
        result = sketch_advanced_modify(action="bogus")
        assert "error" in result


# === sketch_constraint ===

class TestSketchConstraint:
    @pytest.mark.parametrize("disc, method", [
        ("geometric", "add_constraint"),
        ("keypoint", "add_keypoint_constraint"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = sketch_constraint(type=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_geometric_defaults_empty_elements(self, mock_mgr):
        mock_mgr.add_constraint.return_value = {"status": "ok"}
        sketch_constraint(type="geometric", constraint_type="Horizontal")
        mock_mgr.add_constraint.assert_called_once_with("Horizontal", [])

    def test_geometric_passes_elements(self, mock_mgr):
        mock_mgr.add_constraint.return_value = {"status": "ok"}
        sketch_constraint(type="geometric", constraint_type="Horizontal", elements=[1, 2])
        mock_mgr.add_constraint.assert_called_once_with("Horizontal", [1, 2])

    def test_unknown(self, mock_mgr):
        result = sketch_constraint(type="bogus")
        assert "error" in result


# === sketch_project ===

class TestSketchProject:
    @pytest.mark.parametrize("disc, method", [
        ("edge", "project_edge"),
        ("include_edge", "include_edge"),
        ("ref_plane", "project_ref_plane"),
        ("silhouette", "project_silhouette_edges"),
        ("region_faces", "include_region_faces"),
        ("chain", "chain_locate"),
        ("to_curve", "convert_to_curve"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = sketch_project(source=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_region_faces_defaults_empty(self, mock_mgr):
        mock_mgr.include_region_faces.return_value = {"status": "ok"}
        sketch_project(source="region_faces")
        mock_mgr.include_region_faces.assert_called_once_with([])

    def test_region_faces_passes_list(self, mock_mgr):
        mock_mgr.include_region_faces.return_value = {"status": "ok"}
        sketch_project(source="region_faces", face_indices=[1, 2])
        mock_mgr.include_region_faces.assert_called_once_with([1, 2])

    def test_unknown(self, mock_mgr):
        result = sketch_project(source="bogus")
        assert "error" in result
