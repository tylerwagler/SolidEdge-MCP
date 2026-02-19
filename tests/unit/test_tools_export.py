"""Dispatch tests for tools/export.py composite tools."""

from unittest.mock import MagicMock

import pytest

from solidedge_mcp.tools.export import (
    add_2d_dimension,
    add_annotation,
    add_drawing_view,
    add_smart_frame,
    camera_control,
    create_table,
    display_control,
    draft_config,
    export_file,
    manage_annotation_data,
    manage_drawing_view,
    manage_sheet,
    print_control,
    query_sheet,
)


@pytest.fixture
def mock_export(monkeypatch):
    mgr = MagicMock()
    monkeypatch.setattr("solidedge_mcp.tools.export.export_manager", mgr)
    return mgr


@pytest.fixture
def mock_view(monkeypatch):
    mgr = MagicMock()
    monkeypatch.setattr("solidedge_mcp.tools.export.view_manager", mgr)
    return mgr


# === export_file ===

class TestExportFile:
    @pytest.mark.parametrize("disc, method", [
        ("step", "export_step"),
        ("stl", "export_stl"),
        ("iges", "export_iges"),
        ("pdf", "export_pdf"),
        ("dxf", "export_dxf"),
        ("parasolid", "export_parasolid"),
        ("jt", "export_jt"),
        ("flat_dxf", "export_flat_dxf"),
        ("prc", "export_to_prc"),
        ("plmxml", "export_to_plmxml"),
        ("image", "capture_screenshot"),
    ])
    def test_dispatch(self, mock_export, mock_view, disc, method):
        getattr(mock_export, method).return_value = {"status": "ok"}
        result = export_file(format=disc, file_path="out.file")
        getattr(mock_export, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_plmxml_passes_ini(self, mock_export, mock_view):
        mock_export.export_to_plmxml.return_value = {"status": "ok"}
        export_file(format="plmxml", file_path="out.xml", ini_file_path="cfg.ini")
        mock_export.export_to_plmxml.assert_called_once_with("out.xml", "cfg.ini")

    def test_image_passes_dimensions(self, mock_export, mock_view):
        mock_export.capture_screenshot.return_value = {"status": "ok"}
        export_file(format="image", file_path="out.png", width=1920, height=1080)
        mock_export.capture_screenshot.assert_called_once_with("out.png", 1920, 1080)

    def test_unknown(self, mock_export, mock_view):
        result = export_file(format="bogus")
        assert "error" in result


# === add_drawing_view ===

class TestAddDrawingView:
    @pytest.mark.parametrize("disc, method", [
        ("assembly", "add_assembly_drawing_view"),
        ("assembly_ex", "add_assembly_drawing_view_ex"),
        ("with_config", "add_drawing_view_with_config"),
        ("projected", "add_projected_view"),
        ("detail", "add_detail_view"),
        ("auxiliary", "add_auxiliary_view"),
        ("draft", "add_draft_view"),
        ("by_draft_view", "add_by_draft_view"),
        ("section", "add_section_cut"),
    ])
    def test_dispatch(self, mock_export, mock_view, disc, method):
        getattr(mock_export, method).return_value = {"status": "ok"}
        result = add_drawing_view(type=disc)
        getattr(mock_export, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_export, mock_view):
        result = add_drawing_view(type="bogus")
        assert "error" in result


# === manage_drawing_view ===

class TestManageDrawingView:
    @pytest.mark.parametrize("disc, method", [
        ("get_model_link", "get_drawing_view_model_link"),
        ("show_tangent_edges", "show_tangent_edges"),
        ("set_scale", "set_drawing_view_scale"),
        ("delete", "delete_drawing_view"),
        ("update", "update_drawing_view"),
        ("move", "move_drawing_view"),
        ("show_hidden_edges", "show_hidden_edges"),
        ("set_display_mode", "set_drawing_view_display_mode"),
        ("set_orientation", "set_drawing_view_orientation"),
        ("activate", "activate_drawing_view"),
        ("deactivate", "deactivate_drawing_view"),
        ("get_dimensions", "get_drawing_view_dimensions"),
        ("align", "align_drawing_views"),
        ("update_all", "update_all_views"),
    ])
    def test_dispatch(self, mock_export, mock_view, disc, method):
        getattr(mock_export, method).return_value = {"status": "ok"}
        result = manage_drawing_view(action=disc)
        getattr(mock_export, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_export, mock_view):
        result = manage_drawing_view(action="bogus")
        assert "error" in result


# === add_annotation ===

class TestAddAnnotation:
    @pytest.mark.parametrize("disc, method", [
        ("text_box", "add_text_box"),
        ("leader", "add_leader"),
        ("dimension", "add_dimension"),
        ("balloon", "add_balloon"),
        ("note", "add_note"),
        ("angular_dimension", "add_angular_dimension"),
        ("radial_dimension", "add_radial_dimension"),
        ("diameter_dimension", "add_diameter_dimension"),
        ("ordinate_dimension", "add_ordinate_dimension"),
        ("center_mark", "add_center_mark"),
        ("centerline", "add_centerline"),
        ("surface_finish", "add_surface_finish_symbol"),
        ("weld_symbol", "add_weld_symbol"),
        ("geometric_tolerance", "add_geometric_tolerance"),
    ])
    def test_dispatch(self, mock_export, mock_view, disc, method):
        getattr(mock_export, method).return_value = {"status": "ok"}
        result = add_annotation(type=disc)
        getattr(mock_export, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_export, mock_view):
        result = add_annotation(type="bogus")
        assert "error" in result


# === add_2d_dimension ===

class TestAdd2dDimension:
    @pytest.mark.parametrize("disc, method", [
        ("distance", "add_distance_dimension"),
        ("length", "add_length_dimension"),
        ("radius", "add_radius_dimension_2d"),
        ("angle", "add_angle_dimension_2d"),
    ])
    def test_dispatch(self, mock_export, mock_view, disc, method):
        getattr(mock_export, method).return_value = {"status": "ok"}
        result = add_2d_dimension(type=disc)
        getattr(mock_export, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_export, mock_view):
        result = add_2d_dimension(type="bogus")
        assert "error" in result


# === camera_control (uses view_manager) ===

class TestCameraControl:
    @pytest.mark.parametrize("disc, method", [
        ("set_orientation", "set_view"),
        ("zoom_fit", "zoom_fit"),
        ("zoom_to_selection", "zoom_to_selection"),
        ("rotate", "rotate_camera"),
        ("pan", "pan_camera"),
        ("zoom", "zoom_camera"),
        ("refresh", "refresh_view"),
        ("set_camera", "set_camera"),
        ("begin_dynamics", "begin_camera_dynamics"),
        ("end_dynamics", "end_camera_dynamics"),
    ])
    def test_dispatch(self, mock_export, mock_view, disc, method):
        getattr(mock_view, method).return_value = {"status": "ok"}
        result = camera_control(action=disc)
        getattr(mock_view, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_set_orientation_passes_view(self, mock_export, mock_view):
        mock_view.set_view.return_value = {"status": "ok"}
        camera_control(action="set_orientation", view="Top")
        mock_view.set_view.assert_called_once_with("Top")

    def test_zoom_passes_factor(self, mock_export, mock_view):
        mock_view.zoom_camera.return_value = {"status": "ok"}
        camera_control(action="zoom", factor=2.0)
        mock_view.zoom_camera.assert_called_once_with(2.0)

    def test_unknown(self, mock_export, mock_view):
        result = camera_control(action="bogus")
        assert "error" in result


# === display_control (mixed: view_manager + export_manager) ===

class TestDisplayControl:
    @pytest.mark.parametrize("disc, mgr_attr, method", [
        ("set_mode", "mock_view", "set_display_mode"),
        ("set_background", "mock_view", "set_view_background"),
        ("model_to_screen", "mock_view", "transform_model_to_screen"),
        ("screen_to_model", "mock_view", "transform_screen_to_model"),
        ("set_texture", "mock_export", "set_face_texture"),
    ])
    def test_dispatch(self, mock_export, mock_view, disc, mgr_attr, method, request):
        mgr = mock_view if mgr_attr == "mock_view" else mock_export
        getattr(mgr, method).return_value = {"status": "ok"}
        result = display_control(action=disc)
        getattr(mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_set_texture_passes_args(self, mock_export, mock_view):
        mock_export.set_face_texture.return_value = {"status": "ok"}
        display_control(action="set_texture", face_index=2, texture_name="wood")
        mock_export.set_face_texture.assert_called_once_with(2, "wood")

    def test_unknown(self, mock_export, mock_view):
        result = display_control(action="bogus")
        assert "error" in result


# === manage_sheet ===

class TestManageSheet:
    @pytest.mark.parametrize("disc, method", [
        ("activate", "activate_sheet"),
        ("rename", "rename_sheet"),
        ("delete", "delete_sheet"),
        ("create_drawing", "create_drawing"),
        ("add", "add_draft_sheet"),
    ])
    def test_dispatch(self, mock_export, mock_view, disc, method):
        getattr(mock_export, method).return_value = {"status": "ok"}
        result = manage_sheet(action=disc)
        getattr(mock_export, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_export, mock_view):
        result = manage_sheet(action="bogus")
        assert "error" in result


# === print_control ===

class TestPrintControl:
    @pytest.mark.parametrize("disc, method", [
        ("print", "print_drawing"),
        ("set_printer", "set_printer"),
        ("get_printer", "get_printer"),
        ("set_paper_size", "set_paper_size"),
        ("print_full", "print_document"),
    ])
    def test_dispatch(self, mock_export, mock_view, disc, method):
        getattr(mock_export, method).return_value = {"status": "ok"}
        result = print_control(action=disc)
        getattr(mock_export, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_print_full_empty_printer_becomes_none(self, mock_export, mock_view):
        mock_export.print_document.return_value = {"status": "ok"}
        print_control(action="print_full", printer_name="")
        call_kwargs = mock_export.print_document.call_args[1]
        assert call_kwargs["printer"] is None

    def test_print_full_nonempty_printer(self, mock_export, mock_view):
        mock_export.print_document.return_value = {"status": "ok"}
        print_control(action="print_full", printer_name="HP LaserJet")
        call_kwargs = mock_export.print_document.call_args[1]
        assert call_kwargs["printer"] == "HP LaserJet"

    def test_unknown(self, mock_export, mock_view):
        result = print_control(action="bogus")
        assert "error" in result


# === query_sheet ===

class TestQuerySheet:
    @pytest.mark.parametrize("disc, method", [
        ("dimensions", "get_sheet_dimensions"),
        ("balloons", "get_sheet_balloons"),
        ("text_boxes", "get_sheet_text_boxes"),
        ("drawing_objects", "get_sheet_drawing_objects"),
        ("sections", "get_sheet_sections"),
        ("lines2d", "get_lines2d"),
        ("circles2d", "get_circles2d"),
        ("arcs2d", "get_arcs2d"),
        ("section_cuts", "get_section_cuts"),
    ])
    def test_dispatch(self, mock_export, mock_view, disc, method):
        getattr(mock_export, method).return_value = {"status": "ok"}
        result = query_sheet(type=disc)
        getattr(mock_export, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_export, mock_view):
        result = query_sheet(type="bogus")
        assert "error" in result


# === manage_annotation_data ===

class TestManageAnnotationData:
    @pytest.mark.parametrize("disc, method", [
        ("add_symbol", "add_symbol"),
        ("get_symbols", "get_symbols"),
        ("get_pmi", "get_pmi_info"),
        ("set_pmi_visibility", "set_pmi_visibility"),
    ])
    def test_dispatch(self, mock_export, mock_view, disc, method):
        getattr(mock_export, method).return_value = {"status": "ok"}
        result = manage_annotation_data(action=disc)
        getattr(mock_export, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_export, mock_view):
        result = manage_annotation_data(action="bogus")
        assert "error" in result


# === add_smart_frame ===

class TestAddSmartFrame:
    @pytest.mark.parametrize("disc, method", [
        ("two_point", "add_smart_frame"),
        ("by_origin", "add_smart_frame_by_origin"),
    ])
    def test_dispatch(self, mock_export, mock_view, disc, method):
        getattr(mock_export, method).return_value = {"status": "ok"}
        result = add_smart_frame(method=disc)
        getattr(mock_export, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_export, mock_view):
        result = add_smart_frame(method="bogus")
        assert "error" in result


# === draft_config ===

class TestDraftConfig:
    @pytest.mark.parametrize("disc, method", [
        ("get_global", "get_draft_global_parameter"),
        ("set_global", "set_draft_global_parameter"),
        ("get_origin", "get_symbol_file_origin"),
        ("set_origin", "set_symbol_file_origin"),
    ])
    def test_dispatch(self, mock_export, mock_view, disc, method):
        getattr(mock_export, method).return_value = {"status": "ok"}
        result = draft_config(action=disc)
        getattr(mock_export, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_export, mock_view):
        result = draft_config(action="bogus")
        assert "error" in result


# === create_table ===

class TestCreateTable:
    @pytest.mark.parametrize("disc, method", [
        ("parts_list", "create_parts_list"),
        ("bend", "create_bend_table"),
    ])
    def test_dispatch(self, mock_export, mock_view, disc, method):
        getattr(mock_export, method).return_value = {"status": "ok"}
        result = create_table(type=disc)
        getattr(mock_export, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_export, mock_view):
        result = create_table(type="bogus")
        assert "error" in result
