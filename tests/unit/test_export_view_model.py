"""
Unit tests for ViewModel backend methods.

Tests camera control (get/set/rotate/pan/zoom), view refresh,
camera dynamics, and coordinate transforms.
Uses unittest.mock to simulate COM objects.
"""

from unittest.mock import MagicMock

import pytest


@pytest.fixture
def view_mgr():
    """Create ViewModel with mocked dependencies."""
    from solidedge_mcp.backends.export import ViewModel

    dm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc

    window = MagicMock()
    view_obj = MagicMock()
    window.View = view_obj
    windows = MagicMock()
    windows.Count = 1
    windows.Item.return_value = window
    doc.Windows = windows

    return ViewModel(dm), doc, view_obj


class TestGetCamera:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        # GetCamera returns 11-element tuple
        view_obj.GetCamera.return_value = (
            1.0,
            2.0,
            3.0,  # eye
            0.0,
            0.0,
            0.0,  # target
            0.0,
            1.0,
            0.0,  # up
            False,  # perspective
            1.5,  # scale
        )
        result = vm.get_camera()
        assert result["eye"] == [1.0, 2.0, 3.0]
        assert result["target"] == [0.0, 0.0, 0.0]
        assert result["up"] == [0.0, 1.0, 0.0]
        assert result["perspective"] is False
        assert result["scale_or_angle"] == 1.5
        view_obj.GetCamera.assert_called_once()

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.get_camera()
        assert "error" in result

    def test_perspective_mode(self, view_mgr):
        vm, doc, view_obj = view_mgr
        view_obj.GetCamera.return_value = (
            0.5,
            0.5,
            0.5,
            0.0,
            0.0,
            0.0,
            0.0,
            0.0,
            1.0,
            True,
            0.785,  # ~45 degrees FOV
        )
        result = vm.get_camera()
        assert result["perspective"] is True
        assert result["scale_or_angle"] == 0.785


class TestSetCamera:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.set_camera(1.0, 2.0, 3.0, 0.0, 0.0, 0.0)
        assert result["status"] == "camera_set"
        assert result["eye"] == [1.0, 2.0, 3.0]
        assert result["target"] == [0.0, 0.0, 0.0]
        assert result["up"] == [0.0, 1.0, 0.0]  # defaults
        view_obj.SetCamera.assert_called_once_with(
            1.0, 2.0, 3.0, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0, False, 1.0
        )

    def test_with_perspective(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.set_camera(0.5, 0.5, 0.5, 0.0, 0.0, 0.0, 0.0, 0.0, 1.0, True, 0.785)
        assert result["status"] == "camera_set"
        assert result["perspective"] is True
        assert result["up"] == [0.0, 0.0, 1.0]
        view_obj.SetCamera.assert_called_once_with(
            0.5, 0.5, 0.5, 0.0, 0.0, 0.0, 0.0, 0.0, 1.0, True, 0.785
        )

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.set_camera(0, 0, 1, 0, 0, 0)
        assert "error" in result


class TestRotateCamera:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.rotate_camera(0.5, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0)
        assert result["status"] == "camera_rotated"
        assert result["angle_rad"] == 0.5
        assert result["center"] == [0.0, 0.0, 0.0]
        assert result["axis"] == [0.0, 1.0, 0.0]
        view_obj.RotateCamera.assert_called_once_with(0.5, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0)

    def test_custom_axis(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.rotate_camera(1.57, 0.1, 0.2, 0.3, 1.0, 0.0, 0.0)
        assert result["status"] == "camera_rotated"
        assert result["axis"] == [1.0, 0.0, 0.0]
        view_obj.RotateCamera.assert_called_once_with(1.57, 0.1, 0.2, 0.3, 1.0, 0.0, 0.0)

    def test_defaults(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.rotate_camera(0.3)
        assert result["status"] == "camera_rotated"
        # Default center=(0,0,0), axis=(0,1,0)
        view_obj.RotateCamera.assert_called_once_with(0.3, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0)

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.rotate_camera(0.5)
        assert "error" in result


class TestPanCamera:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.pan_camera(100, -50)
        assert result["status"] == "camera_panned"
        assert result["dx"] == 100
        assert result["dy"] == -50
        view_obj.PanCamera.assert_called_once_with(100, -50)

    def test_zero_pan(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.pan_camera(0, 0)
        assert result["status"] == "camera_panned"
        view_obj.PanCamera.assert_called_once_with(0, 0)

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.pan_camera(10, 20)
        assert "error" in result


class TestZoomCamera:
    def test_zoom_in(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.zoom_camera(2.0)
        assert result["status"] == "camera_zoomed"
        assert result["factor"] == 2.0
        view_obj.ZoomCamera.assert_called_once_with(2.0)

    def test_zoom_out(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.zoom_camera(0.5)
        assert result["status"] == "camera_zoomed"
        assert result["factor"] == 0.5
        view_obj.ZoomCamera.assert_called_once_with(0.5)

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.zoom_camera(1.5)
        assert "error" in result


class TestRefreshView:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.refresh_view()
        assert result["status"] == "view_refreshed"
        view_obj.Update.assert_called_once()

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.refresh_view()
        assert "error" in result


class TestBeginCameraDynamics:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.begin_camera_dynamics()
        assert result["status"] == "camera_dynamics_started"
        view_obj.BeginCameraDynamics.assert_called_once()

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.begin_camera_dynamics()
        assert "error" in result


class TestEndCameraDynamics:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.end_camera_dynamics()
        assert result["status"] == "camera_dynamics_ended"
        view_obj.EndCameraDynamics.assert_called_once()

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.end_camera_dynamics()
        assert "error" in result


# ============================================================================
# VIEW COORDINATE TRANSFORMS
# ============================================================================


class TestTransformModelToScreen:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        view_obj.TransformModelToDC.return_value = (400, 300)
        result = vm.transform_model_to_screen(0.1, 0.2, 0.3)
        assert result["status"] == "success"
        assert result["screen_x"] == 400
        assert result["screen_y"] == 300
        assert result["model"] == [0.1, 0.2, 0.3]
        view_obj.TransformModelToDC.assert_called_once_with(0.1, 0.2, 0.3)

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.transform_model_to_screen(0, 0, 0)
        assert "error" in result


class TestTransformScreenToModel:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        view_obj.TransformDCToModel.return_value = (0.05, 0.10, 0.15)
        result = vm.transform_screen_to_model(500, 250)
        assert result["status"] == "success"
        assert result["x"] == 0.05
        assert result["y"] == 0.10
        assert result["z"] == 0.15
        assert result["screen"] == [500, 250]
        view_obj.TransformDCToModel.assert_called_once_with(500, 250)

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.transform_screen_to_model(0, 0)
        assert "error" in result
