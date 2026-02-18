"""
Unit tests for SolidEdgeConnection backend methods.

Tests application management, performance flags, and new property accessors.
Uses unittest.mock to simulate COM objects.
"""

from unittest.mock import MagicMock

import pytest

from solidedge_mcp.backends.connection import SolidEdgeConnection


@pytest.fixture
def conn():
    """Create a connected SolidEdgeConnection with mocked app."""
    c = SolidEdgeConnection()
    c.application = MagicMock()
    c._is_connected = True
    return c


# ============================================================================
# ACTIVATE APPLICATION
# ============================================================================


class TestActivateApplication:
    def test_success(self, conn):
        result = conn.activate_application()
        assert result["status"] == "activated"
        conn.application.Activate.assert_called_once()

    def test_not_connected(self):
        c = SolidEdgeConnection()
        result = c.activate_application()
        assert "error" in result


# ============================================================================
# ABORT COMMAND
# ============================================================================


class TestAbortCommand:
    def test_abort_all(self, conn):
        result = conn.abort_command(abort_all=True)
        assert result["status"] == "aborted"
        assert result["abort_all"] is True
        conn.application.AbortCommand.assert_called_once_with(True)

    def test_abort_single(self, conn):
        result = conn.abort_command(abort_all=False)
        assert result["status"] == "aborted"
        assert result["abort_all"] is False
        conn.application.AbortCommand.assert_called_once_with(False)

    def test_not_connected(self):
        c = SolidEdgeConnection()
        result = c.abort_command()
        assert "error" in result


# ============================================================================
# GET ACTIVE ENVIRONMENT
# ============================================================================


class TestGetActiveEnvironment:
    def test_success(self, conn):
        env = MagicMock()
        env.Name = "Part"
        env.Caption = "Part Environment"
        conn.application.ActiveEnvironment = env

        result = conn.get_active_environment()
        assert result["status"] == "success"
        assert result["name"] == "Part"
        assert result["caption"] == "Part Environment"

    def test_not_connected(self):
        c = SolidEdgeConnection()
        result = c.get_active_environment()
        assert "error" in result


# ============================================================================
# STATUS BAR
# ============================================================================


class TestGetStatusBar:
    def test_success(self, conn):
        conn.application.StatusBar = "Ready"
        result = conn.get_status_bar()
        assert result["status"] == "success"
        assert result["text"] == "Ready"

    def test_not_connected(self):
        c = SolidEdgeConnection()
        result = c.get_status_bar()
        assert "error" in result


class TestSetStatusBar:
    def test_success(self, conn):
        result = conn.set_status_bar("Processing...")
        assert result["status"] == "set"
        assert result["text"] == "Processing..."
        assert conn.application.StatusBar == "Processing..."

    def test_not_connected(self):
        c = SolidEdgeConnection()
        result = c.set_status_bar("test")
        assert "error" in result


# ============================================================================
# VISIBLE
# ============================================================================


class TestGetVisible:
    def test_visible(self, conn):
        conn.application.Visible = True
        result = conn.get_visible()
        assert result["status"] == "success"
        assert result["visible"] is True

    def test_hidden(self, conn):
        conn.application.Visible = False
        result = conn.get_visible()
        assert result["status"] == "success"
        assert result["visible"] is False

    def test_not_connected(self):
        c = SolidEdgeConnection()
        result = c.get_visible()
        assert "error" in result


class TestSetVisible:
    def test_show(self, conn):
        result = conn.set_visible(True)
        assert result["status"] == "set"
        assert result["visible"] is True
        assert conn.application.Visible is True

    def test_hide(self, conn):
        result = conn.set_visible(False)
        assert result["status"] == "set"
        assert result["visible"] is False
        assert conn.application.Visible is False

    def test_not_connected(self):
        c = SolidEdgeConnection()
        result = c.set_visible(True)
        assert "error" in result


# ============================================================================
# GET GLOBAL PARAMETER
# ============================================================================


class TestGetGlobalParameter:
    def test_success(self, conn):
        conn.application.GetGlobalParameter.return_value = 0.005
        result = conn.get_global_parameter(1)
        assert result["status"] == "success"
        assert result["parameter"] == 1
        assert result["value"] == 0.005
        conn.application.GetGlobalParameter.assert_called_once_with(1)

    def test_not_connected(self):
        c = SolidEdgeConnection()
        result = c.get_global_parameter(1)
        assert "error" in result

    def test_invalid_parameter(self, conn):
        conn.application.GetGlobalParameter.side_effect = Exception("Invalid param")
        result = conn.get_global_parameter(999)
        assert "error" in result


# ============================================================================
# SET GLOBAL PARAMETER
# ============================================================================


class TestSetGlobalParameter:
    def test_success(self, conn):
        result = conn.set_global_parameter(10, 0.01)
        assert result["status"] == "set"
        assert result["parameter"] == 10
        assert result["value"] == 0.01
        conn.application.SetGlobalParameter.assert_called_once_with(10, 0.01)

    def test_not_connected(self):
        c = SolidEdgeConnection()
        result = c.set_global_parameter(1, 0.005)
        assert "error" in result

    def test_com_error(self, conn):
        conn.application.SetGlobalParameter.side_effect = Exception("COM error")
        result = conn.set_global_parameter(1, 0.005)
        assert "error" in result


# ============================================================================
# CONVERT BY FILE PATH
# ============================================================================


class TestConvertByFilePath:
    def test_success(self, conn):
        result = conn.convert_by_file_path("C:/input/part.par", "C:/output/part.step")
        assert result["status"] == "converted"
        assert result["input"] == "C:/input/part.par"
        assert result["output"] == "C:/output/part.step"
        conn.application.ConvertByFilePath.assert_called_once_with(
            "C:/input/part.par", "C:/output/part.step"
        )

    def test_not_connected(self):
        c = SolidEdgeConnection()
        result = c.convert_by_file_path("C:/in.par", "C:/out.step")
        assert "error" in result

    def test_com_error(self, conn):
        conn.application.ConvertByFilePath.side_effect = Exception("Conversion failed")
        result = conn.convert_by_file_path("C:/in.par", "C:/out.step")
        assert "error" in result


# ============================================================================
# GET DEFAULT TEMPLATE PATH
# ============================================================================


class TestGetDefaultTemplatePath:
    def test_success(self, conn):
        conn.application.GetDefaultTemplatePath.return_value = "C:/templates/part.par"
        result = conn.get_default_template_path(1)
        assert result["status"] == "success"
        assert result["doc_type"] == 1
        assert result["template_path"] == "C:/templates/part.par"
        conn.application.GetDefaultTemplatePath.assert_called_once_with(1)

    def test_not_connected(self):
        c = SolidEdgeConnection()
        result = c.get_default_template_path(1)
        assert "error" in result

    def test_com_error(self, conn):
        conn.application.GetDefaultTemplatePath.side_effect = Exception("Invalid type")
        result = conn.get_default_template_path(99)
        assert "error" in result


# ============================================================================
# SET DEFAULT TEMPLATE PATH
# ============================================================================


class TestSetDefaultTemplatePath:
    def test_success(self, conn):
        result = conn.set_default_template_path(1, "C:/templates/custom.par")
        assert result["status"] == "set"
        assert result["doc_type"] == 1
        assert result["template_path"] == "C:/templates/custom.par"
        conn.application.SetDefaultTemplatePath.assert_called_once_with(
            1, "C:/templates/custom.par"
        )

    def test_not_connected(self):
        c = SolidEdgeConnection()
        result = c.set_default_template_path(1, "C:/templates/custom.par")
        assert "error" in result

    def test_com_error(self, conn):
        conn.application.SetDefaultTemplatePath.side_effect = Exception("Access denied")
        result = conn.set_default_template_path(1, "C:/bad/path.par")
        assert "error" in result


# ============================================================================
# ARRANGE WINDOWS
# ============================================================================


class TestArrangeWindows:
    def test_tiled(self, conn):
        result = conn.arrange_windows(1)
        assert result["status"] == "arranged"
        assert result["style"] == 1
        assert result["style_name"] == "Tiled"
        conn.application.ArrangeWindows.assert_called_once_with(1)

    def test_cascade(self, conn):
        result = conn.arrange_windows(8)
        assert result["status"] == "arranged"
        assert result["style"] == 8
        assert result["style_name"] == "Cascade"
        conn.application.ArrangeWindows.assert_called_once_with(8)

    def test_default_style(self, conn):
        result = conn.arrange_windows()
        assert result["status"] == "arranged"
        assert result["style"] == 1
        conn.application.ArrangeWindows.assert_called_once_with(1)

    def test_not_connected(self):
        c = SolidEdgeConnection()
        result = c.arrange_windows()
        assert "error" in result


# ============================================================================
# GET ACTIVE COMMAND
# ============================================================================


class TestGetActiveCommand:
    def test_with_command(self, conn):
        cmd = MagicMock()
        cmd.Name = "ExtrudeProtrusion"
        cmd.ID = 42
        conn.application.ActiveCommand = cmd

        result = conn.get_active_command()
        assert result["status"] == "success"
        assert result["has_active_command"] is True
        assert result["name"] == "ExtrudeProtrusion"
        assert result["id"] == 42

    def test_no_command(self, conn):
        conn.application.ActiveCommand = None

        result = conn.get_active_command()
        assert result["status"] == "success"
        assert result["has_active_command"] is False

    def test_not_connected(self):
        c = SolidEdgeConnection()
        result = c.get_active_command()
        assert "error" in result


# ============================================================================
# RUN MACRO
# ============================================================================


class TestRunMacro:
    def test_success(self, conn):
        result = conn.run_macro("C:/macros/my_script.bas")
        assert result["status"] == "executed"
        assert result["filename"] == "C:/macros/my_script.bas"
        conn.application.RunMacro.assert_called_once_with("C:/macros/my_script.bas")

    def test_not_connected(self):
        c = SolidEdgeConnection()
        result = c.run_macro("C:/macros/my_script.bas")
        assert "error" in result

    def test_com_error(self, conn):
        conn.application.RunMacro.side_effect = Exception("Macro not found")
        result = conn.run_macro("C:/macros/nonexistent.bas")
        assert "error" in result
