"""Dispatch tests for tools/connection.py composite tools."""

from unittest.mock import MagicMock

import pytest

from solidedge_mcp.tools.connection import (
    app_command,
    app_config,
    arrange_windows,
    convert_by_file_path,
    get_active_command,
    manage_connection,
    run_macro,
)


@pytest.fixture
def mock_mgr(monkeypatch):
    mgr = MagicMock()
    monkeypatch.setattr("solidedge_mcp.tools.connection.connection", mgr)
    return mgr


# === manage_connection ===

class TestManageConnection:
    @pytest.mark.parametrize("disc, method", [
        ("connect", "connect"),
        ("disconnect", "disconnect"),
        ("quit", "quit_application"),
        ("activate", "activate_application"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = manage_connection(action=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_connect_passes_start_if_needed(self, mock_mgr):
        mock_mgr.connect.return_value = {"status": "ok"}
        manage_connection(action="connect", start_if_needed=False)
        mock_mgr.connect.assert_called_once_with(False)

    def test_unknown(self, mock_mgr):
        result = manage_connection(action="bogus")
        assert "error" in result


# === app_command ===

class TestAppCommand:
    @pytest.mark.parametrize("disc, method", [
        ("start", "start_command"),
        ("abort", "abort_command"),
        ("idle", "do_idle"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = app_command(action=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_start_passes_command_id(self, mock_mgr):
        mock_mgr.start_command.return_value = {"status": "ok"}
        app_command(action="start", command_id=42)
        mock_mgr.start_command.assert_called_once_with(42)

    def test_unknown(self, mock_mgr):
        result = app_command(action="bogus")
        assert "error" in result


# === app_config ===

class TestAppConfig:
    @pytest.mark.parametrize("disc, method", [
        ("set_performance", "set_performance_mode"),
        ("get_environment", "get_active_environment"),
        ("get_status_bar", "get_status_bar"),
        ("set_status_bar", "set_status_bar"),
        ("get_visible", "get_visible"),
        ("set_visible", "set_visible"),
        ("get_global", "get_global_parameter"),
        ("set_global", "set_global_parameter"),
        ("get_template", "get_default_template_path"),
        ("set_template", "set_default_template_path"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = app_config(property=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_set_performance_passes_args(self, mock_mgr):
        mock_mgr.set_performance_mode.return_value = {"status": "ok"}
        app_config(
            property="set_performance",
            delay_compute=True,
            screen_updating=False,
            interactive=True,
            display_alerts=False,
        )
        mock_mgr.set_performance_mode.assert_called_once_with(True, False, True, False)

    def test_set_template_passes_args(self, mock_mgr):
        mock_mgr.set_default_template_path.return_value = {"status": "ok"}
        app_config(property="set_template", doc_type=3, template_path="/t.asm")
        mock_mgr.set_default_template_path.assert_called_once_with(3, "/t.asm")

    def test_unknown(self, mock_mgr):
        result = app_config(property="bogus")
        assert "error" in result


# === Standalone tools ===

class TestStandaloneConnection:
    def test_convert_by_file_path(self, mock_mgr):
        mock_mgr.convert_by_file_path.return_value = {"status": "ok"}
        result = convert_by_file_path("in.par", "out.step")
        mock_mgr.convert_by_file_path.assert_called_once_with("in.par", "out.step")
        assert result == {"status": "ok"}

    def test_arrange_windows(self, mock_mgr):
        mock_mgr.arrange_windows.return_value = {"status": "ok"}
        result = arrange_windows(style=2)
        mock_mgr.arrange_windows.assert_called_once_with(2)
        assert result == {"status": "ok"}

    def test_get_active_command(self, mock_mgr):
        mock_mgr.get_active_command.return_value = {"status": "ok"}
        result = get_active_command()
        mock_mgr.get_active_command.assert_called_once()
        assert result == {"status": "ok"}

    def test_run_macro(self, mock_mgr):
        mock_mgr.run_macro.return_value = {"status": "ok"}
        result = run_macro("test.vba")
        mock_mgr.run_macro.assert_called_once_with("test.vba")
        assert result == {"status": "ok"}
