"""Dispatch tests for tools/connection.py composite tools."""

from unittest.mock import MagicMock

import pytest

import solidedge_mcp.tools.connection as connection_tools
from solidedge_mcp.tools.connection import (
    app_command,
    app_config,
    arrange_windows,
    convert_by_file_path,
    get_active_command,
    manage_connection,
    run_macro,
)
from tests.unit.test_tools_query import assert_literal_discriminators


@pytest.fixture
def mock_mgr(monkeypatch):
    mgr = MagicMock()
    monkeypatch.setattr("solidedge_mcp.tools.connection.connection", mgr)
    return mgr


# === manage_connection ===


class TestManageConnection:
    @pytest.mark.parametrize(
        "disc, method",
        [
            ("connect", "connect"),
            ("disconnect", "disconnect"),
            ("quit", "quit_application"),
            ("activate", "activate_application"),
        ],
    )
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = manage_connection(action=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_connect_passes_start_if_needed(self, mock_mgr):
        mock_mgr.connect.return_value = {"status": "ok"}
        manage_connection(action="connect", start_if_needed=False)
        mock_mgr.connect.assert_called_once_with(start_if_needed=False)

    def test_unknown(self, mock_mgr):
        result = manage_connection(action="bogus")
        assert "error" in result


# === app_command ===


class TestAppCommand:
    @pytest.mark.parametrize(
        "disc, method",
        [
            ("start", "start_command"),
            ("abort", "abort_command"),
            ("idle", "do_idle"),
        ],
    )
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = app_command(action=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_start_passes_command_id(self, mock_mgr):
        mock_mgr.start_command.return_value = {"status": "ok"}
        app_command(action="start", command_id=42)
        mock_mgr.start_command.assert_called_once_with(command_id=42)

    def test_unknown(self, mock_mgr):
        result = app_command(action="bogus")
        assert "error" in result


# === app_config ===


class TestAppConfig:
    @pytest.mark.parametrize(
        "disc, method",
        [
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
        ],
    )
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
        mock_mgr.set_performance_mode.assert_called_once_with(
            delay_compute=True, screen_updating=False, interactive=True, display_alerts=False
        )

    def test_set_template_passes_args(self, mock_mgr):
        mock_mgr.set_default_template_path.return_value = {"status": "ok"}
        app_config(property="set_template", doc_type=3, template_path="/t.asm")
        mock_mgr.set_default_template_path.assert_called_once_with(
            doc_type=3, template_path="/t.asm"
        )

    def test_unknown(self, mock_mgr):
        result = app_config(property="bogus")
        assert "error" in result


# === Standalone tools ===


class TestStandaloneConnection:
    def test_convert_by_file_path(self, mock_mgr, tmp_path):
        """The input must exist: Solid Edge answers a missing one with a modal."""
        mock_mgr.convert_by_file_path.return_value = {"status": "ok"}
        source = tmp_path / "in.par"
        source.write_text("x", encoding="utf-8")
        target = tmp_path / "out.step"

        result = convert_by_file_path(str(source), str(target))

        mock_mgr.convert_by_file_path.assert_called_once_with(
            input_path=str(source), output_path=str(target)
        )
        assert result == {"status": "ok"}

    def test_convert_refuses_a_missing_input(self, mock_mgr, tmp_path):
        result = convert_by_file_path(str(tmp_path / "gone.par"), str(tmp_path / "out.step"))

        assert "error" in result
        mock_mgr.convert_by_file_path.assert_not_called()

    def test_convert_refuses_to_replace_the_output(self, mock_mgr, tmp_path):
        source = tmp_path / "in.par"
        source.write_text("x", encoding="utf-8")
        target = tmp_path / "out.step"
        target.write_text("old", encoding="utf-8")

        result = convert_by_file_path(str(source), str(target))

        assert "already exists" in result["error"]
        mock_mgr.convert_by_file_path.assert_not_called()
        assert target.exists()

    def test_arrange_windows(self, mock_mgr):
        mock_mgr.arrange_windows.return_value = {"status": "ok"}
        result = arrange_windows(style=2)
        mock_mgr.arrange_windows.assert_called_once_with(style=2)
        assert result == {"status": "ok"}

    def test_get_active_command(self, mock_mgr):
        mock_mgr.get_active_command.return_value = {"status": "ok"}
        result = get_active_command()
        mock_mgr.get_active_command.assert_called_once()
        assert result == {"status": "ok"}

    def test_run_macro(self, mock_mgr, tmp_path):
        mock_mgr.run_macro.return_value = {"status": "ok"}
        macro = tmp_path / "test.vba"
        macro.write_text("x", encoding="utf-8")

        result = run_macro(str(macro))

        mock_mgr.run_macro.assert_called_once_with(filename=str(macro))
        assert result == {"status": "ok"}

    def test_run_macro_refuses_a_missing_file(self, mock_mgr, tmp_path):
        """A missing macro raises a modal dialog titled with the path.

        Reproduced on Solid Edge 2026: the dialog blocked the single UI thread
        and the whole server stopped answering.
        """
        result = run_macro(str(tmp_path / "no_such_macro.bas"))

        assert "error" in result
        mock_mgr.run_macro.assert_not_called()


# === Literal discriminator drift ===


def test_connection_discriminators_match_their_cases():
    assert assert_literal_discriminators(connection_tools) == 3
