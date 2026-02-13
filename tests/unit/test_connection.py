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
