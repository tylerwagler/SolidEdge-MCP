"""Tests for cross-cutting infrastructure: error decoding, COM thread, reconnect."""

from __future__ import annotations

import logging
import sys
import threading
from unittest.mock import MagicMock, PropertyMock

import pytest

from solidedge_mcp.backends import errors
from solidedge_mcp.backends.com_thread import ComThread, on_com_thread
from solidedge_mcp.backends.connection import SolidEdgeConnection
from solidedge_mcp.backends.documents import DocumentManager
from solidedge_mcp.backends.logging import ROOT_LOGGER_NAME, configure_logging


def _com_error(hresult: int, description: str | None = None, inner: int = 0):
    """Build a pywintypes.com_error the way pywin32 raises it."""
    import pywintypes

    excepinfo = (0, "SolidEdge", description, None, 0, inner) if description else None
    return pywintypes.com_error(hresult, "Exception occurred.", excepinfo, None)


def _dead_proxy(hresult: int) -> MagicMock:
    """A mock Application whose Version getter raises a COM error."""
    proxy = MagicMock()
    type(proxy).Version = PropertyMock(side_effect=_com_error(hresult))
    return proxy


class TestErrorResult:
    def test_plain_exception_keeps_message(self, monkeypatch):
        monkeypatch.delenv("SOLIDEDGE_MCP_DEBUG", raising=False)
        result = errors.error_result(ValueError("bad input"))
        assert result == {"error": "bad input"}

    def test_traceback_only_in_debug(self, monkeypatch):
        monkeypatch.setenv("SOLIDEDGE_MCP_DEBUG", "1")
        try:
            raise RuntimeError("boom")
        except RuntimeError as e:
            result = errors.error_result(e)
        assert "traceback" in result
        assert "RuntimeError: boom" in result["traceback"]

    def test_com_error_is_decoded(self, monkeypatch):
        monkeypatch.delenv("SOLIDEDGE_MCP_DEBUG", raising=False)
        exc = _com_error(-2147024891)  # 0x80070005
        result = errors.error_result(exc)
        assert result["hresult"] == "0x80070005"
        assert "E_ACCESSDENIED" in result["error"]
        assert "disconnected" not in result

    def test_disp_e_exception_unwraps_inner_code_and_description(self):
        # DISP_E_EXCEPTION wrapping E_INVALIDARG with the server's description
        exc = _com_error(-2147352567, "Profile is not closed", inner=-2147024809)
        assert errors.com_hresult(exc) == 0x80070057
        text = errors.describe_exception(exc)
        assert "E_INVALIDARG" in text
        assert "Profile is not closed" in text

    def test_disconnected_flag(self):
        exc = _com_error(-2147023174)  # RPC_S_SERVER_UNAVAILABLE
        assert errors.is_disconnected_error(exc)
        assert errors.error_result(exc)["disconnected"] is True

    def test_extra_context_merged(self):
        result = errors.error_result(ValueError("x"), note="hint")
        assert result["note"] == "hint"

    def test_context_leads_the_message(self):
        """ "COM error 0x80004005" alone tells a caller nothing.

        The sentence explaining it is the reason a caller passes context, so
        it belongs in the field every consumer reads, not in a key beside it.
        """
        result = errors.error_result(
            ValueError("boom"), context="Mirroring needs a synchronous part"
        )

        assert result["error"] == "Mirroring needs a synchronous part (boom)"
        assert result["com_error"] == "boom"
        assert "context" not in result

    def test_without_context_the_message_is_unchanged(self):
        result = errors.error_result(ValueError("boom"))

        assert result["error"] == "boom"
        assert "com_error" not in result


class TestComThread:
    def test_runs_on_single_dedicated_thread(self):
        ct = ComThread()
        names = set()

        def work():
            names.add(threading.current_thread().name)
            return 42

        try:
            assert all(ct.run(work) == 42 for _ in range(5))
        finally:
            ct.shutdown()
        assert len(names) == 1
        assert next(iter(names)).startswith("solidedge-com")

    def test_nested_call_runs_inline(self):
        ct = ComThread()

        def inner():
            return ct.is_current()

        def outer():
            return ct.run(inner)

        try:
            assert ct.run(outer) is True
        finally:
            ct.shutdown()

    def test_exceptions_propagate(self):
        ct = ComThread()

        def fail():
            raise ValueError("inside")

        try:
            with pytest.raises(ValueError, match="inside"):
                ct.run(fail)
        finally:
            ct.shutdown()

    def test_wrapper_preserves_signature(self):
        def create(method: str = "finite", distance: float = 0.0) -> dict:
            """Doc."""
            return {"m": method, "d": distance}

        wrapped = on_com_thread(create)
        assert wrapped.__name__ == "create"
        assert wrapped.__doc__ == "Doc."
        assert wrapped.__wrapped__ is create
        assert wrapped.__annotations__ == create.__annotations__
        assert wrapped(distance=1.5) == {"m": "finite", "d": 1.5}


class TestLogging:
    def test_logs_go_to_stderr_not_stdout(self):
        root = configure_logging(logging.DEBUG)
        handlers = [h for h in root.handlers if getattr(h, "_solidedge_mcp", False)]
        assert len(handlers) == 1
        assert handlers[0].stream is sys.stderr
        assert root.propagate is False
        assert root.name == ROOT_LOGGER_NAME

    def test_configure_is_idempotent(self):
        before = len(logging.getLogger(ROOT_LOGGER_NAME).handlers)
        configure_logging()
        configure_logging()
        assert len(logging.getLogger(ROOT_LOGGER_NAME).handlers) == before


class TestConnectionReconnect:
    def test_dead_proxy_is_replaced_on_connect(self, monkeypatch):
        conn = SolidEdgeConnection()
        conn.application = _dead_proxy(-2147023174)
        conn._is_connected = True
        disconnected = []
        conn.on_disconnect = lambda: disconnected.append(True)

        fresh = MagicMock()
        fresh.Version = "226"
        monkeypatch.setattr(
            "solidedge_mcp.backends.connection.win32com.client.GetActiveObject",
            lambda progid: fresh,
        )
        result = conn.connect(start_if_needed=False)
        assert result["status"] == "connected"
        assert conn.application is fresh
        assert disconnected == [True]

    def test_check_connection_drops_dead_proxy(self):
        conn = SolidEdgeConnection()
        conn.application = _dead_proxy(-2147417848)
        conn._is_connected = True
        assert conn.check_connection() is False
        assert conn.application is None
        assert conn.is_connected() is False

    def test_check_connection_keeps_live_proxy(self):
        conn = SolidEdgeConnection()
        conn.application = MagicMock(Version="226")
        conn._is_connected = True
        assert conn.check_connection() is True
        assert conn.application is not None


class TestDocumentSwitchDetection:
    def _dm(self):
        conn = MagicMock()
        app = MagicMock()
        conn.get_application.return_value = app
        sketch = MagicMock()
        dm = DocumentManager(conn, sketch)
        return dm, app, sketch

    def test_wires_disconnect_callback(self):
        conn = SolidEdgeConnection()
        sketch = MagicMock()
        dm = DocumentManager(conn, sketch)
        dm.active_document = MagicMock()
        conn.on_disconnect()
        assert dm.active_document is None
        sketch.clear_state.assert_called_once()

    def test_switch_in_ui_clears_sketch_state(self):
        dm, app, sketch = self._dm()
        doc_a = MagicMock(FullName="C:/a.par")
        doc_b = MagicMock(FullName="C:/b.par")
        app.ActiveDocument = doc_a
        assert dm.get_active_document() is doc_a
        sketch.clear_state.assert_not_called()

        app.ActiveDocument = doc_b
        assert dm.get_active_document() is doc_b
        sketch.clear_state.assert_called_once()

    def test_same_document_does_not_clear(self):
        dm, app, sketch = self._dm()
        doc = MagicMock(FullName="C:/a.par")
        app.ActiveDocument = doc
        dm.get_active_document()
        dm.get_active_document()
        sketch.clear_state.assert_not_called()

    def test_falls_back_to_cache_when_app_unreachable(self):
        dm, app, sketch = self._dm()
        cached = MagicMock(FullName="C:/a.par")
        dm.active_document = cached
        dm.connection.get_application.side_effect = Exception("gone")
        assert dm.get_active_document() is cached

    def test_raises_when_nothing_available(self):
        dm, app, sketch = self._dm()
        app.ActiveDocument = None
        with pytest.raises(Exception, match="No active document"):
            dm.get_active_document()


class TestModalDialogSuppression:
    """A modal Solid Edge dialog blocks the COM call that raised it.

    Observed live: export_file on a part with no solid body raised
    "could not be saved because of a file translation error" and the server
    hung until the dialog was dismissed by hand.
    """

    def test_connect_disables_display_alerts(self, monkeypatch):
        conn = SolidEdgeConnection()
        app = MagicMock()
        app.Version = "226"
        monkeypatch.setattr(
            "solidedge_mcp.backends.connection.win32com.client.GetActiveObject",
            lambda progid: app,
        )
        assert conn.connect(start_if_needed=False)["status"] == "connected"
        assert app.DisplayAlerts is False

    def test_connect_survives_a_build_without_display_alerts(self, monkeypatch):
        conn = SolidEdgeConnection()
        app = MagicMock()
        app.Version = "226"
        type(app).DisplayAlerts = PropertyMock(side_effect=AttributeError("no such member"))
        monkeypatch.setattr(
            "solidedge_mcp.backends.connection.win32com.client.GetActiveObject",
            lambda progid: app,
        )
        assert conn.connect(start_if_needed=False)["status"] == "connected"


class TestOverwritePromptRefusal:
    """Saving over an existing file raises a modal prompt DisplayAlerts cannot suppress.

    Observed live: save_document to an existing .par put up "This file exists.
    Do you want to overwrite it?" and the COM call blocked until it was clicked.
    """

    def _manager(self, tmp_path):
        conn = MagicMock()
        dm = DocumentManager(conn)
        dm.active_document = MagicMock()
        dm.active_document.Name = "Part1"
        return dm

    def test_refuses_an_existing_path_by_default(self, tmp_path):
        dm = self._manager(tmp_path)
        target = tmp_path / "part.par"
        target.write_bytes(b"existing")

        result = dm.save_document(str(target))
        assert result["exists"] is True
        assert "overwrite=true" in result["error"]
        dm.active_document.SaveAs.assert_not_called()
        assert target.read_bytes() == b"existing"

    def test_overwrite_replaces_the_file(self, tmp_path):
        dm = self._manager(tmp_path)
        target = tmp_path / "part.par"
        target.write_bytes(b"existing")

        result = dm.save_document(str(target), overwrite=True)
        assert result["status"] == "saved"
        dm.active_document.SaveAs.assert_called_once_with(str(target))
        assert not target.exists()  # removed; Solid Edge writes it via COM

    def test_new_path_saves_without_a_flag(self, tmp_path):
        dm = self._manager(tmp_path)
        target = tmp_path / "fresh.par"

        result = dm.save_document(str(target))
        assert result["status"] == "saved"
        dm.active_document.SaveAs.assert_called_once_with(str(target))


class TestCloseAllGuard:
    """Closing every document without saving destroys other people's work.

    This is not hypothetical: a sweep of mine used scope="all" in a cleanup
    block and closed three unsaved scratch parts that it had not created.
    """

    def _manager(self, names_dirty):
        conn = MagicMock()
        app = MagicMock()
        conn.get_application.return_value = app
        docs = MagicMock()
        docs.Count = len(names_dirty)
        made = []
        for name, dirty in names_dirty:
            d = MagicMock()
            d.Name = name
            d.Dirty = dirty
            made.append(d)
        docs.Item.side_effect = lambda i: made[i - 1]
        app.Documents = docs
        return DocumentManager(conn), made

    def test_refuses_when_any_document_is_unsaved(self):
        dm, docs = self._manager([("Part1", True), ("Part2", False)])

        result = dm.close_all_documents(save=False)
        assert result["unsaved_documents"] == ["Part1"]
        assert "discard_unsaved=true" in result["error"]
        for d in docs:
            d.Close.assert_not_called()

    def test_discard_unsaved_allows_it(self):
        dm, docs = self._manager([("Part1", True)])

        result = dm.close_all_documents(save=False, discard_unsaved=True)
        assert result["closed"] == 1
        docs[0].Close.assert_called_once()

    def test_all_saved_needs_no_flag(self):
        dm, docs = self._manager([("Part1", False), ("Part2", False)])

        result = dm.close_all_documents(save=False)
        assert result["closed"] == 2
