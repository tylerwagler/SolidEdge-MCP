"""Dispatch tests for tools/documents.py composite tools."""

from unittest.mock import MagicMock

import pytest

from solidedge_mcp.tools.documents import (
    activate_document,
    close_document,
    create_document,
    import_file,
    open_document,
    save_document,
    undo_redo,
)


@pytest.fixture
def mock_mgr(monkeypatch):
    mgr = MagicMock()
    monkeypatch.setattr("solidedge_mcp.tools.documents.doc_manager", mgr)
    # Bypass path validation so tests with dummy file paths still pass
    monkeypatch.setattr(
        "solidedge_mcp.tools.documents.validate_path",
        lambda p, **kw: (p, None),
    )
    return mgr


# === create_document ===

class TestCreateDocument:
    @pytest.mark.parametrize("disc, method", [
        ("part", "create_part"),
        ("assembly", "create_assembly"),
        ("sheet_metal", "create_sheet_metal"),
        ("draft", "create_draft"),
        ("weldment", "create_weldment"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = create_document(type=disc, template="t.par")
        getattr(mock_mgr, method).assert_called_once_with("t.par")
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = create_document(type="bogus")
        assert "error" in result


# === open_document ===

class TestOpenDocument:
    @pytest.mark.parametrize("disc, method", [
        ("foreground", "open_document"),
        ("background", "open_in_background"),
        ("with_template", "open_with_template"),
        ("dialog", "open_with_file_open_dialog"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = open_document(method=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = open_document(method="bogus")
        assert "error" in result

    def test_foreground_passes_path(self, mock_mgr):
        mock_mgr.open_document.return_value = {"status": "ok"}
        open_document(method="foreground", file_path="test.par")
        mock_mgr.open_document.assert_called_once_with("test.par")

    def test_with_template_passes_args(self, mock_mgr):
        mock_mgr.open_with_template.return_value = {"status": "ok"}
        open_document(method="with_template", file_path="f.par", template="t.par")
        mock_mgr.open_with_template.assert_called_once_with("f.par", "t.par")


# === close_document ===

class TestCloseDocument:
    @pytest.mark.parametrize("disc, method", [
        ("active", "close_document"),
        ("all", "close_all_documents"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = close_document(scope=disc, save=False)
        getattr(mock_mgr, method).assert_called_once_with(False)
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = close_document(scope="bogus")
        assert "error" in result


# === save_document ===

class TestSaveDocument:
    @pytest.mark.parametrize("disc, method", [
        ("save", "save_document"),
        ("copy_as", "save_copy_as"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = save_document(method=disc, file_path="out.par")
        getattr(mock_mgr, method).assert_called_once_with("out.par")
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = save_document(method="bogus")
        assert "error" in result


# === undo_redo ===

class TestUndoRedo:
    @pytest.mark.parametrize("disc, method", [
        ("undo", "undo"),
        ("redo", "redo"),
    ])
    def test_dispatch(self, mock_mgr, disc, method):
        getattr(mock_mgr, method).return_value = {"status": "ok"}
        result = undo_redo(action=disc)
        getattr(mock_mgr, method).assert_called_once()
        assert result == {"status": "ok"}

    def test_unknown(self, mock_mgr):
        result = undo_redo(action="bogus")
        assert "error" in result


# === Standalone tools ===

class TestStandaloneDocuments:
    def test_activate_document(self, mock_mgr):
        mock_mgr.activate_document.return_value = {"status": "ok"}
        result = activate_document("MyDoc")
        mock_mgr.activate_document.assert_called_once_with("MyDoc")
        assert result == {"status": "ok"}

    def test_import_file(self, mock_mgr):
        mock_mgr.import_file.return_value = {"status": "ok"}
        result = import_file("test.step")
        mock_mgr.import_file.assert_called_once_with("test.step")
        assert result == {"status": "ok"}
