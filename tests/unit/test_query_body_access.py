"""A body that cannot be read must say why.

A part built in ordered mode and then switched to synchronous keeps its
geometry on screen, but Models.Item(n).Features goes empty and both Body and
Body.Faces raise a bare E_FAIL. Verified on Solid Edge 2026. Every
measurement then failed with nothing but an HRESULT, which gave no hint that
the modelling mode was the cause.
"""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from solidedge_mcp.backends.query._base import (
    BodyNotReachableError,
    all_faces,
    body_of,
)


@pytest.fixture
def query_mgr():
    from solidedge_mcp.backends.query import QueryManager

    doc_manager = MagicMock()
    doc = MagicMock()
    doc_manager.get_active_document.return_value = doc
    return QueryManager(doc_manager), doc


class TestBodyOf:
    def test_returns_the_body(self):
        model = MagicMock()

        assert body_of(model) is model.Body

    def test_a_body_that_raises_explains_the_mode(self):
        model = MagicMock()
        type(model).Body = property(
            lambda self: (_ for _ in ()).throw(Exception("COM error 0x80004005"))
        )

        with pytest.raises(BodyNotReachableError, match="synchronous"):
            body_of(model)

    def test_a_missing_body_is_the_same_answer(self):
        model = MagicMock()
        model.Body = None

        with pytest.raises(BodyNotReachableError):
            body_of(model)


class TestAllFaces:
    def test_returns_the_collection(self):
        body = MagicMock()
        body.Faces.return_value.Count = 6

        assert all_faces(body) is body.Faces.return_value

    def test_a_lazy_failure_is_still_caught(self):
        """Faces() returns without touching COM; reading Count is what asks."""
        body = MagicMock()
        faces = MagicMock()
        type(faces).Count = property(
            lambda self: (_ for _ in ()).throw(Exception("COM error 0x80004005"))
        )
        body.Faces.return_value = faces

        with pytest.raises(BodyNotReachableError, match="synchronous"):
            all_faces(body)

    def test_a_failing_call_is_caught(self):
        body = MagicMock()
        body.Faces.side_effect = Exception("COM error 0x80004005")

        with pytest.raises(BodyNotReachableError):
            all_faces(body)


class TestFirstModel:
    def test_no_model_names_the_mode_as_a_possible_cause(self, query_mgr):
        manager, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        with pytest.raises(BodyNotReachableError, match="switched to synchronous"):
            manager._get_first_model()

    def test_a_document_without_models_says_it_holds_no_geometry(self, query_mgr):
        manager, doc = query_mgr
        doc.Models = None

        with pytest.raises(BodyNotReachableError, match="no solid"):
            manager._get_first_model()

    def test_a_model_is_returned_when_there_is_one(self, query_mgr):
        manager, doc = query_mgr
        models = MagicMock()
        models.Count = 1
        doc.Models = models

        returned_doc, model = manager._get_first_model()

        assert returned_doc is doc
        assert model is models.Item.return_value


class TestSurfaceAreaReportsIt:
    def test_the_measurement_explains_itself(self, query_mgr):
        manager, doc = query_mgr
        models = MagicMock()
        models.Count = 1
        model = MagicMock()
        models.Item.return_value = model
        doc.Models = models
        type(model).Body = property(
            lambda self: (_ for _ in ()).throw(Exception("COM error 0x80004005"))
        )

        result = manager.get_surface_area()

        assert "error" in result
        assert "synchronous" in result["error"]
        assert "0x80004005" not in result["error"]
