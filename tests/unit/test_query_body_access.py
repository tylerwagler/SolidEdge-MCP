"""A body that cannot be read must say why, and name the right cause.

A part built in ordered mode and then switched to synchronous keeps its
geometry on screen, but Models.Item(n).Features goes empty and both Body and
Body.Faces raise a bare E_FAIL. Verified on Solid Edge 2026. Every
measurement then failed with nothing but an HRESULT, which gave no hint that
the modelling mode was the cause.

Suppressing every feature empties the body in exactly the same way, and it
wants the opposite advice. Naming the mode there sent the caller to a setting
that could not help, so the feature tree now decides which message to give.
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


class _Feature:
    def __init__(self, suppressed: bool) -> None:
        self.Suppress = suppressed


class _RefPlane:
    """A tree entry with no Suppress member, which is what RefPlane is."""


class _Features:
    def __init__(self, *features: object) -> None:
        self._features = features
        self.Count = len(features)

    def Item(self, index: int) -> object:  # noqa: N802 - COM spelling
        return self._features[index - 1]


class _Model:
    """A model whose Body is gone, with a feature tree that says why."""

    def __init__(self, *features: object) -> None:
        self.Body = None
        self.Document = type("_Doc", (), {"DesignEdgebarFeatures": _Features(*features)})()


class TestWhyTheBodyIsGone:
    """Suppressing every feature empties the body exactly as a mode switch does.

    Both leave Body unreadable, and they want opposite advice. Telling a caller
    to change the modelling mode when the real cause was suppression sends them
    somewhere that cannot help, so the feature tree has to decide between them.
    """

    def test_every_feature_suppressed_says_so(self):
        model = _Model(_Feature(True), _Feature(True))

        with pytest.raises(BodyNotReachableError) as excinfo:
            body_of(model)

        assert "all 2 of its features are suppressed" in str(excinfo.value)
        assert "unsuppress" in str(excinfo.value)
        assert "synchronous" not in str(excinfo.value)

    def test_the_base_reference_planes_do_not_count(self):
        """DesignEdgebarFeatures holds the three base planes, which have no
        Suppress member at all. Asking every entry meant the answer was never
        yes: a part with one suppressed extrusion reads as four entries."""
        model = _Model(_RefPlane(), _RefPlane(), _RefPlane(), _Feature(True))

        with pytest.raises(BodyNotReachableError) as excinfo:
            body_of(model)

        assert "all 1 of its features are suppressed" in str(excinfo.value)

    def test_a_tree_of_nothing_but_planes_falls_back_to_the_mode(self):
        model = _Model(_RefPlane(), _RefPlane(), _RefPlane())

        with pytest.raises(BodyNotReachableError, match="synchronous"):
            body_of(model)

    def test_one_live_feature_falls_back_to_the_mode(self):
        model = _Model(_Feature(True), _Feature(False))

        with pytest.raises(BodyNotReachableError, match="synchronous"):
            body_of(model)

    def test_an_empty_tree_falls_back_to_the_mode(self):
        model = _Model()

        with pytest.raises(BodyNotReachableError, match="synchronous"):
            body_of(model)

    def test_a_tree_that_cannot_be_read_falls_back_to_the_mode(self):
        model = _Model(_Feature(True))
        del model.Document

        with pytest.raises(BodyNotReachableError, match="synchronous"):
            body_of(model)

    def test_all_faces_says_so_too_when_given_the_model(self):
        model = _Model(_Feature(True))
        body = MagicMock()
        body.Faces.side_effect = Exception("COM error 0x80004005")

        with pytest.raises(BodyNotReachableError) as excinfo:
            all_faces(body, model)

        assert "suppressed" in str(excinfo.value)

    def test_all_faces_without_a_model_keeps_the_mode_message(self):
        body = MagicMock()
        body.Faces.side_effect = Exception("COM error 0x80004005")

        with pytest.raises(BodyNotReachableError, match="synchronous"):
            all_faces(body)


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
