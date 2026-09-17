"""A creator that builds no solid must still prove it built *something*.

Reference planes, construction surfaces, sketches and documents change no
face count, so ``verifies_geometry`` has nothing to say about them and 49
creators reported success unverified. Each lands in a COM collection whose
``Count`` is readable, and that is the signal here.

Fakes are hand-written rather than MagicMock on purpose: a MagicMock answers
``.Count`` with something truthy that is not an int, and has already talked a
check into claiming something Solid Edge never said. The decorator's one rule
is that it never judges on a number it did not really read.
"""

from __future__ import annotations

from unittest.mock import MagicMock

from solidedge_mcp.backends.features._base import (
    collection_count,
    verifies_collection_growth,
    verify_collection_growth_on_creators,
)
from solidedge_mcp.backends.features._ref_planes import RefPlaneMixin


class _Counter:
    def __init__(self, count: int) -> None:
        self.Count = count


class _Constructions:
    def __init__(self, extruded: int = 0, revolved: int = 0) -> None:
        self.ExtrudedSurfaces = _Counter(extruded)
        self.RevolvedSurfaces = _Counter(revolved)


class _Doc:
    """A part with a readable plane count and two surface collections."""

    def __init__(self, planes: int = 3) -> None:
        self.RefPlanes = _Counter(planes)
        self.Constructions = _Constructions()


class _Mgr:
    def __init__(self, doc: object) -> None:
        self.doc_manager = MagicMock()
        self.doc_manager.get_active_document.return_value = doc
        self.doc = doc

    @verifies_collection_growth("RefPlanes")
    def create_plane(self, grows: bool = True) -> dict:
        if grows:
            self.doc.RefPlanes.Count += 1  # type: ignore[attr-defined]
        return {"status": "created", "type": "ref_plane"}

    @verifies_collection_growth("RefPlanes")
    def create_failing(self) -> dict:
        return {"error": "COM said no"}

    @verifies_collection_growth("Constructions.ExtrudedSurfaces", "Constructions.RevolvedSurfaces")
    def create_surface(self, into: str | None) -> dict:
        if into is not None:
            getattr(self.doc.Constructions, into).Count += 1  # type: ignore[attr-defined]
        return {"status": "created", "type": "surface"}


class TestGrowthIsSuccess:
    def test_a_grown_collection_passes_the_result_through(self):
        assert _Mgr(_Doc()).create_plane() == {"status": "created", "type": "ref_plane"}

    def test_an_unchanged_collection_becomes_an_error(self):
        result = _Mgr(_Doc(planes=3)).create_plane(grows=False)

        assert "error" in result
        assert "added nothing" in result["error"]
        assert "RefPlanes" in result["error"]
        assert result["count_before"] == 3
        assert result["count_after"] == 3
        assert result["attempted"] == "ref_plane"

    def test_an_error_result_is_left_alone(self):
        assert _Mgr(_Doc()).create_failing() == {"error": "COM said no"}


class TestItNeverJudgesWhatItCannotRead:
    def test_a_missing_collection_passes_through(self):
        """A draft has no RefPlanes; that is not a failed creator."""
        doc = _Doc()
        del doc.RefPlanes

        result = _Mgr(doc).create_plane(grows=False)

        assert result == {"status": "created", "type": "ref_plane"}

    def test_a_count_that_is_not_an_int_passes_through(self):
        doc = _Doc()
        doc.RefPlanes.Count = MagicMock()  # truthy, not a number

        result = _Mgr(doc).create_plane(grows=False)

        assert result == {"status": "created", "type": "ref_plane"}

    def test_a_mocked_document_passes_through(self):
        """Every other unit test drives these with MagicMock; stay inert there."""
        result = _Mgr(MagicMock()).create_plane(grows=False)

        assert result == {"status": "created", "type": "ref_plane"}

    def test_a_document_that_cannot_be_fetched_passes_through(self):
        mgr = _Mgr(_Doc())
        mgr.doc_manager.get_active_document.side_effect = RuntimeError("gone")

        assert mgr.create_plane(grows=False) == {"status": "created", "type": "ref_plane"}


class TestAGroupIsSummed:
    """A surface lands in whichever collection suits it; any of them growing counts."""

    def test_growth_in_the_second_collection_is_growth(self):
        assert _Mgr(_Doc()).create_surface(into="RevolvedSurfaces")["status"] == "created"

    def test_growth_in_the_first_collection_is_growth(self):
        assert _Mgr(_Doc()).create_surface(into="ExtrudedSurfaces")["status"] == "created"

    def test_no_growth_anywhere_is_an_error(self):
        result = _Mgr(_Doc()).create_surface(into=None)

        assert "error" in result
        assert "ExtrudedSurfaces + Constructions.RevolvedSurfaces" in result["error"]

    def test_one_missing_member_of_the_group_withholds_judgement(self):
        doc = _Doc()
        del doc.Constructions.RevolvedSurfaces

        assert _Mgr(doc).create_surface(into=None)["status"] == "created"


class TestTheApplicationRoot:
    def test_documents_are_counted_on_the_application(self):
        class _App:
            Documents = _Counter(2)

        mgr = _Mgr(_Doc())
        mgr.doc_manager.connection.get_application.return_value = _App()

        assert collection_count("Documents", root="application")(mgr) == 2

    def test_a_mocked_application_withholds_judgement(self):
        mgr = _Mgr(_Doc())  # doc_manager.connection is a MagicMock

        assert collection_count("Documents", root="application")(mgr) is None


class TestTheClassDecorator:
    def test_wraps_only_creators(self):
        @verify_collection_growth_on_creators("RefPlanes")
        class Mixin:
            def create_a(self):
                return {"status": "created"}

            def list_things(self):
                return []

        assert hasattr(Mixin.create_a, "__wrapped__")
        assert not hasattr(Mixin.list_things, "__wrapped__")

    def test_every_ref_plane_creator_is_wrapped(self):
        creators = [n for n in vars(RefPlaneMixin) if n.startswith("create_")]

        assert creators, "RefPlaneMixin has no creators?"
        for name in creators:
            assert hasattr(getattr(RefPlaneMixin, name), "__wrapped__"), name
