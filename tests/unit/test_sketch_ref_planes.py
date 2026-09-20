"""A sketch needs the document's base planes, whatever that document calls them.

A part, sheet metal or weldment holds them in ``RefPlanes``. An assembly has no
such member at all -- its three are ``AsmRefPlanes``, in the same 1=Top/XY,
2=Right/YZ, 3=Front/XZ order. Verified on Solid Edge 2026.

Reaching only for RefPlanes made every sketch in an assembly fail with a bare
``<unknown>.RefPlanes``. That was not a cosmetic failure: each assembly-level
creator consumes an accumulated profile, so with no way to make one they could
only ever answer "No profiles available" -- all six were unreachable.
"""

from __future__ import annotations

from unittest.mock import MagicMock

from solidedge_mcp.backends.sketching import ref_planes_of


def _without(*missing: str):
    """A document proxy that raises for the named members, as COM does.

    ``del`` on a MagicMock attribute is what makes a later read raise
    AttributeError, which is how a late-bound COM proxy answers a member the
    document does not have -- ``<unknown>.RefPlanes``.
    """
    doc = MagicMock()
    for name in missing:
        delattr(doc, name)
    return doc


class TestRefPlanesOf:
    def test_a_part_uses_ref_planes(self):
        doc = MagicMock()

        assert ref_planes_of(doc) is doc.RefPlanes

    def test_an_assembly_falls_back_to_asm_ref_planes(self):
        doc = _without("RefPlanes")

        assert ref_planes_of(doc) is doc.AsmRefPlanes

    def test_ref_planes_wins_when_both_answer(self):
        """A part is never sent down the assembly path."""
        doc = MagicMock()

        assert ref_planes_of(doc) is doc.RefPlanes
        assert ref_planes_of(doc) is not doc.AsmRefPlanes

    def test_a_document_with_neither_reports_none(self):
        """A draft has no planes to sketch on; the caller gets a real message."""
        doc = _without("RefPlanes", "AsmRefPlanes")

        assert ref_planes_of(doc) is None


class TestSketchingAnAssembly:
    def test_create_sketch_reaches_asm_ref_planes(self):
        from solidedge_mcp.backends.sketching import SketchManager

        doc = _without("RefPlanes")
        doc_manager = MagicMock()
        doc_manager.get_active_document.return_value = doc
        manager = SketchManager(doc_manager)

        result = manager.create_sketch("Top")

        assert "error" not in result, result
        doc.AsmRefPlanes.Item.assert_called_once_with(1)

    def test_a_document_without_planes_says_so_instead_of_raising(self):
        from solidedge_mcp.backends.sketching import SketchManager

        doc = _without("RefPlanes", "AsmRefPlanes")
        doc_manager = MagicMock()
        doc_manager.get_active_document.return_value = doc
        manager = SketchManager(doc_manager)

        result = manager.create_sketch("Top")

        assert "error" in result
        assert "no reference planes" in result["error"]
