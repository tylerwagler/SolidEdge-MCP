"""The document fixtures open what they say, and the saved part places.

Every assembly, draft and sheet-metal tool had zero integration coverage
because the only document fixture made parts. These prove the new fixtures
hand a test the right kind of document, and that the saved part on disk can
actually be placed -- the first assembly-level integration test in the suite.

Run with: uv run pytest -m integration
"""

from __future__ import annotations

import pytest

from solidedge_mcp.backends.constants import DocumentTypeConstants

pytestmark = pytest.mark.integration


class TestEachFixtureOpensItsKind:
    def test_new_assembly(self, stack, new_assembly):
        doc = stack.doc.get_active_document()
        assert doc.Type == DocumentTypeConstants.igAssemblyDocument
        assert str(doc.Name) == new_assembly

    def test_new_sheet_metal(self, stack, new_sheet_metal):
        doc = stack.doc.get_active_document()
        assert doc.Type == DocumentTypeConstants.igSheetMetalDocument
        assert str(doc.Name) == new_sheet_metal

    def test_new_draft(self, stack, new_draft):
        doc = stack.doc.get_active_document()
        assert doc.Type == DocumentTypeConstants.igDraftDocument
        assert str(doc.Name) == new_draft

    def test_new_part_still_works_through_the_shared_factory(self, stack, new_part):
        doc = stack.doc.get_active_document()
        assert doc.Type == DocumentTypeConstants.igPartDocument


class TestTheSavedPart:
    def test_it_is_on_disk_and_the_scratch_part_is_closed(self, stack, saved_part):
        """SaveAs renames the scratch part to the file name; it must not be
        left open under that name. The first version of this fixture leaked
        it, and this test only warned instead of failing."""
        from tests.integration.conftest import _open_document_names

        assert saved_part.exists()
        assert saved_part.suffix == ".par"
        open_names = _open_document_names(stack.connection.get_application())
        assert saved_part.name not in open_names, open_names

    def test_it_places_into_an_assembly(self, stack, saved_part, new_assembly):
        placed = stack.assembly.add_component(str(saved_part), 0, 0, 0)
        assert "error" not in placed, placed

        listed = stack.assembly.list_components()
        assert "error" not in listed, listed
        names = [c.get("name", "") for c in listed.get("components", [])]
        assert any("box" in n.lower() for n in names), names
