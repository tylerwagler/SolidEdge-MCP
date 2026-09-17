"""Drawing views land where they are asked, face the way they are asked, and
report what they hold.

Every view orientation this server made used to be wrong: the constants were
from an invented enum, so every "Front" view was a bottom view. And the parts
list accepted a position it then let Solid Edge choose. Both were found by
reading the value back from the sheet. These pin them.

Run with: uv run pytest -m integration
"""

from __future__ import annotations

import pytest

from solidedge_mcp.backends.constants import ViewOrientationConstants

pytestmark = pytest.mark.integration


@pytest.fixture
def drawing_of_box(stack, saved_part, solid_edge, preexisting_documents):
    """A draft with four views of the saved box; everything closed afterwards.

    create_drawing works from the active document, so the part is opened
    first. Both the part and the draft it makes are scratch, and both are
    closed at teardown by name -- the shared factory's identity check would
    refuse the draft, which it did not create, so this fixture does its own
    bookkeeping over the same guards.
    """
    from tests.integration.conftest import _open_document_names

    app = solid_edge.get_application()
    before = _open_document_names(app)
    opened = stack.doc.open_document(str(saved_part))
    assert "error" not in opened, opened

    drawing = stack.export.create_drawing(views=["Front", "Top", "Right", "Isometric"])
    assert "error" not in drawing, drawing

    yield drawing

    for name in _open_document_names(app) - before - preexisting_documents:
        try:
            doc = next(
                app.Documents.Item(i)
                for i in range(1, app.Documents.Count + 1)
                if str(app.Documents.Item(i).Name) == name
            )
            doc.Close(False)
        except Exception:  # noqa: BLE001 - best effort; the session guard reports leaks
            pass
    stack.doc.active_document = None


class TestCreateDrawing:
    def test_places_every_requested_view(self, drawing_of_box):
        assert drawing_of_box["total_views"] == 4
        assert not drawing_of_box.get("views_failed"), drawing_of_box


class TestViewOrientationReadsBackWhatWasAsked:
    """Front used to read back as igBottomView; the enum was invented."""

    @pytest.mark.parametrize(
        ("orientation", "want"),
        [
            ("Front", ViewOrientationConstants.igFrontView),
            ("Top", ViewOrientationConstants.igTopView),
            ("Isometric", ViewOrientationConstants.igTopFrontRightView),
        ],
    )
    def test_orientation(self, stack, drawing_of_box, orientation, want):
        r = stack.export.set_drawing_view_orientation(0, orientation)

        assert "error" not in r, r
        assert r["reads_back"] == want, f"{orientation} read back {r['reads_back']}, want {want}"


class TestViewScaleReadsBack:
    """The setter used to echo the value it was asked for; it now reads
    ScaleFactor back, so a write Solid Edge ignored would show."""

    def test_a_quarter_scale(self, stack, drawing_of_box):
        r = stack.export.set_drawing_view_scale(0, 0.25)

        assert "error" not in r, r
        assert r["reads_back"] == pytest.approx(0.25, rel=1e-9), r


class TestThePartsListIsPositioned:
    """The tool accepted x/y and let Solid Edge choose the position."""

    def test_where_it_was_asked(self, stack, drawing_of_box):
        r = stack.export.create_parts_list(auto_balloon=True, x=0.12, y=0.22)

        assert "error" not in r, r
        assert r.get("positioned") is True, r
        assert r.get("position") == pytest.approx([0.12, 0.22], rel=1e-9)
