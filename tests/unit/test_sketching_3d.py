"""SketchManager.draw_line_3d: one 3D sketch per document until asked for another."""

from __future__ import annotations

from unittest.mock import MagicMock

from solidedge_mcp.backends.sketching import SketchManager


def _manager():
    dm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc
    sketch = doc.Sketches3D.Add.return_value
    counter = {"n": 0}

    def add(*_args):
        counter["n"] += 1
        return MagicMock(name=f"line{counter['n']}")

    sketch.Lines3D.Add.side_effect = add
    type(sketch.Lines3D).Count = property(lambda self: counter["n"])
    return SketchManager(dm), doc, sketch


def test_two_lines_share_one_sketch_and_are_indexed():
    mgr, doc, sketch = _manager()
    first = mgr.draw_line_3d(0, 0, 0, 0.3, 0, 0)
    second = mgr.draw_line_3d(0.3, 0, 0, 0.3, 0.2, 0)
    assert (first["status"], first["index"]) == ("created", 0)
    assert second["index"] == 1
    doc.Sketches3D.Add.assert_called_once()
    sketch.Lines3D.Add.assert_any_call(0, 0, 0, 0.3, 0, 0)
    assert len(mgr.lines_3d) == 2


def test_new_sketch_starts_another():
    mgr, doc, _ = _manager()
    mgr.draw_line_3d(0, 0, 0, 0.3, 0, 0)
    mgr.draw_line_3d(0, 0, 0, 0.3, 0, 0, new_sketch=True)
    assert doc.Sketches3D.Add.call_count == 2


def test_clear_state_forgets_the_lines():
    mgr, _, _ = _manager()
    mgr.draw_line_3d(0, 0, 0, 0.3, 0, 0)
    mgr.clear_state()
    assert mgr.lines_3d == [] and mgr.active_sketch3d is None


def test_a_document_without_3d_sketches_is_refused():
    dm = MagicMock()
    doc = MagicMock(spec=[])  # no Sketches3D
    dm.get_active_document.return_value = doc
    result = SketchManager(dm).draw_line_3d(0, 0, 0, 1, 0, 0)
    assert "Sketches3D" in result["error"]
