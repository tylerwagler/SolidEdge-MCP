"""3D sketch lines, and a structural frame along them, on real Solid Edge."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.integration


class TestDraw3dLine:
    def test_a_part_takes_a_3d_line(self, stack, new_part):
        r = stack.sketch.draw_line_3d(0.0, 0.0, 0.0, 0.1, 0.0, 0.05)
        assert r["status"] == "created", r
        assert r["index"] == 0
        assert r["length"] == pytest.approx((0.1**2 + 0.05**2) ** 0.5)
        assert stack.doc.get_active_document().Sketches3D.Count == 1


class TestStructuralFrame:
    def test_a_frame_runs_along_two_lines(self, stack, new_assembly, saved_part):
        doc = stack.doc.get_active_document()
        assert stack.sketch.draw_line_3d(0.0, 0.0, 0.0, 0.3, 0.0, 0.0)["status"] == "created"
        assert stack.sketch.draw_line_3d(0.3, 0.0, 0.0, 0.3, 0.2, 0.0)["status"] == "created"
        occurrences_before = doc.Occurrences.Count

        r = stack.assembly.add_structural_frame(str(saved_part), [0, 1])

        assert r["status"] == "created", r
        assert doc.StructuralFrames.Count == 1
        assert doc.Occurrences.Count > occurrences_before

    def test_by_orientation_runs_along_a_line(self, stack, new_assembly, saved_part):
        doc = stack.doc.get_active_document()
        assert stack.sketch.draw_line_3d(0.0, 0.0, 0.0, 0.3, 0.0, 0.0)["status"] == "created"

        r = stack.assembly.add_structural_frame_by_orientation(str(saved_part), "", [0])

        assert r["status"] == "created", r
        assert doc.StructuralFrames.Count == 1


class TestRelations:
    def test_a_planar_mate_between_two_placed_boxes(self, stack, new_assembly, saved_part):
        doc = stack.doc.get_active_document()
        doc.Occurrences.AddByFilename(str(saved_part))
        second = doc.Occurrences.AddByFilename(str(saved_part))
        second.Move(0.2, 0.0, 0.0)
        before = doc.Relations3d.Count

        r = stack.assembly.add_planar_relation(0, 1, orientation="Antialign")

        assert r["status"] == "created", r
        assert doc.Relations3d.Count == before + 1

    def test_a_mate_constraint_takes_the_same_route(self, stack, new_assembly, saved_part):
        doc = stack.doc.get_active_document()
        doc.Occurrences.AddByFilename(str(saved_part))
        doc.Occurrences.AddByFilename(str(saved_part)).Move(0.2, 0.0, 0.0)
        before = doc.Relations3d.Count

        r = stack.assembly.create_mate("Mate", 0, 1)

        assert r["status"] == "created", r
        assert doc.Relations3d.Count == before + 1
