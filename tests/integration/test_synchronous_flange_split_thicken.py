"""Three creators that were refusals until driven live the right way.

A flange builds only in a synchronous sheet-metal document (the ordered
Flanges.Add* record a feature that never solves); Splits.Add wants the model's
own Body and a reference plane; AddThickenFeature wants the construction body's
faces. Each is judged by the body, not the result dict.
"""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.integration


def _tab(stack):
    stack.sketch.create_sketch("Top")
    stack.sketch.draw_rectangle(0, 0, 0.08, 0.048)
    stack.sketch.close_sketch()
    assert stack.feature.create_base_tab(0.002)["status"] == "created"


def _horizontal_edge(stack):
    """A (face_index, edge_index) pair on a tab edge that runs along X at y=0."""
    doc = stack.doc.get_active_document()
    faces = doc.Models.Item(1).Body.Faces(1)
    for fi in range(1, faces.Count + 1):
        edges = faces.Item(fi).Edges
        for ei in range(1, edges.Count + 1):
            sp = edges.Item(ei).StartVertex.GetPointData([0.0] * 3)
            ep = edges.Item(ei).EndVertex.GetPointData([0.0] * 3)
            sp, ep = (
                [x for p in sp for x in (p if isinstance(p, tuple) else (p,))],
                [x for p in ep for x in (p if isinstance(p, tuple) else (p,))],
            )
            if abs(sp[1]) < 1e-9 and abs(ep[1]) < 1e-9 and abs(sp[2] - ep[2]) < 1e-9:
                return fi - 1, ei - 1
    raise AssertionError("no edge along X at y=0")


class TestSynchronousFlange:
    def test_builds_in_a_synchronous_document(self, stack, new_sheet_metal):
        assert stack.query.set_modeling_mode("synchronous")["status"] == "changed"
        _tab(stack)
        fi, ei = _horizontal_edge(stack)
        before = stack.query.get_face_count()["face_count"]

        r = stack.feature.create_flange_sync(fi, ei, 0.02, inside_radius=0.003, bend_angle=45.0)

        assert r["status"] == "created", r
        assert stack.query.get_face_count()["face_count"] > before
        flange = stack.doc.get_active_document().Models.Item(1).Flanges.Item(1)
        assert flange.BendRadius == pytest.approx(0.003)

    def test_the_basic_method_routes_to_add_sync_in_a_synchronous_document(
        self, stack, new_sheet_metal
    ):
        assert stack.query.set_modeling_mode("synchronous")["status"] == "changed"
        _tab(stack)
        fi, ei = _horizontal_edge(stack)
        before = stack.query.get_face_count()["face_count"]

        r = stack.feature.create_flange(fi, ei, 0.02)

        assert r["status"] == "created", r
        assert stack.query.get_face_count()["face_count"] > before

    def test_an_ordered_document_is_refused_without_a_dead_feature(self, stack, new_sheet_metal):
        _tab(stack)
        fi, ei = _horizontal_edge(stack)

        r = stack.feature.create_flange_sync(fi, ei, 0.02)

        assert r.get("unsupported") is True, r
        assert stack.doc.get_active_document().Models.Item(1).Flanges.Count == 0


class TestSplit:
    def test_a_reference_plane_splits_the_box_into_two_bodies(self, stack, new_part):
        stack.sketch.create_sketch("Top")
        stack.sketch.draw_rectangle(0, 0, 0.08, 0.048)
        stack.sketch.close_sketch()
        assert stack.feature.create_extrude(0.03)["status"] == "created"
        assert stack.feature.create_ref_plane_by_offset(1, 0.015)["status"] == "created"

        r = stack.feature.create_split(plane_index=4)

        assert r["status"] == "created", r
        doc = stack.doc.get_active_document()
        assert doc.Models.Count == 2
        assert doc.Models.Item(1).Splits.Count == 1


class TestThicken:
    @pytest.mark.parametrize("direction", ["Both", "Normal", "Reverse"])
    def test_an_extruded_surface_becomes_a_solid(self, stack, new_part, direction):
        stack.sketch.create_sketch("Top")
        stack.sketch.draw_rectangle(0, 0, 0.08, 0.048)
        stack.sketch.close_sketch()
        assert stack.feature.create_extruded_surface(0.03)["status"] == "created"
        assert stack.doc.get_active_document().Models.Count == 0

        r = stack.feature.thicken_surface(0.002, direction=direction)

        assert r["status"] == "created", r
        assert stack.doc.get_active_document().Models.Count == 1
        assert stack.query.get_face_count()["face_count"] == 12


class TestMirrorSeenByVolume:
    def test_a_mirror_across_the_boxes_own_face_is_real_geometry(self, stack, new_part):
        """Six faces before and after; the volume doubles, and the check must see it."""
        assert stack.query.set_modeling_mode("synchronous")["status"] == "changed"
        stack.sketch.create_sketch("Top")
        stack.sketch.draw_rectangle(0, 0, 0.08, 0.048)
        stack.sketch.close_sketch()
        assert stack.feature.create_extrude(0.03)["status"] == "created"
        name = stack.feature.list_features()["features"][0]["name"]
        volume_before = stack.doc.get_active_document().Models.Item(1).Body.Volume

        r = stack.feature.create_mirror(name, 2)

        assert r["status"] == "created", r
        assert stack.query.get_face_count()["face_count"] == 6
        # re-read: the Body proxy held across the mirror is stale afterwards
        volume_after = stack.doc.get_active_document().Models.Item(1).Body.Volume
        assert volume_after == pytest.approx(2 * volume_before, rel=1e-6)


class TestSlot:
    def test_a_line_path_cuts_a_slot(self, stack, new_sheet_metal):
        _tab(stack)
        stack.sketch.create_sketch("Top")
        stack.sketch.draw_line(0.02, 0.024, 0.06, 0.024)
        stack.sketch.close_sketch(closed=False)
        before = stack.query.get_face_count()["face_count"]

        r = stack.feature.create_slot(0.004, 0.01)

        assert r["status"] == "created", r
        assert stack.query.get_face_count()["face_count"] == before + 4
