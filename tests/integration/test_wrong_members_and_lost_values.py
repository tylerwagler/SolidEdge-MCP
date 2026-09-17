"""Live checks for the defects a mocked test cannot see.

Each of these passed its unit test while doing nothing, or the wrong thing, in
Solid Edge: the right enum with the wrong member (FaceRotates.Add), a call that
accepts a face it cannot use (DeleteBlends.Add on a plane), a property that is
a member of nothing (Variable.Units), a value that never came back through
IDispatch (GetSymbolFileOrigin), a query that looked in one name space
(Variables.Query with NamedBy=user), and a call that crashes the
process (SweptSurfaces.Add). All verified on Solid Edge 2026.
"""

from __future__ import annotations

import math

import pytest

pytestmark = pytest.mark.integration


def _box(stack, height=0.03):
    stack.sketch.create_sketch("Top")
    stack.sketch.draw_rectangle(0, 0, 0.08, 0.048)
    stack.sketch.close_sketch()
    assert stack.feature.create_extrude(height)["status"] == "created"


def _revolve(stack, angle=90.0):
    stack.sketch.create_sketch("Front")
    stack.sketch.draw_circle(0.03, 0.01, 0.004)
    stack.sketch.set_axis_of_revolution(0.0, 0.0, 0.0, 0.03)
    stack.sketch.close_sketch()
    assert stack.feature.create_revolve(angle=angle)["status"] == "created"


class TestFaceRotates:
    """The old literals (1, 1, .., 2) and (2, 1, .., 0) raised E_INVALIDARG."""

    def test_by_edge_records_a_rotate(self, stack, new_part):
        _box(stack)
        r = stack.feature.create_face_rotate_by_edge(1, 0, 5.0)
        assert r["status"] == "created", r
        assert stack.doc.get_active_document().Models.Item(1).FaceRotates.Count == 1

    def test_by_points_records_a_rotate(self, stack, new_part):
        _box(stack)
        r = stack.feature.create_face_rotate_by_points(0, 0, 1, 5.0)
        assert r["status"] == "created", r
        assert stack.doc.get_active_document().Models.Item(1).FaceRotates.Count == 1


class TestDeleteBlend:
    def test_a_planar_face_is_refused_and_a_cylinder_face_removes_the_round(self, stack, new_part):
        _box(stack)
        assert stack.feature.create_round(0.002)["status"] == "created"
        n = stack.query.get_face_count()["face_count"]
        kinds = {i: stack.query.get_face_geometry(i)["geometry_type"] for i in range(n)}
        planar = next(i for i, k in kinds.items() if k == "Plane")
        cylinder = next(i for i, k in kinds.items() if k == "Cylinder")

        refused = stack.feature.create_delete_blend(planar)
        assert "planar" in refused["error"]
        assert stack.query.get_face_count()["face_count"] == n

        r = stack.feature.create_delete_blend(cylinder)
        assert r["status"] == "created", r
        assert stack.query.get_face_count()["face_count"] < n


class TestVariableUnits:
    def test_the_revolve_angle_is_found_reported_in_degrees_and_set_in_degrees(
        self, stack, new_part
    ):
        _revolve(stack, angle=90.0)

        found = stack.query.query_variables("*")["matches"]
        angle = next(v for v in found if v.get("units") == "angle")
        assert angle["value"] == pytest.approx(math.pi / 2)
        assert angle["value_degrees"] == pytest.approx(90.0)
        # the system-named physical properties are found as well, not only user variables
        assert any(v["name"].startswith("PhysicalProperties_") for v in found)

        r = stack.query.set_variable(angle["name"], 45.0)
        assert r["status"] == "updated", r
        after = stack.query.get_variable(angle["name"])
        assert after["value"] == pytest.approx(math.pi / 4)
        assert after["value_degrees"] == pytest.approx(45.0)

    def test_every_listed_variable_carries_units(self, stack, new_part):
        _revolve(stack)
        listed = stack.query.get_variables()["variables"]
        assert listed and all("units" in v for v in listed), listed


class TestSymbolFileOrigin:
    def test_the_origin_round_trips_and_an_unset_one_is_said_plainly(self, stack, new_draft):
        before = stack.export.get_symbol_file_origin()
        assert "no symbol file origin" in before["error"].lower(), before

        assert stack.export.set_symbol_file_origin(0.1, 0.2)["status"] == "set"

        after = stack.export.get_symbol_file_origin()
        assert after == {"status": "success", "x": pytest.approx(0.1), "y": pytest.approx(0.2)}


class TestSweptSurface:
    def test_refuses_without_the_call(self, stack, new_part):
        """SweptSurfaces.Add crashed Solid Edge 2026 on 3 of 5 calls in this work."""
        stack.sketch.create_sketch("Front")
        stack.sketch.draw_line(0.0, 0.0, 0.0, 0.05)
        stack.sketch.close_sketch(closed=False)
        stack.sketch.create_sketch("Top")
        stack.sketch.draw_circle(0.0, 0.0, 0.005)
        stack.sketch.close_sketch()

        r = stack.feature.create_swept_surface()

        assert r.get("unsupported") is True, r
        assert stack.doc.get_active_document().Constructions.SweptSurfaces.Count == 0
