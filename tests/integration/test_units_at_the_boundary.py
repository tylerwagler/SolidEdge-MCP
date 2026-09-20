"""Meters and degrees at the tool boundary -- checked against live values.

Each of these was a radians-or-worse defect found by reading the value back
from Solid Edge rather than trusting the result dict:

    GetDirection{1,2}Extent returns (ExtentType, ExtentSide, value); the
    readers took the side constant as the distance and stringified the real
    value into a face_ref that does not exist. On a revolve the value is an
    angle in radians.
    Cone.GetConeData's HalfAngle is radians and was reported unlabelled.
    GetCamera's ScaleOrAngle is a scale in orthographic and a field-of-view
    angle in radians in perspective; both directions passed it raw.

Run with: uv run pytest -m integration
"""

from __future__ import annotations

import math

import pytest

pytestmark = pytest.mark.integration
EXACT = dict(rel=1e-6, abs=1e-9)


def _box(stack):
    stack.sketch.create_sketch("Top")
    stack.sketch.draw_rectangle(0, 0, 0.08, 0.048)
    stack.sketch.close_sketch()
    assert stack.feature.create_extrude(0.03)["status"] == "created"


def _revolve(stack, angle: float):
    stack.sketch.create_sketch("Front")
    stack.sketch.draw_circle(0.03, 0.01, 0.004)
    stack.sketch.set_axis_of_revolution(0.0, 0.0, 0.0, 0.03)
    stack.sketch.close_sketch()
    assert stack.feature.create_revolve(angle=angle)["status"] == "created"


def _first_feature_name(stack) -> str:
    return stack.feature.list_features()["features"][0]["name"]


class TestExtentSlotsAreNamedCorrectly:
    def test_an_extrude_reports_its_distance_and_side(self, stack, new_part):
        _box(stack)

        r = stack.query.get_direction1_extent(_first_feature_name(stack))

        assert r["distance"] == pytest.approx(0.03, **EXACT)
        assert r["extent_side"] == 2, "igRight"
        assert "face_ref" not in r, "there is no such out-param"

    def test_a_revolve_reports_its_angle_in_degrees(self, stack, new_part):
        _revolve(stack, 90.0)

        r = stack.query.get_direction1_extent(_first_feature_name(stack))

        assert r["angle_degrees"] == pytest.approx(90.0, **EXACT)
        assert r["angle_radians"] == pytest.approx(math.pi / 2, **EXACT)
        assert "distance" not in r

    def test_setting_a_revolve_extent_takes_degrees(self, stack, new_part):
        _revolve(stack, 90.0)
        name = _first_feature_name(stack)

        updated = stack.query.set_direction1_extent(name, 13, 45.0)
        assert "error" not in updated, updated

        assert stack.query.get_direction1_extent(name)["angle_degrees"] == pytest.approx(
            45.0, **EXACT
        )


class TestAConeReportsItsHalfAngleInDegrees:
    def test_a_revolved_triangle(self, stack, new_part):
        stack.sketch.create_sketch("Front")
        stack.sketch.draw_line(0.0, 0.0, 0.02, 0.0)
        stack.sketch.draw_line(0.02, 0.0, 0.0, 0.04)
        stack.sketch.draw_line(0.0, 0.04, 0.0, 0.0)
        stack.sketch.set_axis_of_revolution(0.0, 0.0, 0.0, 0.04)
        stack.sketch.close_sketch()
        assert stack.feature.create_revolve(angle=360.0)["status"] == "created"

        faces = range(int(stack.query.get_face_count()["face_count"]))
        cones = [
            g
            for g in (stack.query.get_face_geometry(i) for i in faces)
            if g.get("geometry_type") == "Cone"
        ]
        assert len(cones) == 1, cones

        expected = math.degrees(math.atan(0.02 / 0.04))
        assert cones[0]["half_angle_degrees"] == pytest.approx(expected, abs=1e-4)
        assert cones[0]["half_angle_radians"] == pytest.approx(math.atan(0.02 / 0.04), **EXACT)
        assert "half_angle" not in cones[0], "the ambiguous key is gone"


class TestTheCameraNamesWhatItHolds:
    def test_orthographic_is_a_scale(self, stack, new_part):
        _box(stack)

        c = stack.view.get_camera()

        assert c["perspective"] is False
        assert "scale" in c
        assert "field_of_view_degrees" not in c

    def test_perspective_takes_and_reports_degrees(self, stack, new_part):
        _box(stack)
        c = stack.view.get_camera()
        e, t, u = c["eye"], c["target"], c["up"]

        stack.view.set_camera(
            e[0],
            e[1],
            e[2],
            t[0],
            t[1],
            t[2],
            u[0],
            u[1],
            u[2],
            perspective=True,
            scale_or_angle=30.0,
        )
        c2 = stack.view.get_camera()

        assert c2["perspective"] is True
        assert c2["field_of_view_degrees"] == pytest.approx(30.0, **EXACT)
        assert c2["field_of_view_radians"] == pytest.approx(math.radians(30.0), **EXACT)
        assert "scale" not in c2
