"""Unit tests for the draft element locator.

Every ``Dimensions.Add*`` method in Solid Edge takes the object being
dimensioned, never bare coordinates. The locator is what lets a caller who
only has coordinates reach those APIs, so its picking rules are the thing
that decides whether a dimension lands on the right element.
"""

from __future__ import annotations

import math

import pytest

from solidedge_mcp.backends.export._locate import (
    Hit,
    element_count,
    nearest_element,
    no_element_error,
)


class FakeCollection:
    """A 1-indexed COM collection."""

    def __init__(self, items):
        self._items = list(items)

    @property
    def Count(self):
        return len(self._items)

    def Item(self, index):
        return self._items[index - 1]


class FakeLine:
    def __init__(self, x1, y1, x2, y2):
        self._start = (x1, y1)
        self._end = (x2, y2)

    def GetStartPoint(self):
        return self._start

    def GetEndPoint(self):
        return self._end


class FakeCircle:
    def __init__(self, cx, cy, radius):
        self._center = (cx, cy)
        self.Radius = radius

    def GetCenterPoint(self):
        return self._center


class FakeSpline:
    """A shape the locator cannot solve, so it falls back to keypoints."""

    def __init__(self, points):
        self._points = list(points)

    @property
    def KeyPointCount(self):
        return len(self._points)

    def GetKeyPoint(self, index):
        x, y = self._points[index]
        return (x, y, 0.0, 0, 0)


class FakeView:
    """A drawing view whose geometry is in view coordinates."""

    def __init__(self, origin, lines=(), circles=()):
        self._origin = origin
        self.DVLines2d = FakeCollection(lines)
        self.DVCircles2d = FakeCollection(circles)

    def ViewToSheet(self, x, y):
        return (self._origin[0] + x, self._origin[1] + y)


class FakeSheet:
    def __init__(self, lines=(), circles=(), arcs=(), curves=(), views=()):
        self.Lines2d = FakeCollection(lines)
        self.Circles2d = FakeCollection(circles)
        self.Arcs2d = FakeCollection(arcs)
        self.BsplineCurves2d = FakeCollection(curves)
        self.DrawingViews = FakeCollection(views)


class TestPickingTheRightElement:
    def test_picks_the_nearest_line(self):
        near = FakeLine(0.0, 0.0, 0.1, 0.0)
        far = FakeLine(0.0, 0.5, 0.1, 0.5)
        sheet = FakeSheet(lines=[far, near])

        hit = nearest_element(sheet, 0.05, 0.001)

        assert hit is not None
        assert hit.obj is near
        assert hit.kind == "line"
        assert hit.distance == pytest.approx(0.001)

    def test_measures_to_the_segment_not_its_endpoints(self):
        """A point beside a line's middle is close to it, not far from both ends."""
        sheet = FakeSheet(lines=[FakeLine(0.0, 0.0, 1.0, 0.0)])

        hit = nearest_element(sheet, 0.5, 0.01)

        assert hit is not None
        assert hit.distance == pytest.approx(0.01)

    def test_a_point_past_the_end_measures_to_the_end(self):
        sheet = FakeSheet(lines=[FakeLine(0.0, 0.0, 0.1, 0.0)])

        hit = nearest_element(sheet, 0.13, 0.04)

        assert hit is not None
        assert hit.distance == pytest.approx(math.hypot(0.03, 0.04))

    def test_a_zero_length_line_does_not_divide_by_zero(self):
        sheet = FakeSheet(lines=[FakeLine(0.1, 0.1, 0.1, 0.1)])

        hit = nearest_element(sheet, 0.1, 0.13)

        assert hit is not None
        assert hit.distance == pytest.approx(0.03)

    def test_a_circle_is_measured_to_its_rim(self):
        """A pick lands on the drawn ring, not at the empty centre."""
        sheet = FakeSheet(circles=[FakeCircle(0.1, 0.1, 0.02)])

        on_rim = nearest_element(sheet, 0.12, 0.1)
        at_centre = nearest_element(sheet, 0.1, 0.1)

        assert on_rim is not None and on_rim.distance == pytest.approx(0.0)
        assert at_centre is not None and at_centre.distance == pytest.approx(0.02)

    def test_a_circle_reports_its_centre_and_radius(self):
        sheet = FakeSheet(circles=[FakeCircle(0.1, 0.2, 0.03)])

        hit = nearest_element(sheet, 0.13, 0.2)

        assert hit is not None
        assert hit.center == (0.1, 0.2)
        assert hit.radius == 0.03

    def test_an_unsolvable_shape_falls_back_to_its_keypoints(self):
        sheet = FakeSheet(curves=[FakeSpline([(0.0, 0.0), (0.2, 0.2)])])

        hit = nearest_element(sheet, 0.2, 0.21)

        assert hit is not None
        assert hit.kind == "curve"
        assert hit.distance == pytest.approx(0.01)


class TestFiltering:
    def test_kinds_excludes_everything_else(self):
        line = FakeLine(0.0, 0.0, 0.1, 0.0)
        circle = FakeCircle(0.5, 0.5, 0.02)
        sheet = FakeSheet(lines=[line], circles=[circle])

        hit = nearest_element(sheet, 0.05, 0.001, kinds=("circle",))

        assert hit is not None
        assert hit.obj is circle

    def test_tolerance_rejects_a_distant_pick(self):
        sheet = FakeSheet(lines=[FakeLine(0.0, 0.0, 0.1, 0.0)])

        assert nearest_element(sheet, 0.05, 0.5, tolerance=0.01) is None
        assert nearest_element(sheet, 0.05, 0.5) is not None

    def test_an_empty_sheet_finds_nothing(self):
        assert nearest_element(FakeSheet(), 0.1, 0.1) is None


class TestDrawingViewGeometry:
    def test_view_coordinates_are_converted_to_sheet_space(self):
        """DVLines2d are in view space; ViewToSheet moves them onto the sheet."""
        view = FakeView((0.2, 0.15), lines=[FakeLine(-0.04, 0.01, 0.04, 0.01)])
        sheet = FakeSheet(views=[view])

        hit = nearest_element(sheet, 0.20, 0.16)

        assert hit is not None
        assert hit.source == "view 0"
        assert hit.distance == pytest.approx(0.0)

    def test_sheet_geometry_wins_when_it_is_closer(self):
        view = FakeView((0.2, 0.15), lines=[FakeLine(-0.04, 0.01, 0.04, 0.01)])
        line = FakeLine(0.20, 0.1601, 0.30, 0.1601)
        sheet = FakeSheet(lines=[line], views=[view])

        hit = nearest_element(sheet, 0.20, 0.16005)

        assert hit is not None
        assert hit.source == "sheet"


class TestIdentity:
    def test_two_picks_on_one_element_compare_equal(self):
        """pywin32 returns a fresh wrapper per Item(), so ``is`` cannot be used."""
        sheet = FakeSheet(lines=[FakeLine(0.0, 0.0, 0.1, 0.0)])

        first = nearest_element(sheet, 0.02, 0.0)
        second = nearest_element(sheet, 0.08, 0.0)

        assert first is not None and second is not None
        assert first.identity == second.identity

    def test_two_picks_on_different_elements_differ(self):
        sheet = FakeSheet(lines=[FakeLine(0.0, 0.0, 0.1, 0.0), FakeLine(0.1, 0.0, 0.1, 0.1)])

        first = nearest_element(sheet, 0.02, 0.0)
        second = nearest_element(sheet, 0.1, 0.08)

        assert first is not None and second is not None
        assert first.identity != second.identity

    def test_describe_reports_a_zero_based_index(self):
        hit = Hit(obj=None, kind="line", distance=0.0, source="sheet", index=1)

        assert hit.describe()["index"] == 0


class TestReporting:
    def test_element_count_sums_every_collection(self):
        view = FakeView((0.0, 0.0), lines=[FakeLine(0, 0, 1, 0)])
        sheet = FakeSheet(
            lines=[FakeLine(0, 0, 1, 0), FakeLine(0, 1, 1, 1)],
            circles=[FakeCircle(0, 0, 1)],
            views=[view],
        )

        assert element_count(sheet) == 4

    def test_the_error_names_the_point_and_what_was_wanted(self):
        error = no_element_error(0.1, 0.2, "circle or arc")

        assert "circle or arc" in error["error"]
        assert error["point"] == [0.1, 0.2]


class TestResilience:
    class Broken:
        @property
        def Count(self):
            raise RuntimeError("COM is unhappy")

    def test_a_collection_that_raises_does_not_sink_the_search(self):
        sheet = FakeSheet(lines=[FakeLine(0.0, 0.0, 0.1, 0.0)])
        sheet.Circles2d = TestResilience.Broken()

        hit = nearest_element(sheet, 0.05, 0.001)

        assert hit is not None
        assert hit.kind == "line"

    def test_an_element_missing_its_accessors_is_skipped(self):
        good = FakeLine(0.0, 0.5, 0.1, 0.5)
        sheet = FakeSheet(lines=[object(), good])

        hit = nearest_element(sheet, 0.05, 0.5)

        assert hit is not None
        assert hit.obj is good

    def test_a_view_without_viewtosheet_uses_its_origin(self):
        class OriginOnly:
            DVLines2d = FakeCollection([FakeLine(0.0, 0.0, 0.1, 0.0)])

            def GetOrigin(self):
                return (0.3, 0.4)

        sheet = FakeSheet(views=[OriginOnly()])

        hit = nearest_element(sheet, 0.35, 0.4)

        assert hit is not None
        assert hit.distance == pytest.approx(0.0)
