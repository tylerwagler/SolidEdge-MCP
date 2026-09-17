"""A vertex must come back as three coordinates, or not at all.

``Vertex.GetPointData`` takes one ``[in, out] SAFEARRAY(VT_R8)*``. pywin32
hands the filled values back as the call's result -- a flat ``(x, y, z)``
tuple, verified on Solid Edge 2026 -- and leaves the passed-in list alone.
The code read ``result[0]``, which is x on its own, so every vertex this
server reported was a one-element point: the eight corners of a 0.08 x 0.048
x 0.03 box came back as the two points {(0.0,), (0.08,)}.

The fakes here return exactly what Solid Edge returns.
"""

from __future__ import annotations

from unittest.mock import MagicMock

from solidedge_mcp.backends.query._brep import _point3

CORNER = (0.08, 0.048, 0.03)


class TestPoint3:
    def test_the_flat_tuple_solid_edge_returns(self):
        assert _point3(CORNER, [0.0, 0.0, 0.0]) == [0.08, 0.048, 0.03]

    def test_a_wrapped_tuple_is_unwrapped(self):
        assert _point3((CORNER,), [0.0, 0.0, 0.0]) == [0.08, 0.048, 0.03]

    def test_a_filled_buffer_is_used_when_there_is_no_result(self):
        assert _point3(None, [0.08, 0.048, 0.03]) == [0.08, 0.048, 0.03]

    def test_a_lone_coordinate_is_not_a_point(self):
        """result[0] of the flat tuple is what the old code read."""
        assert _point3((0.08,), [0.0, 0.0, 0.0]) is None

    def test_non_numbers_are_not_a_point(self):
        assert _point3(("x", "y", "z"), [0.0, 0.0, 0.0]) is None


class _Vertex:
    def __init__(self, point: tuple[float, float, float]) -> None:
        self._point = point

    def GetPointData(self, buffer):  # noqa: N802 - COM spelling
        return self._point  # the flat tuple, and the buffer is left alone


class _Vertices:
    def __init__(self, *points: tuple[float, float, float]) -> None:
        self._items = [_Vertex(p) for p in points]
        self.Count = len(points)

    def Item(self, index: int) -> _Vertex:  # noqa: N802 - COM spelling
        return self._items[index - 1]


class TestGetBodyVertices:
    def _manager(self, *points):
        from solidedge_mcp.backends.query import QueryManager

        body = MagicMock()
        body.Vertices = _Vertices(*points)
        model = MagicMock()
        model.Body = body
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc = MagicMock()
        doc.Models = models
        doc_manager = MagicMock()
        doc_manager.get_active_document.return_value = doc
        return QueryManager(doc_manager)

    def test_every_vertex_has_three_coordinates(self):
        manager = self._manager((0.0, 0.0, 0.0), CORNER)

        result = manager.get_body_vertices()

        assert "error" not in result, result
        assert [item["point"] for item in result["items"]] == [
            [0.0, 0.0, 0.0],
            [0.08, 0.048, 0.03],
        ]

    def test_the_corners_of_a_box_are_eight_distinct_points(self):
        corners = [(x, y, z) for x in (0.0, 0.08) for y in (0.0, 0.048) for z in (0.0, 0.03)]
        manager = self._manager(*corners)

        points = {tuple(item["point"]) for item in manager.get_body_vertices()["items"]}

        assert len(points) == 8
