"""Spatial context: where existing geometry sits before you sketch on it.

An LLM driving Solid Edge cannot see the model. Without knowing where the
existing body is, and which world axes the sketch plane maps onto, it places
features by guesswork — the most common cause of misplaced geometry. This
module answers all of that in a single call.
"""

from typing import Any

from ..logging import get_logger

_logger = get_logger(__name__)

#: Distance (meters) under which a coordinate counts as "on the origin".
ORIGIN_TOLERANCE = 0.001

#: Solid Edge's three default reference planes, 1-based as the COM
#: RefPlanes collection indexes them. Sketch (x, y) maps onto the world axes
#: named here, and ``normal`` is the world direction a positive extrusion
#: travels along.
PLANE_AXIS_MAP: dict[str, dict[str, Any]] = {
    "1": {
        "name": "Top",
        "plane": "XY",
        "normal": [0.0, 0.0, 1.0],
        "sketch_x_axis": "X",
        "sketch_y_axis": "Y",
    },
    "2": {
        "name": "Right",
        "plane": "YZ",
        "normal": [1.0, 0.0, 0.0],
        "sketch_x_axis": "Y",
        "sketch_y_axis": "Z",
    },
    "3": {
        "name": "Front",
        "plane": "XZ",
        "normal": [0.0, 1.0, 0.0],
        "sketch_x_axis": "X",
        "sketch_y_axis": "Z",
    },
}


class SpatialContextMixin:
    """Mixin providing the combined spatial-context query."""

    doc_manager: Any

    def get_spatial_context(self) -> dict[str, Any]:
        """
        Describe where the existing geometry sits, in one call.

        Combines the active body count, the body's bounding box (min, max,
        center and dimensions), whether that body is centered on the origin,
        the reference plane of the open sketch (if any), and the static
        plane-to-world-axis map. Call this before sketching so new geometry
        lands where you intend.

        Never raises: any part that cannot be computed comes back as null and
        the reason is recorded in ``warnings``.

        Returns:
            Dict with body_count, bounding_box, centered_on_origin,
            active_sketch, plane_axis_map, units and warnings
        """
        warnings: list[str] = []
        result: dict[str, Any] = {
            "body_count": None,
            "bounding_box": None,
            "centered_on_origin": None,
            "active_sketch": None,
            "plane_axis_map": PLANE_AXIS_MAP,
            "origin_tolerance": ORIGIN_TOLERANCE,
            "units": "meters",
            "warnings": warnings,
        }

        body_count = self._spatial_body_count(warnings)
        result["body_count"] = body_count

        if body_count:
            box = self._spatial_bounding_box(warnings)
            if box is not None:
                result["bounding_box"] = box
                center = box["center"]
                result["centered_on_origin"] = all(abs(c) < ORIGIN_TOLERANCE for c in center)
        elif body_count == 0:
            warnings.append("No bodies in the active document; nothing to bound.")

        result["active_sketch"] = self._spatial_active_sketch(warnings)
        return result

    # -- helpers ---------------------------------------------------------

    def _spatial_body_count(self, warnings: list[str]) -> int | None:
        """Number of design bodies (Models) in the active document, or None."""
        try:
            doc = self.doc_manager.get_active_document()
            return int(doc.Models.Count)
        except Exception as e:
            warnings.append(f"Body count unavailable: {e}")
            return None

    def _spatial_bounding_box(self, warnings: list[str]) -> dict[str, Any] | None:
        """Bounding box with center, reusing get_bounding_box()."""
        try:
            box = self.get_bounding_box()
        except Exception as e:  # get_bounding_box already traps, but stay safe
            warnings.append(f"Bounding box unavailable: {e}")
            return None

        if not isinstance(box, dict) or "error" in box:
            detail = box.get("error") if isinstance(box, dict) else box
            warnings.append(f"Bounding box unavailable: {detail}")
            return None

        try:
            minimum = [float(v) for v in box["min"][:3]]
            maximum = [float(v) for v in box["max"][:3]]
            if len(minimum) != 3 or len(maximum) != 3:
                raise ValueError(f"expected 3 coordinates, got {minimum} / {maximum}")
        except Exception as e:
            warnings.append(f"Bounding box malformed: {e}")
            return None

        center = [(lo + hi) / 2.0 for lo, hi in zip(minimum, maximum, strict=False)]
        return {
            "min": minimum,
            "max": maximum,
            "center": center,
            "dimensions": {
                "x": maximum[0] - minimum[0],
                "y": maximum[1] - minimum[1],
                "z": maximum[2] - minimum[2],
            },
        }

    def _spatial_active_sketch(self, warnings: list[str]) -> dict[str, Any] | None:
        """Plane index/name of the open sketch, or None when none is open."""
        try:
            # Deferred import: managers.py builds the QueryManager, so importing
            # it at module scope would be circular.
            from solidedge_mcp.managers import sketch_manager

            plane_index = sketch_manager.get_active_plane_index()
        except Exception as e:
            warnings.append(f"Active sketch plane unavailable: {e}")
            return None

        if plane_index is None:
            return None

        entry = PLANE_AXIS_MAP.get(str(plane_index))
        name = entry["name"] if entry else None
        if name is None:
            try:
                doc = self.doc_manager.get_active_document()
                name = doc.RefPlanes.Item(plane_index).Name
            except Exception:
                name = f"RefPlane {plane_index}"

        sketch: dict[str, Any] = {"plane_index": plane_index, "plane_name": name}
        if entry:
            sketch["plane"] = entry["plane"]
            sketch["normal"] = entry["normal"]
        return sketch
