"""Transform operations for assembly components."""

import math
import traceback
from typing import Any

from ..logging import get_logger

_logger = get_logger(__name__)


class TransformsMixin:
    """Mixin providing component transform/move/rotate methods."""

    def update_component_position(
        self, component_index: int, x: float, y: float, z: float
    ) -> dict[str, Any]:
        """
        Update a component's position using a transformation matrix.

        Args:
            component_index: 0-based index of the component
            x, y, z: New position coordinates (meters)

        Returns:
            Dict with status
        """
        try:
            _logger.info(f"Updating component position: index={component_index}, pos=({x},{y},{z})")
            doc = self.doc_manager.get_active_document()
            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}"}

            occurrence = occurrences.Item(component_index + 1)

            # Get current matrix to preserve rotation
            try:
                current = list(occurrence.GetMatrix())
                # Update translation (indices 12, 13, 14 in row-major 4x4)
                current[12] = x
                current[13] = y
                current[14] = z
                occurrence.SetMatrix(current)
                return {
                    "status": "position_updated",
                    "component": component_index,
                    "position": [x, y, z],
                }
            except Exception as e:
                return {
                    "error": f"Could not update position: {e}",
                    "note": "Position update may not be available for grounded components",
                }
        except Exception as e:
            _logger.error(f"Failed to update component position: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def occurrence_move(
        self, component_index: int, dx: float, dy: float, dz: float
    ) -> dict[str, Any]:
        """
        Move a component by a relative delta.

        Uses Occurrence.Move(DeltaX, DeltaY, DeltaZ) for relative translation.

        Args:
            component_index: 0-based index of the component
            dx: X translation in meters
            dy: Y translation in meters
            dz: Z translation in meters

        Returns:
            Dict with status
        """
        try:
            _logger.info(f"Moving component: index={component_index}, delta=({dx},{dy},{dz})")
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)
            occurrence.Move(dx, dy, dz)

            return {"status": "moved", "component_index": component_index, "delta": [dx, dy, dz]}
        except Exception as e:
            _logger.error(f"Failed to move component: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def occurrence_rotate(
        self,
        component_index: int,
        axis_x1: float,
        axis_y1: float,
        axis_z1: float,
        axis_x2: float,
        axis_y2: float,
        axis_z2: float,
        angle: float,
    ) -> dict[str, Any]:
        """
        Rotate a component around an axis.

        Uses Occurrence.Rotate(AxisX1, AxisY1, AxisZ1, AxisX2, AxisY2, AxisZ2, Angle).
        The axis is defined by two 3D points. Angle is in degrees (converted to radians internally).

        Args:
            component_index: 0-based index of the component
            axis_x1, axis_y1, axis_z1: First point of rotation axis (meters)
            axis_x2, axis_y2, axis_z2: Second point of rotation axis (meters)
            angle: Rotation angle in degrees

        Returns:
            Dict with status
        """
        try:
            import math

            _logger.info(f"Rotating component: index={component_index}, angle={angle}")
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)
            angle_rad = math.radians(angle)
            occurrence.Rotate(axis_x1, axis_y1, axis_z1, axis_x2, axis_y2, axis_z2, angle_rad)

            return {
                "status": "rotated",
                "component_index": component_index,
                "axis": [[axis_x1, axis_y1, axis_z1], [axis_x2, axis_y2, axis_z2]],
                "angle_degrees": angle,
            }
        except Exception as e:
            _logger.error(f"Failed to rotate component: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_component_transform(
        self,
        component_index: int,
        origin_x: float,
        origin_y: float,
        origin_z: float,
        angle_x: float,
        angle_y: float,
        angle_z: float,
    ) -> dict[str, Any]:
        """
        Set a component's full transform (position + rotation).

        Uses Occurrence.PutTransform(OriginX, OriginY, OriginZ, AngleX, AngleY, AngleZ).
        Angles are in degrees (converted to radians internally).

        Args:
            component_index: 0-based index of the component
            origin_x: X position in meters
            origin_y: Y position in meters
            origin_z: Z position in meters
            angle_x: Rotation around X axis in degrees
            angle_y: Rotation around Y axis in degrees
            angle_z: Rotation around Z axis in degrees

        Returns:
            Dict with status
        """
        try:
            import math

            _logger.info(f"Setting component transform: index={component_index}")
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)
            ax_rad = math.radians(angle_x)
            ay_rad = math.radians(angle_y)
            az_rad = math.radians(angle_z)
            occurrence.PutTransform(origin_x, origin_y, origin_z, ax_rad, ay_rad, az_rad)

            return {
                "status": "updated",
                "component_index": component_index,
                "origin": [origin_x, origin_y, origin_z],
                "angles_degrees": [angle_x, angle_y, angle_z],
            }
        except Exception as e:
            _logger.error(f"Failed to set component transform: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_component_origin(
        self, component_index: int, x: float, y: float, z: float
    ) -> dict[str, Any]:
        """
        Set a component's origin (position only, no rotation change).

        Uses Occurrence.PutOrigin(OriginX, OriginY, OriginZ).

        Args:
            component_index: 0-based index of the component
            x: X position in meters
            y: Y position in meters
            z: Z position in meters

        Returns:
            Dict with status
        """
        try:
            _logger.info(f"Setting component origin: index={component_index}, pos=({x},{y},{z})")
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)
            occurrence.PutOrigin(x, y, z)

            return {"status": "updated", "component_index": component_index, "origin": [x, y, z]}
        except Exception as e:
            _logger.error(f"Failed to set component origin: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def mirror_component(self, component_index: int, plane_index: int) -> dict[str, Any]:
        """
        Mirror a component across a reference plane.

        Uses Occurrence.Mirror(pPlane).

        Args:
            component_index: 0-based index of the component
            plane_index: 1-based index of the reference plane

        Returns:
            Dict with status
        """
        try:
            _logger.info(f"Mirroring component: index={component_index}, plane={plane_index}")
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            ref_planes = doc.RefPlanes
            if plane_index < 1 or plane_index > ref_planes.Count:
                return {"error": f"Invalid plane_index: {plane_index}. Count: {ref_planes.Count}"}

            occurrence = occurrences.Item(component_index + 1)
            plane = ref_planes.Item(plane_index)
            occurrence.Mirror(plane)

            return {
                "status": "mirrored",
                "component_index": component_index,
                "plane_index": plane_index,
            }
        except Exception as e:
            _logger.error(f"Failed to mirror component: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def put_transform_euler(
        self,
        component_index: int,
        x: float,
        y: float,
        z: float,
        rx: float,
        ry: float,
        rz: float,
    ) -> dict[str, Any]:
        """
        Set a component's transform using Euler angles.

        Uses occurrence.PutTransform(x, y, z, rx_rad, ry_rad, rz_rad).
        Angles are in degrees (converted to radians internally).

        Args:
            component_index: 0-based index of the component
            x: X position in meters
            y: Y position in meters
            z: Z position in meters
            rx: Rotation around X axis in degrees
            ry: Rotation around Y axis in degrees
            rz: Rotation around Z axis in degrees

        Returns:
            Dict with status
        """
        try:
            _logger.info(f"Setting Euler transform: index={component_index}")
            doc = self.doc_manager.get_active_document()
            occurrences, occurrence, err = self._validate_occurrence_index(doc, component_index)
            if err:
                return err

            rx_rad = math.radians(rx)
            ry_rad = math.radians(ry)
            rz_rad = math.radians(rz)
            occurrence.PutTransform(x, y, z, rx_rad, ry_rad, rz_rad)

            return {
                "status": "updated",
                "component_index": component_index,
                "position": [x, y, z],
                "angles_degrees": [rx, ry, rz],
            }
        except Exception as e:
            _logger.error(f"Failed to set Euler transform: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def put_origin(
        self,
        component_index: int,
        x: float,
        y: float,
        z: float,
    ) -> dict[str, Any]:
        """
        Set a component's origin (position only, no rotation change).

        Uses occurrence.PutOrigin(x, y, z).

        Args:
            component_index: 0-based index of the component
            x: X position in meters
            y: Y position in meters
            z: Z position in meters

        Returns:
            Dict with status
        """
        try:
            _logger.info(f"Setting origin: index={component_index}, pos=({x},{y},{z})")
            doc = self.doc_manager.get_active_document()
            occurrences, occurrence, err = self._validate_occurrence_index(doc, component_index)
            if err:
                return err

            occurrence.PutOrigin(x, y, z)

            return {
                "status": "updated",
                "component_index": component_index,
                "origin": [x, y, z],
            }
        except Exception as e:
            _logger.error(f"Failed to set origin: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}
