"""Reference plane creation operations."""

import traceback
from typing import Any

from ..constants import (
    DirectionConstants,
    FaceQueryConstants,
    KeyPointTypeConstants,
    ReferenceElementConstants,
)


class RefPlaneMixin:
    """Mixin providing reference plane creation methods."""

    def create_ref_plane_by_offset(
        self, parent_plane_index: int, distance: float, normal_side: str = "Normal"
    ) -> dict[str, Any]:
        """
        Create a reference plane parallel to an existing plane at an offset distance.

        Uses RefPlanes.AddParallelByDistance(ParentPlane, Distance, NormalSide).
        Useful for creating sketches at different heights/positions.

        Args:
            parent_plane_index: Index of parent plane (1=Top/XZ, 2=Front/XY, 3=Right/YZ,
                                or higher for user-created planes)
            distance: Offset distance in meters
            normal_side: 'Normal' (igRight=2) or 'Reverse' (igLeft=1)

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            ref_planes = doc.RefPlanes

            if parent_plane_index < 1 or parent_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid plane index: {parent_plane_index}. Count: {ref_planes.Count}"
                }

            parent = ref_planes.Item(parent_plane_index)

            side_map = {
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
            }
            side_const = side_map.get(normal_side, DirectionConstants.igRight)

            ref_planes.AddParallelByDistance(parent, distance, side_const)

            return {
                "status": "created",
                "type": "reference_plane",
                "method": "parallel_by_distance",
                "parent_plane": parent_plane_index,
                "distance": distance,
                "normal_side": normal_side,
                "new_plane_index": ref_planes.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_ref_plane_normal_at_distance(
        self, distance: float, curve_end: str = "End", pivot_plane_index: int = 2
    ) -> dict[str, Any]:
        """
        Create a reference plane normal to a curve at a specified distance from an endpoint.

        Uses RefPlanes.AddNormalToCurveAtDistance(pCurve, Distance, bIgnoreNatural,
        NormalSide, [bFlip], [bOrient], [orientSurface]).
        Requires an active sketch profile that defines the curve.

        Args:
            distance: Distance from the curve endpoint in meters
            curve_end: Which end to measure from - 'Start' or 'End'
            pivot_plane_index: 1-based index of the pivot reference plane (default: 2 = Front)

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            ref_planes = doc.RefPlanes

            # igCurveEnd = 2, igCurveStart = 1
            ignore_natural = curve_end == "End"
            # NormalSide: igRight = 2
            normal_side = DirectionConstants.igRight

            ref_planes.AddNormalToCurveAtDistance(profile, distance, ignore_natural, normal_side)

            return {
                "status": "created",
                "type": "ref_plane_normal_at_distance",
                "distance": distance,
                "curve_end": curve_end,
                "new_plane_index": ref_planes.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_ref_plane_normal_at_arc_ratio(
        self, ratio: float, curve_end: str = "End", pivot_plane_index: int = 2
    ) -> dict[str, Any]:
        """
        Create a reference plane normal to a curve at an arc-length ratio.

        Uses RefPlanes.AddNormalToCurveAtArcLengthRatio(pCurve, Ratio, bIgnoreNatural,
        NormalSide, PivotPlane, PivotEnd, [bFlip], [bOrient]).
        Ratio is 0.0 (start) to 1.0 (end) of the curve arc length.

        Args:
            ratio: Arc length ratio (0.0 to 1.0)
            curve_end: Which end is pivot - 'Start' or 'End'
            pivot_plane_index: 1-based index of the pivot reference plane (default: 2 = Front)

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            if ratio < 0.0 or ratio > 1.0:
                return {"error": f"Ratio must be between 0.0 and 1.0, got {ratio}"}

            ref_planes = doc.RefPlanes
            pivot_plane = ref_planes.Item(pivot_plane_index)

            ignore_natural = curve_end == "End"
            normal_side = DirectionConstants.igRight
            # igPivotEnd = 2
            pivot_end_const = 2

            ref_planes.AddNormalToCurveAtArcLengthRatio(
                profile, ratio, ignore_natural, normal_side, pivot_plane, pivot_end_const
            )

            return {
                "status": "created",
                "type": "ref_plane_normal_at_arc_ratio",
                "ratio": ratio,
                "curve_end": curve_end,
                "new_plane_index": ref_planes.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_ref_plane_normal_at_distance_along(
        self, distance_along: float, curve_end: str = "End", pivot_plane_index: int = 2
    ) -> dict[str, Any]:
        """
        Create a reference plane normal to a curve at a distance along the curve.

        Uses RefPlanes.AddNormalToCurveAtDistanceAlongCurve(pCurve, DistAlong,
        bIgnoreNatural, NormalSide, PivotPlane, PivotEnd, [bFlip], [bOrient]).

        Args:
            distance_along: Distance along the curve in meters
            curve_end: Which end to measure from - 'Start' or 'End'
            pivot_plane_index: 1-based index of the pivot reference plane (default: 2 = Front)

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            ref_planes = doc.RefPlanes
            pivot_plane = ref_planes.Item(pivot_plane_index)

            ignore_natural = curve_end == "End"
            normal_side = DirectionConstants.igRight
            pivot_end_const = 2

            ref_planes.AddNormalToCurveAtDistanceAlongCurve(
                profile, distance_along, ignore_natural, normal_side, pivot_plane, pivot_end_const
            )

            return {
                "status": "created",
                "type": "ref_plane_normal_at_distance_along",
                "distance_along": distance_along,
                "curve_end": curve_end,
                "new_plane_index": ref_planes.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_ref_plane_parallel_by_tangent(
        self, parent_plane_index: int, face_index: int, normal_side: str = "Normal"
    ) -> dict[str, Any]:
        """
        Create a reference plane parallel to a parent plane and tangent to a face.

        Uses RefPlanes.AddParallelByTangent(pParentPlane, pFace, NormalSide,
        [bFlip], [bOrient], [orientSurface]).

        Args:
            parent_plane_index: 1-based index of the parent reference plane
            face_index: 0-based index of the face to be tangent to
            normal_side: 'Normal' (igRight=2) or 'Reverse' (igLeft=1)

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            ref_planes = doc.RefPlanes

            if parent_plane_index < 1 or parent_plane_index > ref_planes.Count:
                return {
                    "error": "Invalid parent_plane_index: "
                    f"{parent_plane_index}. "
                    f"Count: {ref_planes.Count}"
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists."}

            model = models.Item(1)
            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)

            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face_index: {face_index}. Body has {faces.Count} faces."}

            parent_plane = ref_planes.Item(parent_plane_index)
            face = faces.Item(face_index + 1)

            side_map = {
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
            }
            side_const = side_map.get(normal_side, DirectionConstants.igRight)

            ref_planes.AddParallelByTangent(parent_plane, face, side_const)

            return {
                "status": "created",
                "type": "ref_plane_parallel_by_tangent",
                "parent_plane_index": parent_plane_index,
                "face_index": face_index,
                "normal_side": normal_side,
                "new_plane_index": ref_planes.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_ref_plane_normal_to_curve(
        self, curve_end: str = "End", pivot_plane_index: int = 2
    ) -> dict[str, Any]:
        """
        Create a reference plane normal (perpendicular) to a curve at its endpoint.

        Useful for creating sweep cross-section sketches perpendicular to a path.
        Requires an active sketch profile that defines the curve.

        Args:
            curve_end: Which end of the curve to place the plane at - 'Start' or 'End'
            pivot_plane_index: 1-based index of the pivot reference plane (default: 2 = Front)

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            ref_planes = doc.RefPlanes
            pivot_plane = ref_planes.Item(pivot_plane_index)

            # igCurveEnd = 2, igCurveStart = 1
            curve_end_const = 2 if curve_end == "End" else 1
            # igPivotEnd = 2
            pivot_end_const = 2

            ref_planes.AddNormalToCurve(
                profile, curve_end_const, pivot_plane, pivot_end_const, True
            )

            new_index = ref_planes.Count

            return {
                "status": "created",
                "type": "ref_plane_normal_to_curve",
                "curve_end": curve_end,
                "pivot_plane_index": pivot_plane_index,
                "new_plane_index": new_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_ref_plane_by_angle(
        self, parent_plane_index: int, angle: float, normal_side: str = "Normal"
    ) -> dict[str, Any]:
        """
        Create a reference plane at an angle to an existing plane.

        Uses RefPlanes.AddAngularByAngle(ParentPlane, Angle, NormalSide).
        Type library: AddAngularByAngle(ParentPlane: IDispatch, Angle: VT_R8,
        NormalSide: FeaturePropertyConstants, [Edge: VT_VARIANT]).

        Args:
            parent_plane_index: Index of parent plane (1=Top/XZ, 2=Front/XY, 3=Right/YZ)
            angle: Angle in degrees from the parent plane
            normal_side: 'Normal' (igRight=2) or 'Reverse' (igLeft=1)

        Returns:
            Dict with status and new plane index
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()
            ref_planes = doc.RefPlanes

            if parent_plane_index < 1 or parent_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid plane index: {parent_plane_index}. Count: {ref_planes.Count}"
                }

            parent = ref_planes.Item(parent_plane_index)

            side_map = {
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
            }
            side_const = side_map.get(normal_side, DirectionConstants.igRight)

            # Angle in radians for the COM API
            angle_rad = math.radians(angle)

            ref_planes.AddAngularByAngle(parent, angle_rad, side_const)

            return {
                "status": "created",
                "type": "reference_plane",
                "method": "angular_by_angle",
                "parent_plane": parent_plane_index,
                "angle_degrees": angle,
                "normal_side": normal_side,
                "new_plane_index": ref_planes.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_ref_plane_by_3_points(
        self,
        x1: float,
        y1: float,
        z1: float,
        x2: float,
        y2: float,
        z2: float,
        x3: float,
        y3: float,
        z3: float,
    ) -> dict[str, Any]:
        """
        Create a reference plane through 3 points in space.

        Uses RefPlanes.AddBy3Points(Point1X, Point1Y, Point1Z, ...).
        Type library: AddBy3Points(9x VT_R8 params) -> RefPlane*.

        Args:
            x1, y1, z1: First point coordinates (meters)
            x2, y2, z2: Second point coordinates (meters)
            x3, y3, z3: Third point coordinates (meters)

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            ref_planes = doc.RefPlanes

            ref_planes.AddBy3Points(x1, y1, z1, x2, y2, z2, x3, y3, z3)

            return {
                "status": "created",
                "type": "reference_plane",
                "method": "by_3_points",
                "point1": [x1, y1, z1],
                "point2": [x2, y2, z2],
                "point3": [x3, y3, z3],
                "new_plane_index": ref_planes.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_ref_plane_midplane(self, plane1_index: int, plane2_index: int) -> dict[str, Any]:
        """
        Create a reference plane midway between two existing planes.

        Uses RefPlanes.AddMidPlane(Plane1, Plane2).
        Useful for symmetry operations.

        Args:
            plane1_index: Index of first plane (1=Top/XZ, 2=Front/XY, 3=Right/YZ)
            plane2_index: Index of second plane

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            ref_planes = doc.RefPlanes

            if plane1_index < 1 or plane1_index > ref_planes.Count:
                return {"error": f"Invalid plane1 index: {plane1_index}. Count: {ref_planes.Count}"}
            if plane2_index < 1 or plane2_index > ref_planes.Count:
                return {"error": f"Invalid plane2 index: {plane2_index}. Count: {ref_planes.Count}"}

            plane1 = ref_planes.Item(plane1_index)
            plane2 = ref_planes.Item(plane2_index)

            ref_planes.AddMidPlane(plane1, plane2)

            return {
                "status": "created",
                "type": "reference_plane",
                "method": "mid_plane",
                "plane1_index": plane1_index,
                "plane2_index": plane2_index,
                "new_plane_index": ref_planes.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_ref_plane_normal_at_keypoint(
        self, keypoint_type: str = "End", pivot_plane_index: int = 2
    ) -> dict[str, Any]:
        """
        Create a reference plane normal to a curve at a keypoint (start or end).

        Uses RefPlanes.AddNormalToCurveAtKeyPoint(Curve, OrientationPlane, KeyPoint,
        KeyPointTypeConstant, XAxisRotation, normalOrientation, selectedCurveEnd).
        Requires an active sketch profile that defines the curve.

        Args:
            keypoint_type: 'Start' or 'End' of the curve
            pivot_plane_index: 1-based index of the orientation reference plane (default: 2 = Front)

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            ref_planes = doc.RefPlanes

            if pivot_plane_index < 1 or pivot_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid pivot_plane_index: {pivot_plane_index}. "
                    f"Count: {ref_planes.Count}"
                }

            orientation_plane = ref_planes.Item(pivot_plane_index)

            # KeyPointTypeConstant: igKeyPointStart=1, igKeyPointEnd=2
            kp_const = (
                KeyPointTypeConstants.igKeyPointStart
                if keypoint_type == "Start"
                else KeyPointTypeConstants.igKeyPointEnd
            )

            # selectedCurveEnd: igCurveStart=14, igCurveEnd=15
            curve_end_const = (
                ReferenceElementConstants.igCurveStart
                if keypoint_type == "Start"
                else ReferenceElementConstants.igCurveEnd
            )

            ref_planes.AddNormalToCurveAtKeyPoint(
                profile,  # Curve
                orientation_plane,  # OrientationPlane
                profile,  # KeyPoint (same as curve for endpoint)
                kp_const,  # KeyPointTypeConstant
                0.0,  # XAxisRotation
                ReferenceElementConstants.igNormalSide,  # normalOrientation
                curve_end_const,  # selectedCurveEnd
            )

            return {
                "status": "created",
                "type": "ref_plane_normal_at_keypoint",
                "keypoint_type": keypoint_type,
                "pivot_plane_index": pivot_plane_index,
                "new_plane_index": ref_planes.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_ref_plane_tangent_cylinder_angle(
        self, face_index: int, angle: float, parent_plane_index: int = 1
    ) -> dict[str, Any]:
        """
        Create a reference plane tangent to a cylindrical or conical face at an angle.

        Uses RefPlanes.AddTangentToCylinderOrConeAtAngle(Face, ParentPlane,
        AngleOfRotation, XAxisAngle, ExtentSide, normalOrientation).

        Args:
            face_index: 0-based index of the cylindrical/conical face
            angle: Angle of rotation in degrees
            parent_plane_index: 1-based index of the parent reference plane (default: 1 = Top)

        Returns:
            Dict with status and new plane index
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()
            ref_planes = doc.RefPlanes

            if parent_plane_index < 1 or parent_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid parent_plane_index: {parent_plane_index}. "
                    f"Count: {ref_planes.Count}"
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a feature first."}

            model = models.Item(1)
            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)

            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face_index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)
            parent_plane = ref_planes.Item(parent_plane_index)

            angle_rad = math.radians(angle)

            ref_planes.AddTangentToCylinderOrConeAtAngle(
                face,  # Face
                parent_plane,  # ParentPlane
                angle_rad,  # AngleOfRotation
                0.0,  # XAxisAngle
                DirectionConstants.igRight,  # ExtentSide
                ReferenceElementConstants.igNormalSide,  # normalOrientation
            )

            return {
                "status": "created",
                "type": "ref_plane_tangent_cylinder_angle",
                "face_index": face_index,
                "angle_degrees": angle,
                "parent_plane_index": parent_plane_index,
                "new_plane_index": ref_planes.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_ref_plane_tangent_cylinder_keypoint(
        self, face_index: int, keypoint_type: str = "End", parent_plane_index: int = 1
    ) -> dict[str, Any]:
        """
        Create a reference plane tangent to a cylindrical or conical face at a keypoint.

        Uses RefPlanes.AddTangentToCylinderOrConeAtKeyPoint(Face, ParentPlane,
        KeyPoint, KeyPointTypeConstant, XAxisAngle, ExtentSide, normalOrientation).

        Args:
            face_index: 0-based index of the cylindrical/conical face
            keypoint_type: 'Start' or 'End' keypoint on the face
            parent_plane_index: 1-based index of the parent reference plane (default: 1 = Top)

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            ref_planes = doc.RefPlanes

            if parent_plane_index < 1 or parent_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid parent_plane_index: {parent_plane_index}. "
                    f"Count: {ref_planes.Count}"
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a feature first."}

            model = models.Item(1)
            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)

            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face_index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)
            parent_plane = ref_planes.Item(parent_plane_index)

            kp_const = (
                KeyPointTypeConstants.igKeyPointStart
                if keypoint_type == "Start"
                else KeyPointTypeConstants.igKeyPointEnd
            )

            ref_planes.AddTangentToCylinderOrConeAtKeyPoint(
                face,  # Face
                parent_plane,  # ParentPlane
                face,  # KeyPoint (use face itself as keypoint reference)
                kp_const,  # KeyPointTypeConstant
                0.0,  # XAxisAngle
                DirectionConstants.igRight,  # ExtentSide
                ReferenceElementConstants.igNormalSide,  # normalOrientation
            )

            return {
                "status": "created",
                "type": "ref_plane_tangent_cylinder_keypoint",
                "face_index": face_index,
                "keypoint_type": keypoint_type,
                "parent_plane_index": parent_plane_index,
                "new_plane_index": ref_planes.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_ref_plane_tangent_surface_keypoint(
        self, face_index: int, keypoint_type: str = "End", parent_plane_index: int = 1
    ) -> dict[str, Any]:
        """
        Create a reference plane tangent to a curved surface at a keypoint.

        Uses RefPlanes.AddTangentToCurvedSurfaceAtKeyPoint(Face, ParentPlane,
        KeyPoint, KeyPointTypeConstant, XAxisAngle, normalOrientation).

        Args:
            face_index: 0-based index of the curved surface face
            keypoint_type: 'Start' or 'End' keypoint on the face
            parent_plane_index: 1-based index of the parent reference plane (default: 1 = Top)

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            ref_planes = doc.RefPlanes

            if parent_plane_index < 1 or parent_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid parent_plane_index: {parent_plane_index}. "
                    f"Count: {ref_planes.Count}"
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a feature first."}

            model = models.Item(1)
            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)

            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face_index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)
            parent_plane = ref_planes.Item(parent_plane_index)

            kp_const = (
                KeyPointTypeConstants.igKeyPointStart
                if keypoint_type == "Start"
                else KeyPointTypeConstants.igKeyPointEnd
            )

            ref_planes.AddTangentToCurvedSurfaceAtKeyPoint(
                face,  # Face
                parent_plane,  # ParentPlane
                face,  # KeyPoint (use face itself as keypoint reference)
                kp_const,  # KeyPointTypeConstant
                0.0,  # XAxisAngle
                ReferenceElementConstants.igNormalSide,  # normalOrientation
            )

            return {
                "status": "created",
                "type": "ref_plane_tangent_surface_keypoint",
                "face_index": face_index,
                "keypoint_type": keypoint_type,
                "parent_plane_index": parent_plane_index,
                "new_plane_index": ref_planes.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_ref_plane_normal_at_distance_v2(
        self,
        curve_edge_index: int,
        orientation_plane_index: int,
        distance: float,
        normal_side: int = 2,
    ) -> dict[str, Any]:
        """
        Create a reference plane normal to a curve at a distance from the curve.

        Uses RefPlanes.AddNormalToCurveAtDistance(Curve, OrientationPlane,
        Distance, normalOrientation, selectedCurveEnd).

        Args:
            curve_edge_index: 0-based edge index on the body to use as curve
            orientation_plane_index: 1-based index of the orientation reference plane
            distance: Distance from curve endpoint in meters
            normal_side: Normal orientation (1=igLeft, 2=igRight)

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            body = model.Body
            edges = body.Edges(FaceQueryConstants.igQueryAll)

            if curve_edge_index < 0 or curve_edge_index >= edges.Count:
                return {
                    "error": f"Invalid curve_edge_index: {curve_edge_index}. "
                    f"Edge count: {edges.Count}"
                }

            ref_planes = doc.RefPlanes
            curve = edges.Item(curve_edge_index + 1)
            orient_plane = ref_planes.Item(orientation_plane_index)

            ref_planes.AddNormalToCurveAtDistance(
                curve, orient_plane, distance, normal_side, ReferenceElementConstants.igCurveEnd
            )

            return {
                "status": "created",
                "type": "ref_plane_normal_at_distance_v2",
                "curve_edge_index": curve_edge_index,
                "distance": distance,
                "new_plane_index": ref_planes.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_ref_plane_normal_at_arc_ratio_v2(
        self,
        curve_edge_index: int,
        orientation_plane_index: int,
        ratio: float,
        normal_side: int = 2,
    ) -> dict[str, Any]:
        """
        Create a reference plane normal to a curve at an arc-length ratio.

        Uses RefPlanes.AddNormalToCurveAtArcLengthRatio(Curve, OrientationPlane,
        Ratio, normalOrientation, selectedCurveEnd).

        Args:
            curve_edge_index: 0-based edge index on the body to use as curve
            orientation_plane_index: 1-based index of the orientation reference plane
            ratio: Arc length ratio (0.0 to 1.0)
            normal_side: Normal orientation (1=igLeft, 2=igRight)

        Returns:
            Dict with status and new plane index
        """
        try:
            if ratio < 0.0 or ratio > 1.0:
                return {"error": f"Ratio must be between 0.0 and 1.0, got {ratio}"}

            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            body = model.Body
            edges = body.Edges(FaceQueryConstants.igQueryAll)

            if curve_edge_index < 0 or curve_edge_index >= edges.Count:
                return {
                    "error": f"Invalid curve_edge_index: {curve_edge_index}. "
                    f"Edge count: {edges.Count}"
                }

            ref_planes = doc.RefPlanes
            curve = edges.Item(curve_edge_index + 1)
            orient_plane = ref_planes.Item(orientation_plane_index)

            ref_planes.AddNormalToCurveAtArcLengthRatio(
                curve, orient_plane, ratio, normal_side, ReferenceElementConstants.igCurveEnd
            )

            return {
                "status": "created",
                "type": "ref_plane_normal_at_arc_ratio_v2",
                "curve_edge_index": curve_edge_index,
                "ratio": ratio,
                "new_plane_index": ref_planes.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_ref_plane_normal_at_distance_along_v2(
        self,
        curve_edge_index: int,
        orientation_plane_index: int,
        distance: float,
        normal_side: int = 2,
    ) -> dict[str, Any]:
        """
        Create a reference plane normal to a curve at a distance along the curve.

        Uses RefPlanes.AddNormalToCurveAtDistanceAlongCurve(Curve, OrientationPlane,
        Distance, normalOrientation, selectedCurveEnd).

        Args:
            curve_edge_index: 0-based edge index on the body to use as curve
            orientation_plane_index: 1-based index of the orientation reference plane
            distance: Distance along the curve in meters
            normal_side: Normal orientation (1=igLeft, 2=igRight)

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            body = model.Body
            edges = body.Edges(FaceQueryConstants.igQueryAll)

            if curve_edge_index < 0 or curve_edge_index >= edges.Count:
                return {
                    "error": f"Invalid curve_edge_index: {curve_edge_index}. "
                    f"Edge count: {edges.Count}"
                }

            ref_planes = doc.RefPlanes
            curve = edges.Item(curve_edge_index + 1)
            orient_plane = ref_planes.Item(orientation_plane_index)

            ref_planes.AddNormalToCurveAtDistanceAlongCurve(
                curve, orient_plane, distance, normal_side, ReferenceElementConstants.igCurveEnd
            )

            return {
                "status": "created",
                "type": "ref_plane_normal_at_distance_along_v2",
                "curve_edge_index": curve_edge_index,
                "distance": distance,
                "new_plane_index": ref_planes.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_ref_plane_tangent_parallel(
        self,
        face_index: int,
        parent_plane_index: int,
        normal_side: int = 2,
    ) -> dict[str, Any]:
        """
        Create a reference plane parallel to parent and tangent to a curved surface.

        Uses RefPlanes.AddParallelByTangent(Surface, ParentPlane, NormalSide).

        Args:
            face_index: 0-based face index of the curved surface
            parent_plane_index: 1-based index of the parent reference plane
            normal_side: Normal orientation (1=igLeft, 2=igRight)

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            ref_planes = doc.RefPlanes

            if parent_plane_index < 1 or parent_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid parent_plane_index: {parent_plane_index}. "
                    f"Count: {ref_planes.Count}"
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists."}

            model = models.Item(1)
            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)

            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face_index: {face_index}. Body has {faces.Count} faces."}

            parent_plane = ref_planes.Item(parent_plane_index)
            face = faces.Item(face_index + 1)

            ref_planes.AddParallelByTangent(face, parent_plane, normal_side)

            return {
                "status": "created",
                "type": "ref_plane_tangent_parallel",
                "face_index": face_index,
                "parent_plane_index": parent_plane_index,
                "new_plane_index": ref_planes.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
