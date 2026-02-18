"""Primitive solid feature operations (box, cylinder, sphere)."""

import traceback
from typing import Any

from ..constants import DirectionConstants


class PrimitiveMixin:
    """Mixin providing primitive solid creation methods."""

    def create_box_by_center(
        self,
        center_x: float,
        center_y: float,
        center_z: float,
        length: float,
        width: float,
        height: float,
        plane_index: int = 1,
    ) -> dict[str, Any]:
        """
        Create a box primitive by center point and dimensions.

        Args:
            center_x, center_y, center_z: Center point coordinates (meters)
            length: Length in meters (X direction)
            width: Width in meters (Y direction)
            height: Height in meters (Z direction)
            plane_index: Reference plane (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

        Returns:
            Dict with status and box info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            top_plane = self._get_ref_plane(doc, plane_index)

            # AddBoxByCenter: x, y, z, dWidth, dHeight, dAngle, dDepth, pPlane,
            #                  ExtentSide, vbKeyPointExtent, pKeyPointObj, pKeyPointFlags
            models.AddBoxByCenter(
                center_x,
                center_y,
                center_z,
                length,  # dWidth
                width,  # dHeight
                0,  # dAngle (rotation)
                height,  # dDepth
                top_plane,  # pPlane
                DirectionConstants.igRight,  # ExtentSide
                False,  # vbKeyPointExtent
                None,  # pKeyPointObj
                0,  # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "box",
                "method": "by_center",
                "center": [center_x, center_y, center_z],
                "dimensions": {"length": length, "width": width, "height": height},
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_box_by_two_points(
        self, x1: float, y1: float, z1: float, x2: float, y2: float, z2: float, plane_index: int = 1
    ) -> dict[str, Any]:
        """
        Create a box primitive by two opposite corners.

        Args:
            x1, y1, z1: First corner coordinates (meters)
            x2, y2, z2: Opposite corner coordinates (meters)
            plane_index: Reference plane (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

        Returns:
            Dict with status and box info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            top_plane = self._get_ref_plane(doc, plane_index)

            # Compute depth from z difference
            depth = abs(z2 - z1) if abs(z2 - z1) > 0 else abs(y2 - y1)

            # AddBoxByTwoPoints: x1, y1, z1, x2, y2, z2, dAngle, dDepth, pPlane,
            #                     ExtentSide, vbKeyPointExtent, pKeyPointObj, pKeyPointFlags
            models.AddBoxByTwoPoints(
                x1,
                y1,
                z1,
                x2,
                y2,
                z2,
                0,  # dAngle
                depth,  # dDepth
                top_plane,  # pPlane
                DirectionConstants.igRight,  # ExtentSide
                False,  # vbKeyPointExtent
                None,  # pKeyPointObj
                0,  # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "box",
                "method": "by_two_points",
                "corner1": [x1, y1, z1],
                "corner2": [x2, y2, z2],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_box_by_three_points(
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
        plane_index: int = 1,
    ) -> dict[str, Any]:
        """
        Create a box primitive by three points.

        Args:
            x1, y1, z1: First corner point (meters)
            x2, y2, z2: Second point defining width (meters)
            x3, y3, z3: Third point defining height (meters)
            plane_index: Reference plane (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

        Returns:
            Dict with status and box info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            top_plane = self._get_ref_plane(doc, plane_index)

            import math

            # Calculate depth from the three points
            dx = x2 - x1
            dy = y2 - y1
            depth = math.sqrt(dx * dx + dy * dy)
            if depth == 0:
                depth = 0.01  # fallback

            # AddBoxByThreePoints: x1,y1,z1, x2,y2,z2, x3,y3,z3, dDepth, pPlane,
            #                       ExtentSide, vbKeyPointExtent, pKeyPointObj, pKeyPointFlags
            models.AddBoxByThreePoints(
                x1,
                y1,
                z1,
                x2,
                y2,
                z2,
                x3,
                y3,
                z3,
                depth,  # dDepth
                top_plane,  # pPlane
                DirectionConstants.igRight,  # ExtentSide
                False,  # vbKeyPointExtent
                None,  # pKeyPointObj
                0,  # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "box",
                "method": "by_three_points",
                "point1": [x1, y1, z1],
                "point2": [x2, y2, z2],
                "point3": [x3, y3, z3],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_cylinder(
        self,
        base_center_x: float,
        base_center_y: float,
        base_center_z: float,
        radius: float,
        height: float,
        plane_index: int = 1,
    ) -> dict[str, Any]:
        """
        Create a cylinder primitive.

        Args:
            base_center_x, base_center_y, base_center_z: Base circle center (meters)
            radius: Cylinder radius (meters)
            height: Cylinder height (meters)
            plane_index: Reference plane (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

        Returns:
            Dict with status and cylinder info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            top_plane = self._get_ref_plane(doc, plane_index)

            # AddCylinderByCenterAndRadius: x, y, z, dRadius,
            # dDepth, pPlane, ExtentSide, vbKeyPointExtent,
            # pKeyPointObj, pKeyPointFlags
            models.AddCylinderByCenterAndRadius(
                base_center_x,
                base_center_y,
                base_center_z,
                radius,
                height,  # dDepth
                top_plane,  # pPlane
                DirectionConstants.igRight,  # ExtentSide
                False,  # vbKeyPointExtent
                None,  # pKeyPointObj
                0,  # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "cylinder",
                "base_center": [base_center_x, base_center_y, base_center_z],
                "radius": radius,
                "height": height,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_sphere(
        self, center_x: float, center_y: float, center_z: float, radius: float, plane_index: int = 1
    ) -> dict[str, Any]:
        """
        Create a sphere primitive.

        Args:
            center_x, center_y, center_z: Sphere center coordinates (meters)
            radius: Sphere radius (meters)
            plane_index: Reference plane (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

        Returns:
            Dict with status and sphere info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            top_plane = self._get_ref_plane(doc, plane_index)

            # AddSphereByCenterAndRadius: x, y, z, dRadius,
            # pPlane, ExtentSide, vbKeyPointExtent,
            # vbCreateLiveSection, pKeyPointObj, pKeyPointFlags
            models.AddSphereByCenterAndRadius(
                center_x,
                center_y,
                center_z,
                radius,
                top_plane,  # pPlane
                DirectionConstants.igRight,  # ExtentSide
                False,  # vbKeyPointExtent
                False,  # vbCreateLiveSection
                None,  # pKeyPointObj
                0,  # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "sphere",
                "center": [center_x, center_y, center_z],
                "radius": radius,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_box_cutout_by_two_points(
        self, x1: float, y1: float, z1: float, x2: float, y2: float, z2: float, plane_index: int = 1
    ) -> dict[str, Any]:
        """
        Create a box-shaped cutout (boolean subtract) by two opposite corners.

        Uses BoxFeatures.AddCutoutByTwoPoints with same params as AddBoxByTwoPoints.
        Requires an existing base feature to cut from.
        Type library: AddCutoutByTwoPoints(6x VT_R8, dAngle, dDepth, pPlane,
        ExtentSide, vbKeyPointExtent, pKeyPointObj, pKeyPointFlags).

        Args:
            x1, y1, z1: First corner coordinates (meters)
            x2, y2, z2: Opposite corner coordinates (meters)
            plane_index: Reference plane (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

        Returns:
            Dict with status and cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            top_plane = self._get_ref_plane(doc, plane_index)
            depth = abs(z2 - z1) if abs(z2 - z1) > 0 else abs(y2 - y1)

            # BoxFeatures is on the Models collection level
            box_features = models.BoxFeatures if hasattr(models, "BoxFeatures") else None
            if box_features is None:
                # Try via the model object
                model = models.Item(1)
                box_features = model.BoxFeatures if hasattr(model, "BoxFeatures") else None

            if box_features is None:
                return {"error": "BoxFeatures collection not accessible"}

            box_features.AddCutoutByTwoPoints(
                x1,
                y1,
                z1,
                x2,
                y2,
                z2,
                0,  # dAngle
                depth,  # dDepth
                top_plane,  # pPlane
                DirectionConstants.igRight,  # ExtentSide
                False,  # vbKeyPointExtent
                None,  # pKeyPointObj
                0,  # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "box_cutout",
                "method": "by_two_points",
                "corner1": [x1, y1, z1],
                "corner2": [x2, y2, z2],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_box_cutout_by_center(
        self,
        center_x: float,
        center_y: float,
        center_z: float,
        length: float,
        width: float,
        height: float,
        plane_index: int = 1,
    ) -> dict[str, Any]:
        """
        Create a box-shaped cutout by center point and dimensions.

        Removes a rectangular volume centered at the given point.
        Requires an existing base feature.

        Args:
            center_x, center_y, center_z: Center point coordinates (meters)
            length: Length in meters (X direction)
            width: Width in meters (Y direction)
            height: Height in meters (Z direction)
            plane_index: Reference plane (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

        Returns:
            Dict with status and cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            top_plane = self._get_ref_plane(doc, plane_index)

            box_features = models.BoxFeatures if hasattr(models, "BoxFeatures") else None
            if box_features is None:
                model = models.Item(1)
                box_features = model.BoxFeatures if hasattr(model, "BoxFeatures") else None

            if box_features is None:
                return {"error": "BoxFeatures collection not accessible"}

            box_features.AddCutoutByCenter(
                center_x,
                center_y,
                center_z,
                length,  # dWidth
                width,  # dHeight
                0,  # dAngle
                height,  # dDepth
                top_plane,  # pPlane
                DirectionConstants.igRight,  # ExtentSide
                False,  # vbKeyPointExtent
                None,  # pKeyPointObj
                0,  # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "box_cutout",
                "method": "by_center",
                "center": [center_x, center_y, center_z],
                "dimensions": {"length": length, "width": width, "height": height},
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_box_cutout_by_three_points(
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
        plane_index: int = 1,
    ) -> dict[str, Any]:
        """
        Create a box-shaped cutout by three points.

        Removes a rectangular volume defined by three corner points.
        Requires an existing base feature.

        Args:
            x1, y1, z1: First corner point (meters)
            x2, y2, z2: Second point defining width (meters)
            x3, y3, z3: Third point defining height (meters)
            plane_index: Reference plane (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

        Returns:
            Dict with status and cutout info
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            top_plane = self._get_ref_plane(doc, plane_index)

            dx = x2 - x1
            dy = y2 - y1
            depth = math.sqrt(dx * dx + dy * dy)
            if depth == 0:
                depth = 0.01

            box_features = models.BoxFeatures if hasattr(models, "BoxFeatures") else None
            if box_features is None:
                model = models.Item(1)
                box_features = model.BoxFeatures if hasattr(model, "BoxFeatures") else None

            if box_features is None:
                return {"error": "BoxFeatures collection not accessible"}

            box_features.AddCutoutByThreePoints(
                x1,
                y1,
                z1,
                x2,
                y2,
                z2,
                x3,
                y3,
                z3,
                depth,  # dDepth
                top_plane,  # pPlane
                DirectionConstants.igRight,  # ExtentSide
                False,  # vbKeyPointExtent
                None,  # pKeyPointObj
                0,  # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "box_cutout",
                "method": "by_three_points",
                "point1": [x1, y1, z1],
                "point2": [x2, y2, z2],
                "point3": [x3, y3, z3],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_cylinder_cutout(
        self,
        center_x: float,
        center_y: float,
        center_z: float,
        radius: float,
        height: float,
        plane_index: int = 1,
    ) -> dict[str, Any]:
        """
        Create a cylindrical cutout (boolean subtract) primitive.

        Uses CylinderFeatures.AddCutoutByCenterAndRadius.
        Requires an existing base feature to cut from.
        Type library: AddCutoutByCenterAndRadius(x, y, z, dRadius, dDepth, pPlane,
        ProfileSide, ExtentSide, vbKeyPointExtent, pKeyPointObj, pKeyPointFlags).

        Args:
            center_x, center_y, center_z: Center coordinates (meters)
            radius: Cylinder radius (meters)
            height: Cylinder/cut depth (meters)
            plane_index: Reference plane (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

        Returns:
            Dict with status and cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            top_plane = self._get_ref_plane(doc, plane_index)

            # CylinderFeatures collection - try on Models first, then model
            cyl_features = models.CylinderFeatures if hasattr(models, "CylinderFeatures") else None
            if cyl_features is None:
                model = models.Item(1)
                cyl_features = (
                    model.CylinderFeatures if hasattr(model, "CylinderFeatures") else None
                )

            if cyl_features is None:
                return {"error": "CylinderFeatures collection not accessible"}

            cyl_features.AddCutoutByCenterAndRadius(
                center_x,
                center_y,
                center_z,
                radius,
                height,  # dDepth
                top_plane,  # pPlane
                DirectionConstants.igLeft,  # ProfileSide (igLeft = inside profile = hole)
                DirectionConstants.igRight,  # ExtentSide
                False,  # vbKeyPointExtent
                None,  # pKeyPointObj
                0,  # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "cylinder_cutout",
                "center": [center_x, center_y, center_z],
                "radius": radius,
                "height": height,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_sphere_cutout(
        self, center_x: float, center_y: float, center_z: float, radius: float, plane_index: int = 1
    ) -> dict[str, Any]:
        """
        Create a spherical cutout (boolean subtract) primitive.

        Uses SphereFeatures.AddCutoutByCenterAndRadius.
        Requires an existing base feature to cut from.
        Type library: AddCutoutByCenterAndRadius(x, y, z, dRadius, pPlane,
        ProfileSide, ExtentSide, vbKeyPointExtent, vbCreateLiveSection,
        pKeyPointObj, pKeyPointFlags).

        Args:
            center_x, center_y, center_z: Sphere center coordinates (meters)
            radius: Sphere radius (meters)
            plane_index: Reference plane (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

        Returns:
            Dict with status and cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            top_plane = self._get_ref_plane(doc, plane_index)

            # SphereFeatures collection
            sph_features = models.SphereFeatures if hasattr(models, "SphereFeatures") else None
            if sph_features is None:
                model = models.Item(1)
                sph_features = model.SphereFeatures if hasattr(model, "SphereFeatures") else None

            if sph_features is None:
                return {"error": "SphereFeatures collection not accessible"}

            sph_features.AddCutoutByCenterAndRadius(
                center_x,
                center_y,
                center_z,
                radius,
                top_plane,  # pPlane
                DirectionConstants.igLeft,  # ProfileSide (igLeft = inside profile = hole)
                DirectionConstants.igRight,  # ExtentSide
                False,  # vbKeyPointExtent
                False,  # vbCreateLiveSection
                None,  # pKeyPointObj
                0,  # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "sphere_cutout",
                "center": [center_x, center_y, center_z],
                "radius": radius,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
