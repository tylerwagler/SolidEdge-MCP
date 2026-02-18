"""Physical properties, measurements, and body appearance operations."""

import math
import traceback
from typing import Any


class PhysicalPropsMixin:
    """Mixin providing physical property queries and body appearance methods."""

    def get_mass_properties(self, density: float = 7850) -> dict[str, Any]:
        """
        Get mass properties of the part.

        Uses Model.ComputePhysicalProperties(status, density, accuracy) which
        returns a tuple: (volume, area, mass, cog_tuple, cov_tuple, moi_tuple, ...)

        Args:
            density: Material density in kg/m³ (default: 7850 for steel)

        Returns:
            Dict with volume, mass, surface area, center of gravity, moments of inertia
        """
        try:
            doc, model = self._get_first_model()

            # ComputePhysicalPropertiesWithSpecifiedDensity(Density, Accuracy)
            # Returns tuple: (volume, area, mass, cog, cov, moi, principal_moi,
            #                  principal_axes, radii_of_gyration, ?, ?)
            result = model.ComputePhysicalPropertiesWithSpecifiedDensity(density, 0.99)

            volume = result[0] if len(result) > 0 else 0
            surface_area = result[1] if len(result) > 1 else 0
            mass_val = result[2] if len(result) > 2 else 0
            cog = result[3] if len(result) > 3 else (0, 0, 0)
            cov = result[4] if len(result) > 4 else (0, 0, 0)
            moi = result[5] if len(result) > 5 else (0, 0, 0, 0, 0, 0)
            principal_moi = result[6] if len(result) > 6 else (0, 0, 0)

            return {
                "status": "computed",
                "density": density,
                "volume": volume,
                "surface_area": surface_area,
                "mass": mass_val,
                "center_of_gravity": list(cog) if cog else [0, 0, 0],
                "center_of_volume": list(cov) if cov else [0, 0, 0],
                "moments_of_inertia": {
                    "Ixx": moi[0] if len(moi) > 0 else 0,
                    "Iyy": moi[1] if len(moi) > 1 else 0,
                    "Izz": moi[2] if len(moi) > 2 else 0,
                    "Ixy": moi[3] if len(moi) > 3 else 0,
                    "Ixz": moi[4] if len(moi) > 4 else 0,
                    "Iyz": moi[5] if len(moi) > 5 else 0,
                },
                "principal_moments": list(principal_moi) if principal_moi else [0, 0, 0],
                "units": {
                    "volume": "m³",
                    "surface_area": "m²",
                    "mass": "kg",
                    "density": "kg/m³",
                    "moments_of_inertia": "kg·m²",
                    "coordinates": "meters",
                },
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_bounding_box(self) -> dict[str, Any]:
        """
        Get the bounding box of the model.

        Uses Body.GetRange() which returns ((min_x, min_y, min_z), (max_x, max_y, max_z)).

        Returns:
            Dict with min/max coordinates and dimensions
        """
        try:
            doc, model = self._get_first_model()

            body = model.Body
            range_data = body.GetRange()

            min_pt = range_data[0]
            max_pt = range_data[1]

            return {
                "status": "computed",
                "min": list(min_pt),
                "max": list(max_pt),
                "dimensions": {
                    "x": max_pt[0] - min_pt[0],
                    "y": max_pt[1] - min_pt[1],
                    "z": max_pt[2] - min_pt[2],
                },
                "units": "meters",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_surface_area(self) -> dict[str, Any]:
        """
        Get the total surface area of the body.

        Returns:
            Dict with total surface area
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body

            # Try body.SurfaceArea first
            try:
                area = body.SurfaceArea
                return {"surface_area": area, "surface_area_mm2": area * 1e6}
            except Exception:
                pass

            # Fallback: sum face areas
            from ..constants import FaceQueryConstants

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            total_area = 0.0
            for i in range(1, faces.Count + 1):
                try:
                    face = faces.Item(i)
                    total_area += face.Area
                except Exception:
                    pass

            return {
                "surface_area": total_area,
                "surface_area_mm2": total_area * 1e6,
                "face_count": faces.Count,
                "method": "sum_of_faces",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_volume(self) -> dict[str, Any]:
        """
        Get the volume of the body.

        Returns:
            Dict with body volume
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body
            volume = body.Volume

            return {
                "volume": volume,
                "volume_mm3": volume * 1e9,  # m³ to mm³
                "volume_cm3": volume * 1e6,  # m³ to cm³
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_face_area(self, face_index: int) -> dict[str, Any]:
        """
        Get the area of a specific face on the body.

        Args:
            face_index: 0-based index of the face

        Returns:
            Dict with face area in square meters
        """
        try:
            from ..constants import FaceQueryConstants

            doc, model = self._get_first_model()
            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)

            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)

            area = face.Area

            return {
                "face_index": face_index,
                "area": area,
                "area_mm2": area * 1e6,  # Convert m² to mm²
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_center_of_gravity(self) -> dict[str, Any]:
        """
        Get the center of gravity (center of mass) of the part.

        Returns:
            Dict with CoG coordinates in meters and mm
        """
        try:
            doc = self.doc_manager.get_active_document()

            # Try using named variables first (most reliable)
            try:
                variables = doc.Variables
                cog_x = None
                cog_y = None
                cog_z = None

                for i in range(1, variables.Count + 1):
                    var = variables.Item(i)
                    name = var.Name
                    if name == "CoMX":
                        cog_x = var.Value
                    elif name == "CoMY":
                        cog_y = var.Value
                    elif name == "CoMZ":
                        cog_z = var.Value

                if cog_x is not None:
                    return {
                        "center_of_gravity": [cog_x, cog_y, cog_z],
                        "center_of_gravity_mm": [cog_x * 1000, cog_y * 1000, cog_z * 1000],
                    }
            except Exception:
                pass

            # Fallback: compute physical properties
            doc, model = self._get_first_model()
            result = model.ComputePhysicalPropertiesWithSpecifiedDensity(7850.0, 0.001)
            # result[3] is the center of gravity tuple
            cog = result[3]
            return {"center_of_gravity": list(cog), "center_of_gravity_mm": [c * 1000 for c in cog]}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_moments_of_inertia(self) -> dict[str, Any]:
        """
        Get the moments of inertia of the part.

        Returns:
            Dict with moments of inertia values
        """
        try:
            doc, model = self._get_first_model()
            result = model.ComputePhysicalPropertiesWithSpecifiedDensity(7850.0, 0.001)
            # result: (volume, area, mass, cog, cov, moi,
            # principal_moi, principal_axes,
            # radii_of_gyration, ?, ?)
            moi = result[5]
            principal_moi = result[6]

            return {
                "moments_of_inertia": list(moi) if hasattr(moi, "__iter__") else moi,
                "principal_moments": (
                    list(principal_moi) if hasattr(principal_moi, "__iter__") else principal_moi
                ),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_user_physical_properties(self) -> dict[str, Any]:
        """
        Get user-overridden physical properties from the active part document.

        Uses PartDocument.GetUserPhysicalProperties() which returns properties
        that have been manually set via PutUserPhysicalProperties.

        Returns:
            Dict with volume, area, mass, center of gravity, moments of inertia
        """
        try:
            doc = self.doc_manager.get_active_document()
            result = doc.GetUserPhysicalProperties()

            # Result is a tuple: (Volume, Area, Mass, CoG[3], CoV[3],
            #   GlobalMOI[6], PrincipalMOI[3], PrincipalAxes[9], RadiiOfGyration[3])
            props: dict[str, Any] = {"status": "success"}
            if isinstance(result, tuple) and len(result) >= 3:
                props["volume"] = result[0]
                props["surface_area"] = result[1]
                props["mass"] = result[2]
                if len(result) > 3:
                    cog = result[3]
                    if hasattr(cog, "__iter__"):
                        props["center_of_gravity"] = list(cog)[:3]
                if len(result) > 4:
                    cov = result[4]
                    if hasattr(cov, "__iter__"):
                        props["center_of_volume"] = list(cov)[:3]
            else:
                props["raw_result"] = str(result)

            return props
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def measure_distance(
        self, x1: float, y1: float, z1: float, x2: float, y2: float, z2: float
    ) -> dict[str, Any]:
        """
        Measure distance between two points.

        Args:
            x1, y1, z1: First point coordinates
            x2, y2, z2: Second point coordinates

        Returns:
            Dict with distance and components
        """
        try:
            dx = x2 - x1
            dy = y2 - y1
            dz = z2 - z1

            distance = math.sqrt(dx**2 + dy**2 + dz**2)

            return {
                "distance": distance,
                "delta": {"x": dx, "y": dy, "z": dz},
                "point1": [x1, y1, z1],
                "point2": [x2, y2, z2],
                "units": "meters",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def measure_angle(
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
        Measure the angle between three points (vertex at point 2).

        Calculates the angle formed by vectors P2->P1 and P2->P3.

        Args:
            x1, y1, z1: First point coordinates
            x2, y2, z2: Vertex point coordinates
            x3, y3, z3: Third point coordinates

        Returns:
            Dict with angle in degrees and radians
        """
        try:
            # Vector from P2 to P1
            v1 = (x1 - x2, y1 - y2, z1 - z2)
            # Vector from P2 to P3
            v2 = (x3 - x2, y3 - y2, z3 - z2)

            # Dot product
            dot = v1[0] * v2[0] + v1[1] * v2[1] + v1[2] * v2[2]

            # Magnitudes
            mag1 = math.sqrt(v1[0] ** 2 + v1[1] ** 2 + v1[2] ** 2)
            mag2 = math.sqrt(v2[0] ** 2 + v2[1] ** 2 + v2[2] ** 2)

            if mag1 == 0 or mag2 == 0:
                return {"error": "One or more vectors have zero length"}

            # Clamp for numerical stability
            cos_angle = max(-1.0, min(1.0, dot / (mag1 * mag2)))
            angle_rad = math.acos(cos_angle)
            angle_deg = math.degrees(angle_rad)

            return {"angle_degrees": angle_deg, "angle_radians": angle_rad, "vertex": [x2, y2, z2]}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_body_color(self, red: int, green: int, blue: int) -> dict[str, Any]:
        """
        Set the body color of the active part.

        Sets the foreground color of the body's style to the specified RGB values.
        Color values are 0-255 for each component.

        Args:
            red: Red component (0-255)
            green: Green component (0-255)
            blue: Blue component (0-255)

        Returns:
            Dict with status and color info
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body

            # Clamp values to 0-255
            red = max(0, min(255, red))
            green = max(0, min(255, green))
            blue = max(0, min(255, blue))

            style = body.Style
            style.SetForegroundColor(red, green, blue)

            return {
                "status": "set",
                "color": {"red": red, "green": green, "blue": blue},
                "hex": f"#{red:02x}{green:02x}{blue:02x}",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_body_color(self) -> dict[str, Any]:
        """
        Get the current body color.

        Returns:
            Dict with RGB color values
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body

            try:
                color = body.Style.ForegroundColor
                # Decompose OLE color
                red = color & 0xFF
                green = (color >> 8) & 0xFF
                blue = (color >> 16) & 0xFF
                return {"red": red, "green": green, "blue": blue, "ole_color": color}
            except Exception:
                # Try alternative
                try:
                    r, g, b = body.GetColor()
                    return {"red": r, "green": g, "blue": b}
                except Exception:
                    return {"error": "Could not determine body color"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_body_opacity(self, opacity: float) -> dict[str, Any]:
        """
        Set the body opacity (transparency).

        Uses model.Body.FaceStyle.Opacity.

        Args:
            opacity: Opacity value from 0.0 (fully transparent) to 1.0 (fully opaque)

        Returns:
            Dict with status
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body

            opacity = max(0.0, min(1.0, opacity))
            body.FaceStyle.Opacity = opacity

            return {"status": "set", "opacity": opacity}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_body_reflectivity(self, reflectivity: float) -> dict[str, Any]:
        """
        Set the body reflectivity.

        Uses model.Body.FaceStyle.Reflectivity.

        Args:
            reflectivity: Reflectivity value from 0.0 to 1.0

        Returns:
            Dict with status
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body

            reflectivity = max(0.0, min(1.0, reflectivity))
            body.FaceStyle.Reflectivity = reflectivity

            return {"status": "set", "reflectivity": reflectivity}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_material_density(self, density: float) -> dict[str, Any]:
        """
        Set the material density for mass property calculations.

        Stores the density value and recalculates mass properties.
        Default steel density is 7850 kg/m³.

        Args:
            density: Material density in kg/m³

        Returns:
            Dict with status and recalculated mass
        """
        try:
            doc, model = self._get_first_model()

            if density <= 0:
                return {"error": f"Density must be positive, got {density}"}

            # Recompute with new density
            result = model.ComputePhysicalPropertiesWithSpecifiedDensity(density, 0.99)

            mass = result[2] if len(result) > 2 else 0
            volume = result[0] if len(result) > 0 else 0

            return {
                "status": "computed",
                "density": density,
                "mass": mass,
                "volume": volume,
                "units": {"density": "kg/m³", "mass": "kg", "volume": "m³"},
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
