"""Physical properties, measurements, and body appearance operations."""

import math
from typing import Any

from solidedge_mcp.backends.errors import error_result

from ..comutil import com_get, owned_style_for
from ..logging import get_logger
from ._base import QueryManagerBase, all_faces, body_of, r8_array

_logger = get_logger(__name__)


class PhysicalPropsMixin(QueryManagerBase):
    """Mixin providing physical property queries and body appearance methods."""

    doc_manager: Any

    @staticmethod
    def _compute_physical_properties(model: Any, density: float, accuracy: float) -> Any:
        """Call Model.ComputePhysicalPropertiesWithSpecifiedDensity correctly.

        Part.tlb Model.ComputePhysicalPropertiesWithSpecifiedDensity(
            Density VT_R8 [in], Accuracy VT_R8 [in],
            Volume VT_R8* [out], Area VT_R8* [out], Mass VT_R8* [out],
            CenterOfGravity SAFEARRAY(VT_R8)* [in,out],
            CenterOfVolume SAFEARRAY(VT_R8)* [in,out],
            GlobalMomentsOfInteria SAFEARRAY(VT_R8)* [in,out],
            PrincipalMomentsOfInteria SAFEARRAY(VT_R8)* [in,out],
            PrincipalAxes SAFEARRAY(VT_R8)* [in,out],
            RadiiOfGyration SAFEARRAY(VT_R8)* [in,out],
            RelativeAccuracyAchieved VT_R8* [out], Status VT_INT* [out])

        The six SAFEARRAY parameters are [in,out]: pywin32 marshals them by
        reference and the caller must supply correctly sized buffers, so
        calling with just (density, accuracy) fails inside COM with
        "Parameter not optional" (0x8002000F). Volume/Area/Mass sit between
        Accuracy and the buffers, hence the keyword arguments — the parameter
        names are spelled exactly as the type library spells them, typo
        included. Returns the [out]/[in,out] values as a tuple:
        (volume, area, mass, cog, cov, global_moi, principal_moi,
        principal_axes, radii_of_gyration, relative_accuracy, status).
        """
        return model.ComputePhysicalPropertiesWithSpecifiedDensity(
            Density=density,
            Accuracy=accuracy,
            CenterOfGravity=r8_array(3),
            CenterOfVolume=r8_array(3),
            GlobalMomentsOfInteria=r8_array(6),
            PrincipalMomentsOfInteria=r8_array(3),
            PrincipalAxes=r8_array(9),
            RadiiOfGyration=r8_array(3),
        )

    def get_mass_properties(self, density: float = 7850) -> dict[str, Any]:
        """
        Get mass properties of the part.

        Uses Model.ComputePhysicalPropertiesWithSpecifiedDensity, which returns
        (volume, area, mass, cog, cov, global_moi, principal_moi,
        principal_axes, radii_of_gyration, relative_accuracy, status).

        Args:
            density: Material density in kg/m³ (default: 7850 for steel)

        Returns:
            Dict with volume, mass, surface area, center of gravity, moments of inertia
        """
        try:
            _logger.info(f"Computing mass properties with density={density} kg/m^3")
            doc, model = self._get_first_model()

            result = self._compute_physical_properties(model, density, 0.99)

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
            _logger.error(f"Mass properties computation failed: {e}")
            return error_result(e)

    def get_bounding_box(self) -> dict[str, Any]:
        """
        Get the bounding box of the model.

        geometry.tlb Body.GetRange(
            MinRangePoint SAFEARRAY(VT_R8)* [in,out],
            MaxRangePoint SAFEARRAY(VT_R8)* [in,out])

        Both parameters are [in,out], so the caller supplies the buffers and
        reads the filled points back out of the returned tuple
        ((min_x, min_y, min_z), (max_x, max_y, max_z)).

        Returns:
            Dict with min/max coordinates and dimensions
        """
        try:
            doc, model = self._get_first_model()

            body = body_of(model)
            range_data = body.GetRange(r8_array(3), r8_array(3))

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
            return error_result(e)

    def get_surface_area(self) -> dict[str, Any]:
        """
        Get the total surface area of the body.

        Returns:
            Dict with total surface area
        """
        try:
            doc, model = self._get_first_model()
            body = body_of(model)

            # Body has no SurfaceArea property, so that attempt always
            # raised and the sum below was the only path that ever ran.
            # Face.Area is real, and summing it gives the true area.

            faces = all_faces(body, model)
            total_area = 0.0
            for i in range(1, faces.Count + 1):
                try:
                    face = faces.Item(i)
                    total_area += face.Area
                except Exception:
                    pass

            if not faces.Count:
                return {
                    "error": ("This body reports no faces, so its surface area cannot be measured.")
                }

            return {
                "surface_area": total_area,
                "surface_area_mm2": total_area * 1e6,
                "face_count": faces.Count,
                "method": "sum_of_faces",
            }
        except Exception as e:
            return error_result(e)

    def get_volume(self) -> dict[str, Any]:
        """
        Get the volume of the body.

        Returns:
            Dict with body volume
        """
        try:
            doc, model = self._get_first_model()
            body = body_of(model)
            volume = body.Volume

            return {
                "volume": volume,
                "volume_mm3": volume * 1e9,  # m³ to mm³
                "volume_cm3": volume * 1e6,  # m³ to cm³
            }
        except Exception as e:
            return error_result(e)

    def get_face_area(self, face_index: int) -> dict[str, Any]:
        """
        Get the area of a specific face on the body.

        Args:
            face_index: 0-based index of the face

        Returns:
            Dict with face area in square meters
        """
        try:
            doc, model = self._get_first_model()
            body = body_of(model)
            faces = all_faces(body, model)

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
            return error_result(e)

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

                if cog_x is not None and cog_y is not None and cog_z is not None:
                    return {
                        "center_of_gravity": [cog_x, cog_y, cog_z],
                        "center_of_gravity_mm": [cog_x * 1000, cog_y * 1000, cog_z * 1000],
                    }
            except Exception:
                pass

            # Fallback: compute physical properties
            doc, model = self._get_first_model()
            result = self._compute_physical_properties(model, 7850.0, 0.001)
            # result[3] is the center of gravity tuple
            cog = result[3]
            return {"center_of_gravity": list(cog), "center_of_gravity_mm": [c * 1000 for c in cog]}
        except Exception as e:
            return error_result(e)

    def get_moments_of_inertia(self) -> dict[str, Any]:
        """
        Get the moments of inertia of the part.

        Returns:
            Dict with moments of inertia values
        """
        try:
            doc, model = self._get_first_model()
            result = self._compute_physical_properties(model, 7850.0, 0.001)
            # result: (volume, area, mass, cog, cov, global_moi, principal_moi,
            # principal_axes, radii_of_gyration, relative_accuracy, status)
            moi = result[5]
            principal_moi = result[6]

            return {
                "moments_of_inertia": list(moi) if hasattr(moi, "__iter__") else moi,
                "principal_moments": (
                    list(principal_moi) if hasattr(principal_moi, "__iter__") else principal_moi
                ),
            }
        except Exception as e:
            return error_result(e)

    def get_user_physical_properties(self, density: float = 7850.0) -> dict[str, Any]:
        """
        Get physical properties of the active part.

        Prefers user-overridden values from PartDocument.GetUserPhysicalProperties()
        (set via PutUserPhysicalProperties). On a part with no such overrides that
        COM call FAULTS, so we fall back to properties COMPUTED from geometry via
        ComputePhysicalPropertiesWithSpecifiedDensity (the same path the working
        solidedge://geometry/* resources use). The returned dict carries a
        ``source`` of "user_override" or "computed".

        Args:
            density: Density (kg/m³) used only for the computed fallback
                (default 7850 = steel). Ignored when user overrides exist.

        Returns:
            Dict with volume, area, mass, center of gravity, etc.
        """
        try:
            doc = self.doc_manager.get_active_document()
            try:
                result = doc.GetUserPhysicalProperties()
            except Exception as e:
                _logger.info(
                    f"GetUserPhysicalProperties unavailable ({e}); "
                    f"computing properties from geometry."
                )
                result = None

            # Result is a tuple: (Volume, Area, Mass, CoG[3], CoV[3],
            #   GlobalMOI[6], PrincipalMOI[3], PrincipalAxes[9], RadiiOfGyration[3])
            if isinstance(result, tuple) and len(result) >= 3:
                props: dict[str, Any] = {"status": "success", "source": "user_override"}
                props["volume"] = result[0]
                props["surface_area"] = result[1]
                props["mass"] = result[2]
                if len(result) > 3 and hasattr(result[3], "__iter__"):
                    props["center_of_gravity"] = list(result[3])[:3]
                if len(result) > 4 and hasattr(result[4], "__iter__"):
                    props["center_of_volume"] = list(result[4])[:3]
                return props

            # No usable user-set overrides -> compute from geometry.
            computed = self.get_mass_properties(density)
            if "error" not in computed:
                computed["source"] = "computed"
                computed["note"] = (
                    "No user-set physical properties were found; values were "
                    f"computed from geometry at density={density} kg/m³. Assign a "
                    "material/density for an exact mass."
                )
            return computed
        except Exception as e:
            return error_result(e)

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
            return error_result(e)

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
            return error_result(e)

    def set_body_color(self, red: int, green: int, blue: int) -> dict[str, Any]:
        """Set the body colour of the active part.

        ``Style.SetForegroundColor`` is in no Solid Edge type library, and
        ``Body.Style`` is None until a style is assigned, so this raised twice
        over. A body is coloured by assigning it a FaceStyle whose diffuse
        colour is what you want. Verified on Solid Edge 2026: creating a style
        with ``doc.FaceStyles.Add(name, "")``, calling ``SetDiffuse``, then
        assigning ``body.Style`` reads back the colour that was set.

        Args:
            red: Red channel, 0-255.
            green: Green channel, 0-255.
            blue: Blue channel, 0-255.

        Returns:
            Dict with status, the colour, and the style that carries it.
        """
        try:
            doc, model = self._get_first_model()
            body = body_of(model)

            red = max(0, min(255, red))
            green = max(0, min(255, green))
            blue = max(0, min(255, blue))

            style, err = owned_style_for(doc, body, com_get(body, "DisplayName", "Body") or "Body")
            if err:
                return err

            # SetDiffuse takes 0.0-1.0 per channel, not 0-255.
            style.SetDiffuse(red / 255.0, green / 255.0, blue / 255.0)

            return {
                "status": "set",
                "color": {"red": red, "green": green, "blue": blue},
                "hex": f"#{red:02x}{green:02x}{blue:02x}",
                "style": com_get(style, "StyleName", ""),
            }
        except Exception as e:
            return error_result(e)

    def get_body_color(self) -> dict[str, Any]:
        """Read the body colour of the active part.

        ``Style.ForegroundColor`` and ``Body.GetColor`` are in no Solid Edge
        type library, so both branches raised and this only ever returned
        "Could not determine body color". ``FaceStyle.GetDiffuse`` is the real
        accessor and reports each channel as 0.0-1.0.

        Returns:
            Dict with the colour as 0-255 channels and as hex, plus the
            opacity and reflectivity carried on the same style, or an error
            when the body has no style of its own.
        """
        try:
            _doc, model = self._get_first_model()
            body = body_of(model)

            style = com_get(body, "Style")
            if style is None:
                return {
                    "error": (
                        "This body has no style of its own, so it is drawn in the "
                        "document default colour. Set one with "
                        "set_appearance(target='body_color', ...)."
                    )
                }

            diffuse = style.GetDiffuse()
            red, green, blue = (int(round(float(channel) * 255)) for channel in diffuse[:3])

            return {
                "red": red,
                "green": green,
                "blue": blue,
                "hex": f"#{red:02x}{green:02x}{blue:02x}",
                "diffuse": [float(channel) for channel in diffuse[:3]],
                "opacity": com_get(style, "Opacity"),
                "reflectivity": com_get(style, "Reflectivity"),
            }
        except Exception as e:
            return error_result(e)

    def set_body_opacity(self, opacity: float) -> dict[str, Any]:
        """Set the body opacity.

        Opacity lives on the body's FaceStyle, alongside its colour and
        reflectivity; see ``comutil.owned_style_for``.

        Args:
            opacity: 0.0 (fully transparent) to 1.0 (fully opaque).

        Returns:
            Dict with status and the opacity Solid Edge reports afterwards.
        """
        try:
            doc, model = self._get_first_model()
            body = body_of(model)

            opacity = max(0.0, min(1.0, opacity))
            style, err = owned_style_for(doc, body, com_get(body, "DisplayName", "Body") or "Body")
            if err:
                return err
            style.Opacity = opacity

            return {
                "status": "set",
                "opacity": opacity,
                "reads_back": com_get(com_get(body, "Style"), "Opacity"),
            }
        except Exception as e:
            return error_result(e)

    def set_body_reflectivity(self, reflectivity: float) -> dict[str, Any]:
        """Set the body reflectivity.

        Reflectivity lives on the body's FaceStyle, alongside its colour and
        opacity; see ``comutil.owned_style_for``.

        Args:
            reflectivity: 0.0 to 1.0.

        Returns:
            Dict with status and the value Solid Edge reports afterwards.
        """
        try:
            doc, model = self._get_first_model()
            body = body_of(model)

            reflectivity = max(0.0, min(1.0, reflectivity))
            style, err = owned_style_for(doc, body, com_get(body, "DisplayName", "Body") or "Body")
            if err:
                return err
            style.Reflectivity = reflectivity

            return {
                "status": "set",
                "reflectivity": reflectivity,
                "reads_back": com_get(com_get(body, "Style"), "Reflectivity"),
            }
        except Exception as e:
            return error_result(e)

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
            result = self._compute_physical_properties(model, density, 0.99)

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
            return error_result(e)
