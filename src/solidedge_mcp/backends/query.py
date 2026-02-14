"""
Solid Edge Query and Inspection Operations

Handles querying model data, measurements, and properties.
"""

import contextlib
import math
import traceback
from typing import Any

import pythoncom
from win32com.client import VARIANT

from .constants import FaceQueryConstants, ModelingModeConstants


class QueryManager:
    """Manages query and inspection operations"""

    def __init__(self, document_manager):
        self.doc_manager = document_manager

    def _get_first_model(self):
        """Get the first model from the active document."""
        doc = self.doc_manager.get_active_document()
        if not hasattr(doc, "Models"):
            raise Exception("Document does not have a Models collection")
        models = doc.Models
        if models.Count == 0:
            raise Exception("No features in document")
        return doc, models.Item(1)

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

    def list_features(self) -> dict[str, Any]:
        """
        List all features in the active document.

        Uses Model.Features collection and DesignEdgebarFeatures for the feature tree.

        Returns:
            Dict with list of features
        """
        try:
            doc, model = self._get_first_model()

            features = []

            # Use DesignEdgebarFeatures for the full feature tree
            if hasattr(doc, "DesignEdgebarFeatures"):
                debf = doc.DesignEdgebarFeatures
                for i in range(1, debf.Count + 1):
                    try:
                        feat = debf.Item(i)
                        feat_info = {
                            "index": i - 1,
                            "name": feat.Name if hasattr(feat, "Name") else f"Feature_{i}",
                        }
                        features.append(feat_info)
                    except Exception:
                        features.append({"index": i - 1, "name": f"Feature_{i}"})
            else:
                # Fallback to Model.Features
                model_features = model.Features
                for i in range(1, model_features.Count + 1):
                    try:
                        feat = model_features.Item(i)
                        feat_info = {
                            "index": i - 1,
                            "name": feat.Name if hasattr(feat, "Name") else f"Feature_{i}",
                        }
                        features.append(feat_info)
                    except Exception:
                        features.append({"index": i - 1, "name": f"Feature_{i}"})

            return {"features": features, "count": len(features)}
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

    def get_document_properties(self) -> dict[str, Any]:
        """Get document properties and metadata"""
        try:
            doc = self.doc_manager.get_active_document()

            properties = {
                "name": doc.Name if hasattr(doc, "Name") else "Unknown",
                "path": doc.FullName if hasattr(doc, "FullName") else "Unsaved",
                "modified": not doc.Saved if hasattr(doc, "Saved") else False,
                "read_only": doc.ReadOnly if hasattr(doc, "ReadOnly") else False,
            }

            # Try to get summary info
            try:
                if hasattr(doc, "SummaryInfo"):
                    summary = doc.SummaryInfo
                    if hasattr(summary, "Title"):
                        properties["title"] = summary.Title
                    if hasattr(summary, "Author"):
                        properties["author"] = summary.Author
                    if hasattr(summary, "Subject"):
                        properties["subject"] = summary.Subject
                    if hasattr(summary, "Comments"):
                        properties["comments"] = summary.Comments
            except Exception:
                pass

            # Add body topology info
            try:
                models = doc.Models
                if models.Count > 0:
                    body = models.Item(1).Body
                    properties["volume_m3"] = body.Volume
            except Exception:
                pass

            return properties
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_feature_count(self) -> dict[str, Any]:
        """Get count of features in the document"""
        try:
            doc = self.doc_manager.get_active_document()

            counts = {}

            if hasattr(doc, "DesignEdgebarFeatures"):
                counts["features"] = doc.DesignEdgebarFeatures.Count

            if hasattr(doc, "Models"):
                counts["models"] = doc.Models.Count

            if hasattr(doc, "ProfileSets"):
                counts["sketches"] = doc.ProfileSets.Count

            if hasattr(doc, "RefPlanes"):
                counts["ref_planes"] = doc.RefPlanes.Count

            if hasattr(doc, "Variables"):
                counts["variables"] = doc.Variables.Count

            return counts
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # REFERENCE PLANES
    # =================================================================

    def get_ref_planes(self) -> dict[str, Any]:
        """
        List all reference planes in the active document.

        Default planes: 1=Top/XZ, 2=Front/XY, 3=Right/YZ.
        Additional planes are created with create_ref_plane_by_offset.

        Returns:
            Dict with list of reference planes and their indices
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "RefPlanes"):
                return {"error": "Document does not have reference planes"}

            ref_planes = doc.RefPlanes
            planes = []

            default_names = {1: "Top (XZ)", 2: "Front (XY)", 3: "Right (YZ)"}

            for i in range(1, ref_planes.Count + 1):
                try:
                    plane = ref_planes.Item(i)
                    plane_info = {
                        "index": i,
                        "is_default": i <= 3,
                    }

                    try:
                        plane_info["name"] = plane.Name
                    except Exception:
                        plane_info["name"] = default_names.get(i, f"RefPlane_{i}")

                    with contextlib.suppress(Exception):
                        plane_info["visible"] = plane.Visible

                    planes.append(plane_info)
                except Exception:
                    planes.append({"index": i, "name": default_names.get(i, f"RefPlane_{i}")})

            return {
                "planes": planes,
                "count": len(planes),
                "note": "Use plane index (1-based) in "
                "create_sketch_on_plane or "
                "create_ref_plane_by_offset",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # VARIABLES
    # =================================================================

    def get_variables(self) -> dict[str, Any]:
        """
        Get all variables from the active document.

        Queries the Variables collection using Query() to list all
        variable names, values, and formulas.

        Returns:
            Dict with list of variables
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            var_list = []
            for i in range(1, variables.Count + 1):
                try:
                    var = variables.Item(i)
                    var_info = {
                        "index": i - 1,
                        "name": var.DisplayName if hasattr(var, "DisplayName") else f"Var_{i}",
                    }
                    with contextlib.suppress(Exception):
                        var_info["value"] = var.Value
                    with contextlib.suppress(Exception):
                        var_info["formula"] = var.Formula
                    with contextlib.suppress(Exception):
                        var_info["units"] = var.Units
                    var_list.append(var_info)
                except Exception:
                    var_list.append({"index": i - 1, "name": f"Var_{i}"})

            return {"variables": var_list, "count": len(var_list)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_variable(self, name: str) -> dict[str, Any]:
        """
        Get a specific variable by name.

        Args:
            name: Variable display name (e.g., 'V1', 'Mass', 'Volume')

        Returns:
            Dict with variable value and info
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            # Search for the variable by display name
            for i in range(1, variables.Count + 1):
                try:
                    var = variables.Item(i)
                    display_name = var.DisplayName if hasattr(var, "DisplayName") else ""
                    if display_name == name:
                        result = {"name": name, "index": i - 1}
                        with contextlib.suppress(Exception):
                            result["value"] = var.Value
                        with contextlib.suppress(Exception):
                            result["formula"] = var.Formula
                        with contextlib.suppress(Exception):
                            result["units"] = var.Units
                        return result
                except Exception:
                    continue

            return {"error": f"Variable '{name}' not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_variable(self, name: str, value: float) -> dict[str, Any]:
        """
        Set a variable's value by name.

        Args:
            name: Variable display name
            value: New value to set

        Returns:
            Dict with status and updated value
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            for i in range(1, variables.Count + 1):
                try:
                    var = variables.Item(i)
                    display_name = var.DisplayName if hasattr(var, "DisplayName") else ""
                    if display_name == name:
                        old_value = var.Value
                        var.Value = value
                        return {
                            "status": "updated",
                            "name": name,
                            "old_value": old_value,
                            "new_value": value,
                        }
                except Exception:
                    continue

            return {"error": f"Variable '{name}' not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_variable(self, name: str, formula: str, units_type: str = None) -> dict[str, Any]:
        """
        Create a new user variable in the active document.

        Uses Variables.Add(pName, pFormula, [UnitsType]).
        Type library: Add(pName: VT_BSTR, pFormula: VT_BSTR, [UnitsType: VT_VARIANT]) -> variable*.

        Args:
            name: Variable name (e.g., 'MyWidth', 'BoltDiameter')
            formula: Variable formula/value as string (e.g., '0.025', 'V1 * 2')
            units_type: Optional units type string

        Returns:
            Dict with status and variable info
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            if units_type is not None:
                var = variables.Add(name, formula, units_type)
            else:
                var = variables.Add(name, formula)

            result = {"status": "created", "name": name, "formula": formula}
            with contextlib.suppress(Exception):
                result["value"] = var.Value
            with contextlib.suppress(Exception):
                result["display_name"] = var.DisplayName

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_variable_formula(self, name: str, formula: str) -> dict[str, Any]:
        """
        Set the formula of an existing variable by display name.

        Args:
            name: Variable display name
            formula: New formula string (e.g., '0.025', 'V1 * 2')

        Returns:
            Dict with old/new formula and current value
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            for i in range(1, variables.Count + 1):
                try:
                    var = variables.Item(i)
                    display_name = var.DisplayName if hasattr(var, "DisplayName") else ""
                    if display_name == name:
                        old_formula = ""
                        with contextlib.suppress(Exception):
                            old_formula = var.Formula
                        var.Formula = formula
                        result = {
                            "status": "updated",
                            "name": name,
                            "old_formula": old_formula,
                            "new_formula": formula,
                        }
                        with contextlib.suppress(Exception):
                            result["value"] = var.Value
                        return result
                except Exception:
                    continue

            return {"error": f"Variable '{name}' not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # CUSTOM PROPERTIES
    # =================================================================

    def get_custom_properties(self) -> dict[str, Any]:
        """
        Get all custom properties from the active document.

        Accesses the PropertySets collection to retrieve custom properties.

        Returns:
            Dict with list of custom properties (name/value pairs)
        """
        try:
            doc = self.doc_manager.get_active_document()
            prop_sets = doc.Properties

            properties = {}

            # Iterate through property sets to find Custom
            for ps_idx in range(1, prop_sets.Count + 1):
                try:
                    ps = prop_sets.Item(ps_idx)
                    ps_name = ps.Name if hasattr(ps, "Name") else f"Set_{ps_idx}"

                    props = {}
                    for p_idx in range(1, ps.Count + 1):
                        try:
                            prop = ps.Item(p_idx)
                            prop_name = prop.Name if hasattr(prop, "Name") else f"Prop_{p_idx}"
                            try:
                                props[prop_name] = prop.Value
                            except Exception:
                                props[prop_name] = None
                        except Exception:
                            continue

                    properties[ps_name] = props
                except Exception:
                    continue

            return {"property_sets": properties, "count": len(properties)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_custom_property(self, name: str, value: str) -> dict[str, Any]:
        """
        Set or create a custom property.

        Creates the property if it doesn't exist, updates it if it does.

        Args:
            name: Property name
            value: Property value (string)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            prop_sets = doc.Properties

            # Find "Custom" property set (typically the last one, index varies)
            custom_ps = None
            for ps_idx in range(1, prop_sets.Count + 1):
                try:
                    ps = prop_sets.Item(ps_idx)
                    if hasattr(ps, "Name") and ps.Name == "Custom":
                        custom_ps = ps
                        break
                except Exception:
                    continue

            if custom_ps is None:
                return {"error": "Custom property set not found"}

            # Check if property exists
            for p_idx in range(1, custom_ps.Count + 1):
                try:
                    prop = custom_ps.Item(p_idx)
                    if hasattr(prop, "Name") and prop.Name == name:
                        old_value = prop.Value
                        prop.Value = value
                        return {
                            "status": "updated",
                            "name": name,
                            "old_value": old_value,
                            "new_value": value,
                        }
                except Exception:
                    continue

            # Property doesn't exist, add it
            custom_ps.Add(name, value)
            return {"status": "created", "name": name, "value": value}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def delete_custom_property(self, name: str) -> dict[str, Any]:
        """
        Delete a custom property by name.

        Args:
            name: Property name to delete

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            prop_sets = doc.Properties

            # Find "Custom" property set
            for ps_idx in range(1, prop_sets.Count + 1):
                try:
                    ps = prop_sets.Item(ps_idx)
                    if hasattr(ps, "Name") and ps.Name == "Custom":
                        for p_idx in range(1, ps.Count + 1):
                            try:
                                prop = ps.Item(p_idx)
                                if hasattr(prop, "Name") and prop.Name == name:
                                    prop.Delete()
                                    return {"status": "deleted", "name": name}
                            except Exception:
                                continue
                        return {"error": f"Property '{name}' not found in Custom set"}
                except Exception:
                    continue

            return {"error": "Custom property set not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # BODY TOPOLOGY QUERIES
    # =================================================================

    def get_body_faces(self) -> dict[str, Any]:
        """
        Get all faces on the model body.

        Uses Body.Faces(igQueryAll=1) to enumerate all faces with their
        geometry type, area, and edge count.

        Face geometry types are determined by checking which query type
        each face belongs to (plane, cylinder, cone, sphere, torus, spline).

        Returns:
            Dict with list of faces and count
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body

            # Query type constants (from SE type library)
            query_types = {
                6: "plane",  # igQueryPlane
                10: "cylinder",  # igQueryCylinder
                7: "cone",  # igQueryCone
                9: "sphere",  # igQuerySphere
                8: "torus",  # igQueryTorus
                5: "spline",  # igQuerySpline
            }

            # Build a set of face indices per geometry type
            geo_type_map = {}  # face_area -> geometry_type (approximate matching)
            for qval, qname in query_types.items():
                try:
                    typed_faces = body.Faces(qval)
                    for j in range(1, typed_faces.Count + 1):
                        try:
                            tf = typed_faces.Item(j)
                            # Use (area, edge_count) as a fingerprint
                            area = tf.Area
                            ec = tf.Edges.Count if hasattr(tf.Edges, "Count") else 0
                            key = (round(area, 12), ec)
                            geo_type_map[key] = qname
                        except Exception:
                            pass
                except Exception:
                    pass

            # Now enumerate all faces
            faces = body.Faces(FaceQueryConstants.igQueryAll)
            face_list = []

            for i in range(1, faces.Count + 1):
                try:
                    face = faces.Item(i)
                    face_info = {"index": i - 1}
                    with contextlib.suppress(Exception):
                        face_info["area"] = face.Area
                    try:
                        edge_count = face.Edges.Count if hasattr(face.Edges, "Count") else 0
                        face_info["edge_count"] = edge_count
                    except Exception:
                        pass
                    # Look up geometry type
                    try:
                        key = (round(face.Area, 12), face_info.get("edge_count", 0))
                        face_info["geometry"] = geo_type_map.get(key, "unknown")
                    except Exception:
                        face_info["geometry"] = "unknown"
                    face_list.append(face_info)
                except Exception:
                    face_list.append({"index": i - 1})

            return {"faces": face_list, "count": len(face_list)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_body_edges(self) -> dict[str, Any]:
        """
        Get all unique edges on the model body.

        Enumerates edges via faces since Body.Edges() doesn't work
        in COM late binding. Deduplicates by collecting from all faces.

        Returns:
            Dict with edge count and face-edge mapping
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            total_edges = 0
            face_edges = []

            for fi in range(1, faces.Count + 1):
                try:
                    face = faces.Item(fi)
                    edges = face.Edges
                    edge_count = edges.Count if hasattr(edges, "Count") else 0
                    total_edges += edge_count
                    face_edges.append({"face_index": fi - 1, "edge_count": edge_count})
                except Exception:
                    face_edges.append({"face_index": fi - 1, "edge_count": 0})

            return {
                "face_edges": face_edges,
                "total_face_count": faces.Count,
                "total_edge_references": total_edges,
                "note": "Edge count includes shared edges (counted once per face)",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_face_info(self, face_index: int) -> dict[str, Any]:
        """
        Get detailed information about a specific face.

        Args:
            face_index: 0-based face index

        Returns:
            Dict with face type, area, edge count, and vertex count
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Count: {faces.Count}"}

            face = faces.Item(face_index + 1)

            info = {"index": face_index}

            with contextlib.suppress(Exception):
                info["type"] = face.Type
            with contextlib.suppress(Exception):
                info["area"] = face.Area
            try:
                edges = face.Edges
                info["edge_count"] = edges.Count if hasattr(edges, "Count") else 0
            except Exception:
                pass
            try:
                vertices = face.Vertices
                info["vertex_count"] = vertices.Count if hasattr(vertices, "Count") else 0
            except Exception:
                pass

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # PERFORMANCE & RECOMPUTE
    # =================================================================

    def get_body_facet_data(self, tolerance: float = 0.0) -> dict[str, Any]:
        """
        Get tessellation/mesh data from the model body.

        Returns triangulated facet data (vertices, normals, face IDs).
        Useful for 3D printing previews and mesh export.

        Args:
            tolerance: Mesh tolerance in meters. If <= 0, returns cached data.
                       If > 0, recomputes from Parasolid (slower but more accurate).

        Returns:
            Dict with facet count, point count, and sample data
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No geometry in document"}

            model = models.Item(1)
            body = model.Body

            import array as arr_mod

            # Prepare out parameters for COM call
            # GetFacetData(Tolerance, FacetCount, Points,
            # Normals, TextureCoords, StyleIDs, FaceIDs,
            # bHonourPrefs)
            points = arr_mod.array("d", [])
            normals = arr_mod.array("d", [])
            texture_coords = arr_mod.array("d", [])
            style_ids = arr_mod.array("i", [])
            face_ids = arr_mod.array("i", [])

            try:
                result_data = body.GetFacetData(
                    tolerance,  # Tolerance
                )

                # GetFacetData returns a tuple of
                # (facetCount, points, normals,
                # textureCoords, styleIds, faceIds)
                if isinstance(result_data, tuple) and len(result_data) >= 2:
                    facet_count = result_data[0] if isinstance(result_data[0], int) else 0
                    pts = result_data[1] if len(result_data) > 1 else []

                    return {
                        "facet_count": facet_count,
                        "point_count": len(pts) // 3 if pts else 0,
                        "tolerance": tolerance,
                        "has_data": facet_count > 0,
                    }
            except Exception:
                pass

            # Alternative: try with explicit out params
            try:
                facet_count = 0
                body.GetFacetData(
                    tolerance,
                    facet_count,
                    points,
                    normals,
                    texture_coords,
                    style_ids,
                    face_ids,
                    False,
                )

                return {
                    "facet_count": facet_count,
                    "point_count": len(points) // 3 if points else 0,
                    "tolerance": tolerance,
                    "has_data": len(points) > 0,
                }
            except Exception as e2:
                return {
                    "error": f"GetFacetData failed: {e2}",
                    "note": "Body facet data may require "
                    "specific COM marshaling. "
                    "Try export_stl() instead.",
                    "traceback": traceback.format_exc(),
                }

        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_solid_bodies(self) -> dict[str, Any]:
        """
        Report all solid bodies in the active part document.

        Lists design bodies and construction bodies with their properties.

        Returns:
            Dict with body info (is_solid, shell count, etc.)
        """
        try:
            doc = self.doc_manager.get_active_document()

            bodies = []

            # Check design bodies (Models collection)
            models = doc.Models
            for i in range(1, models.Count + 1):
                model = models.Item(i)
                try:
                    body = model.Body
                    body_info = {
                        "index": i - 1,
                        "type": "design",
                        "name": model.Name if hasattr(model, "Name") else f"Model_{i}",
                    }

                    try:
                        body_info["is_solid"] = body.IsSolid
                    except Exception:
                        body_info["is_solid"] = True  # Default assumption

                    with contextlib.suppress(Exception):
                        body_info["volume"] = body.Volume

                    # Count shells
                    try:
                        shells = body.Shells
                        body_info["shell_count"] = shells.Count
                    except Exception:
                        pass

                    bodies.append(body_info)
                except Exception:
                    pass

            # Check construction bodies
            try:
                constructions = doc.Constructions
                for i in range(1, constructions.Count + 1):
                    try:
                        cm = constructions.Item(i)
                        body = cm.Body
                        body_info = {
                            "index": len(bodies),
                            "type": "construction",
                            "name": cm.Name if hasattr(cm, "Name") else f"Construction_{i}",
                        }
                        with contextlib.suppress(Exception):
                            body_info["is_solid"] = body.IsSolid
                        bodies.append(body_info)
                    except Exception:
                        pass
            except Exception:
                pass  # No Constructions collection

            return {"total_bodies": len(bodies), "bodies": bodies}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_modeling_mode(self) -> dict[str, Any]:
        """
        Get the current modeling mode (Ordered vs Synchronous).

        Returns:
            Dict with current mode
        """
        try:
            doc = self.doc_manager.get_active_document()

            try:
                mode = doc.ModelingMode
                if mode == ModelingModeConstants.seModelingModeSynchronous:
                    mode_name = "synchronous"
                elif mode == ModelingModeConstants.seModelingModeOrdered:
                    mode_name = "ordered"
                else:
                    mode_name = f"unknown ({mode})"
                return {"mode": mode_name, "mode_value": mode}
            except Exception:
                return {"error": "ModelingMode not available on this document type"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_modeling_mode(self, mode: str) -> dict[str, Any]:
        """
        Set the modeling mode (Ordered vs Synchronous).

        Args:
            mode: 'ordered' or 'synchronous'

        Returns:
            Dict with status and new mode
        """
        try:
            doc = self.doc_manager.get_active_document()

            mode_map = {
                "ordered": ModelingModeConstants.seModelingModeOrdered,
                "synchronous": ModelingModeConstants.seModelingModeSynchronous,
            }

            mode_value = mode_map.get(mode.lower())
            if mode_value is None:
                return {"error": f"Invalid mode: {mode}. Use 'ordered' or 'synchronous'"}

            try:
                old_mode = doc.ModelingMode
                doc.ModelingMode = mode_value
                new_mode = doc.ModelingMode
                return {
                    "status": "changed",
                    "old_mode": (
                        "ordered"
                        if old_mode == ModelingModeConstants.seModelingModeOrdered
                        else "synchronous"
                    ),
                    "new_mode": (
                        "ordered"
                        if new_mode == ModelingModeConstants.seModelingModeOrdered
                        else "synchronous"
                    ),
                }
            except Exception as e:
                return {"error": f"Cannot change modeling mode: {e}"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def suppress_feature(self, feature_name: str) -> dict[str, Any]:
        """
        Suppress a feature by name.

        Suppressed features are hidden and excluded from computation.

        Args:
            feature_name: Name of the feature to suppress

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            features = doc.DesignEdgebarFeatures

            for i in range(1, features.Count + 1):
                feat = features.Item(i)
                try:
                    if feat.Name == feature_name:
                        feat.Suppress()
                        return {"status": "suppressed", "feature": feature_name}
                except Exception:
                    continue

            return {"error": f"Feature '{feature_name}' not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def unsuppress_feature(self, feature_name: str) -> dict[str, Any]:
        """
        Unsuppress a feature by name.

        Args:
            feature_name: Name of the feature to unsuppress

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            features = doc.DesignEdgebarFeatures

            for i in range(1, features.Count + 1):
                feat = features.Item(i)
                try:
                    if feat.Name == feature_name:
                        feat.Unsuppress()
                        return {"status": "unsuppressed", "feature": feature_name}
                except Exception:
                    continue

            return {"error": f"Feature '{feature_name}' not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_select_set(self) -> dict[str, Any]:
        """
        Get the current selection set.

        Returns information about all currently selected objects.

        Returns:
            Dict with selected objects and count
        """
        try:
            doc = self.doc_manager.get_active_document()

            select_set = doc.SelectSet
            count = select_set.Count

            items = []
            for i in range(1, count + 1):
                try:
                    item = select_set.Item(i)
                    item_info = {"index": i - 1}

                    with contextlib.suppress(Exception):
                        item_info["type"] = str(type(item).__name__)

                    try:
                        if hasattr(item, "Name"):
                            item_info["name"] = item.Name
                    except Exception:
                        pass

                    items.append(item_info)
                except Exception:
                    items.append({"index": i - 1, "error": "could not read"})

            return {"count": count, "items": items}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def clear_select_set(self) -> dict[str, Any]:
        """
        Clear the current selection set.

        Removes all objects from the selection.

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            select_set = doc.SelectSet
            old_count = select_set.Count
            select_set.RemoveAll()

            return {"status": "cleared", "items_removed": old_count}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def select_add(self, object_type: str, index: int) -> dict[str, Any]:
        """
        Add an object to the selection set programmatically.

        Uses SelectSet.Add(Dispatch). Resolves the object from the specified
        type and index, then adds it to the current selection.
        Type library: SelectSet.Add(Dispatch: VT_DISPATCH) -> VT_VOID.

        Args:
            object_type: Type of object to select: 'feature', 'face', 'edge', 'plane'
            index: 0-based index of the object

        Returns:
            Dict with status and selection info
        """
        try:
            doc = self.doc_manager.get_active_document()
            select_set = doc.SelectSet

            obj = None
            if object_type == "feature":
                features = doc.DesignEdgebarFeatures
                if index < 0 or index >= features.Count:
                    return {"error": f"Invalid feature index: {index}. Count: {features.Count}"}
                obj = features.Item(index + 1)
            elif object_type == "face":
                models = doc.Models
                if models.Count == 0:
                    return {"error": "No model features exist"}
                model = models.Item(1)
                body = model.Body
                faces = body.Faces(FaceQueryConstants.igQueryAll)
                if index < 0 or index >= faces.Count:
                    return {"error": f"Invalid face index: {index}. Count: {faces.Count}"}
                obj = faces.Item(index + 1)
            elif object_type == "plane":
                ref_planes = doc.RefPlanes
                plane_idx = index + 1
                if plane_idx < 1 or plane_idx > ref_planes.Count:
                    return {"error": f"Invalid plane index: {index}. Count: {ref_planes.Count}"}
                obj = ref_planes.Item(plane_idx)
            else:
                return {
                    "error": f"Unsupported object type: "
                    f"{object_type}. Use 'feature', "
                    "'face', or 'plane'."
                }

            if obj is None:
                return {"error": "Could not resolve object to select"}

            select_set.Add(obj)

            return {
                "status": "added",
                "object_type": object_type,
                "index": index,
                "selection_count": select_set.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def select_remove(self, index: int) -> dict[str, Any]:
        """
        Remove an object from the selection set by index.

        Uses SelectSet.Remove(Index). Index is 1-based in COM.

        Args:
            index: 0-based index of the item to remove

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            select_set = doc.SelectSet

            com_index = index + 1
            if com_index < 1 or com_index > select_set.Count:
                return {"error": f"Invalid index: {index}. Selection has {select_set.Count} items."}

            select_set.Remove(com_index)

            return {"status": "removed", "index": index, "selection_count": select_set.Count}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def select_all(self) -> dict[str, Any]:
        """
        Select all objects in the active document.

        Uses SelectSet.AddAll() to add all selectable objects.

        Returns:
            Dict with status and new selection count
        """
        try:
            doc = self.doc_manager.get_active_document()
            select_set = doc.SelectSet
            select_set.AddAll()

            return {"status": "selected_all", "selection_count": select_set.Count}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def select_copy(self) -> dict[str, Any]:
        """
        Copy the current selection to the clipboard.

        Uses SelectSet.Copy().

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            select_set = doc.SelectSet

            if select_set.Count == 0:
                return {"error": "Nothing selected to copy"}

            select_set.Copy()

            return {"status": "copied", "items_copied": select_set.Count}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def select_cut(self) -> dict[str, Any]:
        """
        Cut the current selection to the clipboard.

        Uses SelectSet.Cut().

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            select_set = doc.SelectSet

            if select_set.Count == 0:
                return {"error": "Nothing selected to cut"}

            count = select_set.Count
            select_set.Cut()

            return {"status": "cut", "items_cut": count}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def select_delete(self) -> dict[str, Any]:
        """
        Delete the currently selected objects.

        Uses SelectSet.Delete().

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            select_set = doc.SelectSet

            if select_set.Count == 0:
                return {"error": "Nothing selected to delete"}

            count = select_set.Count
            select_set.Delete()

            return {"status": "deleted", "items_deleted": count}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def select_suspend_display(self) -> dict[str, Any]:
        """
        Suspend display updates for the selection set.

        Uses SelectSet.SuspendDisplay(). Call before batch selection
        changes to improve performance.

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            doc.SelectSet.SuspendDisplay()
            return {"status": "display_suspended"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def select_resume_display(self) -> dict[str, Any]:
        """
        Resume display updates for the selection set.

        Uses SelectSet.ResumeDisplay(). Call after SuspendDisplay
        to refresh the visual selection highlights.

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            doc.SelectSet.ResumeDisplay()
            return {"status": "display_resumed"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def select_refresh_display(self) -> dict[str, Any]:
        """
        Refresh the display of the selection set.

        Uses SelectSet.RefreshDisplay(). Forces a visual refresh
        of selection highlights without suspend/resume cycle.

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            doc.SelectSet.RefreshDisplay()
            return {"status": "display_refreshed"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # BODY APPEARANCE & MATERIAL
    # =================================================================

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

    def get_edge_count(self) -> dict[str, Any]:
        """
        Get total edge count on the model body.

        Quick count of all edges across all faces. Useful for
        determining if rounds/chamfers can be applied.

        Returns:
            Dict with total edge count and face count
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            total_edges = 0

            for fi in range(1, faces.Count + 1):
                try:
                    face = faces.Item(fi)
                    edges = face.Edges
                    if hasattr(edges, "Count"):
                        total_edges += edges.Count
                except Exception:
                    pass

            return {
                "total_edge_references": total_edges,
                "face_count": faces.Count,
                "note": "Shared edges are counted once per face",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def recompute(self) -> dict[str, Any]:
        """
        Recompute the active document and model.

        Forces recalculation of all features.

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            # Try model-level recompute first
            try:
                models = doc.Models
                if models.Count > 0:
                    model = models.Item(1)
                    model.Recompute()
            except Exception:
                pass

            # Also try document-level recompute
            with contextlib.suppress(Exception):
                doc.Recompute()

            return {"status": "recomputed"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_design_edgebar_features(self) -> dict[str, Any]:
        """
        Get the full feature tree from DesignEdgebarFeatures.

        Unlike list_features() which only shows Models, this shows the
        complete design tree including sketches, reference planes, etc.

        Returns:
            Dict with list of all feature tree entries
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "DesignEdgebarFeatures"):
                return {"error": "DesignEdgebarFeatures not available"}

            features = doc.DesignEdgebarFeatures
            feature_list = []

            for i in range(1, features.Count + 1):
                try:
                    feat = features.Item(i)
                    entry = {"index": i - 1}
                    try:
                        entry["name"] = feat.Name
                    except Exception:
                        entry["name"] = f"Feature_{i}"
                    with contextlib.suppress(Exception):
                        entry["type"] = feat.Type
                    with contextlib.suppress(Exception):
                        entry["suppressed"] = feat.IsSuppressed
                    feature_list.append(entry)
                except Exception:
                    feature_list.append({"index": i - 1, "name": f"Feature_{i}"})

            return {"features": feature_list, "count": len(feature_list)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def rename_feature(self, old_name: str, new_name: str) -> dict[str, Any]:
        """
        Rename a feature in the design tree.

        Args:
            old_name: Current feature name
            new_name: New feature name

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "DesignEdgebarFeatures"):
                return {"error": "DesignEdgebarFeatures not available"}

            features = doc.DesignEdgebarFeatures

            for i in range(1, features.Count + 1):
                try:
                    feat = features.Item(i)
                    if hasattr(feat, "Name") and feat.Name == old_name:
                        feat.Name = new_name
                        return {"status": "renamed", "old_name": old_name, "new_name": new_name}
                except Exception:
                    continue

            return {"error": f"Feature '{old_name}' not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_document_property(self, name: str, value: str) -> dict[str, Any]:
        """
        Set a summary/document property (Title, Subject, Author, etc.).

        Args:
            name: Property name (Title, Subject, Author, Manager, Company,
                  Category, Keywords, Comments)
            value: Property value

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            # SummaryInfo contains standard document properties
            if not hasattr(doc, "SummaryInfo"):
                return {"error": "SummaryInfo not available on this document"}

            summary = doc.SummaryInfo

            prop_map = {
                "Title": "Title",
                "Subject": "Subject",
                "Author": "Author",
                "Manager": "Manager",
                "Company": "Company",
                "Category": "Category",
                "Keywords": "Keywords",
                "Comments": "Comments",
            }

            prop_attr = prop_map.get(name)
            if prop_attr is None:
                return {"error": f"Unknown property: {name}. Valid: {', '.join(prop_map.keys())}"}

            setattr(summary, prop_attr, value)

            return {"status": "set", "property": name, "value": value}
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

    def get_face_count(self) -> dict[str, Any]:
        """
        Get the total number of faces on the body.

        Returns:
            Dict with face count
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)
            return {"face_count": faces.Count}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_edge_info(self, face_index: int, edge_index: int) -> dict[str, Any]:
        """
        Get information about a specific edge on a face.

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index within the face

        Returns:
            Dict with edge type, length, and vertex coordinates
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)

            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Count: {faces.Count}"}

            face = faces.Item(face_index + 1)
            edges = face.Edges

            if edge_index < 0 or edge_index >= edges.Count:
                return {"error": f"Invalid edge index: {edge_index}. Count: {edges.Count}"}

            edge = edges.Item(edge_index + 1)

            info = {
                "face_index": face_index,
                "edge_index": edge_index,
            }

            try:
                info["length"] = edge.Length
                info["length_mm"] = edge.Length * 1000
            except Exception:
                pass

            with contextlib.suppress(Exception):
                info["type"] = edge.Type

            try:
                start = edge.StartVertex
                end = edge.EndVertex
                info["start_vertex"] = [start.X, start.Y, start.Z]
                info["end_vertex"] = [end.X, end.Y, end.Z]
            except Exception:
                pass

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_face_color(self, face_index: int, red: int, green: int, blue: int) -> dict[str, Any]:
        """
        Set the color of a specific face.

        Args:
            face_index: 0-based face index
            red: Red component (0-255)
            green: Green component (0-255)
            blue: Blue component (0-255)

        Returns:
            Dict with status
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)

            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Count: {faces.Count}"}

            face = faces.Item(face_index + 1)

            # SetFaceStyle or put color directly
            try:
                face.SetColor(red, green, blue)
            except Exception:
                # Try alternative: FaceStyle
                try:
                    face.Color = (red << 0) | (green << 8) | (blue << 16)
                except Exception:
                    # Final fallback using OLE color
                    ole_color = red | (green << 8) | (blue << 16)
                    face.Color = ole_color

            return {"status": "updated", "face_index": face_index, "color": [red, green, blue]}
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

    def delete_feature(self, feature_name: str) -> dict[str, Any]:
        """
        Delete a feature by name.

        Finds the feature in the DesignEdgebarFeatures collection and deletes it.

        Args:
            feature_name: Name of the feature to delete

        Returns:
            Dict with deletion status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "DesignEdgebarFeatures"):
                return {"error": "Document does not support feature deletion"}

            debf = doc.DesignEdgebarFeatures
            for i in range(1, debf.Count + 1):
                try:
                    feat = debf.Item(i)
                    if hasattr(feat, "Name") and feat.Name == feature_name:
                        feat.Delete()
                        return {"status": "deleted", "feature_name": feature_name}
                except Exception:
                    continue

            return {"error": f"Feature '{feature_name}' not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_feature_status(self, feature_name: str) -> dict[str, Any]:
        """
        Get the status of a feature (OK, suppressed, failed, etc.).

        Looks up a feature by name and queries its status properties.

        Args:
            feature_name: Name of the feature

        Returns:
            Dict with feature status info
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "DesignEdgebarFeatures"):
                return {"error": "DesignEdgebarFeatures not available"}

            features = doc.DesignEdgebarFeatures
            for i in range(1, features.Count + 1):
                try:
                    feat = features.Item(i)
                    if hasattr(feat, "Name") and feat.Name == feature_name:
                        result = {"feature_name": feature_name, "index": i - 1}
                        with contextlib.suppress(Exception):
                            result["status"] = feat.Status
                        with contextlib.suppress(Exception):
                            result["is_suppressed"] = feat.IsSuppressed
                        try:
                            status_ex = feat.GetStatusEx()
                            result["status_ex"] = status_ex
                        except Exception:
                            pass
                        with contextlib.suppress(Exception):
                            result["type"] = feat.Type
                        return result
                except Exception:
                    continue

            return {"error": f"Feature '{feature_name}' not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_feature_profiles(self, feature_name: str) -> dict[str, Any]:
        """
        Get the sketch profiles associated with a feature.

        Uses Feature.GetProfiles() to find which sketches drive the feature.

        Args:
            feature_name: Name of the feature

        Returns:
            Dict with profile info
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "DesignEdgebarFeatures"):
                return {"error": "DesignEdgebarFeatures not available"}

            features = doc.DesignEdgebarFeatures
            target = None
            for i in range(1, features.Count + 1):
                try:
                    feat = features.Item(i)
                    if hasattr(feat, "Name") and feat.Name == feature_name:
                        target = feat
                        break
                except Exception:
                    continue

            if target is None:
                return {"error": f"Feature '{feature_name}' not found"}

            profiles = []
            try:
                result = target.GetProfiles()
                if result is not None:
                    if isinstance(result, tuple) and len(result) >= 2:
                        result[0]
                        profile_array = result[1]
                    else:
                        profile_array = result

                    if profile_array is not None and hasattr(profile_array, "__iter__"):
                        for p in profile_array:
                            p_info = {}
                            with contextlib.suppress(Exception):
                                p_info["name"] = p.Name
                            with contextlib.suppress(Exception):
                                p_info["status"] = p.Status
                            profiles.append(p_info)
            except Exception as e:
                return {
                    "feature_name": feature_name,
                    "profiles": [],
                    "note": f"GetProfiles not supported: {str(e)}",
                }

            return {"feature_name": feature_name, "profiles": profiles, "count": len(profiles)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_feature_parents(self, feature_name: str) -> dict[str, Any]:
        """
        Get the parent features of a named feature.

        Reads the .Parents collection from DesignEdgebarFeatures to find
        which features were used to create the named feature.

        Args:
            feature_name: Name of the feature to inspect

        Returns:
            Dict with list of parent feature names
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "DesignEdgebarFeatures"):
                return {"error": "DesignEdgebarFeatures not available"}

            features = doc.DesignEdgebarFeatures
            target = None
            for i in range(1, features.Count + 1):
                try:
                    feat = features.Item(i)
                    if hasattr(feat, "Name") and feat.Name == feature_name:
                        target = feat
                        break
                except Exception:
                    continue

            if target is None:
                return {"error": f"Feature '{feature_name}' not found"}

            parents = []
            try:
                parent_coll = target.Parents
                if parent_coll and hasattr(parent_coll, "Count"):
                    for j in range(1, parent_coll.Count + 1):
                        try:
                            parent = parent_coll.Item(j)
                            p_info = {"index": j - 1}
                            with contextlib.suppress(Exception):
                                p_info["name"] = parent.Name
                            parents.append(p_info)
                        except Exception:
                            parents.append({"index": j - 1, "name": "unknown"})
            except Exception as e:
                return {
                    "feature_name": feature_name,
                    "parents": [],
                    "note": f"Parents collection not available: {e}",
                }

            return {"feature_name": feature_name, "parents": parents, "count": len(parents)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_vertex_count(self) -> dict[str, Any]:
        """
        Get the total vertex count on the model body.

        Enumerates vertices via faces (same approach as edge counting).

        Returns:
            Dict with vertex count
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            total_vertices = 0

            for fi in range(1, faces.Count + 1):
                try:
                    face = faces.Item(fi)
                    vertices = face.Vertices
                    if hasattr(vertices, "Count"):
                        total_vertices += vertices.Count
                except Exception:
                    pass

            return {
                "total_vertex_references": total_vertices,
                "face_count": faces.Count,
                "note": "Shared vertices are counted once per face",
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

    def get_material_table(self) -> dict[str, Any]:
        """
        Get the available material properties from the document variables.

        Returns material-related variables from the active document.

        Returns:
            Dict with material properties
        """
        try:
            doc = self.doc_manager.get_active_document()

            material_vars = {}
            material_names = [
                "Density",
                "Mass",
                "Volume",
                "Surface_Area",
                "YoungsModulus",
                "PoissonsRatio",
                "ThermalExpansionCoefficient",
                "ThermalConductivity",
                "Material",
            ]

            try:
                variables = doc.Variables
                for i in range(1, variables.Count + 1):
                    try:
                        var = variables.Item(i)
                        name = var.Name
                        if name in material_names or name.startswith("Material"):
                            try:
                                material_vars[name] = var.Value
                            except Exception:
                                material_vars[name] = (
                                    str(var.Formula) if hasattr(var, "Formula") else "N/A"
                                )
                    except Exception:
                        continue
            except Exception:
                pass

            # Also try to get material name from properties
            try:
                if hasattr(doc, "Properties"):
                    props = doc.Properties
                    for i in range(1, props.Count + 1):
                        try:
                            prop_set = props.Item(i)
                            for j in range(1, prop_set.Count + 1):
                                try:
                                    prop = prop_set.Item(j)
                                    if "material" in prop.Name.lower():
                                        material_vars[prop.Name] = prop.Value
                                except Exception:
                                    continue
                        except Exception:
                            continue
            except Exception:
                pass

            return {"material_properties": material_vars, "property_count": len(material_vars)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_feature_dimensions(self, feature_name: str) -> dict[str, Any]:
        """
        Get the dimensions/parameters of a specific feature.

        Uses Feature.GetDimensions() to retrieve all dimension objects
        associated with a feature, then reads their names and values.

        Args:
            feature_name: Name of the feature (from list_features)

        Returns:
            Dict with feature dimensions (name, value, units)
        """
        try:
            doc = self.doc_manager.get_active_document()

            # Find the feature by name
            features = doc.DesignEdgebarFeatures
            target_feature = None
            for i in range(1, features.Count + 1):
                feat = features.Item(i)
                try:
                    if feat.Name == feature_name:
                        target_feature = feat
                        break
                except Exception:
                    continue

            if target_feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            dimensions = []
            try:
                # GetDimensions returns (count, dim_array) as out-params
                result = target_feature.GetDimensions()
                if result is not None:
                    # result may be a tuple (count, array) or just an array
                    if isinstance(result, tuple) and len(result) >= 2:
                        result[0]
                        dim_array = result[1]
                    else:
                        dim_array = result

                    # Process dimension objects
                    if dim_array is not None:
                        try:
                            # Try iterating as a collection
                            if hasattr(dim_array, "__iter__"):
                                for dim in dim_array:
                                    dim_info = {}
                                    with contextlib.suppress(Exception):
                                        dim_info["name"] = dim.Name
                                    with contextlib.suppress(Exception):
                                        dim_info["value"] = dim.Value
                                    with contextlib.suppress(Exception):
                                        dim_info["formula"] = dim.Formula
                                    dimensions.append(dim_info)
                        except Exception:
                            dimensions.append({"raw": str(dim_array)})
            except Exception as e:
                # Some features may not support GetDimensions
                return {
                    "feature_name": feature_name,
                    "dimensions": [],
                    "note": f"GetDimensions not supported: {str(e)}",
                }

            return {
                "feature_name": feature_name,
                "dimensions": dimensions,
                "count": len(dimensions),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_material_list(self) -> dict[str, Any]:
        """
        Get the list of available materials from the material table.

        Uses MatTable.GetMaterialList() via the Solid Edge application object.

        Returns:
            Dict with list of material names
        """
        try:
            self.doc_manager.get_active_document()

            # Get MatTable from the application
            app = self.doc_manager.connection.get_application()
            mat_table = app.GetMaterialTable()

            result = mat_table.GetMaterialList()
            # Returns (count, list_of_materials) as out-params
            if isinstance(result, tuple) and len(result) >= 2:
                result[0]
                material_list = result[1]
            else:
                material_list = result

            materials = []
            if material_list is not None:
                if isinstance(material_list, (list, tuple)):
                    materials = list(material_list)
                elif hasattr(material_list, "__iter__"):
                    materials = [str(m) for m in material_list]

            return {"materials": materials, "count": len(materials)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_material(self, material_name: str) -> dict[str, Any]:
        """
        Apply a named material to the active document.

        Uses MatTable.ApplyMaterial(pDocument, bstrMatName).

        Args:
            material_name: Name of the material (from get_material_list)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            app = self.doc_manager.connection.get_application()
            mat_table = app.GetMaterialTable()

            mat_table.ApplyMaterial(doc, material_name)

            return {"status": "applied", "material": material_name}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_material_property(self, material_name: str, property_index: int) -> dict[str, Any]:
        """
        Get a specific property value for a material.

        Uses MatTable.GetMatPropValue(bstrMatName, lPropIndex, varPropValue).

        Common property indices:
            0 = Density, 1 = Thermal Conductivity, 2 = Thermal Expansion,
            3 = Specific Heat, 4 = Young's Modulus, 5 = Poisson's Ratio,
            6 = Yield Stress, 7 = Ultimate Stress, 8 = Elongation

        Args:
            material_name: Name of the material
            property_index: Property index (see above)

        Returns:
            Dict with property value
        """
        try:
            app = self.doc_manager.connection.get_application()
            mat_table = app.GetMaterialTable()

            value = mat_table.GetMatPropValue(material_name, property_index)

            return {"material": material_name, "property_index": property_index, "value": value}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def query_variables(self, pattern: str = "*", case_insensitive: bool = True) -> dict[str, Any]:
        """
        Search variables by name pattern.

        Uses Variables.Query(pFindCriterium, NamedBy, VarType, CaseInsensitive)
        to find matching variables. Supports wildcards (* and ?).

        Args:
            pattern: Search pattern with wildcards (e.g., "*Length*", "V?")
            case_insensitive: Whether to ignore case (default True)

        Returns:
            Dict with matching variables
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            # Variables.Query returns a collection of matching variables
            # NamedBy: 0 = DisplayName, 1 = SystemName
            # VarType: 0 = All, 1 = Dimensions, 2 = UserVariables
            try:
                results = variables.Query(pattern, 0, 0, case_insensitive)
            except Exception:
                # Fallback: manual filtering if Query method not available
                matches = []
                import fnmatch

                for i in range(1, variables.Count + 1):
                    try:
                        var = variables.Item(i)
                        name = var.Name if hasattr(var, "Name") else str(i)
                        if case_insensitive:
                            match = fnmatch.fnmatch(name.lower(), pattern.lower())
                        else:
                            match = fnmatch.fnmatch(name, pattern)
                        if match:
                            entry = {"name": name}
                            with contextlib.suppress(Exception):
                                entry["value"] = var.Value
                            with contextlib.suppress(Exception):
                                entry["formula"] = var.Formula
                            matches.append(entry)
                    except Exception:
                        continue

                return {
                    "pattern": pattern,
                    "matches": matches,
                    "count": len(matches),
                    "method": "fallback_fnmatch",
                }

            # Process Query results
            matches = []
            if results is not None:
                try:
                    for i in range(1, results.Count + 1):
                        var = results.Item(i)
                        entry = {}
                        try:
                            entry["name"] = var.Name
                        except Exception:
                            entry["name"] = f"var_{i}"
                        with contextlib.suppress(Exception):
                            entry["value"] = var.Value
                        with contextlib.suppress(Exception):
                            entry["formula"] = var.Formula
                        matches.append(entry)
                except Exception:
                    pass

            return {"pattern": pattern, "matches": matches, "count": len(matches)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_variable_formula(self, name: str) -> dict[str, Any]:
        """
        Get the formula of a variable by name.

        Args:
            name: Variable display name

        Returns:
            Dict with variable formula string
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            for i in range(1, variables.Count + 1):
                try:
                    var = variables.Item(i)
                    display_name = var.DisplayName if hasattr(var, "DisplayName") else ""
                    if display_name == name:
                        result = {"name": name}
                        try:
                            result["formula"] = var.Formula
                        except Exception:
                            result["formula"] = None
                        with contextlib.suppress(Exception):
                            result["value"] = var.Value
                        return result
                except Exception:
                    continue

            return {"error": f"Variable '{name}' not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def rename_variable(self, old_name: str, new_name: str) -> dict[str, Any]:
        """
        Rename a variable by changing its DisplayName.

        Finds the variable by its current display name and sets a new one.

        Args:
            old_name: Current variable display name
            new_name: New display name

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            for i in range(1, variables.Count + 1):
                try:
                    var = variables.Item(i)
                    display_name = var.DisplayName if hasattr(var, "DisplayName") else ""
                    if display_name == old_name:
                        var.DisplayName = new_name
                        return {"status": "renamed", "old_name": old_name, "new_name": new_name}
                except Exception:
                    continue

            return {"error": f"Variable '{old_name}' not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_variable_names(self, name: str) -> dict[str, Any]:
        """
        Get both the DisplayName and SystemName of a variable.

        Useful for understanding how Solid Edge internally identifies
        variables vs. how they appear in the UI.

        Args:
            name: Variable display name to look up

        Returns:
            Dict with display_name and system_name
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            for i in range(1, variables.Count + 1):
                try:
                    var = variables.Item(i)
                    display_name = var.DisplayName if hasattr(var, "DisplayName") else ""
                    if display_name == name:
                        result = {"display_name": display_name}
                        try:
                            result["system_name"] = var.Name
                        except Exception:
                            result["system_name"] = None
                        with contextlib.suppress(Exception):
                            result["value"] = var.Value
                        return result
                except Exception:
                    continue

            return {"error": f"Variable '{name}' not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def translate_variable(self, name: str) -> dict[str, Any]:
        """
        Translate (look up) a variable by name using Variables.Translate().

        Returns the variable dispatch object's Name, Value, and Formula.

        Args:
            name: Variable name to translate

        Returns:
            Dict with variable info (name, value, formula)
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            var = variables.Translate(name)

            result = {"status": "success", "input_name": name}
            with contextlib.suppress(Exception):
                result["name"] = var.Name
            with contextlib.suppress(Exception):
                result["display_name"] = var.DisplayName
            with contextlib.suppress(Exception):
                result["value"] = var.Value
            with contextlib.suppress(Exception):
                result["formula"] = var.Formula
            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def copy_variable_to_clipboard(self, name: str) -> dict[str, Any]:
        """
        Copy a variable definition to the clipboard.

        Uses Variables.CopyToClipboard(name). The variable can then be pasted
        into another document using add_variable_from_clipboard.

        Args:
            name: Variable display name to copy

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables
            variables.CopyToClipboard(name)
            return {"status": "copied", "name": name}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_variable_from_clipboard(self, name: str, units_type: str = None) -> dict[str, Any]:
        """
        Add a variable from the clipboard.

        Uses Variables.AddFromClipboard(name) or AddFromClipboard(name, units).
        The variable must have been previously copied with copy_variable_to_clipboard.

        Args:
            name: Name for the new variable
            units_type: Optional units type string

        Returns:
            Dict with status and variable info
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            if units_type:
                new_var = variables.AddFromClipboard(name, units_type)
            else:
                new_var = variables.AddFromClipboard(name)

            result = {"status": "added", "name": name}
            with contextlib.suppress(Exception):
                result["value"] = new_var.Value
            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # LAYER MANAGEMENT
    # =================================================================

    def get_layers(self) -> dict[str, Any]:
        """
        Get all layers in the active document.

        Iterates doc.Layers and reports Name, Show, Locatable, IsEmpty.

        Returns:
            Dict with list of layers
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Layers"):
                return {"error": "Active document does not support layers"}

            layers_col = doc.Layers
            layers = []

            for i in range(1, layers_col.Count + 1):
                try:
                    layer = layers_col.Item(i)
                    info = {"index": i - 1}
                    try:
                        info["name"] = layer.Name
                    except Exception:
                        info["name"] = f"Layer_{i}"
                    with contextlib.suppress(Exception):
                        info["show"] = bool(layer.Show)
                    with contextlib.suppress(Exception):
                        info["locatable"] = bool(layer.Locatable)
                    with contextlib.suppress(Exception):
                        info["is_empty"] = bool(layer.IsEmpty)
                    layers.append(info)
                except Exception:
                    continue

            return {"layers": layers, "count": len(layers)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_layer(self, name: str) -> dict[str, Any]:
        """
        Add a new layer to the active document.

        Uses doc.Layers.Add(name).

        Args:
            name: Name for the new layer

        Returns:
            Dict with status and layer info
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Layers"):
                return {"error": "Active document does not support layers"}

            layers = doc.Layers
            layers.Add(name)

            return {"status": "added", "name": name, "total_layers": layers.Count}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def activate_layer(self, name_or_index) -> dict[str, Any]:
        """
        Activate a layer by name or 0-based index.

        Finds the layer and calls layer.Activate().

        Args:
            name_or_index: Layer name (str) or 0-based index (int)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Layers"):
                return {"error": "Active document does not support layers"}

            layers = doc.Layers

            if isinstance(name_or_index, int):
                if name_or_index < 0 or name_or_index >= layers.Count:
                    return {"error": f"Invalid layer index: {name_or_index}. Count: {layers.Count}"}
                layer = layers.Item(name_or_index + 1)
            else:
                # Search by name
                layer = None
                for i in range(1, layers.Count + 1):
                    try:
                        lyr = layers.Item(i)
                        if lyr.Name == name_or_index:
                            layer = lyr
                            break
                    except Exception:
                        continue
                if layer is None:
                    return {"error": f"Layer '{name_or_index}' not found"}

            layer.Activate()
            layer_name = layer.Name if hasattr(layer, "Name") else str(name_or_index)

            return {"status": "activated", "name": layer_name}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_layer_properties(
        self, name_or_index, show: bool = None, selectable: bool = None
    ) -> dict[str, Any]:
        """
        Set layer visibility and selectability properties.

        Args:
            name_or_index: Layer name (str) or 0-based index (int)
            show: If provided, set layer visibility
            selectable: If provided, set layer selectability (Locatable)

        Returns:
            Dict with status and updated properties
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Layers"):
                return {"error": "Active document does not support layers"}

            layers = doc.Layers

            if isinstance(name_or_index, int):
                if name_or_index < 0 or name_or_index >= layers.Count:
                    return {"error": f"Invalid layer index: {name_or_index}. Count: {layers.Count}"}
                layer = layers.Item(name_or_index + 1)
            else:
                layer = None
                for i in range(1, layers.Count + 1):
                    try:
                        lyr = layers.Item(i)
                        if lyr.Name == name_or_index:
                            layer = lyr
                            break
                    except Exception:
                        continue
                if layer is None:
                    return {"error": f"Layer '{name_or_index}' not found"}

            updated = {}
            if show is not None:
                layer.Show = show
                updated["show"] = show
            if selectable is not None:
                layer.Locatable = selectable
                updated["selectable"] = selectable

            layer_name = layer.Name if hasattr(layer, "Name") else str(name_or_index)

            return {"status": "updated", "name": layer_name, "properties": updated}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def delete_layer(self, name_or_index) -> dict[str, Any]:
        """
        Delete a layer from the active document.

        Cannot delete the active layer or a layer containing objects.

        Args:
            name_or_index: Layer name (str) or 0-based index (int)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Layers"):
                return {"error": "Active document does not support layers"}

            layers = doc.Layers

            if isinstance(name_or_index, int):
                if name_or_index < 0 or name_or_index >= layers.Count:
                    return {"error": f"Invalid layer index: {name_or_index}. Count: {layers.Count}"}
                layer = layers.Item(name_or_index + 1)
            else:
                layer = None
                for i in range(1, layers.Count + 1):
                    try:
                        lyr = layers.Item(i)
                        if lyr.Name == name_or_index:
                            layer = lyr
                            break
                    except Exception:
                        continue
                if layer is None:
                    return {"error": f"Layer '{name_or_index}' not found"}

            layer_name = layer.Name if hasattr(layer, "Name") else str(name_or_index)
            layer.Delete()

            return {"status": "deleted", "name": layer_name}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # FACESTYLE / APPEARANCE
    # =================================================================

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

    # =================================================================
    # FEATURE EDITING
    # =================================================================

    def _find_feature(self, feature_name: str):
        """Find a feature by name in DesignEdgebarFeatures. Returns (feature, doc) or raises."""
        doc = self.doc_manager.get_active_document()
        features = doc.DesignEdgebarFeatures
        for i in range(1, features.Count + 1):
            feat = features.Item(i)
            if hasattr(feat, "Name") and feat.Name == feature_name:
                return feat, doc
        return None, doc

    def get_direction1_extent(self, feature_name: str) -> dict[str, Any]:
        """
        Get Direction 1 extent data from a named feature.

        Calls feature.GetDirection1Extent() which returns extent type,
        distance, and optionally a face reference.

        Args:
            feature_name: Name of the feature in the design tree

        Returns:
            Dict with extent_type, distance, and feature_name
        """
        try:
            feature, doc = self._find_feature(feature_name)
            if feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            result = {"feature_name": feature_name}
            try:
                extent_data = feature.GetDirection1Extent()
                if isinstance(extent_data, tuple):
                    result["extent_type"] = extent_data[0] if len(extent_data) > 0 else None
                    result["distance"] = extent_data[1] if len(extent_data) > 1 else None
                    has_face = len(extent_data) > 2 and extent_data[2] is not None
                    result["face_ref"] = str(extent_data[2]) if has_face else None
                else:
                    result["extent_type"] = extent_data
            except Exception as e:
                result["error_detail"] = f"GetDirection1Extent failed: {e}"

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_direction1_extent(
        self, feature_name: str, extent_type: int, distance: float = 0.0
    ) -> dict[str, Any]:
        """
        Set Direction 1 extent on a named feature.

        Calls feature.ApplyDirection1Extent(extent_type, distance, None).
        Common extent types: igFinite=13, igThroughAll=16, igNone=44.

        Args:
            feature_name: Name of the feature in the design tree
            extent_type: Extent type constant (13=Finite, 16=ThroughAll, 44=None)
            distance: Extent distance in meters (used when extent_type is Finite)

        Returns:
            Dict with status
        """
        try:
            feature, doc = self._find_feature(feature_name)
            if feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            feature.ApplyDirection1Extent(extent_type, distance, None)

            return {
                "status": "updated",
                "feature_name": feature_name,
                "extent_type": extent_type,
                "distance": distance,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_direction2_extent(self, feature_name: str) -> dict[str, Any]:
        """
        Get Direction 2 extent data from a named feature.

        Calls feature.GetDirection2Extent() which returns extent type,
        distance, and optionally a face reference.

        Args:
            feature_name: Name of the feature in the design tree

        Returns:
            Dict with extent_type, distance, and feature_name
        """
        try:
            feature, doc = self._find_feature(feature_name)
            if feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            result = {"feature_name": feature_name}
            try:
                extent_data = feature.GetDirection2Extent()
                if isinstance(extent_data, tuple):
                    result["extent_type"] = extent_data[0] if len(extent_data) > 0 else None
                    result["distance"] = extent_data[1] if len(extent_data) > 1 else None
                    has_face = len(extent_data) > 2 and extent_data[2] is not None
                    result["face_ref"] = str(extent_data[2]) if has_face else None
                else:
                    result["extent_type"] = extent_data
            except Exception as e:
                result["error_detail"] = f"GetDirection2Extent failed: {e}"

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_direction2_extent(
        self, feature_name: str, extent_type: int, distance: float = 0.0
    ) -> dict[str, Any]:
        """
        Set Direction 2 extent on a named feature.

        Calls feature.ApplyDirection2Extent(extent_type, distance, None).
        Common extent types: igFinite=13, igThroughAll=16, igNone=44.

        Args:
            feature_name: Name of the feature in the design tree
            extent_type: Extent type constant (13=Finite, 16=ThroughAll, 44=None)
            distance: Extent distance in meters (used when extent_type is Finite)

        Returns:
            Dict with status
        """
        try:
            feature, doc = self._find_feature(feature_name)
            if feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            feature.ApplyDirection2Extent(extent_type, distance, None)

            return {
                "status": "updated",
                "feature_name": feature_name,
                "extent_type": extent_type,
                "distance": distance,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_thin_wall_options(self, feature_name: str) -> dict[str, Any]:
        """
        Get thin wall options from a named feature.

        Calls feature.GetThinWallOptions() which returns the wall type
        and thickness values.

        Args:
            feature_name: Name of the feature in the design tree

        Returns:
            Dict with wall_type, thickness1, thickness2
        """
        try:
            feature, doc = self._find_feature(feature_name)
            if feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            result = {"feature_name": feature_name}
            try:
                tw_data = feature.GetThinWallOptions()
                if isinstance(tw_data, tuple):
                    result["wall_type"] = tw_data[0] if len(tw_data) > 0 else None
                    result["thickness1"] = tw_data[1] if len(tw_data) > 1 else None
                    result["thickness2"] = tw_data[2] if len(tw_data) > 2 else None
                else:
                    result["wall_type"] = tw_data
            except Exception as e:
                result["error_detail"] = f"GetThinWallOptions failed: {e}"

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_thin_wall_options(
        self,
        feature_name: str,
        wall_type: int,
        thickness1: float,
        thickness2: float = 0.0,
    ) -> dict[str, Any]:
        """
        Set thin wall options on a named feature.

        Calls feature.SetThinWallOptions(wall_type, thickness1, thickness2).

        Args:
            feature_name: Name of the feature in the design tree
            wall_type: Thin wall type constant
            thickness1: First wall thickness in meters
            thickness2: Second wall thickness in meters (default 0.0)

        Returns:
            Dict with status
        """
        try:
            feature, doc = self._find_feature(feature_name)
            if feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            feature.SetThinWallOptions(wall_type, thickness1, thickness2)

            return {
                "status": "updated",
                "feature_name": feature_name,
                "wall_type": wall_type,
                "thickness1": thickness1,
                "thickness2": thickness2,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_from_face_offset(self, feature_name: str) -> dict[str, Any]:
        """
        Get the 'from face' offset data from a named feature.

        Calls feature.GetFromFaceOffsetData() which returns the offset
        distance and optionally a face reference.

        Args:
            feature_name: Name of the feature in the design tree

        Returns:
            Dict with offset distance and face reference info
        """
        try:
            feature, doc = self._find_feature(feature_name)
            if feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            result = {"feature_name": feature_name}
            try:
                offset_data = feature.GetFromFaceOffsetData()
                if isinstance(offset_data, tuple):
                    result["offset"] = offset_data[0] if len(offset_data) > 0 else None
                    has_face = len(offset_data) > 1 and offset_data[1] is not None
                    result["face_ref"] = str(offset_data[1]) if has_face else None
                else:
                    result["offset"] = offset_data
            except Exception as e:
                result["error_detail"] = f"GetFromFaceOffsetData failed: {e}"

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_from_face_offset(self, feature_name: str, offset: float) -> dict[str, Any]:
        """
        Set the 'from face' offset on a named feature.

        Calls feature.SetFromFaceOffsetData(offset).

        Args:
            feature_name: Name of the feature in the design tree
            offset: Offset distance in meters

        Returns:
            Dict with status
        """
        try:
            feature, doc = self._find_feature(feature_name)
            if feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            feature.SetFromFaceOffsetData(offset)

            return {
                "status": "updated",
                "feature_name": feature_name,
                "offset": offset,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_body_array(self, feature_name: str) -> dict[str, Any]:
        """
        Get the body array from a named feature.

        Calls feature.GetBodyArray() to retrieve the array of body
        references associated with a multi-body feature.

        Args:
            feature_name: Name of the feature in the design tree

        Returns:
            Dict with body count and body info
        """
        try:
            feature, doc = self._find_feature(feature_name)
            if feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            result = {"feature_name": feature_name}
            try:
                body_array = feature.GetBodyArray()
                bodies = []
                if body_array is not None:
                    if isinstance(body_array, (list, tuple)):
                        for idx, body in enumerate(body_array):
                            body_info = {"index": idx}
                            with contextlib.suppress(Exception):
                                body_info["name"] = body.Name
                            with contextlib.suppress(Exception):
                                body_info["volume"] = body.Volume
                            bodies.append(body_info)
                    else:
                        bodies.append({"index": 0, "raw": str(body_array)})
                result["bodies"] = bodies
                result["count"] = len(bodies)
            except Exception as e:
                result["error_detail"] = f"GetBodyArray failed: {e}"
                result["bodies"] = []
                result["count"] = 0

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_body_array(self, feature_name: str, body_indices: list[int]) -> dict[str, Any]:
        """
        Set the body array on a named feature.

        Resolves body objects from the model by index (0-based) and calls
        feature.SetBodyArray(body_array).

        Args:
            feature_name: Name of the feature in the design tree
            body_indices: List of 0-based body indices from the Models collection

        Returns:
            Dict with status
        """
        try:
            feature, doc = self._find_feature(feature_name)
            if feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            # Resolve body references from Models collection
            models = doc.Models
            body_array = []
            for idx in body_indices:
                com_idx = idx + 1  # Convert 0-based to 1-based
                if com_idx < 1 or com_idx > models.Count:
                    return {"error": f"Invalid body index: {idx}. Models count: {models.Count}"}
                model = models.Item(com_idx)
                body_array.append(model.Body)

            feature.SetBodyArray(body_array)

            return {
                "status": "updated",
                "feature_name": feature_name,
                "body_indices": body_indices,
                "body_count": len(body_array),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # MATERIAL LIBRARY (Batch 10)
    # =================================================================

    def get_material_library(self) -> dict[str, Any]:
        """
        Get the full material library from the active document.

        Iterates the material table to list all available materials
        with their names and density values.

        Returns:
            Dict with count and list of material info dicts
        """
        try:
            doc = self.doc_manager.get_active_document()

            mat_table = doc.GetMaterialTable()
            materials = []
            for i in range(1, mat_table.Count + 1):
                mat = mat_table.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    info["name"] = mat.Name
                with contextlib.suppress(Exception):
                    info["density"] = mat.Density
                with contextlib.suppress(Exception):
                    info["youngs_modulus"] = mat.YoungsModulus
                with contextlib.suppress(Exception):
                    info["poissons_ratio"] = mat.PoissonsRatio
                materials.append(info)
            return {"count": len(materials), "materials": materials}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_material_by_name(self, material_name: str) -> dict[str, Any]:
        """
        Look up a material by name in the material table and apply it.

        Searches the material table for the given name and applies it
        to the active document.

        Args:
            material_name: Name of the material to apply

        Returns:
            Dict with status and material info
        """
        try:
            doc = self.doc_manager.get_active_document()

            mat_table = doc.GetMaterialTable()

            # Search for the material by name
            found_mat = None
            for i in range(1, mat_table.Count + 1):
                mat = mat_table.Item(i)
                try:
                    if mat.Name == material_name:
                        found_mat = mat
                        break
                except Exception:
                    continue

            if found_mat is None:
                # Collect available names for error message
                available = []
                for i in range(1, mat_table.Count + 1):
                    with contextlib.suppress(Exception):
                        available.append(mat_table.Item(i).Name)
                return {
                    "error": f"Material '{material_name}' not found",
                    "available": available[:20],
                }

            # Apply the material
            try:
                doc.ApplyStyle(found_mat)
            except Exception:
                # Alternative: try setting via material name property
                try:
                    doc.Material = material_name
                except Exception:
                    # Another fallback
                    found_mat.Apply()

            result: dict[str, Any] = {"status": "applied", "material": material_name}
            with contextlib.suppress(Exception):
                result["density"] = found_mat.Density
            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # HELPERS for B-Rep topology queries
    # =================================================================

    def _get_body(self):
        """Get the body from the first model of the active document."""
        doc, model = self._get_first_model()
        return doc, model, model.Body

    def _get_face(self, face_index: int):
        """Get a specific face by 0-based index. Returns (doc, model, body, face)."""
        doc, model, body = self._get_body()
        faces = body.Faces(FaceQueryConstants.igQueryAll)
        if face_index < 0 or face_index >= faces.Count:
            raise IndexError(f"Invalid face index: {face_index}. Body has {faces.Count} faces.")
        face = faces.Item(face_index + 1)
        return doc, model, body, face

    def _get_face_edge(self, face_index: int, edge_index: int):
        """Get a specific edge on a face. Returns (doc, model, body, face, edge)."""
        doc, model, body, face = self._get_face(face_index)
        edges = face.Edges
        if edge_index < 0 or edge_index >= edges.Count:
            raise IndexError(f"Invalid edge index: {edge_index}. Face has {edges.Count} edges.")
        edge = edges.Item(edge_index + 1)
        return doc, model, body, face, edge

    @staticmethod
    def _to_list(val):
        """Convert a COM value to a Python list (handles iterables and scalars)."""
        if hasattr(val, "__iter__") and not isinstance(val, str):
            return list(val)
        return [val]

    # =================================================================
    # B-REP TOPOLOGY QUERIES (18 tools)
    # =================================================================

    def get_edge_endpoints(self, face_index: int, edge_index: int) -> dict[str, Any]:
        """
        Get the start and end XYZ coordinates of an edge.

        Uses edge.GetEndPoints() with VARIANT SAFEARRAY out params.

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index within the face

        Returns:
            Dict with start and end coordinates as [x, y, z] lists
        """
        try:
            _doc, _model, _body, _face, edge = self._get_face_edge(face_index, edge_index)

            start_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])
            end_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])
            result = edge.GetEndPoints(start_arr, end_arr)

            # GetEndPoints returns (start_arr, end_arr) as a tuple
            if isinstance(result, tuple) and len(result) >= 2:
                start = list(result[0])
                end = list(result[1])
            else:
                start = list(start_arr)
                end = list(end_arr)

            return {
                "start": start,
                "end": end,
                "face_index": face_index,
                "edge_index": edge_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_edge_length(self, face_index: int, edge_index: int) -> dict[str, Any]:
        """
        Get the total length of an edge.

        Uses edge.GetParamExtents() then edge.GetLengthAtParam().

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index within the face

        Returns:
            Dict with edge length in meters
        """
        try:
            _doc, _model, _body, _face, edge = self._get_face_edge(face_index, edge_index)

            # Try direct .Length property first (simpler)
            try:
                length = edge.Length
                return {
                    "length": length,
                    "length_mm": length * 1000.0,
                    "face_index": face_index,
                    "edge_index": edge_index,
                }
            except Exception:
                pass

            # Fallback: parametric approach
            param_result = edge.GetParamExtents()
            if isinstance(param_result, tuple) and len(param_result) >= 2:
                min_p = param_result[0]
                max_p = param_result[1]
            else:
                min_p = 0.0
                max_p = 1.0

            length_result = edge.GetLengthAtParam(min_p, max_p)
            length = length_result[0] if isinstance(length_result, tuple) else float(length_result)

            return {
                "length": length,
                "length_mm": length * 1000.0,
                "face_index": face_index,
                "edge_index": edge_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_edge_tangent(
        self, face_index: int, edge_index: int, param: float = 0.5
    ) -> dict[str, Any]:
        """
        Get the tangent vector at a parameter on an edge.

        Uses edge.GetTangent(numParams, paramsArray, tangentsArray).

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index within the face
            param: Parameter value (0.0 to 1.0, default 0.5 = midpoint)

        Returns:
            Dict with tangent vector [tx, ty, tz] and parameter value
        """
        try:
            _doc, _model, _body, _face, edge = self._get_face_edge(face_index, edge_index)

            params_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [param])
            tangents_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])

            result = edge.GetTangent(1, params_arr, tangents_arr)

            if isinstance(result, tuple) and len(result) >= 2:
                tangent = list(result[-1])
            else:
                tangent = list(tangents_arr)

            return {
                "tangent": tangent[:3],
                "param": param,
                "face_index": face_index,
                "edge_index": edge_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_edge_geometry(self, face_index: int, edge_index: int) -> dict[str, Any]:
        """
        Get the underlying geometry type and data of an edge.

        Inspects edge.Geometry to determine if the edge is a Line,
        Circle, Ellipse, or BSplineCurve, and extracts relevant data.

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index within the face

        Returns:
            Dict with geometry_type and associated data
        """
        try:
            _doc, _model, _body, _face, edge = self._get_face_edge(face_index, edge_index)

            geom = edge.Geometry
            result: dict[str, Any] = {
                "face_index": face_index,
                "edge_index": edge_index,
            }

            # Try to determine geometry type
            geom_type = "Unknown"
            with contextlib.suppress(Exception):
                geom_type_val = geom.Type
                geom_type = str(geom_type_val)

            # Attempt Circle data
            try:
                circle_data = geom.GetCircleData()
                if isinstance(circle_data, tuple) and len(circle_data) >= 3:
                    result["geometry_type"] = "Circle"
                    result["center"] = self._to_list(circle_data[0])
                    result["axis"] = self._to_list(circle_data[1])
                    result["radius"] = circle_data[2]
                    return result
            except Exception:
                pass

            # Attempt Ellipse data
            try:
                ellipse_data = geom.GetEllipseData()
                if isinstance(ellipse_data, tuple) and len(ellipse_data) >= 4:
                    result["geometry_type"] = "Ellipse"
                    result["center"] = self._to_list(ellipse_data[0])
                    result["axis"] = self._to_list(ellipse_data[1])
                    result["major_axis"] = self._to_list(ellipse_data[2])
                    result["minor_major_ratio"] = ellipse_data[3]
                    return result
            except Exception:
                pass

            # Attempt BSplineCurve info
            try:
                bspline_info = geom.GetBSplineInfo()
                if isinstance(bspline_info, tuple) and len(bspline_info) >= 4:
                    result["geometry_type"] = "BSplineCurve"
                    result["order"] = bspline_info[0]
                    result["num_poles"] = bspline_info[1]
                    result["num_knots"] = bspline_info[2]
                    result["rational"] = bool(bspline_info[3])
                    if len(bspline_info) > 4:
                        result["closed"] = bool(bspline_info[4])
                    if len(bspline_info) > 5:
                        result["periodic"] = bool(bspline_info[5])
                    if len(bspline_info) > 6:
                        result["planar"] = bool(bspline_info[6])
                    return result
            except Exception:
                pass

            # Default: Line (no extra data needed beyond endpoints)
            result["geometry_type"] = "Line"
            result["raw_type"] = geom_type
            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_face_normal(self, face_index: int, u: float = 0.5, v: float = 0.5) -> dict[str, Any]:
        """
        Get the normal vector at a parametric point on a face.

        Uses face.GetNormal(numPoints, paramsArray, normalsArray).

        Args:
            face_index: 0-based face index
            u: U parameter (0.0 to 1.0)
            v: V parameter (0.0 to 1.0)

        Returns:
            Dict with normal vector [nx, ny, nz]
        """
        try:
            _doc, _model, _body, face = self._get_face(face_index)

            params_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [u, v])
            normals_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])

            result = face.GetNormal(1, params_arr, normals_arr)

            if isinstance(result, tuple) and len(result) >= 2:
                normal = list(result[-1])
            else:
                normal = list(normals_arr)

            return {
                "normal": normal[:3],
                "params": [u, v],
                "face_index": face_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_face_geometry(self, face_index: int) -> dict[str, Any]:
        """
        Get the underlying geometry type and data of a face.

        Inspects face.Geometry to determine if the face is a Plane,
        Cylinder, Cone, Sphere, Torus, or BSplineSurface.

        Args:
            face_index: 0-based face index

        Returns:
            Dict with geometry_type and associated geometric data
        """
        try:
            _doc, _model, _body, face = self._get_face(face_index)

            geom = face.Geometry
            result: dict[str, Any] = {"face_index": face_index}

            # Try Plane
            try:
                plane_data = geom.GetPlaneData()
                if isinstance(plane_data, tuple) and len(plane_data) >= 2:
                    result["geometry_type"] = "Plane"
                    result["root_point"] = self._to_list(plane_data[0])
                    result["normal"] = self._to_list(plane_data[1])
                    return result
            except Exception:
                pass

            # Try Cylinder
            try:
                cyl_data = geom.GetCylinderData()
                if isinstance(cyl_data, tuple) and len(cyl_data) >= 3:
                    result["geometry_type"] = "Cylinder"
                    result["base_point"] = self._to_list(cyl_data[0])
                    result["axis"] = self._to_list(cyl_data[1])
                    result["radius"] = cyl_data[2]
                    return result
            except Exception:
                pass

            # Try Cone
            try:
                cone_data = geom.GetConeData()
                if isinstance(cone_data, tuple) and len(cone_data) >= 4:
                    result["geometry_type"] = "Cone"
                    result["base_point"] = self._to_list(cone_data[0])
                    result["axis"] = self._to_list(cone_data[1])
                    result["radius"] = cone_data[2]
                    result["half_angle"] = cone_data[3]
                    if len(cone_data) > 4:
                        result["expanding"] = bool(cone_data[4])
                    return result
            except Exception:
                pass

            # Try Sphere
            try:
                sphere_data = geom.GetSphereData()
                if isinstance(sphere_data, tuple) and len(sphere_data) >= 2:
                    result["geometry_type"] = "Sphere"
                    result["center"] = self._to_list(sphere_data[0])
                    result["radius"] = sphere_data[1]
                    return result
            except Exception:
                pass

            # Try Torus
            try:
                torus_data = geom.GetTorusData()
                if isinstance(torus_data, tuple) and len(torus_data) >= 4:
                    result["geometry_type"] = "Torus"
                    result["center"] = self._to_list(torus_data[0])
                    result["axis"] = self._to_list(torus_data[1])
                    result["major_radius"] = torus_data[2]
                    result["minor_radius"] = torus_data[3]
                    return result
            except Exception:
                pass

            # Try BSplineSurface
            try:
                bspline_info = geom.GetBSplineInfo()
                if isinstance(bspline_info, tuple) and len(bspline_info) >= 2:
                    result["geometry_type"] = "BSplineSurface"
                    result["raw_info"] = list(bspline_info)
                    return result
            except Exception:
                pass

            result["geometry_type"] = "Unknown"
            with contextlib.suppress(Exception):
                result["raw_type"] = geom.Type
            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_face_loops(self, face_index: int) -> dict[str, Any]:
        """
        Get loop info for a face (outer boundary vs holes).

        Iterates face.Loops to determine which loop is the outer boundary
        and how many edges each loop contains.

        Args:
            face_index: 0-based face index

        Returns:
            Dict with loop count and details per loop
        """
        try:
            _doc, _model, _body, face = self._get_face(face_index)

            loops_col = face.Loops
            loop_list = []

            for i in range(1, loops_col.Count + 1):
                loop = loops_col.Item(i)
                loop_info: dict[str, Any] = {"index": i - 1}

                with contextlib.suppress(Exception):
                    loop_info["is_outer"] = bool(loop.IsOuterLoop)

                try:
                    loop_info["edge_count"] = loop.Edges.Count
                except Exception:
                    loop_info["edge_count"] = 0

                loop_list.append(loop_info)

            return {
                "face_index": face_index,
                "loop_count": len(loop_list),
                "loops": loop_list,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_face_curvature(self, face_index: int, u: float = 0.5, v: float = 0.5) -> dict[str, Any]:
        """
        Get principal curvatures at a parametric point on a face.

        Uses face.GetCurvatures(numPoints, params, maxTangents,
        maxCurvatures, minCurvatures).

        Args:
            face_index: 0-based face index
            u: U parameter (0.0 to 1.0)
            v: V parameter (0.0 to 1.0)

        Returns:
            Dict with max/min curvature values and max tangent direction
        """
        try:
            _doc, _model, _body, face = self._get_face(face_index)

            params_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [u, v])
            max_tangents_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])
            max_curvatures_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0])
            min_curvatures_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0])

            result = face.GetCurvatures(
                1, params_arr, max_tangents_arr, max_curvatures_arr, min_curvatures_arr
            )

            if isinstance(result, tuple) and len(result) >= 3:
                max_tangent = list(result[0])[:3]
                r1 = result[1]
                max_curvature = r1[0] if hasattr(r1, "__iter__") else float(r1)
                r2 = result[2]
                min_curvature = r2[0] if hasattr(r2, "__iter__") else float(r2)
            else:
                max_tangent = list(max_tangents_arr)[:3]
                max_curvature = list(max_curvatures_arr)[0]
                min_curvature = list(min_curvatures_arr)[0]

            return {
                "max_curvature": max_curvature,
                "min_curvature": min_curvature,
                "max_tangent": max_tangent,
                "params": [u, v],
                "face_index": face_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_vertex_point(
        self, face_index: int, edge_index: int, which: str = "start"
    ) -> dict[str, Any]:
        """
        Get XYZ coordinates of a vertex on an edge.

        Accesses edge.StartVertex or edge.EndVertex and calls
        vertex.GetPointData() to retrieve the 3D coordinates.

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index
            which: 'start' or 'end' to select which vertex

        Returns:
            Dict with vertex point [x, y, z]
        """
        try:
            if which not in ("start", "end"):
                return {"error": f"Invalid vertex selector: '{which}'. Use 'start' or 'end'."}

            _doc, _model, _body, _face, edge = self._get_face_edge(face_index, edge_index)

            vertex = edge.StartVertex if which == "start" else edge.EndVertex

            point_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])
            result = vertex.GetPointData(point_arr)

            point = self._to_list(result[0]) if isinstance(result, tuple) else list(point_arr)

            vertex_id = -1
            with contextlib.suppress(Exception):
                vertex_id = vertex.ID

            return {
                "point": point[:3],
                "vertex_id": vertex_id,
                "which": which,
                "face_index": face_index,
                "edge_index": edge_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_body_extreme_point(
        self, direction_x: float, direction_y: float, direction_z: float
    ) -> dict[str, Any]:
        """
        Get the extreme point of the body in a given direction.

        Uses body.GetExtremePoint(dx, dy, dz, ex, ey, ez) where the
        last three parameters are out-params for the extreme point coords.

        Args:
            direction_x: X component of direction vector
            direction_y: Y component of direction vector
            direction_z: Z component of direction vector

        Returns:
            Dict with extreme point [x, y, z] and direction
        """
        try:
            _doc, _model, body = self._get_body()

            result = body.GetExtremePoint(
                direction_x,
                direction_y,
                direction_z,
                0.0,
                0.0,
                0.0,
            )

            if isinstance(result, tuple) and len(result) >= 3:
                extreme = [result[0], result[1], result[2]]
            else:
                extreme = [0.0, 0.0, 0.0]

            return {
                "extreme_point": extreme,
                "direction": [direction_x, direction_y, direction_z],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_faces_by_ray(
        self,
        origin_x: float,
        origin_y: float,
        origin_z: float,
        direction_x: float,
        direction_y: float,
        direction_z: float,
    ) -> dict[str, Any]:
        """
        Ray-cast query to find faces hit by a ray.

        Uses body.FacesByRay(ox, oy, oz, dx, dy, dz) which returns
        a Faces collection of intersected faces.

        Args:
            origin_x, origin_y, origin_z: Ray origin point
            direction_x, direction_y, direction_z: Ray direction vector

        Returns:
            Dict with hit face count and face info
        """
        try:
            _doc, _model, body = self._get_body()

            faces = body.FacesByRay(
                origin_x,
                origin_y,
                origin_z,
                direction_x,
                direction_y,
                direction_z,
            )

            face_list = []
            for i in range(1, faces.Count + 1):
                face = faces.Item(i)
                face_info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    face_info["area"] = face.Area
                with contextlib.suppress(Exception):
                    face_info["id"] = face.ID
                face_list.append(face_info)

            return {
                "face_count": len(face_list),
                "faces": face_list,
                "ray_origin": [origin_x, origin_y, origin_z],
                "ray_direction": [direction_x, direction_y, direction_z],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_shell_info(self, shell_index: int = 0) -> dict[str, Any]:
        """
        Get topology information about a shell.

        Accesses body.Shells collection and retrieves properties like
        IsClosed, Volume, IsVoid, face count, and edge count.

        Args:
            shell_index: 0-based shell index (default 0 = first shell)

        Returns:
            Dict with shell topology info
        """
        try:
            _doc, _model, body = self._get_body()

            shells = body.Shells
            if shell_index < 0 or shell_index >= shells.Count:
                return {
                    "error": f"Invalid shell index: {shell_index}. Body has {shells.Count} shells."
                }

            shell = shells.Item(shell_index + 1)

            info: dict[str, Any] = {"shell_index": shell_index}

            with contextlib.suppress(Exception):
                info["is_closed"] = bool(shell.IsClosed)
            with contextlib.suppress(Exception):
                info["volume"] = shell.Volume
            with contextlib.suppress(Exception):
                info["is_void"] = bool(shell.IsVoid)
            try:
                info["face_count"] = shell.Faces.Count
            except Exception:
                info["face_count"] = 0
            try:
                info["edge_count"] = shell.Edges.Count
            except Exception:
                info["edge_count"] = 0

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def is_point_inside_body(self, x: float, y: float, z: float) -> dict[str, Any]:
        """
        Test if a 3D point is inside the solid body.

        Uses the first shell's IsPointInside method with a SAFEARRAY param.

        Args:
            x, y, z: 3D point coordinates (meters)

        Returns:
            Dict with is_inside boolean
        """
        try:
            _doc, _model, body = self._get_body()

            shells = body.Shells
            if shells.Count == 0:
                return {"error": "Body has no shells"}

            shell = shells.Item(1)

            point_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [x, y, z])
            is_inside = shell.IsPointInside(point_arr)

            return {
                "is_inside": bool(is_inside),
                "point": [x, y, z],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_body_shells(self) -> dict[str, Any]:
        """
        List all shells in the body with basic properties.

        Iterates body.Shells to get IsClosed and Volume for each shell.

        Returns:
            Dict with shell count and list of shell info
        """
        try:
            _doc, _model, body = self._get_body()

            shells = body.Shells
            shell_list = []

            for i in range(1, shells.Count + 1):
                shell = shells.Item(i)
                shell_info: dict[str, Any] = {"index": i - 1}

                with contextlib.suppress(Exception):
                    shell_info["is_closed"] = bool(shell.IsClosed)
                with contextlib.suppress(Exception):
                    shell_info["volume"] = shell.Volume

                shell_list.append(shell_info)

            return {
                "shell_count": len(shell_list),
                "shells": shell_list,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_body_vertices(self) -> dict[str, Any]:
        """
        Get all vertices of the body with their 3D coordinates.

        Iterates body.Vertices and calls GetPointData() on each to
        extract XYZ coordinates.

        Returns:
            Dict with vertex count and coordinate list
        """
        try:
            _doc, _model, body = self._get_body()

            vertices = body.Vertices
            vertex_list = []

            for i in range(1, vertices.Count + 1):
                try:
                    vertex = vertices.Item(i)
                    point_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])
                    result = vertex.GetPointData(point_arr)

                    point = (
                        self._to_list(result[0]) if isinstance(result, tuple) else list(point_arr)
                    )

                    vertex_list.append({"index": i - 1, "point": point[:3]})
                except Exception:
                    vertex_list.append({"index": i - 1, "point": None})

            return {
                "vertex_count": len(vertex_list),
                "vertices": vertex_list,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_bspline_curve_info(self, face_index: int, edge_index: int) -> dict[str, Any]:
        """
        Get NURBS curve metadata from an edge's underlying geometry.

        Accesses edge.Geometry and calls GetBSplineInfo() to retrieve
        order, pole count, knot count, and flags.

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index

        Returns:
            Dict with BSpline curve properties, or error if not a BSpline
        """
        try:
            _doc, _model, _body, _face, edge = self._get_face_edge(face_index, edge_index)

            geom = edge.Geometry
            bspline_info = geom.GetBSplineInfo()

            if not isinstance(bspline_info, tuple) or len(bspline_info) < 4:
                return {
                    "error": "Edge geometry is not a BSpline curve "
                    "or GetBSplineInfo returned unexpected data",
                    "face_index": face_index,
                    "edge_index": edge_index,
                }

            result: dict[str, Any] = {
                "face_index": face_index,
                "edge_index": edge_index,
                "order": bspline_info[0],
                "num_poles": bspline_info[1],
                "num_knots": bspline_info[2],
                "rational": bool(bspline_info[3]),
            }
            if len(bspline_info) > 4:
                result["closed"] = bool(bspline_info[4])
            if len(bspline_info) > 5:
                result["periodic"] = bool(bspline_info[5])
            if len(bspline_info) > 6:
                result["planar"] = bool(bspline_info[6])

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_bspline_surface_info(self, face_index: int) -> dict[str, Any]:
        """
        Get NURBS surface metadata from a face's underlying geometry.

        Accesses face.Geometry and calls GetBSplineInfo() to retrieve
        order, pole count, knot count, and flags for U and V directions.

        Args:
            face_index: 0-based face index

        Returns:
            Dict with BSpline surface properties, or error if not BSpline
        """
        try:
            _doc, _model, _body, face = self._get_face(face_index)

            geom = face.Geometry
            bspline_info = geom.GetBSplineInfo()

            if not isinstance(bspline_info, tuple) or len(bspline_info) < 2:
                return {
                    "error": "Face geometry is not a BSpline surface "
                    "or GetBSplineInfo returned unexpected data",
                    "face_index": face_index,
                }

            result: dict[str, Any] = {"face_index": face_index}

            # BSplineSurface.GetBSplineInfo returns data for both U and V
            # The exact tuple structure depends on the SE API version.
            # Common pattern: (orderU, orderV, numPolesU, numPolesV,
            #                   numKnotsU, numKnotsV, rational, ...)
            if len(bspline_info) >= 6:
                result["order"] = [bspline_info[0], bspline_info[1]]
                result["num_poles"] = [bspline_info[2], bspline_info[3]]
                result["num_knots"] = [bspline_info[4], bspline_info[5]]
                if len(bspline_info) > 6:
                    result["rational"] = bool(bspline_info[6])
                if len(bspline_info) > 7:
                    result["planar"] = bool(bspline_info[7])
            else:
                # Simpler format - store raw info
                result["raw_info"] = list(bspline_info)

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_edge_curvature(
        self, face_index: int, edge_index: int, param: float = 0.5
    ) -> dict[str, Any]:
        """
        Get curvature at a parameter on an edge.

        Uses edge.GetCurvature(numParams, paramsArray, directionsArray,
        curvaturesArray).

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index
            param: Parameter value (0.0 to 1.0, default 0.5)

        Returns:
            Dict with curvature value, direction vector, and parameter
        """
        try:
            _doc, _model, _body, _face, edge = self._get_face_edge(face_index, edge_index)

            params_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [param])
            directions_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])
            curvatures_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0])

            result = edge.GetCurvature(1, params_arr, directions_arr, curvatures_arr)

            if isinstance(result, tuple) and len(result) >= 2:
                direction = list(result[0])[:3]
                r1 = result[1]
                curvature = r1[0] if hasattr(r1, "__iter__") else float(r1)
            else:
                direction = list(directions_arr)[:3]
                curvature = list(curvatures_arr)[0]

            return {
                "curvature": curvature,
                "direction": direction,
                "param": param,
                "face_index": face_index,
                "edge_index": edge_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
