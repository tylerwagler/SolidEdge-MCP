"""
Solid Edge Query and Inspection Operations

Handles querying model data, measurements, and properties.
"""

from typing import Dict, Any
import math
import traceback

from .constants import FaceQueryConstants, ModelingModeConstants


class QueryManager:
    """Manages query and inspection operations"""

    def __init__(self, document_manager):
        self.doc_manager = document_manager

    def _get_first_model(self):
        """Get the first model from the active document."""
        doc = self.doc_manager.get_active_document()
        if not hasattr(doc, 'Models'):
            raise Exception("Document does not have a Models collection")
        models = doc.Models
        if models.Count == 0:
            raise Exception("No features in document")
        return doc, models.Item(1)

    def get_mass_properties(self, density: float = 7850) -> Dict[str, Any]:
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
            result = model.ComputePhysicalPropertiesWithSpecifiedDensity(
                density, 0.99
            )

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
                    "coordinates": "meters"
                }
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_bounding_box(self) -> Dict[str, Any]:
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
                    "z": max_pt[2] - min_pt[2]
                },
                "units": "meters"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def list_features(self) -> Dict[str, Any]:
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
            if hasattr(doc, 'DesignEdgebarFeatures'):
                debf = doc.DesignEdgebarFeatures
                for i in range(1, debf.Count + 1):
                    try:
                        feat = debf.Item(i)
                        feat_info = {
                            "index": i - 1,
                            "name": feat.Name if hasattr(feat, 'Name') else f"Feature_{i}",
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
                            "name": feat.Name if hasattr(feat, 'Name') else f"Feature_{i}",
                        }
                        features.append(feat_info)
                    except Exception:
                        features.append({"index": i - 1, "name": f"Feature_{i}"})

            return {
                "features": features,
                "count": len(features)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def measure_distance(self, x1: float, y1: float, z1: float,
                        x2: float, y2: float, z2: float) -> Dict[str, Any]:
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
                "delta": {
                    "x": dx,
                    "y": dy,
                    "z": dz
                },
                "point1": [x1, y1, z1],
                "point2": [x2, y2, z2],
                "units": "meters"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_document_properties(self) -> Dict[str, Any]:
        """Get document properties and metadata"""
        try:
            doc = self.doc_manager.get_active_document()

            properties = {
                "name": doc.Name if hasattr(doc, 'Name') else "Unknown",
                "path": doc.FullName if hasattr(doc, 'FullName') else "Unsaved",
                "modified": not doc.Saved if hasattr(doc, 'Saved') else False,
                "read_only": doc.ReadOnly if hasattr(doc, 'ReadOnly') else False
            }

            # Try to get summary info
            try:
                if hasattr(doc, 'SummaryInfo'):
                    summary = doc.SummaryInfo
                    if hasattr(summary, 'Title'):
                        properties["title"] = summary.Title
                    if hasattr(summary, 'Author'):
                        properties["author"] = summary.Author
                    if hasattr(summary, 'Subject'):
                        properties["subject"] = summary.Subject
                    if hasattr(summary, 'Comments'):
                        properties["comments"] = summary.Comments
            except:
                pass

            # Add body topology info
            try:
                models = doc.Models
                if models.Count > 0:
                    body = models.Item(1).Body
                    properties["volume_m3"] = body.Volume
            except:
                pass

            return properties
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_feature_count(self) -> Dict[str, Any]:
        """Get count of features in the document"""
        try:
            doc = self.doc_manager.get_active_document()

            counts = {}

            if hasattr(doc, 'DesignEdgebarFeatures'):
                counts["features"] = doc.DesignEdgebarFeatures.Count

            if hasattr(doc, 'Models'):
                counts["models"] = doc.Models.Count

            if hasattr(doc, 'ProfileSets'):
                counts["sketches"] = doc.ProfileSets.Count

            if hasattr(doc, 'RefPlanes'):
                counts["ref_planes"] = doc.RefPlanes.Count

            if hasattr(doc, 'Variables'):
                counts["variables"] = doc.Variables.Count

            return counts
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # REFERENCE PLANES
    # =================================================================

    def get_ref_planes(self) -> Dict[str, Any]:
        """
        List all reference planes in the active document.

        Default planes: 1=Top/XZ, 2=Front/XY, 3=Right/YZ.
        Additional planes are created with create_ref_plane_by_offset.

        Returns:
            Dict with list of reference planes and their indices
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'RefPlanes'):
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

                    try:
                        plane_info["visible"] = plane.Visible
                    except Exception:
                        pass

                    planes.append(plane_info)
                except Exception:
                    planes.append({"index": i, "name": default_names.get(i, f"RefPlane_{i}")})

            return {
                "planes": planes,
                "count": len(planes),
                "note": "Use plane index (1-based) in create_sketch_on_plane or create_ref_plane_by_offset"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # VARIABLES
    # =================================================================

    def get_variables(self) -> Dict[str, Any]:
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
                        "name": var.DisplayName if hasattr(var, 'DisplayName') else f"Var_{i}",
                    }
                    try:
                        var_info["value"] = var.Value
                    except Exception:
                        pass
                    try:
                        var_info["formula"] = var.Formula
                    except Exception:
                        pass
                    try:
                        var_info["units"] = var.Units
                    except Exception:
                        pass
                    var_list.append(var_info)
                except Exception:
                    var_list.append({"index": i - 1, "name": f"Var_{i}"})

            return {
                "variables": var_list,
                "count": len(var_list)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_variable(self, name: str) -> Dict[str, Any]:
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
                    display_name = var.DisplayName if hasattr(var, 'DisplayName') else ""
                    if display_name == name:
                        result = {"name": name, "index": i - 1}
                        try:
                            result["value"] = var.Value
                        except Exception:
                            pass
                        try:
                            result["formula"] = var.Formula
                        except Exception:
                            pass
                        try:
                            result["units"] = var.Units
                        except Exception:
                            pass
                        return result
                except Exception:
                    continue

            return {"error": f"Variable '{name}' not found"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def set_variable(self, name: str, value: float) -> Dict[str, Any]:
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
                    display_name = var.DisplayName if hasattr(var, 'DisplayName') else ""
                    if display_name == name:
                        old_value = var.Value
                        var.Value = value
                        return {
                            "status": "updated",
                            "name": name,
                            "old_value": old_value,
                            "new_value": value
                        }
                except Exception:
                    continue

            return {"error": f"Variable '{name}' not found"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def add_variable(self, name: str, formula: str, units_type: str = None) -> Dict[str, Any]:
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

            result = {
                "status": "created",
                "name": name,
                "formula": formula
            }
            try:
                result["value"] = var.Value
            except Exception:
                pass
            try:
                result["display_name"] = var.DisplayName
            except Exception:
                pass

            return result
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # CUSTOM PROPERTIES
    # =================================================================

    def get_custom_properties(self) -> Dict[str, Any]:
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
                    ps_name = ps.Name if hasattr(ps, 'Name') else f"Set_{ps_idx}"

                    props = {}
                    for p_idx in range(1, ps.Count + 1):
                        try:
                            prop = ps.Item(p_idx)
                            prop_name = prop.Name if hasattr(prop, 'Name') else f"Prop_{p_idx}"
                            try:
                                props[prop_name] = prop.Value
                            except Exception:
                                props[prop_name] = None
                        except Exception:
                            continue

                    properties[ps_name] = props
                except Exception:
                    continue

            return {
                "property_sets": properties,
                "count": len(properties)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def set_custom_property(self, name: str, value: str) -> Dict[str, Any]:
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
                    if hasattr(ps, 'Name') and ps.Name == "Custom":
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
                    if hasattr(prop, 'Name') and prop.Name == name:
                        old_value = prop.Value
                        prop.Value = value
                        return {
                            "status": "updated",
                            "name": name,
                            "old_value": old_value,
                            "new_value": value
                        }
                except Exception:
                    continue

            # Property doesn't exist, add it
            custom_ps.Add(name, value)
            return {
                "status": "created",
                "name": name,
                "value": value
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def delete_custom_property(self, name: str) -> Dict[str, Any]:
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
                    if hasattr(ps, 'Name') and ps.Name == "Custom":
                        for p_idx in range(1, ps.Count + 1):
                            try:
                                prop = ps.Item(p_idx)
                                if hasattr(prop, 'Name') and prop.Name == name:
                                    prop.Delete()
                                    return {
                                        "status": "deleted",
                                        "name": name
                                    }
                            except Exception:
                                continue
                        return {"error": f"Property '{name}' not found in Custom set"}
                except Exception:
                    continue

            return {"error": "Custom property set not found"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # BODY TOPOLOGY QUERIES
    # =================================================================

    def get_body_faces(self) -> Dict[str, Any]:
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
                6: "plane",      # igQueryPlane
                10: "cylinder",  # igQueryCylinder
                7: "cone",       # igQueryCone
                9: "sphere",     # igQuerySphere
                8: "torus",      # igQueryTorus
                5: "spline",     # igQuerySpline
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
                            ec = tf.Edges.Count if hasattr(tf.Edges, 'Count') else 0
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
                    try:
                        face_info["area"] = face.Area
                    except Exception:
                        pass
                    try:
                        edge_count = face.Edges.Count if hasattr(face.Edges, 'Count') else 0
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

            return {
                "faces": face_list,
                "count": len(face_list)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_body_edges(self) -> Dict[str, Any]:
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
                    edge_count = edges.Count if hasattr(edges, 'Count') else 0
                    total_edges += edge_count
                    face_edges.append({
                        "face_index": fi - 1,
                        "edge_count": edge_count
                    })
                except Exception:
                    face_edges.append({"face_index": fi - 1, "edge_count": 0})

            return {
                "face_edges": face_edges,
                "total_face_count": faces.Count,
                "total_edge_references": total_edges,
                "note": "Edge count includes shared edges (counted once per face)"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_face_info(self, face_index: int) -> Dict[str, Any]:
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

            try:
                info["type"] = face.Type
            except Exception:
                pass
            try:
                info["area"] = face.Area
            except Exception:
                pass
            try:
                edges = face.Edges
                info["edge_count"] = edges.Count if hasattr(edges, 'Count') else 0
            except Exception:
                pass
            try:
                vertices = face.Vertices
                info["vertex_count"] = vertices.Count if hasattr(vertices, 'Count') else 0
            except Exception:
                pass

            return info
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # PERFORMANCE & RECOMPUTE
    # =================================================================

    def get_body_facet_data(self, tolerance: float = 0.0) -> Dict[str, Any]:
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
            # GetFacetData(Tolerance, FacetCount, Points, Normals, TextureCoords, StyleIDs, FaceIDs, bHonourPrefs)
            points = arr_mod.array('d', [])
            normals = arr_mod.array('d', [])
            texture_coords = arr_mod.array('d', [])
            style_ids = arr_mod.array('i', [])
            face_ids = arr_mod.array('i', [])

            try:
                result_data = body.GetFacetData(
                    tolerance,     # Tolerance
                )

                # GetFacetData returns a tuple of (facetCount, points, normals, textureCoords, styleIds, faceIds)
                if isinstance(result_data, tuple) and len(result_data) >= 2:
                    facet_count = result_data[0] if isinstance(result_data[0], int) else 0
                    pts = result_data[1] if len(result_data) > 1 else []

                    return {
                        "facet_count": facet_count,
                        "point_count": len(pts) // 3 if pts else 0,
                        "tolerance": tolerance,
                        "has_data": facet_count > 0
                    }
            except Exception:
                pass

            # Alternative: try with explicit out params
            try:
                facet_count = 0
                body.GetFacetData(tolerance, facet_count, points, normals, texture_coords, style_ids, face_ids, False)

                return {
                    "facet_count": facet_count,
                    "point_count": len(points) // 3 if points else 0,
                    "tolerance": tolerance,
                    "has_data": len(points) > 0
                }
            except Exception as e2:
                return {
                    "error": f"GetFacetData failed: {e2}",
                    "note": "Body facet data may require specific COM marshaling. Try export_stl() instead.",
                    "traceback": traceback.format_exc()
                }

        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_solid_bodies(self) -> Dict[str, Any]:
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
                        "name": model.Name if hasattr(model, 'Name') else f"Model_{i}",
                    }

                    try:
                        body_info["is_solid"] = body.IsSolid
                    except Exception:
                        body_info["is_solid"] = True  # Default assumption

                    try:
                        body_info["volume"] = body.Volume
                    except Exception:
                        pass

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
                            "name": cm.Name if hasattr(cm, 'Name') else f"Construction_{i}",
                        }
                        try:
                            body_info["is_solid"] = body.IsSolid
                        except Exception:
                            pass
                        bodies.append(body_info)
                    except Exception:
                        pass
            except Exception:
                pass  # No Constructions collection

            return {
                "total_bodies": len(bodies),
                "bodies": bodies
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_modeling_mode(self) -> Dict[str, Any]:
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
                return {
                    "mode": mode_name,
                    "mode_value": mode
                }
            except Exception:
                return {"error": "ModelingMode not available on this document type"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def set_modeling_mode(self, mode: str) -> Dict[str, Any]:
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
                    "old_mode": "ordered" if old_mode == ModelingModeConstants.seModelingModeOrdered else "synchronous",
                    "new_mode": "ordered" if new_mode == ModelingModeConstants.seModelingModeOrdered else "synchronous"
                }
            except Exception as e:
                return {"error": f"Cannot change modeling mode: {e}"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def suppress_feature(self, feature_name: str) -> Dict[str, Any]:
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
                        return {
                            "status": "suppressed",
                            "feature": feature_name
                        }
                except Exception:
                    continue

            return {"error": f"Feature '{feature_name}' not found"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def unsuppress_feature(self, feature_name: str) -> Dict[str, Any]:
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
                        return {
                            "status": "unsuppressed",
                            "feature": feature_name
                        }
                except Exception:
                    continue

            return {"error": f"Feature '{feature_name}' not found"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_select_set(self) -> Dict[str, Any]:
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

                    try:
                        item_info["type"] = str(type(item).__name__)
                    except Exception:
                        pass

                    try:
                        if hasattr(item, 'Name'):
                            item_info["name"] = item.Name
                    except Exception:
                        pass

                    items.append(item_info)
                except Exception:
                    items.append({"index": i - 1, "error": "could not read"})

            return {
                "count": count,
                "items": items
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def clear_select_set(self) -> Dict[str, Any]:
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

            return {
                "status": "cleared",
                "items_removed": old_count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def select_add(self, object_type: str, index: int) -> Dict[str, Any]:
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
                return {"error": f"Unsupported object type: {object_type}. Use 'feature', 'face', or 'plane'."}

            if obj is None:
                return {"error": "Could not resolve object to select"}

            select_set.Add(obj)

            return {
                "status": "added",
                "object_type": object_type,
                "index": index,
                "selection_count": select_set.Count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def select_remove(self, index: int) -> Dict[str, Any]:
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

            return {
                "status": "removed",
                "index": index,
                "selection_count": select_set.Count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def select_all(self) -> Dict[str, Any]:
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

            return {
                "status": "selected_all",
                "selection_count": select_set.Count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def select_copy(self) -> Dict[str, Any]:
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

            return {
                "status": "copied",
                "items_copied": select_set.Count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def select_cut(self) -> Dict[str, Any]:
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

            return {
                "status": "cut",
                "items_cut": count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def select_delete(self) -> Dict[str, Any]:
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

            return {
                "status": "deleted",
                "items_deleted": count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def select_suspend_display(self) -> Dict[str, Any]:
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
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def select_resume_display(self) -> Dict[str, Any]:
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
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def select_refresh_display(self) -> Dict[str, Any]:
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
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # BODY APPEARANCE & MATERIAL
    # =================================================================

    def set_body_color(self, red: int, green: int, blue: int) -> Dict[str, Any]:
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
                "hex": f"#{red:02x}{green:02x}{blue:02x}"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def set_material_density(self, density: float) -> Dict[str, Any]:
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
            result = model.ComputePhysicalPropertiesWithSpecifiedDensity(
                density, 0.99
            )

            mass = result[2] if len(result) > 2 else 0
            volume = result[0] if len(result) > 0 else 0

            return {
                "status": "computed",
                "density": density,
                "mass": mass,
                "volume": volume,
                "units": {"density": "kg/m³", "mass": "kg", "volume": "m³"}
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_edge_count(self) -> Dict[str, Any]:
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
                    if hasattr(edges, 'Count'):
                        total_edges += edges.Count
                except Exception:
                    pass

            return {
                "total_edge_references": total_edges,
                "face_count": faces.Count,
                "note": "Shared edges are counted once per face"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def recompute(self) -> Dict[str, Any]:
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
            try:
                doc.Recompute()
            except Exception:
                pass

            return {"status": "recomputed"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_design_edgebar_features(self) -> Dict[str, Any]:
        """
        Get the full feature tree from DesignEdgebarFeatures.

        Unlike list_features() which only shows Models, this shows the
        complete design tree including sketches, reference planes, etc.

        Returns:
            Dict with list of all feature tree entries
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'DesignEdgebarFeatures'):
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
                    try:
                        entry["type"] = feat.Type
                    except Exception:
                        pass
                    try:
                        entry["suppressed"] = feat.IsSuppressed
                    except Exception:
                        pass
                    feature_list.append(entry)
                except Exception:
                    feature_list.append({"index": i - 1, "name": f"Feature_{i}"})

            return {
                "features": feature_list,
                "count": len(feature_list)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def rename_feature(self, old_name: str, new_name: str) -> Dict[str, Any]:
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

            if not hasattr(doc, 'DesignEdgebarFeatures'):
                return {"error": "DesignEdgebarFeatures not available"}

            features = doc.DesignEdgebarFeatures

            for i in range(1, features.Count + 1):
                try:
                    feat = features.Item(i)
                    if hasattr(feat, 'Name') and feat.Name == old_name:
                        feat.Name = new_name
                        return {
                            "status": "renamed",
                            "old_name": old_name,
                            "new_name": new_name
                        }
                except Exception:
                    continue

            return {"error": f"Feature '{old_name}' not found"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def set_document_property(self, name: str, value: str) -> Dict[str, Any]:
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
            if not hasattr(doc, 'SummaryInfo'):
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
                return {
                    "error": f"Unknown property: {name}. "
                             f"Valid: {', '.join(prop_map.keys())}"
                }

            setattr(summary, prop_attr, value)

            return {
                "status": "set",
                "property": name,
                "value": value
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_face_area(self, face_index: int) -> Dict[str, Any]:
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
                return {
                    "error": f"Invalid face index: {face_index}. "
                             f"Body has {faces.Count} faces."
                }

            face = faces.Item(face_index + 1)

            area = face.Area

            return {
                "face_index": face_index,
                "area": area,
                "area_mm2": area * 1e6  # Convert m² to mm²
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_surface_area(self) -> Dict[str, Any]:
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
                return {
                    "surface_area": area,
                    "surface_area_mm2": area * 1e6
                }
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
                "method": "sum_of_faces"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_volume(self) -> Dict[str, Any]:
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
                "volume_cm3": volume * 1e6   # m³ to cm³
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_face_count(self) -> Dict[str, Any]:
        """
        Get the total number of faces on the body.

        Returns:
            Dict with face count
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)
            return {
                "face_count": faces.Count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_edge_info(self, face_index: int, edge_index: int) -> Dict[str, Any]:
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

            try:
                info["type"] = edge.Type
            except Exception:
                pass

            try:
                start = edge.StartVertex
                end = edge.EndVertex
                info["start_vertex"] = [start.X, start.Y, start.Z]
                info["end_vertex"] = [end.X, end.Y, end.Z]
            except Exception:
                pass

            return info
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def set_face_color(self, face_index: int, red: int, green: int, blue: int) -> Dict[str, Any]:
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

            return {
                "status": "updated",
                "face_index": face_index,
                "color": [red, green, blue]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_center_of_gravity(self) -> Dict[str, Any]:
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
                        "center_of_gravity_mm": [
                            cog_x * 1000,
                            cog_y * 1000,
                            cog_z * 1000
                        ]
                    }
            except Exception:
                pass

            # Fallback: compute physical properties
            doc, model = self._get_first_model()
            result = model.ComputePhysicalPropertiesWithSpecifiedDensity(7850.0, 0.001)
            # result[3] is the center of gravity tuple
            cog = result[3]
            return {
                "center_of_gravity": list(cog),
                "center_of_gravity_mm": [c * 1000 for c in cog]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_moments_of_inertia(self) -> Dict[str, Any]:
        """
        Get the moments of inertia of the part.

        Returns:
            Dict with moments of inertia values
        """
        try:
            doc, model = self._get_first_model()
            result = model.ComputePhysicalPropertiesWithSpecifiedDensity(7850.0, 0.001)
            # result: (volume, area, mass, cog, cov, moi, principal_moi, principal_axes, radii_of_gyration, ?, ?)
            moi = result[5]
            principal_moi = result[6]

            return {
                "moments_of_inertia": list(moi) if hasattr(moi, '__iter__') else moi,
                "principal_moments": list(principal_moi) if hasattr(principal_moi, '__iter__') else principal_moi,
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def delete_feature(self, feature_name: str) -> Dict[str, Any]:
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

            if not hasattr(doc, 'DesignEdgebarFeatures'):
                return {"error": "Document does not support feature deletion"}

            debf = doc.DesignEdgebarFeatures
            for i in range(1, debf.Count + 1):
                try:
                    feat = debf.Item(i)
                    if hasattr(feat, 'Name') and feat.Name == feature_name:
                        feat.Delete()
                        return {
                            "status": "deleted",
                            "feature_name": feature_name
                        }
                except Exception:
                    continue

            return {"error": f"Feature '{feature_name}' not found"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_body_color(self) -> Dict[str, Any]:
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
                return {
                    "red": red,
                    "green": green,
                    "blue": blue,
                    "ole_color": color
                }
            except Exception:
                # Try alternative
                try:
                    r, g, b = body.GetColor()
                    return {"red": r, "green": g, "blue": b}
                except Exception:
                    return {"error": "Could not determine body color"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def measure_angle(self, x1: float, y1: float, z1: float,
                      x2: float, y2: float, z2: float,
                      x3: float, y3: float, z3: float) -> Dict[str, Any]:
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
            dot = v1[0]*v2[0] + v1[1]*v2[1] + v1[2]*v2[2]

            # Magnitudes
            mag1 = math.sqrt(v1[0]**2 + v1[1]**2 + v1[2]**2)
            mag2 = math.sqrt(v2[0]**2 + v2[1]**2 + v2[2]**2)

            if mag1 == 0 or mag2 == 0:
                return {"error": "One or more vectors have zero length"}

            # Clamp for numerical stability
            cos_angle = max(-1.0, min(1.0, dot / (mag1 * mag2)))
            angle_rad = math.acos(cos_angle)
            angle_deg = math.degrees(angle_rad)

            return {
                "angle_degrees": angle_deg,
                "angle_radians": angle_rad,
                "vertex": [x2, y2, z2]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_material_table(self) -> Dict[str, Any]:
        """
        Get the available material properties from the document variables.

        Returns material-related variables from the active document.

        Returns:
            Dict with material properties
        """
        try:
            doc = self.doc_manager.get_active_document()

            material_vars = {}
            material_names = ["Density", "Mass", "Volume", "Surface_Area",
                             "YoungsModulus", "PoissonsRatio", "ThermalExpansionCoefficient",
                             "ThermalConductivity", "Material"]

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
                                material_vars[name] = str(var.Formula) if hasattr(var, 'Formula') else "N/A"
                    except Exception:
                        continue
            except Exception:
                pass

            # Also try to get material name from properties
            try:
                if hasattr(doc, 'Properties'):
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

            return {
                "material_properties": material_vars,
                "property_count": len(material_vars)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_feature_dimensions(self, feature_name: str) -> Dict[str, Any]:
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
                        num_dims = result[0]
                        dim_array = result[1]
                    else:
                        dim_array = result
                        num_dims = None

                    # Process dimension objects
                    if dim_array is not None:
                        try:
                            # Try iterating as a collection
                            if hasattr(dim_array, '__iter__'):
                                for dim in dim_array:
                                    dim_info = {}
                                    try:
                                        dim_info["name"] = dim.Name
                                    except Exception:
                                        pass
                                    try:
                                        dim_info["value"] = dim.Value
                                    except Exception:
                                        pass
                                    try:
                                        dim_info["formula"] = dim.Formula
                                    except Exception:
                                        pass
                                    dimensions.append(dim_info)
                        except Exception:
                            dimensions.append({"raw": str(dim_array)})
            except Exception as e:
                # Some features may not support GetDimensions
                return {
                    "feature_name": feature_name,
                    "dimensions": [],
                    "note": f"GetDimensions not supported: {str(e)}"
                }

            return {
                "feature_name": feature_name,
                "dimensions": dimensions,
                "count": len(dimensions)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_material_list(self) -> Dict[str, Any]:
        """
        Get the list of available materials from the material table.

        Uses MatTable.GetMaterialList() via the Solid Edge application object.

        Returns:
            Dict with list of material names
        """
        try:
            doc = self.doc_manager.get_active_document()
            import win32com.client

            # Get MatTable from the application
            app = self.doc_manager.connection.get_application()
            mat_table = app.GetMaterialTable()

            result = mat_table.GetMaterialList()
            # Returns (count, list_of_materials) as out-params
            if isinstance(result, tuple) and len(result) >= 2:
                num_materials = result[0]
                material_list = result[1]
            else:
                material_list = result
                num_materials = None

            materials = []
            if material_list is not None:
                if isinstance(material_list, (list, tuple)):
                    materials = list(material_list)
                elif hasattr(material_list, '__iter__'):
                    materials = [str(m) for m in material_list]

            return {
                "materials": materials,
                "count": len(materials)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def set_material(self, material_name: str) -> Dict[str, Any]:
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

            return {
                "status": "applied",
                "material": material_name
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_material_property(self, material_name: str, property_index: int) -> Dict[str, Any]:
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

            return {
                "material": material_name,
                "property_index": property_index,
                "value": value
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def query_variables(self, pattern: str = "*",
                        case_insensitive: bool = True) -> Dict[str, Any]:
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
                        name = var.Name if hasattr(var, 'Name') else str(i)
                        if case_insensitive:
                            match = fnmatch.fnmatch(name.lower(), pattern.lower())
                        else:
                            match = fnmatch.fnmatch(name, pattern)
                        if match:
                            entry = {"name": name}
                            try:
                                entry["value"] = var.Value
                            except Exception:
                                pass
                            try:
                                entry["formula"] = var.Formula
                            except Exception:
                                pass
                            matches.append(entry)
                    except Exception:
                        continue

                return {
                    "pattern": pattern,
                    "matches": matches,
                    "count": len(matches),
                    "method": "fallback_fnmatch"
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
                        try:
                            entry["value"] = var.Value
                        except Exception:
                            pass
                        try:
                            entry["formula"] = var.Formula
                        except Exception:
                            pass
                        matches.append(entry)
                except Exception:
                    pass

            return {
                "pattern": pattern,
                "matches": matches,
                "count": len(matches)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_variable_formula(self, name: str) -> Dict[str, Any]:
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
                    display_name = var.DisplayName if hasattr(var, 'DisplayName') else ""
                    if display_name == name:
                        result = {"name": name}
                        try:
                            result["formula"] = var.Formula
                        except Exception:
                            result["formula"] = None
                        try:
                            result["value"] = var.Value
                        except Exception:
                            pass
                        return result
                except Exception:
                    continue

            return {"error": f"Variable '{name}' not found"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def rename_variable(self, old_name: str, new_name: str) -> Dict[str, Any]:
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
                    display_name = var.DisplayName if hasattr(var, 'DisplayName') else ""
                    if display_name == old_name:
                        var.DisplayName = new_name
                        return {
                            "status": "renamed",
                            "old_name": old_name,
                            "new_name": new_name
                        }
                except Exception:
                    continue

            return {"error": f"Variable '{old_name}' not found"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_variable_names(self, name: str) -> Dict[str, Any]:
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
                    display_name = var.DisplayName if hasattr(var, 'DisplayName') else ""
                    if display_name == name:
                        result = {"display_name": display_name}
                        try:
                            result["system_name"] = var.Name
                        except Exception:
                            result["system_name"] = None
                        try:
                            result["value"] = var.Value
                        except Exception:
                            pass
                        return result
                except Exception:
                    continue

            return {"error": f"Variable '{name}' not found"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }
