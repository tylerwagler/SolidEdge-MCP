"""
Solid Edge Query and Inspection Operations

Handles querying model data, measurements, and properties.
"""

from typing import Dict, Any
import math
import traceback


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

        Uses Body.Faces(igQueryAll=6) to enumerate faces with their
        type and edge count.

        Returns:
            Dict with list of faces
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body

            faces = body.Faces(6)  # igQueryAll = 6
            face_list = []

            for i in range(1, faces.Count + 1):
                try:
                    face = faces.Item(i)
                    face_info = {"index": i - 1}
                    try:
                        face_info["type"] = face.Type
                    except Exception:
                        pass
                    try:
                        face_info["area"] = face.Area
                    except Exception:
                        pass
                    try:
                        edge_count = face.Edges.Count if hasattr(face.Edges, 'Count') else 0
                        face_info["edge_count"] = edge_count
                    except Exception:
                        pass
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

            faces = body.Faces(6)  # igQueryAll = 6
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

            faces = body.Faces(6)  # igQueryAll = 6
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
                # seModelingModeOrdered = 1, seModelingModeSynchronous = 2
                mode_name = "ordered" if mode == 1 else "synchronous" if mode == 2 else f"unknown ({mode})"
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
                "ordered": 1,      # seModelingModeOrdered
                "synchronous": 2,  # seModelingModeSynchronous
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
                    "old_mode": "ordered" if old_mode == 1 else "synchronous",
                    "new_mode": "ordered" if new_mode == 1 else "synchronous"
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

            faces = body.Faces(6)  # igQueryAll = 6
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
            faces = body.Faces(6)  # igQueryAll = 6

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
            faces = body.Faces(6)
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
