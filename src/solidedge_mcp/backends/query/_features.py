"""Feature tree queries, editing, and extent/treatment operations."""

import contextlib
import traceback
from typing import Any


class FeatureQueryMixin:
    """Mixin providing feature tree query and editing methods."""

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

    def get_vertex_count(self) -> dict[str, Any]:
        """
        Get the total vertex count on the model body.

        Enumerates vertices via faces (same approach as edge counting).

        Returns:
            Dict with vertex count
        """
        try:
            from ..constants import FaceQueryConstants

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

    # =================================================================
    # EXTENTS
    # =================================================================

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

    # =================================================================
    # THIN WALL OPTIONS
    # =================================================================

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

    # =================================================================
    # FACE OFFSET DATA
    # =================================================================

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

    def get_to_face_offset(self, feature_name: str) -> dict[str, Any]:
        """
        Get the 'to face' offset data from a feature.

        Complementary to get_from_face_offset. Returns the to-face/plane,
        offset side, and offset distance for from-to extrusions.

        Args:
            feature_name: Name of the feature

        Returns:
            Dict with to-face offset info (side, distance)
        """
        try:
            feature, doc = self._find_feature(feature_name)
            if feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            result = feature.GetToFaceOffsetData()
            # Returns (ToFaceOrPlane, ToFaceOffsetSide, ToFaceOffsetDistance)
            info: dict[str, Any] = {"feature": feature_name}
            if isinstance(result, tuple) and len(result) >= 3:
                info["to_face_offset_side"] = result[1]
                info["to_face_offset_distance"] = result[2]
            else:
                info["raw_result"] = str(result)

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_to_face_offset(
        self, feature_name: str, offset_side: int, distance: float
    ) -> dict[str, Any]:
        """
        Set the 'to face' offset data on a feature.

        Args:
            feature_name: Name of the feature
            offset_side: Offset side constant (1=igLeft, 2=igRight)
            distance: Offset distance in meters

        Returns:
            Dict with status
        """
        try:
            feature, doc = self._find_feature(feature_name)
            if feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            # Get current to-face to pass it back
            current = feature.GetToFaceOffsetData()
            to_face = current[0] if isinstance(current, tuple) else None

            feature.SetToFaceOffsetData(to_face, offset_side, distance)
            return {
                "status": "set",
                "feature": feature_name,
                "offset_side": offset_side,
                "distance": distance,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # BODY ARRAY
    # =================================================================

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
    # DIRECTION 1 TREATMENT (CROWN/DRAFT)
    # =================================================================

    def get_direction1_treatment(self, feature_name: str) -> dict[str, Any]:
        """
        Get Direction 1 treatment (crown/draft) from a feature.

        Returns treatment type, draft side/angle, crown type/side/curvature/radius/angle.

        Args:
            feature_name: Name of the feature

        Returns:
            Dict with treatment parameters
        """
        try:
            feature, doc = self._find_feature(feature_name)
            if feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            result = feature.GetDirection1Treatment()
            # Returns 8-element tuple:
            # (TreatmentType, DraftSide, DraftAngle, CrownType,
            #  CrownSide, CrownCurvatureSide, CrownRadiusOrOffset, CrownTakeOffAngle)
            info: dict[str, Any] = {"feature": feature_name}
            if isinstance(result, tuple) and len(result) >= 8:
                info["treatment_type"] = result[0]
                info["draft_side"] = result[1]
                info["draft_angle"] = result[2]
                info["crown_type"] = result[3]
                info["crown_side"] = result[4]
                info["crown_curvature_side"] = result[5]
                info["crown_radius_or_offset"] = result[6]
                info["crown_takeoff_angle"] = result[7]
            else:
                info["raw_result"] = str(result)

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def apply_direction1_treatment(
        self,
        feature_name: str,
        treatment_type: int = 0,
        draft_side: int = 0,
        draft_angle: float = 0.0,
        crown_type: int = 0,
        crown_side: int = 0,
        crown_curvature_side: int = 0,
        crown_radius_or_offset: float = 0.0,
        crown_takeoff_angle: float = 0.0,
    ) -> dict[str, Any]:
        """
        Apply Direction 1 treatment (crown/draft) to a feature.

        Args:
            feature_name: Name of the feature
            treatment_type: Treatment type constant
            draft_side: Draft side constant
            draft_angle: Draft angle in radians
            crown_type: Crown type constant
            crown_side: Crown side constant
            crown_curvature_side: Crown curvature side constant
            crown_radius_or_offset: Crown radius or offset in meters
            crown_takeoff_angle: Crown takeoff angle in radians

        Returns:
            Dict with status
        """
        try:
            feature, doc = self._find_feature(feature_name)
            if feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            feature.ApplyDirection1Treatment(
                treatment_type,
                draft_side,
                draft_angle,
                crown_type,
                crown_side,
                crown_curvature_side,
                crown_radius_or_offset,
                crown_takeoff_angle,
            )
            return {
                "status": "applied",
                "feature": feature_name,
                "treatment_type": treatment_type,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
