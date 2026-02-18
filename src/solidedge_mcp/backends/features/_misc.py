"""Miscellaneous feature operations (mirror, patterns, face ops, body ops, simplify, etc.)."""

import contextlib
import math
import traceback
from typing import Any

import pythoncom
from win32com.client import VARIANT

from ..constants import (
    DirectionConstants,
    FaceQueryConstants,
)


class MiscFeaturesMixin:
    """Mixin providing mirror, pattern, face ops, body ops, and simplify methods."""

    def create_pattern(self, pattern_type: str, **kwargs) -> dict[str, Any]:
        """
        Create a pattern of features.

        Note: Feature patterns require SAFEARRAY(IDispatch) marshaling of feature
        objects which is not currently supported via COM late binding. Use assembly-level
        component patterns (pattern_component) instead.

        Args:
            pattern_type: 'Rectangular' or 'Circular'
            **kwargs: Pattern-specific parameters

        Returns:
            Dict with error explaining limitation
        """
        return {
            "error": "Feature patterns (model.Patterns) require SAFEARRAY marshaling of "
            "feature objects which is not supported via COM late binding. "
            "Use assembly-level pattern_component() for component patterns instead.",
            "pattern_type": pattern_type,
        }

    def create_shell(
        self, thickness: float, remove_face_indices: list[int] | None = None
    ) -> dict[str, Any]:
        """
        Create a shell feature (hollow out the part).

        Note: Shell (Thinwalls) requires face selection for open faces which cannot
        be reliably automated via COM late binding. The Thinwalls.Add method requires
        complex VARIANT parameters for face arrays.

        Args:
            thickness: Wall thickness in meters
            remove_face_indices: Indices of faces to remove (optional)

        Returns:
            Dict with error explaining limitation
        """
        return {
            "error": "Shell (Thinwalls) feature requires face selection for open faces "
            "which cannot be reliably automated via COM. Use the Solid Edge UI "
            "to create shell features.",
            "thickness": thickness,
        }

    def list_features(self) -> dict[str, Any]:
        """List all features in the active part"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            features = []
            for i in range(models.Count):
                model = models.Item(i + 1)
                features.append(
                    {
                        "index": i,
                        "name": model.Name if hasattr(model, "Name") else f"Feature {i + 1}",
                        "type": model.Type if hasattr(model, "Type") else "Unknown",
                    }
                )

            return {"features": features, "count": len(features)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_feature_info(self, feature_index: int) -> dict[str, Any]:
        """Get detailed information about a specific feature"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if feature_index < 0 or feature_index >= models.Count:
                return {"error": f"Invalid feature index: {feature_index}"}

            model = models.Item(feature_index + 1)

            info = {
                "index": feature_index,
                "name": model.Name if hasattr(model, "Name") else "Unknown",
                "type": model.Type if hasattr(model, "Type") else "Unknown",
            }

            # Try to get additional properties
            try:
                if hasattr(model, "Visible"):
                    info["visible"] = model.Visible
                if hasattr(model, "Suppressed"):
                    info["suppressed"] = model.Suppressed
            except Exception:
                pass

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_body(self, body_type: str = "Solid") -> dict[str, Any]:
        """
        Add a body to the part.

        Args:
            body_type: Type of body - 'Solid', 'Surface', 'Construction'

        Returns:
            Dict with status and body info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            # AddBody
            models.AddBody()

            return {"status": "created", "type": "body", "body_type": body_type}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def thicken_surface(self, thickness: float, direction: str = "Both") -> dict[str, Any]:
        """
        Thicken a surface to create a solid.

        Args:
            thickness: Thickness (meters)
            direction: 'Both', 'Inside', or 'Outside'

        Returns:
            Dict with status and thicken info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            # AddThickenFeature
            models.AddThickenFeature(Thickness=thickness)

            return {
                "status": "created",
                "type": "thicken",
                "thickness": thickness,
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def auto_simplify(self) -> dict[str, Any]:
        """Auto-simplify the model"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            models.AddAutoSimplify()

            return {"status": "created", "type": "auto_simplify"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def simplify_enclosure(self) -> dict[str, Any]:
        """Create simplified enclosure"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            models.AddSimplifyEnclosure()

            return {"status": "created", "type": "simplify_enclosure"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def simplify_duplicate(self) -> dict[str, Any]:
        """Create simplified duplicate"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            models.AddSimplifyDuplicate()

            return {"status": "created", "type": "simplify_duplicate"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def local_simplify_enclosure(self) -> dict[str, Any]:
        """Create local simplified enclosure"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            models.AddLocalSimplifyEnclosure()

            return {"status": "created", "type": "local_simplify_enclosure"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_mirror(self, feature_name: str, mirror_plane_index: int) -> dict[str, Any]:
        """
        Create a mirror copy of a feature across a reference plane.

        Note: MirrorCopies via COM has known limitations. The ordered-mode
        Add() method creates a feature object but doesn't persist geometry.
        AddSync() persists the feature tree entry but may not compute geometry.
        This is a known Solid Edge COM API limitation.

        Args:
            feature_name: Name of the feature to mirror (from list_features)
            mirror_plane_index: 1-based index of the mirror plane
                (1=Top/XZ, 2=Front/XY, 3=Right/YZ, or higher for user planes)

        Returns:
            Dict with status and mirror info
        """
        try:
            import win32com.client as win32

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists."}

            model = models.Item(1)

            # Find the feature by name in DesignEdgebarFeatures
            features = doc.DesignEdgebarFeatures
            target_feature = None
            for i in range(1, features.Count + 1):
                f = features.Item(i)
                if f.Name == feature_name:
                    target_feature = f
                    break

            if target_feature is None:
                names = []
                for i in range(1, features.Count + 1):
                    names.append(features.Item(i).Name)
                return {
                    "error": f"Feature '{feature_name}' not found.",
                    "available_features": names,
                }

            # Get the mirror plane
            ref_planes = doc.RefPlanes
            if mirror_plane_index < 1 or mirror_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid plane index: {mirror_plane_index}. Count: {ref_planes.Count}"
                }

            mirror_plane = ref_planes.Item(mirror_plane_index)

            # Use AddSync which persists the feature tree entry
            mc = win32.gencache.EnsureDispatch(model.MirrorCopies)
            mirror = mc.AddSync(1, [target_feature], mirror_plane, False)

            return {
                "status": "created",
                "type": "mirror_copy",
                "feature": feature_name,
                "mirror_plane": mirror_plane_index,
                "name": mirror.Name if hasattr(mirror, "Name") else None,
                "note": "Mirror feature created via AddSync. "
                "Geometry may require manual verification "
                "in Solid Edge UI.",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def delete_faces(self, face_indices: list[int]) -> dict[str, Any]:
        """
        Delete faces from the model body.

        Uses model.DeleteFaces collection to remove specified faces.
        Useful for creating openings or removing geometry.

        Args:
            face_indices: List of 0-based face indices to delete

        Returns:
            Dict with status and deletion info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to delete faces from"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if faces.Count == 0:
                return {"error": "No faces on body"}

            face_objs = []
            for idx in face_indices:
                if idx < 0 or idx >= faces.Count:
                    return {"error": f"Invalid face index: {idx}. Body has {faces.Count} faces."}
                face_objs.append(faces.Item(idx + 1))

            delete_faces = model.DeleteFaces
            delete_faces.Add(len(face_objs), face_objs)

            return {
                "status": "created",
                "type": "delete_faces",
                "face_count": len(face_indices),
                "face_indices": face_indices,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def delete_faces_no_heal(self, face_indices: list[int]) -> dict[str, Any]:
        """
        Delete faces from the model body without healing.

        Unlike delete_faces which attempts to heal/close resulting gaps,
        this removes faces leaving the gap open. Useful when you need
        to create deliberate openings.

        Args:
            face_indices: List of 0-based face indices to delete

        Returns:
            Dict with status and deletion info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to delete faces from"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if faces.Count == 0:
                return {"error": "No faces on body"}

            face_objs = []
            for idx in face_indices:
                if idx < 0 or idx >= faces.Count:
                    return {"error": f"Invalid face index: {idx}. Body has {faces.Count} faces."}
                face_objs.append(faces.Item(idx + 1))

            delete_faces = model.DeleteFaces
            delete_faces.AddNoHeal(len(face_objs), face_objs)

            return {
                "status": "created",
                "type": "delete_faces_no_heal",
                "face_count": len(face_indices),
                "face_indices": face_indices,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_body_by_mesh(self) -> dict[str, Any]:
        """Add body by mesh facets"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            models.AddBodyByMeshFacets()

            return {"status": "created", "type": "body_by_mesh"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_body_feature(self) -> dict[str, Any]:
        """Add body feature"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            models.AddBodyFeature()

            return {"status": "created", "type": "body_feature"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_by_construction(self) -> dict[str, Any]:
        """Add construction body"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            models.AddByConstruction()

            return {"status": "created", "type": "construction_body"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_body_by_tag(self, tag: str) -> dict[str, Any]:
        """Add body by tag reference"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            models.AddBodyByTag(tag)

            return {"status": "created", "type": "body_by_tag", "tag": tag}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_face_rotate_by_edge(
        self, face_index: int, edge_index: int, angle: float
    ) -> dict[str, Any]:
        """
        Rotate a face around an edge axis.

        Tilts a face by rotating it around a specified edge. Useful for
        creating draft angles or adjusting face orientations.

        Args:
            face_index: 0-based face index to rotate
            edge_index: 0-based edge index to use as rotation axis
            angle: Rotation angle in degrees

        Returns:
            Dict with status and face rotate info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to rotate faces on"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)

            # Get edge from the face
            face_edges = face.Edges
            if not hasattr(face_edges, "Count") or face_edges.Count == 0:
                return {"error": f"Face {face_index} has no edges"}
            if edge_index < 0 or edge_index >= face_edges.Count:
                return {
                    "error": f"Invalid edge index: {edge_index}. Face has {face_edges.Count} edges."
                }

            edge = face_edges.Item(edge_index + 1)

            angle_rad = math.radians(angle)

            face_rotates = model.FaceRotates
            # igFaceRotateByGeometry = 1, igFaceRotateRecreateBlends = 1, igFaceRotateAxisEnd = 2
            face_rotates.Add(face, 1, 1, None, None, edge, 2, angle_rad)

            return {
                "status": "created",
                "type": "face_rotate",
                "method": "by_edge",
                "face_index": face_index,
                "edge_index": edge_index,
                "angle_degrees": angle,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_face_rotate_by_points(
        self, face_index: int, vertex1_index: int, vertex2_index: int, angle: float
    ) -> dict[str, Any]:
        """
        Rotate a face around an axis defined by two vertex points.

        Args:
            face_index: 0-based face index to rotate
            vertex1_index: 0-based index of first vertex defining rotation axis
            vertex2_index: 0-based index of second vertex defining rotation axis
            angle: Rotation angle in degrees

        Returns:
            Dict with status and face rotate info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to rotate faces on"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)

            # Get vertices from the face
            vertices = face.Vertices
            if vertex1_index < 0 or vertex1_index >= vertices.Count:
                return {
                    "error": f"Invalid vertex1 index: "
                    f"{vertex1_index}. Face has "
                    f"{vertices.Count} vertices."
                }
            if vertex2_index < 0 or vertex2_index >= vertices.Count:
                return {
                    "error": f"Invalid vertex2 index: "
                    f"{vertex2_index}. Face has "
                    f"{vertices.Count} vertices."
                }

            point1 = vertices.Item(vertex1_index + 1)
            point2 = vertices.Item(vertex2_index + 1)

            angle_rad = math.radians(angle)

            face_rotates = model.FaceRotates
            # igFaceRotateByPoints = 2, igFaceRotateRecreateBlends = 1, igFaceRotateNone = 0
            face_rotates.Add(face, 2, 1, point1, point2, None, 0, angle_rad)

            return {
                "status": "created",
                "type": "face_rotate",
                "method": "by_points",
                "face_index": face_index,
                "vertex1_index": vertex1_index,
                "vertex2_index": vertex2_index,
                "angle_degrees": angle,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_draft_angle(
        self, face_index: int, angle: float, plane_index: int = 1
    ) -> dict[str, Any]:
        """
        Add a draft angle to a face.

        Draft angles are used in injection molding to facilitate part removal
        from the mold. Uses the model.Drafts collection.

        Args:
            face_index: 0-based face index to apply draft to
            angle: Draft angle in degrees
            plane_index: 1-based reference plane index for draft direction (default: 1 = Top)

        Returns:
            Dict with status and draft info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add draft to"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)
            ref_plane = doc.RefPlanes.Item(plane_index)

            angle_rad = math.radians(angle)

            # igRight = 2 (draft direction side)
            drafts = model.Drafts
            drafts.Add(ref_plane, 1, [face], [angle_rad], 2)

            return {
                "status": "created",
                "type": "draft_angle",
                "face_index": face_index,
                "angle_degrees": angle,
                "plane_index": plane_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def convert_feature_type(self, feature_name: str, target_type: str) -> dict[str, Any]:
        """
        Convert a feature between cutout and protrusion.

        Uses Feature.ConvertToCutout() or Feature.ConvertToProtrusion()
        to toggle a feature between adding and removing material.

        Args:
            feature_name: Name of the feature (from list_features)
            target_type: 'cutout' or 'protrusion'

        Returns:
            Dict with conversion status and new feature reference
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

            target_type_lower = target_type.lower()
            if target_type_lower == "cutout":
                result = target_feature.ConvertToCutout()
                new_name = None
                with contextlib.suppress(Exception):
                    new_name = result.Name
                return {
                    "status": "converted",
                    "original_name": feature_name,
                    "target_type": "cutout",
                    "new_name": new_name,
                }
            elif target_type_lower == "protrusion":
                result = target_feature.ConvertToProtrusion()
                new_name = None
                with contextlib.suppress(Exception):
                    new_name = result.Name
                return {
                    "status": "converted",
                    "original_name": feature_name,
                    "target_type": "protrusion",
                    "new_name": new_name,
                }
            else:
                return {
                    "error": f"Invalid target_type: {target_type}. Use 'cutout' or 'protrusion'"
                }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_thicken_sync(self, thickness: float, direction: str = "Both") -> dict[str, Any]:
        """
        Create a synchronous thicken feature.

        Uses Thickens.AddSync to thicken a surface body into a solid
        in synchronous modeling mode.

        Args:
            thickness: Thicken thickness in meters
            direction: 'Both', 'Normal', or 'Reverse'

        Returns:
            Dict with status and thicken info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            direction_map = {
                "Both": DirectionConstants.igBoth,
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
            }
            side = direction_map.get(direction, DirectionConstants.igBoth)

            thickens = model.Thickens
            thickens.AddSync(thickness, side)

            return {
                "status": "created",
                "type": "thicken_sync",
                "thickness": thickness,
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_mirror_sync_ex(self, feature_name: str, mirror_plane_index: int) -> dict[str, Any]:
        """
        Create a synchronous mirror copy using the extended AddSyncEx method.

        Looks up the feature by name from DesignEdgebarFeatures and mirrors
        it across the specified reference plane.

        Args:
            feature_name: Name of the feature to mirror (from list_features)
            mirror_plane_index: 1-based index of the mirror plane

        Returns:
            Dict with status and mirror info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists."}

            model = models.Item(1)

            target_feature, error = self._find_feature_by_name(feature_name)
            if error:
                return error

            ref_planes = doc.RefPlanes
            if mirror_plane_index < 1 or mirror_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid plane index: {mirror_plane_index}. Count: {ref_planes.Count}"
                }

            mirror_plane = ref_planes.Item(mirror_plane_index)

            mc = model.MirrorCopies
            mirror = mc.AddSyncEx(1, [target_feature], mirror_plane, False)

            return {
                "status": "created",
                "type": "mirror_sync_ex",
                "feature": feature_name,
                "mirror_plane": mirror_plane_index,
                "name": mirror.Name if hasattr(mirror, "Name") else None,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_pattern_rectangular_ex(
        self,
        feature_name: str,
        x_count: int,
        y_count: int,
        x_spacing: float,
        y_spacing: float,
    ) -> dict[str, Any]:
        """
        Create a rectangular pattern using the extended AddByRectangularEx method.

        The Ex variant may use different marshaling than the original
        AddByRectangular, potentially avoiding SAFEARRAY issues.

        Args:
            feature_name: Name of the feature to pattern
            x_count: Number of instances in X direction
            y_count: Number of instances in Y direction
            x_spacing: Spacing between instances in X (meters)
            y_spacing: Spacing between instances in Y (meters)

        Returns:
            Dict with status and pattern info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists."}

            model = models.Item(1)

            target_feature, error = self._find_feature_by_name(feature_name)
            if error:
                return error

            patterns = model.Patterns
            feature_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [target_feature])
            pattern = patterns.AddByRectangularEx(
                1, feature_arr, x_count, x_spacing, y_count, y_spacing
            )

            return {
                "status": "created",
                "type": "pattern_rectangular_ex",
                "feature": feature_name,
                "x_count": x_count,
                "y_count": y_count,
                "x_spacing": x_spacing,
                "y_spacing": y_spacing,
                "name": pattern.Name if hasattr(pattern, "Name") else None,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_pattern_circular_ex(
        self,
        feature_name: str,
        count: int,
        angle: float,
        axis_face_index: int,
    ) -> dict[str, Any]:
        """
        Create a circular pattern using the extended AddByCircularEx method.

        The Ex variant may use different marshaling than the original
        AddByCircular, potentially avoiding SAFEARRAY issues.

        Args:
            feature_name: Name of the feature to pattern
            count: Number of instances around the circle
            angle: Total angle in degrees
            axis_face_index: 0-based index of the cylindrical face to use as axis

        Returns:
            Dict with status and pattern info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists."}

            model = models.Item(1)

            target_feature, error = self._find_feature_by_name(feature_name)
            if error:
                return error

            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)

            if axis_face_index < 0 or axis_face_index >= faces.Count:
                return {
                    "error": f"Invalid axis_face_index: {axis_face_index}. Count: {faces.Count}"
                }

            axis_face = faces.Item(axis_face_index + 1)
            angle_rad = math.radians(angle)

            patterns = model.Patterns
            feature_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [target_feature])
            pattern = patterns.AddByCircularEx(1, feature_arr, count, angle_rad, axis_face)

            return {
                "status": "created",
                "type": "pattern_circular_ex",
                "feature": feature_name,
                "count": count,
                "angle": angle,
                "axis_face_index": axis_face_index,
                "name": pattern.Name if hasattr(pattern, "Name") else None,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_pattern_duplicate(self, feature_name: str) -> dict[str, Any]:
        """
        Create a duplicate pattern of a feature.

        Uses Patterns.AddDuplicate to create an exact copy of the
        specified feature in the feature tree.

        Args:
            feature_name: Name of the feature to duplicate

        Returns:
            Dict with status and pattern info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists."}

            model = models.Item(1)

            target_feature, error = self._find_feature_by_name(feature_name)
            if error:
                return error

            patterns = model.Patterns
            feature_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [target_feature])
            pattern = patterns.AddDuplicate(1, feature_arr)

            return {
                "status": "created",
                "type": "pattern_duplicate",
                "feature": feature_name,
                "name": pattern.Name if hasattr(pattern, "Name") else None,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_pattern_by_fill(
        self,
        feature_name: str,
        fill_region_face_index: int,
        x_spacing: float,
        y_spacing: float,
    ) -> dict[str, Any]:
        """
        Create a fill pattern of a feature within a face region.

        Uses Patterns.AddByFill to fill a face region with patterned
        copies of the specified feature.

        Args:
            feature_name: Name of the feature to pattern
            fill_region_face_index: 0-based face index defining the fill region
            x_spacing: Spacing in X direction (meters)
            y_spacing: Spacing in Y direction (meters)

        Returns:
            Dict with status and pattern info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists."}

            model = models.Item(1)

            target_feature, error = self._find_feature_by_name(feature_name)
            if error:
                return error

            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)

            if fill_region_face_index < 0 or fill_region_face_index >= faces.Count:
                return {
                    "error": f"Invalid fill_region_face_index: {fill_region_face_index}. "
                    f"Count: {faces.Count}"
                }

            fill_face = faces.Item(fill_region_face_index + 1)

            patterns = model.Patterns
            feature_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [target_feature])
            pattern = patterns.AddByFill(1, feature_arr, fill_face, x_spacing, y_spacing)

            return {
                "status": "created",
                "type": "pattern_by_fill",
                "feature": feature_name,
                "fill_region_face_index": fill_region_face_index,
                "x_spacing": x_spacing,
                "y_spacing": y_spacing,
                "name": pattern.Name if hasattr(pattern, "Name") else None,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_pattern_by_table(
        self,
        feature_name: str,
        x_offsets: list[float],
        y_offsets: list[float],
    ) -> dict[str, Any]:
        """
        Create a table-driven pattern of a feature.

        Uses Patterns.AddPatternByTable to place copies of the feature
        at specific X/Y offset locations.

        Args:
            feature_name: Name of the feature to pattern
            x_offsets: List of X offsets in meters
            y_offsets: List of Y offsets in meters (must match length of x_offsets)

        Returns:
            Dict with status and pattern info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists."}

            model = models.Item(1)

            if len(x_offsets) != len(y_offsets):
                return {
                    "error": f"x_offsets and y_offsets must have same length. "
                    f"Got {len(x_offsets)} and {len(y_offsets)}."
                }

            if len(x_offsets) == 0:
                return {"error": "At least one offset pair is required."}

            target_feature, error = self._find_feature_by_name(feature_name)
            if error:
                return error

            patterns = model.Patterns
            feature_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [target_feature])
            x_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, x_offsets)
            y_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, y_offsets)

            pattern = patterns.AddPatternByTable(1, feature_arr, len(x_offsets), x_arr, y_arr)

            return {
                "status": "created",
                "type": "pattern_by_table",
                "feature": feature_name,
                "point_count": len(x_offsets),
                "x_offsets": x_offsets,
                "y_offsets": y_offsets,
                "name": pattern.Name if hasattr(pattern, "Name") else None,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_pattern_by_table_sync(
        self,
        feature_name: str,
        x_offsets: list[float],
        y_offsets: list[float],
    ) -> dict[str, Any]:
        """
        Create a synchronous table-driven pattern of a feature.

        Uses Patterns.AddPatternByTableSync.

        Args:
            feature_name: Name of the feature to pattern
            x_offsets: List of X offsets in meters
            y_offsets: List of Y offsets in meters

        Returns:
            Dict with status and pattern info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists."}

            model = models.Item(1)

            if len(x_offsets) != len(y_offsets):
                return {
                    "error": f"x_offsets and y_offsets must have same length. "
                    f"Got {len(x_offsets)} and {len(y_offsets)}."
                }

            if len(x_offsets) == 0:
                return {"error": "At least one offset pair is required."}

            target_feature, error = self._find_feature_by_name(feature_name)
            if error:
                return error

            patterns = model.Patterns
            feature_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [target_feature])
            x_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, x_offsets)
            y_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, y_offsets)

            pattern = patterns.AddPatternByTableSync(1, feature_arr, len(x_offsets), x_arr, y_arr)

            return {
                "status": "created",
                "type": "pattern_by_table_sync",
                "feature": feature_name,
                "point_count": len(x_offsets),
                "name": pattern.Name if hasattr(pattern, "Name") else None,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_pattern_by_fill_ex(
        self,
        feature_name: str,
        fill_region_face_index: int,
        x_spacing: float,
        y_spacing: float,
        stagger_offset: float = 0.0,
    ) -> dict[str, Any]:
        """
        Create an extended fill pattern of a feature within a face region.

        Uses Patterns.AddByFillEx.

        Args:
            feature_name: Name of the feature to pattern
            fill_region_face_index: 0-based face index defining the fill region
            x_spacing: X direction spacing in meters
            y_spacing: Y direction spacing in meters
            stagger_offset: Stagger offset for pattern rows in meters

        Returns:
            Dict with status and pattern info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists."}

            model = models.Item(1)

            target_feature, error = self._find_feature_by_name(feature_name)
            if error:
                return error

            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if fill_region_face_index < 0 or fill_region_face_index >= faces.Count:
                return {
                    "error": f"Invalid fill_region_face_index: "
                    f"{fill_region_face_index}. "
                    f"Body has {faces.Count} faces."
                }

            face = faces.Item(fill_region_face_index + 1)

            feature_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [target_feature])
            region_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [face])

            patterns = model.Patterns
            pattern = patterns.AddByFillEx(
                1,
                feature_arr,
                1,
                region_arr,
                0,
                x_spacing,
                y_spacing,
                stagger_offset,
                0.0,
                False,
            )

            return {
                "status": "created",
                "type": "pattern_by_fill_ex",
                "feature": feature_name,
                "fill_region_face_index": fill_region_face_index,
                "x_spacing": x_spacing,
                "y_spacing": y_spacing,
                "name": pattern.Name if hasattr(pattern, "Name") else None,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_pattern_by_curve_ex(
        self,
        feature_name: str,
        curve_edge_index: int,
        count: int,
        spacing: float,
    ) -> dict[str, Any]:
        """
        Create a pattern along a curve using the extended API.

        Uses Patterns.AddByCurveEx.

        Args:
            feature_name: Name of the feature to pattern
            curve_edge_index: 0-based edge index defining the curve path
            count: Number of pattern occurrences
            spacing: Spacing between occurrences in meters

        Returns:
            Dict with status and pattern info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists."}

            model = models.Item(1)

            target_feature, error = self._find_feature_by_name(feature_name)
            if error:
                return error

            body = model.Body
            edges = body.Edges(FaceQueryConstants.igQueryAll)
            if curve_edge_index < 0 or curve_edge_index >= edges.Count:
                return {
                    "error": f"Invalid curve_edge_index: "
                    f"{curve_edge_index}. "
                    f"Body has {edges.Count} edges."
                }

            edge = edges.Item(curve_edge_index + 1)

            feature_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [target_feature])
            curve_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [edge])

            patterns = model.Patterns
            pattern = patterns.AddByCurveEx(
                1,
                feature_arr,
                0,
                1,
                curve_arr,
                None,
                0,
                0.0,
                0,
                count,
                spacing,
            )

            return {
                "status": "created",
                "type": "pattern_by_curve_ex",
                "feature": feature_name,
                "curve_edge_index": curve_edge_index,
                "count": count,
                "spacing": spacing,
                "name": pattern.Name if hasattr(pattern, "Name") else None,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def save_as_mirror_part(
        self,
        new_file_name: str,
        mirror_plane_index: int = 3,
        link_to_original: bool = True,
    ) -> dict[str, Any]:
        """
        Save the active part as a mirrored copy.

        Creates a new part file that is a mirror of the current part about
        the specified reference plane.

        Args:
            new_file_name: Full file path for the mirrored part (.par)
            mirror_plane_index: 1-based index of the mirror plane (1=Top, 2=Front, 3=Right)
            link_to_original: If True, the mirror part maintains a link to the original

        Returns:
            Dict with status and file info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No model found in active document"}

            ref_planes = doc.RefPlanes
            if mirror_plane_index < 1 or mirror_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid mirror_plane_index: {mirror_plane_index}. "
                    f"Document has {ref_planes.Count} reference planes."
                }

            mirror_plane = ref_planes.Item(mirror_plane_index)

            models.SaveAsMirrorPart(new_file_name, mirror_plane, link_to_original)

            return {
                "status": "saved",
                "type": "mirror_part",
                "path": new_file_name,
                "mirror_plane_index": mirror_plane_index,
                "linked": link_to_original,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_user_defined_pattern(
        self, feature_name: str
    ) -> dict[str, Any]:
        """
        Create a user-defined pattern using accumulated profiles as occurrence locations.

        Each accumulated profile defines where a copy of the seed feature will be placed.
        Create 2+ sketches with single points, close each, then call this method.

        Args:
            feature_name: Name of the seed feature to pattern

        Returns:
            Dict with status and pattern info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No model found in active document"}
            model = models.Item(1)

            # Find the seed feature
            features = doc.DesignEdgebarFeatures
            seed_feature = None
            for i in range(1, features.Count + 1):
                feat = features.Item(i)
                with contextlib.suppress(Exception):
                    if feat.Name == feature_name:
                        seed_feature = feat
                        break

            if seed_feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            all_profiles = self.sketch_manager.get_accumulated_profiles()
            if len(all_profiles) < 1:
                return {
                    "error": "User-defined pattern requires at least 1 profile "
                    "to define occurrence locations."
                }

            profiles_var = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, all_profiles
            )

            udp = model.UserDefinedPatterns
            udp.AddByProfiles(len(all_profiles), profiles_var, seed_feature)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "user_defined_pattern",
                "seed_feature": feature_name,
                "num_occurrences": len(all_profiles),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
