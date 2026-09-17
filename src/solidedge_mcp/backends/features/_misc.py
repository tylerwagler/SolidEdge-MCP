"""Miscellaneous feature operations (mirror, patterns, face ops, body ops, simplify, etc.)."""

import contextlib
import math
from typing import Any

from solidedge_mcp.backends.errors import error_result

from ..comutil import com_get
from ..constants import (
    FaceQueryConstants,
    FaceRotateConstants,
    PatternOffsetTypeConstants,
    PatternTypeConstants,
)
from ..logging import get_logger
from ..validation import guard_overwrite
from ._base import verifies_collection_growth, verifies_geometry

_logger = get_logger(__name__)


# constant.tlb > AddBodyTypeConstants. Kept local because backends/constants.py
# does not carry this enum yet.
_ADD_BODY_TYPES = {
    "Solid": 1,  # igPartType
    "Part": 1,  # igPartType
    "SheetMetal": 2,  # igSheetMetalType
    "Construction": 5,  # igConstructionPartType
}

#: Drafts.Add takes a DraftSide from FeaturePropertyConstants. It is inside or
#: outside; igLeft and igRight are not sides a draft has, and passing one
#: fails with a bare E_FAIL.
_DRAFT_SIDES = {"inside": 4, "outside": 5}  # igInside, igOutside


class MiscFeaturesMixin:
    """Mixin providing mirror, pattern, face ops, body ops, and simplify methods."""

    @verifies_geometry
    # self is positional-only so @verifies_geometry's ParamSpec can bind the
    # **kwargs; every other decorated creator has a fixed signature.
    def create_pattern(self, /, pattern_type: str, **kwargs: Any) -> dict[str, Any]:
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
        del remove_face_indices  # no COM call to pass them to; see above
        return {
            "error": "Shell (Thinwalls) feature requires face selection for open faces "
            "which cannot be reliably automated via COM. Use the Solid Edge UI "
            "to create shell features.",
            "unsupported": True,
            "thickness": thickness,
        }

    def list_features(self) -> dict[str, Any]:
        """List the features of every body in the active part, in tree order.

        Reference planes are not features and do not appear here; read
        ``solidedge://model/edgebar-features`` for the full Pathfinder tree.
        """
        try:
            doc = self.doc_manager.get_active_document()
            features = self._enumerate_features(doc)
            return {"features": features, "count": len(features)}
        except Exception as e:
            return error_result(e)

    def get_feature_info(self, feature_index: int) -> dict[str, Any]:
        """Describe one feature, indexed 0-based into list_features()."""
        try:
            doc = self.doc_manager.get_active_document()
            features = self._enumerate_features(doc)

            if feature_index < 0 or feature_index >= len(features):
                return {
                    "error": (
                        f"Invalid feature index: {feature_index}. "
                        f"The active part has {len(features)} feature(s)."
                    )
                }

            entry = features[feature_index]
            model = doc.Models.Item(entry["body_index"] + 1)
            # The entry records its 1-based position in its own body, so the
            # index no longer has to be recomputed from the flat list.
            feature = model.Features.Item(entry["position_in_body"])

            info: dict[str, Any] = dict(entry)
            # Suppress is a read/write VT_BOOL property on part features.
            suppressed = com_get(feature, "Suppress")
            if suppressed is not None:
                info["suppressed"] = bool(suppressed)
            visible = com_get(feature, "Visible")
            if visible is not None:
                info["visible"] = bool(visible)
            return info
        except Exception as e:
            return error_result(e)

    def add_body(self, body_type: str = "Solid", body_name: str = "") -> dict[str, Any]:
        """
        Add a body to the part.

        Uses Models.AddBody(igBodyType, BodyName).

        Args:
            body_type: 'Solid' (or 'Part'), 'SheetMetal', 'Construction'
            body_name: Name for the new body (defaults to 'Body')

        Returns:
            Dict with status and body info
        """
        try:
            ig_body_type = _ADD_BODY_TYPES.get(body_type)
            if ig_body_type is None:
                return {
                    "error": f"Unknown body_type: {body_type}. "
                    f"Use one of {sorted(_ADD_BODY_TYPES)}."
                }

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            name = body_name or "Body"
            # Models.AddBody(igBodyType: AddBodyTypeConstants, BodyName: BSTR)
            models.AddBody(ig_body_type, name)

            return {
                "status": "created",
                "type": "body",
                "body_type": body_type,
                "body_name": name,
            }
        except Exception as e:
            return error_result(e)

    def thicken_surface(self, thickness: float, direction: str = "Both") -> dict[str, Any]:
        """
        Thicken a surface to create a solid.

        NOT AVAILABLE via COM automation. Models.AddThickenFeature takes
        (Side, offsetDistance, NumberOfFaces, Faces) - the Faces argument is a
        SAFEARRAY of the surface faces to thicken, and this server has no way to
        select the faces of a construction/surface body. The call is never made.

        Args:
            thickness: Thickness (meters)
            direction: 'Both', 'Inside', or 'Outside'

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "Thicken is not available through this server: "
                "Models.AddThickenFeature(Side, offsetDistance, NumberOfFaces, "
                "Faces) requires the surface faces to thicken, which cannot be "
                "selected through this API. Thicken the surface in the Solid "
                "Edge UI."
            ),
            "unsupported": True,
            "type": "thicken",
            "thickness": thickness,
            "direction": direction,
        }

    def auto_simplify(self) -> dict[str, Any]:
        """
        Auto-simplify the model.

        NOT AVAILABLE via COM automation. Models.AddAutoSimplify(numInputs,
        Occurrences, vbRemoveInternals, BodyName) needs an array of assembly
        occurrences, which a part document cannot supply. The call is never made.
        """
        return {
            "error": (
                "Auto-simplify is not available through this server: "
                "Models.AddAutoSimplify(numInputs, Occurrences, "
                "vbRemoveInternals, BodyName) requires an array of assembly "
                "occurrences that this server cannot select. Use the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "auto_simplify",
        }

    def simplify_enclosure(self) -> dict[str, Any]:
        """
        Create a simplified enclosure.

        NOT AVAILABLE via COM automation. Models.AddSimplifyEnclosure(numInputs,
        Occurrences, RefPlane, EncloseType, BodyName) needs an array of assembly
        occurrences, which this server cannot select. The call is never made.
        """
        return {
            "error": (
                "Simplify enclosure is not available through this server: "
                "Models.AddSimplifyEnclosure(numInputs, Occurrences, RefPlane, "
                "EncloseType, BodyName) requires an array of assembly "
                "occurrences that this server cannot select. Use the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "simplify_enclosure",
        }

    def simplify_duplicate(self) -> dict[str, Any]:
        """
        Create a simplified duplicate.

        NOT AVAILABLE via COM automation. Models.AddSimplifyDuplicate(NumBodies,
        Bodies, FromOccurrence, numOccurrences, ToOccurrences, Name) needs body
        and occurrence objects this server cannot select. The call is never made.
        """
        return {
            "error": (
                "Simplify duplicate is not available through this server: "
                "Models.AddSimplifyDuplicate(NumBodies, Bodies, FromOccurrence, "
                "numOccurrences, ToOccurrences, Name) requires body and assembly "
                "occurrence objects that this server cannot select. Use the "
                "Solid Edge UI."
            ),
            "unsupported": True,
            "type": "simplify_duplicate",
        }

    def local_simplify_enclosure(self) -> dict[str, Any]:
        """
        Create a local simplified enclosure.

        NOT AVAILABLE via COM automation.
        Models.AddLocalSimplifyEnclosure(numInputs, TopologyProxies, RefPlane,
        EncloseType, BodyName) needs topology proxy objects (a user selection),
        which this server cannot obtain. The call is never made.
        """
        return {
            "error": (
                "Local simplify enclosure is not available through this server: "
                "Models.AddLocalSimplifyEnclosure(numInputs, TopologyProxies, "
                "RefPlane, EncloseType, BodyName) requires topology proxy "
                "objects from a user selection. Use the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "local_simplify_enclosure",
        }

    @verifies_geometry
    def create_mirror(
        self, feature_name: str, mirror_plane_index: int, allow_mode_switch: bool = False
    ) -> dict[str, Any]:
        """
        Create a mirror copy of a feature across a reference plane.

        Note: MirrorCopies via COM has known limitations. The ordered-mode
        Add() method creates a feature object but doesn't persist geometry.
        AddSync() persists the feature tree entry but may not compute geometry.
        This is a known Solid Edge COM API limitation.

        Args:
            feature_name: Name of the feature to mirror (from list_features)
            mirror_plane_index: 1-based index of the mirror plane
                (1=Top/XY, 2=Right/YZ, 3=Front/XZ, or higher for user planes)

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

            # MirrorCopies.AddSync is synchronous-only: in an ordered part it
            # raises a bare E_FAIL. Verified against Solid Edge 2026, where the
            # same call doubles the face count once the mode is switched.
            err = self._require_synchronous(doc, allow_switch=allow_mode_switch)
            if err:
                return err

            mc = win32.gencache.EnsureDispatch(model.MirrorCopies)
            try:
                mirror = mc.AddSync(1, [target_feature], mirror_plane, False)
            except Exception as exc:
                # Switching an existing ordered part to synchronous is not
                # enough: AddSync mirrors synchronous geometry, so a feature
                # that was built in ordered mode still fails. Verified on
                # Solid Edge 2026, where the same mirror succeeds when the
                # part was in synchronous mode before the feature was made.
                return error_result(
                    exc,
                    context=(
                        "MirrorCopies.AddSync could not mirror "
                        f"'{feature_name}'. It mirrors synchronous geometry, so the "
                        "feature must have been created while the part was in "
                        "synchronous mode. Build the part synchronously from the "
                        "start, or mirror the sketch and re-create the feature."
                    ),
                )

            return {
                "status": "created",
                "type": "mirror_copy",
                "feature": feature_name,
                "mirror_plane": mirror_plane_index,
                "name": com_get(mirror, "Name", None),
                "note": "Mirror feature created via AddSync. "
                "Geometry may require manual verification "
                "in Solid Edge UI.",
            }
        except Exception as e:
            return error_result(e)

    def delete_faces(self, face_indices: list[int]) -> dict[str, Any]:
        """
        Delete faces from the model body.

        NOT AVAILABLE via COM automation. DeleteFaces.Add takes a single
        argument, FaceSetToDelete (VT_DISPATCH) - a FaceSet object. No
        collection in the Part type library hands out a FaceSet, so this server
        cannot build the argument. The call is never made.

        Args:
            face_indices: List of 0-based face indices to delete

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "Delete faces is not available through this server: "
                "DeleteFaces.Add(FaceSetToDelete) takes one FaceSet object, and "
                "the Solid Edge Part API exposes no way to build a FaceSet from "
                "face indices. Delete the faces in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "delete_faces",
            "face_indices": list(face_indices),
        }

    def delete_faces_no_heal(self, face_indices: list[int]) -> dict[str, Any]:
        """
        Delete faces from the model body without healing.

        NOT AVAILABLE via COM automation. DeleteFaces.AddNoHeal takes a single
        FaceSetToDelete (VT_DISPATCH) argument, which this server cannot build.
        The call is never made.

        Args:
            face_indices: List of 0-based face indices to delete

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "Delete faces (no heal) is not available through this server: "
                "DeleteFaces.AddNoHeal(FaceSetToDelete) takes one FaceSet "
                "object, and the Solid Edge Part API exposes no way to build a "
                "FaceSet from face indices. Delete the faces in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "delete_faces_no_heal",
            "face_indices": list(face_indices),
        }

    def add_body_by_mesh(self) -> dict[str, Any]:
        """
        Add a body from mesh facets.

        NOT AVAILABLE via COM automation.
        Models.AddBodyByMeshFacets(NumberOfVertices, VertextPostionsOfFacets)
        needs an explicit array of facet vertex coordinates, which this server
        has no source for. The call is never made.
        """
        return {
            "error": (
                "Add body by mesh is not available through this server: "
                "Models.AddBodyByMeshFacets(NumberOfVertices, "
                "VertextPostionsOfFacets) requires an array of facet vertex "
                "coordinates. Import the mesh through the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "body_by_mesh",
        }

    def add_body_feature(self, import_file_name: str = "") -> dict[str, Any]:
        """
        Add a body feature by importing a body from a file.

        Uses Models.AddBodyFeature(ImportFileName).

        Args:
            import_file_name: Full path of the file to import the body from

        Returns:
            Dict with status and body info
        """
        if not import_file_name:
            return {
                "error": (
                    "Models.AddBodyFeature(ImportFileName) needs the path of the "
                    "file to import the body from; pass import_file_name."
                )
            }
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            # Models.AddBodyFeature(ImportFileName: BSTR)
            models.AddBodyFeature(import_file_name)

            return {
                "status": "created",
                "type": "body_feature",
                "import_file_name": import_file_name,
            }
        except Exception as e:
            return error_result(e)

    def add_by_construction(self, construction_index: int = 0) -> dict[str, Any]:
        """
        Add a body from an existing construction (surface) body.

        Uses Models.AddByConstruction(ConstructionSolid).

        Args:
            construction_index: 0-based index into doc.Constructions

        Returns:
            Dict with status and body info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            constructions = doc.Constructions
            if constructions.Count == 0:
                return {"error": "No construction bodies exist in the active document."}
            if construction_index < 0 or construction_index >= constructions.Count:
                return {
                    "error": f"Invalid construction_index: {construction_index}. "
                    f"Document has {constructions.Count} constructions."
                }

            construction = constructions.Item(construction_index + 1)

            # Models.AddByConstruction(ConstructionSolid: VT_DISPATCH)
            models.AddByConstruction(construction)

            return {
                "status": "created",
                "type": "construction_body",
                "construction_index": construction_index,
            }
        except Exception as e:
            return error_result(e)

    def add_body_by_tag(self, tag: str) -> dict[str, Any]:
        """Add body by tag reference"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            models.AddBodyByTag(tag)

            return {"status": "created", "type": "body_by_tag", "tag": tag}
        except Exception as e:
            return error_result(e)

    @verifies_collection_growth("Models.*.FaceRotates")
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
            if com_get(face_edges, "Count", 0) == 0:
                return {"error": f"Face {face_index} has no edges"}
            if edge_index < 0 or edge_index >= face_edges.Count:
                return {
                    "error": f"Invalid edge index: {edge_index}. Face has {face_edges.Count} edges."
                }

            edge = face_edges.Item(edge_index + 1)

            angle_rad = math.radians(angle)

            face_rotates = model.FaceRotates
            # FaceRotates.Add(FacesToBeRotated, FaceRotateType, BlendRecreation,
            #   startPointFor2PointAxis, endPointFor2PointAxis, axisObject,
            #   axisStartOrEnd, Angle). The literals this replaced (1, 1, ..., 2)
            # named the wrong members and raised E_INVALIDARG on every call.
            face_rotates.Add(
                face,
                FaceRotateConstants.igFaceRotateByGeometry,
                FaceRotateConstants.igFaceRotateRecreateBlends,
                None,
                None,
                edge,
                FaceRotateConstants.igFaceRotateAxisEnd,
                angle_rad,
            )

            return {
                "status": "created",
                "type": "face_rotate",
                "method": "by_edge",
                "face_index": face_index,
                "edge_index": edge_index,
                "angle_degrees": angle,
            }
        except Exception as e:
            return error_result(e)

    @verifies_collection_growth("Models.*.FaceRotates")
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
            # Same signature as by edge; (2, 1, ..., 0) raised E_INVALIDARG.
            face_rotates.Add(
                face,
                FaceRotateConstants.igFaceRotateByPoints,
                FaceRotateConstants.igFaceRotateRecreateBlends,
                point1,
                point2,
                None,
                FaceRotateConstants.igFaceRotateNone,
                angle_rad,
            )

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
            return error_result(e)

    @verifies_collection_growth("Models.*.Drafts")
    def create_draft_angle(
        self, face_index: int, angle: float, plane_index: int = 1, side: str = "inside"
    ) -> dict[str, Any]:
        """Add a draft angle to a face.

        Draft angles let a moulded part leave its mould.
        ``Drafts.Add(DraftPlane, NumberOfFaceSets, FaceSetArray,
        DraftAngleArray, DraftSide)``.

        DraftSide takes igInside (4) or igOutside (5). It was passing igRight
        (2), which is not a side a draft has, so every draft failed with a bare
        E_FAIL whatever face or plane was named. Verified on Solid Edge 2026:
        4 and 5 both work on every face of a box, 1, 2 and 3 all fail. The
        face array goes as a plain list; no VARIANT wrapper is needed.

        Args:
            face_index: 0-based face to draft.
            angle: Draft angle in degrees.
            plane_index: 1-based reference plane giving the pull direction
                (1=Top/XY, 2=Right/YZ, 3=Front/XZ).
            side: 'inside' or 'outside'.

        Returns:
            Dict with status and draft info.
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

            side_const = _DRAFT_SIDES.get(side.strip().lower())
            if side_const is None:
                return {
                    "error": f"Invalid side: {side}. Use 'inside' or 'outside'.",
                }

            drafts = model.Drafts
            drafts.Add(ref_plane, 1, [face], [angle_rad], side_const)

            return {
                "status": "created",
                "type": "draft_angle",
                "face_index": face_index,
                "angle_degrees": angle,
                "plane_index": plane_index,
                "side": side.strip().lower(),
            }
        except Exception as e:
            return error_result(e)

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

            target_type_lower = target_type.strip().lower()
            # ConvertToCutoutAllowed and ConvertToProtrusionAllowed are
            # properties, not methods, so read them rather than call them.
            # Absent means the feature type does not convert at all.
            allowed_name = {
                "cutout": "ConvertToCutoutAllowed",
                "protrusion": "ConvertToProtrusionAllowed",
            }.get(target_type_lower)
            if allowed_name is not None:
                if not hasattr(target_feature, f"ConvertTo{target_type_lower.capitalize()}"):
                    return {
                        "error": (
                            f"Feature '{feature_name}' cannot become a "
                            f"{target_type_lower}. Only extruded and revolved "
                            f"protrusions and cutouts convert."
                        )
                    }
                if getattr(target_feature, allowed_name, True) is False:
                    return {
                        "error": (
                            f"Solid Edge will not convert '{feature_name}' to a "
                            f"{target_type_lower}. Converting the feature that "
                            f"creates the material would leave nothing behind."
                        )
                    }

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
            return error_result(e)

    @verifies_geometry
    def create_thicken_sync(self, thickness: float, direction: str = "Both") -> dict[str, Any]:
        """
        Create a synchronous thicken feature.

        NOT AVAILABLE via COM automation. Thickens.AddSync takes
        (Side, dOffsetDistance, Faces, Loop): the Faces argument is a SAFEARRAY
        of the surface faces to thicken and Loop is the bounding loop, neither of
        which this server can select. The call is never made.

        Args:
            thickness: Thicken thickness in meters
            direction: 'Both', 'Normal', or 'Reverse'

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "Synchronous thicken is not available through this server: "
                "Thickens.AddSync(Side, dOffsetDistance, Faces, Loop) requires "
                "the surface faces and bounding loop to thicken, which cannot be "
                "selected through this API. Thicken the surface in the Solid "
                "Edge UI."
            ),
            "unsupported": True,
            "type": "thicken_sync",
            "thickness": thickness,
            "direction": direction,
        }

    @verifies_geometry
    def create_mirror_sync_ex(self, feature_name: str, mirror_plane_index: int) -> dict[str, Any]:
        """
        Create a synchronous mirror copy using the extended AddSyncEx method.

        NOT AVAILABLE via COM automation. MirrorCopies.AddSyncEx takes
        (NumberOfFeatures, FeatureArray, MirrorPlane, MirrorOption,
        MirrorFeatures) where MirrorFeatures is an [in,out] SAFEARRAY that late
        binding cannot supply byref. The call is never made; use create_mirror
        (MirrorCopies.AddSync), which has a usable signature.

        Args:
            feature_name: Name of the feature to mirror (from list_features)
            mirror_plane_index: 1-based index of the mirror plane

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "MirrorCopies.AddSyncEx is not usable through COM late binding: "
                "it needs an [in,out] SAFEARRAY (MirrorFeatures) that cannot be "
                "passed byref. Use create_mirror (method='basic', "
                "MirrorCopies.AddSync) instead, or mirror in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "mirror_sync_ex",
            "feature": feature_name,
            "mirror_plane": mirror_plane_index,
        }

    @verifies_geometry
    def create_pattern_rectangular_ex(
        self,
        feature_name: str,
        x_count: int,
        y_count: int,
        x_spacing: float,
        y_spacing: float,
        plane_index: int = 1,
        rectangle_angle: float = 0.0,
    ) -> dict[str, Any]:
        """
        Create a rectangular pattern using the extended AddByRectangularEx method.

        Full signature: AddByRectangularEx(NumberOfFeatures, FeatureArray,
        ReferencePlane, XDirectionCount, YDirectionCount, XDirectionSpacing,
        YDirectionSpacing, RectangleAngle, PatternMethod, ReferenceIndex,
        PatternType).

        Args:
            feature_name: Name of the feature to pattern
            x_count: Number of instances in X direction
            y_count: Number of instances in Y direction
            x_spacing: Spacing between instances in X (meters)
            y_spacing: Spacing between instances in Y (meters)
            plane_index: 1-based reference plane the pattern is laid out on
            rectangle_angle: Rotation of the pattern grid in degrees

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

            ref_planes = doc.RefPlanes
            if plane_index < 1 or plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid plane_index: {plane_index}. "
                    f"Document has {ref_planes.Count} reference planes."
                }
            ref_plane = ref_planes.Item(plane_index)

            patterns = model.Patterns
            # A plain sequence, not a VARIANT: Patterns.AddByRectangularEx
            # rejects VARIANT(VT_ARRAY | VT_DISPATCH, ...) with "Objects for
            # SAFEARRAYS must be sequences", verified against Solid Edge 2026.
            feature_arr = [target_feature]
            pattern = patterns.AddByRectangularEx(
                1,
                feature_arr,
                ref_plane,
                x_count,
                y_count,
                x_spacing,
                y_spacing,
                math.radians(rectangle_angle),
                PatternOffsetTypeConstants.sePatternFixedOffset,
                0,
                PatternTypeConstants.seSmartPattern,
            )

            return {
                "status": "created",
                "type": "pattern_rectangular_ex",
                "feature": feature_name,
                "x_count": x_count,
                "y_count": y_count,
                "x_spacing": x_spacing,
                "y_spacing": y_spacing,
                "plane_index": plane_index,
                "name": com_get(pattern, "Name", None),
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_pattern_circular_ex(
        self,
        feature_name: str,
        count: int,
        angle: float,
        axis_face_index: int,
    ) -> dict[str, Any]:
        """
        Create a circular pattern using the extended AddByCircularEx method.

        NOT AVAILABLE via COM automation. AddByCircularEx takes
        (NumberOfFeatures, FeatureArray, ReferencePlane, RadialCount,
        AngleSpacing, AxisPoint, PatternMethod, CurveDirection, ArcPattern,
        PatternType). AxisPoint is a SAFEARRAY of doubles - a point on the
        rotation axis - not the cylindrical face this tool is given, and there
        is no way to derive one from the other here. The call is never made.

        Args:
            feature_name: Name of the feature to pattern
            count: Number of instances around the circle
            angle: Total angle in degrees
            axis_face_index: 0-based index of the cylindrical face to use as axis

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "Circular feature patterns are not available through this "
                "server: Patterns.AddByCircularEx needs a reference plane and an "
                "AxisPoint coordinate array, which cannot be derived from a face "
                "index. Use create_pattern(method='rectangular_ex') or pattern "
                "the feature in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "pattern_circular_ex",
            "feature": feature_name,
            "count": count,
            "angle": angle,
            "axis_face_index": axis_face_index,
        }

    @verifies_geometry
    def create_pattern_duplicate(self, feature_name: str) -> dict[str, Any]:
        """
        Create a duplicate pattern of a feature.

        NOT AVAILABLE via COM automation. Patterns.AddDuplicate takes
        (NumberOfFeatures, FeatureArray, FromReference, NumberOfInstanceRefs,
        InstanceRefsArray, PatternType): FromReference and InstanceRefsArray are
        the placement references the copies land on, which this server cannot
        select. The call is never made.

        Args:
            feature_name: Name of the feature to duplicate

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "Duplicate patterns are not available through this server: "
                "Patterns.AddDuplicate needs a FromReference object and an array "
                "of instance references (a user selection) that cannot be built "
                "here. Use create_pattern(method='rectangular_ex') or duplicate "
                "the feature in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "pattern_duplicate",
            "feature": feature_name,
        }

    @verifies_geometry
    def create_pattern_by_fill(
        self,
        feature_name: str,
        fill_region_face_index: int,
        x_spacing: float,
        y_spacing: float,
    ) -> dict[str, Any]:
        """
        Create a fill pattern of a feature within a region.

        NOT AVAILABLE via COM automation. Patterns.AddByFill takes twelve
        arguments and its region is a RegionProfileArray - closed sketch region
        profiles - not the body face this tool is given. The call is never made.

        Args:
            feature_name: Name of the feature to pattern
            fill_region_face_index: 0-based face index defining the fill region
            x_spacing: Spacing in X direction (meters)
            y_spacing: Spacing in Y direction (meters)

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "Fill patterns are not available through this server: "
                "Patterns.AddByFill needs an array of region profiles (closed "
                "sketch regions), not a body face index. Use "
                "create_pattern(method='rectangular_ex') or fill-pattern the "
                "feature in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "pattern_by_fill",
            "feature": feature_name,
            "fill_region_face_index": fill_region_face_index,
            "x_spacing": x_spacing,
            "y_spacing": y_spacing,
        }

    @verifies_geometry
    def create_pattern_by_table(
        self,
        feature_name: str,
        x_offsets: list[float],
        y_offsets: list[float],
    ) -> dict[str, Any]:
        """
        Create a table-driven pattern of a feature.

        NOT AVAILABLE via COM automation. Patterns.AddPatternByTable takes
        fourteen arguments and is driven by an Excel workbook
        (InputExcelPath) plus KeyPoint objects for the from/to point options;
        it never accepts plain X/Y offset arrays. The call is never made.

        Args:
            feature_name: Name of the feature to pattern
            x_offsets: List of X offsets in meters
            y_offsets: List of Y offsets in meters (must match length of x_offsets)

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "Table-driven patterns are not available through this server: "
                "Patterns.AddPatternByTable is driven by an Excel file path and "
                "KeyPoint objects, not X/Y offset lists. Use "
                "create_pattern(method='rectangular_ex') or build the table "
                "pattern in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "pattern_by_table",
            "feature": feature_name,
            "x_offsets": list(x_offsets),
            "y_offsets": list(y_offsets),
        }

    @verifies_geometry
    def create_pattern_by_table_sync(
        self,
        feature_name: str,
        x_offsets: list[float],
        y_offsets: list[float],
    ) -> dict[str, Any]:
        """
        Create a synchronous table-driven pattern of a feature.

        NOT AVAILABLE via COM automation. Patterns.AddPatternByTableSync takes
        thirteen arguments and is driven by an Excel workbook (InputExcelPath)
        plus KeyPoint objects; it never accepts plain X/Y offset arrays. The
        call is never made.

        Args:
            feature_name: Name of the feature to pattern
            x_offsets: List of X offsets in meters
            y_offsets: List of Y offsets in meters

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "Synchronous table-driven patterns are not available through "
                "this server: Patterns.AddPatternByTableSync is driven by an "
                "Excel file path and KeyPoint objects, not X/Y offset lists. Use "
                "create_pattern(method='rectangular_ex') or build the table "
                "pattern in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "pattern_by_table_sync",
            "feature": feature_name,
            "x_offsets": list(x_offsets),
            "y_offsets": list(y_offsets),
        }

    @verifies_geometry
    def create_pattern_by_fill_ex(
        self,
        feature_name: str,
        fill_region_face_index: int,
        x_spacing: float,
        y_spacing: float,
        stagger_offset: float = 0.0,
    ) -> dict[str, Any]:
        """
        Create an extended fill pattern of a feature within a region.

        NOT AVAILABLE via COM automation. Patterns.AddByFillEx takes fifteen
        arguments and its region is a RegionProfileArray - closed sketch region
        profiles - not the body face this tool is given. The call is never made.

        Args:
            feature_name: Name of the feature to pattern
            fill_region_face_index: 0-based face index defining the fill region
            x_spacing: X direction spacing in meters
            y_spacing: Y direction spacing in meters
            stagger_offset: Stagger offset for pattern rows in meters

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "Fill patterns are not available through this server: "
                "Patterns.AddByFillEx needs an array of region profiles (closed "
                "sketch regions), not a body face index. Use "
                "create_pattern(method='rectangular_ex') or fill-pattern the "
                "feature in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "pattern_by_fill_ex",
            "feature": feature_name,
            "fill_region_face_index": fill_region_face_index,
            "x_spacing": x_spacing,
            "y_spacing": y_spacing,
            "stagger_offset": stagger_offset,
        }

    @verifies_geometry
    def create_pattern_by_curve_ex(
        self,
        feature_name: str,
        curve_edge_index: int,
        count: int,
        spacing: float,
    ) -> dict[str, Any]:
        """
        Create a pattern along a curve using the extended API.

        NOT AVAILABLE via COM automation. Patterns.AddByCurveEx takes 23
        required arguments, among them AnchorPointForCurves1 - a KeyPoint on the
        curve - plus a second curve set and a transform plane/surface. None of
        those can be selected here. The call is never made.

        Args:
            feature_name: Name of the feature to pattern
            curve_edge_index: 0-based edge index defining the curve path
            count: Number of pattern occurrences
            spacing: Spacing between occurrences in meters

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "Curve patterns are not available through this server: "
                "Patterns.AddByCurveEx requires 23 arguments including an anchor "
                "KeyPoint on the curve and a transform plane or surface, which "
                "this server cannot select. Use "
                "create_pattern(method='rectangular_ex') or pattern along the "
                "curve in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "pattern_by_curve_ex",
            "feature": feature_name,
            "curve_edge_index": curve_edge_index,
            "count": count,
            "spacing": spacing,
        }

    def save_as_mirror_part(
        self,
        new_file_name: str,
        mirror_plane_index: int = 3,
        link_to_original: bool = True,
        overwrite: bool = False,
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

            err = guard_overwrite(new_file_name, overwrite)
            if err:
                return err

            models.SaveAsMirrorPart(new_file_name, mirror_plane, link_to_original)

            return {
                "status": "saved",
                "type": "mirror_part",
                "path": new_file_name,
                "mirror_plane_index": mirror_plane_index,
                "linked": link_to_original,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_user_defined_pattern(self, feature_name: str) -> dict[str, Any]:
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

            profiles_var = all_profiles

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
            return error_result(e)
