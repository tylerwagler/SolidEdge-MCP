"""Document properties, feature counts, reference planes, modeling mode, and recompute."""

import contextlib
import traceback
from typing import Any

from ..constants import ModelingModeConstants


class DocumentQueryMixin:
    """Mixin providing document-level query and inspection methods."""

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

    def recompute_document(self) -> dict[str, Any]:
        """
        Force a full document-level recompute.

        Unlike recompute() which tries model-level first, this calls
        PartDocument.Recompute() directly for a complete document rebuild.

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            doc.Recompute()
            return {"status": "recomputed_document"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
