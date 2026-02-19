"""Selection set operations."""

import contextlib
import traceback
from typing import Any

from ..constants import FaceQueryConstants
from ..logging import get_logger

_logger = get_logger(__name__)


class SelectionMixin:
    """Mixin providing selection set methods."""

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
