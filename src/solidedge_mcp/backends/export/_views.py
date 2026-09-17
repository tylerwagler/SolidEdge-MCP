"""Drawing view management operations (scale, delete, update, project, move, display, etc.)."""

import contextlib
from typing import Any

from solidedge_mcp.backends.errors import error_result

from ..constants import FoldTypeConstants
from ..logging import get_logger
from ._base import com_get
from ._drawing import VIEW_ORIENTATIONS

_logger = get_logger(__name__)


#: How each display mode maps onto the properties a DrawingView actually has.
#: DrawingView.SetRenderMode and DisplayMode do not exist; SetRenderMode is on
#: the 3D window View in framewrk.tlb.
_DISPLAY_MODES: dict[str, dict[str, Any]] = {
    "Wireframe": {
        "Shading": False,
        "Defaults_ShowHiddenEdges": True,
    },
    "HiddenEdgesVisible": {
        "Shading": False,
        "Defaults_ShowHiddenEdges": False,
    },
    "Shaded": {
        "Shading": True,
        "ShadingShowVisibleEdges": False,
    },
    "ShadedWithEdges": {
        "Shading": True,
        "ShadingShowVisibleEdges": True,
    },
}


def _origin_of(view: Any) -> list[float] | None:
    """The view's origin on the sheet. GetOrigin is pure [out]."""
    try:
        origin = view.GetOrigin()
        return [float(origin[0]), float(origin[1])]
    except Exception:
        return None


class ViewsMixin:
    """Mixin providing drawing view manipulation methods."""

    def set_drawing_view_scale(self, view_index: int, scale: float) -> dict[str, Any]:
        """
        Set the scale of a drawing view.

        Args:
            view_index: 0-based view index
            scale: New scale factor (e.g. 1.0, 0.5, 2.0)

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()
            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)
            view.ScaleFactor = scale

            # Report what the view holds, not what was asked: a write Solid
            # Edge quietly clamps or ignores would otherwise be reported as
            # applied. set_drawing_view_orientation reads back the same way.
            return {
                "status": "set",
                "view_index": view_index,
                "scale": scale,
                "reads_back": view.ScaleFactor,
            }
        except Exception as e:
            return error_result(e)

    def delete_drawing_view(self, view_index: int) -> dict[str, Any]:
        """
        Delete a drawing view from the active sheet.

        Args:
            view_index: 0-based view index

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()
            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)
            view.Delete()

            return {"status": "deleted", "view_index": view_index, "remaining_views": dvs.Count}
        except Exception as e:
            return error_result(e)

    def update_drawing_view(self, view_index: int) -> dict[str, Any]:
        """
        Force update a drawing view to reflect 3D model changes.

        Args:
            view_index: 0-based view index

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()
            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)
            view.Update()

            return {"status": "updated", "view_index": view_index}
        except Exception as e:
            return error_result(e)

    def add_projected_view(
        self, parent_view_index: int, fold_direction: str, x: float, y: float
    ) -> dict[str, Any]:
        """
        Add a projected (folded) drawing view from a parent view.

        Creates an orthographic projection by folding from the parent view
        in the specified direction.

        Args:
            parent_view_index: 0-based index of the parent drawing view
            fold_direction: 'Up', 'Down', 'Left', or 'Right'
            x: X position on sheet (meters)
            y: Y position on sheet (meters)

        Returns:
            Dict with status and view info
        """
        try:
            dvs = self._get_drawing_views()

            if parent_view_index < 0 or parent_view_index >= dvs.Count:
                return {
                    "error": f"Invalid parent view index: {parent_view_index}. Count: {dvs.Count}"
                }

            fold_map = {
                "Up": FoldTypeConstants.igFoldUp,
                "Down": FoldTypeConstants.igFoldDown,
                "Left": FoldTypeConstants.igFoldLeft,
                "Right": FoldTypeConstants.igFoldRight,
            }

            fold_const = fold_map.get(fold_direction)
            if fold_const is None:
                valid = ", ".join(fold_map.keys())
                return {"error": f"Invalid fold_direction: '{fold_direction}'. Valid: {valid}"}

            parent_view = dvs.Item(parent_view_index + 1)
            dvs.AddByFold(parent_view, fold_const, x, y)

            return {
                "status": "added",
                "parent_view_index": parent_view_index,
                "fold_direction": fold_direction,
                "position": [x, y],
                "total_views": dvs.Count,
            }
        except Exception as e:
            return error_result(e)

    def _set_edge_display(self, view_index: int, prop: str, show: bool, key: str) -> dict[str, Any]:
        """Set one edge-display property on a view and confirm it took."""
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)
            setattr(view, prop, show)
            with contextlib.suppress(Exception):
                view.Update()

            result: dict[str, Any] = {"status": "updated", "view_index": view_index, key: show}
            actual = com_get(view, prop)
            if actual is not None and bool(actual) != show:
                # Say so rather than report a success the view did not accept.
                return {
                    "error": (
                        f"This drawing view did not accept {key}={show}; it still "
                        f"reports {bool(actual)}. Some view types fix their edge "
                        f"display."
                    ),
                    "view_index": view_index,
                }
            return result
        except Exception as e:
            return error_result(e)

    def move_drawing_view(self, view_index: int, x: float, y: float) -> dict[str, Any]:
        """Move a drawing view to a new position on the sheet.

        ``DrawingView.SetOrigin(x, y)`` is the only route. There is no
        ``OriginX``/``OriginY`` property to assign, and no ``XPosition``, which
        is what the fallback tried; both raise "can not be set", so this always
        failed. The new origin is read back with ``GetOrigin``.

        Args:
            view_index: 0-based index of the view to move.
            x: New origin X on the sheet, in meters.
            y: New origin Y on the sheet, in meters.

        Returns:
            Dict with status and the origin Solid Edge settled on.
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)
            view.SetOrigin(x, y)

            result: dict[str, Any] = {
                "status": "moved",
                "view_index": view_index,
                "position": [x, y],
            }
            origin = _origin_of(view)
            if origin is not None:
                result["origin"] = origin
            return result
        except Exception as e:
            return error_result(e)

    def show_hidden_edges(self, view_index: int, show: bool = True) -> dict[str, Any]:
        """Toggle hidden edge display on a drawing view.

        ``DrawingView.ShowHiddenEdges`` does not exist and raises "can not be
        set". ``ModelMember.ShowHiddenEdges`` does exist, accepts the write,
        and then reads back unchanged, so routing it there would look like
        success and do nothing. ``DrawingView.Defaults_ShowHiddenEdges`` is the
        one that takes and survives an Update. Verified on Solid Edge 2026.

        Args:
            view_index: 0-based index of the view.
            show: True to draw hidden edges, False to leave them out.

        Returns:
            Dict with status and the value the view reports afterwards.
        """
        return self._set_edge_display(
            view_index, "Defaults_ShowHiddenEdges", show, "show_hidden_edges"
        )

    def set_drawing_view_display_mode(self, view_index: int, mode: str) -> dict[str, Any]:
        """Set how a drawing view is rendered.

        ``SetRenderMode`` belongs to the 3D window ``View`` in framewrk.tlb,
        not to a 2D ``DrawingView``, and ``DrawingView.DisplayMode`` does not
        exist, so both the call and its fallback always raised. A drawing view
        is controlled by ``Shading``, ``ShadingShowVisibleEdges`` and
        ``Defaults_ShowHiddenEdges`` instead. Verified on Solid Edge 2026.

        Args:
            view_index: 0-based index of the view.
            mode: 'Wireframe' shades nothing and draws every edge.
                'HiddenEdgesVisible' shades nothing and hides obscured edges.
                'Shaded' shades without edge lines.
                'ShadedWithEdges' shades and keeps visible edges.

        Returns:
            Dict with status and the properties that were set.
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            settings = _DISPLAY_MODES.get(mode)
            if settings is None:
                valid = ", ".join(_DISPLAY_MODES)
                return {"error": f"Invalid mode: '{mode}'. Valid: {valid}"}

            view = dvs.Item(view_index + 1)
            applied: dict[str, Any] = {}
            refused: list[str] = []
            for prop, value in settings.items():
                try:
                    setattr(view, prop, value)
                    applied[prop] = value
                except Exception:
                    refused.append(prop)

            if not applied:
                return {
                    "error": (
                        f"This drawing view accepted none of the properties for "
                        f"mode '{mode}': {', '.join(refused)}."
                    ),
                    "view_index": view_index,
                }

            with contextlib.suppress(Exception):
                view.Update()

            result: dict[str, Any] = {
                "status": "updated",
                "view_index": view_index,
                "mode": mode,
                "applied": applied,
            }
            if refused:
                result["not_supported_by_this_view"] = refused
            return result
        except Exception as e:
            return error_result(e)

    def get_drawing_view_info(self, view_index: int) -> dict[str, Any]:
        """
        Get detailed information about a drawing view.

        Args:
            view_index: 0-based view index

        Returns:
            Dict with scale, position, display properties, and name
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)

            info: dict[str, Any] = {"view_index": view_index}

            with contextlib.suppress(Exception):
                info["name"] = view.Name
            with contextlib.suppress(Exception):
                info["scale"] = view.ScaleFactor
            # GetOrigin is pure [out]; there is no OriginX/OriginY to read,
            # so these two keys were always missing.
            origin = _origin_of(view)
            if origin is not None:
                info["origin_x"] = origin[0]
                info["origin_y"] = origin[1]
            with contextlib.suppress(Exception):
                info["show_hidden_edges"] = view.Defaults_ShowHiddenEdges
            with contextlib.suppress(Exception):
                info["show_tangent_edges"] = view.Defaults_ShowTangentEdges
            with contextlib.suppress(Exception):
                info["shaded"] = view.Shading
            with contextlib.suppress(Exception):
                info["type"] = view.Type

            return info
        except Exception as e:
            return error_result(e)

    def set_drawing_view_orientation(self, view_index: int, orientation: str) -> dict[str, Any]:
        """
        Change the orientation of a drawing view.

        Args:
            view_index: 0-based view index
            orientation: 'Front', 'Top', 'Right', 'Back', 'Bottom', 'Left', 'Isometric'

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            orient_const = VIEW_ORIENTATIONS.get(orientation)
            if orient_const is None:
                valid = ", ".join(VIEW_ORIENTATIONS)
                return {"error": f"Invalid orientation: '{orientation}'. Valid: {valid}"}

            view = dvs.Item(view_index + 1)
            # ViewOrientation is a method with seven out-parameters, not a
            # settable property: assigning to it raised "Property
            # 'Item.ViewOrientation' can not be set" on every call.
            # SetViewOrientationStandard is the setter.
            view.SetViewOrientationStandard(orient_const)

            return {
                "status": "updated",
                "view_index": view_index,
                "orientation": orientation,
                "reads_back": view.ViewOrientation()[6],
            }
        except Exception as e:
            return error_result(e)

    def get_drawing_view_model_link(self, view_index: int) -> dict[str, Any]:
        """
        Get the model link reference from a drawing view.

        Returns information about which 3D model is associated with
        the specified drawing view.

        Args:
            view_index: 0-based view index

        Returns:
            Dict with model link info (file path, name)
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)

            info = {"view_index": view_index}

            with contextlib.suppress(Exception):
                model_link = view.ModelLink
                info["has_model_link"] = True
                with contextlib.suppress(Exception):
                    info["model_path"] = model_link.FileName
                with contextlib.suppress(Exception):
                    info["model_name"] = model_link.Name
            if "has_model_link" not in info:
                info["has_model_link"] = False

            with contextlib.suppress(Exception):
                info["view_name"] = view.Name
            with contextlib.suppress(Exception):
                info["scale"] = view.ScaleFactor

            return info
        except Exception as e:
            return error_result(e)

    def show_tangent_edges(self, view_index: int, show: bool = True) -> dict[str, Any]:
        """Toggle tangent edge display on a drawing view.

        Tangent edges are where two surfaces meet smoothly, such as where a
        fillet runs into a flat face.

        ``DrawingView.ShowTangentEdges`` does not exist.
        ``ModelMember.ShowTangentEdges`` does, but accepts the write and reads
        back unchanged, so it is a silent no-op.
        ``DrawingView.Defaults_ShowTangentEdges`` is the one that holds.
        Verified on Solid Edge 2026.

        Args:
            view_index: 0-based index of the view.
            show: True to draw tangent edges, False to leave them out.

        Returns:
            Dict with status and the value the view reports afterwards.
        """
        return self._set_edge_display(
            view_index, "Defaults_ShowTangentEdges", show, "show_tangent_edges"
        )

    # =================================================================
    # DRAWING VIEW VARIANTS
    # =================================================================

    def add_detail_view(
        self,
        parent_view_index: int,
        center_x: float,
        center_y: float,
        radius: float,
        x: float,
        y: float,
        scale: float = 2.0,
    ) -> dict[str, Any]:
        """
        Add a detail (zoom) view from a parent drawing view.

        Creates a circular detail envelope on the parent view and places
        an enlarged view at the specified position.

        Args:
            parent_view_index: 0-based index of the parent drawing view
            center_x: Detail envelope center X on parent view (meters)
            center_y: Detail envelope center Y on parent view (meters)
            radius: Detail envelope radius (meters)
            x: Detail view X position on sheet (meters)
            y: Detail view Y position on sheet (meters)
            scale: Detail view scale factor (default 2.0)

        Returns:
            Dict with status and view info
        """
        try:
            dvs = self._get_drawing_views()

            if parent_view_index < 0 or parent_view_index >= dvs.Count:
                return {
                    "error": f"Invalid parent view index: {parent_view_index}. Count: {dvs.Count}"
                }

            parent_view = dvs.Item(parent_view_index + 1)

            try:
                # AddByDetailEnvelope(From, x1, y1, Radius, Scale, x2, y2) --
                # Scale comes before the placement point, not after it.
                dvs.AddByDetailEnvelope(parent_view, center_x, center_y, radius, scale, x, y)
            except Exception:
                # Fallback: AddDetailView(From, x1, y1, Radius, Scale, x2, y2,
                # Independent)
                dvs.AddDetailView(parent_view, center_x, center_y, radius, scale, x, y, False)

            return {
                "status": "added",
                "type": "detail_view",
                "parent_view_index": parent_view_index,
                "center": [center_x, center_y],
                "radius": radius,
                "position": [x, y],
                "scale": scale,
                "total_views": dvs.Count,
            }
        except Exception as e:
            return error_result(e)

    def add_auxiliary_view(
        self,
        parent_view_index: int,
        x: float,
        y: float,
        fold_direction: str = "Up",
    ) -> dict[str, Any]:
        """
        Add an auxiliary (folded) view from a parent drawing view.

        An auxiliary view shows the model from an angle not available
        from standard orthographic projections.

        Args:
            parent_view_index: 0-based index of the parent drawing view
            x: Auxiliary view X position on sheet (meters)
            y: Auxiliary view Y position on sheet (meters)
            fold_direction: Fold direction - 'Up', 'Down', 'Left', 'Right'

        Returns:
            Dict with status and view info
        """
        try:
            dvs = self._get_drawing_views()

            if parent_view_index < 0 or parent_view_index >= dvs.Count:
                return {
                    "error": f"Invalid parent view index: {parent_view_index}. Count: {dvs.Count}"
                }

            fold_map = {
                "Up": FoldTypeConstants.igFoldUp,
                "Down": FoldTypeConstants.igFoldDown,
                "Left": FoldTypeConstants.igFoldLeft,
                "Right": FoldTypeConstants.igFoldRight,
            }

            fold_const = fold_map.get(fold_direction)
            if fold_const is None:
                valid = ", ".join(fold_map.keys())
                return {"error": f"Invalid fold_direction: '{fold_direction}'. Valid: {valid}"}

            parent_view = dvs.Item(parent_view_index + 1)

            # AddByFold(From, foldDir, x, y) is the fold-direction entry point.
            # AddByAuxiliaryFold(From, x1, y1, x2, y2, x3, y3) is a different
            # method that wants a fold line picked on the parent view, so it was
            # never callable with these arguments.
            dvs.AddByFold(parent_view, fold_const, x, y)

            return {
                "status": "added",
                "type": "auxiliary_view",
                "parent_view_index": parent_view_index,
                "fold_direction": fold_direction,
                "position": [x, y],
                "total_views": dvs.Count,
            }
        except Exception as e:
            return error_result(e)

    def add_draft_view(self, x: float, y: float, scale: float = 1.0) -> dict[str, Any]:
        """
        Add an empty draft (sketch) view to the active sheet.

        A draft view is an empty drawing view area where you can add
        free-form sketch geometry and annotations.

        Args:
            x: View X position on sheet (meters)
            y: View Y position on sheet (meters)
            scale: View scale factor (default 1.0)

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()

            # AddDraftView(Scale, x1, y1) - the scale comes first.
            dvs.AddDraftView(scale, x, y)

            return {
                "status": "added",
                "type": "draft_view",
                "position": [x, y],
                "total_views": dvs.Count,
            }
        except Exception as e:
            return error_result(e)

    # =================================================================
    # DRAWING VIEW PROPERTIES
    # =================================================================

    def align_drawing_views(
        self, view_index1: int, view_index2: int, align: bool = True
    ) -> dict[str, Any]:
        """Align or unalign two drawing views.

        ``DrawingView.AlignToView`` and ``RemoveAlignment`` are in no Solid
        Edge type library, so both the call and its fallback always raised.
        The real API is on the collection: ``DrawingViews.Align()`` and
        ``Unalign()``, which act on whatever is in the document select set.
        ``Align()`` raises outright when nothing is selected. Verified on
        Solid Edge 2026.

        Solid Edge can only align views that already have a fold relationship,
        so this reports each view's origin before and after; equal origins mean
        Solid Edge found nothing to align.

        Args:
            view_index1: 0-based index of the first view.
            view_index2: 0-based index of the second view.
            align: True to align the pair, False to release the alignment.

        Returns:
            Dict with status and both origins before and after.
        """
        try:
            dvs = self._get_drawing_views()

            if view_index1 < 0 or view_index1 >= dvs.Count:
                return {"error": f"Invalid view_index1: {view_index1}. Count: {dvs.Count}"}
            if view_index2 < 0 or view_index2 >= dvs.Count:
                return {"error": f"Invalid view_index2: {view_index2}. Count: {dvs.Count}"}
            if view_index1 == view_index2:
                return {
                    "error": (
                        f"view_index1 and view_index2 are both {view_index1}. "
                        f"Alignment needs two different views."
                    )
                }

            view1 = dvs.Item(view_index1 + 1)
            view2 = dvs.Item(view_index2 + 1)
            before = [_origin_of(view1), _origin_of(view2)]

            doc = self.doc_manager.get_active_document()
            select_set = com_get(doc, "SelectSet")
            if select_set is None:
                return {
                    "error": (
                        "This document has no SelectSet, so the views cannot be "
                        "handed to DrawingViews.Align."
                    )
                }
            with contextlib.suppress(Exception):
                select_set.RemoveAll()
            select_set.Add(view1)
            select_set.Add(view2)

            try:
                if align:
                    dvs.Align()
                else:
                    dvs.Unalign()
            finally:
                with contextlib.suppress(Exception):
                    select_set.RemoveAll()

            after = [_origin_of(view1), _origin_of(view2)]
            result: dict[str, Any] = {
                "status": "aligned" if align else "unaligned",
                "view_index1": view_index1,
                "view_index2": view_index2,
                "origins_before": before,
                "origins_after": after,
            }
            if align and before == after:
                result["note"] = (
                    "Neither view moved. Solid Edge aligns views that share a fold "
                    "relationship; unrelated views have nothing to align to. Create "
                    "one with add_drawing_view(type='projected', parent_view_index=...)."
                )
            return result
        except Exception as e:
            return error_result(e)

    def add_by_draft_view(
        self, source_view_index: int, x: float, y: float, scale: float | None = None
    ) -> dict[str, Any]:
        """
        Copy an existing drawing view to a new location on the active sheet.

        Uses DrawingViews.AddByDraftView(From, Scale, x1, y1) from the draft
        type library. The source view is specified by 0-based index.

        Args:
            source_view_index: 0-based index of the source drawing view
            x: X position for the new view on the sheet (meters)
            y: Y position for the new view on the sheet (meters)
            scale: Scale for the new view (default: same as source view)

        Returns:
            Dict with status and new view info
        """
        try:
            dvs = self._get_drawing_views()

            if source_view_index < 0 or source_view_index >= dvs.Count:
                return {
                    "error": f"Invalid source view index: {source_view_index}. Count: {dvs.Count}"
                }

            source_view = dvs.Item(source_view_index + 1)

            # Use source view scale if not specified
            if scale is None:
                try:
                    scale = source_view.ScaleFactor
                except Exception:
                    scale = 1.0

            # AddByDraftView(From: DrawingView*, Scale: VT_R8, x1: VT_R8, y1: VT_R8)
            new_view = dvs.AddByDraftView(source_view, scale, x, y)

            result = {
                "status": "added",
                "source_view_index": source_view_index,
                "position": [x, y],
                "scale": scale,
                "total_views": dvs.Count,
            }

            with contextlib.suppress(Exception):
                result["name"] = new_view.Name

            return result
        except Exception as e:
            return error_result(e)

    def activate_drawing_view(self, view_index: int) -> dict[str, Any]:
        """
        Activate a drawing view by 0-based index.

        An activated view allows editing its contents.

        Args:
            view_index: 0-based view index

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()
            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view_index: {view_index}. Count: {dvs.Count}"}
            view = dvs.Item(view_index + 1)
            view.Activate()
            return {"status": "activated", "view_index": view_index}
        except Exception as e:
            return error_result(e)

    def deactivate_drawing_view(self, view_index: int) -> dict[str, Any]:
        """
        Deactivate a drawing view by 0-based index.

        Deactivating a view returns focus to the sheet.

        Args:
            view_index: 0-based view index

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()
            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view_index: {view_index}. Count: {dvs.Count}"}
            view = dvs.Item(view_index + 1)
            view.Deactivate()
            return {"status": "deactivated", "view_index": view_index}
        except Exception as e:
            return error_result(e)

    # =================================================================
    # SECTION CUTS / DIMENSIONS ON VIEWS
    # =================================================================

    def get_section_cuts(self, view_index: int) -> dict[str, Any]:
        """
        Get section cut (cutting plane) information from a drawing view.

        Accesses DrawingView.CuttingPlanes collection and extracts caption,
        display type, and fold line geometry for each cutting plane.

        Args:
            view_index: 0-based index of the drawing view

        Returns:
            Dict with count and list of section cut info dicts
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)

            cutting_planes = com_get(view, "CuttingPlanes")
            if cutting_planes is None:
                return {"count": 0, "section_cuts": [], "note": "No CuttingPlanes on this view"}
            items = []
            for i in range(1, cutting_planes.Count + 1):
                cp = cutting_planes.Item(i)
                info: dict[str, Any] = {"index": i - 1}

                with contextlib.suppress(Exception):
                    info["caption"] = cp.Caption
                with contextlib.suppress(Exception):
                    info["display_caption"] = cp.DisplayCaption
                with contextlib.suppress(Exception):
                    info["display_type"] = cp.DisplayType
                with contextlib.suppress(Exception):
                    info["style_name"] = cp.StyleName
                with contextlib.suppress(Exception):
                    info["text_height"] = cp.TextHeight

                # Try to get fold line geometry
                with contextlib.suppress(Exception):
                    (
                        line_start_x,
                        line_start_y,
                        line_end_x,
                        line_end_y,
                        view_dir_x,
                        view_dir_y,
                    ) = cp.GetFoldLineWithViewDirection()
                    info["fold_line"] = {
                        "start": [line_start_x, line_start_y],
                        "end": [line_end_x, line_end_y],
                        "view_direction": [view_dir_x, view_dir_y],
                    }

                items.append(info)

            return {"count": len(items), "section_cuts": items, "view_index": view_index}
        except Exception as e:
            return error_result(e)

    def add_section_cut(
        self,
        view_index: int,
        x: float,
        y: float,
        section_type: int = 0,
    ) -> dict[str, Any]:
        """
        Add a section cut (cutting plane) to a drawing view and create the section view.

        Creates a cutting plane on the specified drawing view via
        CuttingPlanes.Add(), then calls CuttingPlane.CreateView(SectionType)
        to generate the section drawing view.

        Args:
            view_index: 0-based index of the source drawing view
            x: X position for the section view on the sheet (meters)
            y: Y position for the section view on the sheet (meters)
            section_type: 0 = standard, 1 = revolved (DraftSectionViewType)

        Returns:
            Dict with status and section view info
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            if section_type not in (0, 1):
                return {
                    "error": f"Invalid section_type: {section_type}. "
                    "Use 0 (standard) or 1 (revolved)."
                }

            view = dvs.Item(view_index + 1)

            cutting_planes = com_get(view, "CuttingPlanes")
            if cutting_planes is None:
                return {"error": "Drawing view does not support CuttingPlanes"}

            # CuttingPlanes.Add() returns a new CuttingPlane object
            cutting_plane = cutting_planes.Add()

            # CuttingPlane.CreateView(SectionType) creates the section view
            section_view = cutting_plane.CreateView(section_type)

            # Move the section view to the desired position. A DrawingView
            # has no OriginX/OriginY to assign, so this quietly left every
            # section view wherever Solid Edge first put it.
            placed = False
            with contextlib.suppress(Exception):
                section_view.SetOrigin(x, y)
                placed = True

            result = {
                "status": "added",
                "source_view_index": view_index,
                "section_type": "standard" if section_type == 0 else "revolved",
                "position": [x, y] if placed else None,
                "total_cutting_planes": cutting_planes.Count,
            }

            with contextlib.suppress(Exception):
                result["caption"] = cutting_plane.Caption
            with contextlib.suppress(Exception):
                result["section_view_name"] = section_view.Name

            return result
        except Exception as e:
            return error_result(e)

    def get_drawing_view_dimensions(self, view_index: int) -> dict[str, Any]:
        """
        Get all dimensions associated with a specific drawing view.

        Accesses DrawingView.Dimensions collection (dispid 120) and iterates
        each Dimension to return its type, value, prefix/suffix strings, and
        override text.

        DimensionType constants (DimTypeConstants):
            1=Linear, 2=Radial, 3=Angular, 4=RadialDiameter,
            5=CircularDiameter, 6=ArcLength, 7=ArcAngle, 8=Coordinate,
            9=SymmetricalDiameter, 10=Chamfer, 11=AngularCoordinate,
            12=CurveLength

        Args:
            view_index: 0-based index of the drawing view

        Returns:
            Dict with count and list of dimension info dicts
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)

            dims = com_get(view, "Dimensions")
            if dims is None:
                return {"count": 0, "dimensions": [], "note": "No Dimensions on this view"}

            dim_type_names = {
                1: "Linear",
                2: "Radial",
                3: "Angular",
                4: "RadialDiameter",
                5: "CircularDiameter",
                6: "ArcLength",
                7: "ArcAngle",
                8: "Coordinate",
                9: "SymmetricalDiameter",
                10: "Chamfer",
                11: "AngularCoordinate",
                12: "CurveLength",
            }

            items = []
            for i in range(1, dims.Count + 1):
                dim = dims.Item(i)
                info: dict[str, Any] = {"index": i - 1}

                with contextlib.suppress(Exception):
                    raw_type = dim.DimensionType
                    info["dimension_type"] = raw_type
                    info["dimension_type_name"] = dim_type_names.get(raw_type, "Unknown")
                with contextlib.suppress(Exception):
                    info["value"] = dim.Value
                with contextlib.suppress(Exception):
                    info["constraint"] = dim.Constraint
                with contextlib.suppress(Exception):
                    info["prefix"] = dim.PrefixString
                with contextlib.suppress(Exception):
                    info["suffix"] = dim.SuffixString
                with contextlib.suppress(Exception):
                    info["override"] = dim.OverrideString
                with contextlib.suppress(Exception):
                    info["subfix"] = dim.SubfixString
                with contextlib.suppress(Exception):
                    info["superfix"] = dim.SuperfixString

                items.append(info)

            return {
                "count": len(items),
                "dimensions": items,
                "view_index": view_index,
            }
        except Exception as e:
            return error_result(e)

    def update_all_views(self, force_update: bool = True) -> dict[str, Any]:
        """
        Update all drawing views on all sheets in the active draft document.

        Uses DraftDocument.UpdateAll(ForceUpdate).

        Args:
            force_update: If True, forces update even if views appear current

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_draft(doc)
            if err:
                return err

            doc.UpdateAll(force_update)
            return {"status": "updated_all", "force_update": force_update}
        except Exception as e:
            return error_result(e)
