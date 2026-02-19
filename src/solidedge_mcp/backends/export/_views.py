"""Drawing view management operations (scale, delete, update, project, move, display, etc.)."""

import contextlib
import traceback
from typing import Any

from ..constants import DrawingViewOrientationConstants, FoldTypeConstants, RenderModeConstants
from ..logging import get_logger

_logger = get_logger(__name__)


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

            return {"status": "set", "view_index": view_index, "scale": scale}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

    def move_drawing_view(self, view_index: int, x: float, y: float) -> dict[str, Any]:
        """
        Reposition a drawing view on the sheet.

        Args:
            view_index: 0-based view index
            x: New X position (meters)
            y: New Y position (meters)

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)

            try:
                view.OriginX = x
                view.OriginY = y
            except Exception:
                view.XPosition = x
                view.YPosition = y

            return {"status": "moved", "view_index": view_index, "position": [x, y]}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def show_hidden_edges(self, view_index: int, show: bool = True) -> dict[str, Any]:
        """
        Toggle hidden edge visibility on a drawing view.

        Args:
            view_index: 0-based view index
            show: True to show hidden edges, False to hide them

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)
            view.ShowHiddenEdges = show

            return {"status": "updated", "view_index": view_index, "show_hidden_edges": show}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_drawing_view_display_mode(self, view_index: int, mode: str) -> dict[str, Any]:
        """
        Set the display/render mode of a drawing view.

        Args:
            view_index: 0-based view index
            mode: 'Wireframe', 'HiddenEdgesVisible', 'Shaded', or 'ShadedWithEdges'

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            mode_map = {
                "Wireframe": RenderModeConstants.seRenderModeWireframe,
                "HiddenEdgesVisible": RenderModeConstants.seRenderModeVHL,
                "Shaded": RenderModeConstants.seRenderModeSmooth,
                "ShadedWithEdges": RenderModeConstants.seRenderModeSmoothBoundary,
            }

            mode_value = mode_map.get(mode)
            if mode_value is None:
                valid = ", ".join(mode_map.keys())
                return {"error": f"Invalid mode: '{mode}'. Valid: {valid}"}

            view = dvs.Item(view_index + 1)

            try:
                view.SetRenderMode(mode_value)
            except Exception:
                view.DisplayMode = mode_value

            return {"status": "updated", "view_index": view_index, "mode": mode}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            info = {"view_index": view_index}

            with contextlib.suppress(Exception):
                info["name"] = view.Name
            with contextlib.suppress(Exception):
                info["scale"] = view.ScaleFactor
            with contextlib.suppress(Exception):
                info["origin_x"] = view.OriginX
            with contextlib.suppress(Exception):
                info["origin_y"] = view.OriginY
            with contextlib.suppress(Exception):
                info["show_hidden_edges"] = view.ShowHiddenEdges
            with contextlib.suppress(Exception):
                info["show_tangent_edges"] = view.ShowTangentEdges
            with contextlib.suppress(Exception):
                info["type"] = view.Type

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            orient_map = {
                "Front": DrawingViewOrientationConstants.Front,
                "Top": DrawingViewOrientationConstants.Top,
                "Right": DrawingViewOrientationConstants.Right,
                "Back": DrawingViewOrientationConstants.Back,
                "Bottom": DrawingViewOrientationConstants.Bottom,
                "Left": DrawingViewOrientationConstants.Left,
                "Isometric": DrawingViewOrientationConstants.Isometric,
                "Iso": DrawingViewOrientationConstants.Isometric,
            }

            orient_const = orient_map.get(orientation)
            if orient_const is None:
                valid = ", ".join(orient_map.keys())
                return {"error": f"Invalid orientation: '{orientation}'. Valid: {valid}"}

            view = dvs.Item(view_index + 1)
            view.ViewOrientation = orient_const

            return {"status": "updated", "view_index": view_index, "orientation": orientation}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

    def show_tangent_edges(self, view_index: int, show: bool = True) -> dict[str, Any]:
        """
        Set tangent edge visibility on a drawing view.

        Tangent edges are edges where two surfaces meet tangentially
        (e.g., where a fillet meets a flat face).

        Args:
            view_index: 0-based view index
            show: True to show tangent edges, False to hide them

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)
            view.ShowTangentEdges = show

            return {
                "status": "updated",
                "view_index": view_index,
                "show_tangent_edges": show,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
                dvs.AddByDetailEnvelope(parent_view, center_x, center_y, radius, x, y, scale)
            except Exception:
                # Fallback: try AddDetailView
                dvs.AddDetailView(parent_view, center_x, center_y, radius, x, y, scale)

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            try:
                dvs.AddByAuxiliaryFold(parent_view, fold_const, x, y)
            except Exception:
                # Fallback: try AddByFold
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
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_draft_view(self, x: float, y: float) -> dict[str, Any]:
        """
        Add an empty draft (sketch) view to the active sheet.

        A draft view is an empty drawing view area where you can add
        free-form sketch geometry and annotations.

        Args:
            x: View X position on sheet (meters)
            y: View Y position on sheet (meters)

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()

            dvs.AddDraftView(x, y)

            return {
                "status": "added",
                "type": "draft_view",
                "position": [x, y],
                "total_views": dvs.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # DRAWING VIEW PROPERTIES
    # =================================================================

    def align_drawing_views(
        self, view_index1: int, view_index2: int, align: bool = True
    ) -> dict[str, Any]:
        """
        Align or unalign two drawing views.

        When aligned, moving one view constrains the other to maintain
        alignment (horizontal or vertical).

        Args:
            view_index1: 0-based index of the first view
            view_index2: 0-based index of the second view
            align: True to align, False to unalign

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()

            if view_index1 < 0 or view_index1 >= dvs.Count:
                return {"error": f"Invalid view_index1: {view_index1}. Count: {dvs.Count}"}
            if view_index2 < 0 or view_index2 >= dvs.Count:
                return {"error": f"Invalid view_index2: {view_index2}. Count: {dvs.Count}"}

            view1 = dvs.Item(view_index1 + 1)
            view2 = dvs.Item(view_index2 + 1)

            if align:
                try:
                    view1.AlignToView(view2)
                except Exception:
                    view2.AlignToView(view1)
            else:
                try:
                    view1.RemoveAlignment()
                except Exception:
                    view2.RemoveAlignment()

            return {
                "status": "aligned" if align else "unaligned",
                "view_index1": view_index1,
                "view_index2": view_index2,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            if not hasattr(view, "CuttingPlanes"):
                return {"count": 0, "section_cuts": [], "note": "No CuttingPlanes on this view"}

            cutting_planes = view.CuttingPlanes
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            if not hasattr(view, "CuttingPlanes"):
                return {"error": "Drawing view does not support CuttingPlanes"}

            cutting_planes = view.CuttingPlanes

            # CuttingPlanes.Add() returns a new CuttingPlane object
            cutting_plane = cutting_planes.Add()

            # CuttingPlane.CreateView(SectionType) creates the section view
            section_view = cutting_plane.CreateView(section_type)

            # Move the section view to the desired position
            with contextlib.suppress(Exception):
                section_view.OriginX = x
                section_view.OriginY = y

            result = {
                "status": "added",
                "source_view_index": view_index,
                "section_type": "standard" if section_type == 0 else "revolved",
                "position": [x, y],
                "total_cutting_planes": cutting_planes.Count,
            }

            with contextlib.suppress(Exception):
                result["caption"] = cutting_plane.Caption
            with contextlib.suppress(Exception):
                result["section_view_name"] = section_view.Name

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            if not hasattr(view, "Dimensions"):
                return {"count": 0, "dimensions": [], "note": "No Dimensions on this view"}

            dims = view.Dimensions

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            doc.UpdateAll(force_update)
            return {"status": "updated_all", "force_update": force_update}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
