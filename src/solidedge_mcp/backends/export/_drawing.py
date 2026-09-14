"""Drawing creation, sheet management, and drawing view creation operations."""

import contextlib
import os
from typing import Any

from solidedge_mcp.backends.errors import describe_exception, error_result

from ..constants import PartDrawingViewTypeConstants, ViewOrientationConstants
from ..logging import get_logger
from ._base import NOT_A_DRAFT, com_get

_logger = get_logger(__name__)

#: constant.tlb > ViewOrientationConstants values accepted by the DrawingViews
#: Add* methods and by DrawingView.SetViewOrientationStandard, keyed by the
#: names this server exposes. This is the one mapping; three copies of it had
#: drifted onto a set of values Solid Edge does not use.
VIEW_ORIENTATIONS: dict[str, int] = {
    "Front": ViewOrientationConstants.igFrontView,
    "Back": ViewOrientationConstants.igBackView,
    "Top": ViewOrientationConstants.igTopView,
    "Bottom": ViewOrientationConstants.igBottomView,
    "Right": ViewOrientationConstants.igRightView,
    "Left": ViewOrientationConstants.igLeftView,
    "Isometric": ViewOrientationConstants.igTopFrontRightView,
    "Iso": ViewOrientationConstants.igTopFrontRightView,
}


class DrawingMixin:
    """Mixin providing drawing creation and sheet management methods."""

    def create_drawing(
        self, template: str | None = None, views: list[str] | None = None
    ) -> dict[str, Any]:
        """
        Create a 2D drawing from the active 3D model.

        The active document must be a saved Part or Assembly. A new Draft document
        is created with a model link to the source, then drawing views are placed.

        Args:
            template: Drawing template path (optional)
            views: List of views to create - ['Front', 'Top', 'Right', 'Isometric']

        Returns:
            Dict with status and drawing info
        """
        try:
            import win32com.client.dynamic as dyn

            source_doc = self.doc_manager.get_active_document()
            app = self.doc_manager.connection.get_application()

            # Source must be saved to disk for ModelLinks.Add
            source_path = None
            with contextlib.suppress(Exception):
                source_path = source_doc.FullName

            if not source_path:
                return {
                    "error": "Active document must be saved "
                    "before creating a drawing. "
                    "Use save_document() first."
                }

            if views is None:
                views = ["Front", "Top", "Right", "Isometric"]

            view_orient_map = VIEW_ORIENTATIONS

            # Create a new draft document
            if template and os.path.exists(template):
                draft_doc = app.Documents.Add(template)
            else:
                draft_doc = app.Documents.Add("SolidEdge.DraftDocument")

            # Add model link to the source part/assembly
            model_link = draft_doc.ModelLinks.Add(source_path)

            # Get active sheet and DrawingViews
            sheet = draft_doc.ActiveSheet
            # DrawingViews may bind to wrong type library (Part instead of Draft)
            # due to gen_py cache. Force late binding on the DrawingViews object.
            dvs_early = sheet.DrawingViews
            dvs = dyn.Dispatch(dvs_early._oleobj_)

            # Place views in a grid layout on the sheet
            # Standard A-size sheet is ~0.279 x 0.216 m (A4 landscape)
            positions = [
                (0.10, 0.15),  # View 1
                (0.10, 0.06),  # View 2
                (0.22, 0.15),  # View 3
                (0.22, 0.06),  # View 4
            ]

            views_added = []
            views_failed: dict[str, str] = {}
            for i, view_name in enumerate(views[:4]):
                orient = view_orient_map.get(view_name)
                if orient is None:
                    continue

                x, y = positions[i] if i < len(positions) else (0.15 + i * 0.08, 0.10)

                try:
                    dvs.AddPartView(model_link, orient, 1.0, x, y, 0)
                    views_added.append(view_name)
                except Exception as part_view_error:
                    # DrawingViews.Add is the generic form of the same call.
                    try:
                        dvs.Add(model_link, orient, 1.0, x, y)
                        views_added.append(view_name)
                    except Exception as add_error:
                        # A dropped view used to leave nothing behind but its
                        # absence from views_added, so a caller could see that
                        # a view was missing and never why.
                        views_failed[view_name] = (
                            f"AddPartView: {describe_exception(part_view_error)}; "
                            f"Add: {describe_exception(add_error)}"
                        )

            result: dict[str, Any] = {
                "status": "created",
                "type": "drawing",
                "draft_name": com_get(draft_doc, "Name", "Draft"),
                "model_link": source_path,
                "views_requested": views,
                "views_added": views_added,
                "total_views": len(views_added),
            }
            if views_failed:
                result["views_failed"] = views_failed
            return result
        except Exception as e:
            return error_result(e)

    def add_model_drawing_view(
        self,
        x: float = 0.15,
        y: float = 0.15,
        orientation: str = "Isometric",
        scale: float = 1.0,
        model: str = "part",
    ) -> dict[str, Any]:
        """Add a drawing view of a part, sheet metal part, or weldment.

        draft.tlb exposes a separate DrawingViews method per model kind:
        AddPartView, AddSheetMetalView and AddWeldmentView, each
        ``(From, Orientation, Scale, x, y, ViewType)``. Only the assembly one
        was wired up, so drafting a part -- the ordinary case -- had no route
        of its own and depended on an untyped fallback.

        Args:
            x: View centre X on the sheet, in meters.
            y: View centre Y on the sheet, in meters.
            orientation: 'Front', 'Back', 'Top', 'Bottom', 'Right', 'Left',
                'Isometric'.
            scale: View scale factor.
            model: 'part', 'sheet_metal' or 'weldment'.

        Returns:
            Dict with status and view info.
        """
        try:
            import win32com.client.dynamic as dyn

            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            methods = {
                "part": "AddPartView",
                "sheet_metal": "AddSheetMetalView",
                "weldment": "AddWeldmentView",
            }
            method_name = methods.get(model)
            if method_name is None:
                return {
                    "error": f"Unknown model kind: {model}. Valid: {', '.join(methods)}",
                }

            if not com_get(com_get(doc, "ModelLinks"), "Count", 0):
                return {
                    "error": "No model link found. Create a drawing with create_drawing() first."
                }
            model_link = doc.ModelLinks.Item(1)

            orient = VIEW_ORIENTATIONS.get(orientation)
            if orient is None:
                valid = ", ".join(VIEW_ORIENTATIONS)
                return {"error": f"Invalid orientation: {orientation}. Valid: {valid}"}

            sheet = doc.ActiveSheet
            dvs = dyn.Dispatch(sheet.DrawingViews._oleobj_)
            getattr(dvs, method_name)(
                model_link,
                orient,
                scale,
                x,
                y,
                PartDrawingViewTypeConstants.sePartDesignedView,
            )

            return {
                "status": "added",
                "model": model,
                "orientation": orientation,
                "scale": scale,
                "position": [x, y],
            }
        except Exception as e:
            return error_result(e)

    def add_assembly_drawing_view(
        self, x: float = 0.15, y: float = 0.15, orientation: str = "Isometric", scale: float = 1.0
    ) -> dict[str, Any]:
        """
        Add an assembly drawing view to the active draft document.

        The active document must be a Draft with a model link to an assembly.

        Args:
            x: View center X position on sheet (meters)
            y: View center Y position on sheet (meters)
            orientation: View orientation - 'Front', 'Top', 'Right', 'Isometric', etc.
            scale: View scale factor

        Returns:
            Dict with status and view info
        """
        try:
            import win32com.client.dynamic as dyn

            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            # Get model link
            if not com_get(com_get(doc, "ModelLinks"), "Count", 0):
                return {
                    "error": "No model link found. Create a drawing with create_drawing() first."
                }

            model_link = doc.ModelLinks.Item(1)

            # View orientation constants
            view_orient_map = {
                "Front": 5,
                "Back": 8,
                "Top": 6,
                "Bottom": 9,
                "Right": 7,
                "Left": 10,
                "Isometric": 12,
                "Iso": 12,
            }

            orient = view_orient_map.get(orientation)
            if orient is None:
                valid = ", ".join(view_orient_map.keys())
                return {"error": f"Invalid orientation: {orientation}. Valid: {valid}"}

            sheet = doc.ActiveSheet
            dvs_early = sheet.DrawingViews
            dvs = dyn.Dispatch(dvs_early._oleobj_)

            # seAssemblyDesignedView = 0
            try:
                dvs.AddAssemblyView(model_link, orient, scale, x, y, 0)
            except Exception:
                # A model link to a part rejects AddAssemblyView. Use the
                # matching API rather than reporting a failure; add_part_view
                # is the explicit route.
                dvs.AddPartView(
                    model_link, orient, scale, x, y, PartDrawingViewTypeConstants.sePartDesignedView
                )

            return {
                "status": "added",
                "orientation": orientation,
                "scale": scale,
                "position": [x, y],
            }
        except Exception as e:
            return error_result(e)

    def create_parts_list(
        self, auto_balloon: bool = True, x: float = 0.15, y: float = 0.25
    ) -> dict[str, Any]:
        """
        Create a parts list (BOM table) on the active draft sheet.

        Requires the active document to be a Draft with at least one drawing view.
        Uses PartsLists.Add(DrawingView, SavedSettings, AutoBalloon, CreatePartsList).

        Args:
            auto_balloon: Whether to auto-generate balloon callouts
            x: Table X position on sheet, in meters, applied with
                PartsList.SetOrigin after the table is added
            y: Table Y position on sheet, in meters

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            import win32com.client.dynamic as dyn

            sheet = doc.ActiveSheet

            # sheet.DrawingViews binds to the Part type library's
            # SketchDrawingViews through the gen_py cache, which does not expose
            # the members needed here. Force late binding, as _get_drawing_views
            # and create_drawing do, and fall back if that is not possible.
            dvs = sheet.DrawingViews
            with contextlib.suppress(Exception):
                dvs = dyn.Dispatch(dvs._oleobj_)
            if dvs.Count == 0:
                return {"error": "No drawing views on active sheet. Add a view first."}

            dv = dvs.Item(1)

            # PartsLists is on DraftDocument, not on Sheet; the sheet
            # lookup could only ever return None.
            parts_lists = com_get(doc, "PartsLists")

            if parts_lists is None:
                return {"error": "PartsLists collection not available"}

            # Add parts list
            # Parameters: DrawingView, SavedSettings, AutoBalloon, CreatePartsList
            # AutoBalloon: 0=No, 1=Yes; CreatePartsList: 0=No, 1=Yes
            parts_list = parts_lists.Add(dv, "", 1 if auto_balloon else 0, 1)

            # PartsLists.Add takes no position, so x/y were accepted and
            # dropped: the table landed wherever Solid Edge chose and the
            # result still reported the position the caller asked for.
            # PartsList.SetOrigin(x, y) is the placement API.
            placed = False
            try:
                parts_list.SetOrigin(x, y)
                placed = True
            except Exception as exc:  # noqa: BLE001
                _logger.warning(f"Parts list created but could not be positioned: {exc}")

            return {
                "status": "created",
                "auto_balloon": auto_balloon,
                "position": [x, y] if placed else None,
                "positioned": placed,
                "total_parts_lists": parts_lists.Count,
            }
        except Exception as e:
            return error_result(e)

    def add_assembly_drawing_view_ex(
        self,
        x: float = 0.15,
        y: float = 0.15,
        orientation: str = "Isometric",
        scale: float = 1.0,
        config: str | None = None,
    ) -> dict[str, Any]:
        """
        Add an extended assembly drawing view with optional configuration.

        Similar to add_assembly_drawing_view but with configuration support.

        Args:
            x: View center X position on sheet (meters)
            y: View center Y position on sheet (meters)
            orientation: View orientation ('Front', 'Top', 'Right', 'Isometric', etc.)
            scale: View scale factor
            config: Optional configuration name

        Returns:
            Dict with status and view info
        """
        try:
            import win32com.client.dynamic as dyn

            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            if not com_get(com_get(doc, "ModelLinks"), "Count", 0):
                return {
                    "error": "No model link found. Create a drawing with create_drawing() first."
                }

            model_link = doc.ModelLinks.Item(1)

            view_orient_map = {
                "Front": 5,
                "Back": 8,
                "Top": 6,
                "Bottom": 9,
                "Right": 7,
                "Left": 10,
                "Isometric": 12,
                "Iso": 12,
            }

            orient = view_orient_map.get(orientation)
            if orient is None:
                valid = ", ".join(view_orient_map.keys())
                return {"error": f"Invalid orientation: {orientation}. Valid: {valid}"}

            sheet = doc.ActiveSheet
            dvs_early = sheet.DrawingViews
            dvs = dyn.Dispatch(dvs_early._oleobj_)

            # There is no AddAssemblyViewWithConfiguration. The configuration
            # is the seventh argument of AddAssemblyView(From, Orientation,
            # Scale, x, y, ViewType, ConfigurationName, Snapshot,
            # SnapShotQuality), so this always raised and fell through, and the
            # requested configuration was silently ignored.
            if config is not None:
                try:
                    dvs.AddAssemblyView(model_link, orient, scale, x, y, 0, config)
                    return {
                        "status": "added",
                        "orientation": orientation,
                        "scale": scale,
                        "position": [x, y],
                        "configuration": config,
                    }
                except Exception:
                    pass  # Fall through to standard method

            # Standard assembly view
            try:
                dvs.AddAssemblyView(model_link, orient, scale, x, y, 0)
            except Exception:
                dvs.AddPartView(model_link, orient, scale, x, y, 0)

            result = {
                "status": "added",
                "orientation": orientation,
                "scale": scale,
                "position": [x, y],
            }
            if config is not None:
                result["configuration"] = config
                result["note"] = "Configuration param not applied; used standard view."
            return result
        except Exception as e:
            return error_result(e)

    def add_drawing_view_with_config(
        self,
        x: float = 0.15,
        y: float = 0.15,
        orientation: str = "Front",
        scale: float = 1.0,
        configuration: str = "Default",
    ) -> dict[str, Any]:
        """
        Add a drawing view with a specific configuration.

        Uses DrawingViews.AddWithConfiguration if available.

        Args:
            x: View center X position on sheet (meters)
            y: View center Y position on sheet (meters)
            orientation: View orientation ('Front', 'Top', 'Right', etc.)
            scale: View scale factor
            configuration: Configuration name

        Returns:
            Dict with status and view info
        """
        try:
            import win32com.client.dynamic as dyn

            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            if not com_get(com_get(doc, "ModelLinks"), "Count", 0):
                return {
                    "error": "No model link found. Create a drawing with create_drawing() first."
                }

            model_link = doc.ModelLinks.Item(1)

            view_orient_map = {
                "Front": 5,
                "Back": 8,
                "Top": 6,
                "Bottom": 9,
                "Right": 7,
                "Left": 10,
                "Isometric": 12,
                "Iso": 12,
            }

            orient = view_orient_map.get(orientation)
            if orient is None:
                valid = ", ".join(view_orient_map.keys())
                return {"error": f"Invalid orientation: {orientation}. Valid: {valid}"}

            sheet = doc.ActiveSheet
            dvs_early = sheet.DrawingViews
            dvs = dyn.Dispatch(dvs_early._oleobj_)

            # draft.tlb names this AddPartViewByConfiguration, not
            # AddPartViewWithConfiguration, so this always raised and the
            # configuration was silently dropped by the fallback.
            configured = False
            if configuration:
                try:
                    dvs.AddPartViewByConfiguration(
                        model_link, orient, scale, x, y, 0, configuration
                    )
                    configured = True
                except Exception:
                    _logger.warning(
                        "Configuration %r was rejected; adding a plain part view",
                        configuration,
                    )
            if not configured:
                dvs.AddPartView(model_link, orient, scale, x, y, 0)

            return {
                "status": "added",
                "orientation": orientation,
                "scale": scale,
                "position": [x, y],
                "configuration": configuration if configured else None,
            }
        except Exception as e:
            return error_result(e)

    # =================================================================
    # SHEET MANAGEMENT
    # =================================================================

    def add_draft_sheet(self) -> dict[str, Any]:
        """
        Add a new sheet to the active draft document.

        The active document must be a Draft document. Creates a new sheet
        and activates it.

        Returns:
            Dict with status and new sheet info
        """
        try:
            doc = self.doc_manager.get_active_document()

            err = self._require_draft(
                doc, "Active document is not a draft document. Create a drawing first."
            )
            if err:
                return err

            sheets = doc.Sheets

            sheet = sheets.AddSheet()
            sheet.Activate()

            return {
                "status": "added",
                "sheet_number": sheets.Count,
                "total_sheets": sheets.Count,
                "name": com_get(sheet, "Name", f"Sheet {sheets.Count}"),
            }
        except Exception as e:
            return error_result(e)

    def get_sheet_info(self) -> dict[str, Any]:
        """
        Get information about the active draft sheet.

        Returns:
            Dict with sheet name, size, scale, and counts of drawing objects
        """
        try:
            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            sheet = doc.ActiveSheet
            sheets = doc.Sheets

            info = {
                "status": "ok",
                "sheet_name": sheet.Name,
                "sheet_count": sheets.Count,
            }

            # SheetWidth and SheetHeight are on SheetSetup, not on Sheet, so
            # both keys were always missing from the result.
            with contextlib.suppress(Exception):
                setup = sheet.SheetSetup
                info["width"] = setup.SheetWidth
                info["height"] = setup.SheetHeight

            # Sheet.Background is a Sheet object, not a string. Putting the
            # proxy in the result made json.dumps raise "Object of type
            # CDispatch is not JSON serializable" for the whole resource.
            with contextlib.suppress(Exception):
                background = sheet.Background
                info["background"] = com_get(background, "Name", str(background))

            # Count drawing objects
            with contextlib.suppress(Exception):
                info["drawing_views"] = sheet.DrawingViews.Count
            with contextlib.suppress(Exception):
                info["text_boxes"] = sheet.TextBoxes.Count
            with contextlib.suppress(Exception):
                info["leaders"] = sheet.Leaders.Count
            with contextlib.suppress(Exception):
                info["dimensions"] = sheet.Dimensions.Count

            return info
        except Exception as e:
            return error_result(e)

    def activate_sheet(self, sheet_index: int) -> dict[str, Any]:
        """
        Activate a specific draft sheet by index.

        Args:
            sheet_index: 0-based sheet index

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            sheets = doc.Sheets
            if sheet_index < 0 or sheet_index >= sheets.Count:
                return {"error": f"Invalid sheet index: {sheet_index}. Count: {sheets.Count}"}

            sheet = sheets.Item(sheet_index + 1)  # COM is 1-indexed
            sheet.Activate()

            return {"status": "activated", "sheet_name": sheet.Name, "sheet_index": sheet_index}
        except Exception as e:
            return error_result(e)

    def rename_sheet(self, sheet_index: int, new_name: str) -> dict[str, Any]:
        """
        Rename a draft sheet.

        Args:
            sheet_index: 0-based sheet index
            new_name: New name for the sheet

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            sheets = doc.Sheets
            if sheet_index < 0 or sheet_index >= sheets.Count:
                return {"error": f"Invalid sheet index: {sheet_index}. Count: {sheets.Count}"}

            sheet = sheets.Item(sheet_index + 1)
            old_name = sheet.Name
            sheet.Name = new_name

            return {
                "status": "renamed",
                "old_name": old_name,
                "new_name": new_name,
                "sheet_index": sheet_index,
            }
        except Exception as e:
            return error_result(e)

    def delete_sheet(self, sheet_index: int) -> dict[str, Any]:
        """
        Delete a draft sheet.

        Args:
            sheet_index: 0-based sheet index

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            sheets = doc.Sheets
            if sheets.Count <= 1:
                return {"error": "Cannot delete the last sheet"}

            if sheet_index < 0 or sheet_index >= sheets.Count:
                return {"error": f"Invalid sheet index: {sheet_index}. Count: {sheets.Count}"}

            sheet = sheets.Item(sheet_index + 1)
            sheet_name = sheet.Name
            sheet.Delete()

            return {"status": "deleted", "sheet_name": sheet_name, "remaining_sheets": sheets.Count}
        except Exception as e:
            return error_result(e)

    # =================================================================
    # SHEET COLLECTION QUERIES
    # =================================================================

    def get_sheet_dimensions(self) -> dict[str, Any]:
        """
        Get all dimensions on the active draft sheet.

        Iterates sheet.Dimensions collection and collects type, value, and
        position where available.

        Returns:
            Dict with count and list of dimension info dicts
        """
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_draft(doc, NOT_A_DRAFT)
            if err:
                return err
            sheet = doc.ActiveSheet
            dims = sheet.Dimensions
            items = []
            for i in range(1, dims.Count + 1):
                dim = dims.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    info["type"] = dim.Type
                with contextlib.suppress(Exception):
                    info["value"] = dim.Value
                with contextlib.suppress(Exception):
                    info["name"] = dim.Name
                items.append(info)
            return {"count": len(items), "dimensions": items}
        except Exception as e:
            return error_result(e)

    def get_sheet_balloons(self) -> dict[str, Any]:
        """
        Get all balloons on the active draft sheet.

        Iterates sheet.Balloons collection and collects text and position
        where available.

        Returns:
            Dict with count and list of balloon info dicts
        """
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_draft(doc, NOT_A_DRAFT)
            if err:
                return err
            sheet = doc.ActiveSheet
            balloons = sheet.Balloons
            items = []
            for i in range(1, balloons.Count + 1):
                balloon = balloons.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    info["text"] = balloon.BalloonText
                with contextlib.suppress(Exception):
                    info["x"] = balloon.x
                with contextlib.suppress(Exception):
                    info["y"] = balloon.y
                items.append(info)
            return {"count": len(items), "balloons": items}
        except Exception as e:
            return error_result(e)

    def get_sheet_text_boxes(self) -> dict[str, Any]:
        """
        Get all text boxes on the active draft sheet.

        Iterates sheet.TextBoxes collection and collects text content
        and position where available.

        Returns:
            Dict with count and list of text box info dicts
        """
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_draft(doc, NOT_A_DRAFT)
            if err:
                return err
            sheet = doc.ActiveSheet
            text_boxes = sheet.TextBoxes
            items = []
            for i in range(1, text_boxes.Count + 1):
                tb = text_boxes.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    info["text"] = tb.Text
                # TextBox has no x/y properties, so both keys were always
                # missing. Its position comes from GetOrigin's out-parameters.
                with contextlib.suppress(Exception):
                    origin = tb.GetOrigin()
                    info["x"] = float(origin[0])
                    info["y"] = float(origin[1])
                with contextlib.suppress(Exception):
                    info["height"] = tb.Height
                items.append(info)
            return {"count": len(items), "text_boxes": items}
        except Exception as e:
            return error_result(e)

    def get_sheet_drawing_objects(self) -> dict[str, Any]:
        """
        Get all drawing objects on the active draft sheet.

        Iterates sheet.DrawingObjects collection and collects type and
        name where available.

        Returns:
            Dict with count and list of drawing object info dicts
        """
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_draft(doc, NOT_A_DRAFT)
            if err:
                return err
            sheet = doc.ActiveSheet
            drawing_objects = sheet.DrawingObjects
            items = []
            for i in range(1, drawing_objects.Count + 1):
                obj = drawing_objects.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    info["type"] = str(type(obj).__name__)
                with contextlib.suppress(Exception):
                    info["name"] = obj.Name
                items.append(info)
            return {"count": len(items), "drawing_objects": items}
        except Exception as e:
            return error_result(e)

    def get_sheet_sections(self) -> dict[str, Any]:
        """
        Get all section views on the active draft sheet.

        Iterates sheet.Sections collection and collects label and
        info where available.

        Returns:
            Dict with count and list of section view info dicts
        """
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_draft(doc, NOT_A_DRAFT)
            if err:
                return err
            # Sections is on DraftDocument, not on Sheet.
            sections = doc.Sections
            items = []
            for i in range(1, sections.Count + 1):
                sec = sections.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                # Section has Name, Index, Type and Sheets. It has no Label,
                # so that key was always missing.
                with contextlib.suppress(Exception):
                    info["name"] = sec.Name
                with contextlib.suppress(Exception):
                    info["type"] = sec.Type
                with contextlib.suppress(Exception):
                    info["sheet_count"] = sec.Sheets.Count
                items.append(info)
            return {"count": len(items), "sections": items}
        except Exception as e:
            return error_result(e)

    def get_drawing_view_count(self) -> dict[str, Any]:
        """
        Get the number of drawing views on the active sheet.

        Returns:
            Dict with view count
        """
        try:
            dvs = self._get_drawing_views()
            return {"count": dvs.Count}
        except Exception as e:
            return error_result(e)

    def get_drawing_view_scale(self, view_index: int) -> dict[str, Any]:
        """
        Get the scale of a drawing view.

        Args:
            view_index: 0-based view index

        Returns:
            Dict with scale value
        """
        try:
            dvs = self._get_drawing_views()
            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)
            scale = view.ScaleFactor

            return {"view_index": view_index, "scale": scale}
        except Exception as e:
            return error_result(e)
