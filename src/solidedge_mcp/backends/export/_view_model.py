"""ViewModel class for view manipulation (orientation, zoom, display, camera)."""

import traceback
from typing import Any

from ..constants import RenderModeConstants


class ViewModel:
    """Manages view manipulation"""

    def __init__(self, document_manager):
        self.doc_manager = document_manager

    def set_view(self, view: str) -> dict[str, Any]:
        """
        Set the viewing orientation.

        Args:
            view: View orientation - 'Iso', 'Top', 'Front', 'Right', 'Bottom', 'Back', 'Left'

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            # Get the window
            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            # Valid view names (discovered via introspection)
            # Note: Bottom, Back, Left may not work in all contexts
            valid_views = ["Iso", "Top", "Front", "Right", "Bottom", "Back", "Left"]

            if view not in valid_views:
                return {"error": f"Invalid view: {view}. Valid: {', '.join(valid_views)}"}

            # Use ApplyNamedView with string name (discovered method!)
            if hasattr(view_obj, "ApplyNamedView"):
                view_obj.ApplyNamedView(view)
                return {"status": "view_set", "view": view}
            else:
                return {
                    "error": "ApplyNamedView not available",
                    "note": "Use View menu in Solid Edge UI",
                }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def zoom_fit(self) -> dict[str, Any]:
        """Zoom to fit all geometry in view"""
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            # Zoom to fit
            if hasattr(view_obj, "Fit"):
                view_obj.Fit()
                return {"status": "zoomed_fit"}
            else:
                return {
                    "error": "Fit method not available",
                    "note": "Use View > Fit in Solid Edge UI",
                }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def zoom_to_selection(self) -> dict[str, Any]:
        """Zoom to fit all geometry (equivalent to View > Fit)."""
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            view_obj.Fit()

            return {"status": "zoomed_to_selection"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_display_mode(self, mode: str) -> dict[str, Any]:
        """
        Set the display mode for the active view.

        Args:
            mode: Display mode - 'Shaded', 'ShadedWithEdges', 'Wireframe', 'HiddenEdgesVisible'

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            mode_map = {
                "Wireframe": RenderModeConstants.seRenderModeWireframe,
                "Shaded": RenderModeConstants.seRenderModeSmooth,
                "ShadedWithEdges": RenderModeConstants.seRenderModeSmoothBoundary,
                "HiddenEdgesVisible": RenderModeConstants.seRenderModeVHL,
            }

            mode_value = mode_map.get(mode)
            if mode_value is None:
                return {
                    "error": f"Invalid mode: {mode}. "
                    "Use 'Shaded', 'ShadedWithEdges', "
                    "'Wireframe', or "
                    "'HiddenEdgesVisible'"
                }

            view_obj.SetRenderMode(mode_value)
            return {"status": "display_mode_set", "mode": mode}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_view_background(self, red: int, green: int, blue: int) -> dict[str, Any]:
        """
        Set the view background color.

        Args:
            red: Red component (0-255)
            green: Green component (0-255)
            blue: Blue component (0-255)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            ole_color = red | (green << 8) | (blue << 16)

            try:
                view_obj.SetBackgroundColor(ole_color)
            except Exception:
                try:
                    view_obj.BackgroundColor = ole_color
                except Exception:
                    view_obj.SetBackgroundGradientColor(ole_color, ole_color)

            return {"status": "updated", "color": [red, green, blue]}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_camera(self) -> dict[str, Any]:
        """
        Get the current camera parameters.

        Returns eye position, target position, up vector, perspective flag,
        and scale/field-of-view angle.

        Returns:
            Dict with camera parameters
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            # GetCamera returns 11 out-params by reference
            result = view_obj.GetCamera()

            # result is a tuple: (EyeX, EyeY, EyeZ, TargetX, TargetY, TargetZ,
            #                     UpX, UpY, UpZ, Perspective, ScaleOrAngle)
            return {
                "eye": [result[0], result[1], result[2]],
                "target": [result[3], result[4], result[5]],
                "up": [result[6], result[7], result[8]],
                "perspective": bool(result[9]),
                "scale_or_angle": result[10],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def rotate_camera(
        self,
        angle: float,
        center_x: float = 0.0,
        center_y: float = 0.0,
        center_z: float = 0.0,
        axis_x: float = 0.0,
        axis_y: float = 1.0,
        axis_z: float = 0.0,
    ) -> dict[str, Any]:
        """
        Rotate the camera around a specified axis through a center point.

        Args:
            angle: Rotation angle in radians
            center_x, center_y, center_z: Center of rotation (meters)
            axis_x, axis_y, axis_z: Rotation axis vector (default: Y-up)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            view_obj.RotateCamera(angle, center_x, center_y, center_z, axis_x, axis_y, axis_z)

            return {
                "status": "camera_rotated",
                "angle_rad": angle,
                "center": [center_x, center_y, center_z],
                "axis": [axis_x, axis_y, axis_z],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def pan_camera(self, dx: int, dy: int) -> dict[str, Any]:
        """
        Pan the camera by pixel offsets.

        Args:
            dx: Horizontal pan in pixels (positive = right)
            dy: Vertical pan in pixels (positive = down)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            view_obj.PanCamera(dx, dy)

            return {"status": "camera_panned", "dx": dx, "dy": dy}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def zoom_camera(self, factor: float) -> dict[str, Any]:
        """
        Zoom the camera by a scale factor.

        Args:
            factor: Zoom scale factor (>1 = zoom in, <1 = zoom out)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            view_obj.ZoomCamera(factor)

            return {"status": "camera_zoomed", "factor": factor}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def refresh_view(self) -> dict[str, Any]:
        """
        Force the active view to refresh/update.

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            view_obj.Update()

            return {"status": "view_refreshed"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def _get_view_object(self):
        """Get the active view object from the first window."""
        doc = self.doc_manager.get_active_document()
        if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
            raise Exception("No window available")
        window = doc.Windows.Item(1)
        view_obj = window.View if hasattr(window, "View") else None
        if not view_obj:
            raise Exception("Cannot access view object")
        return view_obj

    def transform_model_to_screen(self, x: float, y: float, z: float) -> dict[str, Any]:
        """
        Transform 3D model coordinates to 2D screen (device) coordinates.

        Uses View.ModelToScreenTransform or TransformModelToDC.

        Args:
            x: Model X coordinate (meters)
            y: Model Y coordinate (meters)
            z: Model Z coordinate (meters)

        Returns:
            Dict with screen_x and screen_y pixel coordinates
        """
        try:
            view_obj = self._get_view_object()
            result = view_obj.TransformModelToDC(x, y, z)
            return {
                "status": "success",
                "model": [x, y, z],
                "screen_x": result[0],
                "screen_y": result[1],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def transform_screen_to_model(self, screen_x: int, screen_y: int) -> dict[str, Any]:
        """
        Transform 2D screen (device) coordinates to 3D model coordinates.

        Uses View.TransformDCToModel.

        Args:
            screen_x: Screen X pixel coordinate
            screen_y: Screen Y pixel coordinate

        Returns:
            Dict with model x, y, z coordinates (meters)
        """
        try:
            view_obj = self._get_view_object()
            result = view_obj.TransformDCToModel(screen_x, screen_y)
            return {
                "status": "success",
                "screen": [screen_x, screen_y],
                "x": result[0],
                "y": result[1],
                "z": result[2],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def begin_camera_dynamics(self) -> dict[str, Any]:
        """
        Begin camera dynamics mode for smooth multi-step camera manipulation.

        Call this before a sequence of RotateCamera/PanCamera/ZoomCamera calls,
        then call end_camera_dynamics() when done. This prevents intermediate
        view updates and improves performance.

        Returns:
            Dict with status
        """
        try:
            view_obj = self._get_view_object()
            view_obj.BeginCameraDynamics()
            return {"status": "camera_dynamics_started"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def end_camera_dynamics(self) -> dict[str, Any]:
        """
        End camera dynamics mode and apply all pending camera changes.

        Call after a sequence of camera manipulations that were started with
        begin_camera_dynamics().

        Returns:
            Dict with status
        """
        try:
            view_obj = self._get_view_object()
            view_obj.EndCameraDynamics()
            return {"status": "camera_dynamics_ended"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_camera(
        self,
        eye_x: float,
        eye_y: float,
        eye_z: float,
        target_x: float,
        target_y: float,
        target_z: float,
        up_x: float = 0.0,
        up_y: float = 1.0,
        up_z: float = 0.0,
        perspective: bool = False,
        scale_or_angle: float = 1.0,
    ) -> dict[str, Any]:
        """
        Set the camera parameters for the active view.

        Args:
            eye_x, eye_y, eye_z: Camera eye (position) coordinates
            target_x, target_y, target_z: Camera target (look-at) coordinates
            up_x, up_y, up_z: Camera up vector (default: Y-up)
            perspective: True for perspective, False for orthographic
            scale_or_angle: View scale (ortho) or FOV angle in radians (perspective)

        Returns:
            Dict with status and camera settings
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            view_obj.SetCamera(
                eye_x,
                eye_y,
                eye_z,
                target_x,
                target_y,
                target_z,
                up_x,
                up_y,
                up_z,
                perspective,
                scale_or_angle,
            )

            return {
                "status": "camera_set",
                "eye": [eye_x, eye_y, eye_z],
                "target": [target_x, target_y, target_z],
                "up": [up_x, up_y, up_z],
                "perspective": perspective,
                "scale_or_angle": scale_or_angle,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
