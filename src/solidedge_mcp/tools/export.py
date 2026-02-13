"""Export, drawing, and view tools for Solid Edge MCP."""

from typing import Optional, List
from solidedge_mcp.managers import export_manager, view_manager


# === Export Formats ===

def export_step(file_path: str) -> dict:
    """Export the active document to STEP format."""
    return export_manager.export_step(file_path)

def export_stl(file_path: str) -> dict:
    """Export the active document to STL format."""
    return export_manager.export_stl(file_path)

def export_iges(file_path: str) -> dict:
    """Export the active document to IGES format."""
    return export_manager.export_iges(file_path)

def export_pdf(file_path: str) -> dict:
    """Export the active document to PDF format."""
    return export_manager.export_pdf(file_path)

def export_dxf(file_path: str) -> dict:
    """Export the active document to DXF format."""
    return export_manager.export_dxf(file_path)

def export_parasolid(file_path: str) -> dict:
    """Export the active document to Parasolid format (.x_t)."""
    return export_manager.export_parasolid(file_path)

def export_jt(file_path: str) -> dict:
    """Export the active document to JT format."""
    return export_manager.export_jt(file_path)

def export_flat_dxf(file_path: str) -> dict:
    """Export sheet metal flat pattern to DXF format."""
    return export_manager.export_flat_dxf(file_path)


# === Drawing / Draft ===

def create_drawing(template: Optional[str] = None, views: Optional[List[str]] = None) -> dict:
    """Create a 2D drawing from the active 3D model."""
    return export_manager.create_drawing(template, views)

def add_draft_sheet() -> dict:
    """Add a new sheet to the active draft document."""
    return export_manager.add_draft_sheet()

def add_assembly_drawing_view(x: float = 0.15, y: float = 0.15, orientation: str = "Isometric", scale: float = 1.0) -> dict:
    """Add an assembly drawing view to the active draft."""
    return export_manager.add_assembly_drawing_view(x, y, orientation, scale)

def create_parts_list(auto_balloon: bool = True, x: float = 0.15, y: float = 0.25) -> dict:
    """Create a parts list (BOM) on the active draft."""
    return export_manager.create_parts_list(auto_balloon, x, y)

def capture_screenshot(file_path: str, width: int = 1920, height: int = 1080) -> dict:
    """Capture a screenshot of the current view."""
    return export_manager.capture_screenshot(file_path, width, height)


# === Draft Annotations ===

def add_text_box(x: float, y: float, text: str, height: float = 0.005) -> dict:
    """Add a text box to the active draft."""
    return export_manager.add_text_box(x, y, text, height)

def add_leader(x1: float, y1: float, x2: float, y2: float, text: str = "") -> dict:
    """Add a leader annotation to the active draft."""
    return export_manager.add_leader(x1, y1, x2, y2, text)

def add_dimension(x1: float, y1: float, x2: float, y2: float, dim_x: Optional[float] = None, dim_y: Optional[float] = None) -> dict:
    """Add a linear dimension to the active draft."""
    return export_manager.add_dimension(x1, y1, x2, y2, dim_x, dim_y)

def add_balloon(x: float, y: float, text: str = "", leader_x: Optional[float] = None, leader_y: Optional[float] = None) -> dict:
    """Add a balloon annotation to the active draft."""
    return export_manager.add_balloon(x, y, text, leader_x, leader_y)

def add_note(x: float, y: float, text: str, height: float = 0.005) -> dict:
    """Add a note annotation to the active draft."""
    return export_manager.add_note(x, y, text, height)


# === Drawing View Management ===

def get_drawing_view_count() -> dict:
    """Get the number of drawing views on the active sheet."""
    return export_manager.get_drawing_view_count()

def get_drawing_view_scale(view_index: int) -> dict:
    """Get the scale of a drawing view by 0-based index."""
    return export_manager.get_drawing_view_scale(view_index)

def set_drawing_view_scale(view_index: int, scale: float) -> dict:
    """Set the scale of a drawing view by 0-based index."""
    return export_manager.set_drawing_view_scale(view_index, scale)

def delete_drawing_view(view_index: int) -> dict:
    """Delete a drawing view from the active sheet by 0-based index."""
    return export_manager.delete_drawing_view(view_index)

def update_drawing_view(view_index: int) -> dict:
    """Force update a drawing view to reflect 3D model changes."""
    return export_manager.update_drawing_view(view_index)


# === Draft Sheet Management ===

def get_sheet_info() -> dict:
    """Get information about the active draft sheet."""
    return export_manager.get_sheet_info()

def activate_sheet(sheet_index: int) -> dict:
    """Activate a draft sheet by index."""
    return export_manager.activate_sheet(sheet_index)

def rename_sheet(sheet_index: int, new_name: str) -> dict:
    """Rename a draft sheet."""
    return export_manager.rename_sheet(sheet_index, new_name)

def delete_sheet(sheet_index: int) -> dict:
    """Delete a draft sheet."""
    return export_manager.delete_sheet(sheet_index)


# === View Controls ===

def set_view_orientation(view: str) -> dict:
    """Set the viewing orientation ('Iso', 'Top', 'Front', 'Right', 'Left', 'Back', 'Bottom')."""
    return view_manager.set_view(view)

def zoom_fit() -> dict:
    """Zoom to fit all geometry in the view."""
    return view_manager.zoom_fit()

def zoom_to_selection() -> dict:
    """Zoom to the currently selected geometry."""
    return view_manager.zoom_to_selection()

def set_display_mode(mode: str) -> dict:
    """Set the display mode ('Shaded', 'ShadedWithEdges', 'Wireframe', 'HiddenEdgesVisible')."""
    return view_manager.set_display_mode(mode)

def set_view_background(red: int, green: int, blue: int) -> dict:
    """Set the view background color."""
    return view_manager.set_view_background(red, green, blue)

def rotate_view(angle: float,
                center_x: float = 0.0, center_y: float = 0.0, center_z: float = 0.0,
                axis_x: float = 0.0, axis_y: float = 1.0, axis_z: float = 0.0) -> dict:
    """Rotate the view around an axis through a center point.

    Args:
        angle: Rotation angle in radians
        center_x, center_y, center_z: Center of rotation (meters)
        axis_x, axis_y, axis_z: Rotation axis vector (default: Y-up)
    """
    return view_manager.rotate_camera(angle, center_x, center_y, center_z, axis_x, axis_y, axis_z)

def pan_view(dx: int, dy: int) -> dict:
    """Pan the view by pixel offsets.

    Args:
        dx: Horizontal pan in pixels (positive = right)
        dy: Vertical pan in pixels (positive = down)
    """
    return view_manager.pan_camera(dx, dy)

def zoom_view(factor: float) -> dict:
    """Zoom the view by a scale factor (>1 = zoom in, <1 = zoom out)."""
    return view_manager.zoom_camera(factor)

def refresh_view() -> dict:
    """Force the active view to refresh/update."""
    return view_manager.refresh_view()

def get_camera() -> dict:
    """Get the current camera position and orientation."""
    return view_manager.get_camera()

def set_camera(eye_x: float, eye_y: float, eye_z: float,
               target_x: float, target_y: float, target_z: float,
               up_x: float = 0.0, up_y: float = 1.0, up_z: float = 0.0,
               perspective: bool = False, scale_or_angle: float = 1.0) -> dict:
    """Set the camera position and orientation."""
    return view_manager.set_camera(
        eye_x, eye_y, eye_z,
        target_x, target_y, target_z,
        up_x, up_y, up_z,
        perspective, scale_or_angle
    )


# === Registration ===

def register(mcp):
    """Register export, drawing, and view tools with the MCP server."""
    # Export Formats
    mcp.tool()(export_step)
    mcp.tool()(export_stl)
    mcp.tool()(export_iges)
    mcp.tool()(export_pdf)
    mcp.tool()(export_dxf)
    mcp.tool()(export_parasolid)
    mcp.tool()(export_jt)
    mcp.tool()(export_flat_dxf)
    # Drawing / Draft
    mcp.tool()(create_drawing)
    mcp.tool()(add_draft_sheet)
    mcp.tool()(add_assembly_drawing_view)
    mcp.tool()(create_parts_list)
    mcp.tool()(capture_screenshot)
    # Draft Annotations
    mcp.tool()(add_text_box)
    mcp.tool()(add_leader)
    mcp.tool()(add_dimension)
    mcp.tool()(add_balloon)
    mcp.tool()(add_note)
    # Drawing View Management
    mcp.tool()(get_drawing_view_count)
    mcp.tool()(get_drawing_view_scale)
    mcp.tool()(set_drawing_view_scale)
    mcp.tool()(delete_drawing_view)
    mcp.tool()(update_drawing_view)
    # Draft Sheet Management
    mcp.tool()(get_sheet_info)
    mcp.tool()(activate_sheet)
    mcp.tool()(rename_sheet)
    mcp.tool()(delete_sheet)
    # View Controls
    mcp.tool()(set_view_orientation)
    mcp.tool()(zoom_fit)
    mcp.tool()(zoom_to_selection)
    mcp.tool()(set_display_mode)
    mcp.tool()(set_view_background)
    mcp.tool()(rotate_view)
    mcp.tool()(pan_view)
    mcp.tool()(zoom_view)
    mcp.tool()(refresh_view)
    mcp.tool()(get_camera)
    mcp.tool()(set_camera)
