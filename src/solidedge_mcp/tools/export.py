"""Export, drawing, and view tools for Solid Edge MCP."""

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


def create_drawing(template: str | None = None, views: list[str] | None = None) -> dict:
    """Create a 2D drawing from the active 3D model."""
    return export_manager.create_drawing(template, views)


def add_draft_sheet() -> dict:
    """Add a new sheet to the active draft document."""
    return export_manager.add_draft_sheet()


def add_assembly_drawing_view(
    x: float = 0.15, y: float = 0.15, orientation: str = "Isometric", scale: float = 1.0
) -> dict:
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


def add_dimension(
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    dim_x: float | None = None,
    dim_y: float | None = None,
) -> dict:
    """Add a linear dimension to the active draft."""
    return export_manager.add_dimension(x1, y1, x2, y2, dim_x, dim_y)


def add_balloon(
    x: float, y: float, text: str = "", leader_x: float | None = None, leader_y: float | None = None
) -> dict:
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


def add_projected_view(parent_view_index: int, fold_direction: str, x: float, y: float) -> dict:
    """Add a projected (folded) view from a parent view.

    Directions: 'Up', 'Down', 'Left', 'Right'.
    """
    return export_manager.add_projected_view(parent_view_index, fold_direction, x, y)


def move_drawing_view(view_index: int, x: float, y: float) -> dict:
    """Reposition a drawing view on the sheet (meters)."""
    return export_manager.move_drawing_view(view_index, x, y)


def show_hidden_edges(view_index: int, show: bool = True) -> dict:
    """Toggle hidden edge visibility on a drawing view."""
    return export_manager.show_hidden_edges(view_index, show)


def set_drawing_view_display_mode(view_index: int, mode: str) -> dict:
    """Set drawing view render mode.

    Modes: 'Wireframe', 'HiddenEdgesVisible', 'Shaded', 'ShadedWithEdges'.
    """
    return export_manager.set_drawing_view_display_mode(view_index, mode)


def get_drawing_view_info(view_index: int) -> dict:
    """Get detailed info about a drawing view (scale, position, name, properties)."""
    return export_manager.get_drawing_view_info(view_index)


def set_drawing_view_orientation(view_index: int, orientation: str) -> dict:
    """Change drawing view orientation.

    Options: 'Front', 'Top', 'Right', 'Back', 'Bottom', 'Left', 'Isometric'.
    """
    return export_manager.set_drawing_view_orientation(view_index, orientation)


# === Dimension Annotations ===


def add_angular_dimension(
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    x3: float,
    y3: float,
    dim_x: float | None = None,
    dim_y: float | None = None,
) -> dict:
    """Add an angular dimension between three points. Vertex is (x2, y2)."""
    return export_manager.add_angular_dimension(x1, y1, x2, y2, x3, y3, dim_x, dim_y)


def add_radial_dimension(
    center_x: float,
    center_y: float,
    point_x: float,
    point_y: float,
    dim_x: float | None = None,
    dim_y: float | None = None,
) -> dict:
    """Add a radial dimension from center to a point on an arc."""
    return export_manager.add_radial_dimension(center_x, center_y, point_x, point_y, dim_x, dim_y)


def add_diameter_dimension(
    center_x: float,
    center_y: float,
    point_x: float,
    point_y: float,
    dim_x: float | None = None,
    dim_y: float | None = None,
) -> dict:
    """Add a diameter dimension on a circle."""
    return export_manager.add_diameter_dimension(center_x, center_y, point_x, point_y, dim_x, dim_y)


def add_ordinate_dimension(
    origin_x: float,
    origin_y: float,
    point_x: float,
    point_y: float,
    dim_x: float | None = None,
    dim_y: float | None = None,
) -> dict:
    """Add an ordinate dimension from a datum origin to a point."""
    return export_manager.add_ordinate_dimension(origin_x, origin_y, point_x, point_y, dim_x, dim_y)


# === Symbol Annotations ===


def add_center_mark(x: float, y: float) -> dict:
    """Add a center mark annotation at the specified position."""
    return export_manager.add_center_mark(x, y)


def add_centerline(x1: float, y1: float, x2: float, y2: float) -> dict:
    """Add a centerline between two points on the active draft."""
    return export_manager.add_centerline(x1, y1, x2, y2)


def add_surface_finish_symbol(x: float, y: float, symbol_type: str = "machined") -> dict:
    """Add a surface finish symbol. Types: 'machined', 'any', 'prohibited'."""
    return export_manager.add_surface_finish_symbol(x, y, symbol_type)


def add_weld_symbol(x: float, y: float, weld_type: str = "fillet") -> dict:
    """Add a welding symbol. Types: 'fillet', 'groove', 'plug', 'spot', 'seam'."""
    return export_manager.add_weld_symbol(x, y, weld_type)


def add_geometric_tolerance(x: float, y: float, tolerance_text: str = "") -> dict:
    """Add a geometric tolerance (Feature Control Frame / GD&T) annotation."""
    return export_manager.add_geometric_tolerance(x, y, tolerance_text)


# === Drawing View Variants ===


def add_detail_view(
    parent_view_index: int,
    center_x: float,
    center_y: float,
    radius: float,
    x: float,
    y: float,
    scale: float = 2.0,
) -> dict:
    """Add a detail (zoom) view from a parent drawing view."""
    return export_manager.add_detail_view(
        parent_view_index, center_x, center_y, radius, x, y, scale
    )


def add_auxiliary_view(
    parent_view_index: int, x: float, y: float, fold_direction: str = "Up"
) -> dict:
    """Add an auxiliary (folded) view from a parent view.

    Directions: 'Up', 'Down', 'Left', 'Right'.
    """
    return export_manager.add_auxiliary_view(parent_view_index, x, y, fold_direction)


def add_draft_view(x: float, y: float) -> dict:
    """Add an empty draft (sketch) view for annotations."""
    return export_manager.add_draft_view(x, y)


# === Drawing View Properties ===


def align_drawing_views(view_index1: int, view_index2: int, align: bool = True) -> dict:
    """Align or unalign two drawing views."""
    return export_manager.align_drawing_views(view_index1, view_index2, align)


def get_drawing_view_model_link(view_index: int) -> dict:
    """Get the model link reference from a drawing view."""
    return export_manager.get_drawing_view_model_link(view_index)


def show_tangent_edges(view_index: int, show: bool = True) -> dict:
    """Set tangent edge visibility on a drawing view."""
    return export_manager.show_tangent_edges(view_index, show)


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


def rotate_view(
    angle: float,
    center_x: float = 0.0,
    center_y: float = 0.0,
    center_z: float = 0.0,
    axis_x: float = 0.0,
    axis_y: float = 1.0,
    axis_z: float = 0.0,
) -> dict:
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


def set_camera(
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
) -> dict:
    """Set the camera position and orientation."""
    return view_manager.set_camera(
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


def transform_model_to_screen(x: float, y: float, z: float) -> dict:
    """Transform 3D model coordinates to 2D screen pixel coordinates."""
    return view_manager.transform_model_to_screen(x, y, z)


def transform_screen_to_model(screen_x: int, screen_y: int) -> dict:
    """Transform 2D screen pixel coordinates to 3D model coordinates."""
    return view_manager.transform_screen_to_model(screen_x, screen_y)


def begin_camera_dynamics() -> dict:
    """Begin camera dynamics mode for smooth multi-step camera manipulation."""
    return view_manager.begin_camera_dynamics()


def end_camera_dynamics() -> dict:
    """End camera dynamics mode and apply pending camera changes."""
    return view_manager.end_camera_dynamics()


# === Draft Sheet Collections ===


def get_sheet_dimensions() -> dict:
    """Get all dimensions on the active draft sheet."""
    return export_manager.get_sheet_dimensions()


def get_sheet_balloons() -> dict:
    """Get all balloons on the active draft sheet."""
    return export_manager.get_sheet_balloons()


def get_sheet_text_boxes() -> dict:
    """Get all text boxes on the active draft sheet."""
    return export_manager.get_sheet_text_boxes()


def get_sheet_drawing_objects() -> dict:
    """Get all drawing objects on the active draft sheet."""
    return export_manager.get_sheet_drawing_objects()


def get_sheet_sections() -> dict:
    """Get all section views on the active draft sheet."""
    return export_manager.get_sheet_sections()


# === Printing ===


def print_drawing(copies: int = 1, all_sheets: bool = True) -> dict:
    """Print the active draft document."""
    return export_manager.print_drawing(copies, all_sheets)


def set_printer(printer_name: str) -> dict:
    """Set the printer for the active draft document."""
    return export_manager.set_printer(printer_name)


def get_printer() -> dict:
    """Get the current printer for the active draft document."""
    return export_manager.get_printer()


def set_paper_size(width: float, height: float, orientation: str = "Landscape") -> dict:
    """Set the paper size and orientation for printing."""
    return export_manager.set_paper_size(width, height, orientation)


# === Drawing View Tools (Batch 10) ===


def set_face_texture(face_index: int, texture_name: str) -> dict:
    """Apply a texture to a face by 0-based index."""
    return export_manager.set_face_texture(face_index, texture_name)


def add_assembly_drawing_view_ex(
    x: float = 0.15,
    y: float = 0.15,
    orientation: str = "Isometric",
    scale: float = 1.0,
    config: str = None,
) -> dict:
    """Add an extended assembly drawing view with optional configuration."""
    return export_manager.add_assembly_drawing_view_ex(x, y, orientation, scale, config)


def add_drawing_view_with_config(
    x: float = 0.15,
    y: float = 0.15,
    orientation: str = "Front",
    scale: float = 1.0,
    configuration: str = "Default",
) -> dict:
    """Add a drawing view with a specific configuration."""
    return export_manager.add_drawing_view_with_config(x, y, orientation, scale, configuration)


def activate_drawing_view(view_index: int) -> dict:
    """Activate a drawing view by 0-based index for editing."""
    return export_manager.activate_drawing_view(view_index)


def deactivate_drawing_view(view_index: int) -> dict:
    """Deactivate a drawing view by 0-based index."""
    return export_manager.deactivate_drawing_view(view_index)


# === Drawing View Copy / Section / Dimensions (Batch 11) ===


def add_by_draft_view(
    source_view_index: int, x: float, y: float, scale: float | None = None
) -> dict:
    """Copy an existing drawing view to a new sheet location.

    Args:
        source_view_index: 0-based index of the source drawing view.
        x: X position for the new view (meters).
        y: Y position for the new view (meters).
        scale: Optional scale factor (defaults to source view scale).
    """
    return export_manager.add_by_draft_view(source_view_index, x, y, scale)


def get_section_cuts(view_index: int) -> dict:
    """Get section cut (cutting plane) info from a drawing view."""
    return export_manager.get_section_cuts(view_index)


def add_section_cut(view_index: int, x: float, y: float, section_type: int = 0) -> dict:
    """Add a section cut to a drawing view and create the section view.

    Args:
        view_index: 0-based index of the source drawing view.
        x: X position for the section view (meters).
        y: Y position for the section view (meters).
        section_type: 0 = standard, 1 = revolved.
    """
    return export_manager.add_section_cut(view_index, x, y, section_type)


def get_drawing_view_dimensions(view_index: int) -> dict:
    """Get all dimensions associated with a specific drawing view."""
    return export_manager.get_drawing_view_dimensions(view_index)


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
    # Dimension Annotations
    mcp.tool()(add_angular_dimension)
    mcp.tool()(add_radial_dimension)
    mcp.tool()(add_diameter_dimension)
    mcp.tool()(add_ordinate_dimension)
    # Symbol Annotations
    mcp.tool()(add_center_mark)
    mcp.tool()(add_centerline)
    mcp.tool()(add_surface_finish_symbol)
    mcp.tool()(add_weld_symbol)
    mcp.tool()(add_geometric_tolerance)
    # Drawing View Variants
    mcp.tool()(add_detail_view)
    mcp.tool()(add_auxiliary_view)
    mcp.tool()(add_draft_view)
    # Drawing View Properties
    mcp.tool()(align_drawing_views)
    mcp.tool()(get_drawing_view_model_link)
    mcp.tool()(show_tangent_edges)
    # Drawing View Management
    mcp.tool()(get_drawing_view_count)
    mcp.tool()(get_drawing_view_scale)
    mcp.tool()(set_drawing_view_scale)
    mcp.tool()(delete_drawing_view)
    mcp.tool()(update_drawing_view)
    mcp.tool()(add_projected_view)
    mcp.tool()(move_drawing_view)
    mcp.tool()(show_hidden_edges)
    mcp.tool()(set_drawing_view_display_mode)
    mcp.tool()(get_drawing_view_info)
    mcp.tool()(set_drawing_view_orientation)
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
    mcp.tool()(transform_model_to_screen)
    mcp.tool()(transform_screen_to_model)
    mcp.tool()(begin_camera_dynamics)
    mcp.tool()(end_camera_dynamics)
    # Draft Sheet Collections
    mcp.tool()(get_sheet_dimensions)
    mcp.tool()(get_sheet_balloons)
    mcp.tool()(get_sheet_text_boxes)
    mcp.tool()(get_sheet_drawing_objects)
    mcp.tool()(get_sheet_sections)
    # Printing
    mcp.tool()(print_drawing)
    mcp.tool()(set_printer)
    mcp.tool()(get_printer)
    mcp.tool()(set_paper_size)
    # Drawing View Tools (Batch 10)
    mcp.tool()(set_face_texture)
    mcp.tool()(add_assembly_drawing_view_ex)
    mcp.tool()(add_drawing_view_with_config)
    mcp.tool()(activate_drawing_view)
    mcp.tool()(deactivate_drawing_view)
    # Drawing View Copy / Section / Dimensions (Batch 11)
    mcp.tool()(add_by_draft_view)
    mcp.tool()(get_section_cuts)
    mcp.tool()(add_section_cut)
    mcp.tool()(get_drawing_view_dimensions)
