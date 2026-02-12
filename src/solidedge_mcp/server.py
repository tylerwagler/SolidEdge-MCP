"""Solid Edge MCP Server"""

from typing import Optional
from fastmcp import FastMCP

# Import backend managers
from solidedge_mcp.backends.connection import SolidEdgeConnection
from solidedge_mcp.backends.documents import DocumentManager
from solidedge_mcp.backends.sketching import SketchManager
from solidedge_mcp.backends.features import FeatureManager
from solidedge_mcp.backends.assembly import AssemblyManager
from solidedge_mcp.backends.query import QueryManager
from solidedge_mcp.backends.export import ExportManager, ViewModel
from solidedge_mcp.backends.diagnostics import diagnose_document, diagnose_feature

# Create the FastMCP server
mcp = FastMCP("Solid Edge MCP Server")

# Initialize managers (global state)
connection = SolidEdgeConnection()
doc_manager = DocumentManager(connection)
sketch_manager = SketchManager(doc_manager)
feature_manager = FeatureManager(doc_manager, sketch_manager)
assembly_manager = AssemblyManager(doc_manager)
query_manager = QueryManager(doc_manager)
export_manager = ExportManager(doc_manager)
view_manager = ViewModel(doc_manager)


# ============================================================================
# CONNECTION TOOLS
# ============================================================================

@mcp.tool()
def connect_to_solidedge(start_if_needed: bool = True) -> dict:
    """
    Connect to Solid Edge application (start if needed).

    Args:
        start_if_needed: If True, start Solid Edge if not running

    Returns:
        Connection status with version info
    """
    return connection.connect(start_if_needed)


@mcp.tool()
def get_application_info() -> dict:
    """
    Get Solid Edge application information.

    Returns:
        Application version, path, and document count
    """
    return connection.get_info()


# ============================================================================
# DOCUMENT MANAGEMENT TOOLS
# ============================================================================

@mcp.tool()
def create_part_document(template: Optional[str] = None) -> dict:
    """
    Create a new part document.

    Args:
        template: Optional template file path

    Returns:
        Document creation status
    """
    return doc_manager.create_part(template)


@mcp.tool()
def create_assembly_document(template: Optional[str] = None) -> dict:
    """
    Create a new assembly document.

    Args:
        template: Optional template file path

    Returns:
        Document creation status
    """
    return doc_manager.create_assembly(template)


@mcp.tool()
def open_document(file_path: str) -> dict:
    """
    Open an existing document.

    Args:
        file_path: Path to the document file (.par, .asm, .dft)

    Returns:
        Open status and document info
    """
    return doc_manager.open_document(file_path)


@mcp.tool()
def save_document(file_path: Optional[str] = None) -> dict:
    """
    Save the active document.

    Args:
        file_path: Optional save path (if not provided, saves to current location)

    Returns:
        Save status
    """
    return doc_manager.save_document(file_path)


@mcp.tool()
def close_document(save: bool = True) -> dict:
    """
    Close the active document.

    Args:
        save: Whether to save before closing (default: True)

    Returns:
        Close status
    """
    return doc_manager.close_document(save)


@mcp.tool()
def list_documents() -> dict:
    """
    List all open documents.

    Returns:
        List of open documents with their info
    """
    return doc_manager.list_documents()


# ============================================================================
# SKETCHING TOOLS
# ============================================================================

@mcp.tool()
def create_sketch(plane: str = "Top") -> dict:
    """
    Create a new 2D sketch on a reference plane.

    Args:
        plane: Plane name - 'Top', 'Front', 'Right', 'XY', 'XZ', 'YZ'

    Returns:
        Sketch creation status
    """
    return sketch_manager.create_sketch(plane)


@mcp.tool()
def draw_line(x1: float, y1: float, x2: float, y2: float) -> dict:
    """
    Draw a line in the active sketch.

    Args:
        x1, y1: Start point coordinates (meters)
        x2, y2: End point coordinates (meters)

    Returns:
        Line creation status
    """
    return sketch_manager.draw_line(x1, y1, x2, y2)


@mcp.tool()
def draw_circle(center_x: float, center_y: float, radius: float) -> dict:
    """
    Draw a circle in the active sketch.

    Args:
        center_x, center_y: Center point coordinates (meters)
        radius: Circle radius (meters)

    Returns:
        Circle creation status
    """
    return sketch_manager.draw_circle(center_x, center_y, radius)


@mcp.tool()
def draw_rectangle(x1: float, y1: float, x2: float, y2: float) -> dict:
    """
    Draw a rectangle in the active sketch.

    Args:
        x1, y1: First corner coordinates (meters)
        x2, y2: Opposite corner coordinates (meters)

    Returns:
        Rectangle creation status
    """
    return sketch_manager.draw_rectangle(x1, y1, x2, y2)


@mcp.tool()
def draw_arc(
    center_x: float,
    center_y: float,
    radius: float,
    start_angle: float,
    end_angle: float
) -> dict:
    """
    Draw an arc in the active sketch.

    Args:
        center_x, center_y: Arc center coordinates (meters)
        radius: Arc radius (meters)
        start_angle: Start angle in degrees (0 = right, 90 = up)
        end_angle: End angle in degrees

    Returns:
        Arc creation status
    """
    return sketch_manager.draw_arc(center_x, center_y, radius, start_angle, end_angle)


@mcp.tool()
def draw_polygon(center_x: float, center_y: float, radius: float, sides: int) -> dict:
    """
    Draw a regular polygon in the active sketch.

    Args:
        center_x, center_y: Polygon center coordinates (meters)
        radius: Distance from center to vertices (meters)
        sides: Number of sides (minimum 3)

    Returns:
        Polygon creation status
    """
    return sketch_manager.draw_polygon(center_x, center_y, radius, sides)


@mcp.tool()
def draw_ellipse(
    center_x: float,
    center_y: float,
    major_radius: float,
    minor_radius: float,
    angle: float = 0.0
) -> dict:
    """
    Draw an ellipse in the active sketch.

    Args:
        center_x, center_y: Ellipse center coordinates (meters)
        major_radius: Major axis radius (meters)
        minor_radius: Minor axis radius (meters)
        angle: Rotation angle in degrees (default 0)

    Returns:
        Ellipse creation status
    """
    return sketch_manager.draw_ellipse(center_x, center_y, major_radius, minor_radius, angle)


@mcp.tool()
def draw_spline(points: list) -> dict:
    """
    Draw a B-spline curve through a list of points.

    Args:
        points: List of [x, y] coordinate pairs, e.g., [[0, 0], [1, 1], [2, 0]]

    Returns:
        Spline creation status
    """
    return sketch_manager.draw_spline(points)


@mcp.tool()
def add_constraint(constraint_type: str, elements: list) -> dict:
    """
    Add a geometric constraint to sketch elements.

    Args:
        constraint_type: Type of constraint (e.g., 'Horizontal', 'Vertical', 'Parallel', 'Equal')
        elements: List of sketch element references

    Returns:
        Constraint creation status
    """
    return sketch_manager.add_constraint(constraint_type, elements)


@mcp.tool()
def set_axis_of_revolution(x1: float, y1: float, x2: float, y2: float) -> dict:
    """
    Draw an axis of revolution in the active sketch for revolve operations.

    Must be called before close_sketch() when preparing a profile for revolve.
    Draws a construction line and sets it as the revolution axis.

    Args:
        x1: Axis start X coordinate (meters)
        y1: Axis start Y coordinate (meters)
        x2: Axis end X coordinate (meters)
        y2: Axis end Y coordinate (meters)

    Returns:
        Axis creation status
    """
    return sketch_manager.set_axis_of_revolution(x1, y1, x2, y2)


@mcp.tool()
def create_sketch_on_plane(plane_index: int) -> dict:
    """
    Create a new 2D sketch on a reference plane by index.

    Useful for sketching on user-created offset planes (index > 3).
    Default planes: 1=Top/XZ, 2=Front/XY, 3=Right/YZ.

    Args:
        plane_index: 1-based index of the reference plane

    Returns:
        Sketch creation status
    """
    return sketch_manager.create_sketch_on_plane_index(plane_index)


@mcp.tool()
def close_sketch() -> dict:
    """
    Close/finish the active sketch.

    Returns:
        Sketch closure status
    """
    return sketch_manager.close_sketch()


# ============================================================================
# 3D FEATURE TOOLS
# ============================================================================

@mcp.tool()
def create_extrude(distance: float, operation: str = "Add", direction: str = "Normal") -> dict:
    """
    Create an extrusion feature from the active sketch.

    Args:
        distance: Extrusion distance in meters (use negative for reverse)
        operation: 'Add' (protrusion) - Note: 'Cut' is NOT available in COM API
        direction: 'Normal', 'Reverse', or 'Symmetric'

    Returns:
        Extrusion creation status
    """
    return feature_manager.create_extrude(distance, operation, direction)


@mcp.tool()
def create_revolve(angle: float, operation: str = "Add") -> dict:
    """
    Create a revolve feature from the active sketch.

    Args:
        angle: Rotation angle in degrees (360 for full revolution)
        operation: 'Add' (protrusion) - Note: 'Cut' is NOT available in COM API

    Returns:
        Revolve creation status
    """
    return feature_manager.create_revolve(angle, operation)


@mcp.tool()
def create_extrude_thin_wall(distance: float, wall_thickness: float, direction: str = "Normal") -> dict:
    """
    Create a thin-walled extrusion feature.

    Args:
        distance: Extrusion distance in meters
        wall_thickness: Wall thickness in meters
        direction: 'Normal', 'Reverse', or 'Symmetric'

    Returns:
        Thin-wall extrusion creation status
    """
    return feature_manager.create_extrude_thin_wall(distance, wall_thickness, direction)


@mcp.tool()
def create_extrude_infinite(direction: str = "Normal") -> dict:
    """
    Create an infinite extrusion (extends through all).

    Args:
        direction: 'Normal', 'Reverse', or 'Symmetric'

    Returns:
        Infinite extrusion creation status
    """
    return feature_manager.create_extrude_infinite(direction)


@mcp.tool()
def create_revolve_finite(angle: float, axis_type: str = "CenterLine") -> dict:
    """
    Create a finite revolve feature.

    Args:
        angle: Revolution angle in degrees
        axis_type: Type of revolution axis (default: 'CenterLine')

    Returns:
        Finite revolve creation status
    """
    return feature_manager.create_revolve_finite(angle, axis_type)


@mcp.tool()
def create_revolve_thin_wall(angle: float, wall_thickness: float) -> dict:
    """
    Create a thin-walled revolve feature.

    Args:
        angle: Revolution angle in degrees
        wall_thickness: Wall thickness in meters

    Returns:
        Thin-wall revolve creation status
    """
    return feature_manager.create_revolve_thin_wall(angle, wall_thickness)


@mcp.tool()
def create_revolve_sync(angle: float) -> dict:
    """
    Create a synchronous revolve feature.

    Args:
        angle: Revolution angle in degrees

    Returns:
        Synchronous revolve creation status
    """
    return feature_manager.create_revolve_sync(angle)


@mcp.tool()
def create_revolve_finite_sync(angle: float) -> dict:
    """
    Create a finite synchronous revolve feature.

    Args:
        angle: Revolution angle in degrees

    Returns:
        Finite synchronous revolve creation status
    """
    return feature_manager.create_revolve_finite_sync(angle)


# ============================================================================
# CUTOUT (CUT) OPERATIONS
# ============================================================================

@mcp.tool()
def create_extruded_cutout(distance: float, direction: str = "Normal") -> dict:
    """
    Create an extruded cutout (cut) through the part.

    Removes material by extruding the active sketch profile into the part.
    Requires an existing base feature and a closed sketch profile.

    Args:
        distance: Cutout depth in meters
        direction: 'Normal' or 'Reverse'

    Returns:
        Cutout creation status
    """
    return feature_manager.create_extruded_cutout(distance, direction)


@mcp.tool()
def create_extruded_cutout_through_all(direction: str = "Normal") -> dict:
    """
    Create an extruded cutout that goes through the entire part.

    Removes material by extruding the active sketch profile through the whole body.
    Requires an existing base feature and a closed sketch profile.

    Args:
        direction: 'Normal' or 'Reverse'

    Returns:
        Cutout creation status
    """
    return feature_manager.create_extruded_cutout_through_all(direction)


@mcp.tool()
def create_revolved_cutout(angle: float = 360) -> dict:
    """
    Create a revolved cutout (cut) in the part.

    Removes material by revolving the active sketch profile around an axis.
    Requires an existing base feature, a closed sketch profile, and an axis of revolution.

    Args:
        angle: Revolution angle in degrees (360 for full revolution)

    Returns:
        Cutout creation status
    """
    return feature_manager.create_revolved_cutout(angle)


# ============================================================================
# REFERENCE PLANE CREATION
# ============================================================================

@mcp.tool()
def create_ref_plane_by_offset(
    parent_plane_index: int,
    distance: float,
    normal_side: str = "Normal"
) -> dict:
    """
    Create a reference plane parallel to an existing plane at an offset distance.

    Useful for creating sketches at different heights/positions.
    Default planes: 1=Top/XZ, 2=Front/XY, 3=Right/YZ.

    Args:
        parent_plane_index: Index of parent plane (1-based)
        distance: Offset distance in meters
        normal_side: 'Normal' or 'Reverse' for offset direction

    Returns:
        Reference plane creation status with new plane index
    """
    return feature_manager.create_ref_plane_by_offset(parent_plane_index, distance, normal_side)


# ============================================================================
# PRIMITIVE SHAPES
# ============================================================================

@mcp.tool()
def create_box_by_center(
    center_x: float,
    center_y: float,
    center_z: float,
    length: float,
    width: float,
    height: float
) -> dict:
    """
    Create a box primitive by center point and dimensions.

    Args:
        center_x, center_y, center_z: Center point coordinates (meters)
        length: Length in meters (X direction)
        width: Width in meters (Y direction)
        height: Height in meters (Z direction)

    Returns:
        Box creation status
    """
    return feature_manager.create_box_by_center(
        center_x, center_y, center_z, length, width, height
    )


@mcp.tool()
def create_box_by_two_points(
    x1: float, y1: float, z1: float,
    x2: float, y2: float, z2: float
) -> dict:
    """
    Create a box primitive by two opposite corners.

    Args:
        x1, y1, z1: First corner coordinates (meters)
        x2, y2, z2: Opposite corner coordinates (meters)

    Returns:
        Box creation status
    """
    return feature_manager.create_box_by_two_points(x1, y1, z1, x2, y2, z2)


@mcp.tool()
def create_cylinder(
    base_center_x: float,
    base_center_y: float,
    base_center_z: float,
    radius: float,
    height: float
) -> dict:
    """
    Create a cylinder primitive.

    Args:
        base_center_x, base_center_y, base_center_z: Base circle center (meters)
        radius: Cylinder radius (meters)
        height: Cylinder height (meters)

    Returns:
        Cylinder creation status
    """
    return feature_manager.create_cylinder(
        base_center_x, base_center_y, base_center_z, radius, height
    )


@mcp.tool()
def create_box_by_three_points(
    x1: float, y1: float, z1: float,
    x2: float, y2: float, z2: float,
    x3: float, y3: float, z3: float
) -> dict:
    """
    Create a box primitive by three points.

    Args:
        x1, y1, z1: First corner point (meters)
        x2, y2, z2: Second point defining width (meters)
        x3, y3, z3: Third point defining height (meters)

    Returns:
        Box creation status
    """
    return feature_manager.create_box_by_three_points(x1, y1, z1, x2, y2, z2, x3, y3, z3)


@mcp.tool()
def create_sphere(
    center_x: float,
    center_y: float,
    center_z: float,
    radius: float
) -> dict:
    """
    Create a sphere primitive.

    Args:
        center_x, center_y, center_z: Sphere center coordinates (meters)
        radius: Sphere radius (meters)

    Returns:
        Sphere creation status
    """
    return feature_manager.create_sphere(center_x, center_y, center_z, radius)


# ============================================================================
# ADVANCED 3D FEATURES
# ============================================================================

@mcp.tool()
def create_loft(profile_indices: Optional[list] = None) -> dict:
    """
    Create a loft feature between multiple profiles.

    Workflow: Create 2+ sketches on different parallel planes, close each one,
    then call create_loft(). Each close_sketch() accumulates the profile.

    Args:
        profile_indices: List of profile indices to loft between (optional).
            If None, uses all accumulated profiles.

    Returns:
        Loft creation status

    Note: Requires multiple closed profiles to be created first
    """
    return feature_manager.create_loft(profile_indices)


@mcp.tool()
def create_sweep(path_profile_index: Optional[int] = None) -> dict:
    """
    Create a sweep feature along a path.

    Workflow: Create a path sketch (open profile, e.g. a line) on one plane,
    close it, then create a cross-section sketch (closed profile, e.g. a circle)
    on a perpendicular plane, close it, then call create_sweep().

    Args:
        path_profile_index: Index of the path profile in accumulated profiles
            (default: 0, the first accumulated profile)

    Returns:
        Sweep creation status

    Note: Requires a cross-section profile and a path curve
    """
    return feature_manager.create_sweep(path_profile_index)


@mcp.tool()
def create_loft_thin_wall(wall_thickness: float, profile_indices: Optional[list] = None) -> dict:
    """
    Create a thin-walled loft feature.

    Args:
        wall_thickness: Wall thickness in meters
        profile_indices: List of profile indices (optional)

    Returns:
        Thin-wall loft creation status
    """
    return feature_manager.create_loft_thin_wall(wall_thickness, profile_indices)


@mcp.tool()
def create_sweep_thin_wall(wall_thickness: float, path_profile_index: Optional[int] = None) -> dict:
    """
    Create a thin-walled sweep feature.

    Args:
        wall_thickness: Wall thickness in meters
        path_profile_index: Index of path profile (optional)

    Returns:
        Thin-wall sweep creation status
    """
    return feature_manager.create_sweep_thin_wall(wall_thickness, path_profile_index)


# ============================================================================
# HELIX & SPIRAL FEATURES
# ============================================================================

@mcp.tool()
def create_helix(
    pitch: float,
    height: float,
    revolutions: Optional[float] = None,
    direction: str = "Right"
) -> dict:
    """
    Create a helical feature (springs, threads, etc.).

    Args:
        pitch: Distance between coils in meters
        height: Total height of helix in meters
        revolutions: Number of turns (optional, calculated from pitch/height if not given)
        direction: 'Right' or 'Left' hand helix

    Returns:
        Helix creation status
    """
    return feature_manager.create_helix(pitch, height, revolutions, direction)


@mcp.tool()
def create_helix_sync(
    pitch: float,
    height: float,
    revolutions: Optional[float] = None
) -> dict:
    """
    Create a synchronous helical feature.

    Args:
        pitch: Distance between coils in meters
        height: Total height of helix in meters
        revolutions: Number of turns (optional)

    Returns:
        Synchronous helix creation status
    """
    return feature_manager.create_helix_sync(pitch, height, revolutions)


@mcp.tool()
def create_helix_thin_wall(
    pitch: float,
    height: float,
    wall_thickness: float,
    revolutions: Optional[float] = None
) -> dict:
    """
    Create a thin-walled helical feature.

    Args:
        pitch: Distance between coils in meters
        height: Total height of helix in meters
        wall_thickness: Wall thickness in meters
        revolutions: Number of turns (optional)

    Returns:
        Thin-wall helix creation status
    """
    return feature_manager.create_helix_thin_wall(pitch, height, wall_thickness, revolutions)


@mcp.tool()
def create_helix_sync_thin_wall(
    pitch: float,
    height: float,
    wall_thickness: float,
    revolutions: Optional[float] = None
) -> dict:
    """
    Create a synchronous thin-walled helical feature.

    Args:
        pitch: Distance between coils in meters
        height: Total height of helix in meters
        wall_thickness: Wall thickness in meters
        revolutions: Number of turns (optional)

    Returns:
        Synchronous thin-wall helix creation status
    """
    return feature_manager.create_helix_sync_thin_wall(pitch, height, wall_thickness, revolutions)


# ============================================================================
# SHEET METAL FEATURES
# ============================================================================

@mcp.tool()
def create_base_flange(
    width: float,
    thickness: float,
    bend_radius: Optional[float] = None
) -> dict:
    """
    Create a base contour flange (sheet metal).

    Args:
        width: Flange width in meters
        thickness: Material thickness in meters
        bend_radius: Bend radius in meters (optional, defaults to 2x thickness)

    Returns:
        Base flange creation status
    """
    return feature_manager.create_base_flange(width, thickness, bend_radius)


@mcp.tool()
def create_base_tab(thickness: float, width: Optional[float] = None) -> dict:
    """
    Create a base tab (sheet metal).

    Args:
        thickness: Material thickness in meters
        width: Tab width in meters (optional)

    Returns:
        Base tab creation status
    """
    return feature_manager.create_base_tab(thickness, width)


@mcp.tool()
def create_lofted_flange(thickness: float) -> dict:
    """
    Create a lofted flange (sheet metal).

    Args:
        thickness: Material thickness in meters

    Returns:
        Lofted flange creation status
    """
    return feature_manager.create_lofted_flange(thickness)


@mcp.tool()
def create_web_network() -> dict:
    """
    Create a web network (sheet metal).

    Returns:
        Web network creation status
    """
    return feature_manager.create_web_network()


@mcp.tool()
def create_base_contour_flange_advanced(
    thickness: float,
    bend_radius: float,
    relief_type: str = "Default"
) -> dict:
    """
    Create base contour flange with bend deduction or bend allowance (advanced sheet metal).

    Args:
        thickness: Material thickness (meters)
        bend_radius: Bend radius (meters)
        relief_type: Relief type - 'Default', 'Rectangular', 'Obround'

    Returns:
        Base contour flange creation status
    """
    return feature_manager.create_base_contour_flange_advanced(thickness, bend_radius, relief_type)


@mcp.tool()
def create_base_tab_multi_profile(thickness: float) -> dict:
    """
    Create base tab with multiple profiles (sheet metal).

    Args:
        thickness: Material thickness (meters)

    Returns:
        Base tab creation status
    """
    return feature_manager.create_base_tab_multi_profile(thickness)


@mcp.tool()
def create_lofted_flange_advanced(thickness: float, bend_radius: float) -> dict:
    """
    Create lofted flange with bend deduction or bend allowance (advanced sheet metal).

    Args:
        thickness: Material thickness (meters)
        bend_radius: Bend radius (meters)

    Returns:
        Lofted flange creation status
    """
    return feature_manager.create_lofted_flange_advanced(thickness, bend_radius)


@mcp.tool()
def create_lofted_flange_ex(thickness: float) -> dict:
    """
    Create extended lofted flange (sheet metal).

    Args:
        thickness: Material thickness (meters)

    Returns:
        Lofted flange creation status
    """
    return feature_manager.create_lofted_flange_ex(thickness)


# ============================================================================
# BODY OPERATIONS
# ============================================================================

@mcp.tool()
def add_body(body_type: str = "Solid") -> dict:
    """
    Add a body to the part.

    Args:
        body_type: Type of body - 'Solid', 'Surface', 'Construction'

    Returns:
        Body creation status
    """
    return feature_manager.add_body(body_type)


@mcp.tool()
def thicken_surface(thickness: float, direction: str = "Both") -> dict:
    """
    Thicken a surface to create a solid.

    Args:
        thickness: Thickness in meters
        direction: 'Both', 'Inside', or 'Outside'

    Returns:
        Thicken operation status
    """
    return feature_manager.thicken_surface(thickness, direction)


@mcp.tool()
def add_body_by_mesh() -> dict:
    """
    Add a body by mesh facets.

    Returns:
        Body by mesh creation status
    """
    return feature_manager.add_body_by_mesh()


@mcp.tool()
def add_body_feature() -> dict:
    """
    Add a body feature.

    Returns:
        Body feature creation status
    """
    return feature_manager.add_body_feature()


@mcp.tool()
def add_by_construction() -> dict:
    """
    Add a construction body.

    Returns:
        Construction body creation status
    """
    return feature_manager.add_by_construction()


@mcp.tool()
def add_body_by_tag(tag: str) -> dict:
    """
    Add a body by tag reference.

    Args:
        tag: Tag identifier for the body

    Returns:
        Body by tag creation status
    """
    return feature_manager.add_body_by_tag(tag)


# ============================================================================
# SIMPLIFICATION FEATURES
# ============================================================================

@mcp.tool()
def auto_simplify() -> dict:
    """
    Auto-simplify the model (reduce complexity).

    Returns:
        Simplification status
    """
    return feature_manager.auto_simplify()


@mcp.tool()
def simplify_enclosure() -> dict:
    """
    Create a simplified enclosure around the model.

    Returns:
        Simplification status
    """
    return feature_manager.simplify_enclosure()


@mcp.tool()
def simplify_duplicate() -> dict:
    """
    Create a simplified duplicate of the model.

    Returns:
        Simplification status
    """
    return feature_manager.simplify_duplicate()


@mcp.tool()
def local_simplify_enclosure() -> dict:
    """
    Create a local simplified enclosure.

    Returns:
        Simplification status
    """
    return feature_manager.local_simplify_enclosure()


# ============================================================================
# QUERY TOOLS
# ============================================================================

@mcp.tool()
def get_mass_properties() -> dict:
    """
    Get mass properties of the active part.

    Returns:
        Mass, volume, center of gravity, and moments of inertia
    """
    return query_manager.get_mass_properties()


@mcp.tool()
def get_bounding_box() -> dict:
    """
    Get the bounding box of the active part.

    Returns:
        Min and max coordinates in X, Y, Z
    """
    return query_manager.get_bounding_box()


@mcp.tool()
def list_features() -> dict:
    """
    List all features in the active document.

    Returns:
        List of features with their properties
    """
    return query_manager.list_features()


@mcp.tool()
def measure_distance(
    x1: float, y1: float, z1: float,
    x2: float, y2: float, z2: float
) -> dict:
    """
    Measure distance between two points.

    Args:
        x1, y1, z1: First point coordinates (meters)
        x2, y2, z2: Second point coordinates (meters)

    Returns:
        Distance measurement
    """
    return query_manager.measure_distance(x1, y1, z1, x2, y2, z2)


@mcp.tool()
def get_document_properties() -> dict:
    """
    Get properties of the active document.

    Returns:
        Document properties (title, author, creation date, etc.)
    """
    return query_manager.get_document_properties()


@mcp.tool()
def get_feature_count() -> dict:
    """
    Get the count of features in the active document.

    Returns:
        Feature count
    """
    return query_manager.get_feature_count()


# ============================================================================
# EXPORT TOOLS
# ============================================================================

@mcp.tool()
def export_step(file_path: str) -> dict:
    """
    Export the active document to STEP format.

    Args:
        file_path: Output file path (.step or .stp)

    Returns:
        Export status
    """
    return export_manager.export_step(file_path)


@mcp.tool()
def export_stl(file_path: str) -> dict:
    """
    Export the active document to STL format.

    Args:
        file_path: Output file path (.stl)

    Returns:
        Export status
    """
    return export_manager.export_stl(file_path)


@mcp.tool()
def export_iges(file_path: str) -> dict:
    """
    Export the active document to IGES format.

    Args:
        file_path: Output file path (.iges or .igs)

    Returns:
        Export status
    """
    return export_manager.export_iges(file_path)


@mcp.tool()
def export_pdf(file_path: str) -> dict:
    """
    Export the active document to PDF format.

    Args:
        file_path: Output file path (.pdf)

    Returns:
        Export status
    """
    return export_manager.export_pdf(file_path)


@mcp.tool()
def create_drawing(template: Optional[str] = None, views: Optional[list] = None) -> dict:
    """
    Create a 2D drawing from the active 3D model.

    Args:
        template: Drawing template path (optional)
        views: List of views to create - ['Front', 'Top', 'Right', 'Isometric'] (optional)

    Returns:
        Drawing creation status

    Note: Drawing views require manual placement for detailed control
    """
    return export_manager.create_drawing(template, views)


@mcp.tool()
def capture_screenshot(file_path: str, width: int = 1920, height: int = 1080) -> dict:
    """
    Capture a screenshot of the current view.

    Args:
        file_path: Output image file path (.png, .jpg, .jpeg, .bmp)
        width: Image width in pixels (default: 1920)
        height: Image height in pixels (default: 1080)

    Returns:
        Screenshot capture status and file info
    """
    return export_manager.capture_screenshot(file_path, width, height)


@mcp.tool()
def export_dxf(file_path: str) -> dict:
    """
    Export the active document to DXF format.

    Args:
        file_path: Output file path (.dxf)

    Returns:
        Export status
    """
    return export_manager.export_to_dxf(file_path)


@mcp.tool()
def export_parasolid(file_path: str) -> dict:
    """
    Export the active document to Parasolid format.

    Args:
        file_path: Output file path (.x_t or .x_b)

    Returns:
        Export status
    """
    return export_manager.export_to_parasolid(file_path)


@mcp.tool()
def export_jt(file_path: str) -> dict:
    """
    Export the active document to JT format.

    Args:
        file_path: Output file path (.jt)

    Returns:
        Export status
    """
    return export_manager.export_to_jt(file_path)


# ============================================================================
# VIEW & DISPLAY TOOLS
# ============================================================================

@mcp.tool()
def set_view(view: str) -> dict:
    """
    Set the viewing orientation.

    Args:
        view: View orientation - 'Iso', 'Top', 'Front', 'Right', 'Left', 'Back', 'Bottom'

    Returns:
        View change status
    """
    return view_manager.set_view(view)


@mcp.tool()
def zoom_fit() -> dict:
    """
    Zoom to fit all geometry in the view.

    Returns:
        Zoom status
    """
    return view_manager.zoom_fit()


@mcp.tool()
def zoom_to_selection() -> dict:
    """
    Zoom to the currently selected geometry.

    Returns:
        Zoom status
    """
    return view_manager.zoom_to_selection()


@mcp.tool()
def set_display_mode(mode: str) -> dict:
    """
    Set the display mode for the active view.

    Args:
        mode: Display mode - 'Shaded', 'ShadedWithEdges', 'Wireframe', 'HiddenEdgesVisible'

    Returns:
        Display mode setting status
    """
    return view_manager.set_display_mode(mode)


# ============================================================================
# ASSEMBLY TOOLS
# ============================================================================

@mcp.tool()
def place_component(component_path: str, x: float = 0.0, y: float = 0.0, z: float = 0.0) -> dict:
    """
    Place a component in the active assembly.

    Args:
        component_path: Path to the component file (.par or .asm)
        x, y, z: Position coordinates in meters

    Returns:
        Component placement status
    """
    return assembly_manager.place_component(component_path, x, y, z)


@mcp.tool()
def list_assembly_components() -> dict:
    """
    List all components in the active assembly.

    Returns:
        List of components with their properties
    """
    return assembly_manager.list_components()


@mcp.tool()
def create_mate(mate_type: str, component1_index: int, component2_index: int) -> dict:
    """
    Create a mate/assembly relationship between components.

    Args:
        mate_type: Type of mate - 'Planar', 'Axial', 'Insert', 'Match', 'Parallel', 'Angle'
        component1_index: Index of first component (0-based)
        component2_index: Index of second component (0-based)

    Returns:
        Mate creation status

    Note: Actual mate creation requires face/edge selection
    """
    return assembly_manager.create_mate(mate_type, component1_index, component2_index)


@mcp.tool()
def get_component_info(component_index: int) -> dict:
    """
    Get detailed information about a specific component in assembly.

    Args:
        component_index: Index of the component (0-based)

    Returns:
        Component information (name, file path, visibility, etc.)
    """
    return assembly_manager.get_component_info(component_index)


@mcp.tool()
def update_component_position(component_index: int, x: float, y: float, z: float) -> dict:
    """
    Update a component's position in the assembly.

    Args:
        component_index: Index of the component (0-based)
        x, y, z: New position coordinates (meters)

    Returns:
        Position update status

    Note: May require adjusting assembly relationships
    """
    return assembly_manager.update_component_position(component_index, x, y, z)


@mcp.tool()
def add_align_constraint(component1_index: int, component2_index: int) -> dict:
    """
    Add an align constraint between two components.

    Args:
        component1_index: Index of first component (0-based)
        component2_index: Index of second component (0-based)

    Returns:
        Constraint creation status
    """
    return assembly_manager.add_align_constraint(component1_index, component2_index)


@mcp.tool()
def add_angle_constraint(component1_index: int, component2_index: int, angle: float) -> dict:
    """
    Add an angle constraint between two components.

    Args:
        component1_index: Index of first component (0-based)
        component2_index: Index of second component (0-based)
        angle: Angle in degrees

    Returns:
        Constraint creation status
    """
    return assembly_manager.add_angle_constraint(component1_index, component2_index, angle)


@mcp.tool()
def add_planar_align_constraint(component1_index: int, component2_index: int) -> dict:
    """
    Add a planar align constraint between two components.

    Args:
        component1_index: Index of first component (0-based)
        component2_index: Index of second component (0-based)

    Returns:
        Constraint creation status
    """
    return assembly_manager.add_planar_align_constraint(component1_index, component2_index)


@mcp.tool()
def add_axial_align_constraint(component1_index: int, component2_index: int) -> dict:
    """
    Add an axial align constraint between two components.

    Args:
        component1_index: Index of first component (0-based)
        component2_index: Index of second component (0-based)

    Returns:
        Constraint creation status
    """
    return assembly_manager.add_axial_align_constraint(component1_index, component2_index)


@mcp.tool()
def pattern_component(component_index: int, count: int, spacing: float, direction: str = "X") -> dict:
    """
    Create a pattern of a component in the assembly.

    Args:
        component_index: Index of the component to pattern (0-based)
        count: Number of instances in the pattern
        spacing: Distance between instances (meters)
        direction: Pattern direction - 'X', 'Y', or 'Z'

    Returns:
        Pattern creation status
    """
    return assembly_manager.pattern_component(component_index, count, spacing, direction)


@mcp.tool()
def suppress_component(component_index: int, suppress: bool = True) -> dict:
    """
    Suppress or unsuppress a component in the assembly.

    Args:
        component_index: Index of the component (0-based)
        suppress: True to suppress, False to unsuppress

    Returns:
        Component suppression status
    """
    return assembly_manager.suppress_component(component_index, suppress)


# ============================================================================
# DIAGNOSTIC TOOLS
# ============================================================================

@mcp.tool()
def diagnose_api() -> dict:
    """
    Diagnose available Solid Edge API methods for debugging.

    Returns:
        Available collections and their Add methods
    """
    try:
        doc = doc_manager.get_active_document()
        return diagnose_document(doc)
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def diagnose_feature(feature_index: int) -> dict:
    """
    Diagnose properties and methods on a specific feature.

    Args:
        feature_index: Index of the feature to inspect (0-based)

    Returns:
        Feature properties and available methods
    """
    try:
        doc = doc_manager.get_active_document()
        models = doc.Models
        if feature_index < 0 or feature_index >= models.Count:
            return {
                "error": f"Invalid feature index: {feature_index}. "
                        f"Document has {models.Count} features."
            }
        model = models.Item(feature_index + 1)  # COM is 1-based
        return diagnose_feature(model)
    except Exception as e:
        return {"error": str(e)}


def main():
    """Entry point for the MCP server"""
    mcp.run()


if __name__ == "__main__":
    main()
