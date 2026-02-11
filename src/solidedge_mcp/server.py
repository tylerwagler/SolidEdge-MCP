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
