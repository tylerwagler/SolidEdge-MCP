"""MCP Resources — read-only data endpoints for Solid Edge.

52 read-only endpoints (37 static + 15 templates) exposed as MCP Resources
rather than tools, keeping the LLM action space to things that change the model.
Every resource returns a JSON string.
"""

import json
from typing import Any

from solidedge_mcp.managers import (
    connection,
    doc_manager,
    export_manager,
    feature_manager,
    query_manager,
    sketch_manager,
    view_manager,
)
from solidedge_mcp.tools._registry import register_resource

RESOURCE_TAGS = {"query"}


# ===================================================================
# Tier 1: Static resources (no parameters) — 37
# ===================================================================

# --- Application (4) ---


def app_info() -> str:
    """Solid Edge application information (version, path, document count)."""
    return json.dumps(connection.get_info())


def app_install() -> str:
    """Solid Edge installation information (path, language, version)."""
    return json.dumps(connection.get_install_info())


def app_process() -> str:
    """Solid Edge process information (PID, window handle)."""
    return json.dumps(connection.get_process_info())


def app_connection_status() -> str:
    """Whether Solid Edge is currently connected."""
    return json.dumps({"connected": connection.is_connected()})


# --- Document (3) ---


def document_list() -> str:
    """List of all open documents."""
    return json.dumps(doc_manager.list_documents())


def document_active_type() -> str:
    """Type of the currently active document."""
    return json.dumps(doc_manager.get_active_document_type())


def document_count() -> str:
    """Count of open documents."""
    return json.dumps(doc_manager.get_document_count())


# --- Model (11) ---


def model_features() -> str:
    """All features in the active model, in tree order."""
    return json.dumps(feature_manager.list_features())


def model_ref_planes() -> str:
    """Reference planes (1-based: 1=Top/XY, 2=Right/YZ, 3=Front/XZ, 4+ user)."""
    return json.dumps(query_manager.get_ref_planes())


def model_variables() -> str:
    """All variables (dimensions, parameters) in the document."""
    return json.dumps(query_manager.get_variables())


def model_custom_properties() -> str:
    """All custom properties."""
    return json.dumps(query_manager.get_custom_properties())


def model_document_properties() -> str:
    """Document properties (Title, Subject, Author, etc.)."""
    return json.dumps(query_manager.get_document_properties())


def model_layers() -> str:
    """All layers in the active document."""
    return json.dumps(query_manager.get_layers())


def model_mode() -> str:
    """Current modeling mode (Ordered vs Synchronous)."""
    return json.dumps(query_manager.get_modeling_mode())


def model_select_set() -> str:
    """Current selection set."""
    return json.dumps(query_manager.get_select_set())


def model_edgebar_features() -> str:
    """Full feature tree from DesignEdgebarFeatures."""
    return json.dumps(query_manager.get_design_edgebar_features())


def model_feature_count() -> str:
    """Total count of features."""
    return json.dumps(query_manager.get_feature_count())


def model_camera() -> str:
    """Current camera eye/target/up (meters), projection, and scale."""
    return json.dumps(view_manager.get_camera())


# --- Geometry (12) ---


def geometry_bodies() -> str:
    """All solid bodies in the active part."""
    return json.dumps(query_manager.get_solid_bodies())


def geometry_bounding_box() -> str:
    """Bounding box of the model, in meters."""
    return json.dumps(query_manager.get_bounding_box())


def geometry_face_count() -> str:
    """Total face count on the body."""
    return json.dumps(query_manager.get_face_count())


def geometry_edge_count() -> str:
    """Total edge count on the body."""
    return json.dumps(query_manager.get_edge_count())


def geometry_vertex_count() -> str:
    """Total vertex count on the body."""
    return json.dumps(query_manager.get_vertex_count())


def geometry_faces() -> str:
    """All faces on the body with geometry info; 0-based face indices."""
    return json.dumps(query_manager.get_body_faces())


def geometry_edges() -> str:
    """Edge information from the model body; 0-based indices."""
    return json.dumps(query_manager.get_body_edges())


def geometry_body_color() -> str:
    """Current body color (RGB 0-255) of the active part."""
    return json.dumps(query_manager.get_body_color())


def geometry_surface_area() -> str:
    """Total surface area of the active part, in square meters."""
    return json.dumps(query_manager.get_surface_area())


def geometry_volume() -> str:
    """Volume of the active part, in cubic meters."""
    return json.dumps(query_manager.get_volume())


def geometry_center_of_gravity() -> str:
    """Center of gravity of the active part, in meters."""
    return json.dumps(query_manager.get_center_of_gravity())


def geometry_moments_of_inertia() -> str:
    """Moments of inertia of the active part."""
    return json.dumps(query_manager.get_moments_of_inertia())


# --- Material (2) ---


def material_list() -> str:
    """List of available materials."""
    return json.dumps(query_manager.get_material_list())


def material_table() -> str:
    """Full material table with properties."""
    return json.dumps(query_manager.get_material_table())


# --- Sketch (3) ---


def sketch_info() -> str:
    """Geometry counts in the active sketch."""
    return json.dumps(sketch_manager.get_sketch_info())


def sketch_matrix() -> str:
    """Sketch coordinate system matrix (2D-to-3D transformation)."""
    return json.dumps(sketch_manager.get_sketch_matrix())


def sketch_constraints() -> str:
    """Constraints in the active sketch."""
    return json.dumps(sketch_manager.get_sketch_constraints())


# --- Drawing (2) ---


def drawing_sheets() -> str:
    """Information about the active draft sheet."""
    return json.dumps(export_manager.get_sheet_info())


def drawing_view_count() -> str:
    """Number of drawing views on the active sheet."""
    return json.dumps(export_manager.get_drawing_view_count())


# ===================================================================
# Tier 2: Resource templates (parameterized) — 15
# ===================================================================

# --- Model feature templates (5) ---


def model_feature_by_index(index: int) -> str:
    """Detailed info about a feature by 0-based index."""
    return json.dumps(feature_manager.get_feature_info(int(index)))


def model_feature_dimensions(name: str) -> str:
    """Dimensions/parameters of a named feature, in meters."""
    return json.dumps(query_manager.get_feature_dimensions(name))


def model_feature_status(name: str) -> str:
    """Status of a feature (OK, suppressed, failed, etc.)."""
    return json.dumps(query_manager.get_feature_status(name))


def model_feature_profiles(name: str) -> str:
    """Sketch profiles associated with a feature."""
    return json.dumps(query_manager.get_feature_profiles(name))


def model_feature_parents(name: str) -> str:
    """Parent geometry/features of a named feature."""
    return json.dumps(query_manager.get_feature_parents(name))


# --- Geometry templates (3) ---


def geometry_face_by_index(index: int) -> str:
    """Detailed information about a face, by 0-based index."""
    return json.dumps(query_manager.get_face_info(int(index)))


def geometry_face_area(index: int) -> str:
    """Area of a face (square meters), by 0-based index."""
    return json.dumps(query_manager.get_face_area(int(index)))


def geometry_edge_by_face(face: int, edge: int) -> str:
    """Detailed info about an edge: 0-based face index, 0-based edge on it."""
    return json.dumps(query_manager.get_edge_info(int(face), int(edge)))


# --- Variable templates (3) ---


def model_variable_by_name(name: str) -> str:
    """Value of a specific variable by name."""
    return json.dumps(query_manager.get_variable(name))


def model_variable_formula(name: str) -> str:
    """Formula of a variable by name."""
    return json.dumps(query_manager.get_variable_formula(name))


def model_variable_names(name: str) -> str:
    """DisplayName and SystemName of a variable."""
    return json.dumps(query_manager.get_variable_names(name))


# --- Drawing templates (2) ---


def drawing_view_scale(index: int) -> str:
    """Scale of a drawing view, by 0-based index."""
    return json.dumps(export_manager.get_drawing_view_scale(int(index)))


def drawing_view_info(index: int) -> str:
    """Detailed info about a drawing view, by 0-based index."""
    return json.dumps(export_manager.get_drawing_view_info(int(index)))


# --- Material template (1) ---


def material_property(name: str, index: int) -> str:
    """One property of a named material, by 0-based property index."""
    return json.dumps(query_manager.get_material_property(name, int(index)))


# --- Mass properties template (1) ---


def geometry_mass_properties(density: float) -> str:
    """Mass properties for a given density (kg/m3)."""
    return json.dumps(query_manager.get_mass_properties(float(density)))


# ===================================================================
# Registration
# ===================================================================

#: URI -> handler, in registration order. 37 static + 15 templates = 52.
RESOURCES: tuple[tuple[str, Any], ...] = (
    # Application
    ("solidedge://app/info", app_info),
    ("solidedge://app/install", app_install),
    ("solidedge://app/process", app_process),
    ("solidedge://app/connection-status", app_connection_status),
    # Document
    ("solidedge://document/list", document_list),
    ("solidedge://document/active-type", document_active_type),
    ("solidedge://document/count", document_count),
    # Model
    ("solidedge://model/features", model_features),
    ("solidedge://model/ref-planes", model_ref_planes),
    ("solidedge://model/variables", model_variables),
    ("solidedge://model/custom-properties", model_custom_properties),
    ("solidedge://model/document-properties", model_document_properties),
    ("solidedge://model/layers", model_layers),
    ("solidedge://model/mode", model_mode),
    ("solidedge://model/select-set", model_select_set),
    ("solidedge://model/edgebar-features", model_edgebar_features),
    ("solidedge://model/feature-count", model_feature_count),
    ("solidedge://model/camera", model_camera),
    # Geometry
    ("solidedge://geometry/bodies", geometry_bodies),
    ("solidedge://geometry/bounding-box", geometry_bounding_box),
    ("solidedge://geometry/face-count", geometry_face_count),
    ("solidedge://geometry/edge-count", geometry_edge_count),
    ("solidedge://geometry/vertex-count", geometry_vertex_count),
    ("solidedge://geometry/faces", geometry_faces),
    ("solidedge://geometry/edges", geometry_edges),
    ("solidedge://geometry/body-color", geometry_body_color),
    ("solidedge://geometry/surface-area", geometry_surface_area),
    ("solidedge://geometry/volume", geometry_volume),
    ("solidedge://geometry/center-of-gravity", geometry_center_of_gravity),
    ("solidedge://geometry/moments-of-inertia", geometry_moments_of_inertia),
    # Material
    ("solidedge://material/list", material_list),
    ("solidedge://material/table", material_table),
    # Sketch
    ("solidedge://sketch/info", sketch_info),
    ("solidedge://sketch/matrix", sketch_matrix),
    ("solidedge://sketch/constraints", sketch_constraints),
    # Drawing
    ("solidedge://drawing/sheets", drawing_sheets),
    ("solidedge://drawing/view-count", drawing_view_count),
    # Templates: model features
    ("solidedge://model/feature/{index}", model_feature_by_index),
    ("solidedge://model/feature/{name}/dimensions", model_feature_dimensions),
    ("solidedge://model/feature/{name}/status", model_feature_status),
    ("solidedge://model/feature/{name}/profiles", model_feature_profiles),
    ("solidedge://model/feature/{name}/parents", model_feature_parents),
    # Templates: geometry
    ("solidedge://geometry/face/{index}", geometry_face_by_index),
    ("solidedge://geometry/face/{index}/area", geometry_face_area),
    ("solidedge://geometry/face/{face}/edge/{edge}", geometry_edge_by_face),
    # Templates: variables
    ("solidedge://model/variable/{name}", model_variable_by_name),
    ("solidedge://model/variable/{name}/formula", model_variable_formula),
    ("solidedge://model/variable/{name}/names", model_variable_names),
    # Templates: drawing
    ("solidedge://drawing/view/{index}/scale", drawing_view_scale),
    ("solidedge://drawing/view/{index}", drawing_view_info),
    # Template: material
    ("solidedge://material/{name}/property/{index}", material_property),
    # Template: mass properties
    ("solidedge://geometry/mass-properties/{density}", geometry_mass_properties),
)


def register(mcp: Any) -> None:
    """Register read-only MCP resources (37 static + 15 templates)."""
    for uri, fn in RESOURCES:
        register_resource(mcp, uri, fn, tags=RESOURCE_TAGS)
