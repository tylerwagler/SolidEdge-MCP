"""MCP Resources — read-only data endpoints for Solid Edge.

Converts 52 former read-only tools into MCP Resources for cleaner LLM action space.
Static resources require no parameters; resource templates accept URI parameters.
"""

import json

from solidedge_mcp.managers import (
    connection,
    doc_manager,
    export_manager,
    feature_manager,
    query_manager,
    sketch_manager,
    view_manager,
)


def register(mcp):
    """Register read-only MCP resources (37 static + 15 templates)."""

    # ===================================================================
    # Tier 1: Static Resources (no parameters) — 37 resources
    # ===================================================================

    # --- Application (4) ---

    @mcp.resource("solidedge://app/info")
    def app_info() -> str:
        """Solid Edge application information (version, path, document count)."""
        return json.dumps(connection.get_info())

    @mcp.resource("solidedge://app/install")
    def app_install() -> str:
        """Solid Edge installation information (path, language, version)."""
        return json.dumps(connection.get_install_info())

    @mcp.resource("solidedge://app/process")
    def app_process() -> str:
        """Solid Edge process information (PID, window handle)."""
        return json.dumps(connection.get_process_info())

    @mcp.resource("solidedge://app/connection-status")
    def app_connection_status() -> str:
        """Whether Solid Edge is currently connected."""
        return json.dumps({"connected": connection.is_connected()})

    # --- Document (3) ---

    @mcp.resource("solidedge://document/list")
    def document_list() -> str:
        """List of all open documents."""
        return json.dumps(doc_manager.list_documents())

    @mcp.resource("solidedge://document/active-type")
    def document_active_type() -> str:
        """Type of the currently active document."""
        return json.dumps(doc_manager.get_active_document_type())

    @mcp.resource("solidedge://document/count")
    def document_count() -> str:
        """Count of open documents."""
        return json.dumps(doc_manager.get_document_count())

    # --- Model (11) ---

    @mcp.resource("solidedge://model/features")
    def model_features() -> str:
        """All features in the active model."""
        return json.dumps(feature_manager.list_features())

    @mcp.resource("solidedge://model/ref-planes")
    def model_ref_planes() -> str:
        """All reference planes in the active document."""
        return json.dumps(query_manager.get_ref_planes())

    @mcp.resource("solidedge://model/variables")
    def model_variables() -> str:
        """All variables (dimensions, parameters) in the document."""
        return json.dumps(query_manager.get_variables())

    @mcp.resource("solidedge://model/custom-properties")
    def model_custom_properties() -> str:
        """All custom properties."""
        return json.dumps(query_manager.get_custom_properties())

    @mcp.resource("solidedge://model/document-properties")
    def model_document_properties() -> str:
        """Document properties (Title, Subject, Author, etc.)."""
        return json.dumps(query_manager.get_document_properties())

    @mcp.resource("solidedge://model/layers")
    def model_layers() -> str:
        """All layers in the active document."""
        return json.dumps(query_manager.get_layers())

    @mcp.resource("solidedge://model/mode")
    def model_mode() -> str:
        """Current modeling mode (Ordered vs Synchronous)."""
        return json.dumps(query_manager.get_modeling_mode())

    @mcp.resource("solidedge://model/select-set")
    def model_select_set() -> str:
        """Current selection set."""
        return json.dumps(query_manager.get_select_set())

    @mcp.resource("solidedge://model/edgebar-features")
    def model_edgebar_features() -> str:
        """Full feature tree from DesignEdgebarFeatures."""
        return json.dumps(query_manager.get_design_edgebar_features())

    @mcp.resource("solidedge://model/feature-count")
    def model_feature_count() -> str:
        """Total count of features."""
        return json.dumps(query_manager.get_feature_count())

    @mcp.resource("solidedge://model/camera")
    def model_camera() -> str:
        """Current camera position and orientation."""
        return json.dumps(view_manager.get_camera())

    # --- Geometry (14) ---

    @mcp.resource("solidedge://geometry/bodies")
    def geometry_bodies() -> str:
        """All solid bodies in the active part."""
        return json.dumps(query_manager.get_solid_bodies())

    @mcp.resource("solidedge://geometry/bounding-box")
    def geometry_bounding_box() -> str:
        """Bounding box of the model."""
        return json.dumps(query_manager.get_bounding_box())

    @mcp.resource("solidedge://geometry/face-count")
    def geometry_face_count() -> str:
        """Total face count on the body."""
        return json.dumps(query_manager.get_face_count())

    @mcp.resource("solidedge://geometry/edge-count")
    def geometry_edge_count() -> str:
        """Total edge count on the body."""
        return json.dumps(query_manager.get_edge_count())

    @mcp.resource("solidedge://geometry/vertex-count")
    def geometry_vertex_count() -> str:
        """Total vertex count on the body."""
        return json.dumps(query_manager.get_vertex_count())

    @mcp.resource("solidedge://geometry/faces")
    def geometry_faces() -> str:
        """All faces on the body with geometry info."""
        return json.dumps(query_manager.get_body_faces())

    @mcp.resource("solidedge://geometry/edges")
    def geometry_edges() -> str:
        """Edge information from the model body."""
        return json.dumps(query_manager.get_body_edges())

    @mcp.resource("solidedge://geometry/body-color")
    def geometry_body_color() -> str:
        """Current body color of the active part."""
        return json.dumps(query_manager.get_body_color())

    @mcp.resource("solidedge://geometry/surface-area")
    def geometry_surface_area() -> str:
        """Total surface area of the active part."""
        return json.dumps(query_manager.get_surface_area())

    @mcp.resource("solidedge://geometry/volume")
    def geometry_volume() -> str:
        """Volume of the active part."""
        return json.dumps(query_manager.get_volume())

    @mcp.resource("solidedge://geometry/center-of-gravity")
    def geometry_center_of_gravity() -> str:
        """Center of gravity of the active part."""
        return json.dumps(query_manager.get_center_of_gravity())

    @mcp.resource("solidedge://geometry/moments-of-inertia")
    def geometry_moments_of_inertia() -> str:
        """Moments of inertia of the active part."""
        return json.dumps(query_manager.get_moments_of_inertia())

    # --- Material (2) ---

    @mcp.resource("solidedge://material/list")
    def material_list() -> str:
        """List of available materials."""
        return json.dumps(query_manager.get_material_list())

    @mcp.resource("solidedge://material/table")
    def material_table() -> str:
        """Full material table with properties."""
        return json.dumps(query_manager.get_material_table())

    # --- Sketch (3) ---

    @mcp.resource("solidedge://sketch/info")
    def sketch_info() -> str:
        """Geometry counts in the active sketch."""
        return json.dumps(sketch_manager.get_sketch_info())

    @mcp.resource("solidedge://sketch/matrix")
    def sketch_matrix() -> str:
        """Sketch coordinate system matrix (2D-to-3D transformation)."""
        return json.dumps(sketch_manager.get_sketch_matrix())

    @mcp.resource("solidedge://sketch/constraints")
    def sketch_constraints() -> str:
        """Constraints in the active sketch."""
        return json.dumps(sketch_manager.get_sketch_constraints())

    # --- Drawing (2) ---

    @mcp.resource("solidedge://drawing/sheets")
    def drawing_sheets() -> str:
        """Information about the active draft sheet."""
        return json.dumps(export_manager.get_sheet_info())

    @mcp.resource("solidedge://drawing/view-count")
    def drawing_view_count() -> str:
        """Number of drawing views on the active sheet."""
        return json.dumps(export_manager.get_drawing_view_count())

    # ===================================================================
    # Tier 2: Resource Templates (parameterized) — 15 templates
    # ===================================================================

    # --- Model Feature Templates (5) ---

    @mcp.resource("solidedge://model/feature/{index}")
    def model_feature_by_index(index: int) -> str:
        """Detailed info about a feature by index."""
        return json.dumps(feature_manager.get_feature_info(int(index)))

    @mcp.resource("solidedge://model/feature/{name}/dimensions")
    def model_feature_dimensions(name: str) -> str:
        """Dimensions/parameters of a named feature."""
        return json.dumps(query_manager.get_feature_dimensions(name))

    @mcp.resource("solidedge://model/feature/{name}/status")
    def model_feature_status(name: str) -> str:
        """Status of a feature (OK, suppressed, failed, etc.)."""
        return json.dumps(query_manager.get_feature_status(name))

    @mcp.resource("solidedge://model/feature/{name}/profiles")
    def model_feature_profiles(name: str) -> str:
        """Sketch profiles associated with a feature."""
        return json.dumps(query_manager.get_feature_profiles(name))

    @mcp.resource("solidedge://model/feature/{name}/parents")
    def model_feature_parents(name: str) -> str:
        """Parent geometry/features of a named feature."""
        return json.dumps(query_manager.get_feature_parents(name))

    # --- Geometry Templates (4) ---

    @mcp.resource("solidedge://geometry/face/{index}")
    def geometry_face_by_index(index: int) -> str:
        """Detailed information about a specific face."""
        return json.dumps(query_manager.get_face_info(int(index)))

    @mcp.resource("solidedge://geometry/face/{index}/area")
    def geometry_face_area(index: int) -> str:
        """Area of a specific face."""
        return json.dumps(query_manager.get_face_area(int(index)))

    @mcp.resource("solidedge://geometry/face/{face}/edge/{edge}")
    def geometry_edge_by_face(face: int, edge: int) -> str:
        """Detailed info about a specific edge on a face."""
        return json.dumps(query_manager.get_edge_info(int(face), int(edge)))

    # --- Variable Templates (3) ---

    @mcp.resource("solidedge://model/variable/{name}")
    def model_variable_by_name(name: str) -> str:
        """Value of a specific variable by name."""
        return json.dumps(query_manager.get_variable(name))

    @mcp.resource("solidedge://model/variable/{name}/formula")
    def model_variable_formula(name: str) -> str:
        """Formula of a variable by name."""
        return json.dumps(query_manager.get_variable_formula(name))

    @mcp.resource("solidedge://model/variable/{name}/names")
    def model_variable_names(name: str) -> str:
        """DisplayName and SystemName of a variable."""
        return json.dumps(query_manager.get_variable_names(name))

    # --- Drawing Templates (2) ---

    @mcp.resource("solidedge://drawing/view/{index}/scale")
    def drawing_view_scale(index: int) -> str:
        """Scale of a drawing view by index."""
        return json.dumps(export_manager.get_drawing_view_scale(int(index)))

    @mcp.resource("solidedge://drawing/view/{index}")
    def drawing_view_info(index: int) -> str:
        """Detailed info about a drawing view."""
        return json.dumps(export_manager.get_drawing_view_info(int(index)))

    # --- Material Template (1) ---

    @mcp.resource("solidedge://material/{name}/property/{index}")
    def material_property(name: str, index: int) -> str:
        """Specific property of a material."""
        return json.dumps(query_manager.get_material_property(name, int(index)))

    # --- Mass Properties Template (1) ---

    @mcp.resource("solidedge://geometry/mass-properties/{density}")
    def geometry_mass_properties(density: float) -> str:
        """Mass properties for a given density (kg/m3)."""
        return json.dumps(query_manager.get_mass_properties(float(density)))
