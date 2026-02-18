# Add MCP Resources for Read-Only Data

## Goal
Convert pure read-only tools to **MCP Resources** (`@mcp.resource`) to reduce tool noise, enable client-side caching, and make the LLM's decision space cleaner. The existing tools remain registered for backward compatibility (dual-registration) unless opted out.

## Background

FastMCP supports two resource types:
- **Resources** (`solidedge://app/info`): No parameters, returns data about a fixed endpoint
- **Resource Templates** (`solidedge://feature/{index}/info`): URI contains `{param}` placeholders, resolved at read time

Resources use `mcp.resource(uri)(func)` for programmatic registration (same pattern as tools).

## User Review Required

> [!NOTE]
> **Strategy: Full Replacement.** Since this is a v1 with no existing users, we will **remove** the read-only tools and **replace** them with resources. This cuts the tool count by ~52, giving the LLM a much cleaner action space.
>
> If a client doesn't support `resources/read`, these data endpoints simply won't be visible to it.

---

## Proposed Changes

### URI Scheme
All resources use the `solidedge://` scheme with a logical path hierarchy:

| Category | URI Pattern | Example |
|----------|-------------|---------|
| Application | `solidedge://app/{topic}` | `solidedge://app/info` |
| Document | `solidedge://document/{topic}` | `solidedge://document/list` |
| Model | `solidedge://model/{topic}` | `solidedge://model/features` |
| Geometry | `solidedge://geometry/{topic}` | `solidedge://geometry/bodies` |
| Material | `solidedge://material/{topic}` | `solidedge://material/list` |
| Sketch | `solidedge://sketch/{topic}` | `solidedge://sketch/info` |
| Drawing | `solidedge://drawing/{topic}` | `solidedge://drawing/sheets` |

---

### Tier 1: Static Resources (no parameters)

These return data about the current state with no inputs required:

| Current Tool | Resource URI | Module |
|---|---|---|
| `get_application_info` | `solidedge://app/info` | connection |
| `get_install_info` | `solidedge://app/install` | connection |
| `get_process_info` | `solidedge://app/process` | connection |
| `is_connected` | `solidedge://app/connection-status` | connection |
| `list_documents` | `solidedge://document/list` | documents |
| `get_active_document_type` | `solidedge://document/active-type` | documents |
| `get_document_count` | `solidedge://document/count` | documents |
| `list_features` | `solidedge://model/features` | features |
| `get_ref_planes` | `solidedge://model/ref-planes` | query |
| `get_variables` | `solidedge://model/variables` | query |
| `get_custom_properties` | `solidedge://model/custom-properties` | query |
| `get_document_properties` | `solidedge://model/document-properties` | query |
| `get_solid_bodies` | `solidedge://geometry/bodies` | query |
| `get_bounding_box` | `solidedge://geometry/bounding-box` | query |
| `get_face_count` | `solidedge://geometry/face-count` | query |
| `get_edge_count` | `solidedge://geometry/edge-count` | query |
| `get_vertex_count` | `solidedge://geometry/vertex-count` | query |
| `get_body_faces` | `solidedge://geometry/faces` | query |
| `get_body_edges` | `solidedge://geometry/edges` | query |
| `get_body_color` | `solidedge://geometry/body-color` | query |
| `get_surface_area` | `solidedge://geometry/surface-area` | query |
| `get_volume` | `solidedge://geometry/volume` | query |
| `get_center_of_gravity` | `solidedge://geometry/center-of-gravity` | query |
| `get_moments_of_inertia` | `solidedge://geometry/moments-of-inertia` | query |
| `get_material_list` | `solidedge://material/list` | query |
| `get_material_table` | `solidedge://material/table` | query |
| `get_layers` | `solidedge://model/layers` | query |
| `get_modeling_mode` | `solidedge://model/mode` | query |
| `get_select_set` | `solidedge://model/select-set` | query |
| `get_design_edgebar_features` | `solidedge://model/edgebar-features` | query |
| `get_sketch_info` | `solidedge://sketch/info` | sketching |
| `get_sketch_matrix` | `solidedge://sketch/matrix` | sketching |
| `get_sketch_constraints` | `solidedge://sketch/constraints` | sketching |
| `get_sheet_info` | `solidedge://drawing/sheets` | export |
| `get_drawing_view_count` | `solidedge://drawing/view-count` | export |
| `get_camera` | `solidedge://model/camera` | export |
| `get_feature_count` | `solidedge://model/feature-count` | query |

**Total: 37 static resources**

### Tier 2: Resource Templates (parameterized)

These require one or more parameters:

| Current Tool | Resource URI Template | Parameters |
|---|---|---|
| `get_feature_info` | `solidedge://model/feature/{index}` | `index: int` |
| `get_feature_dimensions` | `solidedge://model/feature/{name}/dimensions` | `name: str` |
| `get_feature_status` | `solidedge://model/feature/{name}/status` | `name: str` |
| `get_feature_profiles` | `solidedge://model/feature/{name}/profiles` | `name: str` |
| `get_feature_parents` | `solidedge://model/feature/{name}/parents` | `name: str` |
| `get_face_info` | `solidedge://geometry/face/{index}` | `index: int` |
| `get_face_area` | `solidedge://geometry/face/{index}/area` | `index: int` |
| `get_edge_info` | `solidedge://geometry/face/{face}/edge/{edge}` | `face: int, edge: int` |
| `get_variable` | `solidedge://model/variable/{name}` | `name: str` |
| `get_variable_formula` | `solidedge://model/variable/{name}/formula` | `name: str` |
| `get_variable_names` | `solidedge://model/variable/{name}/names` | `name: str` |
| `get_drawing_view_scale` | `solidedge://drawing/view/{index}/scale` | `index: int` |
| `get_drawing_view_info` | `solidedge://drawing/view/{index}` | `index: int` |
| `get_material_property` | `solidedge://material/{name}/property/{index}` | `name: str, index: int` |
| `get_mass_properties` | `solidedge://geometry/mass-properties/{density}` | `density: float` |

**Total: 15 resource templates**

### Tier 3: Stay as Tools Only

These are either computational or have side-effects despite the `get_` prefix:

| Tool | Reason |
|---|---|
| `get_body_facet_data` | Expensive computation (tessellation), configurable tolerance |
| `measure_distance` | Not a "read" — it's a computation on arbitrary coordinates |
| `measure_angle` | Same as above |
| `diagnose_feature` | Diagnostic action, not a data read |
| `query_list_features` / `query_variables` | Redundant with resources above |

---

### Implementation

#### [NEW] [resources.py](file:///c:/Users/tyler/Dev/repos/SolidEdge_MCP/src/solidedge_mcp/tools/resources.py)

A single new module that registers all resources. Example structure:

```python
"""MCP Resources — read-only data endpoints for Solid Edge."""

import json
from solidedge_mcp.managers import (
    connect, doc_manager, feature_manager,
    query_manager, sketch_manager, view_manager,
    export_manager,
)


def register(mcp):
    """Register read-only MCP resources."""

    # --- Static Resources ---

    @mcp.resource("solidedge://app/info")
    def app_info() -> str:
        return json.dumps(connect.get_application_info())

    @mcp.resource("solidedge://document/list")
    def document_list() -> str:
        return json.dumps(doc_manager.list_documents())

    # ... (remaining static resources)

    # --- Resource Templates ---

    @mcp.resource("solidedge://model/feature/{index}")
    def feature_info(index: int) -> str:
        return json.dumps(feature_manager.get_feature_info(index))

    @mcp.resource("solidedge://model/variable/{name}")
    def variable_value(name: str) -> str:
        return json.dumps(query_manager.get_variable(name))

    # ... (remaining templates)
```

#### [MODIFY] [\_\_init\_\_.py](file:///c:/Users/tyler/Dev/repos/SolidEdge_MCP/src/solidedge_mcp/tools/__init__.py)

Add `resources` to the import list and call `resources.register(mcp)`.

#### [MODIFY] Existing tool modules

**Remove** the 52 read-only functions and their `mcp.tool()` registrations from `connection.py`, `documents.py`, `sketching.py`, `features.py`, `query.py`, and `export.py`. This reduces the total tool count by ~52.

---

## Verification Plan

### Automated
1. Run `verify_tools_list.py` — tool count must not decrease
2. Add a `verify_resources_list.py` script that lists `mcp._resource_manager._resources` and asserts count ≥ 52
3. Run `pytest tests/unit` — all existing tests pass

### Manual
1. Start the server, connect with MCP Inspector
2. Verify resources appear under "Resources" tab
3. Read `solidedge://app/info` and confirm it returns valid JSON
4. Read `solidedge://model/feature/0` to test a template
