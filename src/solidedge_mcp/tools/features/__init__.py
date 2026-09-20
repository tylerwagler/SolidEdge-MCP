"""Feature modeling tools for Solid Edge MCP.

Consolidated composite tools that use a `method` (or `shape`/`type`/`action`)
discriminator parameter to dispatch to the correct backend method.
"""

from typing import Any

from solidedge_mcp.tools._registry import register_tool
from solidedge_mcp.tools.features._cutout import (
    create_extruded_cutout,
    create_helix_cutout,
    create_lofted_cutout,
    create_normal_cutout,
    create_revolved_cutout,
    create_swept_cutout,
)
from solidedge_mcp.tools.features._extrude import create_extrude
from solidedge_mcp.tools.features._holes import create_hole
from solidedge_mcp.tools.features._loft_sweep import create_helix, create_loft, create_sweep
from solidedge_mcp.tools.features._misc import (
    add_body,
    create_draft_angle,
    create_mirror,
    create_pattern,
    face_operation,
    manage_feature,
    simplify,
    thicken,
)
from solidedge_mcp.tools.features._primitives import create_primitive, create_primitive_cutout
from solidedge_mcp.tools.features._ref_planes import (
    create_ref_plane,
    create_ref_plane_on_curve,
    create_ref_plane_tangent,
)
from solidedge_mcp.tools.features._revolve import create_revolve
from solidedge_mcp.tools.features._rounds_chamfers import (
    create_blend,
    create_chamfer,
    create_round,
    delete_topology,
)
from solidedge_mcp.tools.features._sheet_metal import (
    create_bend,
    create_contour_flange,
    create_dimple,
    create_drawn_cutout,
    create_flange,
    create_lofted_flange,
    create_louver,
    create_reinforcement,
    create_sheet_metal_base,
    create_slot,
    create_split,
    create_stamped,
    create_surface_mark,
    create_thread,
    create_web_network,
    sheet_metal_misc,
)
from solidedge_mcp.tools.features._surfaces import (
    create_bounded_surface,
    create_extruded_surface,
    create_lofted_surface,
    create_revolved_surface,
    create_swept_surface,
)

__all__ = [
    "add_body",
    "create_bend",
    "create_blend",
    "create_bounded_surface",
    "create_chamfer",
    "create_contour_flange",
    "create_dimple",
    "create_draft_angle",
    "create_drawn_cutout",
    "create_extrude",
    "create_extruded_cutout",
    "create_extruded_surface",
    "create_flange",
    "create_helix",
    "create_helix_cutout",
    "create_hole",
    "create_loft",
    "create_lofted_cutout",
    "create_lofted_flange",
    "create_lofted_surface",
    "create_louver",
    "create_mirror",
    "create_normal_cutout",
    "create_pattern",
    "create_primitive",
    "create_primitive_cutout",
    "create_ref_plane",
    "create_ref_plane_on_curve",
    "create_ref_plane_tangent",
    "create_reinforcement",
    "create_revolve",
    "create_revolved_cutout",
    "create_revolved_surface",
    "create_round",
    "create_sheet_metal_base",
    "create_slot",
    "create_split",
    "create_stamped",
    "create_surface_mark",
    "create_sweep",
    "create_swept_cutout",
    "create_swept_surface",
    "create_thread",
    "create_web_network",
    "delete_topology",
    "face_operation",
    "manage_feature",
    "sheet_metal_misc",
    "simplify",
    "thicken",
]


def register(mcp: Any) -> None:
    """Register feature tools with the MCP server."""
    part = {"part"}
    sheet = {"part", "sheet_metal"}

    # Solids from sketch profiles
    register_tool(mcp, create_extrude, tags=part)
    register_tool(mcp, create_revolve, tags=part)
    register_tool(mcp, create_helix, tags=part)
    register_tool(mcp, create_loft, tags=part)
    register_tool(mcp, create_sweep, tags=part)
    register_tool(mcp, create_primitive, tags=part)
    register_tool(mcp, add_body, tags=part)

    # Material removal
    register_tool(mcp, create_extruded_cutout, tags=part)
    register_tool(mcp, create_revolved_cutout, tags=part)
    register_tool(mcp, create_normal_cutout, tags=part)
    register_tool(mcp, create_lofted_cutout, tags=part)
    register_tool(mcp, create_swept_cutout, tags=part)
    register_tool(mcp, create_helix_cutout, tags=part)
    register_tool(mcp, create_primitive_cutout, tags=part)
    register_tool(mcp, create_hole, tags=part)

    # Surfaces
    register_tool(mcp, create_extruded_surface, tags=part)
    register_tool(mcp, create_revolved_surface, tags=part)
    register_tool(mcp, create_lofted_surface, tags=part)
    register_tool(mcp, create_swept_surface, tags=part)
    register_tool(mcp, create_bounded_surface, tags=part)
    register_tool(mcp, thicken, tags=part)

    # Dress-up features
    register_tool(mcp, create_round, tags=part)
    register_tool(mcp, create_chamfer, tags=part)
    register_tool(mcp, create_blend, tags=part)
    register_tool(mcp, create_draft_angle, tags=part)
    register_tool(mcp, create_thread, tags=part)
    register_tool(mcp, create_surface_mark, tags=part)
    register_tool(mcp, create_reinforcement, tags=part)
    register_tool(mcp, create_web_network, tags=part)
    register_tool(mcp, create_split, tags=part)
    register_tool(mcp, face_operation, tags=part)
    register_tool(mcp, delete_topology, tags=part, destructive=True)

    # Reference geometry
    register_tool(mcp, create_ref_plane, tags=part)
    register_tool(mcp, create_ref_plane_on_curve, tags=part)
    register_tool(mcp, create_ref_plane_tangent, tags=part)

    # Sheet metal
    register_tool(mcp, create_sheet_metal_base, tags=sheet)
    register_tool(mcp, create_flange, tags=sheet)
    register_tool(mcp, create_contour_flange, tags=sheet)
    register_tool(mcp, create_lofted_flange, tags=sheet)
    register_tool(mcp, create_bend, tags=sheet)
    register_tool(mcp, create_slot, tags=sheet)
    register_tool(mcp, create_drawn_cutout, tags=sheet)
    register_tool(mcp, create_dimple, tags=sheet)
    register_tool(mcp, create_louver, tags=sheet)
    register_tool(mcp, create_stamped, tags=sheet)
    register_tool(mcp, sheet_metal_misc, tags=sheet)

    # Model-wide operations
    register_tool(mcp, create_pattern, tags=part)
    register_tool(mcp, create_mirror, tags=part)
    register_tool(mcp, simplify, tags=part)
    register_tool(mcp, manage_feature, tags=part, destructive=True)
