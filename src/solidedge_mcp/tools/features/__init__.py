"""Feature modeling tools for Solid Edge MCP.

Consolidated composite tools that use a `method` (or `shape`/`type`/`action`)
discriminator parameter to dispatch to the correct backend method.
"""

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
    create_thin_wall,
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
    "create_thin_wall",
    "create_thread",
    "create_web_network",
    "delete_topology",
    "face_operation",
    "manage_feature",
    "sheet_metal_misc",
    "simplify",
    "thicken",
]


def register(mcp):
    """Register feature tools with the MCP server."""
    # Composite tools
    mcp.tool()(create_extrude)
    mcp.tool()(create_revolve)
    mcp.tool()(create_extruded_cutout)
    mcp.tool()(create_revolved_cutout)
    mcp.tool()(create_normal_cutout)
    mcp.tool()(create_lofted_cutout)
    mcp.tool()(create_swept_cutout)
    mcp.tool()(create_helix)
    mcp.tool()(create_helix_cutout)
    mcp.tool()(create_loft)
    mcp.tool()(create_sweep)
    mcp.tool()(create_extruded_surface)
    mcp.tool()(create_revolved_surface)
    mcp.tool()(create_lofted_surface)
    mcp.tool()(create_swept_surface)
    mcp.tool()(thicken)
    mcp.tool()(create_primitive)
    mcp.tool()(create_primitive_cutout)
    mcp.tool()(create_hole)
    mcp.tool()(create_round)
    mcp.tool()(create_chamfer)
    mcp.tool()(create_blend)
    mcp.tool()(delete_topology)
    mcp.tool()(create_ref_plane)
    mcp.tool()(create_ref_plane_on_curve)
    mcp.tool()(create_ref_plane_tangent)
    mcp.tool()(create_flange)
    mcp.tool()(create_contour_flange)
    mcp.tool()(create_sheet_metal_base)
    mcp.tool()(create_lofted_flange)
    mcp.tool()(create_bend)
    mcp.tool()(create_slot)
    mcp.tool()(create_thread)
    mcp.tool()(create_drawn_cutout)
    mcp.tool()(create_dimple)
    mcp.tool()(create_louver)
    mcp.tool()(create_pattern)
    mcp.tool()(create_mirror)
    mcp.tool()(create_thin_wall)
    mcp.tool()(face_operation)
    mcp.tool()(add_body)
    mcp.tool()(simplify)
    mcp.tool()(manage_feature)
    mcp.tool()(sheet_metal_misc)
    mcp.tool()(create_stamped)
    mcp.tool()(create_surface_mark)
    mcp.tool()(create_reinforcement)
    # Standalone tools
    mcp.tool()(create_web_network)
    mcp.tool()(create_split)
    mcp.tool()(create_draft_angle)
    mcp.tool()(create_bounded_surface)
