"""Primitive solid creation and cutout tools."""

from typing import Any, Literal

from solidedge_mcp.backends.validation import validate_numerics
from solidedge_mcp.managers import feature_manager


def create_primitive(
    shape: Literal[
        "box_two_points", "box_center", "box_three_points", "cylinder", "sphere"
    ] = "box_two_points",
    x1: float = 0.0,
    y1: float = 0.0,
    z1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    z2: float = 0.0,
    x3: float = 0.0,
    y3: float = 0.0,
    z3: float = 0.0,
    length: float = 0.0,
    width: float = 0.0,
    height: float = 0.0,
    radius: float = 0.0,
    depth: float = 0.0,
    plane_index: int = 1,
) -> dict[str, Any]:
    """Create a primitive solid (no sketch needed).

    All coordinates/dimensions in meters. box_two_points: (x1,y1,z1)-(x2,y2,z2)
    corners. box_center: center (x1,y1,z1) + length/width/height.
    box_three_points: three corner points. cylinder: center (x1,y1,z1),
    radius, depth. sphere: center (x1,y1,z1), radius.
    plane_index is 1-based (1=Top/XY, 2=Right/YZ, 3=Front/XZ).
    """
    err = validate_numerics(
        x1=x1,
        y1=y1,
        z1=z1,
        x2=x2,
        y2=y2,
        z2=z2,
        x3=x3,
        y3=y3,
        z3=z3,
        length=length,
        width=width,
        height=height,
        radius=radius,
        depth=depth,
    )
    if err:
        return err
    match shape:
        case "box_two_points":
            return feature_manager.create_box_by_two_points(x1, y1, z1, x2, y2, z2, plane_index)
        case "box_center":
            return feature_manager.create_box_by_center(
                x1,
                y1,
                z1,
                length,
                width,
                height,
                plane_index,
            )
        case "box_three_points":
            return feature_manager.create_box_by_three_points(
                x1,
                y1,
                z1,
                x2,
                y2,
                z2,
                x3,
                y3,
                z3,
                plane_index,
            )
        case "cylinder":
            # A cylinder's axial length is "height" to most callers; "depth" is
            # accepted too because that is what the COM parameter is called.
            return feature_manager.create_cylinder(
                x1, y1, z1, radius=radius, height=(height or depth), plane_index=plane_index
            )
        case "sphere":
            return feature_manager.create_sphere(x1, y1, z1, radius, plane_index)
        case _:
            return {"error": f"Unknown shape: {shape}"}


def create_primitive_cutout(
    shape: Literal["box", "cylinder", "sphere"] = "box",
    x1: float = 0.0,
    y1: float = 0.0,
    z1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    z2: float = 0.0,
    radius: float = 0.0,
    height: float = 0.0,
    plane_index: int = 1,
) -> dict[str, Any]:
    """Create a primitive cutout (removes material; no sketch needed).

    All coordinates/dimensions in meters. box: (x1,y1,z1)-(x2,y2,z2) corners.
    cylinder: center (x1,y1,z1), radius, height. sphere: center, radius.
    plane_index is 1-based (1=Top/XY, 2=Right/YZ, 3=Front/XZ).
    """
    err = validate_numerics(
        x1=x1,
        y1=y1,
        z1=z1,
        x2=x2,
        y2=y2,
        z2=z2,
        radius=radius,
        height=height,
    )
    if err:
        return err
    match shape:
        case "box":
            return feature_manager.create_box_cutout_by_two_points(
                x1, y1, z1, x2, y2, z2, plane_index
            )
        case "cylinder":
            return feature_manager.create_cylinder_cutout(x1, y1, z1, radius, height, plane_index)
        case "sphere":
            return feature_manager.create_sphere_cutout(x1, y1, z1, radius, plane_index)
        case _:
            return {"error": f"Unknown shape: {shape}"}
