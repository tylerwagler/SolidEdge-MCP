"""Primitive solid creation and cutout tools."""

from solidedge_mcp.managers import feature_manager


def create_primitive(
    shape: str = "box_two_points",
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
) -> dict:
    """Create a primitive solid shape.

    shape: 'box_two_points' | 'box_center' | 'box_three_points'
        | 'cylinder' | 'sphere'

    Parameters (used per shape):
        x1,y1,z1,x2,y2,z2: Corner coords (box_two_points),
            or center+corner (box_center uses x1,y1,z1 as
            center), or first two points (box_three_points).
        x3,y3,z3: Third point (box_three_points).
        length,width,height: Dimensions (box_center).
        radius: Radius (cylinder, sphere).
        depth: Cylinder height (cylinder).
        plane_index: 1=Top/XY, 2=Right/YZ, 3=Front/XZ.
    """
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
            return feature_manager.create_cylinder(x1, y1, z1, radius, depth, plane_index)
        case "sphere":
            return feature_manager.create_sphere(x1, y1, z1, radius, plane_index)
        case _:
            return {"error": f"Unknown shape: {shape}"}


def create_primitive_cutout(
    shape: str = "box",
    x1: float = 0.0,
    y1: float = 0.0,
    z1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    z2: float = 0.0,
    radius: float = 0.0,
    height: float = 0.0,
    plane_index: int = 1,
) -> dict:
    """Create a primitive cutout shape.

    shape: 'box' | 'cylinder' | 'sphere'

    Parameters (used per shape):
        x1,y1,z1,x2,y2,z2: Corner coords (box).
        x1,y1,z1: Center coords (cylinder, sphere).
        radius: Radius (cylinder, sphere).
        height: Cylinder height (cylinder).
        plane_index: 1=Top/XY, 2=Right/YZ, 3=Front/XZ.
    """
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
