"""Drive every registered tool once against a live Solid Edge, and record it.

This is the step that turns "unknown-broken" into "known". Each case opens a
fresh document context through the real MCP server (stdio, the same path a
client takes), runs the tool, classifies what came back, and closes
everything. Outcomes:

    OK      the tool reported success
    UNSUP   it refused honestly, with ``unsupported: True``
    NOOP    it claimed success and a verification decorator caught nothing
            having changed -- the class of bug this server is prone to
    FAIL    it returned an error
    SETUP   the context it needed could not be built
    EXC     the call raised or timed out on the client side

Rows are appended to ``reference/LIVE_SWEEP.md`` as they complete, so a hang
leaves a partial record rather than none. A watchdog dismisses any modal
dialog and notes its title -- a modal blocks Solid Edge's one UI thread and
looks exactly like a crash.

    uv run python scripts/live_sweep.py            # needs a running Solid Edge
"""

from __future__ import annotations

import asyncio
import ctypes
import datetime as dt
import json
import pathlib
import sys
import tempfile
import threading
import time
from dataclasses import dataclass, field
from typing import Any

from fastmcp import Client
from fastmcp.client.transports import StdioTransport

REPO = pathlib.Path(__file__).resolve().parent.parent
OUT_MD = REPO / "reference" / "LIVE_SWEEP.md"
OUT_JSON = pathlib.Path(tempfile.gettempdir()) / "solidedge_live_sweep.json"
CALL_TIMEOUT = 90

# ---------------------------------------------------------------- watchdog
user32 = ctypes.windll.user32
_stop = threading.Event()
modals: list[str] = []


def _watchdog() -> None:
    proto = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    while not _stop.is_set():

        def cb(hwnd: Any, _: Any) -> bool:
            cls = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, cls, 256)
            if cls.value == "#32770" and user32.IsWindowVisible(hwnd):
                title = ctypes.create_unicode_buffer(512)
                user32.GetWindowTextW(hwnd, title, 512)
                modals.append(title.value)
                btn = user32.GetDlgItem(hwnd, 2) or user32.GetDlgItem(hwnd, 1)
                if btn:
                    user32.SendMessageW(btn, 0x00F5, 0, 0)
            return True

        user32.EnumWindows(proto(cb), 0)
        time.sleep(0.4)


# ---------------------------------------------------------------- contexts
TMP = pathlib.Path(tempfile.mkdtemp(prefix="se_sweep_"))
BOX = TMP / "box.par"

Step = tuple[str, dict[str, Any]]


def _part() -> list[Step]:
    return [("create_document", {"type": "part"})]


def _rect_closed(plane: str = "Top") -> list[Step]:
    return [
        ("manage_sketch", {"action": "create", "plane": plane}),
        ("draw", {"shape": "rectangle", "x1": 0, "y1": 0, "x2": 0.08, "y2": 0.048}),
        ("manage_sketch", {"action": "close"}),
    ]


def _circle_closed(plane: str = "Top", cx: float = 0.02, cy: float = 0.02) -> list[Step]:
    return [
        ("manage_sketch", {"action": "create", "plane": plane}),
        ("draw", {"shape": "circle", "center_x": cx, "center_y": cy, "radius": 0.004}),
        ("manage_sketch", {"action": "close"}),
    ]


def _rev_closed() -> list[Step]:
    return [
        ("manage_sketch", {"action": "create", "plane": "Front"}),
        ("draw", {"shape": "circle", "center_x": 0.03, "center_y": 0.01, "radius": 0.004}),
        ("manage_sketch", {"action": "set_axis", "x1": 0.0, "y1": 0.0, "x2": 0.0, "y2": 0.03}),
        ("manage_sketch", {"action": "close"}),
    ]


def _box() -> list[Step]:
    return _part() + _rect_closed() + [("create_extrude", {"method": "finite", "distance": 0.03})]


def _loft_profiles() -> list[Step]:
    return [
        ("manage_sketch", {"action": "create", "plane": "Top"}),
        ("draw", {"shape": "rectangle", "x1": 0.02, "y1": 0.02, "x2": 0.06, "y2": 0.05}),
        ("manage_sketch", {"action": "close"}),
        ("create_ref_plane", {"method": "offset", "parent_plane_index": 1, "distance": 0.04}),
        ("manage_sketch", {"action": "create_on_plane", "plane_index": 4}),
        ("draw", {"shape": "rectangle", "x1": 0.03, "y1": 0.025, "x2": 0.05, "y2": 0.045}),
        ("manage_sketch", {"action": "close"}),
    ]


def _sweep_profiles() -> list[Step]:
    return [
        ("manage_sketch", {"action": "create", "plane": "Front"}),
        ("draw", {"shape": "line", "x1": 0.02, "y1": 0.0, "x2": 0.02, "y2": 0.03}),
        ("manage_sketch", {"action": "close"}),
        ("manage_sketch", {"action": "create", "plane": "Top"}),
        ("draw", {"shape": "circle", "center_x": 0.02, "center_y": 0.02, "radius": 0.004}),
        ("manage_sketch", {"action": "close"}),
    ]


def _sheet() -> list[Step]:
    return (
        [("create_document", {"type": "sheet_metal"})]
        + _rect_closed()
        + [("create_sheet_metal_base", {"type": "tab", "thickness": 0.002})]
    )


def _asm(n: int = 1) -> list[Step]:
    steps: list[Step] = [("create_document", {"type": "assembly"})]
    for i in range(n):
        steps.append(
            (
                "add_assembly_component",
                {"method": "basic", "file_path": str(BOX), "x": 0.1 * i, "y": 0, "z": 0},
            )
        )
    return steps


def _draft() -> list[Step]:
    return [("create_document", {"type": "draft"})]


def _draft_link() -> list[Step]:
    """A draft made from the saved box, so it carries a model link and a view."""
    return [
        ("open_document", {"method": "foreground", "file_path": str(BOX)}),
        ("manage_sheet", {"action": "create_drawing", "views": ["Front", "Top"]}),
    ]


CONTEXTS: dict[str, list[Step]] = {
    "none": [],
    "part": _part(),
    "part_sketch": _part() + [("manage_sketch", {"action": "create", "plane": "Top"})],
    "part_sketch_rect": _part()
    + [
        ("manage_sketch", {"action": "create", "plane": "Top"}),
        ("draw", {"shape": "rectangle", "x1": 0, "y1": 0, "x2": 0.08, "y2": 0.048}),
    ],
    "part_rect_closed": _part() + _rect_closed(),
    "part_rev_closed": _part() + _rev_closed(),
    "part_loft": _part() + _loft_profiles(),
    "part_sweep": _part() + _sweep_profiles(),
    "part_line_closed": _part()
    + [
        ("manage_sketch", {"action": "create", "plane": "Top"}),
        ("draw", {"shape": "line", "x1": 0.0, "y1": 0.0, "x2": 0.05, "y2": 0.04}),
        ("manage_sketch", {"action": "close"}),
    ],
    "box": _box(),
    "box_circle_closed": _box() + _circle_closed(),
    "box_rev_closed": _box() + _rev_closed(),
    "box_loft": _box() + _loft_profiles(),
    "box_sweep": _box() + _sweep_profiles(),
    "box_round": _box() + [("create_round", {"method": "all_edges", "radius": 0.002})],
    "sheet": _sheet(),
    "sheet_rect_closed": [("create_document", {"type": "sheet_metal"})] + _rect_closed(),
    "sheet_circle_closed": _sheet() + _circle_closed(),
    "sheet_line_closed": _sheet()
    + [
        ("manage_sketch", {"action": "create", "plane": "Top"}),
        ("draw", {"shape": "line", "x1": 0.04, "y1": 0.0, "x2": 0.04, "y2": 0.048}),
        ("manage_sketch", {"action": "close"}),
    ],
    "asm": _asm(1),
    "asm2": _asm(2),
    "asm_circle_closed": _asm(1) + _circle_closed(),
    "draft": _draft(),
    "draft_line": _draft()
    + [("draw_sheet_geometry", {"shape": "line", "x1": 0.02, "y1": 0.02, "x2": 0.12, "y2": 0.02})],
    "draft_link": _draft_link(),
}


# ------------------------------------------------------------------- cases
@dataclass
class Case:
    tool: str
    args: dict[str, Any]
    ctx: str = "box"
    note: str = ""


F = "ExtrudedProtrusion_1"
STEP = str(TMP / "box.step")
CASES: list[Case] = [
    # connection / app
    Case("manage_connection", {"action": "connect", "start_if_needed": False}, "none"),
    Case("get_active_command", {}, "none"),
    Case("arrange_windows", {}, "part"),
    Case("app_config", {"property": "get_visible"}, "none"),
    Case("app_command", {"action": "idle"}, "none"),
    Case(
        "convert_by_file_path",
        {"input_path": str(BOX), "output_path": STEP, "overwrite": True},
        "none",
    ),
    Case("run_macro", {"filename": str(TMP / "nowhere.bas")}, "none", "a missing macro"),
    # documents
    Case("create_document", {"type": "part"}, "none"),
    Case("create_document", {"type": "weldment"}, "none", "refuses without a template"),
    Case("activate_document", {"name_or_index": 0}, "part"),
    Case("close_document", {"scope": "active", "save": False}, "part"),
    Case("open_document", {"method": "foreground", "file_path": str(BOX)}, "none"),
    Case("save_document", {"method": "copy_as", "file_path": str(TMP / "copy.par")}, "box"),
    Case("import_file", {"file_path": str(BOX)}, "none"),
    Case("undo_redo", {"action": "undo"}, "box"),
    # diagnostics
    Case("diagnose_api", {}, "none"),
    Case("diagnose_feature_tool", {"feature_index": 0}, "box"),
    # sketching
    Case("manage_sketch", {"action": "create", "plane": "Top"}, "part"),
    Case(
        "draw",
        {"shape": "circle", "center_x": 0.02, "center_y": 0.02, "radius": 0.005},
        "part_sketch",
    ),
    Case("sketch_modify", {"action": "fillet", "radius": 0.002}, "part_sketch_rect"),
    Case(
        "sketch_constraint",
        {
            "type": "geometric",
            "constraint_type": "Horizontal",
            "element1_type": "line",
            "element1_index": 0,
        },
        "part_sketch_rect",
    ),
    Case(
        "sketch_advanced_modify",
        {
            "action": "offset_2d",
            "offset_distance": 0.002,
            "offset_side_x": 0.1,
            "offset_side_y": 0.1,
        },
        "part_sketch_rect",
    ),
    Case("sketch_project", {"source": "ref_plane", "plane_index": 2}, "part_sketch"),
    # part features
    Case("create_extrude", {"method": "finite", "distance": 0.03}, "part_rect_closed"),
    Case("create_extruded_cutout", {"method": "finite", "distance": 0.01}, "box_circle_closed"),
    Case(
        "create_normal_cutout",
        {"method": "finite", "distance": 0.01},
        "box_circle_closed",
        "refuses on a part",
    ),
    Case("create_normal_cutout", {"method": "finite", "distance": 0.01}, "sheet_circle_closed"),
    Case("create_revolve", {"method": "finite", "angle": 90.0}, "part_rev_closed"),
    # A 90-degree sweep from the Front plane heads into -Y, away from the box
    # (the Front-plane normal quirk), and removes nothing; a full revolve
    # cannot miss.
    Case("create_revolved_cutout", {"method": "finite", "angle": 360.0}, "box_rev_closed"),
    Case("create_hole", {"method": "through_all", "x": 0.02, "y": 0.02, "diameter": 0.006}, "box"),
    Case("create_round", {"method": "all_edges", "radius": 0.002}, "box"),
    Case("create_chamfer", {"method": "equal", "distance": 0.002}, "box"),
    Case("create_draft_angle", {"face_index": 0, "angle": 5.0}, "box"),
    Case(
        "create_ref_plane", {"method": "offset", "parent_plane_index": 1, "distance": 0.02}, "part"
    ),
    Case(
        "create_ref_plane_on_curve",
        {"method": "normal_to_curve", "curve_end": "End", "pivot_plane_index": 2},
        "part_line_closed",
    ),
    Case(
        "create_ref_plane_tangent",
        {"method": "tangent_cylinder_angle", "face_index": 0, "angle": 0.0},
        "box",
        "a box has no cylinder",
    ),
    Case(
        "create_primitive",
        {"shape": "box_two_points", "x1": 0, "y1": 0, "z1": 0, "x2": 0.05, "y2": 0.04, "z2": 0.03},
        "part",
    ),
    Case(
        "create_primitive_cutout",
        {"shape": "cylinder", "x1": 0.02, "y1": 0.02, "z1": 0.0, "radius": 0.005, "height": 0.03},
        "box",
    ),
    Case("create_loft", {"method": "solid"}, "part_loft"),
    Case("create_lofted_cutout", {"method": "basic"}, "box_loft"),
    Case("create_sweep", {"method": "solid", "path_profile_index": 0}, "part_sweep"),
    Case("create_swept_cutout", {"method": "basic", "path_profile_index": 0}, "box_sweep"),
    Case("create_helix", {"method": "finite", "pitch": 0.005, "height": 0.02}, "part_rev_closed"),
    Case(
        "create_helix_cutout",
        {"method": "finite", "pitch": 0.005, "height": 0.02},
        "box_rev_closed",
    ),
    Case("create_extruded_surface", {"method": "finite", "distance": 0.02}, "part_rect_closed"),
    Case("create_revolved_surface", {"method": "finite", "angle": 90.0}, "part_rev_closed"),
    Case("create_lofted_surface", {"method": "basic"}, "part_loft"),
    Case("create_swept_surface", {"method": "basic", "path_profile_index": 0}, "part_sweep"),
    Case("create_bounded_surface", {}, "part_rect_closed"),
    Case("create_mirror", {"method": "basic", "feature_name": F, "mirror_plane_index": 2}, "box"),
    Case(
        "create_pattern",
        {
            "method": "rectangular_ex",
            "feature_name": F,
            "x_count": 2,
            "y_count": 1,
            "x_spacing": 0.1,
            "y_spacing": 0.1,
            "plane_index": 1,
        },
        "box",
    ),
    Case(
        "create_thread",
        {"method": "basic", "face_index": 0, "thread_diameter": 0.006, "thread_depth": 0.01},
        "box",
    ),
    Case("create_blend", {"method": "basic", "radius": 0.002, "face_index": 0}, "box"),
    Case("create_split", {}, "box"),
    Case("thicken", {"method": "basic", "thickness": 0.002}, "box"),
    Case("create_web_network", {"thickness": 0.002, "depth": 0.01}, "box_circle_closed"),
    Case("create_reinforcement", {"type": "rib", "thickness": 0.002}, "box_circle_closed"),
    Case("add_body", {"method": "basic", "body_type": "Solid", "body_name": "B2"}, "box"),
    Case("delete_topology", {"type": "blend"}, "box_round"),
    Case(
        "face_operation",
        {"type": "rotate_by_edge", "face_index": 0, "edge_index": 0, "angle": 5.0},
        "box",
    ),
    Case("manage_feature", {"action": "rename", "index": 0, "new_name": "Base"}, "box"),
    Case("simplify", {"method": "auto"}, "box"),
    Case("edit_feature_extent", {"property": "get_direction1", "feature_name": F}, "box"),
    # sheet metal
    Case(
        "create_sheet_metal_base",
        {"type": "tab", "thickness": 0.002},
        "sheet_rect_closed",
    ),
    Case(
        "create_flange",
        {"method": "basic", "face_index": 0, "edge_index": 0, "flange_length": 0.02},
        "sheet",
    ),
    Case("create_bend", {"method": "basic", "bend_angle": 90.0}, "sheet_line_closed"),
    Case(
        "create_contour_flange",
        {
            "method": "ex",
            "thickness": 0.002,
            "bend_radius": 0.002,
            "face_index": 0,
            "edge_index": 0,
        },
        "sheet",
    ),
    Case("create_dimple", {"method": "basic", "depth": 0.005}, "sheet_circle_closed"),
    Case("create_drawn_cutout", {"method": "basic", "depth": 0.005}, "sheet_circle_closed"),
    Case(
        "create_louver", {"method": "basic", "depth": 0.003, "height": 0.005}, "sheet_circle_closed"
    ),
    Case(
        "create_lofted_flange",
        {"method": "basic", "thickness": 0.002, "bend_radius": 0.002},
        "sheet",
    ),
    Case("create_slot", {"method": "basic", "width": 0.004, "depth": 0.01}, "sheet_circle_closed"),
    Case("create_stamped", {"type": "bead", "depth": 0.002}, "sheet_line_closed"),
    Case("create_surface_mark", {"type": "etch", "face_indices": [0]}, "sheet_circle_closed"),
    Case(
        "sheet_metal_misc",
        {"action": "hem", "face_index": 0, "edge_index": 0, "hem_width": 0.005},
        "sheet",
    ),
    # query
    Case("query_body", {"property": "faces"}, "box"),
    Case("query_face", {"property": "normal", "face_index": 0}, "box"),
    Case("query_edge", {"property": "length", "face_index": 0, "edge_index": 0}, "box"),
    Case("query_bspline", {"type": "surface", "face_index": 0}, "box", "a plane is not a bspline"),
    Case(
        "measure",
        {"type": "distance", "x1": 0, "y1": 0, "z1": 0, "x2": 0.08, "y2": 0.048, "z2": 0.03},
        "box",
    ),
    Case("manage_variable", {"action": "add", "name": "W", "formula": "20 mm"}, "box"),
    Case("manage_property", {"action": "set_custom", "name": "Batch", "value": "X-1"}, "box"),
    Case("manage_material", {"action": "set_density", "density": 2700.0}, "box"),
    Case("manage_layer", {"action": "add", "name_or_index": "L1"}, "box"),
    Case("set_appearance", {"target": "body_color", "red": 200, "green": 30, "blue": 30}, "box"),
    Case("select_set", {"action": "all"}, "box"),
    Case("recompute", {"scope": "model"}, "box"),
    Case("manage_feature_tree", {"action": "rename", "feature_name": F, "new_name": "B"}, "box"),
    # view / display
    Case("camera_control", {"action": "zoom_fit"}, "box"),
    Case(
        "set_camera",
        {"eye_x": 1, "eye_y": -1, "eye_z": 1, "target_x": 0, "target_y": 0, "target_z": 0},
        "box",
    ),
    Case("display_control", {"action": "set_mode", "mode": "Shaded"}, "box"),
    Case(
        "export_file",
        {"format": "step", "file_path": str(TMP / "out.step"), "overwrite": True},
        "box",
    ),
    # assembly
    Case(
        "add_assembly_component",
        {"method": "basic", "file_path": str(BOX), "x": 0.2, "y": 0, "z": 0},
        "asm",
    ),
    Case("query_component", {"property": "list"}, "asm"),
    Case("query_component", {"property": "info", "component_index": 0}, "asm"),
    Case(
        "transform_component",
        {"method": "move", "component_index": 0, "dx": 0.01, "dy": 0, "dz": 0},
        "asm",
    ),
    Case(
        "rotate_component",
        {
            "component_index": 0,
            "axis_x1": 0,
            "axis_y1": 0,
            "axis_z1": 0,
            "axis_x2": 0,
            "axis_y2": 0,
            "axis_z2": 1,
            "angle": 30.0,
        },
        "asm",
    ),
    Case(
        "set_component_orientation",
        {"method": "put_euler", "component_index": 0, "angle_z": 30.0},
        "asm",
    ),
    Case(
        "set_component_appearance",
        {"property": "color", "component_index": 0, "red": 200, "green": 0, "blue": 0},
        "asm",
    ),
    Case("manage_component", {"action": "ground", "component_index": 0, "ground": True}, "asm"),
    Case("manage_relation", {"action": "list"}, "asm"),
    Case(
        "add_assembly_relation",
        {"type": "planar", "occurrence1_index": 0, "occurrence2_index": 1},
        "asm2",
    ),
    Case(
        "add_assembly_constraint",
        {"type": "mate", "component1_index": 0, "component2_index": 1},
        "asm2",
    ),
    Case(
        "assembly_feature",
        {
            "type": "extruded_cutout",
            "scope_parts": [0],
            "extent_type": "ThroughAll",
            "distance": 0.02,
        },
        "asm_circle_closed",
    ),
    Case("virtual_component", {"method": "new", "name": "VC1", "component_type": "Part"}, "asm"),
    Case(
        "structural_frame",
        {"method": "basic", "part_filename": str(BOX), "path_indices": [0]},
        "asm",
    ),
    Case("wiring", {"type": "wire", "path_indices": [0]}, "asm"),
    # draft
    Case("manage_sheet", {"action": "add"}, "draft"),
    Case("query_sheet", {"type": "dimensions"}, "draft"),
    Case(
        "draw_sheet_geometry",
        {"shape": "line", "x1": 0.02, "y1": 0.02, "x2": 0.12, "y2": 0.02},
        "draft",
    ),
    Case("add_annotation", {"type": "text_box", "x": 0.05, "y": 0.05, "text": "Hello"}, "draft"),
    Case(
        "add_2d_dimension",
        {
            "type": "distance",
            "x1": 0.02,
            "y1": 0.02,
            "x2": 0.12,
            "y2": 0.02,
            "x3": 0.07,
            "y3": 0.04,
        },
        "draft_line",
    ),
    Case(
        "add_dimension_annotation",
        {
            "type": "dimension",
            "x1": 0.02,
            "y1": 0.02,
            "x2": 0.12,
            "y2": 0.02,
            "dim_x": 0.07,
            "dim_y": 0.04,
        },
        "draft_line",
    ),
    Case("add_symbol_annotation", {"type": "center_mark", "x": 0.05, "y": 0.05}, "draft"),
    Case(
        "add_smart_frame",
        {"method": "two_point", "style_name": "A4", "x1": 0.01, "y1": 0.01, "x2": 0.2, "y2": 0.15},
        "draft",
    ),
    Case("manage_annotation_data", {"action": "get_symbols"}, "draft"),
    Case("print_control", {"action": "get_printer"}, "draft"),
    Case("draft_config", {"action": "get_origin"}, "draft"),
    Case(
        "manage_sheet",
        {"action": "create_drawing", "views": ["Front", "Top"]},
        "draft_link",
        "the context already made one",
    ),
    Case(
        "add_drawing_view",
        {"type": "part", "orientation": "Isometric", "x": 0.15, "y": 0.1},
        "draft_link",
    ),
    Case(
        "manage_drawing_view", {"action": "set_scale", "view_index": 0, "scale": 0.5}, "draft_link"
    ),
    Case("create_table", {"type": "parts_list"}, "draft_link"),
    Case("create_table", {"type": "bend"}, "draft_link"),
    Case(
        "export_file",
        {"format": "pdf", "file_path": str(TMP / "out.pdf"), "overwrite": True},
        "draft_link",
    ),
]


# --------------------------------------------------------------------- run
def unwrap(res: Any) -> Any:
    if getattr(res, "data", None) is not None:
        return res.data
    try:
        return json.loads(res.content[0].text)
    except Exception:
        return res.content[0].text if res.content else None


def classify(payload: Any) -> tuple[str, str]:
    if isinstance(payload, dict) and "error" in payload:
        msg = str(payload["error"])
        if payload.get("unsupported"):
            return "UNSUP", msg
        if "reported success but" in msg or "added nothing" in msg or "changed no geometry" in msg:
            return "NOOP", msg
        return "FAIL", msg
    return "OK", json.dumps(payload)[:160] if payload is not None else ""


@dataclass
class Row:
    tool: str
    args: str
    ctx: str
    outcome: str
    detail: str
    note: str = ""
    modals: list[str] = field(default_factory=list)


def write_report(rows: list[Row], started: dt.datetime, version: str, done: bool) -> None:
    counts: dict[str, int] = {}
    for r in rows:
        counts[r.outcome] = counts.get(r.outcome, 0) + 1
    tools_hit = len({r.tool for r in rows})
    lines = [
        "# Live sweep",
        "",
        f"**{'Completed' if done else 'IN PROGRESS'}** {started:%Y-%m-%d %H:%M} "
        f"against Solid Edge {version}.",
        "Generated by `scripts/live_sweep.py`; regenerate rather than edit.",
        "",
        f"{len(rows)} cases across {tools_hit} of 118 tools. "
        + ", ".join(f"**{k}** {v}" for k, v in sorted(counts.items()))
        + ".",
        "",
        "| outcome | meaning |",
        "|---|---|",
        "| OK | reported success |",
        "| UNSUP | refused honestly with `unsupported: True` |",
        "| NOOP | claimed success; a verification decorator found nothing changed |",
        "| FAIL | returned an error |",
        "| SETUP | the context it needed could not be built |",
        "| EXC | raised or timed out on the client side |",
        "",
        "| tool | args | context | outcome | detail |",
        "|---|---|---|---|---|",
    ]
    for r in rows:
        detail = r.detail.replace("|", "\\|").replace("\n", " ")[:150]
        note = f" *({r.note})*" if r.note else ""
        modal = f" **modal:** {', '.join(r.modals)}" if r.modals else ""
        lines.append(
            f"| `{r.tool}` | `{r.args}` | {r.ctx} | **{r.outcome}** | {detail}{note}{modal} |"
        )
    if modals:
        lines += ["", f"Modal dialogs dismissed during the run: {len(modals)}."]
    OUT_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")
    OUT_JSON.write_text(json.dumps([r.__dict__ for r in rows], indent=1), encoding="utf-8")


async def main() -> int:
    started = dt.datetime.now()
    transport = StdioTransport(command="uv", args=["run", "solidedge-mcp"], cwd=str(REPO))
    rows: list[Row] = []
    async with Client(transport) as c:

        async def raw(tool: str, **kw: Any) -> Any:
            try:
                return unwrap(await asyncio.wait_for(c.call_tool(tool, kw), timeout=CALL_TIMEOUT))
            except Exception as exc:  # noqa: BLE001
                return {
                    "error": f"{type(exc).__name__}: {str(exc)[:160]}",
                    "_client_exception": True,
                }

        conn = await raw("manage_connection", action="connect", start_if_needed=False)
        if not isinstance(conn, dict) or conn.get("status") != "connected":
            print("SOLID EDGE NOT RUNNING:", conn)
            return 2
        version = str(conn.get("version", "?"))

        async def close_all() -> None:
            # discard_unsaved is required: every scratch document is dirty, and
            # close_all_documents rightly refuses to throw away unsaved work
            # without it. Omitting it made each close after the first dirty
            # document a no-op, so documents piled up across the run.
            await raw("close_document", scope="all", save=False, discard_unsaved=True)

        # the saved part every assembly and draft case needs
        await close_all()
        for tool, args in _box():
            await raw(tool, **args)
        saved = await raw("save_document", method="save", file_path=str(BOX))
        await close_all()
        if not BOX.exists():
            print("could not save the box part:", saved)
            return 2
        print(f"box saved: {BOX}", flush=True)

        for i, case in enumerate(CASES, 1):
            seen_before = len(modals)
            await close_all()
            setup_failed = None
            for tool, args in CONTEXTS[case.ctx]:
                r = await raw(tool, **args)
                if isinstance(r, dict) and "error" in r:
                    setup_failed = f"{tool}: {str(r['error'])[:120]}"
                    break
            args_str = json.dumps(case.args, default=str)[:110]
            if setup_failed:
                row = Row(case.tool, args_str, case.ctx, "SETUP", setup_failed, case.note)
            else:
                payload = await raw(case.tool, **case.args)
                if isinstance(payload, dict) and payload.get("_client_exception"):
                    outcome, detail = "EXC", str(payload["error"])
                else:
                    outcome, detail = classify(payload)
                row = Row(case.tool, args_str, case.ctx, outcome, detail, case.note)
            row.modals = modals[seen_before:]
            rows.append(row)
            print(f"  {i:3d}/{len(CASES)} [{row.outcome:5s}] {case.tool} ({case.ctx})", flush=True)
            write_report(rows, started, version, done=False)

        await close_all()
    write_report(rows, started, version, done=True)
    print(f"\nwrote {OUT_MD}")
    return 0


if __name__ == "__main__":
    threading.Thread(target=_watchdog, daemon=True).start()
    try:
        code = asyncio.run(main())
    finally:
        _stop.set()
    sys.exit(code)
