"""Connection tools for Solid Edge MCP."""

from typing import Any, Literal

from solidedge_mcp.managers import connection
from solidedge_mcp.tools._registry import register_tool

# === Composite: manage_connection ===


def manage_connection(
    action: Literal["connect", "disconnect", "quit", "activate"] = "connect",
    start_if_needed: bool = True,
) -> dict[str, Any]:
    """Manage the Solid Edge application connection.

    connect: attach to a running instance (start_if_needed launches one).
    disconnect: drop the COM reference. quit: exit Solid Edge (unsaved work
    may be lost). activate: bring the window to the foreground.
    """
    match action:
        case "connect":
            return connection.connect(start_if_needed)
        case "disconnect":
            return connection.disconnect()
        case "quit":
            return connection.quit_application()
        case "activate":
            return connection.activate_application()
        case _:
            return {"error": f"Unknown action: {action}"}


# === Composite: app_command ===


def app_command(
    action: Literal["start", "abort", "idle"],
    command_id: int = 0,
    abort_all: bool = True,
) -> dict[str, Any]:
    """Drive the Solid Edge command engine.

    start: run command_id (Solid Edge command constant).
    abort: cancel the running command (abort_all cancels nested ones too).
    idle: let Solid Edge process pending events.
    """
    match action:
        case "start":
            return connection.start_command(command_id)
        case "abort":
            return connection.abort_command(abort_all)
        case "idle":
            return connection.do_idle()
        case _:
            return {"error": f"Unknown action: {action}"}


# === Composite: app_config ===


def app_config(
    property: Literal[
        "set_performance",
        "get_environment",
        "get_status_bar",
        "set_status_bar",
        "get_visible",
        "set_visible",
        "get_global",
        "set_global",
        "get_template",
        "set_template",
    ],
    delay_compute: bool | None = None,
    screen_updating: bool | None = None,
    interactive: bool | None = None,
    display_alerts: bool | None = None,
    text: str = "",
    visible: bool = True,
    parameter: int = 0,
    value: float = 0.0,
    doc_type: int = 1,
    template_path: str = "",
) -> dict[str, Any]:
    """Get or set application-level settings.

    set_performance: delay_compute/screen_updating/interactive/display_alerts
    (None leaves a flag unchanged). set_status_bar: text. set_visible: visible.
    get_global/set_global: parameter (ApplicationGlobalConstants) [+ value].
    get_template/set_template: doc_type 1=Part, 2=Draft, 3=Assembly,
    4=SheetMetal [+ template_path].
    """
    match property:
        case "set_performance":
            return connection.set_performance_mode(
                delay_compute,
                screen_updating,
                interactive,
                display_alerts,
            )
        case "get_environment":
            return connection.get_active_environment()
        case "get_status_bar":
            return connection.get_status_bar()
        case "set_status_bar":
            return connection.set_status_bar(text)
        case "get_visible":
            return connection.get_visible()
        case "set_visible":
            return connection.set_visible(visible)
        case "get_global":
            return connection.get_global_parameter(parameter)
        case "set_global":
            return connection.set_global_parameter(parameter, value)
        case "get_template":
            return connection.get_default_template_path(doc_type)
        case "set_template":
            return connection.set_default_template_path(doc_type, template_path)
        case _:
            return {"error": f"Unknown property: {property}"}


# === Standalone tools ===


def convert_by_file_path(input_path: str, output_path: str) -> dict[str, Any]:
    """Convert a CAD file to another format by extension (e.g. .par -> .step)."""
    return connection.convert_by_file_path(input_path, output_path)


def arrange_windows(style: int = 1) -> dict[str, Any]:
    """Arrange open document windows. style: 1=Tiled, 2=Horizontal, 4=Vertical, 8=Cascade."""
    return connection.arrange_windows(style)


def get_active_command() -> dict[str, Any]:
    """Return the currently active Solid Edge command (read-only)."""
    return connection.get_active_command()


def run_macro(filename: str) -> dict[str, Any]:
    """Run a VBA macro file (.vba/.exe path) in Solid Edge."""
    return connection.run_macro(filename)


# === Registration ===


def register(mcp: Any) -> None:
    """Register connection tools with the MCP server."""
    tags = {"app"}
    # Composite tools
    register_tool(mcp, manage_connection, tags=tags, destructive=True)
    register_tool(mcp, app_command, tags=tags)
    register_tool(mcp, app_config, tags=tags)
    # Standalone tools
    register_tool(mcp, convert_by_file_path, tags=tags)
    register_tool(mcp, arrange_windows, tags=tags, idempotent=True)
    register_tool(mcp, get_active_command, tags=tags, read_only=True)
    register_tool(mcp, run_macro, tags=tags)
