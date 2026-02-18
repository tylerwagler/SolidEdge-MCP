"""Connection tools for Solid Edge MCP."""

from solidedge_mcp.managers import connection

# === Composite: manage_connection ===


def manage_connection(
    action: str = "connect",
    start_if_needed: bool = True,
) -> dict:
    """Manage the Solid Edge application connection.

    action: 'connect' | 'disconnect' | 'quit' | 'activate'

    - connect: Connect (start if needed when start_if_needed=True)
    - disconnect: Disconnect without closing Solid Edge
    - quit: Quit the Solid Edge application
    - activate: Bring the Solid Edge window to foreground
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
    action: str,
    command_id: int = 0,
    abort_all: bool = True,
) -> dict:
    """Execute an application command.

    action: 'start' | 'abort' | 'idle'

    - start: Execute a command by its ID
    - abort: Abort the current command
    - idle: Process pending background operations
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
    property: str,
    delay_compute: bool = None,
    screen_updating: bool = None,
    interactive: bool = None,
    display_alerts: bool = None,
    text: str = "",
    visible: bool = True,
    parameter: int = 0,
    value: float = 0.0,
    doc_type: int = 1,
    template_path: str = "",
) -> dict:
    """Get or set application configuration properties.

    property: 'set_performance' | 'get_environment'
      | 'get_status_bar' | 'set_status_bar'
      | 'get_visible' | 'set_visible'
      | 'get_global' | 'set_global'
      | 'get_template' | 'set_template'

    - get_template/set_template: doc_type 1=Part, 2=Draft,
      3=Assembly, 4=SheetMetal
    """
    match property:
        case "set_performance":
            return connection.set_performance_mode(
                delay_compute, screen_updating,
                interactive, display_alerts,
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
            return connection.set_default_template_path(
                doc_type, template_path
            )
        case _:
            return {"error": f"Unknown property: {property}"}


# === Standalone tools ===


def convert_by_file_path(input_path: str, output_path: str) -> dict:
    """Batch-convert CAD files between formats."""
    return connection.convert_by_file_path(input_path, output_path)


def arrange_windows(style: int = 1) -> dict:
    """Arrange document windows.

    style: 1=Tiled, 2=Horizontal, 4=Vertical, 8=Cascade
    """
    return connection.arrange_windows(style)


def get_active_command() -> dict:
    """Get the currently active Solid Edge command."""
    return connection.get_active_command()


def run_macro(filename: str) -> dict:
    """Run a VBA macro file in Solid Edge."""
    return connection.run_macro(filename)


# === Registration ===


def register(mcp):
    """Register connection tools with the MCP server."""
    # Composite tools
    mcp.tool()(manage_connection)
    mcp.tool()(app_command)
    mcp.tool()(app_config)
    # Standalone tools
    mcp.tool()(convert_by_file_path)
    mcp.tool()(arrange_windows)
    mcp.tool()(get_active_command)
    mcp.tool()(run_macro)
