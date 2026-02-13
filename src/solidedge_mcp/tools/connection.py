"""Connection tools for Solid Edge MCP."""

from solidedge_mcp.managers import connection

def connect_to_solidedge(start_if_needed: bool = True) -> dict:
    """Connect to Solid Edge application. Starts the application if not running and `start_if_needed` is True."""
    return connection.connect(start_if_needed)

def quit_application() -> dict:
    """Quit the Solid Edge application. Closes all documents and shuts down Solid Edge completely."""
    return connection.quit_application()

def get_application_info() -> dict:
    """Get Solid Edge application information (version, path, document count)."""
    return connection.get_info()

def disconnect_from_solidedge() -> dict:
    """Disconnect from the Solid Edge application without closing it."""
    return connection.disconnect()

def is_connected() -> dict:
    """Check if currently connected to Solid Edge."""
    return {"connected": connection.is_connected()}

def get_process_info() -> dict:
    """Get Solid Edge process information (PID, window handle)."""
    return connection.get_process_info()

def get_install_info() -> dict:
    """Get Solid Edge installation information (path, language, version)."""
    return connection.get_install_info()

def start_command(command_id: int) -> dict:
    """Execute a Solid Edge command by its ID."""
    return connection.start_command(command_id)

def set_performance_mode(delay_compute: bool = None, screen_updating: bool = None, interactive: bool = None, display_alerts: bool = None) -> dict:
    """Set application performance flags for batch operations."""
    return connection.set_performance_mode(delay_compute, screen_updating, interactive, display_alerts)

def do_idle() -> dict:
    """Process pending background operations."""
    return connection.do_idle()

def activate_application() -> dict:
    """Activate (bring to foreground) the Solid Edge application window."""
    return connection.activate_application()

def abort_command(abort_all: bool = True) -> dict:
    """Abort the current Solid Edge command."""
    return connection.abort_command(abort_all)

def get_active_environment() -> dict:
    """Get the currently active environment (Part, Assembly, Draft, etc.)."""
    return connection.get_active_environment()

def get_status_bar() -> dict:
    """Get the current status bar text."""
    return connection.get_status_bar()

def set_status_bar(text: str) -> dict:
    """Set the status bar text."""
    return connection.set_status_bar(text)

def get_visible() -> dict:
    """Get the visibility state of the Solid Edge window."""
    return connection.get_visible()

def set_visible(visible: bool) -> dict:
    """Set the visibility of the Solid Edge window."""
    return connection.set_visible(visible)

def register(mcp):
    """Register connection tools with the MCP server."""
    mcp.tool()(connect_to_solidedge)
    mcp.tool()(quit_application)
    mcp.tool()(get_application_info)
    mcp.tool()(disconnect_from_solidedge)
    mcp.tool()(is_connected)
    mcp.tool()(get_process_info)
    mcp.tool()(get_install_info)
    mcp.tool()(start_command)
    mcp.tool()(set_performance_mode)
    mcp.tool()(do_idle)
    mcp.tool()(activate_application)
    mcp.tool()(abort_command)
    mcp.tool()(get_active_environment)
    mcp.tool()(get_status_bar)
    mcp.tool()(set_status_bar)
    mcp.tool()(get_visible)
    mcp.tool()(set_visible)
