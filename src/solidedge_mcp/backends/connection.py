"""
Solid Edge Connection Management

Handles connecting to and managing Solid Edge application instances.
"""

import contextlib
import traceback
from typing import Any

import win32com.client

from .logging import get_logger

_logger = get_logger(__name__)


class SolidEdgeConnection:
    """Manages connection to Solid Edge application"""

    def __init__(self):
        self.application: Any | None = None
        self._is_connected: bool = False

    def connect(self, start_if_needed: bool = True) -> dict[str, Any]:
        """
        Connect to Solid Edge application instance.

        Args:
            start_if_needed: If True, start Solid Edge if not running

        Returns:
            Dict with connection status and info
        """
        try:
            if self.application is None:
                try:
                    # Try to connect to existing instance
                    self.application = win32com.client.GetActiveObject("SolidEdge.Application")
                    _logger.info("Connected to existing Solid Edge instance")
                except Exception:
                    if start_if_needed:
                        # Start new instance with early binding if possible
                        try:
                            self.application = win32com.client.gencache.EnsureDispatch(
                                "SolidEdge.Application"
                            )
                        except Exception:
                            # Fall back to late binding
                            self.application = win32com.client.Dispatch("SolidEdge.Application")

                        self.application.Visible = True
                        _logger.info("Started new Solid Edge instance")
                    else:
                        raise Exception(
                            "No Solid Edge instance found and start_if_needed=False"
                        ) from None

            self._is_connected = True

            # Get version info
            version = self.application.Version

            return {
                "status": "connected",
                "version": version,
                "visible": self.application.Visible,
                "caption": self.application.Caption,
            }
        except Exception as e:
            self._is_connected = False
            _logger.error(f"Connection failed: {e}")
            return {"status": "error", "message": str(e), "traceback": traceback.format_exc()}

    def disconnect(self) -> dict[str, Any]:
        """Disconnect from Solid Edge (does not close the application)"""
        self.application = None
        self._is_connected = False
        _logger.info("Disconnected from Solid Edge")
        return {"status": "disconnected"}

    def get_info(self) -> dict[str, Any]:
        """Get information about the connected Solid Edge instance"""
        if not self._is_connected or self.application is None:
            return {"error": "Not connected to Solid Edge"}

        try:
            info = {
                "version": self.application.Version,
                "caption": self.application.Caption,
                "visible": self.application.Visible,
                "documents_count": self.application.Documents.Count,
            }

            # Path property may not exist in all Solid Edge versions
            try:
                info["path"] = self.application.Path
            except Exception:
                info["path"] = "N/A"

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_application_info(self) -> dict[str, Any]:
        """Alias for get_info() for consistency with MCP tool name"""
        return self.get_info()

    def is_connected(self) -> bool:
        """Check if connected to Solid Edge"""
        return self._is_connected and self.application is not None

    def ensure_connected(self) -> None:
        """Ensure connection exists, raise exception if not"""
        if not self.is_connected():
            raise Exception("Not connected to Solid Edge. Call connect() first.")

    def get_application(self):
        """Get the application object"""
        self.ensure_connected()
        return self.application

    def quit_application(self) -> dict[str, Any]:
        """
        Quit the Solid Edge application.

        Closes all documents and shuts down Solid Edge.

        Returns:
            Dict with quit status
        """
        try:
            if not self._is_connected or self.application is None:
                return {"error": "Not connected to Solid Edge"}

            self.application.Quit()
            self.application = None
            self._is_connected = False

            _logger.info("Solid Edge application quit")
            return {"status": "quit", "message": "Solid Edge has been closed"}
        except Exception as e:
            self.application = None
            self._is_connected = False
            _logger.error(f"Failed to quit Solid Edge: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_process_info(self) -> dict[str, Any]:
        """
        Get Solid Edge process information (PID, window handle).

        Returns:
            Dict with process_id and window_handle
        """
        try:
            self.ensure_connected()
            app = self.application

            info = {}
            try:
                info["process_id"] = app.ProcessID
            except Exception:
                info["process_id"] = None

            try:
                info["window_handle"] = app.hWnd
            except Exception:
                info["window_handle"] = None

            return {"status": "success", **info}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_install_info(self) -> dict[str, Any]:
        """
        Get Solid Edge installation information (path, language).

        Uses the SEInstallData COM library to read install location and language.

        Returns:
            Dict with install_path and language
        """
        try:
            info = {}

            try:
                install_data = win32com.client.Dispatch("SEInstallDataLib.SEInstallData")
                with contextlib.suppress(Exception):
                    info["install_path"] = install_data.GetInstalledPath()
                with contextlib.suppress(Exception):
                    info["language"] = install_data.GetInstalledLanguage()
                with contextlib.suppress(Exception):
                    info["version"] = install_data.GetInstalledVersion()
            except Exception:
                # SEInstallData may not be registered; fall back to Application.Path
                if self._is_connected and self.application is not None:
                    with contextlib.suppress(Exception):
                        info["install_path"] = self.application.Path

            if not info:
                return {
                    "error": "Could not retrieve installation"
                    " info. SEInstallData COM library may "
                    "not be registered."
                }

            return {"status": "success", **info}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_performance_mode(
        self,
        delay_compute: bool = None,
        screen_updating: bool = None,
        interactive: bool = None,
        display_alerts: bool = None,
    ) -> dict[str, Any]:
        """
        Set application performance flags for batch operations.

        These flags can significantly speed up batch operations by disabling
        UI updates and delayed computation. Remember to restore defaults after.

        Args:
            delay_compute: If True, delays feature recomputation until reset
            screen_updating: If False, disables screen refreshes
            interactive: If False, suppresses all UI dialogs
            display_alerts: If False, suppresses alert dialogs

        Returns:
            Dict with status and current settings
        """
        try:
            self.ensure_connected()
            app = self.application

            settings = {}

            if delay_compute is not None:
                try:
                    app.DelayCompute = delay_compute
                    settings["delay_compute"] = delay_compute
                except Exception as e:
                    settings["delay_compute_error"] = str(e)

            if screen_updating is not None:
                try:
                    app.ScreenUpdating = screen_updating
                    settings["screen_updating"] = screen_updating
                except Exception as e:
                    settings["screen_updating_error"] = str(e)

            if interactive is not None:
                try:
                    app.Interactive = interactive
                    settings["interactive"] = interactive
                except Exception as e:
                    settings["interactive_error"] = str(e)

            if display_alerts is not None:
                try:
                    app.DisplayAlerts = display_alerts
                    settings["display_alerts"] = display_alerts
                except Exception as e:
                    settings["display_alerts_error"] = str(e)

            return {"status": "updated", "settings": settings}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def start_command(self, command_id: int) -> dict[str, Any]:
        """
        Execute a Solid Edge command by its command ID.

        Invokes Application.StartCommand(CommandID) to programmatically trigger
        any Solid Edge menu/ribbon command. Command IDs are from the
        SolidEdgeCommandConstants enum in the type library.

        Args:
            command_id: Integer command ID (from SolidEdgeCommandConstants)

        Returns:
            Dict with status
        """
        try:
            self.ensure_connected()
            self.application.StartCommand(command_id)
            return {"status": "success", "command_id": command_id}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def do_idle(self) -> dict[str, Any]:
        """
        Allow Solid Edge to process pending operations.

        Calls Application.DoIdle() to give Solid Edge a chance to complete
        background processing. Useful after batch operations or before
        querying results that depend on recomputation.

        Returns:
            Dict with status
        """
        try:
            self.ensure_connected()
            self.application.DoIdle()
            return {"status": "success"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def activate_application(self) -> dict[str, Any]:
        """
        Activate (bring to foreground) the Solid Edge application window.

        Uses Application.Activate().

        Returns:
            Dict with status
        """
        try:
            self.ensure_connected()
            self.application.Activate()
            return {"status": "activated"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def abort_command(self, abort_all: bool = True) -> dict[str, Any]:
        """
        Abort the current Solid Edge command.

        Uses Application.AbortCommand(AbortAll).

        Args:
            abort_all: If True, aborts all pending commands. If False,
                       aborts only the most recent command.

        Returns:
            Dict with status
        """
        try:
            self.ensure_connected()
            self.application.AbortCommand(abort_all)
            return {"status": "aborted", "abort_all": abort_all}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_active_environment(self) -> dict[str, Any]:
        """
        Get the currently active environment in Solid Edge.

        The active environment determines which commands and menus
        are available (e.g., Part, Assembly, Draft).

        Returns:
            Dict with environment info
        """
        try:
            self.ensure_connected()
            env = self.application.ActiveEnvironment

            result = {"status": "success"}
            try:
                result["name"] = env.Name
            except Exception:
                result["name"] = str(env)
            with contextlib.suppress(Exception):
                result["caption"] = env.Caption

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_status_bar(self) -> dict[str, Any]:
        """
        Get the current status bar text.

        Returns:
            Dict with status bar text
        """
        try:
            self.ensure_connected()
            text = self.application.StatusBar
            return {"status": "success", "text": text}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_status_bar(self, text: str) -> dict[str, Any]:
        """
        Set the status bar text.

        Args:
            text: Text to display in the status bar

        Returns:
            Dict with status
        """
        try:
            self.ensure_connected()
            self.application.StatusBar = text
            return {"status": "set", "text": text}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_visible(self) -> dict[str, Any]:
        """
        Get the visibility state of the Solid Edge application window.

        Returns:
            Dict with visible state
        """
        try:
            self.ensure_connected()
            visible = self.application.Visible
            return {"status": "success", "visible": visible}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_visible(self, visible: bool) -> dict[str, Any]:
        """
        Set the visibility of the Solid Edge application window.

        Args:
            visible: True to show, False to hide

        Returns:
            Dict with status
        """
        try:
            self.ensure_connected()
            self.application.Visible = visible
            return {"status": "set", "visible": visible}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_global_parameter(self, parameter: int) -> dict[str, Any]:
        """
        Get an application-level global parameter.

        Uses Application.GetGlobalParameter(param). Parameter IDs come from
        the AssemblyGlobalConstants enum (1-21).

        Args:
            parameter: Global parameter ID (from AssemblyGlobalConstants)

        Returns:
            Dict with parameter value
        """
        try:
            self.ensure_connected()
            value = self.application.GetGlobalParameter(parameter)
            return {"status": "success", "parameter": parameter, "value": value}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_global_parameter(self, parameter: int, value) -> dict[str, Any]:
        """
        Set an application-level global parameter.

        Uses Application.SetGlobalParameter(param, value). Parameter IDs come
        from the AssemblyGlobalConstants enum (1-21).

        Args:
            parameter: Global parameter ID (from AssemblyGlobalConstants)
            value: New value for the parameter

        Returns:
            Dict with status
        """
        try:
            self.ensure_connected()
            self.application.SetGlobalParameter(parameter, value)
            return {"status": "set", "parameter": parameter, "value": value}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def convert_by_file_path(
        self, input_path: str, output_path: str
    ) -> dict[str, Any]:
        """
        Batch-convert CAD files between formats.

        Uses Application.ConvertByFilePath to convert individual files or
        entire folders. The format is determined by file extensions.

        Args:
            input_path: Input file or folder path
            output_path: Output file or folder path

        Returns:
            Dict with status
        """
        try:
            self.ensure_connected()
            self.application.ConvertByFilePath(input_path, output_path)
            return {
                "status": "converted",
                "input": input_path,
                "output": output_path,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_default_template_path(self, doc_type: int) -> dict[str, Any]:
        """
        Get the default template file path for a document type.

        Args:
            doc_type: Document type constant (1=Part, 2=Draft, 3=Assembly, 4=SheetMetal)

        Returns:
            Dict with template path
        """
        try:
            self.ensure_connected()
            path = self.application.GetDefaultTemplatePath(doc_type)
            return {"status": "success", "doc_type": doc_type, "template_path": path}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_default_template_path(self, doc_type: int, template_path: str) -> dict[str, Any]:
        """
        Set the default template file path for a document type.

        Args:
            doc_type: Document type constant (1=Part, 2=Draft, 3=Assembly, 4=SheetMetal)
            template_path: Path to the template file

        Returns:
            Dict with status
        """
        try:
            self.ensure_connected()
            self.application.SetDefaultTemplatePath(doc_type, template_path)
            return {
                "status": "set",
                "doc_type": doc_type,
                "template_path": template_path,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def arrange_windows(self, style: int = 1) -> dict[str, Any]:
        """
        Arrange document windows in the Solid Edge application.

        Args:
            style: Window arrangement style
                   (1=Tiled, 2=Horizontal, 4=Vertical, 8=Cascade)

        Returns:
            Dict with status
        """
        try:
            self.ensure_connected()
            style_names = {1: "Tiled", 2: "Horizontal", 4: "Vertical", 8: "Cascade"}
            self.application.ArrangeWindows(style)
            return {
                "status": "arranged",
                "style": style,
                "style_name": style_names.get(style, f"Unknown({style})"),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_active_command(self) -> dict[str, Any]:
        """
        Get the currently active command in Solid Edge.

        Returns:
            Dict with active command info
        """
        try:
            self.ensure_connected()
            cmd = self.application.ActiveCommand
            result: dict[str, Any] = {"status": "success"}
            if cmd is not None:
                result["has_active_command"] = True
                with contextlib.suppress(Exception):
                    result["name"] = cmd.Name
                with contextlib.suppress(Exception):
                    result["id"] = cmd.ID
            else:
                result["has_active_command"] = False
            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def run_macro(self, filename: str) -> dict[str, Any]:
        """
        Run a VBA macro file in Solid Edge.

        Args:
            filename: Path to the VBA macro file (.bas, .exe, etc.)

        Returns:
            Dict with status
        """
        try:
            self.ensure_connected()
            self.application.RunMacro(filename)
            return {"status": "executed", "filename": filename}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
