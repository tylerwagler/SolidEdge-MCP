"""
Solid Edge Connection Management

Handles connecting to and managing Solid Edge application instances.
"""

import contextlib
from collections.abc import Callable
from typing import Any

import win32com.client

from solidedge_mcp.backends.errors import error_result, is_disconnected_error

from .comutil import com_get
from .logging import get_logger

_logger = get_logger(__name__)

_DEAD_MARKERS = (
    "RPC server is unavailable",
    "object invoked has disconnected",
    "-2147023174",
    "-2147417848",
)


def _looks_dead(exc: BaseException) -> bool:
    """Heuristic for dead-proxy errors that are not typed com_error."""
    text = str(exc)
    return any(m in text for m in _DEAD_MARKERS)


class SolidEdgeConnection:
    """Manages connection to Solid Edge application"""

    def __init__(self) -> None:
        self.application: Any | None = None
        self._is_connected: bool = False
        # Called when a dead connection is detected so managers can drop
        # cached document/profile pointers. Set by DocumentManager.
        self.on_disconnect: Callable[[], None] | None = None

    def _is_alive(self) -> bool:
        """Cheap liveness probe of the cached Application proxy."""
        if self.application is None:
            return False
        try:
            _ = self.application.Version
            return True
        except Exception as exc:
            return not is_disconnected_error(exc) and not _looks_dead(exc)

    def _notify_disconnect(self) -> None:
        if self.on_disconnect is not None:
            try:
                self.on_disconnect()
            except Exception as exc:  # never let cleanup mask the real error
                _logger.debug(f"on_disconnect callback failed: {exc}")

    def connect(self, start_if_needed: bool = True) -> dict[str, Any]:
        """
        Connect to Solid Edge application instance.

        Args:
            start_if_needed: If True, start Solid Edge if not running

        Returns:
            Dict with connection status and info
        """
        try:
            # Drop a cached proxy whose server has gone away (SE closed/crashed)
            if self.application is not None and not self._is_alive():
                _logger.warning("Cached Solid Edge proxy is dead; reconnecting")
                self.application = None
                self._is_connected = False
                self._notify_disconnect()

            if self.application is None:
                try:
                    # Try to connect to existing instance
                    self.application = win32com.client.GetActiveObject("SolidEdge.Application")
                    _logger.info("Connected to existing Solid Edge instance")
                except Exception:
                    if start_if_needed:
                        # Late binding everywhere so object semantics never differ
                        # between "attached to running" and "started by us".
                        self.application = win32com.client.Dispatch("SolidEdge.Application")
                        self.application.Visible = True
                        _logger.info("Started new Solid Edge instance")
                    else:
                        raise Exception(
                            "No Solid Edge instance found and start_if_needed=False"
                        ) from None

            self._is_connected = True

            # Solid Edge raises modal dialogs for things an automation client
            # cannot answer: a failed STEP translation, an overwrite prompt, a
            # rebuild warning. The dialog blocks the COM call that triggered it
            # for as long as it is on screen, which hangs this server with no
            # error and no timeout. Suppressing alerts turns those into ordinary
            # COM failures we can report. Observed with export_file on a part
            # with no solid body: "could not be saved because of a file
            # translation error" sat there until dismissed by hand.
            try:
                self.application.DisplayAlerts = False
            except Exception as exc:  # not fatal; older builds may not expose it
                _logger.debug(f"Could not disable DisplayAlerts: {exc}")

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
            return {"status": "error", **error_result(e)}

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

            # Application has no Path property on Solid Edge 2026, so this
            # always fell through to "N/A". AppDataFolder and RegistryPath are
            # what it does report about where it lives.
            for key, member in (
                ("app_data_folder", "AppDataFolder"),
                ("registry_path", "RegistryPath"),
            ):
                value = com_get(self.application, member)
                if value is not None:
                    info[key] = value

            return info
        except Exception as e:
            return error_result(e)

    def get_application_info(self) -> dict[str, Any]:
        """Alias for get_info() for consistency with MCP tool name"""
        return self.get_info()

    def is_connected(self) -> bool:
        """Check if connected to Solid Edge (flag only; no COM round trip)."""
        return self._is_connected and self.application is not None

    def check_connection(self) -> bool:
        """Probe the live connection; drops a dead proxy and returns False if gone."""
        if not self.is_connected():
            return False
        if self._is_alive():
            return True
        _logger.warning("Solid Edge connection lost")
        self.application = None
        self._is_connected = False
        self._notify_disconnect()
        return False

    def ensure_connected(self) -> None:
        """Ensure connection exists, raise exception if not"""
        if not self.is_connected():
            raise Exception(
                "Not connected to Solid Edge. Call manage_connection(action='connect') first."
            )

    def get_application(self) -> Any:
        """Get the application object"""
        self.ensure_connected()
        return self.application

    def _get_app(self) -> Any:
        """Return the connected application, or raise if not connected."""
        app = self.application
        if app is None:
            raise RuntimeError("Not connected to Solid Edge. Call connect() first.")
        return app

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
            return error_result(e)

    def get_process_info(self) -> dict[str, Any]:
        """
        Get Solid Edge process information (PID, window handle).

        Returns:
            Dict with process_id and window_handle
        """
        try:
            app = self._get_app()

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
            return error_result(e)

    def get_install_info(self) -> dict[str, Any]:
        """
        Get Solid Edge installation information (path, language).

        Uses the SEInstallData COM library to read install location and language.

        Returns:
            Dict with install_path and language
        """
        try:
            info = {}

            # SEInstallDataLib.SEInstallData is not registered on Solid Edge
            # 2026 -- Dispatch raises "Invalid class string" -- and neither
            # GetInstalledLanguage nor GetInstalledVersion is in any scraped
            # type library. Application.Path does not exist either, so the
            # fallback was dead too. These four do work.
            if self._is_connected and self.application is not None:
                for key, member in (
                    ("version", "Version"),
                    ("name", "Name"),
                    ("app_data_folder", "AppDataFolder"),
                    ("registry_path", "RegistryPath"),
                ):
                    value = com_get(self.application, member)
                    if value is not None:
                        info[key] = value

            if not info:
                return {
                    "error": "Could not retrieve installation"
                    " info. SEInstallData COM library may "
                    "not be registered."
                }

            return {"status": "success", **info}
        except Exception as e:
            return error_result(e)

    def set_performance_mode(
        self,
        delay_compute: bool | None = None,
        screen_updating: bool | None = None,
        interactive: bool | None = None,
        display_alerts: bool | None = None,
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
            app = self._get_app()

            settings: dict[str, Any] = {}

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
            return error_result(e)

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
            app = self._get_app()
            app.StartCommand(command_id)
            return {"status": "success", "command_id": command_id}
        except Exception as e:
            return error_result(e)

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
            app = self._get_app()
            app.DoIdle()
            return {"status": "success"}
        except Exception as e:
            return error_result(e)

    def activate_application(self) -> dict[str, Any]:
        """
        Activate (bring to foreground) the Solid Edge application window.

        Uses Application.Activate().

        Returns:
            Dict with status
        """
        try:
            app = self._get_app()
            app.Activate()
            return {"status": "activated"}
        except Exception as e:
            return error_result(e)

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
            app = self._get_app()
            app.AbortCommand(abort_all)
            return {"status": "aborted", "abort_all": abort_all}
        except Exception as e:
            return error_result(e)

    def get_active_environment(self) -> dict[str, Any]:
        """
        Get the currently active environment in Solid Edge.

        The active environment determines which commands and menus
        are available (e.g., Part, Assembly, Draft).

        Returns:
            Dict with environment info
        """
        try:
            app = self._get_app()
            env = app.ActiveEnvironment

            result = {"status": "success"}
            try:
                result["name"] = env.Name
            except Exception:
                result["name"] = str(env)
            with contextlib.suppress(Exception):
                result["caption"] = env.Caption

            return result
        except Exception as e:
            return error_result(e)

    def get_status_bar(self) -> dict[str, Any]:
        """
        Get the current status bar text.

        Returns:
            Dict with status bar text
        """
        try:
            app = self._get_app()
            text = app.StatusBar
            return {"status": "success", "text": text}
        except Exception as e:
            return error_result(e)

    def set_status_bar(self, text: str) -> dict[str, Any]:
        """
        Set the status bar text.

        Args:
            text: Text to display in the status bar

        Returns:
            Dict with status
        """
        try:
            app = self._get_app()
            app.StatusBar = text
            return {"status": "set", "text": text}
        except Exception as e:
            return error_result(e)

    def get_visible(self) -> dict[str, Any]:
        """
        Get the visibility state of the Solid Edge application window.

        Returns:
            Dict with visible state
        """
        try:
            app = self._get_app()
            visible = app.Visible
            return {"status": "success", "visible": visible}
        except Exception as e:
            return error_result(e)

    def set_visible(self, visible: bool) -> dict[str, Any]:
        """
        Set the visibility of the Solid Edge application window.

        Args:
            visible: True to show, False to hide

        Returns:
            Dict with status
        """
        try:
            app = self._get_app()
            app.Visible = visible
            return {"status": "set", "visible": visible}
        except Exception as e:
            return error_result(e)

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
            app = self._get_app()
            value = app.GetGlobalParameter(parameter)
            return {"status": "success", "parameter": parameter, "value": value}
        except Exception as e:
            return error_result(e)

    def set_global_parameter(self, parameter: int, value: Any) -> dict[str, Any]:
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
            app = self._get_app()
            app.SetGlobalParameter(parameter, value)
            return {"status": "set", "parameter": parameter, "value": value}
        except Exception as e:
            return error_result(e)

    def convert_by_file_path(self, input_path: str, output_path: str) -> dict[str, Any]:
        """
        Batch-convert CAD files between formats -- refused, with the evidence.

        Application.ConvertByFilePath returns in about a second reporting
        nothing wrong and writes nothing. Verified on Solid Edge 2026 with a
        saved part and nine output extensions (stp, step, x_t, igs, jt, pdf,
        dxf, stl, par): no file appeared for any of them, while SaveCopyAs on
        the opened document wrote the STEP at once. export_file is that route.

        Args:
            input_path: Input file or folder path
            output_path: Output file or folder path

        Returns:
            Dict with the refusal
        """
        return {
            "error": (
                "Application.ConvertByFilePath writes nothing on Solid Edge 2026 "
                "(nine output formats tried; SaveCopyAs writes them). Open the file "
                "and use export_file instead."
            ),
            "unsupported": True,
            "input": input_path,
            "output": output_path,
        }

    def get_default_template_path(self, doc_type: int) -> dict[str, Any]:
        """
        Get the default template file path for a document type.

        Args:
            doc_type: Document type constant (1=Part, 2=Draft, 3=Assembly, 4=SheetMetal)

        Returns:
            Dict with template path
        """
        try:
            app = self._get_app()
            path = app.GetDefaultTemplatePath(doc_type)
            return {"status": "success", "doc_type": doc_type, "template_path": path}
        except Exception as e:
            return error_result(e)

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
            app = self._get_app()
            app.SetDefaultTemplatePath(doc_type, template_path)
            return {
                "status": "set",
                "doc_type": doc_type,
                "template_path": template_path,
            }
        except Exception as e:
            return error_result(e)

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
            app = self._get_app()
            style_names = {1: "Tiled", 2: "Horizontal", 4: "Vertical", 8: "Cascade"}
            app.ArrangeWindows(style)
            return {
                "status": "arranged",
                "style": style,
                "style_name": style_names.get(style, f"Unknown({style})"),
            }
        except Exception as e:
            return error_result(e)

    def get_active_command(self) -> dict[str, Any]:
        """
        Get the currently active command in Solid Edge.

        Returns:
            Dict with active command info
        """
        try:
            app = self._get_app()
            # framewrk.tlb exposes GetActiveCommand() as a method; there is
            # no ActiveCommand property, so this always raised.
            cmd = app.GetActiveCommand()
            result: dict[str, Any | None] = {"status": "success"}
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
            return error_result(e)

    def run_macro(self, filename: str) -> dict[str, Any]:
        """
        Run a VBA macro file in Solid Edge.

        Args:
            filename: Path to the VBA macro file (.bas, .exe, etc.)

        Returns:
            Dict with status
        """
        try:
            app = self._get_app()
            app.RunMacro(filename)
            return {"status": "executed", "filename": filename}
        except Exception as e:
            return error_result(e)
