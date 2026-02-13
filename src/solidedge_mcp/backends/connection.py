"""
Solid Edge Connection Management

Handles connecting to and managing Solid Edge application instances.
"""

import win32com.client
import pythoncom
from typing import Optional, Dict, Any
import traceback


class SolidEdgeConnection:
    """Manages connection to Solid Edge application"""

    def __init__(self):
        self.application: Optional[Any] = None
        self._is_connected: bool = False

    def connect(self, start_if_needed: bool = True) -> Dict[str, Any]:
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
                    print("Connected to existing Solid Edge instance")
                except:
                    if start_if_needed:
                        # Start new instance with early binding if possible
                        try:
                            self.application = win32com.client.gencache.EnsureDispatch(
                                "SolidEdge.Application"
                            )
                        except:
                            # Fall back to late binding
                            self.application = win32com.client.Dispatch("SolidEdge.Application")

                        self.application.Visible = True
                        print("Started new Solid Edge instance")
                    else:
                        raise Exception("No Solid Edge instance found and start_if_needed=False")

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
            return {
                "status": "error",
                "message": str(e),
                "traceback": traceback.format_exc()
            }

    def disconnect(self) -> Dict[str, Any]:
        """Disconnect from Solid Edge (does not close the application)"""
        self.application = None
        self._is_connected = False
        return {"status": "disconnected"}

    def get_info(self) -> Dict[str, Any]:
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
            except:
                info["path"] = "N/A"

            return info
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_application_info(self) -> Dict[str, Any]:
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

    def quit_application(self) -> Dict[str, Any]:
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

            return {"status": "quit", "message": "Solid Edge has been closed"}
        except Exception as e:
            self.application = None
            self._is_connected = False
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_process_info(self) -> Dict[str, Any]:
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

    def get_install_info(self) -> Dict[str, Any]:
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
                try:
                    info["install_path"] = install_data.GetInstalledPath()
                except Exception:
                    pass
                try:
                    info["language"] = install_data.GetInstalledLanguage()
                except Exception:
                    pass
                try:
                    info["version"] = install_data.GetInstalledVersion()
                except Exception:
                    pass
            except Exception:
                # SEInstallData may not be registered; fall back to Application.Path
                if self._is_connected and self.application is not None:
                    try:
                        info["install_path"] = self.application.Path
                    except Exception:
                        pass

            if not info:
                return {"error": "Could not retrieve installation info. SEInstallData COM library may not be registered."}

            return {"status": "success", **info}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_performance_mode(
        self,
        delay_compute: bool = None,
        screen_updating: bool = None,
        interactive: bool = None,
        display_alerts: bool = None
    ) -> Dict[str, Any]:
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

            return {
                "status": "updated",
                "settings": settings
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def do_idle(self) -> Dict[str, Any]:
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
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }
