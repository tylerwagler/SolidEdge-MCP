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
