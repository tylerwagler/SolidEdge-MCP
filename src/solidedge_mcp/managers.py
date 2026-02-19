"""Global manager instances for Solid Edge MCP."""

from solidedge_mcp.backends.assembly import AssemblyManager
from solidedge_mcp.backends.connection import SolidEdgeConnection
from solidedge_mcp.backends.diagnostics import diagnose_document, diagnose_feature
from solidedge_mcp.backends.documents import DocumentManager
from solidedge_mcp.backends.export import ExportManager, ViewModel
from solidedge_mcp.backends.features import FeatureManager
from solidedge_mcp.backends.query import QueryManager
from solidedge_mcp.backends.sketching import SketchManager

# Initialize managers (global state)
connection = SolidEdgeConnection()
doc_manager = DocumentManager(connection)
sketch_manager = SketchManager(doc_manager)
doc_manager.sketch_manager = sketch_manager  # Provide reference for sketch state clearing
feature_manager = FeatureManager(doc_manager, sketch_manager)
assembly_manager = AssemblyManager(doc_manager, sketch_manager)
query_manager = QueryManager(doc_manager)
export_manager = ExportManager(doc_manager)
view_manager = ViewModel(doc_manager)

# Re-export diagnostics functions if needed by tools directly
__all__ = [
    "connection",
    "doc_manager",
    "sketch_manager",
    "feature_manager",
    "assembly_manager",
    "query_manager",
    "export_manager",
    "view_manager",
    "diagnose_document",
    "diagnose_feature",
]
