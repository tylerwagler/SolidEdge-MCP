"""Document management tools for Solid Edge MCP."""

from typing import Optional
from solidedge_mcp.managers import doc_manager

def create_part_document(template: Optional[str] = None) -> dict:
    """Create a new part document."""
    return doc_manager.create_part(template)

def create_assembly_document(template: Optional[str] = None) -> dict:
    """Create a new assembly document."""
    return doc_manager.create_assembly(template)

def create_sheet_metal_document(template: Optional[str] = None) -> dict:
    """Create a new sheet metal document."""
    return doc_manager.create_sheet_metal(template)

def open_document(file_path: str) -> dict:
    """Open an existing document (.par, .asm, .dft)."""
    return doc_manager.open_document(file_path)

def save_document(file_path: Optional[str] = None) -> dict:
    """Save the active document. If file_path is not provided, saves to current location."""
    return doc_manager.save_document(file_path)

def close_document(save: bool = True) -> dict:
    """Close the active document. Saves before closing if `save` is True."""
    return doc_manager.close_document(save)

def activate_document(name_or_index) -> dict:
    """Activate a specific open document by name or index."""
    return doc_manager.activate_document(name_or_index)

def undo() -> dict:
    """Undo the last operation on the active document."""
    return doc_manager.undo()

def redo() -> dict:
    """Redo the last undone operation on the active document."""
    return doc_manager.redo()

def create_draft_document(template: Optional[str] = None) -> dict:
    """Create a new draft (drawing) document."""
    return doc_manager.create_draft(template)

def list_documents() -> dict:
    """List all open documents."""
    return doc_manager.list_documents()

def get_active_document_type() -> dict:
    """Get the type of the currently active document."""
    return doc_manager.get_active_document_type()

def create_weldment_document(template: Optional[str] = None) -> dict:
    """Create a new weldment document."""
    return doc_manager.create_weldment(template)

def import_file(file_path: str) -> dict:
    """Import an external CAD file (STEP, IGES, Parasolid, etc.)."""
    return doc_manager.import_file(file_path)

def get_document_count() -> dict:
    """Get the count of open documents."""
    return doc_manager.get_document_count()

def open_in_background(file_path: str) -> dict:
    """Open a document in the background without showing a window."""
    return doc_manager.open_in_background(file_path)

def close_all_documents(save: bool = False) -> dict:
    """Close all open documents. Optionally save them."""
    return doc_manager.close_all_documents(save)

def register(mcp):
    """Register document tools with the MCP server."""
    mcp.tool()(create_part_document)
    mcp.tool()(create_assembly_document)
    mcp.tool()(create_sheet_metal_document)
    mcp.tool()(open_document)
    mcp.tool()(save_document)
    mcp.tool()(close_document)
    mcp.tool()(activate_document)
    mcp.tool()(undo)
    mcp.tool()(redo)
    mcp.tool()(create_draft_document)
    mcp.tool()(list_documents)
    mcp.tool()(get_active_document_type)
    mcp.tool()(create_weldment_document)
    mcp.tool()(import_file)
    mcp.tool()(get_document_count)
    mcp.tool()(open_in_background)
    mcp.tool()(close_all_documents)
