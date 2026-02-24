"""Document management tools for Solid Edge MCP."""

from typing import Any

from solidedge_mcp.backends.validation import validate_path
from solidedge_mcp.managers import doc_manager

# === Composite: create_document ===


def create_document(
    type: str = "part",
    template: str | None = None,
) -> dict[str, Any]:
    """Create a new Solid Edge document.

    type: 'part' | 'assembly' | 'sheet_metal' | 'draft' | 'weldment'
    """
    match type:
        case "part":
            return doc_manager.create_part(template)
        case "assembly":
            return doc_manager.create_assembly(template)
        case "sheet_metal":
            return doc_manager.create_sheet_metal(template)
        case "draft":
            return doc_manager.create_draft(template)
        case "weldment":
            return doc_manager.create_weldment(template)
        case _:
            return {"error": f"Unknown document type: {type}"}


# === Composite: open_document ===


def open_document(
    method: str = "foreground",
    file_path: str = "",
    template: str = "",
    filename: str | None = None,
    dialog_title: str | None = None,
) -> dict[str, Any]:
    """Open a document.

    method: 'foreground' | 'background' | 'with_template' | 'dialog'
    """
    if method in ("foreground", "background", "with_template") and file_path:
        file_path, err = validate_path(file_path, must_exist=True)
        if err:
            return err
    match method:
        case "foreground":
            return doc_manager.open_document(file_path)
        case "background":
            return doc_manager.open_in_background(file_path)
        case "with_template":
            return doc_manager.open_with_template(
                file_path, template
            )
        case "dialog":
            return doc_manager.open_with_file_open_dialog(
                filename, dialog_title
            )
        case _:
            return {"error": f"Unknown method: {method}"}


# === Composite: close_document ===


def close_document(
    scope: str = "active",
    save: bool = True,
) -> dict[str, Any]:
    """Close documents.

    scope: 'active' | 'all'
    """
    match scope:
        case "active":
            return doc_manager.close_document(save)
        case "all":
            return doc_manager.close_all_documents(save)
        case _:
            return {"error": f"Unknown scope: {scope}"}


# === Composite: save_document ===


def save_document(
    method: str = "save",
    file_path: str | None = None,
) -> dict[str, Any]:
    """Save the active document.

    method: 'save' | 'copy_as'
    """
    if file_path:
        file_path, err = validate_path(file_path, must_exist=False)
        if err:
            return err
    match method:
        case "save":
            return doc_manager.save_document(file_path)
        case "copy_as":
            if file_path is None:
                return {"error": "file_path is required for 'copy_as' method"}
            return doc_manager.save_copy_as(file_path)
        case _:
            return {"error": f"Unknown method: {method}"}


# === Composite: undo_redo ===


def undo_redo(action: str = "undo") -> dict[str, Any]:
    """Undo or redo the last operation.

    action: 'undo' | 'redo'
    """
    match action:
        case "undo":
            return doc_manager.undo()
        case "redo":
            return doc_manager.redo()
        case _:
            return {"error": f"Unknown action: {action}"}


# === Standalone tools ===


def activate_document(name_or_index: str | int) -> dict[str, Any]:
    """Activate a specific open document by name or index."""
    return doc_manager.activate_document(name_or_index)


def import_file(file_path: str) -> dict[str, Any]:
    """Import an external CAD file (STEP, IGES, Parasolid, etc.)."""
    file_path, err = validate_path(file_path, must_exist=True)
    if err:
        return err
    return doc_manager.import_file(file_path)


# === Registration ===


def register(mcp: Any) -> None:
    """Register document tools with the MCP server."""
    # Composite tools
    mcp.tool()(create_document)
    mcp.tool()(open_document)
    mcp.tool()(close_document)
    mcp.tool()(save_document)
    mcp.tool()(undo_redo)
    # Standalone tools
    mcp.tool()(activate_document)
    mcp.tool()(import_file)
