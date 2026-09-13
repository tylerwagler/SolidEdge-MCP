"""Document management tools for Solid Edge MCP."""

from typing import Any, Literal

from solidedge_mcp.backends.validation import validate_path
from solidedge_mcp.managers import doc_manager
from solidedge_mcp.tools._registry import register_tool

# === Composite: create_document ===


def create_document(
    type: Literal["part", "assembly", "sheet_metal", "draft", "weldment"] = "part",
    template: str | None = None,
) -> dict[str, Any]:
    """Create a new document of the given type; it becomes the active document.

    template: optional template file path (None uses the Solid Edge default).
    """
    match type:
        case "part":
            return doc_manager.create_part(template=template)
        case "assembly":
            return doc_manager.create_assembly(template=template)
        case "sheet_metal":
            return doc_manager.create_sheet_metal(template=template)
        case "draft":
            return doc_manager.create_draft(template=template)
        case "weldment":
            return doc_manager.create_weldment(template=template)
        case _:
            return {"error": f"Unknown document type: {type}"}


# === Composite: open_document ===


def open_document(
    method: Literal["foreground", "background", "with_template", "dialog"] = "foreground",
    file_path: str = "",
    template: str = "",
    filename: str | None = None,
    dialog_title: str | None = None,
) -> dict[str, Any]:
    """Open an existing document.

    foreground/background: file_path (must exist). with_template: file_path +
    template. dialog: shows the Solid Edge open dialog (optional filename
    preset, dialog_title).
    """
    if method in ("foreground", "background", "with_template") and file_path:
        file_path, err = validate_path(file_path, must_exist=True)
        if err:
            return err
    match method:
        case "foreground":
            return doc_manager.open_document(file_path=file_path)
        case "background":
            return doc_manager.open_in_background(file_path=file_path)
        case "with_template":
            return doc_manager.open_with_template(file_path=file_path, template=template)
        case "dialog":
            return doc_manager.open_with_file_open_dialog(
                filename=filename, dialog_title=dialog_title
            )
        case _:
            return {"error": f"Unknown method: {method}"}


# === Composite: close_document ===


def close_document(
    scope: Literal["active", "all"] = "active",
    save: bool = True,
) -> dict[str, Any]:
    """Close the active document or all documents.

    save=False discards unsaved changes.
    """
    match scope:
        case "active":
            return doc_manager.close_document(save=save)
        case "all":
            return doc_manager.close_all_documents(save=save)
        case _:
            return {"error": f"Unknown scope: {scope}"}


# === Composite: save_document ===


def save_document(
    method: Literal["save", "copy_as"] = "save",
    file_path: str | None = None,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Save the active document.

    save: in place, or Save As when file_path is given.
    copy_as: write a copy to file_path (required) and keep the current file active.
    overwrite: required to replace an existing file. Without it the call is
    refused, because Solid Edge would raise a modal overwrite prompt that
    blocks the server until someone clicks it.
    """
    if file_path:
        file_path, err = validate_path(file_path, must_exist=False)
        if err:
            return err
    match method:
        case "save":
            return doc_manager.save_document(file_path=file_path, overwrite=overwrite)
        case "copy_as":
            if file_path is None:
                return {"error": "file_path is required for 'copy_as' method"}
            return doc_manager.save_copy_as(file_path=file_path)
        case _:
            return {"error": f"Unknown method: {method}"}


# === Composite: undo_redo ===


def undo_redo(action: Literal["undo", "redo"] = "undo") -> dict[str, Any]:
    """Undo or redo the last operation in the active document."""
    match action:
        case "undo":
            return doc_manager.undo()
        case "redo":
            return doc_manager.redo()
        case _:
            return {"error": f"Unknown action: {action}"}


# === Standalone tools ===


def activate_document(name_or_index: str | int) -> dict[str, Any]:
    """Make an open document active, by document name or 0-based index."""
    return doc_manager.activate_document(name_or_index=name_or_index)


def import_file(file_path: str) -> dict[str, Any]:
    """Import an external CAD file (STEP, IGES, Parasolid, ...) as a new document."""
    file_path, err = validate_path(file_path, must_exist=True)
    if err:
        return err
    return doc_manager.import_file(file_path=file_path)


# === Registration ===


def register(mcp: Any) -> None:
    """Register document tools with the MCP server."""
    tags = {"document"}
    # Composite tools
    register_tool(mcp, create_document, tags=tags)
    register_tool(mcp, open_document, tags=tags)
    register_tool(mcp, close_document, tags=tags, destructive=True)
    register_tool(mcp, save_document, tags=tags, idempotent=True)
    register_tool(mcp, undo_redo, tags=tags)
    # Standalone tools
    register_tool(mcp, activate_document, tags=tags, idempotent=True)
    register_tool(mcp, import_file, tags=tags)
