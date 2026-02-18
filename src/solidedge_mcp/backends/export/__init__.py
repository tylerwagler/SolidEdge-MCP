"""
Solid Edge Export and Visualization Operations

Handles exporting to various formats, creating drawings, and view manipulation.
"""

from ._annotations import AnnotationsMixin
from ._base import ExportManagerBase
from ._draft import DraftMixin
from ._drawing import DrawingMixin
from ._file_export import FileExportMixin
from ._view_model import ViewModel
from ._views import ViewsMixin


class ExportManager(
    FileExportMixin,
    DrawingMixin,
    ViewsMixin,
    AnnotationsMixin,
    DraftMixin,
    ExportManagerBase,
):
    """Manages export and drawing operations"""

    pass


__all__ = ["ExportManager", "ViewModel"]
