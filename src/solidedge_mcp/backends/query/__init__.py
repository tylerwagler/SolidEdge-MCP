"""
Solid Edge Query and Inspection Operations

Handles querying model data, measurements, and properties.
"""

from ._base import DEFAULT_PAGE_LIMIT, MAX_PAGE_LIMIT, QueryManagerBase
from ._brep import BRepMixin
from ._document import DocumentQueryMixin
from ._features import FeatureQueryMixin
from ._materials import MaterialsMixin
from ._physical_props import PhysicalPropsMixin
from ._selection import SelectionMixin
from ._spatial import SpatialContextMixin
from ._variables import VariablesMixin


class QueryManager(
    PhysicalPropsMixin,
    DocumentQueryMixin,
    VariablesMixin,
    BRepMixin,
    SelectionMixin,
    FeatureQueryMixin,
    MaterialsMixin,
    SpatialContextMixin,
    QueryManagerBase,
):
    """Manages query and inspection operations"""

    pass


__all__ = ["DEFAULT_PAGE_LIMIT", "MAX_PAGE_LIMIT", "QueryManager"]
