"""
Solid Edge Query and Inspection Operations

Handles querying model data, measurements, and properties.
"""

from ._base import QueryManagerBase
from ._brep import BRepMixin
from ._document import DocumentQueryMixin
from ._features import FeatureQueryMixin
from ._materials import MaterialsMixin
from ._physical_props import PhysicalPropsMixin
from ._selection import SelectionMixin
from ._variables import VariablesMixin


class QueryManager(
    PhysicalPropsMixin,
    DocumentQueryMixin,
    VariablesMixin,
    BRepMixin,
    SelectionMixin,
    FeatureQueryMixin,
    MaterialsMixin,
    QueryManagerBase,
):
    """Manages query and inspection operations"""

    pass
