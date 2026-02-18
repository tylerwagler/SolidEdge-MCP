"""
Solid Edge Assembly Operations

Handles assembly creation and component management.
"""

from ._base import AssemblyManagerBase
from ._features import AssemblyFeaturesMixin
from ._placement import PlacementMixin
from ._properties import PropertiesMixin
from ._query import QueryMixin
from ._relations import RelationsMixin
from ._specialized import SpecializedMixin
from ._transforms import TransformsMixin


class AssemblyManager(
    PlacementMixin,
    QueryMixin,
    TransformsMixin,
    PropertiesMixin,
    RelationsMixin,
    AssemblyFeaturesMixin,
    SpecializedMixin,
    AssemblyManagerBase,
):
    """Manages assembly operations"""

    pass
