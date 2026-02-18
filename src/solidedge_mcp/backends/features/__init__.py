"""
Solid Edge Feature Operations

Handles creating 3D features like extrusions, revolves, holes, fillets, etc.
"""

from ._base import FeatureManagerBase
from ._cutout import CutoutMixin
from ._extrude import ExtrudeMixin
from ._holes import HolesMixin
from ._loft_sweep import LoftSweepMixin
from ._misc import MiscFeaturesMixin
from ._primitives import PrimitiveMixin
from ._ref_planes import RefPlaneMixin
from ._revolve import RevolveMixin
from ._rounds_chamfers import RoundsChamfersMixin
from ._sheet_metal import SheetMetalMixin
from ._surfaces import SurfacesMixin


class FeatureManager(
    ExtrudeMixin,
    RevolveMixin,
    CutoutMixin,
    PrimitiveMixin,
    RefPlaneMixin,
    LoftSweepMixin,
    RoundsChamfersMixin,
    HolesMixin,
    SheetMetalMixin,
    SurfacesMixin,
    MiscFeaturesMixin,
    FeatureManagerBase,
):
    """Manages 3D feature creation"""

    pass
