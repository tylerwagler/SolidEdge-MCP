"""
Solid Edge COM API Constants

These constants are defined by the Solid Edge API.
References: Solid Edge API documentation
"""


class RefPlaneConstants:
    """Reference plane type constants"""
    seRefPlaneTop = 1
    seRefPlaneFront = 2
    seRefPlaneRight = 3


class DocumentTypeConstants:
    """Document type constants"""
    igUnknownDocument = 0
    igPartDocument = 1
    igSheetMetalDocument = 2
    igAssemblyDocument = 3
    igDraftDocument = 4
    igWeldmentDocument = 5
    igWeldmentAssemblyDocument = 6


class FeaturePropertyConstants:
    """Feature property constants"""
    igStatusNormal = 0
    igStatusSuppressed = 1
    igStatusRollback = 2


class FeatureOperationConstants:
    """Feature operation type constants"""
    igFeatureAdd = 0
    igFeatureCut = 1
    igFeatureIntersect = 2
    igFeatureJoin = 3


class ExtrudedProtrusion:
    """Extrusion/Revolve direction constants (from Solid Edge type library)"""
    igLeft = 1       # Left/Reverse direction
    igRight = 2      # Right/Normal direction
    igSymmetric = 3  # Symmetric (both directions)


class ProfileValidationConstants:
    """Profile.End() validation flag constants (bitfield)"""
    igProfileClosed = 1              # Profile must be closed
    igProfileSingle = 4             # Single profile only
    igProfileNoSelfIntersect = 8    # No self-intersection
    igProfileRefAxisRequired = 16   # Reference axis required (for revolve)
    igProfileNoRefAxisIntersect = 32  # Profile must not intersect axis
    igProfileAllowNested = 8192     # Allow nested profiles

    # Common combinations
    igProfileDefault = 0                                  # Default (extrude)
    igProfileForRevolve = 1 | 16   # igProfileClosed | igProfileRefAxisRequired = 17


class ExtentTypeConstants:
    """Extent type constants"""
    igFinite = 13
    igNone = 44


class HoleTypeConstants:
    """Hole type constants"""
    igRegularHole = 0
    igCounterboreHole = 1
    igCountersinkHole = 2
    igVHole = 3


class MateTypeConstants:
    """Assembly mate type constants"""
    igMate = 0
    igPlanarAlign = 1
    igAxialAlign = 2
    igInsert = 3
    igAngle = 4
    igTangent = 5
    igCam = 6
    igGear = 7
    igParallel = 8
    igConnect = 9
    igMatchCoordSys = 10


class ViewOrientationConstants:
    """View orientation constants"""
    seIsoView = 1
    seTopView = 2
    seFrontView = 3
    seRightView = 4
    seLeftView = 5
    seBackView = 6
    seBottomView = 7


class SaveAsConstants:
    """File save format constants"""
    igNormalSave = 0
    igSaveAsCopy = 1
