"""
Solid Edge COM API Constants

All values verified against the Solid Edge type library
(gencache.EnsureModule) unless otherwise noted.
"""


class RefPlaneConstants:
    """Reference plane index constants (1-based collection indices)"""

    seRefPlaneTop = 1  # Top/XZ plane
    seRefPlaneFront = 2  # Front/XY plane
    seRefPlaneRight = 3  # Right/YZ plane


class DocumentTypeConstants:
    """Document type constants (from type library)"""

    igPartDocument = 1
    igDraftDocument = 2
    igAssemblyDocument = 3
    igSheetMetalDocument = 4
    igUnknownDocument = 5
    igWeldmentDocument = 6
    igWeldmentAssemblyDocument = 7


class DirectionConstants:
    """Extrusion/Revolve direction constants (from type library)"""

    igLeft = 1  # Left/Reverse direction
    igRight = 2  # Right/Normal direction (also igNormalSide)
    igSymmetric = 3  # Symmetric (both directions)
    igBoth = 6  # Both directions


class ProfileValidationConstants:
    """Profile.End() validation flag constants (bitfield, from type library)"""

    igProfileClosed = 1  # Profile must be closed
    igProfileSingle = 4  # Single profile only
    igProfileNoSelfIntersect = 8  # No self-intersection
    igProfileRefAxisRequired = 16  # Reference axis required (for revolve)
    igProfileNoRefAxisIntersect = 32  # Profile must not intersect axis
    igProfileAllowNested = 8192  # Allow nested profiles

    # Common combinations
    igProfileDefault = 0  # Default (extrude)
    igProfileForRevolve = 17  # igProfileClosed | igProfileRefAxisRequired


class ExtentTypeConstants:
    """Extent type constants (from type library)"""

    igFinite = 13
    igThroughAll = 16
    igNone = 44


class HoleTypeConstants:
    """Hole type constants (from type library)"""

    igRegularHole = 33
    igCounterboreHole = 34
    igCountersinkHole = 35
    igCounterdrillHole = 36
    igTappedHole = 37
    igTaperedHole = 38


class FaceQueryConstants:
    """Body.Faces() query type constants (from type library)"""

    igQueryAll = 1
    igQueryRoundable = 2
    igQueryStraight = 3
    igQueryEllipse = 4
    igQuerySpline = 5
    igQueryPlane = 6
    igQueryCone = 7
    igQueryTorus = 8
    igQuerySphere = 9
    igQueryCylinder = 10


class ViewOrientationConstants:
    """Standard orthographic view orientation constants (from type library)"""

    igTopView = 1
    igRightView = 2
    igLeftView = 3
    igFrontView = 4
    igBottomView = 5
    igBackView = 6
    # Pictorial views use a separate numbering scheme:
    igTopFrontLeftView = 8  # Standard isometric
    igTopFrontRightView = 9
    igTopBackLeftView = 7
    igTopBackRightView = 10


class DrawingViewTypeConstants:
    """Drawing view type constants (from type library - separate enum)"""

    igPrincipleView = 1
    igIsometricView = 2
    igAuxiliaryView = 3
    igXSectionView = 4
    igDetailView = 5
    igIsoXSectionView = 6


class DrawingViewOrientationConstants:
    """Drawing view orientation constants for AddPartView (empirically verified)"""

    Front = 5
    Top = 6
    Right = 7
    Back = 8
    Bottom = 9
    Left = 10
    Isometric = 12


class RenderModeConstants:
    """View render mode constants (from type library: seRenderMode*)"""

    seRenderModeUndefined = 0
    seRenderModeWireframe = 1
    seRenderModeWiremesh = 2
    seRenderModeOutline = 3
    seRenderModeBoundary = 4
    seRenderModeVHL = 6  # Hidden edges visible
    seRenderModeSmooth = 8  # Shaded
    seRenderModeSmoothMesh = 9
    seRenderModeSmoothVHL = 10
    seRenderModeSmoothBoundary = 11  # Shaded with edges


class AssemblyRelationConstants:
    """Assembly 3D relation type constants (from type library)

    Note: These are large integers used as COM type identifiers,
    not small sequential enum values.
    """

    igGroundRelation3d = 1959028688
    igPlanarRelation3d = -2058948880
    igAxialRelation3d = 1472929712
    igAngularRelation3d = 1290792304
    igTangentRelation3d = 918452310

    # Relation orientation
    igRelation3dOrientationAlign = 1
    igRelation3dOrientationAntialign = 2
    igRelation3dOrientationNotspecified = 0


class ModelingModeConstants:
    """Modeling mode constants (from type library)"""

    seModelingModeSynchronous = 1
    seModelingModeOrdered = 2


class TreatmentTypeConstants:
    """Treatment type constants for extruded surfaces (from type library)"""

    seTreatmentNone = 0
    seTreatmentCrown = 1
    seTreatmentDraft = 2
    seTreatmentCrownAndDraft = 3


class DraftSideConstants:
    """Draft side constants (from type library)"""

    seDraftNone = 0
    seDraftInside = 1
    seDraftOutside = 2


class TreatmentCrownTypeConstants:
    """Crown type constants (from type library)"""

    seTreatmentCrownByRadius = 0
    seTreatmentCrownByOffset = 1


class TreatmentCrownSideConstants:
    """Crown side constants (from type library)"""

    seTreatmentCrownSideInside = 0
    seTreatmentCrownSideOutside = 1


class TreatmentCrownCurvatureSideConstants:
    """Crown curvature side constants (from type library)"""

    seTreatmentCrownCurvatureInside = 0
    seTreatmentCrownCurvatureOutside = 1


class OffsetSideConstants:
    """Offset side constants (from type library)"""

    seOffsetNone = 0
    seOffsetInside = 1
    seOffsetOutside = 2


class KeyPointExtentConstants:
    """Keypoint extent constants (from type library)"""

    igTangentNormal = 0


class KeyPointTypeConstants:
    """Keypoint type constants (from constant.tlb > KeyPointType enum)"""

    igKeyPointStart = 1
    igKeyPointEnd = 2


class ReferenceElementConstants:
    """Reference element constants (from constant.tlb > ReferenceElementConstants enum)"""

    igRefEleInit = 0
    igReverseNormalSide = 1
    igNormalSide = 2
    igPivotStart = 3
    igPivotEnd = 4
    igCurveStart = 14
    igCurveEnd = 15
    igNormalToCurveAtKeyPoint = 22
    igTangentToSurfaceAtAngle = 25
    igTangentToSurfaceAtKeypoint = 26


class LoftSweepConstants:
    """Loft and sweep profile type constants (from type library)"""

    igProfileBasedCrossSection = 48


# === Legacy aliases for backward compatibility with existing imports ===
# These preserve the old class names used in backend code.

# ExtrudedProtrusion was renamed to DirectionConstants
ExtrudedProtrusion = DirectionConstants


# FeatureOperationConstants - names not in type library, values unverified.
# Kept as alias since features.py imports it (though the values are never
# passed to any API call).
class FeatureOperationConstants:
    """Feature operation type constants (NOT in type library - unverified)"""

    igFeatureAdd = 0
    igFeatureCut = 1
    igFeatureIntersect = 2
    igFeatureJoin = 3


# MateTypeConstants - names not in type library.
# Assembly relations use AssemblyRelationConstants instead.
# Kept as alias since it may be imported elsewhere.
class MateTypeConstants:
    """Assembly mate type constants (NOT in type library - unverified)

    For verified constants, use AssemblyRelationConstants instead.
    """

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


# SaveAsConstants - names not in type library, but values match
# (SaveAs=0, SaveCopyAs=1)
class AssemblyGlobalConstants:
    """Assembly global parameter constants (from type library: seAssemblyGlobal*)

    Used with Application.GetGlobalParameter / SetGlobalParameter.
    """

    seAssemblyGlobalTubeWallThickness = 1
    seAssemblyGlobalTubeOuterDiameter = 2
    seAssemblyGlobalMiterClearance = 3
    seAssemblyGlobalTrimExtendLength = 4
    seAssemblyGlobalCopeClearance = 5
    seAssemblyGlobalNotchPlateLength = 6
    seAssemblyGlobalWeldOffset = 7
    seAssemblyGlobalWeldAngle = 8
    seAssemblyGlobalWeldSize = 9
    seAssemblyGlobalBoltDiameter = 10
    seAssemblyGlobalHoleDiameter = 11
    seAssemblyGlobalHoleDiameterOffset = 12
    seAssemblyGlobalBoltHeadDiameter = 13
    seAssemblyGlobalBoltLength = 14
    seAssemblyGlobalNutDiameter = 15
    seAssemblyGlobalNutHeight = 16
    seAssemblyGlobalWasherOuterDiameter = 17
    seAssemblyGlobalWasherThickness = 18
    seAssemblyGlobalDefaultMaterial = 19
    seAssemblyGlobalEndCapType = 20
    seAssemblyGlobalAdjustableTubes = 21


class FoldTypeConstants:
    """Drawing view fold direction constants (from type library)"""

    igFoldUp = 1
    igFoldDown = 2
    igFoldRight = 3
    igFoldLeft = 4


class SaveAsConstants:
    """File save format constants (values verified, names approximate)"""

    SaveAs = 0
    SaveCopyAs = 1
