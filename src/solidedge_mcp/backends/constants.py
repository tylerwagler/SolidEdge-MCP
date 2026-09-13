"""
Solid Edge COM API Constants

All values verified against the Solid Edge type library
(gencache.EnsureModule) unless otherwise noted.
"""


class RefPlaneConstants:
    """Reference plane index constants (1-based collection indices)"""

    seRefPlaneTop = 1  # Top/XY plane
    seRefPlaneRight = 2  # Right/YZ plane
    seRefPlaneFront = 3  # Front/XZ plane


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
    """constant.tlb > TreatmentTypeConstants."""

    seTreatmentNone = 44
    seTreatmentDraft = 173
    seTreatmentCrown = 174
    # Not a member of the type library enum; Solid Edge has no combined value.
    # Kept so existing callers resolve, but it will be rejected by COM.
    seTreatmentCrownAndDraft = 3


class DraftSideConstants:
    """constant.tlb > DraftSideConstants."""

    seDraftNone = 44
    seDraftInside = 4
    seDraftOutside = 5


class TreatmentCrownTypeConstants:
    """constant.tlb > TreatmentCrownTypeConstants."""

    seTreatmentCrownNone = 0
    seTreatmentCrownByRadius = 1
    seTreatmentCrownByRadiusAndTakeOffAngle = 2
    seTreatmentCrownByOffset = 3
    seTreatmentCrownByOffsetAndTakeOffAngle = 4


class TreatmentCrownSideConstants:
    """constant.tlb > TreatmentCrownSideConstants."""

    seTreatmentCrownSideNone = 44
    seTreatmentCrownSideInside = 4
    seTreatmentCrownSideOutside = 5


class TreatmentCrownCurvatureSideConstants:
    """constant.tlb > TreatmentCrownCurvatureSideConstants."""

    seTreatmentCrownCurvatureNone = 44
    seTreatmentCrownCurvatureInside = 4
    seTreatmentCrownCurvatureOutside = 5


class OffsetSideConstants:
    """constant.tlb > OffsetSideConstants.

    The enum is None/Left/Right; there are no Inside/Outside members.
    """

    seOffsetNone = 44
    seOffsetLeft = 1
    seOffsetRight = 2


class AddBodyTypeConstants:
    """constant.tlb > AddBodyTypeConstants."""

    igPartType = 1
    igSheetMetalType = 2
    igSubdivisionType = 3
    igSubdivisionControlCageType = 4
    igConstructionPartType = 5
    igConstructionSheetMetalType = 6
    igConstructionSubdivisionType = 7


class DrawnCutoutFeatureConstants:
    """constant.tlb > DrawnCutoutFeatureConstants."""

    seDrawnCutoutDepthLeft = 1
    seDrawnCutoutDepthRight = 2
    seDrawnCutoutMaterialInside = 3
    seDrawnCutoutMaterialOutside = 4
    seDrawnCutoutProfileLeft = 5
    seDrawnCutoutProfileRight = 6
    seDrawnCutoutRoundEdges = 7
    seDrawnCutoutNoRoundEdges = 8
    seDrawnCutoutRoundCorners = 9
    seDrawnCutoutNoRoundCorners = 10


class LouverFeatureConstants:
    """constant.tlb > LouverFeatureConstants."""

    seLouverDepthDirectionLeft = 1
    seLouverDepthDirectionRight = 2
    seLouverDimensionOffset = 3
    seLouverDimensionFull = 4
    seLouverFormedEnd = 5
    seLouverLancedEnd = 6
    seLouverHeightNormal = 7
    seLouverHeightReverseNormal = 8
    seLouverRound = 9
    seLouverNoRound = 10


class GNTTypePropertyConstants:
    """geometry.tlb > GNTTypePropertyConstants.

    Values of Face.GeometryForm and Edge.GeometryForm. Reading that one
    property is what makes paged face queries cheap: the alternative is
    re-querying Body.Faces() once per geometry type.
    """

    igBody = 167551091
    igShell = 167551088
    igFace = 167551075
    igLoop = 167551097
    igEdgeUse = 167551099
    igEdge = 167551093
    igVertex = 167551101
    igBSplineSurface = 1465959633
    igCylinder = -114972029
    igCone = -114972031
    igPlane = -1909484335
    igMesh = -2071771273
    igSphere = -114972027
    igTorus = -114972025
    igBSplineCurve = 167551103
    igCircle = 167551105
    igEllipse = 167551107
    igLine = 167551109
    igPLine = 434530178
    igParamBSplineCurve = -1811952078
    igCurveBody = -1020639371
    igCurvePath = -1020639369
    igCurve = -1020639367
    igCurveVertex = -1020639365
    igShells = 167551078
    igFaces = 167551073
    igLoops = 167551080
    igEdgeUses = 167551095
    igEdges = 167551084
    igVertices = 167551086


class SeGradientType:
    """constant.tlb > SeGradientType."""

    seGradientTypeHorizontal = 1
    seGradientTypeVertical = 2
    seGradientTypeDiagonalUp = 3
    seGradientTypeDiagonalDown = 4
    seGradientTypeSquareSpot = 5
    seGradientTypeCircularSpot = 6
    seGradientTypeCustom = 7


class DimWeldTypeConstants:
    """constant.tlb > DimWeldTypeConstants (top-symbol values).

    The full enum also carries igDimWeldBottom* and modifier values; add them
    from the type library if a caller needs them.
    """

    igDimWeldTypeNone = 0
    igDimWeldTopFillet = 1
    igDimWeldTopSpot = 2
    igDimWeldTopSeam = 3
    igDimWeldTopBevel = 4
    igDimWeldTopVGroove = 5
    igDimWeldTopSlot = 6
    igDimWeldTopSquare = 7
    igDimWeldTopUGroove = 8
    igDimWeldTopFlangeEdge = 9
    igDimWeldTopFlangeCorner = 10
    igDimWeldTopBacking = 11
    igDimWeldTopJGroove = 12
    igDimWeldTopFlareV = 13
    igDimWeldTopFlareBevel = 14
    igDimWeldTopSurfacing = 15
    igDimWeldTopSteepFlankedBevel = 16
    igDimWeldTopSteepFlankedV = 17
    igDimWeldTopEdgeWeld = 18
    igDimWeldTopSurfaceJoint = 19
    igDimWeldTopInclinedJoint = 20
    igDimWeldTopFoldJoint = 21
    igDimWeldTopSingleVButt = 43
    igDimWeldTopSingleBevelButt = 44
    igDimWeldTopCompoundSquareGroove = 54
    igDimWeldTopCompoundJGroove = 55
    igDimWeldTopCompoundBevel = 56
    igDimWeldTopCompoundFlareBevel = 57
    igDimWeldTopMeltThrough = 62
    igDimWeldTopKeyhole = 65
    igDimWeldTopScarf = 67
    igDimWeldTopContinuationFillet = 70
    igDimWeldTopGroove = 73
    igDimWeldTopHYWeld = 75
    igDimWeldTopSingleUButt = 78
    igDimWeldTopPermRemBacking = 80
    igDimWeldTopConsumableInsert = 82
    igDimWeldTopCompoundFlareBevelDashed = 84


class AxisEndConstants:
    """Which end of an axis a helix grows from.

    Aliases of constant.tlb > FeaturePropertyConstants. Used for the AxisStart
    argument of the helix APIs, whose declared type is FeaturePropertyConstants.
    """

    igStart = 29
    igEnd = 30


class ThicknessSideConstants:
    """Which side of the profile a thin wall is offset to.

    Aliases of constant.tlb > FeaturePropertyConstants. Used for the
    ThicknessSide argument of the thin-wall APIs.
    """

    igInside = 4
    igOutside = 5


class NormalCutoutMethodConstants:
    """Method argument of the NormalCutouts.Add* calls.

    Aliases of constant.tlb > FeaturePropertyConstants. Normal cutouts are a
    sheet metal feature: on an ordinary part document the call succeeds but
    removes no material.
    """

    igSMClearanceCutout = 181
    igSMMidPlaneCutout = 182
    igSMFaceCutout = 205


class KeyPointExtentConstants:
    """constant.tlb > KeyPointExtentConstants."""

    igTangentNormal = 1
    igReverseTangentNormal = 2
    igInteriorTangentNormal = 3
    igInteriorReverseTangentNormal = 4


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
    seAssemblyGlobalTubeBendRadius = 2
    seAssemblyGlobalTubeOuterDiameter = 3
    seAssemblyGlobalTubeMinimumFlatLength = 4
    seAssemblyGlobalTubeEndTreatmentOutsideDiameter = 5
    seAssemblyGlobalTubeEndTreatmentInsideDiameter = 6
    seAssemblyGlobalTubeEndTreatmentDepth = 7
    seAssemblyGlobalTubeEndTreatmentAngle = 8
    seAssemblyGlobalTubeEndTreatmentRadius = 9
    seAssemblyGlobalDefaultPartDensity = 10
    seAssemblyGlobalDefaultAccuracyForPartDensity = 11
    seAssemblyGlobalWireHarnessDefaultSlackCompensation = 12
    seAssemblyGlobalWireHarnessDefaultHoleClearance = 13
    seAssemblyGlobalWireHarnessDefaultBundleClearance = 14
    seAssemblyGlobalWireHarnessDefaultWireAdder = 15
    seAssemblyGlobalWireHarnessDefaultCableAdder = 16
    seAssemblyGlobalWireHarnessDefaultBundleAdder = 17
    seAssemblyGlobalUpdatePhysicalPropertiesOnSave = 18
    seAssemblyGlobalAutomaticUpdate = 19
    seAssemblyGlobalAdjustableAsm = 20
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


class AssemblyFeaturePropertyConstants:
    """Assembly feature property constants (from type library)"""

    # Feature extent side
    igAssemblyFeatureBothSides = 0
    igAssemblyFeatureOneSide = 1

    # Feature profile side
    igAssemblyFeatureProfileLeft = 0
    igAssemblyFeatureProfileRight = 1
    igAssemblyFeatureProfileSymmetric = 2

    # Assembly feature types
    igAssemblyFeatureExtrudedCutout = 1
    igAssemblyFeatureRevolvedCutout = 2
    igAssemblyFeatureHole = 3
    igAssemblyFeatureExtrudedProtrusion = 4
    igAssemblyFeatureRevolvedProtrusion = 5
    igAssemblyFeatureMirror = 6
    igAssemblyFeaturePattern = 7


class PatternTypeConstants:
    """constant.tlb > PatternTypeConstants."""

    seSmartPattern = 0
    seFastPattern = 1


class PatternCurveAnchorSideConstants:
    """constant.tlb > PatternCurveAnchorSideConstants."""

    sePatternCurveLeftSide = 1
    sePatternCurveRightSide = 2


class PatternOffsetTypeConstants:
    """assembly.tlb > PatternOffsetTypeConstants."""

    sePatternFitOffset = 0
    sePatternFillOffset = 1
    sePatternFixedOffset = 2
    sePatternChordLengthOffset = 3


class PatternTransformTypeConstants:
    """constant.tlb > PatternTransformTypeConstants."""

    sePatternTransformLinear = 0
    sePatternTransformFullRotation = 1
    sePatternTransformProjectedRotation = 2
    sePatternTransformFullRotationFromSurface = 3


class PatternTransformRotateTypeConstants:
    """constant.tlb > PatternTransformRotateTypeConstants."""

    sePatternTransformRotateOnCurvePosition = 0
    sePatternTransformRotateOnFeaturePosition = 1
