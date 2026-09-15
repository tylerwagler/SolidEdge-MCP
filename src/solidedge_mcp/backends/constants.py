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
    """constant.tlb > ViewOrientationConstants.

    This is the enum ``DrawingViews.AddPartView`` and
    ``DrawingView.SetViewOrientationStandard`` declare. The drawing code used
    to carry its own ``DrawingViewOrientationConstants`` instead, with Front=5
    and Top=6, excused in the constants test as "empirically verified". Solid
    Edge 2026 disagrees: ``DrawingView.ViewOrientation()`` reads back exactly
    the value passed in, 4 for a front view and 1 for a top view. Front=5 is
    igBottomView, so every "Front" drawing view this server made was a bottom
    view and every "Top" a back view.
    """

    igTopView = 1
    igRightView = 2
    igLeftView = 3
    igFrontView = 4
    igBottomView = 5
    igBackView = 6
    # Pictorial views continue the same numbering. Solid Edge's own isometric
    # view looks from the top-front-right corner.
    igTopBackLeftView = 7
    igTopFrontLeftView = 8
    igTopFrontRightView = 9
    igTopBackRightView = 10


class PartDrawingViewTypeConstants:
    """draft.tlb > PartDrawingViewTypeConstants."""

    sePartDesignedView = 0
    sePartSimplifiedView = 1


class AssemblyDrawingViewTypeConstants:
    """draft.tlb > AssemblyDrawingViewTypeConstants."""

    seAssemblyDesignedView = 0
    seAssemblySimplifiedView = 1
    seAssemblyConfigurationSimplifiedView = 2


class DrawingViewTypeConstants:
    """Drawing view type constants (from type library - separate enum)"""

    igPrincipleView = 1
    igIsometricView = 2
    igAuxiliaryView = 3
    igXSectionView = 4
    igDetailView = 5
    igIsoXSectionView = 6


class DraftPrintOrientationConstants:
    """draft.tlb > DraftPrintOrientationConstants.

    The print code used to pass 1 for Portrait and 2 for Landscape, calling
    them "typical COM constants". Portrait is 0 and Landscape is 1, so
    "Portrait" set landscape, and 2 is not a member at all -- Solid Edge 2026
    rejects it and leaves the orientation as it was.
    """

    igDraftPrintPortrait = 0
    igDraftPrintLandscape = 1


class VariableNameBy:
    """constant.tlb > VariableNameBy: which name Variables.Query matches on."""

    seVariableNameByUser = 0
    seVariableNameBySystem = 1
    seVariableNameByBoth = 2


class seVariableTypeConstants:  # noqa: N801
    """framewrk.tlb > seVariableTypeConstants (named to match the enum).

    Variables.Query's VarType takes one of these. It used to be passed 0,
    which is not a member of this enum, and Solid Edge answered every query
    with an empty collection -- so variable search never returned anything.
    There is no "all" member; query each type and merge.
    """

    seVariableType_Dimension = 1661573600
    seVariableType_UserDefined = 1560616706
    seVariableType_Simulation = 215773802
    seVariableType_Text = -170730141


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
    """Assembly feature side constants.

    These are NOT a type library enum of their own -- there is no
    AssemblyFeaturePropertyConstants in any Solid Edge library. Every side
    argument on AssemblyFeatures*.Add is typed FeaturePropertyConstants, the
    same enum the part-level features use, so the values below are its
    members and match DirectionConstants.

    They were previously an invented 0/1/2 numbering: "Left" was 0, which is
    igNullConstant and not a side at all, and "Right" and "Symmetric" were
    each one short of the member they named. Verified on Solid Edge 2026,
    where Add accepts any of 1, 2, 3 and 6 without complaint -- so this was
    never going to surface as an error, only as the wrong side.
    """

    # Feature extent side (FeaturePropertyConstants)
    igAssemblyFeatureOneSide = 2  # igRight
    igAssemblyFeatureBothSides = 3  # igSymmetric

    # Feature profile side (FeaturePropertyConstants)
    igAssemblyFeatureProfileLeft = 1  # igLeft
    igAssemblyFeatureProfileRight = 2  # igRight
    igAssemblyFeatureProfileSymmetric = 3  # igSymmetric

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


class MatTablePropIndexConstants:
    """Property indices for MatTable.GetMaterialPropValueFromDoc.

    From Program/constant.tlb > MatTablePropIndexConstants. The numbers are
    not 0..8 as this module once assumed; they start at 3 and jump to 20.
    """

    seMaterialName = 3
    seFaceStyle = 20
    seFillStyle = 21
    seVSPlusStyle = 22
    seDensity = 23
    seCoefOfThermalExpansion = 24
    seThermalConductivity = 25
    seSpecificHeat = 26
    seModulusElasticity = 27
    sePoissonRatio = 28
    seYieldStress = 29
    seUltimateStress = 30
    seElongation = 31


class FeatureStatusConstants:
    """Feature.Status values, from constant.tlb > FeatureStatusConstants."""

    igFeatureOK = 1216476310
    igFeatureFailed = 1216476311
    igFeatureWarned = 1216476312
    igFeatureSuppressed = 1216476313
    igFeatureRolledBack = 1216476314


#: Feature.Type values, from constant.tlb > FeatureTypeConstants (129 members,
#: a few sharing a value; the first name wins). A caller cannot do anything
#: with a bare 462094706, so list_features reports the name alongside it.
FEATURE_TYPE_NAMES: dict[int, str] = {
    -2101998503: "emboss",
    -2101194894: "swept protrusion",
    -2099521208: "bead",
    -1994967864: "weld pattern",
    -1974090952: "united body",
    -1934200528: "revolved surface",
    -1906466896: "bend",
    -1901787677: "weld round",
    -1887900773: "resize hole",
    -1871757644: "copied part",
    -1752010637: "flange",
    -1746885455: "asm extruded cutout",
    -1625749114: "weld mirror",
    -1621393684: "face rotate",
    -1575914081: "extend surface",
    -1553526524: "recovered body",
    -1489778563: "body",
    -1483539936: "surface by boundary",
    -1472076592: "weld extruded cutout",
    -1470838992: "tab",
    -1468087919: "user defined pattern",
    -1444717528: "intersection point",
    -1444717526: "project curve",
    -1431348508: "live section",
    -1385224025: "split face",
    -1364719728: "blank",
    -1223346452: "parting split",
    -1204891230: "helix protrusion",
    -1199320480: "tube",
    -1132312974: "model copy",
    -1106524128: "pattern copy geometry",
    -1105952133: "weld bead by revolved protrusion",
    -1078402089: "delete hole",
    -1072882524: "label weld",
    -1042199833: "feature group",
    -1033222194: "weld revolved cutout",
    -869608965: "asm pattern",
    -848873917: "boolean",
    -818300041: "wire",
    -615522252: "break corner",
    -591258000: "mid surface",
    -573089291: "weld hole",
    -504033272: "rebend",
    -489073541: "asm label weld",
    -483645223: "lip",
    -461605375: "fillet weld",
    -416228998: "pattern",
    -398746894: "swept cutout",
    -396962225: "stitch surface",
    -326994721: "interpart construction",
    -313264769: "relief patch",
    -292547215: "normal cutout",
    -266365000: "asm stitch weld",
    -109105343: "blue surf",
    -85880079: "vent",
    -48339624: "tab and slot",
    45018492: "asm mirror",
    66247736: "mirror copy",
    96955616: "copy construction",
    97588292: "derived curve",
    138029688: "close corner",
    144813157: "asm swept protrusion",
    176240055: "pattern part",
    276679016: "split curve",
    281089316: "contour flange",
    313074017: "weld bead by swept protrusion",
    316506019: "louver",
    339062508: "asm revolved cutout",
    339115113: "slot group",
    395285420: "jog",
    421475437: "lofted surface",
    421475439: "swept surface",
    424701353: "parting surface",
    438063868: "lofted flange",
    438630050: "thin region",
    462094706: "extruded protrusion",
    462094710: "revolved protrusion",
    462094714: "extruded cutout",
    462094718: "revolved cutout",
    462094722: "hole",
    462094730: "rib",
    462094734: "thinwall",
    462094738: "round",
    462094742: "chamfer",
    462094746: "draft",
    483231468: "asm thread",
    499918320: "extruded surface",
    611936790: "drawn cutout",
    630099633: "replace face",
    721904086: "dimple",
    731482635: "asm fillet weld",
    741566449: "weld chamfer",
    758717329: "gusset",
    759182187: "delete blend",
    785064758: "normal to face cutout",
    785064761: "normal to face protrusion",
    839887038: "wrap sketch",
    841561903: "intersect surface",
    963690016: "resize round",
    980322967: "mirror copy geometry",
    990795800: "thicken",
    1036932736: "weld bead by extruded protrusion",
    1100034230: "face move",
    1106163506: "resize bend",
    1180468550: "face offset",
    1189339109: "coordinate system",
    1197717883: "helix cutout",
    1323346977: "unbend",
    1336556513: "slot",
    1484085415: "blank surface",
    1498218798: "asm groove weld",
    1501757264: "intersection curve",
    1532784456: "asm hole",
    1541369857: "asm extruded protrusion",
    1587872528: "offset surface",
    1647712262: "match flange face",
    1718424353: "web network",
    1752082343: "copy surface",
    1786053696: "etch",
    1854650269: "duplicate",
    1908287958: "mirror part",
    1917772600: "trim surface",
    2009702983: "redefine face",
    2057842144: "lofted protrusion",
    2057842149: "lofted cutout",
    2079590812: "asm revolved protrusion",
    2113150980: "delete region",
    2113150981: "delete face",
    2137685236: "assembly weldment",
}


#: Document.Type values, from constant.tlb > DocumentTypeConstants. Reported
#: by name because a bare 4 does not say "sheet metal".
DOCUMENT_TYPE_NAMES: dict[int, str] = {
    1: "part",
    2: "draft",
    3: "assembly",
    4: "sheet metal",
    5: "unknown",
    6: "weldment",
    7: "weldment assembly",
    8: "synchronous part",
    9: "synchronous sheet metal",
    10: "synchronous assembly",
    11: "assembly viewer",
    12: "part viewer",
    13: "draft viewer",
}
