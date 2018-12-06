' Author: Andy Perez
' Date: 2018-11
' Femap Version: 12.0

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Working View")
    App.feAppMessage(FCM_NORMAL, "Changes the settings of the active view.")

    Dim rc As Long

    Dim curView As femap.View
    Set curView = App.feView

    curView.Get(curView.Active)

    With curView
        .WindowShadeBack = True  'Use a shaded background
        .WindowShadeBackOpt = 6  'Use horizon option
        .WindowBackColor = 136
        .WindowBackColor2 = 132
        .WindowBackColor3 = 136
        .WindowShadeHorizon = 40

        '=============================
        ' Labels, Entities, and Colors
        '=============================
        .Draw(FVI_LABEL) = True
        .Label(FVI_LABEL) = 38
        .ColorMode(FVI_LABEL) = 1
        .Color(FVI_LABEL) = 149
        .RenderPushLabel = 1

        .Draw(FVI_CSYS) = False
        .Label(FVI_CSYS) = 1
        .ColorMode(FVI_CSYS) = 7
        .Color(FVI_CSYS) = 60

        .Draw(FVI_POINT) = False
        .Label(FVI_POINT) = 0
        .ColorMode(FVI_POINT) = 0
        .Color(FVI_POINT) = 24

        .Draw(FVI_CURVE) = False
        .Label(FVI_CURVE) = 0
        .ColorMode(FVI_CURVE) = 0
        .Color(FVI_CURVE) = 120

        .Draw(FVI_COMPOSITE_CURVE) = False
        .Label(FVI_COMPOSITE_CURVE) = 0
        .ColorMode(FVI_COMPOSITE_CURVE) = 0
        .Color(FVI_COMPOSITE_CURVE) = 20590

        .Draw(FVI_CURVE_MESHSIZE) = False
        .Label(FVI_CURVE_MESHSIZE) = 1
        .ColorMode(FVI_CURVE_MESHSIZE) = 1
        .Color(FVI_CURVE_MESHSIZE) = 123

        .ColorMode(FVI_CURVE_ACCURACY) = 0

        .Draw(FVI_SURFACE) = False
        .Label(FVI_SURFACE) = 0
        .ColorMode(FVI_SURFACE) = 3
        .Color(FVI_SURFACE) = 60

        .Draw(FVI_BOUNDARY) = False
        .Label(FVI_BOUNDARY) = 0
        .ColorMode(FVI_BOUNDARY) = 3
        .Color(FVI_BOUNDARY) = 24648

        .Draw(FVI_VOLUME) = False
        .Label(FVI_VOLUME) = 0
        .ColorMode(FVI_VOLUME) = 3
        .Color(FVI_VOLUME) = 66

        .Draw(FVI_MESH_POINT) = True
        .Label(FVI_MESH_POINT) = 0
        .ColorMode(FVI_MESH_POINT) = 0
        .Color(FVI_MESH_POINT) = 120

        .Draw(FVI_TEXT) = True
        .Label(FVI_TEXT) = 0
        .ColorMode(FVI_TEXT) = 0
        .Color(FVI_TEXT) = 149

        .Draw(FVI_NODE) = False
        .Label(FVI_NODE) = 0
        .ColorMode(FVI_NODE) = 1
        .Color(FVI_NODE) = 20526

        .Draw(FVI_NODE_PERMCONSTRAINT) = False
        .Label(FVI_NODE_PERMCONSTRAINT) = 1
        .ColorMode(FVI_NODE_PERMCONSTRAINT) = 0
        .Color(FVI_NODE_PERMCONSTRAINT) = 120

        .Draw(FVI_ELEMENT) = True
        .Label(FVI_ELEMENT) = 0
        .ColorMode(FVI_ELEMENT) = 3
        .Color(FVI_ELEMENT) = 124

        .Draw(FVI_ELEMENT_DIRECTION) = False
        .Label(FVI_ELEMENT_DIRECTION) = 1
        .ColorMode(FVI_ELEMENT_DIRECTION) = 1
        .Color(FVI_ELEMENT_DIRECTION) = 0

        .Draw(FVI_ELEMENT_OFFSET) = False
        .Label(FVI_ELEMENT_OFFSET) = 0
        .ColorMode(FVI_ELEMENT_OFFSET) = 1
        .Color(FVI_ELEMENT_OFFSET) = 14

        .Draw(FVI_ELEMENT_ORIENT) = False
        .Label(FVI_ELEMENT_ORIENT) = 0
        .ColorMode(FVI_ELEMENT_ORIENT) = 0
        .Color(FVI_ELEMENT_ORIENT) = 62

        .Draw(FVI_BEAM_YAXIS) = False
        .ColorMode(FVI_BEAM_YAXIS) = 0
        .Color(FVI_BEAM_YAXIS) = 62

        .Draw(FVI_ELEM_CSYS) = False
        .Label(FVI_ELEM_CSYS) = 1
        .ColorMode(FVI_ELEM_CSYS) = 2
        .Color(FVI_ELEM_CSYS) = 124

        .Draw(FVI_ELEM_MATL_DIR) = True
        .ColorMode(FVI_ELEM_MATL_DIR) = 0
        .Color(FVI_ELEM_MATL_DIR) = 124

        .Draw(FVI_ELEMENT_WELD) = False

        .Draw(FVI_RIGID) = True
        .Label(FVI_RIGID) = 0

        .Draw(FVI_ELEMENT_COHESIVE) = True
        .Label(FVI_ELEMENT_COHESIVE) = 1
        .ColorMode(FVI_ELEMENT_COHESIVE) = 0
        .Color(FVI_ELEMENT_COHESIVE) = 140

        .Draw(FVI_ELEMENT_SHELL) = False

        .Draw(FVI_LOAD_VECTORS) = False
        .Label(FVI_LOAD_VECTORS) = 1
        .LoadVecLength = 100
        .LoadVectorMinScale = 25

        .Draw(FVI_LOAD_BODY) = True
        .Label(FVI_LOAD_BODY) = 1
        .ColorMode(FVI_LOAD_BODY) = 3
        .Color(FVI_LOAD_BODY) = 12326

        .Draw(FVI_LOAD_FORCE) = False
        .Label(FVI_LOAD_FORCE) = 1
        .ColorMode(FVI_LOAD_FORCE) = 0
        .Color(FVI_LOAD_FORCE) = 10

        .Draw(FVI_LOAD_MOMENT) = False
        .Label(FVI_LOAD_MOMENT) = 1
        .ColorMode(FVI_LOAD_MOMENT) = 0
        .Color(FVI_LOAD_MOMENT) = 52

        .Draw(FVI_LOAD_THERMAL) = False
        .Label(FVI_LOAD_THERMAL) = 1
        .ColorMode(FVI_LOAD_THERMAL) = 0
        .Color(FVI_LOAD_THERMAL) = 4

        .Draw(FVI_LOAD_DISTRIBUTEDLOAD) = False
        .Label(FVI_LOAD_DISTRIBUTEDLOAD) = 1
        .ColorMode(FVI_LOAD_DISTRIBUTEDLOAD) = 0
        .Color(FVI_LOAD_DISTRIBUTEDLOAD) = 90

        .Draw(FVI_LOAD_PRESSURE) = False
        .Label(FVI_LOAD_PRESSURE) = 1
        .ColorMode(FVI_LOAD_PRESSURE) = 0
        .Color(FVI_LOAD_PRESSURE) = 120

        .Draw(FVI_LOAD_ACCELERATION) = False
        .Label(FVI_LOAD_ACCELERATION) = 1
        .ColorMode(FVI_LOAD_ACCELERATION) = 0
        .Color(FVI_LOAD_ACCELERATION) = 12

        .Draw(FVI_LOAD_VELOCITY) = False
        .Label(FVI_LOAD_VELOCITY) = 1
        .ColorMode(FVI_LOAD_VELOCITY) = 0
        .Color(FVI_LOAD_VELOCITY) = 14

        .Draw(FVI_LOAD_ENFORCEDDISP) = False
        .Label(FVI_LOAD_ENFORCEDDISP) = 1
        .ColorMode(FVI_LOAD_ENFORCEDDISP) = 0
        .Color(FVI_LOAD_ENFORCEDDISP) = 2

        .Draw(FVI_LOAD_NONLINEARFORCE) = False
        .Label(FVI_LOAD_NONLINEARFORCE) = 1
        .ColorMode(FVI_LOAD_NONLINEARFORCE) = 0
        .Color(FVI_LOAD_NONLINEARFORCE) = 8261

        .Draw(FVI_LOAD_HEATGEN) = False
        .Label(FVI_LOAD_HEATGEN) = 1
        .ColorMode(FVI_LOAD_HEATGEN) = 0
        .Color(FVI_LOAD_HEATGEN) = 100

        .Draw(FVI_LOAD_HEATFLUX) = False
        .Label(FVI_LOAD_HEATFLUX) = 1
        .ColorMode(FVI_LOAD_HEATFLUX) = 0
        .Color(FVI_LOAD_HEATFLUX) = 24

        .Draw(FVI_LOAD_CONVECTION) = False
        .Label(FVI_LOAD_CONVECTION) = 1
        .ColorMode(FVI_LOAD_CONVECTION) = 0
        .Color(FVI_LOAD_CONVECTION) = 60

        .Draw(FVI_LOAD_RADIATION) = False
        .Label(FVI_LOAD_RADIATION) = 1
        .ColorMode(FVI_LOAD_RADIATION) = 0
        .Color(FVI_LOAD_RADIATION) = 104

        .Draw(FVI_BOLT_PRELOAD) = False
        .Label(FVI_BOLT_PRELOAD) = 1
        .ColorMode(FVI_BOLT_PRELOAD) = 0
        .Color(FVI_BOLT_PRELOAD) = 48

        .Draw(FVI_LOAD_FLUIDTRACKING) = False
        .Label(FVI_LOAD_FLUIDTRACKING) = 1
        .ColorMode(FVI_LOAD_FLUIDTRACKING) = 0
        .Color(FVI_LOAD_FLUIDTRACKING) = 20

        .Draw(FVI_LOAD_UNKNOWN) = False
        .ColorMode(FVI_LOAD_UNKNOWN) = 0
        .Color(FVI_LOAD_UNKNOWN) = 34

        .Draw(FVI_LOAD_SLIPWALL) = False
        .ColorMode(FVI_LOAD_SLIPWALL) = 0
        .Color(FVI_LOAD_SLIPWALL) = 9

        .Draw(FVI_LOAD_FANCURVE) = False
        .Label(FVI_LOAD_FANCURVE) = 1
        .ColorMode(FVI_LOAD_FANCURVE) = 0
        .Color(FVI_LOAD_FANCURVE) = 49

        .Draw(FVI_LOAD_PERIODIC) = False
        .ColorMode(FVI_LOAD_PERIODIC) = 0
        .Color(FVI_LOAD_PERIODIC) = 105

        .Draw(FVI_CONSTRAINT) = False
        .Label(FVI_CONSTRAINT) = 3
        .ColorMode(FVI_CONSTRAINT) = 3
        .Color(FVI_CONSTRAINT) = 120

        .Draw(FVI_CONSTRAINT_EQUATION) = False
        .Label(FVI_CONSTRAINT_EQUATION) = 1
        .ColorMode(FVI_CONSTRAINT_EQUATION) = 0
        .Color(FVI_CONSTRAINT_EQUATION) = 8312

        .Draw(FVI_CONNECTOR) = False
        .Label(FVI_CONNECTOR) = 0
        .ColorMode(FVI_CONNECTOR) = 0
        .Color(FVI_CONNECTOR) = 14

        .Draw(FVI_CONNECTION_REGION) = False
        .Label(FVI_CONNECTION_REGION) = 0
        .ColorMode(FVI_CONNECTION_REGION) = 0
        .Color(FVI_CONNECTION_REGION) = 20488

        .Draw(FVI_INTERIOR_POINT) = False
        .ColorMode(FVI_INTERIOR_POINT) = 1
        .Color(FVI_INTERIOR_POINT) = 97

        .Draw(FVI_INTERIOR_CURVE) = False
        .ColorMode(FVI_INTERIOR_CURVE) = 1
        .Color(FVI_INTERIOR_CURVE) = 8252

        .Draw(FVI_INTERIOR_SURFACE) = False
        .ColorMode(FVI_INTERIOR_SURFACE) = 0
        .Color(FVI_INTERIOR_SURFACE) = 66

        .Draw(FVI_AERO_PANEL) = False
        .Label(FVI_AERO_PANEL) = 0
        .ColorMode(FVI_AERO_PANEL) = 0
        .Color(FVI_AERO_PANEL) = 28676

        .Draw(FVI_AERO_MESH) = False
        .Label(FVI_AERO_MESH) = 0
        .ColorMode(FVI_AERO_MESH) = 3
        .Color(FVI_AERO_MESH) = 4

        .Draw(FVI_AERO_INTERFERENCE) = False
        .Label(FVI_AERO_INTERFERENCE) = 1
        .ColorMode(FVI_AERO_INTERFERENCE) = 0
        .Color(FVI_AERO_INTERFERENCE) = 124

        .Draw(FVI_AERO_SPLINE) = False
        .Label(FVI_AERO_SPLINE) = 0
        .ColorMode(FVI_AERO_SPLINE) = 0
        .Color(FVI_AERO_SPLINE) = 24696

        .Draw(FVI_AERO_SURFACE) = False
        .Label(FVI_AERO_SURFACE) = 0
        .ColorMode(FVI_AERO_SURFACE) = 0
        .Color(FVI_AERO_SURFACE) = 24590

        '====================
        'Tools and View Style
        '====================
        .Draw(FVI_FREE_EDGE) = False
        .Label(FVI_FREE_EDGE) = 0
        .ColorMode(FVI_FREE_EDGE) = 0
        .Color(FVI_FREE_EDGE) = 24600

        .Draw(FVI_SHRINK_ELEM) = False
        .ShrinkFactor = 70

        .Draw(FVI_FILL_ELEM) = True
        .Label(FVI_FILL_ELEM) = 0
        .ColorMode(FVI_FILL_ELEM) = 0

        .Draw(FVI_FILLED_EDGES) = True
        .Label(FVI_FILLED_EDGES) = 0
        .ColorMode(FVI_FILLED_EDGES) = 3
        .Color(FVI_FILLED_EDGES) = 0

        .Draw(FVI_FEATURE_LINE) = True
        .Label(FVI_FEATURE_LINE) = 1
        .ColorMode(FVI_FEATURE_LINE) = 1
        .Color(FVI_FEATURE_LINE) = 28676
        .FeatureLineAngle = 5

        .Draw(FVI_RENDER_OPTIONS) = False
        .Label(FVI_RENDER_OPTIONS) = 1
        .ColorMode(FVI_RENDER_OPTIONS) = 0
        .RenderPushForward = 50  ' offset factor
        .RenderPushUnit = 100  ' offset

        .Draw(FVI_PERF_GRAPHICS) = False
        .Label(FVI_PERF_GRAPHICS) = 3
        .ColorMode(FVI_PERF_GRAPHICS) = 0
        '.FilledEdgeOffset = 1000  ' unknown property

        .Draw(FVI_TRANSPARENCY) = False
        .TransparencyPercent = 75

        .Draw(FVI_SHADING) = True
        .Label(FVI_SHADING) = 0
        .ColorMode(FVI_SHADING) = 3
        .Color(FVI_SHADING) = 124
        .AmbientLight = 25
        .SpotlightDist = 1000

        .Draw(FVI_PERSPECTIVE) = False
        .PerspectiveDist = 22.5
        .PerspectiveAngle = 22.5

        ' not used
        '.Draw(FVI_STEREO)
        '.Label(FVI_STEREO)
        '.ColorMode(FVI_STEREO)
        '.Color(FVI_STEREO)

        .Draw(FVI_AXISYM_AXIS) = False
        .Label(FVI_AXISYM_AXIS) = 4
        .ColorMode(FVI_AXISYM_AXIS) = 0
        .Color(FVI_AXISYM_AXIS) = 60

        .Draw(FVI_VIEW_LEGEND) = False
        .Label(FVI_VIEW_LEGEND) = 0
        .ColorMode(FVI_VIEW_LEGEND) = 0
        .Color(FVI_VIEW_LEGEND) = 149

        .Draw(FVI_VIEW_AXES) = True
        .Label(FVI_VIEW_AXES) = 1
        .ColorMode(FVI_VIEW_AXES) = 5
        .Color(FVI_VIEW_AXES) = 60
        '.AxisSize = 13  ' unknown property name

        .Draw(FVI_ORIGIN) = True
        .Color(FVI_ORIGIN) = 56

        .Draw(FVI_WORKPLANE) = False
        .Label(FVI_WORKPLANE) = 0
        .ColorMode(FVI_WORKPLANE) = 1
        .Color(FVI_WORKPLANE) = 2328

        .Draw(FVI_WORKPLANE_GRID) = False
        .Color(FVI_WORKPLANE_GRID) = 8216

        .Draw(FVI_CLIPPING_PLANE) = False
        .Color(FVI_CLIPPING_PLANE) = 4

        .Draw(FVI_SYMBOL) = True
        .Label(FVI_SYMBOL) = 3
        .ColorMode(FVI_SYMBOL) = 2
        .LoadVecLength = 100
        .OtherVecLength = 100

        .Label(FVI_PREVIEW) = 2
        .Color(FVI_PREVIEW) = 0

        .Draw(FVI_ASPECT_RATIO) = True
        .Aspect = 1  ' aspect ratio

        .Draw(FVI_CLIP_PLANE) = False
        .Label(FVI_CLIP_PLANE) = 0

        '==============
        'PostProcessing
        '==============
        .Draw(FVI_POST_TITLES) = True
        .Label(FVI_POST_TITLES) = 2
        .ColorMode(FVI_POST_TITLES) = 3
        .Color(FVI_POST_TITLES) = 149

        .Label(FVI_DEFORMED_STYLE) = True
        .ColorMode(FVI_DEFORMED_STYLE) = 1
        .DeformedRelScale = 3  ' scale %
        .DeformedAbsScale = 1  ' scale act

        .Draw(FVI_VECTOR_STYLE) = True
        .Label(FVI_VECTOR_STYLE) = 0
        .ColorMode(FVI_VECTOR_STYLE) = 0
        .Color(FVI_VECTOR_STYLE) = 24
        .ArrowsToLabel = 24  ' Top Percent

        .Draw(FVI_ANIMATED_STYLE) = False
        .Label(FVI_ANIMATED_STYLE) = 3
        .ColorMode(FVI_ANIMATED_STYLE) = 0
        .Color(FVI_ANIMATED_STYLE) = 124
        .AnimationFrames = 24
        .AnimationDelay = 101  ' in ms

        .Label(FVI_DEFORMED) = 0
        .ColorMode(FVI_DEFORMED) = 0
        .Color(FVI_DEFORMED) = 100
        .DeformRelativeToNode = 0

        .Draw(FVI_UNDEFORMED) = False
        .ColorMode(FVI_UNDEFORMED) = 1
        .Color(FVI_UNDEFORMED) = 100
        .RenderPushUndeformed = -100  ' render offset %

        .Draw(FVI_TRACE_STYLE) = False
        .Label(FVI_TRACE_STYLE) = 0
        .ColorMode(FVI_TRACE_STYLE) = 0
        .Color(FVI_TRACE_STYLE) = 19

        .Draw(FVI_CONTOUR_TYPE) = True
        .Label(FVI_CONTOUR_TYPE) = 2
        .ColorMode(FVI_CONTOUR_TYPE) = 0
        .ContourElemUnaveraged = True
        .ContourGroup = True
        .ContourDataSelection = 1  ' none/visible groups

        .Draw(FVI_CONTOUR_STYLE) = True
        .Label(FVI_CONTOUR_STYLE) = 2
        .ColorMode(FVI_CONTOUR_STYLE) = 1
        .ContourLabelFreq = 5
        .ContourLabelDigits = 0

        .Draw(FVI_CONTOUR_LEVELS) = True
        .Label(FVI_CONTOUR_LEVELS) = 0
        .ColorMode(FVI_CONTOUR_LEVELS) = 0
        .Color(FVI_CONTOUR_LEVELS) = 10
        .UserContourValue(11) = 0
        .UserContourValue(0) = 0

        .Draw(FVI_CONTOUR_LEGEND) = True
        .Label(FVI_CONTOUR_LEGEND) = 7
        .ColorMode(FVI_CONTOUR_LEGEND) = 1
        .Color(FVI_CONTOUR_LEGEND) = 149
        .ContourLegendFreq = 0
        .ContourLegendDigits = 0

        .Draw(FVI_CRITERIA_LIMITS) = False
        .Label(FVI_CRITERIA_LIMITS) = 0
        .CriteriaMin = 0
        .CriteriaMax = 0

        .Draw(FVI_CRITERIA_PASS) = True
        .Label(FVI_CRITERIA_PASS) = 1
        .ColorMode(FVI_CRITERIA_PASS) = 2
        .Color(FVI_CRITERIA_PASS) = 124

        .Draw(FVI_CRITERIA_FAIL) = False
        .Label(FVI_CRITERIA_FAIL) = 0
        .ColorMode(FVI_CRITERIA_FAIL) = 1
        .Color(FVI_CRITERIA_FAIL) = 8254

        .Draw(FVI_BEAM_DIAGRAM) = False
        .Label(FVI_BEAM_DIAGRAM) = 0
        .ColorMode(FVI_BEAM_DIAGRAM) = 0
        .Color(FVI_BEAM_DIAGRAM) = 62
        .BeamDiagramScale = 10

        .Draw(FVI_ISOSURFACE) = False
        .Label(FVI_ISOSURFACE) = 0
        .ColorMode(FVI_ISOSURFACE) = 0
        .Color(FVI_ISOSURFACE) = 4
        .IsoSurfaceValue = 0

        .Draw(FVI_ISOLINE) = False
        .Label(FVI_ISOLINE) = 1
        .ColorMode(FVI_ISOLINE) = 0
        .Color(FVI_ISOLINE) = 0

        .Draw(FVI_STREAMLINE) = False
        .ColorMode(FVI_STREAMLINE) = 0
        .Color(FVI_STREAMLINE) = 4
        .StreamLineMaxLenFactor = 4
        .StreamLineMinSpeedPct = 0.01

        .Draw(FVI_CONTOUR_VECTOR) = True
        .Label(FVI_CONTOUR_VECTOR) = 7
        .ColorMode(FVI_CONTOUR_VECTOR) = 0
        .ContourVecDigits = 1.E-8
        .ContourVecLength = 100.

        .Draw(FVI_CONTOUR_ARROW_OPT) = True
        .Label(FVI_CONTOUR_ARROW_OPT) = 1
        .ColorMode(FVI_CONTOUR_ARROW_OPT) = 0

        .Draw(FVI_FREEBODY) = True
        .Label(FVI_FREEBODY) = 0

        .Draw(FVI_FREEBODY_NODEMARKER) = True
        .Label(FVI_FREEBODY_NODEMARKER) = 2
        .ColorMode(FVI_FREEBODY_NODEMARKER) = 0
        .Color(FVI_FREEBODY_NODEMARKER) = 20

        .Draw(FVI_FREEBODY_VECTOR) = True
        .Label(FVI_FREEBODY_VECTOR) = 1
        .ColorMode(FVI_FREEBODY_VECTOR) = 0
        .FreebodyUnitsPerLength = 100.
        .FreebodyLength = 100.

        .Label(FVI_FREEBODY_TOTALF) = 2
        .ColorMode(FVI_FREEBODY_TOTALF) = 2
        .Color(FVI_FREEBODY_TOTALF) = 15
        .FreebodyTotalForceFactor = 1.

        .Label(FVI_FREEBODY_TOTALM) = 2
        .ColorMode(FVI_FREEBODY_TOTALM) = 2
        .Color(FVI_FREEBODY_TOTALM) = 15
        .FreebodyTotalMomentFactor = 1.

        .Draw(FVI_FREEBODY_NODAL_F) = 0
        .Label(FVI_FREEBODY_NODAL_F) = 2
        .ColorMode(FVI_FREEBODY_NODAL_F) = 2
        .Color(FVI_FREEBODY_NODAL_F) = 24
        .FreebodyNodalForceFactor = 1.

        .Draw(FVI_FREEBODY_NODAL_M) = 0
        .Label(FVI_FREEBODY_NODAL_M) = 2
        .ColorMode(FVI_FREEBODY_NODAL_M) = 2
        .Color(FVI_FREEBODY_NODAL_M) = 24
        .FreebodyNodalMomentFactor = 1.

        .Put(.Active)

    End With

    App.feViewRegenerate(0)

End Sub