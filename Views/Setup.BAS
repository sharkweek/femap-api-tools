Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Setup View")

    Dim viewID As Long
    Dim rc As femap.zReturnCode

    With App
        ' FEMAP Object Properties
        .Info_GroupAutomaticAdd = -1     ' Automatic add to active group
        '.Pref_AltContourPalette = True  ' Use alternate color palette for contour plots
        .feAppGetActiveView(viewID)
    End With

    Dim View As femap.View
    Set View = App.feView

    With View
        .Get(viewID)

        '.Draw
        'Labels, Entities, & Colors
        .Draw(FVI_LABEL)=False                   'Label Parameters (Index=0)
        .Draw(FVI_CSYS)=False                     'Coordinate System (Index=1)
        .Draw(FVI_POINT)=False                   'Point (Index=2)
        .Draw(FVI_CURVE)=False                   'Curve (Index=3)
        .Draw(FVI_COMPOSITE_CURVE)=False         'Combined Curve (Index=95)
        .Draw(FVI_CURVE_MESHSIZE)=False          'Curve - Mesh Size (Index=24)
        .Draw(FVI_CURVE_ACCURACY)=False          'Curve/Surface Directions (Index=39)
        .Draw(FVI_SURFACE)=False                 'Surface (Index=4)
        .Draw(FVI_VOLUME) =False                 'Volume (Index=5)
        .Draw(FVI_TEXT) =False                   'Text (Index=6)
        .Draw(FVI_BOUNDARY)=False                'Boundary (Index=27)
        .Draw(FVI_NODE)=False                     'Node (Index=7)
        .Draw(FVI_NODE_PERMCONSTRAINT)=False     'Node - Perm Constraint (Index=8)
        .Draw(FVI_ELEMENT)=True                   'Element (Index=9)
        .Draw(FVI_ELEMENT_DIRECTION)=False       'Element - Directions (Index=10)
        .Draw(FVI_ELEMENT_OFFSET)=False           'Element - Offsets/Releases (Index=11)
        .Draw(FVI_ELEMENT_ORIENT)=False          'Element - Orientation/Shape (Index=12)
        .Draw(FVI_BEAM_YAXIS)=False              'Element - Beam Y-Axis (Index=13)
        .Draw(FVI_ELEM_CSYS)=False               'Element - Coordinate System (Index=96)
        .Draw(FVI_ELEM_MATL_DIR)=False           'Element - Material Direction (Index=115)
        .Draw(FVI_ELEMENT_WELD)=False            'Element - Weld (Index=90)
        .Draw(FVI_RIGID)=True                   'Element - Rigid (Index=94)
        .Draw(FVI_LOAD_VECTORS)=False            'Load Vectors (Index=78)
        .Draw(FVI_LOAD_BODY)=False               'Load Body (Index=100)
        .Draw(FVI_LOAD_FORCE)=False              'Load Force (Index=14)
        .Draw(FVI_LOAD_MOMENT)=False             'Load Moment and Torque (Index=15)
        .Draw(FVI_LOAD_THERMAL)=False            'Load Temperature (Index=16)
        .Draw(FVI_LOAD_DISTRIBUTEDLOAD)=False    'Load Distributed (Index=71)
        .Draw(FVI_LOAD_PRESSURE)=False           'Load Pressure (Index=17)
        .Draw(FVI_LOAD_ACCELERATION)=False       'Load Acceleration (Index=18)
        .Draw(FVI_LOAD_VELOCITY)=False           'Load Velocity (Index=72)
        .Draw(FVI_LOAD_ENFORCEDDISP)=False       'Load Enforced Displacement (Index=19)
        .Draw(FVI_LOAD_NONLINEARFORCE)=False     'Load Nonlinear Force (Index=73)
        .Draw(FVI_LOAD_HEATGEN)=False            'Load Heat Generation (Index=66)
        .Draw(FVI_LOAD_HEATFLUX)=False           'Load Heat Flux (Index=67)
        .Draw(FVI_LOAD_CONVECTION)=False         'Load Convection (Index=68)
        .Draw(FVI_LOAD_RADIATION)=False          'Load Radiation (Index=69)
        .Draw(FVI_BOLT_PRELOAD)=False            'Load Bolt Preload (Index=93)
        .Draw(FVI_LOAD_FLUIDTRACKING)=False      'Load Fluid Tracking (Index=81)
        .Draw(FVI_LOAD_UNKNOWN)=False            'Load Unknown Condition (Index=82)
        .Draw(FVI_LOAD_SLIPWALL)=False           'Load Slip Wall Condition (Index=83)
        .Draw(FVI_LOAD_FANCURVE)=False           'Load Fan Curve (Index=84)
        .Draw(FVI_LOAD_PERIODIC)=False           'Load Periodic Condition (Index=85)
        .Draw(FVI_CONSTRAINT)=False              'Constraint (Index=20)
        .Draw(FVI_CONSTRAINT_EQUATION)=False     'Constraint Equation (Index=21)
        .Draw(FVI_CONNECTOR)=False               'Connector (Index=79)
        .Draw(FVI_CONNECTION_REGION)=False       'Connection Region (Index=91)
        .Draw(FVI_INTERIOR_POINT)=False          'Combined - Eliminated Points (Index=99)
        .Draw(FVI_INTERIOR_CURVE)=False          'Combined - Eliminated Curves (Index=97)
        .Draw(FVI_INTERIOR_SURFACE)=False        'Combined - Eliminated Surfaces (Index=98)
        .Draw(FVI_AERO_PANEL)=False              'Aero Panel (Index=102)
        .Draw(FVI_AERO_MESH)=False               'Aero Mesh (Index=103)
        .Draw(FVI_AERO_INTERFERENCE)=False       'Aero Interference (Index=108)
        .Draw(FVI_AERO_SPLINE)=False             'Aero Splines (Index=104)
        .Draw(FVI_AERO_SURFACE)=False            'Aero Control Surfaces (Index=105)

        'Tools & View Style
        .Draw(FVI_FREE_EDGE)=False               'Free Edge and Face (Index=22)
        .Draw(FVI_SHRINK_ELEM)=False             'Shrink Elements (Index=23)
        .Draw(FVI_FILL_ELEM)=True               'Fill, Backfaces and Hidden (Index=25)
        .Draw(FVI_FILLED_EDGES)=True            'Filled Edges (Index=26)
        '.Draw(FVI_RENDER_OPTIONS)=False          'Graphics Options (Index=77)
        '.Draw(FVI_PERF_GRAPHICS)=False           'Performance Graphics (Index=116)
        '.Draw(FVI_TRANSPARENCY)=False            'Transparency (Index=92)
        '.Draw(FVI_SHADING)=False                 'Shading (Index=28)
        '.Draw(FVI_PERSPECTIVE)=False             'Perspective (Index=29)
        '.Draw(FVI_STEREO)=False                  'Stereo (Index=30)
        '.Draw(FVI_AXISYM_AXIS)=False             'Axisymmetric Axis (Index=87)
        .Draw(FVI_VIEW_LEGEND)=True               'View Legend (Index=31)
        .Draw(FVI_VIEW_AXES)=True               'View Axes (Index=32)
        .Draw(FVI_ORIGIN)=True                  'Origin (Index=33)
        .Draw(FVI_WORKPLANE)=False                'Workplane and Rulers (Index=34)
        .Draw(FVI_WORKPLANE_GRID)=False           'Workplane Grid (Index=35)
        .Draw(FVI_CLIPPING_PLANE)=False           'Clipping Planes (Index=36)
        .Draw(FVI_SYMBOL)=True                  'Symbols (Index=37)
        '.Draw(FVI_PREVIEW)=False                 'Preview (Index=109)
        '.Draw(FVI_ASPECT_RATIO)=False            'View Aspect Ratio (Index=38)
        '.Draw(FVI_CLIP_PLANE)=False               'Model Clipping Plane (Index=101)

        'Post-processing
        '.Draw(FVI_POST_TITLES)=False             'Post Titles (Index=40)
        '.Draw(FVI_DEFORMED_STYLE)=False          'Deformed Style (Index=41)
        '.Draw(FVI_VECTOR_STYLE)=False            'Vector Style (Index=42)
        '.Draw(FVI_ANIMATED_STYLE)=False          'Animated Style (Index=43)
        '.Draw(FVI_DEFORMED)=False                'Deformed Model (Index=44)
        .Draw(FVI_UNDEFORMED)=False              'Undeformed Model (Index=45)
        '.Draw(FVI_TRACE_STYLE)=False             'Trace Style (Index=74)
        .Draw(FVI_CONTOUR_TYPE)=True             'Contour Type (Index=80)
        '.Draw(FVI_CONTOUR_STYLE)=False           'Contour/Criteria Style (Index=46)
        '.Draw(FVI_CONTOUR_LEVELS)=False          'Contour/Criteria Levels (Index=47)
        .Draw(FVI_CONTOUR_LEGEND)=True            'Contour/Criteria Legend (Index=48)
        '.Draw(FVI_CRITERIA_LIMITS)=False         'Criteria Limits (Index=49)
        '.Draw(FVI_CRITERIA_PASS)=False           'Criteria - Elements that Pass (Index=50)
        '.Draw(FVI_CRITERIA_FAIL)=False           'Criteria - Elements that Fail (Index=51)
        '.Draw(FVI_BEAM_DIAGRAM)=False            'Beam Diagram (Index=86)
        '.Draw(FVI_ISOSURFACE)=False              'IsoSurface (Index=76)
        '.Draw(FVI_ISOLINE)=False                 'IsoLine (Index=88)
        '.Draw(FVI_STREAMLINE)=False              'Streamline (Index=89)
        '.Draw(FVI_CONTOUR_VECTOR)=False          'Contour Arrow Style (Index=75)
        '.Draw(FVI_CONTOUR_ARROW_OPT)=False       'Contour Arrow Options (Index=117)
        '.Draw(FVI_FREEBODY)=False                'Freebody (Index=114)
        '.Draw(FVI_FREEBODY_NODEMARKER)=False     'Freebody Node Markers (Index=106)
        '.Draw(FVI_FREEBODY_VECTOR)=False         'Freebody Vectors (Index=113)
        '.Draw(FVI_FREEBODY_TOTALF)=False         'Freebody Total Force (Index=107)
        '.Draw(FVI_FREEBODY_TOTALM)=False         'Freebody Total Moment (Index=110)
        '.Draw(FVI_FREEBODY_NODAL_F)=False        'Freebody Nodal Force (Index=111)
        '.Draw(FVI_FREEBODY_NODAL_M)=False        'Freebody Nodal Force (Index=112)
        '.Draw(FVI_XY_TITLES)=False               'XY Titles (obsolete) (Index=52)
        '.Draw(FVI_XY_LEGEND)=False               'XY Legend (obsolete) (Index=53)
        '.Draw(FVI_XY_AXES_STYLE)=False           'XY Axes Style (obsolete) (Index=54)
        '.Draw(FVI_XY_XAXIS)=False                'XY X Range/Grid (obsolete) (Index=55)
        '.Draw(FVI_XY_YAXIS)=False                'XY Y Range/Grid (obsolete) (Index=56)
        '.Draw(FVI_XY_CURVE1)=False               'XY Curve 1 (obsolete) (Index=57)
        '.Draw(FVI_XY_CURVE2)=False               'XY Curve 2 (obsolete) (Index=58)
        '.Draw(FVI_XY_CURVE3)=False               'XY Curve 3 (obsolete) (Index=59)
        '.Draw(FVI_XY_CURVE4)=False               'XY Curve 4 (obsolete) (Index=60)
        '.Draw(FVI_XY_CURVE5)=False               'XY Curve 5 (obsolete) (Index=61)
        '.Draw(FVI_XY_CURVE6)=False               'XY Curve 6 (obsolete) (Index=62)
        '.Draw(FVI_XY_CURVE7)=False               'XY Curve 7 (obsolete) (Index=63)
        '.Draw(FVI_XY_CURVE8)=False               'XY Curve 8 (obsolete) (Index=64)
        '.Draw(FVI_XY_CURVE9)=False               'XY Curve 9 (obsolete) (Index=65)

        'Background
        .WindowShadeBack = False    'Use a shaded background
        .WindowBackcolor = 55       'Set first background color to neutral blue

        'Text Colors
        .color(FVI_LABEL) = 149     'label color to grey/white
        .color(FVI_TEXT) = 149      'text color to grey/white
        .Label(FVI_LABEL) = 38      'Change Text to 14 pt Arial
        .Color(FVI_VIEW_LEGEND) = 149 'Change view legend text color to grey/white

        'Properties
        .ColorMode(FVI_ELEMENT) = 3         'Use property colors for elements
        .ColorMode(FVI_FILLED_EDGES) = 3    'use contrasting color for filled  edges

        'Set Contour/Criteria Setting
        .ContourElemUnaveraged = True        'No Averaging

        .Label(FVI_POST_TITLES) = 7          'Move Post Titles to Bottom Left
        .color(FVI_POST_TITLES) = 149        'Post titles color to white
        .ColorMode(FVI_POST_TITLES) = 3      'Set legend style to titles and max/min data

        .Label(FVI_CONTOUR_LEGEND) = 2       'Post contour legend to the top right
        .ColorMode(FVI_CONTOUR_LEGEND) = 3   'Set label color to view color with exponent labels
        .Color(FVI_CONTOUR_LEGEND) = 149     'Set view color of contour legend labels to white
        .ContourLegendShrinkPct = 95         'Shrink the contour Bar to 95%

        .Label(FVI_CONTOUR_TYPE) = 1         'Elemental Contour Data

        .ColorMode(FVI_CONTOUR_STYLE) = 3    'Max Element No Corner Data
        .Label(FVI_CONTOUR_STYLE) = 2        'Max-Min Contour values is shown on FEM

        .Label(FVI_CONTOUR_LEVELS) = 0       'Automatic level mode
        .ColorMode(FVI_CONTOUR_LEVELS) = 0   'Use Standard pallete for contour levels
        .color(FVI_CONTOUR_LEVELS) = 10      'number of contour levels to 10

        'Contour Arrows
        .Label(FVI_CONTOUR_ARROW_OPT) = 6      'Solid Center Dual arrow style (Index=117)

        'Original FeMap Colors
        .ContourColor(FVI_LABEL) = 4
        .ContourColor(FVI_CSYS) = 16408
        .ContourColor(FVI_POINT) = 20
        .ContourColor(FVI_CURVE) = 16504
        .ContourColor(FVI_SURFACE) = 100
        .ContourColor(FVI_VOLUME) = 16488

        'Set Contour Colors
        '.ContourColor(FVI_LABEL) = 4
        '.ContourColor(FVI_CSYS) = 16398
        '.ContourColor(FVI_POINT) = 24
        '.ContourColor(FVI_CURVE) = 16404
        '.ContourColor(FVI_SURFACE) = 120
        '.ContourColor(FVI_VOLUME) = 16484

        'CSys
        .ColorMode(FVI_VIEW_AXES) = 5  'View Axes as RGB Solid
        .ColorMode(FVI_CSYS) = 7       'CSys is RGB Arrow
        .AxisLocation(FVI_LABEL) = 1   'Move Axis to Bottom Left Corner
        .AxisLocation(FVI_CSYS) = 99   'Move Axis to Bottom Left Corner

        'Element Stuff
        '.Label(FVI_ELEMENT_ORIENT) = 3      'Turn on Element Thickness
        .ColorMode(FVI_ELEMENT_DIRECTION) = 1  'Set Normal Style to Normal Vectors
        .Label(FVI_SYMBOL) = 1 'Change symbol size to 1
        .ColorMode(FVI_SYMBOL) = 0 'Nodes, Points and Mesh changed to "Node Dots"

        .Put(viewID)

    End With

    'Set element Color for all elements
    Dim eset As femap.Set  'Set Containing all elements
    Set eset = App.feSet()
    rc = eset.AddAll(FT_ELEM)
    App.feModifyColor(FT_ELEM,eset.ID, 75, False)

    '
    '*********** Set RBE2 and RBE3 to different colors ********************
    '
    Dim rigidSet As femap.Set
    Set rigidSet = App.feSet
    Dim eid1 As Long
        rc = rigidSet.AddRule( femap.FET_L_RIGID, FGD_ELEM_BYTYPE )
        eid1 = rigidSet.First

    Dim elem1 As femap.Elem
    Set elem1 = App.feElem
    Dim elem2 As femap.Elem
    Set elem2 = App.feElem


    While eid1 > 0
        If eid1 = 0 Then Exit While

        rc = elem1.Get( eid1)
        If elem1.RigidInterpolate = True Then
            '*** RBE3 Color ***
            elem1.color =8
        Else
            '*** RBE2 Color ***
            elem1.color =75
        End If
        elem1.Put (eid1)
            eid1= rigidSet.Next
    Wend

    Dim Color As Long
    Color = 80
    Dim Index As Long

    'Randomize the Propety Colors
    Dim pset As femap.Set  'Set containing all properties
    Dim propset As Integer
    Set pset = App.feSet()
    propset = pset.ID
    rc = pset.AddAll(FT_PROP)
    If pset.count<=0 Then Exit Sub

    'Regenerate and redraw view
    View.Regenerate
    View.Redraw

End Sub
