Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Setup View")

    Dim viewID As Long
    Dim rc As femap.zReturnCode

    With App
        .Info_GroupAutomaticAdd = -1     ' Automatic add to active group
        .feAppGetActiveView(viewID)
    End With

    Dim slctr As femap.Selector
    Set slctr = App.feSelector

    Dim View As femap.View
    Set View = App.feView

    Dim presentationView As femap.View
    Set presentationView = App.feView

    ' Setup the selector
    With slctr
        .MultipleMode = True
        .AddMode = True
        .DrillMode = 0
        .RelatedMode = False
        .SelectEntity = FT_ELEM
        .Active = False
    End With

    'Set RBE2 and RBE3 to different colors
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
            eid1 = rigidSet.Next
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

    With View
        .Get(viewID)
        .Title = "Working"

        .WindowShadeBack = True                   'Use a shaded background
        .WindowShadeBackOpt = 6                   'Use horizon option
        .WindowBackColor2 = 136
        .WindowBackcolor = 136
        .WindowBackColor3 = 132
        .WindowShadeHorizon = 40

        .Draw(FVI_AERO_INTERFERENCE)=False        'Aero Interference (Index=108)
        .Draw(FVI_AERO_MESH)=False                'Aero Mesh (Index=103)
        .Draw(FVI_AERO_PANEL)=False               'Aero Panel (Index=102)
        .Draw(FVI_AERO_SPLINE)=False              'Aero Splines (Index=104)
        .Draw(FVI_AERO_SURFACE)=False             'Aero Control Surfaces (Index=105)
        .Draw(FVI_BEAM_YAXIS)=False               'Element - Beam Y-Axis (Index=13)
        .Draw(FVI_BOLT_PRELOAD)=False             'Load Bolt Preload (Index=93)
        .Draw(FVI_BOUNDARY)=False                 'Boundary (Index=27)
        .Draw(FVI_CLIPPING_PLANE)=False           'Clipping Planes (Index=36)
        .Draw(FVI_COMPOSITE_CURVE)=False          'Combined Curve (Index=95)
        .Draw(FVI_CONNECTION_REGION)=False        'Connection Region (Index=91)
        .Draw(FVI_CONNECTOR)=False                'Connector (Index=79)
        .Draw(FVI_CONSTRAINT_EQUATION)=False      'Constraint Equation (Index=21)
        .Draw(FVI_CONSTRAINT)=False               'Constraint (Index=20)
        .Draw(FVI_CONTOUR_LEGEND)=True            'Contour/Criteria Legend (Index=48)
        .Color(FVI_CONTOUR_LEGEND) = 149          'Set view color of contour legend labels to white
        .ColorMode(FVI_CONTOUR_LEGEND) = 3        'Set label color to view color with exponent labels
        .Label(FVI_CONTOUR_LEGEND) = 2            'Post contour legend to the top right
        .Draw(FVI_CONTOUR_TYPE)=True              'Contour Type (Index=80)
        .Label(FVI_CONTOUR_TYPE) = 1              'Elemental Contour Data
        .Draw(FVI_CSYS)=False                     'Coordinate System (Index=1)
        .AxisLocation(FVI_CSYS) = 99              'Move Axis to Bottom Left Corner
        .ColorMode(FVI_CSYS) = 7                  'CSys is RGB Arrow
        .ContourColor(FVI_CSYS) = 16408
        .Draw(FVI_CURVE_ACCURACY)=False           'Curve/Surface Directions (Index=39)
        .Draw(FVI_CURVE)=False                    'Curve (Index=3)
        .ContourColor(FVI_CURVE) = 16504
        .Draw(FVI_ELEM_CSYS)=False                'Element - Coordinate System (Index=96)
        .Draw(FVI_ELEM_MATL_DIR)=False            'Element - Material Direction (Index=115)
        .Draw(FVI_ELEMENT_DIRECTION)=False        'Element - Directions (Index=10)
        .ColorMode(FVI_ELEMENT_DIRECTION) = 1     'Set Normal Style to Normal Vectors
        .Draw(FVI_ELEMENT_OFFSET)=False           'Element - Offsets/Releases (Index=11)
        .Draw(FVI_ELEMENT_ORIENT)=False           'Element - Orientation/Shape (Index=12)
        .Draw(FVI_ELEMENT_WELD)=False             'Element - Weld (Index=90)
        .Draw(FVI_ELEMENT)=True                   'Element (Index=9)
        .ColorMode(FVI_ELEMENT) = 3               'Use property colors for elements
        .Draw(FVI_FILL_ELEM)=True                 'Fill, Backfaces and Hidden (Index=25)
        .Draw(FVI_FILLED_EDGES)=True              'Filled Edges (Index=26)
        .ColorMode(FVI_FILLED_EDGES) = 3          'use contrasting color for filled  edges
        .Draw(FVI_FREE_EDGE)=False                'Free Edge and Face (Index=22)
        .Draw(FVI_INTERIOR_CURVE)=False           'Combined - Eliminated Curves (Index=97)
        .Draw(FVI_INTERIOR_POINT)=False           'Combined - Eliminated Points (Index=99)
        .Draw(FVI_INTERIOR_SURFACE)=False         'Combined - Eliminated Surfaces (Index=98)
        .Draw(FVI_LABEL)=True                     'Label Parameters (Index=0)
        .AxisLocation(FVI_LABEL) = 1              'Move Axis to Bottom Left Corner
        .Color(FVI_LABEL) = 149                   'label color to grey/white
        .ContourColor(FVI_LABEL) = 4
        .Label(FVI_LABEL) = 38                    'Change Text to 14 pt Arial
        .Draw(FVI_LOAD_ACCELERATION)=False        'Load Acceleration (Index=18)
        .Draw(FVI_LOAD_BODY)=False                'Load Body (Index=100)
        .Draw(FVI_LOAD_CONVECTION)=False          'Load Convection (Index=68)
        .Draw(FVI_LOAD_DISTRIBUTEDLOAD)=False     'Load Distributed (Index=71)
        .Draw(FVI_LOAD_ENFORCEDDISP)=False        'Load Enforced Displacement (Index=19)
        .Draw(FVI_LOAD_FANCURVE)=False            'Load Fan Curve (Index=84)
        .Draw(FVI_LOAD_FLUIDTRACKING)=False       'Load Fluid Tracking (Index=81)
        .Draw(FVI_LOAD_FORCE)=False               'Load Force (Index=14)
        .Draw(FVI_LOAD_HEATFLUX)=False            'Load Heat Flux (Index=67)
        .Draw(FVI_LOAD_HEATGEN)=False             'Load Heat Generation (Index=66)
        .Draw(FVI_LOAD_MOMENT)=False              'Load Moment and Torque (Index=15)
        .Draw(FVI_LOAD_NONLINEARFORCE)=False      'Load Nonlinear Force (Index=73)
        .Draw(FVI_LOAD_PERIODIC)=False            'Load Periodic Condition (Index=85)
        .Draw(FVI_LOAD_PRESSURE)=False            'Load Pressure (Index=17)
        .Draw(FVI_LOAD_RADIATION)=False           'Load Radiation (Index=69)
        .Draw(FVI_LOAD_SLIPWALL)=False            'Load Slip Wall Condition (Index=83)
        .Draw(FVI_LOAD_THERMAL)=False             'Load Temperature (Index=16)
        .Draw(FVI_LOAD_UNKNOWN)=False             'Load Unknown Condition (Index=82)
        .Draw(FVI_LOAD_VECTORS)=False             'Load Vectors (Index=78)
        .Draw(FVI_LOAD_VELOCITY)=False            'Load Velocity (Index=72)
        .Draw(FVI_NODE_PERMCONSTRAINT)=False      'Node - Perm Constraint (Index=8)
        .Draw(FVI_NODE)=False                     'Node (Index=7)
        .Draw(FVI_ORIGIN)=True                    'Origin (Index=33)
        .Draw(FVI_POINT)=False                    'Point (Index=2)
        .ContourColor(FVI_POINT) = 20
        .Draw(FVI_RIGID)=True                     'Element - Rigid (Index=94)
        .Draw(FVI_SHRINK_ELEM)=False              'Shrink Elements (Index=23)
        .Draw(FVI_SURFACE)=False                  'Surface (Index=4)
        .ColorMode(FVI_SURFACE)=3                 'Color Surfaces by Property
        .ContourColor(FVI_SURFACE) = 100
        .Draw(FVI_SYMBOL)=True                    'Symbols (Index=37)
        .ColorMode(FVI_SYMBOL) = 0                'Nodes, Points and Mesh changed to "Node Dots"
        .Label(FVI_SYMBOL) = 1                    'Change symbol size to 1
        .Draw(FVI_TEXT) =False                    'Text (Index=6)
        .Color(FVI_TEXT) = 149                    'text color to grey/white
        .Draw(FVI_UNDEFORMED)=False               'Undeformed Model (Index=45)
        .Draw(FVI_VIEW_AXES)=True                 'View Axes (Index=32)
        .ColorMode(FVI_VIEW_AXES) = 5             'View Axes as RGB Solid
        .Draw(FVI_VIEW_LEGEND)=True               'View Legend (Index=31)
        .Color(FVI_VIEW_LEGEND) = 149             'Change view legend text color to grey/white
        .ColorMode(FVI_VIEW_LEGEND) = 0           'Change view legend text to IDs only
        .Draw(FVI_VOLUME) =False                  'Volume (Index=5)
        .ContourColor(FVI_VOLUME) = 16488
        .Draw(FVI_WORKPLANE_GRID)=False           'Workplane Grid (Index=35)
        .Draw(FVI_WORKPLANE)=False                'Workplane and Rulers (Index=34)
        .Color(FVI_CONTOUR_LEVELS) = 10           'number of contour levels to 10
        .ColorMode(FVI_CONTOUR_LEVELS) = 0        'Use Standard pallete for contour levels
        .Label(FVI_CONTOUR_LEVELS) = 0            'Automatic level mode
        .Color(FVI_POST_TITLES) = 149             'Post titles color to white
        .ColorMode(FVI_POST_TITLES) = 3           'Set legend style to titles and max/min data
        .Label(FVI_POST_TITLES) = 7               'Move Post Titles to Bottom Left
        .ColorMode(FVI_CONTOUR_STYLE) = 3         'Max Element No Corner Data
        .Label(FVI_CONTOUR_STYLE) = 2             'Max-Min Contour values is shown on FEM
        .ContourElemUnaveraged = True             'No Averaging
        .ContourLegendShrinkPct = 95              'Shrink the contour Bar to 95%
        .Label(FVI_CONTOUR_ARROW_OPT) = 6         'Solid Center Dual arrow style (Index=117)
        .Draw(FVI_CURVE_MESHSIZE) = True          'Show mesh sizing
        .Label(FVI_CURVE_MESHSIZE) = 3            'Curve mesh sizing shown as symbols and count

        .Put(viewID)
    End With

    'Regenerate and redraw view
    View.Regenerate
    View.Redraw

    ' Create Presentation View
    With View
        .Get(viewID)
        .Title = "Presentation"

        .WindowBackColor = 124              'Set the background to white
        .WindowShadeBack = False            'Use a shaded background

        .Label(FVI_LABEL) = 39              'Change Text to 18 pt Arial
        .ContourLegendShrinkPct = 90        'Shrink the contour Bar to 90%
        .ColorMode(FVI_CONTOUR_LEVELS) = 0  'Use Standard pallete for contour levels
        .ColorMode(FVI_CONTOUR_STYLE) = 3   'Max Element No Corner Data
        .Draw(FVI_CURVE_MESHSIZE) = False   'Curve - Mesh Size (Index=24)
        .Draw(FVI_POINT) = False            'Turn off points
        .Draw(FVI_CURVE) = False            'Turn off curves
        .Draw(FVI_SURFACE) = False          'Turn off surfaces

        .Color(FVI_TEXT) = 128              'text color to dark grey
        .Color(FVI_VIEW_LEGEND) = 128       'Change view legend text color to dark grey
        .ColorMode(FVI_FILLED_EDGES) = 1    'use view color for filled  edges
        .Color(FVI_FILLED_EDGES) = 128
        .Draw(FVI_ORIGIN) = False           'Origin (Index=33)
        .Color(FVI_LABEL) = 128             'label colors to dark grey

        .Put(View.NextEmptyID)
    End With

End Sub
