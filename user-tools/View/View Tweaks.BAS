Sub Main
    Dim App As femap.model
    'Set App = GetObject(,"femap.model")
        Set App = feFemap()


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

            'Background
            '.WindowBackColor = 124   ' Set the background to white
            .WindowShadeBack = False ' Use a shaded background
            '.WindowShadeBack = True  ' Use a shaded background
            .WindowBackcolor = 55   ' Set first background color to neutral blue
            '.WindowBackColor2 = 55  ' Set second background color to dark grey

            'Text Colors
            .color(0) = 250   ' label color to white
            .color(6) = 250   ' text color to white
            .Label(0) = 2  'Change Text to 12 pt MS Sans Serif

            'Properties
            .ColorMode(9) = 3 ' Use property colors for elements
            .ColorMode(26) = 1 ' use view color for filled  edges
            '.Label(26) = 2  'Filled Edges Use View Color
            '.color(26) = 0 ' Set Filled Edges View color

            .Draw(45) = False ' turn off display of undeformed model

            'Set Contour/Criteria Setting
            .ContourElemUnaveraged = True  'No Averaging

            '.color(40) = 0      'Post titles color
            .Label(40) = 0      'Move Post Titles to Top Left
            '.Label(40) = 2      'Move Post Titles to Top Right
            .Label(80) = 1      'Elemental Contour Data
            .Draw(80) = True    'Set to Double Sided
            .ColorMode(46) = 3  'Max Element No Corner Data
            .Label(46) = 2      'Max-Min Contour values is shown on FEM
            .Label(47) = 1      'Contour Level is Auto-Group
            .Label(48) = 4      'Move Contour Bar to Center Right
            '.Label(48) = 7      'Move Contour Bar to Bottom Right
            '.ColorMode(48) = 1  'Use view color for contour legend
            '.color(48) = 0      'Black view color set for Contour
            '.ContourLegendShrinkPct = 95 'Shrink the contour Bar to 95%


            'Original FeMap Colors
            .ContourColor(0) = 4
            .ContourColor(1) = 16408
            .ContourColor(2) = 20
            .ContourColor(3) = 16504
            .ContourColor(4) = 100
            .ContourColor(5) = 16488

            'Set Contour Colors
            '.ContourColor(0) = 4
            '.ContourColor(1) = 16398
            '.ContourColor(2) = 24
            '.ContourColor(3) = 16404
            '.ContourColor(4) = 120
            '.ContourColor(5) = 16484

            'CSys
            '.ColorMode(32) = 5 'Axis is RGB Color
            '.ColorMode(1) = 7 'CSys is RGB Arrow Color
            '.ColorMode(1) = 11 'CSys is RGB Solid Color
            '.AxisLocation(0) = 1   'Move Axis to Bottom Left Corner
            '.AxisLocation(1) = 99  'Move Axis to Bottom Left Corner

            'Element Stuff
            .Draw(11) = True   'Element-Offsets/Releases
            '.Label(12) = 3      'Turn on Element Thickness
            .ColorMode(10) = 1  'Set Normal Style to Normal Vectors

            'Turn Off All Extra Stuff
            .Draw(1) = True   'Coordinate Systems
            .Draw(9) = True   'Element



            .Draw(31) = True ' turn on view legend
            .Draw(33) = True ' turn on origin
            .Draw(34) = False ' turn off workplane and rulers
            .Draw(35) = False ' turn off workplane grid
            .Draw(1) = False  ' turn off Coordinate Systems
            '.Draw(0) = False  'Label Parameters
            '.Draw(2) = False  'Point
            '.Draw(3) = False  'Curve
            '.Draw(4) = False  'Surface
            '.Draw(5) = False  'Volume
            '.Draw(6) = False  'Text
            '.Draw(27) = False 'Boundry
            '.Draw(95) = False 'Combined Curve
            '.Draw(7) = False   'Node
            '.Draw(8) = False  'Node - PPerm Constraint
            '.Draw(9) = False   'Element
            '.Draw(10) = False  'Element-Directions
            '.Draw(11) = False   'Element-Offsets/Releases
            '.Draw(12) = False  'Element Orientation/Shape
            '.Draw(13) = False  'Element-Beam Y Axis
            '.Draw(90) = False  'Element-Weld
            '.Draw(78) = False  'Load Vectors
            '.Draw(14) = False  'Load Force
            '.Draw(15) = False  'Load Moment
            '.Draw(16) = False  'Load Thermal
            '.Draw(71) = False  'Load Distributed
            '.Draw(17) = False  'Load Pressure
            '.Draw(18) = False  'Load Acceleration
            '.Draw(72) = False  'Load Velocity
            '.Draw(19) = False  'Load Engorced Displacement
            '.Draw(73) = False  'Load NonLinear Force
            '.Draw(66) = False  'Load Heat Generation
            '.Draw(67) = False  'Load Heat Flux
            '.Draw(68) = False  'Load Convection
            '.Draw(69) = False  'Load Radiation
            '.Draw(81) = False  'Load Fluid Tracking
            '.Draw(82) = False  'Load Unknown Condition
            '.Draw(83) = False  'Load Slip Wall Condition
            '.Draw(84) = False  'Load Fan Curve
            '.Draw(85) = False  'Load Periodic Condition
            '.Draw(93) = False  'Load Bolt Preload
            '.Draw(100) = False 'Load Body
            '.Draw(20) = False  'Constraint
            '.Draw(21) = False  'Constraint Equation
            '.Draw(79) = False  'Connector
            '.Draw(91) = False  'Connection Region

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

        rc = App.feWindowRegenerate (1)


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

ReTry_Colors:
        App.feModifyColor(FT_PROP, propset, Color, True)

        'Make Sure no properties use Dark Colors
        '0,1,2,125-130
        Dim Prope As femap.Prop
        Set Prope = App.feProp
        Dim Prop_ID As Long
        Dim PColor As Integer
        Dim I As Long
        Dim Prop_IDs As Variant
        rc = pset.GetArray(I,Prop_IDs)
        rc = pset.Clear

        For I = 0  To UBound(Prop_IDs,1)
            Prop_ID = Prop_IDs(I)
            Prope.Get(Prop_ID)
            PColor = Prope.color
            If PColor >=0 And PColor <= 2 Then GoTo Change_Color
            If PColor >=5 And PColor <= 7 Then GoTo Change_Color
            If PColor >=25 And PColor <= 27 Then GoTo Change_Color
            If PColor >=30 And PColor <= 32 Then GoTo Change_Color
            If PColor >=50 And PColor <= 52 Then GoTo Change_Color
            If PColor >=55 And PColor <= 56 Then GoTo Change_Color
            If PColor >=75 And PColor <= 76 Then GoTo Change_Color
            If PColor >=80 And PColor <= 81 Then GoTo Change_Color
            If PColor >=124 And PColor <= 149 Then GoTo Change_Color
Next_i:
        Next


        If Index > 0 Then
            Color = Color +10
            Index = 0
            GoTo Retry_Colors
       End If


        Exit Sub

Change_Color:
                    Index = Index +1
                   rc = pset.Add(Prop_ID)
                   GoTo Next_i
End Sub
