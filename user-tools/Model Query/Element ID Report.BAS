Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Element ID Report")

    Dim rc As Long

    Dim p As femap.Prop
    Set p = App.feProp

    Dim propSet As femap.Set
    Set propSet = App.feSet

    Dim eSet As femap.Set
    Set eSet = App.feSet

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim vPropID As Variant
    Dim vElemCount As Variant
    Dim columnLength As Long
    Dim i As Long

    Dim PropID As String
    Dim PropType As String
    Dim Order As String
    Dim FirstEID As String
    Dim LastEID  As String
    Dim ElmCount As String
    Dim FirstNID As String
    Dim LastNID  As String
    Dim NdCount As String
    Dim rigids As Long
    Dim typeStr As Variant

    ' prompt to select properties for report
    rc = App.feAppMessageBox(3, "Report full model?")
    If rc = FE_FAIL Then
        PICK_PROPERTIES:
        If propSet.Select(FT_PROP, True, "Select properties to report...") = FE_NOT_EXIST Then
            If App.feAppMessageBox(1, "One or more selected properties do not exist.") = FE_OK Then
                GoTo PICK_PROPERTIES
            Else
                Exit Sub
            End If
        End If
    ElseIf rc = FE_OK Then
        propSet.AddAll(FT_PROP)
        If propSet.Count() = 0 Then
            App.feAppMessage(FCM_ERROR, "No properties exist in model.")
            Exit Sub
        End If
    Else
        App.feAppMessage(FCM_COMMAND, "Command canceled...")
        Exit Sub
    End If

    rigids = App.feAppMessageBox(2, "Include rigid/interpolation elements?")

    ' find element counts on properties
    rc = App.feElemCountOnProp(propSet.ID, vPropID, vElemCount)

    ' print header (column width = 8)
    App.feAppMessage(FCM_COMMAND, "  PropID," & _
                                  "                           Type," & _
                                  "     Order," & _
                                  "   MinEID," & _
                                  "   MaxEID," & _
                                  " ElmCount," & _
                                  "   MinNID," & _
                                  "   MaxNID," & _
                                  "  NdCount")

    ' include rigid elements
    If rigids = FE_OK Then
        Dim reSet As femap.Set
        Set reSet = App.feSet

        Dim rnSet As femap.Set
        Set rnSet = App.feSet

        Dim rbe2Set As femap.Set
        Set rbe2Set = App.feSet

        Dim rbe3Set As femap.Set
        Set rbe3Set = App.feSet

        Dim rbe2NodeSet As femap.Set
        Set rbe2NodeSet = App.feSet

        Dim rbe3NodeSet As femap.Set
        Set rbe3NodeSet = App.feSet

        Dim rbe As femap.Elem
        Set rbe = App.feElem

        ' find all rigid elements
        reSet.Clear()
        reSet.AddRule(29, FGD_ELEM_BYTYPE)

        ' find all nodes attached to elements
        rnSet.Clear()
        rnset.AddSetRule(reSet.ID, FGD_NODE_ONELEM)

        ' separate out interpolation elements
        Do While rbe.NextInSet(reSet.ID)
            rbe.Get(rbe.ID)
            If rbe.RigidInterpolate Then
                rbe3Set.Add(rbe.ID)
                rbe3NodeSet.AddSetRule(rbe3Set.ID, FGD_NODE_ONELEM)
            Else
                rbe2Set.Add(rbe.ID)
                rbe2NodeSet.AddSetRule(rbe2Set.ID, FGD_NODE_ONELEM)
            End If
        Loop

        If rbe2Set.Count() > 0 Then

            ' create set length strings
            PropID = Right(Space(8) & "0", 8)
            PropType = Right(Space(30) & "Rigid", 30)
            Order = Right(Space(30) & "", 9)
            FirstEID = Right(Space(8) & CStr(rbe2Set.First()), 8)
            LastEID = Right(Space(8) & CStr(rbe2Set.Last()), 8)
            ElmCount = Right(Space(8) & CStr(rbe2Set.Count()), 8)
            FirstNID = Right(Space(8) & CStr(rbe2NodeSet.First()), 8)
            LastNID = Right(Space(8) & CStr(rbe2NodeSet.Last()), 8)
            NdCount = Right(Space(8) & CStr(rbe2NodeSet.Count()), 8)

            App.feAppMessage(FCM_NORMAL, PropID & ", " & _
                                         PropType & ", " & _
                                         Order & ", " & _
                                         FirstEID & ", " & _
                                         LastEID & ", " & _
                                         ElmCount & ", " & _
                                         FirstNID & ", " & _
                                         LastNID & ", " & _
                                         NdCount)

        ElseIf rbe3Set.Count() > 0 Then
            ' create set length strings
            PropID = Right(Space(8) & "0", 8)
            PropType = Right(Space(30) & "Interpolation", 30)
            Order = Right(Space(30) & "", 9)
            FirstEID = Right(Space(8) & CStr(rbe3Set.First()), 8)
            LastEID = Right(Space(8) & CStr(rbe3Set.Last()), 8)
            ElmCount = Right(Space(8) & CStr(rbe3Set.Count()), 8)
            FirstNID = Right(Space(8) & CStr(rbe3NodeSet.First()), 8)
            LastNID = Right(Space(8) & CStr(rbe3NodeSet.Last()), 8)
            NdCount = Right(Space(8) & CStr(rbe3NodeSet.Count()), 8)

            App.feAppMessage(FCM_NORMAL, PropID & ", " & _
                                         PropType & ", " & _
                                         Order & ", " & _
                                         FirstEID & ", " & _
                                         LastEID & ", " & _
                                         ElmCount & ", " & _
                                         FirstNID & ", " & _
                                         LastNID & ", " & _
                                         NdCount)
        End If
    End If

    ' print line for each property
    For i = 0 To propSet.Count() - 1
        ' find all elements for given property
        eSet.Clear()
        eSet.AddRule(vPropID(i), FGD_ELEM_BYPROP)

        ' find all nodes attached to elements
        nSet.Clear()
        nset.AddSetRule(eSet.ID, FGD_NODE_ONELEM)

        ' get property type and order
        p.Get(vPropID(i))
        typeStr = PropTypeAsString(p.type)

        ' create set length strings
        PropID = Right(Space(8) & CStr(vPropID(i)), 8)
        PropType = Right(Space(30) & typeStr(0), 30)
        Order = Right(Space(30) & typeStr(1), 9)
        FirstEID = Right(Space(8) & CStr(eSet.First()), 8)
        LastEID = Right(Space(8) & CStr(eSet.Last()), 8)
        ElmCount = Right(Space(8) & CStr(eSet.Count()), 8)
        FirstNID = Right(Space(8) & CStr(nSet.First()), 8)
        LastNID = Right(Space(8) & CStr(nSet.Last()), 8)
        NdCount = Right(Space(8) & CStr(nSet.Count()), 8)

        App.feAppMessage(FCM_NORMAL, PropID & ", " & _
                                     PropType & ", " & _
                                     Order & ", " & _
                                     FirstEID & ", " & _
                                     LastEID & ", " & _
                                     ElmCount & ", " & _
                                     FirstNID & ", " & _
                                     LastNID & ", " & _
                                     NdCount)
    Next i

    ' add all elements and nodes to sets for totals
    eSet.Clear()
    eSet.AddSetRule(propSet.ID, FGD_ELEM_BYPROP)
    nSet.Clear()
    nSet.AddSetRule(eSet.ID, FGD_NODE_ONELEM)

    ' add rigids to sets for totals
    If rigids = FE_OK Then
        propSet.Add(0)
        eSet.AddSet(reSet.ID)
        nSet.AddSet(rnSet.ID)
    End If

    App.feAppMessage(FCM_NORMAL, "")

    ' print totals header
    App.feAppMessage(FCM_COMMAND, "Totals")
    App.feAppMessage(FCM_NORMAL, "No. of properties = " & CStr(propSet.Count()))
    App.feAppMessage(FCM_NORMAL, "No. of elements   = " & CStr(eSet.Count()))
    App.feAppMessage(FCM_NORMAL, "No. of nodes      = " & CStr(nSet.Count()))

End Sub


Public Function PropTypeAsString(typeVal As Long)

    Dim typeArr As Variant
    ReDim typeArr(2)

    Select Case typeVal
        Case 1
            typeArr(0) = "Rod"
            typeArr(1) = ""
        Case 2
            typeArr(0) = "Bar"
            typeArr(1) = ""
        Case 3
            typeArr(0) = "Tube"
            typeArr(1) = ""
        Case 4
            typeArr(0) = "Link"
            typeArr(1) = ""
        Case 5
            typeArr(0) = "Beam"
            typeArr(1) = ""
        Case 6
            typeArr(0) = "Spring/Damper"
            typeArr(1) = ""
        Case 7
            typeArr(0) = "DOF Spring"
            typeArr(1) = ""
        Case 8
            typeArr(0) = "Curved Beam"
            typeArr(1) = ""
        Case 9
            typeArr(0) = "Gap"
            typeArr(1) = ""
        Case 10
            typeArr(0) = "Plot Only"
            typeArr(1) = ""
        Case 11
            typeArr(0) = "Shear Panel"
            typeArr(1) = "Linear"
        Case 12
            typeArr(0) = "Shear Panel"
            typeArr(1) = "Parabolic"
        Case 13
            typeArr(0) = "Membrane"
            typeArr(1) = "Linear"
        Case 14
            typeArr(0) = "Membrane"
            typeArr(1) = "Parabolic"
        Case 15
            typeArr(0) = "Bending Only"
            typeArr(1) = "Linear"
        Case 16
            typeArr(0) = "Bending Only"
            typeArr(1) = "Parabolic"
        Case 17
            typeArr(0) = "Plate"
            typeArr(1) = "Linear"
        Case 18
            typeArr(0) = "Plate"
            typeArr(1) = "Parabolic"
        Case 19
            typeArr(0) = "Plane Strain"
            typeArr(1) = "Linear"
        Case 20
            typeArr(0) = "Plane Strain"
            typeArr(1) = "Parabolic"
        Case 21
            typeArr(0) = "Laminate"
            typeArr(1) = "Linear"
        Case 22
            typeArr(0) = "Laminate"
            typeArr(1) = "Parabolic"
        Case 23
            typeArr(0) = "Axisymmetric"
            typeArr(1) = "Linear"
        Case 24
            typeArr(0) = "Axisymmetric"
            typeArr(1) = "Parabolic"
        Case 25
            typeArr(0) = "Solid"
            typeArr(1) = "Linear"
        Case 26
            typeArr(0) = "Solid"
            typeArr(1) = "Parabolic"
        Case 27
            typeArr(0) = "Mass"
            typeArr(1) = ""
        Case 28
            typeArr(0) = "Mass Matrix"
            typeArr(1) = ""
        Case 29
            typeArr(0) = "Rigid"
            typeArr(1) = ""
        Case 30
            typeArr(0) = "Stiffness Matrix"
            typeArr(1) = ""
        Case 31
            typeArr(0) = "Curved Tube"
            typeArr(1) = ""
        Case 32
            typeArr(0) = "Plot Only Plate"
            typeArr(1) = ""
        Case 33
            typeArr(0) = "Side Line"
            typeArr(1) = ""
        Case 34
            typeArr(0) = "Contact"
            typeArr(1) = ""
        Case 35
            typeArr(0) = "Axisymmetric Shell"
            typeArr(1) = "Linear"
        Case 36
            typeArr(0) = "Axisymmetric Shell"
            typeArr(1) = "Parabolic"
        Case 37
            typeArr(0) = "Para Beam"
            typeArr(1) = ""
        Case 38
            typeArr(0) = "Weld/Fastener"
            typeArr(1) = ""
        Case 39
            typeArr(0) = "Solid Laminate"
            typeArr(1) = "Linear"
        Case 40
            typeArr(0) = "Solid Laminate"
            typeArr(1) = "Parabolic"
        Case 41
            typeArr(0) = "Spring/Damper to Ground"
            typeArr(1) = ""
        Case 42
            typeArr(0) = "DOF Spring/Damper to Ground"
            typeArr(1) = ""
    End Select

    PropTypeAsString = typeArr

End Function
