Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim rc As Long

    Dim crvSet As femap.Set
    Set crvSet = App.feSet

    Dim crv As femap.Curve
    Set crv = App.feCurve

    Dim bmProp As femap.Prop
    Set bmProp = App.feProp

    Dim plotOnly As Boolean

    Dim vecLength As Double
    Dim vecBase As Variant
    Dim vecDir As Variant

    ' Prompt user to select curves
    rc = crvSet.Select(FT_CURVE, True, "Select straight curves to convert")
    Select Case rc
        Case FE_CANCEL
            Exit Sub
        Case FE_NOT_EXIST
            App.feAppMessage(FCM_ERROR, "No entities of selected type exist.")
        Case Else
            ' Check for straight curves
            Do While crv.NextInSet(crvSet.ID)
                If crv.IsStraight() = FE_FAIL Then  ' remove non-straight curves
                    App.feAppMessage(FCM_ERROR, "Curve " & crv.ID & " is not a straight line. Removed from selection set.")
                    crvSet.Remove(crv.ID)
                ElseIf crv.IsCombinedCurve() = FE_OK Then  ' remove combined curves
                    App.feAppMessage(FCM_ERROR, "Curve " & crv.ID & " is a combined curve. Removed from selection set.")
                    crvSet.Remove(crv.ID)
                End If
            Loop
    End Select

    ' Prompt user to select 1D elm property
    If bmProp.SelectID("Select Property") Then
        ' Make sure 1D property was selected
        Select Case bmProp.type
            Case 1 To 9  ' Do nothing if 1D property is selected
                plotOnly = False
            Case Else  ' Prompt for plot-only elements
                If App.feAppMessageBox(1, "Selected property is not a 1D property. Do you wish to create plot-only elements?") Then
                    plotOnly = True
                Else
                    Exit Sub
                End If
        End Select
    Else
        Exit Sub
    End If

    ' Prompt user to specify y-vector for curves
    App.feVectorPick("Select orientation Vector", True, vecLength, vecBase, vecDir)

    ' Size curves
    App.feMeshSizeCurve(crvSet.ID, 1, 1, 1, 1, 1, 0, 0, 1.0, 2, False)

    ' Mesh curves
    If plotOnly Then
        App.feMeshCurve(crvSet.ID, True, -1, vecDir)
    Else
        App.feMeshCurve(crvSet.ID, True, bmProp.ID, vecDir)
    End If

End Sub
