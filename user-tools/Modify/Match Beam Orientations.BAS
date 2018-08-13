Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Match Beam Orientations")

    Dim rc As Long

    Dim beamSet As femap.Set
    Set beamSet = App.feSet

    Dim bSource As femap.Elem
    Set bSource = App.feElem

    Dim beam As femap.Elem
    Set beam = App.feElem

    Dim counter As Long

    rc = beamSet.Select(FT_ELEM, True, "Select beam elements...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "One or more selected elements do not exist")
        Exit Sub
    End If

    ' prompt user to pick element
    PICK_SOURCE:
    bSource.SelectID("Select beam element to match orientation...")
    bSource.Get(bSource.ID)

    ' make sure element is a bar or beam type element
    Select Case bSource.type
        Case 2, 5, 8, 37
            ' do nothing if beam element
        Case Else
            rc = App.feAppMessageBox(1, "Selected element is not a beam element")
            If rc = FE_OK Then
                GoTo PICK_SOURCE
            Else
                Exit Sub
            End If
    End Select

    ' orient each beam to the same as the source beam
    counter = 0
    Do While beam.NextInSet(beamSet.ID)
        beam.Get(beam.ID)

        ' make sure element is a beam type
        Select Case beam.type
            Case 2, 5, 8, 37
                beam.vorient = bSource.vorient
                If beam.Put(beam.ID) <> FE_FAIL Then
                    counter = counter + 1
                Else
                    App.feAppMessage(FCM_ERROR, "Unable to reorient Element " & CStr(beam.ID) & ".")
                End If
            Case Else
                App.feAppMessage(FCM_ERROR, "Element " & CStr(beam.ID) & " is not a beam element. Element skipped...")
        End Select
    Loop

    App.feViewRegenerate(0)
    App.feAppMessage(FCM_NORMAL, CStr(counter) & " elements updated.")

End Sub
