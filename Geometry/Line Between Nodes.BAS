Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Line Between Nodes")

    Dim rc As Long

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim n1 As femap.Node
    Set n1 = App.feNode

    Dim n2 As femap.Node
    Set n2 = App.feNode

    Dim cTrack As femap.TrackData
    Set cTrack = App.feTrackData

    Dim cSet As femap.Set
    Set cSet = App.feSet

    Dim sSet As femap.Set
    Set sSet = App.feSet

    Dim along(2) As Variant

    cTrack.Start(4)
    rc = FE_OK

    Do While rc = FE_OK
        ' prompt to select two nodes
        nSet.Clear()
        rc = nSet.SelectList(2, FT_NODE)

        n1.Get(nSet.First)
        n2.Get(nSet.Last)

        'If n1.SelectID("Select first node...") = FE_CANCEL Then
        '    Exit Do
        'End If
        'If n2.SelectID("Select second node...") = FE_CANCEL Then
        '    Exit Do
        'End If

        ' create line
        If App.feLinePoints(False, n1.xyz, n2.xyz, True) = FE_FAIL Then
            App.feAppMessage(FCM_ERROR, "Could not create line.")
            Exit Sub
        End If

        App.feViewRegenerate(0)
    Loop

    cTrack.Created(4, cSet.ID, True)

    If App.feAppMessageBox(2, "Project created curves onto surfaces?") Then
        rc = sSet.Select(FT_SURFACE, True, "Select surfaces to project onto...")
        If rc = FE_CANCEL Then
            Exit Sub
        ElseIf rc = FE_NOT_EXIST Then
            App.feAppMessage(FCM_ERROR, "One or more of the selected surfaces do not exist. No curves projected.")
            Exit Sub
        End If

        ' Project curves onto surfaces
        App.feCurveProjectOntoSurfaces(True, along, sSet.ID, cSet.ID, True)

    End If

End Sub
