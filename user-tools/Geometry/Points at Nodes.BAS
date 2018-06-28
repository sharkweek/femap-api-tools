Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Create Points at Nodes")

    ' Node set
    Dim nSet As femap.Set
    Set nSet = App.feSet

    ' Node
    Dim n As femap.Node
    Set n = App.feNode

    ' Point
    Dim p As femap.Point
    Set p = App.fePoint

    ' definition coordinate system flag
    Dim def_csys As Boolean

    ' point creation tracker
    Dim p_track As femap.TrackData
    Set p_track = App.feTrackData

    ' set for created points
    Dim pSet As femap.Set
    Set pSet = App.feSet

    ' Prompt to select nodes
    If nSet.Select(FT_NODE, False, "Select nodes...") = FE_CANCEL Then
        Exit Sub
    End If

    rc = App.feAppMessageBox(3, "Create in points in the nodal coordinate system?")
    If rc = FE_OK Then
        def_csys = True
    ElseIf rc = FE_FAIL Then
        def_csys = False
    Else
        Exit Sub
    End If

    ' Initiate tracker
    p_track.Start(3)

    Do While n.NextInSet(nSet.ID)
        p.xyz = n.xyz

        If def_csys Then
            p.defCSys = n.defCSys
        Else
            p.defCSys = 0
        End If

        p.Put(p.NextEmptyID)
    Loop

    p_track.Created(3, pSet.ID, False)

    App.feAppMessage(FCM_NORMAL, Cstr(pSet.Count()) & " points created...")

End Sub
