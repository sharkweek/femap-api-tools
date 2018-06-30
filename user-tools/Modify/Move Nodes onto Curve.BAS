Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Move Nodes onto Curve")

    Dim rc As Long

    ' point set
    Dim nSet As femap.Set
    Set nSet = App.feSet

    ' curve to project onto
    Dim c As femap.Curve
    Set c = App.feCurve

    ' dummy point object
    Dim n As femap.Node
    Set n = App.feNode

    Dim success As Long

    rc = nSet.Select(FT_NODE, True, "Select nodes...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected nodes do not exist.")
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    rc = c.SelectID("Select curve to project onto...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected curve does not exist.")
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    success = 0

    Do While n.NextInSet(nSet.ID)
        rc = App.feProjectOntoCurve(FT_NODE, nSet.ID, c.ID)
        If rc = FE_OK Then
            success = success + 1
        End If
    Loop

    App.feAppMessage(FCM_NORMAL, Cstr(success) & " nodes moved onto curve...")

End Sub