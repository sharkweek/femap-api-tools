Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Move Points Onto Curves")

    Dim rc As Long

    ' point set
    Dim pSet As femap.Set
    Set pSet = App.feSet

    ' curve to project onto
    Dim c As femap.Curve
    Set c = App.feCurve

    ' dummy point object
    Dim p As femap.Point
    Set p = App.fePoint

    Dim success As Long

    rc = pSet.Select(FT_POINT, True, "Select points...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected points do not exist.")
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

    Do While p.NextInSet(pSet.ID)
        rc = App.feProjectOntoCurve(FT_POINT, pSet.ID, c.ID)
        If rc = FE_OK Then
            success = success + 1
        End If
    Loop

    App.feAppMessage(FCM_NORMAL, Cstr(success) & " points moved onto curve...")

End Sub