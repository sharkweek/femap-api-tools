Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Copy Node IDs on Elements")

    Dim rc As Long

    Dim elmSet As femap.Set
    Set elmSet = App.feSet

    Dim nodeSet As femap.Set
    Set nodeSet = App.feSet

    rc = elmSet.Select(FT_ELEM, True, "Select elements...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "None of the selected elements exist...")
        Exit Sub
    ElseIf rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "No elements selected")
        Exit Sub
    End If

    nodeSet.AddSetRule(elmSet.ID, FGD_Node_onElem)

    nodeSet.CopyToClipboard(False)
    App.feAppMessage(FCM_NORMAL, Cstr(nodeSet.Count()) + " Node IDs copied to clipboard...")

End Sub
