Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Nodes From Selector To Clipboard")

    Dim rc As Long

    Dim slct As femap.Selector
    Set slct = App.feSelector

    Dim nodeSet As femap.Set
    Set nodeSet = slct.GetSelected(FT_NODE)

    nodeSet.CopyToClipboard(False)
    App.feAppMessage(FCM_NORMAL, Cstr(nodeSet.Count()) + " Node IDs copied to clipboard...")

End Sub
