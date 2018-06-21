Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Points From Selector To Clipboard")

    Dim rc As Long

    Dim slct As femap.Selector
    Set slct = App.feSelector

    Dim pointSet As femap.Set
    Set pointSet = slct.GetSelected(FT_POINT)

    pointSet.CopyToClipboard(False)
    App.feAppMessage(FCM_NORMAL, Cstr(pointSet.Count()) + " Point IDs copied to clipboard...")

End Sub
