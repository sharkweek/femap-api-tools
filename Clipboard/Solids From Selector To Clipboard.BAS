Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Solids From Selector To Clipboard")

    Dim rc As Long

    Dim slct As femap.Selector
    Set slct = App.feSelector

    Dim solidSet As femap.Set
    Set solidSet = slct.GetSelected(FT_SOLID)

    solidSet.CopyToClipboard(False)
    App.feAppMessage(FCM_NORMAL, Cstr(solidSet.Count()) + " Solid IDs copied to clipboard...")

End Sub
