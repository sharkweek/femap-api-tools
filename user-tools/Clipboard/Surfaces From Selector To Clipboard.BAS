Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Surfaces From Selector To Clipboard")

    Dim rc As Long

    Dim slct As femap.Selector
    Set slct = App.feSelector

    Dim surfSet As femap.Set
    Set surfSet = slct.GetSelected(FT_SURFACE)

    surfSet.CopyToClipboard(False)
    App.feAppMessage(FCM_NORMAL, Cstr(surfSet.Count()) + " Surface IDs copied to clipboard...")

End Sub
