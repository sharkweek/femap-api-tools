Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim rc As Long

    Dim slct As femap.Selector
    Set slct = App.feSelector

    Dim propSet As femap.Set
    Set propSet = slct.GetSelected(FT_PROP)

    propSet.CopyToClipboard(False)
    App.feAppMessage(FCM_NORMAL, Cstr(propSet.Count()) + " Property IDs copied to clipboard...")

End Sub
