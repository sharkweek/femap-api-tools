Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim rc As Long

    Dim slct As femap.Selector
    Set slct = App.feSelector

    Dim matSet As femap.Set
    Set matSet = slct.GetSelected(FT_MATL)

    matSet.CopyToClipboard(False)
    App.feAppMessage(FCM_NORMAL, Cstr(matSet.Count()) + " Material IDs copied to clipboard...")

End Sub
