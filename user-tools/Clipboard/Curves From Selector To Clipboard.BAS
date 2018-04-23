Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Curves From Selector To Clipboard")

    Dim rc As Long

    Dim slct As femap.Selector
    Set slct = App.feSelector

    Dim curveSet As femap.Set
    Set curveSet = slct.GetSelected(FT_CURVE)

    curveSet.CopyToClipboard(False)
    App.feAppMessage(FCM_NORMAL, Cstr(curveSet.Count()) + " Curve IDs copied to clipboard...")

End Sub
