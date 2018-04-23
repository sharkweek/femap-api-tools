Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Elements From Selector To Clipboard")

    Dim rc As Long

    Dim slct As femap.Selector
    Set slct = App.feSelector

    Dim elmSet As femap.Set
    Set elmSet = slct.GetSelected(FT_ELEM)

    elmSet.CopyToClipboard(False)
    App.feAppMessage(FCM_NORMAL, Cstr(elmSet.Count()) + " Element IDs copied to clipboard...")

End Sub
