' Title: Curves From Selector To Clipboard.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2306

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

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
