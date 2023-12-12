' Title: Properties From Selector To Clipboard.BAS
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

    Dim rc As Long

    Dim slct As femap.Selector
    Set slct = App.feSelector

    Dim propSet As femap.Set
    Set propSet = slct.GetSelected(FT_PROP)

    propSet.CopyToClipboard(False)
    App.feAppMessage(FCM_NORMAL, Cstr(propSet.Count()) + " Property IDs copied to clipboard...")

End Sub
