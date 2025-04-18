' Title: Hidden Solids To Clipboard.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2306

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' public declarations
Dim App As femap.model

Sub Main
    ' `Main` declarations
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Hidden Solids To Clipboard")

    Dim rc As Long

    Dim solid_set As femap.Set
    Set solid_set = App.feSet()

    Dim sol As femap.Solid
    Set sol = App.feSolid()

    solid_set.AddAll(FT_SOLID)
    App.feEntityGetVisibility(FVIS_SOLID, solid_set.ID, False)
    solid_set.CopyToClipboard(False)

    App.feViewRegenerate(0)

End Sub