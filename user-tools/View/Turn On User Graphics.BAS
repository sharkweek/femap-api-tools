' Title: Turn On User Graphics.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 12.0

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim gfxSet As femap.Set
    Set gfxSet = App.feSet

    gfxSet.AddRange(0, 1e4, 1)
    App.feGFXSelect(gfxSet.ID, True, True)
    App.feGFXEnableView(False, App.feView.Active)
    App.feAppMessage(FCM_COMMAND, "User Graphics enabled...")

End Sub