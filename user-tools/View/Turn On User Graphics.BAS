' Title: Turn On User Graphics.BAS
' Author: Andy Perez
' Date: May 2019
' Femap API Version: 12.0

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