' Title: Turn Off User Graphics.BAS
' Author: Andy Perez
' Date: May 2019
' Femap API Version: 12.0

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feGFXReset() = FE_OK
    App.feAppMessage(FCM_COMMAND, "User Graphics disabled...")
End Sub