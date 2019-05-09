' Title: Turn Off User Graphics.BAS
' Author: Andy Perez
' Date: May 2019
' Femap API Version: 12.0

Sub Main
    If feFemap().feGFXReset() = FE_OK Then
        feFemap().feAppMessage(FCM_COMMAND, "User Graphics disabled...")
    End If
End Sub