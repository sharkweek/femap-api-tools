' Title: Turn On User Graphics.BAS
' Author: Andy Perez
' Date: May 2019
' Femap API Version: 12.0

Sub Main
    If feFemap().feGFXEnableView(True, 0) = FE_OK Then
        feFemap().feAppMessage(FCM_COMMAND, "User Graphics enabled...")
    End If
End Sub