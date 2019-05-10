' Title: Delete All User Graphics.BAS
' Author: Andy Perez
' Date: May 2019
' Femap API Version: 12.0

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    With App
        .feAppMessage(FCM_COMMAND, "Delete All User Graphics")

        If .feAppMessageBox(2, "Are you sure you wish to delete all user graphics?") = FE_OK Then
            .feGFXDelete(True, 0)
        End If

        .feGFXReset()
        .feViewRegenerate(0)
    End With
End Sub
