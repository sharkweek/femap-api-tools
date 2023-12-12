' Title: Toggle Node Labels.BAS
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

    Dim View As femap.View
    Set View = App.feView

    Dim viewID As Long

    App.feAppGetActiveView(viewID)

    With View
        .Get(viewID)

        If .Label(FVI_NODE) = 0 Then
            .Label(FVI_NODE) = 1
        ElseIf .Label(FVI_NODE) <> 0 Then
            .Label(FVI_NODE) = 0
        End If

        .Put(viewID)
        .Regenerate
    End With

    If View.Label(FVI_NODE) Then
        App.feAppMessage(FCM_COMMAND, "Toggle Node Labels On")
    Else
        App.feAppMessage(FCM_COMMAND, "Toggle Node Labels Off")
    End If

End Sub
