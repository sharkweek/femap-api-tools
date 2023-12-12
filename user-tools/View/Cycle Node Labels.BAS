' Title: Cycle Node Labels.BAS
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
    Dim labelType As String

    App.feAppGetActiveView(viewID)

    With View
        .Get(viewID)

        If .Label(FVI_NODE) < 4 Then
            .Label(FVI_NODE) = .Label(FVI_NODE) + 1
        ElseIf .Label(FVI_NODE) = 4 Then
            .Label(FVI_NODE) = 1
        End If

        Select Case .Label(FVI_NODE)
            Case 1
                App.feAppMessage(FCM_COMMAND, "Node IDs On")
            Case 2
                App.feAppMessage(FCM_COMMAND, "Node Definition CSYS IDs On")
            Case 3
                App.feAppMessage(FCM_COMMAND, "Node Output CSYS IDs On")
            Case 4
                App.feAppMessage(FCM_COMMAND, "Node Superelement IDs On")
        End Select

        .Put(viewID)
        .Regenerate
    End With

End Sub
