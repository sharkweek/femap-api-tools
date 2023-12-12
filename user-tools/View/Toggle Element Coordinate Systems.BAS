' Title: Toggle Element Coordinate Systems.BAS
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
        .Draw(FVI_ELEM_CSYS) = Not .Draw(FVI_ELEM_CSYS)
        .Put(viewID)

        If .Draw(FVI_ELEM_CSYS) Then
            App.feAppMessage(FCM_COMMAND, "Element Coordinate Systems On")
        Else
            App.feAppMessage(FCM_COMMAND, "Element Coordinate Systems Off")
        End If

        .Regenerate
    End With

End Sub
