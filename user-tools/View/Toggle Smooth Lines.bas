' Title: Toggle Smooth Lines.bas
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

    App.Pref_RenderSmoothLines = Not App.Pref_RenderSmoothLines
    App.Pref_RenderRotate(19) = App.Pref_RenderSmoothLines
    App.feViewRegenerate(0)

    If App.Pref_RenderSmoothLines Then
        App.feAppMessage(FCM_COMMAND, "Smooth Lines ON")
    Else
        App.feAppMessage(FCM_COMMAND, "Smooth Lines OFF")
    End If

End Sub