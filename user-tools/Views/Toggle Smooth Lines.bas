' Toggle smooth lines on or off

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