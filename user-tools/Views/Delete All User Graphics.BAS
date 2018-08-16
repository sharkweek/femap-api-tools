Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Delete All User Graphics")

    App.feGFXDelete(True, 0)
    App.feGFXReset()
    App.feViewRegenerate(0)

End Sub
