' Toggle Element Coordinate Systems

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
