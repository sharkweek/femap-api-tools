Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Parse Into Layers by Entity Type")
    App.feAppMessage(FCM_NORMAL, "Reassigns entities on selected layer to new layers based on entity type.")

    Dim rc As Long

    Dim propLayID As Long
    Dim meshLayID As Long
    Dim loadLayID As Long
    Dim geomLayID As Long
    Dim matlLayID As Long
    Dim csysLayID As Long
    Dim cnstLayID As Long

    Dim lyr As femap.Layer
    Set lyr = App.feLayer

    Dim lyrGroup As femap.Group
    Set lyrGroup = App.feGroup

    Dim lyrID As Long

    ' prompt user to select layer
    rc = lyr.SelectID("Select layer...")
    If rc = FE_CANCEL Then
        Exit Sub
    End If

    

End Sub

