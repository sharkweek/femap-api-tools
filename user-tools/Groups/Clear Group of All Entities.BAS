' Clear Group

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Clear Group")

    Dim rc As Long

    Dim gr As femap.Group
    Set gr = App.feGroup

    Dim grID As Long

    ' Select group to purge
    rc = App.feSet.SelectID(FT_GROUP, "Select group to clear...", grID)
    If rc = FE_CANCEL Then
        Exit Sub
    End If

    ' Get group
    gr.Get(grID)

    'Purge all
    gr.RangeDeleteAll(-1)

    ' Assign back to group
    gr.Put(grID)

    ' Regenerate
    App.feViewRegenerate(0)

End Sub
