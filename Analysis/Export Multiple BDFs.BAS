Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Export Multiple BDFs")

    Dim rc As Long

    Dim deckSet As femap.Set
    Set deckSet = App.feSet

    Dim runDeck As femap.AnalysisMgr
    Set runDeck = App.feAnalysisMgr

    deck


    runDeck.Active = deckSet.CurrentID
    App.feFileWriteNastran(8, runDeck.title)

End Sub
