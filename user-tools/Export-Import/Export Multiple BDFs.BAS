Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Export Multiple BDFs")

    Dim rc As Long

    Dim deckSet As femap.Set
    Set deckSet = App.feSet

    Dim deck As femap.AnalysisMgr
    Set deck = App.feAnalysisMgr
    Dim deckID As Long
    Dim deckTitle as String 'Placeholder variable to replace spaces with underscores
    Dim deckFormat As Integer  'field format for write deck

    Dim emptyGroup As femap.Group
    Set emptyGroup = App.feGroup

    Dim analysisGroupID As Long

    'Get model name and path
    Dim model As Huge_
    Dim modelFileFull As String
    Dim modelFileName As String
    Dim modelFilePath As String
    App.feAppGetModel(model)
    rc = App.feAppGetModelName(model,modelFileFull)
    modelFileName = Dir(modelFileFull)
    modelFilePath = Left(modelFileFull, Len(modelFileFull)-Len(modelFileName))

    rc = deckSet.SelectMultiID(FT_AMGR_DIR, 1, "Select analysis sets to export...")
    If rc = FE_CANCEL Then
        App.feAppMessage(FCM_NORMAL, "Export canceled...")
        Exit Sub
    End If

    rc = App.feAppMessageBox(3, "Do you wish to export with empty Bulk Data sections?")

    If rc = FE_FAIL Then  'If user decides to keep Bulk Data sections
        While deck.NextInSet(deckSet.ID) = FE_OK

            deck.Active = deckSet.CurrentID
            'clean title
            deckTitle = Replace(deck.Title," ", "_") & ".dat"
            deckTitle = Replace(deck.Title,"/", "-") & ".dat"
            deckTitle = Replace(deck.Title,"\", "-") & ".dat"
            deckTitle = Replace(deck.Title,":", "-") & ".dat"
            deckTitle = Replace(deck.Title,"*", "-") & ".dat"
            deckTitle = Replace(deck.Title,"%", "percent") & ".dat"
            App.feFileWriteNastran(8, deckTitle)

        Wend

        App.feAppMessage(FCM_NORMAL, Cstr(deckSet.Count) & " analysis sets exported...")

    ElseIf rc = FE_OK Then

        'Create empty group for writing out decks
        emptyGroup.Put(emptyGroup.NextEmptyID)

        'Write out each deck in the set
        While deck.NextInSet(deckSet.ID) = FE_OK

            'store original deck properties
            deck.Get(deckSet.CurrentID)
            analysisGroupID = deck.NasBulkGroupID
            App.feAppMessage(FCM_NORMAL, "Analysis Group ID = " & Cstr(analysisGroupID))

            'Reassign deck group to empty group
            deck.NasBulkGroupID = emptyGroup.ID
            deck.Put(deck.ID)

            'Write out deck
            deck.Active = deck.ID
            deckTitle = Replace(deck.Title," ", "_") & ".dat"
            deckTitle = Replace(deck.Title,"/", "-") & ".dat"
            deckTitle = Replace(deck.Title,"\", "-") & ".dat"
            deckTitle = Replace(deck.Title,":", "-") & ".dat"
            deckTitle = Replace(deck.Title,"*", "-") & ".dat"
            deckTitle = Replace(deck.Title,"%", "percent") & ".dat"
            App.feFileWriteNastran(8, deckTitle)

            'Restore original deck output group
            deck.NasBulkGroupID = analysisGroupID
            deck.Put(deck.ID)

        Wend

        'Delete empty group
        emptyGroup.Delete(emptyGroup.ID)

        App.feAppMessage(FCM_NORMAL, Cstr(deckSet.Count) & " exported...")

    ElseIf rc = FE_CANCEL Then  'If user cancels export

        App.feAppMessage(FCM_NORMAL, "Export canceled...")
        Exit Sub

    End If

    'Reset active analysis set to original

End Sub
