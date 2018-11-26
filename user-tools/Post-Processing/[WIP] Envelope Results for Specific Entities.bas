Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Envelope Results for Specific Entities")
    App.feAppMessage(FCM_NORMAL, "Envelopes results for user-selected elements and/or nodes.")

    Dim rc As Long

    App.feSelectOutput2( title, nBaseOutputSetID, nPreCheckedSetSetID, nPreCheckedVectorSetID, limitOut­putType, limitComplex, limitToEntity, includeCorner, pOutputSets, pOutputVecs )

End Sub