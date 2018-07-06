Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Merge Two Nodes (Keep Lower)")

    Dim rc As Long

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim n As femap.Node
    Set n = App.feNode

    Dim vecBase As Variant
    Dim vecTip As Variant
    Dim vecDist As Variant
    Dim mergeTol As Double

    rc = FE_OK

    Do While rc = FE_OK
        ' prompt to select two nodes
        If n.SelectID("Select first node...") = FE_CANCEL Then
            Exit Do
        Else
            nSet.Add(n.ID)
        End If

        If n.SelectID("Select second node...") = FE_CANCEL Then
            Exit Do
        Else
            nSet.Add(n.ID)
        End If

        ' calculate merge tolerance
        rc = App.feMeasureDistanceBetweenNodes2(nSet.First(), _
                                                nSet.Last(), _
                                                0, _
                                                0, _
                                                0, _
                                                vecBase, _
                                                vecTip, _
                                                vecDist, _
                                                mergeTol)
        mergeTol = 1.1 * mergeTol

        ' merge nodes
        rc = App.feCheckCoincidentNode3(nSet.ID, _
                                        mergeTol, _
                                        True, _
                                        1, _
                                        1, _
                                        False, _
                                        False, _
                                        -1, _
                                        1, _
                                        False)

        App.feViewRegenerate(0)
        nSet.Clear()
    Loop

End Sub
