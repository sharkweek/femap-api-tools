Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Replace Node in Multiple Elements")

    Dim rc As Long

    Dim n_1 As femap.Node
    Set n_1 = App.feNode
    Dim nSet_1 As femap.Set
    Set nSet_1 = App.feSet

    Dim n_2 As femap.Node
    Set n_2 = App.feNode
    Dim nSet_2 As femap.Set
    Set nSet_2 = App.feSet

    Dim e As femap.Elem
    Set e = App.feElem

    Dim eSet As femap.Set
    Set eSet = App.feSet

    Dim i As Long
    Dim j As Long

    ' prompt to select elements
    rc = eSet.Select(FT_ELEM, True, "Select elements...")

    ' prompt for first node
    rc = n_1.SelectID("Select node to replace...")
    If rc = FE_CANCEL Then
        Exit Sub
    End If

    ' prompt for second node
    rc = n_2.SelectID("Select new node...")
    If rc = FE_CANCEL Then
        Exit Sub
    End If

    ' Swap node for first set of elements
    i = 0
    eSet.Reset()
    Do While e.NextInSet(eSet.ID)
        'find node in element
        e.Get(e.ID)
        j = 0
        Do While e.node(j) <> n_1.ID  'find node in element
            j = j + 1
        Loop

        ' move to next element if node isn't in last index
        If e.node(j) <> n_1.ID Then
            App.feAppMessage(FCM_ERROR, "Node " & CStr(n_1.ID) & " does not exist in element " & CStr(e.ID) & ". Element skipped...")
        Else
            e.node(j) = n_2.ID  'swap for new node
            e.Put(e.ID)
        End If
    Loop

    If n_1.Deletable(n_1.ID) Then
        If App.feAppMessageBox(2, "Delete old node?") = FE_OK Then
            n_1.Delete(n_1.ID)
        End If
    End If

    App.feViewRegenerate(0)

End Sub
