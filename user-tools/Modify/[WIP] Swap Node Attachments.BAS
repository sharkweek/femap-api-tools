' TODO: account for elements that have both nodes in them (i.e. beams, etc.)

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Swap Node Attachments")

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

    Dim eSet_1 As femap.Set
    Set eSet_1 = App.feSet

    Dim eSet_2 As femap.Set
    Set eSet_2 = App.feSet

    Dim i As Long
    Dim j As Long

    ' prompt for first node
    rc = n_1.SelectID("Select first node...")
    If rc = FE_CANCEL Then
        Exit Sub
    Else
        eSet_1.AddRule(n_1.ID, FGD_ELEM_BYNODE)
    End If

    ' prompt for second node
    rc = n_2.SelectID("Select second node...")
    If rc = FE_CANCEL Then
        Exit Sub
    Else
        eSet_2.AddRule(n_2.ID, FGD_ELEM_BYNODE)
    End If

    ' Swap node for first set of elements
    i = 0
    eSet_1.Reset()
    Do While e.NextInSet(eSet_1.ID)
        'find node in element
        e.Get(e.ID)
        j = 0
        Do While e.node(j) <> n_1.ID  'find node in element
            j = j + 1
        Loop
        e.node(j) = n_2.ID  'swap for new node
        e.Put(e.ID)
    Loop

    ' Swap node for second set of elements
    i = 0
    eSet_2.Reset()
    Do While e.NextInSet(eSet_2.ID)
        e.Get(e.ID)
        j = 0
        Do While e.node(j) <> n_2.ID  'find node in element
            j = j + 1
        Loop
        e.node(j) = n_1.ID  'swap for new node
        e.Put(e.ID)
    Loop

End Sub
