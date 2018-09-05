Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Show Element Property Boundaries")
    App.feAppMessage(FCM_NORMAL, "Highlights all nodes shared by elements of different properties.")
    App.feAppMessage(FCM_HIGHLIGHT, "NOTE: This command takes a significant amount of time. Be patient!")

    Dim rc As Long

    Dim n As femap.Node
    Set n = App.feNode

    Dim e As femap.Elem
    Set e = App.feElem

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim bNodeSet As femap.Set
    Set bNodeSet = App.feSet

    Dim tSet As femap.Set
    Set tSet = App.feSet

    Dim eSet As femap.Set
    Set eSet = App.feSet

    Dim pid as Long

    ' add nodes to set
    rc = App.feAppMessageBox(3, "Evaluate full model?")
    If rc = FE_OK Then
        nSet.AddAll(FT_NODE)
    ElseIf rc = FE_FAIL Then
        PICK_NODES:
        If nSet.Select(FT_NODE, True, "Select nodes to evaluate...") = FE_NOT_EXIST Then
            App.feAppMessageBox(0, "One or more selected nodes do not exist.")
            GoTo PICK_NODES
        End If
    Else
        Exit Sub
    End If

    ' find nodes with multiple property elements
    Do While n.NextInSet(nSet.ID)
        ' add node to temp set tSet
        tSet.Clear()
        tSet.Add(n.ID)
        n.Get(n.ID)

        ' add all elements attached to node to eSet
        eSet.Clear()
        eSet.AddSetRule(tSet.ID, FGD_ELEM_BYNODE)
        If eSet.Count() < 2 Then
        	GoTo NEXT_NODE
        End If

        ' set first prop for comparison
        e.Get(eSet.First())
        pid = e.propID

        ' add nodes with multiple properties elements to bNodeSet
        Do While e.NextInSet(eSet.ID)
            e.Get(e.ID)
            If e.propID <> pid Then
                bNodeSet.Add(n.ID)
                GoTo NEXT_NODE
            End If
        Loop
        NEXT_NODE:
    Loop

    App.feViewShow(FT_NODE, bNodeSet.ID)

    ' prompt to copy to clipboard
    rc = App.feAppMessageBox(1, "Copy list of nodes to clipboard?")
    If rc = FE_OK Then
        bNodeSet.CopyToClipboard(False)
        App.feAppMessage(FCM_NORMAL, CStr(eSet.Count()) & " node IDs copied to clipboard...")
    End If

End Sub
