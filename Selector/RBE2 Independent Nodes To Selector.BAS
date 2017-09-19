Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "RBE2 Independent Nodes To Selector")

    Dim rc As Long

    Dim rigidSet As femap.Set
    Set rigidSet = App.feSet

    Dim nodeSet As femap.Set
    Set nodeSet = App.feSet

    Dim selector As femap.Selector
    Set selector = App.feSelector

    Dim elm As femap.Elem
    Set elm = App.feElem

    rc = rigidSet.Select(FT_ELEM, True, "Select elements...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "None of the selected elements exist...")
        Exit Sub
    ElseIf rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "No elements selected")
        Exit Sub
    End If

    rigidSet.Reset()
    Do While elm.NextInSet(rigidSet.ID)
        ' Remove non-rigid elements from selection set
        If elm.type <> 29 Then
            rigidSet.Remove(elm.ID)
        ' Remove RBE3 elements from selection set
        ElseIf elm.RigidInterpolate Then
            rigidSet.Remove(elm.ID)
        ' Add independent nodes of RBE2s to new set
        Else
            nodeSet.Add(elm.node(0))
        End If
    Loop

    ' Add independent nodes to Selector
    selector.DrillMode = False
    selector.AddSet(FT_NODE, nodeSet.ID, False)

    App.feAppMessage(FCM_NORMAL, Cstr(nodeSet.Count()) + " Nodes added to selector...")

End Sub
