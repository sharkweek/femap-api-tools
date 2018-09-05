Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Show Free 1D Elements")
    App.feAppMessage(FCM_NORMAL, "Highlights all 1D elements that have one or more free ends.")

    Dim rc As Long

    Dim e As femap.Elem
    Set e = App.feElem

    Dim n As femap.Node
    Set n = App.feNode

    Dim eSet As femap.Set
    Set eSet = App.feSet

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim i As Long

    ' add all 1D elements to eSet
    For i = 1 To 10
        eSet.AddRule(i, FGD_ELEM_BYTYPE)
    Next i
    eSet.AddRule(37, FGD_ELEM_BYTYPE)

    ' add all nodes on 1D elements to nSet
    nSet.AddSetRule(eSet.ID, FGD_NODE_ONELEM)

    ' remove all nodes with more than one element attached
    nSet.Reset()
    Do While n.NextInSet(nSet.ID)
        n.Get(n.ID)
        If n.NumberOfElements > 1 Then
            nSet.Remove(n.ID)
        End If
    Loop

    ' add only free elements to eSet
    eSet.Clear()
    eSet.AddSetRule(nSet.ID, FGD_ELEM_BYNODE)

    ' show elements
    App.feViewRegenerate(0)
    App.feViewShow(FT_ELEM, eSet.ID)

    App.feAppMessage(FCM_NORMAL, CStr(eSet.Count()) & " elements highlighted...")

    ' prompt to copy to clipboard
    rc = App.feAppMessageBox(2, "Copy list of elements to clipboard?")
    If rc = FE_OK Then
        eSet.CopyToClipboard(False)
        App.feAppMessage(FCM_NORMAL, CStr(eSet.Count()) & " element IDs copied to clipboard...")
    End If

End Sub
