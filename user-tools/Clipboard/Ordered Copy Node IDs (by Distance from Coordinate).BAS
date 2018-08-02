Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Ordered Copy Node IDs (by Distance from Coordinate)")

    Dim rc As Long

    Dim n As femap.Node
    Set n = App.feNode

    Dim basePoint As Variant
    Dim i As Long

    Dim nList As String

    Dim nSort As femap.SortSet
    Set nSort = App.feSort

    Dim nSet As femap.Set
    Set nSet = App.feSet

    rc = nSet.Select(FT_NODE, True, "Select nodes to be copied")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "One or more selected nodes do not exist.")
    End If

    ' prompt to select location
    rc = App.feCoordPick("Select coordinate...", basePoint)
    If rc = FE_CANCEL Then
        Exit Sub
    End If

    ' add to sort object and sort
    nSort.AddAroundPoint(FT_NODE, nSet.ID, basePoint)
    nSort.SortRemoveDuplicates(True)

    ' add ordered nodes to string
    nSort.Reset()
    nList = ""
    For i = 1 To nSort.Count()
        nSort.Next()
        nList = nList & CStr(nSort.ID) & vbCrLf
    Next i

    ' copy to clipboard
    Clipboard(nList)

End Sub
