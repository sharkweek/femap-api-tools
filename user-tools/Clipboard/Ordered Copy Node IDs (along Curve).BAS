Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Ordered Copy Nodes (along Curve)")

    Dim rc As Long

    Dim crv As femap.Curve
    Set crv = App.feCurve

    Dim crvNodes As femap.Set
    Set crvNodes = App.feSet

    Dim n As femap.Node
    Set n = App.feNode

    Dim nSort As femap.SortSet
    Set nSort = App.feSort

    Dim xyz As Variant
    Dim s As Double
    Dim nList As String
    Dim i As Long

    ' pick curve
    rc = crv.SelectID("Select curve...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected curve does not exist.")
        Exit Sub
    End If

    ' get nodes on curve
    crv.NodesAsSet(True, True, crvNodes.ID, True)
    If crvNodes.Count() < 2 Then
        App.feAppMessage(FCM_ERROR, "Curve must have 2 or more nodes associated with it.")
        Exit Sub
    End If

    ' get xyz-coords from nodes
    Do While n.NextInSet(crvNodes.ID)
        crv.XYZToParam(n.xyz, s)
        nSort.AddReal(n.ID, s, 0, 0)
    Loop

    ' sort by curve coordinate
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
