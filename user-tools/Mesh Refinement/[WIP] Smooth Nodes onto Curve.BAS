' Spaces the interior nodes on a curve out evenly with respect to the curve's
' length. If nodes are selected manually, select ONLY the interior nodes.

' TODO: add end point node detection
' TODO: rework so that outermost nodes are considered the endpoints of the
'       smoothing
' TODO: Add biasing; l_i / l_(i-1) = B^(1/(n-1)) for the i-th element where
'       `B` is the biasing ratio and `n` is the number of elements

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Smooth Nodes onto Curve")

    Dim rc As Long

    Dim crv As femap.Curve
    Set crv = App.feCurve

    Dim crvNodes As femap.Set
    Set crvNodes = App.feSet

    Dim n As femap.Node
    Set n = App.feNode

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim nSort As femap.SortSet
    Set nSort = App.feSort

    Dim xyz As Variant
    Dim s As Double
    Dim spacing As Double
    Dim dist As Double
    Dim i As Long

    ' select first node
    PICK_NODE:
    rc = n.SelectID("Select first node...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessageBox(0, "Selected node does not exist...")
        GoTo PICK_NODE
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If
    n.Get(n.ID)

    ' select remaining nodes
    REMAINING_NODES:
    nSet.Clear()
    rc = nSet.Select(FT_NODE, True, "Remaining nodes...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessageBox(0, "One or more selected nodes do not exist.")
        GoTo REMAINING_NODES
    End If

    ' pick curve
    rc = crv.SelectID("Select curve...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected curve does not exist.")
        Exit Sub
    End If

    ' convert xyz-coords to parametric values and sort
    Do While n.NextInSet(crvNodes.ID)
        crv.XYZToParam(n.xyz, s)
        nSort.AddReal(n.ID, s, 0, 0)
    Loop

    nSort.SortRemoveDuplicates(True)

    ' space out nodes evenly
    spacing = 1 / (crvNodes.Count() + 1)
    dist = 0
    nSort.Reset()
    For i = 1 To nSort.Count()
        nSort.Next()
        n.Get(nSort.Current())
        dist = dist + spacing
        crv.ParamToXYZ(dist, xyz)
        n.xyz = xyz
        n.Put(n.ID)
    Next i

End Sub
