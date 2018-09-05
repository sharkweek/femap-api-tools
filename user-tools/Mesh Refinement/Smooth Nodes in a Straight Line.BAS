' Spaces the nodes evenly in a straight line between the first selected node
' and the selected node furthest from it

' TODO: Add biasing; l_i / l_(i-1) = B^(1/(n-1)) for the i-th element where
'       `B` is the biasing ratio and `n` is the number of elements

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Smooth Nodes in a Straight Line")
    App.feAppMessage(FCM_NORMAL, "This command spaces nodes out evenly in a straight line between the first selected node and the selected node that is")
    App.feAppMessage(FCM_NORMAL, "furthest from the first.")

    Dim rc As Long

    Dim crv As femap.Curve
    Set crv = App.feCurve

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim n As femap.Node
    Set n = App.feNode

    Dim nSort As femap.SortSet
    Set nSort = App.feSort

    Dim cSet As femap.Set
    Set cSet = App.feSet

    Dim pSet As femap.Set
    Set pSet = App.feSet

    Dim tracker As femap.TrackData
    Set tracker = App.feTrackData

    Dim xyz As Variant
    Dim s As Double
    Dim spacing As Double
    Dim dist As Double
    Dim i As Long

    Dim coord1 As Variant
    Dim coord2 As Variant

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
    rc = nSet.Select(FT_NODE, True, "Select remaining nodes...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "One or more selected nodes do not exist.")
        Exit Sub
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    ElseIf nSet.Count() < 2 Then
        App.feAppMessage(FCM_ERROR, "Must select 2 or more additional nodes...")
        Exit Sub
    End If
    nSet.Add(n.ID)

    ' add nodes to sort object and sort by distance from first node
    nSort.AddAroundPoint(FT_NODE, nSet.ID, n.xyz)
    nSort.SortRemoveDuplicates(True)

    ' create xyz array
    ReDim Preserve xyz(nSort.Count() * 3)
    For i = 1 To nSort.Count()
        nSort.Next()
        n.Get(nSort.Current())
        xyz(i * 3 - 3) = n.x
        xyz(i * 3 - 2) = n.y
        xyz(i * 3 - 1) = n.z
        nSort.SetData(n.ID, n.x, n.y, n.z, i, 0, 0, FT_NODE)
    Next i

    ' assign coordinates
    n.Get(nSort.First())
    coord1 = n.xyz
    n.Get(nSort.Last())
    coord2 = n.xyz

    ' create line
    tracker.StartGeometry()
    App.feLinePoints(False, coord1, coord2, False)
    tracker.StopAll()
    tracker.Created(FT_CURVE, cSet.ID, False)
    tracker.Created(FT_POINT, pSet.ID, False)
    pSet.AddRule(cSet.First(), FGD_POINT_ONCURVE)
    crv.Get(cSet.First())

    ' convert xyz-coords to parametric values and sort
    nSort.Reset()
    For i = 1 To nSort.Count()
        nSort.Next()
        n.Get(nSort.Current())
        crv.XYZToParam(n.xyz, s)
        nSort.SetReal(n.ID, s, 0, 0)
    Next i

    nSort.SortRemoveDuplicates(False)

    ' space out nodes evenly (excluding end nodes)
    spacing = 1 / (nSort.Count() - 1)
    dist = 0
    nSort.First()
    App.feAppMessage(FCM_COMMAND, CStr(nSort.Current()))
    For i = 1 To nSort.Count() - 2
        nSort.Next()
        n.Get(nSort.Current())
        dist = dist + spacing
        crv.ParamToXYZ(dist, xyz)
        n.xyz = xyz
        n.Put(n.ID)
    Next i

    ' cleanup dummy geometry
    App.feDelete(FT_CURVE, cSet.ID)
    App.feDelete(FT_POINT, pSet.ID)

    App.feViewRegenerate(0)
    App.feAppMessage(FCM_NORMAL, CStr(nSort.Count() - 2) & " nodes smoothed...")

End Sub
