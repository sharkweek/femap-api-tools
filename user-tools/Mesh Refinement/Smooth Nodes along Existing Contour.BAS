' Spaces the nodes evenly along their existing contour. Only possible up to
' 109 nodes. End noes are *NOT* relocated, but are used to calculate the total
' length over which to space out the nodes. Also note that nodes must be
' selected in the order they are to be spaced.

' TODO: Add biasing; l_i / l_(i-1) = B^(1/(n-1)) for the i-th element where
'       `B` is the biasing ratio and `n` is the number of elements

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Smooth Nodes along Existing Contour")

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

    ' select nodes
    rc = nSort.Select(FT_NODE, True, "Select nodes...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "One or more selected nodes do not exist.")
        Exit Sub
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    ElseIf nSort.Count() < 4 Then
        App.feAppMessage(FCM_ERROR, "Must select 4 or more nodes...")
        Exit Sub
    End If

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

    ' create dummy spline
    tracker.StartGeometry()
    rc = App.feSplinePoints(False, False, nSort.Count(), xyz, False)
    If rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "Unable to create dummy spline. Command canceled...")
        Exit Sub
    Else
        tracker.StopAll()
        tracker.Created(FT_CURVE, cSet.ID, False)
        tracker.Created(FT_POINT, pSet.ID, False)
        pSet.AddRule(cSet.First(), FGD_POINT_ONCURVE)
        crv.Get(cSet.First())
    End If

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
    App.feAppMessage(FCM_NORMAL, CStr(nSort.Count() - 2) & " nodes moved...")

End Sub
