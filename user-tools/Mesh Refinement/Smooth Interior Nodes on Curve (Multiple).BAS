' Spaces the interior nodes on a curve out evenly with respect to the curve's
' length. If nodes are selected manually, select ONLY the interior nodes.

' TODO: add end point node detection
' TODO: Add biasing; l_i / l_(i-1) = B^(1/(n-1)) for the i-th element where
'       `B` is the biasing ratio and `n` is the number of elements

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Smooth Interior Nodes on Curve (Multiple)")

    Dim rc As Long

    Dim crv As femap.Curve
    Set crv = App.feCurve

    Dim crvNodes As femap.Set
    Set crvNodes = App.feSet

    Dim crvSet As femap.Set
    Set crvSet = App.feSet

    Dim n As femap.Node
    Set n = App.feNode

    Dim nSort As femap.SortSet
    Set nSort = App.feSort

    Dim xyz As Variant
    Dim s As Double
    Dim spacing As Double
    Dim dist As Double
    Dim i As Long

    ' pick curves
    PICK_CURVES:
    rc = crvSet.Select(FT_CURVE, True, "Select curves...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessageBox(0, "One or more selected curves do not exist.")
        GoTo PICK_CURVES
    End If

    crvSet.Reset()
    Do While crv.NextInSet(crvSet.ID)

        ' get nodes on curve (not including ends)
        crv.NodesAsSet(False, True, crvNodes.ID, True)
        If crvNodes.Count() < 2 Then
            App.feAppMessage(FCM_ERROR, "Curve " & CStr(crv.ID) & " has less than two nodes associated with it. Curve skipped...")
            GoTo CURVE_SKIPPED
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

        CURVE_SKIPPED:
    Loop

End Sub
