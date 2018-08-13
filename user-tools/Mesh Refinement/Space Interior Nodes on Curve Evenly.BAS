' Spaces the interior nodes on a curve out evenly with respect to the curve's
' length. If nodes are selected manually, select ONLY the interior nodes.

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Space Nodes on Curve Evenly")

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
    Dim spacing As Double
    Dim dist As Double
    Dim i As Long

    ' pick curve
    rc = crv.SelectID("Select curve...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected curve does not exist.")
        Exit Sub
    End If

    ' get nodes on curve (not including ends)
    crv.NodesAsSet(False, True, crvNodes.ID, True)
    If crvNodes.Count() < 2 Then
        rc = App.feAppMessageBox(1, "Curve has less than two nodes associated with it. Do you wish to select nodes?")
        If rc = FE_OK Then
            crvNodes.Select(FT_NODE, False, "Select nodes")
        Else
            Exit Sub
        End If
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
