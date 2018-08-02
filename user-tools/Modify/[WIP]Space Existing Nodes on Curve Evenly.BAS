Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Space Existing Nodes on Curve Evenly")

    Dim rc As Long

    Dim crv As femap.Curve
    Set crv = App.feCurve

    ' pick curve
    rc = crv.SelectID()

    ' get nodes on curv
    crv.Nodes(getBoundary, getInterior, numNodes, nodeID)

    ' get xyz-coords from nodes

    ' transform node coordinates into parametric values
    crv.XYZToParam(xyz, s)

    ' sort nodes by parametric values

    ' create evenly spaced parametric values based on number of nodes

    ' convert new parametric values to xyz-coords
    crv.ParamToXYZ(s, xyz)

    ' move nodes to new xyz-coordinates

End Sub
