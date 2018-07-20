Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Surface Match Mesh Attributes")

    Dim rc As Long

    feMeshAttrSurface(surfaceSET, propID, offset)
    feMeshSizeSurface(surfaceSET, replaceAll, meshSize, minEdge, maxAngle, maxOnSmall, smallSize, vertexAspect, mappedRefinement, growthFactor, refineRatio, refineSurf)

    s.attrPID
    s.attrTopology
    s.attrMaxQuads


End Sub