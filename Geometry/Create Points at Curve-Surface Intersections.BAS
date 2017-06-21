' This script breaks selected curves at the where they intersect with a
' selected broup of surfaces

' TODO: Add feature to add new points to surface meshes

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Create Points at Curve-Surface Intersections")

    Dim tracker As femap.TrackData
    Dim cSet As femap.Set
    Dim sSet As femap.Set
    Dim s As femap.Surface
    Dim c As femap.Curve
    Dim p As femap.Point
    Dim createdPoints As femap.Set

    Set tracker = App.feTrackData
    Set cSet = App.feSet
    Set sSet = App.feSet
    Set s = App.feSurface
    Set c = App.feCurve
    Set p = App.fePoint
    Set createdPoints = App.feSet

    Dim nearXYZ(2) As Double
    Dim pCoord As Variant

    ' Prompt to select surfaces and curves
    If cSet.Select(FT_CURVE,False,"Select curves") = FE_CANCEL Then
        Exit Sub
    ElseIf sSet.Select(FT_SURFACE, False, "Select surfaces") = FE_CANCEL Then
        Exit Sub
    End If

    ' Begin tracking points
    tracker.Start(3)

    ' for each surface....
    Do While s.NextInSet(sSet.ID)
        ' for each curve
        While cSet.Next
            If App.feCoordSurfaceIntersect(cSet.CurrentID, s.ID, nearXYZ, pCoord) <> FE_FAIL Then
                p.xyz = pCoord
                p.Put(p.NextEmptyID)
            End If
        Wend
        cSet.Reset()
    Loop

    ' Report number of points created
    tracker.Created(3, createdPoints.ID, False)
    App.feAppMessage(FCM_NORMAL, createdPoints.Count() & " Point(s) created...")

End Sub
