' This script breaks selected curves at the where they intersect with a
' selected broup of surfaces

' TODO: Add feature to add new points to surface meshes

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Break Curves at Surface Intersections")

    Dim cSet As femap.Set
    Dim sSet As femap.Set
    Dim s As femap.Surface
    Dim c As femap.Curve

    Set cSet = App.feSet
    Set sSet = App.feSet
    Set s = App.feSurface
    Set c = App.feCurve

    Dim nearXYZ(2) As Double
    Dim pCoord As Variant
    Dim breakFail As Long

    ' Prompt to select surfaces and curves
    If cSet.Select(FT_CURVE,False,"Select curves to break") = FE_CANCEL Then
        Exit Sub
    ElseIf sSet.Select(FT_SURFACE, False, "Select intersecting surfaces") = FE_CANCEL Then
        Exit Sub
    End If

    breakFail = 0

    ' for each surface....
    Do While s.NextInSet(sSet.ID)
        ' for each curve
        While cSet.Next
            If App.feCoordSurfaceIntersect(cSet.CurrentID, s.ID, nearXYZ, pCoord) <> FE_FAIL Then
                ' break curve at point
                rc = App.feCurveBreak(cSet.CurrentID, pCoord)

                ' track unsuccessful breaks
                If rc = FE_FAIL Then
                    breakFail = breakFail + 1
                End If
            End If
        Wend
        cSet.Reset()
    Loop

    ' Report number of points created
    If breakFail > 0 Then
        App.feAppMessage(FCM_ERROR, "None of the selected curves intersect with the selected surfaces. No curves broken.")
    Else
        App.feAppMessage(FCM_NORMAL, cSet.Count() & " Curve(s) broken...")
    End If

    App.feViewRegenerate(0)

End Sub
