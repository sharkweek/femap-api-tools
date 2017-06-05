' This script creates points at the intersections between curves and
' surfaces. Options are provided to add created points to groups, add newly
' created points to the intersected surfaces as mesh points, and break the
' intersected curves at newly created points'

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim tracker As femap.TrackData
    Set tracker = App.feTrackData

    Dim p1(2) As Double
    Dim p2 As Variant
    Dim newPointID As Long

    Dim cSet As femap.Set
    Dim s As femap.Surface
    Dim p As femap.Point
    Dim sSet As femap.Set

    Dim createdPoints As femap.Set
    Set createdPoints = App.feSet

    Set cSet = App.feSet
    Set sSet = App.feSet
    Set s = App.feSurface
    Set p = App.fePoint

    ' Prompt to select surfaces and curves
    If sSet.Select(FT_SURFACE, False, "Select surfaces") = FE_CANCEL Then
        Exit Sub
    ElseIf cSet.Select(FT_CURVE,False,"Select curves") = FE_CANCEL Then
        Exit Sub
    End If

    ' Prompt user for options
    Begin Dialog UserDialog 280,105,"Intersection Point Options" ' %GRID:10,7,1,1
        OKButton 40,77,90,21
        CancelButton 150,77,90,21
        CheckBox 20,14,240,14,"Create group from created points",.groupPoints
        CheckBox 20,35,220,14,"Add points to surface mesh",.meshPoints
        CheckBox 20,56,230,14,"Break curves at points",.breakCurves
    End Dialog
    Dim dlg As UserDialog
    Dialog dlg

    ' Begin tracking created points
    tracker.Start(3)

    ' Create points
    Do While s.NextInSet(sSet.ID)
        While cSet.Next
            If App.feCoordSurfaceIntersect(cSet.CurrentID,s.ID,p1,p2) <> FE_FAIL Then
                'Create point
                p.xyz = p2
                newPointID = p.NextEmptyID
                p.Put(newPointID)

                ' Add point as mesh point to surface
                If dlg.meshPoints And s.AddMeshPoint(newPointID) = FE_INVALID Then
                    App.feAppMessage(FCM_ERROR, "Maximum number of mesh points for surface has been exceeded")
                End If

                ' break curve at point
                If dlg.breakCurves Then
                    App.feCurveBreak(cSet.CurrentID, p.xyz)
                End If
            End If
        Wend
        cSet.Reset
    Loop

    ' Report number of points created
    tracker.Stop(3)
    If tracker.Created(3, createdPoints.ID, False) = FE_NOT_AVAILABLE Then
        App.feAppMessage(FCM_ERROR, "None of the selected curves intersect with the selected surfaces.")
    Else
        App.feAppMessage(FCM_NORMAL, createdPoints.Count() & " Point(s) created...")

        ' Add created points to group
        If dlg.groupPoints Then
            Dim grp As femap.Group
            Set grp = App.feGroup

            ' Create group
            grp.SetAdd(FT_POINT, createdPoints.ID)
            grp.title = "Points from Surface-Curve Intersections"
            grp.Put(grp.NextEmptyID)
            App.feAppMessage(FCM_NORMAL, "Group created...")
        End If
    End If

End Sub
