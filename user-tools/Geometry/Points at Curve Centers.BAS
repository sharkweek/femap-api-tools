Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Points at Curve Centers")

    Dim rc As Long

    Dim cSet As femap.Set
    Set cSet = App.feSet

    Dim pSort As femap.SortSet
    Set pSort = App.feSort

    Dim crv As femap.Curve
    Set crv = App.feCurve

    Dim pnt As femap.Point
    Set pnt = App.fePoint

    Dim tracker As femap.TrackData
    Set tracker = App.feTrackData

    Dim createdPoints As femap.Set
    Set createdPoints = App.feSet

    Dim dTab As femap.DataTable
    Set dTab = App.feDataTable

    Dim cntr As Variant
    Dim tol As Double
    Dim coinCount As Long
    Dim coinPoints As Variant
    Dim firstID As Long

    ' Select curves
    rc = cSet.Select(FT_CURVE, True, "Select curves...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected curves do not exist.")
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    ' Clear tracker and begin tracking points
    tracker.Clear()
    tracker.Start(3)
    firstID = pnt.NextEmptyID

    ' Create points at circular curves
    Do While crv.NextInSet(cSet.ID)
        If crv.Center(cntr) Then
            pnt.xyz = cntr
            pnt.Put(pnt.NextEmptyID)
        Else
            App.feAppMessage(FCM_ERROR, "Curve " & Str(crv.ID) & " is not circular. Curve skipped")
        End If
    Loop

    'Stop tracking
    tracker.Stop(3)
    tracker.Created(3, createdPoints.ID, False)

    ' If no circular curves are selected
    If createdPoints.Count() = 0 Then
        App.feAppMessage(FCM_ERROR, "No circular curves selected")
        Exit Sub

    ' merge coincident points
    ElseIf createdPoints.Count() >= 2 Then
        ' prompt user
        rc = App.feAppMessageBox(3, "Merge coincident points?")
        If rc = FE_OK Then
            If App.feGetReal("Merge Tolerance", 1E-20, 1E+20, tol) = FE_CANCEL Then
                App.feDelete(FT_POINT, createdPoints.ID)
                App.feAppMessage(FCM_NORMAL, "User canceled. No points created...")
                Exit Sub
            End If

            ' Merge
            App.feCheckCoincidentPoint( createdPoints.ID, createdPoints.ID, tol, True, False, coinCount, coinPoints )

            ' Remove orphaned point IDs from createdPoints
            If coinCount > 0 Then
                createdPoints.Reset()
                Do While pnt.NextInSet(createdPoints.ID)
                    If pnt.Exist(pnt.ID) = FE_FAIL Then
                        createdPoints.Remove(pnt.ID)
                    End If
                Loop
            End If

            ' renumber remaining points
            App.feRenumber(FT_POINT, createdPoints.ID, firstID)

        ElseIf rc = FE_CANCEL Then
            App.feDelete(FT_POINT, createdPoints.ID)
            App.feAppMessage(FCM_NORMAL, "User canceled. No points created...")
            Exit Sub

        End If
    End If

    ' Regenerate view and report
    App.feViewRegenerate(0)
    App.feAppMessage(FCM_NORMAL, Str(createdPoints.Count()) & " Point(s) created...")

    ' Add created points to group
    If App.feAppMessageBox(2, "Add created points to group?") = FE_OK Then
        Dim grp As femap.Group
        Set grp = App.feGroup

        ' Create group
        grp.SetAdd(FT_POINT, createdPoints.ID)
        grp.title = "Center Points"
        grp.Put(grp.NextEmptyID)
        App.feAppMessage(FCM_NORMAL, "Group created...")
    End If

End Sub
