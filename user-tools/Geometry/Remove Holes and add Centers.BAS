' Title: Remove Holes and add Centers.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2301

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.
' public declarations

Dim App As femap.model

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Remove Holes and add Centers")

    Dim rc As Long
    Dim precision As Long
    precision = 4

    Dim cSet As femap.Set
    Set cSet = App.feSet

    Dim radii As femap.SortSet
    Set radii = App.feSort

    Dim crv As femap.Curve
    Set crv = App.feCurve

    Dim pnt As femap.Point
    Set pnt = App.fePoint

    Dim tracker As femap.TrackData
    Set tracker = App.feTrackData

    Dim createdPoints As femap.Set
    Set createdPoints = App.feSet

    Dim cntr As Variant
    Dim tol As Double
    Dim coinCount As Long
    Dim coinPoints As Variant
    Dim firstID As Long
    Dim norm As Variant
    Dim spt As Variant
    Dim ept As Variant
    Dim ang As Double
    Dim radius As Double
    Dim groups As Boolean

    ' Select curves
    rc = cSet.Select(FT_CURVE, True, "Select curves...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected curves do not exist.")
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    ' prompt to create groups
    groups = App.feAppMessageBox(2, "Add created points to groups?")

    ' Clear tracker and begin tracking points
    tracker.Clear()
    tracker.Start(3)
    firstID = pnt.NextEmptyID

    ' Create points at circular curves
    Do While crv.NextInSet(cSet.ID)
        If crv.ArcCircleInfo(cntr, normal, spt, ept, ang, radius) Then
            pnt.xyz = cntr
            pnt.ID = pnt.NextEmptyID()
            pnt.Put(pnt.ID)
            radius = Round(radius, precision)
            radii.AddReal(pnt.ID, radius, 0, 0)
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
            App.feCheckCoincidentPoint(createdPoints.ID, createdPoints.ID, tol, True, False, coinCount, coinPoints )

            ' Remove orphaned point IDs from createdPoints
            If coinCount > 0 Then
                createdPoints.Reset()
                Do While pnt.NextInSet(createdPoints.ID)
                    If pnt.Exist(pnt.ID) = FE_FAIL Then
                        createdPoints.Remove(pnt.ID)
                        radii.Remove(pnt.ID)
                    End If
                Loop
            End If

        ElseIf rc = FE_CANCEL Then
            App.feDelete(FT_POINT, createdPoints.ID)
            App.feAppMessage(FCM_NORMAL, "User canceled. No points created...")
            Exit Sub

        End If
    End If

    ' Remove curves
    App.feSurfaceRemoveHole(cSet.ID)

    ' Regenerate view and report
    App.feViewRegenerate(0)
    App.feAppMessage(FCM_NORMAL, Str(createdPoints.Count()) & " Point(s) created...")

    ' Add created points to group
    If groups = FE_OK Then
        Dim grp As femap.group
        Set grp = App.feGroup()
        Dim real2 As Double
        Dim real3 As Double
        Dim new_radius As Double

        ' first radius group
        radii.Sort(True)  ' sort by radius magnitude
        radii.First()
        radii.Current()
        radius = radii.dReal1
        grp.ID = grp.NextEmptyID()
        grp.title = "Hole Centers - " & CStr(radius * 2) & " dia"
        grp.Put(grp.ID)

        ' create groups for each point
        radii.Reset()
        Do While radii.Next()
            new_radius = radii.dReal1
            If Abs(new_radius - radius) > (1 * 10^(-precision)) Then  ' if current curve has different radius than previous
                'grp.Get(grp.ID)
            'Else
                grp.clear()
                grp.ID = grp.NextEmptyID()
                grp.title = "Hole Centers - " & CStr(new_radius * 2) & " dia"
            End If
            grp.Add(FT_POINT, radii.Current())
            grp.Put(grp.ID)
            radius = new_radius
        Loop
    End If

End Sub
