' Title: Add Diameter Callouts at Curve Centers.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2301

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' Usage Notes
' -----------
' The primary use case for this script is for creating points at fastener hole
' centers when fasteners can be modeled as simple beam or CBUSH connections and
' the geometry of the hole is undesired. This script will place points at the
' centers of user-selected curves. The points are then optionally added to
' groups by curve radius. Non-circular curves are ignored.
'
' The equivalence tolerance for evaluating curve radii is set by the `precision`
' parameter. The default is 4 decimal places, assuming an Imperial unit system,
' but can be changed to any desired value.

' public declarations
Dim App As femap.model

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Add Diameter Callouts at Curve Centers")

    Dim rc As Long
    Dim precision As Long
    precision = 4

    Dim cSet As femap.Set
    Set cSet = App.feSet

    Dim radii As femap.SortSet
    Set radii = App.feSort

    Dim points As femap.SortSet
    Set points = App.feSort

    Dim crv As femap.Curve
    Set crv = App.feCurve

    Dim pnt As femap.Point
    Set pnt = App.fePoint

    Dim txt As femap.Text
    Set txt = App.feText()

    Dim x As Double
    Dim y As Double
    Dim z As Double

    Dim cntr As Variant
    Dim norm As Variant
    Dim spt As Variant
    Dim ept As Variant
    Dim ang As Double
    Dim radius As Double
    Dim groups As Boolean
    Dim diameter As Boolean
    Dim height As Double
    Dim bold As Boolean
    Dim italic As Boolean
    Dim font As String

    ' Select curves
    rc = cSet.Select(FT_CURVE, True, "Select curves...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected curves do not exist.")
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    ' Create points at circular curves
    Do While crv.NextInSet(cSet.ID)
        If crv.ArcCircleInfo(cntr, normal, spt, ept, ang, radius) = FE_OK Then
            pnt.xyz = cntr
            pnt.ID = pnt.NextEmptyID()
            pnt.Put(pnt.ID)
            radius = Round(radius, precision)
            radii.AddReal(pnt.ID, radius, 0, 0)
            points.AddReal(pnt.ID, pnt.x, pnt.y, pnt.z)
        Else
            App.feAppMessage(FCM_ERROR, "Curve " & Str(crv.ID) & " is not circular. Curve skipped")
        End If
    Loop

    ' If no circular curves are selected
    If radii.Count() = 0 Then
        App.feAppMessage(FCM_ERROR, "No circular curves selected")
        Exit Sub
    End If

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

    ' add diameter callouts
    radii.Reset()
    txt.ModelPosition = True
    txt.AllViews = True
    txt.DrawPointer = False
    txt.DrawBorder = False
    txt.color = 124
    txt.FontNumber = 1
    txt.SetFontData(10, False, False, "Microsoft Sans Serif")
    Do While radii.Next()
        pnt.Get(radii.Current())
        txt.TextPosition(0) = pnt.x
        txt.TextPosition(1) = pnt.y
        txt.TextPosition(2) = pnt.z
        txt.text = CStr(radii.dReal1) & " dia"
        txt.Put(txt.NextEmptyID())
    Loop

    ' Regenerate view and report
    App.feViewRegenerate(0)

End Sub
