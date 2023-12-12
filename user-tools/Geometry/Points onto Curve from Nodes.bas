' Title: Points onto Curve from Nodes.bas
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2306

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' Usage Notes
' -----------
' Project points onto curve from a selected group of points

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Points onto Curve from Nodes")

    Dim rc As Long

    ' node set
    Dim nSet As femap.Set
    Set nSet = App.feSet

    ' point set
    Dim pSet As femap.Set
    Set pSet = App.feSet

    ' curve to project onto
    Dim c As femap.Curve
    Set c = App.feCurve

    ' dummy point object
    Dim n As femap.Node
    Set n = App.feNode

    ' point
    Dim p As femap.Point
    Set p = App.fePoint

    Dim pid As Long

    Dim success As Long

    rc = nSet.Select(FT_NODE, True, "Select nodes...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected nodes do not exist.")
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    rc = c.SelectID("Select curve to project onto...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected curve does not exist.")
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    success = 0

    Do While n.NextInSet(nSet.ID)
        p.xyz = n.xyz
        pid = p.NextEmptyID
        p.Put(pid)

        ' add new point to dummy set for feProjectOntoCurve arg
        pSet.Clear()
        pSet.Add(pid)

        rc = App.feProjectOntoCurve(FT_POINT, pSet.ID, c.ID)
        If rc = FE_OK Then
            success = success + 1
        End If
    Loop

    App.feAppMessage(FCM_NORMAL, Cstr(success) & " points projected onto curve...")

End Sub