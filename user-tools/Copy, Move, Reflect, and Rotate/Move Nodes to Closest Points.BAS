' Title: Move Nodes to Closest Points.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 12.0

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Move Nodes to Closest Points")
    App.feAppMessage(FCM_NORMAL, "Moves nearest nodes to selected points.")

    Dim rc As Long

    Dim pSet As femap.Set
    Set pSet = App.feSet

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim p As femap.Point
    Set p = App.fePoint

    Dim n As femap.Node
    Set n = App.feNode

    ' prompt user to select points
    rc = pSet.Select(FT_POINT, True, "Select points...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_FAIL Then
        rc = App.feAppMessage(FCM_ERROR, "Selected points do not exist.")
    End If

    ' get nodes
    pSet.Reset()
    Do While p.NextInSet(pSet.ID)
        p.Get(p.ID)
        n.GetClosest(p.xyz)
        n.xyz = p.xyz
        n.Put(n.ID)
    Loop

    App.feViewRegenerate(0)

End Sub