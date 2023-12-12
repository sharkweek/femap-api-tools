' Title: Closest Nodes to Points.BAS
' Author: Andy Perez
' Femap API Version: 12.0

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Closest Nodes to Points")
    App.feAppMessage(FCM_NORMAL, "Copies the IDs of nodes closest to a set of user-selected points to clipboard. Selects one node per point.")

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
    Do While p.NextInSet(pSet.ID)
        p.Get(p.ID)
        n.GetClosest(p.xyz)
        nSet.Add(n.ID)
    Loop

    nSet.CopyToClipboard(False)
    App.feAppMessage(FCM_NORMAL, Cstr(nSet.Count()) + " Node IDs copied to clipboard...")

    App.feViewRegenerate(0)

End Sub