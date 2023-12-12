' Title: Restore Nodes to Associated Curves.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2306

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Restore Nodes to Associated Curves")

    Dim rc As Long

    ' point set
    Dim nSet As femap.Set
    Set nSet = App.feSet

    ' surface set
    Dim cSet As femap.Set
    Set cSet = App.feSet

    ' surface to project onto
    Dim c As femap.Curve
    Set c = App.feCurve

    ' node count
    Dim n_count As Long

    ' prompt to select surfaces
    rc = cSet.Select(FT_CURVE, True, "Select curves...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "One or more selected curves do not exist.")
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    n_count = 0
    Do While c.NextInSet(cSet.ID)
        nSet.Clear()
        rc = nSet.AddRule(c.ID, FGD_NODE_ATCURVE)
        If rc <> FE_OK Then
            App.feAppMessage(FCM_ERROR, "Curve " & Cstr(c.ID) & " has no nodes attached. Curve skipped.")
        End If

        ' restore nodes for current surface
        rc = App.feProjectOntoCurve(FT_NODE, nSet.ID, c.ID)
        If rc = FE_FAIL Then
            App.feAppMessage(FCM_ERROR, "Nodes could not be moved to curve" & Cstr(c.ID) & ".")
        Else
            n_count = n_count + nSet.Count()
        End If
    Loop

    App.feAppMessage(FCM_NORMAL, Cstr(n_count) & " nodes moved onto curves...")

End Sub