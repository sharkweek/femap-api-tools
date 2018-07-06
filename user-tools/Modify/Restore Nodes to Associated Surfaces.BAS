Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Restore Nodes to Associated Surfaces")

    Dim rc As Long

    ' point set
    Dim nSet As femap.Set
    Set nSet = App.feSet

    ' surface set
    Dim sSet As femap.Set
    Set sSet = App.feSet

    ' surface to project onto
    Dim s As femap.Surface
    Set s = App.feSurface

    ' node count
    Dim n_count As Long

    ' prompt to select surfaces
    rc = sSet.Select(FT_SURFACE, True, "Select surfaces...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "One or more selected surfaces do not exist.")
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If


    n_count = 0
    Do While s.NextInSet(sSet.ID)
        nSet.Clear()
        rc = nSet.AddRule(s.ID, FGD_NODE_ATSURFACE)
        If rc <> FE_OK Then
            App.feAppMessage(FCM_ERROR, Cstr(s.ID) & " has no nodes attached. Surface skipped.")
        End If

        ' restore nodes for current surface
        rc = App.feProjectOntoSurface(FT_NODE, nSet.ID, s.ID)
        If rc = FE_FAIL Then
            App.feAppMessage(FCM_ERROR, "Nodes could not be moved to surface" & Cstr(s.ID) & ".")
        Else
            n_count = n_count + nSet.Count()
        End If
    Loop

    App.feAppMessage(FCM_NORMAL, Cstr(n_count) & " nodes moved onto surfaces...")

End Sub