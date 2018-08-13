Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Move Nodes onto Surface")

    Dim rc As Long

    ' point set
    Dim nSet As femap.Set
    Set nSet = App.feSet

    ' surface to project onto
    Dim s As femap.Surface
    Set s = App.feSurface

    ' dummy point object
    Dim n As femap.Node
    Set n = App.feNode

    Dim i As Long
    Dim attach As Boolean
    Dim n_count As Long
    Dim attach_id() As Long
    Dim attach_type() As Long
    Dim n_array As Variant

    ' prompt user to select nodes
    rc = nSet.Select(FT_NODE, True, "Select nodes...")
    If rc = FE_OK Then
        n_count = nSet.Count()
        ReDim attach_id(n_count)
        ReDim attach_type(n_count)
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected nodes do not exist.")
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    ' prompt to select surface
    rc = s.SelectID("Select surface to project onto...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected surface does not exist.")
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    ' Choose whether or not to attach nodes to surface
    rc = App.feAppMessageBox(3, "Associate nodes with surface?")
    If rc = FE_OK Then
        attach = True
    ElseIf rc = FE_FAIL Then
        attach = False
    Else
        Exit Sub
    End If

    ' move nodes onto surface
    i = 0
    Do While n.NextInSet(nSet.ID)
        attach_type(i) = FT_SURFACE
        attach_id(i) = s.ID
        i = i + 1
    Loop
    rc = App.feProjectOntoSurface(FT_NODE, nSet.ID, s.ID)
    If rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "Nodes could not be moved to surface.")
    End If

    App.feAppMessage(FCM_NORMAL, Cstr(i) & " nodes moved onto surface...")

    ' associate nodes with surface
    If attach Then
        nSet.GetArray(n_count, n_array)

        App.feMeshDetach(FT_NODE, nSet.ID, 0, 0, False)
        rc = App.feMeshAttachNodes(n_count, n_array, attach_type, attach_id)
        If rc = FE_OK Then
        App.feAppMessage(FCM_NORMAL, Cstr(i) & " nodes associated with surface...")
        End If
    End If

End Sub