Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Grounded CBUSHs at Nodes")
    App.feAppMessage(FCM_NORMAL, "Creates very stiff, grounded CBUSHs at selected nodes. Note: this command differs from the default")
    App.feAppMessage(FCM_NORMAL, "Femap script in that it creates a grounded Nastran CBUSH instead of creating a normal CBUSH")
    App.feAppMessage(FCM_NORMAL, "attached to a node fixed with permanent constraints. See QRG entry for CBUSH for more information.")

    Dim rc As Long

    Dim n As femap.Node
    Set n = App.feNode

    Dim e As femap.Elem
    Set e = App.feElem

    Dim p As femap.Prop
    Set p = App.feProp

    Dim g As femap.Group
    Set g = App.feGroup

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim eSet As femap.Set
    Set eSet = App.feSet

    Dim pSet As femap.Set
    Set pSet = App.feSet

    Dim tracker As femap.TrackData
    Set tracker = App.feTrackData

    Dim existProp As Boolean
    Dim pid As Long
    Dim eid As Long

    ' check if existing grounded CBUSH property exists in model
    rc = pSet.AddRule(41, FGD_PROP_BYTYPE)
    If pSet.Count() > 0 Then
        rc = App.feAppMessageBox(3, "Use existing CBUSH property?")
        Select Case rc
            Case FE_OK
                existProp = True
            Case FE_FAIL
                existProp = False
            Case FE_CANCEL
                Exit Sub
        End Select
    End If

    If existProp Then
        ' prompt to select existing property
        If p.SelectMultiID(FT_PROP, pSet.ID, "Select property...") = FE_CANCEL Then
            Exit Sub
        Else
            g.Add(FT_PROP, p.ID)  ' add property to group
        End If
    Else
        ' create grounded CBUSH property
        With p
            .title = "Grounded CBUSH"
            .type = 41 ' grounded spring/damper property type
            .cbush = 1
            .flag(3) = True

            ' set spring stiffnesses for each DOF
            For i = 36 To 41
                .pval(i) = 1000000000
            Next i

            ' store in model
            pid = p.NextEmptyID
            .Put(pid)
        End With

        ' add property to group
        g.Add(FT_PROP, pid)
    End If

    ' prompt user to select nodes
    rc = nSet.Select(FT_NODE, True, "Select nodes...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_FAIL Then
        rc = App.feAppMessage(FCM_ERROR, "Selected nodes do not exist.")
    End If

    ' create CBUSHs at nodes
    nSet.Reset()
    Do While n.NextInSet(nSet.ID)
        With e
            ' set properties
            .type = 41
            .propID = pid
            .topology = 9
            .node(0) = n.ID
            .SpringNoOrient = False
            .SpringUseCID = False
            .SpringUseLocation = False
            .SpringLocation = 0
            .SpringUsePropCID = True
            .SpringUsePropLocation = True

            ' store in model
            eid = e.NextEmptyID()
            .Put(eid)
        End With

        ' add property to group
        g.Add(FT_ELEM, eid)
    Loop

    ' put group in model
    g.title = "Grounded CBUSHs"
    g.AddRelated()
    g.Put(g.NextEmptyID)

    App.feViewRegenerate(0)

End Sub