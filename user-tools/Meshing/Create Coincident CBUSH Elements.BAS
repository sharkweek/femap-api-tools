' This API script will create a "coincident node" CBUSH Element at each
' selected node. The 2nd node of the CBUSH will be created. A CBUSH property
' with an Orinetation CSys must be selected in order for this API to run to
' completion.

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Create Coincident CBUSH Elements")

    Dim n As femap.Node
    Set n = App.feNode

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim p As femap.Prop
    Set p = App.feProp

    Dim pSet As femap.Set
    Set pSet = App.feSet

    Dim e As femap.Elem
    Set e = App.feElem

    Dim pid As Long

    ' Create set of spring/damper properties
    pSet.AddRule(6, FGD_Prop_byType)
    If pSet.Count = 0 Then
        App.feAppMessage(FCM_ERROR, "No Spring/Damper properties exist. Create one and try again.")
        GoTo Done
    End If

    Start:

    pSet.SelectID(FT_PROP, "Select Spring/Damper Property", pid)

    p.Get(pid)
    ptype = p.type

    If ptype = 6 Then
        If p.flag(0) = False Then
            rc = App.feAppMessageBox(2,"Spring/Dmaper Property does not have an Orientation Coordinate System. Choose Another Property? (No = Cancel)")
            If rc = FE_OK Then
                GoTo Start
            Else
                GoTo Done
            End If
        End If

    Else
        rc = App.feAppMessageBox(2, "Not a Spring/Dmaper Property. Choose Another Property? (No = Cancel)")
        If rc = FE_OK Then
            GoTo Start
        Else
            GoTo Done
        End If
    End If

    nSet.Select(FT_NODE, True, "Select Nodes for Coincident CBUSH Elements")

    While nSet.Next
        nid = nSet.CurrentID
        n.Get(nid)
        nidG = n.NextEmptyID
        n.Put(nidG)
        e.type = FET_L_SPRING
        e.topology = FTO_LINE2
        e.Node(0) = nid
        e.Node(1) = nidG
        e.propID = pid
        e.formulation(1) = 1
        e.Put(e.NextEmptyID)
    Wend

    Done:

End Sub
