' Title: Associative Move Nodes onto Curve.BAS
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

    App.feAppMessage(FCM_COMMAND, "Move Nodes onto Curve")

    Dim rc As Long

    ' point set
    Dim nSet As femap.Set
    Set nSet = App.feSet

    ' curve to project onto
    Dim c As femap.Curve
    Set c = App.feCurve

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

    ' prompt user to select curve
    rc = c.SelectID("Select curve to project onto...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected curve does not exist.")
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    ' Choose whether or not to attach nodes to curve
    rc = App.feAppMessageBox(3, "Associate nodes with curve?")
    If rc = FE_OK Then
        attach = True
    ElseIf rc = FE_FAIL Then
        attach = False
    Else
        Exit Sub
    End If

    ' move nodes onto curve
    i = 0
    Do While n.NextInSet(nSet.ID)
        attach_type(i) = FT_CURVE
        attach_id(i) = c.ID
        i = i + 1
    Loop
    rc = App.feProjectOntoCurve(FT_NODE, nSet.ID, c.ID)
    If rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "Nodes could not be moved to curve.")
    End If

    App.feAppMessage(FCM_NORMAL, Cstr(i) & " nodes moved onto curve...")

    ' associate nodes with curve
    If attach Then
        nSet.GetArray(n_count, n_array)

        App.feMeshDetach(FT_NODE, nSet.ID, 0, 0, False)
        rc = App.feMeshAttachNodes(n_count, n_array, attach_type, attach_id)
        If rc = FE_OK Then
        App.feAppMessage(FCM_NORMAL, Cstr(i) & " nodes associated with curve...")
        End If
    End If

End Sub