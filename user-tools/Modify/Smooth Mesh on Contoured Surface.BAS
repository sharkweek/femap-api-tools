' dimension as public for DialogFunc to access
Dim App As femap.model
Dim fixed_nodes as femap.Set

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Smooth Mesh on Contoured Surface")

    Dim rc As Long

    Dim laplacian As Boolean

    Dim e As femap.Elem
    Set e = App.feElem

    Dim eSet As femap.Set
    Set eSet = App.feSet

    Dim n As femap.Node
    Set n = App.feNode

    Set fixed_nodes = App.feSet

    Dim smoothed_nodes As femap.Set
    Set smoothed_nodes = App.feSet

    Dim fixedSet As femap.Set
    Set fixedSet = App.feSet

    Dim tolerance As Variant
    Dim iterations As Variant

    Dim s As femap.Surface
    Set s = App.feSurface

    Dim i As Long
    Dim nFreeCount As Long
    Dim nFreeData As Variant
    Dim bLineElem As Variant
    Dim n_count As Long
    Dim attach As Boolean
    Dim attach_id() As Long
    Dim attach_type() As Long

    ' prompt to select elements
    rc = eSet.Select(FT_ELEM, True, "Select elements to smooth...")
    If rc = FE_OK and eSet.Count() > 1 Then
        ' make sure all elements are shell elements
        Do While e.NextInSet(eSet.ID)
            Select Case e.topology
                Case 2 To 5  'do nothing if element is a shell element
                Case Else
                    App.feAppMessage(FCM_ERROR, "One or more selected elements are not shell elements.")
                    Exit Sub
            End Select
        Loop
    ElseIf rc = FE_OK and eSet.Count() <= 1 Then
        App.feAppMessage(FCM_ERROR, "Must select more than one element.")
    ElseIf rc =FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "One or more selected elements do not exist.")
        Exit Sub
    End If

    ' prompt to select surface
    rc = s.SelectID("Select surface to project onto...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected surface does not exist.")
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    ' get boundary nodes too fix during smoothing
    App.feElementFreeEdge(eSet.ID, True, False, nFreeCount, nFreeData)
    i = 0
    Do While i < (nFreeCount * 3)
        If i Mod 3 <> 0 Then
            fixed_nodes.Add(nFreeData(i))
        End If
        i = i + 1
    Loop

    ' prompt for smoothing options
    Begin Dialog UserDialog 470,91,"Smoothing Options",.DialogFunc
        GroupBox 20,7,130,77,"Method",.method_group_box
        OptionGroup .method
            OptionButton 40,28,100,14,"Laplacian",.laplacian
            OptionButton 40,49,90,14,"Centroidal",.centroidal
        GroupBox 160,7,190,77,"Iterations",.iterations_group_box
        text 170,28,90,14,"Iterations:",.iterations_text
        TextBox 250,28,90,14,.iterations
        text 170,49,90,14,"Tolerance:",.tolerance_text
        TextBox 250,49,90,14,.tolerance
        OKButton 360,35,100,21
        CancelButton 360,63,100,21
        PushButton 360,7,100,21,"Fix Nodes...",.fix_nodes
    End Dialog
    Dim dlg As UserDialog
    Dialog dlg

    ' validate smoothing tolerance
    If IsNumeric(dlg.tolerance) Then
        tolerance = CDbl(dlg.tolerance)
    Else
        App.feAppMessage(FCM_ERROR, "Tolerance must be a non-zero number.")
    End If

    ' validate iterations
    If IsNumeric(dlg.iterations) Then
        If (CLng(dlg.iterations) - CDbl(dlg.iterations) = 0) And _
        (CLng(dlg.iterations) > 0) Then
            iterations = CLng(dlg.iterations)
        Else
            App.feAppMessage(FCM_ERROR, "Iterations must be a positive non-zero integer.")
        End If
    Else
        App.feAppMessage(FCM_ERROR, "Iterations must be a positive non-zero integer.")
    End If

    ' set laplacian flag
    If dlg.method Then
        laplacian = False
    Else
        laplacian = True
    End If

    App.feMeshSmooth2(eSet.ID, fixed_nodes.ID, laplacian, iterations, tolerance)

    ' get internal, non-fixed nodes
    eSet.Reset()
    Do While e.NextInSet(eSet.ID)  ' add all nodes from selected elements
        smoothed_nodes.AddRule(FGD_Node_onElem, e.ID)
    Loop
    smoothed_nodes.RemoveSet(fixed_nodes.ID)  ' remove the fixed nodes

    n_count = smoothed_nodes.Count()
    ReDim attach_id(n_count)
    ReDim attach_type(n_count)

    ' project smoothed nodes to surface
    i = 0
    smoothed_nodes.Reset()
    Do While n.NextInSet(smoothed_nodes.ID)
        rc = App.feProjectOntoSurface(FT_NODE, smoothed_nodes.ID, s.ID)
        If rc = FE_OK Then
            attach_type(i) = FT_SURFACE
            attach_id(i) = s.ID
            i = i + 1
        End If
    Loop

    ' associate nodes with surface
    If attach Then
        smoothed_nodes.GetArray(n_count, n_array)

        App.feMeshDetach(FT_NODE, smoothed_nodes.ID, 0, 0, False)
        rc = App.feMeshAttachNodes(n_count, n_array, attach_type, attach_id)
        If rc = FE_OK Then
        App.feAppMessage(FCM_NORMAL, Cstr(success) & " nodes associated with surface...")
        End If
    End If

    App.feViewRegenerate(0)

End Sub


Function DialogFunc(DlgItem As String, Action as Integer, SuppValue As Long) As Boolean
    Select Case Action
        Case 1  ' dialog box initialization
            Beep
        Case 2  ' value changing or button pressed
            If DlgItem = "fix_nodes" Then
                fixed_nodes.Select(FT_NODE, False, "Select fixed nodes...")
                DialogFunc = True  ' do not exit dialog
            End If
    End Select
End Function