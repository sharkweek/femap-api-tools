' Title: Remesh on Contoured Surface.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 12.0

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.


Dim App As femap.model
Dim boundary_nodes As femap.Set
Dim e As femap.Elem
Dim eSet As femap.Set
Dim n As femap.Node
Dim nSet as femap.Set
Dim s As femap.Surface

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Remesh on Contoured Surface")
    App.feAppMessage(FCM_NORMAL, "Remesh shell elements on a contoured surface. The new mesh will derive its properties from the original")
    App.feAppMessage(FCM_NORMAL, "surface. New elements and nodes are automatically associated with the surface.")

    '===========================================================================
    ' Declarations
    '===========================================================================

    Dim i As Long
    Dim iterations As Variant
    Dim mesh_approach As Integer
    Dim nFreeCount As Long
    Dim nFreeData As Variant
    Dim p As femap.Prop
    Dim pick_geom As Boolean
    Dim prop_id As Long
    Dim prop_ids() As Long
    Dim prop_titles() As String
    Dim pSet As femap.Set
    Dim rc As Long
    Dim sid As Long
    Dim tolerance As Variant
    Dim tracker As femap.TrackData
    Set boundary_nodes = App.feSet
    Set e = App.feElem
    Set eSet = App.feSet
    Set n = App.feNode
    Set nSet = App.feSet
    Set p = App.feProp
    Set pSet = App.feSet
    Set s = App.feSurface
    Set tracker = App.feTrackData

    '===========================================================================
    ' Entity Selection Prompts
    '===========================================================================

    ' prompt to select elements
    rc = eSet.Select(FT_ELEM, True, "Select elements to remesh...")
    If rc = FE_OK and eSet.Count() > 1 Then
        ' remove non-planar elements
        removed = False
        For i = 1 To 44
            Select Case i
                Case 11 To 22, 32, 35, 36  'do nothing if element is a shell element
                Case Else
                    eSet.RemoveRule(i, FGD_ELEM_BYTYPE)
            End Select
        Next i

    ElseIf rc = FE_OK and eSet.Count() <= 1 Then
        App.feAppMessage(FCM_ERROR, "Must select more than one element.")
        Exit Sub
    ElseIf rc =FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "One or more selected elements do not exist.")
        Exit Sub
    Else
        Exit Sub
    End If

    ' create remesh node set and check for surface association
    e.Get(eSet.First())
    sid = e.geomID
    pick_geom = False
    eSet.Reset()
    Do While e.NextInSet(eSet.ID)
        'create remesh node set
        nSet.AddRule(e.ID, FGD_NODE_ONELEM)

        'check for surface association
        e.Get(e.ID)
        If sid <> e.geomID Then
            pick_geom = True
        End If
    Loop

    ' create set with only boundary nodes
    App.feElementFreeEdge(eSet.ID, True, False, nFreeCount, nFreeData)
    i = 0
    Do While i < (nFreeCount * 3)
        If i Mod 3 <> 0 Then
            boundary_nodes.Add(nFreeData(i))
        End If
        i = i + 1
    Loop
    nSet.RemoveSet(boundary_nodes.ID)

    ' prompt to select surface
    rc = s.SelectID("Select surface...")
    If rc = FE_OK Then
        s.Get(s.ID)
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected surface does not exist.")
        Exit Sub
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    '===========================================================================
    ' Dialog Box Prompts
    '===========================================================================

    ' populate prop_titles() and prop_ids()
    pSet.AddAll(FT_PROP)
    pSet.Reset()

    If s.attrPID <> 0 Then
        Redim prop_ids(2)
        Redim prop_titles(2)
        prop_ids(0) = -1
        prop_ids(1) = 0
        prop_titles(0) = "-1..Plot Only"
        prop_titles(1) = "0..Surface Property"
        i = 3
    Else
        Redim prop_ids(1)
        Redim prop_titles(1)
        prop_ids(0) = 0
        prop_titles(0) = "0..Plot Only"
        i = 2
    End If

    Do While p.NextInSet(pSet.ID)
        p.Get(p.ID)
        Select Case p.type
            Case 11 To 22  'shell property types
                i += 1
                Redim Preserve prop_ids(i)
                Redim Preserve prop_titles(i)
                prop_ids(i-1) = p.id
                prop_titles(i-1) = CStr(p.id) & ".." & p.title
            Case Else
                ' do nothing
        End Select
    Loop

    ' user dialog
    Begin Dialog UserDialog 670,266,"Refine/Remesh on Surface Options",.DialogFunc ' %GRID:10,7,1,1
        ' remeshing options
        GroupBox 10,7,340,84,"Remeshing Options",.remesh_option_groupbox
        OptionGroup .remesh_mode
            OptionButton 30,21,90,21,"Refine",.refine
            OptionButton 30,42,90,21,"Remesh",.remesh
            OptionButton 30,63,90,21,"Unrefine",.unrefine
        text 120,28,130,14,"Refinement Ratio",.refinement_ratio_text1,1
        text 320,28,20,14,":1",.refinement_ratio_text2
        TextBox 260,28,50,14,.refinement_ratio
        text 160,49,90,14,"Break Angle",.break_angle_text,1
        TextBox 260,49,50,14,.break_angle

        ' mesh smoothing
        GroupBox 360,7,300,84,"Mesh Smoothing",.smoothing_groupbox
        OptionGroup .smooth_type
            OptionButton 380,28,90,14,"Laplacian",.laplacian
            OptionButton 380,49,90,14,"Centroidal",.centroidal
        text 500,28,70,14,"Iterations",.iterations_text,1
        TextBox 580,28,60,14,.iterations
        text 500,49,70,14,"Tolerance",.tolerance_text,1
        TextBox 580,49,60,14,.smooth_tol

        ' element shape
        GroupBox 10,91,260,84,"Element Shape",.elem_shape_groupbox
        OptionGroup .elem_shape
            OptionButton 30,105,90,21,"All Triangles",.tri
            OptionButton 30,126,230,21,"Quad (when internal angles are",.quad
        text 50,147,40,14,"within",.quad_angle_text1
        TextBox 90,147,50,14,.quad_angle
        text 150,147,110,14,"of 90 degrees)",.quad_angle_text2
        CheckBox 140,105,110,21,"Fast Tri Mesh",.fast_tri

        ' mesh control
        GroupBox 280,91,380,84,"Mesh Control",.mesh_control_groupbox
        text 310,112,220,14,"Min Elements Between Boundaries",.min_between_text,1
        TextBox 540,112,50,14,.min_between
        text 400,133,130,14,"Max Aspect Ratio",.aspect_ratio_text1,1
        TextBox 540,133,50,14,.aspect_ratio
        text 600,133,20,14,":1",.aspect_ratio_text2
        CheckBox 300,154,230,14,"Quick-cut boundaries more than",.quick_cut
        TextBox 540,154,50,14,.quick_nodes
        text 600,154,50,14,"nodes.",.quick_node_text

        ' node and element options
        GroupBox 10,175,380,84,"Node and Element Options",.node_elem_options_groupbox
        text 20,196,60,14,"Node ID",.start_node_text,1
        TextBox 90,196,100,14,.start_node
        text 210,196,50,14,"Elem ID",.start_elem_text,1
        TextBox 270,196,100,14,.start_elem
        CheckBox 400,175,260,21,"Delete Original Nodes and Elements",.delete_orig
        text 20,231,60,14,"Property",.prop_id_text,1
        DropListBox 90,224,290,21,prop_titles(),.prop_id

        ' no group
        PushButton 400,217,260,21,"Exclude Nodes from Boundary >>",.exclude_nodes
        CheckBox 400,196,210,14,"Associate entities to surface",.associate
        OKButton 400,238,120,21
        CancelButton 530,238,130,21
    End Dialog
    Dim dlg As UserDialog

    ' prompt user to select smoothing options
    If Dialog(dlg) = 0 Then
        Exit Sub  ' exit if user hits 'Cancel'
    End If

    '===========================================================================
    ' User Input Processing
    '===========================================================================

    ' set property
    Select Case prop_ids(dlg.prop_id)
        Case 0
            p.ID = s.attrPID
        Case Else
            p.ID = prop_ids(dlg.prop_id)
    End Select

    ' set mesh_approach
    Select Case dlg.elem_shape
        Case 0  'tri
            If dlg.fast_tri Then
                mesh_approach = 2
            Else
                mesh_approach = 0
            End If
        Case 1  'quad
            mesh_approach = 1
    End Select

    '===========================================================================
    ' Remeshing
    '===========================================================================

    tracker.StartMesh()
    App.feMeshRemesh(eSet.ID, _
                     nSet.ID, _
                     dlg.remesh_mode, _
                     dlg.refinement_ratio, _
                     dlg.break_angle, _
                     dlg.delete_orig, _
                     p.ID, _
                     dlg.min_between, _
                     dlg.aspect_ratio, _
                     dlg.quick_nodes, _
                     mesh_approach, _
                     dlg.quad_angle, _
                     dlg.smooth_type, _
                     dlg.iterations, _
                     dlg.smooth_tol)
    tracker.Created(FT_NODE, nSet.ID, False)
    tracker.Created(FT_ELEM, eSet.ID, False)

    '===========================================================================
    ' Cleanup
    '===========================================================================

    ' move nodes to surface and associate
    App.feProjectOntoSurface(FT_NODE, nSet.ID, s.ID)
    App.feMeshAttach(FT_ELEM, eSet.ID, FT_SURFACE, s.ID, True)

    App.feViewRegenerate(0)

End Sub


Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Set App = feFemap()
    Set e = App.feElem
    Set n = App.feNode

    Select Case Action%
        Case 1  ' dialog box initialization
            ' starting values
            DlgText "refinement_ratio", "2"
            DlgText "break_angle", "45"
            DlgText "iterations", "1000"
            DlgText "smooth_tol", "0.001"
            DlgText "quad_angle", "60"
            DlgText "min_between", "1"
            DlgText "aspect_ratio", "2"
            DlgText "quick_nodes", "300"
            DlgText "start_node", CStr(n.NextEmptyID)
            DlgText "start_elem", CStr(e.NextEmptyID)
            DlgValue "delete_orig", True
            DlgValue "smooth_type", 0
            DlgValue "elem_shape", 1
            DlgValue "associate", True
            If s.attrPID = 0 Then
                DlgValue "prop_id", 0
            Else
                DlgValue "prop_id", 1
            End If

            ' set usable starting fields
            DlgEnable "break_angle_text", False
            DlgEnable "break_angle", False

            DlgEnable "elem_shape", False
            DlgEnable "tri", False
            DlgEnable "quad", False
            DlgEnable "quad_angle_text1", False
            DlgEnable "quad_angle", False
            DlgEnable "quad_angle_text2", False
            DlgEnable "fast_tri", False

            DlgEnable "min_between_text", False
            DlgEnable "min_between", False
            DlgEnable "aspect_ratio_text1", False
            DlgEnable "aspect_ratio", False
            DlgEnable "aspect_ratio_text2", False
            DlgEnable "quick_cut", False
            DlgEnable "quick_nodes", False
            DlgEnable "quick_node_text", False

            DlgEnable "start_node_text", False
            DlgEnable "start_node", False
            DlgEnable "start_elem_text", False
            DlgEnable "start_elem", False

            DlgEnable "exclude_nodes", False

        Case 2  ' value changing or button pressed
            Select Case DlgItem$
                Case "remesh_mode"
                    Select Case SuppValue?
                        Case 0  'refine
                            ' set usable fields
                            DlgEnable "refinement_ratio_text1", True
                            DlgEnable "refinement_ratio_text2", True
                            DlgEnable "refinement_ratio", True
                            DlgEnable "break_angle_text", False
                            DlgEnable "break_angle", False

                            DlgEnable "elem_shape", False
                            DlgEnable "tri", False
                            DlgEnable "quad", False
                            DlgEnable "quad_angle_text1", False
                            DlgEnable "quad_angle", False
                            DlgEnable "quad_angle_text2", False
                            DlgEnable "fast_tri", False

                            DlgEnable "min_between_text", False
                            DlgEnable "min_between", False
                            DlgEnable "aspect_ratio_text1", False
                            DlgEnable "aspect_ratio", False
                            DlgEnable "aspect_ratio_text2", False
                            DlgEnable "quick_cut", False
                            DlgEnable "quick_nodes", False
                            DlgEnable "quick_node_text", False

                            DlgEnable "start_node_text", False
                            DlgEnable "start_node", False
                            DlgEnable "start_elem_text", False
                            DlgEnable "start_elem", False

                            DlgEnable "exclude_nodes", False

                        Case 1  'remesh
                            '  set usable fields
                            DlgEnable "refinement_ratio_text1", False
                            DlgEnable "refinement_ratio_text2", False
                            DlgEnable "refinement_ratio", False
                            DlgEnable "break_angle_text", False
                            DlgEnable "break_angle", False

                            DlgEnable "elem_shape", True
                            DlgEnable "tri", True
                            DlgEnable "quad", True
                            DlgEnable "quad_angle_text1", True
                            DlgEnable "quad_angle", True
                            DlgEnable "quad_angle_text2", True
                            DlgEnable "fast_tri", False

                            DlgEnable "min_between_text", True
                            DlgEnable "min_between", True
                            DlgEnable "aspect_ratio_text1", True
                            DlgEnable "aspect_ratio", True
                            DlgEnable "aspect_ratio_text2", True
                            DlgEnable "quick_cut", True
                            DlgEnable "quick_nodes", True
                            DlgEnable "quick_node_text", True

                            DlgEnable "start_node_text", True
                            DlgEnable "start_node", True
                            DlgEnable "start_elem_text", True
                            DlgEnable "start_elem", True

                            DlgEnable "exclude_nodes", True

                        Case 2  'unrefine
                            ' set usable fields
                            DlgEnable "refinement_ratio_text1", True
                            DlgEnable "refinement_ratio_text2", True
                            DlgEnable "refinement_ratio", True
                            DlgEnable "break_angle_text", True
                            DlgEnable "break_angle", True

                            DlgEnable "elem_shape", False
                            DlgEnable "tri", False
                            DlgEnable "quad", False
                            DlgEnable "quad_angle_text1", False
                            DlgEnable "quad_angle", False
                            DlgEnable "quad_angle_text2", False
                            DlgEnable "fast_tri", False

                            DlgEnable "min_between_text", False
                            DlgEnable "min_between", False
                            DlgEnable "aspect_ratio_text1", False
                            DlgEnable "aspect_ratio", False
                            DlgEnable "aspect_ratio_text2", False
                            DlgEnable "quick_cut", False
                            DlgEnable "quick_nodes", False
                            DlgEnable "quick_node_text", False

                            DlgEnable "start_node_text", False
                            DlgEnable "start_node", False
                            DlgEnable "start_elem_text", False
                            DlgEnable "start_elem", False

                            DlgEnable "exclude_nodes", False
                    End Select
                    DialogFunc = True  'do not exit the dialog

                Case "elem_shape"
                    Select Case SuppValue?
                        Case 0  'tri
                            DlgEnable "quad_angle", False
                            DlgEnable "fast_tri", True
                        Case 1  'quad
                            DlgEnable "fast_tri", False
                            DlgEnable "quad_angle", True
                    End Select
                    DialogFunc = True  'do not exit the dialog

                Case "fast_tri"
                    If SuppValue? Then
                        DlgEnable "mesh_control_groupbox", False
                        DlgEnable "min_between_text", False
                        DlgEnable "min_between", False
                        DlgEnable "aspect_ratio_text1", False
                        DlgEnable "aspect_ratio", False
                        DlgEnable "aspect_ratio_text2", False
                        DlgEnable "quick_cut", False
                        DlgEnable "quick_nodes", False
                        DlgEnable "quick_node_text", False
                    Else
                        DlgEnable "mesh_control_groupbox", True
                        DlgEnable "min_between_text", True
                        DlgEnable "min_between", True
                        DlgEnable "aspect_ratio_text1", True
                        DlgEnable "aspect_ratio", True
                        DlgEnable "aspect_ratio_text2", True
                        DlgEnable "quick_cut", True
                        DlgEnable "quick_nodes", True
                        DlgEnable "quick_node_text", True
                    End If

                Case "exclude_nodes"
                    boundary_nodes.Select(FT_NODE, False, "Nodes to exclude from remesh...")
                    ' recreate interior node set
                    eSet.Reset()
                    Do While e.NextInSet(eSet.ID)
                        nSet.AddRule(e.ID, FGD_NODE_ONELEM)
                    Loop
                    nSet.RemoveSet(boundary_nodes.ID)
                    DialogFunc = True  'do not exit the dialog

            End Select
    End Select
End Function