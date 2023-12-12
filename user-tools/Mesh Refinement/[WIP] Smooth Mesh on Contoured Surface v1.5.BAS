' Title: Smooth Mesh on Contoured Surface v1.5.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 12.0

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.


' dimension as public for DialogFunc to access
Dim App As femap.model
Dim fixed_nodes As femap.Set
Dim tolerance As Variant
Dim iterations As Variant

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Smooth Mesh on Contoured Surface")

'   Declarations
'   ------------
    Dim rc As Long

    Dim e As femap.Elem
    Set e = App.feElem

    Dim eSet As femap.Set
    Set eSet = App.feSet

    Dim n As femap.Node
    Set n = App.feNode

    Dim propSet As femap.Set
    Set propSet = App.feSet

    Set fixed_nodes = App.feSet

    Dim smoothed_nodes As femap.Set
    Set smoothed_nodes = App.feSet

    Dim s As femap.Surface
    Set s = App.feSurface

    Dim smooth_s As femap.Surface
    Set smooth_s = App.feSurface

    Dim i As Long
    Dim nFreeCount As Long
    Dim nFreeData As Variant
    Dim removed As Boolean

    Dim tracker As femap.TrackData
    Set tracker = App.feTrackData
    Dim avgLength As Double
    Dim faceArea As Double
    Dim totalLength As Double

    Dim sSet As femap.Set
    Set sSet = App.feSet

    Dim cSet As femap.Set
    Set cSet = App.feSet

    Dim pSet As femap.Set
    Set pSet = App.feSet

    Dim n_normals As femap.SortSet
    Set n_normals = App.feSort

    Dim u As Double
    Dim v As Double

    Dim norm_num As Long
    Dim norm_og As Variant
    Dim norm_new As Variant
    Dim sum_vec_1 As Variant
    Dim sum_vec_2 As Variant
    Dim proj_vec(3) As Double
    Dim vec_length_og As Double
    Dim vec_length_new As Double
    Dim og_coord(3) As Double
    Dim x As Double
    Dim y As Double
    Dim z As Double

'   Entity Selection
'   ----------------
    ' prompt to select elements
    rc = eSet.Select(FT_ELEM, True, "Select elements to smooth...")
    If rc = FE_OK And eSet.Count() > 1 Then
        ' remove non-planar elements
        removed = False
        Do While e.NextInSet(eSet.ID)
            Select Case e.type
                Case 10 To 22  'do nothing if element is a shell element
                Case Else
                    propSet.RemoveRule(e.type, FGD_ELEM_BYTYPE)
                    removed = True
            End Select
        Loop

        If removed Then
            App.feAppMessage(FCM_HIGHLIGHT, "Non-planar elements removed from selection set...")
        End If

    ElseIf rc = FE_OK And eSet.Count() <= 1 Then
        App.feAppMessage(FCM_ERROR, "Must select more than one element.")
        Exit Sub

    ElseIf rc =FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "One or more selected elements do not exist.")
        Exit Sub

    Else
        Exit Sub

    End If

    ' prompt to select surface
    rc = s.SelectID("Select surface...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected surface does not exist.")
        Exit Sub
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    ' get boundary nodes to fix during smoothing
    App.feElementFreeEdge(eSet.ID, True, False, nFreeCount, nFreeData)
    i = 0
    Do While i < (nFreeCount * 3)
        If i Mod 3 <> 0 Then  ' add node IDs to fixed_nodes
            fixed_nodes.Add(nFreeData(i))
        End If
        i = i + 1
    Loop

'   Dialog prompt
'   -------------
    ' set initial smoothing values
    tolerance = 0.001
    iterations = 1000

    ' user dialog for smoothing options
    Begin Dialog UserDialog 1050,845,470,91,"Smoothing Options",.DialogFunc
        GroupBox 20,7,130,77,"Method",.method_group_box
        OptionGroup .method
            OptionButton 40,49,90,14,"Centroidal",.centroidal
            OptionButton 40,28,100,14,"Laplacian",.laplacian
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

    ' prompt user to select smoothing options
    If Dialog(dlg) = 0 Then
    	Exit Sub  ' exit if user hits 'Cancel'
    End If

    ' get internal, non-fixed nodes
    rc = smoothed_nodes.AddSetRule(eSet.ID, FGD_NODE_ONELEM)
    rc = smoothed_nodes.RemoveSet(fixed_nodes.ID)  ' remove the fixed nodes from set

    ' store original node coordinates
    Dim node_coords As femap.SortSet
    Set node_coords = App.feSort

    'smoothed_nodes.Reset()
    Do While n.NextInSet(smoothed_nodes.ID)
        n.Get(n.ID)
        node_coords.AddReal(n.ID, n.x, n.y, n.z)
    Loop

'   Mesh Smoothing
'   --------------
    rc = App.feMeshSmooth2(eSet.ID, fixed_nodes.ID, dlg.method, iterations, tolerance)
    If rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "Unable to smooth selected elements.")
        Exit Sub
    End If

'   Create Surface from smoothed mesh
'   ---------------------------------
    ' find average element edge length
    totalLength = 0
    eSet.Reset()
    Do While e.NextInSet(eSet.ID)
        e.Get(e.ID)
        e.GetFaceArea(1, faceArea)
        totalLength += Sqr(faceArea)
    Loop

    avgLength = totalLength / eSet.Count()
    tracker.StartGeometry()
    App.feSurfaceFromMesh2(eSet.ID, False, False, avgLength/100, False, avgLength*1.5, 1.0)
    tracker.Created(FT_SURFACE, sSet.ID, False)
    tracker.Created(FT_CURVE, cSet.ID, False)
    tracker.Created(FT_POINT, pSet.ID, False)

    smooth_s.Get(sSet.First())

    ' project nodes
    i = 0
    Do While i <= 2  ' three iterations is sufficient for most cases
        node_coords.Reset()
        Do While node_coords.Next()
            n.Get(node_coords.Current)
            ' find vector bisecting the normals of the two surfaces for projection
            node_coords.GetReal(n.ID, x, y, z)
            og_coord(0) = x
            og_coord(1) = y
            og_coord(2) = z
            App.feVectorNormalToSurface(s.ID, og_coord, norm_og)
            App.feVectorNormalToSurface(smooth_s.ID, n.xyz, norm_new)
            App.feVectorLength(norm_og, vec_length_og)
            App.feVectorLength(norm_new, vec_length_new)
            For i = 0 To 2
                proj_vec(i) = vec_length_new * norm_og(i) + vec_length_og * norm_new(i)
            Next i
            App.feProjectAlongVector(FT_NODE, -n.ID, -s.ID, proj_vec)
        Loop

        i += 1
    Loop

    ' delete created geometry
    App.feDelete(FT_SURFACE, sSet.ID)
    App.feDelete(FT_CURVE, cSet.ID)
    App.feDelete(FT_POINT, pSet.ID)

    ' Choose whether or not to attach mesh to surface
    rc = App.feAppMessageBox(2, "Associate elements and nodes with surface?")
    If rc = FE_OK Then
        App.feMeshAttach(FT_ELEM, eSet.ID, FT_SURFACE, s.ID, True)
    End If

    App.feViewRegenerate(0)

End Sub


Function DialogFunc(DlgItem As String, Action As Integer, SuppValue As Long) As Boolean
    Select Case Action
        Case 1  ' dialog box initialization
            Beep
            DlgText "tolerance", CStr(tolerance)
            DlgText "iterations", CStr(iterations)
            DlgValue "method", 1

        Case 2  ' value changing or button pressed
            If DlgItem = "fix_nodes" Then
                fixed_nodes.Select(FT_NODE, False, "Select fixed nodes...")
                DialogFunc = True  ' do not exit dialog
            End If

        Case 3  ' TextBox or ComboBox text changed
            'validate input values
            Select Case DlgItem
                Case "iterations"
                    ' validate iterations
                    If IsNumeric(DlgText(DlgItem)) Then
                        iterations = CDbl(DlgText(DlgItem))

                        ' throw error if negative or non-integer
                        If iterations <= 0 Or (CLng(iterations) - iterations = 0) Then
                            App.feAppMessageBox(0, "ERROR: Iterations must be a positive, non-zero integer.")
                            DlgFocus(6)
                        End If

                    Else  ' throw error if non-numeric value
                        App.feAppMessageBox(0, "ERROR: Iterations must be a positive, non-zero integer.")
                        DlgFocus(6)

                    End If

                    DialogFunc = True

                Case "tolerance"
                    'validate tolerance
                    If IsNumeric(DlgText(DlgItem)) Then
                        tolerance = CDbl(DlgText(DlgItem))

                        If tolerance <= 0 Then  ' throw error if negative
                            App.feAppMessageBox(0, "ERROR: Tolerance must be a positive, non-zero integer.")
                            DlgFocus(8)
                        End If

                    Else  ' throw error if non-numeric value
                        App.feAppMessageBox(0, "ERROR: Tolerance must be a positive, non-zero integer.")
                        DlgFocus(8)

                    End If

                    DialogFunc = True

            End Select
    End Select
End Function
