' TODO: add edge smoothing option

' dimension as public for DialogFunc to access
Dim App As femap.model
Dim fixed_nodes as femap.Set

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Smooth Free Surface Mesh")
    App.feAppMessage(FCM_NORMAL, "Smooth a free surface mesh onto its existing surface profile. Works best for meshes with consistent element sizes and")
    App.feAppMessage(FCM_NORMAL, " subtle contours. Use caution when on meshes with complex contours. Meshes with transitioning densities may yield")
    App.feAppMessage(FCM_NORMAL, " inconsistent results.")

    Dim rc As Long

    Dim laplacian As Boolean

    Dim e As femap.Elem
    Set e = App.feElem

    Dim eSet As femap.Set
    Set eSet = App.feSet

    Dim avgLength As Double
    Dim faceArea As Double
    Dim totalLength As Double

    Dim n As femap.Node
    Set n = App.feNode

    Dim propSet As femap.Set
    Set propSet = App.feSet

    Set fixed_nodes = App.feSet

    Dim smoothed_nodes As femap.Set
    Set smoothed_nodes = App.feSet

    Dim fixedSet As femap.Set
    Set fixedSet = App.feSet

    Dim tracker As femap.TrackData
    Set tracker = App.feTrackData

    Dim s As femap.Surface
    Set s = App.feSurface

    Dim sSet As femap.Set
    Set sSet = App.feSet

    Dim cSet As femap.Set
    Set cSet = App.feSet

    Dim pSet As femap.Set
    Set pSet = App.feSet

    Dim tolerance As Variant
    Dim iterations As Variant
    Dim i As Long
    Dim nFreeCount As Long
    Dim nFreeData As Variant
    Dim removed As Boolean

    ' prompt to select elements
    rc = eSet.Select(FT_ELEM, True, "Select elements to smooth...")
    If rc = FE_OK and eSet.Count() > 1 Then
        ' remove non-planar elements
        removed = False
        Do While e.NextInSet(eSet.ID)
            Select Case e.type
                Case 11 To 22  'do nothing if element is a shell element
                Case Else
                    propSet.RemoveRule(e.type, FGD_ELEM_BYTYPE)
                    removed = True
            End Select
        Loop

        If removed Then
            App.feAppMessage(FCM_HIGHLIGHT, "Non-planar elements removed from selection set...")
        End If

    ElseIf rc = FE_OK and eSet.Count() <= 1 Then
        App.feAppMessage(FCM_ERROR, "Must select more than one element.")
        Exit Sub

    ElseIf rc =FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "One or more selected elements do not exist.")
        Exit Sub

    Else
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

    ' user dialog for smoothing options
    Begin Dialog UserDialog 1050,845,470,91,"Smoothing Options",.DialogFunc
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

    ' prompt user to select smoothing options
    If Dialog(dlg) = 0 Then
    	Exit Sub  ' exit if user hits 'Cancel'
    End If

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

    ' smooth mesh
    rc = App.feMeshSmooth2(eSet.ID, fixed_nodes.ID, laplacian, iterations, tolerance)
    If rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "Unable to smooth selected elements.")
        Exit Sub
    End If

    ' get internal, non-fixed nodes
    smoothed_nodes.AddSetRule(eSet.ID, FGD_NODE_ONELEM)
    smoothed_nodes.RemoveSet(fixed_nodes.ID)  ' remove the fixed nodes from set

    ' find average element edge length
    totalLength = 0
    Do While e.NextInSet(eSet.ID)
        e.Get(e.ID)
        e.GetFaceArea(1, faceArea)
        totalLength = totalLength + Sqr(faceArea)
    Loop

    avgLength = totalLength / eSet.Count()

    ' create surface approximation
    ' TODO: prompt user for tolerances
    tracker.StartGeometry()
    App.feSurfaceFromMesh2(eSet.ID, False, False, avgLength/100, False, avgLength*1.5, 1.0)
    tracker.Created(FT_SURFACE, sSet.ID, False)
    tracker.Created(FT_CURVE, cSet.ID, False)
    tracker.Created(FT_POINT, pSet.ID, False)

    s.Get(sSet.First())

    ' project smoothed nodes to surface
    smoothed_nodes.Reset()
    rc = App.feProjectOntoSurface(FT_NODE, smoothed_nodes.ID, s.ID)
    If rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "Nodes could not be moved to surface.")
    End If

    ' delete created geometry
    App.feDelete(FT_SURFACE, sSet.ID)
    App.feDelete(FT_CURVE, cSet.ID)
    App.feDelete(FT_POINT, pSet.ID)

    App.feViewRegenerate(0)

End Sub


Function DialogFunc(DlgItem As String, Action as Integer, SuppValue As Long) As Boolean
    Select Case Action
        Case 1  ' dialog box initialization
            Beep
            DlgText "iterations", "1000"
            DlgText "tolerance", "0.001"
        Case 2  ' value changing or button pressed
            If DlgItem = "fix_nodes" Then
                fixed_nodes.Select(FT_NODE, False, "Select fixed nodes...")
                DialogFunc = True  ' do not exit dialog
            End If
    End Select
End Function
