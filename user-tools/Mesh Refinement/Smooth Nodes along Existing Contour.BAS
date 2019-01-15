' Spaces the nodes evenly along their existing contour. Only possible up to
' 109 nodes. End noes are *NOT* relocated, but are used to calculate the total
' length over which to space out the nodes. Also note that nodes must be
' selected in the order they are to be spaced.

' TODO: Add biasing; l_i / l_(i-1) = B^(1/(n-1)) for the i-th element where
'       `B` is the biasing ratio and `n` is the number of elements
' TODO: have nodes automatically sorted by their distance from the first node
Dim App As femap.model
Dim bias_ratio As Double
Dim bias_option As Integer

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Smooth Nodes along Existing Contour")

    Dim rc As Long

    Dim crv As femap.Curve
    Set crv = App.feCurve

    Dim n As femap.Node
    Set n = App.feNode

    Dim nSort As femap.SortSet
    Set nSort = App.feSort

    Dim cSet As femap.Set
    Set cSet = App.feSet

    Dim pSet As femap.Set
    Set pSet = App.feSet

    Dim tracker As femap.TrackData
    Set tracker = App.feTrackData

    Dim xyz As Variant
    Dim s As Double
    Dim spacing As Double
    Dim i As Long
    Dim B As Double
    Dim L_prime As Double  ' iterator for total length
    Dim l1 As Double
    Dim li As Double
    Dim d2 As Double
    Dim d3 As Double
    Dim sn As Double
    Dim L As Double
    Dim n_elm As Long
    Dim s0 As Double

    ' select nodes
    rc = nSort.Select(FT_NODE, True, "Select nodes...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "One or more selected nodes do not exist.")
        Exit Sub
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    ElseIf nSort.Count() < 4 Then
        App.feAppMessage(FCM_ERROR, "Must select 4 or more nodes...")
        Exit Sub
    End If

    ' create xyz array
    ReDim Preserve xyz(nSort.Count() * 3)
    For i = 1 To nSort.Count()
        nSort.Next()
        n.Get(nSort.Current())
        xyz(i * 3 - 3) = n.x
        xyz(i * 3 - 2) = n.y
        xyz(i * 3 - 1) = n.z
        nSort.SetData(n.ID, n.x, n.y, n.z, i, 0, 0, FT_NODE)
    Next i

    ' create dummy spline
    tracker.StartGeometry()
    rc = App.feSplinePoints(False, False, nSort.Count(), xyz, False)
    If rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "Unable to create dummy spline. Command canceled...")
        Exit Sub
    Else
        tracker.StopAll()
        tracker.Created(FT_CURVE, cSet.ID, False)
        tracker.Created(FT_POINT, pSet.ID, False)
        pSet.AddRule(cSet.First(), FGD_POINT_ONCURVE)
        crv.Get(cSet.First())
    End If

    ' convert xyz-coords to parametric values and sort
    nSort.Reset()
    For i = 1 To nSort.Count()
        nSort.Next()
        n.Get(nSort.Current())
        crv.XYZToParam(n.xyz, s)
        nSort.SetReal(n.ID, s, 0, 0)
    Next i

    nSort.SortRemoveDuplicates(False)

    ' calculate base values
    nSort.First()
    nSort.GetReal(nSort.Current(), s, d2, d3)
    s0 = s
    nSort.Last()
    nSort.GetReal(nSort.Current(), s, d2, d3)
    sn = s
    L = sn - s0
    n_elm = nSort.Count() - 1

    ' spacing options dialog
    bias_ratio = "2"
    bias_option = False

    BIAS_OPTIONS:
    Begin Dialog UserDialog 220,98,"Biasing Options",.DialogFunc ' %GRID:10,7,1,1
        OptionGroup .spacing
            OptionButton 20,7,120,21,"Equal Spacing",.equal
            OptionButton 20,35,90,21,"Biased",.biased
        TextBox 120,35,90,21,.b
        OKButton 20,70,90,21
        CancelButton 120,70,90,21
    End Dialog
    Dim dlg As UserDialog

    ' prompt user to select biasing options
    If Dialog(dlg) = 0 Then
        Exit Sub  ' exit if user hits 'Cancel'
    Else
        bias_ratio = dlg.b
        bias_option = dlg.spacing
    End If

    ' space nodes out evenly
    If dlg.spacing = False Then
        li = L / n_elm

        ' relocate nodes
        nSort.First()
        s = s0 + li
        For i = 1 To n_elm - 1
            nSort.Next()
            n.Get(nSort.Current())
            crv.ParamToXYZ(s, xyz)
            n.xyz = xyz
            n.Put(n.ID)
            s += li
        Next i

    ' biased spacing
    Else
        ' make sure bias ratio is a valid number
        If dlg.b * 1 = 1 Or dlg.b * 1 <= 0 Then
            If App.feAppMessageBox(1, "Error: Bias ratio must be greater than 0 and not equal to 1.") Then
                GoTo BIAS_OPTIONS
            Else
                Exit Sub
            End If
        End If

        B = dlg.b * 1  ' ensure that dlg.b is a number
        l1 = L * (1 - B^(1/(n_elm-1)))/(1-B^(n_elm/(n_elm-1)))

        ' relocate nodes
        nSort.First()
        li = l1
        s = s0 + li
        For i = 1 To n_elm - 1
            nSort.Next()
            n.Get(nSort.Current())
            crv.ParamToXYZ(s, xyz)
            n.xyz = xyz
            n.Put(n.ID)
            li *= B^(1/(n_elm-1))
            s += li
        Next i

    End If

    ' cleanup dummy geometry
    App.feDelete(FT_CURVE, cSet.ID)
    App.feDelete(FT_POINT, pSet.ID)

    App.feViewRegenerate(0)
    App.feAppMessage(FCM_NORMAL, CStr(nSort.Count() - 2) & " nodes moved...")

End Sub


Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Select Case Action%
        Case 1  ' dialog box initialization
            DlgText "b", CStr(bias_ratio)
            DlgValue "spacing", bias_option
            DlgEnable "b", bias_option
        Case 2  ' value changing or button pressed
            Select Case DlgItem$
                Case "spacing"
                    If SuppValue? Then
                        DlgEnable "b", True
                    Else
                        DlgEnable "b", False
                    End If
                    DialogFunc = True  ' prevent from closing dialog
                Case "b"
                    DialogFunc = True  ' prevent from closing the dialog
            End Select
    End Select
End Function