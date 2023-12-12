' Title: Smooth Nodes onto Curve.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 12.0

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' Usage Notes
' -----------
' This script spaces out a selected set of nodes out on a curve. Allows for
' equal spacing or biased. End nodes are not relocated. Also note that the bias
' ratio is the last element length over the first, measured in curve's
' parametric coordinates.

Dim App As femap.model
Dim bias_ratio As Double
Dim bias_option As Integer

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Smooth Nodes onto Curve")
    App.feAppMessage(FCM_NORMAL, "Smooths a set of nodes onto a selected curve equally or using a user-specified bias.")

    Dim rc As Long

    Dim crv As femap.Curve
    Set crv = App.feCurve

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim n As femap.Node
    Set n = App.feNode

    Dim nSort As femap.SortSet
    Set nSort = App.feSort

    Dim xyz As Variant
    Dim s As Double
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
    rc = nSet.Select(FT_NODE, True, "Select nodes...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "One or more selected nodes do not exist.")
        Exit Sub
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    ElseIf nSet.Count() < 4 Then
        App.feAppMessage(FCM_ERROR, "Must select 4 or more nodes...")
        Exit Sub
    End If

    ' pick curve
    rc = crv.SelectID("Select curve...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected curve does not exist.")
        Exit Sub
    End If

    ' convert xyz-coords to parametric values and sort
    nSet.Reset()
    Do While n.NextInSet(nSet.ID)
        n.Get(n.ID)
        crv.XYZToParam(n.xyz, s)
        nSort.AddReal(n.ID, s, 0, 0)  'd1 is the original s, d2 will be the biased s
    Loop
    nSort.SortRemoveDuplicates(True)

    ' calculate base values
    nSort.First()
    nSort.GetReal(nSort.Current(), s, d2, d3)
    s0 = s
    nSort.Last()
    nSort.GetReal(nSort.Current(), s, d2, d3)
    sn = s
    L = sn - s0
    n_elm = nSet.Count() - 1

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

    ' equal spacing
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

    App.feViewRegenerate(0)

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
