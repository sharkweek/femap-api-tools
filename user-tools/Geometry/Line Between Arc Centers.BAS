' Title: Line Between Arc Centers.BAS
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

    Dim crv1 As femap.Curve
    Set crv1 = App.feCurve

    Dim crv2 As femap.Curve
    Set crv2 = App.feCurve

    Dim c1 As Variant
    Dim c2 As Variant

    Dim rc As Long

    ' Select Curve 1
    If crv1.SelectID("Pick Curve 1") = FE_CANCEL Then
        Exit Sub
    ElseIf crv1.center(c1) = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "Curve is non-circular or does not exist")
    End If

    ' Select Curve 2
    If crv2.SelectID("Pick Curve 2") = FE_CANCEL Then
        Exit Sub
    ElseIf crv2.center(c2) = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "Curve is non-circular or does not exist")
    End If

    ' Draw line between centers
    App.feLinePoints(False, c1, c2, True)

End Sub
