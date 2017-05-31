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
    Else If crv1.center(c1) = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "Curve is non-circular or does not exist")
    End If

    ' Select Curve 2
    If crv2.SelectID("Pick Curve 2") = FE_CANCEL Then
        Exit Sub
    Else If crv2.center(c2) = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "Curve is non-circular or does not exist")
    End If

    ' Draw line between centers
    App.feLinePoints(False, c1, c2, True)

End Sub
