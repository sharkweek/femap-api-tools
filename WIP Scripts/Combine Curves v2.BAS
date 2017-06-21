' Create a combined curve. Useful for creating a custom toolbar button.

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim rc As Long

    Dim cSet As femap.Set
    Set cSet = App.feSet

    Dim cSort As femap.Sort
    Set cSort = App.feSort

    Dim crv As femap.Curve
    Set crv = App.feCurve

    Dim numCurves As Long
    Dim crvList As Variant

    ' Initialize rc
    rc = FE_OK

    ' Loop to continue creating new composite curves
    Do While rc = FE_OK
        ' Select curves
        rc = cSort.Clear()
        rc = cSort.Select(FT_CURVE, True, "Select curves to combine...")
        If rc = FE_NOT_EXIST Then
            App.feAppMessage(FCM_ERROR, "Selected curves do not exist.")
        ElseIf rc = FE_CANCEL Then
            Exit Sub
        End If

        ' Create composite curve
        rc = cSort.GetSet(cSet.ID, True)
        numCurves = cSet.Count()
        cSet.GetArray(numCurves, crvList)
        crv.CombineCurves(cSet.Count(), crvList)
        crv.Put(crv.NextEmptyID)

        App.feViewRegenerate(0)
    Loop

End Sub
