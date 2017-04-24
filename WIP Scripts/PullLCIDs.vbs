Sub PullLCIDs()

    Dim App As femap.Model
    Set App = GetObject(, "femap.model")

    Dim feOutputSet As femap.OutputSet
    Set feOutputSet = App.feOutputSet

    Dim rc As Long

    Dim i As Long
    Dim obj As Variant
    Dim headers() As String
    Dim count As Long
    Dim ID As Variant
    Dim Title As Variant

    ' redimension headers to include all LCs
    ReDim headers(feOutputSet.CountSet() + 1)
    headers(0) = "Combined LCID"
    headers(1) = "Description"

    rc = feOutputSet.GetTitleIDList(False, 0, 0, count, ID, Title)

    i = 2
    feOutputSet.First
    Do While i < feOutputSet.CountSet() + 2
        headers(i) = "LC " + CStr(feOutputSet.ID)
        i = i + 1
        If i <> feOutputSet.CountSet() + 2 Then
            feOutputSet.Next
        End If
    Loop

    For Each obj In headers
        ActiveCell.Value = obj
        ActiveCell.Offset(0, 1).Select
    Next

End Sub
