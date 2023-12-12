' Title: Pull Femap LCIDs.vbs
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 12.0

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

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