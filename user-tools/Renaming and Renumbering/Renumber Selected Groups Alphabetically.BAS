' Renumber a selected set of groups alphabetically. This script requires that
' group IDs 99999000 and up be unoccupied depending on the number of groups
' selected

Sub Main
    Dim App As femap.model
    Set App = feFemap()

        App.feAppMessage(FCM_COMMAND, "Renumber Selected Groups Alphabetically")

    Dim g As femap.Group
    Set g = App.feGroup

    Dim groupSet As femap.Set
    Set groupSet = App.feSet

    Dim newGroupSet As femap.Set
    Set newGroupSet = App.feSet

    Dim listID As Variant
    Dim startID As Long
    Dim gtitles() As Variant

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim compVal As Long
    Dim tempID As Long
    Dim title1 As String
    Dim title2 As String
    Dim tempTitle As String

    ' Check that model contains groups
    If App.Info_Count(FT_GROUP) = 0 Then
        App.feAppMessage(FCM_ERROR, "No Groups in Model, Exiting...")
        Exit Sub
    End If

    ' Select groups to alphabetize
    rc = groupSet.SelectMultiID(FT_GROUP,1,"Select groups to alphabetize")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "No groups in selected set exist.")
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    ' Prompt for new starting ID
    rc = App.feGetInt("Starting ID", 1, 99998999, startID )

    ' Get a list of titles for specified group set
    groupSet.Reset()
    k = 0
    Do While g.NextInSet(groupSet.ID)
        ReDim Preserve gtitles(k)
        Redim Preserve listID(k)
        g.Get(groupSet.CurrentID)
        gtitles(k) = g.title
        listID(k) = g.ID
        k = k + 1
    Loop

    'Bubble Sort
    For i = groupSet.Count() - 1 To 0 Step -1
       For j = 1 To i
        title1 = gtitles(j)
        title2 = gtitles(j-1)
        compVal = StrComp(title1, title2, vbTextCompare)
            If compVal = -1 Then
                tempID = listID(j-1)
                tempTitle = gtitles(j-1)
                listID(j-1) = listID(j)
                gtitles(j-1) = gtitles(j)
                listID(j) = tempID
                gtitles(j) = tempTitle
            End If
        Next j
    Next i

    ' move groups to higher ID block in alphabetic order
    n = 0
    For i = 0 To groupSet.Count() - 1
        n = n + 1
        g.Get(listID(i))
        g.Put(99999000 + n)
        newGroupSet.Add(g.ID)
        g.Delete(listID(i))
    Next i

    ' renumber IDs to selected block in consecutive order
    App.feRenumber(FT_GROUP, newGroupSet.ID, startID)

End Sub
