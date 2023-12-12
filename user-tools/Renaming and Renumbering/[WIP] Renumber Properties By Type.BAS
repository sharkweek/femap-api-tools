' Title: Renumber Properties By Type.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2306

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' Renumber a selected set of groups alphabetically. This script requires that
' group IDs 99999000 and up be unoccupied depending on the number of groups
' selected

' TODO: add functions and boundary constraints

' module scope variables
Dim App As femap.model
Dim obj As Object
Dim select_prompt As String

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Renumber Entities Alphabetically")

    Dim pSet As femap.Set
    Set pSet = App.feSet

    Dim newpSet As femap.Set
    Set newpSet = App.feSet

    Dim dSet As femap.Set
    Set dSet = App.feSet

    Dim listID As Variant
    Dim startID As Long
    Dim types() As Variant

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim n As Long
    Dim swapped As Boolean
    Dim tSwap As Variant
    Dim idSwap As Variant
    Dim iMax As Long

    Dim pSet As femap.Set
    Set pSet = App.feSet

    ' prompt user to select properties
    rc = pSet.Select(FT_PROP, True, "Select properties...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_FAIL Then
        rc = App.feAppMessage(FCM_ERROR, "Selected properties do not exist.")
    End If

    ' Prompt for new starting ID
    rc = App.feGetInt("Starting ID", 1, 99998999, startID)

    ' Get a list of property types for specified group set
    pSet.Reset()
    k = 0
    Do While p.NextInSet(pSet.ID)
        ReDim Preserve types(k)
        ReDim Preserve listID(k)
        p.Get(p.ID)
        types(k) = p.type
        listID(k) = p.ID
        k = k + 1
    Loop

    ' bubble sort
    iMax = pSet.Count() - 2
    Do
        swapped = False
        For i = 0 To iMax
            If types(i) > types(i + 1) Then
                ' swap property types
                tSwap = types(i)
                types(i) = types(i + 1)
                types(i + 1) = tSwap

                ' swap IDs
                idSwap = listID(i)
                listID(i) = listID(i + 1)
                listID(i + 1) = idSwap

                swapped = True
            End If
        Next i
        iMax = iMax - 1
    Loop Until swapped = False

    ' move groups to higher ID block in alphabetic order
    n = 99999000
    newpSet.Clear()
    For i = 0 To pSet.Count() - 1
        p.Get(listID(i))
        dSet.Clear()
        dSet.Add(listID(i))

        App.feRenumber(FT_PROP, dSet.ID, n)
        newpSet.Add(n)
        n = n + 1
    Next i

    ' renumber IDs to selected block in consecutive order
    App.feRenumber(FT_PROP, newpSet.ID, startID)

End Sub
