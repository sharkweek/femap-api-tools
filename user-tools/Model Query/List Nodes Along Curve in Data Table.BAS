' Title: List Nodes Along Curve in Data Table.BAS
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

    App.feAppMessage(FCM_COMMAND, "List Nodes Along Curve in Data Table")

    Dim rc As Long

    Dim dTable As femap.DataTable
    Set dTable = App.feDataTable

    Dim crv As femap.Curve
    Set crv = App.feCurve

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim nArray() As Long
    Dim sArray() As Variant

    Dim n As femap.Node
    Set n = App.feNode

    Dim i As Long
    Dim s As Double
    Dim newColID As Long

    ' prompt to clear data table
    rc = App.feAppMessageBox(2, "OK to clear Data Table?")
    If rc <> FE_OK Then
        Exit Sub
    End If

    ' prep data table
    App.feAppManagePanes("Data Table", 1)
    dTable.Lock(False)
    dTable.Clear()

    ' pick curve
    rc = crv.SelectID("Select curve...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected curve does not exist.")
        Exit Sub
    End If

    ' get nodes on curve (not including ends)
    crv.NodesAsSet(True, True, nSet.ID, True)
    If nSet.Count() = 0 Then
        App.feAppMessage(FCM_ERROR, "No nodes associated with curve.")
        Exit Sub
    End If

    ' get xyz-coords as parametric values
    ReDim nArray(nSet.Count()) As Long
    ReDim sArray(nSet.Count()) As Variant
    i = 0
    Do While n.NextInSet(nSet.ID)
        crv.XYZToParam(n.xyz, s)
        nArray(i) = n.ID
        sArray(i) = s * crv.Length()
        i = i + 1
    Loop

    ' add node data to data table and sort
    dTable.AddEntityData(True, FT_NODE, 0, nSet.ID)
    dTable.AddColumn(False, False, FT_NODE, 0, "Distance along Curve " & CStr(crv.ID), 2, nSet.Count(), nArray, sArray, newColID)
    dTable.SetColumnPosition(newColID, 2, True)
    dTable.AddSort(newColID, True)

    App.feAppMessage(FCM_NORMAL, CStr(nSet.Count()) & " nodes added to Data Table...")

End Sub
