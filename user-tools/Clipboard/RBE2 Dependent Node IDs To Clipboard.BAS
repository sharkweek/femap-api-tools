' Title: RBE2 Dependent Node IDs To Clipboard.BAS
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

    App.feAppMessage(FCM_COMMAND, "RBE2 Dependent Node IDs to Clipboard")

    Dim rc As Long

    Dim rigidSet As femap.Set
    Set rigidSet = App.feSet

    Dim iNodeSet As femap.Set
    Set iNodeSet = App.feSet

    Dim dNodeSet As femap.Set
    Set dNodeSet = App.feSet

    Dim elm As femap.Elem
    Set elm = App.feElem

    rc = rigidSet.Select(FT_ELEM, True, "Select elements...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "None of the selected elements exist...")
        Exit Sub
    ElseIf rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "No elements selected")
        Exit Sub
    End If

    rigidSet.Reset()
    Do While elm.NextInSet(rigidSet.ID)
        ' Remove non-rigid elements from selection set
        If elm.type <> 29 Then
            rigidSet.Remove(elm.ID)
        ' Remove RBE3 elements from selection set
        ElseIf elm.RigidInterpolate Then
            rigidSet.Remove(elm.ID)
        ' Add independent nodes of RBE2s to new set
        Else
            iNodeSet.Add(elm.node(0))
        End If
    Loop

    'add all nodes from rigidSet to dNodeSet
    dNodeSet.AddSetRule(rigidSet.ID, FGD_Node_onElem)
    dNodeSet.RemoveSet(iNodeSet.ID)

    ' Add independent nodes to Selector
    dNodeSet.CopyToClipboard(False)
    App.feAppMessage(FCM_NORMAL, Cstr(dNodeSet.Count()) + " Nodes copied to clipboard...")

End Sub
