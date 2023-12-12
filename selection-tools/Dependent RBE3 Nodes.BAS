' Title: Dependent RBE3 Nodes.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2301

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' public declarations
Dim App As femap.model

Sub Main
    Set App = feFemap()

    Dim rc As Long

    Dim etype As femap.zDataType

    Dim rigidSet As femap.Set
    Set rigidSet = App.feSet

    Dim nodeSet As femap.Set
    Set nodeSet = App.feSet

    Dim elm As femap.Elem
    Set elm = App.feElem

    ' start selection tool
    App.feStartSelectionTool(0, etype)
    If etype <> FT_NODE Then
        App.feAbortSelectionTool()
    End If

    rc = rigidSet.Select(FT_ELEM, True, "Select elements...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "None of the selected elements exist...")
        Exit Sub
    ElseIf rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "No elements selected")
        Exit Sub
    Else
        ' remove non-rigid elements from selection
        rigidSet.RemoveRule(29, FGC_ELEM_BYTYPE)
    End If

    ' remove non-rigid elements
    rigidSet.Reset()
    Do While elm.NextInSet(rigidSet.ID)
        If elm.RigidInterpolate Then
            nodeSet.Add(elm.node(0))
        End If
    Loop

    ' add nodes
    App.feEndSelectionTool(nodeSet.ID, 0)

End Sub
