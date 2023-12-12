' Title: Dependent RBE2 Nodes.BAS
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

    Dim etype As femap.zDataType

    Dim rc As Long

    Dim rigid_set As femap.Set
    Set rigid_set = App.feSet

    Dim node_set As femap.Set
    Set node_set = App.feSet()

    Dim iNodeSet As femap.Set
    Set iNodeSet = App.feSet

    Dim dNodeSet As femap.Set
    Set dNodeSet = App.feSet

    Dim elm As femap.Elem
    Set elm = App.feElem

    ' start selection tool
    App.feStartSelectionTool(0, etype)
    If etype <> FT_NODE Then
        App.feAbortSelectionTool()
    End If

    rc = rigid_set.Select(FT_ELEM, True, "Select rigid elements...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "None of the selected elements exist...")
        Exit Sub
    ElseIf rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "No elements selected")
        Exit Sub
    Else
        ' remove non-rigid elements from set
        rigid_set.RemoveRule(29, FGD_ELEM_BYTYPE)
    End If

    rigid_set.Reset()
    Do While rigid_set.Next()
        elm.Get(rigid_set.CurrentID)
        If Not elm.RigidInterpolate then  ' check for RBE2
            iNodeSet.Add(elm.node(0))
        End If
    Loop

    'add all nodes from rigid_set to dNodeSet
    dNodeSet.AddSetRule(rigid_set.ID, FGD_Node_onElem)
    dNodeSet.RemoveSet(iNodeSet.ID)

    ' add nodes
    App.feEndSelectionTool(dNodeSet.ID, 0)

End Sub
