' Title: Independent RBE2 Nodes.BAS
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

    Dim rigid_set As femap.Set
    Set rigid_set = App.feSet

    Dim node_set As femap.Set
    Set node_set = App.feSet

    Dim elm As femap.Elem
    Set elm = App.feElem

    ' start selection tool
    App.feStartSelectionTool(0, etype)
    If etype <> FT_NODE Then
        App.feAbortSelectionTool()
    End If

    rc = rigid_set.Select(FT_ELEM, True, "Select elements...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "None of the selected elements exist...")
        Exit Sub
    ElseIf rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "No elements selected")
        Exit Sub
    Else
        ' remove non-rigid elements from selection
        rigid_set.RemoveRule(29, FGC_ELEM_BYTYPE)
    End If

    ' remove non-rigid elements
    rigid_set.Reset()
    Do While rigid_set.Next()
        elm.Get(rigid_set.CurrentID)
        If Not elm.RigidInterpolate Then  ' check for RBE2
            node_set.Add(elm.node(0))
        End If
    Loop

    ' add nodes
    App.feEndSelectionTool(node_set.ID, 0)

End Sub
