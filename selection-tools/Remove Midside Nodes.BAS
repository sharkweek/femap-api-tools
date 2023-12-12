' Title: Remove Midside Nodes.BAS
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

    Dim elm_set As femap.Set
    Set elm_set = App.feSet()

    Dim mid_node_set As femap.Set
    Set mid_node_set = App.feSet()

    Dim node_set As femap.Set
    Set node_set = App.feSet()

    ' start selection tool
    ' etype = FT_NODE
    App.feStartSelectionTool(node_set.ID, etype)
    If etype <> FT_NODE Then
        App.feAbortSelectionTool()
    End If

    ' ' clear out midside nodes
    elm_set.AddSetRule(node_set.ID, FGD_ELEM_BYNODE)
    mid_node_set.AddMidsideNodes(elm_set.ID)

    ' return filtered nodes
    App.feEndSelectionTool(0, mid_node_set.ID)

End Sub
