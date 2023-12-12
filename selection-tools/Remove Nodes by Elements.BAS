' Title: Remove Nodes by Elements.BAS
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

    Dim elm_set As femap.Set
    Set elm_set = App.feSet

    Dim node_set As femap.Set
    Set node_set = App.feSet()

    Dim elm As femap.Elem
    Set elm = App.feElem

    ' start selection tool
    App.feStartSelectionTool(0, etype)
    If etype <> FT_NODE Then
        App.feAbortSelectionTool()
    End If

    rc = elm_set.Select(FT_ELEM, True, "Select elements...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "None of the selected elements exist...")
        Exit Sub
    ElseIf rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "No elements selected")
        Exit Sub
    End If

    node_set.AddSetRule(elm_set.ID, FGD_NODE_ONELEM)

    ' add nodes
    App.feEndSelectionTool(0, node_set.ID)

End Sub
