' Title: Copy Element IDs by Nodes.BAS
' Author: Andy Perez
' License: OSL-3.0
' Date: November 2023
' Femap API Version: 2301

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Copy Element IDs by Nodes")

    Dim rc As Long

    Dim nodeSet As femap.Set
    Set nodeSet = App.feSet

    Dim elmSet As femap.Set
    Set elmSet = App.feSet

    rc = nodeSet.Select(FT_NODE, True, "Select nodes...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "None of the selected nodes exist...")
        Exit Sub
    ElseIf rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "No nodes selected")
        Exit Sub
    End If

    elmSet.AddSetRule(nodeSet.ID, FGD_Elem_byNode)

    elmSet.CopyToClipboard(False)
    App.feAppMessage(FCM_NORMAL, Cstr(elmSet.Count()) + " Element IDs copied to clipboard...")

End Sub
