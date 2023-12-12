' Title: Copy Node IDs on Elements.BAS
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

    App.feAppMessage(FCM_COMMAND, "Copy Node IDs on Elements")

    Dim rc As Long

    Dim elmSet As femap.Set
    Set elmSet = App.feSet

    Dim nodeSet As femap.Set
    Set nodeSet = App.feSet

    rc = elmSet.Select(FT_ELEM, True, "Select elements...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "None of the selected elements exist...")
        Exit Sub
    ElseIf rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "No elements selected")
        Exit Sub
    End If

    nodeSet.AddSetRule(elmSet.ID, FGD_Node_onElem)

    nodeSet.CopyToClipboard(False)
    App.feAppMessage(FCM_NORMAL, Cstr(nodeSet.Count()) + " Node IDs copied to clipboard...")

End Sub
