' Title: Nodes Around Points to Clipboard.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 12.0

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Nodes Around Points to Clipboard")
    App.feAppMessage(FCM_NORMAL, "Copies IDs of visible nodes around user-selected points to clipboard.")

    Dim rc As Long

    Dim pSet As femap.Set
    Set pSet = App.feSet

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim p As femap.Point
    Set p = App.fePoint

    Dim radius As Double

    ' prompt user to select points
    rc = pSet.Select(FT_POINT, True, "Select points...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_FAIL Then
        rc = App.feAppMessage(FCM_ERROR, "Selected points do not exist.")
    End If

    ' prompt for radius
    If App.feGetReal("Radius", 0, 1e9, radius) = FE_CANCEL Then
        Exit Sub
    End If

    ' get nodes
    Do While p.NextInSet(pSet.ID)
        p.Get(p.ID)
        nSet.AddAroundPoint(FT_NODE, p.xyz, 1, radius, 0)
    Loop

    nSet.CopyToClipboard(False)
    App.feAppMessage(FCM_NORMAL, Cstr(nSet.Count()) + " Node IDs copied to clipboard...")

    App.feViewRegenerate(0)

End Sub