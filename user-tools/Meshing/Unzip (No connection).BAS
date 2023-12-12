' Title: Unzip (No connection).BAS
' Author: Andy Perez
' License: OSL-3.0
' Date: December 2023
' Femap API Version: 2306

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' public declarations
Dim App As femap.model

Sub Main
    ' `Main` declarations
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Unzip (No connection)")

    Dim rc As Long

    Dim elm_set As femap.Set
    Set elm_set = App.feSet()

    Dim node_set As femap.Set
    Set node_set = App.feSet()

    Dim dummy_vec(2) As Double

    rc = FE_OK
    Do While rc = FE_OK
            ' prompt user to select boundary elements
            rc = elm_set.Select(FT_ELEM, True, "Select boundary elements...")
            If rc = FE_CANCEL Then
                Exit Sub
            ElseIf rc = FE_NOT_EXIST Then
                App.feAppMessage(FCM_ERROR, "No elements selected. Routine canceled...")
            End If

            ' prompt user to select nodes along boundary
            rc = node_set.Select(FT_NODE, True, "Select nodes along boundary...")
            If rc = FE_CANCEL Then
                Exit Sub
            ElseIf rc = FE_NOT_EXIST Then
                App.feAppMessage(FCM_ERROR, "No nodes selected. Routine canceled...")
            End If

            rc = App.feMeshUnzip2(elm_set.ID, node_set.ID, 0, 2, 0, dummy_vec, 0, False)
            App.feViewRegenerate(0)
    Loop
End Sub