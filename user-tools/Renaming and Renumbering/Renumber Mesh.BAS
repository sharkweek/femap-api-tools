' Title: Renumber Mesh.BAS
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
    ' `Main` declarations
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Renumber Mesh")
    App.feAppMessage(FCM_NORMAL, "Renumber selected mesh entities and, optionally, related entities.")

    Dim rc As Long

    Dim elm_set As femap.Set
    Set elm_set = App.feSet()

    Dim entity_set As femap.Set
    Set entity_set = App.feSet()

    Dim grp As femap.Group
    Set grp = App.feGroup()

    ' prompt user to select elements
    rc = elm_set.Select(FT_ELM, True, "Select eleme...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "No elements selected. Routine canceled...")
    End If

    grp.SetAdd(FT_ELEM, elm_set.ID)
    grp.AddRelated()
    ' grp.SelectMesh(elm_set.ID)

    App.feViewRegenerate(0)

End Sub