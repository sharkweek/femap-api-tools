' Title: Clear Group Entities and Rules.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2301

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Clear Group Entities and Rules")

    Dim rc As Long

    Dim gr As femap.Group
    Set gr = App.feGroup

    Dim grID As Long
    Dim title As String

    ' Select group to purge
    rc = App.feSet.SelectID(FT_GROUP, "Select group to clear...", grID)
    If rc = FE_CANCEL Then
        Exit Sub
    End If

    ' Get group
    gr.Get(grID)
    title = gr.Title

    'Purge all
    gr.Clear()
    gr.title = title

    ' Assign back to group
    gr.Put(grID)

    ' Regenerate
    App.feViewRegenerate(0)

End Sub
