' Title: Purge Non-Mesh Entities from Group.BAS
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

    App.feAppMessage(FCM_COMMAND, "Purge Non-Mesh Entities from Group")

    Dim rc As Long

    Dim gr As femap.Group
    Set gr = App.feGroup

    Dim nodeSet As femap.Set
    Set nodeSet = App.feSet

    Dim elmSet As femap.Set
    Set elmSet = App.feSet

    Dim grID As Long

    ' Select group to purge
    rc = App.feSet.SelectID(FT_GROUP, "Select group to purge", grID)
    If rc = FE_CANCEL Then
        Exit Sub
    End If

    'Retrieve nodes and elements
    nodeSet.AddGroup(FT_NODE, grID)
    elmSet.AddGroup(FT_ELEM, grID)

    ' Get group
    gr.Get(grID)

    'Purge all
    rc = gr.RangeDeleteAll(-1)

    'Add nodes and elements
    rc = gr.SetAdd(FT_NODE, nodeSet.ID)
    rc = gr.SetAdd(FT_ELEM, elmSet.ID)

    ' Assign back to group
    gr.Put(grID)

    ' Regenerate
    App.feViewRegenerate(0)

End Sub
