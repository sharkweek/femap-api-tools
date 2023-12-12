' Title: Break 1D Element At Location.BAS
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

    App.feAppMessage(FCM_COMMAND, "Break 1D Element At Location")

    Dim rc As Long

    Dim elm As femap.Elem
    Set elm = App.feElem()

    Dim new_elm As femap.Elem
    Set new_elm = App.feElem()

    Dim node_a As femap.Node
    Set node_a = App.feNode()

    Dim node_b As femap.Node
    Set node_b = App.feNode()

    Dim new_node As femap.Node
    Set new_node = App.feNode()

    Dim xyz As Variant
    Dim break_xyz As Variant

    ' select a linear 1D element
    rc = elm.SelectID("Pick element to break")
    Do While rc <> FE_OK
        App.feAppMessage(FCM_ERROR, "Only linear 1D elements are supported. Select new element")
        rc = elm.SelectID("Pick element to break...")
        If rc = FE_CANCEL Then
            Exit Sub
        End If
    Loop

    App.feCoordPick("Pick location to break element...", xyz)

    ' get entity information
    elm.Get(elm.ID)
    node_a.Get(elm.node(0))
    node_b.Get(elm.node(1))

    ' initialize new node and element
    new_node.Get(node_a.ID)
    new_node.ID = new_node.NextEmptyID()
    App.feCoordOntoLine(xyz, node_a.xyz, node_b.xyz, break_xyz)
    new_node.xyz = break_xyz
    new_node.Put(new_node.ID)
    new_elm.Get(elm.ID)
    new_elm.ID = new_elm.NextEmptyID()

    ' split element
    elm.node(1) = new_node.ID
    elm.Put(elm.ID)
    new_elm.node(0) = new_node.ID
    new_elm.Put(new_elm.ID)

    App.feViewRegenerate(0)

End Sub