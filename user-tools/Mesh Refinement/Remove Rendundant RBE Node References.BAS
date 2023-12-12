' Title: Remove Rendundant RBE Node References.BAS
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

    App.feAppMessage(FCM_COMMAND, "Remove Rendundant RBE Node References")

    Dim rc As Long

    Dim rbe_set As femap.Set
    Set rbe_set = App.feSet()

    Dim elm As femap.Elem
    Set elm = App.feElem()

    Dim node_set As femap.Set
    Set node_set = App.feSet()

    Dim node_count As Long
    Dim node As Variant
    Dim face As Variant
    Dim weight As Variant
    Dim dof As Variant

    rbe_set.AddRule(29, FGD_ELEM_BYTYPE)

    Do While rbe_set.Next()
        elm.Get(rbe_set.CurrentID)
        elm.GetNodeList(0, node_count, node, face, weight, dof)
        node_set.AddArray(node)
    Loop

    App.feViewRegenerate(0)

End Sub