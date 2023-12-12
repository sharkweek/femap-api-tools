' Title: Remove RBE2s.BAS
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

    Dim etype As femap.zDataType

    Dim elm_set As femap.Set
    Set elm_set = App.feSet()

    Dim rbe_set As femap.Set
    Set rbe_set = App.feSet()

    Dim elm As femap.Elem
    Set elm = App.feElem()

    ' start selection tool
    App.feStartSelectionTool(elm_set.ID, etype)
    If etype <> FT_ELEM Then
        App.feAbortSelectionTool()
    End If

    ' get all rigid elements in selection
    rbe_set.AddRule(13, FGD_ELEM_BYSHAPE)
    rbe_set.RemoveNotCommon(elm_set.ID)

    ' RBE2s in selection
    rbe_set.Reset()
    Do While rbe_set.Next()
        elm.Get(rbe_set.CurrentID)
        ' remove RBE3s from set
        If elm.RigidInterpolate Then
            rbe_set.Remove(elm.ID)
        End If
    Loop

    ' remove RBE2s
    App.feEndSelectionTool(0, rbe_set.ID)

End Sub