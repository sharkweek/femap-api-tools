' Title: Next Empty Mesh ID.BAS
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

    App.feAppMessage(FCM_COMMAND, "Next Empty Mesh ID")

    Dim min_id As Long

    Dim elm As femap.Elem
    Set elm = App.feElem()

    Dim nd As femap.Node
    Set nd = App.feNode()

    Dim col_width As Long
    Dim next_elm As Long
    Dim prev_elm As Long
    Dim next_nd As Long
    Dim prev_nd As Long
    Dim next_elm_text As String
    Dim prev_elm_text As String
    Dim next_nd_text As String
    Dim prev_nd_text As String

    col_width = 12

    App.feGetInt("Minimum ID", 1, 9999999, min_id)
    elm.ID = min_id
    nd.ID = min_id
    next_elm = elm.NextEmptyID()
    next_nd = nd.NextEmptyID()
    prev_elm = elm.PrevEmptyID()
    prev_nd = nd.PrevEmptyID()

    ' format ids for width
    next_elm_text = Right(Space(col_width) & CStr(next_elm), col_width)
    prev_elm_text = Right(Space(col_width) & CStr(prev_elm), col_width)
    next_nd_text =  Right(Space(col_width) & CStr(next_nd), col_width)
    prev_nd_text =  Right(Space(col_width) & CStr(prev_nd), col_width)

    ' print out next and previous IDs
    App.feAppMessage(FCM_NORMAL, "              " & _
                     Right(Space(col_width) & "Prev ID", col_width) & " " & _
                     Right(Space(col_width) & "Next ID", col_width))
    App.feAppMessage(FCM_NORMAL, "Element ID:   " & prev_elm_text & " " & next_elm_text)
    App.feAppMessage(FCM_NORMAL, "Node ID:      " & prev_nd_text & " " & next_nd_text)

    ' set next IDs for element and node creation automatically
    App.Info_NextID(FT_ELEM) = next_elm
    App.Info_NextID(FT_NODE) = next_nd

    App.feViewRegenerate(0)

End Sub