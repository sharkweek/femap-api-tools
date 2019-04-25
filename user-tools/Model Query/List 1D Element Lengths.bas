' Title: List 1D Element Lengths.bas
' Author: Andy Perez
' Date: April 2019
' Femap API Version: 12.0

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "List 1D Element Lengths")
    App.feAppMessage(FCM_NORMAL, "List the lengths of multiple 1D elements using the active coordinate system")

    Dim rc As Long

    Dim elmSet As femap.Set
    Set elmSet = App.feSet

    Dim c_sys As femap.CSys
    Set c_sys = App.feCSys

    Dim skip As Boolean

    Dim e As femap.Elem
    Set e = App.feElem

    Dim n1 As femap.Node
    Set n1 = App.feNode

    Dim n2 As femap.Node
    Set n2 = App.feNode

    Dim _vecBase As Variant
    Dim _vecTip As Variant
    Dim vecDist As Variant
    Dim dist As Double

    Dim pString As String

    ' prompt user to select elements
    rc = elmSet.Select(FT_ELEM, True, "Select elements...")
    If rc = FE_OK Then
        ' remove non-1D elements
        For i = 2 To 19
            elmSet.RemoveRule(i, FGD_ELEM_BYSHAPE)
        Next i
        App.feAppMessage(FCM_NORMAL, "Non-1D elements removed...")

    ElseIf rc = FE_CANCEL Then
        Exit Sub

    ElseIf rc = FE_FAIL Then
        rc = App.feAppMessage(FCM_ERROR, "Selected elements do not exist.")

    End If

    ' print header
    pString = "EID      X        Y        Z           Total"
    App.feAppMessage(FCM_COMMAND, pString)

    ' print distances for each element
    Do While e.NextInSet(elmSet.ID)
        n1.Get(e.nodes(0))
        n2.Get(e.nodes(1))
        App.feMeasureDistanceBetweenNodes2(n1.ID, n2.ID, 0, 0, c_sys.Active, _
                                           _vecBase, _vecTip, vecDist, dist)

        pString = Left(CStr(e.ID) & Space(8), 8) & " " & _
                  Left(CStr(vecDist(0)) & Space(8), 8) & " " & _
                  Left(CStr(vecDist(1)) & Space(8), 8) & " " & _
                  Left(CStr(vecDist(2)) & Space(8), 8) & " " & _
                  Left(CStr(dist) & Space(8), 8) & " "

        App.feAppMessage(FCM_NORMAL, pString)
    Loop

    App.feViewRegenerate(0)

End Sub