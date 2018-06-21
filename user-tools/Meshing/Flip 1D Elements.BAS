Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Flip 1D Element Orientation")

    Dim rc As Long

    Dim elmSet As femap.Set
    Set elmSet = App.feSet

    Dim elm As femap.Elem
    Set elm = App.feElem

    Dim node1 As Long
    Dim node2 As Long

    ' Prompt user to select 1D elements
    rc = elmSet.Select(FT_ELEM, True, "Select 1D elements...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "None of the selected elements exist...")
        Exit Sub
    ElseIf rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "No elements selected")
        Exit Sub
    End If

    ' Check that all elements selected are 1D elements
    elmSet.Reset()
    Do While elm.NextInSet(elmSet.ID)
        elm.Get(elm.ID)

        Select Case elm.topology

            Case 0 to 1 'If element is a Line2 or Line3 element
                node1 = elm.node(0)
                node2 = elm.node(1)
                elm.node(0) = node2
                elm.node(1) = node1
                elm.Put(elm.ID)

            Case Else
                App.feAppMessage(FCM_NORMAL, "Element " & Cstr(elm.ID) & " is not a 1D element. Element is not flipped.")
        End Select

    Loop

End Sub
