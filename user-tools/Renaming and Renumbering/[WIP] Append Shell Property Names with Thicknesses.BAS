Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Append Shell Property Names with Thickness")

    Dim rc As Long

    Dim append_all As Boolean

    Dim pSet As femap.Set
    Set pSet = App.feSet

    Dim p As femap.Prop
    Set p = App.feProp

    rc = App.feAppMessageBox(3, "Append all property names? (No=selected properties only)")
    If rc = FE_OK Then
        append_all = True
    ElseIf rc = FE_FAIL Then
        append_all = False
    Else
        Exit Sub
    End If

    If append_all Then
        ' Add all shell property types to group
        For i = 11 to 20  ' do not include laminates
            pSet.AddRule(i, FGD_Prop_byType)
        Next i

    Else
        rc = pSet.Select(FT_PROP, )
    End If



End Sub