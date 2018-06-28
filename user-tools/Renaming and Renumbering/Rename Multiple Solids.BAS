Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Rename Multiple Solids")

    Dim rc As Long

    Dim objSet As femap.Set
    Set objSet = App.feSet
    Dim obj As Object
    Set obj = App.feSolid       'turn obj into a solid
    Dim objID As Long
    Dim newName As String       'new name from user
    Dim prevName As String      'existing name

    'User inputs
    Begin Dialog UserDialog 310,301,"Title for Selected Solid(s)" ' %GRID:10,7,1,1
        text 20,14,260,14,"Enter new title:",.TextBox_Title,2
        TextBox 30,35,250,14,.NewName
        GroupBox 20,70,120,77,"Append",.GroupBoxAppend
        GroupBox 170,70,120,77,"Overwrite",.GroupBoxReplace
        GroupBox 20,161,270,98,"Find and Replace",.ReplaceBox
        text 40,210,120,14,"Text to replace:"
        OptionGroup .RenameType
            OptionButton 40,91,70,14,"Prefix"
            OptionButton 40,119,70,14,"Suffix"
            OptionButton 190,91,90,14,"Indexed"
            OptionButton 190,119,90,14,"No Index"
            OptionButton 40,182,120,14,"Replace"
        TextBox 40,231,230,14,.ToReplace
        CheckBox 190,182,90,14,"Indexed",.ReplaceIndex
        OKButton 50,266,100,28
        CancelButton 160,266,100,28
    End Dialog

    Dim dlg As UserDialog
    Dialog dlg

    newName = dlg.NewName

    'prompt user to select solids for renaming
    rc = objSet.Select ( FT_SOLID, True, "Select Solids for Renaming" )
    If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
        App.feAppMessage ( FCM_HIGHLIGHT, "No objs selected")
        GoTo EndMacro
    End If

    objSet.Reset
    objID = objSet.First

    'rename object according to user's selections
    Select Case dlg.RenameType
    Case 0 'Prefix
        For i = 0 To objSet.Count-1
            obj.Get ( objID )
            prevName = obj.title
            obj.title = newName + prevName
            obj.Put ( objID )
            objID = objSet.Next
        Next

    Case 1 'Suffix
        For i = 0 To objSet.Count-1
            obj.Get ( objID )
            prevName = obj.title
            obj.title = prevName + newName
            obj.Put ( objID )
            objID = objSet.Next
        Next

    Case 2 ' Overwrite (Indexed)
        For i = 0 To objSet.Count-1
            obj.Get ( objID )
            obj.title = newName + Str$(i+1)
            obj.Put ( objID )
            objID = objSet.Next
        Next

    Case 3 'Overwrite (No Index)
        For i = 0 To objSet.Count-1
            obj.Get ( objID )
            obj.title = newName
            obj.Put ( objID )
            objID = objSet.Next
        Next

    Case 4 'Find and Replace
        If dlg.ReplaceIndex Then
            For i = 0 To objSet.Count-1
                obj.Get ( objID )
                prevName = obj.title
                obj.title = Replace(prevName, dlg.ToReplace, newName + Str$(i+1))
                obj.Put ( objID )
                objID = objSet.Next
            Next

        Else
            For i = 0 To objSet.Count-1
                obj.Get ( objID )
                prevName = obj.title
                obj.title = Replace(prevName, dlg.ToReplace, newName)
                obj.Put ( objID )
                objID = objSet.Next
            Next
        End If
    End Select

    EndMacro:
    App.feViewRegenerate ( 0 )

End Sub
