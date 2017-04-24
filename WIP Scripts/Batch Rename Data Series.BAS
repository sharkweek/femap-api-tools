Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim rc As Long

    Dim objSet As femap.Set     'Object set
    Set objSet = App.feSet
    Dim obj As Object
    Set obj = App.feChartSeries
    Dim objID As Long             'Individual solid ID
    Dim newName As String         'Individual solid name
    Dim prevName As String		'Existing indivdual solid name

    rc = objSet.Select ( FT_CHART_SERIES, True, "Select Chart Series for Renaming" )
    If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
        App.feAppMessage ( FCM_HIGHLIGHT, "No Chart Series selected")
        GoTo EndMacro
    End If

	Begin Dialog UserDialog 310,203,"Title for Selected Chart Series" ' %GRID:10,7,1,1
		text 20,14,260,14,"Enter new title:",.seriesName_Text,2
		TextBox 20,42,270,21,.objTitle_TextBox
		GroupBox 20,77,120,77,"Append",.GroupBoxAppend
		GroupBox 170,77,120,77,"Replace",.GroupBoxReplace
		OptionGroup .RenameType
			OptionButton 40,98,70,14,"Prefix"
			OptionButton 40,126,70,14,"Suffix"
			OptionButton 190,98,90,14,"Indexed"
			OptionButton 190,126,90,14,"No Index"
		OKButton 30,168,100,28
		CancelButton 180,168,100,28
	End Dialog
    Dim dlg As UserDialog
    Dialog dlg

    newName = dlg.objTitle_TextBox

    objSet.Reset
    objID = objSet.First

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

	Case 2 'Indexed
	    For i = 0 To objSet.Count-1
	        obj.Get ( objID )
	        obj.title = newName + Str$(i+1)
	        obj.Put ( objID )
	        objID = objSet.Next
	    Next

	Case 3 'No Index
	    For i = 0 To objSet.Count-1
	        obj.Get ( objID )
	        obj.title = newName
	        obj.Put ( objID )
	        objID = objSet.Next
	    Next
	End Select

    EndMacro:
    App.feViewRegenerate ( 0 )

End Sub
