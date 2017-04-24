'Macro Export by Group
'A. Pavageau 05/10
'v2 04/2012 - Tel: 5721

Sub Main
	Dim App As femap.model
	Set App = feFemap
	App.feAppMessage(FCM_ERROR,"Macro Export by Group")

	Dim doublePrec As Long
	Dim grSet As femap.Set, gr As femap.Group, grRefSet As femap.Set, bdfName As String
	Set grSet = App.feSet
	Set gr = App.feGroup
	If grSet.SelectMultiID(FT_GROUP,1,"Select groups to export") = FE_CANCEL Then Exit Sub

	doublePrec = App.feAppMessageBox(3,"Export in double precision?")
	If doublePrec = FE_CANCEL Then End

'retrieve folder path method
    Dim objShell, objFolder, objFolderItem
	Set objShell = CreateObject("Shell.Application")
	Set objFolder = objShell.BrowseForFolder(0, "Select folder to write file(s)",&H1)
	If objFolder Is Nothing Then
		Exit Sub
	Else
		Set objFolderItem = objFolder.Self
		ChDir(objFolderItem.Path)
	End If
	Set objFolderItem = Nothing : Set objFolder = Nothing : Set objShell = Nothing
'end folder path method

	Dim anSet As femap.AnalysisMgr
	Set anSet = App.feAnalysisMgr
	With anSet
		.ID = .NextEmptyID : .title = "Macro Export by Group"
		.InitAnalysisMgr(FAM_MSC_NASTRAN,FAT_MODES,True)
		.NasBulkOn = True
		If doublePrec = FE_OK Then .NasBulkLargeField = 4
		.Active = .ID
	End With

	While gr.NextInSet(grSet.ID)
		Set grRefSet = gr.ReferencedGroups
		If grRefSet.Count <> 0 Then GoTo SKIPGR

		anSet.NasBulkGroupID = gr.ID
		anSet.Put(anSet.ID)

		'deal with group names...
		bdfName = gr.title
		If bdfName = "" Then bdfName = "Group" & CStr(gr.ID) & ".bdf"
		If InStr(bdfName,".") = 0 Then bdfName = bdfName & ".bdf"
		While InStr(bdfName,"/") <> 0
			bdfName = Mid(bdfName,InStr(bdfName,"/")+1)
		Wend

		App.feFileWriteNastran(0, bdfName)

		Call modOutputFiles(bdfName,doublePrec)
		App.feAppMessage(0, "File" & bdfName & " created")
	SKIPGR:
	Wend

	App.feDelete(FT_AMGR_DIR,-anSet.ID)
	App.feAppMessage(FCM_ERROR,"Macro is done")
End Sub

'sub which deletes useless lines automatically writtenby FEMAP at the start and end of files
Private Sub modOutputFiles(file As String, doubleP As Long)
	Dim curline As String, s1 As String, l1 As Long
	If doubleP = False Then
		s1 = "CORD2S         2" : l1 = 16
	Else
		s1 = "CORD2S*                2" : l1 = 24
	End If

	Open file For Input As #1
	Open "temp" For Output As #2

	'look for definition of CSys 2
	Do
		Line Input #1, curline
	Loop Until Left(curline,l1) = s1
	Line Input #1, curline

	If doubleP = True Then
		Line Input #1, curline
		Line Input #1, curline
	End If

	Line Input #1, curline
	While Left(curline,7) <> "ENDDATA"
		Print #2, curline
		Line Input #1, curline
	Wend

	Close #1
	Close #2

	Kill file
	Name "temp" As file
End Sub
