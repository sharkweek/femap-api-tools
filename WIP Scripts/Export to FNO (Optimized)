'Export Results to Femap Neutral Output
'If automatic "Attach to Results" is desired, uncomment line 55
Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim ouSets As femap.Set
    Set ouSets = App.feSet

    Dim gr As femap.Group
    Set gr = App.feGroup

    Dim filePrefix As String
    Dim fileSuffix As String

    Dim fno_Name As String

    Dim ouSetID As Long

    Dim tempSet As femap.Set
    Set tempSet = App.feSet

    Dim groupList() As Variant
    Dim groupCount As Integer

    Dim grCount As Long
    Dim grIDs As Variant
    Dim grTitles As Variant

    rc = gr.GetTitleIDList(True, 0, 0, grCount, grIDs, grTitles)

    'rc = gr.SelectID("Select Group to Export")
    rc = ouSets.Select( FT_OUT_CASE, True, "Select Output Sets to Export" )

    'Ask for output file prefixes/suffixes
    Begin Dialog UserDialog 320,238,"Append File Names" ' %GRID:10,7,1,1
		TextBox 20,28,280,21,.Prefix
		TextBox 20,77,280,21,.Suffix
		OKButton 50,210,90,21
		CancelButton 180,210,90,21
		text 20,7,150,14,"Prefix for *.fno files",.PrefixTitle
		text 20,56,150,14,"Suffix for *.fno files",.SuffixTitle
		GroupBox 10,112,300,91,"Portion of model to export",.GroupBox1
		DropListBox 20,175,280,21,grTitles(),.Group
		OptionGroup .ModelGroup
			OptionButton 30,133,130,14,"Full Model",.FullModelOption
			OptionButton 30,154,90,14,"Group",.GroupOption
    End Dialog
    Dim dlg As UserDialog
    rc = Dialog(dlg)

    filePrefix = dlg.Prefix + "OutputSet_"

    If fileSuffix = Nothing Then
        fileSuffix = ""
    Else
        fileSuffix = dlg.Suffix
    End If

    ' Write out *.fno
    If ouSets.Count > 0 Then
            ouSetID = ouSets.First
            While ouSetID > 0
                fno_Name = filePrefix + Str$(ouSetID) + fileSuffix + ".fno"
                tempSet.Clear
                tempSet.Add( ouSetID )
                rc = App.feFileWriteFNO( tempSet.ID, 0, gr.ID, fno_Name )
                'App.feFileAttachResults( FAP_NE_NASTRAN, fno_Name, False )
                ouSetID = ouSets.Next
            Wend
    End If

End Sub
