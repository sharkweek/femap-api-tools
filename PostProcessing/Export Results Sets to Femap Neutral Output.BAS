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

    rc = gr.SelectID("Select Group to Export (Cancel for Full Model)")
    rc = ouSets.Select( FT_OUT_CASE, True, "Select Output Sets to Export" )

    'Ask for output file prefixes/suffixes
    Begin Dialog UserDialog 320,154,"Append File Names" ' %GRID:10,7,1,1
        TextBox 20,35,280,21,.Prefix
        TextBox 20,91,280,21,.Suffix
        OKButton 50,126,90,21
        CancelButton 180,126,90,21
        text 20,14,150,14,"Prefix for *.fno files",.PrefixTitle
        text 20,70,150,14,"Suffix for *.fno files",.SuffixTitle
    End Dialog
    Dim dlg As UserDialog
    Dialog dlg

    If Right(dlg.Prefix) = " " 0R Right(dlg.Prefix) = "_" Then
        filePrefix = dlg.Prefix + "OutputSet_"
    Else
        filePrefix = dlg.Prefix + "_OutputSet_"
    End If

    If Left(dlg.Suffix) = " " OR Left(dlg.Suffix) = "_" Then
        fileSuffix = dlg.Suffix
    Else
        fileSuffix = "_" + dlg.Suffix
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
