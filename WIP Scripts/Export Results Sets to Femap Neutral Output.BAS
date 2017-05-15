'Export Results to Femap Neutral Output

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim rc As Long

    Dim ouSets As femap.Set
    Set ouSets = App.feSet

    Dim gr As femap.Group
    Set gr = App.feGroup

    Dim exportGroupID As Long

    Dim filePrefix As String
    Dim fileSuffix As String

    Dim fno_Name As String

    Dim ouSetID As Long

    Dim tempSet As femap.Set
    Set tempSet = App.feSet

    ' Prompt user to select group to export
    rc = gr.SelectID("Select Group to Export (Cancel for Full Model)")
    If rc = FE_CANCEL Then
        exportGroupID = 0
    Else
        expotGroupID = gr.ID
    End If

    ' Prompt user to select output sets
    rc = ouSets.Select( FT_OUT_CASE, True, "Select Output Sets" )
    If rc = FE_CANCEL Then
        Exit Sub
    End If

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

    If Right(dlg.Prefix,1) = " " _
        Or Right(dlg.Prefix,1) = "_" _
        Or dlg.Prefix = "" Then
        filePrefix = dlg.Prefix + "OutputSet_"
    Else
        filePrefix = dlg.Prefix + "_OutputSet_"
    End If

    If Left(dlg.Suffix,1) = " " _
        Or Left(dlg.Suffix,1) = "_" _
        Or dlg.Suffix = "" Then
        fileSuffix = dlg.Suffix
    Else
        fileSuffix = "_" + dlg.Suffix
    End If

    ' Write out *.fno
    If ouSets.Count > 0 Then
            ouSetID = ouSets.First
            While ouSetID > 0
                fno_Name = filePrefix + Str(ouSetID) + fileSuffix + ".fno"
                tempSet.Clear
                tempSet.Add( ouSetID )
                rc = App.feFileWriteFNO( tempSet.ID, 0, exportGroupID, fno_Name )
                ouSetID = ouSets.Next
            Wend
    End If

End Sub
