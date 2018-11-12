'Export Results to Femap Neutral Output

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Results to Femap Neutral Output")

    Dim rc As Long

    Dim modelOutputSets As femap.OutputSet
    Set modelOutputSets = App.feOutputSet

    Dim ouSets As femap.Set
    Set ouSets = App.feSet

    Dim gr As femap.Group
    Set gr = App.feGroup

    Dim ouSetID As Long

    Dim passCount As Long

    Dim tempSet As femap.Set
    Set tempSet = App.feSet

    Dim longTitles As Boolean

    Dim fno_Name As String

    Dim model As Huge_
    Dim modelFileFull As String
    Dim modelFileName As String
    Dim modelFilePath As String

    'Get model name and path
    App.feAppGetModel(model)
    App.feAppGetModelName(model,modelFileFull)
    modelFileName = Dir(modelFileFull)
    modelFilePath = Left(modelFileFull, Len(modelFileFull)-Len(modelFileName))

    ' Prompt user to select output sets
    rc = ouSets.Select( FT_OUT_CASE, True, "Select Output Sets." )
    If rc = FE_CANCEL Then
        Exit Sub
    End If

    ' Export full model or group only
    rc = App.feAppMessageBox(2, "Do you wish to export the full model?")
    If rc = FE_OK Then
        gr.ID = 0
    Else
        gr.SelectID("Select a group to export.")
    End If

    ' Use output set titles in .fno names or only IDs
    rc = App.feAppMessageBox(3, "Use output set titles?")
    If rc = FE_OK Then
        longTitles = True
    ElseIf rc = FE_FAIL Then
        longTitles = False
    Else
        Exit Sub
    End If

    passCount = 0

    ' Write out *.fno
    If ouSets.Count > 0 Then
        ouSetID = ouSets.First
        While ouSetID > 0
            ' find corresponding output set from model
            modelOutputSets.Get(ouSetID)

            ' create file name based on output set ID and title
            If longTitles Then
                fno_Name = Str$(ouSetID) + ".." + modelOutputSets.title + ".fno"
            Else
                fno_Name = Str$(ouSetID) + ".fno"
            End If

            ' Create dummy set to include the id of the output set
            tempSet.Clear
            tempSet.Add( ouSetID )

            ' Write out the output set in an individual file
            rc = App.feFileWriteFNO( tempSet.ID, 0, gr.ID, modelFilePath + fno_Name )
            If rc = FE_OK Then
                App.feAppMessage(FCM_NORMAL, fno_Name + " successfully written to model directory...")
                passCount = passCount + 1
            Else
                App.feAppMessage(FCM_ERROR, "Unable to write " + fno_Name + " to model directory.")
            End If

            ' Cycle set ID
            ouSetID = ouSets.Next
        Wend
    End If

    App.feAppMessage(FCM_NORMAL, CStr(passCount) + " results set(s) successfully written.")

End Sub
