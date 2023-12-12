' Title: Results to Femap Neutral Output.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 12.0

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.


Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Results to Femap Neutral Output")
    App.feAppMessage(FCM_NORMAL, "Exports results in model to a Femap Neutral Output (*.fno) file. One output set is exported per *.fno.")

    Dim rc As Long

    Dim modelOutputSet As femap.OutputSet
    Set modelOutputSet = App.feOutputSet

    Dim osSet As femap.Set
    Set osSet = App.feSet

    Dim gr As femap.Group
    Set gr = App.feGroup

    Dim passCount As Long

    Dim tempSet As femap.Set
    Set tempSet = App.feSet

    Dim longTitles As Boolean

    Dim fnoName As String

    Dim model As Huge_
    Dim modelFileFull As String
    Dim modelFileName As String
    Dim modelFilePath As String

    Dim startTime As Single
    Dim finishTime As Single
    Dim totalTime As Single

    'Get model name and path
    App.feAppGetModel(model)
    App.feAppGetModelName(model, modelFileFull)
    modelFileName = Dir(modelFileFull)
    modelFilePath = Left(modelFileFull, Len(modelFileFull)-Len(modelFileName))

    ' Prompt user to select output sets
    rc = osSet.Select( FT_OUT_CASE, True, "Select Output Sets..." )
    If rc = FE_CANCEL Then
        Exit Sub
    End If

    ' Export full model or group only
    rc = App.feAppMessageBox(2, "Export the full model?")
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
    totaltime = 0

    ' Write out *.fno
    osSet.Reset()
    App.feAppMessage(FCM_NORMAL, "Exporting results to " & modelFilePath)
    Do While modelOutputSet.NextInSet(osSet.ID)
        ' find corresponding output set from model
        modelOutputSet.Get(modelOutputSet.ID)

        ' create file name based on output set ID and title
        If longTitles Then
            fnoName = Str$(modelOutputSet.ID) + ".." + modelOutputSet.title + ".fno"
        Else
            fnoName = Str$(modelOutputSet.ID) + ".fno"
        End If

        ' Create dummy set to include the id of the output set
        tempSet.Clear
        tempSet.Add(modelOutputSet.ID)

        ' Write out the output set in an individual file
        startTime = Time
        rc = App.feFileWriteFNO(tempSet.ID, 0, gr.ID, modelFilePath + fnoName)
        finishTime = Time
        totalTime += finishTime - startTime
        If rc = FE_OK Then
            App.feAppMessage(FCM_NORMAL, fnoName + " successfully written in " & Format(Cstr(finishTime - startTime), "HH:MM:SS") & "...")
            passCount += 1
        Else
            App.feAppMessage(FCM_ERROR, "Unable to write " + fnoName)
        End If
    Loop

    App.feAppMessage(FCM_NORMAL, CStr(passCount) + " results set(s) successfully written in " & Format(CStr(totalTime), "HH:MM:SS"))

End Sub
