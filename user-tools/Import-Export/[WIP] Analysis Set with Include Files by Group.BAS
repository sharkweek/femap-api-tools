' Title: Analysis Set with Include Files by Group.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2306

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' Usage Notes
' -----------
' This script writes out a selected analysis set as an NX Nastran *.dat file
' substituting all bulk data for `INCLUDE` statements. User-selected groups
' are written out as a separate bulk data file as referred to by the `INCLUDE`
' statements in the master *.dat file.
'
' The general flow of the script is as follows:
'   1. Prompt user for analysis set
'   2. Prompt user for field type
'   3. Prompt user for groups to write to bulk data section
'   4. Create a dummy analysis set for writing out INCLUDE files
'   5. Write out selected analysis set with empty bulk data section
'   6. Loop through each layer, writing out a *.bdf for each one making sure
'      only to write out entities that are contained in the bulk data group
'      request and stripping all case control statements, parameters, etc.
'   7. Rewrite the analysis *.dat file, substituting in INCLUDE statements
'      for each of the created *.bdfs

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Analysis Set with Include Files by Group")
    App.feAppMessage(FCM_NORMAL, "Write out a *.dat file in which all BULK DATA entities have been parsed out by Femap groups. Each")
    App.feAppMessage(FCM_NORMAL, "layer is written out to a separate *.bdf and an INCLUDE statement substituted in its place within")
    App.feAppMessage(FCM_NORMAL, "the *.dat file.")

    Dim rc As Long

    Dim dummyDeck As femap.AnalysisMgr
    Set dummyDeck = App.feAnalysisMgr
    Dim dummyDeckID As Long
    Dim dummyDeckTitle as String 'Placeholder variable to replace spaces with underscores
    Dim field_format As Integer  'field format for write deck

    Dim pickedDeck As femap.AnalysisMgr  ' The run deck to parse
    Set pickedDeck = App.feAnalysisMgr
    Dim pickedDeckTitle As String
    Dim pickedDeckFileName As String

    Dim groupSet As femap.Set
    Set groupSet = App.feSet

    Dim grp As femap.Group
    Set grp = App.feGroup

    Dim origGroupID As Long

    Dim layer As femap.Layer
    Set layer = App.feLayer

    Dim layerSet As femap.Set  'dummy layer set
    Set layerSet = App.feSet

    Dim includeGroupSet As femap.Set  'group set
    Set includeGroupSet = App.feSet

    Dim includeGroup As femap.Group
    Set includeGroup = App.feGroup

    Dim includeFileName As String

    Dim emptyGroup As femap.Group  'empty group for selected run deck
    Set emptyGroup = App.feGroup
    Dim emptyGroupID As Long  'to store the empty group ID for rewriting

    Dim allLayerSet As femap.Set
    Set allLayerSet = App.feSet

    Dim includeFiles() As String

    Dim i As Long  ' counter for layer entity types
    Dim j As Long  ' counter for includeFiles ReDim

    'Write in INCLUDE statements into picked deck here
    Dim readDeck As Integer
    Dim writeDeck As Integer
    Dim textRow As String
    Dim endLine As String
    Dim newDeckName As String
    Dim includeStatement As String

    'Get model name and path
    Dim model As Huge_
    Dim modelFileFull As String
    Dim modelFileName As String
    Dim modelFilePath As String
    App.feAppGetModel(model)
    rc = App.feAppGetModelName(model,modelFileFull)
    modelFileName = Dir(modelFileFull)
    modelFilePath = Left(modelFileFull, Len(modelFileFull)-Len(modelFileName))

    'Prompt user to pick an analysis deck
    rc = pickedDeck.SelectID("Select analysis deck export...")
    If rc = FE_CANCEL Then
        App.feAppMessage(FCM_NORMAL, "Export canceled...")
        Exit Sub
    End If
    pickedDeck.Get(pickedDeck.ID)

    'All field format types stored in strings
    Dim FFormat(5) As String
    FFormat(0) = "Small Field"
    FFormat(1) = "Large Csys, Material, Property"
    FFormat(2) = "Large All But Elements"
    FFormat(3) = "Large"
    FFormat(4) = "Large Csys, Node"

    ' Prompt user to select field format for deck
    Begin Dialog UserDialog 310,119,"Field Format Options" 'NasBulkLargeField %GRID:10,7,1,1
        ListBox 20,14,270,70,FFormat(),.ListBox1
        OKButton 40,84,100,28
        CancelButton 170,84,100,28
    End Dialog
    Dim dlg As UserDialog
    field_format = Dialog(dlg)

    If field_format = -1 Then   'user pressed okay
        field_format = dlg.ListBox1
    ElseIf field_format = 0 Then 'terminates macro if user cancels
        App.feAppMessage(FCM_NORMAL,"User pressed Cancel...")
        End
    End If

    ' prompt user to select groups for bulk data section
    rc = groupSet.Select(FT_GROUP, True, "Select groups...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_FAIL Then
        rc = App.feAppMessage(FCM_ERROR, "Selected groups do not exist.")
    End If

    'Set pickedDeck to write out an empty group for minimal data
    emptyGroup.title = "EMPTY"
    emptyGroupID = emptyGroup.NextEmptyID
    emptyGroup.Put(emptyGroupID)

    origGroupID = pickedDeck.NasBulkGroupID
    pickedDeck.NasBulkGroupID = emptyGroup.ID
    pickedDeck.Put(pickedDeck.ID)
    pickedDeck.Active = pickedDeck.ID  ' feFileWriteNastran uses the active analysis set

    ' scrub name for illegal filename characters
    pickedDeckTitle = Replace(pickedDeck.title," ", "_")
    pickedDeckTitle = Replace(pickedDeckTitle, "_-_", "-")
    pickedDeckTitle = Replace(pickedDeckTitle, "\", "-")
    pickedDeckTitle = Replace(pickedDeckTitle, "/", "-")
    pickedDeckTitle = Replace(pickedDeckTitle, "*", "-")
    pickedDeckTitle = Replace(pickedDeckTitle, ":", "-")
    pickedDeckTitle = Replace(pickedDeckTitle, "&", "and")
    pickedDeckTitle = Replace(pickedDeckTitle, "%", "")

    ' write deck
    pickedDeckFileName = modelFilePath & pickedDeckTitle
    App.feFileWriteNastran(8, pickedDeckFileName)

    ' prep dummy deck for writing out groups
    With dummyDeck
        .ID = dummyDeck.NextEmptyID
        dummyDeckID = dummyDeck.ID
        .NasBulkLargeField = field_format
        .title = "Dummy Deck for Group Export"
        .Solver = 36  'NX Nastran solver
        .AnalysisType = 2  ' modes solution for creation of less data
        .NasCnlSkipNLPARM = True  ' minimizes data in *.dat
        .SkipStandard = True  ' minimizes data in *.dat
        .NasBulkOn = True
        .Put(dummyDeckID)
    End With

    ' create *.bdfs from selected groups
    j = 0
    While grp.NextInSet(groupSet.ID) = FE_OK
        grp.Get(grp.ID)

        ' scrub group name for illegal filename characters
        dummyDeckTitle = grp.title
        dummyDeckTitle = Replace(pickedDeck.title," ", "_")
        dummyDeckTitle = Replace(dummyDeckTitle, "_-_", "-")
        dummyDeckTitle = Replace(dummyDeckTitle, "\", "-")
        dummyDeckTitle = Replace(dummyDeckTitle, "/", "-")
        dummyDeckTitle = Replace(dummyDeckTitle, "*", "-")
        dummyDeckTitle = Replace(dummyDeckTitle, ":", "-")
        dummyDeckTitle = Replace(dummyDeckTitle, "&", "and")
        dummyDeckTitle = Replace(dummyDeckTitle, "%", "")

        includeFileName = modelFilePath & dummyDeckTitle & ".include"

        With dummyDeck
            .Get(dummyDeckID)
            .NasBulkGroupID = grp.ID
            .Put(dummyDeckID)
            .Active = dummyDeckID
        End With

        ' write out include file
        rc = App.feFileWriteNastran(8, includeFileName)
        If rc = FE_OK Then
            App.feAppMessage(FCM_COMMAND,"Exported to file: " & includeFileName)
            CleanBulk(includeFileName) 'Function that writes correct data to new final file
            DeleteTempFile(includeFileName)  'Function that deletes the temporary file
        Else
            App.feAppMessage(FCM_ERROR, "Error writing out " & includeFileName)
        End If

        ' Add file name to include statement array
        ReDim Preserve includeFiles(j + 1)
        includeFiles(j) = "INCLUDE '" & dummyDeckTitle & "'"
        j = j + 1
    Wend

    'Reassign the original analysis group to the analysis set when done
    pickedDeck.NasBulkGroupID = origGroupID
    pickedDeck.Put(pickedDeck.ID)

    'Delete empty group
    emptyGroup.Delete(emptyGroup.ID)

    'Deletes dummy analysis Set after use
    If dummyDeck.Deletable(dummyDeckID) Then
           dummyDeck.Delete(dummyDeckID)
    End If

    ' Write INCLUDE statements into the deck
    newDeckName = pickedDeckFileName & ".dat"
    readDeck = FreeFile
    writeDeck = FreeFile + 1
    Open pickedDeckFileName For Input As #readDeck
    Open newDeckName For Output As #writeDeck
    Do While Not EOF(readDeck)
        Line Input #readDeck, textRow
        If InStr(1, textRow, "ENDDATA") = 1 Then
            i = 0
            Do While i <= (j - 1)  'use final j as length of includeFiles array
                includeStatement = includeFiles(i)
                Print #writeDeck, includeStatement
                i = i + 1
            Loop
            Print #writeDeck, textRow  'place ENDDATA line back at end of file
        ElseIf InStr(1, textRow, "$ Femap") = 1 Then
            'do nothing
        Else
            Print #writeDeck, textRow
        End If
    Loop
    Close #readDeck
    Close #writeDeck

    'Delete original file
    DeleteTempFile(pickedDeckFileName)

    App.feAppMessage(FCM_NORMAL, "Export completed.")

End Sub


Function CleanBulk(fileName As String)

    Set readOld = CreateObject("Scripting.FileSystemObject")
    Set writeNew = CreateObject("Scripting.FileSystemObject")
    Set readOld = readOld.OpenTextFile(fileName, 1)

    Dim found As Boolean  'for "CORD2S" keyword search
    found = False
    Dim foundPlus As Boolean  'for "+" keyword search
    foundPlus = False
    Dim textLine As String 'stores line read in

    Do While Not found And Not readOld.AtEndOfStream   'Searches for keyword
        'Line Input #1, textLine
        textLine = readOld.ReadLine
        If InStr(1, textLine, "CORD2S") > 0 AND InStr(1, textLine, "2") > 0 Then
            found = True
        End If
    Loop

    Do While Not foundPlus And Not readOld.AtEndOfStream   'Searches for first line without a + or *
        textLine = readOld.ReadLine
        If InStr(1, left(textLine,1), "+") = 0 And InStr(1, left(textLine,1), "*") = 0 Then
            foundPlus = True
        End If
    Loop

    Dim newFileName As String   'This file will be the final file for the user
    newFileName = fileName & ".bdf"

    'Creates new text file to write to
    Set createF = CreateObject("Scripting.FileSystemObject")
    Set writeNew = createF.CreateTextFile(newFileName, True)

    While Not readOld.AtEndOfStream  'Loop logic flipped to purposely skip last line
        'Print #2, textLine
        writeNew.WriteLine textLine
        'Line Input #1, textLine
        textLine = readOld.ReadLine

    Wend

    'Close #1  'closing is necessary for deletion
    readOld.Close
    'Close #2  'this file is the final file and will not be deleted
    writeNew.Close

End Function


Function DeleteTempFile(FileToDelete As String)

    SetAttr FileToDelete, vbNormal  'Makes sure it is not in an undeletable view
    Kill FileToDelete   'Deletes file

End Function
