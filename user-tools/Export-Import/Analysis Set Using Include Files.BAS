Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Analysis Set Using Include Files")
    App.feAppMessage(FCM_NORMAL, "Write out a *.dat file in which all BULK DATA entities have been parsed out by Femap layers. Each")
    App.feAppMessage(FCM_NORMAL, "layer is written out to a separate *.bdf and an INCLUDE statement substituted in its place within")
    App.feAppMessage(FCM_NORMAL, "the *.dat file.")

    Dim rc As Long

    Dim dummyDeck As femap.AnalysisMgr
    Set dummyDeck = App.feAnalysisMgr
    Dim dummyDeckID As Long
    Dim dummyDeckTitle as String 'Placeholder variable to replace spaces with underscores
    Dim dummyDeckFormat As Integer  'field format for write deck

    Dim pickedDeck As femap.AnalysisMgr  ' The run deck to parse
    Set pickedDeck = App.feAnalysisMgr
    Dim pickedDeckTitle As String
    Dim pickedDeckFileName As String

    Dim pickedGroup As femap.Group
    Set pickedGroup = App.feGroup

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
    dummyDeckFormat = Dialog(dlg)

    If dummyDeckFormat = -1 Then   'user pressed okay
        dummyDeckFormat = dlg.ListBox1
    ElseIf dummyDeckFormat = 0 Then 'terminates macro if user cancels
        App.feAppMessage(FCM_NORMAL,"User pressed Cancel...")
        End
    End If

    'Set properties of dummy analysis set
    dummyDeck.ID = dummyDeck.NextEmptyID
    dummyDeck.NasBulkLargeField = dummyDeckFormat  'Updates field format
    dummyDeck.title = "Dummy Deck for Group Export"  'Names dummy set
    dummyDeck.Solver = 36        'Sets Solver to NX Nastran
    dummyDeck.AnalysisType = 2     'Sets AnalysisType to Modes for creation of less data
    dummyDeck.NasCnlSkipNLPARM = True   'Skips more unneeded data to be written out
    dummyDeck.SkipStandard = True   'Skips more unneeded data to be written out

    'Set pickedDeck write group to an empty group so it doens't write out the full BDF before stripping it
    pickedDeck.Get(pickedDeck.ID)
    pickedGroup.ID = pickedDeck.NasBulkGroupID
    emptyGroup.title = "EMPTY"
    emptyGroup.Put(emptyGroup.NextEmptyID)  'creates the empty group in the model
    pickedDeck.NasBulkGroupID = emptyGroup.ID  'assign the empty group to the selected run deck
    pickedDeck.Put(pickedDeck.ID)
    pickedDeck.Active = pickedDeck.ID  'Activate pickedDeck

    'Write out picked deck here and store name
    pickedDeckTitle = Replace(pickedDeck.title," ", "_") 'Replace spaces with underscores
    pickedDeckTitle = Replace(pickedDeckTitle, "_-_", "-")
    pickedDeckTitle = Replace(pickedDeckTitle, "\", "-")
    pickedDeckTitle = Replace(pickedDeckTitle, "/", "-")
    pickedDeckTitle = Replace(pickedDeckTitle, "*", "-")
    pickedDeckTitle = Replace(pickedDeckTitle, ":", "-")
    pickedDeckTitle = Replace(pickedDeckTitle, "&", "and")
    pickedDeckTitle = Replace(pickedDeckTitle, "%", "")
    pickedDeckFileName = modelFilePath & pickedDeckTitle 'Sets file name with path for picked deck
    App.feFileWriteNastran(8, pickedDeckFileName)

    j = 0
    allLayerSet.AddAll(FT_LAYER)
    includeGroup.ID = includeGroup.NextEmptyID
    'Loop reads out all layer information seperately
    While layer.NextInSet(allLayerSet.ID) = FE_OK

        ' Group objects into includeGroup by layer
        i = 0  'reset counter
        layerSet.Clear
        layerSet.Add(layer.ID)  'Populate a layer set for adding to includeGroupSet
        Do While i <= 168  'Maximum entityType value is 168
            includeGroupSet.AddEntitiesOnLayer(layerSet.ID, i) 'Add entities form layer to set
            includeGroup.SetAdd(i, includeGroupSet.ID)  'Add entities from set to group
            includeGroupSet.Clear  'Clear all entities from set
            i = i + 1  'increment counter
        Loop

        includeGroupSet.Add(includeGroup.ID)  'add group with all layer entities to dummy group set
        App.feGroupBoolean2(includeGroupSet.ID, pickedGroup.ID, 2, includeGroup.ID,"INCLUDE_DUMMY_GROUP")  'add entities only in both group and current layer to new group
        includeGroup.Put(includeGroup.ID)

        ' Set dummyDeck properties
        dummyDeck.NasBulkOn = True
        dummyDeck.NasBulkGroupID = includeGroup.ID 'Set group to includeGroup just created by the boolean operation
        dummyDeck.Put(dummyDeck.ID)    'Adds dummy analysis to femap session
        dummyDeck.Active = dummyDeck.ID  'Activates dummy analysis set
        dummyDeckTitle = layer.title
        dummyDeckTitle = Replace(dummyDeckTitle," ", "_") 'Replace spaces with underscores
        includeFileName = modelFilePath & dummyDeckTitle 'Sets temp file name

        App.feAppMessage(FCM_COMMAND,"Exported to file: " & includeFileName & ".include")
        App.feFileWriteNastran(8, includeFileName)

        CleanBulk(includeFileName) 'Function that writes correct data to new final file
        DeleteTempFile(includeFileName)  'Function that deletes the temporary file
        includeGroup.RangeDeleteAll(-1)  'Delete all entities from group
        includeGroup.Delete(includeGroup.ID)

        ' Add file name to include statement array
        ReDim Preserve includeFiles(j + 1)
        includeFiles(j) = "INCLUDE '" & dummyDeckTitle & ".include'"
        j = j + 1

    Wend

    includeGroup.Delete(includeGroup.ID)' Delete dummy group

    'Reassign the original analysis group to the analysis set when done
    pickedDeck.NasBulkGroupID = pickedGroup.ID
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
