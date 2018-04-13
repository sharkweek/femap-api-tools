Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Export Layers to BDFs")

    Dim rc As Long

    Dim runDeck As femap.AnalysisMgr
    Set runDeck = App.feAnalysisMgr
    Dim runDeckID As Long
    Dim runDeckTitle as String 'Placeholder variable to replace spaces with underscores
    Dim runDeckFormat As Integer  'field format for write deck

    Dim layer As femap.Layer
    Set layer = App.feLayer

    Dim layerSet As femap.Set
    Set layerSet = App.feSet

    Dim dlSet As femap.Set
    Set dlSet = App.feSet

    Dim grSet As femap.Set
    Set grSet = App.feSet

    Dim gr As femap.Group
    Set gr = App.feGroup
    Dim grName As String
    Dim grFileName As String

    Dim i As Long  ' Dummy counter

    'Get model name and path
    Dim model As Huge_
    Dim modelFileFull As String
    Dim modelFileName As String
    Dim modelFilePath As String
    App.feAppGetModel(model)
    rc = App.feAppGetModelName(model,modelFileFull)
    modelFileName = Dir(modelFileFull)
    modelFilePath = Left(modelFileFull, Len(modelFileFull)-Len(modelFileName))

    rc = layerSet.SelectMultiID(FT_LAYER, 1, "Select layers to export...")
    If rc = FE_CANCEL Then
        App.feAppMessage(FCM_NORMAL, "Export canceled...")
        Exit Sub
    End If

    'All field format types stored in strings
    Dim FFormat(5) As String
    FFormat(0) ="Small Field"
    FFormat(1) ="Large Csys, Material, Property"
    FFormat(2) ="Large All But Elements"
    FFormat(3) ="Large"
    FFormat(4) ="Large Csys, Node"

    'User options
    Begin Dialog UserDialog 310,119,"Field Format Options" 'NasBulkLargeField %GRID:10,7,1,1
        ListBox 20,14,270,70,FFormat(),.ListBox1
        OKButton 40,84,100,28
        CancelButton 170,84,100,28
    End Dialog
    Dim dlg As UserDialog
    runDeckFormat = Dialog(dlg)

    If runDeckFormat = -1 Then   'user pressed okay
        runDeckFormat = dlg.ListBox1
    ElseIf runDeckFormat = 0 Then 'terminates macro if user cancels
        App.feAppMessage(FCM_NORMAL,"User pressed Cancel...")
        End
    End If

    'Set properties of analysis set
    runDeckID = runDeck.NextEmptyID   'Finds next avaliable ID to make analysis set in
    runDeck.NasBulkLargeField = runDeckFormat  'Updates field format
    runDeck.title = "Dummy Set for Group Export"     'Names dummy set
    runDeck.Solver = 36        'Sets Solver to NX Nastran
    runDeck.AnalysisType = 2     'Sets AnalysisType to Modes for creation of less data
    runDeck.NasCnlSkipNLPARM = True   'Skips more unneeded data to be written out
    runDeck.SkipStandard = True   'Skips more unneeded data to be written out

    'Loop reads out all layer information seperately
    While layer.NextInSet(layerSet.ID) = FE_OK

        'Populate dummy layer set for adding to grSet
        dlSet.Clear
        dlSet.Add(layer.ID)

        ' Get objects on specific layers onto dummy group
        i = 1
        Do While i <= 168  'Maximum entityType value is 168
            grSet.AddEntitiesOnLayer(dlSet.ID, i) 'Add entities form layer to set
            gr.SetAdd(i, grSet.ID)  'Add entities from set to group
            grSet.Clear  'Clear all entities from set
            i = i + 1  'increment counter
        Loop

        gr.Put(gr.NextEmptyID)  'create group in model for analysis deck

        ' Set runDeck properties
        runDeck.NasBulkOn = True
        runDeck.NasBulkGroupID = gr.ID 'User selected group
        runDeck.Put(runDeckID)    'Adds dummy analysis to femap session
        runDeck.Active = runDeckID  'Activates dummy analysis set
        runDeckTitle = layer.title
        runDeckTitle = Replace(runDeckTitle," ", "_") 'Replace space with underscore
        grFileName = modelFilePath & runDeckTitle 'Sets temp file name

        App.feAppMessage(FCM_COMMAND,"Exported to file: " & grFileName & ".bdf")
        App.feFileWriteNastran(8,grFileName)

        FixFile(grFileName) 'Function that writes correct data to new final file
        DeleteTempFile(grFileName)  'Function that deletes the temporary file

        gr.Delete(gr.ID)  ' Delete dummy group

    Wend

    'Deletes dummy analysis Set after use
    If runDeck.Deletable(runDeckID) Then
           runDeck.Delete(runDeckID)
    End If

    App.feAppMessage(FCM_NORMAL, Cstr(layerSet.Count) & " layers exported.")
End Sub


Function FixFile(fileName As String)

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
