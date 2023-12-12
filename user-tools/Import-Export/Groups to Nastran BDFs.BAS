' Title: Groups to Nastran BDFs.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2306

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' This script is a modified version of `WriteGroupsToNastranFiles_r2.BAS` by
' Ryan Blelloch and Michael Palodichuk at ATA Engineering Description: This API
' allows user to select one or more FEMAP groups and exports their bulk data to
' individual Nastran INCLUDE files.

'Usage Notes
'====================================
' The active FEMAP model is automatically saved upon running this API The files
' written out are put in the same directory as the FEMAP model The files are
' stripped of extraneous data including the basic coordinate system definitions
' and the ENDDATA tag.

'Copyright (c) <2015> <ATA Engineering, Inc.>

' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software") to use,
' copy, modify, merge, and/or distribute royalty-free copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to
' the following conditions: The above copyright notice and this permission
' notice shall be included in all copies or substantial portions of the
' Software. THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
' EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
' MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO
' EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES
' OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
' ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.

Dim App As femap.model

Sub Main

    'Load model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND,"Groups to Nastran BDFs")

    'Get model name and path
    Dim model As Huge_
    Dim modelFileFull As String
    Dim modelFileName As String
    Dim modelFilePath As String
    App.feAppGetModel(model)
    App.feAppGetModelName(model,modelFileFull)
    modelFileName = Dir(modelFileFull)
    modelFilePath = Left(modelFileFull, Len(modelFileFull)-Len(modelFileName))

    'Make group set object
    Dim grSet As femap.Set
    Set grSet =App.feSet

    ' group obj
    Dim gr As femap.Group
    Set gr = App.feGroup
    Dim grName As String
    Dim grFileName As String

    ' dummy analysis set from which to write deck
    Dim sao As femap.AnalysisMgr
    Set sao = App.feAnalysisMgr
    Dim saoID As Long

    'Set properties of analysis set
    saoID = sao.NextEmptyID                   'Finds next avaliable ID to make analysis set in
    sao.title = "Dummy Set for Group Export"  'Names dummy set
    sao.Solver = 36                           'Sets Solver to NX Nastran
    sao.AnalysisType = 2                      'Sets AnalysisType to Modes for creation of less data
    sao.NasCnlSkipNLPARM = True               'Skips more unneeded data to be written out
    sao.SkipStandard = True                   'Skips more unneeded data to be written out

    'All field format types stored in strings
    Dim FFormat(5) As String
    FFormat(0) ="Small Field"
    FFormat(1) ="Large Csys, Material, Property"
    FFormat(2) ="Large All But Elements"
    FFormat(3) ="Large"
    FFormat(4) ="Large Csys, Node"

    ' prompt user for groups to export
    If grSet.SelectMultiID(FT_GROUP, 1, "Select Group(s) to Write to Nastran File(s)") = FE_CANCEL Then
        App.feAppMessage(FCM_NORMAL,"Canceled...")
        End 'terminates macro if user cancels
    End If

    'Prompts user for field format
    Dim UserFormat As Integer

    Begin Dialog UserDialog 290,100,"Field Format Options" 'NasBulkLargeField %GRID:10,5,1,1
        ListBox 10,7,270,70,FFormat(),.format
        OKButton 40,75,100,20
        CancelButton 150,75,100,20
    End Dialog
    Dim dlg As UserDialog

    ' prompt user
    If Dialog(dlg) = 0 Then
        Exit Sub  ' exit if user cancels
    Else
        UserFormat = dlg.format
    End If

    'Updates field format
    sao.NasBulkLargeField = dlg.format
    Dim Title as String 'Placeholder variable to replace spaces with underscores

    ' write out each group
    Do While grSet.Next()
        gr.Get(grSet.CurrentID)

        sao.NasBulkOn = True
        sao.NasBulkGroupID = gr.ID 'User selected group
        sao.Put(saoID)             'Adds dummy analysis to femap session
        sao.Active = saoID         'Activates dummy analysis set

        Title = gr.title
        Title = Replace(Title," ", "_") 'Replace space with underscore
        grFileName = modelFilePath & Title 'Sets temp file name

        ' create and trim file
        App.feFileWriteNastran(8, grFileName)
        FixFile(grFileName)
        App.feAppMessage(FCM_COMMAND,"Exported to file: " & grFileName & ".bdf")
    Loop

    'Deletes dummy analysis Set after use
    If sao.Deletable(saoID) Then
        sao.Delete(saoID)
    End If

End Sub


Function FixFile(fileName As String)
' scrubs all information not in Bulk Data section of bdf

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

    Dim newFileName As String   ' final file for the user
    newFileName = fileName & ".bdf"

    'Creates new text file to write to
    Set createF = CreateObject("Scripting.FileSystemObject")
    Set writeNew = createF.CreateTextFile(newFileName, True)

    Do While Not readOld.AtEndOfStream  'Loop logic flipped to purposely skip last line
        'Print #2, textLine
        writeNew.WriteLine textLine
        'Line Input #1, textLine
        textLine = readOld.ReadLine
    Loop

    'Close #1  'closing is necessary for deletion
    readOld.Close
    'Close #2  'this file is the final file and will not be deleted
    writeNew.Close

    ' delete old file
    SetAttr fileName, vbNormal
    Kill fileName

End Function
