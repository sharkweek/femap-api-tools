' Modified from WriteGroupsToNastranFiles_r2.BAS by ATA Engineering. Original file info below.
    ' File: WriteGroupsToNastranFiles_r2.BAS
    ' Version: Version 2
    ' Date: January 12, 2015
    ' Author: Ryan Blelloch, Michael Palodichuk, ATA Engineering
    ' www.ata-plmsoftware.com
    ' Description: This API allows user to select one or more FEMAP groups and exports their bulk data to individual
    ' Nastran INCLUDE files.
    '
    'Usage Notes
    '====================================
    ' The active FEMAP model is automatically saved upon running this API
    ' The files written out are put in the same directory as the FEMAP model
    ' The files are stripped of extraneous data including the basic coordinate system definitions and the ENDDATA tag.
    '
    'Copyright (c) <2015> <ATA Engineering, Inc.>
    '
    'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
    'associated documentation files (the "Software") to use, copy, modify, merge, and/or distribute
    'royalty-free copies of the Software, and to permit persons to whom the Software is furnished
    'to do so, subject to the following conditions:
    'The above copyright notice and this permission notice shall be included in all copies or
    'substantial portions of the Software.
    'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
    'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
    'IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
    'WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
    'OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

'-----------------------------------------------------------'
'USER DEFINITIONS
'
'Text editor w/ absolute path
'ex: "C:\WINDOWS\system32\notepad.exe"
Const TEXT_EDITOR = "C:\Program Files\IDM Computer Solutions\UltraEdit\uedit64.exe"
'
'-----------------------------------------------------------'
'
Sub Main

    'Load model
    Dim App As femap.model
    Set App = feFemap()

    'Indicate command has begun
    App.feAppMessage(FCM_COMMAND, "Export Groups to BDF")

    'Get model name and path
    Dim model As Huge_
    Dim modelFileFull As String
    Dim modelFileName As String
    Dim modelFilePath As String
    App.feAppGetModel(model)
    rc = App.feAppGetModelName(model,modelFileFull)
    modelFileName = Dir(modelFileFull)
    modelFilePath = Left(modelFileFull, Len(modelFileFull)-Len(modelFileName))

    ' Working directory
    Dim workDir As String
    workDir = CurDir$ + "\"

    'Make group set object
    Dim grSet As femap.Set
    Set grSet =App.feSet

    'System call to open up text editor
    Dim openFileWithTextEditor As String

    'Make group object
    Dim gr As femap.Group
    Set gr = App.feGroup
    Dim grName As String
    Dim grFileName As String

    'task ID of text editor shell call
    Dim taskID As Long
    taskID = -1

    'Make analysis set object
    Dim sao As femap.AnalysisMgr
    Set sao = App.feAnalysisMgr
    Dim saoID As Long

    'Set properties of analysis set
    saoID = sao.NextEmptyID   'Finds next avaliable ID to make analysis set in
    sao.title = "Dummy Set for Group Export"     'Names dummy set
    sao.Solver = 36        'Sets Solver to NX Nastran
    sao.AnalysisType = 2     'Sets AnalysisType to Modes for creation of less data
    sao.NasCnlSkipNLPARM = True   'Skips more unneeded data to be written out
    sao.SkipStandard = True   'Skips more unneeded data to be written out

    'All field format types stored in strings
    Dim FFormat(5) As String
    FFormat(0) ="Small Field"
    FFormat(1) ="Large Csys, Material, Property"
    FFormat(2) ="Large All But Elements"
    FFormat(3) ="Large"
    FFormat(4) ="Large Csys, Node"


    'Prompt user for groups to export
    If grSet.SelectMultiID(FT_GROUP, 1, "Select Group(s) to Write to Nastran File(s)") = FE_CANCEL Then
        App.feAppMessage(FCM_NORMAL,"Canceled...")
        End 'terminates macro if user cancels
    End If

    'Prompts user for field format
    Dim UserFormat As Integer

    Begin Dialog UserDialog 280,126,"Select Field Format..." 'NasBulkLargeField %GRID:10,7,1,1
        ListBox 20,14,240,70,FFormat(),.ListBox1
        OKButton 20,91,110,28
        CancelButton 150,91,110,28
    End Dialog
    Dim dlg As UserDialog
    UserFormat = Dialog(dlg)

    If UserFormat = -1 Then   'user pressed okay
        UserFormat = dlg.ListBox1
    ElseIf UserFormat = 0 Then 'terminates macro if user cancels
        App.feAppMessage(FCM_NORMAL,"Command canceled...")
        End
    End If

    'Updates field format
    sao.NasBulkLargeField = UserFormat
    Dim Title as String 'Placeholder variable to replace spaces with underscores
    'Loop reads out all group information seperately
    While gr.NextInSet( grSet.ID ) = FE_OK

        sao.NasBulkOn = True
        sao.NasBulkGroupID = gr.ID 'User selected group
        sao.Put (saoID)    'Adds dummy analysis to femap session
        sao.Active = saoID  'Activates dummy analysis set
        Title = gr.title
        fixGroupName(Title)
        grFileName = modelFilePath & Title 'Sets temp file name

        App.feAppMessage(FCM_COMMAND,"Exported to file: " & grFileName & ".bdf")
        App.feFileWriteNastran(8,grFileName)

        FixFile(grFileName) 'Function that writes correct data to new final file
        DeleteTempFile(grFileName)  'Function that deletes the temporary file

    Wend

    'Deletes dummy analysis Set after use
    If sao.Deletable(saoID) Then
           sao.Delete(saoID)
    End If

    ' Preview files
    If grSet.Count = 1 Then
        If App.feAppMessageBox(2, "Preview file(s)?") Then
            'If export was a success and only one group was made, attempt to open file in chosen text editor
            openFileWithTextEditor = """" + TEXT_EDITOR + """ """ + grFileName + """"
            On Error Resume Next
            taskID = Shell(openFileWithTextEditor, vbNormalFocus)

            'Make sure the preview opened correctly
            If( taskID < 0 ) Then
                App.feAppMessage(FCM_ERROR, "Failed to open preview of bulk data file")
                App.feAppMessage(FCM_ERROR, "Attempted to run: " + openFileWithTextEditor)
            End If
        End If
    ElseIf grSet.Count > 1 and App.feAppMessageBox(2, "Show files in Explorer?") Then
        ' If multiple files were created, open explorer at location of files
        Shell "C:\WINDOWS\explorer.exe" & " " & workDir, vbNormalFocus
    End If

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


Public Sub fixGroupName(ByRef gName As String)
' Removes invalid characters from file path gName

    Dim replaceChar As String
    replaceChar = "-"

    gName = Replace( gName, "\", replaceChar,1, -1)
    gName = Replace( gName, "/", replaceChar,1, -1)
    gName = Replace( gName, ":", replaceChar,1, -1)
    gName = Replace( gName, "~", replaceChar,1, -1)
    gName = Replace( gName, "#", replaceChar,1, -1)
    gName = Replace( gName, "%", replaceChar,1, -1)
    gName = Replace( gName, "&", replaceChar,1, -1)
    gName = Replace( gName, "*", replaceChar,1, -1)
    gName = Replace( gName, "{", replaceChar,1, -1)
    gName = Replace( gName, "}", replaceChar,1, -1)
    gName = Replace( gName, "<", replaceChar,1, -1)
    gName = Replace( gName, ">", replaceChar,1, -1)
    gName = Replace( gName, "?", replaceChar,1, -1)
    gName = Replace( gName, "+", replaceChar,1, -1)
    gName = Replace( gName, "|", replaceChar,1, -1)
    gName = Replace( gName, """", replaceChar,1, -1)

End Sub


Function DeleteTempFile(FileToDelete As String)

    SetAttr FileToDelete, vbNormal  'Makes sure it is not in an undeletable view
    Kill FileToDelete   'Deletes file

End Function
