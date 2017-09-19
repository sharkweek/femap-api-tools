' Export Groups to BDF.BAS
' modified from PrevGrpBlk.BAS by Ryan Blelloch, Copyright (c) <2017>
' '<ATA Engineering, Inc.> and from 'Export by Group' by A. Pavageau 04/2012
' https://community.plm.automation.siemens.com/t5/3D-Simulation-Femap-Forum/Dealing-with-includes/td-p/309370
'
'-----------------------------------------------------------'
'USER DEFINITIONS
'
'Text editor w/ absolute path
'ex: "C:\WINDOWS\system32\notepad.exe"
Const TEXT_EDITOR = "C:\Program Files\IDM Computer Solutions\UltraEdit\uedit64.exe"
'
'Field format
' 0 = Small Field
' 1 = Large Field (CSys, Material, Prop)
' 2 = Large Field (All but Elements)
' 3 = Large Field
' 4 = Large Field (CSys, Node)

' Only small field format is supported currently
Const FIELD_FORMAT = 0
'-----------------------------------------------------------'

Sub Main
    'FEMAP objects
    'get active femap model
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Export Groups to BDF")

    ' Selected groups to export
    Dim grSet As femap.Set
    Set grSet = App.feSet

    'selected user group
    Dim blkGrp As femap.Group
    Set blkGrp = App.feGroup

    Dim expTitle As String

    'temp analysis set for exporting group bulk data
    Dim tmpAnalysis As femap.AnalysisMgr
    Set tmpAnalysis = App.feAnalysisMgr

    'full path and name of group bulk data file
    Dim absolutePath As String
    ' Used to change absolutePath to something does not already exist
    Dim updatedFilename As String
    Dim fileNum As Long

    'System call to open up text editor
    Dim openFileWithTextEditor As String

    'task ID of text editor shell call
    Dim taskID As Long
    taskID = -1

    'Femap return codes
    Dim rc As femap.zReturnCode

    ' To Current directory
    rc = App.feAppMessageBox(3, "Export to current directory?")
    If rc = FE_FAIL Then
        'retrieve folder path
        Dim objShell
        Dim objFolder
        Dim objFolderItem

        Set objShell = CreateObject("Shell.Application")
        Set objFolder = objShell.BrowseForFolder(0, "Select folder to write file(s)",&H1)
        If objFolder Is Nothing Then
            Exit Sub
        Else
            Set objFolderItem = objFolder.Self
            ChDir(objFolderItem.Path)
        End If

        Set objFolderItem = Nothing
        Set objFolder = Nothing
        Set objShell = Nothing
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    workDir = CurDir$ + "\"

    'Prompt User with User Definition Settings
    App.feAppMessage( FCM_NORMAL, "-----Preview Group Bulk Data-----" )
    App.feAppMessage( FCM_NORMAL, "-(Change settings by editing file)-" )
    App.feAppMessage( FCM_NORMAL, "Text Editor:       " + CStr( TEXT_EDITOR ) )
    App.feAppMessage( FCM_NORMAL, "Output Folder: " + CStr( workDir ) )

    ' prompt user to select groups
    If grSet.SelectMultiID(FT_GROUP,1,"Select groups to export") = FE_CANCEL Then
        Exit Sub
    End If

    'Create dummy analysis set manager
    With tmpAnalysis
        .ID = .NextEmptyID
        .AnalysisType = FAT_STATIC
        .NasBulkRigidThermalExp = True
        .NasExecSkipStandard = True
        .NasBulkSkipStandard = False
        .SkipStandard = True
        .Solver = FAM_NX_NASTRAN
        .NasBulkLargeField = FIELD_FORMAT
        .title = "Export Group"
    End With

    Do While blkGrp.NextInSet(grSet.ID)
        ' Assign group as analysis group for dummy analysis manager
        With tmpAnalysis
            .NasBulkGroupID = blkGrp.ID
            .Put(tmpAnalysis.ID)
            'make sure newly created analysis is active for export
            .Active = tmpAnalysis.ID
        End With

        'remove bad characters from filename
        expTitle = blkGrp.title
        fixFileRoot(expTitle)
        absolutePath = workDir + expTitle

        'change filename until it is not a preexisting file
        fileNum = 0
        updatedFilename = absolutePath + ".bdf"
        While (FileExists(updatedFilename))
            fileNum = fileNum + 1
            updatedFilename = absolutePath
            updatedFilename = updatedFilename + " (" + CStr(fileNum) + ").bdf"
        Wend
        absolutePath = updatedFilename

        'Attempt to export bulk data to file (NX Nastran Format)
        rc = App.feFileWriteNastran( 8, absolutePath )

        'Check for failed export
        If( rc <> FE_OK) Then
            ' If export fails, clean up analysis set and exit
            App.feAppMessage(FCM_ERROR, "Exporting preview to " + absolutePath + " failed")
            deleteAnalysisSet(App, tmpAnalysis)
            Exit Sub
        Else
            ' If export succeeds, clean file of global c-cys's, parameters, etc.
            stripExtras(absolutePath)
            App.feAppMessage(FCM_NORMAL,"'" & absolutePath & "' created...")
        End If

    Loop

    rc = App.feAppMessageBox(2, "Preview file(s)?")
    If rc = FE_OK Then
        If grSet.Count = 1 Then
            'If export was a success and only one group was made, attempt to open file in chosen text editor
            openFileWithTextEditor = """" + TEXT_EDITOR + """ """ + absolutePath + """"
            On Error Resume Next
            taskID = Shell(openFileWithTextEditor, vbNormalFocus)

            'Make sure the preview opened correctly
            If( taskID < 0 ) Then
                App.feAppMessage(FCM_ERROR, "Failed to open preview of bulk data file")
                App.feAppMessage(FCM_ERROR, "Attempted to run: " + openFileWithTextEditor)
            Else
                'Inform user where the file has been made
                App.feAppMessage(FCM_NORMAL, "Preview file created: " + absolutePath)
            End If
        ElseIf grSet.Count > 1 Then
            ' If multiple files were created, open explorer at location of files
            Shell "C:\WINDOWS\explorer.exe" & " " & workDir, vbNormalFocus
        End If
    End If

    'Clean up dummy analysis set
    deleteAnalysisSet(App, tmpAnalysis)

End Sub


Public Function FileExists( ByRef absolutePath As String ) As Boolean
'Check if the file exists. If a file exists but can't be opened, it will assume the file does not exist.

    Dim testFile As Object
    Set testFile = CreateObject("Scripting.FileSystemObject")
    FileExists = testFile.FileExists(absolutePath)

End Function


Public Sub deleteAnalysisSet(ByRef App As femap.model, ByRef tmpAnalysis As femap.AnalysisMgr)
'Thorough deletion of the dummy analysis set. If it fails, prompts user to manually delete the file

    rc = tmpAnalysis.Deletable( tmpAnalysis.ID )
    If( rc = femap.zReturnCode.FE_OK ) Then
        rc = tmpAnalysis.Delete( tmpAnalysis.ID )
    End If

    'If the analysis set is unable to be deleted, prompt the user to delete it themselves
    If(rc <> femap.zReturnCode.FE_OK) Then
                App.feAppMessage( FCM_WARNING, _
                    "Unable to delete dummy analysis set: " + CStr(tmpAnalysis.ID) + ". " + tmpAnalysis.title)
    End If

End Sub


Public Sub fixFileRoot(ByRef fRoot As String)
' Removes invalid characters from file path fRoot

    Dim replaceChar As String
    replaceChar = "-"

    fRoot = Replace( fRoot, "\", replaceChar,1, -1)
    fRoot = Replace( fRoot, "/", replaceChar,1, -1)
    fRoot = Replace( fRoot, ":", replaceChar,1, -1)
    fRoot = Replace( fRoot, "~", replaceChar,1, -1)
    fRoot = Replace( fRoot, "#", replaceChar,1, -1)
    fRoot = Replace( fRoot, "%", replaceChar,1, -1)
    fRoot = Replace( fRoot, "&", replaceChar,1, -1)
    fRoot = Replace( fRoot, "*", replaceChar,1, -1)
    fRoot = Replace( fRoot, "{", replaceChar,1, -1)
    fRoot = Replace( fRoot, "}", replaceChar,1, -1)
    fRoot = Replace( fRoot, "<", replaceChar,1, -1)
    fRoot = Replace( fRoot, ">", replaceChar,1, -1)
    fRoot = Replace( fRoot, "?", replaceChar,1, -1)
    fRoot = Replace( fRoot, "+", replaceChar,1, -1)
    fRoot = Replace( fRoot, "|", replaceChar,1, -1)
    fRoot = Replace( fRoot, """", replaceChar,1, -1)

End Sub


Private Sub stripExtras(file As String)
'Deletes all lines for entities not inclued in group (PARAM, C-Sys, etc.)
    Dim curLine As String
    Dim s1 As String
    Dim l1 As Long
    Dim i As Long

    ' Initialize s1 and l1 for locating Femap default spherical coordinate system
    s1 = "CORD2S         2"
    l1 = 16

    Open file For Input As #1
    Open "temp" For Output As #2



    'look for definition of CSys 2
    Do
        Line Input #1, curLine
    Loop Until Left(curLine,l1) = s1
    Line Input #1, curLine

    Line Input #1, curLine
    While Left(curLine,7) <> "ENDDATA"
        Print #2, curLine
        Line Input #1, curLine
    Wend

    Close #1
    Close #2

    Kill file
    Name "temp" As file
End Sub
