' Title: Import Multiple Nastran Analysis Results.BAS
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

    App.feAppMessage(FCM_COMMAND, "Import Multiple Nastran Analysis Results")
    App.feAppMessage(FCM_NORMAL, "Import multiple Nastran analysis results in a single command. Accepts OP2 and F06 formats.")

    Dim rc As Long

    Dim objShell
    Dim objFolder
    Dim objFolderItem
    Dim startDir As String
    Dim F$
    Dim results$()
    Dim i As Long
    Dim j As Long

    ' retrieve folder path
    Set objShell = CreateObject("Shell.Application")
    startDir = CurDir$()  ' save starting directory
    Set objFolder = objShell.BrowseForFolder(0, "Select folder containing results.", BIF_EDITBOX, startDir)
    If objFolder Is Nothing Then
        Exit Sub
    Else
        Set objFolderItem = objFolder.Self
        ChDir(objFolderItem.Path)
        Set objFolderItem = Nothing
        Set objFolder = Nothing
        Set objShell = Nothing
    End If

    ' populate results list with F06s
    i = 0
    F$ = Dir$("*.f06")
    Do While F$ <> ""
        ReDim Preserve results$(i)
        results$(i) = F$
        F$ = Dir$()
        i += 1
    Loop

    ' populate results list with OP2s
    F$ = Dir$("*.op2")
    Do While F$ <> ""
        ReDim Preserve results$(i)
        results$(i) = F$
        F$ = Dir$()
        i += 1
    Loop

    ' prompt user to select which files in folder to import
    Begin Dialog UserDialog 100,100,500,378,"Import results from: " & CurDir$(),.DlgFunc ' %GRID:10,7,1,1
        MultiListBox 10,42,480,308,results$(),.Results
        OKButton 170,350,80,21
        CancelButton 280,350,90,21
        CheckBox 10,14,90,14,"Select All",.AllChecked
    End Dialog
    Dim dlg As UserDialog
    Dialog dlg ' show dialog (wait for ok)

    ' import files
    If dlg.AllChecked Then
        For j = 0 To UBound(results$)
            App.feFileReadNastranResults(8, results$(j))
        Next j
    Else
        For j = 0 To UBound(dlg.Results)
            App.feFileReadNastranResults(8, results$(dlg.Results(j)))
        Next j
    End If

    ChDir(startDir)  ' reset working directory to starting directory

    App.feAppUpdatePanes(True) ' reload model info tree

End Sub


'See DialogFunc help topic for more information.
Function DlgFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Select Case Action%
    Case 1 ' Dialog box initialization
        'Beep
    Case 2 ' Value changing or button pressed
        Select Case DlgItem$
            Case "Cancel","Esc"
                End
            Case Else
        End Select
    Case 3 ' TextBox or ComboBox text changed
        rc = 1
    Case 4 ' Focus changed
        DlgFunc = True
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function