' Title: [WIP] Results Set Importer.vbs
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 12.0

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

Dim App As femap.Model
Dim resultsSets() As String

Sub ImportResultsSets()
    ' adds a list of selected files to the table
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    Dim fileName As Variant

    With fd
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Femap Neutral Output", "*.fno", 1
        .Filters.Add "Nastran Output", ""
        .Filters.Add "All Files", "*.*", 2
        .Show

        For Each fileName In .SelectedItems
            ActiveCell.Value = fileName
            ActiveCell.Offset(1, 0).Select
        Next
    End With

End Sub

Sub AttachResults()

    Set App = GetObject(, "femap.model")

End Sub
