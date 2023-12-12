' Title: Load Case Renumberer.vbs
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 12.0

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

'Connect to FEMAP
Dim App As femap.Model
Dim caseTable As ListObject
Dim r As ListRow
Dim lc As femap.OutputSet
Dim rc As Long
Dim oldIDColumn As Range
Dim oldTitleColumn As Range
Dim newIDColumn As Range
Dim newTitleColumn As Range
Dim allOn As Boolean


Sub GetLoadCases()

    ' Local variables
    Dim load_cases As femap.Set

    ' Set global variable cases
    Set App = GetObject(, "femap.model")
    Set load_cases = App.feSet
    Set caseTable = ActiveSheet.ListObjects("LoadCases")
    Set lc = App.feOutputSet

    ' Prompt user to select load cases to pull from Femap
    rc = load_cases.Select(FT_OUT_CASE, True, "Select output sets...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_FAIL Then
        rc = App.feAppMessage(FCM_ERROR, "Selected output sets do not exist.")
    End If

    ' Populate first two columns with Femap data
    caseTable.DataBodyRange(1, 1).Select
    Do While lc.NextInSet(load_cases.ID)
        ActiveCell.Value = lc.ID
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = lc.Title
        ActiveCell.Offset(1, -1).Select
    Loop

End Sub

Sub RenameLoadCases()
    ' Rename and renumber load cases in Femap model based on user-filled data

    ' Local variables
    Dim rc_dummy As Long  ' second rc is needed because Excel requires an `=` when using App methods
    Dim oldID As Long
    Dim newID As Long
    Dim nextID As Long
    Dim newTitle As String

    ' Set global variable references
    Set App = GetObject(, "femap.model")
    Set lc = App.feOutputSet
    Set caseTable = ActiveSheet.ListObjects("LoadCases")
    Set oldIDColumn = caseTable.ListColumns("Existing LCID").Range
    Set newIDColumn = caseTable.ListColumns("New LCID").Range

    ' loop through table and renumber each load case
    For Each r In caseTable.ListRows
        oldID = CLng(Intersect(r.Range, oldIDColumn).Value)
        newID = CLng(Intersect(r.Range, newIDColumn).Value)

        ' get load case data from model
        If lc.Get(oldID) = FE_FAIL Then
            rc = App.feAppMessage(FCM_ERROR, "Load case " & CStr(oldID) & " does not exist in model. Correct and re-run.")
            Exit Sub

        ' If new ID already exists in model, prompt user for how to proceed
        ElseIf lc.Exist(newID) Then
            lc.Title = newTitle
            rc = MsgBox("Load case " & CStr(newID) & " already exists in model. Do you wish to overwrite?", vbYesNoCancel)

            ' overwrite
            If rc = vbYes Then
                If App.feRenumber(FT_OUT_CASE, -oldID, newID) = FE_OK Then
                    rc_dummy = App.feAppMessage(FCM_NORMAL, "Results set " & CStr(oldID) & " renumbered to " & CStr(newID))
                End If


            ' use next empty ID
            ElseIf rc = vbNo Then
                nextID = lc.NextEmptyID
                rc_dummy = App.feAppMessage(FCM_HIGHLIGHT, "Load case " & CStr(newID) & " imported as load case " & CStr(nextID))
                If App.feRenumber(FT_OUT_CASE, -oldID, nextID) = FE_OK Then
                    rc_dummy = App.feAppMessage(FCM_NORMAL, "Results set " & CStr(oldID) & " renumbered to " & CStr(nextID))
                End If

            ' cancel sub
            ElseIf rc = vbCancel Then
                Exit Sub

            End If

        ' if new ID does not yet exist, add to model and delete old one
        Else
            If App.feRenumber(FT_OUT_CASE, -oldID, newID) = FE_OK Then
                rc_dummy = App.feAppMessage(FCM_NORMAL, "Results set " & CStr(oldID) & " renumbered to " & CStr(newID))
            End If
        End If

    Next r

End Sub
