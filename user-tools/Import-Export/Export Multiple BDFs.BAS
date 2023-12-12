' Title: Export Multiple BDFs.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2301

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

Dim App As femap.model
Dim out_dir As String
Dim new_text As String
Dim model_dir As String

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Export Multiple BDFs")
    App.feAppMessage(FCM_NORMAL, "Batch export multiple analysis sets as Nastran *.bdfs.")

    Dim rc As Long

    Dim deck_set As femap.Set
    Set deck_set = App.feSet

    Dim deck As femap.AnalysisMgr
    Set deck = App.feAnalysisMgr
    Dim deck_name as String 'Placeholder variable to replace spaces with underscores
    Dim deckFormat As Integer  'field format for write deck

    Dim empty_group As femap.Group
    Set empty_group = App.feGroup

    Dim analysis_group_id As Long
    Dim init_deck_id As Long

    Dim bc_sets As femap.Set
    Set bc_sets = App.feSet()

    Dim model As Huge_
    Dim model_name_full As String
    Dim model_name As String
    Dim file_path As String

    'Get model name and path
    App.feAppGetModel(model)
    rc = App.feAppGetModelName(model, model_name_full)
    model_name = Dir(model_name_full)
    model_dir = Left(model_name_full, Len(model_name_full) - Len(model_name))

    rc = deck_set.SelectMultiID(FT_AMGR_DIR, 1, "Select analysis sets to export...")
    If rc = FE_CANCEL Then
        App.feAppMessage(FCM_NORMAL, "Export canceled...")
        Exit Sub
    End If

    Begin Dialog UserDialog 400,112,"Output Options",.DialogFunc ' %GRID:10,7,1,1
        GroupBox 6,3,384,80,"Location",.location_options_box               '0
        OptionGroup .location                                              '1
            OptionButton 20,17,200,14,"Direct output directories",.direct  '2
            OptionButton 20,31,160,14,"Model directory",.model             '3
            OptionButton 20,45,200,14,"Specify location",.user             '4
        TextBox 20,63,300,14,.user_dir                                     '5
        CheckBox 20,91,140,14,"Empty Bulk Data",.empty_bulk                '6
        PushButton 330,62,50,16,"Path",.get_path                           '7
        OKButton 210,87,90,21                                              '8
        CancelButton 300,87,90,21                                          '9
    End Dialog
    Dim dlg As UserDialog

    ' prompt user
    If Dialog(dlg) = 0 Then
        Exit Sub  ' exit if user cancels
    End If

    ' create empty group if "Empty Bulk Data" is selected
    If dlg.empty_bulk Then
        empty_group.Put(empty_group.NextEmptyID())
    End If

    ' write procedure based on output options
    Select Case dlg.location
        Case 0  ' direct
            init_deck_id = deck.Active
            deck_set.Reset()
            Do While deck_set.Next()  ' output each deck to the output directory spec'd by the analysis manager
                deck.Get(deck_set.CurrentID)
                deck.Active = deck.ID
                deck_name = SanitizeFilename(deck.NasExecAnalyzeFilename)
                file_path = deck.NasExecOutDir & "\" & deck_name

                ' for assign empty group for empty_bulk
                If dlg.empty_bulk Then
                    analysis_group_id = deck.NasBulkGroupID
                    deck.NasBulkGroupID = empty_group.ID
                    deck.Put(deck.ID)
                End If

                ' write and report result
                If App.feFileWriteNastran(8, file_path) = FE_OK Then
                    App.feAppMessage(FCM_COMMAND, deck.Title & " exported as " & file_path)
                Else
                    App.feAppMessage(FCM_ERROR, deck.Title & " failed to export.")
                End If

                ' reset original group for empty_bulk
                If dlg.empty_bulk Then
                    deck.NasBulkGroupID = analysis_group_id
                    deck.Put(deck.ID)
                End If
            Loop

            deck.Active = init_deck_id

        Case 1  ' model
            deck_set.Reset()
            Do While deck_set.Next()  ' output all decks to the model parent directory
                deck.Get(deck_set.CurrentID)
                deck.Active = deck.ID
                deck_name = SanitizeFilename(deck.NasExecAnalyzeFilename)
                file_path = model_dir & "\" & deck_name

                ' for assign empty group for empty_bulk
                If dlg.empty_bulk Then
                    analysis_group_id = deck.NasBulkGroupID
                    deck.NasBulkGroupID = empty_group.ID
                    deck.Put(deck.ID)
                End If

                ' write and report result
                If App.feFileWriteNastran(8, file_path) = FE_OK Then
                    App.feAppMessage(FCM_COMMAND, deck.Title & " exported as " & file_path)
                Else
                    App.feAppMessage(FCM_ERROR, deck.Title & " failed to export.")
                End If

                ' reset original group for empty_bulk
                If dlg.empty_bulk Then
                    deck.NasBulkGroupID = analysis_group_id
                    deck.Put(deck.ID)
                End If
            Loop

            deck.Active = init_deck_id

        Case 2  ' custom
            deck_set.Reset()
            Do While deck_set.Next()  ' output each deck to the output directory spec'd by the analysis manager
                deck.Get(deck_set.CurrentID)
                deck.Active = deck.ID
                deck_name = SanitizeFilename(deck.title)
                file_path = dlg.user_dir & "\" & deck_name

                ' for assign empty group for empty_bulk
                If dlg.empty_bulk Then
                    analysis_group_id = deck.NasBulkGroupID
                    deck.NasBulkGroupID = empty_group.ID
                    deck.Put(deck.ID)
                End If

                ' write and report result
                If App.feFileWriteNastran(8, file_path) = FE_OK Then
                    App.feAppMessage(FCM_COMMAND, deck.Title & " exported as " & file_path)
                Else
                    App.feAppMessage(FCM_ERROR, deck.Title & " failed to export.")
                End If

                ' reset original group for empty_bulk
                If dlg.empty_bulk Then
                    deck.NasBulkGroupID = analysis_group_id
                    deck.Put(deck.ID)
                End If
            Loop

            deck.Active = init_deck_id
    End Select

    'Delete empty group
    empty_group.Delete(empty_group.ID)

End Sub


Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Select Case Action%
        Case 1 ' Dialog box initialization
            DlgText "user_dir", ""
            DlgEnable "user_dir", False
            DlgEnable "get_path", False
            DlgText "user_dir", model_dir
        Case 2 ' Value changing or button pressed
            Select Case DlgItem
                Case "get_path"
                    Dim objShell As Object
                    Set objShell  = CreateObject("Shell.Application")
                    Set objFolder = objShell.BrowseForFolder(0, "Pick Folder", &H100, )
                    If objFolder Is Nothing Then
                        DialogFunc = True
                        Exit Function
                    End If
                    DlgText "user_dir", objFolder.self.path
                    DialogFunc = True
                Case "location"
                    Select Case SuppValue
                        Case 0, 1  ' model paths
                            DlgEnable "user_dir", False
                            DlgEnable "get_path", False
                        Case 2  ' custom path
                            DlgEnable "user_dir", True
                            DlgEnable "get_path", True
                    End Select
                    DialogFunc = True
                Case "user_dir"
                    If Right(DlgValue("user_dir"), 1) = "\" Then
                        new_text = Left(DlgText("user_dir"), Len(DlgText("user_dir")) - 1)
                        DlgText "user_dir", new_text
                    End If
                    DialogFunc = True
            End Select
        Case 3 ' TextBox or ComboBox text changed
        Case 4 ' Focus changed
        Case 5 ' Idle
        Case 6 ' Function key
    End Select
End Function


Private Function SanitizeFilename(filename As String) As String

    Dim s_filename As String

    s_filename = Replace(filename," ", "-")
    s_filename = Replace(s_filename,"/", "-")
    s_filename = Replace(s_filename,"\", "-")
    s_filename = Replace(s_filename,":", "-")
    s_filename = Replace(s_filename,"*", "-")
    s_filename = Replace(s_filename,"%", "percent")

    SanitizeFilename = s_filename & ".dat"

End Function
