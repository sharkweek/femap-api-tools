' Title: Write-Read Nastran Sets.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2301

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' This scripts uses logic from the "WriteReadSets.BAS" script produced by ATA
' Engineering with the following heading and copyright.

' File: WriteReadSets.BAS
' Version: Version 1
' Date: January 12, 2015
' Author: Ali Gokce, ATA Engineering
' www.ata-plmsoftware.com
' Description: This API reads Nastran set definitions into FEMAP as groups,
'              or writes out Nastran sets
'  from FEMAP groups in the message window.

' Copyright (c) <2008> <ATA Engineering, Inc.>

' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software") to use,
' copy, modify, merge, and/or distribute royalty-free copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to
' the following conditions:
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

' public declarations
Dim App As femap.model
Dim model_dir As String
Dim entity_set As femap.Set

Sub Main
    ' `Main` declarations
    Set App = feFemap()

    Dim model As Huge_
    Dim model_name_full As String
    Dim model_name As String

    App.feAppGetModel(model)
    rc = App.feAppGetModelName(model, model_name_full)
    model_name = Dir(model_name_full)
    model_dir = Left(model_name_full, Len(model_name_full) - Len(model_name))

    App.feAppMessage(FCM_NORMAL,"Write-Read Nastran Sets")

    Begin Dialog UserDialog 790,189,"Read/Write Nastran Sets",.DialogFunc ' %GRID:5,3,0,1
        GroupBox 15,3,120,51,"Action",.action_groupbox
            OptionGroup .action
                OptionButton 25,18,90,15,"Write",.action_write
                OptionButton 25,33,90,15,"Read",.action_read
        GroupBox 10,57,125,51,"Entity",.entities_groupbox
            OptionGroup .entity
                OptionButton 30,72,90,15,"Nodes",.nos
                OptionButton 30,87,90,15,"Elements",.elements
        GroupBox 145,3,630,39,"Write Parameters",.writeparameters_groupbox
            TextBox 690,18,70,15,.setID
            Text 630,21,50,15,"Set ID:",.setid_text,1
            Text 160,21,75,15,"File name:",.file_name_text,1
            TextBox 245,18,370,15,.file_name
        GroupBox 150,45,625,111,"Read parameters",.read_parameters_groupbox
            TextBox 165,75,600,75,.set_text,1
            OptionGroup .read_from
                OptionButton 165,57,90,15,"From text",.read_from_text
                OptionButton 270,57,90,15,"From file",.read_from_file
        OKButton 15,114,120,21
        CancelButton 15,135,120,21
    End Dialog
    Dim dlg As UserDialog

    ' prompt user
    If Dialog(dlg) = 0 Then
        Exit Sub  ' exit if user cancels
    End If

    Set entity_set = App.feSet()

    Dim file_path As String
    Dim entity_type As Long

    Select Case dlg.action
        Case 0  ' Write
            Dim out_string As String
            Dim entity_string As String
            Select Case dlg.entity
                Case 0  ' nodes
                    entity_type = FT_NODE
                    entity_string = "nodes"
                Case 1  ' elements
                    entity_type = FT_ELEM
                    entity_string = "elements"
            End Select

            ' prompt user to select entities
            rc = entity_set.Select(entity_type, True, "Select " & entity_string & "...")
            If rc = FE_CANCEL Then
                Exit Sub
            ElseIf rc = FE_NOT_EXIST Then
                App.feAppMessage(FCM_ERROR, "No " & entity_string & " selected. Routine canceled...")
            End If

            out_string = SetToString(entity_set, dlg.setID)
            file_path = model_dir & dlg.file_name

            Dim fso As Object
            Dim tf As Object
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set tf = fso.CreateTextFile("D:\loft\set-1.bdf", True)
            tf.write(out_string)
            tf.Close

        Case 1  ' Read
            Dim set_string As String
            Select Case dlg.read_from
                Case 0  ' from text
                    StringToGroup(dlg.set_text, entity_type)
                Case 1  ' from file
                    file_path = GetFilePath(,"Bulk Data File|*.bdf|Any|*.*", "Pick file...", 0)
                    set_string = StringFromFile(file_path)
                    StringToGroup(set_string, dlg.entity)
            End Select
    End Select
End Sub


Function SetToString(entity_set, set_id)
    Dim line As String
    Dim out_string As String
    Dim add_str As String
    Dim pointer As Long
    Dim begin_id As Long

    ' initialize first string
    entity_set.Reset()
    pointer = entity_set.First()
    begin_id = pointer
    line = "SET " & set_id & "= "
    add_str = "" & begin_id

    Do While entity_set.NextID()
        If entity_set.NextID > pointer + 1 Then
            If pointer = begin_id + 1 Then
                add_str = add_str & ", " & pointer
            ElseIf pointer > begin_id Then
                add_str = add_str & " THRU " & pointer
            End If

            begin_id = entity_set.NextID()
            If Len(line) + Len(add_str) < 72 Then
                line = line & add_str & ", "
            Else
                out_string = out_string & RTrim(line) & vbCRLF
                line = "    " & add_str & ", "
            End If
            add_str = "" & begin_id
        End If

        pointer = entity_set.Next()
    Loop

    '  last ID
    If pointer = begin_id + 1 Then
        add_str = add_str & ", " & pointer
    ElseIf pointer > begin_id Then
        add_str = add_str & " THRU " & pointer
    End If

    If Len(line) + Len(add_str) < 72 Then
        line = line & add_str
    Else
        out_string = out_string & line & vbCRLF
        line = "    " & add_str
    End If

    out_string = out_string & line & vbCRLF

    SetToString = out_string
End Function


Sub StringToGroup(set_string As String, entity_type)
    Dim str_array As Variant
    Dim set_str_curr As String
    Dim curr_text As String
    Dim grp_id As Long
    Dim ind As Long
    Dim str_len As Long
    Dim idx1 As Long
    Dim idx2 As Long
    Dim idx_eq As Long
    Dim start_idx As long
    Dim end_idx As Long
    Dim set_id As Long
    Dim thru_idx As Long
    Dim type_str As String
    Dim idx As Long

    Dim grp As femap.Group
    Set grp = App.feGroup()

    set_string = Replace(UCase$(set_string), vbNewLine, "")
    idx = 1
    str_len = Len(set_string)
    Do While idx < str_len
        entity_set.Clear
        grp_id = grp.NextEmptyID
        grp.ID = grp_id
        type_str = ""
        idx1 = InStr(idx, set_string, "SET")
        idx_eq = InStr(idx, set_string, " = ")
        set_id = Val(Trim(Mid(set_string, idx1 + 3, idx_eq - idx1 - 3)))
        idx2= InStr(idx_eq+1, set_string, "SET")

        If IsNull(idx2) Or idx2 = 0 Then
            idx2 = str_len+1
        End If

        set_str_curr = Mid$(set_string, idx_eq + 2, idx2 - idx_eq - 2)
        str_array = Split(set_str_curr, ",")

        For i = 0 To UBound(str_array)
            curr_text = Trim(str_array(i))
            thru_idx = InStr(1, curr_text, "THRU")
            If IsNull(thru_idx) Or thru_idx = 0 Then
                start_idx = Val(curr_text)
                If start_idx > 0 Then
                    entity_set.Add(start_idx)
                End If
            Else
                start_idx = Val(Trim(Split(curr_text, "THRU")(0)))
                end_idx= Val(Trim(Split(curr_text, "THRU")(1)))
                If start_idx > 0 And end_idx > 0 Then
                    entity_set.AddRange(start_idx, end_idx, 1)
                End If
            End If
        Next i

        grp.SetAdd(type_str, entity_set.ID)
        grp.Put(grp_id)

        idx = idx2
    Loop
End Sub


Function StringFromFile(file_path as String)
    Dim reader As femap.Read
    Set reader = App.feRead()

    Dim file_string As String
    file_string = ""

    With reader
        .Open(file_path, 80)
        .SetTrimSpace(True)
        .SetNoBlankLines(True)
        .ShowStatus(True)
        Do Until .AtEOF()
            .Read()
            file_string = file_string & .line & vbCrLf
        Loop
    End With

    StringFromFile = file_string

End Function


Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Select Case Action%
        Case 1 ' Dialog box initialization
            ' write parameters
            DlgText "setID", "1"
            DlgText "file_name", "set-1.bdf"
            ' read parameters
            DlgText "set_text", ""
            DlgEnable "set_text", False
            DlgEnable "read_from_file", False
            DlgEnable "read_from_text", False
        Case 2 ' Value changing or button pressed
            Select Case DlgItem
                Case "action"
                    Select Case SuppValue
                        Case 0  ' write
                            'write group
                            DlgEnable "setID", True
                            DlgEnable "setid_text", True
                            DlgEnable "file_name", True
                            DlgEnable "file_name_text", True

                            'read group
                            DlgEnable "set_text", False
                            DlgEnable "read_from_file", False
                            DlgEnable "read_from_text", False

                        Case 1  ' read
                            'write group
                            DlgEnable "setID", False
                            DlgEnable "setid_text", False
                            DlgEnable "file_name", False
                            DlgEnable "file_name_text", False

                            'read group
                            DlgEnable "read_from_file", True
                            DlgEnable "read_from_text", True
                            If DlgValue("read_from") = 0 Then
                                DlgEnable "set_text", True
                            Else
                                DlgEnable "set_text", False
                            End If

                    End Select

                    DialogFunc = True

                Case "read_from"
                    Select Case SuppValue
                        Case 0  ' from text
                            DlgEnable "set_text", True
                        Case 1  ' from file
                            DlgEnable "set_text", False
                    End Select

                    DialogFunc = True
            End Select
        Case 3 ' TextBox or ComboBox text changed
        Case 4 ' Focus changed
        Case 5 ' Idle
        Case 6 ' Function key
    End Select
End Function
