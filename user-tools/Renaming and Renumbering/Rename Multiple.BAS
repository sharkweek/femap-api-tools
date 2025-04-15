' Title: Rename Multiple.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2022.1

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' Usage Notes
' -----------
' "Replace" feature supports regular expression pattern matching. Regex patterns
' follow the VBScript spec (https://www.tutorialspoint.com/vbscript/vbscript_reg_expressions.htm)

' TODO: Regex replacement still not working

' module scope variables
Dim App As femap.model
Dim obj as Object
Dim select_prompt As String
Dim entity_type As Long
Dim select_multi_id As Boolean
Dim new_val As Long

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Rename Multiple")

    Dim rc As Long

    Dim objSet As femap.Set
    Set objSet = App.feSet

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    Dim objid As Long
    Dim obj_count As Long
    Dim prev_name As String     'existing name
    Dim index As Long
    Dim step As Long
    Dim zero_format As String
    Dim success As Long

    ' arrays for dropdown box
    Dim entity_names(24) As String
    entity_names(0) = "1..Coordinate Systems"
    entity_names(1) = "2..Solids"
    entity_names(2) = "3..Connection Properties"
    entity_names(3) = "4..Connection Regions"
    entity_names(4) = "5..Connectors"
    entity_names(5) = "6..Materials"
    entity_names(6) = "7..Properties"
    entity_names(7) = "8..Layups"
    entity_names(8) = "9..Load Sets"
    entity_names(9) = "10..Constraint Sets"
    entity_names(10) = "11..Functions"
    entity_names(11) = "12..Tables"
    entity_names(12) = "13..Data Surfaces"
    entity_names(13) = "14..Analysis Sets"
    entity_names(14) = "15..Analysis Studies"
    entity_names(15) = "16..Output Sets"
    entity_names(16) = "17..Views"
    entity_names(17) = "18..Groups"
    entity_names(18) = "19..Layers"
    entity_names(19) = "20..Freebodies"
    entity_names(20) = "21..Charts"
    entity_names(21) = "22..Chart Data Series"
    entity_names(22) = "23..Planes"
    entity_names(23) = "24..Vectors"

    ' number of leading zeros for indexing
    Dim zero_arr(8) As String
    For i = 0 To 7
        zero_arr(i) = CStr(i)
    Next

    'User inputs
    Begin Dialog UserDialog 320,305,"Rename Entities...",.DialogFunc ' %GRID:10,5,1,1
        Text 20,5,90,15,"Entity Type:",.entity_type_text
        DropListBox 20,20,280,20,entity_names(),.entity_type
        Text 20,245,260,15,"New name:",.title_text
        TextBox 20,260,280,15,.new_name
        GroupBox 20,40,140,65,"Rename Type",.rename_groupbox
        OptionGroup .rename_option
            OptionButton 40,55,90,15,"Overwrite",.overwrite
            OptionButton 40,70,90,15,"Append",.append
            OptionButton 40,85,90,15,"Replace",.find_replace
        GroupBox 170,40,130,65,"Append Options",.append_groupbox
        OptionGroup .append_option
            OptionButton 190,55,70,15,"Prefix",.prefix
            OptionButton 190,70,70,15,"Suffix",.suffix
        Text 20,190,280,15,"Text to Replace:",.replace_text_text
        TextBox 20,205,280,15,.replace_text
        CheckBox 40,225,200,15,"Use regex",.use_regex
        GroupBox 20,105,280,85,"Indexing",.idx_groupbox
        CheckBox 30,125,90,15,"Index",.indexed
        TextBox 30,165,110,15,.idx_start
        TextBox 170,165,110,15,.idx_step
        Text 30,150,60,15,"Start ID:",.idx_start_text
        Text 170,150,50,15,"Step:",.idx_step_text
        Text 170,115,100,15,"Leading zeros:",.idx_zeros_text
        DropListBox 170,130,110,15,zero_arr(),.idx_zeros
        OKButton 50,280,100,20
        CancelButton 170,280,100,20
    End Dialog
    Dim dlg As UserDialog

    ' prompt user
    If Dialog(dlg) = 0 Then
        Exit Sub  ' exit if user cancels
    End If

    ' set the entity type using the defined function below
    SetEntityType(dlg.entity_type)

    'prompt user to select entities for renaming
    If select_multi Then
        rc = objSet.SelectMultiIDV2(entity_type, 1, select_prompt)
    Else
        rc = objSet.Select(entity_type, True, select_prompt)
    End If

    ' rc = objSet.Select(entity_type, True, select_prompt )
    If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
        App.feAppMessage(FCM_ERROR, "Nothing selected...")
        Exit Sub
    End If

    ' leading zeros format string
    If dlg.idx_zeros = 0 Or dlg.indexed = False Then
        zero_format = "0"
    Else
        zero_format = ""
        For i = 0 To dlg.idx_zeros
            zero_format = zero_format & "0"
        Next i
    End If

    ' setup for iterations
    objSet.Reset()
    objid = objSet.First()
    index = CLng(dlg.idx_start)  ' set first index value
    obj_count = objSet.Count()
    succes = 0

    'rename propect according to user's selections
    ' note: code intentionally uses conditionals first to partition multiple
    '       do-loops to avoid loops with repetetive conditional evaluations
    '       that would slow things down.
    Select Case dlg.rename_option
        Case 0  'overwrite
            ' with indices
            If dlg.indexed Then
                Do While obj.NextInSet(objSet.ID)
                    obj.Get(obj.ID)
                    If dlg.append_option Then
                        obj.title = Format(index, zero_format) + dlg.new_name
                    Else
                        obj.title = dlg.new_name & Format(index, zero_format)
                    End If
                    If obj.Put(obj.ID) = FE_OK Then
                        success += 1
                    End If
                    index += CLng(dlg.idx_step)
                Loop
            ' no indices
            Else
                Do While obj.NextInSet(objSet.ID)
                    obj.Get(obj.ID)
                    obj.title = dlg.new_name
                    If obj.Put(obj.ID) = FE_OK Then
                        success += 1
                    End If
                Loop
            End If

        Case 1  'append
            ' prefix option
            If dlg.append_option = 0 Then
                ' with indices
                If dlg.indexed Then
                    Do While obj.NextInSet(objSet.ID)
                        obj.Get(obj.ID)
                        prev_name = obj.title
                        obj.title = dlg.new_name & Format(index, zero_format) & prev_name
                        If obj.Put(obj.ID) = FE_OK Then
                            success += 1
                        End If
                        index += CLng(dlg.idx_step)
                    Loop
                ' no indices
                Else
                    Do While obj.NextInSet(objSet.ID)
                        obj.Get(obj.ID)
                        prev_name = obj.title
                        obj.title = dlg.new_name & prev_name
                        If obj.Put(obj.ID) = FE_OK Then
                            success += 1
                        End If
                    Loop
                End If
            ' suffix option
            Else
                ' with indices
                If dlg.indexed Then
                    Do While obj.NextInSet(objSet.ID)
                        obj.Get(obj.ID)
                        prev_name = obj.title
                        obj.title = prev_name & dlg.new_name & Format(index, zero_format)
                        If obj.Put(obj.ID) = FE_OK Then
                            success += 1
                        End If
                        index += CLng(dlg.idx_step)
                    Loop
                ' no indices
                Else
                    Do While obj.NextInSet(objSet.ID)
                        obj.Get(obj.ID)
                        prev_name = obj.title
                        obj.title = prev_name & dlg.new_name
                        If obj.Put(obj.ID) = FE_OK Then
                            success += 1
                        End If
                    Loop
                End If
            End If

        ' find and replace
        Case 2
            ' don't use regular expressions
            If dlg.use_regex Then
            ' use regular expressions
                ' with indices
                If dlg.indexed Then
                    Do While obj.NextInSet(objSet.ID)
                        obj.Get(obj.ID)
                        prev_name = obj.title
                        regex.Pattern = dlg.replace_text
                        obj.title = regex.Replace(prev_name, dlg.new_name & Format(index, zero_format))
                        If obj.Put(obj.ID) = FE_OK Then
                            success += 1
                        End If
                        index += CLng(dlg.idx_step)
                    Loop
                ' no indices
                Else
                    Do While obj.NextInSet(objSet.ID)
                        obj.Get(obj.ID)
                        prev_name = obj.title
                        regex.Pattern = dlg.replace_text
                        obj.title = regex.Replace(prev_name, dlg.new_name)
                        If obj.Put(obj.ID) = FE_OK Then
                            success += 1
                        End If
                    Loop
                End If
            Else
                ' with indices
                If dlg.indexed Then
                    Do While obj.NextInSet(objSet.ID)
                        obj.Get(obj.ID)
                        prev_name = obj.title
                        obj.title = Replace(prev_name, dlg.replace_text & Format(index, zero_format), dlg.new_name)
                        If obj.Put(obj.ID) = FE_OK Then
                            success += 1
                        End If
                        index += CLng(dlg.idx_step)
                    Loop
                ' no indices
                Else
                    Do While obj.NextInSet(objSet.ID)
                        obj.Get(obj.ID)
                        prev_name = obj.title
                        obj.title = Replace(prev_name, dlg.replace_text, dlg.new_name)
                        If obj.Put(obj.ID) = FE_OK Then
                            success += 1
                        End If
                    Loop
                End If
            End If
    End Select

    App.feViewRegenerate (0)
    App.feAppMessage(FCM_NORMAL, CStr(success) & " entities renamed...")

End Sub


Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Select Case Action%
        Case 1  ' dialog box initialization
            DlgEnable "append_option", False
            DlgEnable "replace_text", False
            DlgText "idx_step", "1"
            DlgText "idx_start", "1"
            DlgEnable "indexed", True
            DlgEnable "idx_step", False
            DlgEnable "idx_step_text", False
            DlgEnable "idx_start", False
            DlgEnable "idx_start_text", False
            DlgEnable "idx_zeros_text", False
            DlgEnable "idx_zeros", False
            Dlgenable "use_regex", False
            ' DlgEnable "case_option", False
        Case 2  ' value changing or button pressed
            Select Case DlgItem
                Case "rename_option"
                    ' Grey out DlgItems that aren't available
                    Select Case SuppValue
                        Case 0  ' overwrite
                            DlgText "title_text", "New name:"
                            DlgEnable "replace_text", False
                            DlgEnable "use_regex", False
                            DlgEnable "append_option", False
                            ' DlgEnable "indexed", True
                            If DlgValue("indexed") Then
                                DlgEnable "idx_start", True
                                DlgEnable "idx_step", True
                                DlgEnable "idx_start_text", True
                                DlgEnable "idx_step_text", True
                                DlgEnable "idx_zeros_text", True
                                DlgEnable "idx_zeros", True
                            End If
                            DialogFunc = True  'do not exit dialog
                        Case 1  ' append
                            DlgText "title_text", "Text to append:"
                            DlgEnable "replace_text", False
                            DlgEnable "use_regex", False
                            DlgEnable "append_option", True
                            ' DlgEnable "indexed", True
                            If DlgValue("indexed") Then
                                DlgEnable "idx_start", True
                                DlgEnable "idx_step", True
                                DlgEnable "idx_start_text", True
                                DlgEnable "idx_step_text", True
                                DlgEnable "idx_zeros_text", True
                                DlgEnable "idx_zeros", True
                            End If
                            DialogFunc = True  'do not exit dialog
                        Case 2  'find and replace
                            DlgText "title_text", "New text:"
                            DlgEnable "replace_text", True
                            DlgEnable "use_regex", True
                            DlgEnable "append_option", False
                            ' DlgEnable "indexed", False
                            DlgEnable "idx_start", False
                            DlgEnable "idx_step", False
                            DlgEnable "idx_start_text", False
                            DlgEnable "idx_step_text", False
                            DlgEnable "idx_zeros_text", False
                            DlgEnable "idx_zeros", False
                            DialogFunc = True  'do not exit dialog
                    End Select
                Case "indexed"
                    Select Case SuppValue
                        Case 0  ' no index
                            DlgEnable "idx_start", False
                            DlgEnable "idx_step", False
                            DlgEnable "idx_start_text", False
                            DlgEnable "idx_step_text", False
                            DlgEnable "idx_zeros_text", False
                            DlgEnable "idx_zeros", False
                            DialogFunc = True  'do not exit dialog
                        Case 1  ' with index
                            DlgEnable "idx_start", True
                            DlgEnable "idx_step", True
                            DlgEnable "idx_start_text", True
                            DlgEnable "idx_step_text", True
                            DlgEnable "idx_zeros_text", True
                            DlgEnable "idx_zeros", True
                            DialogFunc = True  'do not exit dialog
                    End Select
            End Select
        Case 4 ' Focus changed
            Select Case SuppValue
                Case 18, 19 ' ensure indexing values are integers
                    On Error Resume Next
                        new_val = Clng(Eval(DlgText(SuppValue)))

                    If Err <> 0 Then
                        Err.Clear()
                        App.feAppMessageBox(0, "Value must be numeric")
                        DlgText SuppValue, "1"
                        DlgFocus(SuppValue)
                    ElseIf DlgText(SuppValue) = "" Then
                        DlgText SuppValue, "1"
                    Else
                        DlgText SuppValue, CStr(new_val)
                    End If
                    DialogFunc=True
            End Select
    End Select
End Function


Private Sub SetEntityType(type_option As Long)
    'set object type based on selection
    Select Case type_option
        Case 0
            Set obj = App.feCsys()
            select_prompt = "Coordinate Systems..."
            entity_type = FT_CSYS
            select_multi_id = True
        Case 1
            Set obj = App.feSolid()
            select_prompt = "Geometry..."
            entity_type = FT_SOLID
            select_multi_id = False
        Case 2
            Set obj = App.feConnectionProp()
            select_prompt = "Connection Properties..."
            entity_type = FT_CONNECTION_PROP
            select_multi_id = True
        Case 3
            Set obj = App.feConnectionRegion()
            select_prompt = "Connection Regions..."
            entity_type = FT_CONTACT
            select_multi_id = True
        Case 4
            Set obj = App.feConnection()
            select_prompt = "Connectors..."
            entity_type = FT_CONNECTION
            select_multi_id = True
        Case 5
            Set obj = App.feMatl()
            select_prompt = "Materials..."
            entity_type = FT_MATL
            select_multi_id = False
        Case 6
            Set obj = App.feProp()
            select_prompt = "Properties..."
            entity_type = FT_PROP
            select_multi_id = False
        Case 7
            Set obj = App.feLayup()
            select_prompt = "Layups..."
            entity_type = FT_LAYUP
            select_multi_id = False
        Case 8
            Set obj = App.feLoadSet()
            select_prompt = "Load Sets..."
            entity_type = FT_LOAD_DIR
            select_multi_id = True
        Case 9
            Set obj = App.feBCSet()
            select_prompt = "Constraint Sets..."
            entity_type = FT_BC_DIR
            select_multi_id = True
        Case 10
            Set obj = App.feFunction()
            select_prompt = "Functions..."
            entity_type = FT_FUNCTION_DIR
            select_multi_id = True
        Case 11
            Set obj = App.feTableData()
            select_prompt = "Tables..."
            entity_type = FT_TABLE_DATA
            select_multi_id = True
        Case 12
            Set obj = App.feDataSurf()
            select_prompt = "Data Surfaces..."
            entity_type = FT_DATA_SURF
            select_multi_id = True
        Case 13
            Set obj = App.feAnalysisMgr()
            select_prompt = "Analysis Sets..."
            entity_type = FT_AMGR_DIR
            select_multi_id = True
        Case 14
            Set obj = App.feAnalysisStudy()
            select_prompt = "Analysis Studies..."
            entity_type = FT_ANALYSIS_STUDY
            select_multi_id = True
        Case 15
            Set obj = App.feOutputSet()
            select_prompt = "output Sets..."
            entity_type = FT_OUT_CASE
            select_multi_id = True
        Case 16
            Set obj = App.feView()
            select_prompt = "Views..."
            entity_type = FT_VIEW
            select_multi_id = True
        Case 17
            Set obj = App.feGroup()
            select_prompt = "Groups..."
            entity_type = FT_GROUP
            select_multi_id = True
        Case 18
            Set obj = App.feLayer()
            select_prompt = "Layers..."
            entity_type = FT_LAYER
            select_multi_id = True
        Case 19
            Set obj = App.feFreebody()
            select_prompt = "Freebodies..."
            entity_type = FT_FREEBODY
            select_multi_id = True
        Case 20
            Set obj = App.feChart()
            select_prompt = "Charts..."
            entity_type = FT_CHART
            select_multi_id = True
        Case 21
            Set obj = App.feChartSeries()
            select_prompt = "Chart Series..."
            entity_type = FT_CHART_SERIES
            select_multi_id = True
        Case 22
            Set obj = App.fePlane()
            select_prompt = "Planes..."
            entity_type = FT_PLANE
            select_multi_id = True
        Case 23
            Set obj = App.feVector()
            select_prompt = "Vectors..."
            entity_type = FT_VECTOR
            select_multi_id = True
    End Select
End Sub


Function ProperCase(str As String)
    Dim parsed As Variant
    str = LCase(str)
    parsed = Split(str)

    For i = 0 to Ubound(parsed)
        If Len(parsed(i)) > 1 Then
            parsed(i) = UCase(Left(parsed(i), 1)) & LCase(Right(parsed(i), Len(parsed(i)) - 1))
        End If
    Next i

    ProperCase = Join(parsed, " ")
End Function
