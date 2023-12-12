' Title: Rename Multiple.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2301

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' Usage Notes
' -----------
' "Replace" feature supports regular expression pattern matching

' module scope variables
Dim App As femap.model
Dim obj as Object
Dim select_prompt As String
Dim entity_type As Long

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

    ' arrays for dropdown box
    Dim entity_names(22) As String
    entity_names(0) = "1..Coordinate Systems"
    entity_names(1) = "2..Geometry"
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

    'User inputs
    Begin Dialog UserDialog 320,200,"Rename Entities...",.DialogFunc ' %GRID:10,5,1,1
        Text 20,5,90,15,"Entity Type:",.entity_type_text
        DropListBox 20,20,280,20,entity_names(),.entity_type
        Text 20,41,260,13,"New name:",.title_text
        TextBox 20,55,280,15,.new_name
        GroupBox 20,75,140,65,"Rename Type",.rename_groupbox
            OptionGroup .rename_option
                OptionButton 40,90,90,15,"Overwrite",.overwrite
                OptionButton 40,105,90,15,"Append",.append
                OptionButton 40,120,90,15,"Replace",.find_replace
        GroupBox 170,75,130,65,"Append Options",.append_groupbox
            OptionGroup .append_option
                OptionButton 190,90,70,15,"Prefix",.prefix
                OptionButton 190,105,70,15,"Suffix",.suffix
            CheckBox 180,120,90,15,"Indexed",.indexed
        Text 20,140,280,20,"Text to Replace:",.replace_text_text
        TextBox 20,155,280,15,.replace_text
        OKButton 50,176,100,18
        CancelButton 160,176,100,18
    End Dialog
    Dim dlg As UserDialog

    ' prompt user
    If Dialog(dlg) = 0 Then
        Exit Sub  ' exit if user cancels
    End If

    ' set the entity type using the defined function below
    SetEntityType(obj, dlg.entity_type, select_prompt)

    'prompt user to select entities for renaming
    ' rc = objSet.Select (entity_ids(dlg.entity_type), True, select_prompt )
    rc = objSet.Select (entity_type, True, select_prompt )
    If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
        App.feAppMessage (FCM_ERROR, "Nothing selected...")
        Exit Sub
    End If

    objSet.Reset()
    objid = objSet.First()
    index = 1  ' set first index value
    obj_count = objSet.Count()

    'rename propect according to user's selections
    Select Case dlg.rename_option
        Case 0  'overwrite
            ' with indices
            If dlg.indexed Then
                Do While obj.NextInSet(objSet.ID)
                    obj.Get(obj.ID)
                    obj.title = dlg.new_name & CStr(index)
                    obj.Put(obj.ID)
                    index = index + 1
                Loop
            ' no indices
            Else
                Do While obj.NextInSet(objSet.ID)
                    obj.Get(obj.ID)
                    obj.title = dlg.new_name
                    obj.Put(obj.ID)
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
                        obj.title = dlg.new_name & prev_name & CStr(index)
                        obj.Put(obj.ID)
                        index = index + 1
                    Loop
                ' no indices
                Else
                    Do While obj.NextInSet(objSet.ID)
                        obj.Get(obj.ID)
                        prev_name = obj.title
                        obj.title = dlg.new_name & prev_name
                        obj.Put(obj.ID)
                    Loop
                End If
            ' suffix option
            Else
                ' with indices
                If dlg.indexed Then
                    Do While obj.NextInSet(objSet.ID)
                        obj.Get(obj.ID)
                        prev_name = obj.title
                        obj.title = prev_name & dlg.new_name + CStr(index)
                        obj.Put(obj.ID)
                        index = index + 1
                    Loop
                ' no indices
                Else
                    Do While obj.NextInSet(objSet.ID)
                        obj.Get(obj.ID)
                        prev_name = obj.title
                        obj.title = prev_name & dlg.new_name
                        obj.Put(obj.ID)
                    Loop
                End If
            End If

        ' find and replace
        Case 2
            ' with indices
            If dlg.indexed Then
                Do While obj.NextInSet(objSet.ID)
                    obj.Get(obj.ID)
                    prev_name = obj.title
                    regex.Pattern = dlg.replace_text
                    obj.title = regex.Replace(obj.title, dlg.new_name) & CStr(index)
                    obj.Put(obj.ID)
                    index = index + 1
                Loop
            ' no indices
            Else
                Do While obj.NextInSet(objSet.ID)
                    obj.Get(obj.ID)
                    prev_name = obj.title
                    regex.Pattern = dlg.replace_text
                    obj.title = regex.Replace(obj.title, dlg.new_name)
                    obj.Put(obj.ID)
                Loop
            End If
    End Select

    App.feViewRegenerate (0)
    App.feAppMessage(FCM_NORMAL, CStr(obj_count) & " entities renamed...")

End Sub


Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Select Case Action%
        Case 1  ' dialog box initialization
            DlgEnable "append_option", False
            DlgEnable "replace_text", False
        Case 2  ' value changing or button pressed
            If DlgItem = "rename_option" Then
                ' Grey out DlgItems that aren't available
                Select Case SuppValue
                    Case 0  ' overwrite
                        DlgText "title_text", "New name:"
                        DlgEnable "replace_text", False
                        DlgEnable "append_option", False
                        DialogFunc = True  'do not exit dialog
                    Case 1  ' append
                        DlgText "title_text", "Text to append:"
                        DlgEnable "replace_text", False
                        DlgEnable "append_option", True
                        DialogFunc = True  'do not exit dialog
                    Case 2  'find and replace
                        DlgText "title_text", "New text:"
                        DlgEnable "replace_text", True
                        DlgEnable "append_option", False
                        DialogFunc = True  'do not exit dialog
                End Select
            End If
    End Select
End Function


Private Function SetEntityType(obj As Object, type_option As Long, prompt As String)
'set object type based on selection
    Select Case type_option
        Case 0
            Set obj = App.feCsys()
            prompt = "Coordinate Systems..."
            entity_type = FT_CSYS
        Case 1
            Set obj = App.feSolid()
            prompt = "Geometry..."
            entity_type = FT_SOLID
        Case 2
            Set obj = App.feConnectionProp()
            prompt = "Connection Properties..."
            entity_type = FT_CONNECTION_PROP
        Case 3
            Set obj = App.feConnectionRegion()
            prompt = "Connection Regions..."
            entity_type = FT_CONTACT
        Case 4
            Set obj = App.feConnection()
            prompt = "Connectors..."
            entity_type = FT_CONNECTION
        Case 5
            Set obj = App.feMatl()
            prompt = "Materials..."
            entity_type = FT_MATL
        Case 6
            Set obj = App.feProp()
            prompt = "Properties..."
            entity_type = FT_PROP
        Case 7
            Set obj = App.feLayup()
            prompt = "Layups..."
            entity_type = FT_LAYUP
        Case 8
            Set obj = App.feLoadSet()
            prompt = "Load Sets..."
            entity_type = FT_LOAD_DIR
        Case 9
            Set obj = App.feBCSet()
            prompt = "Constraint Sets..."
            entity_type = FT_BC_DIR
        Case 10
            Set obj = App.feFunction()
            prompt = "Functions..."
            entity_type = FT_FUNCTION_DIR
        Case 11
            Set obj = App.feTableData()
            prompt = "Tables..."
            entity_type = FT_TABLE_DATA
        Case 12
            Set obj = App.feDataSurf()
            prompt = "Data Surfaces..."
            entity_type = FT_DATA_SURF
        Case 13
            Set obj = App.feAnalysisMgr()
            prompt = "Analysis Sets..."
            entity_type = FT_AMGR_DIR
        Case 14
            Set obj = App.feAnalysisStudy()
            prompt = "Analysis Studies..."
            entity_type = FT_ANALYSIS_STUDY
        Case 15
            Set obj = App.feOutputSet()
            prompt = "Output Sets..."
            entity_type = FT_OUT_CASE
        Case 16
            Set obj = App.feView()
            prompt = "Views..."
            entity_type = FT_VIEW
        Case 17
            Set obj = App.feGroup()
            prompt = "Groups..."
            entity_type = FT_GROUP
        Case 18
            Set obj = App.feLayer()
            prompt = "Layers..."
            entity_type = FT_LAYER
        Case 19
            Set obj = App.feFreebody()
            prompt = "Freebodies..."
            entity_type = FT_FREEBODY
        Case 20
            Set obj = App.feChart()
            prompt = "Charts..."
            entity_type = FT_CHART
        Case 21
            Set obj = App.feChartSeries()
            prompt = "Chart Series..."
            entity_type = FT_CHART_SERIES
    End Select
End Function
