' module scope variables
Dim App As femap.model
Dim obj as Object
Dim select_prompt As String

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Batch Rename Entities")

    Dim rc As Long

    Dim objSet As femap.Set
    Set objSet = App.feSet

    Dim objid As Long
    Dim obj_count As Long
    Dim new_name As String      'new name from user
    Dim prev_name As String     'existing name
    Dim index As Long

    ' arrays for dropdown box
    Dim entity_names(19) As String
    entity_names(0) = "9..Coordinate Systems"
    entity_names(1) = "10..Materials"
    entity_names(2) = "11..Properties"
    entity_names(3) = "12..Load Set"
    entity_names(4) = "22..Views"
    entity_names(5) = "24..Groups"
    entity_names(6) = "28..Output Sets"
    entity_names(7) = "33..Layers"
    entity_names(8) = "39..Solids"
    entity_names(9) = "46..Load Definitions"
    entity_names(10) = "58..Connection Region"
    entity_names(11) = "60..Analysis Set"
    entity_names(12) = "69..Data Surfaces"
    entity_names(13) = "72..Connection Property"
    entity_names(14) = "81..Layups"
    entity_names(15) = "102..Freebodies"
    entity_names(16) = "106..Charts"
    entity_names(17) = "107..Chart Data Series"
    entity_names(18) = "113..Analysis Studies"

    Dim entity_ids(19) As Long
    entity_ids(0) = FT_CSYS
    entity_ids(1) = FT_MATL
    entity_ids(2) = FT_PROP
    entity_ids(3) = FT_LOAD_DIR
    entity_ids(4) = FT_VIEW
    entity_ids(5) = FT_GROUP
    entity_ids(6) = FT_OUT_CASE
    entity_ids(7) = FT_LAYER
    entity_ids(8) = FT_SOLID
    entity_ids(9) = FT_LOAD_DEFINITION
    entity_ids(10) = FT_CONTACT
    entity_ids(11) = FT_AMGR_DIR
    entity_ids(12) = FT_DATA_SURF
    entity_ids(13) = FT_CONNECTION_PROP
    entity_ids(14) = FT_LAYUP
    entity_ids(15) = FT_FREEBODY
    entity_ids(16) = FT_CHART
    entity_ids(17) = FT_CHART_SERIES
    entity_ids(18) = FT_ANALYSIS_STUDY

    'User inputs
    Begin Dialog UserDialog 1250,650,320,270,"Rename Entities...",.DialogFunc ' %GRID:10,5,1,1
        text 20,50,260,14,"New name:",.title_text
        GroupBox 20,91,140,86,"Rename Type",.rename_groupbox
        GroupBox 180,91,120,64,"Append Options",.append_groupbox
        GroupBox 20,182,280,47,"Text to Be Replaced:",.replace_groupbox
        OptionGroup .rename_option
            OptionButton 40,110,90,14,"Overwrite",.overwrite
            OptionButton 40,131,90,14,"Append",.append
            OptionButton 40,152,90,14,"Replace",.find_replace
        OptionGroup .append_option
            OptionButton 200,110,70,14,"Prefix",.prefix
            OptionButton 200,131,70,14,"Suffix",.suffix
        TextBox 40,203,240,14,.replace_text
        TextBox 20,65,280,15,.new_name
        CheckBox 190,160,90,15,"Indexed",.indexed
        OKButton 50,235,100,30
        CancelButton 170,235,100,30
        DropListBox 20,21,280,20,entity_names(),.entity_type
        text 20,5,90,15,"Entity Type:",.entity_type_text
    End Dialog
    Dim dlg As UserDialog

    ' prompt user
    If Dialog(dlg) = 0 Then
        Exit Sub  ' exit if user cancels
    End If

    ' set the entity type using the defined function below
    SetEntityType(obj, dlg.entity_type, select_prompt)

    'prompt user to select entities for renaming
    rc = objSet.Select (entity_ids(dlg.entity_type), True, select_prompt )
    If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
        App.feAppMessage (FCM_ERROR, "Nothing selected...")
        Exit Sub
    End If

    new_name = dlg.new_name
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
                    obj.title = new_name & CStr(index)
                    obj.Put(obj.ID)
                    index = index + 1
                Loop
            ' no indices
            Else
                Do While obj.NextInSet(objSet.ID)
                    obj.Get(obj.ID)
                    obj.title = new_name
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
                        obj.title = new_name & prev_name & CStr(index)
                        obj.Put(obj.ID)
                        index = index + 1
                    Loop
                ' no indices
                Else
                    Do While obj.NextInSet(objSet.ID)
                        obj.Get(obj.ID)
                        prev_name = obj.title
                        obj.title = new_name & prev_name
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
                        obj.title = prev_name & new_name + CStr(index)
                        obj.Put(obj.ID)
                        index = index + 1
                    Loop
                ' no indices
                Else
                    Do While obj.NextInSet(objSet.ID)
                        obj.Get(obj.ID)
                        prev_name = obj.title
                        obj.title = prev_name & new_name
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
                    obj.title = Replace(prev_name, dlg.replace_text, new_name) & CStr(index)
                    obj.Put(obj.ID)
                    index = index + 1
                Loop
            ' no indices
            Else
                Do While obj.NextInSet(objSet.ID)
                    obj.Get(obj.ID)
                    prev_name = obj.title
                    obj.title = Replace(prev_name, dlg.replace_text, new_name)
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
            Beep
            DlgEnable "prefix", False
            DlgEnable "suffix", False
            DlgEnable "replace_text", False
        Case 2  ' value changing or button pressed
            If DlgItem = "rename_option" Then
                ' Grey out DlgItems that aren't available
                Select Case SuppValue
                    Case 0  ' overwrite
                        DlgText "title_text", "New name:"
                        DlgEnable "prefix", False
                        DlgEnable "suffix", False
                        DlgEnable "replace_text", False
                        DialogFunc = True  'do not exit dialog
                    Case 1  ' append
                        DlgText "title_text", "Text to append:"
                        DlgEnable "prefix", True
                        DlgEnable "suffix", True
                        DlgEnable "replace_text", False
                        DialogFunc = True  'do not exit dialog
                    Case 2  'find and replace
                        DlgText "title_text", "Replacement text:"
                        DlgEnable "prefix", False
                        DlgEnable "suffix", False
                        DlgEnable "replace_text", True
                        DialogFunc = True  'do not exit dialog
                End Select
            End If
    End Select
End Function


Private Function SetEntityType(obj As Object, type_option As Long, prompt As String)
        Select Case type_option
        Case 0
            Set obj = App.feCsys
            select_prompt = "Coordinate Systems..."
        Case 1
            Set obj = App.feMatl
            select_prompt = "Materials..."
        Case 2
            Set obj = App.feProp
            select_prompt = "Properties..."
        Case 3
            Set obj = App.feLoadSet
            select_prompt = "Load Sets..."
        Case 4
            Set obj = App.feView
            select_prompt = "Views..."
        Case 5
            Set obj = App.feGroup
            select_prompt = "Groups..."
        Case 6
            Set obj = App.feOutputSet
            select_prompt = "output Sets..."
        Case 7
            Set obj = App.feLayer
            select_prompt = "Layers..."
        Case 8
            Set obj = App.feSolid
            select_prompt = "Solids..."
        Case 9
            Set obj = App.feLoadDefinition
            select_prompt = "Load Definitions..."
        Case 10
            Set obj = App.feConnectionRegion
            select_prompt = "Connection Regions..."
        Case 11
            Set obj = App.feAnalysisMgr
            select_prompt = "Analysis Sets..."
        Case 12
            Set obj = App.feDataSurf
            select_prompt = "Data Surfaces..."
        Case 13
            Set obj = App.feConnectionProp
            select_prompt = "Connection Properties..."
        Case 14
            Set obj = App.feLayup
            select_prompt = "Layups..."
        Case 15
            Set obj = App.feFreebody
            select_prompt = "Freebodies..."
        Case 16
            Set obj = App.feChart
            select_prompt = "Charts..."
        Case 17
            Set obj = App.feChartSeries
            select_prompt = "Chart Series..."
        Case 18
            Set obj = App.feAnalysisStudy
            select_prompt = "Analysis Studies..."
    End Select
End Function