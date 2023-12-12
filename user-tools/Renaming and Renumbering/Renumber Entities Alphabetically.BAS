' Title: Renumber Entities Alphabetically.BAS
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
' Renumber a selected set of groups alphabetically. This script requires that
' group IDs 99999000 and up be unoccupied depending on the number of groups
' selected

' TODO: add functions and boundary constraints
' TODO: retool to use next available ids instead of a 99999000 block for renumbering

' module scope variables
Dim App As femap.model
Dim obj As Object
Dim select_prompt As String
Dim entity_type As Long

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Renumber Entities Alphabetically")

    Dim objSet As femap.Set
    Set objSet = App.feSet

    Dim newObjSet As femap.Set
    Set newObjSet = App.feSet

    Dim dSet As femap.Set
    Set dSet = App.feSet

    Dim listID As Variant
    Dim startID As Long
    Dim titles() As Variant

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim n As Long
    Dim swapped As Boolean
    Dim tSwap As Variant
    Dim idSwap As Variant
    Dim iMax As Long

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

    Begin Dialog UserDialog 360,77,"Entity Type" ' %GRID:10,7,1,1
        DropListBox 20,14,320,21,entity_names(),.entity_type
        OKButton 60,49,100,21
        CancelButton 190,49,90,21
    End Dialog
    Dim dlg As UserDialog

    ' prompt user
    If Dialog(dlg) = 0 Then
        Exit Sub  ' exit if user cancels
    End If

    ' set the entity type using the defined function below
    SetEntityType(obj, dlg.entity_type, select_prompt)

    'prompt user to select entities for renaming
    rc = objSet.Select(entity_type, True, select_prompt)
    If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
        App.feAppMessage(FCM_ERROR, "Nothing selected...")
        Exit Sub
    End If

    ' Prompt for new starting ID
    startID = objSet.First()
    rc = App.feGetInt("Starting ID", 1, 99998999, startID)

    ' Get a list of titles for specified group set
    objSet.Reset()
    k = 0
    Do While obj.NextInSet(objSet.ID)
        ReDim Preserve titles(k)
        ReDim Preserve listID(k)
        obj.Get(obj.ID)
        titles(k) = obj.title
        listID(k) = obj.ID
        k = k + 1
    Loop

    ' bubble sort
    iMax = objSet.Count() - 2
    Do
        swapped = False
        For i = 0 To iMax
            If titles(i) > titles(i + 1) Then
                ' swap titles
                tSwap = titles(i)
                titles(i) = titles(i + 1)
                titles(i + 1) = tSwap

                ' swap IDs
                idSwap = listID(i)
                listID(i) = listID(i + 1)
                listID(i + 1) = idSwap

                swapped = True
            End If
        Next i
        iMax = iMax - 1
    Loop Until swapped = False

    ' move groups to higher ID block in alphabetic order
    n = 99999000
    newObjSet.Clear()
    For i = 0 To objSet.Count() - 1
        obj.Get(listID(i))
        dSet.Clear()
        dSet.Add(listID(i))

        App.feRenumber(entity_type, dSet.ID, n)
        newObjSet.Add(n)
        n = n + 1
    Next i

    ' renumber IDs to selected block in consecutive order
    App.feRenumber(entity_type, newObjSet.ID, startID)

End Sub


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
            prompt = "output Sets..."
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
