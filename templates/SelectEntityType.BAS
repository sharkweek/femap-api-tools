' Title: SelectEntityType.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2306

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' Usage Notes
' -----------
' This is a template for an entity type selection drop down for in UserDialog
' objects.

' module scope variables - these variables have to be scoped outside of `Main`
' for the drop down to work.
Dim App As femap.model
Dim obj as Object
Dim entity_type As Long
Dim select_prompt As String

Sub Main

    ' arrays for dropdown box
    Dim entity_names(20) As String
    entity_names(0) = "1..Coordinate Systems"
    entity_names(1) = "2..Materials"
    entity_names(2) = "3..Properties"
    entity_names(3) = "4..Load Set"
    entity_names(4) = "5..Views"
    entity_names(5) = "6..Groups"
    entity_names(6) = "7..Output Sets"
    entity_names(7) = "8..Layers"
    entity_names(8) = "9..Solids"
    entity_names(9) = "10..Load Definitions"
    entity_names(10) = "11..Connection Region"
    entity_names(11) = "12..Analysis Set"
    entity_names(12) = "13..Data Surfaces"
    entity_names(13) = "14..Connection Property"
    entity_names(14) = "15..Layups"
    entity_names(15) = "16..Freebodies"
    entity_names(16) = "17..Charts"
    entity_names(17) = "18..Chart Data Series"
    entity_names(18) = "19..Analysis Studies"
    entity_names(19) = "20..Functions"

    ' prompt user to select entity type
    Begin Dialog UserDialog 360,77,"Entity Type" ' %GRID:10,7,1,1
        DropListBox 20,14,320,21,entity_names(),.entity_name
        OKButton 60,49,100,21
        CancelButton 190,49,90,21
    End Dialog
    Dim dlg As UserDialog

    ' prompt user
    If Dialog(dlg) = 0 Then
        Exit Sub  ' exit if user cancels
    End If

    SetEntityType(obj, dlg.entity_name, select_prompt)

End Sub

Private Function SetEntityType(obj As Object, type_option As Long, prompt As String)
' sets the provided object to the type selected by the dialog box
    Select Case type_option
        Case 0
            Set obj = App.feCsys
            prompt = "Coordinate Systems..."
            entity_type = FT_CSYS
        Case 1
            Set obj = App.feMatl
            prompt = "Materials..."
            entity_type = FT_MATL
        Case 2
            Set obj = App.feProp
            prompt = "Properties..."
            entity_type = FT_PROP
        Case 3
            Set obj = App.feLoadSet
            prompt = "Load Sets..."
            entity_type = FT_LOAD_DIR
        Case 4
            Set obj = App.feView
            prompt = "Views..."
            entity_type = FT_VIEW
        Case 5
            Set obj = App.feGroup
            prompt = "Groups..."
            entity_type = FT_GROUP
        Case 6
            Set obj = App.feOutputSet
            prompt = "output Sets..."
            entity_type = FT_OUT_CASE
        Case 7
            Set obj = App.feLayer
            prompt = "Layers..."
            entity_type = FT_LAYER
        Case 8
            Set obj = App.feSolid
            prompt = "Solids..."
            entity_type = FT_SOLID
        Case 9
            Set obj = App.feLoadDefinition
            prompt = "Load Definitions..."
            entity_type = FT_LOAD_DEFINITION
        Case 10
            Set obj = App.feConnectionRegion
            prompt = "Connection Regions..."
            entity_type = FT_CONTACT
        Case 11
            Set obj = App.feAnalysisMgr
            prompt = "Analysis Sets..."
            entity_type = FT_AMGR_DIR
        Case 12
            Set obj = App.feDataSurf
            prompt = "Data Surfaces..."
            entity_type = FT_DATA_SURF
        Case 13
            Set obj = App.feConnectionProp
            prompt = "Connection Properties..."
            entity_type = FT_CONNECTION_PROP
        Case 14
            Set obj = App.feLayup
            prompt = "Layups..."
            entity_type = FT_LAYUP
        Case 15
            Set obj = App.feFreebody
            prompt = "Freebodies..."
            entity_type = FT_FREEBODY
        Case 16
            Set obj = App.feChart
            prompt = "Charts..."
            entity_type = FT_CHART
        Case 17
            Set obj = App.feChartSeries
            prompt = "Chart Series..."
            entity_type = FT_CHART_SERIES
        Case 18
            Set obj = App.feAnalysisStudy
            prompt = "Analysis Studies..."
            entity_type = FT_ANALYSIS_STUDY
        Case 19
            Set obj = App.feFunction
            prompt = "Functions..."
            entity_type = FT_FUNCTION_DIR
    End Select
End Function
