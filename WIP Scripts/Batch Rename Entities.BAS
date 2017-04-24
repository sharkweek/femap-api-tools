Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim rc As Long

    Dim objSet As femap.Set     'Object set
    Set objSet = App.feSet
    Dim obj As Object           'Individual object
    Dim objID As Long           'Individual object ID
    Dim newName As String       'Individual object name
    Dim prevName As String      'Existing indivdual solid name
    Dim Types(20) As String     'Array of entity types for selection box

    Types(0) = "Coordinate Systems"
    Types(1) = "Solids"
    Types(2) = "Connection Properties"
    Types(3) = "Connection Regions"
    Types(4) = "Connectors"
    Types(5) = "Aero Panels/Bodies"
    Types(6) = "Aero Properties"
    Types(7) = "Aero Splines"
    Types(8) = "Aero Control Surfaces"
    Types(9) = "Materials"
    Types(10) = "Properties"
    Types(11) = "Layups"
    Types(12) = "Load Sets"
    Types(13) = "Constraint Sets"
    Types(14) = "Functions"
    Types(15) = "Analysis Studies"
    Types(16) = "Layers"
    Types(17) = "Groups"
    Types(18) = "Charts"
    Types(19) = "Data Series"
    Types(20) = "Output Sets"

    'Request info from user
    'Request info from user
	Begin Dialog UserDialog 310,322,"Title for Selected Solid(s)" ' %GRID:10,7,1,1
		text 20,14,270,14,"Enter new title:",.objName_Text,2
		TextBox 20,35,270,21,.objTitle_TextBox
		GroupBox 20,105,120,77,"Append"
		GroupBox 170,105,120,77,"Replace"
		GroupBox 20,189,270,77,"Find and Replace"
		OptionGroup .RenameType
			OptionButton 40,126,70,14,"Prefix"
			OptionButton 40,154,70,14,"Suffix"
			OptionButton 190,126,90,14,"Indexed"
			OptionButton 190,154,90,14,"No Index"
			OptionButton 40,210,130,14,"Replace with..."
		CheckBox 190,210,80,14,"Indexed",.ReplaceIndex
		DropListBox 20,70,270,21,Types(),.ObjectType
		TextBox 40,238,230,14,.ReplaceWith
		OKButton 40,280,100,28
		CancelButton 170,280,100,28
	End Dialog

    Dim dlg As UserDialog
    Dialog dlg

    Select Case dlg.ObjectType
        Case 0 'Coordinate Systems
            Set obj = App.feCSys

            'Prompt user to select objects
            rc = objSet.Select ( FT_CSYS, True, "Select Coodrinate Systems for Renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No coordinate systems selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 1 'Solids
            Set obj = App.feSolid

            'Prompt user to select objects
            rc = objSet.Select ( FT_SOLID, True, "Select Solids for Renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No solids selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 2 'Connection Properties
            Set obj = App.feConnectionProp

            'Prompt user to select objects
            rc = objSet.Select ( FT_CONNECTION_PROP, True, "Select Connection Properties for Renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No coordinate systems selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 3 'Connection Regions
            Set obj = App.feConnectionRegion

            'Prompt user to select objects
            rc = objSet.Select ( FT_CONTACT, True, "Select connection region for Renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No regions selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 4 'Connectors
            Set obj = App.feConnection

            'Prompt user to select objects
            rc = objSet.Select ( FT_CONNECTION, True, "Select connector for Renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No connectors selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 5 'Aero Panels/Bodies
            Set obj = App.feAeroPanel

            'Prompt user to select objects
            rc = objSet.Select ( FT_AERO_PANEL, True, "Select aero panels/bodies for renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No aero panels/bodies selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 6 'Aero Properties
            Set obj = App.feAeroProp

            'Prompt user to select objects
            rc = objSet.Select ( FT_AERO_PROP, True, "Select aero properties for renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No aero properties selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 7 'Aero Splines
            Set obj = App.feAeroSpline

            'Prompt user to select objects
            rc = objSet.Select ( FT_AERO_SPLINE, True, "Select aero splines for renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No aero splines selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 8 'Aero Control Surfaces
            Set obj = App.feAeroSurf

            'Prompt user to select objects
            rc = objSet.Select ( FT_AERO_SURF, True, "Select aero control surfaces for renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No aero control surfaces selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 9 'Materials
            Set obj = App.feMatl

            'Prompt user to select objects
            rc = objSet.Select ( FT_MATL, True, "Select materials for renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No materials selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 10 'Properties
            Set obj = App.feProp

            'Prompt user to select objects
            rc = objSet.Select ( FT_PROP, True, "Select properties for renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No properties selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 11 'Layups
            Set obj = App.feLayup

            'Prompt user to select objects
            rc = objSet.Select ( FT_LAYUP, True, "Select layups for renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No layups selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 12 'Load Sets
            Set obj = App.feLoadSet

            'Prompt user to select objects
            rc = objSet.Select ( FT_LOAD_DIR, True, "Select load sets for renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No load sets selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 13 'Constraint Sets
            Set obj = App.feBCDSet

            'Prompt user to select objects
            rc = objSet.Select ( FT_BC_DIR, True, "Select constraint sets for renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No constraint sets selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 14 'Functions
            Set obj = App.feFunction

            'Prompt user to select objects
            rc = objSet.Select ( FT_FUNCTION_DIR, True, "Select functions for renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No functions selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 15 'Analysis Studies
            Set obj = App.feAnalysisStudy

            'Prompt user to select objects
            rc = objSet.Select ( FT_ANALYSIS_STUDY, True, "Select analysis studies for renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No analysis studies selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 16 'Layers
            Set obj = App.feLayer

            'Prompt user to select objects
            rc = objSet.Select ( FT_LAYER, True, "Select Layers for renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No layers selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 17 'Groups
            Set obj = App.feGroup

            'Prompt user to select objects
            rc = objSet.Select ( FT_GROUP, True, "Select groups for renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No groups selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 18 'Charts
            Set obj = App.feChart

            'Prompt user to select objects
            rc = objSet.Select ( FT_CHART, True, "Select charts for renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No charts selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 19 'Data Series
            Set obj = App.feChartSeries

            'Prompt user to select objects
            rc = objSet.Select ( FT_CHART_SERIES, True, "Select chart series for renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No chart series selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

        Case 20 'Output Sets
            Set obj = App.feChartSeries

            'Prompt user to select objects
            rc = objSet.Select ( FT_OUT_CASE, True, "Select output sets for renaming" )
            If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
                App.feAppMessage ( FCM_HIGHLIGHT, "No output sets selected")
                GoTo EndMacro
            End If

            BatchRename(obj, objSet, dlg.objTitle_TextBox, dlg.RenameType, dlg.ReplaceWith, dlg.ReplaceIndex)

    End Select

    EndMacro:
    App.feViewRegenerate ( 0 )

End Sub

Function BatchRename(obj As Object, _
                     objSet As Object, _
                     newName As String, _
                     renameType As Integer, _
                     Optional replaceWith As String, _
                     Optional replaceIndex As Boolean)

    objSet.Reset
    objID = objSet.First

    Select Case renameType
        Case 0 'Prefix
            For i = 0 To objSet.Count-1
                obj.Get ( objID )
                prevName = obj.title
                obj.title = newName + prevName
                obj.Put ( objID )
                objID = objSet.Next
            Next

        Case 1 'Suffix
            For i = 0 To objSet.Count-1
                obj.Get ( objID )
                prevName = obj.title
                obj.title = prevName + newName
                obj.Put ( objID )
                objID = objSet.Next
            Next

        Case 2 'Replace (Indexed)
            For i = 0 To objSet.Count-1
                obj.Get ( objID )
                obj.title = newName + Str$(i+1)
                obj.Put ( objID )
                objID = objSet.Next
            Next

        Case 3 'Replace (No Index)
            For i = 0 To objSet.Count-1
                obj.Get ( objID )
                obj.title = newName
                obj.Put ( objID )
                objID = objSet.Next
            Next

        Case 4 'Find and Replace
            If replaceIndex Then
        	    For i = 0 To objSet.Count-1
        	        obj.Get ( objID )
        	        prevName = Replace(obj.title, newName, replaceWith + Str$(i+1))
        	        obj.title = prevName
        	        obj.Put ( objID )
        	        objID = objSet.Next
        	    Next

            Else
                For i = 0 To objSet.Count-1
                    obj.Get ( objID )
                    prevName = Replace(obj.title, newName, replaceWith)
                    obj.title = prevName
                    obj.Put ( objID )
                    objID = objSet.Next
                Next
            End If
    End Select

End Function
