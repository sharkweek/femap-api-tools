' Title: Show Where Used.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 12.0

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

Private App As femap.model

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Show Where Used")
    App.feAppMessage(FCM_NORMAL, "Reports which entities use a user-selected entity" & vbFormFeed)
    App.feAppMessage(FCM_WARNING, "WARNING: Take care when selecting objects referenced by large numbers" _
                     & vbFormFeed & "         of other entities such as coordinate systems." & vbFormFeed)

    Dim rc As Long
    Dim query_obj As Object
    Dim fe_type As Long

    Dim sel As femap.Selector
    Set sel = App.feSelector

    Dim query_set As femap.Set
    Set query_set = App.feSet

    Dim gr As femap.Group
    Set gr = App.feGroup

    Dim gr_set As femap.Set
    Set gr_set = App.feSet

    Dim i As Long
    Dim rules() As Long
    Dim types() As String
    Dim select_prompt As String
    Dim sel_gid As Long
    Dim sel_clear As Boolean
    Dim report As String
    Dim type_name As String
    Dim c_width As Long
    Dim lines As Long
    Dim skip_list As Boolean

    ' arrays for dropdown box
    Dim entity_names(10) As String
    entity_names(0) = "Coordinate System"
    entity_names(1) = "Material"
    entity_names(2) = "Property"
    entity_names(3) = "Layup"
    entity_names(4) = "Point"
    entity_names(5) = "Curve"
    entity_names(6) = "Surface"
    entity_names(7) = "Solid"
    entity_names(8) = "Node"
    entity_names(9) = "Element"

    ' dialog
    Begin Dialog UserDialog 330,105,"Show Where Used...",.DialogFunc ' %GRID:10,7,1,1
        DropListBox 10,25,310,21,entity_names(),.entity_type
        text 10,7,90,14,"Entity type:",.entity_type_text
        OKButton 50,77,110,21
        CancelButton 170,77,110,21
        CheckBox 20,53,140,14,"Populate Selector",.pop_selector
        CheckBox 170,53,140,14,"Highlight entities",.highlight
    End Dialog
    Dim dlg As UserDialog

    ' prompt user
    If Dialog(dlg) = 0 Then
        Exit Sub  ' exit if user cancels
    End If

    ' get associated entity types for selected entity
    SetEntityType(query_obj, dlg.entity_type, fe_type, type_name, rules, types)

    ' prompt user to select query_obj
    rc = query_obj.SelectID("Select " & type_name & "...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf query_obj.CountSet() = 0 Then
        App.feAppMessage(FCM_ERROR, "No " & type_name & " entities exist in model. Script aborted...")
        Exit Sub
    Else
        query_obj.Get(query_obj.ID)
    End If

    ' prep selector
    sel_gid = gr.NextEmptyID
    sel_clear = False  ' do not clear the selector by default
    If dlg.pop_selector Then
        ' prompt to clear
        rc = App.feAppMessageBox(3, "Do you wish to clear the selector? (No=Append)")
        If rc = FE_OK Then
            sel_clear = True
        ElseIf rc = FE_FAIL Then  ' save as group to repopulate at end
            sel.SaveGroup(sel_gid)
        ElseIf rc = FE_CANCEL Then
            Exit Sub
        End If

    Else
        sel.SaveGroup(sel_gid)

    End If

    ' get referenced entities
    sel.ClearAll()
    c_width = 9
    lines = 0
    skip_list = False
    App.feAppMessage(FCM_NORMAL, vbFormFeed & type_name & " " & CStr(query_obj.ID) & " is used in the following entities: " & vbFormFeed)
    App.feAppMessage(FCM_COMMAND, Right(Space(c_width + 4) & "Entity ID", c_width) & Right(Space(c_width) & "Type", c_width))

    For i = 0 To Ubound(rules, 1)
        ' add each entity type to the selector
        query_set.Clear()
        query_set.AddRule(query_obj.ID, rules(i, 0))
        sel.AddSet(rules(i, 1), query_set.ID, False)

        ' begin printing out values
        query_set.Reset()
        Do While query_set.Next() And skip_list = False  ' stop at first 1000 lines
            App.feAppMessage(FCM_NORMAL, Right(Space(c_width + 4) & CStr(query_set.CurrentID), c_width) & " " & _
                                         Right(Space(c_width) & types(i), c_width))

            ' prompt to populate selector at 1000 lines
            lines += 1
            If lines = 1000 And skip_list = False Then
                rc = App.feAppMessageBox(3, "More than 1,000 referenced entities found. Do you wish to clear and" _
                                            & vbCrLf & "add entities to the selector? (No=Continue)" _
                                            & vbCrLf & vbCrLf & "WARNING: Listing all entities may take a significant amount of time.")
                If rc = FE_OK Then
                    skip_list = True
                    sel_clear = True
                ElseIf rc = FE_FAIL Then
                    skip_list = False
                ElseIf rc = FE_CANCEL Then
                    GoTo CLEANUP
                End If
            End If
        Loop

    Next i

    CLEANUP:
    If gr.Exist(sel_id) = FE_OK And sel_clear = False Then
        sel.LoadGroup(sel_gid)
    End If


    gr.Delete(sel_gid)


    App.feViewRegenerate(0)

    If dlg.pop_selector And dlg.highlight Then
        sel.Show(True)
    End If

End Sub


Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Select Case Action
        Case 1  ' dialog box initialization
            DlgEnable "highlight", False
        Case 2  ' value changing or button pressed
            Select Case DlgItem
                Case "pop_selector"
                    If SuppValue = 1 Then
                        DlgEnable "highlight", True
                        DialogFunc = True
                    Else
                        DlgEnable "highlight", False
                        DialogFunc = True
                    End If
            End Select
    End Select
End Function


Private Sub SetEntityType(obj As Object, type_option As Long, entity_type As Long, type_name As String, rule_array() As Long, type_array() As String)
    Select Case type_option
        Case 0
            Set obj = App.feCsys
            entity_type = FT_CSYS
            type_name = "CSYS"

            ReDim rule_array(5, 3)
            ' rules
            rule_array(0, 0) = FGD_POINT_BYDEFCSYS
            rule_array(1, 0) = FGD_NODE_BYDEFCSYS
            rule_array(2, 0) = FGD_NODE_BYOUTCSYS
            rule_array(3, 0) = FGD_CSYS_ONCSYS
            rule_array(4, 0) = FGD_CSYS_BYDEFCSYS

            ' entity types
            rule_array(0, 1) = FT_POINT
            rule_array(1, 1) = FT_NODE
            rule_array(2, 1) = FT_NODE
            rule_array(3, 1) = FT_CSYS
            rule_array(4, 0) = FT_CSYS

            ' entity type names
            ReDim type_array(5)
            type_array(0) = "POINT"
            type_array(1) = "NODE"
            type_array(2) = "NODE"
            type_array(3) = "CSYS"
            type_array(4) = "CSYS"

        Case 1
            Set obj = App.feMatl
            entity_type = FT_MATL
            type_name = "Material"

            ReDim rule_array(3, 3)
            ' rules
            rule_array(0, 0) = FGD_ELEM_BYMATL
            rule_array(1, 0) = FGD_PROP_BYMATL
            rule_array(2, 0) = FGD_LAYUP_BYMATL

            ' entities
            rule_array(0, 1) = FT_ELEM
            rule_array(1, 1) = FT_PROP
            rule_array(2, 1) = FT_LAYUP

            ' entity type names
            ReDim type_array(3)
            type_array(0) = "ELEM"
            type_array(1) = "PROP"
            type_array(2) = "LAYUP"

        Case 2
            Set obj = App.feProp
            entity_type = FT_PROP
            type_name = "Property"

            ReDim rule_array(9, 3)
            ' rules
            rule_array(0, 0) = FGD_ELEM_BYPROP
            rule_array(1, 0) = FGD_MATL_ONPROP
            rule_array(2, 0) = FGD_POINT_BYPROP
            rule_array(3, 0) = FGD_CURVE_BYPROP
            rule_array(4, 0) = FGD_SURFACE_BYPROP
            rule_array(5, 0) = FGD_VOLUME_BYPROP
            rule_array(6, 0) = FGD_SOLID_BYPROP
            rule_array(7, 0) = FGD_LAYUP_ONPROP
            rule_array(8, 0) = FGD_CSYS_ONPROP

            ' entities
            rule_array(0, 1) = FT_ELEM
            rule_array(1, 1) = FT_MATL
            rule_array(2, 1) = FT_POINT
            rule_array(3, 1) = FT_CURVE
            rule_array(4, 1) = FT_SURFACE
            rule_array(5, 1) = FT_VOLUME
            rule_array(6, 1) = FT_SOLID
            rule_array(7, 1) = FT_LAYUP
            rule_array(8, 1) = FT_CSYS

            ' entity type names
            ReDim type_array(9)
            type_array(0) = "ELEM"
            type_array(1) = "MATL"
            type_array(2) = "POINT"
            type_array(3) = "CURVE"
            type_array(4) = "SURFACE"
            type_array(5) = "VOLUME"
            type_array(6) = "SOLID"
            type_array(7) = "LAYUP"
            type_array(8) = "CSYS"

        Case 3
            Set obj = App.feLayup
            entity_type = FT_LAYUP
            type_name = "Layup"

            ReDim rule_array(2, 3)
            ' rules
            rule_array(0, 0) = FGD_ELEM_BYLAYUP
            rule_array(1, 0) = FGD_PROP_BYLAYUP

            ' entities
            rule_array(0, 1) = FT_ELEM
            rule_array(1, 1) = FT_PROP

            ' entity type names
            ReDim type_array(2)
            type_array(0) = "ELEM"
            type_array(1) = "PROP"

        Case 4
            Set obj = App.fePoint
            entity_type = FT_POINT
            type_name = "Point"

            ReDim rule_array(7, 3)
            ' rules
            rule_array(0, 0) = FGD_CURVE_BYPOINT
            rule_array(1, 0) = FGD_NODE_ATPOINT
            rule_array(2, 0) = FGD_ELEM_ATPOINT
            rule_array(3, 0) = FGD_LOAD_BYPOINT
            rule_array(4, 0) = FGD_BCO_BYPOINT
            rule_array(5, 0) = FGD_CSYS_ONPOINT
            rule_array(6, 0) = FGD_PROP_ONPOINT

            ' entities
            rule_array(0, 1) = FT_CURVE
            rule_array(1, 1) = FT_NODE
            rule_array(2, 1) = FT_ELEM
            rule_array(3, 1) = FT_LOAD
            rule_array(4, 1) = FT_BCO
            rule_array(5, 1) = FT_CSYS
            rule_array(6, 1) = FT_PROP

            ' entity type names
            ReDim type_array(7)
            type_array(0) = "CURVE"
            type_array(1) = "NODE"
            type_array(2) = "ELEM"
            type_array(3) = "LOAD"
            type_array(4) = "BC"
            type_array(5) = "CSYS"
            type_array(6) = "PROP"

        Case 5
            Set obj = App.feCurve
            entity_type = FT_CURVE
            type_name = "Curve"

            ReDim rule_array(10, 3)
            ' rules
            rule_array(0, 0) = FGD_POINT_ONCURVE
            rule_array(1, 0) = FGD_SURFACE_BYCURVE
            rule_array(2, 0) = FGD_BOUNDARY_BYCURVE
            rule_array(3, 0) = FGD_NODE_ATCURVE
            rule_array(4, 0) = FGD_ELEM_ATCURVE
            rule_array(5, 0) = FGD_LOAD_BYCURVE
            rule_array(6, 0) = FGD_BCO_BYCURVE
            rule_array(7, 0) = FGD_SOLID_BYCURVE
            rule_array(8, 0) = FGD_CONTACT_BYCURVE
            rule_array(9, 0) = FGD_PROP_ONCURVE

            ' entities
            rule_array(0, 1) = FT_POINT
            rule_array(1, 1) = FT_SURFACE
            rule_array(2, 1) = FT_BOUNDARY
            rule_array(3, 1) = FT_NODE
            rule_array(4, 1) = FT_ELEM
            rule_array(5, 1) = FT_LOAD
            rule_array(6, 1) = FT_BCO
            rule_array(7, 1) = FT_SOLID
            rule_array(8, 1) = FT_CONTACT
            rule_array(9, 1) = FT_PROP

            ' entity type names
            ReDim type_array(10)
            type_array(0) = "POINT"
            type_array(1) = "SURFACE"
            type_array(2) = "BOUNDARY"
            type_array(3) = "NODE"
            type_array(4) = "ELEM"
            type_array(5) = "LOAD"
            type_array(6) = "BCO"
            type_array(7) = "SOLID"
            type_array(8) = "CONTACT"
            type_array(9) = "PROP"

        Case 6
            Set obj = App.feSurface
            entity_type = FT_SURFACE
            type_name = "Surface"

            ReDim rule_array(10, 3)
            ' rules
            rule_array(0, 0) = FGD_CURVE_ONSURFACE
            rule_array(1, 0) = FGD_VOLUME_BYSURFACE
            rule_array(2, 0) = FGD_NODE_ATSURFACE
            rule_array(3, 0) = FGD_ELEM_ATSURFACE
            rule_array(4, 0) = FGD_LOAD_BYSURFACE
            rule_array(5, 0) = FGD_BCO_BYSURFACE
            rule_array(6, 0) = FGD_SOLID_BYSURFACE
            rule_array(7, 0) = FGD_CONTACT_BYSURFACE
            rule_array(8, 0) = FGD_PROP_ONSURFACE
            rule_array(9, 0) = FGD_POINT_ONSURFACE

            ' entities
            rule_array(0, 1) = FT_CURVE
            rule_array(1, 1) = FT_VOLUME
            rule_array(2, 1) = FT_NODE
            rule_array(3, 1) = FT_ELEM
            rule_array(4, 1) = FT_LOAD
            rule_array(5, 1) = FT_BCO
            rule_array(6, 1) = FT_SOLID
            rule_array(7, 1) = FT_CONTACT
            rule_array(8, 1) = FT_PROP
            rule_array(9, 1) = FT_POINT

            ' entity type names
            ReDim type_array(10)
            type_array(0) = "CURVE"
            type_array(1) = "VOLUME"
            type_array(2) = "NODE"
            type_array(3) = "ELEM"
            type_array(4) = "LOAD"
            type_array(5) = "BCO"
            type_array(6) = "SOLID"
            type_array(7) = "CONTACT"
            type_array(8) = "PROP"
            type_array(9) = "POINT"

        Case 7
            Set obj = App.feSolid
            entity_type = FT_SOLID
            type_name = "Solid"

            ReDim rule_array(7, 3)
            ' rules
            rule_array(0, 0) = FGD_NODE_ATSOLID
            rule_array(1, 0) = FGD_ELEM_ATSOLID
            rule_array(2, 0) = FGD_CURVE_ONSOLID
            rule_array(3, 0) = FGD_SURFACE_ONSOLID
            rule_array(4, 0) = FGD_PROP_ONSOLID
            rule_array(5, 0) = FGD_POINT_ONSOLID
            rule_array(6, 0) = FGD_CONTACT_ONSOLID

            ' entities
            rule_array(0, 1) = FT_NODE
            rule_array(1, 1) = FT_ELEM
            rule_array(2, 1) = FT_CURVE
            rule_array(3, 1) = FT_SURFACE
            rule_array(4, 1) = FT_PROP
            rule_array(5, 1) = FT_POINT
            rule_array(6, 1) = FT_CONTACT

            ' entity type names
            ReDim type_array(7)
            type_array(0) = "NODE"
            type_array(1) = "ELEM"
            type_array(2) = "CURVE"
            type_array(3) = "SURFACE"
            type_array(4) = "PROP"
            type_array(5) = "POINT"
            type_array(6) = "CONTACT"

        Case 8
            Set obj = App.feNode
            entity_type = FT_NODE
            type_name = "Node"

            ReDim rule_array(5, 3)
            ' rules
            rule_array(0, 0) = FGD_ELEM_BYNODE
            rule_array(1, 0) = FGD_LOAD_BYNODE
            rule_array(2, 0) = FGD_BEQ_BYNODE
            rule_array(3, 0) = FGD_CSYS_ONNODE
            rule_array(4, 0) = FGD_CONTACT_BYNODE

            ' entities
            rule_array(0, 1) = FT_ELEM
            rule_array(1, 1) = FT_LOAD
            rule_array(2, 1) = FT_BEQ
            rule_array(3, 1) = FT_CSYS
            rule_array(4, 1) = FT_CONTACT

            ' entity type names
            ReDim type_array(5)
            type_array(0) = "ELEM"
            type_array(1) = "LOAD"
            type_array(2) = "BEQ"
            type_array(3) = "CSYS"
            type_array(4) = "CONTACT"

        Case 9
            Set obj = App.feElem
            entity_type = FT_ELEM
            type_name = "Element"

            ReDim rule_array(6, 3)
            ' rules
            rule_array(0, 0) = FGD_NODE_ONELEM
            rule_array(1, 0) = FGD_MATL_ONELEM
            rule_array(2, 0) = FGD_PROP_ONELEM
            rule_array(3, 0) = FGD_LOAD_BYELEM
            rule_array(4, 0) = FGD_CONTACT_BYELEM
            rule_array(5, 0) = FGD_CSYS_ONELEM

            ' entities
            rule_array(0, 1) = FT_NODE
            rule_array(1, 1) = FT_MATL
            rule_array(2, 1) = FT_PROP
            rule_array(3, 1) = FT_LOAD
            rule_array(4, 1) = FT_CONTACT
            rule_array(5, 1) = FT_CSYS

            ' entity type names
            ReDim type_array(6)
            type_array(0) = "NODE"
            type_array(1) = "MATL"
            type_array(2) = "PROP"
            type_array(3) = "LOAD"
            type_array(4) = "CONTACT"
            type_array(5) = "CSYS"

    End Select
End Sub