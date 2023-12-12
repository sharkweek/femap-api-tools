' Title: Assign Mesh Points to Geometry.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 12.0

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

Dim App As femap.model
Dim geom as Object
Dim select_prompt As String

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Assign Mesh Points to Geometry")
    App.feAppMessage(FCM_NORMAL, "Assign multiple mesh points to a single geometric entity.")

    Dim rc As Long

    Dim mpSet As femap.Set
    Set mpSet = App.feSet

    Dim mp As femap.MeshHardPoint
    Set mp = App.feMeshHardPoint

    ' arrays for dropdown box
    Dim geom_names(19) As String
    geom_names(0) = "0..None"
    geom_names(1) = "1..Point"
    geom_names(2) = "2..Curve"
    geom_names(3) = "3..Surface"
    geom_names(4) = "4..Solid"

    Dim geom_type(19) As Long
    geom_type(0) = 0
    geom_type(1) = FT_POINT
    geom_type(2) = FT_CURVE
    geom_type(3) = FT_SURFACE
    geom_type(4) = FT_SOLID

    ' prompt user to select mesh points
    rc = mpSet.Select(FT_MESH_POINT, True, "Select mesh points...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_FAIL Then
        rc = App.feAppMessage(FCM_ERROR, "Selected mesh points do not exist.")
    End If

    'User inputs
    Begin Dialog UserDialog 1250,650,320,75,"Assign Meshpoints to Geometry" ' %GRID:10,5,1,1
        text 20,5,90,15,"Entity Type:",.geom_type_text
        DropListBox 20,21,280,20,geom_names(),.geom_type
        OKButton 55,47,100,25
        CancelButton 165,47,100,25
    End Dialog
    Dim dlg As UserDialog

    ' prompt user
    If Dialog(dlg) = 0 Then
        Exit Sub  ' exit if user cancels
    End If

    ' set geometry object type
    Select Case dlg.geom_type
        Case 1
            Set geom = App.fePoint
            select_prompt = "Point..."
        Case 2
            Set geom = App.feCurve
            select_prompt = "Curve..."
        Case 3
            Set geom = App.feSurface
            select_prompt = "Surface..."
        Case 4
            Set geom = App.feSolid
            select_prompt = "Solid..."
    End Select

    ' prompt user to select geometry
    If dlg.geom_type <> 0 Then
        rc = geom.SelectID("Select " & select_prompt)
        If rc = FE_CANCEL Then
            Exit Sub
        End If
    End If

    ' assign mesh points to geometry
    Do While mp.NextInSet(mpSet.ID)
        mp.Get(mp.ID)
        mp.OnGeometryType = geom_type(dlg.geom_type)
        If mp.OnGeometryType <> 0 Then
            mp.OnGeometryID = geom.ID
        End If
        mp.Put(mp.ID)
    Loop

    App.feViewRegenerate(0)

End Sub
