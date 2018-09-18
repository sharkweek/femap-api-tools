' Renumber a selected set of groups alphabetically. This script requires that
' group IDs 99999000 and up be unoccupied depending on the number of groups
' selected

' TODO: add functions and boundary constraints

' module scope variables
Dim App As femap.model
Dim obj As Object
Dim select_prompt As String

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
    Dim entity_names(12) As String
    entity_names(0)  = "1..Coordinate Systems"
    entity_names(1)  = "2..Materials"
    entity_names(2)  = "3..Properties"
    entity_names(3)  = "4..Load Set"
    entity_names(4)  = "5..Groups"
    entity_names(5)  = "6..Output Sets"
    entity_names(6)  = "7..Layers"
    entity_names(7)  = "8..Solids"
    entity_names(8)  = "9..Connection Region"
    entity_names(9)  = "10..Analysis Set"
    entity_names(10) = "11..Connection Property"
    entity_names(11) = "12..Layups"

    Dim entity_ids(12) As Long
    entity_ids(0)  = FT_CSYS             ' 9
    entity_ids(1)  = FT_MATL             ' 10
    entity_ids(2)  = FT_PROP             ' 11
    entity_ids(3)  = FT_LOAD_DIR         ' 12
    entity_ids(4)  = FT_GROUP            ' 24
    entity_ids(5)  = FT_OUT_CASE         ' 28
    entity_ids(6)  = FT_LAYER            ' 33
    entity_ids(7)  = FT_SOLID            ' 39
    entity_ids(8)  = FT_CONTACT          ' 58
    entity_ids(9)  = FT_AMGR_DIR         ' 60
    entity_ids(10) = FT_CONNECTION_PROP  ' 72
    entity_ids(11) = FT_LAYUP            ' 81

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
    rc = objSet.Select(entity_ids(dlg.entity_type), True, select_prompt)
    If rc = 2 Or rc = 4 Or objSet.Count = 0 Then
        App.feAppMessage(FCM_ERROR, "Nothing selected...")
        Exit Sub
    End If

    ' Prompt for new starting ID
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

        App.feRenumber(entity_ids(dlg.entity_type), dSet.ID, n)
        newObjSet.Add(n)
        n = n + 1
    Next i

    ' renumber IDs to selected block in consecutive order
    App.feRenumber(entity_ids(dlg.entity_type), newObjSet.ID, startID)

End Sub


Private Function SetEntityType(obj As Object, type_option As Long, prompt As String)
        Select Case type_option
        Case 0
            Set obj = App.feCSys
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
            Set obj = App.feGroup
            select_prompt = "Groups..."
        Case 5
            Set obj = App.feOutputSet
            select_prompt = "output Sets..."
        Case 6
            Set obj = App.feLayer
            select_prompt = "Layers..."
        Case 7
            Set obj = App.feSolid
            select_prompt = "Solids..."
        Case 8
            Set obj = App.feConnectionRegion
            select_prompt = "Connection Regions..."
        Case 9
            Set obj = App.feAnalysisMgr
            select_prompt = "Analysis Sets..."
        Case 10
            Set obj = App.feConnectionProp
            select_prompt = "Connection Properties..."
        Case 11
            Set obj = App.feLayup
            select_prompt = "Layups..."
    End Select
End Function
