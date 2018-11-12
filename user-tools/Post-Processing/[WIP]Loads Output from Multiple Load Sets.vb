Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Loads Output from Multiple Load Sets")
    App.feAppMessage(FCM_NORMAL, "Batch create output set data from multiple load sets.")

    Dim rc As Long

    Dim lsSet As femap.Set
    Set lsSet = App.feSet

    Dim ls As femap.LoadSet
    Set ls = App.feLoadSet

    Dim lmesh As femap.LoadMesh
    Set lmesh = App.feLoadMesh

    Dim eSet As femap.Set
    Set eSet = App.feSet

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim currentls As Long

    Dim loadTypes(17) As Long
    loadtypes(0) = 16 ' force
    loadtypes(1) = 32 ' displacement
    loadtypes(2) = 2048 ' velocity
    loadtypes(3) = 64 ' acceleration
    loadtypes(4) = 4096 ' nodal temperature
    loadtypes(5) = 2 ' heat generation
    loadtypes(6) = 4 ' heat flux
    loadtypes(7) = 128 ' distributed load
    loadtpyes(8) = 256 ' pressure
    loadtypes(9) = 8192 ' element temperature
    loadtypes(10) = 512 ' element heat generation
    loadtypes(11) = 1024 ' element heat flux
    loadtypes(12) = 16384 ' element convection
    loadtypes(13) = 32768 ' element radiation
    loadtypes(14) = 131072 ' fluid pressure
    loadtypes(15) = 262144 ' fluid tracking
    loadtypes(16) = 2097152 ' fluid fan curve

    ' save current load set
    currentls = ls.Active

    ' prompt user to select load sets
    rc = lsSet.Select(FT_LOAD_DIR, True, "Select load sets...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_FAIL Then
        rc = App.feAppMessage(FCM_ERROR, "Selected load sets do not exist.")
    End If

    ' loop through load sets
    Do While ls.NextInSet(lsSet.ID)
        ' get all elements and nodes with loads applied
        For i = 0 To 16
        
        Next i
        App.feOutputFromLoad(elmnodeSetID, lmesh.LoadType)
    Loop

End Sub
