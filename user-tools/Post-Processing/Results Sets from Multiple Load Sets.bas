' Title: Results Sets from Multiple Load Sets.bas
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 12.0

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Results Sets from Multiple Load Sets")
    App.feAppMessage(FCM_NORMAL, "Create output set data from multiple load sets.")

    Dim rc As Long

    Dim lsSet As femap.Set
    Set lsSet = App.feSet

    Dim ls As femap.LoadSet
    Set ls = App.feLoadSet

    Dim os As femap.OutputSet
    Set os = App.feOutputSet

    Dim eSet As femap.Set
    Set eSet = App.feSet

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim og_ls As Long
    Dim og_os As Long

    Dim osid As Long
    Dim written As Long

    Dim loadType(14) As Long
    ' nodal loads
    loadType(0) = 16  ' force
    loadType(1) = 32  ' displacement
    loadType(2) = 2048  ' velocity
    loadType(3) = 64  ' acceleration
    loadType(4) = 4096  ' temperature
    loadType(5) = 2  ' heat generation
    loadType(6) = 4  ' heat flux
    ' elemental loads
    loadType(7) = 128  ' elemental distributed load
    loadType(8) = 256  ' pressure
    loadType(9) = 8192  ' temperature
    loadType(10) = 512  ' heat generation
    loadType(11) = 1024  ' heat flux
    loadType(12) = 16384  ' convection
    loadType(13) = 32768  ' radiation

    Dim loadTypeName(14) As String
    ' nodal loads
    loadTypeName(0) = "1..Nodal Force"
    loadTypeName(1) = "2..Nodal Displacement"
    loadTypeName(2) = "3..Nodal Velocity"
    loadTypeName(3) = "4..Nodal Acceleration"
    loadTypeName(4) = "5..Nodal Temperature"
    loadTypeName(5) = "6..Nodal Heat Generation"
    loadTypeName(6) = "7..Nodal Heat Flux"
    ' elemental loads
    loadTypeName(7) = "8..Elemental Distributed Load"
    loadTypeName(8) = "9..Elemental Pressure"
    loadTypeName(9) = "10..Elemental Temperature"
    loadTypeName(10) = "11..Elemental Heat Generation"
    loadTypeName(11) = "12..Elemental Heat Flux"
    loadTypeName(12) = "13..Elemental Convection"
    loadTypeName(13) = "14..Elemental Radiation"

    Begin Dialog UserDialog 320,70 ' %GRID:10,7,1,1
        DropListBox 10,14,300,21,loadTypeName(),.LoadType
        OKButton 40,42,110,21
        CancelButton 170,42,110,21
    End Dialog
    Dim dlg As UserDialog
    Dialog dlg

    ' save current load and output sets
    og_ls = ls.Active
    og_os = os.Active

    ' prompt user if they wish to create vectors in new output set
    rc = App.feAppMessageBox(3, "Create new output set?")
    If rc = FE_OK Then
        ' get title for new output set
        Begin Dialog UserDialog 320,70,"New Output Set Title"
            TextBox 10,14,300,21,.NewTitle
            OKButton 40,42,110,21
            CancelButton 170,42,110,21
        End Dialog
        Dim title_dlg As UserDialog
        Dialog title_dlg

        osid = os.NextEmptyID()
        os.title = title_dlg.NewTitle
        os.Put(osid)

    ElseIf rc = FE_FAIL Then
        ' prompt user to select output set
        If os.SelectID("Select output set...") = FE_CANCEL Then
            Exit Sub
        End If
        osid = os.ID

    Else
        Exit Sub

    End If

    os.Active = osid

    ' prompt user to select load sets
    rc = lsSet.Select(FT_LOAD_DIR, True, "Select load sets...")
    If rc = FE_CANCEL Then
        os.Delete(osid)
        Exit Sub
    ElseIf rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "Selected load sets do not exist.")
    End If

    ' prompt user to select elements
    rc = eSet.Select(FT_ELEM, True, "Select elements...")
    If rc = FE_CANCEL Then
        os.Delete(osid)
        Exit Sub
    ElseIf rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "Selected elements do not exist.")
    End If

    ' find nodes on elements
    nSet.AddRule(FT_NODE, FGD_NODE_ONELEM)

    ' create output vectors from each load set
    written = 0
    App.feAppMessage(FCM_NORMAL, "Writing output vectors...")
    Do While ls.NextInSet(lsSet.ID)
        ls.Active = ls.ID
        If App.feOutputFromLoad(eSet.ID, loadType(dlg.LoadType)) = FE_OK
            written += 1
        End If
    Loop
    App.feAppMessage(FCM_NORMAL, CStr(written) & " output vectors written to output set " & os.Title)

    ' reset to original load and output sets
    ls.Active = og_ls
    os.Active = og_os

    App.feViewRegenerate(0)

End Sub
