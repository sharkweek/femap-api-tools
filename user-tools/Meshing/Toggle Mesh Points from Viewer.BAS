' Title: Toggle Mesh Points from Viewer.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2021.2

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' public declarations
Dim App As femap.model

Sub Main
    ' `Main` declarations
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Toggle Mesh Points from Viewer")

    Dim rc As Long

    Dim mp As femap.MeshHardPoint
    Set mp = App.feMeshHardPoint()

    Dim mp_set As femap.Set
    Set mp_set = App.feSet()

    ' prompt user to select mesh points
    rc = mp_set.Select(FT_MESH_POINT, True, "Select mesh points...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "No mesh points selected. Routine canceled...")
    End If

    Do While mp_set.Next()
        mp.Get(mp_set.CurrentID)
        If mp.IsEnabled Then
            mp.Enable(False)
            mp.color = 26882
        Else
            mp.Enable(True)
            mp.color = 24578
        End If
        mp.Put(mp.ID)
    Loop

    App.feViewRegenerate(0)

End Sub