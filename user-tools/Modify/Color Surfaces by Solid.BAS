' Title: Color Surfaces by Solid.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2301

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

    App.feAppMessage(FCM_COMMAND, "Color Surfaces by Solid")
    App.feAppMessage(FCM_NORMAL, "Matches color of surfaces to the solid to which they belong.")

    Dim rc As Long

    Dim solid_set As femap.Set
    Set solid_set = App.feSet()

    Dim surf_set As femap.Set
    Set surf_set = App.feSet()

    Dim sol As femap.Solid
    Set sol = App.feSolid()

    Dim surf As femap.Surface
    Set surf = App.feSurface()

    ' prompt user to select solids
    rc = solid_set.Select(FT_SOLID, True, "Select solids...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "No solids selected. Routine canceled...")
    End If

    ' loop through solids and color surfaces accordingly
    solid_set.Reset()
    Do While solid_set.Next()
        sol.Get(solid_set.CurrentID)
        surf_set.Clear()
        surf_set.AddRule(solid_set.CurrentID, FGD_SURFACE_ONSOLID)
        Do While surf_set.Next()
            surf.Get(surf_set.CurrentID)
            surf.color = sol.color
            surf.Put(surf.ID)
        Loop
    Loop

    App.feViewRegenerate(0)

End Sub