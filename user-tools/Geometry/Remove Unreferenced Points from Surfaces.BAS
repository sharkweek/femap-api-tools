' Title: Remove Unreferenced Points from Surfaces.BAS
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

    App.feAppMessage(FCM_COMMAND, "Remove Unreferenced Points from Surfaces")
    App.feAppMessage(FCM_NORMAL, "Removes all unreferenced from surfaces in the model, saves, and rebuilds the *.modfem.")

    Dim rc As Long

    Dim surf_set As femap.Set
    Set surf_set = App.feSet

    Dim surf As femap.Surface
    Set surf = App.feSurface

    Dim pnt As femap.Point
    Set pnt = App.fePoint

    Dim purge_set As femap.Set
    Set purge_set = App.feSet

    Dim point_set As femap.Set
    Set point_set = App.feSet

    ' get all surfaces
    surf_set.AddAll(FT_SURFACE)
    surf_set.Reset()

    ' check each surface for orphaned points
    Do While surf.NextInSet(surf_set.ID)
        point_set.Clear()
        surf.Get(surf.ID)
        surf.PointsAsSet(2, point_set.ID, True)

        ' add orphaned points to set for deletion
        Do While pnt.NextInSet(point_set.ID)
            if pnt.Exist(pnt.ID) = FE_FAIL Then
                purge_set.Add(pnt.ID)
            End If
        Loop
    Loop

    If purge_set.Count() = 0 Then
        App.feAppMessage(FCM_ERROR, "No orphaned points in model.")
    ElseIf App.feAppMessageBox(1, "Are you sure you wish to delete " & CStr(purge_set.Count()) & " points?") = FE_OK Then
        App.feDelete(FT_POINT, purge_set.ID)
    End If

    If App.feAppMessageBox(1, "Save and rebuild?") = FE_OK Then
        App.feFileSave(True)
        App.feFileRebuild(True, True)
    End If

    App.feViewRegenerate(0)

End Sub