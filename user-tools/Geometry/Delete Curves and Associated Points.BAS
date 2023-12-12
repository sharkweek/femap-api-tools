' Title: Delete Curves and Associated Points.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2020.2

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

    App.feAppMessage(FCM_COMMAND, "Delete Curves and Associated Points")

    Dim rc As Long

    Dim cset As femap.Set
    Set cset = App.feSet()

    Dim crv As femap.Curve
    Set crv = App.feCurve()

    Dim pset As femap.Set
    Set pset = App.feSet()

    ' prompt user to select curves
    rc = cset.Select(FT_CURVE, True, "Select curves...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "No curves selected. Routine canceled...")
    End If

    cset.Reset()
    Do While cset.Next()
        crv.Get(cset.CurrentID)
        pset.Clear()
        crv.PointsAsSet(pset.ID, True)
        crv.Delete(crv.ID)
        App.feDelete(FT_POINT, pset.ID)
    Loop

    App.feViewRegenerate(0)

End Sub