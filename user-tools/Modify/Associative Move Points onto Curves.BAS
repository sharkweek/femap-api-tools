' Title: Associative Move Points onto Curves.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2306

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Move Points Onto Curves")

    Dim rc As Long

    ' point set
    Dim pSet As femap.Set
    Set pSet = App.feSet

    ' curve to project onto
    Dim c As femap.Curve
    Set c = App.feCurve

    ' dummy point object
    Dim p As femap.Point
    Set p = App.fePoint

    Dim success As Long

    ' prompt user to select points
    rc = pSet.Select(FT_POINT, True, "Select points...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected points do not exist.")
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    ' prompt for curve selection
    rc = c.SelectID("Select curve to project onto...")
    If rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected curve does not exist.")
    ElseIf rc = FE_CANCEL Then
        Exit Sub
    End If

    ' move to curve
    rc = App.feProjectOntoCurve(FT_POINT, pSet.ID, c.ID)
    If rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "Points could not be moved to curve.")
    End If

    ' report success
    success = pSet.Count()
    App.feAppMessage(FCM_NORMAL, Cstr(success) & " points moved onto curve...")

End Sub