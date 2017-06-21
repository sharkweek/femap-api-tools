' Create Axes for Cylinders
' This macro creates lines through the axes of cylindrical
' surfaces on solids

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Axes from Cylindrical Solid Surfaces")

    Dim surf As femap.Surface
    Set surf = App.feSurface

    Dim surfSet As femap.Set
    Set surfSet = App.feSet

    Dim skippedSet As femap.Set
    Set skippedSet = App.feSet

    Dim rad As Double
    Dim aLength As Double
    Dim cBase As Variant
    Dim nBase(2) As Variant
    Dim aVect As Variant
    Dim cEnd(2) As Variant
    Dim pOffset As Double

    Dim tracker As femap.TrackData
    Set tracker = App.feTrackData

    Dim cTrack As femap.Set
    Set cTrack = App.feSet


    Dim checkTol As Double

    Dim rc As Long

    ' Begin tracking created curves
    tracker.Start(4)

    ' Select cylindrical surfaces
    rc = surfSet.Select(FT_SURFACE, True, "Select cylindrical surfaces to create axes")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "One or more surfaces do not exist")
        Exit Sub
    End If

    ' Prompt user to select length of axis and where to center
    Begin Dialog UserDialog 270,105,"Axis Length" ' %GRID:10,7,1,1
        text 30,14,100,21,"Axis length",.Text1
        text 30,42,100,14,"Percent Offset",.Text2
        TextBox 140,14,100,14,.axisLength
        TextBox 140,42,100,14,.offsetValue
        OKButton 20,70,100,28
        CancelButton 140,70,100,28
    End Dialog
    Dim dlg As UserDialog
    Dialog dlg

    ' Check axis length
    If IsNumeric(dlg.axisLength) Then
        aLength = CDbl(dlg.axisLength)
    Else
        App.feAppMessage(FCM_ERROR, "Axis length must be a numeric, non-zero value.")
        Exit Sub
    End If

    ' Check for baseOffset option and offsetValue
    If dlg.offsetValue = "" Then
        pOffset = 0.0
    ElseIf IsNumeric(dlg.offsetValue) = False _
      Or CDbl(dlg.offsetValue) > 1.0 _
      Or CDbl(dlg.offsetValue) < 0.0 Then
        App.feAppMessage(FCM_ERROR, "Offset must be a value >=0 and <=1.")
        Exit Sub
    Else
        pOffset = CDbl(dlg.offsetValue)
    End If

    ' Create axis base point and vector
    Do While surf.NextInSet(surfSet.ID)
        If surf.Cylindrical(rad, cBase, aVect) = FE_FAIL Then
            skippedSet.Add(surfSet.CurrentID)
        Else
            ' Determine end points for axis
            For i = 0 To 2
                nBase(i) = cBase(i) - aVect(i) * aLength * pOffset
                cEnd(i) = cBase(i) + aVect(i) * aLength * (1 - pOffset)
            Next i

            ' Create line
            App.feLinePoints(False, nBase, cEnd, False)
        End If
    Loop

    ' Set from created curves
    tracker.Stop(4)
    tracker.Created(4, cTrack.ID, False)

    ' Merge created coincident points and curves
    checkTol = 0.01 * aLength  ' check tolerance as 1% of specified axis length
    App.feCheckCoincidentCurve(cTrack.ID, checkTol)

    App.feAppMessage(FCM_NORMAL, cTrack.Count() & " curve(s) created...")

    If skippedSet.Count() > 0 Then
        App.feAppMessage(FCM_ERROR, skippedSet.count() & " surfaces were non-cylindrical and skipped.")
    End If

End Sub
