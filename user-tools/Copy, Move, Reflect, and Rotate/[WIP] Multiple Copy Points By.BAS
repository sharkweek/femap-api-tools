' Title: [WIP] Multiple Copy Points By.BAS
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

    App.feAppMessage(FCM_COMMAND, "Multiple Copy Points By")
    App.feAppMessage(FCM_NORMAL, "Continue copying multiple points to selected locations.")

    Dim rc As Long

    Dim pSet As femap.Set
    Set pSet = App.feSet

    Dim copy As femap.CopyTool
    Set copy = App.feCopyTool

    Dim vecBase As Variantn
    Dim vecTip As Variant
    Dim alongVector(2) As Double

    ' prompt user to select points to copy
    rc = pSet.Select(FT_POINT, True, "Select points to copy...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "One or more selected points do not exist.")
        Exit Sub
    End If

    ' select base location for all copy operations
    App.feCoordPick("Select location to copy from...", vecBase)

    ' keep copying until user cancels
    Do While App.feCoordPick("Select location to copy to...", vecTip) <> FE_CANCEL
        ' calculate new vector based on copy-to location
        alongVector(0) = vecTip(0) - vecBase(0)
        alongVector(1) = vecTip(1) - vecBase(1)
        alongVector(2) = vecTip(2) - vecBase(2)

        ' copy
        With copy
            .Repetitions = 1
            .CopyInSameLocation = False
            .AlwaysCreateParentCSys = False
            .UsePattern =
        End With
        rc = App.feGenerateCopy2(FT_ELEM, pSet.ID, alongVector, 0, False, True)
        If rc = FE_FAIL Then
            App.feAppMessage(FCM_ERROR, "Unable to copy points.")
        End If
        App.feViewRegenerate(0)
    Loop

End Sub
