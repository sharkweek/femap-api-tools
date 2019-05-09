' Title: Move Annotation Pointers.BAS
' Author: Andy Perez
' Date: May 2019
' Femap API Version: 12.0

Sub Main
    feFemap().feAppMessage(FCM_COMMAND,"Create Dimension")

    Dim rc As Long

    Dim txt As femap.text
    Set txt = feFemap().feText

    Dim arrow As femap.GFXArrow
    Set arrow = feFemap().feGFXArrow

    Dim textLoc As Variant
    Dim v_base As Variant
    Dim v_dir As Variant
    Dim dist As Double
    Dim distString As String

    'prompt user to pick annotation points
    If feFemap().feVectorPick("Select dimension...", False, dist, v_base, v_dir) = FE_CANCEL Then
        Exit Sub
    End If

    ' determine midpoint for text location and distance between coords
    ReDim textLoc(3)
    textLoc(0) = dist * v_dir(0) / 2 + v_base(0)
    textLoc(1) = dist * v_dir(1) / 2 + v_base(1)
    textLoc(2) = dist * v_dir(2) / 2 + v_base(2)

    ' rc = feFemap().feFormatReal(dist, 8, 8, 0, distString)
    distString = Format$(dist, "0.0##E+00")

    'create text and arrow
    feFemap().feTextPut(txt.NextEmptyID, False, True, True, 0, 0, 0, feFemap().feLayer.Active, 1, 0, FCL_YELLOWLIMEGREEN, 124, textLoc, textLoc, distString)
    arrow.PutAll(arrow.NextEmptyID, v_base(0), v_base(1), v_base(2), v_dir(0), v_dir(1), v_dir(2), dist, 0, feFemap().feLayer.Active, FCL_YELLOWLIMEGREEN, 2)

    ' enable gfx objects
    feFemap().feGFXEnableView(False, feFemap().feView.Active)
    feFemap().feGFXSelect(arrow.ID, True, True)

    feFemap().feViewRegenerate(0)  'regenerate active view so as to display textboxes
    feFemap().feAppMessage(FCM_NORMAL, "Dimension created...")

End Sub
