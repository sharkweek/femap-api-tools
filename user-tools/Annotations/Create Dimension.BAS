' Title: Move Annotation Pointers.BAS
' Author: Andy Perez
' Femap API Version: 12.0

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND,"Create Dimension")
    App.feAppMessage(FCM_ERROR, "Creates a dimension across the selected vector.")
    App.feAppMessage(FCM_WARNING, "Ensure that the current view is activated, or dimensions arrows will not be displayed.")

    Dim rc As Long

    Dim txt As femap.text
    Set txt = App.feText

    Dim arrow As femap.GFXArrow
    Set arrow = App.feGFXArrow

    Dim gfxSet As femap.Set
    Set gfxSet = App.feSet

    Dim textLoc As Variant
    Dim v_base As Variant
    Dim v_dir As Variant
    Dim dist As Double
    Dim distString As String

    rc = FE_OK
    Do While rc = FE_OK
    'prompt user to pick annotation points
        rc = App.feVectorPick("Select dimension...", False, dist, v_base, v_dir)
        If rc = FE_CANCEL Then
            Exit Sub
        End If

        ' determine midpoint for text location and distance between coords
        ReDim textLoc(3)
        textLoc(0) = dist * v_dir(0) / 2 + v_base(0)
        textLoc(1) = dist * v_dir(1) / 2 + v_base(1)
        textLoc(2) = dist * v_dir(2) / 2 + v_base(2)

        ' rc = App.feFormatReal(dist, 8, 8, 0, distString)
        distString = Format$(dist, "0.0##E+00")

        'create text and arrow
        App.feTextPut(txt.NextEmptyID, False, True, True, 0, 0, 0, App.feLayer.Active, 1, 0, 20509, 124, textLoc, textLoc, distString)
        arrow.PutAll(arrow.NextEmptyID, v_base(0), v_base(1), v_base(2), v_dir(0), v_dir(1), v_dir(2), dist, 0, App.feLayer.Active, 20509, 2)

        ' enable gfx objects
        gfxSet.AddRange(0, 1e4, 1)
        App.feGFXSelect(gfxSet.ID, True, True)
        App.feGFXEnableView(False, App.feView.Active)
        App.feViewRegenerate(0)  'regenerate active view so as to display textboxes
        App.feAppMessage(FCM_NORMAL, "Dimension created...")
    Loop
End Sub
