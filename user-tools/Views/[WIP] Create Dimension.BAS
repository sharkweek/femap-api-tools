' TODO: create saved set for user graphics objects for preserving dimensions
'       and deleting user-specified ones

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND,"Create Dimension")

    Dim rc As Long

    Dim txt As femap.text
    Set txt = App.feText

    Dim arrow As femap.GFXArrow
    Set arrow = App.feGFXArrow

    Dim aSet As femap.Set
    Set aSet = App.feSet

    Dim textLoc As Variant
    Dim pointerLoc1 As Variant
    Dim pointerLoc2 As Variant
    Dim vec As Variant
    Dim dist As Double
    Dim aid As Long

    'prompt user to pick annotation points
    App.feCoordPick("First location...", pointerLoc1)
    App.feCoordPick("Second location...", pointerLoc2)

    ' determine midpoint for text location and distance between coords
    ReDim textLoc(3)
    ReDim vec(3)
    textLoc(0) = (pointerLoc2(0) - pointerLoc1(0)) / 2 + pointerLoc1(0)
    textLoc(1) = (pointerLoc2(1) - pointerLoc1(1)) / 2 + pointerLoc1(1)
    textLoc(2) = (pointerLoc2(2) - pointerLoc1(2)) / 2 + pointerLoc1(2)
    dist = Sqr((pointerLoc2(0) - pointerLoc1(0))^2 + (pointerLoc2(1) - pointerLoc1(1))^2 + (pointerLoc2(2) - pointerLoc1(2))^2)
    vec(0) = (pointerLoc2(0) - pointerLoc1(0)) / dist
    vec(1) = (pointerLoc2(1) - pointerLoc1(1)) / dist
    vec(2) = (pointerLoc2(2) - pointerLoc1(2)) / dist

    'set annotation properties
    With txt
        .ModelPosition = True
        .AllViews = True
        .DrawBorder = True
        .BorderColor = FCL_YELLOWLIMEGREEN
        .BackColor = 0
        .color = 124
        .layer = App.feLayer.Active
        .FontNumber = 1
        .DrawPointer = False
        .ID = App.Info_NextID(FT_TEXT)
        .vTextPosition = textLoc
        .text = Format$(dist, "0.0##E+0")
        .Put(txt.ID) 'save annotation to model
    End With

    aid = arrow.NextEmptyID
    With arrow
        .vxyz = pointerLoc1
        .vdxyz = vec
        .length = dist
        .lengthmode = GAM_ABSOLUTE
        .style = GAS_DOUBLE_LINE
        .color = FCL_YELLOWLIMEGREEN
        aSet.Add(aid)
        .Put(aid)
    End With

    App.feGFXEnableView(False, App.feView.Active)
    App.feGFXSelect(aid, True, True)

    App.feViewRegenerate(0)  'regenerate active view so as to display textboxes
    App.feAppMessage(FCM_NORMAL, "Dimension created...")

End Sub
