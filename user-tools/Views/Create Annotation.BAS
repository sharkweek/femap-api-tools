Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_ERROR,"Create Annotation")

    Dim rc As Long

    Dim txt As femap.text
    Set txt = App.feText

    Dim textLoc As Variant
    Dim pointerLoc As Variant

    Dim curLayer As femap.layer
    Set curLayer = App.feLayer

    'prompt user to pick annotation points
    App.feCoordPick("Location of annotation text.", textLoc)
    App.feCoordPick("Location of pointer", pointerLoc)

    'prompt user to enter annotation text
    Begin Dialog UserDialog 380,70,"Enter annotation text..." ' %GRID:10,7,1,1
        TextBox 20,14,340,21,.Annotation
        OKButton 40,42,130,21
        CancelButton 210,42,130,21
    End Dialog
    Dim dlg As UserDialog
    Dialog dlg

    'set annotation properties
    With txt
        .ModelPosition = True
        .AllViews = True
        .DrawBorder = True
        .BorderColor = FCL_BLACK
        .BackColor = FCL_WHITE
        .color = FCL_BLACK
        .layer = curLayer.Active
        .FontNumber = 4
        .DrawPointer = True
        .ID = App.Info_NextID(FT_TEXT)
        .vTextPosition = textLoc
        .vPointerPosition = pointerLoc
        .text = dlg.Annotation
        .Put(txt.ID) 'save annotation to model
        App.feAppMessage(FCM_NORMAL, "Annotation created...")
    End With

    App.feViewRegenerate(0)  'regenerate active view so as to display textboxes

End Sub
