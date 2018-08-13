Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Smart Mesh Region on Surface")

    Dim rc As Long

    Dim propArray As Variant
    Dim propIDArray() As Long

    Dim propSet As femap.Set
    Set propSet = App.feSet

    Dim p As femap.Prop
    Set p = App.feProp

    Dim i As Long

    ' create property array for dialog box
    For i = 10 To 22
        propSet.AddRule(i, FGD_Prop_byType)
    Next i
    propSet.AddRule(22, FGD_Prop_byType)  ' add plot-only plates

    rc = propSet.GetArray(i, propArray)

    For i = 1 to propSet.Count()
        ReDim Preserve propIDArray(i) As Long

        propIDArray(i - 1) =

    Begin Dialog UserDialog 0,0,660,203,"Smart Surface Mesh Region",.DialogFunc ' %GRID:10,7,1,1
        DropListBox 40,14,300,21,propArray(),.prop
        OptionGroup .meshApproach
            OptionButton 20,126,90,14,"Triangle",.triangle
            OptionButton 20,154,90,14,"Quad",.quad
            OptionButton 20,182,90,14,"Fast Triangle",.fastTriangle
        TextBox 140,98,90,21,.minBetween
        TextBox 140,140,90,21,.maxAspectRatio
        TextBox 260,63,90,21,.maxQuickNodes
        TextBox 260,98,90,21,.quadAngle
        OptionGroup .smoothing
            OptionButton 20,49,90,14,"Laplacian",.Laplacian
            OptionButton 20,70,90,14,"Centroidal",.Centroidal
        TextBox 140,63,90,21,.MeshSize
        TextBox 270,140,90,21,.smoothIter
        TextBox 390,70,90,21,.smoothTol
        PushButton 500,70,90,21,"Node Set 1",.nodeSet1
        PushButton 500,119,90,21,"Node Set 2",.nodeSet2
        OKButton 120,168,90,21
        CancelButton 260,168,90,21
    End Dialog
    Dim dlg As UserDialog
    Dialog dlg

    App.feMeshRemeshRegions(meshSize,
                            propArray(dlg.prop),
                            dlg.minBetween,
                            dlg.maxAspect,
                            dlg.maxQuickNodes,
                            dlg.meshApproach,
                            dlg.quadAngle,
                            dlg.Laplacian,
                            dlg.smoothIter,
                            dlg.smoothTol)

End Sub


Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
