Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim rc As Long

    Dim ldSet As femap.Set
    Dim spcSet As femap.Set
    Dim elmSet As femap.Set
    Dim nodeSet As femap.Set
    Dim matSet As femap.Set
    Dim propSet As femap.Set
    Dim csysSet As femap.Set
    Dim lyrSet As femap.Set

    Set ldSet = App.feSet
    Set spcSet = App.feSet
    Set elmSet = App.feSet
    Set nodeSet = App.feSet
    Set matSet = App.feSet
    Set propSet = App.feSet
    Set csysSet = App.feSet
	Set lyrSet = App.feSet

    Dim lyr As femap.layer
    Set lyr = App.feLayer

    ' Add entities to sets
    ldSet.AddAll(FT_LOAD_DEFINITION)
    spcSet.AddAll(FT_BC_DEFINITION)
    elmSet.AddAll(FT_ELEM)
    nodeSet.AddAll(FT_NODE)
    matSet.AddAll(FT_MATL)
    propSet.AddAll(FT_PROP)
    csysSet.AddAll(FT_CSYS)
    lyrSet.AddAll(FT_LAYER)

	rc = lyr.SelectID("Pick a layer, any layer.")
	MsgBox(Str(lyr.ID))


End Sub
