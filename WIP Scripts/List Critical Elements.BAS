'TODO: have the user select the envelope output set, and then the envelop info set to match critical element to value and load case

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim rc As Long

    Dim pickedSets As femap.Set
    Dim pickedVectors As femap.Set
    Dim e As femap.Set
    Set e = App.feSet
    Dim ov As femap.output
    Set ov = App.feOutput
    Dim minID As Long
    Dim maxID As Long
    Dim minVAL As Double
    Dim maxVAL As Double

   ' Create a set to contain the selected elements
    Dim elmSet As femap.Set
    Set elmSet = App.feSet

    rc = App.feSelectOutput( "Select Output Vectors", 0, FOT_ANY, FOC_ANY, FT_ELEM, False, pickedSets, pickedVectors )

    ' Check for selected output
    If rc <> FE_OK Then
        Exit Sub
    End If

    ' Select elements to include
    rc = elmSet.Select( FT_ELEM, True, "Select Elements" )

    ' Check for selected output
    If rc <> FE_OK Then
        Exit Sub
    End If

    While pickedSets.Next()
        pickedVectors.Reset()
        ' Go to the next output vector
        While pickedVectors.Next()
            ' Get the set and vector ID's
            OutputVector.GetFromSet( pickedSets.CurrentID(), pickedVectors.CurrentID() )
            ' Get the set and vector names
            SetName.SetTitle( pickedSets.CurrentID , SNListTitle )
            SetName.VectorTitle( pickedSets.CurrentID, pickedVectors.CurrentID, VNListTitle )
            ' Determine the max and min values per set and per loadcase
            OutputVector.FindMaxMin(elmSet.ID, True, minID, minVal, maxID, maxVal)
        Wend
    Wend


    'write to message window
    While pickedSets.Next()
        pickedVectors.Reset()
        While pickedVectors.Next()
            ov.GetFromSet( pickedSets.CurrentID(), pickedVectors.CurrentID() )
            ov.FindMaxMin(e.ID, True, minID, minVAL, maxID, maxVAL)
            App.feAppMessage( FCM_NORMAL, Str$(pickedSets.CurrentID)+" "+Str$(pickedVectors.CurrentID)+" "+Str$(minID)+" "+Str$(minVAL)+" "+Str$(maxID)+" "+Str$(maxVAL) )
        Wend
    Wend

End Sub