Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim rc As Long

    Dim pickedSet As femap.Set
    Set pickedSet = App.feSet

    Dim pickedVect As femap.Set
    Set pickedVect = App.feSet

    Dim outSet As femap.OutputSet
    Set outSet = App.feOutputSet

    Dim cRes As Boolean

    rc = App.feSelectOutputSets( "Select output sets", pickedSet )
    
    rc = App.feResults

    ' Prompt for corner results
    If MsgBox("Would you like to include corner results?", vbYesNo, "Corner Results") = vbYes Then
        cRes = True
    Else
        cRes = False
    End If

    ' Select output sets and vectorsa
    rc = App.feSelectOutput( "Select Output Sets and Vectors", App.feOutputSet.Active, FOT_ANY, FOC_ANY, 0, cRes, pickedSet, pickedVect )

    ' Exit sub if no output sets are selected
    If pickedSet.Count < 1 Then
        Exit Sub
    End If

    pickedSet.Reset()

    While outSet.Next()

        v = App.feOutputProcessRSSCombination(False,FOPC_ALL,outSet.ID,1,outSet.ID, list,0)

        With outSet
            .ID = outSet.NextEmptyID
            .title = "OUTPUT " & outSet.ID
            .Value = v
            .Put(outSet.ID)
        End With

    Wend

End Sub