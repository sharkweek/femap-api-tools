' Export Results to Femap Neutral Output (Full Model)

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim rc As Long

    Dim modelOutputSets As femap.OutputSet
    Set modelOutputSets = App.feOutputSet

    Dim ouSets As femap.Set
    Set ouSets = App.feSet

    Dim fno_Name As String

    Dim ouSetID As Long

    Dim tempSet As femap.Set
    Set tempSet = App.feSet

    ' Prompt user to select output sets
    rc = ouSets.Select( FT_OUT_CASE, True, "Select Output Sets" )
    If rc = FE_CANCEL Then
        Exit Sub
    End If

    ' Write out *.fno
    If ouSets.Count > 0 Then
            ouSetID = ouSets.First
            While ouSetID > 0
                ' find corresponding output set from model
                modelOutputSets.Get(ouSetID)

                ' create file name based on output set ID and title
                fno_Name = Str(ouSetID) + ".." + modelOutputSets.title + ".fno"

                ' Create dummy set to include the id of the output set
                tempSet.Clear
                tempSet.Add( ouSetID )

                ' Write out the output set in an individual file
                rc = App.feFileWriteFNO( tempSet.ID, 0, 0, fno_Name )

                ' Cycle set ID
                ouSetID = ouSets.Next
            Wend
    End If

End Sub
