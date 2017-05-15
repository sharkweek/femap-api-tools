Sub Main

    Dim App As femap.model
    Set App = feFemap()

    Dim os As femap.OutputSet
    Set os = App.feOutputSet

    ' Prompt user to select output sets
    rc = os.Select( FT_OUT_CASE, True, "Select Output Sets" )
    If rc = FE_CANCEL Then
        Exit Sub
    End If

    While os.Next
        If os.Value <> 0 Then
            os.title = Str(os.Value)
            os.Put(os.ID)
        End If
    Wend

End Sub
