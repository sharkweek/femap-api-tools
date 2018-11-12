Public Class Form1
    Dim App As femap.model              'Global Variable to the FEMAP Connection/Sessions
    Dim AppLoaded As Boolean = False    'Global Variable True once the Pane has been registered and connected

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim rc As Long

        Dim pSet As femap.Set
        pSet = App.feSet

        Dim properties As Object

        Dim p As femap.Prop
        p = App.feProp

        Try
            App = GetObject(, "femap.model") 'Connect to the active running FEMAP Session
        Catch ex As Exception
            MsgBox("Error Connecting to FEMAP, Exiting.....", MsgBoxStyle.OkOnly) 'If we can't connect exit gracefully
            Environment.Exit(0)
        End Try

        rc = App.feAppRegisterAddInPane(True, Me.Handle, Me.Handle, True, True, 2, 0) 'Hook our Form in as a FEMAP pane, last two arguments are 2 - dock right, 0 - to the FEMAP Main Windows
        If rc = femap.zReturnCode.FE_OK Then
            AppLoaded = True 'Yes we did it
        End If

        rc = pSet.AddAll(femap.zDataType.FT_PROP)
        rc = pSet.GetArray(pSet.Count, properties)



        For Each prp As Long In properties
            p.Get(prp)
            Me.DataGridView1.Rows.Add(p.ID, p.title)
        Next

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        If AppLoaded Then
            App.feAppRegisterAddInPane(False, Me.Handle, Me.Handle, True, True, 2, 0) 'Disconnect our pane
        End If
        Close() 'Exit
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

End Class
