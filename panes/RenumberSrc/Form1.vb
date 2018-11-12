Imports System.ComponentModel

Public Class Form1

    Dim App As femap.model
    Dim AppLoaded As Boolean = False
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim rc As Long

        'Try to connect to active FEMAP session' 
        Try
            App = GetObject(, "femap.model")
        Catch ex As Exception
            MsgBox("Error Connecting to FEMAP, Exiting...", MsgBoxStyle.OkOnly)
            Environment.Exit(0)
        End Try

        'Hook form in as FEMAP pane docked right to the main windows'
        rc = App.feAppRegisterAddInPane(True, Me.Handle, Me.Handle, True, True, 2, 0)

        If rc = femap.zReturnCode.FE_OK Then
            AppLoaded = True
        End If

        'Position buttons' 
        PositionButton()

        'Fill grid'
        ResultsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill

    End Sub

    'Exit'
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles ExitButton.Click

        'Disconnect Pane'
        If AppLoaded Then
            App.feAppRegisterAddInPane(False, Me.Handle, Me.Handle, True, True, 2, 0)
        End If

        End

    End Sub

    'Position Buttons' 
    Private Sub PositionButton()

        ResultsGrid.Width = Me.ClientSize.Width - 40
        ResultsGrid.Height = Me.ClientSize.Height - 80

        ResultsGrid.Top = 15
        ResultsGrid.Left = (Me.ClientSize.Width / 2) - (ResultsGrid.Width / 2)

        ExitButton.Top = ResultsGrid.Top + ResultsGrid.Height + 10
        ExitButton.Left = ResultsGrid.Left + ResultsGrid.Width - ExitButton.Width

        LoadButton.Top = ResultsGrid.Top + ResultsGrid.Height + 10
        LoadButton.Left = ResultsGrid.Left

        RenumberButton.Top = ResultsGrid.Top + ResultsGrid.Height + 10
        RenumberButton.Left = LoadButton.Left + LoadButton.Width + 5

        ResetButton.Top = ResultsGrid.Top + ResultsGrid.Height + 10
        ResetButton.Left = RenumberButton.Left + RenumberButton.Width + 5

        ResultsGrid.AutoResizeColumn(0)
        ResultsGrid.AutoResizeColumn(1)
        ResultsGrid.AutoResizeColumn(2)

    End Sub

    'Resize Callback'
    Private Sub Form1_Resize(ByVal sender As Object,
        ByVal e As System.EventArgs) Handles MyBase.Resize
        PositionButton()
    End Sub

    'Load Results'
    Private Sub LoadButton_Click(sender As Object, e As EventArgs) Handles LoadButton.Click

        Dim oSet As femap.Set
        Dim outputSet As femap.OutputSet
        oSet = App.feSet
        outputSet = App.feOutputSet

        oSet.AddAll(femap.zDataType.FT_OUT_CASE)

        If oSet.Count = 0 Then
            App.feAppMessage(femap.zMessageColor.FCM_NORMAL, "No output sets available for loading...")
            GoTo Done
        End If

        While outputSet.Next
            Me.ResultsGrid.Rows.Add(outputSet.ID, outputSet.title)
        End While

        ResultsGrid.AutoResizeColumn(0)
        ResultsGrid.AutoResizeColumn(1)
        ResultsGrid.AutoResizeColumn(2)

Done:
    End Sub
    Private Sub RenumberButton_Click(sender As Object, e As EventArgs) Handles RenumberButton.Click

        Dim numRows As Long
        Dim oldID As Long
        Dim newID As Long
        Dim rc As Long
        Dim rowIndex As Long
        Dim outputSet As femap.OutputSet
        Dim oSet As femap.Set
        outputSet = App.feOutputSet
        oSet = App.feSet

        numRows = ResultsGrid.Rows.Count - 1

        If numRows = 0 Then
            App.feAppMessage(femap.zMessageColor.FCM_NORMAL, "No output sets available for renumbering...")
            GoTo Done
        End If

        rowIndex = 0

        oSet.AddAll(femap.zDataType.FT_OUT_CASE)
        oSet.Reset()

        While oSet.Next
            rc = outputSet.Get(oSet.CurrentID)
            oldID = ResultsGrid.Rows(rowIndex).Cells(0).Value

            If outputSet.ID = oldID Then
                Try
                    newID = ResultsGrid.Rows(rowIndex).Cells(2).Value
                Catch ex As Exception
                    App.feAppMessage(femap.zMessageColor.FCM_ERROR, "ID not valid, please enter valid ID...")
                    GoTo Done
                End Try

                rc = App.feRenumber(femap.zDataType.FT_OUT_CASE, -outputSet.ID, newID)
                ResultsGrid.Rows(rowIndex).Cells(0).Value = newID
            End If

            rowIndex += 1
        End While

Done:

    End Sub

    'Reset'
    Private Sub ResetButton_Click(sender As Object, e As EventArgs) Handles ResetButton.Click
        ResultsGrid.Rows.Clear()
    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        Dim rc As Long

        Try
            App = GetObject(, "femap.model") 'Attempting to re-establish connection with FEMAP, just to unregister
            rc = App.feAppRegisterAddInPane(False, Me.Handle, Me.Handle, True, True, 2, 0) 'Disconnect our pane
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        End

    End Sub
End Class
