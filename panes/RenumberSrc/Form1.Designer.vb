<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ExitButton = New System.Windows.Forms.Button()
        Me.ResultsGrid = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LoadButton = New System.Windows.Forms.Button()
        Me.RenumberButton = New System.Windows.Forms.Button()
        Me.ResetButton = New System.Windows.Forms.Button()
        CType(Me.ResultsGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ExitButton
        '
        Me.ExitButton.Location = New System.Drawing.Point(358, 778)
        Me.ExitButton.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.ExitButton.Name = "ExitButton"
        Me.ExitButton.Size = New System.Drawing.Size(88, 26)
        Me.ExitButton.TabIndex = 0
        Me.ExitButton.Text = "Exit"
        Me.ExitButton.UseVisualStyleBackColor = True
        '
        'ResultsGrid
        '
        Me.ResultsGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.ResultsGrid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.Column3})
        Me.ResultsGrid.Location = New System.Drawing.Point(9, 23)
        Me.ResultsGrid.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.ResultsGrid.Name = "ResultsGrid"
        Me.ResultsGrid.RowTemplate.Height = 24
        Me.ResultsGrid.Size = New System.Drawing.Size(437, 740)
        Me.ResultsGrid.TabIndex = 1
        '
        'Column1
        '
        Me.Column1.HeaderText = "Result ID"
        Me.Column1.Name = "Column1"
        '
        'Column2
        '
        Me.Column2.HeaderText = "Title"
        Me.Column2.Name = "Column2"
        '
        'Column3
        '
        Me.Column3.HeaderText = "New Result ID"
        Me.Column3.Name = "Column3"
        '
        'LoadButton
        '
        Me.LoadButton.Location = New System.Drawing.Point(9, 778)
        Me.LoadButton.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.LoadButton.Name = "LoadButton"
        Me.LoadButton.Size = New System.Drawing.Size(88, 26)
        Me.LoadButton.TabIndex = 2
        Me.LoadButton.Text = "Load Results"
        Me.LoadButton.UseVisualStyleBackColor = True
        '
        'RenumberButton
        '
        Me.RenumberButton.Location = New System.Drawing.Point(108, 778)
        Me.RenumberButton.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.RenumberButton.Name = "RenumberButton"
        Me.RenumberButton.Size = New System.Drawing.Size(88, 26)
        Me.RenumberButton.TabIndex = 3
        Me.RenumberButton.Text = "Renumber"
        Me.RenumberButton.UseVisualStyleBackColor = True
        '
        'ResetButton
        '
        Me.ResetButton.Location = New System.Drawing.Point(201, 778)
        Me.ResetButton.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.ResetButton.Name = "ResetButton"
        Me.ResetButton.Size = New System.Drawing.Size(88, 26)
        Me.ResetButton.TabIndex = 4
        Me.ResetButton.Text = "Reset"
        Me.ResetButton.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(455, 813)
        Me.Controls.Add(Me.ResetButton)
        Me.Controls.Add(Me.RenumberButton)
        Me.Controls.Add(Me.LoadButton)
        Me.Controls.Add(Me.ResultsGrid)
        Me.Controls.Add(Me.ExitButton)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Name = "Form1"
        Me.Text = "Renumber Output Sets"
        CType(Me.ResultsGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ExitButton As Button
    Friend WithEvents ResultsGrid As DataGridView
    Friend WithEvents LoadButton As Button
    Friend WithEvents RenumberButton As Button
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
    Friend WithEvents Column2 As DataGridViewTextBoxColumn
    Friend WithEvents Column3 As DataGridViewTextBoxColumn
    Friend WithEvents ResetButton As Button
End Class
