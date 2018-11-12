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
        Me.CaptureMsgButton = New System.Windows.Forms.Button()
        Me.ListModelInfoButton = New System.Windows.Forms.Button()
        Me.ReleaseMsgButton = New System.Windows.Forms.Button()
        Me.FemapMsgHandlerListBox = New System.Windows.Forms.ListBox()
        Me.ExitButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'CaptureMsgButton
        '
        Me.CaptureMsgButton.Location = New System.Drawing.Point(12, 12)
        Me.CaptureMsgButton.Name = "CaptureMsgButton"
        Me.CaptureMsgButton.Size = New System.Drawing.Size(75, 23)
        Me.CaptureMsgButton.TabIndex = 0
        Me.CaptureMsgButton.Text = "Capture"
        Me.CaptureMsgButton.UseVisualStyleBackColor = True
        '
        'ListModelInfoButton
        '
        Me.ListModelInfoButton.Location = New System.Drawing.Point(12, 41)
        Me.ListModelInfoButton.Name = "ListModelInfoButton"
        Me.ListModelInfoButton.Size = New System.Drawing.Size(75, 23)
        Me.ListModelInfoButton.TabIndex = 1
        Me.ListModelInfoButton.Text = "List Info"
        Me.ListModelInfoButton.UseVisualStyleBackColor = True
        '
        'ReleaseMsgButton
        '
        Me.ReleaseMsgButton.Location = New System.Drawing.Point(12, 70)
        Me.ReleaseMsgButton.Name = "ReleaseMsgButton"
        Me.ReleaseMsgButton.Size = New System.Drawing.Size(75, 23)
        Me.ReleaseMsgButton.TabIndex = 2
        Me.ReleaseMsgButton.Text = "Release"
        Me.ReleaseMsgButton.UseVisualStyleBackColor = True
        '
        'FemapMsgHandlerListBox
        '
        Me.FemapMsgHandlerListBox.Font = New System.Drawing.Font("Consolas", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FemapMsgHandlerListBox.FormattingEnabled = True
        Me.FemapMsgHandlerListBox.Location = New System.Drawing.Point(93, 12)
        Me.FemapMsgHandlerListBox.Name = "FemapMsgHandlerListBox"
        Me.FemapMsgHandlerListBox.Size = New System.Drawing.Size(398, 173)
        Me.FemapMsgHandlerListBox.TabIndex = 3
        '
        'ExitButton
        '
        Me.ExitButton.Location = New System.Drawing.Point(12, 161)
        Me.ExitButton.Name = "ExitButton"
        Me.ExitButton.Size = New System.Drawing.Size(75, 23)
        Me.ExitButton.TabIndex = 4
        Me.ExitButton.Text = "Exit"
        Me.ExitButton.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(500, 196)
        Me.Controls.Add(Me.ExitButton)
        Me.Controls.Add(Me.FemapMsgHandlerListBox)
        Me.Controls.Add(Me.ReleaseMsgButton)
        Me.Controls.Add(Me.ListModelInfoButton)
        Me.Controls.Add(Me.CaptureMsgButton)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents CaptureMsgButton As Button
    Friend WithEvents ListModelInfoButton As Button
    Friend WithEvents ReleaseMsgButton As Button
    Friend WithEvents FemapMsgHandlerListBox As ListBox
    Friend WithEvents ExitButton As Button
End Class
