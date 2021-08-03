<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReplaceTextForm
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
      Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ReplaceTextForm))
      Me.Label1 = New System.Windows.Forms.Label()
      Me.Label2 = New System.Windows.Forms.Label()
      Me.FromTextBox = New System.Windows.Forms.TextBox()
      Me.ToTextBox = New System.Windows.Forms.TextBox()
      Me.OkayButton = New System.Windows.Forms.Button()
      Me.CancelButton = New System.Windows.Forms.Button()
      Me.SuspendLayout()
      '
      'Label1
      '
      Me.Label1.AutoSize = True
      Me.Label1.Location = New System.Drawing.Point(3, 12)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(59, 13)
      Me.Label1.TabIndex = 0
      Me.Label1.Text = "Find What:"
      '
      'Label2
      '
      Me.Label2.AutoSize = True
      Me.Label2.Location = New System.Drawing.Point(3, 35)
      Me.Label2.Name = "Label2"
      Me.Label2.Size = New System.Drawing.Size(75, 13)
      Me.Label2.TabIndex = 1
      Me.Label2.Text = "Replace With:"
      '
      'FromTextBox
      '
      Me.FromTextBox.Location = New System.Drawing.Point(78, 5)
      Me.FromTextBox.Name = "FromTextBox"
      Me.FromTextBox.Size = New System.Drawing.Size(215, 20)
      Me.FromTextBox.TabIndex = 1
      '
      'ToTextBox
      '
      Me.ToTextBox.Location = New System.Drawing.Point(78, 35)
      Me.ToTextBox.Name = "ToTextBox"
      Me.ToTextBox.Size = New System.Drawing.Size(215, 20)
      Me.ToTextBox.TabIndex = 2
      '
      'OkayButton
      '
      Me.OkayButton.Location = New System.Drawing.Point(299, 2)
      Me.OkayButton.Name = "OkayButton"
      Me.OkayButton.Size = New System.Drawing.Size(75, 23)
      Me.OkayButton.TabIndex = 3
      Me.OkayButton.Text = "Okay"
      Me.OkayButton.UseVisualStyleBackColor = True
      '
      'CancelButton
      '
      Me.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
      Me.CancelButton.Location = New System.Drawing.Point(299, 32)
      Me.CancelButton.Name = "CancelButton"
      Me.CancelButton.Size = New System.Drawing.Size(75, 23)
      Me.CancelButton.TabIndex = 4
      Me.CancelButton.Text = "Cancel"
      Me.CancelButton.UseVisualStyleBackColor = True
      '
      'ReplaceTextForm
      '
      Me.AcceptButton = Me.OkayButton
      Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
      Me.ClientSize = New System.Drawing.Size(378, 60)
      Me.Controls.Add(Me.CancelButton)
      Me.Controls.Add(Me.OkayButton)
      Me.Controls.Add(Me.ToTextBox)
      Me.Controls.Add(Me.FromTextBox)
      Me.Controls.Add(Me.Label2)
      Me.Controls.Add(Me.Label1)
      Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.MaximizeBox = False
      Me.Name = "ReplaceTextForm"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
      Me.Text = "Replace Text in Descriptions"
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents FromTextBox As System.Windows.Forms.TextBox
   Friend WithEvents ToTextBox As System.Windows.Forms.TextBox
   Friend WithEvents OkayButton As System.Windows.Forms.Button
   Friend WithEvents CancelButton As System.Windows.Forms.Button
End Class
