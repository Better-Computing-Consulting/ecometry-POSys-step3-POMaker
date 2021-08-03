<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ManagerApprovalWindow
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
      Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ManagerApprovalWindow))
      Me.PasswordTextBox = New System.Windows.Forms.MaskedTextBox()
      Me.OKButton = New System.Windows.Forms.Button()
      Me.CancelButton = New System.Windows.Forms.Button()
      Me.InfoLabel = New System.Windows.Forms.TextBox()
      Me.SuspendLayout()
      '
      'PasswordTextBox
      '
      Me.PasswordTextBox.Location = New System.Drawing.Point(12, 199)
      Me.PasswordTextBox.Name = "PasswordTextBox"
      Me.PasswordTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
      Me.PasswordTextBox.Size = New System.Drawing.Size(291, 20)
      Me.PasswordTextBox.TabIndex = 0
      '
      'OKButton
      '
      Me.OKButton.Location = New System.Drawing.Point(310, 196)
      Me.OKButton.Name = "OKButton"
      Me.OKButton.Size = New System.Drawing.Size(75, 23)
      Me.OKButton.TabIndex = 2
      Me.OKButton.Text = "OK"
      Me.OKButton.UseVisualStyleBackColor = True
      '
      'CancelButton
      '
      Me.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
      Me.CancelButton.Location = New System.Drawing.Point(391, 196)
      Me.CancelButton.Name = "CancelButton"
      Me.CancelButton.Size = New System.Drawing.Size(75, 23)
      Me.CancelButton.TabIndex = 3
      Me.CancelButton.Text = "Cancel"
      Me.CancelButton.UseVisualStyleBackColor = True
      '
      'InfoLabel
      '
      Me.InfoLabel.Font = New System.Drawing.Font("Lucida Sans Typewriter", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.InfoLabel.Location = New System.Drawing.Point(12, 0)
      Me.InfoLabel.Multiline = True
      Me.InfoLabel.Name = "InfoLabel"
      Me.InfoLabel.ReadOnly = True
      Me.InfoLabel.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
      Me.InfoLabel.Size = New System.Drawing.Size(576, 193)
      Me.InfoLabel.TabIndex = 4
      '
      'ManagerApprovalWindow
      '
      Me.AcceptButton = Me.OKButton
      Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
      Me.ClientSize = New System.Drawing.Size(590, 231)
      Me.Controls.Add(Me.InfoLabel)
      Me.Controls.Add(Me.CancelButton)
      Me.Controls.Add(Me.OKButton)
      Me.Controls.Add(Me.PasswordTextBox)
      Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.MaximizeBox = False
      Me.MinimizeBox = False
      Me.Name = "ManagerApprovalWindow"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
      Me.Text = "Cost Increase Manager Approval"
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents PasswordTextBox As System.Windows.Forms.MaskedTextBox
   Friend WithEvents OKButton As System.Windows.Forms.Button
   Shadows WithEvents CancelButton As System.Windows.Forms.Button
   Friend WithEvents InfoLabel As System.Windows.Forms.TextBox
End Class
