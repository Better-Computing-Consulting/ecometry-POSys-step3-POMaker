<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ActivityHistoryForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ActivityHistoryForm))
        Me.ActivityHistoryTextBox = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'ActivityHistoryTextBox
        '
        Me.ActivityHistoryTextBox.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ActivityHistoryTextBox.Location = New System.Drawing.Point(0, 0)
        Me.ActivityHistoryTextBox.Multiline = True
        Me.ActivityHistoryTextBox.Name = "ActivityHistoryTextBox"
        Me.ActivityHistoryTextBox.Size = New System.Drawing.Size(982, 394)
        Me.ActivityHistoryTextBox.TabIndex = 0
        '
        'ActivityHistoryForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(982, 394)
        Me.Controls.Add(Me.ActivityHistoryTextBox)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "ActivityHistoryForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Activity History"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ActivityHistoryTextBox As System.Windows.Forms.TextBox
End Class
