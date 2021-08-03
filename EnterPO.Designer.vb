<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EnterPOForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EnterPOForm))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.EnterPOToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.CancelButton = New System.Windows.Forms.Button()
        Me.POCommentsTextBox = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.FinalReportWebBrowser = New System.Windows.Forms.WebBrowser()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EnterPOToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(943, 27)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'EnterPOToolStripMenuItem
        '
        Me.EnterPOToolStripMenuItem.Name = "EnterPOToolStripMenuItem"
        Me.EnterPOToolStripMenuItem.Size = New System.Drawing.Size(76, 23)
        Me.EnterPOToolStripMenuItem.Text = "Enter PO"
        '
        'ListBox1
        '
        Me.ListBox1.AllowDrop = True
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(365, 50)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(424, 95)
        Me.ListBox1.TabIndex = 2
        '
        'CancelButton
        '
        Me.CancelButton.Location = New System.Drawing.Point(795, 50)
        Me.CancelButton.Name = "CancelButton"
        Me.CancelButton.Size = New System.Drawing.Size(147, 95)
        Me.CancelButton.TabIndex = 3
        Me.CancelButton.Text = "Cancel"
        Me.CancelButton.UseVisualStyleBackColor = True
        '
        'POCommentsTextBox
        '
        Me.POCommentsTextBox.AllowDrop = True
        Me.POCommentsTextBox.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.POCommentsTextBox.Location = New System.Drawing.Point(0, 50)
        Me.POCommentsTextBox.MaxLength = 290
        Me.POCommentsTextBox.Multiline = True
        Me.POCommentsTextBox.Name = "POCommentsTextBox"
        Me.POCommentsTextBox.Size = New System.Drawing.Size(359, 99)
        Me.POCommentsTextBox.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(-3, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(233, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "PO Comments for Receiving ticket (6 lines max):"
        '
        'FinalReportWebBrowser
        '
        Me.FinalReportWebBrowser.AllowNavigation = False
        Me.FinalReportWebBrowser.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.FinalReportWebBrowser.Location = New System.Drawing.Point(0, 151)
        Me.FinalReportWebBrowser.MinimumSize = New System.Drawing.Size(20, 20)
        Me.FinalReportWebBrowser.Name = "FinalReportWebBrowser"
        Me.FinalReportWebBrowser.Size = New System.Drawing.Size(943, 613)
        Me.FinalReportWebBrowser.TabIndex = 6
        '
        'EnterPOForm
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(943, 764)
        Me.Controls.Add(Me.FinalReportWebBrowser)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.POCommentsTextBox)
        Me.Controls.Add(Me.CancelButton)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.Name = "EnterPOForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Finalize Purchase Order"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
   Friend WithEvents EnterPOToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
   Friend WithEvents CancelButton As System.Windows.Forms.Button
   Friend WithEvents POCommentsTextBox As System.Windows.Forms.TextBox
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents FinalReportWebBrowser As System.Windows.Forms.WebBrowser
End Class
