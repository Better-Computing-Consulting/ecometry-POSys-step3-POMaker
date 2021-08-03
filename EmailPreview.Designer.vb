<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EmailPreviewForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EmailPreviewForm))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.SendToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.HTMLPreviewWebBrowser = New System.Windows.Forms.WebBrowser()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.AttachmentLink = New System.Windows.Forms.LinkLabel()
        Me.SubjectTextBox = New System.Windows.Forms.TextBox()
        Me.BccTextBox = New System.Windows.Forms.TextBox()
        Me.ToTextBox = New System.Windows.Forms.TextBox()
        Me.FromTextBox = New System.Windows.Forms.TextBox()
        Me.SubjectLabel = New System.Windows.Forms.Label()
        Me.AttachmentLabel = New System.Windows.Forms.Label()
        Me.BccLabel = New System.Windows.Forms.Label()
        Me.ToLabel = New System.Windows.Forms.Label()
        Me.FromLabel = New System.Windows.Forms.Label()
        Me.MenuStrip1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SendToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(762, 27)
        Me.MenuStrip1.TabIndex = 4
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'SendToolStripMenuItem
        '
        Me.SendToolStripMenuItem.Name = "SendToolStripMenuItem"
        Me.SendToolStripMenuItem.Size = New System.Drawing.Size(51, 23)
        Me.SendToolStripMenuItem.Text = "Send"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.HTMLPreviewWebBrowser, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.GroupBox1, 0, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 27)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 150.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(762, 672)
        Me.TableLayoutPanel1.TabIndex = 7
        '
        'HTMLPreviewWebBrowser
        '
        Me.HTMLPreviewWebBrowser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.HTMLPreviewWebBrowser.Location = New System.Drawing.Point(3, 153)
        Me.HTMLPreviewWebBrowser.MinimumSize = New System.Drawing.Size(20, 20)
        Me.HTMLPreviewWebBrowser.Name = "HTMLPreviewWebBrowser"
        Me.HTMLPreviewWebBrowser.Size = New System.Drawing.Size(756, 516)
        Me.HTMLPreviewWebBrowser.TabIndex = 8
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.AttachmentLink)
        Me.GroupBox1.Controls.Add(Me.SubjectTextBox)
        Me.GroupBox1.Controls.Add(Me.BccTextBox)
        Me.GroupBox1.Controls.Add(Me.ToTextBox)
        Me.GroupBox1.Controls.Add(Me.FromTextBox)
        Me.GroupBox1.Controls.Add(Me.SubjectLabel)
        Me.GroupBox1.Controls.Add(Me.AttachmentLabel)
        Me.GroupBox1.Controls.Add(Me.BccLabel)
        Me.GroupBox1.Controls.Add(Me.ToLabel)
        Me.GroupBox1.Controls.Add(Me.FromLabel)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(756, 144)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        '
        'AttachmentLink
        '
        Me.AttachmentLink.AutoSize = True
        Me.AttachmentLink.Location = New System.Drawing.Point(85, 92)
        Me.AttachmentLink.Name = "AttachmentLink"
        Me.AttachmentLink.Size = New System.Drawing.Size(61, 13)
        Me.AttachmentLink.TabIndex = 16
        Me.AttachmentLink.TabStop = True
        Me.AttachmentLink.Text = "Attachment"
        '
        'SubjectTextBox
        '
        Me.SubjectTextBox.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SubjectTextBox.Location = New System.Drawing.Point(88, 111)
        Me.SubjectTextBox.Name = "SubjectTextBox"
        Me.SubjectTextBox.Size = New System.Drawing.Size(659, 20)
        Me.SubjectTextBox.TabIndex = 15
        '
        'BccTextBox
        '
        Me.BccTextBox.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BccTextBox.Location = New System.Drawing.Point(88, 59)
        Me.BccTextBox.Name = "BccTextBox"
        Me.BccTextBox.ReadOnly = True
        Me.BccTextBox.Size = New System.Drawing.Size(659, 20)
        Me.BccTextBox.TabIndex = 13
        '
        'ToTextBox
        '
        Me.ToTextBox.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ToTextBox.Location = New System.Drawing.Point(88, 36)
        Me.ToTextBox.Name = "ToTextBox"
        Me.ToTextBox.Size = New System.Drawing.Size(659, 20)
        Me.ToTextBox.TabIndex = 2
        '
        'FromTextBox
        '
        Me.FromTextBox.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FromTextBox.Location = New System.Drawing.Point(88, 13)
        Me.FromTextBox.Name = "FromTextBox"
        Me.FromTextBox.ReadOnly = True
        Me.FromTextBox.Size = New System.Drawing.Size(659, 20)
        Me.FromTextBox.TabIndex = 11
        '
        'SubjectLabel
        '
        Me.SubjectLabel.AutoSize = True
        Me.SubjectLabel.Location = New System.Drawing.Point(12, 118)
        Me.SubjectLabel.Name = "SubjectLabel"
        Me.SubjectLabel.Size = New System.Drawing.Size(46, 13)
        Me.SubjectLabel.TabIndex = 10
        Me.SubjectLabel.Text = "Subject:"
        '
        'AttachmentLabel
        '
        Me.AttachmentLabel.AutoSize = True
        Me.AttachmentLabel.Location = New System.Drawing.Point(12, 92)
        Me.AttachmentLabel.Name = "AttachmentLabel"
        Me.AttachmentLabel.Size = New System.Drawing.Size(64, 13)
        Me.AttachmentLabel.TabIndex = 9
        Me.AttachmentLabel.Text = "Attachment:"
        '
        'BccLabel
        '
        Me.BccLabel.AutoSize = True
        Me.BccLabel.Location = New System.Drawing.Point(12, 66)
        Me.BccLabel.Name = "BccLabel"
        Me.BccLabel.Size = New System.Drawing.Size(29, 13)
        Me.BccLabel.TabIndex = 8
        Me.BccLabel.Text = "Bcc:"
        '
        'ToLabel
        '
        Me.ToLabel.AutoSize = True
        Me.ToLabel.Location = New System.Drawing.Point(12, 39)
        Me.ToLabel.Name = "ToLabel"
        Me.ToLabel.Size = New System.Drawing.Size(23, 13)
        Me.ToLabel.TabIndex = 7
        Me.ToLabel.Text = "To:"
        '
        'FromLabel
        '
        Me.FromLabel.AutoSize = True
        Me.FromLabel.Location = New System.Drawing.Point(12, 16)
        Me.FromLabel.Name = "FromLabel"
        Me.FromLabel.Size = New System.Drawing.Size(33, 13)
        Me.FromLabel.TabIndex = 6
        Me.FromLabel.Text = "From:"
        '
        'EmailPreviewForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(762, 699)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "EmailPreviewForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Email To Vendor Preview"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
   Friend WithEvents SendToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
   Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
   Friend WithEvents SubjectTextBox As System.Windows.Forms.TextBox
   Friend WithEvents BccTextBox As System.Windows.Forms.TextBox
   Friend WithEvents ToTextBox As System.Windows.Forms.TextBox
   Friend WithEvents FromTextBox As System.Windows.Forms.TextBox
   Friend WithEvents SubjectLabel As System.Windows.Forms.Label
   Friend WithEvents AttachmentLabel As System.Windows.Forms.Label
   Friend WithEvents BccLabel As System.Windows.Forms.Label
   Friend WithEvents ToLabel As System.Windows.Forms.Label
   Friend WithEvents FromLabel As System.Windows.Forms.Label
   Friend WithEvents AttachmentLink As System.Windows.Forms.LinkLabel
   Friend WithEvents HTMLPreviewWebBrowser As System.Windows.Forms.WebBrowser
End Class
