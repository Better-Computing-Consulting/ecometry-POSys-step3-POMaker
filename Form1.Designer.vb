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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.VendorTabs = New System.Windows.Forms.TabControl()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.CreateQuoteForVendorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FinalizePOToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DataGridViewMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.MoveToVendorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.MoveToShortListToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.DataGridViewMenuStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'VendorTabs
        '
        Me.VendorTabs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.VendorTabs.Location = New System.Drawing.Point(0, 27)
        Me.VendorTabs.Name = "VendorTabs"
        Me.VendorTabs.SelectedIndex = 0
        Me.VendorTabs.Size = New System.Drawing.Size(1017, 695)
        Me.VendorTabs.TabIndex = 0
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CreateQuoteForVendorToolStripMenuItem, Me.FinalizePOToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1017, 27)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'CreateQuoteForVendorToolStripMenuItem
        '
        Me.CreateQuoteForVendorToolStripMenuItem.Name = "CreateQuoteForVendorToolStripMenuItem"
        Me.CreateQuoteForVendorToolStripMenuItem.Size = New System.Drawing.Size(176, 23)
        Me.CreateQuoteForVendorToolStripMenuItem.Text = "Create Quote For Vendor"
        '
        'FinalizePOToolStripMenuItem
        '
        Me.FinalizePOToolStripMenuItem.Name = "FinalizePOToolStripMenuItem"
        Me.FinalizePOToolStripMenuItem.Size = New System.Drawing.Size(88, 23)
        Me.FinalizePOToolStripMenuItem.Text = "Finalize PO"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(49, 23)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(116, 24)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'DataGridViewMenuStrip
        '
        Me.DataGridViewMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MoveToVendorToolStripMenuItem, Me.ToolStripSeparator1, Me.MoveToShortListToolStripMenuItem})
        Me.DataGridViewMenuStrip.Name = "DataGridViewMenuStrip"
        Me.DataGridViewMenuStrip.Size = New System.Drawing.Size(193, 58)
        '
        'MoveToVendorToolStripMenuItem
        '
        Me.MoveToVendorToolStripMenuItem.Name = "MoveToVendorToolStripMenuItem"
        Me.MoveToVendorToolStripMenuItem.Size = New System.Drawing.Size(192, 24)
        Me.MoveToVendorToolStripMenuItem.Text = "Move to Vendor..."
        Me.MoveToVendorToolStripMenuItem.ToolTipText = "Move current item to a different vendor"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(189, 6)
        '
        'MoveToShortListToolStripMenuItem
        '
        Me.MoveToShortListToolStripMenuItem.Name = "MoveToShortListToolStripMenuItem"
        Me.MoveToShortListToolStripMenuItem.ShortcutKeyDisplayString = ""
        Me.MoveToShortListToolStripMenuItem.Size = New System.Drawing.Size(192, 24)
        Me.MoveToShortListToolStripMenuItem.Text = "Move to Short List"
        Me.MoveToShortListToolStripMenuItem.ToolTipText = "Move item to Short List Tab"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1017, 722)
        Me.Controls.Add(Me.VendorTabs)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.Text = "Purchase Order Maker"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.DataGridViewMenuStrip.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents VendorTabs As System.Windows.Forms.TabControl
   Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
   Friend WithEvents CreateQuoteForVendorToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents FinalizePOToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents DataGridViewMenuStrip As System.Windows.Forms.ContextMenuStrip
   Friend WithEvents MoveToVendorToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents MoveToShortListToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
   Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
