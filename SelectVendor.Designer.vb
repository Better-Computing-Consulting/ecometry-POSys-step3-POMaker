<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SelectVendorWindow
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SelectVendorWindow))
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.btn = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.Vendor = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Rank = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Cost = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CancelButton = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.btn, Me.Vendor, Me.Rank, Me.Cost})
        Me.DataGridView1.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.Size = New System.Drawing.Size(341, 180)
        Me.DataGridView1.TabIndex = 0
        '
        'btn
        '
        Me.btn.HeaderText = ""
        Me.btn.Name = "btn"
        Me.btn.ReadOnly = True
        Me.btn.Text = "Select"
        Me.btn.UseColumnTextForButtonValue = True
        '
        'Vendor
        '
        DataGridViewCellStyle1.Format = "C2"
        DataGridViewCellStyle1.NullValue = Nothing
        Me.Vendor.DefaultCellStyle = DataGridViewCellStyle1
        Me.Vendor.HeaderText = "Vendor"
        Me.Vendor.Name = "Vendor"
        Me.Vendor.ReadOnly = True
        Me.Vendor.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        '
        'Rank
        '
        Me.Rank.HeaderText = "Rank"
        Me.Rank.Name = "Rank"
        Me.Rank.ReadOnly = True
        Me.Rank.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Rank.Width = 50
        '
        'Cost
        '
        DataGridViewCellStyle2.Format = "C2"
        DataGridViewCellStyle2.NullValue = Nothing
        Me.Cost.DefaultCellStyle = DataGridViewCellStyle2
        Me.Cost.HeaderText = "Cost"
        Me.Cost.Name = "Cost"
        Me.Cost.ReadOnly = True
        Me.Cost.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Cost.Width = 70
        '
        'CancelButton
        '
        Me.CancelButton.Location = New System.Drawing.Point(0, 186)
        Me.CancelButton.Name = "CancelButton"
        Me.CancelButton.Size = New System.Drawing.Size(75, 23)
        Me.CancelButton.TabIndex = 2
        Me.CancelButton.Text = "Cancel"
        Me.CancelButton.UseVisualStyleBackColor = True
        '
        'SelectVendorWindow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(341, 212)
        Me.Controls.Add(Me.CancelButton)
        Me.Controls.Add(Me.DataGridView1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimizeBox = False
        Me.Name = "SelectVendorWindow"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Select Vendor"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
   Friend WithEvents CancelButton As System.Windows.Forms.Button
   Friend WithEvents btn As System.Windows.Forms.DataGridViewButtonColumn
   Friend WithEvents Vendor As System.Windows.Forms.DataGridViewTextBoxColumn
   Friend WithEvents Rank As System.Windows.Forms.DataGridViewTextBoxColumn
   Friend WithEvents Cost As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
