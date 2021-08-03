Public Class SelectVendorWindow
   Public NewVendorSelected As String = ""
   Private Sub SelectVendor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      NewVendorSelected = ""
   End Sub
   Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
      If e.RowIndex < 0 OrElse Not e.ColumnIndex = DataGridView1.Columns(0).Index Then Return
      Dim v As String = DataGridView1(1, e.RowIndex).Value
      'SelectedVendor = v
      NewVendorSelected = v
      Me.DialogResult = System.Windows.Forms.DialogResult.OK
      Me.Close()
   End Sub
   Private Sub CancelButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelButton.Click
      Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
      Me.Close()
   End Sub
End Class