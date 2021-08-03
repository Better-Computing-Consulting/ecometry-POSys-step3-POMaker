Public Class ReplaceTextForm
   Public FromText As String = ""
   Public ToText As String = ""
   Private Sub OkayButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OkayButton.Click
      FromText = FromTextBox.Text
      ToText = ToTextBox.Text
      Debug.Print("In Form: from: " & FromText & " to: " & ToText)
      DialogResult = DialogResult.OK
   End Sub
   Private Sub ReplaceTextForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
      FromText = ""
      ToText = ""
   End Sub
   Private Sub CancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelButton.Click
      Me.Close()
   End Sub
End Class