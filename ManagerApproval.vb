Public Class ManagerApprovalWindow
   Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
      If Me.PasswordTextBox.Text = password Then
         Me.DialogResult = System.Windows.Forms.DialogResult.Yes
      Else
         MsgBox("Bad password", MsgBoxStyle.Exclamation, "Password Error")
         Me.DialogResult = System.Windows.Forms.DialogResult.No
      End If
      Me.Close()
   End Sub
   Private Sub CancelButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelButton.Click
      Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
      Me.Close()
   End Sub

   Private Sub ManagerApprovalWindow_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

   End Sub
End Class