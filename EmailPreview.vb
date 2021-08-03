Imports System.Net.Mail
Imports System.Net.Mime

Public Class EmailPreviewForm
   Public FullFilePath As String = ""
   Private Sub AttachmentLink_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles AttachmentLink.LinkClicked
      Process.Start(FullFilePath)
   End Sub
   Private Sub SendToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SendToolStripMenuItem.Click
      SendToolStripMenuItem.Enabled = False
      Dim anEmail As New MailMessage
      With anEmail
         .From = User.MAILADDRESS
         .Subject = SubjectTextBox.Text
         If ToTextBox.Text.Contains(",") Then
            Dim emails() As String = ToTextBox.Text.Trim.Split(",")
            For Each s As String In emails
               .To.Add(s.Trim)
            Next
         Else
            .To.Add(ToTextBox.Text.Trim)
         End If
         .Bcc.Add(BccTextBox.Text)
         .Body = HTMLPreviewWebBrowser.DocumentText
         .IsBodyHtml = True
         .Attachments.Add(New Attachment(FullFilePath))
      End With
      Dim SMTPClient As New SmtpClient("172.16.92.10")
      Try
         SMTPClient.Send(anEmail)
         Me.DialogResult = System.Windows.Forms.DialogResult.OK
         MsgBox("Message send succesfully", MsgBoxStyle.Information, "Email Sent")
      Catch ex As Exception
         MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
      End Try
      SendToolStripMenuItem.Enabled = False
   End Sub
End Class