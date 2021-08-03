Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Text.RegularExpressions
Public Class EnterPOForm
   Public aVendor As ItemVendor
   Public POItems As New List(Of RecommendedBuyItem)
   Public DestinationFolder As String = ""
   Public Vendor As String = ""
   Dim PONumber As String = ""
   Dim TotalCost As Decimal = 0
   Public POCommentLines As New List(Of String)
   Public Sub New(ByVal vi As VendorItems)
      InitializeComponent()
      POItems = vi.WorkingItems
      PONumber = vi.PONUMBER
      Vendor = vi.VendorNumber
      TotalCost = vi.WorkingTotalPurchase
   End Sub
   Private Sub EnterPOToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnterPOToolStripMenuItem.Click
      Me.ValidateChildren()
      DialogResult = DialogResult.OK
      Dim s As String = ""
      For Each l As String In POCommentLines
         s &= l & Environment.NewLine
      Next
      Me.Close()
   End Sub
   Private Sub EnterPOForm_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
      ExecuteDragDrop(sender, e)
   End Sub
   Private Sub EnterPOForm_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter
      ExecuteDragEnter(sender, e)
   End Sub
   Private Sub EnterPOForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      Dim sFormat As String = "{0,-20}{1,-55}{2,-10}{3,-10}{4}" & vbCrLf
      EnterPOToolStripMenuItem.Enabled = False
      ListBox1.Items.Add("Drag and Drop files or emails in the window to activate the Enter PO Button.")
      ListBox1.Items.Add("")
   End Sub
   Private Sub ListBox1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragDrop
      ExecuteDragDrop(sender, e)
   End Sub
   Private Sub ExecuteDragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
      If e.Data.GetDataPresent(DataFormats.FileDrop) Then
         Dim MyFiles() As String
         Dim i As Integer
         MyFiles = e.Data.GetData(DataFormats.FileDrop)
         For i = 0 To MyFiles.Length - 1
            ListBox1.Items.Add(MyFiles(i))
            File.Copy(MyFiles(i), DestinationFolder & New FileInfo(MyFiles(i)).Name, True)
            EnterPOToolStripMenuItem.Enabled = True
         Next
      ElseIf e.Data.GetDataPresent("FileGroupDescriptor") Then
         Dim anOutlookEmail As Outlook.MailItem = Nothing
         Dim OutlookApp As New Outlook.Application
         For Each MailItem As Outlook.MailItem In OutlookApp.ActiveExplorer.Selection
            Dim tmpFilename As String = Regex.Replace(MailItem.Subject, "[<>:""/\\|?*]", "_") & ".msg"
            Try
               MailItem.SaveAs(DestinationFolder & tmpFilename)
               If File.Exists(DestinationFolder & tmpFilename) Then
                  EnterPOToolStripMenuItem.Enabled = True
                  ListBox1.Items.Add(MailItem.Subject)
               Else
                  MsgBox("Email did not save.  Maybe you did not grant access in Outlook.  Please, try again.", MsgBoxStyle.Exclamation, "Email Copy Error")
               End If
            Catch ex As Exception
               MsgBox("Email did not save.  Maybe you did not grant access in Outlook.  Please, try again.", MsgBoxStyle.Exclamation, "Email Copy Error")
            End Try
         Next
      End If
   End Sub
   Private Sub ListBox1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragEnter
      ExecuteDragEnter(sender, e)
   End Sub
   Private Sub ExecuteDragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
      If e.Data.GetDataPresent(DataFormats.FileDrop) Then
         e.Effect = DragDropEffects.All
      ElseIf e.Data.GetDataPresent("FileGroupDescriptor") Then
         e.Effect = DragDropEffects.All
      End If
   End Sub
   Private Sub CancelButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelButton.Click
      DialogResult = DialogResult.Cancel
      Me.Close()
   End Sub
   Private Sub POCommentsTextBox_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles POCommentsTextBox.DragDrop
      ExecuteDragDrop(sender, e)
   End Sub
   Private Sub POCommentsTextBox_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles POCommentsTextBox.DragEnter
      ExecuteDragEnter(sender, e)
   End Sub
   Private Sub POCommentsTextBox_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles POCommentsTextBox.Validating
      Debug.Print("validating")
      Dim alllines As New List(Of String)
      For Each s As String In POCommentsTextBox.Lines
         If Not String.IsNullOrEmpty(s) Then alllines.Add(s)
      Next
      If alllines.Count > 6 Then
         MsgBox("Six line limit violated.", MsgBoxStyle.Exclamation, "PO Comment limit alert")
         e.Cancel = True
      End If
      Dim tmplines As New List(Of String)
      For Each line As String In alllines
         If line.Length > 50 Then tmplines.AddRange(BrakeLine(line)) Else tmplines.Add(line)
      Next
      If tmplines.Count > 6 Then
         MsgBox("Six line limit violated.", MsgBoxStyle.Exclamation, "PO Comment limit alert")
         e.Cancel = True
      End If
      POCommentLines.Clear()
      For Each s As String In tmplines
         If s.Length > 50 Then
            MsgBox("Line """ & s & """ is over the 50 character limit.", MsgBoxStyle.Exclamation, "PO Comment limit alert")
            e.Cancel = True
         End If
         Debug.Print(s)
         POCommentLines.Add(s.Trim)
      Next
   End Sub
   Private Function BrakeLine(ByVal line As String) As List(Of String)
      Dim tmpResult As New List(Of String)
      If line.Length <= 50 Then
         tmpResult.Add(line)
         Return tmpResult
      Else
         Dim s1 As String = line.Substring(0, 50)
         Dim i As Integer = s1.TrimEnd.LastIndexOf(" ")
         If i = -1 Then i = 50 Else i += 1
         tmpResult.Add(s1.Substring(0, i))
         Dim left1 As String = line.Substring(i)
         If left1.Length > 50 Then tmpResult.AddRange(BrakeLine(left1)) Else tmpResult.Add(left1)
         Return tmpResult
      End If
   End Function
   Private Sub POItemsTextBox_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
      ExecuteDragDrop(sender, e)
   End Sub
   Private Sub POItemsTextBox_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
      ExecuteDragEnter(sender, e)
   End Sub
End Class