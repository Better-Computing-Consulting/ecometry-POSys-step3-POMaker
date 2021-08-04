#Region "Imports"
Imports System.IO
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports Microsoft.Office.Interop
Imports System.Text.RegularExpressions
Imports System.Collections
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.Serialization
#End Region
Public Class Form1
#Region "Class Private Variables"
   ' Dim RcmdItems As New List(Of RecommendedBuyItem)
   'Public OneFinalPOItems As New List(Of RecommendedBuyItem)
   Dim AppUser As String = ""
   'Dim CurrentUser As CurrentUser
   'Dim UserVendors As New List(Of ItemVendor)
   'Dim UserVendorItems As New List(Of VendorItems)
   'Dim TabVendors As New List(Of String)
   Dim ApplicationVersion As String = ""
   'Public SelectedVendor As String
   'Public VERCONN As String = "Data Source=ecom-db2;Initial Catalog=ECOMVER;UID=mgr;PWD=doall"
#End Region
   Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      Debug.Print(vbCrLf & Now.ToString)
      AppUser = My.Application.GetEnvironmentVariable("USERNAME")
      If AppUser.ToUpper = "FEDERICO" Then AppUser = InputBox("User", "Enter Runas user", "", , )
      If AppUser.ToUpper = "TSCHMIDT" Then AppUser = InputBox("User", "Enter Runas user", "", , )
      If Not {"BRIAN", "SCOTTM", "PAUL", "CATHY", "EDLYN", "TRACEY"}.Contains(AppUser.ToUpper) Then
         MsgBox("You are not authorized to access this application.", MsgBoxStyle.Exclamation, "Access Error")
         Me.Close()
      End If
      User = New CurrentUser(AppUser)
      UserVendorItemsList = New List(Of VendorItems)
      For Each Vendor As String In User.Vendors
         'Debug.Print(Vendor)
         Dim tmpVI As New VendorItems(Vendor)
         If tmpVI.WorkingItems.Count > 0 Then
            UserVendorItemsList.Add(tmpVI)
            VendorTabs.TabPages.Add(tmpVI.TabPage)
            AddHandler tmpVI.VendorAdded, AddressOf VendorAdded
            AddHandler tmpVI.VendorEmpty, AddressOf VendorEmpty
         End If
      Next
   End Sub
   Private Sub Form1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
      If VendorTabs.TabPages.Count > 0 Then
         If VendorTabs.SelectedTab.Name = User.ID Then
            FinalizePOToolStripMenuItem.Enabled = False
         End If
      End If
   End Sub
   Private Sub VendorAdded(ByVal vi As VendorItems)
      UserVendorItemsList.Add(vi)
      If vi.VendorNumber = User.ID Then
         VendorTabs.TabPages.Add(vi.TabPage)
         AddHandler vi.VendorAdded, AddressOf VendorAdded
         AddHandler vi.VendorEmpty, AddressOf VendorEmpty
         Exit Sub
      End If
      Dim tmplist As New List(Of String)
      For Each v As VendorItems In UserVendorItemsList
         If v.VendorNumber <> User.ID Then
            tmplist.Add(v.VendorNumber)
         End If
      Next
      tmplist.Sort()
      Dim i As Integer = 0
      For Each s As String In tmplist
         If s = vi.VendorNumber Then
            VendorTabs.TabPages.Insert(i, vi.TabPage)
         End If
         i += 1
      Next
      AddHandler vi.VendorAdded, AddressOf VendorAdded
      AddHandler vi.VendorEmpty, AddressOf VendorEmpty
   End Sub
   Private Sub VendorEmpty(ByVal vi As VendorItems)
      Dim id As String = vi.VendorNumber
      For Each t As TabPage In VendorTabs.TabPages
         If t.Name = vi.VendorNumber Then
            VendorTabs.TabPages.Remove(t)
         End If
      Next
      UserVendorItemsList.RemoveAll(VendorItems.FindPredicateByVendorId(vi.VendorNumber))
   End Sub
   Private Sub FinalizePOToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FinalizePOToolStripMenuItem.Click
      Dim CurrentVendorItems As VendorItems = UserVendorItemsList.Find(VendorItems.FindPredicateByVendorId(VendorTabs.SelectedTab.Name))
      If CurrentVendorItems.VendorNumber = User.ID Then
         MsgBox("You cannot create a PO for the short list.", MsgBoxStyle.Exclamation, "In Short List")
         Exit Sub
      End If
      If CurrentVendorItems.ItemsWithoutRECDATE.Length > 1 Then
         MsgBox("These items do not have a Due Date set." & vbCrLf & CurrentVendorItems.ItemsWithoutRECDATE, MsgBoxStyle.Exclamation, "Items withour Due Date")
         Exit Sub
      End If
      Dim tmpListOfHigherCostItems As New List(Of String)
      Dim CostierItems As New List(Of RecommendedBuyItem)
      For Each i As RecommendedBuyItem In CurrentVendorItems.WorkingItems
         Dim increase As Decimal = 0
         If i.OriginalVendorCost > 0 Then
            increase = ((i.VendorCost - i.OriginalVendorCost) / i.OriginalVendorCost) * 100
         Else
            increase = 99999
         End If
         If increase >= 2 Then
            Dim mrgnPrior As Decimal = 0
            tmpListOfHigherCostItems.Add(String.Format("{0,-21}{1,-9}{2,-9}{3,-9}{4,-15}{5}", i.ITEMNO, i.OriginalVendorCost, i.VendorCost, increase.ToString("f2") & "%", i.OriginalMargin.ToString("f2") & "%", i.Margin.ToString("f2") & "%"))
            CostierItems.Add(i)
         End If
      Next
      If tmpListOfHigherCostItems.Count > 0 Then
         Dim sRprt As String = String.Format("{0,-21}{1,-9}{2,-9}{3,-9}{4,-15}{5}" & Environment.NewLine, "ITEM", "Old Cost", "New Cost", "Increase", "Margin Before", "Margin After")
         For Each i As String In tmpListOfHigherCostItems
            sRprt &= i & Environment.NewLine
         Next
         Dim ApprovalWindow As New ManagerApprovalWindow
         ApprovalWindow.InfoLabel.Text = "These items have cost increases of over 2%:" & Environment.NewLine & Environment.NewLine & sRprt & Environment.NewLine & "Please enter Manager approval password:"
         If DialogResult.Yes = ApprovalWindow.ShowDialog Then
            For Each i As RecommendedBuyItem In CostierItems
               LogThis("ITMCOSTAPPRV", i.ITEMNO & " manager approved item cost increase from " & i.OriginalVendorCost & " to " & i.VendorCost, i.ITEMID)
            Next
         Else
            Exit Sub
         End If
      End If
      If CurrentVendorItems.TotalCostHigher() Then
         Exit Sub
      End If
      'after checking price increase manager approval
      Dim FinalDocumentPath As String = CurrentVendorItems.FinalDocumentPath
      If FinalDocumentPath = "" Then
         MsgBox("Error creating final document.", MsgBoxStyle.Exclamation, "Error creating final document")
         Exit Sub
      End If
      Dim FinalizeWindow As New EnterPOForm(CurrentVendorItems)
      Dim result As DialogResult
      With FinalizeWindow
         .DestinationFolder = FinalDocumentPath
         .POCommentsTextBox.Text = CurrentVendorItems.Vendor.POComments
         .FinalReportWebBrowser.DocumentText = CurrentVendorItems.FinalHTMLReport
         result = .ShowDialog()
         CurrentVendorItems.FinalPOComments = .POCommentLines
      End With
      If Not result = DialogResult.OK Then Return
      If CurrentVendorItems.FinalizePurchaseOrderSuccess() Then
         Using sr As New StreamWriter(FinalDocumentPath & "FINALPURCHASEORDER.HTML")
            sr.WriteLine(CurrentVendorItems.FinalHTMLReport)
            sr.Flush()
         End Using
         MsgBox("PO " & CurrentVendorItems.PONUMBER & " entered sucessfully")
      Else
         MsgBox("something went wrong entering po")
      End If
   End Sub
   Private Sub CreateQuoteForVendorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreateQuoteForVendorToolStripMenuItem.Click
      Dim CurrentVendorItems As VendorItems = UserVendorItemsList.Find(VendorItems.FindPredicateByVendorId(VendorTabs.SelectedTab.Name))
      Dim ProposalDocument As String = CurrentVendorItems.ProposalDocumentPath
      If ProposalDocument = "" Then
         MsgBox("Error creating proposal document.", MsgBoxStyle.Exclamation, "Error creating proposal document")
         Exit Sub
      End If
      Dim FileName As String = New FileInfo(ProposalDocument).Name
      Dim anEmailPreviewWindow As New EmailPreviewForm
      Dim result As DialogResult
      With anEmailPreviewWindow
         .FullFilePath = ProposalDocument
         .FromTextBox.Text = User.MAILADDRESS.ToString
         .ToTextBox.Text = CurrentVendorItems.Vendor.Email
         .BccTextBox.Text = User.Email
         .AttachmentLink.Text = FileName
         .SubjectTextBox.Text = "PRICE POINT PO " & FileName.Substring(2, 9)
         .HTMLPreviewWebBrowser.DocumentText = CurrentVendorItems.EmailBodyInHTML
         result = .ShowDialog
      End With
      If result = DialogResult.OK Then CurrentVendorItems.UpdateLastEmailDate()
   End Sub
   Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
      With My.Application.Info.Version
         ApplicationVersion = .Major & "." & .MajorRevision & "." & .Minor & "." & .MinorRevision
      End With
      With My.Application.Info
         MsgBox(.AssemblyName & vbCrLf & .CompanyName & vbCrLf & ApplicationVersion, MsgBoxStyle.Information, "About PO Maker")
      End With
   End Sub
   Private Sub VendorTabs_Selected(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlEventArgs) Handles VendorTabs.Selected
      If e.TabPage.Name = User.ID Then
         FinalizePOToolStripMenuItem.Enabled = False
      Else
         FinalizePOToolStripMenuItem.Enabled = True
      End If
   End Sub
End Class
Public Class VendorItems
   Implements IComparable(Of VendorItems)
#Region "Class Private Variables"
   Private oTab As New TabPage
   Private WithEvents oDataViewGrid As New DataGridView
   Private oAddressLabel As New Label
   Private oCommentsLabel1 As New Label
   Private oVendorCommentslabel As New TextBox
   Private oCommentsLabel2 As New Label
   Private oPOCommentslabel As New TextBox
   Private oRequestLabel As New Label
   Private oRequestText As New TextBox
   Private oDueDatePicker As New DateTimePicker
   Private oSetPODateButton As New Button
   Private WithEvents oEditDescriptionButton As New Button
   Private oCountlabel As New ToolStripStatusLabel
   Private oTotallabel As New ToolStripStatusLabel
   Private oPOlabel As New ToolStripStatusLabel
   Private oInfolabel As New ToolStripStatusLabel
   Private oEmailLabel As New ToolStripStatusLabel
   Private oStatusStrip As New StatusStrip
   Private oVendorNumber As String = ""
   Private oVendor As ItemVendor
   Private oWorkingItems As New List(Of RecommendedBuyItem)
   Private oInitialItems As New List(Of RecommendedBuyItem)
   Private oFinalPOComments As New List(Of String)
   Private oPONumber As String = ""
   Private oPODate As String = ""
   Private oLastEmailDate As String = ""
   Private oPopUpMenu As New ContextMenuStrip
   Private WithEvents oMoveToVendor As New ToolStripMenuItem
   Private WithEvents oMoveToShortList As New ToolStripMenuItem
   Private WithEvents oItemEventHistory As New ToolStripMenuItem
    Private VERCONN As String = "Data Source=ecom-db2;Initial Catalog=ECOMVER;UID=xxx;PWD=xxxx"
#End Region
    Public Event VendorAdded(ByVal vi As VendorItems)
   Public Event VendorEmpty(ByVal vi As VendorItems)
   Sub New(ByVal VendorNumber As String)
      oVendorNumber = VendorNumber.ToUpper
      oWorkingItems = GetVendorItems()
      If oWorkingItems.Count > 0 Then
         oVendor = New ItemVendor(VendorNumber)
         SetTempPONumberDate()
         SetLastEmailDate()
         SetDataViewGrid()
         SetStatusStrip()
         SetNewTab()
         oInitialItems = DeepCopy(oWorkingItems)
      End If
   End Sub
#Region "Class Properties"
   ReadOnly Property VendorNumber() As String
      Get
         Return oVendorNumber
      End Get
   End Property
   ReadOnly Property Vendor() As ItemVendor
      Get
         Return oVendor
      End Get
   End Property
   ReadOnly Property WorkingItems() As List(Of RecommendedBuyItem)
      Get
         Return oWorkingItems
      End Get
   End Property
   ReadOnly Property WorkingTotalPurchase() As Decimal
      Get
         Dim i As Decimal = 0
         For Each anitem As RecommendedBuyItem In oWorkingItems
            i += anitem.VendorCost * anitem.TotalQtyToBuy
         Next
         Return i
      End Get
   End Property
   ReadOnly Property PONUMBER() As String
      Get
         Return oPONumber.Trim
      End Get
   End Property
   ReadOnly Property TabPage() As TabPage
      Get
         Return oTab
      End Get
   End Property
   ReadOnly Property ProposalDocumentPath As String
      Get
         Dim TmpPO As String = ValidatePONumber()
         If TmpPO = "" Then Return ""
         Dim tmpFolderName As String = RootDocFolder & oVendorNumber & "\" & TmpPO & "\"
         Dim tmpFileName = "PO" & TmpPO & "." & Now.ToString("yyyyMMddhhmm") & ".xls"
         If File.Exists("C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE") Or File.Exists("C:\Program Files\Microsoft Office\Office14\EXCEL.EXE") Then tmpFileName &= "x"
         'If {"PAUL", "KATIE"}.Contains(User.ID) Then tmpFileName &= "x"
         If Not Directory.Exists(tmpFolderName) Then
            Try
               Directory.CreateDirectory(tmpFolderName)
            Catch ex As Exception
               MsgBox(ex.Message, MsgBoxStyle.Exclamation, "PO Directory Creation Error")
               tmpFolderName = My.Computer.FileSystem.SpecialDirectories.Temp & "\"
            End Try
         End If
         Dim tmpResult As String = tmpFolderName & tmpFileName
         Dim xlsApp As New Excel.Application
         Dim xlsWrkbk As Excel.Workbook = xlsApp.Workbooks.Add
         Dim xlsWSheet As Excel.Worksheet = xlsWrkbk.Worksheets(1)
         Dim rowcntr As Integer = 1
         With xlsWSheet
            .Range("A1").Value = "Product Number"
            .Range("B1").Value = "Vendor Model Number"
            .Range("C1").Value = "Description"
            .Range("D1").Value = "Qty"
            .Range("D1").ColumnWidth = 8.57
            .Range("E1").Value = "Confirm"
            .Range("E1").ColumnWidth = 8.14
            .Range("F1").Value = "Cost"
            .Range("F1").ColumnWidth = 10
            .Range("F1").EntireColumn.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            .Range("G1").Value = "Confirm"
            .Range("G1").ColumnWidth = 8.14
            .Range("H1").Value = "ETA's"
            .Range("H1").ColumnWidth = 13.43
            .Range("I1").Value = "Comments"
            .Range("I1").ColumnWidth = 22.29
            .Range("A1:I1").Font.Bold = True
            .Range("A1:I1").HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
            For Each oneitem As RecommendedBuyItem In oWorkingItems
               If oneitem.TotalQtyToBuy > 0 Then
                  rowcntr += 1
                  .Range("A" & rowcntr).Value = oneitem.ITEMNO
                  .Range("B" & rowcntr).Value = oneitem.VendorItemNo
                  .Range("C" & rowcntr).Value = oneitem.Description
                  .Range("D" & rowcntr).Value = oneitem.TotalQtyToBuy
                  .Range("F" & rowcntr).Value = oneitem.VendorCost
               End If
            Next
            .Range("A1:I" & rowcntr).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = BorderStyle.FixedSingle
            .Range("A1:I" & rowcntr).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = BorderStyle.FixedSingle
            .Range("A1:I" & rowcntr).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = BorderStyle.FixedSingle
            .Range("A1:I" & rowcntr).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = BorderStyle.FixedSingle
            .Range("A1:I" & rowcntr).Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = BorderStyle.FixedSingle
            .Range("A1:I" & rowcntr).Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = BorderStyle.FixedSingle
            .Range("A1").EntireColumn.AutoFit()
            .Range("B1").EntireColumn.AutoFit()
            .Range("C1").EntireColumn.AutoFit()
         End With
         Try
            xlsWrkbk.SaveAs(tmpResult)
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error Saving Vendor Quote Spreadsheet")
         End Try
         xlsApp.Quit()
         Return tmpResult
      End Get
   End Property
   ReadOnly Property FinalDocumentPath As String
      Get
         Dim TmpPO As String = ValidatePONumber()
         If TmpPO = "" Then Return ""
         Dim tmpFolderName As String = RootDocFolder & oVendorNumber & "\" & TmpPO & "\"
         If Not Directory.Exists(tmpFolderName) Then
            Try
               Directory.CreateDirectory(tmpFolderName)
            Catch ex As Exception
               MsgBox(ex.Message, MsgBoxStyle.Exclamation, "PO Directory Creation Error")
               tmpFolderName = My.Computer.FileSystem.SpecialDirectories.Temp & "\"
            End Try
         End If
         Return tmpFolderName
      End Get
   End Property
   ReadOnly Property ItemsWithoutRECDATE() As String
      Get
         Dim result As String = ""
         For Each i As RecommendedBuyItem In oWorkingItems
            If String.IsNullOrWhiteSpace(i.RECDATE) Then
               result &= i.ITEMNO & vbCrLf
            End If
         Next
         Return result
      End Get
   End Property
   Public Property FinalPOComments() As List(Of String)
      Get
         Return oFinalPOComments
      End Get
      Set(ByVal value As List(Of String))
         oFinalPOComments = value
      End Set
   End Property
   ReadOnly Property FinalHTMLReport As String
      Get
         Dim htmlcontent As New System.Text.StringBuilder
         With htmlcontent
            .Append("<!DOCTYPE html><html><head><style>body{Font-family:courier}td{font-size:14px;}</style><title>PO " & PONUMBER & "</title></head><body>")
            .Append("<h3><p>PO Number: " & PONUMBER & "&nbsp &nbsp &nbsp PO Date: " & CDate(oPODate).ToShortDateString & "&nbsp &nbsp &nbsp Finalized Date: " & Now.ToShortDateString & "</p></h3>")
            .Append("<p><h4>" & Vendor.ContactNameAndAddressHTML & "<h4>")
            If oFinalPOComments.Count > 0 Then
               .Append("<p><h4>Final PO Comments for Receving Ticket:<br><br>")
               For Each s As String In FinalPOComments
                  .Append(s & "<br>")
               Next
               .Append("<br></h4>")
            End If
            .Append("</td></tr></table><br>")
            .Append("<table border=""1""><col align=""left"" /><col align=""left"" /><col align=""center"" /><col align=""center"" /><col align=""right"" /><col align=""right"" />")
            .Append("<tr><th>Item Number</th><th>Description</th><th>Due Date</th><th>Quantity</th><th>Cost</th><th>Total</th></tr>")
            For Each r As RecommendedBuyItem In oWorkingItems
               .Append("<tr><td>" & r.ITEMNO & "</td><td>" & r.Description & "</td><td>" & r.RECDATE & "</td><td>" & r.TotalQtyToBuy & "</td><td>" & r.VendorCost.ToString("f2") & "</td><td>" & r.TotalVendorCost.ToString("f2") & "</td></tr>")
            Next
            .Append("<tfoot><tr><td colspan=""5""></td><th>" & WorkingTotalPurchase.ToString("f2") & "</th></tr></tfoot></table></body></html>")
            Return .ToString
         End With
      End Get
   End Property
   ReadOnly Property EmailBodyInHTML As String
      Get
         Dim tmpResult As New System.Text.StringBuilder
         tmpResult.Append("<p>Hello")
         If Vendor.ContactFullName.Length > 3 Then tmpResult.Append(" " & Vendor.ContactFullName)
            tmpResult.Append(",</p><p>Attached is a new purchase order, ecommerce company PO " & oPONumber & "</p>")
            If Not String.IsNullOrEmpty(oRequestText.Text.Trim) Then tmpResult.Append("<p>" & oRequestText.Text.Trim & "</p>")
         If oVendor.POCommentsHTML.Length > 3 Then tmpResult.Append("<p>The payment terms are as follow: <br><br>" & oVendor.POCommentsHTML & "<br>")
         tmpResult.Append("</p><p>Please open the attached spreadsheet to confirm pricing, availability and delivery date as soon as possible. " & _
                          "<u>Reply to this email with confirmation of receipt.</u>  <u>If any changes need to be made to the order, please note them in the attached spreadsheet.</u> " & _
                          "If you do not provide confirmation via email and ship the order it will be refused.</p>" & _
                          "<p>Any pricing, availability and/or delivery discrepancies need to be agreed upon before the order is finalized. " & _
                          "<u>Please note we will only pay the agreed upon costs and accept the agreed upon quantities.</u> Any excess quantities, unordered items or prices those not " & _
                          "matching the final order confirmation will be cancelled and returned at your expense. This is a <u>legally binding agreement</u> to purchase/provide the " & _
                          "products in the attached purchase order. Any variance from the order confirmation must be agreed to via email or writing. Any verbal agreements may or may " & _
                          "not be honored at our sole discretion.</p>")
         Dim HTMSignature As String = GetUserOutlookHTMLSignature()
         If HTMSignature.ToUpper.Contains("BODY") Then
            Dim i As Integer = HTMSignature.ToUpper.IndexOf("<BODY") + 4
            Return HTMSignature.Insert(HTMSignature.IndexOf(">", i) + 1, tmpResult.ToString)
         Else
            Return "<HTML><BODY>" & tmpResult.ToString & "</BODY></HTML>"
         End If
      End Get
   End Property
#End Region
#Region "Class Private Subs"
   Private Sub SetNewTab()
      Dim adlbW As Integer = 230
      Dim cW As Integer = 240
      Dim cLbH As Integer = 17
      With oAddressLabel
         .Text = oVendor.ContactNameAndAddress
         .Dock = DockStyle.None
         .Location = New Point(0, 5)
         .Size = New Size(adlbW, 125)
         .BorderStyle = BorderStyle.Fixed3D
      End With
      With oEditDescriptionButton
         .Name = "MultyDescriptionEdit"
         .Location = New Point(0, 130)
         .Size = New Size(155, 20)
         .Text = "Replace Text in Descriptions"
      End With
      With oCommentsLabel1
         .Text = "Vendor Comments"
         .Dock = DockStyle.None
         .Location = New Point(adlbW + 5, 5)
         .Size = New Size(cW, cLbH)
         .BorderStyle = BorderStyle.FixedSingle
      End With
      With oVendorCommentslabel
         .Multiline = True
         .ReadOnly = True
         .Text = oVendor.VendorComments
         .Dock = DockStyle.None
         .WordWrap = False
         .ScrollBars = ScrollBars.Horizontal
         .Location = New Point(adlbW + 5, 25)
         .Size = New Size(cW, 130)
         .BorderStyle = BorderStyle.Fixed3D
      End With
      With oCommentsLabel2
         .Text = "PO Comments"
         .Dock = DockStyle.None
         .Location = New Point(cW + adlbW + 10, 5)
         .Size = New Size(cW, cLbH)
         .BorderStyle = BorderStyle.FixedSingle
      End With
      With oPOCommentslabel
         .Multiline = True
         .ReadOnly = True
         .Text = oVendor.POComments
         .Dock = DockStyle.None
         .WordWrap = False
         .ScrollBars = ScrollBars.Horizontal
         .Location = New Point(cW + adlbW + 10, 25)
         .Size = New Size(cW, 130)
         .BorderStyle = BorderStyle.Fixed3D
      End With
      With oRequestLabel
         .Text = "Special Requests"
         .Dock = DockStyle.None
         .Location = New Point(cW + cW + adlbW + 15, 5)
         .Size = New Size(280, cLbH)
         .BorderStyle = BorderStyle.FixedSingle
      End With
      With oRequestText
         .Name = "RequestTextBox"
         .Multiline = True
         .Dock = DockStyle.None
         .WordWrap = False
         .ScrollBars = ScrollBars.Both
         .Location = New Point(cW + cW + adlbW + 15, 25)
         .Size = New Size(280, 110)
         .BorderStyle = BorderStyle.Fixed3D
      End With
      With oDueDatePicker
         .Name = "RECDATE"
         .Location = New Point(cW + cW + adlbW + 15, cLbH + 120)
         .Format = DateTimePickerFormat.Short
         .Size = New Size(120, 20)
         .Value = Today.AddDays(14)
      End With
      With oSetPODateButton
         .Name = "SetPODateButton"
         .Location = New Point(cW + cW + adlbW + 15 + 125, cLbH + 120)
         .Size = New Size(155, 20)
         .Text = "Set PO Items Due Dates"
         AddHandler .Click, AddressOf SetPODatesButton_Click
      End With
      If oVendorNumber = User.ID Then
         oTab.Text = "Short List"
      Else
         oTab.Text = oVendorNumber
      End If
      oMoveToShortList.Text = "Move To Short List"
      oMoveToVendor.Text = "Move to Vendor"
      oItemEventHistory.Text = "Item Activity History"
      With oPopUpMenu
         .Items.Add(oMoveToVendor)
         If Not oVendorNumber = User.ID Then
            .Items.Add(New ToolStripSeparator)
            .Items.Add(oMoveToShortList)
         End If
         .Items.Add(New ToolStripSeparator)
         .Items.Add(oItemEventHistory)
      End With
      oTab.Name = oVendorNumber
      With oTab.Controls
         .Add(oAddressLabel)
         .Add(oEditDescriptionButton)
         .Add(oCommentsLabel1)
         .Add(oVendorCommentslabel)
         .Add(oCommentsLabel2)
         .Add(oPOCommentslabel)
         .Add(oRequestLabel)
         .Add(oRequestText)
         If Not oVendorNumber = User.ID Then
            .Add(oDueDatePicker)
            .Add(oSetPODateButton)
         End If
         .Add(oDataViewGrid)
         .Add(oStatusStrip)
      End With
   End Sub
   Private Sub SetDataViewGrid()
      Dim ItemNoColumn As New DataGridViewTextBoxColumn
      With ItemNoColumn
         .Name = "cItemNo"
         .HeaderText = "Item Number"
         .CellTemplate.ValueType = GetType(String)
         .Width = 145
         .ReadOnly = True
      End With
      Dim VendorItemNoColumn As New DataGridViewTextBoxColumn
      With VendorItemNoColumn
         .Name = "cVendorItemNo"
         .HeaderText = "Vendor Item Number"
         .CellTemplate.ValueType = GetType(String)
         .Width = 145
      End With
      Dim ItemDescColumn As New DataGridViewTextBoxColumn
      With ItemDescColumn
         .Name = "cDescription"
         .HeaderText = "Description"
         .CellTemplate.ValueType = GetType(String)
         .Width = 340
         .MaxInputLength = 70
      End With
      Dim VendorNameColumn As New DataGridViewTextBoxColumn
      With VendorNameColumn
         .Name = "cVendor"
         .HeaderText = "Vendor"
         .CellTemplate.ValueType = GetType(String)
         .Width = 50
         .ReadOnly = True
      End With
      Dim DueDateColumn As New DataGridViewTextBoxColumn
      With DueDateColumn
         .Name = "cDueDate"
         .HeaderText = "Due Date"
         .Width = 70
         .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
         .CellTemplate.ValueType = GetType(Date)
         .DefaultCellStyle.Format = "d"
         If oVendorNumber = User.ID Then
            .ReadOnly = True
         End If
      End With
      Dim VendorPriceColumn As New DataGridViewTextBoxColumn
      With VendorPriceColumn
         .Name = "cVendorPrice"
         .HeaderText = "Cost"
         .CellTemplate.ValueType = GetType(Decimal)
         .Width = 50
         If oVendorNumber = User.ID Then
            .ReadOnly = True
         End If
         .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
      End With
      Dim TotalItemsToBuyColumn As New DataGridViewTextBoxColumn
      With TotalItemsToBuyColumn
         .Name = "cFinalNumberToBuy"
         .HeaderText = "Total to Buy"
         .CellTemplate.ValueType = GetType(Integer)
         .Width = 80
         If oVendorNumber = User.ID Then
            .ReadOnly = True
         End If
         .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
      End With
      Dim TotalItemsCostColumn As New DataGridViewTextBoxColumn
      With TotalItemsCostColumn
         .Name = "cTotalItemsCost"
         .HeaderText = "Total Cost"
         .Width = 70
         .ReadOnly = True
         .Visible = True
         .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
      End With
      Dim ItemsIDColumn As New DataGridViewTextBoxColumn
      With ItemsIDColumn
         .Name = "cItemID"
         .HeaderText = "Item ID"
         .Width = 50
         .ReadOnly = True
         .Visible = False
      End With
      Dim ItemsEDPNOColumn As New DataGridViewTextBoxColumn
      With ItemsEDPNOColumn
         .Name = "cItemEDPNO"
         .HeaderText = "Item EDPNO"
         .Width = 50
         .ReadOnly = True
         .Visible = False
      End With
      Dim tmpResult As New DataGridView
      With oDataViewGrid
         .TabIndex = 1
         .Location = New Point(0, 160)
         .Size = New Size(1007, 485)
         .Dock = DockStyle.Bottom
         .Name = "DataGridView"
         .AllowUserToAddRows = False
         .AllowUserToDeleteRows = False
         .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
         .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
         .EditMode = DataGridViewEditMode.EditOnEnter
         .RowHeadersWidth = 40
         .MultiSelect = False
         .Columns.Insert(0, ItemNoColumn)
         .Columns.Insert(1, VendorItemNoColumn)
         .Columns.Insert(2, ItemDescColumn)
         .Columns.Insert(3, VendorNameColumn)
         .Columns.Insert(4, DueDateColumn)
         .Columns.Insert(5, VendorPriceColumn)
         .Columns.Insert(6, TotalItemsToBuyColumn)
         .Columns.Insert(7, TotalItemsCostColumn)
         .Columns.Insert(8, ItemsIDColumn)
         .Columns.Insert(9, ItemsEDPNOColumn)
         For Each item As RecommendedBuyItem In oWorkingItems
            .Rows.Add(item.DataGridViewRow)
         Next
      End With
   End Sub
   Private Sub SetStatusStrip()
      oCountlabel.Text = "Items: " & oWorkingItems.Count
      oTotallabel.Text = "Total: " & WorkingTotalPurchase
      oPOlabel.Text = "PO: " & oPONumber & " " & oPODate
      oEmailLabel.Text = "Last Email: " & oLastEmailDate
      With oStatusStrip
         .Name = "StatusStrip"
         .Dock = DockStyle.Bottom
         .Items.Insert(0, oCountlabel)
         .Items.Insert(1, New ToolStripSeparator)
         .Items.Insert(2, oTotallabel)
         .Items.Insert(3, New ToolStripSeparator)
         .Items.Insert(4, oPOlabel)
         .Items.Insert(5, New ToolStripSeparator)
         .Items.Insert(6, oEmailLabel)
         .Items.Insert(7, New ToolStripSeparator)
         .Items.Insert(8, oInfolabel)
      End With
   End Sub
   Private Sub SetVendor()
      If oVendorNumber = User.ID Then
         oVendor = New ItemVendor(User)
      Else
         Using conn As New SqlConnection(SELECTVENDORSDB)
            Dim QueryString As String = "SELECT * FROM dbo.VENDORDFROM WHERE VENDORNO = '" & oVendorNumber & "'"
            Dim cmd As New SqlCommand(QueryString, conn)
            Try
               conn.Open()
               Dim r As SqlDataReader = cmd.ExecuteReader
               If r.HasRows Then
                  r.Read()
                  oVendor = New ItemVendor(r)
               End If
            Catch ex As Exception
               MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error getting vendor info")
            End Try
         End Using
      End If
   End Sub
   Private Sub SetTempPONumberDate()
      If oVendorNumber = User.ID Then
         oPONumber = "SHORTLIST"
         Return
      End If
      Dim tmpPO As Integer = 0
      For Each i As RecommendedBuyItem In oWorkingItems
         If i.TempPONumber <> "" Then
            If CInt(i.TempPONumber) > tmpPO Then
               tmpPO = CInt(i.TempPONumber)
               oPONumber = i.TempPONumber
               oPODate = i.TempPODateTime
            End If
         End If
      Next
   End Sub
   Private Sub SetLastEmailDate()
      If oVendorNumber = User.ID Then
         oLastEmailDate = ""
         Return
      End If
      For Each i As RecommendedBuyItem In oWorkingItems
         If i.LastEmailDate <> "" Then oLastEmailDate = i.LastEmailDate
      Next
   End Sub
   Private Sub SetPODatesButton_Click()
      Dim aDate As String = oDueDatePicker.Value.ToString("M/d/yy")
      For Each r As DataGridViewRow In oDataViewGrid.Rows
         Dim c As DataGridViewCell = r.Cells.Item("cDueDate")
         Dim tmpValue As String = ""
         If Trim(c.Value.ToString) = "" Then
            c.Value = aDate
            oWorkingItems.Find(RecommendedBuyItem.FindPredicateByItemId(r.Cells("cItemID").Value)).UpdateRecDateInfo(aDate)
         End If
      Next
   End Sub
   Private Sub ValidateVendorItemNumber(ByVal sender As Object, ByVal e As DataGridViewCellValidatingEventArgs)
      Dim anItem As RecommendedBuyItem = oWorkingItems.Find(RecommendedBuyItem.FindPredicateByItemId(oDataViewGrid.Rows(e.RowIndex).Cells("cItemID").Value))
      Dim ErrorText As String = ""
      Dim NewValue As String = ""
      NewValue = e.FormattedValue.ToString()
      If NewValue.Length > 20 Then
         ErrorText = "Vendor Item Number cannot be more than 20 characters long"
         oDataViewGrid.Rows(e.RowIndex).ErrorText = ErrorText
         MsgBox(ErrorText, MsgBoxStyle.Exclamation, "Data Error")
         e.Cancel = True
      Else
         anItem.UpdateVendorItemNumber(NewValue)
      End If
   End Sub
   Private Sub ValidateDescription(ByVal sender As Object, ByVal e As DataGridViewCellValidatingEventArgs)
      Dim anItem As RecommendedBuyItem = oWorkingItems.Find(RecommendedBuyItem.FindPredicateByItemId(oDataViewGrid.Rows(e.RowIndex).Cells("cItemID").Value))
      Dim ErrorText As String = ""
      Dim NewValue As String = ""
      NewValue = e.FormattedValue.ToString()
      If (String.IsNullOrEmpty(NewValue)) Then
         ErrorText = "Description must not be empty"
         oDataViewGrid.Rows(e.RowIndex).ErrorText = ErrorText
         MsgBox(ErrorText, MsgBoxStyle.Exclamation, "Data Error")
         e.Cancel = True
      ElseIf NewValue.Length > 70 Then
         ErrorText = "Description cannot be more than 70 characters long"
         oDataViewGrid.Rows(e.RowIndex).ErrorText = ErrorText
         MsgBox(ErrorText, MsgBoxStyle.Exclamation, "Data Error")
         e.Cancel = True
      Else
         anItem.UpdateDescription(NewValue)
      End If
   End Sub
   Private Sub ValidateDueDate(ByVal sender As Object, ByVal e As DataGridViewCellValidatingEventArgs)
      If oVendorNumber = User.ID Then Return
      Dim anItem As RecommendedBuyItem = oWorkingItems.Find(RecommendedBuyItem.FindPredicateByItemId(oDataViewGrid.Rows(e.RowIndex).Cells("cItemID").Value))
      Dim ErrorText As String = ""
      Dim NewValue As String = ""
      NewValue = e.FormattedValue.ToString()
      If Not (String.IsNullOrEmpty(NewValue)) Then
         If IsDate(NewValue) Then
            Dim ss As DateTime = CDate(NewValue)
            If ss.Ticks < Now.Ticks Then
               ErrorText = "Due date must not be in the past."
               oDataViewGrid.Rows(e.RowIndex).ErrorText = ErrorText
               MsgBox(ErrorText, MsgBoxStyle.Exclamation, "Date Format Error")
               e.Cancel = True
            ElseIf NewValue <> anItem.RECDATE Then
               anItem.UpdateRecDateInfo(NewValue)
            End If
         Else
            ErrorText = NewValue & " is not a date"
            oDataViewGrid.Rows(e.RowIndex).ErrorText = ErrorText
            MsgBox(ErrorText, MsgBoxStyle.Exclamation, "Date Format Error")
            e.Cancel = True
         End If
      Else
         anItem.UpdateRecDateInfo(NewValue)
      End If
   End Sub
   Private Sub ValidateCost(ByVal sender As Object, ByVal e As DataGridViewCellValidatingEventArgs)
      Dim anItemRow As DataGridViewRow = oDataViewGrid.Rows(e.RowIndex)
      Dim anItem As RecommendedBuyItem = oWorkingItems.Find(RecommendedBuyItem.FindPredicateByItemId(anItemRow.Cells("cItemID").Value))
      Dim ErrorText As String = ""
      Dim NewValue As String = ""
      NewValue = e.FormattedValue.ToString()
      If (String.IsNullOrEmpty(NewValue)) Then
         ErrorText = "Cost must not be empty"
         oDataViewGrid.Rows(e.RowIndex).ErrorText = ErrorText
         MsgBox(ErrorText, MsgBoxStyle.Exclamation, "Data Error")
         e.Cancel = True
      ElseIf Not (IsNumeric(NewValue.ToString)) Then
         ErrorText = "Cost must be a number"
         oDataViewGrid.Rows(e.RowIndex).ErrorText = ErrorText
         MsgBox(ErrorText, MsgBoxStyle.Exclamation, "Data Error")
         e.Cancel = True
      Else ' new value is number
         Dim oldcost As Decimal = anItem.VendorCost
         Dim newcost As Decimal = CDec(NewValue)
         If oldcost = newcost Then
            Debug.Print("same cost")
            Return
         End If
         If newcost = 0 Then
            If MsgBoxResult.Yes = MsgBox("Okay to set price of 0 dollars for the item?", MsgBoxStyle.YesNoCancel, "Zero Cost Confirm") Then
               UpdateCost(NewValue, anItem, anItemRow)
            Else
               ErrorText = "Item Zero Cost Error"
               oDataViewGrid.Rows(e.RowIndex).ErrorText = ErrorText
               e.Cancel = True
            End If
         End If
         If newcost > oldcost Then
            Debug.Print("Validating Cost.  Cost higher.  New: " & newcost & " old: " & oldcost)
            Dim increase As Decimal = 0
            If oldcost > 0 Then
               increase = ((newcost - oldcost) / oldcost) * 100
            Else
               increase = 99999
            End If
            Debug.Print("Cost percent increase: " & increase.ToString("f2"))
            If increase >= 2 Then
               'Dim ApprovalWindow As New ManagerApprovalWindow
               'ApprovalWindow.InfoLabel.Text =
               '   "Old Cost: " & oldcost & vbCrLf & "New Cost: " & newcost & vbCrLf & _
               '   "Increase: " & increase.ToString("f2") & "%" & vbCrLf & vbCrLf & _
               '   "Cost Increases equal or over 2% requiere manager approval." & vbCrLf & vbCrLf & _
               '   "Please enter manager password below to continue." & vbCrLf & _
               '   "Or click Cancel to abort change."

               Dim s As String =
                  "Old Cost: " & oldcost & vbCrLf & "New Cost: " & newcost & vbCrLf & _
                  "Increase: " & increase.ToString("f2") & "%" & vbCrLf & vbCrLf & _
                  "You will need manager approval to exit vendor if any item with an increase of over 2% percent remains in the PO."
               MsgBox(s, MsgBoxStyle.Exclamation, "PO Item Cost Increase Alert")
               oInfolabel.Text = "At least one item with over 2% cost increase in PO."
               UpdateCost(NewValue, anItem, anItemRow)

               'If DialogResult.Yes = ApprovalWindow.ShowDialog Then
               '   LogThis("ITMCOSTAPPRV", "Manager approved item cost increase from " & oldcost & " to " & newcost, anItem.ITEMID)
               '   UpdateCost(NewValue, anItem, anItemRow)
               'Else
               '   ErrorText = "Cost increases of 2% or more require manager approval"
               '   oDataViewGrid.Rows(e.RowIndex).ErrorText = ErrorText
               '   e.Cancel = True
               'End If

            Else ' increase is less than 2%
               Debug.Print("new cost lower then 2%")
               UpdateCost(NewValue, anItem, anItemRow)
            End If
         Else ' newcost is lower than old cost
            UpdateCost(NewValue, anItem, anItemRow)
         End If
      End If
   End Sub
   Private Sub UpdateCost(ByVal NewValue As Decimal, ByVal item As RecommendedBuyItem, ByVal itemRow As DataGridViewRow)
      item.UpdateVendorCost(NewValue)
      itemRow.Cells("cTotalItemsCost").Value = item.TotalVendorCost
      oInitialItems.Find(RecommendedBuyItem.FindPredicateByItemId(item.ITEMID)).UpdateVendorCost(NewValue, False)
      Debug.Print("Initial: " & InitialTotalPurchase & " Working: " & WorkingTotalPurchase)
      oTotallabel.Text = "Total:  " & WorkingTotalPurchase
   End Sub
   Private Sub ValidateTotalToBuy(ByVal sender As Object, ByVal e As DataGridViewCellValidatingEventArgs)
      Dim anItemRow As DataGridViewRow = oDataViewGrid.Rows(e.RowIndex)
      Dim anItem As RecommendedBuyItem = oWorkingItems.Find(RecommendedBuyItem.FindPredicateByItemId(anItemRow.Cells("cItemID").Value))
      Dim ErrorText As String = ""
      Dim NewValue As String = ""
      NewValue = e.FormattedValue.ToString()
      If (String.IsNullOrEmpty(NewValue)) Then
         ErrorText = "Quantity to buy must not be empty"
         oDataViewGrid.Rows(e.RowIndex).ErrorText = ErrorText
         MsgBox(ErrorText, MsgBoxStyle.Exclamation, "Data Error")
         e.Cancel = True
      ElseIf Not (IsNumeric(NewValue.ToString)) Then
         ErrorText = "Quantity must be a number"
         oDataViewGrid.Rows(e.RowIndex).ErrorText = ErrorText
         MsgBox(ErrorText, MsgBoxStyle.Exclamation, "Data Error")
         e.Cancel = True
      Else 'Quantity is a number
         Dim oldqty As Integer = anItem.TotalQtyToBuy
         Dim newqty As Integer = CInt(NewValue)
         If oldqty = newqty Then Return
         If newqty = 0 Then
            MoveToOtherVendor(anItem.CreateLeftOverQuantityItem(oldqty, User.ID))
            anItemRow.Cells.Item("cFinalNumberToBuy").Value = 0
            anItemRow.Cells.Item("cTotalItemsCost").Value = 0
            anItem.UpdatePOQtyToBuy("0", True)
            Exit Sub
         End If
         If newqty < oldqty Then
            Dim difference As Integer = oldqty - newqty
            MoveToOtherVendor(anItem.CreateLeftOverQuantityItem(difference, User.ID))
            anItem.UpdatePOQtyToBuy(newqty)
            anItemRow.Cells("cTotalItemsCost").Value = anItem.TotalVendorCost
            CheckTotalCostDifference()
         Else ' new qty is greater than old qty
            Dim QtyFromOtherTab As Integer = ExtraQuantityFromOtherTab(anItem, newqty - oldqty)
            Debug.Print("Qty from other tab: " & QtyFromOtherTab)
            If QtyFromOtherTab > 0 Then
               oInitialItems.Find(RecommendedBuyItem.FindPredicateByItemId(anItem.ITEMID)).UpdatePOQtyToBuy(anItem.TotalQtyToBuy + QtyFromOtherTab, False)
            End If
            anItem.UpdatePOQtyToBuy(newqty)
            anItemRow.Cells("cTotalItemsCost").Value = anItem.TotalVendorCost
            CheckTotalCostDifference()
         End If
         oTotallabel.Text = "Total:  " & WorkingTotalPurchase
      End If
   End Sub
   Private Sub MoveToOtherVendor(ByVal anItem As RecommendedBuyItem)
      Dim OtherVendorItemsList As VendorItems = UserVendorItemsList.Find(VendorItems.FindPredicateByVendorId(anItem.Vendor))
      If IsNothing(OtherVendorItemsList) Then
         Debug.Print("Destination Vendor " & anItem.Vendor & " not found")
         OtherVendorItemsList = New VendorItems(anItem.Vendor)
         RaiseEvent VendorAdded(OtherVendorItemsList)
      Else
         Debug.Print("Destination Vendor " & anItem.Vendor & " found")
         Dim tmpItem As RecommendedBuyItem = Nothing
         Try
            tmpItem = OtherVendorItemsList.WorkingItems.Find(RecommendedBuyItem.FindPredicateByEDPNO(anItem.EDPNO))
         Catch ex As Exception
         End Try
         If IsNothing(tmpItem) Then
            OtherVendorItemsList.AddItem(anItem)
         Else
            OtherVendorItemsList.DeleteItem(tmpItem)
            OtherVendorItemsList.AddItem(anItem.MergeItem(tmpItem))
         End If
      End If
   End Sub
   Private Sub oDataViewGrid_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles oDataViewGrid.CellEndEdit
      oDataViewGrid.Rows(e.RowIndex).ErrorText = String.Empty
   End Sub
   Private Sub oDataViewGrid_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles oDataViewGrid.CellMouseClick
      oDataViewGrid.ClearSelection()
      If e.RowIndex >= 0 And e.ColumnIndex >= 0 And e.Button = MouseButtons.Right Then
         oDataViewGrid.Rows(e.RowIndex).Selected = True
         Dim r As Rectangle = oDataViewGrid.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True)
         oPopUpMenu.Show(sender, r.Left + e.X, r.Top + e.Y)
      End If
   End Sub
   Private Sub oDataViewGrid_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles oDataViewGrid.CellValidating
      Select Case oDataViewGrid.Columns(e.ColumnIndex).HeaderText
         Case "Vendor Item Number"
            ValidateVendorItemNumber(sender, e)
         Case "Description"
            ValidateDescription(sender, e)
         Case "Due Date"
            ValidateDueDate(sender, e)
         Case "Cost"
            ValidateCost(sender, e)
         Case "Total to Buy"
            ValidateTotalToBuy(sender, e)
         Case Else
            Return
      End Select
   End Sub
   Private Sub oDataViewGrid_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles oDataViewGrid.Validating
      For Each row As DataGridViewRow In oDataViewGrid.Rows
         If row.Cells.Item("cFinalNumberToBuy").Value = 0 Then
            Dim itmid As String = row.Cells.Item("cItemID").Value
            oDataViewGrid.Rows.Remove(row)
            oInitialItems.RemoveAll(RecommendedBuyItem.FindPredicateByItemId(itmid))
            oWorkingItems.RemoveAll(RecommendedBuyItem.FindPredicateByItemId(itmid))
         End If
      Next
   End Sub
   Private Sub oMoveToShortList_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles oMoveToShortList.Click
      Dim anItem As RecommendedBuyItem = oWorkingItems.Find(RecommendedBuyItem.FindPredicateByItemId(oDataViewGrid.SelectedRows.Item(0).Cells("cItemID").Value))
      anItem.UpdateVendor(User.ID, True)
      MoveToOtherVendor(anItem)
      DeleteItem(anItem)
      If oWorkingItems.Count = 0 Then
         RaiseEvent VendorEmpty(Me)
         Return
      End If
   End Sub
   Private Sub oMoveToVendor_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles oMoveToVendor.Click
      Dim anItem As RecommendedBuyItem = oWorkingItems.Find(RecommendedBuyItem.FindPredicateByItemId(oDataViewGrid.SelectedRows.Item(0).Cells("cItemID").Value))
      Dim result As DialogResult
      Dim NewVendor As String = ""
      Dim SelectVendorForm As New SelectVendorWindow
      With SelectVendorForm
         With .DataGridView1.Rows
            For Each v As String() In anItem.ITEMVENDORS
               .Add(v)
            Next
            For Each s As String In {"PAUL", "SCOTTM", "BRIAN", "EDLYN", "TRACEY"}
               If User.ID <> s Then .Add({"", s, "", ""})
            Next
         End With
         .Text = anItem.ITEMNO & " Vendors"
         result = .ShowDialog
         NewVendor = .NewVendorSelected
      End With
      If Not result = DialogResult.OK Then Return

      If {"PAUL", "SCOTTM", "BRIAN", "EDLYN", "TRACEY"}.Contains(NewVendor) Then
         anItem.UpdateVendor(NewVendor, True)
      Else
         anItem.UpdateVendor(NewVendor, False)
      End If
      DeleteItem(anItem)
      If User.Vendors.Contains(NewVendor) Then
         MoveToOtherVendor(anItem)
      Else
         Dim nVendor As String = ""
         If {"PAUL", "SCOTTM", "BRIAN", "EDLYN", "TRACEY"}.Contains(NewVendor) Then
            nVendor = "Short List"
         Else
            nVendor = NewVendor
         End If
         MsgBox("Item " & anItem.ITEMNO & " moved to " & nVendor & " under buyer " & GetVerndorBuyer(NewVendor), MsgBoxStyle.Information, "Item moved")
      End If
      If oWorkingItems.Count = 0 Then
         RaiseEvent VendorEmpty(Me)
         Return
      End If
   End Sub
   Private Sub oItemEventHistory_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles oItemEventHistory.Click
      Dim anItem As RecommendedBuyItem = oWorkingItems.Find(RecommendedBuyItem.FindPredicateByItemId(oDataViewGrid.SelectedRows.Item(0).Cells("cItemID").Value))
      Dim ReportWindow As New ActivityHistoryForm
      With ReportWindow
         .Text = "Item Activity History"
         .ActivityHistoryTextBox.Text = anItem.ActivityHistory
         .ActivityHistoryTextBox.SelectionStart = 0
         .ActivityHistoryTextBox.SelectionLength = 0
         .Show()
      End With
   End Sub
   Private Sub oEditDescriptionButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles oEditDescriptionButton.Click
      Dim EditTextWindow As New ReplaceTextForm
      Dim FromText As String = ""
      Dim ToText As String = ""
      Dim result As DialogResult = DialogResult.Ignore
      With EditTextWindow
         result = .ShowDialog
         If Not result = DialogResult.OK Then Return
         FromText = .FromText
         ToText = .ToText
      End With
      Debug.Print("changing test from " & FromText & " to " & ToText)
      For Each r As DataGridViewRow In oDataViewGrid.Rows
         Dim tmpDescription As String = r.Cells.Item("cDescription").Value
         If tmpDescription.Contains(FromText) Then
            Dim tmpItem As RecommendedBuyItem = oWorkingItems.Find(RecommendedBuyItem.FindPredicateByItemId(r.Cells.Item("cItemID").Value))
            Dim newDesc As String = tmpDescription.Replace(FromText, ToText)
            r.Cells.Item("cDescription").Value = newDesc
            tmpItem.UpdateDescription(newDesc.Trim)
         End If
      Next
   End Sub
#End Region
#Region "Class Private Functions"
   Private Function InitialTotalPurchase() As Decimal
      Dim i As Decimal = 0
      For Each anitem As RecommendedBuyItem In oInitialItems
         i += anitem.VendorCost * anitem.TotalQtyToBuy
      Next
      Return i
   End Function
   Private Function ExtraQuantityFromOtherTab(ByVal anItem As RecommendedBuyItem, ByVal NeededQty As Integer) As Integer
      Debug.Print("Qty needed: " & NeededQty)
      Dim tmpResult As Integer = 0
      Dim ShortList As VendorItems = UserVendorItemsList.Find(VendorItems.FindPredicateByVendorId(User.ID))
      If Not IsNothing(ShortList) Then
         Dim tmpItem As RecommendedBuyItem = ShortList.WorkingItems.Find(RecommendedBuyItem.FindPredicateByEDPNO(anItem.EDPNO))
         If Not IsNothing(tmpItem) Then
            If tmpItem.TotalQtyToBuy <= NeededQty Then
               tmpResult = tmpItem.TotalQtyToBuy
               ShortList.DeleteItem(tmpItem)
               tmpItem.DeleteItem()
               Return tmpResult
            Else
               tmpItem.UpdatePOQtyToBuy(tmpItem.TotalQtyToBuy - NeededQty)
               ShortList.UpdateItemQuantityInGrid(tmpItem)
               Return NeededQty
            End If
         End If
      End If
      For Each vi As VendorItems In UserVendorItemsList
         If Not vi.VendorNumber = VendorNumber Then
            Dim tmpItem As RecommendedBuyItem = vi.WorkingItems.Find(RecommendedBuyItem.FindPredicateByEDPNO(anItem.EDPNO))
            If Not IsNothing(tmpItem) Then
               If tmpItem.TotalQtyToBuy <= NeededQty Then
                  tmpResult = tmpItem.TotalQtyToBuy
                  vi.DeleteItem(tmpItem)
                  tmpItem.DeleteItem()
                  Return tmpResult
               Else
                  tmpItem.UpdatePOQtyToBuy(tmpItem.TotalQtyToBuy - NeededQty)
                  vi.UpdateItemQuantityInGrid(anItem)
                  Return NeededQty
               End If
            End If
         End If
      Next
      Return tmpResult
   End Function
   Private Function CheckTotalCostDifference() As Boolean
      Dim oldc As Decimal = InitialTotalPurchase
      Dim newc As Decimal = WorkingTotalPurchase
      Debug.Print("Old Total Cost: " & oldc & " New Total Cost: " & newc)
      If newc > oldc Then
         Dim increase As Decimal = 0
         If oldc > 0 Then
            increase = ((newc - oldc) / oldc) * 100
         Else
            increase = 99999
         End If
         If increase >= 10 Then
            Dim s As String = "New total cost of " & newc & " is " & increase.ToString("f2") & "% higher than the initial cost of " & newc & "." & vbCrLf & "You will need manager approval to save the PO if its cost increase is equal or over 10%."
            MsgBox(s, MsgBoxStyle.Exclamation, "PO Cost Alert")
            oInfolabel.Text = "Inital Cost:  " & oldc & ".  New total cost over 10% of inital cost!"
            Return True
         Else
            oInfolabel.Text = ""
         End If
      Else
         oInfolabel.Text = ""
      End If
      Return False
   End Function
   Private Function GetOtherVendorItems(ByVal NewVendorNumber) As VendorItems
      Dim tmpResult As VendorItems = UserVendorItemsList.Find(VendorItems.FindPredicateByVendorId(NewVendorNumber))
      If IsNothing(tmpResult) Then
         tmpResult = New VendorItems(NewVendorNumber)
      End If
      Return tmpResult
   End Function
   Private Function GetVendorItems() As List(Of RecommendedBuyItem)
      Dim tmpResult As New List(Of RecommendedBuyItem)
      Dim QueryString As String = "SELECT * FROM RECMDBUYITEMS WHERE RBI_VENDOR = '" & oVendorNumber & "' AND RBI_PO_ISSUE_DATE IS NULL AND RBI_TOTAL_TO_BUY > 0 AND RBI_MGR_APPRVD IS NOT NULL"
      Using conn As New SqlConnection(ITEMSDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         Try
            conn.Open()
            Dim r As SqlDataReader = cmd.ExecuteReader
            If r.HasRows Then
               Do While r.Read
                  tmpResult.Add(New RecommendedBuyItem(r))
               Loop
            End If
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error Getting Items for " & oVendorNumber)
         End Try
      End Using
      Return tmpResult
   End Function
   Private Function UpdateItemMasterTable(ByVal i As RecommendedBuyItem) As Integer
      Dim FinalReqDate As Date = Date.Parse(i.RECDATE)
      Dim PON1 As Integer = 0
      Dim POD1 As String = ""
      Dim POQ1 As Integer = 0
      Dim PON2 As Integer = 0
      Dim POD2 As String = ""
      Dim POQ2 As Integer = 0
      Dim PON3 As Integer = 0
      Dim POD3 As String = ""
      Dim POQ3 As Integer = 0
      Dim PON4 As Integer = 0
      Dim POD4 As String = ""
      Dim POQ4 As Integer = 0
      Dim QueryString = "SELECT " & _
         "PONUMBERS_001,EXPECTEDDATE_001,NEXTQTY_001," & _
         "PONUMBERS_002,EXPECTEDDATE_002,NEXTQTY_002," & _
         "PONUMBERS_003,EXPECTEDDATE_003,NEXTQTY_003," & _
         "PONUMBERS_004,EXPECTEDDATE_004,NEXTQTY_004 " &
         "FROM ITEMMAST WHERE EDPNO = " & i.EDPNO
      Using conn As New SqlConnection(POENTRYDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         Try
            conn.Open()
            Dim r As SqlDataReader = cmd.ExecuteReader
            If r.HasRows Then
               r.Read()
               PON1 = r.Item("PONUMBERS_001")
               PON2 = r.Item("PONUMBERS_002")
               PON3 = r.Item("PONUMBERS_003")
               PON4 = r.Item("PONUMBERS_004")
               POD1 = r.Item("EXPECTEDDATE_001")
               POD2 = r.Item("EXPECTEDDATE_002")
               POD3 = r.Item("EXPECTEDDATE_003")
               POD4 = r.Item("EXPECTEDDATE_004")
               POQ1 = r.Item("NEXTQTY_001")
               POQ2 = r.Item("NEXTQTY_002")
               POQ3 = r.Item("NEXTQTY_003")
               POQ4 = r.Item("NEXTQTY_004")
            Else
               MsgBox(i.ITEMNO & " not found in Master Table", MsgBoxStyle.Exclamation, "Item Master table get exiting POs error")
               Return 0
            End If
            r.Close()
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Item Master table get exiting POs error")
         End Try
      End Using
      Using conn As New SqlConnection(POENTRYDB)
         QueryString = "UPDATE ITEMMAST SET " & _
            "PONUMBERS_001=@PONUMBERS_001," & _
            "PONUMBERS_002=@PONUMBERS_002," & _
            "PONUMBERS_003=@PONUMBERS_003," & _
            "PONUMBERS_004=@PONUMBERS_004," & _
            "PONUMBERS_005=@PONUMBERS_005," & _
            "EXPECTEDDATE_001=@EXPECTEDDATE_001," & _
            "EXPECTEDDATE_002=@EXPECTEDDATE_002," & _
            "EXPECTEDDATE_003=@EXPECTEDDATE_003," & _
            "EXPECTEDDATE_004=@EXPECTEDDATE_004," & _
            "EXPECTEDDATE_005=@EXPECTEDDATE_005," & _
            "NEXTQTY_001=@NEXTQTY_001," & _
            "NEXTQTY_002=@NEXTQTY_002," & _
            "NEXTQTY_003=@NEXTQTY_003," & _
            "NEXTQTY_004=@NEXTQTY_004," & _
            "NEXTQTY_005=@NEXTQTY_005 " & _
            "WHERE EDPNO=@EDPNO"
         Dim UpdateCmd As New SqlCommand(QueryString, conn)
         With UpdateCmd.Parameters
            .Add("@PONUMBERS_001", SqlDbType.BigInt).Value = PONUMBER
            .Add("@PONUMBERS_002", SqlDbType.BigInt).Value = PON1
            .Add("@PONUMBERS_003", SqlDbType.BigInt).Value = PON2
            .Add("@PONUMBERS_004", SqlDbType.BigInt).Value = PON3
            .Add("@PONUMBERS_005", SqlDbType.BigInt).Value = PON4
            .Add("@EXPECTEDDATE_001", SqlDbType.Char, 8).Value = FinalReqDate.ToString("yyyyMMdd")
            .Add("@EXPECTEDDATE_002", SqlDbType.Char, 8).Value = POD1
            .Add("@EXPECTEDDATE_003", SqlDbType.Char, 8).Value = POD2
            .Add("@EXPECTEDDATE_004", SqlDbType.Char, 8).Value = POD3
            .Add("@EXPECTEDDATE_005", SqlDbType.Char, 8).Value = POD4
            .Add("@NEXTQTY_001", SqlDbType.BigInt).Value = i.TotalQtyToBuy
            .Add("@NEXTQTY_002", SqlDbType.BigInt).Value = POQ1
            .Add("@NEXTQTY_003", SqlDbType.BigInt).Value = POQ2
            .Add("@NEXTQTY_004", SqlDbType.BigInt).Value = POQ3
            .Add("@NEXTQTY_005", SqlDbType.BigInt).Value = POQ4
            .Add("@EDPNO", SqlDbType.BigInt).Value = i.EDPNO
         End With
         Try
            conn.Open()
            Return UpdateCmd.ExecuteNonQuery()
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error Updating Item Master PO info")
            Return 0
         End Try
      End Using
      Return 0
   End Function
   Private Function UpdateVendorCostTable(ByVal i As RecommendedBuyItem) As Integer
      Dim QueryString As String = "UPDATE VENDORITEMS SET DOLLARCOST=@DOLLARCOST WHERE EDPNO=@EDPNO AND VENDORNO=@VENDORNO"
      Using conn As New SqlConnection(POENTRYDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         With cmd.Parameters
            .Add("@EDPNO", SqlDbType.BigInt).Value = i.EDPNO
            .Add("@VENDORNO", SqlDbType.Char, 10).Value = VendorNumber.PadRight(10, " ")
            .Add("@DOLLARCOST", SqlDbType.BigInt).Value = i.VendorCost * 10000
         End With
         Try
            conn.Open()
            Return cmd.ExecuteNonQuery
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error updating vendor cost")
         End Try
      End Using
      Return 0
   End Function
   Private Function GetFinalPODate() As String
      Using conn As New SqlConnection(POENTRYDB)
         Dim cmd As New SqlCommand("SELECT DATEX FROM POHEADER WHERE PONUMBER = " & PONUMBER, conn)
         Try
            conn.Open()
            Dim r As SqlDataReader = cmd.ExecuteReader
            If r.HasRows Then
               r.Read()
               Return r.Item("DATEX")
            End If
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Get Temp PO Date Error")
         End Try
      End Using
      Return ""
   End Function
   Private Function InsertPartIntoPODETAILS(ByVal itm As RecommendedBuyItem, ByVal line As String, ByVal FinalPODate As String) As Integer
      Dim VendorItemDetails As New VendorItemTableDetails
      Dim FinalReqDate As Date = Date.Parse(itm.RECDATE)
      VendorItemDetails = GetVendorItemsTableDetails(itm.EDPNO)
      Dim tmpResult As Integer = 0
      Dim QueryString As String = "INSERT INTO PODETAILS (PONUMBER,LINENOX,EDPNO,VENDORNO,PODATE,POQTY,REUSEDQTY,ORIGREQDATE,NEWREQDATE,REASON,FIRSTRECQTY,FIRSTRECDATE,TOTALRECQTY," & _
         "CANCELQTY,RETURNQTY,DAMAGEDQTY,LASTRECDATE,AUTOORDERNO,UNITOFMEAS,VENDUNITFACTOR,ACTUALCOST,LANDEDCOST,NODATECHG,ADDITIONALDATA,DATEX,STATUS,ADDPRODUCTCOST,PERCENTFLAG) " & _
         "VALUES(@PONUMBER,@LINENOX,@EDPNO,@VENDORNO,@PODATE,@POQTY,@REUSEDQTY,@ORIGREQDATE,@NEWREQDATE,@REASON,@FIRSTRECQTY,@FIRSTRECDATE,@TOTALRECQTY,@CANCELQTY,@RETURNQTY,@DAMAGEDQTY," & _
         "@LASTRECDATE,@AUTOORDERNO,@UNITOFMEAS,@VENDUNITFACTOR,@ACTUALCOST,@LANDEDCOST,@NODATECHG,@ADDITIONALDATA,@DATEX,@STATUS,@ADDPRODUCTCOST,@PERCENTFLAG)"
      Using conn As New SqlConnection(POENTRYDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         With cmd.Parameters
            .Add("@PONUMBER", SqlDbType.BigInt).Value = PONUMBER
            .Add("@LINENOX", SqlDbType.Char, 4).Value = line
            .Add("@EDPNO", SqlDbType.BigInt).Value = itm.EDPNO
            .Add("@VENDORNO", SqlDbType.Char, 10).Value = VendorNumber.PadRight(10, " ")
            .Add("@PODATE", SqlDbType.Char, 8).Value = FinalPODate
            .Add("@POQTY", SqlDbType.BigInt).Value = itm.TotalQtyToBuy
            .Add("@REUSEDQTY", SqlDbType.BigInt).Value = 0
            .Add("@ORIGREQDATE", SqlDbType.Char, 8).Value = FinalReqDate.ToString("yyyyMMdd")
            .Add("@NEWREQDATE", SqlDbType.Char, 8).Value = FinalReqDate.ToString("yyyyMMdd")
            .Add("@REASON", SqlDbType.Char, 4).Value = "02  "
            .Add("@FIRSTRECQTY", SqlDbType.BigInt).Value = 0
            .Add("@FIRSTRECDATE", SqlDbType.Char, 8).Value = "00000000"
            .Add("@TOTALRECQTY", SqlDbType.BigInt).Value = 0
            .Add("@CANCELQTY", SqlDbType.BigInt).Value = 0
            .Add("@RETURNQTY", SqlDbType.BigInt).Value = 0
            .Add("@DAMAGEDQTY", SqlDbType.BigInt).Value = 0
            .Add("@LASTRECDATE", SqlDbType.Char, 8).Value = "00000000"
            .Add("@AUTOORDERNO", SqlDbType.BigInt).Value = 0
            .Add("@UNITOFMEAS", SqlDbType.Char, 4).Value = VendorItemDetails.UNITOFMEAS
            .Add("@VENDUNITFACTOR", SqlDbType.BigInt).Value = VendorItemDetails.VENDUNITFACTOR
            .Add("@ACTUALCOST", SqlDbType.BigInt).Value = itm.ACTUALCOST
            .Add("@LANDEDCOST", SqlDbType.BigInt).Value = itm.ACTUALCOST
            .Add("@NODATECHG", SqlDbType.Char, 4).Value = "0000"
            .Add("@ADDITIONALDATA", SqlDbType.Char, 100).Value = "                        00000000@            0000                                                   "
            .Add("@DATEX", SqlDbType.Char, 8).Value = "00000000"
            .Add("@STATUS", SqlDbType.Char, 2).Value = "OP"
            .Add("@ADDPRODUCTCOST", SqlDbType.BigInt).Value = 0
            .Add("@PERCENTFLAG", SqlDbType.Char, 2).Value = "  "
         End With
         Try
            conn.Open()
            tmpResult = cmd.ExecuteNonQuery
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error Entering " & itm.ITEMNO & " into PO")
         End Try
      End Using
      Return tmpResult
   End Function
   Private Function RemoveDummyOK() As Boolean
      'I added extra select statment to double check the dummy PO was removed from the PO on 1/22/13
      'I have not tested the changes in live, only in ver and they worked
      Using conn As New SqlConnection(POENTRYDB)
         Dim cmd As New SqlCommand("DELETE FROM PODETAILS WHERE PONUMBER=@PONUMBER AND EDPNO=@EDPNO", conn)
         With cmd.Parameters
            .Add("@PONUMBER", SqlDbType.BigInt).Value = PONUMBER
            .Add("@EDPNO", SqlDbType.BigInt).Value = 21101
         End With
         Try
            conn.Open()
            If cmd.ExecuteNonQuery = 1 Then
               LogThis("DUMMYITMRM", "Dummy item removed okay for PO " & PONUMBER, 1000000)
               'Return True
            End If
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Delete Dummy From PO Error")
            LogThis("DUMMYITMRM", "Dummy item remove failed for PO " & PONUMBER, 1000000)
            Return False
         End Try
      End Using
      Using conn As New SqlConnection(POENTRYDB)
         Dim cmd As New SqlCommand("SELECT * FROM PODETAILS WHERE PONUMBER=@PONUMBER AND EDPNO=@EDPNO", conn)
         With cmd.Parameters
            .Add("@PONUMBER", SqlDbType.BigInt).Value = PONUMBER
            .Add("@EDPNO", SqlDbType.BigInt).Value = 21101
         End With
         Try
            conn.Open()
            Dim r As SqlDataReader = cmd.ExecuteReader
            If r.HasRows Then
               LogThis("DUMMYITMRM", "Dummy item remove failed for PO " & PONUMBER, 1000000)
               Return False
            Else
               LogThis("DUMMYITMRM", "Checked dummy item removed okay for PO " & PONUMBER, 1000000)
               Return True
            End If
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Delete Dummy From PO Error")
            LogThis("DUMMYITMRM", "Dummy item remove failed for PO " & PONUMBER, 1000000)
            Return False
         End Try
      End Using
      Return False
   End Function
   Private Function InsertNewTempPO(ByVal PONumber As Integer) As Integer
      Dim tmpResult As Integer = 0
      Dim v As ItemVendor = Vendor
      Dim QueryString As String = "INSERT INTO POHEADER(PONUMBER,VENDORNO,APVENDOR,COMPANY,DIVISION,SHIPTOADDRESS,TERMSPCT,TERMSDAYS,SHIPMETHOD,CITY,USERID,POTERMSCD,POTERMS,POBUYERCD," & _
               "POBUYER,DATEX,REQDATE,REQDATETYPE,POPAYMETHOD,LASTSHIPDATE,REASON,FIRSTSHIPDATE,CANAFTERDATE,CANAFTERCODE,NODATECHG,ADDITIONALDATA,STATUS)VALUES" & _
               "(@PONUMBER,@VENDORNO,@APVENDOR,@COMPANY,@DIVISION,@SHIPTOADDRESS,@TERMSPCT,@TERMSDAYS,@SHIPMETHOD,@CITY,@USERID,@POTERMSCD,@POTERMS,@POBUYERCD,@POBUYER," & _
               "@DATEX,@REQDATE,@REQDATETYPE,@POPAYMETHOD,@LASTSHIPDATE,@REASON,@FIRSTSHIPDATE,@CANAFTERDATE,@CANAFTERCODE,@NODATECHG,@ADDITIONALDATA,@STATUS)"
      Using conn As New SqlConnection(POENTRYDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         Dim TmpPODate As String = ""
         With cmd.Parameters
            .Add("@PONUMBER", SqlDbType.BigInt).Value = PONumber
            .Add("@VENDORNO", SqlDbType.Char, 10).Value = v.Number.PadRight(10, " ")
            .Add("@APVENDOR", SqlDbType.Char, 20).Value = v.APVENDOR.PadRight(20, " ")
            .Add("@COMPANY", SqlDbType.Char, 2).Value = "01"
            .Add("@DIVISION", SqlDbType.Char, 2).Value = "01"
            .Add("@SHIPTOADDRESS", SqlDbType.Char, 4).Value = "01  "
            .Add("@TERMSPCT", SqlDbType.Char, 4).Value = v.TERMSPCT
            .Add("@TERMSDAYS", SqlDbType.BigInt).Value = v.TERMSDAYS
            .Add("@SHIPMETHOD", SqlDbType.Char, 2).Value = "02"
            .Add("@CITY", SqlDbType.Char, 30).Value = v.FOBCITY.PadRight(30, " ")
            .Add("@USERID", SqlDbType.Char, 8).Value = User.ID.PadRight(8, " ")
            .Add("@POTERMSCD", SqlDbType.Char, 4).Value = "    "
            .Add("@POTERMS", SqlDbType.Char, 20).Value = "".PadRight(20, " ")
            .Add("@POBUYERCD", SqlDbType.Char, 4).Value = User.POBUYERCD
            .Add("@POBUYER", SqlDbType.Char, 20).Value = User.FullNameUpper.PadRight(20, " ")
            .Add("@DATEX", SqlDbType.Char, 8).Value = Now.ToString("yyyyMMdd")
            .Add("@REQDATE", SqlDbType.Char, 8).Value = Now.AddDays(30).ToString("yyyyMMdd")
            .Add("@REQDATETYPE", SqlDbType.Char, 2).Value = "  "
            .Add("@POPAYMETHOD", SqlDbType.Char, 2).Value = "  "
            .Add("@LASTSHIPDATE", SqlDbType.Char, 8).Value = "00000000"
            .Add("@REASON", SqlDbType.Char, 4).Value = "02  "
            .Add("@FIRSTSHIPDATE", SqlDbType.Char, 8).Value = "00000000"
            .Add("@CANAFTERDATE", SqlDbType.Char, 8).Value = "00000000"
            .Add("@CANAFTERCODE", SqlDbType.Char, 2).Value = "  "
            .Add("@NODATECHG", SqlDbType.Char, 4).Value = "0000"
            .Add("@ADDITIONALDATA", SqlDbType.Char, 100).Value = v.ADDITIONALDATA
            .Add("@STATUS", SqlDbType.Char, 2).Value = "  "
         End With
         Try
            conn.Open()
            tmpResult = cmd.ExecuteNonQuery
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error Entering Temp PO")
         End Try
      End Using
      Dim dummyitm As Integer = 0
      If tmpResult <> 0 Then
         dummyitm = InsertDummyPartIntoPODETAILS(PONumber, v)
      End If
      If dummyitm = 0 Then
         Return dummyitm
      End If
      Return tmpResult
   End Function
   Private Function InsertDummyPartIntoPODETAILS(ByVal PONumber As Integer, ByVal v As ItemVendor) As Integer
      Dim tmpResult As Integer = 0
      Dim QueryString As String = "INSERT INTO PODETAILS (PONUMBER,LINENOX,EDPNO,VENDORNO,PODATE,POQTY,REUSEDQTY,ORIGREQDATE,NEWREQDATE,REASON,FIRSTRECQTY,FIRSTRECDATE,TOTALRECQTY," & _
         "CANCELQTY,RETURNQTY,DAMAGEDQTY,LASTRECDATE,AUTOORDERNO,UNITOFMEAS,VENDUNITFACTOR,ACTUALCOST,LANDEDCOST,NODATECHG,ADDITIONALDATA,DATEX,STATUS,ADDPRODUCTCOST,PERCENTFLAG) " & _
         "VALUES(@PONUMBER,@LINENOX,@EDPNO,@VENDORNO,@PODATE,@POQTY,@REUSEDQTY,@ORIGREQDATE,@NEWREQDATE,@REASON,@FIRSTRECQTY,@FIRSTRECDATE,@TOTALRECQTY,@CANCELQTY,@RETURNQTY,@DAMAGEDQTY," & _
         "@LASTRECDATE,@AUTOORDERNO,@UNITOFMEAS,@VENDUNITFACTOR,@ACTUALCOST,@LANDEDCOST,@NODATECHG,@ADDITIONALDATA,@DATEX,@STATUS,@ADDPRODUCTCOST,@PERCENTFLAG)"
      Using conn As New SqlConnection(POENTRYDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         With cmd.Parameters
            .Add("@PONUMBER", SqlDbType.BigInt).Value = PONumber
            .Add("@LINENOX", SqlDbType.Char, 4).Value = "0001"
            .Add("@EDPNO", SqlDbType.BigInt).Value = 21101
            .Add("@VENDORNO", SqlDbType.Char, 10).Value = v.Number.PadRight(10, " ")
            .Add("@PODATE", SqlDbType.Char, 8).Value = Now.ToString("yyyyMMdd")
            .Add("@POQTY", SqlDbType.BigInt).Value = 1
            .Add("@REUSEDQTY", SqlDbType.BigInt).Value = 0
            .Add("@ORIGREQDATE", SqlDbType.Char, 8).Value = Now.AddDays(30).ToString("yyyyMMdd")
            .Add("@NEWREQDATE", SqlDbType.Char, 8).Value = Now.AddDays(30).ToString("yyyyMMdd")
            .Add("@REASON", SqlDbType.Char, 4).Value = "02  "
            .Add("@FIRSTRECQTY", SqlDbType.BigInt).Value = 0
            .Add("@FIRSTRECDATE", SqlDbType.Char, 8).Value = "00000000"
            .Add("@TOTALRECQTY", SqlDbType.BigInt).Value = 0
            .Add("@CANCELQTY", SqlDbType.BigInt).Value = 0
            .Add("@RETURNQTY", SqlDbType.BigInt).Value = 0
            .Add("@DAMAGEDQTY", SqlDbType.BigInt).Value = 0
            .Add("@LASTRECDATE", SqlDbType.Char, 8).Value = "00000000"
            .Add("@AUTOORDERNO", SqlDbType.BigInt).Value = 0
            .Add("@UNITOFMEAS", SqlDbType.Char, 4).Value = "    "
            .Add("@VENDUNITFACTOR", SqlDbType.BigInt).Value = 1
            .Add("@ACTUALCOST", SqlDbType.BigInt).Value = 0
            .Add("@LANDEDCOST", SqlDbType.BigInt).Value = 0
            .Add("@NODATECHG", SqlDbType.Char, 4).Value = "0000"
            .Add("@ADDITIONALDATA", SqlDbType.Char, 100).Value = "                        00000000@            0000                                                   "
            .Add("@DATEX", SqlDbType.Char, 8).Value = "00000000"
            .Add("@STATUS", SqlDbType.Char, 2).Value = "OP"
            .Add("@ADDPRODUCTCOST", SqlDbType.BigInt).Value = 0
            .Add("@PERCENTFLAG", SqlDbType.Char, 2).Value = "  "
         End With
         Try
            conn.Open()
            tmpResult = cmd.ExecuteNonQuery
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error Entering DUMMYPART into PO")
         End Try
      End Using
      Return tmpResult
   End Function
   Private Function DeepCopy(ByVal ObjectToCopy As Object) As Object
      Using mem As New MemoryStream
         Dim bf As New BinaryFormatter
         bf.Serialize(mem, ObjectToCopy)
         mem.Seek(0, SeekOrigin.Begin)
         Return bf.Deserialize(mem)
      End Using
   End Function
   Private Function GetVendorItemsTableDetails(ByVal EDPNO As String) As VendorItemTableDetails
      Dim tmpResult As New VendorItemTableDetails
      tmpResult.UNITOFMEAS = "    "
      tmpResult.VENDUNITFACTOR = 1
      Dim QueryString As String = "SELECT UNITOFMEAS,VENDUNITFACTOR FROM VENDORITEMS WHERE EDPNO=" & EDPNO & " AND VENDORNO='" & VendorNumber & "'"
      Using conn As New SqlConnection(POENTRYDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         Try
            conn.Open()
            Dim r As SqlDataReader = cmd.ExecuteReader
            If r.HasRows Then
               r.Read()
               tmpResult.UNITOFMEAS = r.Item("UNITOFMEAS")
               tmpResult.VENDUNITFACTOR = r.Item("VENDUNITFACTOR")
            End If
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Get Vendor Item Details Error")
         End Try
      End Using
      Return tmpResult
   End Function
   Private Function GetVerndorBuyer(ByVal VendorNumber As String) As String
      Dim QueryString As String = "SELECT RBV_VENDOR_BUYER FROM RECMDBUYVENDORS WHERE RBV_VENDOR_NUMBER = '" & VendorNumber & "'"
      Using conn As New SqlConnection(ITEMSDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         Try
            conn.Open()
            Dim r As SqlDataReader = cmd.ExecuteReader
            If r.HasRows Then
               r.Read()
               Return r.Item("RBV_VENDOR_BUYER")
            End If
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error getting vendor buyer")
         End Try
      End Using
      Return "PAUL"
   End Function
   Public Function TotalCostHigher() As Boolean
      Dim oldc As Decimal = InitialTotalPurchase
      Dim newc As Decimal = WorkingTotalPurchase
      If newc > oldc Then
         Dim increase As Decimal = 0
         If oldc > 0 Then
            increase = ((newc - oldc) / oldc) * 100
         Else
            increase = 99999
         End If
         If increase >= 10 Then
            oInfolabel.Text = "Inital Cost:  " & oldc & ".  New total cost over 10% of inital cost!"
            Dim ApprovalWindow As New ManagerApprovalWindow
            ApprovalWindow.InfoLabel.Text =
               "Old Cost: " & oldc & vbCrLf & _
               "New Cost: " & newc & vbCrLf & _
               "Increase: " & increase.ToString("f2") & "%" & vbCrLf & vbCrLf & _
               "Cost Increases equal or over 10% require manager approval." & vbCrLf & vbCrLf & _
               "Please enter manager password below to continue."
            If Not DialogResult.Yes = ApprovalWindow.ShowDialog Then
               Return True
            Else
               Dim oldCost As Decimal = InitialTotalPurchase()
               For Each i As RecommendedBuyItem In oInitialItems
                  i.UpdatePOQtyToBuy(oWorkingItems.Find(RecommendedBuyItem.FindPredicateByItemId(i.ITEMID)).TotalQtyToBuy, False)
                  LogThis("TOTCOSTAPPRV", i.ITEMNO & " manager approved total PO cost increase from " & oldCost & " to " & WorkingTotalPurchase, i.ITEMID)
               Next
               oInfolabel.Text = ""
            End If
         Else
            oInfolabel.Text = ""
         End If
      Else
         oInfolabel.Text = ""
      End If
      Return False
   End Function
   Private Function ValidatePONumber() As String
      If oVendorNumber = User.ID Then Return "SHORTLIST"
      If String.IsNullOrEmpty(oPONumber) Then Return CreateTempPurchaseOrder()
      Dim QueryString As String = "SELECT PONUMBER FROM POHEADER WHERE PONUMBER = " & oPONumber & " AND VENDORNO = '" & oVendorNumber & "' AND USERID = '" & User.ID & "'"
      Using conn As New SqlConnection(POENTRYDB)
         Dim cmdHeader As New SqlCommand(QueryString, conn)
         Try
            conn.Open()
            Dim r As SqlDataReader = cmdHeader.ExecuteReader
            If r.HasRows Then
               r.Read()
               Debug.Print("IN DB: " & r.Item("PONUMBER") & " In Grid: " & oPONumber & " If not equal, creating a new po number, else checking po item is dummy part")
               If r.Item("PONUMBER") <> oPONumber Then Return CreateTempPurchaseOrder()
            Else
               Debug.Print(oPONumber & " not found in db.  Creating a new one...")
               Return CreateTempPurchaseOrder()
            End If
            r.Close()
         Catch ex As Exception
            MsgBox(QueryString & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error Validating PO")
         End Try
         Dim cmdDetail As New SqlCommand("SELECT EDPNO FROM PODETAILS WHERE PONUMBER = " & oPONumber, conn)
         Try
            Dim r As SqlDataReader = cmdDetail.ExecuteReader
            Dim tmpString As String = ""
            Dim tmpCount As Integer = 0
            If r.HasRows Then
               While r.Read
                  tmpString = r.Item("EDPNO")
                  tmpCount += 1
               End While
               If tmpCount <> 1 Then Return CreateTempPurchaseOrder()
               If tmpString = "21101" Then Return oPONumber
            End If
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error Validating PO")
         End Try
      End Using
      Debug.Print("Error getting valid po, creating new temp PO")
      Return CreateTempPurchaseOrder()
   End Function
   Private Function CreateTempPurchaseOrder() As String
      Dim TempPO As Integer = 0
      Using conn As New SqlConnection(POENTRYDB)
         Dim cmd As New SqlCommand("SELECT CTLDATA FROM CTLMAST WHERE CTLID = '0000PO-NUMBER'", conn)
         Try
            conn.Open()
            Dim r As SqlDataReader = cmd.ExecuteReader
            If r.HasRows Then
               r.Read()
               TempPO = r.Item(0)
            Else
               TempPO = 0
            End If
         Catch ex As Exception
            TempPO = 0
         End Try
      End Using
      If TempPO > 0 Then
         Dim NEXTPO As Integer = TempPO + 1
         Using conn As New SqlConnection(POENTRYDB)
            Dim cmd As New SqlCommand("UPDATE CTLMAST SET CTLDATA = '" & NEXTPO.ToString.PadRight(20, " ") & "' WHERE CTLID = '0000PO-NUMBER'", conn)
            Try
               conn.Open()
               Dim i As Integer = cmd.ExecuteNonQuery
            Catch ex As Exception
               MsgBox("Error setting next PO", MsgBoxStyle.Exclamation, "PO Entry")
            End Try
         End Using
         InsertNewTempPO(TempPO)
      Else
         MsgBox("Temp PO No: " & TempPO, MsgBoxStyle.Exclamation, "Error Allocatting temp PO Number")
      End If
      For Each item As RecommendedBuyItem In oWorkingItems
         item.UpdateTempPOInfo(TempPO)
      Next
      oPONumber = TempPO
      oPODate = Now.ToString("yyyy-MM-dd HH:mm:ss")
      oPOlabel.Text = "PO: " & oPONumber & " " & oPODate
      Debug.Print("New PO: " & TempPO)
      Return TempPO
   End Function
   Private Function UpdatePOCommentsTable() As Integer
      Dim ThereIsOneComment As Boolean = False
      Dim lines(5) As String
      For i As Integer = 0 To 5
         If i <= oFinalPOComments.Count - 1 Then
            lines(i) = oFinalPOComments(i).PadRight(50, " ")
            ThereIsOneComment = True
         Else : lines(i) = "".PadRight(50, " ")
         End If
      Next
      If Not ThereIsOneComment Then Return 0
      Using conn As New SqlConnection(POENTRYDB)
         Dim QueryString As String =
            "INSERT INTO POCOMMENTS(PONOLINENO,USERID,STATUS,POCOMMENTS_001,POCOMMENTS_002,POCOMMENTS_003,POCOMMENTS_004,POCOMMENTS_005,POCOMMENTS_006,ADDITIONALDATA)VALUES" & _
            "(@PONOLINENO,@USERID,@STATUS,@POCOMMENTS_001,@POCOMMENTS_002,@POCOMMENTS_003,@POCOMMENTS_004,@POCOMMENTS_005,@POCOMMENTS_006,@ADDITIONALDATA)"
         Dim cmd As New SqlCommand(QueryString, conn)
         With cmd.Parameters
            .Add("@PONOLINENO", SqlDbType.Char, 12).Value = PONUMBER & "000"
            .Add("@USERID", SqlDbType.Char, 8).Value = User.ID.PadRight(8, " ")
            .Add("@STATUS", SqlDbType.Char, 2).Value = "  "
            .Add("@POCOMMENTS_001", SqlDbType.Char, 50).Value = lines(0)
            .Add("@POCOMMENTS_002", SqlDbType.Char, 50).Value = lines(1)
            .Add("@POCOMMENTS_003", SqlDbType.Char, 50).Value = lines(2)
            .Add("@POCOMMENTS_004", SqlDbType.Char, 50).Value = lines(3)
            .Add("@POCOMMENTS_005", SqlDbType.Char, 50).Value = lines(4)
            .Add("@POCOMMENTS_006", SqlDbType.Char, 50).Value = lines(5)
            .Add("@ADDITIONALDATA", SqlDbType.Char, 100).Value = "".PadRight(100, " ")
         End With
         Try
            conn.Open()
            Dim result As Integer = cmd.ExecuteNonQuery
            Debug.Print("Entered comments into " & PONUMBER)
            Return result
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error Entering Comments into PO")
            Return 0
         End Try
      End Using
   End Function
#End Region
   Public Sub UpdateLastEmailDate()
      For Each i As RecommendedBuyItem In oWorkingItems
         i.UpdateLastEmailDate()
      Next
      oLastEmailDate = Now.ToString("yyyy-MM-dd HH:mm:ss")
      oEmailLabel.Text = "Last Email: " & oLastEmailDate
   End Sub
   Public Sub UpdateItemQuantityInGrid(ByVal Item As RecommendedBuyItem)
      For Each r As DataGridViewRow In oDataViewGrid.Rows
         If r.Cells.Item("cItemEDPNO").Value = Item.EDPNO Then
            r.Cells.Item("cFinalNumberToBuy").Value = Item.TotalQtyToBuy
            r.Cells.Item("cTotalItemsCost").Value = Item.TotalVendorCost
         End If
      Next
      oTotallabel.Text = WorkingTotalPurchase
   End Sub
   Public Sub AddItem(ByVal anItem As RecommendedBuyItem)
      oWorkingItems.Add(anItem)
      oInitialItems.Add(anItem)
      oDataViewGrid.Rows.Add(anItem.DataGridViewRow)
      oTotallabel.Text = "Cost:  " & WorkingTotalPurchase
      oCountlabel.Text = "Items:  " & oWorkingItems.Count
   End Sub
   Public Sub DeleteItem(ByVal anItem As RecommendedBuyItem)
      oWorkingItems.Remove(anItem)
      oInitialItems.RemoveAll(RecommendedBuyItem.FindPredicateByItemId(anItem.ITEMID))
      'oInitialItems.Remove(anItem)
      For Each r As DataGridViewRow In oDataViewGrid.Rows
         If r.Cells.Item("cItemID").Value = anItem.ITEMID Then
            oDataViewGrid.Rows.Remove(r)
         End If
      Next
      oTotallabel.Text = "Cost:  " & WorkingTotalPurchase
      oCountlabel.Text = "Items:  " & oWorkingItems.Count
   End Sub
   Public Function FinalizePurchaseOrderSuccess() As Boolean
      If RemoveDummyOK() Then
         Dim FinalPODate As String = GetFinalPODate()
         If Not String.IsNullOrEmpty(FinalPODate.Trim) Then
            Dim line As Integer = 1
            For Each i As RecommendedBuyItem In oWorkingItems
               If InsertPartIntoPODETAILS(i, line.ToString("D4"), FinalPODate) <> 1 Then Return False
               line += 1
            Next
            For Each i As RecommendedBuyItem In oWorkingItems
               UpdateItemMasterTable(i)
               UpdateVendorCostTable(i)
               i.UpdateFinalPOInfo(oPONumber)
               LogThis("POFINALIZED", oPONumber & " finalized", i.ITEMID)
            Next
            UpdatePOCommentsTable()
            RaiseEvent VendorEmpty(Me)
            Return True
         End If
      End If
      Return False
   End Function
   Public Function CompareTo(ByVal other As VendorItems) As Integer Implements System.IComparable(Of VendorItems).CompareTo
      Return New CaseInsensitiveComparer().Compare(VendorNumber, other.VendorNumber)
   End Function
   Public Shared Function FindPredicate(ByVal VI As VendorItems) As Predicate(Of VendorItems)
      Return Function(VI2 As VendorItems) VI.VendorNumber = VI2.VendorNumber
   End Function
   Public Shared Function FindPredicateByVendorId(ByVal VI As String) As Predicate(Of VendorItems)
      Return Function(VI2 As VendorItems) VI = VI2.VendorNumber
   End Function
   Private Structure VendorItemTableDetails
      Dim UNITOFMEAS As String
      Dim VENDUNITFACTOR As Integer
   End Structure
   Private Function GetUserOutlookHTMLSignature() As String
      Dim mostRecentfile As String = ""
      Dim mostRecentTime As ULong = 0
      Try
         Dim SigFolder As String = "C:\Users\" & User.ID & "\AppData\Roaming\Microsoft\Signatures\"
         If Directory.Exists(SigFolder) Then
            For Each f As String In Directory.GetFiles("C:\Users\" & User.ID & "\AppData\Roaming\Microsoft\Signatures\", "*.htm")
               If File.GetLastWriteTime(f).Ticks > mostRecentTime Then
                  mostRecentTime = File.GetLastWriteTime(f).Ticks
                  mostRecentfile = f
               End If
            Next
         End If
      Catch ex As Exception
         MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
      End Try
      If mostRecentfile <> "" Then
         Return File.ReadAllText(mostRecentfile)
      Else
         Return ""
      End If
   End Function
End Class
<Serializable()> _
Public Class RecommendedBuyItem
   Implements IComparable(Of RecommendedBuyItem)
#Region "Class Private Variables"
   Private oItemID As Integer = 0
   Private oItemNo As String = ""
   Private oEDPNO As String = ""
   Private oItemDesc As String = ""
   Private oVendor As String = ""
   Private oOriginalVendor As String = ""
   Private oVendorNo As String = ""
   Private oVendorDesc As String = ""
   Private oStatus As String = ""
   Private oPrice As String = ""
   Private oMarginPercent As String = ""
   Private oAvg52Wk As String = ""
   Private oAvg26Wk As String = ""
   Private oAvg13Wk As String = ""
   Private oAvg8Wk As String = ""
   Private oAvg4Wk As String = ""
   Private oLastWk As String = ""
   Private oPONum1 As String = ""
   Private oPOExpDate1 As String = ""
   Private oPOQty1 As String = ""
   Private oPONum2 As String = ""
   Private oPOExpDate2 As String = ""
   Private oPOQty2 As String = ""
   Private oPONum3 As String = ""
   Private oPOExpDate3 As String = ""
   Private oPOQty3 As String = ""
   Private oPONum4 As String = ""
   Private oPOExpDate4 As String = ""
   Private oPOQty4 As String = ""
   Private oTotalDue As String = ""
   Private oOnHand As String = ""
   Private oWeeksOfStock As String = ""
   Private oReorderLevel As String = ""
   Private oVendorPrice As String = ""
   Private oOriginalVendorPrice As String = ""
   Private oMinQty As String = ""
   Private oBOQty As String = ""
   Private oRecmdBuy As String = ""
   Private oEnteredDate As String = ""
   Private oMgrRecmdBuy As String = ""
   Private oMgrAppredDate As String = ""
   Private oTonyRecmdBuy As String = ""
   Private oTonyApprovedDate As String = ""
   Private oTotalToBuy As Integer = -1
   Private oTempPONumber As String = ""
   Private oTempPODate As String = ""
   Private oRecDate As String = ""
   Private oFinalPONumber As String = ""
   Private oFinalPODate As String = ""
   Private oLastEmailDate As String = ""
#End Region
   Sub New(ByVal r As SqlDataReader)
      oItemID = r.Item("RBITEM_ID")
      oItemNo = r.Item("RBI_NUMBER")
      oEDPNO = r.Item("RBI_EDPNO_NUMBER")
      oItemDesc = r.Item("RBI_DESCRIPTION")
      oAvg52Wk = r.Item("RBI_AVG52WK")
      oAvg26Wk = r.Item("RBI_AVG26WK")
      oAvg13Wk = r.Item("RBI_AVG13WK")
      oAvg8Wk = r.Item("RBI_AVG8WK")
      oAvg4Wk = r.Item("RBI_AVG4WK")
      oLastWk = r.Item("RBI_LASTWK")
      oPONum1 = r.Item("RBI_PO_NUMBER1")
      oPOExpDate1 = r.Item("RBI_PO_EXPDATE1")
      oPOQty1 = r.Item("RBI_PO_QTY1")
      oPONum2 = r.Item("RBI_PO_NUMBER2")
      oPOExpDate2 = r.Item("RBI_PO_EXPDATE2")
      oPOQty2 = r.Item("RBI_PO_QTY2")
      oPONum3 = r.Item("RBI_PO_NUMBER3")
      oPOExpDate3 = r.Item("RBI_PO_EXPDATE3")
      oPOQty3 = r.Item("RBI_PO_QTY3")
      oPONum4 = r.Item("RBI_PO_NUMBER4")
      oPOExpDate4 = r.Item("RBI_PO_EXPDATE4")
      oPOQty4 = r.Item("RBI_PO_QTY4")
      oTotalDue = r.Item("RBI_TOTAL_DUE")
      oOnHand = r.Item("RBI_ONHAND")
      oWeeksOfStock = r.Item("RBI_WEEKSOFSTOCK")
      oReorderLevel = r.Item("RBI_REORDERLEVEL")
      oVendorPrice = r.Item("RBI_VENDOR_PRICE")
      oOriginalVendorPrice = r.Item("RBI_ORIGINAL_VENDOR_PRICE")
      oMinQty = r.Item("RBI_MINQTY")
      oBOQty = r.Item("RBI_BOQTY")
      oRecmdBuy = r.Item("RBI_RECMDBUY")
      oVendor = r.Item("RBI_VENDOR")
      oVendorNo = r.Item("RBI_VENDOR_ITM_NUMBER")
      oStatus = r.Item("RBI_STATUS")
      oPrice = r.Item("RBI_PRICE")
      oMarginPercent = r.Item("RBI_MARGIN")
      oVendorDesc = r.Item("RBI_VENDOR_DESCRIPTION")
      oEnteredDate = r.Item("RBI_ADDED_DATE")
      If Not IsDBNull(r.Item("RBI_MGR_RECMDBUY")) Then
         oMgrRecmdBuy = r.Item("RBI_MGR_RECMDBUY")
      End If
      If Not IsDBNull(r.Item("RBI_MGR_APPRVD")) Then
         oMgrAppredDate = r.Item("RBI_MGR_APPRVD")
      End If
      If Not IsDBNull(r.Item("RBI_TONY_RECMDBUY")) Then
         oTonyRecmdBuy = r.Item("RBI_TONY_RECMDBUY")
      End If
      If Not IsDBNull(r.Item("RBI_TONY_APPRVD")) Then
         oTonyApprovedDate = r.Item("RBI_TONY_APPRVD")
      End If
      If Not IsDBNull(r.Item("RBI_TOTAL_TO_BUY")) Then
         oTotalToBuy = r.Item("RBI_TOTAL_TO_BUY")
      End If
      If Not IsDBNull(r.Item("RBI_TMP_PO_NUMBER")) Then
         oTempPONumber = r.Item("RBI_TMP_PO_NUMBER")
      End If
      If Not IsDBNull(r.Item("RBI_TMP_PO_DATE")) Then
         oTempPODate = r.Item("RBI_TMP_PO_DATE")
      End If
      If Not IsDBNull(r.Item("RBI_REQDATE")) Then
         oRecDate = r.Item("RBI_REQDATE")
      End If
      If Not IsDBNull(r.Item("RBI_ORIGINAL_VENDOR")) Then
         oOriginalVendor = r.Item("RBI_ORIGINAL_VENDOR")
      End If
      If Not IsDBNull(r.Item("RBI_LAST_EMAIL_DATE")) Then
         oLastEmailDate = r.Item("RBI_LAST_EMAIL_DATE")
      End If
      'LogThis("ITMCREATED", oItemNo & " " & oTotalToBuy & " qty created for " & oVendor & " on qty change", oItemID)
   End Sub
   Public Function CreateLeftOverQuantityItem(ByVal LeftOverQtyToBuy As Integer, ByVal NewVendor As String) As RecommendedBuyItem
      Dim tmpResult As RecommendedBuyItem = Nothing
      Dim result As Integer = 0
      Dim QueryString As String =
         "INSERT INTO RECMDBUYITEMS" & _
"(RBI_NUMBER,RBI_EDPNO_NUMBER,RBI_DESCRIPTION,RBI_STATUS,RBI_PRICE,RBI_MARGIN,RBI_VENDOR,RBI_VENDOR_ITM_NUMBER,RBI_VENDOR_DESCRIPTION,RBI_VENDOR_PRICE,RBI_AVG52WK,RBI_AVG26WK" & _
",RBI_AVG13WK,RBI_AVG8WK,RBI_AVG4WK,RBI_LASTWK,RBI_PO_NUMBER1,RBI_PO_EXPDATE1,RBI_PO_QTY1,RBI_PO_NUMBER2,RBI_PO_EXPDATE2,RBI_PO_QTY2,RBI_PO_NUMBER3,RBI_PO_EXPDATE3,RBI_PO_QTY3" & _
",RBI_PO_NUMBER4,RBI_PO_EXPDATE4,RBI_PO_QTY4,RBI_TOTAL_DUE,RBI_ONHAND,RBI_WEEKSOFSTOCK,RBI_REORDERLEVEL,RBI_MINQTY,RBI_BOQTY,RBI_RECMDBUY,RBI_ADDED_DATE,RBI_MGR_RECMDBUY,RBI_MGR_APPRVD" & _
",RBI_TONY_RECMDBUY,RBI_TONY_APPRVD,RBI_TOTAL_TO_BUY,RBI_PO_ISSUE_NUMBER,RBI_REQDATE,RBI_VENDOR_QTY_CONFIRM,RBI_VENDOR_COST_CONFIRM,RBI_TMP_PO_NUMBER,RBI_ORIGINAL_VENDOR,RBI_ORIGINAL_VENDOR_PRICE)VALUES" & _
"(@RBI_NUMBER,@RBI_EDPNO_NUMBER,@RBI_DESCRIPTION,@RBI_STATUS,@RBI_PRICE,@RBI_MARGIN,@RBI_VENDOR,@RBI_VENDOR_ITM_NUMBER,@RBI_VENDOR_DESCRIPTION,@RBI_VENDOR_PRICE,@RBI_AVG52WK,@RBI_AVG26WK" & _
",@RBI_AVG13WK,@RBI_AVG8WK,@RBI_AVG4WK,@RBI_LASTWK,@RBI_PO_NUMBER1,@RBI_PO_EXPDATE1,@RBI_PO_QTY1,@RBI_PO_NUMBER2,@RBI_PO_EXPDATE2,@RBI_PO_QTY2,@RBI_PO_NUMBER3,@RBI_PO_EXPDATE3,@RBI_PO_QTY3" & _
",@RBI_PO_NUMBER4,@RBI_PO_EXPDATE4,@RBI_PO_QTY4,@RBI_TOTAL_DUE,@RBI_ONHAND,@RBI_WEEKSOFSTOCK,@RBI_REORDERLEVEL,@RBI_MINQTY,@RBI_BOQTY,@RBI_RECMDBUY,@RBI_ADDED_DATE,@RBI_MGR_RECMDBUY,@RBI_MGR_APPRVD" & _
",@RBI_TONY_RECMDBUY,@RBI_TONY_APPRVD,@RBI_TOTAL_TO_BUY,@RBI_PO_ISSUE_NUMBER,@RBI_REQDATE,@RBI_VENDOR_QTY_CONFIRM,@RBI_VENDOR_COST_CONFIRM,@RBI_TMP_PO_NUMBER,@RBI_ORIGINAL_VENDOR,@RBI_ORIGINAL_VENDOR_PRICE);SELECT Scope_Identity()"
      Using conn As New SqlConnection(ITEMSDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         With cmd.Parameters
            .Add("@RBI_NUMBER", SqlDbType.Char, 20).Value = oItemNo
            .Add("@RBI_EDPNO_NUMBER", SqlDbType.BigInt).Value = oEDPNO
            .Add("@RBI_DESCRIPTION", SqlDbType.Char, 50).Value = oItemDesc
            .Add("@RBI_STATUS", SqlDbType.Char, 2).Value = oStatus
            .Add("@RBI_PRICE", SqlDbType.Char, 8).Value = oPrice
            .Add("@RBI_MARGIN", SqlDbType.Char, 8).Value = oMarginPercent
            .Add("@RBI_VENDOR", SqlDbType.Char, 10).Value = NewVendor
            .Add("@RBI_VENDOR_ITM_NUMBER", SqlDbType.Char, 20).Value = oVendorNo
            .Add("@RBI_VENDOR_DESCRIPTION", SqlDbType.Char, 50).Value = oVendorDesc
            .Add("@RBI_VENDOR_PRICE", SqlDbType.Char, 8).Value = oVendorPrice
            .Add("@RBI_ORIGINAL_VENDOR_PRICE", SqlDbType.Char, 8).Value = oOriginalVendorPrice
            .Add("@RBI_AVG52WK", SqlDbType.Char, 8).Value = oAvg52Wk
            .Add("@RBI_AVG26WK", SqlDbType.Char, 8).Value = oAvg26Wk
            .Add("@RBI_AVG13WK", SqlDbType.Char, 8).Value = oAvg13Wk
            .Add("@RBI_AVG8WK", SqlDbType.Char, 8).Value = oAvg8Wk
            .Add("@RBI_AVG4WK", SqlDbType.Char, 8).Value = oAvg4Wk
            .Add("@RBI_LASTWK", SqlDbType.Char, 8).Value = oLastWk
            .Add("@RBI_PO_NUMBER1", SqlDbType.Char, 9).Value = oPONum1
            .Add("@RBI_PO_EXPDATE1", SqlDbType.Char, 8).Value = oPOExpDate1
            .Add("@RBI_PO_QTY1", SqlDbType.Char, 4).Value = oPOQty1
            .Add("@RBI_PO_NUMBER2", SqlDbType.Char, 9).Value = oPONum2
            .Add("@RBI_PO_EXPDATE2", SqlDbType.Char, 8).Value = oPOExpDate2
            .Add("@RBI_PO_QTY2", SqlDbType.Char, 4).Value = oPOQty2
            .Add("@RBI_PO_NUMBER3", SqlDbType.Char, 9).Value = oPONum3
            .Add("@RBI_PO_EXPDATE3", SqlDbType.Char, 8).Value = oPOExpDate3
            .Add("@RBI_PO_QTY3", SqlDbType.Char, 4).Value = oPOQty3
            .Add("@RBI_PO_NUMBER4", SqlDbType.Char, 9).Value = oPONum4
            .Add("@RBI_PO_EXPDATE4", SqlDbType.Char, 8).Value = oPOExpDate4
            .Add("@RBI_PO_QTY4", SqlDbType.Char, 4).Value = oPOQty4
            .Add("@RBI_TOTAL_DUE", SqlDbType.Char, 7).Value = oTotalDue
            .Add("@RBI_ONHAND", SqlDbType.Char, 7).Value = oOnHand
            .Add("@RBI_WEEKSOFSTOCK", SqlDbType.Char, 4).Value = oWeeksOfStock
            .Add("@RBI_REORDERLEVEL", SqlDbType.Char, 4).Value = oReorderLevel
            .Add("@RBI_MINQTY", SqlDbType.Char, 7).Value = oMinQty
            .Add("@RBI_BOQTY", SqlDbType.Char, 7).Value = oBOQty
            .Add("@RBI_RECMDBUY", SqlDbType.Char, 7).Value = oRecmdBuy
            .Add("@RBI_MGR_APPRVD", SqlDbType.DateTime).Value = Date.Parse(oMgrAppredDate).ToString("yyyy-MM-ddTHH:mm:ss")
            .Add("@RBI_TONY_APPRVD", SqlDbType.DateTime).Value = Date.Parse(oTonyApprovedDate).ToString("yyyy-MM-ddTHH:mm:ss")
            .Add("@RBI_ADDED_DATE", SqlDbType.DateTime).Value = Date.Parse(oEnteredDate).ToString("yyyy-MM-ddTHH:mm:ss")
            .Add("@RBI_MGR_RECMDBUY", SqlDbType.Char, 7).Value = oMgrRecmdBuy
            .Add("@RBI_TONY_RECMDBUY", SqlDbType.Char, 7).Value = oTonyRecmdBuy
            .Add("@RBI_TOTAL_TO_BUY", SqlDbType.Int).Value = LeftOverQtyToBuy
            .Add("@RBI_PO_ISSUE_NUMBER", SqlDbType.Char, 10).Value = ""
            .Add("@RBI_REQDATE", SqlDbType.Char, 8).Value = ""
            .Add("@RBI_VENDOR_QTY_CONFIRM", SqlDbType.Char, 4).Value = ""
            .Add("@RBI_VENDOR_COST_CONFIRM", SqlDbType.Char, 4).Value = ""
            .Add("@RBI_TMP_PO_NUMBER", SqlDbType.Char, 10).Value = ""
            .Add("@RBI_ORIGINAL_VENDOR", SqlDbType.Char, 10).Value = oVendor
         End With
         Try
            conn.Open()
            result = cmd.ExecuteScalar()
            LogThis("ITMMOVED", oItemNo & " " & LeftOverQtyToBuy & " moved to vendor " & NewVendor, oItemID)
         Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.ToString, MsgBoxStyle.Exclamation, "Left Over Item entry Error")
         End Try
         Dim cmd2 As New SqlCommand("SELECT * FROM RECMDBUYITEMS WHERE RBITEM_ID= " & result, conn)
         Try
            Dim r As SqlDataReader = cmd2.ExecuteReader
            If r.HasRows Then
               r.Read()
               Return New RecommendedBuyItem(r)
            End If
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Left Over Item entry Error")
         End Try
      End Using
      Return tmpResult
   End Function
   ReadOnly Property DataGridViewRow() As String()
      Get
         If oOriginalVendor.Trim.Length > 2 Then
            Dim tmpstring() As String = {ITEMNO, oVendorNo.Trim, oItemDesc.Trim, oOriginalVendor, RECDATE, oVendorPrice.Trim, oTotalToBuy, TotalVendorCost, oItemID, oEDPNO}
            Return tmpstring
         Else
            Dim tmpstring() As String = {ITEMNO, oVendorNo.Trim, oItemDesc.Trim, Vendor, RECDATE, oVendorPrice.Trim, oTotalToBuy, TotalVendorCost, oItemID, oEDPNO}
            Return tmpstring
         End If
      End Get
   End Property
   ReadOnly Property ITEMNO() As String
      Get
         Return oItemNo.Trim
      End Get
   End Property
   ReadOnly Property EDPNO() As Integer
      Get
         Return CInt(oEDPNO)
      End Get
   End Property
   ReadOnly Property VendorItemNo() As String
      Get
         Return oVendorNo.Trim
      End Get
   End Property
   ReadOnly Property Description() As String
      Get
         Return oItemDesc.Trim
      End Get
   End Property
   ReadOnly Property TotalQtyToBuy() As Integer
      Get
         Return oTotalToBuy
      End Get
   End Property
   ReadOnly Property ACTUALCOST() As UInteger
      Get
         Return VendorCost * 1000000
      End Get
   End Property
   ReadOnly Property OriginalVendorCost() As Decimal
      Get
         Return CDec(oOriginalVendorPrice.Trim)
      End Get
   End Property
   ReadOnly Property VendorCost() As Decimal
      Get
         Return CDec(oVendorPrice.Trim)
      End Get
   End Property
   ReadOnly Property Price() As Decimal
      Get
         Return CDec(oPrice.Trim)
      End Get
   End Property
   ReadOnly Property Margin() As Decimal
      Get
         Return GetMargin(Price, VendorCost)
      End Get
   End Property
   ReadOnly Property OriginalMargin() As Decimal
      Get
         Return GetMargin(Price, OriginalVendorCost)
      End Get
   End Property
   ReadOnly Property TotalVendorCost() As Decimal
      Get
         Return VendorCost * TotalQtyToBuy
      End Get
   End Property
   ReadOnly Property Vendor() As String
      Get
         Return oVendor.Trim
      End Get
   End Property
   ReadOnly Property ITEMID() As String
      Get
         Return oItemID
      End Get
   End Property
   ReadOnly Property TempPONumber() As String
      Get
         Return oTempPONumber.Trim
      End Get
   End Property
   ReadOnly Property RECDATE() As String
      Get
         Return oRecDate
      End Get
   End Property
   ReadOnly Property LastEmailDate() As String
      Get
         Return oLastEmailDate
      End Get
   End Property
   ReadOnly Property TempPODateTime() As String
      Get
         Return oTempPODate.Replace("T", " ")
      End Get
   End Property
   ReadOnly Property ITEMVENDORS() As List(Of String())
      Get
         Dim tmpResutl As New List(Of String())
         Dim QueryString As String = "SELECT VENDORNO,PREFERENCE,CAST(DOLLARCOST AS MONEY) /10000 AS COST FROM VENDORITEMS WHERE EDPNO=" & oEDPNO & " ORDER BY PREFERENCE"
         Using conn As New SqlConnection(SELECTVENDORSDB)
            Dim cmd As New SqlCommand(QueryString, conn)
            Try
               conn.Open()
               Dim r As SqlDataReader = cmd.ExecuteReader
               If r.HasRows Then
                  Do While r.Read
                     Dim cost As Decimal = CDec(r.Item("COST"))
                     tmpResutl.Add({"", Trim(r.Item("VENDORNO")), r.Item("PREFERENCE"), cost.ToString("f2")})
                  Loop
               End If
            Catch ex As Exception
               MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error getting item vendors")
            End Try
         End Using
         Return tmpResutl
      End Get
   End Property
   ReadOnly Property ActivityHistory() As String
      Get
         Dim tmpResult As String = ""
         Using conn As New SqlConnection(ITEMSDB)
            Dim QueryString As String = "SELECT * FROM RECMDBUYITEMACTIONS WHERE RBITEM_ID = @RBITEM_ID"
            Dim cmd As New SqlCommand(QueryString, conn)
            cmd.Parameters.Add("@RBITEM_ID", SqlDbType.BigInt).Value = oItemID
            Try
               conn.Open()
               Dim r As SqlDataReader = cmd.ExecuteReader
               If r.HasRows Then
                  While r.Read
                     tmpResult &= String.Format("{0,-20}{1}" & Environment.NewLine, CDate(r.Item("RBIA_ACTION_DATETIME")).ToString, r.Item("RBIA_ACTION_LONG_TEXT"))
                  End While
               End If
            Catch ex As Exception
               Debug.Print("Error Entering Item Action into DB: " & ex.Message)
            End Try
         End Using
         Return tmpResult
      End Get
   End Property
   Private Function GetMargin(ByVal price As String, ByVal cost As String) As Decimal
      If price < 1 Then Return 0
      Return ((price - cost) / price) * 100
   End Function
   Public Function MergeItem(ByVal Item As RecommendedBuyItem) As RecommendedBuyItem
      Dim itemIdToKeep As Integer = 0
      Dim itemIdToDelete As Integer = 0
      If oItemID < Item.ITEMID Then
         UpdatePOQtyToBuy(oTotalToBuy + Item.TotalQtyToBuy)
         Item.DeleteItem()
         Debug.Print("Merged: " & oItemNo & " Kept: " & oItemID)
         Return Me
      Else
         Item.UpdatePOQtyToBuy(oTotalToBuy + Item.TotalQtyToBuy)
         DeleteItem()
         Debug.Print("Merged: " & oItemNo & " Kept: " & Item.oItemID)
         Return Item
      End If
   End Function
   Public Function DeleteItem() As Integer
      Dim QueryString As String = "DELETE FROM RECMDBUYITEMS WHERE RBITEM_ID = " & oItemID
      Using conn As New SqlConnection(ITEMSDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         Try
            conn.Open()
            Debug.Print("Deleted: " & oItemID)
            Dim result As Integer = cmd.ExecuteNonQuery()
            LogThis("ITMDELETED", oItemNo & " with ID " & oItemID & " was deleted on merge", oItemID)
            Return result
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error Deleting Merged Item")
            Return 0
         End Try
      End Using
   End Function
   Public Function UpdateTempPOInfo(ByVal aPONumber) As Integer
      Dim tmpResult As Integer = 0
      oTempPONumber = aPONumber
      oTempPODate = Now.ToString("yyyy-MM-ddTHH:mm:ss")
      Dim QueryString As String = "UPDATE RECMDBUYITEMS SET RBI_TMP_PO_NUMBER = '" & oTempPONumber & "',RBI_TMP_PO_DATE = '" & oTempPODate & "' WHERE RBITEM_ID = " & oItemID
      Using conn As New SqlConnection(ITEMSDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         Try
            conn.Open()
            Dim result As Integer = cmd.ExecuteNonQuery()
            LogThis("ITMTEMPPO", oItemNo & " had temp PO " & oTempPONumber & " created", oItemID)
            Return result
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error Updating Temp PO Info")
            Return 0
         End Try
      End Using
      Return tmpResult
   End Function
   Public Function UpdateFinalPOInfo(ByVal aPONumber) As Integer
      Dim tmpResult As Integer = 0
      Dim QueryString As String = "UPDATE RECMDBUYITEMS SET RBI_PO_ISSUE_NUMBER = '" & aPONumber & "',RBI_PO_ISSUE_DATE = '" & Now.ToString("yyyy-MM-ddTHH:mm:ss") & "' WHERE RBITEM_ID = " & oItemID
      Using conn As New SqlConnection(ITEMSDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         Try
            conn.Open()
            Dim result As Integer = cmd.ExecuteNonQuery()
            LogThis("ITMFINALPO", oItemNo & " was included in PO " & aPONumber, oItemID)
            Return result
         Catch ex As Exception
            MsgBox(ex.Message)
            Return 0
         End Try
      End Using
      Return tmpResult
   End Function
   Public Function UpdateRecDateInfo(ByVal aRecDate As String) As Integer
      Dim tmpResult As Integer = 0
      oRecDate = aRecDate
      Dim QueryString As String = "UPDATE RECMDBUYITEMS SET RBI_REQDATE = '" & oRecDate & "' WHERE RBITEM_ID = " & oItemID
      Using conn As New SqlConnection(ITEMSDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         Try
            conn.Open()
            Dim result As Integer = cmd.ExecuteNonQuery()
            LogThis("NEWITMREQDATE", oItemNo & " has new reqdate of " & oRecDate, oItemID)
            Return result
         Catch ex As Exception
            MsgBox(ex.Message)
            Return 0
         End Try
      End Using
      Return tmpResult
   End Function
   Public Function UpdateLastEmailDate() As Integer
      Dim tmpResult As Integer = 0
      Dim QueryString As String = "UPDATE RECMDBUYITEMS SET RBI_LAST_EMAIL_DATE = '" & Now.ToString("yyyy-MM-ddTHH:mm:ss") & "' WHERE RBITEM_ID = " & oItemID
      Using conn As New SqlConnection(ITEMSDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         Try
            conn.Open()
            Dim result As Integer = cmd.ExecuteNonQuery()
            LogThis("NEWEMAILDATE", oItemNo & " included in email sent to vendor " & oVendor, oItemID)
            Return result
         Catch ex As Exception
            MsgBox(ex.Message)
            Return 0
         End Try
      End Using
      Return tmpResult
   End Function
   Public Function UpdateVendorItemNumber(ByVal aVendorItemNumber As String) As Integer
      Dim tmpResult As Integer = 0
      oVendorNo = aVendorItemNumber
      Dim QueryString As String = "UPDATE RECMDBUYITEMS SET RBI_VENDOR_ITM_NUMBER = '" & oVendorNo.Replace("'", "''") & "' WHERE RBITEM_ID = " & oItemID
      Using conn As New SqlConnection(ITEMSDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         Try
            conn.Open()
            LogThis("NEWVITMNO", oItemNo & " vendor item number changed to " & oVendorNo, oItemID)
            Return cmd.ExecuteNonQuery()
         Catch ex As Exception
            MsgBox(ex.Message & "  " & oItemNo, MsgBoxStyle.Exclamation, "Error Updating Vendor Item Number")
            Return 0
         End Try
      End Using
      Return tmpResult
   End Function
   Public Function UpdateDescription(ByVal aDescription As String) As Integer
      Dim tmpResult As Integer = 0
      oItemDesc = aDescription
      Dim QueryString As String = "UPDATE RECMDBUYITEMS SET RBI_DESCRIPTION = '" & oItemDesc.Replace("'", "''") & "' WHERE RBITEM_ID = " & oItemID
      Using conn As New SqlConnection(ITEMSDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         Try
            conn.Open()
            LogThis("NEWITMDESC", oItemNo & " item description changed to " & oItemDesc, oItemID)
            Return cmd.ExecuteNonQuery()
         Catch ex As Exception
            MsgBox(ex.Message & "  " & oItemNo, MsgBoxStyle.Exclamation, "Error Updating Item Description")
            Return 0
         End Try
      End Using
      Return tmpResult
   End Function
   Public Function UpdatePOQtyToBuy(ByVal NewQty As String, Optional ByVal inDB As Boolean = True) As Integer
      Dim tmpResult As Integer = 0
      Dim oldQty As Integer = oTotalToBuy
      oTotalToBuy = NewQty
      If Not inDB Then Return 0
      Dim QueryString As String = "UPDATE RECMDBUYITEMS SET RBI_TOTAL_TO_BUY = " & oTotalToBuy & " WHERE RBITEM_ID = " & oItemID
      Using conn As New SqlConnection(ITEMSDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         Try
            conn.Open()
            Dim result As Integer = cmd.ExecuteNonQuery()
            LogThis("NEWITMQTY", oItemNo & " qty to buy went from " & oldQty & " to " & oTotalToBuy, oItemID)
            Return result
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error Updating Vendor Cost")
            Return 0
         End Try
      End Using
      Return tmpResult
   End Function
   Public Function UpdateVendorCost(ByVal NewCost As String, Optional ByVal InDB As Boolean = True) As Integer
      Dim tmpResult As Integer = 0
      Dim oldPrice As Decimal = oVendorPrice
      oVendorPrice = NewCost
      If Not InDB Then
         Debug.Print("in db is false item:" & oItemNo)
         Return 0
      End If
      Dim QueryString As String = "UPDATE RECMDBUYITEMS SET RBI_VENDOR_PRICE = '" & oVendorPrice & "' WHERE RBITEM_ID = " & oItemID
      Using conn As New SqlConnection(ITEMSDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         Try
            conn.Open()
            Dim i As Integer = 0
            i = cmd.ExecuteNonQuery()
            Debug.Print("Update Vendor Cost Records affected: " & i & " item:" & oItemNo)
            LogThis("NEWITMCOST", oItemNo & " cost went from " & oldPrice & " to " & oVendorPrice, oItemID)
            Return i
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error Updating Vendor Cost")
            Return 0
         End Try
      End Using
      Return tmpResult
   End Function
   Public Function UpdateVendor(ByVal NewVendor As String, ByVal ToShortList As Boolean, Optional ByVal TempPONumber As String = "", Optional ByVal TempPODate As String = "") As Integer
      Dim tmpResult As Integer = 0
      Dim tmpOldVendor As String = oVendor
      If ToShortList Then
         oOriginalVendor = oVendor
      Else
         oOriginalVendor = ""
      End If
      oVendor = NewVendor
      oRecDate = ""
      oTempPONumber = TempPONumber
      oTempPODate = TempPODate
      Dim tmpDate As DateTime = Nothing
      If oTempPONumber.Length > 7 Then
         If IsDate(TempPODate) Then
            oTempPODate = TempPODate
            tmpDate = Date.Parse(oTempPODate)
         End If
      Else
         oTempPONumber = ""
      End If
      Dim QueryString As String = "UPDATE RECMDBUYITEMS SET RBI_VENDOR=@RBI_VENDOR,RBI_ORIGINAL_VENDOR=@RBI_ORIGINAL_VENDOR,RBI_REQDATE=@RBI_REQDATE,RBI_TMP_PO_NUMBER=@RBI_TMP_PO_NUMBER,RBI_TMP_PO_DATE=@RBI_TMP_PO_DATE,RBI_LAST_EMAIL_DATE=@RBI_LAST_EMAIL_DATE WHERE RBITEM_ID =@RBITEM_ID"
      Using conn As New SqlConnection(ITEMSDB)
         Dim cmd As New SqlCommand(QueryString, conn)
         With cmd.Parameters
            .Add("@RBI_VENDOR", SqlDbType.Char, 10).Value = oVendor
            .Add("@RBI_ORIGINAL_VENDOR", SqlDbType.Char, 10).Value = oOriginalVendor
            .Add("@RBI_REQDATE", SqlDbType.Char, 8).Value = ""
            .Add("@RBI_TMP_PO_NUMBER", SqlDbType.Char, 10).Value = oTempPONumber
            If oTempPODate.Length < 7 Then
               .Add("@RBI_TMP_PO_DATE", SqlDbType.DateTime).Value = DBNull.Value
            Else
               .Add("@RBI_TMP_PO_DATE", SqlDbType.DateTime).Value = tmpDate.ToString("yyyy-MM-ddTHH:mm:ss")
            End If
            .Add("@RBITEM_ID", SqlDbType.BigInt).Value = oItemID
            .Add("@RBI_LAST_EMAIL_DATE", SqlDbType.DateTime).Value = DBNull.Value
         End With
         Try
            conn.Open()
            Debug.Print(oItemNo & " is now under vendor " & oVendor)
            Dim result As Integer = cmd.ExecuteNonQuery()
            LogThis("ITMVENDORMV", oItemNo & " changed vendor from " & tmpOldVendor & " to " & oVendor, oItemID)
            Return result
         Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error Updating Item Vendor")
            Return 0
         End Try
      End Using
      Return tmpResult
   End Function
   Public Function CompareTo(ByVal other As RecommendedBuyItem) As Integer Implements System.IComparable(Of RecommendedBuyItem).CompareTo
      Return New CaseInsensitiveComparer().Compare(ITEMID, other.ITEMID)
   End Function
   Public Shared Function FindPredicate(ByVal Item As RecommendedBuyItem) As Predicate(Of RecommendedBuyItem)
      Return Function(Item2 As RecommendedBuyItem) Item.ITEMID = Item.ITEMID
   End Function
   Public Shared Function FindPredicateByItemId(ByVal Item As String) As Predicate(Of RecommendedBuyItem)
      Return Function(Item2 As RecommendedBuyItem) Item = Item2.ITEMID
   End Function
   Public Shared Function FindPredicateByEDPNO(ByVal Item As String) As Predicate(Of RecommendedBuyItem)
      Return Function(Item2 As RecommendedBuyItem) Item = Item2.EDPNO
   End Function
End Class
Public Class ItemVendor
#Region "Class Private Variables"
   Private oNumber As String = ""
   Private oAPVendor As String = ""
   Private oName As String = ""
   Private oContactFirstName As String = ""
   Private oContactLastName As String = ""
   Private oInitial As String = ""
   Private oRef1 As String = ""
   Private oRef2 As String = ""
   Private oStreet As String = ""
   Private oCity As String = ""
   Private oFOBCITY As String = ""
   Private oState As String = ""
   Private oZip As String = ""
   Private oDayPhone As String = ""
   Private oFaxNumber As String = ""
   Private oTermDays As Integer = 0
   Private oTemsPercent As String = ""
   Private oStandardDays As Integer = 0
   Private oComments1 As String = ""
   Private oComments2 As String = ""
   Private oComments3 As String = ""
   Private oComments4 As String = ""
   Private oComments5 As String = ""
   Private oComments6 As String = ""
   Private oComments7 As String = ""
   Private oComments8 As String = ""
   Private oEmail As String = ""
   Private oAllComments As New List(Of String)
   Private oRegularComments As New List(Of String)
   Private oPOComments As New List(Of String)
   Private oCommentIDString As String = ""
#End Region
   Sub New(ByVal r As SqlDataReader)
      SetVendorFromDataReader(r)
   End Sub
   Sub New(ByVal aUser As CurrentUser)
      oNumber = aUser.ID
      oName = "Short List"
      oEmail = aUser.Email
   End Sub
   Sub New(ByVal VendorNumber As String)
      If VendorNumber = User.ID Then
         oNumber = User.ID
         oName = "Short List"
         oEmail = User.Email
      Else
         Using conn As New SqlConnection(SELECTVENDORSDB)
            Dim QueryString As String = "SELECT * FROM VENDORDFROM  WHERE VENDORNO = '" & VendorNumber & "'"
            Dim cmd As New SqlCommand(QueryString, conn)
            Try
               conn.Open()
               Dim r As SqlDataReader = cmd.ExecuteReader
               If r.HasRows Then
                  r.Read()
                  SetVendorFromDataReader(r)
               Else
                  oNumber = VendorNumber
               End If
            Catch ex As Exception
               MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error getting new vendor information")
            End Try
         End Using
      End If
   End Sub
   Private Sub SetVendorFromDataReader(ByVal r As SqlDataReader)
      oNumber = Trim(r.Item("VENDORNO"))
      oAPVendor = Trim(r.Item("APVENDOR"))
      oName = Trim(r.Item("NAMEX"))
      oContactFirstName = Trim(r.Item("FIRSTNAME"))
      oContactLastName = Trim(r.Item("LASTNAME"))
      oInitial = Trim(r.Item("INITIALX"))
      oRef1 = Trim(r.Item("REF1"))
      oRef2 = Trim(r.Item("REF2"))
      oStreet = Trim(r.Item("STREET"))
      oCity = Trim(r.Item("CITY"))
      oFOBCITY = Trim(r.Item("FOBCITY"))
      oState = Trim(r.Item("STATE"))
      oZip = Trim(r.Item("ZIP"))
      oDayPhone = Trim(r.Item("DAYPHONE"))
      oFaxNumber = Trim(r.Item("FAXNO"))
      oTermDays = r.Item("TERMSDAYS")
      oTemsPercent = r.Item("TERMSPCT")
      oStandardDays = r.Item("STANDARDDAYS")
      oComments1 = r.Item("VENDORCOMMENTS_001")
      oComments2 = r.Item("VENDORCOMMENTS_002")
      oComments3 = r.Item("VENDORCOMMENTS_003")
      oComments4 = r.Item("VENDORCOMMENTS_004")
      oComments5 = r.Item("VENDORCOMMENTS_005")
      oComments6 = r.Item("VENDORCOMMENTS_006")
      oComments7 = r.Item("VENDORCOMMENTS_007")
      oComments8 = r.Item("VENDORCOMMENTS_008")
      With oAllComments
         .Add(oComments1)
         .Add(oComments2)
         .Add(oComments3)
         .Add(oComments4)
         .Add(oComments5)
         .Add(oComments6)
         .Add(oComments7)
         .Add(oComments8)
      End With
      oCommentIDString = r.Item("MISCDATA40")
      oEmail = Trim(r.Item("EMAIL"))
      SetComments()
   End Sub
   Private Sub SetComments()
      Dim s As String = oCommentIDString.Substring(27, 8)
      For i As Integer = 0 To 7
         If s.Chars(i) = "X" Then
            oPOComments.Add(oAllComments.Item(i))
         Else
            oRegularComments.Add(oAllComments.Item(i))
         End If
      Next
   End Sub
   ReadOnly Property Number() As String
      Get
         Return oNumber
      End Get
   End Property
   ReadOnly Property Email() As String
      Get
         If isEmail(oEmail) Then
            Return oEmail
         Else
            Return ""
         End If
      End Get
   End Property
   ReadOnly Property APVENDOR() As String
      Get
         Return oAPVendor
      End Get
   End Property
   ReadOnly Property TERMSPCT() As String
      Get
         Return oTemsPercent
      End Get
   End Property
   ReadOnly Property TERMSDAYS() As Integer
      Get
         Return oTermDays
      End Get
   End Property
   ReadOnly Property FOBCITY() As String
      Get
         Return oFOBCITY
      End Get
   End Property
   ReadOnly Property ADDITIONALDATA() As String
      Get
         If oStandardDays < 100 And oStandardDays > 9 Then
            Return "  0" & oStandardDays & "".PadRight(95, " ")
         ElseIf oStandardDays < 10 Then
            Return "  00" & oStandardDays & "".PadRight(95, " ")
         Else
            Return "  " & oStandardDays & "".PadRight(95, " ")
         End If
      End Get
   End Property
   ReadOnly Property FullName() As String
      Get
         Return oName
      End Get
   End Property
   ReadOnly Property ContactFullName() As String
      Get
         Return Trim(oContactFirstName & " " & oContactLastName)
      End Get
   End Property
   ReadOnly Property Address() As String
      Get
         Dim tmpresult As String = ""
         Dim ref1 As String = ""
         If oRef1 <> "" Then
            If Not isEmail(oRef1) Then
               ref1 = oRef1
            End If
         End If
         Dim ref2 As String = ""
         If oRef2 <> "" Then
            If Not isEmail(oRef2) Then
               ref2 = oRef2
            End If
         End If
         tmpresult = oStreet & vbCrLf
         If ref1 <> "" Then
            tmpresult &= ref1 & vbCrLf
         End If
         If ref2 <> "" Then
            tmpresult &= ref2 & vbCrLf
         End If
         tmpresult &= oCity & ", " & oState & " " & oZip
         Return tmpresult
      End Get
   End Property
   ReadOnly Property ContactNameAndAddress As String
      Get
         Dim tmpResult As String = ""
         tmpResult &= FullName & vbCrLf
         tmpResult &= Address & vbCrLf
         If ContactFullName <> "" Then
            tmpResult &= ContactFullName & vbCrLf
         End If
         If Phone <> "" Then
            tmpResult &= Phone & vbCrLf
         End If
         tmpResult &= "Email: " & Email & Environment.NewLine
         If User.ID = oNumber Then Return tmpResult
         tmpResult &= "Standard Days: " & CInt(oStandardDays) & " Term Days: " & CInt(oTermDays) & " Terms %: " & CInt(oTemsPercent) / 100
         Return tmpResult
      End Get
   End Property
   ReadOnly Property ContactNameAndAddressHTML As String
      Get
         Dim ref1 As String = ""
         If oRef1 <> "" Then
            If Not isEmail(oRef1) Then
               ref1 = oRef1
            End If
         End If
         Dim ref2 As String = ""
         If oRef2 <> "" Then
            If Not isEmail(oRef2) Then
               ref2 = oRef2
            End If
         End If
         Dim tmpResult As String = ""
         tmpResult &= FullName & "<br>"
         tmpResult &= oStreet & "<br>"
         If ref1 <> "" Then
            tmpResult &= ref1 & "<br>"
         End If
         If ref2 <> "" Then
            tmpResult &= ref2 & "<br>"
         End If
         tmpResult &= oCity & ", " & oState & " " & oZip & "<br>"
         If ContactFullName <> "" Then
            tmpResult &= ContactFullName & "<br>"
         End If
         If Phone <> "" Then
            tmpResult &= Phone & "<br>"
         End If
         If Email <> "" Then
            tmpResult &= Email & "<br>"
         End If
         If User.ID = oNumber Then Return tmpResult
         tmpResult &= "Standard Days: " & CInt(oStandardDays) & " Term Days: " & CInt(oTermDays) & " Terms Percent: " & CInt(oTemsPercent) / 100 & "<br>"
         Return tmpResult
      End Get
   End Property
   ReadOnly Property VendorEmails() As List(Of String)
      Get
         Return GetAllEMailAddresses(oEmail & " " & oRef2 & " " & oRef1 & " " & oComments1 & " " & oComments2 & " " & oComments3 & " " & oComments4 & " " & oComments5 & " " & oComments6 & " " & oComments7 & " " & oComments8)
      End Get
   End Property
   ReadOnly Property Phone() As String
      Get
         If oDayPhone <> "" Then
            Dim s As String = Regex.Replace(oDayPhone, "[^0-9]", "")
            If s.Length = 10 Then
               Return String.Format("{0:(###) ###-####}", Long.Parse(s))
            ElseIf s.Length = 11 And s.Chars(0) = "1" Then
               Return String.Format("{0:(###) ###-####}", Long.Parse(s.Substring(1, 10)))
            Else
               Return ""
            End If
         Else
            Return ""
         End If
      End Get
   End Property
   ReadOnly Property Comments() As String
      Get
         Return "Vendor Comments: " & vbCrLf & VendorComments & vbCrLf & vbCrLf & "PO Comments: " & vbCrLf & POComments
      End Get
   End Property
   ReadOnly Property POComments() As String
      Get
         Dim tmpresult As String = ""
         For Each s As String In oPOComments
            If s.Trim <> "" Then
               tmpresult &= s.Trim & Environment.NewLine
            End If
         Next
         Return tmpresult
      End Get
   End Property
   ReadOnly Property POCommentsHTML() As String
      Get
         Dim tmpresult As String = ""
         For Each s As String In oPOComments
            If s.Trim <> "" Then
               tmpresult &= s.Trim & "<br>"
            End If
         Next
         Return tmpresult
      End Get
   End Property
   ReadOnly Property VendorComments() As String
      Get
         Dim tmpresult As String = ""
         For Each s As String In oRegularComments
            If s.Trim <> "" Then
               tmpresult &= s.Trim & vbCrLf
            End If
         Next
         Return tmpresult
      End Get
   End Property
   Private Function isEmail(ByVal aString As String) As Boolean
      If GetAllEMailAddresses(aString).Count > 0 Then
         Return True
      Else
         Return False
      End If
   End Function
   Private Function GetAllEMailAddresses(ByVal Input As String) As List(Of String)
      Dim Results As New List(Of String)
      Dim MC As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Input, "\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*")
      For i As Integer = 0 To MC.Count - 1
         If Results.Contains(MC(i).Value) = False Then
            Results.Add(MC(i).Value)
         End If
      Next
      Return Results
   End Function
End Class
Public Class CurrentUser
#Region "Class Private Variables"
   Private oID As String = ""
   Private oFullName As String = ""
   Private oPOBUYERCD As String = ""
   Private oemail As String
#End Region
   Sub New(ByVal aUser As String)
      oID = aUser
      SetApplicationUser()
   End Sub
   Private Sub SetApplicationUser()
      Select Case UCase(oID)
         Case "BRIAN"
            oID = "BRIAN"
            oPOBUYERCD = "05"
            oFullName = "Brian Cleveland"
                oemail = "brian@ecommerce.com"
            Case "SCOTTM"
            oID = "SCOTTM"
            oPOBUYERCD = "04"
            oFullName = "Scott McLean"
                oemail = "scottm@ecommerce.com"
            Case "PAUL"
            oID = "PAUL"
            oPOBUYERCD = "06"
            oFullName = "Paul Cusick"
                oemail = "paul@ecommerce.com"
            Case "CATHY"
            oID = "CATHY"
            oPOBUYERCD = "18"
            oFullName = "Cathy Brown"
                oemail = "cathy@ecommerce.com"
            Case "EDLYN"
            oID = "EDLYN"
            oPOBUYERCD = "12"
            oFullName = "Edlyn Chavez"
                oemail = "edlyn@ecommerce.com"
            Case "TRACEY"
            oID = "TRACEY"
            oPOBUYERCD = "14"
            oFullName = "Tracey Bueno"
                oemail = "tracey@ecommerce.com"
        End Select
   End Sub
   ReadOnly Property ID() As String
      Get
         Return oID
      End Get
   End Property
   ReadOnly Property FullName() As String
      Get
         Return oFullName
      End Get
   End Property
   ReadOnly Property FullNameUpper() As String
      Get
         Return oFullName.ToUpper
      End Get
   End Property
   ReadOnly Property POBUYERCD() As String
      Get
         Return oPOBUYERCD.PadRight(4, " ")
      End Get
   End Property
   ReadOnly Property Email() As String
      Get
         Return oemail
      End Get
   End Property
   ReadOnly Property MAILADDRESS() As MailAddress
      Get
         Return New MailAddress(oemail, FullName)
      End Get
   End Property
   ReadOnly Property Vendors() As List(Of String)
      Get
         Dim tmpResult As New List(Of String)
         Dim QueryString As String = ""
         If oID = "PAUL" Then
            Dim tmpVendors As New List(Of String)
            QueryString = "SELECT RBV_VENDOR_NUMBER FROM RECMDBUYVENDORS"
            Using conn As New SqlConnection(ITEMSDB)
               Dim cmd As New SqlCommand(QueryString, conn)
               Try
                  conn.Open()
                  Dim r As SqlDataReader = cmd.ExecuteReader
                  If r.HasRows Then
                     While r.Read
                        tmpVendors.Add(Trim(r.Item("RBV_VENDOR_NUMBER")).ToUpper)
                     End While
                  End If
               Catch ex As Exception
                  MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error getting assigned vendors")
               End Try
            End Using
            Using conn As New SqlConnection(SELECTVENDORSDB)
               QueryString = "SELECT VENDORNO FROM VENDORDFROM"
               Dim cmd As New SqlCommand(QueryString, conn)
               Try
                  conn.Open()
                  Dim r As SqlDataReader = cmd.ExecuteReader
                  If r.HasRows Then
                     While r.Read
                        Dim v As String = Trim(r.Item("VENDORNO")).ToUpper
                        If Not tmpVendors.Contains(v) Then
                           tmpResult.Add(v)
                        End If
                     End While
                  End If
               Catch ex As Exception
                  MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error reading VENDORDFROM table")
               End Try
            End Using
            tmpResult.Sort()
            tmpResult.Add("PAUL")
         Else
            QueryString = "SELECT RBV_VENDOR_NUMBER FROM RECMDBUYVENDORS WHERE RBV_VENDOR_BUYER = '" & oID & "'"
            Using conn As New SqlConnection(ITEMSDB)
               Dim cmd As New SqlCommand(QueryString, conn)
               Try
                  conn.Open()
                  Dim r As SqlDataReader = cmd.ExecuteReader
                  If r.HasRows Then
                     While r.Read
                        Dim v As String = Trim(r.Item("RBV_VENDOR_NUMBER")).ToUpper
                        If Not v = oID Then
                           tmpResult.Add(Trim(r.Item("RBV_VENDOR_NUMBER")).ToUpper)
                        End If
                     End While
                  End If
               Catch ex As Exception
                  MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error user getting assigned vendors")
               End Try
            End Using
            tmpResult.Sort()
            tmpResult.Add(oID)
         End If
         Return tmpResult
      End Get
   End Property
End Class