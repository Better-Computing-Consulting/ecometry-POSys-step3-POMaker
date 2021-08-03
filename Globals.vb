Imports System.Data.SqlClient
Module Globals
    Public password As String = "dummypw"
    Public User As CurrentUser
   Public UserVendorItemsList As List(Of VendorItems)

   Public RootDocFolder As String = "\\usa2\Purchasing\Finished_POs\"
    'Public RootDocFolder As String = "\\usa2\Projects\Purchasing\POTESTING\"

    Public ITEMSDB As String = "Data Source=ECOMDB3;Initial Catalog=RECBUYS;UID=ss;PWD=sss"
    'Public ITEMSDB As String = "Data Source=ECOMDB3;Initial Catalog=RECBYSDEV;UID=ss;PWD=sss"

    'Public SELECTVENDORSDB As String = "Data Source=ecom-db2;Initial Catalog=ECOMVER;UID=ss;PWD=ssss"
    Public SELECTVENDORSDB As String = "Data Source=ecom-db1;Initial Catalog=ECOMLIVE;UID=ss;PWD=ssss"

    Public POENTRYDB As String = SELECTVENDORSDB
   Public Sub LogThis(ByVal ShortText As String, ByVal LongText As String, ByVal ItemID As Integer)
      Dim entryDateTime As DateTime = Now
      Dim logfile As String = My.Application.Info.DirectoryPath & "\" & My.Application.Info.ProductName & ".log"
      My.Computer.FileSystem.WriteAllText(logfile, "[" & entryDateTime.ToString("yyyy-MM-dd HH:mm:ss") & "]" & " " & String.Format("{0,-15}{1}" & Environment.NewLine, ShortText, LongText), True)
      Using conn As New SqlConnection(ITEMSDB)
         Dim QueryString As String =
            "INSERT INTO RECMDBUYITEMACTIONS(RBITEM_ID,RBIA_ACTION_SHORT_TEXT,RBIA_ACTION_LONG_TEXT,RBIA_ACTION_DATETIME)" & _
            "VALUES(@RBITEM_ID,@RBIA_ACTION_SHORT_TEXT,@RBIA_ACTION_LONG_TEXT,@RBIA_ACTION_DATETIME)"
         Dim cmd As New SqlCommand(QueryString, conn)
         With cmd.Parameters
            .Add("@RBITEM_ID", SqlDbType.BigInt).Value = ItemID
            .Add("@RBIA_ACTION_SHORT_TEXT", SqlDbType.Char, 15).Value = ShortText
            .Add("@RBIA_ACTION_LONG_TEXT", SqlDbType.Char, 80).Value = LongText
            .Add("@RBIA_ACTION_DATETIME", SqlDbType.DateTime).Value = entryDateTime.ToString("yyyy-MM-ddTHH:mm:ss")
         End With
         Try
            conn.Open()
            cmd.ExecuteNonQuery()
         Catch ex As Exception
            Debug.Print("Error Entering Item Action into DB: " & ex.Message)
         End Try
      End Using
   End Sub
End Module
