Imports System.IO
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports Microsoft.Office.Interop
Module Module1
    Sub Main()
        Dim s As String = GetDupMattersReport()
        Dim Message As New MailMessage
        With Message
            .From = New MailAddress("administrator@lawfirm.com")
            .To.Add("fcanton@lawfirm.com")
            .To.Add("records@lawfirm.com")
            .Attachments.Add(New Attachment(s))
            .Subject = "Duplicate DMS Matters Report"
            .Body = "Report Attached"
        End With
        Dim SMTPClient As New SmtpClient("SMTP1")
        Try
            SMTPClient.Send(Message)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
    Function GetDupMattersReport() As String
        Dim rptPath As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\duplicated.matters." & Now.ToString("yyyyMMdd") & ".xlsx"
        Dim aExcelApp As New Excel.Application
        Dim aExcelWrkbook As Excel.Workbook = aExcelApp.Workbooks.Add
        Dim aExcelWrkSheet As Excel.Worksheet = aExcelWrkbook.Worksheets.Add
        With aExcelWrkSheet
            .Name = "Duplicated Matters"
            .Range("A1").Value = "Client"
            .Range("A1").ColumnWidth = 15
            .Range("A1").NumberFormat = "@"
            .Range("B1").Value = "Matter"
            .Range("B1").ColumnWidth = 15
            .Range("B1").NumberFormat = "@"
            .Range("A1:B1").Font.Bold = True
            Dim connString1 As String = "Data Source=sql;Initial Catalog=iManage_Active;Integrated Security=SSPI"
            Dim queryString As String = "declare @matters table (client varchar(32), matter varchar(32)) " &
                                        "insert into @matters " &
                                        "select distinct C1ALIAS,C2ALIAS from MHGROUP.DOCMASTER " &
                                        "declare @mcounts table (mcount NUMERIC(18,0), mmatter varchar(32)) " &
                                        "insert into @mcounts " &
                                        "select COUNT(matter), matter from @matters group by matter having COUNT(matter) > 1 " &
                                        "select * from @matters where matter in (select mmatter from @mcounts) order by matter"
            Using conn As New SqlConnection(connString1)
                Dim cmd As New SqlCommand(queryString, conn)
                conn.Open()
                Dim r As SqlDataReader = cmd.ExecuteReader()
                If r.HasRows Then
                    Dim c As Integer = 1
                    Try
                        While r.Read
                            c += 1
                            Dim tClient As String = r("client")
                            Dim tMatter As String = r("matter")
                            .Range("A" & c).NumberFormat = "@"
                            .Range("A" & c).Value = tClient
                            .Range("B" & c).NumberFormat = "@"
                            .Range("B" & c).Value = tMatter
                        End While
                    Catch ex As Exception
                    End Try
                End If
            End Using
        End With
        aExcelWrkSheet.Range("A2").Select()
        aExcelApp.ActiveWindow.SplitColumn = 0
        aExcelApp.ActiveWindow.SplitRow = 1
        aExcelApp.ActiveWindow.FreezePanes = True
        aExcelWrkbook.Sheets.Item("Sheet1").delete()
        aExcelWrkbook.Sheets.Item("Sheet2").delete()
        aExcelWrkbook.Sheets.Item("Sheet3").delete()
        Dim tmpFilePath As String = rptPath
        aExcelWrkbook.SaveAs(tmpFilePath)
        aExcelWrkbook.Close()
        aExcelApp.Quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(aExcelWrkbook)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(aExcelApp)
        aExcelWrkbook = Nothing
        aExcelApp = Nothing
        GC.Collect()
        Return rptPath
    End Function
End Module
