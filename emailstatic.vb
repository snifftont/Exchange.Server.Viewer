Imports Microsoft.VisualBasic
Imports System
Imports System.Threading
Imports System.Net.Mail
Imports System.Data.SqlClient
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Xml.XPath

Public Class emailstatic



    'Private Delegate Sub voidFunc(Of P1, P3)(ByVal p1 As P1, ByVal p3 As P3)
    'Shared Sub StartTimer(ByVal startTime As DateTime, ByVal action As Boolean)
    '    Dim Timer As voidFunc(Of DateTime, Boolean)
    '    Timer = AddressOf TimedThread
    '    Timer.BeginInvoke(startTime, action, Nothing, Nothing)
    'End Sub
  
    Shared Function entry() As Boolean
        Try

        
        Dim i As Integer
            For i = 0 To get_user_stt().Tables(0).Rows.Count - 1
                If (get_email(get_user_stt().Tables(0).Rows(i)("username").ToString()).Tables(0).Rows(0)("body").ToString() <> "0") Then
                    ' SendMail(ConfigurationManager.AppSettings("EmailFrom").ToString(), get_email(get_user_stt().Tables(0).Rows(i)("username").ToString()).Tables(0).Rows(0)("email").ToString(), "Daily Inbox Mesage", get_email(get_user_stt().Tables(0).Rows(i)("username").ToString()).Tables(0).Rows(0)("body").ToString())

                    SendMessage(ConfigurationManager.AppSettings("EmailFrom").ToString(), get_email(get_user_stt().Tables(0).Rows(i)("username").ToString()).Tables(0).Rows(0)("email").ToString(), "Daily Inbox Mesage", get_email(get_user_stt().Tables(0).Rows(i)("username").ToString()).Tables(0).Rows(0)("body").ToString())

                End If
                delete_daily_email()
            Next
        Catch ex As Exception

        End Try
        Return True
    End Function
    Shared Function SendMail(ByVal [From] As String, ByVal [To] As String, ByVal Subject As String, ByVal Body As String) As Boolean
        Dim r As Boolean = False
        Try
            Dim mailClient As New SmtpClient()
            'NetworkCredential basicAuthenticationInfo = new NetworkCredential();
            Dim gridHtml As String = "<!DOCTYPE html PUBLIC" + """"c + "-//W3C//DTD XHTML 1.0 Transitional//EN" + """"c + """"c + "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd" + """"c + ">"
            gridHtml += "<html xmlns=" + """"c + "http://www.w3.org/1999/xhtml" + """"c + ">" + "<head></head><body>"
            gridHtml += Body
            gridHtml += "</body></html>"
            Dim mailmessage As New System.Net.Mail.MailMessage([From], [To], Subject, gridHtml)
            'mailmessage.CC = "jaswinder.s@nucleon-tech.com"
            mailmessage.IsBodyHtml = True
            mailClient.Send(mailmessage)
            r = True
        Catch ex As System.Net.Mail.SmtpException
            r = False
            Throw
        End Try
        Return r
    End Function
    Shared Function get_user_stt() As DataSet
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("exchangeconnection").ToString())
        conn.Open()
        Dim cmd As New SqlCommand("dbo.ev_get_all_active_user", conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Dim ds As New DataSet()
        Dim ad As New SqlDataAdapter(cmd)
        ad.Fill(ds)
        conn.Close()
        Return ds
    End Function
    Shared Function get_email(ByVal username As String) As DataSet
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("exchangeconnection").ToString())
        conn.Open()
        Dim cmd As New SqlCommand("dbo.ev_daily_email", conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@username", username)
        cmd.Parameters.AddWithValue("@root", ConfigurationManager.AppSettings("root_url").ToString() + "/deliver.aspx?msgid=")
        cmd.ExecuteNonQuery()
        Dim ds As New DataSet()
        Dim ad As New SqlDataAdapter(cmd)
        ad.Fill(ds)
        conn.Close()
        Return ds
    End Function
    Shared Function email_at() As DataSet
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("exchangeconnection").ToString())
        conn.Open()
        Dim cmd As New SqlCommand("dbo.ev_mail_at", conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Dim ds As New DataSet()
        Dim ad As New SqlDataAdapter(cmd)
        ad.Fill(ds)
        conn.Close()
        Return ds
    End Function
    Shared Function save_mail_at(ByVal time_at As String) As Boolean
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("exchangeconnection").ToString())
        conn.Open()
        Dim cmd1 As New SqlCommand("dbo.ev_update_mail_at", conn)
        cmd1.CommandType = CommandType.StoredProcedure
        cmd1.Parameters.AddWithValue("@send_mail_at", time_at)
        cmd1.ExecuteNonQuery()
        conn.Close()
        Return True
    End Function
    Shared Function delete_daily_email() As Boolean
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("exchangeconnection").ToString())
        conn.Open()
        Dim cmd1 As New SqlCommand("dbo.ev_deleteable_files", conn)
        cmd1.CommandType = CommandType.StoredProcedure
        cmd1.Parameters.AddWithValue("@delete_after", Convert.ToInt32(ConfigurationManager.AppSettings("delete_after")))
        cmd1.ExecuteNonQuery()
        Dim ds As DataSet = New DataSet()
        Dim ad As SqlDataAdapter = New SqlDataAdapter(cmd1)
        ad.Fill(ds)
        conn.Close()
        conn.Open()
        Dim cmd As New SqlCommand("dbo.ev_delete_daily_email", conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@delete_after", Convert.ToInt32(ConfigurationManager.AppSettings("delete_after")))
        cmd.ExecuteNonQuery()
        conn.Close()
        
        Dim ia As Integer = 0
        For ia = 0 To ds.Tables(0).Rows.Count - 1
            Dim source As String = ConfigurationManager.AppSettings("email_save_at") + "\" + ds.Tables(0).Rows(ia)(0).ToString() + ".eml"
            Dim file As New System.IO.FileInfo(source)
            If (file.Exists) Then
                file.Delete()
            End If
        Next
        Return True
    End Function

    'Shared Sub TimedThread(ByVal startTime As DateTime, ByVal action As Boolean)

    '    Dim keepRunning As Boolean = True
    '    Dim NextExecute As DateTime = startTime
    '    While keepRunning
    '        If DateTime.Now > NextExecute Then
    '            keepRunning = action
    '            'If Convert.ToDateTime(DateTime.Now.ToShortDateString() + " " + ConfigurationManager.AppSettings("send_mail_at").ToString()) > NextExecute Then
    '            '    NextExecute = Convert.ToDateTime(Convert.ToDateTime(DateTime.Now.ToShortDateString() + " " + ConfigurationManager.AppSettings("send_mail_at").ToString()) - NextExecute)
    '            'Else
    '            '  NextExecute = Convert.ToDateTime(DateTime.Now.ToShortDateString() + " " + ConfigurationManager.AppSettings("send_mail_at").ToString()).AddDays(1)
    '            NextExecute = NextExecute.AddMinutes(2)
    '        End If
    '        Thread.Sleep(1000)
    '    End While

    'End Sub

    Shared Function entry(ByVal files As String()) As Boolean
        ' Try
        ' Get the object used to communicate with the server.

        'Dim dirInfo As New DirectoryInfo(System.Configuration.ConfigurationSettings.AppSettings("spam_email_dir").ToString())
        'Dim fileList As FileInfo() = dirInfo.GetFiles()
        'Dim numberOfFiles As Integer = fileList.Length

        For Each file As String In files
            Dim request As FtpWebRequest = DirectCast(WebRequest.Create(ConfigurationManager.AppSettings("ftp_host").ToString() + file), FtpWebRequest)
            request.Method = WebRequestMethods.Ftp.DownloadFile
            ' This example assumes the FTP site uses anonymous logon.
            request.Credentials = New NetworkCredential(ConfigurationManager.AppSettings("ftp_username").ToString(), System.Configuration.ConfigurationManager.AppSettings("ftp_password").ToString())
            ' Copy the contents of the file to the request stream.
            Dim response As FtpWebResponse = DirectCast(request.GetResponse(), FtpWebResponse)
            Dim responseStream As Stream = response.GetResponseStream()
            Dim writestream As FileStream = New FileStream(ConfigurationManager.AppSettings("spam_email_dir").ToString() + file, FileMode.Create)
            Dim Length As Integer = 2048
            Dim buffer As [Byte]() = New [Byte](Length) {}
            Dim bytesRead As Integer = responseStream.Read(buffer, 0, Length)
            While bytesRead > 0
                writestream.Write(buffer, 0, bytesRead)
                bytesRead = responseStream.Read(buffer, 0, Length)
            End While
            writestream.Close()
            response.Close()
            ' fileList(i).Delete()
        Next
        '  Catch ex As Exception

        '  End Try
        Return True
    End Function
    Shared Function GetFileList() As String()
        Dim downloadFiles As String()
        Dim result As New StringBuilder()
        Dim response As WebResponse = Nothing
        Dim reader As StreamReader = Nothing
        Try
            Dim reqFTP As FtpWebRequest
            reqFTP = DirectCast(FtpWebRequest.Create(New Uri(System.Configuration.ConfigurationSettings.AppSettings("ftp_host").ToString())), FtpWebRequest)
            reqFTP.UseBinary = True
            reqFTP.Credentials = New NetworkCredential(System.Configuration.ConfigurationSettings.AppSettings("ftp_username").ToString(), System.Configuration.ConfigurationSettings.AppSettings("ftp_password").ToString())
            reqFTP.Method = WebRequestMethods.Ftp.ListDirectory
            reqFTP.Proxy = Nothing
            reqFTP.KeepAlive = False
            reqFTP.UsePassive = False
            response = reqFTP.GetResponse()
            reader = New StreamReader(response.GetResponseStream())
            Dim line As String = reader.ReadLine()
            While line <> Nothing
                result.Append(line)
                result.Append(vbLf)
                line = reader.ReadLine()
            End While
            ' to remove the trailing '\n'
            result.Remove(result.ToString().LastIndexOf(ControlChars.Lf), 1)
            Return result.ToString().Split(ControlChars.Lf)
        Catch ex As Exception
            If reader IsNot Nothing Then
                reader.Close()
            End If
            If response IsNot Nothing Then
                response.Close()
            End If
            downloadFiles = Nothing
            Return downloadFiles
        End Try
    End Function

    Shared Function SendMessage(ByVal strSendFrom As String, ByVal strSendTo As String, ByVal strSendSubject As String, ByVal strSendBody As String) As Boolean
        Dim p_strAlias As String = ConfigurationManager.AppSettings("daily_sender_alias").ToString()
        Dim p_strServer As String = ConfigurationManager.AppSettings("Exchange_server").ToString() + p_strAlias + "/"
        Dim p_strUserName As String = ConfigurationManager.AppSettings("EmailFrom").ToString()
        Dim p_strPassword As String = ConfigurationManager.AppSettings("daily_sender_password").ToString()
        Dim p_strInboxURL As String = ConfigurationManager.AppSettings("inbox_name").ToString()
       
        Dim reply As Boolean = False
        Dim strSubURI As String
        Dim strTempURI As String
        Dim strAlias As String
        Dim strUserName As String
        Dim strPassWord As String
        Dim strExchSvrName As String
        Dim strTo As String
        Dim strSubject As String
        Dim strBody As String
        Dim strText As String
        Dim bRequestSuccessful As Boolean

        ' To use MSXML 3.0, use the following Dim statements.
        Dim xmlReq As MSXML2.IXMLHTTPRequest
        Dim xmlReq2 As MSXML2.IXMLHTTPRequest

        ' To use MSXML 4.0, use the following Dim statements.
        ' Dim xmlReq As MSXML2.XMLHTTP40
        ' Dim xmlReq2 As MSXML2.XMLHTTP40



        ' Exchange server name.
        strExchSvrName = p_strServer

        ' Alias of the sender.
        strAlias = p_strAlias

        ' User name of the sender.
        strUserName = p_strUserName

        ' Password of the sender.
        strPassWord = p_strPassword

        ' E-mail address of the sender.
        strTo = strSendTo

        ' Subject of the mail.
        strSubject = strSendSubject

        ' Text body of the mail.
        strBody = HttpUtility.HtmlDecode(strSendBody)

        ' Build the submission URI. If Secure Sockets Layer (SSL)
        ' is set up on the server, use "https://" instead of "http://".
        strSubURI = strExchSvrName + "%23%23DavMailSubmissionURI%23%23/"

        ' Build the temporary URI. If SSL is set up on the
        ' server, use "https://" instead of "http://".
        strTempURI = strExchSvrName + "drafts/" & strSubject & ".eml"

        ' Construct the body of the PUT request.
        ' Note: If the From: header is included here,
        ' the MOVE method request will return a
        ' 403 (Forbidden) status. The From address will
        ' be generated by the Exchange server.
        strText = "To: " & strTo & vbNewLine & _
                  "Subject: " & strSubject & vbNewLine & _
                  "Date: " & Now & _
                  "X-Mailer: test mailer" & vbNewLine & _
                  "MIME-Version: 1.0" & vbNewLine & _
                  "Content-Type: text/html;" & vbNewLine & _
                  "Charset = ""iso-8859-1""" & vbNewLine & _
                  "Content-Transfer-Encoding: 7bit" & vbNewLine & _
                  vbNewLine & strBody

        ' Initialize.
        bRequestSuccessful = False

        ' To use MSXML 3.0, use the following Set statement.
        xmlReq = CreateObject("Microsoft.XMLHTTP")

        ' To use MSXML 4.0, use the following Set statement.
        ' Set xmlReq = CreateObject("Msxml2.XMLHTTP.4.0")

        ' Open the request object with the PUT method and
        ' specify that it will be sent asynchronously. The
        ' message will be saved to the drafts folder of the
        ' specified user's inbox.
        xmlReq.open("PUT", strTempURI, False, strUserName, strPassWord)

        ' Set the Content-Type header to the RFC 822 message format.
        xmlReq.setRequestHeader("Content-Type", "message/rfc822")

        ' Send the request with the message body.
        xmlReq.send(strText)

        ' The PUT request was successful.












        If (xmlReq.status >= 200 And xmlReq.status < 300) Then
            bRequestSuccessful = True
            reply = True
            If bRequestSuccessful Then

                ' To use MSXML 3.0, use the following Set statement.
                xmlReq2 = CreateObject("Microsoft.XMLHTTP")

                ' To use MSXML 4.0, use the following Set statement.
                ' Set xmlReq2 = CreateObject("Msxml2.XMLHTTP.4.0")

                ' Open the request object with the PUT method and
                ' specify that it will be sent asynchronously.
                xmlReq2.open("MOVE", strTempURI, False, strUserName, strPassWord)
                xmlReq2.setRequestHeader("Destination", strSubURI)
                xmlReq2.send()

                ' The MOVE request was successful.
                If (xmlReq2.status >= 200 And xmlReq2.status < 300) Then
                    reply = True

                    ' An error occurred on the server.
                ElseIf (xmlReq.status = 500) Then
                    reply = False


                Else
                    reply = False

                End If

            End If
            ' An error occurred on the server.
        ElseIf (xmlReq.status = 500) Then
            reply = False


        Else
            reply = False

        End If

        ' If the PUT request was successful,
        ' MOVE the message to the mailsubmission URI.


        ' Clean up.
        xmlReq = Nothing
        xmlReq2 = Nothing


        Return reply
    End Function

   

End Class


