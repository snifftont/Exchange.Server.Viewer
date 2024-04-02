Public Class deliver
    Inherits System.Web.UI.Page
    Dim m As New Mails()
    Dim em As New email_class()
    Dim acc As New Account_DAL()
    Dim ereader As New EMLReader()
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        
        If (Request.QueryString("msgid") IsNot Nothing) Then




            Dim source As String = ConfigurationManager.AppSettings("email_save_at") + "\" + Request.QueryString("msgid").ToString() + ".eml"

            Try

                Dim iDropDir As New CDO.DropDirectory()

                Dim iMsgs As CDO.IMessages
                Dim iMsgReply As CDO.IMessage
                Dim iMsgReplyAll As CDO.IMessage
                Dim iMsgForward As CDO.IMessage
                ' Get the messages from the Drop directory.
                iMsgs = iDropDir.GetMessages(ConfigurationManager.AppSettings("email_save_at").ToString())
                '  Console.WriteLine("Messages Count : " + iMsgs.Count.ToString())
                For Each iMsg As CDO.IMessage In iMsgs
                    Try
                        ' If (iMsg.Fields("urn:schemas:mailheader:message-id").Value = em.get_single_email_received(Convert.ToInt32(Request.QueryString("msgid").ToString())).Tables(0).Rows(0)("msg_id").ToString()) Then
                        If (iMsgs.FileName(iMsg) = source) Then
                            Dim subject As String = iMsg.Subject.Substring(iMsg.Subject.IndexOf("-") + 2)
                            Dim tomail As String = iMsg.Subject.Substring(0, iMsg.Subject.IndexOf(" - "))
                            Dim username As String = tomail.Substring(0, tomail.IndexOf("@"))
                            SetMailClass(username)
                            If (m.SendMessage(tomail, tomail, subject, iMsg.HTMLBody) = True) Then
                                'If (m.move_to_inbox(source, detination) = True) Then  iMsg.Fields("urn:schemas:httpmail:from").Value
                                em.Delete_email_received(Convert.ToInt32(Request.QueryString("msgid").ToString()))
                                Dim file As New System.IO.FileInfo(source)
                                If (file.Exists) Then
                                    file.Delete()
                                End If
                                lblMsg.Text = "Mail Delivered to your inbox successfully."
                            End If
                        End If
                    Catch ex As Exception

                    End Try
                Next
                ' Clean up memory.
                iMsgs = Nothing
                iMsgReply = Nothing
                iMsgReplyAll = Nothing
                iMsgForward = Nothing
            Catch ex1 As Exception
                lblMsg.Text = "Mail Delivery Fail."
            End Try
        End If

    End Sub
    Private Sub SetMailClass(ByVal username As String)
        m.p_strAlias = username
        m.p_strServer = ConfigurationManager.AppSettings("Exchange_server").ToString() + m.p_strAlias + "/"
        m.p_strUserName = username
        m.p_strPassword = acc.get_user(username).Tables(0).Rows(0)("password").ToString()
        m.p_strInboxURL = ConfigurationManager.AppSettings("inbox_name").ToString()
        m.p_strSpam = ConfigurationManager.AppSettings("spam_name").ToString()
        m.p_strSpamServer = ConfigurationManager.AppSettings("Spam_Exchange_server").ToString() + ConfigurationManager.AppSettings("spam_alias").ToString() + "/"
        m.p_strSpamPwd = ConfigurationManager.AppSettings("spam_pwd").ToString()
        m.p_strSpamUserName = ConfigurationManager.AppSettings("spam_username").ToString()
    End Sub
End Class