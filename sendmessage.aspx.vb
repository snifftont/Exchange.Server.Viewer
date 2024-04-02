Public Class sendmessage
    Inherits System.Web.UI.Page
    Dim m As New Mails()
    Dim em As New email_class()
    Dim acc As New Account_DAL()
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then

            If Request.QueryString("val") IsNot Nothing And Request.QueryString("id") IsNot Nothing Then
                If Request.QueryString("val").ToString() = "forward" Then
                    Dim path As String = ConfigurationManager.AppSettings("temp_mail") + "\" + Request.QueryString("id").ToString() + ".eml"
                    'lblContent.Text = ereader.get_email_body(path)
                    Try
                        Dim iDropDir As New CDO.DropDirectory()

                        Dim iMsgs As CDO.IMessages
                        Dim iMsgReply As CDO.IMessage
                        Dim iMsgReplyAll As CDO.IMessage
                        Dim iMsgForward As CDO.IMessage
                        ' Get the messages from the Drop directory.
                        iMsgs = iDropDir.GetMessages(ConfigurationManager.AppSettings("temp_mail").ToString())
                        '  Console.WriteLine("Messages Count : " + iMsgs.Count.ToString())
                        For Each iMsg As CDO.IMessage In iMsgs
                            If (iMsgs.FileName(iMsg) = path) Then
                                'lblContent.Text = iMsg.TextBody
                                If iMsg.HTMLBody <> Nothing AndAlso Not String.IsNullOrWhiteSpace(iMsg.HTMLBody) Then
                                    txtBody.Text = HttpUtility.HtmlDecode(iMsg.HTMLBody)
                                End If

                                txtSubject.Text = "Fwd: " + iMsg.Subject
                                txtFrom.Text = acc.get_user(User.Identity.Name).Tables(0).Rows(0)("email").ToString()
                            End If
                        Next
                        ' Clean up memory.
                        iMsgs = Nothing
                        iMsgReply = Nothing
                        iMsgReplyAll = Nothing
                        iMsgForward = Nothing
                    Catch ex1 As Exception
                        'Console.WriteLine("{0} Exception caught.", e)
                    End Try
                Else
                    If Request.QueryString("val").ToString() = "reply" Then
                        Dim path As String = ConfigurationManager.AppSettings("temp_mail") + "\" + Request.QueryString("id").ToString() + ".eml"
                        'lblContent.Text = ereader.get_email_body(path)
                        Try
                            Dim iDropDir As New CDO.DropDirectory()

                            Dim iMsgs As CDO.IMessages
                            Dim iMsgReply As CDO.IMessage
                            Dim iMsgReplyAll As CDO.IMessage
                            Dim iMsgForward As CDO.IMessage
                            ' Get the messages from the Drop directory.
                            iMsgs = iDropDir.GetMessages(ConfigurationManager.AppSettings("temp_mail").ToString())
                            '  Console.WriteLine("Messages Count : " + iMsgs.Count.ToString())
                            For Each iMsg As CDO.IMessage In iMsgs
                                If (iMsgs.FileName(iMsg) = path) Then
                                    'lblContent.Text = iMsg.TextBody
                                    txtBody.Text = ""
                                    txtSubject.Text = "Re: " + iMsg.Subject
                                    txtFrom.Text = acc.get_user(User.Identity.Name).Tables(0).Rows(0)("email").ToString()
                                    Dim tomail As String = iMsg.Fields("urn:schemas:httpmail:from").Value
                                    tomail = tomail.Substring(tomail.IndexOf("<") + 1)
                                    tomail = tomail.Substring(0, tomail.IndexOf(">"))
                                    txtTo.Text = tomail
                                End If
                            Next
                            ' Clean up memory.
                            iMsgs = Nothing
                            iMsgReply = Nothing
                            iMsgReplyAll = Nothing
                            iMsgForward = Nothing
                        Catch ex1 As Exception
                            'Console.WriteLine("{0} Exception caught.", e)
                        End Try
                    End If
                End If
            Else
                Response.Redirect("emailviewer.aspx")
            End If
        End If
    End Sub
    Protected Sub btnSend_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSend.Click
        SetMailClass()
        Dim source As String = ConfigurationManager.AppSettings("temp_mail") + "\" + Request.QueryString("id").ToString() + ".eml"

        Try
            Dim iDropDir As New CDO.DropDirectory()

            Dim iMsgs As CDO.IMessages
            Dim iMsgReply As CDO.IMessage
            Dim iMsgReplyAll As CDO.IMessage
            Dim iMsgForward As CDO.IMessage
            ' Get the messages from the Drop directory.
            iMsgs = iDropDir.GetMessages(ConfigurationManager.AppSettings("temp_mail").ToString())
            '  Console.WriteLine("Messages Count : " + iMsgs.Count.ToString())
            For Each iMsg As CDO.IMessage In iMsgs
                If (iMsg.Fields("urn:schemas:mailheader:message-id").Value = em.get_single_email_received(Convert.ToInt32(Request.QueryString("id").ToString())).Tables(0).Rows(0)("msg_id").ToString()) Then
                    ' If (m.SendMessage(iMsg.Fields("urn:schemas:httpmail:from").Value, iMsg.Fields("urn:schemas:httpmail:to").Value, iMsg.Subject, ereader.StripTagsCharArray(iMsg.HTMLBody).Replace("&nbsp;", " ")) = True) Then
                    'If (m.move_to_inbox(source, detination) = True) Then
                    If Request.QueryString("val").ToString() = "forward" Then
                        Dim i As Integer
                        Dim arratt As New ArrayList
                        For i = 1 To iMsg.Attachments.Count
                            arratt.Add(iMsg.Attachments(i).FileName)
                            iMsg.Attachments(i).SaveToFile(Server.MapPath("Upload") + "\\" + iMsg.Attachments(i).FileName)
                        Next
                        'Dim arrfiles As New ArrayList
                        'For i = 1 To iMsg.Attachments.Count
                        '    arrfiles.Add(iMsg.Attachments(i))
                        '    iMsg.Attachments(i).SaveToFile(Server.MapPath("Upload") + "\\" + iMsg.Attachments(i).FileName)
                        'Next
                        '  If (m.SendEmailMessage(iMsg.Fields("urn:schemas:httpmail:from").Value, iMsg.Fields("urn:schemas:httpmail:to").Value, iMsg.Subject, iMsg.HTMLBody, arratt) = True) Then
                        If (m.SendEmailMessage(txtFrom.Text, txtTo.Text, txtSubject.Text, HttpUtility.HtmlDecode(txtBody.Text), arratt) = True) Then


                            'em.Delete_email_received(Convert.ToInt32(Request.QueryString("id")))

                            Dim file As New System.IO.FileInfo(source)
                            If (file.Exists) Then
                                file.Delete()
                            End If
                            'End If
                            ' Load_email()
                            Dim strScript As String = "<script>"
                            strScript += "alert('Message: Mail sent successfully.');"
                            strScript += "window.location='emailviewer.aspx';"
                            strScript += "</script>"
                            Me.ClientScript.RegisterStartupScript(Me.[GetType](), "Startup", strScript)
                        End If
                    End If
                    If Request.QueryString("val").ToString() = "reply" Then
                        Dim i As Integer
                        Dim arratt As New ArrayList
                        'For i = 1 To iMsg.Attachments.Count
                        '    arratt.Add(iMsg.Attachments(i).FileName)
                        '    iMsg.Attachments(i).SaveToFile(Server.MapPath("Upload") + "\\" + iMsg.Attachments(i).FileName)
                        'Next
                        'Dim arrfiles As New ArrayList
                        'For i = 1 To iMsg.Attachments.Count
                        '    arrfiles.Add(iMsg.Attachments(i))
                        '    iMsg.Attachments(i).SaveToFile(Server.MapPath("Upload") + "\\" + iMsg.Attachments(i).FileName)
                        'Next
                        '  If (m.SendEmailMessage(iMsg.Fields("urn:schemas:httpmail:from").Value, iMsg.Fields("urn:schemas:httpmail:to").Value, iMsg.Subject, iMsg.HTMLBody, arratt) = True) Then
                        If (m.SendEmailMessage(txtFrom.Text, txtTo.Text, txtSubject.Text, HttpUtility.HtmlDecode(txtBody.Text), arratt) = True) Then


                            'em.Delete_email_received(Convert.ToInt32(Request.QueryString("id")))

                            Dim file As New System.IO.FileInfo(source)
                            If (file.Exists) Then
                                file.Delete()
                            End If
                            'End If
                            ' Load_email()
                            Dim strScript As String = "<script>"
                            strScript += "alert('Message: Mail sent successfully.');"
                            strScript += "window.location='emailviewer.aspx';"
                            strScript += "</script>"
                            Me.ClientScript.RegisterStartupScript(Me.[GetType](), "Startup", strScript)
                        End If
                    End If
                End If
            Next
            ' Clean up memory.
            iMsgs = Nothing
            iMsgReply = Nothing
            iMsgReplyAll = Nothing
            iMsgForward = Nothing
        Catch ex1 As Exception
            'Console.WriteLine("{0} Exception caught.", e)
        End Try

    End Sub

    Private Sub SetMailClass()
        m.p_strAlias = acc.get_user(User.Identity.Name).Tables(0).Rows(0)("alias").ToString()
        m.p_strServer = ConfigurationManager.AppSettings("Exchange_server").ToString() + m.p_strAlias + "/"
        m.p_strUserName = acc.get_user(User.Identity.Name).Tables(0).Rows(0)("alias").ToString()
        m.p_strPassword = acc.get_user(User.Identity.Name).Tables(0).Rows(0)("password").ToString()
        m.p_strInboxURL = ConfigurationManager.AppSettings("inbox_name").ToString()
        m.p_strSpam = ConfigurationManager.AppSettings("spam_name").ToString()
        m.p_strSpamServer = ConfigurationManager.AppSettings("Spam_Exchange_server").ToString() + ConfigurationManager.AppSettings("spam_alias").ToString() + "/"
        m.p_strSpamPwd = ConfigurationManager.AppSettings("spam_pwd").ToString()
        m.p_strSpamUserName = ConfigurationManager.AppSettings("spam_username").ToString()
    End Sub
End Class