Imports System.Xml
Imports System.IO
Imports System.Threading

Public Class emailviewer
    Inherits System.Web.UI.Page
    Dim m As New Mails()
    Dim em As New email_class()
    Dim acc As New Account_DAL()
    Dim ereader As New EMLReader()
    Dim starting As Integer
    Dim ending As Integer
    Public emailpro1 As System.Threading.Thread = Nothing
    Dim total_inbox As Integer
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If (User.Identity.IsAuthenticated = True) Then
            If (User.Identity.Name = "admin") Then
                Dim strScript As String = "<script>"
                strScript += "alert('Message: Please Login as a user.');"
                strScript += "window.location='/admin/admin.aspx';"
                strScript += "</script>"
                Me.ClientScript.RegisterStartupScript(Me.[GetType](), "Startup", strScript)
            Else
                If (Not IsPostBack) Then
                    lblTotal.Text = ""
                    emailstatic.delete_daily_email()
                    ddlEmailof.SelectedIndex = 1
                    Panel1.Visible = True
                    Panel2.Visible = False
                    del_all()
                    Load_email(0, Convert.ToInt32(ddlPage.SelectedItem.Text) - 1)
                    GridView2.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
                    GridView1.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
                    GridView2.DataSource = bind_data()
                    GridView1.DataBind()
                    GridView2.DataSource = bind_data()
                    GridView2.DataBind()
                    'emailpro1 = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf RunEmailProcess))
                    'emailpro1.Start()
                End If
                If (ddlEmailof.SelectedIndex = 1) Then
                    Panel2.Visible = True
                    Panel1.Visible = False
                Else
                    Panel1.Visible = True
                    Panel2.Visible = False
                End If
                If total_inbox = 0 Then
                    lblTotal.Text = "Not Available"
                End If
            End If
        Else
            Response.Redirect("Account/login.aspx")
        End If
    End Sub
    Public Sub del_all()
        Try


            Dim em As New email_class()
            Dim acc As New Account_DAL()
            Dim ds As DataSet = em.get_email_received(0, 0, 1, Convert.ToInt32(acc.get_user(HttpContext.Current.User.Identity.Name).Tables(0).Rows(0)("user_id")), Convert.ToInt32(acc.get_user(HttpContext.Current.User.Identity.Name).Tables(0).Rows(0)("user_e_id")))
            Dim i As Integer
            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim path As String = ""
                Try
                    path = ConfigurationManager.AppSettings("temp_mail") + "\" + ds.Tables(0).Rows(i)("er_id").ToString() + ".eml"
                    File.Delete(path)
                Catch ex As Exception

                End Try
            Next
            em.Delete_email_received_user(Convert.ToInt32(acc.get_user(HttpContext.Current.User.Identity.Name).Tables(0).Rows(0)("user_id")))
        Catch ex As Exception

        End Try
    End Sub
    Protected Sub ddlEmailof_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlEmailof.SelectedIndexChanged

        If (ddlEmailof.SelectedIndex = 1) Then
            GridView2.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
            GridView2.DataSource = bind_data()
            GridView2.DataBind()
            Panel1.Visible = False
            Panel2.Visible = True
        Else
            RunEmailProcess()
            GridView1.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
            GridView1.DataSource = bind_data()
            GridView1.DataBind()
            Panel1.Visible = True
            Panel2.Visible = False
        End If
    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As EventArgs)

    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If (e.CommandName = "mopen") Then
            If pnlpop2.Visible = True Then
                pnlpop2.Visible = False
            Else

                pnlpop2.Visible = True

                hlnkForward.NavigateUrl = "sendmessage.aspx?val=forward&id=" + e.CommandArgument.ToString()
                hlnkReply.NavigateUrl = "sendmessage.aspx?val=reply&id=" + e.CommandArgument.ToString()
                Dim path As String = ConfigurationManager.AppSettings("temp_mail") + "\" + e.CommandArgument.ToString() + ".eml"
                'lblContent.Text = ereader.get_email_body(path)

                Try
                    SetMailClass()
                    m.save_temp_eml(em.get_single_email_received(Convert.ToInt32(e.CommandArgument.ToString())).Tables(0).Rows(0)("email_link").ToString(), e.CommandArgument.ToString())
                    Session("mid") = path
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
                            lblicontent.Text = iMsg.HTMLBody
                            iMsg = Nothing
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
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim lblread As Label = CType(e.Row.FindControl("lblread"), Label)
            If (lblread.Text = "0") Then
                For Each cell In e.Row.Cells
                    cell.Font.Bold = True
                Next
            End If
        End If
    End Sub
    'Private Function bind_data() As DataSet
    '    Dim ds As DataSet
    '    If (ddlEmailof.SelectedIndex = 1) Then
    '        ds = em.get_email_received(1, 0, 0, Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id")), Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id")))

    '    Else

    '        ds = em.get_email_received(0, 0, 1, Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id")), Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id")))

    '        End If

    '    Return ds
    'End Function
    Private Function bind_data_search() As DataSet
        Dim ds As DataSet
        If (ddlEmailof.SelectedIndex = 1) Then
            If (rbtnSender.Checked = True) Then
                ds = em.get_email_received_search_sender(1, 0, 0, Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id")), Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id")), txtSearch.Text)
            Else
                If (rbtnSubject.Checked = True) Then
                    ds = em.get_email_received_search_subject(1, 0, 0, Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id")), Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id")), txtSearch.Text)
                Else
                    ds = em.get_email_received(0, 0, 1, Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id")), Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id")))
                End If
            End If
        Else
            If (rbtnSender.Checked = True) Then
                ds = em.get_email_received_search_sender(0, 0, 1, Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id")), Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id")), txtSearch.Text)
            Else
                If (rbtnSubject.Checked = True) Then
                    ds = em.get_email_received_search_subject(0, 0, 1, Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id")), Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id")), txtSearch.Text)
                Else
                    ds = em.get_email_received(0, 0, 1, Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id")), Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id")))
                End If
            End If
        End If

        Return ds
    End Function

    Private Function bind_data() As DataSet
        emailstatic.delete_daily_email()
        Dim ds As DataSet
        If (ddlEmailof.SelectedIndex = 1) Then
            If (ddlFilter.SelectedIndex = 1) Then
                ds = em.get_email_received_search_filter(1, 0, 0, Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id")), Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id")), System.DateTime.Now.AddDays(-7))
            Else
                If (ddlFilter.SelectedIndex = 2) Then
                    ds = em.get_email_received_search_filter(1, 0, 0, Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id")), Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id")), System.DateTime.Now.AddMonths(-1))
                Else
                    ds = em.get_email_received(1, 0, 0, Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id")), Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id")))

                End If
            End If
            lblTotal.Text = ds.Tables(0).Rows.Count.ToString()
        Else
            If (ddlFilter.SelectedIndex = 1) Then
                ds = em.get_email_received_search_filter(0, 0, 1, Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id")), Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id")), System.DateTime.Now.AddDays(-7))
            Else
                If (ddlFilter.SelectedIndex = 2) Then
                    ds = em.get_email_received_search_filter(0, 0, 1, Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id")), Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id")), System.DateTime.Now.AddMonths(-1))
                Else
                    ds = em.get_email_received(0, 0, 1, Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id")), Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id")))

                End If
            End If
            If total_inbox = 0 Then
                lblTotal.Text = "Not Available"
            Else
                lblTotal.Text = total_inbox.ToString()
            End If

        End If

        Return ds
    End Function
    Private Sub DisplayXML(ByVal xmlDoc As XmlDocument, ByVal folder As String)

        Try

            AddNode2(xmlDoc.GetElementsByTagName("a:response"), folder)

        Catch xmlEx As XmlException
            ' Response.Write(xmlEx.Message)
        Catch ex As Exception
            ' Response.Write(ex.Message)
        End Try
    End Sub
    Private Sub AddNode2(ByVal NodeList As XmlNodeList, ByVal folder As String)

        Dim i As Integer
       
        For i = 0 To NodeList.Count - 1
            Try
                If NodeList(i).ChildNodes(0) IsNot Nothing Then
                    Dim arr_data As New ArrayList
                    arr_data.Add(NodeList(i).ChildNodes(0).InnerText)
                    arr_data.Add(NodeList(i).ChildNodes(1).ChildNodes(1).ChildNodes(0).InnerText)
                    arr_data.Add(NodeList(i).ChildNodes(1).ChildNodes(1).ChildNodes(2).InnerText.Replace(">", "").Replace("<", "<br/>"))
                    arr_data.Add(NodeList(i).ChildNodes(1).ChildNodes(1).ChildNodes(3).InnerText)
                    arr_data.Add(NodeList(i).ChildNodes(1).ChildNodes(1).ChildNodes(4).InnerText)
                    arr_data.Add(NodeList(i).ChildNodes(1).ChildNodes(1).ChildNodes(5).InnerText)
                    arr_data.Add(NodeList(i).ChildNodes(1).ChildNodes(1).ChildNodes(6).InnerText)
                    arr_data.Add(NodeList(i).ChildNodes(1).ChildNodes(1).ChildNodes(1).InnerText)
                    arr_data.Add(NodeList(i).ChildNodes(1).ChildNodes(1).ChildNodes(7).InnerText)

                    If em.check_duplicate_email_received(arr_data(0).ToString()).Tables(0).Rows.Count = 1 Then
                        If (folder = m.p_strInboxURL) Then
                            em.Update_email_received(Convert.ToInt32(em.check_duplicate_email_received(arr_data(0).ToString()).Tables(0).Rows(0)("er_id")), 0, Convert.ToInt32(arr_data(5)), 0, 0, 1)
                        Else
                            em.Update_email_received(Convert.ToInt32(em.check_duplicate_email_received(arr_data(0).ToString()).Tables(0).Rows(0)("er_id")), 0, Convert.ToInt32(arr_data(5)), 1, 0, 0)
                        End If
                    Else

                        ' Coordinated Universal Time string from DateTime.Now.ToUniversalTime().ToString("u");
                        Dim localDateTime As DateTime = DateTime.Parse(arr_data(4).ToString())
                        ' Local .NET timeZone.
                        Dim utcDateTime As DateTime = localDateTime.ToUniversalTime().AddHours(-5)
                        Dim dt As DateTime = DateTime.Parse(utcDateTime.ToString().Replace("#", ""))
                        Dim datepart As String = dt.ToShortDateString()
                        Dim timepart As String = dt.ToShortTimeString()
                        Dim newtime As String() = timepart.Split(" ")
                        timepart = newtime(0) + ":00" + " " + newtime(1)
                        Dim a As String = datepart + " " + timepart
                        dt = Convert.ToDateTime(a)
                        Dim user_e_id As Integer = Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id"))
                        Dim user_id As Integer = Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id"))
                        Dim filename As String = arr_data(2).ToString().Substring(arr_data(2).ToString().IndexOf(">") + 1) + "_" + acc.get_user(User.Identity.Name).Tables(0).Rows(0)("email").ToString() + "_" + arr_data(3).ToString().Replace(" ", "").Replace("/", "-").Replace("/", "-").Replace(".", "-") + "_" + dt.ToString().Replace(" ", "").Replace("/", "-") + ".eml"

                        If (folder = m.p_strInboxURL) Then
                            Dim user_email As String = acc.get_user(User.Identity.Name).Tables(0).Rows(0)("email").ToString()
                            em.Insert_email_received(user_e_id, a, arr_data(2).ToString(), arr_data(3).ToString(), 0, Convert.ToInt32(arr_data(5)), 0, 0, 1, user_id, Convert.ToInt32(arr_data(1)), arr_data(0).ToString(), arr_data(6).ToString() + "B", filename, arr_data(8).ToString(), user_email)
                            Dim er_id As String = em.check_duplicate_email_received(arr_data(0).ToString()).Tables(0).Rows(0)("er_id").ToString()

                            'm.get_eml(folder, arr_data(0).ToString(), er_id)
                        Else
                            'Dim tomail As String = NodeList(i).ChildNodes(1).ChildNodes(1).ChildNodes(8).InnerText.ToString()
                            'tomail = tomail.Substring(tomail.IndexOf("<") + 1)
                            'tomail = tomail.Substring(0, tomail.IndexOf(">"))

                            'tomail = tomail.Substring(tomail.IndexOf("<") + 1)
                            'tomail = tomail.Substring(0, tomail.IndexOf(">"))
                            Dim user_email As String = acc.get_user(User.Identity.Name).Tables(0).Rows(0)("email").ToString()
                            Dim subject As String = arr_data(3).ToString().Substring(arr_data(3).ToString().IndexOf("-") + 2)
                            Dim tomail As String = arr_data(3).ToString().Substring(0, arr_data(3).ToString().IndexOf(" - "))
                            If (tomail = user_email) Then
                                em.Insert_email_received(user_e_id, a, arr_data(2).ToString(), subject, 0, Convert.ToInt32(arr_data(5)), 1, 0, 0, user_id, Convert.ToInt32(arr_data(1)), arr_data(0).ToString(), arr_data(6).ToString() + "B", filename, arr_data(8).ToString(), user_email)
                                Dim er_id As String = em.check_duplicate_email_received(arr_data(0).ToString()).Tables(0).Rows(0)("er_id").ToString()
                                m.get_eml(folder, arr_data(0).ToString(), er_id)
                                m.delete_mail(arr_data(0).ToString())
                                'Else
                                '  m.get_eml(folder, arr_data(0).ToString())
                            End If

                        End If

                    End If
                End If
            Catch ex As Exception

            End Try


        Next
    End Sub

    Protected Sub lnlEmailChecknew_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnlEmailChecknew.Click
        RunEmailProcess()
        Load_email(0, Convert.ToInt32(ddlPage.SelectedItem.Text) - 1)
        GridView2.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
        GridView1.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
        GridView1.DataSource = bind_data()
        GridView1.DataBind()
        If ddlEmailof.SelectedIndex = 1 Then
            Try
                SetMailClass()
                Dim xmlDoc2 As XmlDocument = m.GetjunkMailAll(starting, ending)
                DisplayXML(xmlDoc2, m.p_strSpam)
            Catch ex As Exception

            End Try
        End If
        
        GridView2.DataSource = bind_data()
        GridView2.DataBind()
    End Sub
    Private Sub Load_email(ByVal starting, ByVal ending)
        Try
            em.check_login_user_email(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("email").ToString(), Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id")), Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id")))
            SetMailClass()
            Dim xmlDoc As XmlDocument = m.GetUnreadMailAll(starting, ending)
            ' Dim xmlDoc2 As XmlDocument = m.GetjunkMailAll(starting, ending)
            DisplayXML(xmlDoc, m.p_strInboxURL)
            ' DisplayXML(xmlDoc2, m.p_strSpam)
            'read_all_spam()
            emailstatic.delete_daily_email()
            GridView2.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
            GridView1.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
            GridView1.DataSource = bind_data()
            GridView1.DataBind()
            GridView2.DataSource = bind_data()
            GridView2.DataBind()
            'lblTotal.Text = bind_data().Tables(0).Rows.Count.ToString()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub RunEmailProcess()
        SetMailClass()
        Dim xmlDoc As XmlDocument = m.GetUnreadMailAll()
        ' DisplayXML(xmlDoc, m.p_strInboxURL)
        'GridView1.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
        'GridView1.DataSource = bind_data()
        'GridView1.DataBind() GetElementsByTagName("a:response").Count
        'DisplayXML(xmlDoc, m.p_strInboxURL)
        total_inbox = xmlDoc.GetElementsByTagName("a:response").Count
        ' lblTotal.Text = total_inbox.ToString()
        ' total_inbox = m.GetMailboxSize()
    End Sub
    
    Public Sub read_all_spam()
        
        Dim reply As String = ""
        '  Dim _to As String, _from As String, _subject As String, _msgid As String, _date As String
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
                'Console.WriteLine(iMsgs.FileName(iMsg))
                '' Output some common properties of the extracted message.
                'Console.WriteLine("Subject: " + iMsg.Subject)
                'Console.WriteLine("TextBody: " + iMsg.TextBody)
                'Console.WriteLine("datereceived: " + iMsg.Fields("urn:schemas:httpmail:datereceived").Value)

                'Console.WriteLine("senderemail: " + iMsg.Fields("urn:schemas:httpmail:senderemail").Value)
                'Console.WriteLine("from: " + iMsg.Fields("urn:schemas:httpmail:from").Value)
                'Console.WriteLine("to: " + iMsg.Fields("urn:schemas:httpmail:to").Value)
                Dim to1 As String = iMsg.Fields("urn:schemas:httpmail:to").Value
                to1 = to1.Substring(to1.IndexOf("<") + 1)
                to1 = to1.Substring(0, to1.IndexOf(">"))

                Dim user_email As String = acc.get_user(User.Identity.Name).Tables(0).Rows(0)("email").ToString()
                If (to1 = user_email) Then
                    If em.check_duplicate_spam_received(iMsg.Fields("urn:schemas:mailheader:message-id").Value).Tables(0).Rows.Count = 0 Then
                        Dim a As String = iMsg.Fields("urn:schemas:httpmail:datereceived").Value.ToString().Replace("#", "")



                        Dim user_e_id As Integer = Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id"))
                        Dim user_id As Integer = Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id"))
                        ' Dim filename As String = arr_data(2).ToString().Substring(arr_data(2).ToString().IndexOf(">") + 1) + "_" + acc.get_user(User.Identity.Name).Tables(0).Rows(0)("email").ToString() + "_" + arr_data(3).ToString().Replace(" ", "").Replace("/", "-").Replace("/", "-").Replace(".", "-") + "_" + dt.ToString().Replace(" ", "").Replace("/", "-") + ".eml"
                        Dim filename As String = ""
                        Dim ca As Integer = 0
                        If (iMsg.Attachments.Count > 0) Then
                            ca = 1
                        End If
                        Dim fl As String = iMsgs.FileName(iMsg)
                        'em.Insert_email_received(user_e_id, a, arr_data(2).ToString(), arr_data(3).ToString(), 0, Convert.ToInt32(arr_data(5)), 1, 0, 0, user_id, Convert.ToInt32(arr_data(1)), arr_data(0).ToString(), arr_data(6).ToString() + "B", filename, arr_data(8).ToString())
                        em.Insert_email_received(user_e_id, a, iMsg.Fields("urn:schemas:httpmail:from").Value, iMsg.Subject, 0, 1, 1, 0, 0, user_id, ca, fl, ereader.get_email_size(fl) + "B", filename, iMsg.Fields("urn:schemas:mailheader:message-id").Value, user_email)

                        Dim er_id As String = em.check_duplicate_spam_received(iMsg.Fields("urn:schemas:mailheader:message-id").Value).Tables(0).Rows(0)("er_id").ToString()
                        m.rename_eml(fl, er_id)



                    End If
                End If
                '' Reply.
                'iMsgReply = iMsg.Reply()
                '' TODO: Change "rhaddock@northwindtraders.com" to your e-mail address.
                'iMsgReply.From = "rhaddock@northwindtraders.com"
                'iMsgReply.TextBody = "I agree. You can continue." + vbLf & vbLf + iMsgReply.TextBody
                'iMsgReply.Send()
                '' This is ReplyAll.
                'iMsgReplyAll = iMsg.ReplyAll()
                '' TODO: Change "rhaddock@northwindtraders.com" to your e-mail address.
                'iMsgReplyAll.From = "rhaddock@northwindtraders.com"
                'iMsgReplyAll.TextBody = "I agree. You can continue" + vbLf & vbLf + iMsgReplyAll.TextBody
                'iMsgReplyAll.Send()
                '' This is Forward.
                'iMsgForward = iMsg.Forward()
                '' TODO: Change "rhaddock@northwindtraders.com" to your e-mail address.
                'iMsgForward.From = "rhaddock@northwindtraders.com"
                '' TODO: Change "Jonathan@northwindtraders.com" to the address that you want to forward to.
                'iMsgForward.[To] = "Jonathan@northwindtraders.com"
                'iMsgForward.TextBody = "You missed this." + vbLf & vbLf + iMsgForward.TextBody
                'iMsgForward.Send()
            Next
            ' Clean up memory.
            iMsgs = Nothing
            iMsgReply = Nothing
            iMsgReplyAll = Nothing
            iMsgForward = Nothing
        Catch e As Exception
            'Console.WriteLine("{0} Exception caught.", e)
        End Try
        'For Each sFile As String In Directory.GetFiles(ConfigurationManager.AppSettings("email_save_at").ToString())
        '    Try

        '        Dim st As New StreamReader(sFile)
        '        Dim fc As String = st.ReadToEnd()

        '        _from = Regex.Matches(fc, "From: (.+)", RegexOptions.IgnoreCase)(0).ToString().Substring(6).TrimStart().TrimEnd()
        '        _to = Regex.Matches(fc, "To: (.+)", RegexOptions.IgnoreCase)(0).ToString().Substring(4).TrimStart().TrimEnd()
        '        _to = _to.Substring(_to.IndexOf("<") + 1)
        '        _to = _to.Substring(0, _to.IndexOf(">"))
        '        _subject = Regex.Matches(fc, "Subject: (.+)", RegexOptions.IgnoreCase)(0).ToString().Substring(9).TrimStart().TrimEnd()
        '        _msgid = Regex.Matches(fc, "Message-ID: (.+)", RegexOptions.IgnoreCase)(0).ToString().Substring(11).TrimStart().TrimEnd()
        '        _date = Regex.Matches(fc, "Date: (.+)", RegexOptions.IgnoreCase)(0).ToString().Substring(6).TrimStart().TrimEnd()
        '        st.Close()
        '    Dim localDateTime As DateTime = DateTime.Parse(_date)
        '    Dim dt As DateTime = DateTime.Parse(localDateTime.ToString().Replace("#", ""))
        '    Dim datepart As String = dt.ToShortDateString()
        '    Dim timepart As String = dt.ToShortTimeString()
        '    Dim newtime As String() = timepart.Split(" ")
        '    timepart = newtime(0) + ":00" + " " + newtime(1)
        '    Dim a As String = datepart + " " + timepart

        '    Dim filename As String = ""
        '    If em.check_duplicate_spam_received(_msgid).Tables(0).Rows.Count = 0 Then
        '            Dim user_e_id As Integer = Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id"))
        '            Dim user_email As String = acc.get_user(User.Identity.Name).Tables(0).Rows(0)("email").ToString()
        '            Dim user_id As Integer = Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id"))
        '            If (user_email = _to) Then
        '                em.Insert_email_received(user_e_id, a, _from, _subject, 0, 1, 1, 0, 0, user_id, 0, sFile, ereader.get_email_size(sFile) + "B", filename, _msgid)
        '                Dim er_id As String = em.check_duplicate_spam_received(_msgid).Tables(0).Rows(0)("er_id").ToString()
        '                m.rename_eml(sFile, er_id)
        '            End If
        '        End If
        '    Catch ex As Exception

        '    End Try
        'Next












        'For Each sFile As String In Directory.GetFiles(ConfigurationManager.AppSettings("email_save_at").ToString())
        '    Try
        '        Dim mime As New MimeReader()

        '        Dim mm As RxMailMessage = mime.GetEmail(sFile)

        'Dim arr_data As New ArrayList

        'arr_data.Add(sFile)
        'arr_data.Add(0)
        'arr_data.Add(ereader.get_email_From(sFile).Replace(">", "").Replace("<", "<br/>"))
        'arr_data.Add(ereader.get_email_Subject(sFile))
        'arr_data.Add(ereader.get_email_date(sFile))
        'arr_data.Add(1)
        'arr_data.Add(ereader.get_email_size(sFile))
        'arr_data.Add("")
        'arr_data.Add(ereader.get_email_mesgid(sFile))
        'If em.check_duplicate_spam_received(arr_data(8).ToString()).Tables(0).Rows.Count = 0 Then

        ' Dim localDateTime As DateTime = DateTime.Parse(arr_data(4).ToString())

        'Dim dt As DateTime = DateTime.Parse(localDateTime.ToString().Replace("#", ""))
        'Dim datepart As String = dt.ToShortDateString()
        'Dim timepart As String = dt.ToShortTimeString()
        'Dim newtime As String() = timepart.Split(" ")
        'timepart = newtime(0) + ":00" + " " + newtime(1)
        'Dim a As String = datepart + " " + timepart
        'dt = Convert.ToDateTime(a)
        'If em.check_duplicate_spam_received(mm.MessageId).Tables(0).Rows.Count = 0 Then
        '    Dim a As String = mm.DeliveryDate.ToString()
        '    Dim user_e_id As Integer = Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_e_id"))
        '    Dim user_id As Integer = Convert.ToInt32(acc.get_user(User.Identity.Name).Tables(0).Rows(0)("user_id"))
        '    ' Dim filename As String = arr_data(2).ToString().Substring(arr_data(2).ToString().IndexOf(">") + 1) + "_" + acc.get_user(User.Identity.Name).Tables(0).Rows(0)("email").ToString() + "_" + arr_data(3).ToString().Replace(" ", "").Replace("/", "-").Replace("/", "-").Replace(".", "-") + "_" + dt.ToString().Replace(" ", "").Replace("/", "-") + ".eml"
        '    Dim filename As String = ""
        '    Dim ca As Integer = 0
        '    If (mm.Attachments.Count > 0) Then
        '        ca = 1
        '    End If
        '    'em.Insert_email_received(user_e_id, a, arr_data(2).ToString(), arr_data(3).ToString(), 0, Convert.ToInt32(arr_data(5)), 1, 0, 0, user_id, Convert.ToInt32(arr_data(1)), arr_data(0).ToString(), arr_data(6).ToString() + "B", filename, arr_data(8).ToString())
        '    em.Insert_email_received(user_e_id, a, mm.From.Address, mm.Subject, 0, 1, 1, 0, 0, user_id, ca, sFile, ereader.get_email_size(sFile) + "B", filename, mm.MessageId)

        '    Dim er_id As String = em.check_duplicate_spam_received(mm.MessageId).Tables(0).Rows(0)("er_id").ToString()
        '    m.rename_eml(sFile, er_id)


        'End If


        '    Catch err As System.IO.IOException
        '    'Debug.WriteLine("File " + sFile + " is currently in use.")
        'End Try
        'Next

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
    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound

        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim lblread As Label = CType(e.Row.FindControl("lblread"), Label)
            If (lblread.Text = "0") Then
                For Each cell In e.Row.Cells
                    cell.Font.Bold = True
                Next
            End If
        End If
    End Sub
    Protected Sub chkall_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim chkall As CheckBox = DirectCast(sender, CheckBox)
        If chkall.Checked Then
            For Each row In GridView2.Rows
                Dim chk As CheckBox = CType(row.FindControl("chkSelect"), CheckBox)
                chk.Checked = True
            Next
        ElseIf Not chkall.Checked Then
            For Each row In GridView2.Rows
                Dim chk As CheckBox = CType(row.FindControl("chkSelect"), CheckBox)
                chk.Checked = False
            Next
        End If
    End Sub
    Protected Sub chkSelect_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)

    End Sub
    Protected Sub GridView2_PageIndexChanged(ByVal sender As Object, ByVal e As EventArgs)

    End Sub

    Protected Sub GridView2_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView2.RowCommand
        If (e.CommandName = "deliv") Then
            SetMailClass()
            Dim source As String = ConfigurationManager.AppSettings("email_save_at") + "\" + e.CommandArgument.ToString() + ".eml"

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
                    '  If (iMsg.Fields("urn:schemas:mailheader:message-id").Value = em.get_single_email_received(Convert.ToInt32(e.CommandArgument.ToString())).Tables(0).Rows(0)("msg_id").ToString()) Then
                    ' If (m.SendMessage(iMsg.Fields("urn:schemas:httpmail:from").Value, iMsg.Fields("urn:schemas:httpmail:to").Value, iMsg.Subject, ereader.StripTagsCharArray(iMsg.HTMLBody).Replace("&nbsp;", " ")) = True) Then
                    'If (m.move_to_inbox(source, detination) = True) Then
                    If (iMsgs.FileName(iMsg) = source) Then
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
                        Dim subject As String = iMsg.Subject.Substring(iMsg.Subject.IndexOf("-") + 2)
                        If (m.SendEmailMessage(iMsg.Fields("urn:schemas:httpmail:from").Value, iMsg.Fields("urn:schemas:httpmail:to").Value, subject, iMsg.HTMLBody, arratt) = True) Then


                            em.Delete_email_received(Convert.ToInt32(e.CommandArgument))

                            Dim file As New System.IO.FileInfo(source)
                            If (file.Exists) Then
                                file.Delete()
                            End If
                            'End If
                            ' Load_email()
                            Dim strScript As String = "<script>"
                            strScript += "alert('Message: Mail Delivered to your inbox successfully.');"
                            strScript += "window.location='emailviewer.aspx';"
                            strScript += "</script>"
                            Me.ClientScript.RegisterStartupScript(Me.[GetType](), "Startup", strScript)
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








            '' Dim source As String = em.get_single_email_received(Convert.ToInt32(e.CommandArgument)).Tables(0).Rows(0)("email_link").ToString()

            '' Dim detination As String = source.Replace("/" + ConfigurationManager.AppSettings("spam_name").ToString().Replace(" ", "%20") + "/", "/" + ConfigurationManager.AppSettings("inbox_name").ToString().Replace(" ", "%20") + "/")
            'Dim detination As String = m.p_strServer + ConfigurationManager.AppSettings("inbox_name").ToString().Replace(" ", "%20") + "/" + e.CommandArgument.ToString() + ".eml"

            'Dim _to As String, _from As String, _subject As String, _msgid As String, _date As String
            'Dim fc As String = New StreamReader(source).ReadToEnd()
            '_from = Regex.Matches(fc, "From: (.+)")(0).ToString().Substring(6).TrimStart().TrimEnd()
            '_to = Regex.Matches(fc, "To: (.+)")(0).ToString().Substring(4).TrimStart().TrimEnd()
            '_to = _to.Substring(_to.IndexOf("<"))
            '_subject = Regex.Matches(fc, "Subject: (.+)")(0).ToString().Substring(9).TrimStart().TrimEnd()
            '_msgid = Regex.Matches(fc, "Message-ID: (.+)")(0).ToString().Substring(11).TrimStart().TrimEnd()
            '_date = Regex.Matches(fc, "Date: (.+)")(0).ToString().Substring(6).TrimStart().TrimEnd()
            'Dim localDateTime As DateTime = DateTime.Parse(_date)
            'Dim dt As DateTime = DateTime.Parse(localDateTime.ToString().Replace("#", ""))
            'Dim datepart As String = dt.ToShortDateString()
            'Dim timepart As String = dt.ToShortTimeString()
            'Dim newtime As String() = timepart.Split(" ")
            'timepart = newtime(0) + ":00" + " " + newtime(1)
            'Dim a As String = datepart + " " + timepart
            'If (m.SendMessage(_from, _to, _subject, ereader.StripTagsCharArray(ereader.get_email_body(source)).Replace("&nbsp;", " ")) = True) Then
            '    'If (m.move_to_inbox(source, detination) = True) Then
            '    em.Delete_email_received(Convert.ToInt32(e.CommandArgument))
            '    Dim file As New System.IO.FileInfo(source)
            '    If (file.Exists) Then
            '        file.Delete()
            '    End If
            '    'End If
            '    Load_email()
            'End If
        End If
        If (e.CommandName = "viewcont") Then
            If pnlpop1.Visible = True Then
                pnlpop1.Visible = False
            Else

                pnlpop1.Visible = True
                Dim path As String = ConfigurationManager.AppSettings("email_save_at") + "\" + e.CommandArgument.ToString() + ".eml"
                'lblContent.Text = ereader.get_email_body(path)
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
                        If (iMsgs.FileName(iMsg) = path) Then
                            'lblContent.Text = iMsg.TextBody
                            lblContent.Text = iMsg.HTMLBody
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
        If (e.CommandName = "mopn") Then
            Dim path As String = ConfigurationManager.AppSettings("email_save_at") + "\" + e.CommandArgument.ToString() + ".eml"
            Dim file As New System.IO.FileInfo(path)
            If (file.Exists) Then
                Response.Clear()
                Response.ClearContent()
                Response.ClearHeaders()
                Response.ContentType = "application/octet-stream; name=" & file.Name
                Response.AddHeader("content-transfer-encoding", "binary")
                Response.AppendHeader("Content-Disposition", "attachment;filename=" & file.Name)
                Response.ContentEncoding = System.Text.Encoding.GetEncoding(1251)
                Response.WriteFile(file.FullName)
                Response.End()
            End If
        End If
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
        GridView1.PageIndex = e.NewPageIndex
        GridView1.DataSource = bind_data()
        GridView1.DataBind()
    End Sub

    Protected Sub GridView2_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView2.PageIndexChanging
        GridView2.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
        GridView2.PageIndex = e.NewPageIndex
        GridView2.DataSource = bind_data()
        GridView2.DataBind()
    End Sub

    Protected Sub btnDelSelected_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDelSelected.Click
        For Each row In GridView2.Rows
            Dim chkbk As CheckBox = CType(row.Cells(0).FindControl("ChkSelect"), CheckBox)
            If (chkbk.Checked = True) Then


                Dim lblid As Label = CType(row.Cells(0).FindControl("lblid"), Label)
                ' Dim source As String = em.get_single_email_received(Convert.ToInt32(lblid.Text)).Tables(0).Rows(0)("email_link").ToString()

                ' Dim detination As String = source.Replace("/Junk%20E-mail/", "/Inbox/")
                ' SetMailClass()
                ' If (m.move_to_inbox(source, detination) = True) Then
                em.Delete_email_received(Convert.ToInt32(lblid.Text))
                Dim path As String = ConfigurationManager.AppSettings("email_save_at") + "\" + lblid.Text + ".eml"
                Dim file As New System.IO.FileInfo(path)
                If (file.Exists) Then
                    file.Delete()
                End If
                ' End If
            End If
        Next
        RunEmailProcess()
        Load_email(0, Convert.ToInt32(ddlPage.SelectedItem.Text) - 1)
    End Sub

    Protected Sub btnDelAll_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDelAll.Click
        For Each row In GridView2.Rows
            Dim chkbk As CheckBox = CType(row.Cells(0).FindControl("ChkSelect"), CheckBox)
            ' If (chkbk.Checked = True) Then


            Dim lblid As Label = CType(row.Cells(0).FindControl("lblid"), Label)
            ' Dim source As String = em.get_single_email_received(Convert.ToInt32(lblid.Text)).Tables(0).Rows(0)("email_link").ToString()

            ' Dim detination As String = source.Replace("/Junk%20E-mail/", "/Inbox/")
            ' SetMailClass()
            ' If (m.move_to_inbox(source, detination) = True) Then
            em.Delete_email_received(Convert.ToInt32(lblid.Text))
            Dim path As String = ConfigurationManager.AppSettings("email_save_at") + "\" + lblid.Text + ".eml"
            Dim file As New System.IO.FileInfo(path)
            If (file.Exists) Then
                file.Delete()
            End If
            ' End If
            ' End If
        Next
        RunEmailProcess()
        Load_email(0, Convert.ToInt32(ddlPage.SelectedItem.Text) - 1)
    End Sub

    Protected Sub ddlPage_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlPage.SelectedIndexChanged
        GridView2.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
        GridView2.DataSource = bind_data()
        GridView2.DataBind()
        GridView1.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
        GridView1.DataSource = bind_data()
        GridView1.DataBind()
    End Sub

    Protected Sub btnGoSearch_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnGoSearch.Click
        If (ddlEmailof.SelectedIndex = 0) Then
            GridView1.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
            GridView1.DataSource = bind_data_search()
            GridView1.DataBind()
            lblTotal.Text = bind_data_search().Tables(0).Rows.Count.ToString()
        Else
            GridView2.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
            GridView2.DataSource = bind_data_search()
            GridView2.DataBind()
            lblTotal.Text = bind_data_search().Tables(0).Rows.Count.ToString()
        End If
    End Sub

    Protected Sub btnGoFilter_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnGoFilter.Click
        If (ddlEmailof.SelectedIndex = 0) Then
            GridView1.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
            GridView1.DataSource = bind_data()
            GridView1.DataBind()
        Else
            GridView2.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
            GridView2.DataSource = bind_data()
            GridView2.DataBind()
        End If
    End Sub

    Protected Sub ddlFilter_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlFilter.SelectedIndexChanged

    End Sub

    Protected Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
        lblContent.Text = ""
        pnlpop1.Visible = False

    End Sub
    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose2.Click
        lblicontent.Text = ""
        pnlpop2.Visible = False
        Try
            If (Session("mid") IsNot Nothing) Then
                Dim stream As FileStream = Nothing
                Try
                    stream = File.Open(Session("mid").ToString(), FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                Catch generatedExceptionName As IOException
                    'the file is unavailable because it is:
                    'still being written to
                    'or being processed by another thread
                    'or does not exist (has already been processed)

                Finally
                    If stream IsNot Nothing Then
                        stream.Close()
                    End If
                End Try
                File.Delete(Session("mid").ToString())
            End If
        Catch ex As Exception

        End Try
    End Sub
    Protected Sub btnClose3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose3.Click
        lblContent.Text = ""
        pnlpop1.Visible = False

    End Sub
    Protected Sub btnClose4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose4.Click
        lblicontent.Text = ""
        pnlpop2.Visible = False
        Try
            If (Session("mid") IsNot Nothing) Then
                Dim stream As FileStream = Nothing
                Try
                    stream = File.Open(Session("mid").ToString(), FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                Catch generatedExceptionName As IOException

                Finally
                    If stream IsNot Nothing Then
                        stream.Close()
                    End If
                End Try
                File.Delete(Session("mid").ToString())
            End If
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub lnkAllload_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAllload.Click
        Try
            'SetMailClass()
            'Dim xmlDoc As XmlDocument = m.GetMailAll()
            'Dim xmlDoc2 As XmlDocument = m.GetspamMailAll()
            'DisplayXML(xmlDoc, m.p_strInboxURL)
            'DisplayXML(xmlDoc2, m.p_strSpam)
            ''read_all_spam()
            'emailstatic.delete_daily_email()
            'GridView1.DataSource = bind_data()
            'GridView1.DataBind()
            'GridView2.DataSource = bind_data()
            'GridView2.DataBind()
            'lblTotal.Text = bind_data().Tables(0).Rows.Count.ToString()
            Dim i As Integer = 1
            Dim j As Integer = 1
            If (ddlEmailof.SelectedIndex = 0) Then
                If GridView1.PageCount > 0 Then
                    i = GridView1.PageCount
                End If
                If GridView1.PageIndex > 0 Then
                    j = GridView1.PageIndex
                End If
                RunEmailProcess()
                Load_email(j * GridView1.PageSize, GridView1.PageSize + (i * GridView1.PageSize) - 1)
            Else
                If GridView1.PageCount > 0 Then
                    i = GridView2.PageCount
                End If
                ' Load_email(0, 1000)
                Try
                    ' SetMailClass()
                    ' Dim xmlDoc As XmlDocument = m.GetUnreadMailAll(starting, ending)
                    '  Dim xmlDoc2 As XmlDocument = m.GetjunkMailAll(0, 1000)
                    ' DisplayXML(xmlDoc, m.p_strInboxURL)
                    '  DisplayXML(xmlDoc2, m.p_strSpam)
                    'read_all_spam()
                        Try
                            SetMailClass()
                        Dim xmlDoc2 As XmlDocument = m.GetjunkMailAll()
                            DisplayXML(xmlDoc2, m.p_strSpam)
                        Catch ex As Exception

                        End Try

                        emailstatic.delete_daily_email()

                        GridView2.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
                        GridView1.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
                        GridView1.DataSource = bind_data()
                        GridView1.DataBind()
                        GridView2.DataSource = bind_data()
                        GridView2.DataBind()
                        lblTotal.Text = bind_data().Tables(0).Rows.Count.ToString()
                Catch ex As Exception

                End Try
            End If

        Catch ex As Exception

        End Try
    End Sub

    Protected Sub btndelivers_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btndelivers.Click
        For Each row In GridView2.Rows
            Dim chkbk As CheckBox = CType(row.Cells(0).FindControl("ChkSelect"), CheckBox)
            If (chkbk.Checked = True) Then


                Dim lblid As Label = CType(row.Cells(0).FindControl("lblid"), Label)
                'Dim source As String = em.get_single_email_received(Convert.ToInt32(lblid.Text)).Tables(0).Rows(0)("email_link").ToString()

                'Dim detination As String = source.Replace("/Junk%20E-mail/", "/Inbox/")
                'SetMailClass()
                'If (m.move_to_inbox(source, detination) = True) Then
                '    em.Delete_email_received(Convert.ToInt32(lblid.Text))
                '    Dim path As String = ConfigurationManager.AppSettings("email_save_at") + "\" + lblid.Text + ".eml"
                '    Dim file As New System.IO.FileInfo(path)
                '    If (file.Exists) Then
                '        file.Delete()
                '    End If
                'End If
                Dim source As String = ConfigurationManager.AppSettings("email_save_at") + "\" + lblid.Text + ".eml"
                'Dim _to As String, _from As String, _subject As String, _msgid As String, _date As String
                'Dim fc As String = New StreamReader(source).ReadToEnd()
                '_from = Regex.Matches(fc, "From: (.+)")(0).ToString().Substring(6).TrimStart().TrimEnd()
                '_to = Regex.Matches(fc, "To: (.+)")(0).ToString().Substring(4).TrimStart().TrimEnd()
                '_to = _to.Substring(_to.IndexOf("<"))
                '_subject = Regex.Matches(fc, "Subject: (.+)")(0).ToString().Substring(9).TrimStart().TrimEnd()
                '_msgid = Regex.Matches(fc, "Message-ID: (.+)")(0).ToString().Substring(11).TrimStart().TrimEnd()
                '_date = Regex.Matches(fc, "Date: (.+)")(0).ToString().Substring(6).TrimStart().TrimEnd()
                'Dim localDateTime As DateTime = DateTime.Parse(_date)
                'Dim dt As DateTime = DateTime.Parse(localDateTime.ToString().Replace("#", ""))
                'Dim datepart As String = dt.ToShortDateString()
                'Dim timepart As String = dt.ToShortTimeString()
                'Dim newtime As String() = timepart.Split(" ")
                'timepart = newtime(0) + ":00" + " " + newtime(1)
                'Dim a As String = datepart + " " + timepart
                'If (m.SendMessage(_from, _to, _subject, ereader.StripTagsCharArray(ereader.get_email_body(source)).Replace("&nbsp;", " ")) = True) Then
                '    'If (m.move_to_inbox(source, detination) = True) Then
                '    em.Delete_email_received(Convert.ToInt32(lblid.Text))
                '    Dim file As New System.IO.FileInfo(source)
                '    If (file.Exists) Then
                '        file.Delete()
                '    End If
                '    'End If
                '    Load_email()
                'End If
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
                        ' If (iMsg.Fields("urn:schemas:mailheader:message-id").Value = em.get_single_email_received(Convert.ToInt32(lblid.Text.ToString())).Tables(0).Rows(0)("msg_id").ToString()) Then
                        If (iMsgs.FileName(iMsg) = source) Then
                            Dim subject As String = iMsg.Subject.Substring(iMsg.Subject.IndexOf("-") + 2)
                           
                            If (m.SendMessage(iMsg.Fields("urn:schemas:httpmail:from").Value, iMsg.Fields("urn:schemas:httpmail:to").Value, subject, iMsg.TextBody) = True) Then
                                'If (m.move_to_inbox(source, detination) = True) Then
                                em.Delete_email_received(Convert.ToInt32(lblid.Text))
                                Dim file As New System.IO.FileInfo(source)
                                If (file.Exists) Then
                                    file.Delete()
                                End If
                                'End If
                                Dim strScript As String = "<script>"
                                strScript += "alert('Message: Mail Delivered to your inbox successfully.');"
                                strScript += "window.location='emailviewer.aspx';"
                                strScript += "</script>"
                                Me.ClientScript.RegisterStartupScript(Me.[GetType](), "Startup", strScript)
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
            End If
        Next
        RunEmailProcess()
        Load_email(0, Convert.ToInt32(ddlPage.SelectedItem.Text) - 1)
    End Sub

    Protected Sub LinkButton1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles LinkButton1.Click
        Try
            'SetMailClass()
            'Dim xmlDoc As XmlDocument = m.GetMailAll()
            'Dim xmlDoc2 As XmlDocument = m.GetspamMailAll()
            'DisplayXML(xmlDoc, m.p_strInboxURL)
            'DisplayXML(xmlDoc2, m.p_strSpam)
            ''read_all_spam()
            'emailstatic.delete_daily_email()
            'GridView1.DataSource = bind_data()
            'GridView1.DataBind()
            'GridView2.DataSource = bind_data()
            'GridView2.DataBind()
            'lblTotal.Text = bind_data().Tables(0).Rows.Count.ToString()
            Dim i As Integer = 1
            Dim j As Integer = 1
            If (ddlEmailof.SelectedIndex = 0) Then
                If GridView1.PageCount > 0 Then
                    i = GridView1.PageCount
                End If
                If GridView1.PageIndex > 0 Then
                    j = GridView1.PageIndex
                End If
                RunEmailProcess()
                Load_email(j * GridView1.PageSize, GridView1.PageSize + (i * GridView1.PageSize) - 1)
            Else
                If GridView1.PageCount > 0 Then
                    i = GridView2.PageCount
                End If
                ' Load_email(0, 1000)
                Try
                    ' SetMailClass()
                    ' Dim xmlDoc As XmlDocument = m.GetUnreadMailAll(starting, ending)
                    ' Dim xmlDoc2 As XmlDocument = m.GetjunkMailAll(0, 1000)
                    ' DisplayXML(xmlDoc, m.p_strInboxURL)
                    ' DisplayXML(xmlDoc2, m.p_strSpam)
                    'read_all_spam()
                    Try
                        SetMailClass()
                        Dim xmlDoc2 As XmlDocument = m.GetjunkMailAll()
                        DisplayXML(xmlDoc2, m.p_strSpam)
                    Catch ex As Exception

                    End Try
                    emailstatic.delete_daily_email()
                    RunEmailProcess()
                    GridView2.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
                    GridView1.PageSize = Convert.ToInt32(ddlPage.SelectedItem.Text)
                    GridView1.DataSource = bind_data()
                    GridView1.DataBind()
                    GridView2.DataSource = bind_data()
                    GridView2.DataBind()
                    lblTotal.Text = bind_data().Tables(0).Rows.Count.ToString()
                Catch ex As Exception

                End Try
            End If

        Catch ex As Exception

        End Try
    End Sub
End Class