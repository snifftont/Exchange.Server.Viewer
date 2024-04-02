Imports System.Web.SessionState
Imports System.Threading
Imports System.Xml

Public Class Global_asax
    Inherits System.Web.HttpApplication
    Public emailSender As System.Threading.Thread = Nothing
    Public SpamSender As System.Threading.Thread = Nothing
    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        emailSender = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf RunEmailSender))
        emailSender.Start()
        SpamSender = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf RunSpamSender))
        SpamSender.Start()
    End Sub
    Private Sub RunEmailSender()

        Dim Start As DateTime = DateTime.Now.AddSeconds(0)
        Dim keepRunning As Boolean = True
        Dim NextExecute As DateTime = Start
        'If (Convert.ToDateTime(DateTime.Now.ToShortDateString() + " " + emailstatic.email_at().Tables(0).Rows(0)(0).ToString())) > NextExecute Then
        '    NextExecute = Convert.ToDateTime(DateTime.Now.ToShortDateString() + " " + emailstatic.email_at().Tables(0).Rows(0)(0).ToString())
        'Else
        '    NextExecute = Convert.ToDateTime(DateTime.Now.ToShortDateString() + " " + emailstatic.email_at().Tables(0).Rows(0)(0).ToString()).AddDays(1)
        '    ' NextExecute = NextExecute.AddMinutes(2)
        'End If
        While keepRunning
            If (NextExecute.ToShortTimeString() = emailstatic.email_at().Tables(0).Rows(0)(0).ToString()) Then
                If DateTime.Now > NextExecute Then
                    keepRunning = emailstatic.entry()
                    NextExecute = Convert.ToDateTime(DateTime.Now.ToShortDateString() + " " + emailstatic.email_at().Tables(0).Rows(0)(0).ToString()).AddDays(1)
                    'NextExecute = NextExecute.AddMinutes(2)
                End If
            Else
                If (NextExecute.ToShortDateString = DateTime.Now.ToShortDateString) Then
                    If (NextExecute > Convert.ToDateTime(DateTime.Now.ToShortDateString() + " " + emailstatic.email_at().Tables(0).Rows(0)(0).ToString())) Then
                        NextExecute = Convert.ToDateTime(DateTime.Now.ToShortDateString() + " " + emailstatic.email_at().Tables(0).Rows(0)(0).ToString()).AddDays(1)
                    Else
                        keepRunning = emailstatic.entry()
                        NextExecute = Convert.ToDateTime(DateTime.Now.ToShortDateString() + " " + emailstatic.email_at().Tables(0).Rows(0)(0).ToString()).AddDays(1)
                    End If
                Else
                    NextExecute = Convert.ToDateTime(DateTime.Now.ToShortDateString() + " " + emailstatic.email_at().Tables(0).Rows(0)(0).ToString()).AddDays(1)
                End If
            End If
            Thread.Sleep(100)
        End While

    End Sub
    Private Sub RunSpamSender()


        Dim Start As DateTime = DateTime.Now.AddSeconds(0)
        Dim keepRunning As Boolean = True
        Dim NextExecute As DateTime = Start
        While keepRunning
            If DateTime.Now > NextExecute Then
                Dim m As New Mails()
                SetMailClass()
                Dim xmlDoc2 As XmlDocument = m.GetjunkMailAll()
                DisplayXML(xmlDoc2, m.p_strSpam)
                emailstatic.delete_daily_email()

                NextExecute = DateTime.Now.AddMinutes(15)
            End If
            System.Threading.Thread.Sleep(100)
            ' keepRunning = False
        End While
    End Sub
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
        Dim m As New Mails()
        Dim em As New email_class()
        Dim acc As New Account_DAL()
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

                        em.Update_email_received(Convert.ToInt32(em.check_duplicate_email_received(arr_data(0).ToString()).Tables(0).Rows(0)("er_id")), 0, Convert.ToInt32(arr_data(5)), 1, 0, 0)

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
                       
                        Dim subject As String = arr_data(3).ToString().Substring(arr_data(3).ToString().IndexOf("-") + 2)
                        Dim tomail As String = arr_data(3).ToString().Substring(0, arr_data(3).ToString().IndexOf(" - "))
                        Dim filename As String = arr_data(2).ToString().Substring(arr_data(2).ToString().IndexOf(">") + 1) + "_" + tomail + "_" + arr_data(3).ToString().Replace(" ", "").Replace("/", "-").Replace("/", "-").Replace(".", "-") + "_" + dt.ToString().Replace(" ", "").Replace("/", "-") + ".eml"

                        em.Insert_email_received(0, a, arr_data(2).ToString(), subject, 0, Convert.ToInt32(arr_data(5)), 1, 0, 0, 0, Convert.ToInt32(arr_data(1)), arr_data(0).ToString(), arr_data(6).ToString() + "B", filename, arr_data(8).ToString(), tomail)
                        Dim er_id As String = em.check_duplicate_email_received(arr_data(0).ToString()).Tables(0).Rows(0)("er_id").ToString()
                        m.get_eml(folder, arr_data(0).ToString(), er_id)
                        m.delete_mail(arr_data(0).ToString())
                       

                    End If
                End If
            Catch ex As Exception

            End Try


        Next
    End Sub
    Private Sub SetMailClass()
        Dim m As New Mails()
        Dim em As New email_class()
        Dim acc As New Account_DAL()
        m.p_strInboxURL = ConfigurationManager.AppSettings("inbox_name").ToString()
        m.p_strSpam = ConfigurationManager.AppSettings("spam_name").ToString()
        m.p_strSpamServer = ConfigurationManager.AppSettings("Spam_Exchange_server").ToString() + ConfigurationManager.AppSettings("spam_alias").ToString() + "/"
        m.p_strSpamPwd = ConfigurationManager.AppSettings("spam_pwd").ToString()
        m.p_strSpamUserName = ConfigurationManager.AppSettings("spam_username").ToString()
    End Sub
    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the session is started
    End Sub

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires at the beginning of each request
    End Sub

    Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires upon attempting to authenticate the use
    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when an error occurs
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the session ends
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the application ends
        'emailSender.Abort()
    End Sub
End Class