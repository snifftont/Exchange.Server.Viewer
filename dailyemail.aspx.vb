
Public Class dailyemail
    Inherits System.Web.UI.Page
    'Dim keepRunning As Boolean
    'Public emailpro As System.Threading.Thread = Nothing
    'Dim count As Integer = 1
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Dim em As New email_class
        'em.Delete_email_received(10920)
        If (Request.QueryString("myid") IsNot Nothing) Then
            '    count = 0
            '    Threading.Thread.Sleep(1000)

            '    emailpro = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf RunEmailSender))
            '    emailpro.Start()
            '    'keepRunning = False
            '    'System.Threading.Thread.Sleep(1000)
            '    'RunEmailSender()
        End If
    End Sub
    'Private Sub RunEmailSender()
    '    'count = 1

    '    'Dim Start As DateTime = DateTime.Now.AddSeconds(0)
    '    'keepRunning = True
    '    'Dim NextExecute As DateTime = Start
    '    'If (Convert.ToDateTime(DateTime.Now.ToShortDateString() + " " + ConfigurationManager.AppSettings("send_mail_at").ToString())) > NextExecute Then
    '    '    NextExecute = Convert.ToDateTime(DateTime.Now.ToShortDateString() + " " + ConfigurationManager.AppSettings("send_mail_at").ToString())
    '    'Else
    '    '    NextExecute = Convert.ToDateTime(DateTime.Now.ToShortDateString() + " " + ConfigurationManager.AppSettings("send_mail_at").ToString()).AddDays(1)
    '    '    ' NextExecute = NextExecute.AddMinutes(2)
    '    'End If
    '    'Dim i As Long = System.DateTime.Now.Ticks
    '    'While (count = 1)
    '    '    If DateTime.Now > NextExecute Then
    '    '        keepRunning = emailstatic.entry()
    '    '        NextExecute = Convert.ToDateTime(DateTime.Now.ToShortDateString() + " " + ConfigurationManager.AppSettings("send_mail_at").ToString()).AddDays(1)
    '    '        'NextExecute = NextExecute.AddMinutes(2)
    '    '    End If
    '    'End While
    '    'emailpro.Abort()
    'End Sub
End Class