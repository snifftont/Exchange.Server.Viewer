Imports System.IO

Public Class Site
    Inherits System.Web.UI.MasterPage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim est As TimeZoneInfo
        Try
            est = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")
            '  DateTime targetTime = TimeZoneInfo.ConvertTime(System.DateTime.Now, est);
            lblzone.Text = est.StandardName.ToString()
        Catch
        End Try
    End Sub

    Protected Sub LoginStatus1_LoggingOut(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.LoginCancelEventArgs) Handles LoginStatus1.LoggingOut
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
End Class