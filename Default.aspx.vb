Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Configuration
Imports System.Xml
Imports System.IO
Imports System.Net
Imports System.Globalization


Public Class _Default
    Inherits System.Web.UI.Page
    Dim m As New Mails()
    Dim em As New email_class()
    Dim acc As New Account_DAL()
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (User.Identity.IsAuthenticated = True) Then
            If (User.Identity.Name = "admin") Then
                Response.Redirect("Admin/admin.aspx")
            Else
                If (Request.QueryString("id") Is Nothing) Then
                    Response.Redirect("Emailviewer.aspx")
                End If
            End If

        End If
    End Sub
End Class