' MailMessage with Email Stream and File support
' =================
'
' copyright by Peter Huber, Singapore, 2006
' this code is provided as is, bugs are probable, free for any use at own risk, no 
' responsibility accepted. All rights, title and interest in and to the accompanying content retained.  :-)
'
' based on Standard for ARPA Internet Text Messages, http://rfc.net/rfc822.html
' based on MIME Standard,  Internet Message Bodies, http://rfc.net/rfc2045.html
' based on MIME Standard, Media Types, http://rfc.net/rfc2046.html
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Text

''' <summary>
''' Stores all MIME decoded information of a received email. One email might consist of
''' several MIME entities, which have a very similar structure to an email. A RxMailMessage
''' can be a top most level email or a MIME entity the emails contains.
''' 
''' According to various RFCs, MIME entities can contain other MIME entities 
''' recursively. However, they usually need to be mapped to alternative views and 
''' attachments, which are non recursive.
'''
''' RxMailMessage inherits from System.Net.MailMessage, but provides additional receiving related information 
''' </summary>
Public Class RxMailMessage
    Inherits MailMessage
    ''' <summary>
    ''' To whom the email was delivered to
    ''' </summary>
    Public DeliveredTo As MailAddress
    ''' <summary>
    ''' To whom the email was
    ''' </summary>
    Public ReturnPath As MailAddress
    ''' <summary>
    ''' 
    ''' </summary>
    Public DeliveryDate As DateTime
    ''' <summary>
    ''' Date when the email was received
    ''' </summary>
    Public MessageId As String
    ''' <summary>
    ''' probably '1,0'
    ''' </summary>
    Public MimeVersion As String
    ''' <summary>
    ''' It may be desirable to allow one body to make reference to another. Accordingly, 
    ''' bodies may be labelled using the "Content-ID" header field.    
    ''' </summary>
    Public ContentId As String
    ''' <summary>
    ''' some descriptive information for body
    ''' </summary>
    Public ContentDescription As String
    ''' <summary>
    ''' ContentDisposition contains normally redundant information also stored in the 
    ''' ContentType. Since ContentType is more detailed, it is enough to analyze ContentType
    ''' 
    ''' something like:
    ''' inline
    ''' inline; filename="image001.gif
    ''' attachment; filename="image001.jpg"
    ''' </summary>
    Public ContentDisposition As ContentDisposition
    ''' <summary>
    ''' something like "7bit" / "8bit" / "binary" / "quoted-printable" / "base64"
    ''' </summary>
    Public TransferType As String
    ''' <summary>
    ''' similar as TransferType, but .NET supports only "7bit" / "quoted-printable"
    ''' / "base64" here, "bit8" is marked as "bit7" (i.e. no transfer encoding needed), 
    ''' "binary" is illegal in SMTP
    ''' </summary>
    Public ContentTransferEncoding As TransferEncoding
    ''' <summary>
    ''' The Content-Type field is used to specify the nature of the data in the body of a
    ''' MIME entity, by giving media type and subtype identifiers, and by providing 
    ''' auxiliary information that may be required for certain media types. Examples:
    ''' text/plain;
    ''' text/plain; charset=ISO-8859-1
    ''' text/plain; charset=us-ascii
    ''' text/plain; charset=utf-8
    ''' text/html;
    ''' text/html; charset=ISO-8859-1
    ''' image/gif; name=image004.gif
    ''' image/jpeg; name="image005.jpg"
    ''' message/delivery-status
    ''' message/rfc822
    ''' multipart/alternative; boundary="----=_Part_4088_29304219.1115463798628"
    ''' multipart/related; 	boundary="----=_Part_2067_9241611.1139322711488"
    ''' multipart/mixed; 	boundary="----=_Part_3431_12384933.1139387792352"
    ''' multipart/report; report-type=delivery-status; boundary="k04G6HJ9025016.1136391237/carbon.singnet.com.sg"
    ''' </summary>
    Public ContentType As ContentType
    ''' <summary>
    ''' .NET framework combines MediaType (text) with subtype (plain) in one property, but
    ''' often one or the other is needed alone. MediaMainType in this example would be 'text'.
    ''' </summary>
    Public MediaMainType As String
    ''' <summary>
    ''' .NET framework combines MediaType (text) with subtype (plain) in one property, but
    ''' often one or the other is needed alone. MediaSubType in this example would be 'plain'.
    ''' </summary>
    Public MediaSubType As String
    ''' <summary>
    ''' RxMessage can be used for any MIME entity, as a normal message body, an attachement or an alternative view. ContentStream
    ''' provides the actual content of that MIME entity. It's mainly used internally and later mapped to the corresponding 
    ''' .NET types.
    ''' </summary>
    Public ContentStream As Stream
    ''' <summary>
    ''' A MIME entity can contain several MIME entities. A MIME entity has the same structure
    ''' like an email. 
    ''' </summary>
    Public Entities As List(Of RxMailMessage)
    ''' <summary>
    ''' This entity might be part of a parent entity
    ''' </summary>
    Public Parent As RxMailMessage
    ''' <summary>
    ''' The top most MIME entity this MIME entity belongs to (grand grand grand .... parent)
    ''' </summary>
    Public TopParent As RxMailMessage
    ''' <summary>
    ''' The complete entity in raw content. Since this might take up quiet some space, the raw content gets only stored if the
    ''' Pop3MimeClient.isGetRawEmail is set.
    ''' </summary>
    Public RawContent As String
    ''' <summary>
    ''' Headerlines not interpretable by Pop3ClientEmail
    ''' <example></example>
    ''' </summary>
    Public UnknowHeaderlines As List(Of String)
    '
    ' Constructors
    ' ------------
    ''' <summary>
    ''' default constructor
    ''' </summary>
    Public Sub New()
        'for the moment, we assume to be at the top
        'should this entity become a child, TopParent will be overwritten
        TopParent = Me
        Entities = New List(Of RxMailMessage)()
        UnknowHeaderlines = New List(Of String)()
    End Sub
    ''' <summary>
    ''' Set all content disposition related fields
    ''' </summary>
    Public Sub SetContentDisposition(ByVal headerLineContent As String)
        ' example Content-Disposition: inline; filename="PilotsEy.gif"; size=7242; creation-date="Thu, 13 Nov 2008 14:03:50 GMT"; modification-date="Thu, 13 Nov 2008 14:03:50 GMT"
        Dim saParms As String() = headerLineContent.Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)
        If saParms.Length = 0 Then
            Me.ContentDisposition = New ContentDisposition("inline")
            Return
        End If
        ' do the type and create the object
        Me.ContentDisposition = New ContentDisposition(saParms(0).Trim())
        ' now for the parms (skip the first array value since the RFC says is has to be the type and is done already)
        Dim i As Integer = 1
        While i < saParms.Length
            Dim saNameValue As String() = saParms(i).Split(New Char() {"="c})
            If saNameValue.Length <> 2 Then
                Continue While
            End If
            ' shouldn't happen
            Dim sName As String = saNameValue(0).Trim().ToLower()
            Dim sValue As String = saNameValue(1).Trim()
            sValue = sValue.Replace("""", "")
            Select Case sName
                Case "filename"
                    Me.ContentDisposition.FileName = sValue
                    Exit Select
                Case "size"
                    Me.ContentDisposition.Size = Long.Parse(sValue)
                    Exit Select
                Case "creation-date"
                    Me.ContentDisposition.CreationDate = DateTime.Parse(sValue)
                    Exit Select
                Case "modification-date"
                    Me.ContentDisposition.ModificationDate = DateTime.Parse(sValue)
                    Exit Select
                Case "read-date"
                    Me.ContentDisposition.ReadDate = DateTime.Parse(sValue)
                    Exit Select
            End Select
            System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
        End While
    End Sub
    ''' <summary>
    ''' Set all content type related fields
    ''' </summary>
    Public Sub SetContentTypeFields(ByVal contentTypeString As String)
        contentTypeString = contentTypeString.Trim()
        'set content type
        If contentTypeString = Nothing OrElse contentTypeString.Length < 1 Then
            ContentType = New ContentType("text/plain; charset=us-ascii")
        Else
            ContentType = New ContentType(contentTypeString)
        End If
        'set encoding (character set)
        If ContentType.CharSet = Nothing Then
            BodyEncoding = Encoding.ASCII
        Else
            Try
                BodyEncoding = Encoding.GetEncoding(ContentType.CharSet)
            Catch
                BodyEncoding = Encoding.ASCII
            End Try
        End If
        'set media main and sub type
        If ContentType.MediaType = Nothing OrElse ContentType.MediaType.Length < 1 Then
            'no mediatype found
            ContentType.MediaType = "text/plain"
        Else
            Dim mediaTypeString As String = ContentType.MediaType.Trim().ToLowerInvariant()
            Dim slashPosition As Integer = ContentType.MediaType.IndexOf("/")
            If slashPosition < 1 Then
                'only main media type found
                MediaMainType = mediaTypeString
                System.Diagnostics.Debugger.Break()
                'didn't have a sample email to test this
                If MediaMainType = "text" Then
                    MediaSubType = "plain"
                Else
                    MediaSubType = ""
                End If
            Else
                'also submedia found
                MediaMainType = mediaTypeString.Substring(0, slashPosition)
                If mediaTypeString.Length > slashPosition Then
                    MediaSubType = mediaTypeString.Substring(slashPosition + 1)
                Else
                    If MediaMainType = "text" Then
                        MediaSubType = "plain"
                    Else
                        MediaSubType = ""
                        'didn't have a sample email to test this
                        System.Diagnostics.Debugger.Break()
                    End If
                End If
            End If
        End If
        IsBodyHtml = MediaSubType = "html"
    End Sub
    ''' <summary>
    ''' Creates an empty child MIME entity from the parent MIME entity.
    ''' 
    ''' An email can consist of several MIME entities. A entity has the same structure
    ''' like an email, that is header and body. The child inherits few properties 
    ''' from the parent as default value.
    ''' </summary>
    Public Function CreateChildEntity() As RxMailMessage
        Dim child As New RxMailMessage()
        child.Parent = Me
        child.TopParent = Me.TopParent
        child.ContentTransferEncoding = Me.ContentTransferEncoding
        Return child
    End Function
    Public Shared Function CreateFromFile(ByVal sEmlPath As String) As RxMailMessage
        Dim mimeDecoder As New MimeReader()
        Return mimeDecoder.GetEmail(sEmlPath)
    End Function
    Public Shared Function CreateFromFile(ByVal mimeDecoder As MimeReader, ByVal sEmlPath As String) As RxMailMessage
        Return mimeDecoder.GetEmail(sEmlPath)
    End Function
    Public Shared Function CreateFromStream(ByVal EmailStream As Stream) As RxMailMessage
        Dim mimeDecoder As New MimeReader()
        Return mimeDecoder.GetEmail(EmailStream)
    End Function
    Public Shared Function CreateFromStream(ByVal mimeDecoder As MimeReader, ByVal EmailStream As Stream) As RxMailMessage
        Return mimeDecoder.GetEmail(EmailStream)
    End Function
#If UNEEDED_CODE Then
		Private mailStructure As StringBuilder
		Private Sub AppendLine(format As String, arg As Object)
			If arg <> Nothing Then
				Dim argString As String = arg.ToString()
				If argString.Length > 0 Then
					mailStructure.AppendLine(String.Format(format, argString))
				End If
			End If
		End Sub
		Private Sub decodeEntity(entity As RxMailMessage)
			AppendLine("From  : {0}", entity.From)
			AppendLine("Sender: {0}", entity.Sender)
			AppendLine("To    : {0}", entity.[To])
			AppendLine("CC    : {0}", entity.CC)
			AppendLine("ReplyT: {0}", entity.ReplyTo)
			AppendLine("Sub   : {0}", entity.Subject)
			AppendLine("S-Enco: {0}", entity.SubjectEncoding)
			If entity.DeliveryDate > DateTime.MinValue Then
				AppendLine("Date  : {0}", entity.DeliveryDate)
			End If
			If entity.Priority <> MailPriority.Normal Then
				AppendLine("Priori: {0}", entity.Priority)
			End If
			If entity.Body.Length > 0 Then
				AppendLine("Body  : {0} byte(s)", entity.Body.Length)
				AppendLine("B-Enco: {0}", entity.BodyEncoding)
			Else
				If entity.BodyEncoding <> Encoding.ASCII Then
					AppendLine("B-Enco: {0}", entity.BodyEncoding)
				End If
			End If
			AppendLine("T-Type: {0}", entity.TransferType)
			AppendLine("C-Type: {0}", entity.ContentType)
			AppendLine("C-Desc: {0}", entity.ContentDescription)
			AppendLine("C-Disp: {0}", entity.ContentDisposition)
			AppendLine("C-Id  : {0}", entity.ContentId)
			AppendLine("M-ID  : {0}", entity.MessageId)
			AppendLine("Mime  : Version {0}", entity.MimeVersion)
			If entity.ContentStream <> Nothing Then
				AppendLine("Stream: Length {0}", entity.ContentStream.Length)
			End If
			'decode all shild MIME entities
			For Each child As RxMailMessage In entity.Entities
				mailStructure.AppendLine("------------------------------------")
				decodeEntity(child)
			Next
			If entity.ContentType <> Nothing AndAlso entity.ContentType.MediaType <> Nothing AndAlso entity.ContentType.MediaType.StartsWith("multipart") Then
				AppendLine("End {0}", entity.ContentType.ToString())
			End If
		End Sub
		''' <summary>
		''' Convert structure of message into a string
		''' </summary>
		''' <returns></returns>
		Public Function MailStructure() As String
			mailStructure = New StringBuilder(1000)
			decodeEntity(Me)
			mailStructure.AppendLine("====================================")
			Return mailStructure.ToString()
		End Function
#End If
End Class