
' MimeReader
' =================
'
' copyright by Peter Huber, Singapore, 2006
' this code is provided as is, bugs are probable, free for any use at own risk, no 
' responsibility accepted. All rights, title and interest in and to the accompanying content retained.  :-)
'
' based on Standard for ARPA Internet Text Messages, http://rfc.net/rfc822.html
' based on MIME Standard,  Internet Message Bodies, http://rfc.net/rfc2045.html
' based on MIME Standard, Media Types, http://rfc.net/rfc2046.html
' based on QuotedPrintable Class from ASP emporium, http://www.aspemporium.com/classes.aspx?cid=6
' based on MIME Standard, E-mail Encapsulation of HTML (MHTML), http://rfc.net/rfc2110.html
' based on MIME Standard, Multipart/Related Content-type, http://rfc.net/rfc2112.html
' ?? RFC 2557       MIME Encapsulation of Aggregate Documents http://rfc.net/rfc2557.html
Imports System
Imports System.Collections.Generic
Imports System.Runtime
Imports System.Globalization
Imports System.IO
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Text
Imports System.Text.RegularExpressions

' Mime Reader
' =================
''' <summary>
''' Reads MIME based emails streams and files
''' </summary>
Public Class MimeReader
    ''' <summary>
    ''' Reader for Email streams
    ''' </summary>
    Protected EmailStreamReader As StreamReader
    ''' <summary>
    ''' char 'array' for carriage return / line feed
    ''' </summary>
    Protected CRLF As String = vbCr & vbLf
    'character array 'constants' used for analysing MIME
    '----------------------------------------------------------
    Private Shared BlankChars As Char() = {" "c}
    Private Shared BracketChars As Char() = {"("c, ")"c}
    Private Shared ColonChars As Char() = {":"c}
    Private Shared CommaChars As Char() = {","c}
    Private Shared EqualChars As Char() = {"="c}
    Private Shared ForwardSlashChars As Char() = {"/"c}
    Private Shared SemiColonChars As Char() = {";"c}
    Private Shared WhiteSpaceChars As Char() = {" "c, ControlChars.Tab}
    Private Shared NonValueChars As Char() = {""""c, "("c, ")"c}
    'Help for debugging
    '------------------
    ''' <summary>
    ''' used for debugging. Collects all unknown header lines for all (!) emails received
    ''' </summary>
    Public Shared isCollectUnknowHeaderLines As Boolean = True
    ''' <summary>
    ''' list of all unknown header lines received, for all (!) emails 
    ''' </summary>
    Public Shared AllUnknowHeaderLines As New List(Of String)()
#If UNEEDED_CODE Then
		''' <summary>
		''' Set this flag, if you would like to get also the email in the raw US-ASCII format
		''' as received.
		''' Good for debugging, but takes quiet some space.
		''' </summary>
		Public Property IsCollectRawEmail() As Boolean
			Get
				Return isGetRawEmail
			End Get
			Set
				isGetRawEmail = value
			End Set
		End Property
		Private isGetRawEmail As Boolean = False
#End If
    ' MimeReader Constructor
    '---------------------------
    ''' <summary>
    ''' constructor
    ''' </summary>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Gets an email from the supplied Email Stream and processes it.
    ''' </summary>
    ''' <param name="sEmlPath">string that designates the path to a .EML file</param>
    ''' <returns>RxMailMessage or null if email not properly formatted</returns>
    Public Function GetEmail(ByVal sEmlPath As String) As RxMailMessage
        Dim mm As RxMailMessage = Nothing
        Using EmailStream As Stream = File.Open(sEmlPath, FileMode.Open)
            mm = GetEmail(EmailStream)
        End Using
        Return mm
    End Function
    ''' <summary>
    ''' Gets an email from the supplied Email Stream and processes it.
    ''' </summary>
    ''' <param name="sEmlPath">Stream that designates the an Email stream</param>
    ''' <returns>RxMailMessage or null if email not properly formatted</returns>
    Public Function GetEmail(ByVal EmailStream As Stream) As RxMailMessage
        EmailStreamReader = New StreamReader(EmailStream, Encoding.ASCII)
        'prepare message, set defaults as specified in RFC 2046
        'although they get normally overwritten, we have to make sure there are at least defaults
        Dim Message As New RxMailMessage()
        Message.ContentTransferEncoding = TransferEncoding.SevenBit
        Message.TransferType = "7bit"
        'convert received email into RxMailMessage
        Dim MessageMimeReturnCode As MimeEntityReturnCode = ProcessMimeEntity(Message, "")
        If MessageMimeReturnCode = MimeEntityReturnCode.bodyComplete OrElse MessageMimeReturnCode = MimeEntityReturnCode.parentBoundaryEndFound Then
            ' I've seen EML files that don't have a "To: entity but have "x-receiver:" entity set to the recipient. check and use that if need be
            If Message.[To].Count = 0 Then
                ' do something with
                Dim sTo As String = Message.Headers("x-receiver")
                If Not String.IsNullOrEmpty(sTo) Then
                    Message.[To].Add(sTo)
                End If
            End If
            ' From: maybe also but never have seen it missing
            If Message.From Is Nothing Then
                ' do something with
                Dim sFrom As String = Message.Headers("x-sender")
                If Not String.IsNullOrEmpty(sFrom) Then
                    Message.From = New MailAddress(sFrom)
                End If
            End If
            'TraceFrom("email with {0} body chars received", Message.Body.Length);
            Return Message
        End If
        Return Nothing
    End Function
    Private Sub callGetEmailWarning(ByVal warningText As String, ByVal ParamArray warningParameters As Object())
        Dim warningString As String
        Try
            warningString = String.Format(warningText, warningParameters)
        Catch generatedExceptionName As Exception
            'some strange email address can give string.Format() a problem
            warningString = warningText
        End Try
        'CallWarning("GetEmail", "", "Problem EmailNo {0}: " + warningString, messageNo);
    End Sub
    ''' <summary>
    ''' indicates the reason how a MIME entity processing has terminated
    ''' </summary>
    Private Enum MimeEntityReturnCode
        undefined = 0
        'meaning like null
        bodyComplete
        'end of message line found
        parentBoundaryStartFound
        parentBoundaryEndFound
        problem
        'received message doesn't follow MIME specification
    End Enum
    'buffer used by every ProcessMimeEntity() to store  MIME entity
    Private MimeEntitySB As New StringBuilder(100000)
    ''' <summary>
    ''' Process a MIME entity
    ''' 
    ''' A MIME entity consists of header and body.
    ''' Separator lines in the body might mark children MIME entities
    ''' </summary>
    Private Function ProcessMimeEntity(ByVal message As RxMailMessage, ByVal parentBoundaryStart As String) As MimeEntityReturnCode
        Dim hasParentBoundary As Boolean = parentBoundaryStart.Length > 0
        Dim parentBoundaryEnd As String = parentBoundaryStart + "--"
        Dim boundaryMimeReturnCode As MimeEntityReturnCode
        'some format fields are inherited from parent, only the default for
        'ContentType needs to be set here, otherwise the boundary parameter would be
        'inherited too !
        message.SetContentTypeFields("text/plain; charset=us-ascii")
        'get header
        '----------
        Dim completeHeaderField As String = Nothing
        'consists of one start line and possibly several continuation lines
        Dim response As String
        ' read header lines until empty line is found (end of header)
        While True
            If Not readMultiLine(response) Then
                callGetEmailWarning("incomplete MIME entity header received")
                'empty this message
                While readMultiLine(response)
                End While
                'System.Diagnostics.Debugger.Break()
                'didn't have a sample email to test this
                Return MimeEntityReturnCode.problem
            End If
            If response.Length < 1 Then
                'empty line found => end of header
                If completeHeaderField <> Nothing Then
                    ProcessHeaderField(message, completeHeaderField)
                    'there was only an empty header.
                Else
                End If
                Exit While
            End If
            'check if there is a parent boundary in the header (wrong format!)
            If hasParentBoundary AndAlso parentBoundaryFound(response, parentBoundaryStart, parentBoundaryEnd, boundaryMimeReturnCode) Then
                callGetEmailWarning("MIME entity header  prematurely ended by parent boundary")
                'empty this message
                While readMultiLine(response)
                End While
                System.Diagnostics.Debugger.Break()
                'didn't have a sample email to test this
                Return boundaryMimeReturnCode
            End If
            'read header field
            'one header field can extend over one start line and multiple continuation lines
            'a continuation line starts with at least 1 blank (' ') or tab
            If response(0) = " "c OrElse response(0) = ControlChars.Tab Then
                'continuation line found.
                If completeHeaderField = Nothing Then
                    callGetEmailWarning("Email header starts with continuation line")
                    'empty this message
                    While readMultiLine(response)
                    End While
                    System.Diagnostics.Debugger.Break()
                    'didn't have a sample email to test this
                    Return MimeEntityReturnCode.problem
                Else
                    ' append space, if needed, and continuation line
                    If completeHeaderField(completeHeaderField.Length - 1) <> " "c Then
                        'previous line did not end with a whitespace
                        'need to replace CRLF with a ' '
                        completeHeaderField += " "c + response.TrimStart(WhiteSpaceChars)
                    Else
                        'previous line did end with a whitespace
                        completeHeaderField += response.TrimStart(WhiteSpaceChars)
                    End If
                End If
            Else
                'a new header field line found
                If completeHeaderField = Nothing Then
                    'very first field, just copy it and then check for continuation lines
                    completeHeaderField = response
                Else
                    'new header line found
                    ProcessHeaderField(message, completeHeaderField)
                    'save the beginning of the next line
                    completeHeaderField = response
                End If
            End If
        End While
        'end while read header lines
        'process body
        '------------
        MimeEntitySB.Length = 0
        'empty StringBuilder. For speed reasons, reuse StringBuilder defined as member of class
        Dim BoundaryDelimiterLineStart As String = Nothing
        Dim isBoundaryDefined As Boolean = False
        If message.ContentType.Boundary <> Nothing Then
            isBoundaryDefined = True
            BoundaryDelimiterLineStart = "--" + message.ContentType.Boundary
        End If
        'prepare return code for the case there is no boundary in the body
        boundaryMimeReturnCode = MimeEntityReturnCode.bodyComplete
        'read body lines
        While readMultiLine(response)
            'check if there is a boundary line from this entity itself in the body
            If isBoundaryDefined AndAlso response.TrimEnd() = BoundaryDelimiterLineStart Then
                'boundary line found.
                'stop the processing here and start a delimited body processing
                Return ProcessDelimitedBody(message, BoundaryDelimiterLineStart, parentBoundaryStart, parentBoundaryEnd)
            End If
            'check if there is a parent boundary in the body
            If hasParentBoundary AndAlso parentBoundaryFound(response, parentBoundaryStart, parentBoundaryEnd, boundaryMimeReturnCode) Then
                'a parent boundary is found. Decode the content of the body received so far, then end this MIME entity
                'note that boundaryMimeReturnCode is set here, but used in the return statement
                Exit While
            End If
            'process next line
            MimeEntitySB.Append(response + CRLF)
        End While
        'a complete MIME body read
        'convert received US ASCII characters to .NET string (Unicode)
        Dim TransferEncodedMessage As String = MimeEntitySB.ToString()
        Dim isAttachmentSaved As Boolean = False
        Select Case message.ContentTransferEncoding
            Case TransferEncoding.SevenBit
                'nothing to do
                saveMessageBody(message, TransferEncodedMessage)
                Exit Select
            Case TransferEncoding.Base64
                'convert base 64 -> byte[]
                Dim bodyBytes As Byte() = System.Convert.FromBase64String(TransferEncodedMessage)
                message.ContentStream = New MemoryStream(bodyBytes, False)
                If message.MediaMainType = "text" Then
                    'convert byte[] -> string
                    message.Body = DecodeByteArryToString(bodyBytes, message.BodyEncoding)
                ElseIf message.MediaMainType = "image" OrElse message.MediaMainType = "application" Then
                    SaveAttachment(message)
                    isAttachmentSaved = True
                End If
                Exit Select
            Case TransferEncoding.QuotedPrintable
                saveMessageBody(message, QuotedPrintable.Decode(TransferEncodedMessage))
                Exit Select
            Case Else
                saveMessageBody(message, TransferEncodedMessage)
                'no need to raise a warning here, the warning was done when analising the header
                Exit Select
        End Select
        If message.ContentDisposition IsNot Nothing AndAlso message.ContentDisposition.DispositionType.ToLowerInvariant() = "attachment" AndAlso Not isAttachmentSaved Then
            SaveAttachment(message)
            isAttachmentSaved = True
        End If
        Return boundaryMimeReturnCode
    End Function
    ''' <summary>
    ''' Check if the response line received is a parent boundary 
    ''' </summary>
    Private Function parentBoundaryFound(ByVal response As String, ByVal parentBoundaryStart As String, ByVal parentBoundaryEnd As String, ByRef boundaryMimeReturnCode As MimeEntityReturnCode) As Boolean
        boundaryMimeReturnCode = MimeEntityReturnCode.undefined
        If response = Nothing OrElse response.Length < 2 OrElse response(0) <> "-"c OrElse response(1) <> "-"c Then
            'quick test: reponse doesn't start with "--", so cannot be a separator line
            Return False
        End If
        If response = parentBoundaryStart Then
            boundaryMimeReturnCode = MimeEntityReturnCode.parentBoundaryStartFound
            Return True
        ElseIf response = parentBoundaryEnd Then
            boundaryMimeReturnCode = MimeEntityReturnCode.parentBoundaryEndFound
            Return True
        End If
        Return False
    End Function
    ''' <summary>
    ''' Convert one MIME header field and update message accordingly
    ''' </summary>
    Private Sub ProcessHeaderField(ByVal message As RxMailMessage, ByVal headerField As String)
        Dim headerLineType As String
        Dim headerLineContent As String
        Dim separatorPosition As Integer = headerField.IndexOf(":"c)
        If separatorPosition < 1 Then
            ' header field type not found, skip this line
            callGetEmailWarning("character ':' missing in header format field: '{0}'", headerField)
        Else
            'process header field type
            headerLineType = headerField.Substring(0, separatorPosition).ToLowerInvariant()
            headerLineContent = headerField.Substring(separatorPosition + 1).Trim(WhiteSpaceChars)
            If headerLineType = "" OrElse headerLineContent = "" Then
                '1 of the 2 parts missing, drop the line
                Return
            End If
            ' add header line to headers
            message.Headers.Add(headerLineType, headerLineContent)
            'interpret if possible
            Select Case headerLineType
                Case "bcc"
                    AddMailAddresses(headerLineContent, message.Bcc)
                    Exit Select
                Case "cc"
                    AddMailAddresses(headerLineContent, message.CC)
                    Exit Select
                Case "content-description"
                    message.ContentDescription = headerLineContent
                    Exit Select
                Case "content-disposition"
                    'message.ContentDisposition = new ContentDisposition(headerLineContent);
                    message.SetContentDisposition(headerLineContent)
                    Exit Select
                Case "content-id"
                    message.ContentId = headerLineContent
                    Exit Select
                Case "content-transfer-encoding"
                    message.TransferType = headerLineContent
                    message.ContentTransferEncoding = ConvertToTransferEncoding(headerLineContent)
                    Exit Select
                Case "content-type"
                    message.SetContentTypeFields(headerLineContent)
                    Exit Select
                Case "date"
                    message.DeliveryDate = ConvertToDateTime(headerLineContent)
                    Exit Select
                Case "delivered-to"
                    message.DeliveredTo = ConvertToMailAddress(headerLineContent)
                    Exit Select
                Case "from"
                    Dim address As MailAddress = ConvertToMailAddress(headerLineContent)
                    If address IsNot Nothing Then
                        message.From = address
                    End If
                    Exit Select
                Case "message-id"
                    message.MessageId = headerLineContent
                    Exit Select
                Case "mime-version"
                    message.MimeVersion = headerLineContent
                    'message.BodyEncoding = new Encoding();
                    Exit Select
                Case "sender"
                    message.Sender = ConvertToMailAddress(headerLineContent)
                    Exit Select
                Case "subject"
                    message.Subject = headerLineContent
                    Exit Select
                Case "received"
                    'throw mail routing information away
                    Exit Select
                Case "reply-to"
                    message.ReplyTo = ConvertToMailAddress(headerLineContent)
                    Exit Select
                Case "return-path"
                    message.ReturnPath = ConvertToMailAddress(headerLineContent)
                    Exit Select
                Case "to"
                    AddMailAddresses(headerLineContent, message.[To])
                    Exit Select
                Case Else
                    message.UnknowHeaderlines.Add(headerField)
                    If isCollectUnknowHeaderLines Then
                        AllUnknowHeaderLines.Add(headerField)
                    End If
                    Exit Select
            End Select
        End If
    End Sub
    ''' <summary>
    ''' find individual addresses in the string and add it to address collection
    ''' </summary>
    ''' <param name="Addresses">string with possibly several email addresses</param>
    ''' <param name="AddressCollection">parsed addresses</param>
    Private Sub AddMailAddresses(ByVal Addresses As String, ByVal AddressCollection As MailAddressCollection)
        Dim adr As MailAddress
        Try
            ' I copped out on this regex - trying to figure out how to do it in just a single regex was giving me a headache
            ' just replace the comma that's inside quotes with a char (^c) thats not going to occur in a legal email name (or the 2183 RFC for that matter)
            Dim regexObj As New Regex("""[^""]*""")
            Dim mc3 As MatchCollection = regexObj.Matches(Addresses)
            For Each match As Match In mc3
                Dim sQuotedString As String = match.Value.Replace(","c, CChar(CStr(3)))
                Addresses = Addresses.Replace(match.Value, sQuotedString)
            Next
            Dim AddressSplit As String() = Addresses.Split(","c)
            For Each adrString As String In AddressSplit
                ' be sure to add the comma back if it was replaced
                adr = ConvertToMailAddress(adrString.Replace(CChar(CStr(3)), ","c))
                If adr IsNot Nothing Then
                    AddressCollection.Add(adr)
                End If
            Next
        Catch
            'didn't have a sample email to test this
            System.Diagnostics.Debugger.Break()
        End Try
    End Sub
    ''' <summary>
    ''' Tries to convert a string into an email address
    ''' </summary>
    Public Function ConvertToMailAddress(ByVal address As String) As MailAddress
        ' just remove the quotes since they are not valid in email addresses and the MailAdress parser doesn't need them. 
        ' this will handles both 
        ' ->>>    "name@host.com" 
        ' and 
        ' ->>>    "LName,FName (name@host.com)" <name@host.com>
        ' formats
        address = address.Replace("""", "")
        address = address.Trim()
        If address = "<>" Then
            'empty email address, not recognised a such by .NET
            Return Nothing
        End If
        Try
            'return new MailAddress(address.Trim(new char[] { '"' }));
            Return New MailAddress(address)
        Catch
            callGetEmailWarning("address format not recognized: '" + address.Trim() + "'")
        End Try
        Return Nothing
    End Function
    Private culture As IFormatProvider = New CultureInfo("en-US", True)
    ''' <summary>
    ''' Tries to convert string to date
    ''' If there is a run time error, the smallest possible date is returned
    ''' <example>Wed, 04 Jan 2006 07:58:08 -0800</example>
    ''' </summary>
    Public Function ConvertToDateTime(ByVal [date] As String) As DateTime
        Dim ReturnDateTime As DateTime
        Try
            'sample; 'Wed, 04 Jan 2006 07:58:08 -0800 (PST)'
            'remove day of the week before ','
            'remove date zone in '()', -800 indicates the zone already
            'remove day of week
            Dim cleanDateTime As String = [date]
            Dim DateSplit As String() = cleanDateTime.Split(CommaChars, 2)
            If DateSplit.Length > 1 Then
                cleanDateTime = DateSplit(1)
            End If
            'remove time zone (PST)
            DateSplit = cleanDateTime.Split(BracketChars)
            If DateSplit.Length > 1 Then
                cleanDateTime = DateSplit(0)
            End If
            'convert to DateTime
            If Not DateTime.TryParse(cleanDateTime, culture, DateTimeStyles.AdjustToUniversal Or DateTimeStyles.AllowWhiteSpaces, ReturnDateTime) Then
                'try just to convert the date
                Dim DateLength As Integer = cleanDateTime.IndexOf(":"c) - 3
                cleanDateTime = cleanDateTime.Substring(0, DateLength)
                If DateTime.TryParse(cleanDateTime, culture, DateTimeStyles.AdjustToUniversal Or DateTimeStyles.AllowWhiteSpaces, ReturnDateTime) Then
                    callGetEmailWarning("got only date, time format not recognised: '" + [date] + "'")
                Else
                    callGetEmailWarning("date format not recognised: '" + [date] + "'")
                    System.Diagnostics.Debugger.Break()
                    'didn't have a sample email to test this
                    Return DateTime.MinValue
                End If
            End If
        Catch
            callGetEmailWarning("date format not recognised: '" + [date] + "'")
            Return DateTime.MinValue
        End Try
        Return ReturnDateTime
    End Function
    ''' <summary>
    ''' converts TransferEncoding as defined in the RFC into a .NET TransferEncoding
    ''' 
    ''' .NET doesn't know the type "bit8". It is translated here into "bit7", which
    ''' requires the same kind of processing (none).
    ''' </summary>
    ''' <param name="TransferEncodingString"></param>
    ''' <returns></returns>
    Private Function ConvertToTransferEncoding(ByVal TransferEncodingString As String) As TransferEncoding
        ' here, "bit8" is marked as "bit7" (i.e. no transfer encoding needed)
        ' "binary" is illegal in SMTP
        ' something like "7bit" / "8bit" / "binary" / "quoted-printable" / "base64"
        Select Case TransferEncodingString.Trim().ToLowerInvariant()
            Case "7bit", "8bit"
                Return TransferEncoding.SevenBit
            Case "quoted-printable"
                Return TransferEncoding.QuotedPrintable
            Case "base64"
                Return TransferEncoding.Base64
            Case "binary"
                Throw New Exception("SMPT does not support binary transfer encoding")
            Case Else
                callGetEmailWarning("not supported content-transfer-encoding: " + TransferEncodingString)
                Return TransferEncoding.Unknown
        End Select
    End Function
    ''' <summary>
    ''' Copies the content found for the MIME entity to the RxMailMessage body and creates
    ''' a stream which can be used to create attachements, alternative views, ...
    ''' </summary>
    Private Sub saveMessageBody(ByVal message As RxMailMessage, ByVal contentString As String)
        message.Body = contentString
        Dim ascii As System.Text.Encoding = System.Text.Encoding.ASCII
        Dim bodyStream As New MemoryStream(ascii.GetBytes(contentString), 0, contentString.Length)
        message.ContentStream = bodyStream
    End Sub
    ''' <summary>
    ''' each attachement is stored in its own MIME entity and read into this entity's
    ''' ContentStream. SaveAttachment creates an attachment out of the ContentStream
    ''' and attaches it to the parent MIME entity.
    ''' </summary>
    Private Sub SaveAttachment(ByVal message As RxMailMessage)
        If message.Parent Is Nothing Then
            'didn't have a sample email to test this
            System.Diagnostics.Debugger.Break()
        Else
            Dim thisAttachment As New Attachment(message.ContentStream, message.ContentType)
            'no idea why ContentDisposition is read only. on the other hand, it is anyway redundant
            If message.ContentDisposition IsNot Nothing Then
                Dim messageContentDisposition As ContentDisposition = message.ContentDisposition
                Dim AttachmentContentDisposition As ContentDisposition = thisAttachment.ContentDisposition
                If messageContentDisposition.CreationDate > DateTime.MinValue Then
                    AttachmentContentDisposition.CreationDate = messageContentDisposition.CreationDate
                End If
                AttachmentContentDisposition.DispositionType = messageContentDisposition.DispositionType
                AttachmentContentDisposition.FileName = messageContentDisposition.FileName
                AttachmentContentDisposition.Inline = messageContentDisposition.Inline
                If messageContentDisposition.ModificationDate > DateTime.MinValue Then
                    AttachmentContentDisposition.ModificationDate = messageContentDisposition.ModificationDate
                End If
                ' see note below
                'AttachmentContentDisposition.Parameters.Clear();
                If messageContentDisposition.ReadDate > DateTime.MinValue Then
                    AttachmentContentDisposition.ReadDate = messageContentDisposition.ReadDate
                End If
                If messageContentDisposition.Size > 0 Then
                    AttachmentContentDisposition.Size = messageContentDisposition.Size
                    ' I think this is a bug. Setting the ContentDisposition values above had 
                    ' already set the Parameters collection so I got an error when the attempt 
                    ' was made to add the same parameter again
                    'foreach (string key in messageContentDisposition.Parameters.Keys)
                    '{
                    '    AttachmentContentDisposition.Parameters.Add(key, messageContentDisposition.Parameters[key]);
                    '}
                End If
            End If
            'get ContentId
            Dim contentIdString As String = message.ContentId
            If contentIdString <> Nothing Then
                thisAttachment.ContentId = RemoveBrackets(contentIdString)
            End If
            thisAttachment.TransferEncoding = message.ContentTransferEncoding
            message.Parent.Attachments.Add(thisAttachment)
        End If
    End Sub
    ''' <summary>
    ''' removes leading '&lt;' and trailing '&gt;' if both exist
    ''' </summary>
    ''' <param name="parameterString"></param>
    ''' <returns></returns>
    Private Function RemoveBrackets(ByVal parameterString As String) As String
        If parameterString = Nothing Then
            Return Nothing
        End If
        If parameterString.Length < 1 OrElse parameterString(0) <> "<"c OrElse parameterString(parameterString.Length - 1) <> ">"c Then
            System.Diagnostics.Debugger.Break()
            'didn't have a sample email to test this
            Return parameterString
        Else
            Return parameterString.Substring(1, parameterString.Length - 2)
        End If
    End Function
    Private Function ProcessDelimitedBody(ByVal message As RxMailMessage, ByVal BoundaryStart As String, ByVal parentBoundaryStart As String, ByVal parentBoundaryEnd As String) As MimeEntityReturnCode
        Dim response As String
        If BoundaryStart.Trim() = parentBoundaryStart.Trim() Then
            'Mime entity boundaries have to be unique
            callGetEmailWarning("new boundary same as parent boundary: '{0}'", parentBoundaryStart)
            'empty this message
            While readMultiLine(response)
            End While
            Return MimeEntityReturnCode.problem
        End If
        Dim ReturnCode As MimeEntityReturnCode
        Do
            'empty StringBuilder
            MimeEntitySB.Length = 0
            Dim ChildPart As RxMailMessage = message.CreateChildEntity()
            'recursively call MIME part processing
            ReturnCode = ProcessMimeEntity(ChildPart, BoundaryStart)
            If ReturnCode = MimeEntityReturnCode.problem Then
                'it seems the received email doesn't follow the MIME specification. Stop here
                Return MimeEntityReturnCode.problem
            End If
            'add the newly found child MIME part to the parent
            AddChildPartsToParent(ChildPart, message)
        Loop While ReturnCode <> MimeEntityReturnCode.parentBoundaryEndFound
        'disregard all future lines until parent boundary is found or end of complete message
        Dim boundaryMimeReturnCode As MimeEntityReturnCode
        Dim hasParentBoundary As Boolean = parentBoundaryStart.Length > 0
        While readMultiLine(response)
            If hasParentBoundary AndAlso parentBoundaryFound(response, parentBoundaryStart, parentBoundaryEnd, boundaryMimeReturnCode) Then
                Return boundaryMimeReturnCode
            End If
        End While
        Return MimeEntityReturnCode.bodyComplete
    End Function
    ''' <summary>
    ''' Add all attachments and alternative views from child to the parent
    ''' </summary>
    Private Sub AddChildPartsToParent(ByVal child As RxMailMessage, ByVal parent As RxMailMessage)
        'add the child itself to the parent
        parent.Entities.Add(child)
        'add the alternative views of the child to the parent
        If child.AlternateViews IsNot Nothing Then
            For Each childView As AlternateView In child.AlternateViews
                parent.AlternateViews.Add(childView)
            Next
        End If
        'add the body of the child as alternative view to parent
        'this should be the last view attached here, because the POP 3 MIME client
        'is supposed to display the last alternative view
        If child.MediaMainType = "text" AndAlso child.ContentStream IsNot Nothing AndAlso child.Parent.ContentType IsNot Nothing AndAlso child.Parent.ContentType.MediaType.ToLowerInvariant() = "multipart/alternative" Then
            Dim thisAlternateView As New AlternateView(child.ContentStream)
            thisAlternateView.ContentId = RemoveBrackets(child.ContentId)
            thisAlternateView.ContentType = child.ContentType
            thisAlternateView.TransferEncoding = child.ContentTransferEncoding
            parent.AlternateViews.Add(thisAlternateView)
        End If
        'add the attachments of the child to the parent
        If child.Attachments IsNot Nothing Then
            For Each childAttachment As Attachment In child.Attachments
                parent.Attachments.Add(childAttachment)
            Next
        End If
    End Sub
    ''' <summary>
    ''' Converts byte array to string, using decoding as requested
    ''' </summary>
    Public Function DecodeByteArryToString(ByVal ByteArry As Byte(), ByVal ByteEncoding As Encoding) As String
        If ByteArry Is Nothing Then
            'no bytes to convert
            Return Nothing
        End If
        Dim byteArryDecoder As Decoder
        If ByteEncoding Is Nothing Then
            'no encoding indicated. Let's try UTF7
            System.Diagnostics.Debugger.Break()
            'didn't have a sample email to test this
            byteArryDecoder = Encoding.UTF7.GetDecoder()
        Else
            byteArryDecoder = ByteEncoding.GetDecoder()
        End If
        Dim charCount As Integer = byteArryDecoder.GetCharCount(ByteArry, 0, ByteArry.Length)
        Dim bodyChars As Char() = New [Char](charCount) {}
        Dim charsDecodedCount As Integer = byteArryDecoder.GetChars(ByteArry, 0, ByteArry.Length, bodyChars, 0)
        'convert char[] to string
        Return New String(bodyChars)
    End Function
    ''' <summary>
    ''' read one line in multiline mode from the Email stream. 
    ''' </summary>
    ''' <param name="response">line received</param>
    ''' <returns>false: end of message</returns>
    ''' <returns></returns>
    Protected Function readMultiLine(ByRef response As String) As Boolean
        response = Nothing
        response = EmailStreamReader.ReadLine()
        If response = Nothing Then
            Return False
        End If
        ' if we can't read anymore from the stream then the stream is ended since this version is reading from a file and not the net
        'check for byte stuffing, i.e. if a line starts with a '.', another '.' is added, unless
        'it is the last line
        If response.Length > 0 AndAlso response(0) = "."c Then
            If response = "." Then
                'closing line found
                Return False
            End If
            'remove the first '.'
            response = response.Substring(1, response.Length - 1)
        End If
        Return True
    End Function
End Class
