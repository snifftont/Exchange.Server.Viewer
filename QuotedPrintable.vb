' POP3 Quoted Printable
' =====================
'
' copyright by Peter Huber, Singapore, 2006
' this code is provided as is, bugs are probable, free for any use, no responsiblity accepted :-)
'
' based on QuotedPrintable Class from ASP emporium, http://www.aspemporium.com/classes.aspx?cid=6
Imports System
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Security

''' <summary>
''' <para>
''' Robust and fast implementation of Quoted Printable
''' Multipart Internet Mail Encoding (MIME) which encodes every 
''' character, not just "special characters" for transmission over SMTP.
''' </para>
''' <para>
''' More information on the quoted-printable encoding can be found
''' here: http://www.freesoft.org/CIE/RFC/1521/6.htm
''' </para>
''' </summary>
''' <remarks>
''' <para>
''' detailed in: RFC 1521
''' </para>
''' <para>
''' more info: http://www.freesoft.org/CIE/RFC/1521/6.htm
''' </para>
''' <para>
''' The QuotedPrintable class encodes and decodes strings and files
''' that either were encoded or need encoded in the Quoted-Printable
''' MIME encoding for Internet mail. The encoding methods of the class
''' use pointers wherever possible to guarantee the fastest possible 
''' encoding times for any size file or string. The decoding methods 
''' use only the .NET framework classes.
''' </para>
''' <para>
''' The Quoted-Printable implementation
''' is robust which means it encodes every character to ensure that the
''' information is decoded properly regardless of machine or underlying
''' operating system or protocol implementation. The decode can recognize
''' robust encodings as well as minimal encodings that only encode special
''' characters and any implementation in between. Internally, the
''' class uses a regular expression replace pattern to decode a quoted-
''' printable string or file.
''' </para>
''' </remarks>
''' <example>
''' This example shows how to quoted-printable encode an html file and then
''' decode it.
''' <code>
''' string encoded = QuotedPrintable.EncodeFile(
''' 	@"C:\WEBS\wwwroot\index.html"
''' 	);
''' 
''' string decoded = QuotedPrintable.Decode(encoded);
''' 
''' Console.WriteLine(decoded);
''' </code>
''' </example>
Class QuotedPrintable
    Private Sub New()
    End Sub
    ''' <summary>
    ''' Gets the maximum number of characters per quoted-printable
    ''' line as defined in the RFC minus 1 to allow for the =
    ''' character (soft line break).
    ''' </summary>
    ''' <remarks>
    ''' (Soft Line Breaks): The Quoted-Printable encoding REQUIRES 
    ''' that encoded lines be no more than 76 characters long. If 
    ''' longer lines are to be encoded with the Quoted-Printable 
    ''' encoding, 'soft' line breaks must be used. An equal sign 
    ''' as the last character on a encoded line indicates such a 
    ''' non-significant ('soft') line break in the encoded text.
    ''' </remarks>
    Public Const RFC_1521_MAX_CHARS_PER_LINE As Integer = 75
    Shared Function HexDecoderEvaluator(ByVal m As Match) As String
        Dim hex As String = m.Groups(2).Value
        Dim iHex As Integer = Convert.ToInt32(hex, 16)
        Dim c As Char = CChar(CStr(iHex))
        Return c.ToString()
    End Function
    Shared Function HexDecoder(ByVal line As String) As String
        If line = Nothing Then
            Throw New ArgumentNullException()
        End If
        'parse looking for =XX where XX is hexadecimal
        Dim re As New Regex("(\=([0-9A-F][0-9A-F]))", RegexOptions.IgnoreCase)
        Return re.Replace(line, New MatchEvaluator(AddressOf HexDecoderEvaluator))
    End Function
    ''' <summary>
    ''' decodes an entire file's contents into plain text that 
    ''' was encoded with quoted-printable.
    ''' </summary>
    ''' <param name="filepath">
    ''' The path to the quoted-printable encoded file to decode.
    ''' </param>
    ''' <returns>The decoded string.</returns>
    ''' <exception cref="ObjectDisposedException">
    ''' A problem occurred while attempting to decode the 
    ''' encoded string.
    ''' </exception>
    ''' <exception cref="OutOfMemoryException">
    ''' There is insufficient memory to allocate a buffer for the
    ''' returned string. 
    ''' </exception>
    ''' <exception cref="ArgumentNullException">
    ''' A string is passed in as a null reference.
    ''' </exception>
    ''' <exception cref="IOException">
    ''' An I/O error occurs, such as the stream being closed.
    ''' </exception>  
    ''' <exception cref="FileNotFoundException">
    ''' The file was not found.
    ''' </exception>
    ''' <exception cref="SecurityException">
    ''' The caller does not have the required permission to open
    ''' the file specified in filepath.
    ''' </exception>
    ''' <exception cref="UnauthorizedAccessException">
    ''' filepath is read-only or a directory.
    ''' </exception>
    ''' <remarks>
    ''' Decodes a quoted-printable encoded file into a string
    ''' of unencoded text of any size.
    ''' </remarks>
    Public Shared Function DecodeFile(ByVal filepath As String) As String
        If filepath = Nothing Then
            Throw New ArgumentNullException()
        End If
        Dim decodedHtml As String = "", line As String
        Dim f As New FileInfo(filepath)
        If Not f.Exists Then
            Throw New FileNotFoundException()
        End If
        Dim sr As StreamReader = f.OpenText()
        Try
            While (line = sr.ReadLine()) <> Nothing
                decodedHtml += Decode(line)
            End While
            Return decodedHtml
        Finally
            sr.Close()
            sr = Nothing
            f = Nothing
        End Try
    End Function
    ''' <summary>
    ''' Decodes a Quoted-Printable string of any size into 
    ''' it's original text.
    ''' </summary>
    ''' <param name="encoded">
    ''' The encoded string to decode.
    ''' </param>
    ''' <returns>The decoded string.</returns>
    ''' <exception cref="ArgumentNullException">
    ''' A string is passed in as a null reference.
    ''' </exception>
    ''' <remarks>
    ''' Decodes a quoted-printable encoded string into a string
    ''' of unencoded text of any size.
    ''' </remarks>
    Public Shared Function Decode(ByVal encoded As String) As String
        If encoded = Nothing Then
            Throw New ArgumentNullException()
        End If
        Dim line As String
        Dim sw As New StringWriter()
        Dim sr As New StringReader(encoded)
        Try
            While (line = sr.ReadLine() <> Nothing)
                If line.EndsWith("=") Then
                    sw.Write(HexDecoder(line.Substring(0, line.Length - 1)))
                Else
                    sw.WriteLine(HexDecoder(line))
                End If
                sw.Flush()
            End While
            Return sw.ToString()
        Finally
            sw.Close()
            sr.Close()
            sw = Nothing
            sr = Nothing
        End Try
    End Function
End Class

