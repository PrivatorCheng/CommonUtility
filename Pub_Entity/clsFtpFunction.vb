Imports System.IO
Imports System.Net
Imports System.Text

Public Class clsFtpFunction
    Public Sub UploadFile(ByVal serverIP As String, ByVal remoteFilePath As String, ByVal userID As String, ByVal passwd As String, ByVal fileContents As Byte())
        Dim request As FtpWebRequest = Nothing
        Dim response As FtpWebResponse = Nothing
        Try
            'Get the object used to communicate with the server.
            'Dim request As FtpWebRequest = CType(WebRequest.Create("ftp://www.contoso.com/test.htm"), FtpWebRequest)

            request = CType(WebRequest.Create("ftp://" & serverIP & "/" & remoteFilePath), FtpWebRequest)
            request.Method = WebRequestMethods.Ftp.UploadFile

            'This example assumes the FTP site uses anonymous logon.
            request.Credentials = New NetworkCredential(userID, passwd)

            '' Copy the contents of the file to the request stream.
            'Dim sourceStream As New StreamReader("testfile.txt")
            'Dim fileContents As Byte() = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd())
            'sourceStream.Close()

            request.ContentLength = fileContents.Length
            Dim requestStream As Stream = request.GetRequestStream()
            requestStream.Write(fileContents, 0, fileContents.Length)
            requestStream.Close()

            ' Get Response
            response = CType(request.GetResponse(), FtpWebResponse)
            Console.WriteLine("Upload File Complete, status {0}", response.StatusDescription)
            If response.StatusCode <> FtpStatusCode.ClosingData Then
                Throw New Exception("Upload file error:[" & response.StatusCode.ToString & "]" & response.StatusDescription)
            End If
            response.Close()

        Catch ex As Exception
            Throw ex
        Finally
            If Not IsNothing(request) Then request = Nothing
            If Not IsNothing(response) Then response = Nothing
        End Try
    End Sub
End Class
