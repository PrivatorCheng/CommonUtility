Imports System.IO
Imports System.IO.Compression
Imports System.Security.Principal

Public Class clsFileHandler
    Public Sub CopyFile(ByVal localFilePath As String, ByVal remoteFilePath As String, ByVal userID As String, ByVal userPasswd As String, ByVal domainName As String)
        Dim admin_token As IntPtr
        Dim wid_current As WindowsIdentity = WindowsIdentity.GetCurrent()
        Dim wid_admin As WindowsIdentity = Nothing
        Dim wic As WindowsImpersonationContext = Nothing
        Try

            If CommLib.LogonUser(userID, domainName, userPasswd, 9, 0, admin_token) <> 0 Then
                wid_admin = New WindowsIdentity(admin_token)
                wic = wid_admin.Impersonate()
                If Not System.IO.File.Exists(localFilePath) Then
                    Throw New Exception("File " & localFilePath & " not found!")
                End If
                File.Copy(localFilePath, remoteFilePath, True)
            Else
                Throw New Exception("Copy Failed")
            End If
        Catch ex As Exception
            Throw New Exception("Copy file error: " + ex.ToString)
        Finally
            If Not IsNothing(wic) Then
                wic.Undo()
            End If
        End Try
    End Sub

    Public Function Read(ByVal fsSource As FileStream) As Byte()
        Dim bytes() As Byte = New Byte((CType(fsSource.Length, Integer)) - 1) {}
        Dim numBytesToRead As Integer = CType(fsSource.Length, Integer)
        Dim numBytesRead As Integer = 0

        While (numBytesToRead > 0)
            ' Read may return anything from 0 to numBytesToRead.
            Dim n As Integer = fsSource.Read(bytes, numBytesRead, _
                numBytesToRead)
            ' Break when the end of the file is reached.
            If (n = 0) Then
                Exit While
            End If
            numBytesRead = (numBytesRead + n)
            numBytesToRead = (numBytesToRead - n)

        End While
        numBytesToRead = bytes.Length

        Return bytes
    End Function

    Public Sub Zip(ByVal srcPath As String, ByVal targetPath As String)
        If File.Exists(targetPath) Then
            File.Delete(targetPath)
        End If
        'ZipFile.CreateFromDirectory(srcPath, targetPath)

        Dim srcInfo As New FileInfo(srcPath)
        Dim fileName As String = srcInfo.Name
        Using archive As ZipArchive = ZipFile.Open(targetPath, ZipArchiveMode.Create)
            archive.CreateEntryFromFile(srcPath, fileName, CompressionLevel.Fastest)
        End Using



    End Sub
End Class
