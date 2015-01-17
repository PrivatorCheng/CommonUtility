Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography

Public Class CommLib
#Region "kernel32 Function Declaration"
    <DllImport("advapi32.DLL", SetLastError:=True)> _
    Friend Shared Function LogonUser(ByVal lpszUsername As String, ByVal lpszDomain As String, _
        ByVal lpszPassword As String, ByVal dwLogonType As Integer, ByVal dwLogonProvider As Integer, _
        ByRef phToken As IntPtr) As Integer
    End Function

    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (<MarshalAs(UnmanagedType.LPStr)> ByVal lpApplicationName As String, <MarshalAs(UnmanagedType.LPStr)> ByVal lpKeyName As String, <MarshalAs(UnmanagedType.LPStr)> ByVal lpDefault As String, <MarshalAs(UnmanagedType.LPStr)> ByVal lpReturnedString As StringBuilder, ByVal nSize As UInt32, <MarshalAs(UnmanagedType.LPStr)> ByVal lpFileName As String) As UInt32
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (<MarshalAs(UnmanagedType.LPStr)> ByVal lpApplicationName As String, <MarshalAs(UnmanagedType.LPStr)> ByVal lpKeyName As String, <MarshalAs(UnmanagedType.LPStr)> ByVal lpReturnedString As StringBuilder, <MarshalAs(UnmanagedType.LPStr)> ByVal lpFileName As String) As UInt32
#End Region

#Region ".INI Read/Write"
    Public Function myGetPrivateProfileString(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As String
        'Return myGetPrivateProfileString(lpApplicationName, lpKeyName, lpDefault, lpReturnedString, nSize, GetCurrDir() & "\Ini\DFA2.INI")
        Return myGetPrivateProfileString(lpApplicationName, lpKeyName, lpDefault, lpReturnedString, nSize, "DFA2.INI")
    End Function

    Public Function myGetPrivateProfileString(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As String
        Dim d As Integer
        'Dim s As String * 256
        Dim s As New StringBuilder(256)

        d = CType(GetPrivateProfileString(lpApplicationName, lpKeyName, lpDefault, s, CType(nSize, UInteger), lpFileName), Integer)
        'myGetPrivateProfileString = Left(s.ToString, d)
        Return Left(s.ToString, d)

    End Function

    Public Function myWritePrivateProfileString(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String) As Long
        'Return myWritePrivateProfileString(lpApplicationName, lpKeyName, lpString, GetCurrDir() & "\Ini\DFA2.INI")
        Return myWritePrivateProfileString(lpApplicationName, lpKeyName, lpString, "DFA2.INI")
    End Function

    Public Function myWritePrivateProfileString(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
        Dim d As UInteger
        Dim s As New StringBuilder(256)
        s.Append(lpString)

        'd = WritePrivateProfileString(lpApplicationName, lpKeyName, lpString, lpFileName)
        d = WritePrivateProfileString(lpApplicationName, lpKeyName, s, lpFileName)

        Return d
        'myWritePrivateProfileString = d
    End Function
#End Region

#Region "Form Template Function"
    Public Function GetMySetting(ByVal sSettingName As String) As String
        Return GetMySetting("DFA Parameter", sSettingName)
    End Function

    Public Function GetMySetting(ByVal sApplication As String, ByVal sSettingName As String) As String
        Return myGetPrivateProfileString(sApplication, sSettingName, "", "", 50)
    End Function

    Public Function GetMySetting(ByVal sApplication As String, ByVal sSettingName As String, ByVal sFileName As String) As String
        Return myGetPrivateProfileString(sApplication, sSettingName, "", "", 50, sFileName)
    End Function

    Public Sub SetMySetting(ByVal sSettingName As String, ByVal sValue As String)
        'myWritePrivateProfileString("DFA Parameter", sSettingName, sValue)
        SetMySetting("DFA Parameter", sSettingName, sValue)
    End Sub

    Public Sub SetMySetting(ByVal sApplication As String, ByVal sSettingName As String, ByVal sValue As String)
        myWritePrivateProfileString(sApplication, sSettingName, sValue)
    End Sub

#End Region

#Region "Text File IO"

    Public Function ReadTxtFile(ByVal sFileName As String) As String
        Dim sLine As String = ""
        Try
            Dim sr As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(sFileName, System.Text.Encoding.Default)
            sLine = sr.ReadToEnd
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        Return sLine
    End Function

    Public Function ReadTxtFileToArray(ByVal sFileName As String) As String(,)
        Dim sArr2 As String(,) = Nothing

        Dim s As String = ReadTxtFile(sFileName)
        Dim sArr1 As String() = s.Split(Chr(13))
        Dim iCol As Integer
        If sArr1.Length > 0 Then
            Dim sArr1T() As String = sArr1(0).Split(",".ToCharArray)
            If sArr1T.Length > 0 Then
                iCol = sArr1T.Length
                Dim i As Integer, j As Integer
                For i = 1 To sArr1.Length
                    sArr1T = sArr1(i - 1).Split(",".ToCharArray)
                    If IsNothing(sArr2) Then
                        ReDim sArr2(sArr1.Length + 1, sArr1T.Length + 1)
                    ElseIf sArr1T.Length + 1 > sArr2.GetUpperBound(1) Then
                        ReDim Preserve sArr2(sArr1.Length + 1, sArr1T.Length + 1)
                    End If
                    For j = 1 To sArr1T.Length
                        sArr2(i, j) = sArr1T(j - 1)
                    Next
                Next
            End If
        End If

        Return sArr2
    End Function

    Public Sub WriteTxtFile(ByVal sFileName As String, ByVal sText As String)
        Dim sr As System.IO.StreamWriter
        sr = Nothing
        Try
            sr = System.IO.File.CreateText(sFileName)
            sr.Write(sText)
            sr.Flush()
        Catch ex As Exception
            MsgBox("無法儲存檔案, 檢查是否有寫入的權限與檔案名稱. " + ex.Message)
        Finally
            If Not sr Is Nothing Then
                sr.Close()
            End If
        End Try
    End Sub

    Public Sub WriteTxtFile(ByVal sFileName As String, ByVal sTextArr As String())
        Dim sr As System.IO.StreamWriter
        sr = Nothing
        Try
            sr = System.IO.File.CreateText(sFileName)
            For Each sText As String In sTextArr
                If Not IsNothing(sText) AndAlso sText <> "" Then
                    sr.Write(sText)
                    sr.WriteLine()
                End If
            Next
            sr.Flush()
        Catch ex As Exception
            MsgBox("無法儲存檔案, 檢查是否有寫入的權限與檔案名稱. " + ex.Message)
        Finally
            If Not sr Is Nothing Then
                sr.Close()
            End If
        End Try
    End Sub


#End Region

#Region "Array Function"

#End Region

#Region "字串函式"
    '計算字串長度(byte)
    Public Function ansiLen(ByVal s As String) As Integer
        Dim i, iAsc, iUnc As Integer
        iUnc = Len(s)
        iAsc = 0
        For i = 1 To iUnc
            If AscW(Mid(s, i, 1)) > 127 Then
                iAsc = iAsc + 2
            Else
                iAsc = iAsc + 1
            End If
        Next
        Return iAsc
    End Function

    '中文截切(左切) Byte 處理
    Public Function ansiStr(ByVal strText As String, ByVal intLength As Integer) As String
        Return ansiStr(strText, 1, intLength)
    End Function

    '中文裁切(中間切) Byte 處理
    Public Function ansiStr(ByVal strText As String, ByVal intNum As Integer, ByVal intLength As Integer) As String
        If intNum < 1 Then Throw New Exception("Start Position must bigger than or equal to 1")
        If intLength < 1 Then Throw New Exception("Length must bigger than or equal to 1")
        Dim i, iAsc, iUnc As Integer
        Dim strRt As String = ""
        iUnc = Len(strText)
        iAsc = 0
        For i = 1 To iUnc
            If AscW(Mid(strText, i, 1)) > 127 Then
                iAsc = iAsc + 2
            Else
                iAsc = iAsc + 1
            End If
            If iAsc >= intNum And iAsc < (intNum + intLength) Then
                strRt += Mid(strText, i, 1)
            ElseIf iAsc >= (intNum + intLength) Then
                Exit For
            End If
        Next
        Return strRt
    End Function

    Public Function ansiFillStr(ByVal TarGetStr As String, ByVal StrLen As Integer) As String
        Return ansiFillStr(TarGetStr, StrLen, " ")
    End Function

    Public Function ansiFillStr(ByVal TarGetStr As String, ByVal FillLen As Integer, ByVal RepStr As String) As String
        If Len(RepStr) <> 1 Then
            Throw New Exception("Replace string must be one character")
        End If
        Dim StrLength As Integer = ansiLen(TarGetStr)
        Dim i As Integer
        For i = StrLength + 1 To FillLen
            TarGetStr = TarGetStr & " "
        Next
        Return TarGetStr
    End Function

    Public Function AddQuote(ByVal s As String) As String
        Dim sTmp As String = ""
        s = s.Replace("'", "''")
        sTmp = "'" & s & "'"

        Return sTmp
    End Function

#End Region

#Region "Others"
    Public Function DynNewClass(ByVal strClassname As String, ByVal strAssembly As String) As Object
        Dim o As Object
        Try
            Dim userAssembly As System.Reflection.Assembly = System.Reflection.Assembly.Load(strAssembly)
            Dim t As Type = userAssembly.GetType(strClassname, True)
            o = Activator.CreateInstance(t, True)
        Catch ex As Exception
            Dim sMsg As String = ex.Message
            Throw New Exception(sMsg)
            'MsgBox(sMsg, MsgBoxStyle.Question)
            'Return Nothing
            'Exit Function
        End Try
        Return o
    End Function

    Public Function GetCurrDir() As String
        Return My.Application.Info.DirectoryPath
    End Function

    Public Sub ClearDir(ByVal oDir As System.IO.DirectoryInfo)
        Dim oSubDirArr As System.IO.DirectoryInfo() = oDir.GetDirectories
        For Each oSubDir As System.IO.DirectoryInfo In oSubDirArr
            ClearDir(oSubDir)
            ''Dim oFileArr As System.IO.FileInfo() = oSubDir.GetFiles
            ''For Each oFile As System.IO.FileInfo In oFileArr
            ''    oFile.Delete()
            ''Next
            'oSubDir.Delete()
        Next
        Dim oFileArr2 As System.IO.FileInfo() = oDir.GetFiles
        For Each oFile2 As System.IO.FileInfo In oFileArr2
            oFile2.Delete()
        Next
        oDir.Delete()
    End Sub

    Public Function IsYrQtr(ByVal s As String) As Boolean
        Dim b As Boolean = True
        If Len(s) <> 7 OrElse Not IsNumeric(Left(s, 4)) OrElse Not IsNumeric(Left(s, 2)) _
        OrElse Mid(s, 5, 1) <> "Q" OrElse CType(Left(s, 2), Integer) < 4 Then
            b = False
        End If
        Return b
    End Function

    Public Function IsBiggerYrQtr(ByVal sBgn As String, ByVal sEnd As String) As Boolean
        Dim b As Boolean = False
        If CType(Microsoft.VisualBasic.Left(sEnd, 4), Integer) < CType(Microsoft.VisualBasic.Left(sBgn, 4), Integer) OrElse _
        (CType(Microsoft.VisualBasic.Left(sEnd, 4), Integer) = CType(Microsoft.VisualBasic.Left(sBgn, 4), Integer) AndAlso _
        CType(Microsoft.VisualBasic.Right(sEnd, 2), Integer) < CType(Microsoft.VisualBasic.Right(sBgn, 2), Integer)) Then
            b = True
        End If
        Return b
    End Function

    Public Function YrQtrAdd(ByVal sYrQtr As String, ByVal iAdd As Integer) As String
        Dim sAdd As String = ""

        If Not IsYrQtr(sYrQtr) Then
            sAdd = ""
        Else
            Dim iYear As Integer = CType(Left(sYrQtr, 4), Integer)
            Dim iMonth As Integer = CType(Right(sYrQtr, 2), Integer)
            iMonth += iAdd
            Do While iMonth > 4
                iMonth -= 4
                iYear += 1
            Loop
            sAdd = iYear.ToString("0000") & "Q" & iMonth.ToString("00")
        End If
        Return sAdd
    End Function

#End Region

#Region "中文函式"
    Public Function TransStrChs(ByVal strSrc As String) As String
        Return TransStr(strSrc, VbStrConv.SimplifiedChinese, 2052)
    End Function

    Private Function TransStr(ByVal strSrc As String, ByVal type As VbStrConv, ByVal localeId As Integer) As String
        Return StrConv(strSrc, type, localeId)
    End Function

    Public Function ConvertBig5(ByVal strUtf As String) As String
        Dim utf81 As Encoding = Encoding.GetEncoding("utf-8")
        Dim big51 As Encoding = Encoding.GetEncoding("big5")
        Dim strUtf81 As Byte() = utf81.GetBytes(strUtf.Trim())
        Dim strBig51 As Byte() = Encoding.Convert(utf81, big51, strUtf81)

        Dim big5Chars1 As Char() = New Char(big51.GetCharCount(strBig51, 0, strBig51.Length) - 1) {}
        big51.GetChars(strBig51, 0, strBig51.Length, big5Chars1, 0)
        Dim tempString1 As New String(big5Chars1)
        Return tempString1
    End Function

#End Region

    ''' <summary>
    ''' 用指定的登錄事件來源，將有指定訊息文字的錯誤、警告、資訊、成功稽核或失敗稽核項目寫入事件記錄檔。
    ''' </summary>
    ''' <param name="sSrc">將應用程式登錄在指定電腦上的來源</param>
    ''' <param name="sDesc">要寫入事件記錄檔的字串</param>
    ''' <param name="t">其中一個 EventLogEntryType 值</param>
    ''' <remarks></remarks>
    Public Sub WriteEventLog(ByVal sSrc As String, ByVal sDesc As String, ByVal t As Diagnostics.EventLogEntryType)
        System.Diagnostics.EventLog.WriteEntry(sSrc, sDesc, t)
    End Sub

    Public Sub WriteDataSetXMLFile(ByVal sFile As String, ByVal ds As DataSet)
        Dim fr As New System.IO.FileStream(sFile, IO.FileMode.Create)
        Dim xtw As New System.Xml.XmlTextWriter(fr, System.Text.UnicodeEncoding.Unicode)
        xtw.WriteProcessingInstruction("xml", "version='1.0'")
        ds.WriteXml(xtw)
    End Sub

    Public Function GetMD5Str(ByVal sTarget As String) As String
        Dim o As New MD5CryptoServiceProvider
        Dim bOutput As Byte() = o.ComputeHash(Encoding.Default.GetBytes(sTarget))

        Return BitConverter.ToString(bOutput).Replace("-", "")
    End Function

    Public Function TextToByte(ByVal sText As String) As Byte()
        Dim b As Byte()
        b = System.Text.Encoding.Default.GetBytes(sText)
        Return b
    End Function

    Public Shared Function GetBytes(ByVal sText As String) As Byte()
        Dim b As Byte()
        b = System.Text.Encoding.Default.GetBytes(sText)
        Return b
    End Function

    Public Shared Function GetByteCount(ByVal sText As String) As Integer
        Dim b As Integer
        b = System.Text.Encoding.Default.GetByteCount(sText)

        Return b
    End Function

End Class
