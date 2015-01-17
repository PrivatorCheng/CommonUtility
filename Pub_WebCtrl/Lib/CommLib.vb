Imports System.Text
Imports System.Runtime.InteropServices

Public Class CommLib

    'Private SLib As DFA2_PubLib.CommLib
    Private ELib As Pub_Entity.CommLib

    Public Sub New()
        'SLib = New DFA2_PubLib.CommLib
        ELib = New Pub_Entity.CommLib
    End Sub

#Region "kernel32 Function Declaration"
    Const IDC_ARROW As Long = 32512&

    Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
    Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long ' modified
    'Declare Function LoadCursorFromFile Lib "user32" (ByVal lpFileName As String) As Long
    'Declare Function LoadCursor Lib "user32" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long ' modified
    Declare Function SetSystemCursor Lib "user32" (ByVal hcur As Long, ByVal id As Long) As Long
    'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (<MarshalAs(UnmanagedType.LPStr)> ByVal lpApplicationName As String, <MarshalAs(UnmanagedType.LPStr)> ByVal lpKeyName As String, <MarshalAs(UnmanagedType.LPStr)> ByVal lpDefault As String, <MarshalAs(UnmanagedType.LPStr)> ByVal lpReturnedString As StringBuilder, ByVal nSize As UInt32, <MarshalAs(UnmanagedType.LPStr)> ByVal lpFileName As String) As UInt32
    'Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (<MarshalAs(UnmanagedType.LPStr)> ByVal lpApplicationName As String, <MarshalAs(UnmanagedType.LPStr)> ByVal lpKeyName As String, <MarshalAs(UnmanagedType.LPStr)> ByVal lpReturnedString As StringBuilder, <MarshalAs(UnmanagedType.LPStr)> ByVal lpFileName As String) As UInt32
    'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer

    ' Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Integer, ByRef lpdwProcessId As IntPtr) As IntPtr

#End Region

#Region ".INI Read/Write"
    Public Function myGetPrivateProfileString(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As String
        Return ELib.myGetPrivateProfileString(lpApplicationName, lpKeyName, lpDefault, lpReturnedString, nSize)
        'Return myGetPrivateProfileString(lpApplicationName, lpKeyName, lpDefault, lpReturnedString, nSize, SLib.GetCurrDir & "\Ini\DFA2.INI")
    End Function

    Public Function myGetPrivateProfileString(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As String
        Return ELib.myGetPrivateProfileString(lpApplicationName, lpKeyName, lpDefault, lpReturnedString, nSize, lpFileName)
        'Dim d As Integer
        ''Dim s As String * 256
        'Dim s As New StringBuilder(256)
        'd = CType(GetPrivateProfileString(lpApplicationName, lpKeyName, lpDefault, s, CType(nSize, UInteger), lpFileName), Integer)
        'myGetPrivateProfileString = Left(s.ToString, d)
    End Function

    Public Function myWritePrivateProfileString(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String) As Long
        Return ELib.myWritePrivateProfileString(lpApplicationName, lpKeyName, lpString)
        'Return myWritePrivateProfileString(lpApplicationName, lpKeyName, lpString, SLib.GetCurrDir & "\Ini\DFA2.INI")
    End Function

    Public Function myWritePrivateProfileString(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
        Return ELib.myWritePrivateProfileString(lpApplicationName, lpKeyName, lpString, lpFileName)
        'Dim d As UInteger
        'Dim s As New StringBuilder(256)
        's.Append(lpString)
        ''d = WritePrivateProfileString(lpApplicationName, lpKeyName, lpString, lpFileName)
        'd = WritePrivateProfileString(lpApplicationName, lpKeyName, s, lpFileName)
        'myWritePrivateProfileString = d
    End Function
#End Region

#Region "字串函式"
    '計算字串長度(byte)
    Public Function ansiLen(ByVal s As String) As Integer
        Return ELib.ansiLen(s)
    End Function

    '中文截切(左切) Byte 處理
    Public Function ansiStr(ByVal strText As String, ByVal intLength As Integer) As String
        Return ELib.ansiStr(strText, intLength)
    End Function

    '中文裁切(中間切) Byte 處理
    Public Function ansiStr(ByVal strText As String, ByVal intNum As Integer, ByVal intLength As Integer) As String
        Return ELib.ansiStr(strText, intNum, intLength)
    End Function

    Public Function ansiFillStr(ByVal TarGetStr As String, ByVal StrLen As Integer) As String
        Return ELib.ansiFillStr(TarGetStr, StrLen)
    End Function

    Public Function ansiFillStr(ByVal TarGetStr As String, ByVal FillLen As Integer, ByVal RepStr As String) As String
        Return ELib.ansiFillStr(TarGetStr, FillLen, RepStr)
    End Function

#End Region

#Region "日期函式"
    ''' <summary>
    ''' 傳入YYYYQQQ,算出當季最後一天之DATE值
    ''' 
    ''' </summary>
    ''' <param name="YrQtr">EX: "2009Q04" string </param>
    ''' <returns>date:2009/12/31</returns>
    ''' <remarks></remarks>
    Public Function GetQtrLastDay(ByVal YrQtr As String) As Date
        Dim dYy As Integer = CType(Left(YrQtr, 4), Integer)
        Dim dMm As Integer = (CType(Right(YrQtr, 2), Integer) - 1) * 3 + 1
        Select Case CInt(Right(YrQtr, 2))
            Case 1
            Case 2
            Case 3
            Case 4
        End Select
        Dim dDate As Date = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Quarter, 1, New Date(dYy, dMm, 1)))
        Return dDate

    End Function

#End Region

#Region "Text File IO"

    Public Function ReadTxtFile(ByVal sFileName As String) As String
        Return ELib.ReadTxtFile(sFileName)
        'Dim sLine As String = ""
        'Try
        '    Dim sr As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(sFileName, System.Text.Encoding.Default)
        '    sLine = sr.ReadToEnd
        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'End Try

        'Return sLine
    End Function

    Public Sub WriteTxtFile(ByVal sFileName As String, ByVal sText As String)
        ELib.WriteTxtFile(sFileName, sText)
        'Dim sr As System.IO.StreamWriter
        'sr = Nothing
        'Try
        '    sr = System.IO.File.CreateText(sFileName)
        '    sr.Write(sText)
        '    sr.Flush()
        'Catch ex As Exception
        '    MsgBox("無法儲存檔案, 檢查是否有寫入的權限與檔案名稱. " + ex.Message)
        'Finally
        '    If Not sr Is Nothing Then
        '        sr.Close()
        '    End If
        'End Try
    End Sub

#End Region

#Region "Form Template Function"

    Public Sub SetMySetting(ByVal sSettingName As String, ByVal sValue As String)
        ELib.SetMySetting(sSettingName, sValue)
        'myWritePrivateProfileString("DFA Parameter", sSettingName, sValue)
        ''DFA2_PubControl.My.Settings(sSettingName) = sValue
    End Sub

    Public Function GetMySetting(ByVal sSettingName As String) As String
        Return ELib.GetMySetting(sSettingName)
        'Return myGetPrivateProfileString("DFA Parameter", sSettingName, "", "", 50)
        ''DFA2_PubControl.My.Settings(sSettingName) = sValue
    End Function

    Public Sub ClearProjectSetting()
        SetMySetting("ProjCode", "")
        SetMySetting("ProjName", "")
        SetMySetting("UnderExpBgn", "")
        SetMySetting("UnderExpEnd", "")
        SetMySetting("FinaExpBgn", "")
        SetMySetting("FinaExpEnd", "")
        SetMySetting("ProjBgn", "")
        SetMySetting("ProjEnd", "")
    End Sub

#End Region

#Region "Debug"

    Public Function GetTime() As String
        Return Now.Hour.ToString("00") & Now.Minute.ToString("00") & Now.Second.ToString("00") & Now.Millisecond.ToString("000")
    End Function

#End Region

#Region "Cursor Setting"

    'Public Function LoadCursor(ByVal sFileName As String) As Cursor
    '    Dim b As Drawing.Bitmap = CType(Drawing.Image.FromFile(sFileName), Drawing.Bitmap)
    '    'b.MakeTransparent(Color::FromArgb(r, g, b)); //Put the color you want to make transparent here. May not be neccessary if you're loading an already transparent file<br/>
    '    Dim ptr As IntPtr = b.GetHicon()
    '    'Cursor^ c = gcnew System::Windows::Forms::Cursor( ptr );<br/>
    '    Dim c As Cursor = New Cursor(ptr)
    '    'this->Cursor = c; 
    '    Return c

    '    'Dim hCursor As Long
    '    'hCursor = LoadCursorFromFile(SLib.GetCurrDir & sFileName)
    '    'Call SetSystemCursor(hCursor, IDC_ARROW)
    'End Function

#End Region

#Region "Process conducting"
    Public Sub KillProcessByHwnd(ByVal hwnd As Integer)

        Try
            ''Dim iProcessID As IntPtr
            ''GetWindowThreadProcessId(hwnd, iProcessID)

            ''Dim p As Process = Process.GetProcessById(iProcessID.ToInt32())
            ''p.Kill()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function GetWindowProcessID(ByVal hwnd As Integer) As Integer
        ''Dim iProcessID As IntPtr
        ''GetWindowThreadProcessId(hwnd, iProcessID)
        ''Return iProcessID.ToInt32
        Return 1
    End Function
#End Region

#Region "Recycle Can"
#End Region

End Class
