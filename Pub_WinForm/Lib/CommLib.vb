Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Class CommLib

    Private SLib As Pub_Library.CommLib
    Private ELib As Pub_Entity.CommLib

    Public Sub New()
        SLib = New Pub_Library.CommLib
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

#Region "Menu TreeView Sub/Function"

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

#Region "Dynamic Class"

    Public Overloads Function DynNewForm(ByVal strClassname As String, ByVal strAssembly As String, ByVal sUserID As String) As System.Windows.Forms.Form
        Dim f As System.Windows.Forms.Form = DynFindForm(strClassname)

        If IsNothing(f) Then
            f = CType(ELib.DynNewClass(strClassname, strAssembly), System.Windows.Forms.Form)
        End If
        Return f
    End Function


    Public Function DynFindForm(ByVal strClassName As String) As System.Windows.Forms.Form
        Dim f As System.Windows.Forms.Form = Nothing
        For Each f2 As System.Windows.Forms.Form In System.Windows.Forms.Application.OpenForms
            If f2.GetType.ToString = strClassName Then
                Exit For
            End If
        Next
        Return f
    End Function

#End Region

#Region "Debug"

    Public Function GetTime() As String
        Return Now.Hour.ToString("00") & Now.Minute.ToString("00") & Now.Second.ToString("00") & Now.Millisecond.ToString("000")
    End Function

#End Region

#Region "Hierarchic Tree Function"
    'Public Function GetTreeFullFile(ByVal ProjCode As String, ByVal AnaModule As String, ByVal LineGroup As String, ByVal AnaVersion As String, ByVal FunctionName As String) As String
    '    Return ELib.GetTreeFullFile(ProjCode, AnaModule, LineGroup, AnaVersion, FunctionName)
    '    'Return GetTreeDir(ProjCode, AnaModule) & "\" & GetTreeFile(LineGroup, AnaVersion, FunctionName)
    'End Function

    'Public Function GetTreeDir(ByVal ProjCode As String, ByVal AnaModule As String, ByVal AnaSubModule As String) As String
    '    Return ELib.GetTreeDir(ProjCode, AnaModule, AnaSubModule)
    '    'Dim sDir As String
    '    'sDir = GetMySetting("DataPath")
    '    'sDir = sDir & "\Report\" & ProjCode & "\" & AnaModule
    '    'Return sDir
    'End Function

    'Public Function GetTreeFile(ByVal LineGroup As String, ByVal AnaVersion As String, ByVal FunctionName As String) As String
    '    Return ELib.GetTreeFile(LineGroup, AnaVersion, FunctionName)
    '    'Dim sFileName As String
    '    'sFileName = LineGroup & AnaVersion & FunctionName & ".txt"
    '    'Return sFileName
    'End Function

    Public Sub SaveTableToTextFile(ByVal FullFileName As String, ByVal dt As DataTable, ByVal sDilimiter As String)
        SaveTableToTextFile(FullFileName, dt, sDilimiter, False)
    End Sub

    Public Sub SaveTableToTextFile(ByVal FullFileName As String, ByVal dt As DataTable, ByVal sDilimiter As String, ByVal fAddQuote As Boolean)
        SaveTableToTextFile(FullFileName, dt, sDilimiter, False, Nothing)
    End Sub

    Public Sub SaveTableToTextFile(ByVal FullFileName As String, ByVal dt As DataTable, ByVal sDilimiter As String, ByVal fAddQuote As Boolean, ByVal sHeader As String())
        Dim sArr() As String = Nothing
        Dim s As String = ""
        Dim i As Integer, j As Integer
        i = 0

        Dim sStep As String = ""
        ReDim sArr(dt.Rows.Count + 1)
        Try
            Try
                If Not IsNothing(sHeader) AndAlso sHeader.Length > 0 Then
                    Dim iColIdx As Integer = 0
                    For Each sh As String In sHeader
                        iColIdx += 1
                        If fAddQuote Then
                            s = s & """" & "HeaderCol" & CStr(iColIdx) & """" & sDilimiter
                        Else
                            s = s & "HeaderCol" & CStr(iColIdx) & sDilimiter
                        End If
                    Next
                End If
                For Each dc As DataColumn In dt.Columns
                    Dim sh As String = ""
                    If dc.Caption = "" Then
                        sh = dc.ColumnName
                    Else
                        sh = dc.Caption
                    End If
                    If fAddQuote Then
                        s = s & """" & sh & """" & sDilimiter
                    Else
                        s = s & sh & sDilimiter
                    End If
                Next
                s = Mid(s, 1, Len(s) - Len(sDilimiter))
                's = s & vbCrLf
                sArr(0) = s
            Catch ex As Exception
                Throw New Exception("[Setting Header]" & ex.Message)
            End Try

            Try
                sStep = "S1"
                For Each dr As DataRow In dt.Rows
                    If dr.RowState <> DataRowState.Deleted Then
                        s = ""
                        sStep = "S2_" & CStr(i)
                        If Not IsNothing(sHeader) AndAlso sHeader.Length > 0 Then
                            For Each sh As String In sHeader
                                If fAddQuote Then
                                    s = s & """" & sh & """" & sDilimiter
                                Else
                                    s = s & sh & sDilimiter
                                End If
                            Next
                        End If
                        j = 0
                        sStep = "S3_" & CStr(i) & "_" & CStr(j)
                        For Each dc As DataColumn In dt.Columns
                            sStep = "S3_" & CStr(i) & "_" & CStr(j)
                            Dim sValue As String = ""
                            If dc.DataType Is Type.GetType("System.Date") OrElse dc.DataType Is Type.GetType("System.DateTime") Then
                                If Not IsNothing(dr.Item(dc)) AndAlso Not IsDBNull(dr.Item(dc)) Then
                                    sValue = CType(dr.Item(dc), Date).ToShortDateString
                                Else
                                    sValue = ""
                                End If
                            Else
                                sValue = dr.Item(dc).ToString
                            End If
                            sStep = "S4_" & CStr(i) & "_" & CStr(j) & "_" & sValue
                            If fAddQuote Then
                                s = s & """" & sValue & """" & sDilimiter
                            Else
                                s = s & sValue & sDilimiter
                            End If
                            sStep = "S5_" & CStr(i) & "_" & CStr(j) & "_" & s
                            j = j + 1
                        Next
                        sStep = "S6_" & CStr(i) & "_" & CStr(j) & "_" & s
                        s = Mid(s, 1, Len(s) - Len(sDilimiter))
                        sArr(i + 1) = s
                        's = s & vbCrLf
                        sStep = "S7_" & CStr(i) & "_" & CStr(j) & "_" & s
                        i = i + 1
                    End If
                Next
            Catch ex As Exception
                Throw New Exception("[Setting Content]" & ex.Message)
            End Try
            Try
                'WriteTxtFile(FullFileName, s)
                ELib.WriteTxtFile(FullFileName, sArr)
            Catch ex As Exception
                Throw New Exception("[WriteTextFile] " & ex.Message)
            End Try
        Catch ex As Exception
            Throw New Exception("[SaveTableToTextFile] " & "[" & FullFileName & "][" & dt.TableName & "][" & sStep & "] " & ex.Message)
        End Try
    End Sub

    Public Sub SaveDataViewToTextFile(ByVal FullFileName As String, ByVal dv As DataView)
        SaveDataViewToTextFile(FullFileName, dv, True, 1)
    End Sub

    Public Sub SaveDataViewToTextFile(ByVal FullFileName As String, ByVal dv As DataView, ByVal fHDR As Boolean, ByVal iStartCol As Integer)
        Dim sArr() As String = Nothing
        Dim s As String = ""
        Dim i As Integer, j As Integer
        i = 0
        Dim dr As DataRowView

        ReDim sArr(dv.Count + 1)
        If fHDR Then
            j = 0
            For Each dc As DataColumn In dv.Table.Columns
                j = j + 1
                If j >= iStartCol Then
                    s = s & dc.ColumnName.ToString & ","
                End If
            Next
            If s <> "" Then
                s = Mid(s, 1, Len(s) - Len(","))
                's = s & vbCrLf
            End If
            sArr(0) = s
        End If

        For i = 0 To dv.Count - 1
            dr = dv.Item(i)
            j = 0
            s = ""
            For Each dc As DataColumn In dv.Table.Columns
                j = j + 1
                If j >= iStartCol Then
                    Dim sh As String = ""
                    If dc.Caption = "" Then
                        sh = dc.ColumnName
                    Else
                        sh = dc.Caption
                    End If
                    Dim sValue As String = ""
                    If dc.DataType Is Type.GetType("System.Date") OrElse dc.DataType Is Type.GetType("System.DateTime") Then
                        If Not IsNothing(dr.Item(dc.ColumnName)) AndAlso Not IsDBNull(dr.Item(dc.ColumnName)) Then
                            sValue = CType(dr.Item(dc.ColumnName), Date).ToShortDateString
                        Else
                            sValue = ""
                        End If
                    Else
                        sValue = dr.Item(dc.ColumnName).ToString
                    End If
                    's = s & dr.Item(dc.ColumnName).ToString & ","
                    s = s & sValue & ","
                End If
            Next
            s = Mid(s, 1, Len(s) - Len(","))
            's = s & vbCrLf
            sArr(i + 1) = s
            'i = i + 1
        Next
        'WriteTxtFile(FullFileName, s)
        ELib.WriteTxtFile(FullFileName, sArr)
    End Sub
#End Region

#Region "Cursor Setting"

    Public Function LoadCursor(ByVal sFileName As String) As Cursor


        Dim b As Drawing.Bitmap = CType(Drawing.Image.FromFile(sFileName), Drawing.Bitmap)
        'b.MakeTransparent(Color::FromArgb(r, g, b)); //Put the color you want to make transparent here. May not be neccessary if you're loading an already transparent file<br/>
        Dim ptr As IntPtr = b.GetHicon()
        'Cursor^ c = gcnew System::Windows::Forms::Cursor( ptr );<br/>
        Dim c As Cursor = New Cursor(ptr)
        'this->Cursor = c; 
        Return c

        'Dim hCursor As Long
        'hCursor = LoadCursorFromFile(SLib.GetCurrDir & sFileName)
        'Call SetSystemCursor(hCursor, IDC_ARROW)
    End Function

#End Region

#Region "Process conducting"
    Public Sub KillProcessByHwnd(ByVal hwnd As Integer)
        Try
            ELib.KillProcessByHwnd(hwnd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function GetWindowProcessID(ByVal hwnd As Integer) As Integer
        Return ELib.GetWindowProcessID(hwnd)
    End Function
#End Region

#Region "Excel Function"
    Public Sub FillCopyData(ByVal dgv As DataGridView)
        FillCopyData(dgv, Nothing)
    End Sub

    Public Sub FillCopyData(ByVal dgv As DataGridView, ByVal sHeader() As String)
        Dim o As IDataObject = Nothing
        o = System.Windows.Forms.Clipboard.GetDataObject()
        'Dim oMs As IO.MemoryStream = CType(o.GetData(DataFormats.CommaSeparatedValue, True), IO.MemoryStream)
        Dim oMs2 As String = CType(o.GetData(DataFormats.Text), String)
        'If Not IsNothing(oMs) Then
        If Not IsNothing(oMs2) AndAlso oMs2 <> "" Then
            'Dim oReader As New IO.StreamReader(oMs)
            'Dim s As String = oReader.ReadToEnd
            Dim s As String = oMs2
            Dim sArr As String() = s.Split(vbLf.ToCharArray)
            Dim iDataX As Integer, iHeaderLen As Integer
            Dim iDataY As Integer
            iDataY = 0
            For Each s2 As String In sArr
                If s2 <> "" Then
                    If InStr(s2, vbCr) = 1 Then
                        s2 = ""
                    ElseIf InStr(s2, vbCr) = 0 Then
                    Else
                        s2 = Mid(s2, 1, InStr(s2, vbCr) - 1)
                    End If
                End If
                If s2 <> "" Then
                    iDataY += 1
                    'Dim sCol() As String = s2.Split(",".ToCharArray)
                    Dim sCol() As String = s2.Split(vbTab.ToCharArray)
                    If Not IsNothing(sCol) Then
                        If iDataX < sCol.Length Then iDataX = sCol.Length
                    End If
                End If
            Next
            If IsNothing(sHeader) OrElse sHeader.Length = 0 Then
                iHeaderLen = 0
            Else
                iHeaderLen = sHeader.Length
                iDataX = iDataX + sHeader.Length
            End If

            Dim sData(iDataY, iDataX) As String
            Dim j As Integer
            'For j = 0 To sArr.Length - 1
            For j = 0 To iDataY - 1
                Dim s2 As String
                Dim sCol() As String
                Dim i As Integer
                Try
                    s2 = sArr(j)
                    For i = 1 To iHeaderLen
                        sData(j + 1, i) = sHeader(i - 1)
                    Next
                    If s2 <> "" Then
                        'sCol = s2.Split(",".ToCharArray)
                        sCol = s2.Split(vbTab.ToCharArray)
                        For i = 1 To sCol.Length
                            If i <= sCol.Length Then
                                sData(j + 1, iHeaderLen + i) = sCol(i - 1)
                            End If
                        Next
                    End If
                Catch ex As Exception
                    Throw New Exception(ex.Message)
                End Try
            Next

            Dim oTbl As DataTable = Nothing
            If dgv.DataSource.GetType.ToString = (New DataView).GetType.ToString Then
                Dim dv As DataView = CType(dgv.DataSource, DataView)
                oTbl = dv.Table
                For Each dvi As DataRowView In dv
                    dvi.Delete()
                Next
            Else
                oTbl = CType(dgv.DataSource, DataTable)
                For Each dr As DataRow In oTbl.Rows
                    dr.Delete()
                Next
            End If
            Dim oArr As New Pub_Entity.ClsArrayFunction
            oArr.ToNewTable(sData, oTbl)
            'dgv.DataSource = oTbl.DefaultView
            dgv.Refresh()
        End If

    End Sub

#End Region

#Region "Recycle Can"
    'no use now
    'Public Shared Function CheckCitizenID(ByVal sID As String) As Boolean
    '    Dim checkno As Integer, X As Integer, x1 As Integer, x2 As Integer, d() As Integer, i As Integer
    '    ReDim d(10)
    '    Dim s As String = CType(sID, String)
    '    Dim t As String
    '    Dim fPidCorrect As Boolean = True

    '    If Len(s) <> 10 Then
    '        fPidCorrect = False
    '    End If
    '    s = UCase(s)
    '    t = Mid(s, 1, 1)
    '    X = Asc(t) - Asc("A") + 10
    '    If t < "I" Then
    '        x2 = X Mod 10
    '        x1 = CType((X - x2) / 10, Integer)
    '    ElseIf t = "I" Then
    '        x1 = 3
    '        x2 = 4
    '    ElseIf t < "O" Then
    '        X = X - 1
    '        x2 = X Mod 10
    '        x1 = CType((X - x2) / 10, Integer)
    '    ElseIf t = "O" Then
    '        x1 = 3
    '        x2 = 5
    '    ElseIf t < "W" Then
    '        X = X - 2
    '        x2 = X Mod 10
    '        x1 = CType((X - x2) / 10, Integer)
    '    ElseIf t = "W" Then
    '        x1 = 3
    '        x2 = 2
    '    ElseIf t = "X" Then
    '        x1 = 3
    '        x2 = 0
    '    ElseIf t = "Y" Then
    '        x1 = 3
    '        x2 = 1
    '    ElseIf t = "Z" Then
    '        x1 = 3
    '        x2 = 3
    '    End If
    '    checkno = 0
    '    For i = 1 To 9
    '        If Not IsNumeric(Mid(s, i + 1, 1)) Then
    '            fPidCorrect = False
    '        End If
    '        d(i) = CInt(Mid(s, i + 1, 1))
    '    Next i
    '    For i = 1 To 8
    '        checkno = checkno + d(i) * (9 - i)
    '    Next i
    '    checkno = checkno + x1 + 9 * x2 + d(9)
    '    If (checkno Mod 10 = 0) Then
    '        fPidCorrect = True
    '    Else
    '        fPidCorrect = False
    '    End If
    '    Return fPidCorrect
    'End Function
#End Region

End Class
