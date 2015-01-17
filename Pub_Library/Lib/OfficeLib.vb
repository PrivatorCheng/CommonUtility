Public Class OfficeLib

#Region " Excel Application"
    ''creat by scott 20100630
    ''Public Function getExcelApplication() As Excel.Application
    ''    '當一個執行階段錯誤產生時，程式控制立刻到發生錯誤陳述式接下去的陳述式，而繼續執行下去
    ''    On Error Resume Next
    ''    '#一部電腦僅執行一個Excel Application, 就算中突開啟Excel也不會影響程式執行
    ''    '#在工作管理員中只會看見一個EXCEL.exe在執行，不會浪費電腦資源
    ''    '#引用正在執行的Excel Application
    ''    Dim xlApp As Excel.Application  '= New Excel.Application
    ''    xlApp = CType(GetObject(, "Excel.Application"), Excel.Application)
    ''    '#若發生錯誤表示電腦沒有Excel正在執行，需重新建立一個新的應用程式
    ''    If Err.Number() <> 0 Then
    ''        Err.Clear()
    ''        '#執行一個新的Excel Application
    ''        xlApp = CType(CreateObject("Excel.Application"), Excel.Application)
    ''        If Err.Number() <> 0 Then
    ''            MsgBox("Not install Microsoft Excel ")
    ''            Return Nothing
    ''        End If
    ''    End If
    ''    Err.Clear()
    ''    Return xlApp
    ''End Function
    'Public Sub ReleaseExcel(ByRef Reference As Excel.Application)
    '    Try
    '        Dim iHwnd As Integer = Reference.Hwnd
    '        Do Until _
    '         System.Runtime.InteropServices.Marshal.ReleaseComObject(Reference) <= 0
    '        Loop
    '        Me.KillProcessByHwnd(iHwnd)
    '    Catch
    '    Finally
    '        Reference = Nothing
    '    End Try
    'End Sub
    'Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Integer, ByRef lpdwProcessId As IntPtr) As IntPtr

    'Public Sub KillProcessByHwnd(ByVal hwnd As Integer)
    '    Try
    '        Dim iProcessID As IntPtr
    '        GetWindowThreadProcessId(hwnd, iProcessID)

    '        Dim p As Process = Process.GetProcessById(iProcessID.ToInt32())
    '        p.Kill()
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try
    'End Sub
    'Public Function GetWindowProcessID(ByVal hwnd As Integer) As Integer
    '    Dim iProcessID As IntPtr
    '    GetWindowThreadProcessId(hwnd, iProcessID)

    '    Return iProcessID.ToInt32
    'End Function
#End Region

#Region "Access Function"
    'Public Sub AccessLinkTable(ByVal srcMdb As String, ByVal srcTbl() As String)
    '    AccessLinkTable(srcMdb, srcTbl, MainDBName(True))
    'End Sub

    'Public Sub AccessLinkTable(ByVal srcMdb As String, ByVal srcTbl() As String, ByVal targetMdb As String, ByVal TType As Access.AcDataTransferType, ByVal OType As Access.AcObjectType)
    '    Dim oApp As New Access.Application
    '    With oApp
    '        .OpenCurrentDatabase(targetMdb) ' 開啟 MDB 資料庫檔案

    '        For Each t As String In srcTbl
    '            Try
    '                .DoCmd.DeleteObject(Access.AcObjectType.acTable, t)
    '            Catch ex As Exception
    '            End Try
    '            .DoCmd.TransferDatabase(TType, "Microsoft Access", srcMdb, OType, t, t)
    '            '.ImportXML(strXML, 1) ' 建立資料結構 ( Table ) , 並匯入 XML 資料 ( 1 = acStructureAndData )
    '        Next
    '        .CloseCurrentDatabase() ' 關閉資料庫
    '        .Quit() ' 關閉 Access 執行個體
    '    End With
    'End Sub

    'Public Sub AccessLinkTable(ByVal srcMdb As String, ByVal srcTbl() As String, ByVal targetMdb As String, ByVal OType As Access.AcObjectType)
    '    AccessLinkTable(srcMdb, srcTbl, targetMdb, Access.AcDataTransferType.acLink, OType)
    'End Sub

    'Public Sub AccessLinkTable(ByVal srcMdb As String, ByVal srcTbl() As String, ByVal targetMdb As String)
    '    AccessLinkTable(srcMdb, srcTbl, targetMdb, Access.AcObjectType.acTable)
    'End Sub

    'Public Sub AccessLinkTable(ByVal sServer As String, ByVal sUser As String, ByVal sPwd As String, ByVal sDB As String, ByVal targetMdb As String, ByVal srcTable As String, ByVal targetTable As String)
    '    '    DoCmd.TransferDatabase acLink, "ODBC", "ODBC;Driver={SQL Server};Server=X3650;Database=DFA2_DataTrans;Uid=sa;Pwd=1234", acTable, "HIS_Cmn_Premium", "HIS_Cmn_Premium"
    '    Dim sLink As String = "ODBC;Driver={SQL Server};Server=" & sServer & ";Database=" & sDB & ";Uid=" & sUser & ";Pwd=" & sPwd
    '    AccessLinkTable(sLink, targetMdb, srcTable, targetTable)
    'End Sub

    'Public Sub AccessLinkTable(ByVal ODBCLinkString As String, ByVal targetMdb As String, ByVal srcTable As String, ByVal targetTable As String)
    '    Dim oApp As New Access.Application
    '    With oApp
    '        .OpenCurrentDatabase(targetMdb) ' 開啟 MDB 資料庫檔案
    '        Try
    '            .DoCmd.DeleteObject(Access.AcObjectType.acTable, targetTable)
    '        Catch ex As Exception
    '        End Try
    '        .DoCmd.TransferDatabase(Access.AcDataTransferType.acLink, "ODBC", ODBCLinkString, Access.AcObjectType.acTable, srcTable, targetTable)

    '        .CloseCurrentDatabase() ' 關閉資料庫
    '        .Quit() ' 關閉 Access 執行個體
    '    End With
    'End Sub

    'Public Sub AccessTransferDataBase(ByVal oApp As Access.Application, ByVal TType As Access.AcDataTransferType, ByVal OType As Access.AcObjectType, ByVal srcTbl() As String, ByVal targetMdb As String)
    '    With oApp
    '        For Each t As String In srcTbl
    '            Try
    '                .DoCmd.DeleteObject(Access.AcObjectType.acTable, t)
    '            Catch ex As Exception
    '            End Try
    '            .DoCmd.TransferDatabase(TType, "Microsoft Access", targetMdb, OType, t, t)
    '        Next
    '        '.CloseCurrentDatabase() ' 關閉資料庫
    '        '.Quit() ' 關閉 Access 執行個體
    '    End With
    'End Sub

    'Public Sub AccessRemoveTable(ByVal srcTbl() As String, ByVal targetMdb As String, ByVal OType As Access.AcObjectType)
    '    Dim oApp As New Access.Application
    '    With oApp
    '        .OpenCurrentDatabase(targetMdb) ' 開啟 MDB 資料庫檔案
    '        For Each t As String In srcTbl
    '            Try
    '                .DoCmd.DeleteObject(OType, t)
    '            Catch ex As Exception
    '            End Try
    '        Next
    '        .CloseCurrentDatabase() ' 關閉資料庫
    '        .Quit() ' 關閉 Access 執行個體
    '    End With
    'End Sub

    'Public Sub AccessRemoveTable(ByVal srcTbl() As String, ByVal targetMdb As String)
    '    AccessRemoveTable(srcTbl, targetMdb, Access.AcObjectType.acTable)
    'End Sub

    'Public Sub AccessOpenQuery(ByVal tarQry() As String, ByVal targetMdb As String)
    '    Dim oApp As New Access.Application
    '    Dim sMsg As String = ""
    '    With oApp
    '        .OpenCurrentDatabase(targetMdb) ' 開啟 MDB 資料庫檔案
    '        For Each t As String In tarQry
    '            Try
    '                .DoCmd.OpenQuery(t)
    '            Catch ex As Exception
    '                'Dim sMsg As String = ex.Message
    '                sMsg = sMsg & "[Open Query Error][" & t & "][" & ex.Message & "]" & vbCr
    '            End Try
    '        Next
    '        .CloseCurrentDatabase() ' 關閉資料庫
    '        .Quit() ' 關閉 Access 執行個體
    '    End With
    '    If sMsg <> "" Then
    '        Throw New Exception(sMsg)
    '    End If
    'End Sub

    ''Public Sub AccessRunFunction(ByVal targetMdb As String, ByVal sFuncName As String, ByVal Para() As Object)
    ''    Dim oApp As New Access.Application
    ''    Dim sErr As String = ""
    ''    With oApp
    ''        Try
    ''            .OpenCurrentDatabase(targetMdb) ' 開啟 MDB 資料庫檔案
    ''            .Run(sFuncName, Para)
    ''        Catch ex As Exception
    ''        Finally
    ''            .CloseCurrentDatabase() ' 關閉資料庫
    ''            .Quit() ' 關閉 Access 執行個體
    ''            If sErr <> "" Then
    ''                Throw New Exception(sErr)
    ''            End If
    ''        End Try
    ''    End With
    ''End Sub

    'Public Sub AccessRunMacro(ByVal tarQry() As String, ByVal targetMdb As String)
    '    Dim oApp As New Access.Application
    '    Dim sErr As String = ""
    '    With oApp
    '        Try
    '            .OpenCurrentDatabase(targetMdb) ' 開啟 MDB 資料庫檔案
    '            For Each t As String In tarQry
    '                Try
    '                    .DoCmd.RunMacro(t)
    '                Catch ex As Exception
    '                    sErr += ex.Message & vbCr
    '                End Try
    '            Next
    '        Catch ex As Exception
    '        Finally
    '            .CloseCurrentDatabase() ' 關閉資料庫
    '            .Quit() ' 關閉 Access 執行個體
    '            If sErr <> "" Then
    '                Throw New Exception(sErr)
    '            End If
    '        End Try
    '    End With
    'End Sub

    'Public Sub AccessRunSQL(ByVal tarQry() As String, ByVal targetMdb As String)

    '    Dim oApp As New Access.Application
    '    Dim sErr As String = ""
    '    With oApp
    '        Try
    '            .OpenCurrentDatabase(targetMdb) ' 開啟 MDB 資料庫檔案
    '            For Each t As String In tarQry
    '                Try
    '                    .DoCmd.RunSQL(t)
    '                Catch ex As Exception
    '                    sErr += ex.Message & vbCr
    '                End Try
    '            Next
    '        Catch ex As Exception
    '        Finally
    '            .CloseCurrentDatabase() ' 關閉資料庫
    '            .Quit() ' 關閉 Access 執行個體
    '            If sErr <> "" Then
    '                Throw New Exception(sErr)
    '            End If
    '        End Try
    '    End With
    'End Sub

#End Region


End Class
