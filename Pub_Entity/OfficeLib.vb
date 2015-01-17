Imports Microsoft.Office.Interop
Public Class OfficeLib

#Region " Excel Excel.Application"
    Public Sub RangeImport(ByVal srcTbl As System.Data.DataTable, ByVal desRange As Excel.Range)
        Dim cls As New ClsArrayFunction
        Dim RowLength As Integer = CType(IIf(desRange.Rows.Count > srcTbl.Rows.Count, desRange.Rows.Count, srcTbl.Rows.Count), Integer)
        Dim ColumnLength As Integer = CType(IIf(desRange.Columns.Count > srcTbl.Columns.Count, desRange.Columns.Count, srcTbl.Columns.Count), Integer)
        'Dim cellArr As Object(,) = cls.ToArray(srcTbl, False, 0, desRange.Rows.Count, desRange.Columns.Count)
        Dim cellArr As Object(,) = cls.ToArray(srcTbl, False, 0, RowLength, ColumnLength)
        For i As Integer = srcTbl.Rows.Count + 1 To desRange.Rows.Count
            For j As Integer = srcTbl.Columns.Count + 1 To desRange.Columns.Count
                'CType(desRange.Cells(i, j), Excel.Range).Value = ""
                cellArr(i, j) = ""
            Next
        Next
        desRange.Value2 = cellArr

        'For i As Integer = 0 To srcTbl.Rows.Count - 1
        '    For j As Integer = 0 To srcTbl.Columns.Count - 1
        '        CType(desRange.Cells(i + 1, j + 1), Excel.Range).Value = srcTbl.Rows(i).Item(j)
        '    Next
        'Next
    End Sub

#Region "ToArray"
    Public Function ToArray(ByVal r As Excel.Range) As Object(,)
        Dim i As Integer, j As Integer
        Dim iRow As Integer = r.Rows.Count
        Dim iCol As Integer = r.Columns.Count

        Dim arr(iRow, iCol) As Object
        For i = 1 To iRow
            'ReDim arr(i)(iCol)
            For j = 1 To iCol
                arr(i, j) = CType(r.Cells(i, j), Excel.Range).Value
            Next
        Next
        Return arr
    End Function

    Public Function ToArray(ByVal w As Excel.Worksheet, ByVal iRow As Integer, ByVal iCol As Integer) As Object(,)
        Dim i As Integer, j As Integer
        Dim arr(iRow, iCol) As Object
        For i = 1 To iRow - 1
            'ReDim arr(i)(iCol)
            For j = 1 To iCol - 1
                arr(i, j) = CType(w.Cells(i, j), Excel.Range).Value
            Next
        Next
        Return arr
    End Function
#End Region


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


#Region "Excel Utility marked temperately"

    'Public Function getExcelApplication() As Excel.Application
    '    Try
    '        Dim xlApp As Excel.Application = New Excel.Application
    '        '停用警告訊息
    '        xlApp.DisplayAlerts = False
    '        '設置EXCEL對象可見
    '        xlApp.Visible = True

    '        Return xlApp
    '    Catch ex As Exception
    '        Throw New Exception(ex.ToString, ex.InnerException)
    '        Return Nothing
    '    End Try
    'End Function

    'Public Function getExcelWorkbook(ByVal sExcelPath As String) As Excel.Workbook
    '    Dim xlBook As Excel.Workbook = Nothing
    '    Try
    '        If Not My.Computer.FileSystem.FileExists(sExcelPath) Then
    '            Throw New Exception("Excel File path is nothing:" & sExcelPath)
    '        End If

    '        Dim xlApp As Excel.Application = New Excel.Application
    '        '打開已經存在的EXCEL工件簿文件
    '        xlBook = xlApp.Workbooks.Open(sExcelPath)
    '        '停用警告訊息
    '        xlApp.DisplayAlerts = False
    '        '設置EXCEL對象可見
    '        xlApp.Visible = True


    '        Return xlBook
    '    Catch ex As Exception
    '        Throw New Exception(ex.ToString, ex.InnerException)
    '        Return Nothing
    '    End Try
    'End Function

    ''Overloads Shared Function getExcelWorksheet(ByVal sExcelPath As String, ByVal strSheetName As String) As Excel.Worksheet
    ''    Dim xlBook As Excel.Workbook = getExcelWorkbook(sExcelPath)
    ''    Dim xlSheet As Excel.Worksheet = getExcelWorksheet(xlBook, strSheetName)
    ''    If IsNothing(xlSheet) Then
    ''        Debug.Print("WorkSheet is nothing")
    ''        Return Nothing
    ''    Else
    ''        Return xlSheet
    ''    End If
    ''End Function
    ''Overloads Shared Function getExcelWorksheet(ByVal xlBook As Excel.Workbook, ByVal strSheetName As String) As Excel.Worksheet
    ''    Dim xlSheet As Excel.Worksheet = Nothing
    ''    If IsNothing(xlBook) Then
    ''        Debug.Print("Excel.Workbook Is Nothing")
    ''        Return Nothing
    ''    End If
    ''    If strSheetName.Trim = "" Then
    ''        Debug.Print("Excel worksheet name is nothing")
    ''        Return Nothing
    ''    End If
    ''    For Each xx As Excel.Worksheet In xlBook.Worksheets
    ''        If xx.Name = strSheetName Then
    ''            xlSheet = xx
    ''            Exit For
    ''        End If
    ''    Next

    ''    Return xlSheet
    ''End Function

    ''Overloads Shared Function getExcelRange(ByVal sExcelPath As String, ByVal strRangeName As String) As Excel.Range
    ''    Dim xlBook As Excel.Workbook = getExcelWorkbook(sExcelPath)
    ''    If IsNothing(xlBook) Then
    ''        Debug.Print("WorkBook is nothing")
    ''        Return Nothing
    ''    End If

    ''    Dim xlRange As Excel.Range = getExcelRange(xlBook, strRangeName)
    ''    If IsNothing(xlRange) Then
    ''        Return Nothing
    ''    Else
    ''        Return xlRange
    ''    End If
    ''End Function
    ''Overloads Shared Function getExcelRange(ByVal xlBook As Excel.Workbook, ByVal strRangeName As String) As Excel.Range
    ''    If IsNothing(xlBook) Then
    ''        Debug.Print("WorkBook is nothing")
    ''        Return Nothing
    ''    End If
    ''    Dim xlApp As Excel.Application = xlBook.Application
    ''    If IsNothing(xlApp) Then
    ''        Debug.Print("Excel Excel.Application is nothing")
    ''        Return Nothing
    ''    End If
    ''    '' ''Dim xlSheet As Excel.Worksheet
    ''    '' ''For Each xlSheet In xlBook.Worksheets
    ''    '' ''    xlsheet.Cells.Find(xlapp.Range
    ''    '' ''Next

    ''    Dim xlRange As Excel.Range = xlApp.Range(strRangeName)
    ''    If IsNothing(xlRange) Then
    ''        Return Nothing
    ''    Else
    ''        Return xlRange
    ''    End If
    ''End Function

#End Region
End Class
