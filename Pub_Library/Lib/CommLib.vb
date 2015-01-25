Imports System.Configuration

Public Class CommLib

    Private ELib As New Pub_Entity.CommLib
    Private _cn As OleDb.OleDbConnection 'for SetPrimaryKey
    Private _trn As OleDb.OleDbTransaction 'for SetPrimaryKey
    Private Const DB_DateFormat As String = "YYYY-MM-DD"

#Region "資料庫函式"

#Region "CreateConnection"
    'Public Function CreateConnection(ByVal ExcelFileName As String, ByVal fHDR As Boolean, ByVal fIMEX As Boolean) As OleDb.OleDbConnection
    '    Return CreateConnection(ExcelFileName, fHDR, fIMEX, False)
    'End Function

    ''' <summary>
    ''' ExcelFile的存取連線
    ''' </summary>
    ''' <param name="ExcelFileName"></param>
    ''' <param name="fHDR">Tree:工作表中第一列為標題列,False:無標題</param>
    ''' <param name="fIMEX">表示是否強制轉換為文本 True:IMEX=1 False:IMEX=""
    ''' IMEX　表示是否強制轉換為文本
    '''Extended Properties='Excel 8.0;HDR=yes;IMEX=1'
    '''A： HDR ( HeaDer Row )設置
    '''若指定值為Yes，代表 Excel 檔中的工作表第一行是欄位名稱
    '''若指定值為 No，代表 Excel 檔中的工作表第一行就是資料了，沒有欄位名稱
    '''B：IMEX ( IMport EXport mode )設置
    '''IMEX 有三種模式，各自引起的讀寫行為也不同，容後再述：
    '''0 is Export mode
    '''1 is Import mode
    '''2 is Linked mode (full update capabilities)
    '''我這裏特別要說明的就是 IMEX 參數了，因為不同的模式代表著不同的讀寫行為：
    '''當 IMEX=0 時為“匯出模式”，這個模式開啟的 Excel 檔案只能用來做“寫入”用途。
    '''當 IMEX=1 時為“匯入模式”，這個模式開啟的 Excel 檔案只能用來做“讀取”用途。
    '''當 IMEX=2 時為“連結模式”，這個模式開啟的 Excel 檔案可同時支援“讀取”與“寫入”用途。
    ''' </param>
    ''' <param name="fReadOnly">目前不論其值為何,都只允許readonly</param>
    ''' <returns>OleDb.OleDbConnection</returns>
    ''' <remarks></remarks>
    Public Function CreateConnection(ByVal ExcelFileName As String, ByVal fHDR As Boolean, ByVal fIMEX As Boolean, ByVal fReadOnly As Boolean) As OleDb.OleDbConnection
        Try
            Dim oFile As New System.IO.FileInfo(ExcelFileName)
            If Not oFile.Exists Then
                Throw New Exception("Excel File " & ExcelFileName & " not exists!")
            End If

            Dim sHDR As String, sIMEX As String, sReadOnly As String
            If fHDR Then
                sHDR = "yes"
            Else
                sHDR = "no"
            End If
            If fIMEX Then
                sIMEX = "IMEX=1;"
            Else
                sIMEX = ""
            End If
            If fReadOnly Then
                sReadOnly = ""
            Else
                sReadOnly = ""
                'sReadOnly = "Mode=ReadWrite;ReadOnly=false;"
                'sReadOnly = "Mode=ReadWrite;"
                'sReadOnly = "ReadOnly=False;"
            End If
            Dim sConn As String
            If System.IO.File.Exists(ExcelFileName) Then
            Else
                ExcelFileName = GetCurrDir() & "\Excel\" & ExcelFileName
                If System.IO.File.Exists(ExcelFileName) Then
                Else
                    Throw New Exception("File " & ExcelFileName & " not exists!")
                End If
            End If
            sConn = "provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & ExcelFileName & "; " & sReadOnly & " Extended Properties=""Excel 8.0;HDR=" & sHDR & ";" & sIMEX & """"
            Return New System.Data.OleDb.OleDbConnection(sConn)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function CreateConnection(ByVal IServer As String, ByVal DBName As String, _
                                     ByVal UserID As String, ByVal Passwd As String) As OleDb.OleDbConnection
        '"Provider=MSDAORA; Data Source=ORACLE8i7;Persist Security Info=False;Integrated Security=Yes"
        '"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=c:\bin\LocalAccess的檔案名稱.mdb"
        '"Provider=SQLOLEDB;Data Source=(local);Integrated Security=SSPI"
        'Provider=any oledb provider's name;OledbKey1=someValue;OledbKey2=someValue;
        Dim sConn As String = "Provider=SQLOLEDB;Server=" & IServer & _
            ";Database=" & DBName & ";Uid=" & UserID & ";Pwd=" & Passwd & ";"
        Return CreateConnection(sConn)
    End Function

    Public Function CreateConnection(ByVal ConnString As String) As OleDb.OleDbConnection
        Try
            Return New OleDb.OleDbConnection(ConnString)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function CreateConnection(ByVal DbNAme As String, ByVal ConnString As String) As OleDb.OleDbConnection
        Try
            Dim cn As OleDb.OleDbConnection
            If ConnString = "" Then

                Dim conss As ConnectionStringSettings = ConfigurationManager.ConnectionStrings(DbNAme)
                ConnString = conss.ConnectionString
            End If
            cn = New OleDb.OleDbConnection(ConnString)
            Return cn
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

#End Region

#Region "DataSetLoad"

    Public Overloads Sub DataSetLoad(ByVal ds As DataSet, ByVal sql As String, ByVal TableName As String, ByVal cn As OleDb.OleDbConnection, ByVal trn As OleDb.OleDbTransaction, ByVal FlagFreeConnection As Boolean)
        Dim dt As New DataTable
        Dim dc As New OleDb.OleDbCommand
        Dim dr As OleDb.OleDbDataReader

        Try
            Try
                If cn.State <> ConnectionState.Open Then
                    cn.Open()
                End If
            Catch ex As Exception
                'MsgBox("[" & Err.Source & "][" & Err.Number & "][" & Err.Description)
                Throw New Exception("[Connection Open Error]" & ex.Message)
            End Try
            dc.Connection = cn
            If sql = "" Then
                sql = "SELECT * FROM " & TableName
            End If
            dc.CommandText = sql
            'KEVIN111122_1
            If Not IsNothing(trn) Then
                dc.Transaction = trn
            End If
            'Debug.Print(vbNewLine & sql)
            dr = dc.ExecuteReader
            If ds.Tables.Contains(TableName) Then
                ds.Tables(TableName).Rows.Clear()
            End If
            ds.Load(dr, LoadOption.PreserveChanges, New String() {TableName})
        Catch e As OleDb.OleDbException
            Dim errorMessages As String = ""
            ''For i As Integer = 0 To e.Errors.Count - 1
            ''    errorMessages += "Index #" & i.ToString() & ControlChars.Cr _
            ''                   & "Message: " & e.Errors(i).Message & ControlChars.Cr _
            ''                   & "NativeError: " & e.Errors(i).NativeError & ControlChars.Cr _
            ''                   & "Source: " & e.Errors(i).Source & ControlChars.Cr _
            ''                   & "SQLState: " & e.Errors(i).SQLState & ControlChars.Cr
            ''Next i
            Debug.Print(errorMessages)
            Throw New Exception(e.Message & "[" & sql & "]", e.InnerException)
        Catch ex As Exception
            Throw New Exception(ex.Message, ex.InnerException)
        Finally
            If FlagFreeConnection Then
                cn.Close()
                cn = Nothing
            End If
            'If Not IsNothing(cn) AndAlso cn.State <> ConnectionState.Closed Then
            '    cn.Close()
            'End If
            'cn = Nothing
        End Try
    End Sub
    Public Overloads Sub DataSetLoad(ByVal dt As DataTable, ByVal cn As OleDb.OleDbConnection)

        Dim dc As New OleDb.OleDbCommand
        Dim dr As OleDb.OleDbDataReader
        Try
            Try
                If cn.State <> ConnectionState.Open Then
                    cn.Open()
                End If
            Catch ex As Exception
                Debug.Print(ex.ToString)
            End Try

            dc.Connection = cn
            dc.CommandText = "SELECT * FROM " & dt.TableName

            dr = dc.ExecuteReader
            dt.Clear()
            dt.Load(dr, LoadOption.PreserveChanges)

        Catch e As OleDb.OleDbException
            Dim errorMessages As String = ""
            For i As Integer = 0 To e.Errors.Count - 1
                errorMessages += "Index #" & i.ToString() & ControlChars.Cr _
                               & "Message: " & e.Errors(i).Message & ControlChars.Cr _
                               & "NativeError: " & e.Errors(i).NativeError & ControlChars.Cr _
                               & "Source: " & e.Errors(i).Source & ControlChars.Cr _
                               & "SQLState: " & e.Errors(i).SQLState & ControlChars.Cr
            Next i
            Debug.Print(errorMessages)
            Throw New Exception(e.Message, e.InnerException)
        Catch ex As Exception
            Throw New Exception(ex.Message, ex.InnerException)
        End Try
    End Sub

    Public Overloads Sub DataSetLoad(ByVal ds As DataSet, ByVal sql As String, ByVal TableName As String, ByVal cn As OleDb.OleDbConnection, ByVal trn As OleDb.OleDbTransaction)
        DataSetLoad(ds, sql, TableName, cn, trn, False)
    End Sub

    Public Overloads Sub DataSetLoad(ByVal ds As DataSet, ByVal sql As String, ByVal TableName As String, ByVal sConnection As String)
        Dim conss As ConnectionStringSettings = ConfigurationManager.ConnectionStrings(sConnection)
        Dim ConnString As String = conss.ConnectionString
        Dim cn As OleDb.OleDbConnection
        cn = New OleDb.OleDbConnection(ConnString)
        If sql = "" Then
            sql = "SELECT * FROM " & TableName
        End If
        DataSetLoad(ds, sql, TableName, cn, Nothing)
    End Sub

    Public Overloads Sub DataSetLoad(ByVal ds As DataSet, ByVal TableName As String, ByVal cn As OleDb.OleDbConnection, ByVal trn As OleDb.OleDbTransaction)
        Dim sql As String
        sql = "SELECT * FROM " & TableName
        DataSetLoad(ds, sql, TableName, cn, trn)
    End Sub

#End Region

#Region "DataSetSave"
    Public Sub DataSetSave(ByVal ds As DataSet, ByVal sTableName As String, ByVal cn As OleDb.OleDbConnection, ByVal trn As OleDb.OleDbTransaction, ByVal FlagFreeConnection As Boolean, ByVal sTableMapping() As String)
        Dim da As OleDb.OleDbDataAdapter
        Try
            If cn.State <> ConnectionState.Open Then
                cn.Open()
            End If
            'da = New OleDb.OleDbDataAdapter()
            da = New OleDb.OleDbDataAdapter("SELECT * FROM " & sTableName, cn)

            If Not IsNothing(cn) Then
                _cn = cn
                _trn = trn
                AddHandler da.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf SetPrimaryKey)
            End If

            Dim cb As New OleDb.OleDbCommandBuilder(da)

            'If IsNothing(da.UpdateCommand) Then
            '    BuildDACommand(da, ds.Tables(sTableName), cn)
            'End If

            If Not IsNothing(sTableMapping) AndAlso sTableMapping.Length = 2 Then
                da.TableMappings.Add(sTableMapping(0), sTableMapping(1))
            End If

            'da.SelectCommand = New OleDb.OleDbCommand(GetSelectCommand(ds.Tables(sTableName), cn))
            'da.SelectCommand = New OleDb.OleDbCommand("SELECT * FROM " & sTableName, cn)

            da.SelectCommand = GetSelectCommand(ds.Tables(sTableName), cn)


            If Not IsNothing(trn) Then
                da.SelectCommand.Transaction = trn
            End If

            'da.Fill(ds.Tables(sTableName))
            'cb.RefreshSchema()

            If Not IsNothing(trn) Then
                'da.InsertCommand = cb.GetInsertCommand
                'da.InsertCommand.Transaction = trn
                'da.UpdateCommand = cb.GetUpdateCommand
                'da.UpdateCommand.Transaction = trn
                'da.DeleteCommand.Transaction = trn
                'da.DeleteCommand = cb.GetDeleteCommand

                da.InsertCommand = GetInsertCommand(ds.Tables(sTableName), cn)
                da.InsertCommand.Transaction = trn
                da.UpdateCommand = GetUpdateCommand(ds.Tables(sTableName), cn)
                da.UpdateCommand.Transaction = trn
                da.DeleteCommand = GetDeleteCommand(ds.Tables(sTableName), cn)
                da.DeleteCommand.Transaction = trn

                'cb.GetInsertCommand.Transaction = trn
                'cb.GetUpdateCommand.Transaction = trn
                'cb.GetDeleteCommand.Transaction = trn
            End If

            da.Update(ds, sTableName)

            If Not IsNothing(sTableMapping) AndAlso sTableMapping.Length = 2 Then
                ds.Tables(sTableMapping(1)).AcceptChanges()
            Else
                ds.Tables(sTableName).AcceptChanges()
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message & vbCr & "Table name: [" & sTableName & "]. Please inform IT staff for help")
        Finally
            If FlagFreeConnection Then
                If Not IsNothing(cn) Then
                    cn.Close()
                End If
                cn = Nothing
            End If
        End Try
    End Sub

    Public Sub DataSetSave(ByVal ds As DataSet, ByVal sTableName As String, ByVal cn As OleDb.OleDbConnection, ByVal trn As OleDb.OleDbTransaction, ByVal sTableMapping() As String)
        DataSetSave(ds, sTableName, cn, trn, False, sTableMapping)
    End Sub

    Public Sub DataSetSave(ByVal ds As DataSet, ByVal sTableName As String, ByVal cn As OleDb.OleDbConnection, ByVal trn As OleDb.OleDbTransaction)
        DataSetSave(ds, sTableName, cn, trn, False, Nothing)
    End Sub

    Public Sub DataSetSave(ByVal ds As DataSet, ByVal sTableName As String, ByVal sConnection As String, ByVal trn As OleDb.OleDbTransaction)
        Dim cn As OleDb.OleDbConnection
        'cn = CreateConnection(sConnection, "")
        Dim conss As ConnectionStringSettings = ConfigurationManager.ConnectionStrings(sConnection)
        Dim ConnString As String = conss.ConnectionString
        cn = New OleDb.OleDbConnection(ConnString)
        DataSetSave(ds, sTableName, cn, trn)
    End Sub

    Public Sub DataSetSave(ByVal ds As DataSet, ByVal cn As OleDb.OleDbConnection, ByVal trn As OleDb.OleDbTransaction)
        For Each dt As DataTable In ds.Tables
            DataSetSave(ds, dt.TableName, cn, trn)
        Next
    End Sub

    Public Sub DataSetSave(ByVal ds As DataSet, ByVal sTableName As String, ByVal cn As OleDb.OleDbConnection, ByVal trn As OleDb.OleDbTransaction, ByVal FlagFreeConnection As Boolean)
        DataSetSave(ds, sTableName, cn, trn, FlagFreeConnection, New String() {})
    End Sub

    Private Function GetSelectCommand(ByVal dt As DataTable, ByVal cn As OleDb.OleDbConnection) As OleDb.OleDbCommand
        Dim oCmd As OleDb.OleDbCommand
        Dim sSQL As String = ""
        Dim sWhere As String = "", sSelCol As String = ""
        Dim i As Integer
        For i = 0 To dt.Columns.Count - 1
            If i = dt.Columns.Count - 1 Then
                sSelCol += dt.Columns(i).ColumnName
            Else
                sSelCol += dt.Columns(i).ColumnName & ", "
            End If
        Next

        sSQL = "SELECT " & sSelCol & " FROM " & dt.TableName & " "
        If dt.PrimaryKey.Length > 0 Then
            For i = 0 To dt.PrimaryKey.Length - 1
                sWhere = AppendWhere(sWhere, dt.PrimaryKey(i).ColumnName & " = ? ")
            Next
        End If
        sSQL += sWhere

        oCmd = New OleDb.OleDbCommand(sSQL, cn)
        For i = 0 To dt.PrimaryKey.Length - 1
            AddCmdParameter(oCmd, dt.PrimaryKey(i))
            'Dim sCol As String = dt.PrimaryKey(i).ColumnName
            'Dim oType As OleDb.OleDbType
            'Dim iSize As Integer
            'Select Case dt.PrimaryKey(i).DataType.ToString
            '    Case "System.String"
            '        oType = OleDb.OleDbType.VarChar
            '        iSize = dt.PrimaryKey(i).MaxLength
            '    Case "System.DateTime"
            '        oType = OleDb.OleDbType.Date
            '    Case "System.Decimal"
            '        oType = OleDb.OleDbType.Decimal
            '    Case Else
            'End Select
            'oCmd.Parameters.Add(sCol, oType, iSize, sCol)
        Next

        Return oCmd
    End Function

    Private Function GetInsertCommand(ByVal dt As DataTable, ByVal cn As OleDb.OleDbConnection) As OleDb.OleDbCommand
        Dim oCmd As OleDb.OleDbCommand
        Dim sCol As String = "", sPara As String = ""
        Dim sWhere As String = "", sSelCol As String = ""
        Dim i As Integer
        For i = 0 To dt.Columns.Count - 1
            'If i = dt.Columns.Count - 1 Then
            '    sCol += dt.Columns(i).ColumnName
            '    sPara += "?"
            'Else
            '    sCol += dt.Columns(i).ColumnName & ", "
            '    sPara += "?, "
            'End If
            If dt.Columns(i).AutoIncrement = False Then
                sCol += dt.Columns(i).ColumnName & ", "
                sPara += "?, "
            End If
        Next
        If Right(sCol, 2) = ", " Then sCol = Left(sCol, Len(sCol) - 2)
        If Right(sPara, 2) = ", " Then sPara = Left(sPara, Len(sPara) - 2)

        Dim sSQL As String = "INSERT INTO " & dt.TableName & "(" & sCol & ") VALUES (" & sPara & ")"
        'sSQL = "SELECT " & sSelCol & " FROM " & dt.TableName & " "

        oCmd = New OleDb.OleDbCommand(sSQL, cn)
        For i = 0 To dt.Columns.Count - 1
            If dt.Columns(i).AutoIncrement = False Then
                AddCmdParameter(oCmd, dt.Columns(i))
            End If
        Next

        Return oCmd
    End Function

    Private Function GetUpdateCommand(ByVal dt As DataTable, ByVal cn As OleDb.OleDbConnection) As OleDb.OleDbCommand
        Dim oCmd As OleDb.OleDbCommand
        Dim sSQL As String = ""
        Dim sWhere As String = "", sSelCol As String = ""
        Dim i As Integer
        For i = 0 To dt.Columns.Count - 1
            'If i = dt.Columns.Count - 1 Then
            '    sSelCol += dt.Columns(i).ColumnName & " = ? "
            'Else
            '    sSelCol += dt.Columns(i).ColumnName & " = ?, "
            'End If
            If dt.Columns(i).AutoIncrement = False Then
                sSelCol += dt.Columns(i).ColumnName & " = ?, "
            End If
        Next
        If Right(sSelCol, 2) = ", " Then sSelCol = Left(sSelCol, Len(sSelCol) - 2)

        If dt.PrimaryKey.Length > 0 Then
            For i = 0 To dt.PrimaryKey.Length - 1
                sWhere = AppendWhere(sWhere, dt.PrimaryKey(i).ColumnName & " = ? ")
            Next
        Else
            For i = 0 To dt.Columns.Count - 1
                sWhere = AppendWhere(sWhere, dt.Columns(i).ColumnName & " = ? ")
            Next
        End If

        sSQL = "UPDATE " & dt.TableName & " SET " & sSelCol & " " & sWhere

        oCmd = New OleDb.OleDbCommand(sSQL, cn)
        For i = 0 To dt.Columns.Count - 1
            If dt.Columns(i).AutoIncrement = False Then
                AddCmdParameter(oCmd, dt.Columns(i))
            End If
        Next

        If dt.PrimaryKey.Length > 0 Then
            For i = 0 To dt.PrimaryKey.Length - 1
                AddCmdParameter(oCmd, dt.PrimaryKey(i), DataRowVersion.Original)
            Next
        Else
            For i = 0 To dt.Columns.Count - 1
                AddCmdParameter(oCmd, dt.Columns(i), DataRowVersion.Original)
            Next
        End If

        Return oCmd
    End Function

    Private Function GetDeleteCommand(ByVal dt As DataTable, ByVal cn As OleDb.OleDbConnection) As OleDb.OleDbCommand
        Dim oCmd As OleDb.OleDbCommand
        Dim sSQL As String = ""
        Dim sWhere As String = "", sSelCol As String = ""
        Dim i As Integer

        If dt.PrimaryKey.Length > 0 Then
            For i = 0 To dt.PrimaryKey.Length - 1
                sWhere = AppendWhere(sWhere, dt.PrimaryKey(i).ColumnName & " = ? ")
            Next
        Else
            For i = 0 To dt.Columns.Count - 1
                sWhere = AppendWhere(sWhere, dt.Columns(i).ColumnName & " = ? ")
            Next
        End If

        sSQL = "DELETE FROM " & dt.TableName & " " & sWhere
        oCmd = New OleDb.OleDbCommand(sSQL, cn)

        If dt.PrimaryKey.Length > 0 Then
            For i = 0 To dt.PrimaryKey.Length - 1
                AddCmdParameter(oCmd, dt.PrimaryKey(i), DataRowVersion.Original)
            Next
        Else
            For i = 0 To dt.Columns.Count - 1
                AddCmdParameter(oCmd, dt.Columns(i), DataRowVersion.Original)
            Next
        End If

        Return oCmd
    End Function


    Private Sub AddCmdParameter(ByVal oCmd As OleDb.OleDbCommand, ByVal oCol As DataColumn)
        AddCmdParameter(oCmd, oCol, DataRowVersion.Current)
    End Sub

    Private Sub AddCmdParameter(ByVal oCmd As OleDb.OleDbCommand, ByVal oCol As DataColumn, ByVal oRowVer As DataRowVersion)
        Dim sCol As String = oCol.ColumnName
        Dim oType As OleDb.OleDbType
        Dim iSize As Integer
        Select Case oCol.DataType.ToString
            Case "System.String"
                oType = OleDb.OleDbType.VarChar
                iSize = oCol.MaxLength
            Case "System.DateTime"
                oType = OleDb.OleDbType.Date
            Case "System.Decimal"
                oType = OleDb.OleDbType.Decimal
            Case "System.Int16"
                oType = OleDb.OleDbType.SmallInt
            Case "System.Int32"
                oType = OleDb.OleDbType.Integer
            Case "System.Int64"
                oType = OleDb.OleDbType.BigInt
            Case "System.Byte[]"
                oType = OleDb.OleDbType.Binary
            Case Else
                Throw New Exception("Parameter datatype not correct " & oCol.DataType.ToString)

        End Select
        Dim para As OleDb.OleDbParameter = oCmd.Parameters.Add(sCol, oType, iSize, sCol)
        para.SourceVersion = oRowVer
    End Sub
#End Region


#Region "ExecSQLCmd"

    Public Sub ExecSQLCmd(ByVal sSQL As String, ByVal trn As OleDb.OleDbTransaction)
        ' Dim cmd As New OleDb.OleDbCommand
        'cmd.Connection = cn
        'cmd.CommandType = Type
        Using cmd As New OleDb.OleDbCommand
            cmd.Transaction = trn
            cmd.Connection = trn.Connection
            cmd.CommandText = sSQL
            'Debug.Print(sSQL + vbNewLine)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Sub ExecSQLCmd(ByVal sSQL As String, ByVal cn As OleDb.OleDbConnection, ByVal FlagFreeConnection As Boolean)
        ExecSQLCmd(CommandType.Text, sSQL, cn, FlagFreeConnection)
    End Sub

    Public Sub ExecSQLCmd(ByVal type As CommandType, ByVal sSQL As String, ByVal cn As OleDb.OleDbConnection, ByVal FlagFreeConnection As Boolean)
        ExecSQLCmd(type, sSQL, cn, FlagFreeConnection, Nothing)
    End Sub

    Public Sub ExecSQLCmd(ByVal type As CommandType, ByVal sSQL As String, ByVal cn As OleDb.OleDbConnection, ByVal FlagFreeConnection As Boolean, ByVal iTimeout As Integer)
        Try
            If cn.State <> ConnectionState.Open Then
                cn.Open()
            End If

            Using cmd As New OleDb.OleDbCommand
                cmd.Connection = cn
                cmd.CommandType = type
                cmd.CommandText = sSQL
                If Not IsNothing(iTimeout) Then
                    cmd.CommandTimeout = iTimeout
                End If
                'Debug.Print(sSQL.ToString + vbNewLine)
                cmd.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            'Dim sMsg As String = ex.Message
            If FlagFreeConnection AndAlso Not IsNothing(cn) Then
                cn.Close()
                cn = Nothing
            End If
            Throw New Exception(ex.Message & vbCr & "SQL String: [" & sSQL & "]. Please inform IT staff for help")
            'Throw New Exception(ex.Message)
        Finally
            If FlagFreeConnection Then
                cn.Close()
                cn = Nothing
            End If
        End Try
    End Sub
#End Region

#Region "GetValue"
    'KEVIN121105_1
    Public Function GetDataRow(ByVal sSQL As String, ByVal oConn As OleDb.OleDbConnection) As DataRow
        Return GetDataRow(sSQL, oConn, Nothing)
    End Function
    Public Function GetDataRow(ByVal sSQL As String, ByVal oConn As OleDb.OleDbConnection, ByVal oTrn As OleDb.OleDbTransaction) As DataRow
        Dim dr As DataRow = Nothing
        Dim dsTmp As New DataSet
        DataSetLoad(dsTmp, sSQL, "TABLE1", oConn, oTrn)
        If dsTmp.Tables.Contains("TABLE1") AndAlso dsTmp.Tables("TABLE1").Rows.Count > 0 Then
            dr = dsTmp.Tables("TABLE1").Rows(0)
        End If
        Return dr
    End Function


    'KEVIN120430_5
    Public Function GetStrValue(ByVal sSQL As String, ByVal oConn As OleDb.OleDbConnection) As String
        Return GetStrValue(sSQL, oConn, Nothing)
    End Function

    'KEVIN120430_5
    'Public Function GetStrValue(ByVal sSQL As String, ByVal oConn As OleDb.OleDbConnection) As String
    Public Function GetStrValue(ByVal sSQL As String, ByVal oConn As OleDb.OleDbConnection, ByVal oTrn As OleDb.OleDbTransaction) As String
        Dim sVal As String = ""
        Dim dsTmp As New DataSet
        'KEVIN120430_5
        'DataSetLoad(dsTmp, sSQL, "TABLE1", oConn, Nothing, False)
        DataSetLoad(dsTmp, sSQL, "TABLE1", oConn, oTrn)

        If dsTmp.Tables.Contains("TABLE1") AndAlso dsTmp.Tables("TABLE1").Rows.Count > 0 AndAlso dsTmp.Tables("TABLE1").Columns.Count > 0 Then
            sVal = dsTmp.Tables("TABLE1").Rows(0).Item(0).ToString
        Else
            sVal = ""
        End If
        Return sVal
    End Function

    'KEVIN120430_5
    Public Function GetDecValue(ByVal sSQL As String, ByVal oConn As OleDb.OleDbConnection) As Decimal
        Return GetDecValue(sSQL, oConn, Nothing)
    End Function

    'KEVIN120430_5
    'Public Function GetDecValue(ByVal sSQL As String, ByVal oConn As OleDb.OleDbConnection) As Decimal
    Public Function GetDecValue(ByVal sSQL As String, ByVal oConn As OleDb.OleDbConnection, ByVal oTrn As OleDb.OleDbTransaction) As Decimal
        Dim iVal As Decimal = 0
        Dim dsTmp As New DataSet
        'KEVIN120430_5
        'DataSetLoad(dsTmp, sSQL, "TABLE1", oConn, Nothing, False)
        DataSetLoad(dsTmp, sSQL, "TABLE1", oConn, oTrn)
        If dsTmp.Tables.Contains("TABLE1") AndAlso dsTmp.Tables("TABLE1").Rows.Count > 0 AndAlso dsTmp.Tables("TABLE1").Columns.Count > 0 Then
            If IsDBNull(dsTmp.Tables("TABLE1").Rows(0).Item(0)) Then
                iVal = 0
            Else
                iVal = CType(dsTmp.Tables("TABLE1").Rows(0).Item(0), Decimal)
            End If
        Else
            iVal = 0
        End If
        Return iVal
    End Function

    'KEVIN120509_4
    Public Function GetDateValue(ByVal sSQL As String, ByVal oConn As OleDb.OleDbConnection) As Date
        Return GetDateValue(sSQL, oConn, Nothing)
    End Function

    Public Function GetDateValue(ByVal sSQL As String, ByVal oConn As OleDb.OleDbConnection, ByVal oTrn As OleDb.OleDbTransaction) As Date
        Dim dDate As Date
        Dim dsTmp As New DataSet
        DataSetLoad(dsTmp, sSQL, "TABLE1", oConn, oTrn)
        If dsTmp.Tables.Contains("TABLE1") AndAlso dsTmp.Tables("TABLE1").Rows.Count > 0 AndAlso dsTmp.Tables("TABLE1").Columns.Count > 0 Then
            If IsDBNull(dsTmp.Tables("TABLE1").Rows(0).Item(0)) Then
                dDate = Nothing
            Else
                dDate = CType(dsTmp.Tables("TABLE1").Rows(0).Item(0), Date)
            End If
        Else
            dDate = Nothing
        End If
        Return dDate
    End Function

#End Region

#Region "Command Builder"
    Private Sub BuildDACommand(ByVal da As OleDb.OleDbDataAdapter, ByVal dt As DataTable, ByVal cn As OleDb.OleDbConnection)
        Dim sCmdPre As String = ""
        Dim sSQL As String
        Dim sTmp1 As String, sTmp2 As String
        Dim cmd As OleDb.OleDbCommand

        ' InsertCommand
        sTmp1 = ""
        sTmp2 = ""
        For Each dc As DataColumn In dt.Columns
            If sTmp1 <> "" Then sTmp1 += ","
            If sTmp2 <> "" Then sTmp2 += ","
            sTmp1 += dc.ColumnName
            sTmp2 += "?"
            'sTmp2 += "@" & dc.ColumnName
        Next
        sSQL = "INSERT INTO " & dt.TableName & "(" & sTmp1 & ") VALUES (" & sTmp2 & ")"
        cmd = New OleDb.OleDbCommand(sSQL, cn)
        For Each dc As DataColumn In dt.Columns
            sTmp2 += sCmdPre & dc.ColumnName
            cmd.Parameters.Add(sTmp2, OleDb.OleDbType.Double)
            cmd.Parameters(sTmp2).SourceColumn = dc.ColumnName
        Next
        da.InsertCommand = cmd

        '' Create the UpdateCommand.
        sTmp1 = ""
        sTmp2 = ""
        For Each dc As DataColumn In dt.Columns
            If sTmp1 <> "" Then sTmp1 += ","
            If sTmp2 <> "" Then sTmp2 += " AND "
            sTmp1 += dc.ColumnName & " = ?"
            sTmp2 += dc.ColumnName & " = ?"
        Next
        sSQL = "UPDATE " & dt.TableName & " SET " & sTmp1 & " WHERE " & sTmp2
        cmd = New OleDb.OleDbCommand(sSQL, cn)
        For Each dc As DataColumn In dt.Columns
            sTmp2 += sCmdPre & dc.ColumnName
            cmd.Parameters.Add(sTmp2, OleDb.OleDbType.Double)
            cmd.Parameters(sTmp2).SourceColumn = dc.ColumnName
            cmd.Parameters(sTmp2).SourceVersion = DataRowVersion.Current
        Next
        For Each dc As DataColumn In dt.Columns
            sTmp2 += sCmdPre & dc.ColumnName & "_Ori"
            cmd.Parameters.Add(sTmp2, OleDb.OleDbType.Double)
            cmd.Parameters(sTmp2).SourceColumn = dc.ColumnName
            cmd.Parameters(sTmp2).SourceVersion = DataRowVersion.Original
        Next
        da.UpdateCommand = cmd
    End Sub

    'Private Function GetOleDBType(ByVal t As System.Type) As OleDb.OleDbType
    '    Select Case LCase(t.ToString)
    '        Case "system.string"
    '            Return OleDb.OleDbType.VarChar
    '        Case Else
    '            Return OleDb.OleDbType.Double
    '    End Select
    'End Function
#End Region

#Region "SQL Function"

    Public Function AppendWhere(ByVal ori_s As String, ByVal s As String) As String
        If ori_s = "" Then
            s = " WHERE " & s
        Else
            s = ori_s & " AND " & s
        End If
        Return s
    End Function

    Public Function AddQuote(ByVal s As String) As String
        Return ELib.AddQuote(s)
        'Dim sTmp As String = ""
        's = s.Replace("'", "''")
        'sTmp = "'" & s & "'"

        'Return sTmp
    End Function

    'todo: [Common]for SQL Server only currently
    Public Sub SetPrimaryKey(ByVal sender As Object, _
                                    ByVal e As OleDb.OleDbRowUpdatedEventArgs)

        ' If this is an INSERT operation...
        If e.Status = UpdateStatus.Continue AndAlso _
           e.StatementType = StatementType.Insert Then

            Dim pk = e.Row.Table.PrimaryKey
            ' and a primary key PK column exists...
            If pk IsNot Nothing AndAlso pk.Count = 1 Then
                'Set up the post-update query to fetch new @@Identity
                Dim cmdGetIdentity As New OleDb.OleDbCommand("SELECT @@IDENTITY", _cn, _trn)
                'Execute the command and set the result identity value to the PK
                e.Row(pk(0)) = CInt(cmdGetIdentity.ExecuteScalar)
                e.Row.AcceptChanges()
            End If
        End If
    End Sub

#End Region


#End Region

#Region "DataSet Function"

    Public Sub CopyTableData(ByVal srcTbl As DataTable, ByVal desTbl As DataTable)
        For Each dr As DataRow In srcTbl.Rows
            desTbl.Rows.Add(dr.ItemArray)
        Next
    End Sub

    Public Sub DeleteTableData(ByVal tbl As DataTable)
        Dim i As Integer
        For i = 0 To tbl.Rows.Count - 1
            If tbl.Rows(i).RowState <> DataRowState.Deleted Then
                tbl.Rows(i).Delete()
            End If
        Next
    End Sub
#End Region

    Public Function GetCurrDir() As String
        Return ELib.GetCurrDir()
        'Return My.Application.Info.DirectoryPath
    End Function

#Region "Recycle Can"
    ''No Use
    'Public Function GetRangeDataDS(ByVal FuncName As String, ByVal Fioc As String) As ExcelFnData
    '    If FuncName <> "" And Fioc <> "" Then
    '        Dim sWhere As String = ""
    '        sWhere = sWhere & "WHERE FunctionName ='" & Trim(FuncName) & "'"
    '        sWhere = sWhere & "AND Fioc ='" & Trim(Fioc) & "'"


    '        Dim cn As OleDb.OleDbConnection = Nothing
    '        Dim ds As New ExcelFnData
    '        Try
    '            'cn = CreateConnection("DFA2_DLL_SQL", "")
    '            cn = GetRangeConn()
    '            cn.Open()
    '            DataSetLoad(ds, "SELECT * FROM RangeData " & sWhere & "ORDER BY FunctionName, FIOC, SeqNo", "RangeData", cn)
    '        Catch ex As Exception
    '        Finally
    '            If Not IsNothing(cn) Then
    '                cn.Close()
    '            End If
    '            cn = Nothing
    '        End Try
    '        Return ds
    '    End If
    '    Return Nothing
    'End Function

    ''NoUse
    'Public Function GetRangeDataDS() As ExcelFnData
    '    Return GetRangeDataDS(Nothing)
    'End Function

    '    'No Use Now
    '#Region "Main"
    '    '    Private Function GetMainConn() As OleDb.OleDbConnection
    '    '        Dim cn As OleDb.OleDbConnection
    '    '        cn = CreateConnection("DFA2", "")
    '    '        Return cn
    '    '    End Function

    '    '    Public Function GetMainDS() As dsMainPara
    '    '        Return GetMainDS("")
    '    '    End Function

    '    '    Public Function GetMainDS(ByVal ProjName As String) As dsMainPara
    '    '        Dim sWhere As String = ""
    '    '        If ProjName <> "" Then
    '    '            sWhere = sWhere & "WHERE ProjName ='" & Trim(ProjName) & "'"
    '    '        End If

    '    '        Dim cn As OleDb.OleDbConnection = Nothing
    '    '        Dim ds As New dsMainPara
    '    '        Try
    '    '            cn = GetMainConn()
    '    '            cn.Open()
    '    '            DataSetLoad(ds, "SELECT * FROM MainProjData " & sWhere & "ORDER BY ProjName", "MainProjData", cn)
    '    '            DataSetLoad(ds, "SELECT * FROM MainProdData " & sWhere & "ORDER BY ProdName", "MainProdData", cn)
    '    '            DataSetLoad(ds, "SELECT * FROM MainRIProgram " & sWhere & "ORDER BY ProgName", "MainRIProgram", cn)
    '    '            DataSetLoad(ds, "SELECT * FROM MainPeriodData " & sWhere & "ORDER BY CalcType, VersionNo", "MainPeriodData", cn)
    '    '        Catch ex As Exception
    '    '        Finally
    '    '            If Not IsNothing(cn) Then
    '    '                cn.Close()
    '    '            End If
    '    '            cn = Nothing
    '    '        End Try
    '    '        Return ds
    '    '        Return Nothing
    '    '    End Function

    '    '    Public Function SaveMainDS(ByVal dsPara As dsMainPara) As dsMainPara
    '    '        Dim da As OleDb.OleDbDataAdapter = Nothing
    '    '        Dim ssql As String
    '    '        Dim cn As New OleDb.OleDbConnection
    '    '        Try
    '    '            cn = GetMainConn()
    '    '            For Each dt As DataTable In dsPara.Tables
    '    '                ssql = "SELECT * FROM " & dt.TableName
    '    '                da = New OleDb.OleDbDataAdapter(ssql, cn)
    '    '                Dim cb As New OleDb.OleDbCommandBuilder(da)
    '    '                da.Update(dsPara, dt.TableName)
    '    '                dt.AcceptChanges()
    '    '            Next
    '    '        Catch ex As Exception

    '    '        Finally
    '    '            If Not IsNothing(cn) Then
    '    '                cn.Close()
    '    '            End If
    '    '            cn = Nothing
    '    '        End Try
    '    '        Return dsPara
    '    '    End Function

    '#End Region

#End Region

#Region "Under Experiment"

    Public Function FindProcess(ByVal procname As String) As System.Diagnostics.Process()
        Return System.Diagnostics.Process.GetProcessesByName(procname)
    End Function

    Public Sub KillProcess(ByVal ProcName As String)
        Dim thisProc As System.Diagnostics.Process
        Dim allRelationalProcs() As Process = FindProcess(ProcName)

        For Each thisProc In allRelationalProcs
            Try
                'Dim thds As ProcessThreadCollection
                'thds = thisProc.Threads
                'For Each thd As ProcessThread In thds
                '    Dim s As String = thd.ToString
                'Next
                If Not thisProc.CloseMainWindow() Then
                    thisProc.Kill()
                End If
            Catch ex As Exception
                'MsgBox(ex.GetBaseException.ToString)
                Throw New Exception("[KillProcessS error] " & ex.Message)
            End Try
        Next
    End Sub

    Public Sub KillProcess(ByVal ProcID As Integer)
        Dim thisProc As System.Diagnostics.Process = Nothing
        Try
            thisProc = System.Diagnostics.Process.GetProcessById(ProcID)
        Catch ex As Exception
        End Try

        'For Each thisProc In allRelationalProcs
        Try
            'Dim thds As ProcessThreadCollection
            'thds = thisProc.Threads
            'For Each thd As ProcessThread In thds
            '    Dim s As String = thd.ToString
            'Next
            If Not IsNothing(thisProc) AndAlso Not thisProc.CloseMainWindow() Then
                thisProc.Kill()
            End If
        Catch ex As Exception
            'MsgBox(ex.GetBaseException.ToString)
            Throw New Exception("[KillProcessI error] " & ex.Message)
        End Try
        'Next
    End Sub
#End Region

    'Todo: [Common]For Oracle Only now
#Region "SQL String Function"
    Public Function GetDateSQLStringBetween(ByVal dDate1 As Date, ByVal dDate2 As Date) As String
        Dim s As String = ""
        s = " BETWEEN TO_DATE(" & AddQuote(dDate1.ToString(DB_DateFormat) & " 00:00:00") & ", " & AddQuote(DB_DateFormat & " :HH24:MI:SS") & ")"
        s += " AND TO_DATE(" & AddQuote(dDate2.ToString(DB_DateFormat) & " 23:59:59") & ", " & AddQuote(DB_DateFormat & " :HH24:MI:SS") & ")"
        Return s
    End Function

    Public Function GetDateSQLString(ByVal dDate As Date) As String
        Return GetDateSQLString(dDate.ToString(DB_DateFormat))
    End Function

    Public Function GetDateSQLString(ByVal sDate As String) As String
        Dim s As String = ""

        s = "TO_DATE(" & AddQuote(sDate) & "," & AddQuote(DB_DateFormat) & ") "
        'SqlStr7 = " AND TRANS_DATE = TO_DATE('" & Trim(sTransFromDate) & "','" & DB_DateFormat & "') "
        Return s
    End Function


    'CYGOO130711_1
    Public Function GetDateTimeSQLString(ByVal sDate As Date) As String
        Dim s As String = ""
        s = "TO_DATE(" & AddQuote(sDate.ToString(DB_DateFormat) & " 23:59:59") & ", " & AddQuote(DB_DateFormat & " :HH24:MI:SS") & ")"
        Return s
    End Function

#End Region

    Public Sub WriteToWindowsEvent(ByVal aSource As String, ByVal aMessage As String)

        'If (Not EventLog.SourceExists(aSource)) Then
        '    EventLog.CreateEventSource(aSource, "Application")
        'End If
        ''Dim oLog As EventLog = New EventLog("Application")
        ''oLog.Source = aSource
        ''oLog.WriteEntry(aMessage, EventLogEntryType.Error)


        'Dim oLog As New EventLog
        'oLog.Log = "Application"
        'oLog.WriteEntry(aSource, aMessage, EventLogEntryType.Error)

    End Sub

    '#Region "DataTableSave"
    '    ''' <summary>
    '    ''' To Sync datatable with Database 
    '    ''' </summary>
    '    ''' <param name="dt">DataTable </param>
    '    ''' <param name="sConnection">connection string</param>
    '    ''' <remarks>create by scott 20100508</remarks>
    '    Public Overloads Sub DataTableSave(ByVal dt As DataTable, ByVal sConnection As String)
    '        Dim cn As OleDb.OleDbConnection

    '        Dim conss As ConnectionStringSettings = ConfigurationManager.ConnectionStrings(sConnection)
    '        Dim ConnString As String = conss.ConnectionString
    '        cn = New OleDb.OleDbConnection(ConnString)
    '        DataTableSave(dt, cn)
    '    End Sub
    '    ''' <summary>
    '    ''' To Sync datatable with Database
    '    ''' </summary>
    '    ''' <param name="dt">DataTable</param>
    '    ''' <param name="cn">OLE DB Connection</param>
    '    ''' <remarks>create by scott 20100508</remarks>
    '    Public Overloads Sub DataTableSave(ByVal dt As DataTable, ByVal cn As OleDb.OleDbConnection)
    '        DataTableSave(dt, cn, False, Nothing)
    '    End Sub
    '    ''' <summary>
    '    ''' To Sync datatable with Database
    '    ''' </summary>
    '    ''' <param name="dt">DataTable</param>
    '    ''' <param name="cn">OLE DB Connection</param>
    '    ''' <param name="sDbTableName">要對應的資料庫內來源資料行之區分大小寫名稱</param>
    '    ''' <remarks>create by scott 20100508</remarks>
    '    Public Overloads Sub DataTableSave(ByVal dt As DataTable, ByVal cn As OleDb.OleDbConnection, ByVal sDbTableName As String)
    '        DataTableSave(dt, cn, False, sDbTableName)
    '    End Sub
    '    Public Overloads Sub DataTableSave(ByVal dt As DataTable, ByVal cn As OleDb.OleDbConnection, ByVal FlagFreeConnection As Boolean)
    '        DataTableSave(dt, cn, FlagFreeConnection, "")
    '    End Sub
    '    Public Overloads Sub DataTableSave(ByVal dt As DataTable)
    '        DataTableSave(dt, CreateConnection, True, Nothing)
    '    End Sub
    '    Public Overloads Sub DataTableSave(ByVal dt As DataTable, ByVal cn As OleDb.OleDbConnection, ByVal FlagFreeConnection As Boolean, ByVal sDbTableName As String)
    '        Dim da As OleDb.OleDbDataAdapter
    '        Try
    '            If cn.State <> ConnectionState.Open Then
    '                cn.Open()
    '            End If
    '            If sDbTableName.Trim <> "" Then
    '                da = New OleDb.OleDbDataAdapter("SELECT * FROM " & sDbTableName, cn)
    '                da.TableMappings.Add(sDbTableName, dt.TableName)
    '            Else
    '                da = New OleDb.OleDbDataAdapter("SELECT * FROM " & dt.TableName, cn)
    '            End If

    '            Try
    '                Debug.Print("[A]:" & da.SelectCommand.CommandText.ToString)
    '                Debug.Print("[U]:" & da.UpdateCommand.CommandText.ToString)
    '                Debug.Print("[D]:" & da.DeleteCommand.CommandText.ToString)
    '            Catch ex As Exception
    '            End Try
    '            da.Update(dt)
    '            dt.AcceptChanges()

    '        Catch ex As Exception
    '            Throw New Exception(ex.Message)
    '        Finally
    '            If FlagFreeConnection Then
    '                If Not IsNothing(cn) Then
    '                    cn.Close()
    '                End If
    '                cn = Nothing
    '            End If
    '        End Try
    '    End Sub
    '#End Region

    '#Region "DataSetLoad"

    '    Public Overloads Sub DataSetLoad(ByVal ds As DataSet, ByVal sql As String, ByVal TableName As String, ByVal cn As OleDb.OleDbConnection, ByVal FlagFreeConnection As Boolean)
    '        Dim dt As New DataTable
    '        Dim dc As New OleDb.OleDbCommand
    '        Dim dr As OleDb.OleDbDataReader

    '        Try
    '            Try
    '                If cn.State <> ConnectionState.Open Then
    '                    cn.Open()
    '                End If
    '            Catch ex As Exception
    '                'MsgBox("[" & Err.Source & "][" & Err.Number & "][" & Err.Description)
    '                Throw New Exception("[Connection Open Error]" & ex.Message)
    '            End Try
    '            dc.Connection = cn
    '            If sql = "" Then
    '                sql = "SELECT * FROM " & TableName
    '            End If
    '            dc.CommandText = sql
    '            'Debug.Print(vbNewLine & sql)
    '            dr = dc.ExecuteReader
    '            If ds.Tables.Contains(TableName) Then
    '                ds.Tables(TableName).Rows.Clear()
    '            End If
    '            ds.Load(dr, LoadOption.PreserveChanges, New String() {TableName})
    '        Catch e As OleDb.OleDbException
    '            Dim errorMessages As String = ""
    '            ''For i As Integer = 0 To e.Errors.Count - 1
    '            ''    errorMessages += "Index #" & i.ToString() & ControlChars.Cr _
    '            ''                   & "Message: " & e.Errors(i).Message & ControlChars.Cr _
    '            ''                   & "NativeError: " & e.Errors(i).NativeError & ControlChars.Cr _
    '            ''                   & "Source: " & e.Errors(i).Source & ControlChars.Cr _
    '            ''                   & "SQLState: " & e.Errors(i).SQLState & ControlChars.Cr
    '            ''Next i
    '            Debug.Print(errorMessages)
    '            Throw New Exception(e.Message, e.InnerException)
    '        Catch ex As Exception
    '            Throw New Exception(ex.Message, ex.InnerException)
    '        Finally
    '            If FlagFreeConnection Then
    '                cn.Close()
    '                cn = Nothing
    '            End If
    '            'If Not IsNothing(cn) AndAlso cn.State <> ConnectionState.Closed Then
    '            '    cn.Close()
    '            'End If
    '            'cn = Nothing
    '        End Try
    '    End Sub
    '    Public Overloads Sub DataSetLoad(ByVal dt As DataTable, ByVal cn As OleDb.OleDbConnection)

    '        Dim dc As New OleDb.OleDbCommand
    '        Dim dr As OleDb.OleDbDataReader
    '        Try
    '            Try
    '                If cn.State <> ConnectionState.Open Then
    '                    cn.Open()
    '                End If
    '            Catch ex As Exception
    '                Debug.Print(ex.ToString)
    '            End Try

    '            dc.Connection = cn
    '            dc.CommandText = "SELECT * FROM " & dt.TableName

    '            dr = dc.ExecuteReader
    '            dt.Clear()
    '            dt.Load(dr, LoadOption.PreserveChanges)

    '        Catch e As OleDb.OleDbException
    '            Dim errorMessages As String = ""
    '            For i As Integer = 0 To e.Errors.Count - 1
    '                errorMessages += "Index #" & i.ToString() & ControlChars.Cr _
    '                               & "Message: " & e.Errors(i).Message & ControlChars.Cr _
    '                               & "NativeError: " & e.Errors(i).NativeError & ControlChars.Cr _
    '                               & "Source: " & e.Errors(i).Source & ControlChars.Cr _
    '                               & "SQLState: " & e.Errors(i).SQLState & ControlChars.Cr
    '            Next i
    '            Debug.Print(errorMessages)
    '            Throw New Exception(e.Message, e.InnerException)
    '        Catch ex As Exception
    '            Throw New Exception(ex.Message, ex.InnerException)
    '        End Try
    '    End Sub
    '    Public Overloads Sub DataSetLoad(ByVal ds As DataSet, ByVal sql As String, ByVal TableName As String, ByVal cn As OleDb.OleDbConnection)
    '        DataSetLoad(ds, sql, TableName, cn, False)
    '    End Sub


    '    Public Overloads Sub DataSetLoad(ByVal ds As DataSet, ByVal sql As String, ByVal TableName As String, ByVal sConnection As String)
    '        Dim conss As ConnectionStringSettings = ConfigurationManager.ConnectionStrings(sConnection)
    '        Dim ConnString As String = conss.ConnectionString
    '        Dim cn As OleDb.OleDbConnection
    '        cn = New OleDb.OleDbConnection(ConnString)
    '        If sql = "" Then
    '            sql = "SELECT * FROM " & TableName
    '        End If
    '        DataSetLoad(ds, sql, TableName, cn)
    '    End Sub

    '    Public Overloads Sub DataSetLoad(ByVal ds As DataSet, ByVal TableName As String, ByVal cn As OleDb.OleDbConnection)
    '        Dim sql As String
    '        sql = "SELECT * FROM " & TableName
    '        DataSetLoad(ds, sql, TableName, cn)
    '    End Sub

    '#End Region

    '#Region "DataSetSave"
    '    ''' <summary>
    '    ''' 
    '    ''' </summary>
    '    ''' <param name="ds">DataSet</param>
    '    ''' <param name="sTableName">DB Table name</param>
    '    ''' <param name="cn">DB source oledbconnection</param>
    '    ''' <param name="FlagFreeConnection">Boolean Value</param>
    '    ''' <param name="sTableMapping">DataTableMapping=(sourceTable As String, dataSetTable As String )  sourceTable:要對應的來源資料行之區分大小寫名稱。 dataSetTable要對應至的 DataSet 資料表名稱，不區分大小寫。</param>
    '    ''' <remarks></remarks>
    '    Public Sub DataSetSave(ByVal ds As DataSet, ByVal sTableName As String, ByVal cn As OleDb.OleDbConnection, ByVal FlagFreeConnection As Boolean, ByVal sTableMapping() As String)
    '        Dim da As OleDb.OleDbDataAdapter
    '        Try
    '            If cn.State <> ConnectionState.Open Then
    '                cn.Open()
    '            End If
    '            da = New OleDb.OleDbDataAdapter("SELECT * FROM " & sTableName, cn)
    '            Dim cb As New OleDb.OleDbCommandBuilder(da)

    '            'If IsNothing(da.UpdateCommand) Then
    '            '    BuildDACommand(da, ds.Tables(sTableName), cn)
    '            'End If

    '            If Not IsNothing(sTableMapping) AndAlso sTableMapping.Length = 2 Then
    '                da.TableMappings.Add(sTableMapping(0), sTableMapping(1))
    '            End If

    '            da.Update(ds, sTableName)
    '            If Not IsNothing(sTableMapping) AndAlso sTableMapping.Length = 2 Then
    '                ds.Tables(sTableMapping(1)).AcceptChanges()
    '            Else
    '                ds.Tables(sTableName).AcceptChanges()
    '            End If
    '        Catch ex As Exception
    '            Throw New Exception(ex.Message & vbCr & "Table name: [" & sTableName & "]. Please inform IT staff for help")
    '        Finally
    '            If FlagFreeConnection Then
    '                If Not IsNothing(cn) Then
    '                    cn.Close()
    '                End If
    '                cn = Nothing
    '            End If
    '        End Try
    '    End Sub

    '    Public Sub DataSetSave(ByVal ds As DataSet, ByVal sTableName As String, ByVal cn As OleDb.OleDbConnection, ByVal sTableMapping() As String)
    '        DataSetSave(ds, sTableName, cn, False, sTableMapping)
    '    End Sub

    '    Public Sub DataSetSave(ByVal ds As DataSet, ByVal sTableName As String, ByVal cn As OleDb.OleDbConnection)
    '        DataSetSave(ds, sTableName, cn, False, Nothing)
    '    End Sub

    '    Public Sub DataSetSave(ByVal ds As DataSet, ByVal sTableName As String, ByVal sConnection As String)
    '        Dim cn As OleDb.OleDbConnection
    '        'cn = CreateConnection(sConnection, "")
    '        Dim conss As ConnectionStringSettings = ConfigurationManager.ConnectionStrings(sConnection)
    '        Dim ConnString As String = conss.ConnectionString
    '        cn = New OleDb.OleDbConnection(ConnString)
    '        DataSetSave(ds, sTableName, cn)
    '    End Sub


    '    Public Sub DataSetSave(ByVal ds As DataSet, ByVal cn As OleDb.OleDbConnection)
    '        For Each dt As DataTable In ds.Tables
    '            DataSetSave(ds, dt.TableName, cn)
    '        Next
    '    End Sub

    '    Public Sub DataSetSave(ByVal ds As DataSet, ByVal sTableName As String, ByVal cn As OleDb.OleDbConnection, ByVal FlagFreeConnection As Boolean)
    '        DataSetSave(ds, sTableName, cn, FlagFreeConnection, New String() {})
    '    End Sub
    '#End Region

    '#Region "ExecSQLCmd"

    '    Public Sub ExecSQLCmd(ByVal sSQL As String, ByVal trn As OleDb.OleDbTransaction)
    '        ' Dim cmd As New OleDb.OleDbCommand
    '        'cmd.Connection = cn
    '        'cmd.CommandType = Type
    '        Using cmd As New OleDb.OleDbCommand
    '            cmd.Transaction = trn
    '            cmd.Connection = trn.Connection
    '            cmd.CommandText = sSQL
    '            'Debug.Print(sSQL + vbNewLine)
    '            cmd.ExecuteNonQuery()
    '        End Using

    '    End Sub

    '    Public Sub ExecSQLCmd(ByVal sSQL As String, ByVal cn As OleDb.OleDbConnection, ByVal FlagFreeConnection As Boolean)
    '        ExecSQLCmd(CommandType.Text, sSQL, cn, FlagFreeConnection)
    '    End Sub

    '    Public Sub ExecSQLCmd(ByVal type As CommandType, ByVal sSQL As String, ByVal cn As OleDb.OleDbConnection, ByVal FlagFreeConnection As Boolean)
    '        ExecSQLCmd(type, sSQL, cn, FlagFreeConnection, Nothing)
    '    End Sub

    '    Public Sub ExecSQLCmd(ByVal type As CommandType, ByVal sSQL As String, ByVal cn As OleDb.OleDbConnection, ByVal FlagFreeConnection As Boolean, ByVal iTimeout As Integer)
    '        Try
    '            If cn.State <> ConnectionState.Open Then
    '                cn.Open()
    '            End If

    '            Using cmd As New OleDb.OleDbCommand
    '                cmd.Connection = cn
    '                cmd.CommandType = type
    '                cmd.CommandText = sSQL
    '                If Not IsNothing(iTimeout) Then
    '                    cmd.CommandTimeout = iTimeout
    '                End If
    '                'Debug.Print(sSQL.ToString + vbNewLine)
    '                cmd.ExecuteNonQuery()
    '            End Using
    '        Catch ex As Exception
    '            'Dim sMsg As String = ex.Message
    '            If FlagFreeConnection AndAlso Not IsNothing(cn) Then
    '                cn.Close()
    '                cn = Nothing
    '            End If
    '            Throw New Exception(ex.Message & vbCr & "SQL String: [" & sSQL & "]. Please inform IT staff for help")
    '            'Throw New Exception(ex.Message)
    '        Finally
    '            If FlagFreeConnection Then
    '                cn.Close()
    '                cn = Nothing
    '            End If
    '        End Try
    '    End Sub
    '#End Region


End Class
