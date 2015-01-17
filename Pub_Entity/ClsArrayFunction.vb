
Public Class ClsArrayFunction

#Region "ToArray"

    Public Function ToArray(ByVal dt As System.Data.DataTable, ByVal FlagColumnInclude As Boolean) As Object(,)
        Return ToArray(dt, FlagColumnInclude, 0)
    End Function

    Public Function ToArray(ByVal dt As System.Data.DataTable, ByVal FlagColumnInclude As Boolean, ByVal StartIdx As Integer) As Object(,)
        Return ToArray(dt, FlagColumnInclude, StartIdx, 0, 0)
    End Function

    Public Function ToArray(ByVal dt As System.Data.DataTable, ByVal FlagColumnInclude As Boolean, ByVal StartIdx As Integer, ByVal RowLength As Integer, ByVal ColumnLength As Integer) As Object(,)
        Dim iRow As Integer = dt.Rows.Count
        Dim iCol As Integer = dt.Columns.Count
        Dim i As Integer, j As Integer

        Dim arr(,) As Object
        If RowLength <> 0 OrElse ColumnLength <> 0 Then
            ReDim arr(RowLength, ColumnLength)
        Else
            If FlagColumnInclude Then
                ReDim arr(iRow, iCol - 1)
            Else
                ReDim arr(iRow - 1, iCol - 1)
            End If
        End If

        If FlagColumnInclude Then
            For i = StartIdx To iCol - 1 + StartIdx
                arr(StartIdx, i) = dt.Columns(i - StartIdx).ColumnName
            Next
        End If
        For i = StartIdx To iRow - 1 + StartIdx
            'ReDim arr(i)(iCol)
            For j = StartIdx To iCol - 1 + StartIdx
                If FlagColumnInclude Then
                    arr(i + 1, j) = dt.Rows(i - StartIdx)(j - StartIdx)
                Else
                    arr(i, j) = dt.Rows(i - StartIdx)(j - StartIdx)
                End If
            Next
        Next
        Return arr
    End Function

    Public Function ToArray(ByVal oArr As Object(,)) As String(,)
        Dim s(,) As String
        ReDim s(oArr.GetUpperBound(0), oArr.GetUpperBound(1))
        Dim i As Integer
        Dim j As Integer
        For i = s.GetLowerBound(0) To s.GetUpperBound(0)
            For j = s.GetLowerBound(0) To s.GetUpperBound(1)
                If Not IsNothing(oArr(i, j)) Then
                    s(i, j) = oArr(i, j).ToString
                Else
                    s(i, j) = ""
                End If
            Next
        Next
        Return s
    End Function

    Public Function ToArrayDblM1(ByVal oTbl As DataTable, ByVal sCol As String) As Double()
        Return ToArrayDblM1(oTbl, sCol, 0)
    End Function

    Public Function ToArrayDblM1(ByVal oTbl As DataTable, ByVal sCol As String, ByVal StartIdx As Integer) As Double()
        Dim dM1() As Double = Nothing
        If Not IsNothing(oTbl) AndAlso oTbl.Columns.Contains(sCol) Then
            ReDim dM1(oTbl.Rows.Count - (1 - StartIdx))
            Dim i As Integer
            For i = 0 To oTbl.Rows.Count - 1
                Try
                    dM1(i + StartIdx) = CType(oTbl.Rows(i).Item(sCol), Double)
                Catch ex As Exception

                End Try
            Next
        End If
        Return dM1
    End Function

    Public Function ToArrayM1(ByVal oTbl As DataTable, ByVal sCol As String) As Object()
        Return ToArrayM1(oTbl, sCol, 0)
    End Function

    Public Function ToArrayM1(ByVal oTbl As DataTable, ByVal sCol As String, ByVal StartIdx As Integer) As Object()
        Dim dM1() As Object = Nothing
        If Not IsNothing(oTbl) AndAlso oTbl.Columns.Contains(sCol) Then
            ReDim dM1(oTbl.Rows.Count - (1 - StartIdx))
            Dim i As Integer
            For i = 0 To oTbl.Rows.Count - 1
                Try
                    dM1(i + StartIdx) = CType(oTbl.Rows(i).Item(sCol), Double)
                Catch ex As Exception

                End Try
            Next
        End If
        Return dM1
    End Function


#End Region

#Region "ToTable"

    Public Sub ToNewTable(ByVal dArr(,) As Object, ByVal dt As DataTable)
        ToNewTable(dArr, dt, 1, False, False)
    End Sub

    Public Sub ToNewTable(ByVal dArr(,) As Object, ByVal dt As DataTable, ByVal StartIdx As Integer, ByVal fNewColumn As Boolean, ByVal fHDR As Boolean)
        ToNewTable(dArr, dt, StartIdx, fNewColumn, fHDR, True)
    End Sub

    Public Sub ToNewTable(ByVal dArr(,) As Object, ByVal dt As DataTable, ByVal StartIdx As Integer, ByVal fNewColumn As Boolean, ByVal fHDR As Boolean, ByVal fSetDecColumn As Boolean)
        ToNewTable(dArr, dt, StartIdx, fNewColumn, fHDR, fSetDecColumn, 1)
    End Sub

    Public Sub ToNewTable(ByVal dArr(,) As Object, ByVal dt As DataTable, ByVal StartIdx As Integer, ByVal fNewColumn As Boolean, ByVal fHDR As Boolean, ByVal fSetDecColumn As Boolean, ByVal dNoNullCol As Integer)
        Dim i As Integer, j As Integer
        Try
            Dim ArrRowCnt As Integer = dArr.GetUpperBound(0)
            Dim ArrColCnt As Integer = dArr.GetUpperBound(1)
            Dim iHDR As Integer = 0


            If fHDR Then
                iHDR = 1
            End If

            If fNewColumn Then
                Dim t As System.Type
                For i = dt.Columns.Count + StartIdx To ArrColCnt
                    t = Type.GetType("System.String")
                    If fSetDecColumn Then
                        Dim fDec As Boolean = True
                        Dim isInt As Boolean = True
                        For j = StartIdx + iHDR To ArrRowCnt
                            If Not IsNothing(dArr(j, i)) AndAlso Not IsNumeric(dArr(j, i)) Then
                                fDec = False
                                Exit For
                            End If
                            If Not IsNothing(dArr(j, i)) AndAlso IsNumeric(dArr(j, i)) Then
                                If dArr(j, i).ToString.Split(".".ToCharArray).Length > 1 Then
                                    isInt = False
                                    Exit For
                                ElseIf Len(dArr(j, i).ToString) >= 10 Then
                                    isInt = False
                                    Exit For
                                End If
                            End If
                        Next
                        If fDec Then
                            If isInt = True Then
                                t = Type.GetType("System.Int64")
                            Else
                                t = Type.GetType("System.Double")
                            End If
                        End If
                    End If
                    If fHDR Then
                        If IsNothing(dArr(StartIdx, i)) Then
                            ArrColCnt = i - 1
                            Exit For
                        Else
                            dt.Columns.Add(dArr(StartIdx, i).ToString.Trim, t)
                        End If
                    Else
                        dt.Columns.Add()
                    End If
                Next
            End If

            Dim dr As DataRow = Nothing
            For i = StartIdx + iHDR To ArrRowCnt
                If Not IsNothing(dArr(i, dNoNullCol)) AndAlso dArr(i, dNoNullCol).ToString.Trim <> "" Then
                    dr = dt.NewRow()
                    For j = StartIdx To dt.Columns.Count - 1 + StartIdx
                        If j <= ArrColCnt AndAlso Not IsNothing(dArr(i, j)) Then
                            If LCase(dt.Columns(j - StartIdx).DataType.ToString) <> "system.string" AndAlso IsNumeric(dArr(i, j).ToString) Then
                                Try
                                    'dr.Item(j - StartIdx) = AddDot(dArr(i, j).ToString.Trim)
                                    dr.Item(j - StartIdx) = dArr(i, j)
                                Catch ex As Exception
                                    Throw New Exception(ex.Message)
                                End Try
                            Else
                                Try
                                    If LCase(dt.Columns(j - StartIdx).DataType.ToString) = "system.string" OrElse dArr(i, j).ToString.Trim <> "" Then
                                        dr.Item(j - StartIdx) = dArr(i, j).ToString.Trim
                                    End If
                                Catch ex As Exception
                                    Throw New Exception(ex.Message)
                                End Try
                            End If
                        End If
                    Next
                    dt.Rows.Add(dr)
                End If
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub ToNewTable(ByVal dArr(,) As Double, ByVal dt As DataTable, ByVal StartIdx As Integer, ByVal fNewColumn As Boolean, ByVal fHDR As Boolean, ByVal dNoNullCol As Integer)
        Dim i As Integer, j As Integer
        Try
            Dim ArrRowCnt As Integer = dArr.GetUpperBound(0)
            Dim ArrColCnt As Integer = dArr.GetUpperBound(1)
            Dim iHDR As Integer = 0

            If fNewColumn Then
                For i = dt.Columns.Count + StartIdx To ArrColCnt
                    If fHDR Then
                        dt.Columns.Add(dArr(StartIdx, i).ToString.Trim, Type.GetType("System.Double"))
                        iHDR = 1
                    Else
                        dt.Columns.Add("Col" & CStr(i), Type.GetType("System.Double"))
                    End If
                Next
            End If

            Dim dr As DataRow = Nothing
            For i = StartIdx + iHDR To ArrRowCnt
                If Not IsNothing(dArr(i, dNoNullCol)) AndAlso dArr(i, dNoNullCol).ToString.Trim <> "" Then
                    dr = dt.NewRow()
                    For j = StartIdx To dt.Columns.Count - 1 + StartIdx
                        If Not IsNothing(dArr(i, j)) Then
                            dr.Item(j - StartIdx) = dArr(i, j)
                        End If
                    Next
                    dt.Rows.Add(dr)
                End If
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub ToTable(ByVal dArr(,) As Object, ByVal dt As DataTable, ByVal StartRow As Integer, ByVal StartCol As Integer, ByVal RowCnt As Integer, ByVal ColCnt As Integer)
        Dim i As Integer, j As Integer
        Dim ArrRowCnt As Integer = dArr.GetUpperBound(0)
        Dim ArrColCnt As Integer = dArr.GetUpperBound(1)

        If RowCnt = 0 Then
            RowCnt = dt.Rows.Count
            For i = RowCnt To ArrRowCnt - 1
                dt.Rows.Add()
            Next
            RowCnt = dt.Rows.Count
            If RowCnt > ArrRowCnt Then
                RowCnt = ArrRowCnt
            End If
        End If
        If ColCnt = 0 Then
            ColCnt = dt.Columns.Count
            If ColCnt > ArrColCnt Then
                ColCnt = ArrColCnt
            End If
        End If
        For i = 1 To RowCnt
            For j = 1 To ColCnt
                dt.Rows(i + StartRow - 2).Item(j + StartCol - 2) = dArr(i, j)
            Next
        Next
    End Sub

    Public Sub ToTable(ByVal dArr(,) As Object, ByVal dt As DataTable)
        ToTable(dArr, dt, 1, 1, 0, 0)
    End Sub

    Public Sub ToTable(ByVal dArr(,) As Double, ByVal dt As DataTable, ByVal StartRow As Integer, ByVal StartCol As Integer, ByVal RowCnt As Integer, ByVal ColCnt As Integer)
        Dim i As Integer, j As Integer
        Dim ArrRowCnt As Integer = dArr.GetUpperBound(0)
        Dim ArrColCnt As Integer = dArr.GetUpperBound(1)

        If RowCnt = 0 Then
            RowCnt = dt.Rows.Count
            For i = RowCnt To ArrRowCnt - 1
                dt.Rows.Add()
            Next
            RowCnt = dt.Rows.Count
            If RowCnt > ArrRowCnt Then
                RowCnt = ArrRowCnt
            End If
        End If
        If ColCnt = 0 Then
            ColCnt = dt.Columns.Count
            If ColCnt > ArrColCnt Then
                ColCnt = ArrColCnt
            End If
        End If
        For i = 1 To RowCnt
            For j = 1 To ColCnt
                dt.Rows(i + StartRow - 2).Item(j + StartCol - 2) = dArr(i, j)
            Next
        Next
    End Sub

    Public Sub ToTable(ByVal dArr(,) As Double, ByVal dt As DataTable)
        ToTable(dArr, dt, 1, 1, 0, 0)
    End Sub

#End Region
    Private Function AddDot(ByVal s As String) As String
        Dim sVal As String
        Try
            sVal = ""
            If s = "" Then
            Else
                sVal = CDbl(s).ToString("###,###,###,###,##0.####")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return sVal
    End Function

#Region "ToCsvFile"
    Public Sub ToCsvFile(ByVal dt As DataTable, ByVal csvFilePath As String)
        Dim tableData As String(,) = ToArray(ToArray(dt, True, 0))
        Dim rowUpperBound As Integer = tableData.GetUpperBound(0)
        Dim colUpperBound As Integer = tableData.GetUpperBound(1)
        Dim f As New System.IO.FileStream(csvFilePath, IO.FileMode.Create)
        For i As Integer = 0 To rowUpperBound
            Dim csvText As String = ""
            For j As Integer = 0 To colUpperBound
                If csvText = "" Then
                    csvText = Chr(34) & tableData(i, j) & Chr(34)
                Else
                    csvText += "," & Chr(34) & tableData(i, j) & Chr(34)
                End If
            Next
            csvText += vbCrLf
            f.Write(CommLib.GetBytes(csvText), 0, CommLib.GetByteCount(csvText))
        Next
        f.Close()
    End Sub
#End Region
End Class
