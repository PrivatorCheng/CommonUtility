Public MustInherit Class ClsOLEDataAccess
    Inherits ClsDataAccess
    Protected MyTrn As OleDb.OleDbTransaction
    Protected MyConn As OleDb.OleDbConnection

    Protected Sub MyFinal()
        If Not IsNothing(MyTrn) Then
            MyTrn.Dispose()
            MyTrn = Nothing
        End If
        If Not IsNothing(MyConn) Then
            MyConn.Close()
            MyConn.Dispose()
            MyConn = Nothing
        End If
    End Sub
End Class
