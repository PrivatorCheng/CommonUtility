Public MustInherit Class ClsDataAccess
    Protected SLib As CommLib
    Protected ELib As Pub_Entity.CommLib
    Protected MustOverride Sub InitConn()
    Protected FlagTransaction As Boolean = True

    Public Sub New()
        SLib = New CommLib
        ELib = New Pub_Entity.CommLib
    End Sub

End Class
