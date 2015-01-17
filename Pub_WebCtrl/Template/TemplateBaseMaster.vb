Public MustInherit Class TemplateBaseMaster
    Inherits System.Web.UI.MasterPage
    Protected ELib As Pub_Entity.CommLib
    Protected CLib As CommLib
    Protected WLib As WebCommLib
    'Protected MustOverride ReadOnly Property PageTitle() As String
    Protected ReadOnly Property PageTitle As String
        Get
            Return Me.Page.Title
        End Get
    End Property


    Private Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ELib = New Pub_Entity.CommLib
        CLib = New CommLib
        WLib = New WebCommLib
    End Sub

    Protected Sub ShowMessage(ByVal s As String)
        MsgBox(s, MsgBoxStyle.OkOnly, PageTitle)
    End Sub
    Protected Sub ShowError(ByVal ex As Exception)
        MsgBox(ex.ToString, MsgBoxStyle.Critical, PageTitle)
    End Sub

End Class
