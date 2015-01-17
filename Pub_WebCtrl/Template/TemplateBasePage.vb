Imports System.Web.UI
Imports System.Web.UI.WebControls
Public MustInherit Class TemplateBasePage
    Inherits System.Web.UI.Page
    Protected ELib As Pub_Entity.CommLib
    Protected CLib As CommLib
    Protected WLib As WebCommLib

    Public MustOverride ReadOnly Property LoginSessionName As String
    Public MustOverride Sub LoginProcess()
    Protected Overridable ReadOnly Property MustLogin() As Boolean
        Get
            Return True
        End Get
    End Property

    Private Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ELib = New Pub_Entity.CommLib
        CLib = New CommLib
        WLib = New WebCommLib

        If IsNothing(Session(LoginSessionName)) AndAlso MustLogin Then
            LoginProcess()
        End If
    End Sub

    Protected Sub ShowMessage(ByVal s As String)
        ShowMessage(Nothing, s)
    End Sub

    Protected Sub ShowMessage(ByVal ctl As Control, ByVal s As String)
        ''MsgBox(s, MsgBoxStyle.OkOnly, Me.Title)
        'Dim AlertMsg As String = Replace(s, """", " ")
        'AlertMsg = Replace(AlertMsg, "''", " ")
        Dim AlertMsg As String = s
        AlertMsg = Replace(AlertMsg, Chr(13), " ")
        AlertMsg = Replace(AlertMsg, Chr(34), " ")
        AlertMsg = Replace(AlertMsg, Chr(39), " ")
        AlertMsg = "alert(" & ELib.AddQuote(AlertMsg) & ");"
        AddJavaScript(ctl, "AlertMsg", AlertMsg)
    End Sub

    Protected Sub ShowError(ByVal ex As Exception)
        ShowError(Nothing, ex)
    End Sub

    Protected Sub ShowError(ByVal ctl As Control, ByVal ex As Exception)
        'MsgBox(ex.ToString, MsgBoxStyle.Critical, Me.Title)
        Dim AlertMsg As String = ex.Message
        AlertMsg = Replace(AlertMsg, Chr(13), " ")
        AlertMsg = Replace(AlertMsg, Chr(34), " ")
        AlertMsg = Replace(AlertMsg, Chr(39), " ")
        Dim s As String = "alert(" & ELib.AddQuote(AlertMsg) & ");"

        AddJavaScript(ctl, "AlertMsg", s)
    End Sub

    Protected Sub AddJavaScript(ByVal sType As String, ByVal sScript As String)
        AddJavaScript(Nothing, sType, sScript)
    End Sub

    Protected Sub AddJavaScript(ByVal ctl As Control, ByVal sType As String, ByVal sScript As String)
        'Dim cs As ClientScriptManager = Page.ClientScript
        'If Not cs.IsClientScriptBlockRegistered("AlertMsg") Then
        '    cs.RegisterStartupScript(Me.GetType, "AlertMsg", "alert(" & ELib.AddQuote(AlertMsg) & ");", True)
        'End If
        If Not IsNothing(ctl) Then
            ScriptManager.RegisterStartupScript(ctl, ctl.GetType, sType, sScript, True)
        Else
            Dim cs As ClientScriptManager = Page.ClientScript
            If Not cs.IsClientScriptBlockRegistered(sType) Then
                cs.RegisterStartupScript(Me.GetType, sType, sScript, True)
            End If
        End If
    End Sub

    'Protected ReadOnly Property LoginUser() As String
    '    Get
    '        If Not IsNothing(Session("CMN_LoginUSer")) Then
    '            Return CType(Session("CMN_LoginUser"), String)
    '        Else
    '            Return ""
    '        End If
    '    End Get
    'End Property

    Private Sub Page_PreRender(sender As Object, e As EventArgs) Handles Me.PreRender
        PagePrerender(sender, e)
    End Sub

    Protected Overridable Sub PagePrerender(sender As Object, e As EventArgs)
        Dim s As String = "XX"
    End Sub

End Class
