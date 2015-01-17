Imports System.Web
Public Class WebCommLib
    Public Function DatatableToHTMLTable(ByVal dt As DataTable, ByVal dPageIndex As Integer, ByVal dPageSize As Integer) As String
        Dim sb As New System.Text.StringBuilder
        sb.Append("<TABLE>" & vbCrLf)
        sb.Append(vbTab & "<TR>" & vbCrLf)
        For Each dc As DataColumn In dt.Columns
            sb.Append(vbTab & vbTab & "<TD>" & vbCrLf)
            If dc.Caption <> "" Then
                sb.Append(dc.Caption)
            Else
                sb.Append(dc.ColumnName)
            End If
            sb.Append(vbTab & vbTab & "</TD>" & vbCrLf)
        Next
        sb.Append(vbTab & "</TR>" & vbCrLf)

        Dim dRowIndex As Integer = 0
        For Each dr As DataRow In dt.Rows
            dRowIndex += 1
            If dPageSize = 0 Or _
                (dRowIndex > (dPageIndex - 1) * dPageSize AndAlso dRowIndex <= dPageIndex * dPageSize) Then
                sb.Append(vbTab & "<TR>" & vbCrLf)
                For Each dc As DataColumn In dt.Columns
                    sb.Append(vbTab & vbTab & "<TD>" & vbCrLf)
                    sb.Append(dr.Item(dc.ColumnName).ToString)
                    sb.Append(vbTab & vbTab & "</TD>" & vbCrLf)
                Next
                sb.Append(vbTab & "</TR>" & vbCrLf)
            End If
        Next

        sb.Append("</TABLE>" & vbCrLf)
        Return sb.ToString
    End Function

    Public Function GetAnchor(ByVal sText As String, ByVal sLink As String) As String
        Dim s As String = ""
        s = "<a href='" & sLink & "'>" & sText & "</a>"
        Return s
    End Function

    Public Sub FillImageBinary(ByVal img As System.Web.UI.WebControls.Image, ByVal bi As Byte())
        Dim strBase64 As String = Convert.ToBase64String(bi, 0, bi.Length)
        img.ImageUrl = "data:image/png;base64, " & strBase64
    End Sub

    Public Function TransHTMLNewLine(ByVal sSrc As String) As String
        Return sSrc.Replace(vbCr, "<br />")
    End Function

    Public Function GetClientIP() As String
        Dim VisitorsIPAddr As String = ""
        If (Not IsNothing(HttpContext.Current.Request.ServerVariables("HTTP_X_FORWARDED_FOR"))) Then
            VisitorsIPAddr = HttpContext.Current.Request.ServerVariables("HTTP_X_FORWARDED_FOR").ToString()
        ElseIf (HttpContext.Current.Request.UserHostAddress.Length <> 0) Then
            VisitorsIPAddr = HttpContext.Current.Request.UserHostAddress
        End If

        Return VisitorsIPAddr

    End Function
End Class
