Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Text
Imports System.Xml.Serialization

Public Class UtilObjSerialization
    Public Shared Function Obj2Xml(ByVal objType As Type, ByVal obj As Object) As String
        Try
            '將物件進行序列化
            Dim oText As String = String.Empty
            Dim mySerializer As XmlSerializer = New XmlSerializer(objType)
            Dim writer As New IO.StringWriter
            mySerializer.Serialize(writer, obj)
            oText = writer.ToString
            writer.Close()
            Return oText
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function Xml2Obj(ByVal objType As Type, ByVal oText As String) As Object
        Try
            '將Xml轉為物件並傳回
            Dim mySerializer As XmlSerializer = New XmlSerializer(objType)
            Dim reader As New IO.StringReader(oText)
            Return mySerializer.Deserialize(reader)
        Catch ex As Exception
            Throw ex
        End Try
    End Function


End Class
