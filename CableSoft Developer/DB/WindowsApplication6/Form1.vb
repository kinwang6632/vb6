Imports System.Xml
Public Class Form1
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim x As New XmlDocument()

        x.Load("C:\Documents and Settings\Administrator\орн▒\Net\XML Sample.xml")
        Dim j As XmlNode = x.DocumentElement
        j.ChildNodes(0).Value = "888"


    End Sub
End Class
