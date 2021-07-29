Imports System.Xml
Public Class Form1
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim x As New XmlDocument()
        x.Load("D:\CableSoft Document\Net\XML Sample2.xml")
        Dim j As XmlNode = x.SelectSingleNode("//EXT")
        If j.SchemaInfo.IsNil Then
            MsgBox("yes")
        Else
            j.InnerText = "888"
        End If
        x.Save("C:\TEST.XML")
        MsgBox("งนฆจ")
    End Sub
End Class
