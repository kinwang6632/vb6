Public Class clsDataToXML
    Implements IDisposable

    Public Sub Test()
        Dim objXml As New Xml.XmlTextWriter("C:\Test.xml", System.Text.Encoding.UTF8)
        objXml.Formatting = Xml.Formatting.Indented
        objXml.Indentation = 4
        objXml.WriteStartDocument()
        objXml.WriteStartElement("Invoice")
        objXml.WriteAttributeString("xsi:schemaLocation", "urn:GEINV:eInvoiceMessage:C0401:3.0 C0401.xsd")
        objXml.WriteAttributeString("xmlns", "urn:GEINV:eInvoiceMessage:C0401:3.0")
        objXml.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
        objXml.WriteStartElement("Main")
        objXml.WriteStartElement("InvoiceNumber")
        objXml.WriteValue("1111111")
        objXml.WriteEndElement()
        objXml.WriteEndElement()
        objXml.WriteEndDocument()
        objXml.Flush()
        objXml.Close()

        ' Writer.WriteStartDocument()  //開始寫xml，在最後有一個與之匹配的
        'Writer.WriteStartElement(”MapDefinition”)//生成一個節點

        'Writer.WriteAttributeString(”xmlns:xsi”, “http://www.w3.org/2001/XMLSchema-instance“)　//MapDefinition節點的屬性
        'Writer.WriteAttributeString("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
        'Writer.WriteAttributeString("xsi:noNamespaceSchemaLocation", "MapDefinition-1.0.0.xsd")
       
    End Sub
    Private Function WriteStartElement(ByVal aXML As Xml.XmlTextWriter, ByVal aElementName As String) As Boolean
        Try
            aXML.WriteStartElement(aElementName)

        Catch ex As Exception
            Throw New Exception("Function: WriteStartElement 訊息: " & ex.ToString)
        End Try
        Return True
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下面的 Finalize()。
            ' TODO: 將大型欄位設定為 null。
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: 只有當上面的 Dispose(ByVal disposing As Boolean) 有可釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 請勿變更此程式碼。在上面的 Dispose(ByVal disposing As Boolean) 中輸入清除程式碼。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' 由 Visual Basic 新增此程式碼以正確實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更此程式碼。在以上的 Dispose 置入清除程式碼 (ByVal 視為布林值處置)。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
