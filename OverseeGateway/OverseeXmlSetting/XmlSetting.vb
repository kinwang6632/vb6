Imports System
Imports System.Text
Imports System.IO
Imports System.Xml
Imports CableSoft.NSTV.Encry.Encryption
Public Class XmlSetting
    Implements IDisposable
    Private FileName As String
    Private Const SettingName As String = "OverseeSetting.set"
    Private Const SODBName As String = "SODB"
    Private Const CommonName As String = "Common"
    Private Const EncryKey As String = "CABLESOFT1234567"
    Public Sub New()

    End Sub
    Public Sub New(ByVal FileName As String)
        Me.FileName = FileName
    End Sub
    Public Function WriteSetting(ByVal tbSODB As DataTable, ByVal tbCommon As DataTable) As Boolean
        If String.IsNullOrEmpty(FileName) Then
            FileName = String.Format("{0}\{1}",
                                     Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory),
                                     SettingName)
        End If
        Try
            Dim xmlWriter As New XmlTextWriter(FileName, New UTF8Encoding(False))
            Dim intSequence As Integer = 0
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 4
            xmlWriter.WriteStartDocument()
            xmlWriter.WriteStartElement("CableSoft")
            For Each rw As DataRow In tbSODB.Rows
                If Not DBNull.Value.Equals(rw.Item("CompCode")) Then
                    xmlWriter.WriteStartElement("SODB")                    
                    xmlWriter.WriteAttributeString("CompCode", rw.Item("CompCode").ToString)
                    xmlWriter.WriteAttributeString("SelectID", rw.Item("SelectID").ToString)
                    xmlWriter.WriteAttributeString("CompName", rw.Item("CompName").ToString)
                    xmlWriter.WriteAttributeString("prgName", rw.Item("prgName").ToString)
                    xmlWriter.WriteAttributeString("Sid", rw.Item("Sid").ToString)
                    xmlWriter.WriteAttributeString("DBUser", rw.Item("DBUser").ToString)
                    xmlWriter.WriteAttributeString("DBPassword", rw.Item("DBPassword").ToString)
                    If Not DBNull.Value.Equals(rw.Item("DBMinPool")) Then
                        xmlWriter.WriteAttributeString("DBMinPool", rw.Item("DBMinPool").ToString)
                    Else
                        xmlWriter.WriteAttributeString("DBMinPool", "1")
                    End If
                    If Not DBNull.Value.Equals(rw.Item("DBMaxPool")) Then
                        xmlWriter.WriteAttributeString("DBMaxPool", rw.Item("DBMaxPool").ToString)
                    Else
                        xmlWriter.WriteAttributeString("DBMaxPool", "10")
                    End If
                    If Not DBNull.Value.Equals(rw.Item("DBLifetime")) Then
                        xmlWriter.WriteAttributeString("DBLifetime", rw.Item("DBLifetime").ToString)
                    Else
                        xmlWriter.WriteAttributeString("DBLifetime", "5")
                    End If
                    xmlWriter.WriteAttributeString("MonitorType", rw.Item("MonitorType").ToString)
                    If Not DBNull.Value.Equals(rw.Item("Runtimer")) Then
                        xmlWriter.WriteAttributeString("Runtimer", rw.Item("Runtimer").ToString)
                    Else
                        xmlWriter.WriteAttributeString("Runtimer", "3")
                    End If
                    If Not DBNull.Value.Equals(rw.Item("Timeout")) Then
                        xmlWriter.WriteAttributeString("Timeout", rw.Item("Timeout").ToString)
                    Else
                        xmlWriter.WriteAttributeString("Timeout", "10")
                    End If
                    xmlWriter.WriteAttributeString("SeqSql", rw.Item("SeqSql").ToString)
                    xmlWriter.WriteAttributeString("InsSql", rw.Item("InsSql").ToString)
                    xmlWriter.WriteAttributeString("GetSql", rw.Item("GetSql").ToString)
                    xmlWriter.WriteAttributeString("WebServiceName", rw.Item("WebServiceName").ToString)
                    xmlWriter.WriteAttributeString("WebSrvMethod", rw.Item("WebSrvMethod").ToString)
                    xmlWriter.WriteAttributeString("WebServicePara", rw.Item("WebServicePara").ToString)
                    xmlWriter.WriteAttributeString("url", rw.Item("url").ToString)
                    xmlWriter.WriteAttributeString("WebGetPara", rw.Item("WebGetPara").ToString)
                    xmlWriter.WriteAttributeString("WebPostPara", rw.Item("WebPostPara").ToString)
                    ' xmlWriter.WriteAttributeString("Sequence", intSequence)
                    xmlWriter.WriteValue(intSequence)
                    intSequence += 1
                    xmlWriter.WriteEndElement()   'SODB
                End If
            Next
            xmlWriter.WriteStartElement("Common") 'Common
            If tbCommon.Rows.Count = 0 Then
                xmlWriter.WriteAttributeString("MaxThread", "10")
                xmlWriter.WriteAttributeString("ErrHistorical", "1")
            End If
            For Each rw As DataRow In tbCommon.Rows
                xmlWriter.WriteAttributeString("MaxThread", rw("MaxThread"))
                xmlWriter.WriteAttributeString("ErrHistorical", rw("ErrHistorical"))
            Next
            xmlWriter.WriteEndElement() 'Common
            xmlWriter.WriteEndElement() 'CableSoft
            xmlWriter.WriteEndDocument()
            xmlWriter.Flush()
            xmlWriter.Close()
            Dim strContext As String
            Using smRd As New System.IO.StreamReader(FileName)
                strContext = smRd.ReadToEnd
                strContext = CableSoft.NSTV.Encry.Encryption.Encrypt3DES(strContext, EncryKey)
                smRd.Close()
                smRd.Dispose()
            End Using
            Using smWr As New System.IO.StreamWriter(FileName)
                smWr.Write(strContext)
                smWr.Close()
                smWr.Dispose()
            End Using

        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try
        Return True
    End Function
    Public Function ReadSetting(ByVal tbSODB As DataTable, ByVal tbCommon As DataTable) As Boolean
        If String.IsNullOrEmpty(FileName) Then
            FileName = String.Format("{0}\{1}",
                                     Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory),
                                     SettingName)
        End If
        Try
            Using red As System.IO.TextReader = New System.IO.StreamReader(FileName, New UTF8Encoding(False))
                Dim context As String
                context = red.ReadToEnd
                context = CableSoft.NSTV.Encry.Encryption.Decrypt3DES(context, EncryKey)
                Using stm As Stream = New MemoryStream(Encoding.UTF8.GetBytes(context))
                    Dim xmlRed As New XmlTextReader(stm)
                    tbSODB.Rows.Clear()
                    tbCommon.Rows.Clear()
                    tbSODB.AcceptChanges()
                    tbCommon.AcceptChanges()
                    xmlRed.WhitespaceHandling = WhitespaceHandling.None
                    While xmlRed.Read
                        Select Case xmlRed.NodeType
                            Case XmlNodeType.Text

                            Case XmlNodeType.Element
                                Select Case xmlRed.Name.ToUpper
                                    Case "SODB".ToUpper
                                        If xmlRed.HasAttributes Then
                                            Dim rwSODB As DataRow = tbSODB.NewRow
                                                         
                                            If xmlRed.MoveToAttribute("SelectID") Then
                                                rwSODB.Item("SelectID") = xmlRed.GetAttribute("SelectID")
                                            Else
                                                rwSODB.Item("SelectID") = 0
                                            End If

                                            rwSODB.Item("CompCode") = xmlRed.GetAttribute("CompCode")
                                            rwSODB.Item("CompName") = xmlRed.GetAttribute("CompName")
                                            rwSODB.Item("prgName") = xmlRed.GetAttribute("prgName")
                                            rwSODB.Item("Sid") = xmlRed.GetAttribute("Sid")
                                            rwSODB.Item("DBUser") = xmlRed.GetAttribute("DBUser")
                                            rwSODB.Item("DBPassword") = xmlRed.GetAttribute("DBPassword")
                                            rwSODB.Item("DBMinPool") = xmlRed.GetAttribute("DBMinPool")
                                            rwSODB.Item("DBMaxPool") = xmlRed.GetAttribute("DBMaxPool")
                                            rwSODB.Item("DBLifetime") = xmlRed.GetAttribute("DBLifetime")
                                            rwSODB.Item("MonitorType") = xmlRed.GetAttribute("MonitorType")
                                            rwSODB.Item("Runtimer") = xmlRed.GetAttribute("Runtimer")
                                            rwSODB.Item("Timeout") = xmlRed.GetAttribute("Timeout")
                                            rwSODB.Item("SeqSql") = xmlRed.GetAttribute("SeqSql")
                                            rwSODB.Item("InsSql") = xmlRed.GetAttribute("InsSql")
                                            rwSODB.Item("GetSql") = xmlRed.GetAttribute("GetSql")
                                            rwSODB.Item("WebServiceName") = xmlRed.GetAttribute("WebServiceName")
                                            rwSODB.Item("WebSrvMethod") = xmlRed.GetAttribute("WebSrvMethod")
                                            rwSODB.Item("WebServicePara") = xmlRed.GetAttribute("WebServicePara")
                                            rwSODB.Item("url") = xmlRed.GetAttribute("url")
                                            rwSODB.Item("WebGetPara") = xmlRed.GetAttribute("WebGetPara")
                                            rwSODB.Item("WebPostPara") = xmlRed.GetAttribute("WebPostPara")
                                            Dim sequence As String = xmlRed.ReadString

                                            If Not String.IsNullOrEmpty(sequence) Then
                                                rwSODB.Item("Sequence") = Integer.Parse(sequence)
                                            End If
                                            tbSODB.Rows.Add(rwSODB)
                                            tbSODB.AcceptChanges()
                                        End If
                                    Case "Common".ToUpper
                                        If xmlRed.HasAttributes Then
                                            Dim rwCommon As DataRow = tbCommon.NewRow
                                            rwCommon.Item("MaxThread") = xmlRed.GetAttribute("MaxThread")
                                            rwCommon.Item("ErrHistorical") = xmlRed.GetAttribute("ErrHistorical")
                                            tbCommon.Rows.Add(rwCommon)
                                            tbCommon.AcceptChanges()
                                        End If
                                End Select

                            Case Else

                        End Select
                    End While
                    stm.Close()
                    stm.Dispose()
                End Using
                red.Close()
                red.Dispose()
            End Using
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try

        Return True
    End Function
    Public Function CreateCommon() As DataTable
        Using tbCommon As New DataTable(CommonName)
            tbCommon.Columns.Add("MaxThread", GetType(Integer))
            tbCommon.Columns.Add("ErrHistorical", GetType(Integer))
            Return tbCommon.Copy
        End Using
    End Function
    Public Function CreateSODB() As DataTable
        Using tbSO As New DataTable(SODBName)
            tbSO.Columns.Add("sequence", GetType(Integer))
            tbSO.Columns.Add("SelectID", GetType(Integer))
            tbSO.Columns.Add("CompCode", GetType(String))
            tbSO.Columns.Add("CompName", GetType(String))
            tbSO.Columns.Add("prgName", GetType(String))
            tbSO.Columns.Add("Sid", GetType(String))
            tbSO.Columns.Add("DBUser", GetType(String))
            tbSO.Columns.Add("DBPassword", GetType(String))
            tbSO.Columns.Add("DBMinPool", GetType(Integer))
            tbSO.Columns.Add("DBMaxPool", GetType(Integer))
            tbSO.Columns.Add("DBLifetime", GetType(Integer))
            tbSO.Columns.Add("MonitorType", GetType(String)) '0.Web Service  1.Post  2.Get  3.DB
            tbSO.Columns.Add("Runtimer", GetType(Integer))
            tbSO.Columns.Add("Timeout", GetType(Integer))
            tbSO.Columns.Add("SeqSql", GetType(String))
            tbSO.Columns.Add("InsSql", GetType(String))
            tbSO.Columns.Add("GetSql", GetType(String))
            tbSO.Columns.Add("WebServiceName", GetType(String))
            tbSO.Columns.Add("WebSrvMethod", GetType(String))
            tbSO.Columns.Add("WebServicePara", GetType(String))
            tbSO.Columns.Add("url", GetType(String))
            tbSO.Columns.Add("WebGetPara", GetType(String))
            tbSO.Columns.Add("WebPostPara", GetType(String))

            Return tbSO.Copy
        End Using
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
