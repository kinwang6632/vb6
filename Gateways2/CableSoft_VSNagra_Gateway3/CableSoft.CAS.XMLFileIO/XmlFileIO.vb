Imports System.Windows.Forms
Imports System.IO
Imports System.Text
Imports Cablesoft.CAS.CryptUtil
Public Class XmlFileIO
    Implements IDisposable
    Private Shared XmlFileName As String = "Setting.bin"
    Private Const Key As String = "CABLESOFT1234567"
    Private xmlDoc As Xml.XmlDocument = Nothing
    Private Const Transaction_Number_File As String = "Nagra_Transaction_Number.xml"
    Private Const Transaction_Number_ParentNote As String = "Nagra"
    Private Const Transaction_Number_ChildNote As String = "Transaction_Number"
    Private Shared xmlDBIdDoc As Xml.XmlDocument = Nothing
    Private Shared mtxWait As New System.Threading.Mutex
    Sub New()
        XmlFileName = Application.StartupPath & "\" & XmlFileName
    End Sub
    Sub New(ByVal XmlFile As String)
        XmlFileName = XmlFile
    End Sub
    '讀取GateWay系統參數
    Public Function ReadSystem(ByVal IsEncryFile As Boolean) As DataTable
        Dim tbSystem As New DataTable
        tbSystem.Columns.Add(New DataColumn("ReadTime", GetType(Int32)))
        tbSystem.Columns.Add(New DataColumn("ProcCmdCount", GetType(Int32)))
        tbSystem.Columns.Add(New DataColumn("MaxThread", GetType(Int32)))
        tbSystem.Columns.Add(New DataColumn("DBMinPool", GetType(Int32)))
        tbSystem.Columns.Add(New DataColumn("DBMaxPool", GetType(Int32)))
        tbSystem.Columns.Add(New DataColumn("DBPoolLiveTime", GetType(Int32)))
        Try
            Dim rwNew As DataRow = tbSystem.NewRow
            rwNew.Item("ReadTime") = Decimal.Parse(ReadXmlNodeValue("CableSoft/System", "ReadTime", IsEncryFile))
            rwNew.Item("ProcCmdCount") = Decimal.Parse(ReadXmlNodeValue("CableSoft/System", "ProcCmdCount", IsEncryFile))
            rwNew.Item("MaxThread") = Decimal.Parse(ReadXmlNodeValue("CableSoft/System", "MaxThread", IsEncryFile))
            rwNew.Item("DBMinPool") = Decimal.Parse(ReadXmlNodeValue("CableSoft/System", "DBMinPool", IsEncryFile))
            rwNew.Item("DBMaxPool") = Decimal.Parse(ReadXmlNodeValue("CableSoft/System", "DBMaxPool", IsEncryFile))
            rwNew.Item("DBPoolLiveTime") = Decimal.Parse(ReadXmlNodeValue("CableSoft/System", "DBPoolLiveTime", IsEncryFile))
            tbSystem.Rows.Add(rwNew)
            tbSystem.AcceptChanges()
        Catch ex As Exception
            Throw
        End Try
        Return tbSystem
    End Function
    '讀取系統台資訊
    Public Function ReadSO(ByVal IsEncryFile As Boolean) As DataTable
        Dim tbSO As New DataTable("SO")
        tbSO.Columns.Add(New DataColumn("IsSelect", GetType(String)))
        tbSO.Columns.Add(New DataColumn("CompCode", GetType(String)))
        tbSO.Columns.Add(New DataColumn("CompName", GetType(String)))
        tbSO.Columns.Add(New DataColumn("SID", GetType(String)))
        tbSO.Columns.Add(New DataColumn("User", GetType(String)))
        tbSO.Columns.Add(New DataColumn("Password", GetType(String)))

        Try
            Dim lstDBValue As List(Of Object) = ReadXmlNodeAttributes("CableSoft/SODB", "DB", "Value", IsEncryFile)
            Dim lstCompCode As List(Of Object) = ReadXmlNodeAttributes("CableSoft/SODB", "DB", "CompCode", IsEncryFile)
            Dim lstCompName As List(Of Object) = ReadXmlNodeValues("CableSoft/SODB", "DB", IsEncryFile)
            Dim lstSID As List(Of Object) = ReadXmlNodeAttributes("CableSoft/SODB", "DB", "SID", IsEncryFile)
            Dim lstUSER As List(Of Object) = ReadXmlNodeAttributes("CableSoft/SODB", "DB", "USER", IsEncryFile)
            Dim lstPassword As List(Of Object) = ReadXmlNodeAttributes("CableSoft/SODB", "DB", "PASSWORD", IsEncryFile)
            If lstDBValue IsNot Nothing Then
                For i As Int32 = 0 To lstDBValue.Count - 1
                    Dim data(5) As Object
                    data(0) = lstDBValue.Item(i)
                    data(1) = lstCompCode.Item(i)
                    data(2) = lstCompName.Item(i)
                    data(3) = lstSID.Item(i)
                    data(4) = lstUSER.Item(i)
                    data(5) = lstPassword.Item(i)
                    Dim rwNew As DataRow = tbSO.NewRow()
                    rwNew.Item("IsSelect") = Int32.Parse(data(0))
                    rwNew.Item("CompCode") = data(1)
                    rwNew.Item("CompName") = data(2)
                    rwNew.Item("SID") = data(3)
                    rwNew.Item("User") = data(4)
                    rwNew.Item("Password") = data(5)
                    tbSO.Rows.Add(rwNew)
                    tbSO.AcceptChanges()
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return tbSO
    End Function
    '讀取錯誤清單
    Public Function ReadErrorList(ByVal IsEncryFile As Boolean) As DataTable
        Dim tbError As New DataTable("Error")
        tbError.Columns.Add(New DataColumn("ErrorCode", GetType(String)))
        tbError.Columns.Add(New DataColumn("ErrorName", GetType(String)))

        Try
            Dim CodeNo As List(Of Object) = ReadXmlNodeAttributes("CableSoft/ErrorList", "Error", "ErrorCode", IsEncryFile)
            Dim Descript As List(Of Object) = ReadXmlNodeValues("CableSoft/ErrorList", "Error", IsEncryFile)
            Dim value As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            If CodeNo IsNot Nothing Then
                For i As Integer = 0 To CodeNo.Count - 1
                    value.Clear()
                    value.Add("ErrorCode", CodeNo.Item(i))
                    value.Add("ErrorName", Descript.Item(i))
                    Dim rwNew As DataRow = tbError.NewRow()
                    rwNew.Item("ErrorCode") = value.Item("ErrorCode")
                    rwNew.Item("ErrorName") = value.Item("ErrorName")
                    tbError.Rows.Add(rwNew)
                    tbError.AcceptChanges()
                Next
            End If

        Catch ex As Exception
            Throw
        End Try
        Return tbError
    End Function
    '讀取Nagra參數
    Public Function ReadNagra(ByVal IsEncryFile As Boolean) As DataTable
        Dim tbNagra As New DataTable("Nagra")
        tbNagra.Columns.Add(New DataColumn("IPAddress", GetType(String)))
        tbNagra.Columns.Add(New DataColumn("Port", GetType(Integer)))
        tbNagra.Columns.Add(New DataColumn("CheckTime", GetType(Integer)))
        tbNagra.Columns.Add(New DataColumn("SendTimeOut", GetType(Integer)))
        tbNagra.Columns.Add(New DataColumn("ReceiveTimeout", GetType(Integer)))
        tbNagra.Columns.Add(New DataColumn("DisconnectRetryTime", GetType(Integer)))
        tbNagra.Columns.Add(New DataColumn("UTC", GetType(Integer)))
        tbNagra.Columns.Add(New DataColumn("Encoding", GetType(String)))
        tbNagra.Columns.Add(New DataColumn("ProcKind", GetType(Integer)))
        tbNagra.Columns.Add(New DataColumn("RetryProcCount", GetType(Integer)))
        Try
            Dim rwNew As DataRow = tbNagra.NewRow
            rwNew.Item("CheckTime") = Integer.Parse(ReadXmlNodeValue("CableSoft/Nagra", "CheckTime", IsEncryFile))
            rwNew.Item("IPAddress") = ReadXmlNodeValue("CableSoft/Nagra", "IPAddress", IsEncryFile)
            rwNew.Item("Port") = Integer.Parse("0" & ReadXmlNodeValue("CableSoft/Nagra", "Port", IsEncryFile))
            rwNew.Item("SendTimeOut") = Decimal.Parse(ReadXmlNodeValue("CableSoft/Nagra", "SendTimeOut", IsEncryFile))
            rwNew.Item("ReceiveTimeout") = Decimal.Parse(ReadXmlNodeValue("CableSoft/Nagra", "ReceiveTimeout", IsEncryFile))
            rwNew.Item("RetryProcCount") = Decimal.Parse("0" & ReadXmlNodeValue("CableSoft/Nagra", "RetryProcCount", IsEncryFile))
            rwNew.Item("DisconnectRetryTime") = Decimal.Parse("0" & ReadXmlNodeValue("CableSoft/Nagra", "DisconnectRetryTime", IsEncryFile))

            rwNew.Item("UTC") = Decimal.Parse("0" & ReadXmlNodeValue("CableSoft/Nagra", "UTC", IsEncryFile))
            rwNew.Item("Encoding") = ReadXmlNodeValue("CableSoft/Nagra", "Encoding", IsEncryFile)
            rwNew.Item("ProcKind") = Integer.Parse("0" & ReadXmlNodeValue("CableSoft/Nagra", "ProcKind", IsEncryFile))
            tbNagra.Rows.Add(rwNew)
            tbNagra.AcceptChanges()
        Catch ex As Exception
            Throw
        End Try
        Return tbNagra
    End Function
    '讀取低階命令列表
    Public Function ReadLowCmd(ByVal IsEncryFile As Boolean) As DataTable
        Dim tbLowCMD As New DataTable("LowCmd")
        tbLowCMD.Columns.Add(New DataColumn("CodeNo", GetType(String)))
        tbLowCMD.Columns.Add(New DataColumn("MsgId", GetType(String)))
        tbLowCMD.Columns.Add(New DataColumn("MsgName", GetType(String)))
        tbLowCMD.Columns.Add(New DataColumn("DefaultValue", GetType(String)))
        tbLowCMD.Columns.Add(New DataColumn("RealField", GetType(String)))
        Try
            Dim CodeNo As List(Of Object) = ReadXmlNodeAttributes("CableSoft/LowCmd", "MsgId", "CodeNo", IsEncryFile)
            Dim DefaultValue As List(Of Object) = ReadXmlNodeAttributes("CableSoft/LowCmd", "MsgId", "DefaultValue", IsEncryFile)
            Dim RealField As List(Of Object) = ReadXmlNodeAttributes("CableSoft/LowCmd", "MsgId", "RealField", IsEncryFile)
            Dim Value As List(Of Object) = ReadXmlNodeAttributes("CableSoft/LowCmd", "MsgId", "Value", IsEncryFile)
            Dim Descript As List(Of Object) = ReadXmlNodeValues("CableSoft/LowCmd", "MsgId", IsEncryFile)
            If CodeNo IsNot Nothing Then
                For i As Int32 = 0 To CodeNo.Count - 1
                    Dim data(4) As Object
                    data(0) = CodeNo.Item(i)
                    data(1) = Value.Item(i)
                    data(2) = Descript.Item(i)
                    If DefaultValue.Count >= i Then
                        data(3) = DefaultValue.Item(i)
                    End If
                    If RealField.Count >= i Then
                        data(4) = RealField.Item(i)
                    End If
                    Dim rwNew As DataRow = tbLowCMD.NewRow
                    rwNew.Item("CodeNo") = data(0)
                    rwNew.Item("MsgId") = data(1)
                    rwNew.Item("MsgName") = data(2)
                    rwNew.Item("DefaultValue") = data(3)
                    rwNew.Item("RealField") = data(4)
                    tbLowCMD.Rows.Add(rwNew)
                    tbLowCMD.AcceptChanges()
                Next
            End If

        Catch ex As Exception
            Throw
        End Try
        Return tbLowCMD
    End Function
    '讀取高階命令列表
    Public Function ReadHightCmd(ByVal IsEncryFile As Boolean) As DataTable
        Dim tbHighCMD As New DataTable("HighCMD")
        tbHighCMD.Columns.Add(New DataColumn("CodeNo", GetType(String)))
        tbHighCMD.Columns.Add(New DataColumn("MsgName", GetType(String)))
        tbHighCMD.Columns.Add(New DataColumn("RunLowCmd", GetType(String)))
        tbHighCMD.Columns.Add(New DataColumn("RunLowCmd2", GetType(String)))

        Try
            Dim CodeNo As List(Of Object) = ReadXmlNodeAttributes("CableSoft/HightCmd", "Cmd", "CodeNo", IsEncryFile)
            Dim Cmd As List(Of Object) = ReadXmlNodeAttributes("CableSoft/HightCmd", "Cmd", "RunLowCmd", IsEncryFile)
            Dim Cmd2 As List(Of Object) = ReadXmlNodeAttributes("CableSoft/HightCmd", "Cmd", "RunLowCmd2", IsEncryFile)
            Dim Descript As List(Of Object) = ReadXmlNodeValues("CableSoft/HightCmd", "Cmd", IsEncryFile)
            If CodeNo IsNot Nothing Then
                For i As Int32 = 0 To CodeNo.Count - 1
                    Dim data(3) As Object
                    data(0) = CodeNo.Item(i)
                    data(1) = Descript.Item(i)
                    data(2) = Cmd.Item(i)
                    If Cmd2 IsNot Nothing AndAlso Cmd2.Count >= i Then
                        data(3) = Cmd2.Item(i)
                    Else
                        data(3) = DBNull.Value
                    End If
                    Dim rwNew As DataRow = tbHighCMD.NewRow
                    rwNew("CodeNo") = data(0)
                    rwNew("RunLowCmd") = data(2)
                    rwNew("MsgName") = data(1)
                    rwNew("RunLowCmd2") = data(3)
                    tbHighCMD.Rows.Add(rwNew)
                    tbHighCMD.AcceptChanges()
                Next
            End If

        Catch ex As Exception
            Throw
        End Try
        Return tbHighCMD
    End Function
    '讀取SourceId資訊
    Public Function ReadServer(ByVal IsEncryFile As Boolean, ByVal Server As String) As DataTable
        Dim tbServer As New DataTable("Server")
        tbServer.Columns.Add(New DataColumn("SourceId", GetType(String)))
        tbServer.Columns.Add(New DataColumn("DestId", GetType(String)))
        tbServer.Columns.Add(New DataColumn("MoppId", GetType(String)))
        Try
            Dim MoppId As List(Of Object) = ReadXmlNodeAttributes("CableSoft/Nagra", "Server", "MoppId", IsEncryFile)
            Dim DestId As List(Of Object) = ReadXmlNodeAttributes("CableSoft/Nagra", "Server", "DestId", IsEncryFile)
            Dim SourceId As List(Of Object) = ReadXmlNodeValues("CableSoft/Nagra", "Server", IsEncryFile)
            Dim value As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            If MoppId IsNot Nothing Then
                For i As Int32 = 0 To MoppId.Count - 1
                    value.Clear()
                    value.Add("SourceId", SourceId.Item(i))
                    value.Add("MoppId", MoppId.Item(i))
                    value.Add("DestId", DestId.Item(i))
                    'Dim data(2) As Object
                    'If String.IsNullOrEmpty(MoppId.Item(i).ToString) Then
                    '    data(0) = "127.0.0.1"
                    'Else
                    '    data(0) = MoppId.Item(i)
                    'End If
                    'If String.IsNullOrEmpty(DestId.Item(i).ToString) Then
                    '    data(1) = "7364"
                    'Else
                    '    data(1) = DestId.Item(i)
                    'End If

                    'If String.IsNullOrEmpty(SourceId.Item(i).ToString) Then
                    '    data(2) = "ABCDEFGHABCDEFGH"
                    'Else
                    '    data(2) = SourceId.Item(i)
                    'End If
                    Dim rwNew As DataRow = tbServer.NewRow
                    rwNew.Item("SourceId") = value("SourceId")
                    rwNew.Item("DestId") = value("DestId")
                    rwNew.Item("MoppId") = value("MoppId")
                    tbServer.Rows.Add(rwNew)
                    tbServer.AcceptChanges()
                    'DataToTreList(TreLstSourceId, data, i)
                Next
            End If

        Catch ex As Exception
            Throw
        End Try
        Return tbServer
    End Function
    '重新讀取XML檔案
    Public Function RetryReadXml(ByVal IsDecry As Boolean) As Boolean
        Try
            xmlDoc = Nothing
            xmlDoc = New Xml.XmlDocument()
            Using stm As Stream = New FileStream(XmlFileName,
                                                 FileMode.Open, FileAccess.ReadWrite)
                Using rder As New System.IO.StreamReader(stm)
                    If IsDecry Then
                        xmlDoc.LoadXml(CableSoft.NSTV.Encry.Encryption.Decrypt3DES(rder.ReadToEnd(), Key))
                    Else
                        xmlDoc.LoadXml(rder.ReadToEnd())
                    End If
                End Using
            End Using
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Public Function ReadXmlAttributeValue(ByVal ParentNode As String, ByVal NodeName As String,
                                          ByVal AttributeName As String, ByVal IsDecry As Boolean) As Object
        Dim aResult As String = Nothing
        Try
            If xmlDoc Is Nothing Then

                xmlDoc = New Xml.XmlDocument()
                Using stm As Stream = New FileStream(XmlFileName,
                                                     FileMode.Open, FileAccess.ReadWrite)
                    Using rder As New System.IO.StreamReader(stm)
                        If IsDecry Then
                            Dim redString As String = rder.ReadToEnd()
                            Dim xmlStr As String = CableSoft.NSTV.Encry.Encryption.Decrypt3DES(redString, Key)
                            xmlDoc.LoadXml(xmlStr)
                        Else
                            xmlDoc.LoadXml(rder.ReadToEnd())
                        End If

                    End Using
                End Using
            End If

            Dim Node As Xml.XmlNode = xmlDoc.SelectSingleNode(ParentNode)
            If Node Is Nothing Then
                Return Nothing
            End If
            Dim ChildNode As Xml.XmlNode = Node.SelectSingleNode(NodeName)

            If ChildNode Is Nothing Then
                Return Nothing
            End If
            Dim xmlAtr As Xml.XmlAttribute = ChildNode.Attributes(AttributeName)
            If xmlAtr Is Nothing Then
                Return Nothing
            End If
            aResult = xmlAtr.Value
        Catch ex As Exception
            Throw
            'MessageBox.Show(ex.ToString, "訊息", MessageBoxButtons.OK)
        End Try
        Return aResult
    End Function
    Private Function ReadNodeName(ByVal lstNode As Xml.XmlNodeList, ByVal NodeName As String) As String
        Dim aResult As String = Nothing
        For Each node As Xml.XmlNode In lstNode
            If node.Name.ToUpper = NodeName.ToUpper Then
                aResult = node.Name
            Else
                ReadNodeName(node.ChildNodes, NodeName)
            End If
        Next
        Return aResult
    End Function

    Public Function WriteXmlNodeValue(ByVal ParentNode As String, ByVal NodeName As String,
                            ByVal Value As Object,
                              ByVal IsDecry As Boolean) As Boolean
        'mtxWait.WaitOne()
        Dim aResult As String = Nothing

        Try
            If xmlDoc Is Nothing Then
                xmlDoc = New Xml.XmlDocument()
                xmlDoc.XmlResolver = Nothing


                Using stm As Stream = New FileStream(XmlFileName,
                                                     FileMode.OpenOrCreate, FileAccess.ReadWrite)

                    Using rder As New System.IO.StreamReader(stm, New UTF8Encoding(False))
                        If IsDecry Then
                            Dim redString As String = rder.ReadToEnd()
                            'Dim xmlStr As String = CableSoft.NSTV.Encry.Encryption.Decrypt3DES(redString, Key)
                            Dim xmlStr As String = CableSoft.NSTV.Encry.Encryption.Decrypt3DES(redString, Key)
                        Else
                            xmlDoc.LoadXml(rder.ReadToEnd())
                            '照他的嘗試，他只要先將第一個字元強迫刪掉，
                            '也就是寫成 xmldoc.LoadXml(s.Substring(1)); ，
                            '問題雖然解決，但確不知道為什麼，
                            '而且這樣寫很low，我們討論了一會兒，
                            '猜測是否UTF8文件的前置表頭(BOM)搞鬼，
                            '沒多久他就找到了另一個解決方法，
                            '將 settings.Encoding = Encoding.UTF8;
                            ' 改成 settings.Encoding = new UTF8Encoding(false);，
                            '雖然目的都是要避開因為utf8表頭造成的解析錯誤，
                            '但至少這樣的寫法是明確且易懂的。
                            '我個人會希望 XmlDocument 要再聰明一點。
                        End If

                    End Using
                End Using
            End If


            Dim Node As Xml.XmlNode = xmlDoc.SelectSingleNode(ParentNode)
            If Node Is Nothing Then
                Return True
            End If
            Dim ChildNode As Xml.XmlNode = Node.SelectSingleNode(NodeName)

            If ChildNode Is Nothing Then
                Return True
            End If
            ChildNode.InnerText = Value

            xmlDoc.Save(XmlFileName)
        Catch ex As Exception
            MessageBox.Show(ex.ToString, "訊息", MessageBoxButtons.OK)
        Finally
            'mtxWait.ReleaseMutex()
        End Try
        Return True
    End Function
    Public Function ReadXmlNodeAttributes(ByVal ParentNode As String, ByVal NodeName As String,
                              ByVal AttributeName As String, ByVal IsDecry As Boolean) As List(Of Object)
        Dim aResult As New List(Of Object)
        Try
            If xmlDoc Is Nothing Then
                xmlDoc = New Xml.XmlDocument()
                xmlDoc.XmlResolver = Nothing


                Using stm As Stream = New FileStream(XmlFileName,
                                                     FileMode.OpenOrCreate, FileAccess.ReadWrite)


                    Using rder As New System.IO.StreamReader(stm, New UTF8Encoding(False))
                        If IsDecry Then
                            'Dim xmlStr As String = UTF8Encoding.UTF8.GetString(Convert.FromBase64String(CableSoft.NSTV.Encry.Encryption.Decrypt3DES(rder.ReadToEnd(), Key)))
                            Dim redString As String = rder.ReadToEnd()
                            Dim xmlStr As String = CableSoft.NSTV.Encry.Encryption.Decrypt3DES(redString, Key)
                            xmlDoc.LoadXml(xmlStr)
                        Else
                            xmlDoc.LoadXml(rder.ReadToEnd())
                        End If
                    End Using
                End Using
            End If

            Dim Node As Xml.XmlNode = xmlDoc.SelectSingleNode(ParentNode)
            If Node Is Nothing Then
                Return Nothing
            End If
            Dim lstChildNode As Xml.XmlNodeList = Node.SelectNodes(NodeName)


            If lstChildNode Is Nothing OrElse lstChildNode.Count = 0 Then
                Return Nothing
            End If
            For Each aNode As Xml.XmlNode In lstChildNode
                For i As Int32 = 0 To aNode.ChildNodes.Count - 1
                    Dim aAttribute As Xml.XmlAttribute = aNode.Attributes(AttributeName)
                    If aAttribute IsNot Nothing Then
                        aResult.Add(aAttribute.Value)
                    Else
                        aResult.Add(String.Empty)
                    End If
                Next

            Next
            'aResult = ChildNode.InnerText
        Catch ex As Exception
            Throw ex
            'MessageBox.Show(ex.ToString, "訊息", MessageBoxButtons.OK)
        End Try
        Return aResult
    End Function
    Public Function ReadXmlNodeValues(ByVal ParentNode As String, ByVal NodeName As String,
                              ByVal IsDecry As Boolean) As List(Of Object)
        Dim aResult As New List(Of Object)
        Try
            If xmlDoc Is Nothing Then
                xmlDoc = New Xml.XmlDocument()
                xmlDoc.XmlResolver = Nothing


                Using stm As Stream = New FileStream(XmlFileName,
                                                     FileMode.OpenOrCreate, FileAccess.ReadWrite)


                    Using rder As New System.IO.StreamReader(stm, New UTF8Encoding(False))
                        If IsDecry Then
                            'Dim xmlStr As String = UTF8Encoding.UTF8.GetString(Convert.FromBase64String(CableSoft.NSTV.Encry.Encryption.Decrypt3DES(rder.ReadToEnd(), Key)))
                            Dim redString As String = rder.ReadToEnd()
                            Dim xmlStr As String = CableSoft.NSTV.Encry.Encryption.Decrypt3DES(redString, Key)
                            xmlDoc.LoadXml(xmlStr)
                        Else
                            xmlDoc.LoadXml(rder.ReadToEnd())
                        End If
                    End Using
                End Using
            End If

            Dim Node As Xml.XmlNode = xmlDoc.SelectSingleNode(ParentNode)
            If Node Is Nothing Then
                Return Nothing
            End If
            Dim lstChildNode As Xml.XmlNodeList = Node.SelectNodes(NodeName)


            If lstChildNode Is Nothing OrElse lstChildNode.Count = 0 Then
                Return Nothing
            End If
            For Each aNode As Xml.XmlNode In lstChildNode
                For i As Int32 = 0 To aNode.ChildNodes.Count - 1
                    aResult.Add(aNode.ChildNodes(i).InnerText)
                Next

            Next
            'aResult = ChildNode.InnerText
        Catch ex As Exception
            Throw
            'MessageBox.Show(ex.ToString, "訊息", MessageBoxButtons.OK)
        End Try
        Return aResult
    End Function
    Public Function ReadXmlNodeValue(ByVal ParentNode As String, ByVal NodeName As String,
                              ByVal IsDecry As Boolean) As Object
        Dim aResult As String = Nothing
        Try
            If xmlDoc Is Nothing Then
                xmlDoc = New Xml.XmlDocument()
                xmlDoc.XmlResolver = Nothing


                Using stm As Stream = New FileStream(XmlFileName,
                                                     FileMode.OpenOrCreate, FileAccess.ReadWrite)


                    Using rder As New System.IO.StreamReader(stm, New UTF8Encoding(False))
                        If IsDecry Then

                            'Dim xmlStr As String = CableSoft.CAS.CryptUtil.CryptUtil.Decrypt(rder.ReadToEnd())
                            'Dim xmlStr As String = UTF8Encoding.UTF8.GetString(Convert.FromBase64String(CableSoft.NSTV.Encry.Encryption.Decrypt3DES(rder.ReadToEnd(), Key)))
                            Dim redString As String = rder.ReadToEnd()
                            Dim xmlStr As String = CableSoft.NSTV.Encry.Encryption.Decrypt3DES(redString, Key)
                            xmlDoc.LoadXml(xmlStr)
                        Else
                            xmlDoc.LoadXml(rder.ReadToEnd())
                            '照他的嘗試，他只要先將第一個字元強迫刪掉，
                            '也就是寫成 xmldoc.LoadXml(s.Substring(1)); ，
                            '問題雖然解決，但確不知道為什麼，
                            '而且這樣寫很low，我們討論了一會兒，
                            '猜測是否UTF8文件的前置表頭(BOM)搞鬼，
                            '沒多久他就找到了另一個解決方法，
                            '將 settings.Encoding = Encoding.UTF8;
                            ' 改成 settings.Encoding = new UTF8Encoding(false);，
                            '雖然目的都是要避開因為utf8表頭造成的解析錯誤，
                            '但至少這樣的寫法是明確且易懂的。
                            '我個人會希望 XmlDocument 要再聰明一點。
                        End If

                    End Using
                End Using
            End If

            'Debug.Print(ReadNodeName(xmlDoc.ChildNodes, NodeName))
            'ParentNode = ReadNodeName(xmlDoc.ChildNodes, ParentNode)
            'NodeName = ReadNodeName(xmlDoc.ChildNodes, NodeName)
            Dim Node As Xml.XmlNode = xmlDoc.SelectSingleNode(ParentNode)
            If Node Is Nothing Then
                Return Nothing
            End If
            Dim ChildNode As Xml.XmlNode = Node.SelectSingleNode(NodeName)

            If ChildNode Is Nothing Then
                Return Nothing
            End If
            aResult = ChildNode.InnerText
        Catch ex As Exception
            Throw ex
            'MessageBox.Show(ex.ToString, "訊息", MessageBoxButtons.OK)
        End Try
        Return aResult
    End Function
    Public Shared Function Create_Transaction_Number_File() As Boolean
        Try
            Try
                Using mem As New MemoryStream
                    Using XmlFile As New Xml.XmlTextWriter(mem, New UTF8Encoding(False))
                        XmlFile.Formatting = Xml.Formatting.Indented
                        XmlFile.Indentation = 4
                        XmlFile.WriteStartDocument()
                        XmlFile.WriteStartElement(Transaction_Number_ParentNote)  'Nagra
                        XmlFile.WriteStartElement(Transaction_Number_ChildNote) 'Transaction_Number
                        XmlFile.WriteAttributeString("Max", "999999999")
                        XmlFile.WriteValue("1")
                        XmlFile.WriteEndElement() 'Transaction_Number
                        XmlFile.WriteEndElement() 'Nagra
                        XmlFile.WriteEndDocument()
                        XmlFile.Flush()
                        XmlFile.Close()
                        Using stm As StreamWriter = New StreamWriter(String.Format("{0}{1}",
                                                                                   AppDomain.CurrentDomain.BaseDirectory,
                                                                                Transaction_Number_File), False, Encoding.UTF8)
                            stm.WriteLine(Encoding.UTF8.GetString(mem.ToArray))
                            stm.Flush()
                            stm.Close()
                            stm.Dispose()
                        End Using
                    End Using
                    mem.Close()
                    mem.Dispose()
                End Using
                Return True
            Catch ex As Exception
                Throw
            End Try
            Return True
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Public Overloads Shared Function Read_Transaction_Number() As UInteger
        Return Read_Transaction_Number(String.Format("{0}{1}",
                                       AppDomain.CurrentDomain.BaseDirectory,
                                       Transaction_Number_File),
                                   Transaction_Number_ParentNote,
                                   Transaction_Number_ChildNote)
    End Function
    Public Overloads Shared Function Read_Transaction_Number(ByVal xmlDBIdFile As String,
                                    ByVal ParentNodeName As String,
                                    ByVal ChildNodeName As String) As UInteger
        Dim aResult As String = Nothing
        Try
            If xmlDBIdDoc Is Nothing Then
                xmlDBIdDoc = New Xml.XmlDocument()
                xmlDBIdDoc.XmlResolver = Nothing


                Using stm As Stream = New FileStream(xmlDBIdFile,
                                                     FileMode.OpenOrCreate, FileAccess.ReadWrite)


                    Using rder As New System.IO.StreamReader(stm, New UTF8Encoding(False))

                        xmlDBIdDoc.LoadXml(rder.ReadToEnd())
                        '照他的嘗試，他只要先將第一個字元強迫刪掉，
                        '也就是寫成 xmldoc.LoadXml(s.Substring(1)); ，
                        '問題雖然解決，但確不知道為什麼，
                        '而且這樣寫很low，我們討論了一會兒，
                        '猜測是否UTF8文件的前置表頭(BOM)搞鬼，
                        '沒多久他就找到了另一個解決方法，
                        '將 settings.Encoding = Encoding.UTF8;
                        ' 改成 settings.Encoding = new UTF8Encoding(false);，
                        '雖然目的都是要避開因為utf8表頭造成的解析錯誤，
                        '但至少這樣的寫法是明確且易懂的。
                        '我個人會希望 XmlDocument 要再聰明一點。


                    End Using
                End Using
            End If


            Dim Node As Xml.XmlNode = xmlDBIdDoc.SelectSingleNode(ParentNodeName)
            If Node Is Nothing Then
                Return Nothing
            End If
            Dim ChildNode As Xml.XmlNode = Node.SelectSingleNode(ChildNodeName)

            If ChildNode Is Nothing Then
                Return Nothing
            End If
            aResult = ChildNode.InnerText
        Catch ex As Exception
            Throw ex
        End Try
        Return aResult
    End Function
    Public Shared Function Write_Transaction_Number(ByVal Value As UInteger) As Boolean
        Return Write_Transaction_Number(String.Format("{0}{1}",
                                               AppDomain.CurrentDomain.BaseDirectory,
                                               Transaction_Number_File),
                                           Value,
                                           Transaction_Number_ParentNote,
                                           Transaction_Number_ChildNote)
    End Function
    Public Shared Function Write_Transaction_Number(ByVal xmlDBIdFile As String,
                                     ByVal Value As UInteger,
                                    ByVal ParentNodeName As String,
                                    ByVal ChildNodeName As String) As Boolean
        'If mtxWait Is Nothing Then
        '    mtxWait = New System.Threading.Mutex()
        'End If

        mtxWait.WaitOne()

        Try
            If xmlDBIdDoc Is Nothing Then
                xmlDBIdDoc = New Xml.XmlDocument()
                xmlDBIdDoc.XmlResolver = Nothing

                Using stm As Stream = New FileStream(XmlFileName,
                                                     FileMode.OpenOrCreate, FileAccess.ReadWrite)

                    Using rder As New System.IO.StreamReader(stm, New UTF8Encoding(False))
                        xmlDBIdDoc.LoadXml(rder.ReadToEnd())
                    End Using
                End Using
            End If
            Dim Node As Xml.XmlNode = xmlDBIdDoc.SelectSingleNode(ParentNodeName)
            If Node Is Nothing Then
                Return True
            End If
            Dim ChildNode As Xml.XmlNode = Node.SelectSingleNode(ChildNodeName)

            If ChildNode Is Nothing Then
                Return True
            End If
            ChildNode.InnerText = Value

            xmlDBIdDoc.Save(xmlDBIdFile)
        Catch ex As Exception
            Throw
            'MessageBox.Show(ex.ToString, "訊息", MessageBoxButtons.OK)
        Finally
            mtxWait.ReleaseMutex()
        End Try
        Return True


    End Function


    Public Function WriteXmlFile(ByVal FileName As String,
                                  ByVal stmWrite As MemoryStream, ByVal IsEncryFile As Boolean) As Boolean
        Try
            Using stm As StreamWriter = New StreamWriter(FileName, False, Encoding.UTF8)
                If IsEncryFile Then
                    'Dim strWrite As String = CableSoft.CAS.Encry.Encryption.Encrypt3DES(
                    '    Convert.ToBase64String(Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(stmWrite.ToArray))), Key)
                    Dim strWrite As String = CableSoft.NSTV.Encry.Encryption.Encrypt3DES(
                       Encoding.UTF8.GetString(stmWrite.ToArray), Key)
                    stm.WriteLine(strWrite)
                    stm.Flush()
                    stm.Close()
                Else
                    stm.WriteLine(Encoding.UTF8.GetString(stmWrite.ToArray))
                    stm.Flush()
                    stm.Close()
                End If
            End Using
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function

    Public Function CreateDefaultXMLFile(ByVal IsEncryFile As Boolean) As Boolean
        Using mem As New MemoryStream




            Using XmlFile As New Xml.XmlTextWriter(mem, New UTF8Encoding(False))
                XmlFile.Formatting = Xml.Formatting.Indented
                XmlFile.Indentation = 4
                XmlFile.WriteStartDocument()
                XmlFile.WriteStartElement("CableSoft")  'CableSoft
                XmlFile.WriteStartElement("System")     'System
                XmlFile.WriteStartElement("ReadTime")
                XmlFile.WriteValue(3)
                XmlFile.WriteEndElement() 'ReadTime
                XmlFile.WriteStartElement("ProcCmdCount")
                XmlFile.WriteValue(100)
                XmlFile.WriteEndElement() 'ProcCmdCount
                XmlFile.WriteStartElement("MaxThread")
                XmlFile.WriteValue(50)
                XmlFile.WriteEndElement() 'MaxThread
                XmlFile.WriteStartElement("DBMinPool")
                XmlFile.WriteValue(0)
                XmlFile.WriteEndElement() 'DBMinPool
                XmlFile.WriteStartElement("DBMaxPool")
                XmlFile.WriteValue(200)
                XmlFile.WriteEndElement() 'DBMaxPool
                XmlFile.WriteStartElement("DBPoolLiveTime")
                XmlFile.WriteValue(0)
                XmlFile.WriteEndElement() 'DBPoolLiveTime
                XmlFile.WriteEndElement()   'System
                '---------------------------------------------------------------------------------------------------------------------------
                XmlFile.WriteStartElement("SODB")   'SODB
                XmlFile.WriteStartElement("DB")     'DB
                XmlFile.WriteAttributeString("CompCode", 7)
                XmlFile.WriteAttributeString("SID", "RDKNET")
                XmlFile.WriteAttributeString("USER", "CCTV")
                XmlFile.WriteAttributeString("PASSWORD", "CCTV")
                XmlFile.WriteValue(0)
                XmlFile.WriteEndElement()            'DB
                XmlFile.WriteStartElement("DB")     'DB
                XmlFile.WriteAttributeString("CompCode", 11)
                XmlFile.WriteAttributeString("SID", "RDKNET")
                XmlFile.WriteAttributeString("USER", "TCBSH")
                XmlFile.WriteAttributeString("PASSWORD", "TCBSH")
                XmlFile.WriteValue(0)
                XmlFile.WriteEndElement()            'DB
                XmlFile.WriteEndElement() 'SODB
                '------------------------------------------------------------------------------------------------------------------------------
                XmlFile.WriteStartElement("NSTV")   'NSTV
                XmlFile.WriteStartElement("SMS")     'SMS
                XmlFile.WriteAttributeString("OPE_ID", 1)
                XmlFile.WriteValue(1)
                XmlFile.WriteEndElement()            'SMS
                XmlFile.WriteStartElement("CheckTime")     'CheckTime
                XmlFile.WriteValue(900)
                XmlFile.WriteEndElement()       'CheckTime
                XmlFile.WriteStartElement("Server") 'Server
                XmlFile.WriteValue("127.0.0.1")
                XmlFile.WriteEndElement()     'Server
                XmlFile.WriteStartElement("Port") 'Port
                XmlFile.WriteValue("7364")
                XmlFile.WriteEndElement()     'Port
                XmlFile.WriteStartElement("RootKey") 'RootKey
                XmlFile.WriteValue("ABCDEFGHABCDEFGH")
                XmlFile.WriteEndElement()     'RootKey
                XmlFile.WriteStartElement("CryptVer") 'CryptVer
                XmlFile.WriteValue(1)
                XmlFile.WriteEndElement()     'CryptVer
                XmlFile.WriteStartElement("ProtoVer") 'ProtoVer
                XmlFile.WriteValue(1)
                XmlFile.WriteEndElement()     'ProtoVer
                XmlFile.WriteStartElement("KeyType") 'KeyType
                XmlFile.WriteValue(2)
                XmlFile.WriteEndElement()     'KeyType
                XmlFile.WriteStartElement("SendDelayTime") 'SendDelayTime
                XmlFile.WriteValue(0.1)
                XmlFile.WriteEndElement()     'SendDelayTime
                XmlFile.WriteEndElement() 'NSTV
                '------------------------------------------------------------------------------------------------------------------------------
                XmlFile.WriteStartElement("ErrorList") 'ErrorList
                XmlFile.WriteStartElement("Error")      'Error
                XmlFile.WriteAttributeString("Value", "0x0000")
                XmlFile.WriteValue("TFCA_OK")
                XmlFile.WriteEndElement()       'Error
                XmlFile.WriteStartElement("Error")      'Error
                XmlFile.WriteAttributeString("Value", "0x0002")
                XmlFile.WriteValue("CARD_NOT_EXISTS")
                XmlFile.WriteEndElement()       'Error
                XmlFile.WriteEndElement() 'ErrorList
                '------------------------------------------------------------------------------------------------------------------------------
                XmlFile.WriteStartElement("LowCmd") 'LowCmd
                XmlFile.WriteStartElement("MsgId")    'MsgId
                XmlFile.WriteAttributeString("CodeNo", "1")
                XmlFile.WriteAttributeString("Value", "0x0001")
                XmlFile.WriteValue("SMS_CA_CREATE_SESSION_REQUEST")
                XmlFile.WriteEndElement()                 'MsgId
                XmlFile.WriteStartElement("MsgId")    'MsgId
                XmlFile.WriteAttributeString("CodeNo", "2")
                XmlFile.WriteAttributeString("Value", "0x0201")
                XmlFile.WriteValue("SMS_CA_OPEN_ACCOUNT_REQUEST")
                XmlFile.WriteEndElement()                 'MsgId
                XmlFile.WriteStartElement("MsgId")    'MsgId
                XmlFile.WriteAttributeString("CodeNo", "3")
                XmlFile.WriteAttributeString("Value", "0x0202")
                XmlFile.WriteValue("SMS_CA_STOP_ACCOUNT_REQUEST")
                XmlFile.WriteEndElement()                 'MsgId
                XmlFile.WriteStartElement("MsgId")    'MsgId
                XmlFile.WriteAttributeString("CodeNo", "4")
                XmlFile.WriteAttributeString("Value", "0x0203")
                XmlFile.WriteValue("SMS_CA_SET_LOCK_REQUEST")
                XmlFile.WriteEndElement()                 'MsgId
                XmlFile.WriteStartElement("MsgId")    'MsgId
                XmlFile.WriteAttributeString("CodeNo", "5")
                XmlFile.WriteAttributeString("Value", "0x0204")
                XmlFile.WriteValue("SMS_CA_SET_UNLOCK_REQUEST")
                XmlFile.WriteEndElement()                 'MsgId
                XmlFile.WriteStartElement("MsgId")    'MsgId
                XmlFile.WriteAttributeString("CodeNo", "6")
                XmlFile.WriteAttributeString("Value", "0x0206")
                XmlFile.WriteValue("SMS_CA_REPAIR_REQUEST")
                XmlFile.WriteEndElement()                 'MsgId
                XmlFile.WriteStartElement("MsgId")    'MsgId
                XmlFile.WriteAttributeString("CodeNo", "7")
                XmlFile.WriteAttributeString("Value", "0x020D")
                XmlFile.WriteValue("SMS_CA_ENTITLE_REQUEST")
                XmlFile.WriteEndElement()                 'MsgId
                XmlFile.WriteStartElement("MsgId")    'MsgId
                XmlFile.WriteAttributeString("CodeNo", "8")
                XmlFile.WriteAttributeString("Value", "0x0402")
                XmlFile.WriteValue("SMS_CA_SET_CHILD_REQUEST")
                XmlFile.WriteEndElement()                 'MsgId
                XmlFile.WriteStartElement("MsgId")    'MsgId
                XmlFile.WriteAttributeString("CodeNo", "9")
                XmlFile.WriteAttributeString("Value", "0x0403")
                XmlFile.WriteValue("SMS_CA_CANCEL_CHILD_REQUEST")
                XmlFile.WriteEndElement()                 'MsgId
                XmlFile.WriteStartElement("MsgId")    'MsgId
                XmlFile.WriteAttributeString("CodeNo", "10")
                XmlFile.WriteAttributeString("Value", "0x0207")
                XmlFile.WriteValue("SMS_CA_RESETCARDPIN_REQUEST")
                XmlFile.WriteEndElement()                 'MsgId
                XmlFile.WriteStartElement("MsgId")    'MsgId
                XmlFile.WriteAttributeString("CodeNo", "11")
                XmlFile.WriteAttributeString("Value", "0x020C")
                XmlFile.WriteValue("SMS_CA_SET_CHARACTER_REQUEST")
                XmlFile.WriteEndElement()                 'MsgId
                XmlFile.WriteEndElement()         'LowCmd
                '--------------------------------------------------------------------------------------------------------------------------------
                XmlFile.WriteStartElement("HightCmd") 'HightCmd
                XmlFile.WriteStartElement("Cmd")        'Cmd
                XmlFile.WriteAttributeString("CodeNo", "A1")
                XmlFile.WriteAttributeString("RunLowCmd", "2,5,3")
                XmlFile.WriteValue("測試用1")
                XmlFile.WriteEndElement()                     'Cmd
                XmlFile.WriteStartElement("Cmd")        'Cmd
                XmlFile.WriteAttributeString("CodeNo", "A2")
                XmlFile.WriteAttributeString("RunCmd", "5,3,8")
                XmlFile.WriteValue("測試用2")
                XmlFile.WriteEndElement()                     'Cmd
                XmlFile.WriteEndElement()             'HightCmd
                '--------------------------------------------------------------------------------------------------------------------------------
                XmlFile.WriteEndElement() 'CableSoft
                XmlFile.WriteEndDocument()
                XmlFile.Flush()

                XmlFile.Close()

                Using stm As StreamWriter = New StreamWriter(XmlFileName, False, Encoding.UTF8)
                    If IsEncryFile Then
                        ' stm.Write(CableSoft.CAS.Encry.Encryption.Encrypt3DES( Convert.ToBase64String(mem.ToArray), Key))
                        Dim strWrite As String = CableSoft.NSTV.Encry.Encryption.Encrypt3DES(Convert.ToBase64String(Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(mem.ToArray))), Key)
                        'Dim strWrite As String = CableSoft.CAS.CryptUtil.CryptUtil.Encrypt(Encoding.UTF8.GetString(mem.ToArray))
                        stm.WriteLine(strWrite)
                        stm.Flush()
                        stm.Close()
                    Else
                        stm.WriteLine(Encoding.UTF8.GetString(mem.ToArray))
                    End If
                End Using


            End Using
        End Using
        Return True
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If mtxWait IsNot Nothing Then
                    mtxWait.Dispose()
                    mtxWait = Nothing
                End If
                If xmlDoc IsNot Nothing Then
                    xmlDoc = Nothing
                End If
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
