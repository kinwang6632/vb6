Imports System.IO
Imports System.Text
Imports DevExpress.XtraEditors
Imports DevExpress.XtraTreeList
Imports DevExpress.XtraTreeList.Nodes
Imports CableSoft.CardLess

Public Class frmSetting
    Private Const xmlFileName As String = "CardLess.Set"
    Private XmlFile As Xml.XmlTextWriter = Nothing
    Private xmlDoc As System.Xml.XmlDocument = Nothing
    Private xmlIO As XMLFileIO.XmlFileIO = Nothing
    Private Const IsEncryFile As Boolean = True
    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        Try
            If WriteData() Then

                XtraMessageBox.Show("儲存成功!", "訊息", MessageBoxButtons.OK)
            Else
                XtraMessageBox.Show("儲存失敗!", "訊息", MessageBoxButtons.OK)
            End If

        Catch ex As Exception

            XtraMessageBox.Show(ex.ToString, "訊息", MessageBoxButtons.OK)
        End Try


    End Sub
    Private Function WriteData() As Boolean
        Try
            Using mem As New MemoryStream
                Using XmlFile As New Xml.XmlTextWriter(mem, New UTF8Encoding(False))
                    XmlFile.Formatting = Xml.Formatting.Indented
                    XmlFile.Indentation = 4
                    XmlFile.WriteStartDocument()
                    XmlFile.WriteStartElement("CableSoft")  'CableSoft
                    XmlFile.WriteStartElement("System")     'System
                    XmlFile.WriteStartElement("ReadTime")
                    XmlFile.WriteValue(spinFreq.Value)
                    XmlFile.WriteEndElement() 'ReadTime
                    XmlFile.WriteStartElement("ProcCmdCount")
                    XmlFile.WriteValue(spinProc.Value)
                    XmlFile.WriteEndElement() 'ProcCmdCount
                    XmlFile.WriteStartElement("MaxThread")
                    XmlFile.WriteValue(spinMaxThreads.Value)
                    XmlFile.WriteEndElement() 'MaxThread
                    XmlFile.WriteStartElement("DBMinPool")
                    XmlFile.WriteValue(spinMinP.Value)
                    XmlFile.WriteEndElement() 'DBMinPool
                    XmlFile.WriteStartElement("DBMaxPool")
                    XmlFile.WriteValue(spinMaxP.Value)
                    XmlFile.WriteEndElement() 'DBMaxPool
                    XmlFile.WriteStartElement("DBPoolLiveTime")
                    XmlFile.WriteValue(spinLifeP.Value)
                    XmlFile.WriteEndElement() 'DBPoolLiveTime
                    XmlFile.WriteEndElement()   'System
                    '---------------------------------------------------------------------------------------------------------------------------
                    XmlFile.WriteStartElement("SODB")   'SODB
                    For i As Int32 = 0 To TreLstDB.Nodes.Count - 1
                        Dim Node As TreeListNode = TreLstDB.Nodes(i)
                        XmlFile.WriteStartElement("DB")     'DB
                        XmlFile.WriteAttributeString("Value", Node.Item("SelectID"))
                        XmlFile.WriteAttributeString("CompCode", Node.Item("CompID"))
                        XmlFile.WriteAttributeString("SID", Node.Item("DBSid"))
                        XmlFile.WriteAttributeString("USER", Node.Item("DBUser"))
                        XmlFile.WriteAttributeString("PASSWORD", Node.Item("DBPassword"))
                        XmlFile.WriteValue(Node.Item("CompName"))
                        XmlFile.WriteEndElement()            'DB
                    Next
                    XmlFile.WriteEndElement()                 'SODB
                    '------------------------------------------------------------------------------------------------------------------------------
                    XmlFile.WriteStartElement("Nagra")   'Nagra
                    For Each Node As TreeListNode In TreLstSourceId.Nodes
                        XmlFile.WriteStartElement("Server")     'Server
                        XmlFile.WriteAttributeString("DestId", Node.Item(colDestId))
                        XmlFile.WriteAttributeString("MoppId", Node.Item(colMoppId))
                        XmlFile.WriteValue(Node.Item(colSourceId))
                        XmlFile.WriteEndElement()            'Server
                    Next
                    XmlFile.WriteStartElement("IPAddress")    'IPAddress
                    If String.IsNullOrEmpty(txtNagraIPAddress.Text) Then
                        XmlFile.WriteValue("127.0.0.1")
                    Else
                        XmlFile.WriteValue(txtNagraIPAddress.Text)
                    End If
                    XmlFile.WriteEndElement()                                             'IPAddress
                    XmlFile.WriteStartElement("Port")        'Port                    
                    XmlFile.WriteValue(spinNagraPort.Value)
                    XmlFile.WriteEndElement()                   'Port
                  
                    XmlFile.WriteStartElement("SendTimeOut")    'SendTimeOut
                    XmlFile.WriteValue(spinSendTimeout.Value)
                    XmlFile.WriteEndElement()                                     'SendTimeOut
                    XmlFile.WriteStartElement("ReceiveTimeout")  'ReceiveTimeout
                    XmlFile.WriteValue(spinReceiveTimeout.Value)
                    XmlFile.WriteEndElement()                                         'ReceiveTimeout                   
                    XmlFile.WriteStartElement("CheckTime")     'CheckTime
                    XmlFile.WriteValue(spinChkStatus.Value)
                    XmlFile.WriteEndElement()       'CheckTime

                    XmlFile.WriteStartElement("DisconnectRetryTime") 'DisconnectRetryTime
                    XmlFile.WriteValue(spinDisconnectRetry.Value)
                    XmlFile.WriteEndElement()     'DisconnectRetryTime                   
                    XmlFile.WriteStartElement("UTC")     'UTC
                    If chkUseUTC.Checked Then
                        XmlFile.WriteValue("1")
                    Else
                        XmlFile.WriteValue("0")
                    End If
                    XmlFile.WriteEndElement()
                    XmlFile.WriteStartElement("Encoding")       'Encoding
                    XmlFile.WriteValue(cmbEncoding.Text)
                    XmlFile.WriteEndElement()                                 'Encoding

                    XmlFile.WriteStartElement("RetryProcCount")     'RetryProcCount
                    XmlFile.WriteValue(spinRetryCount.Value)
                    XmlFile.WriteEndElement()                           'RetryProcCount

                    XmlFile.WriteStartElement("ProcKind")                       'ProcKind
                    XmlFile.WriteValue(cobProcKind.SelectedIndex)
                    XmlFile.WriteEndElement()                                                 'ProcKind
                    XmlFile.WriteEndElement() 'Nagra
                    '------------------------------------------------------------------------------------------------------------------------------
                    XmlFile.WriteStartElement("ErrorList") 'ErrorList
                    For Each Node As TreeListNode In TreLstErrCode.Nodes
                        XmlFile.WriteStartElement("Error")      'Error                    
                        XmlFile.WriteAttributeString("ErrorCode", Node.Item("ErrorCode"))
                        XmlFile.WriteValue(Node.Item("ErrorName"))
                        XmlFile.WriteEndElement()       'Error
                    Next
                    XmlFile.WriteEndElement() 'ErrorList
                    '------------------------------------------------------------------------------------------------------------------------------
                    'XmlFile.WriteStartElement("LowCmd") 'LowCmd
                    'For Each Node As TreeListNode In TreLstLowCmd.Nodes
                    '    XmlFile.WriteStartElement("MsgId")    'MsgId
                    '    XmlFile.WriteAttributeString("CodeNo", Node.Item("Code"))
                    '    XmlFile.WriteAttributeString("Value", Node.Item("Value"))
                    '    XmlFile.WriteAttributeString("DefaultValue", Node.Item("DefaultValue"))
                    '    XmlFile.WriteAttributeString("RealField", Node.Item("RealField"))
                    '    XmlFile.WriteValue(Node.Item("Description"))
                    '    XmlFile.WriteEndElement()                 'MsgId
                    'Next
                    'XmlFile.WriteEndElement()         'LowCmd
                    '--------------------------------------------------------------------------------------------------------------------------------
                    'XmlFile.WriteStartElement("HightCmd") 'HightCmd
                    'For Each node As TreeListNode In TreLstHightCmd.Nodes
                    '    XmlFile.WriteStartElement("Cmd")        'Cmd
                    '    XmlFile.WriteAttributeString("CodeNo", node.Item("CodeNo"))
                    '    XmlFile.WriteAttributeString("RunLowCmd", node.Item("RunLowCmd"))
                    '    XmlFile.WriteAttributeString("RunLowCmd2", node.Item("RunLowCmd2"))
                    '    XmlFile.WriteValue(node.Item("HightDesc"))
                    '    XmlFile.WriteEndElement()                     'Cmd
                    'Next
                    'XmlFile.WriteEndElement()             'HightCmd
                    '--------------------------------------------------------------------------------------------------------------------------------
                    XmlFile.WriteEndElement() 'CableSoft
                    XmlFile.WriteEndDocument()
                    XmlFile.Flush()
                    XmlFile.Close()
                    xmlIO.WriteXmlFile(Application.StartupPath & "\" & xmlFileName, mem, IsEncryFile)
                End Using
            End Using
            Return True
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Private Function ReadData() As Boolean
        Try
            Dim tbsystem As DataTable = xmlIO.ReadSystem(IsEncryFile)
            If tbsystem.Rows.Count > 0 Then
                With tbsystem.Rows(0)
                    spinFreq.Value = Decimal.Parse(.Item("ReadTime"))
                    spinProc.Value = Decimal.Parse(.Item("ProcCmdCount"))
                    spinMaxThreads.Value = Decimal.Parse(.Item("MaxThread"))
                    spinMinP.Value = Decimal.Parse(.Item("DBMinPool"))
                    spinMaxP.Value = Decimal.Parse(.Item("DBMaxPool"))
                    spinLifeP.Value = Decimal.Parse(.Item("DBPoolLiveTime"))
                End With
            End If

            ReadSO()
            ReadErrorList()
            ReadServer()
            ReadNstv()
            ReadLowCmd()
            ReadHightCmd()
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Private Function ReadHightCmd() As Boolean
        Try
            Dim tbHighCMD As DataTable = xmlIO.ReadHightCmd(IsEncryFile)
            For i As Int32 = 0 To tbHighCMD.Rows.Count - 1
                DataToTreList(TreLstHightCmd, tbHighCMD.Rows(i).ItemArray.Take(6).ToArray, i)
            Next
        Catch ex As Exception
            Throw
        End Try
        'Dim tbHighCMD As New DataTable("HighCMD")
        'tbHighCMD.Columns.Add(New DataColumn("CodeNo", GetType(String)))
        'tbHighCMD.Columns.Add(New DataColumn("RunLowCmd", GetType(String)))
        'tbHighCMD.Columns.Add(New DataColumn("MsgName", GetType(String)))
        'Try
        '    Dim CodeNo As List(Of Object) = xmlIO.ReadXmlNodeAttributes("CableSoft/HightCmd", "Cmd", "CodeNo", IsEncryFile)
        '    Dim Cmd As List(Of Object) = xmlIO.ReadXmlNodeAttributes("CableSoft/HightCmd", "Cmd", "RunLowCmd", IsEncryFile)
        '    Dim Descript As List(Of Object) = xmlIO.ReadXmlNodeValues("CableSoft/HightCmd", "Cmd", IsEncryFile)

        '    For i As Int32 = 0 To CodeNo.Count - 1
        '        Dim data(2) As Object
        '        data(0) = CodeNo.Item(i)
        '        data(1) = Descript.Item(i)
        '        data(2) = Cmd.Item(i)
        '        DataToTreList(TreLstHightCmd, data, i)
        '        Dim rwNew As DataRow = tbHighCMD.NewRow
        '        rwNew("CodeNo") = data(0)
        '        rwNew("RunLowCmd") = data(1)
        '        rwNew("MsgName") = data(2)
        '        tbHighCMD.Rows.Add(rwNew)
        '        tbHighCMD.AcceptChanges()
        '    Next
        'Catch ex As Exception
        '    Throw
        'End Try
        Return True
    End Function
    Private Function ReadLowCmd() As Boolean
        Try
            Dim tbLowCmd As DataTable = xmlIO.ReadLowCmd(IsEncryFile)
            For i As Int32 = 0 To tbLowCmd.Rows.Count - 1
                DataToTreList(TreLstLowCmd, tbLowCmd.Rows(i).ItemArray.Take(6).ToArray, i)
            Next
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Private Function ReadServer() As Boolean
        Try
            Dim tbServer As DataTable = xmlIO.ReadServer(IsEncryFile, "Server")
            For i As Int32 = 0 To tbServer.Rows.Count - 1
                DataToTreList(TreLstSourceId, tbServer.Rows(i).ItemArray, i)
            Next
        Catch ex As Exception
            Throw
        End Try


        'Dim tbSMS As New DataTable("SMS")
        'tbSMS.Columns.Add(New DataColumn("OPE_ID", GetType(String)))
        'tbSMS.Columns.Add(New DataColumn("SMS", GetType(String)))
        'Try
        '    Dim OPE_ID As List(Of Object) = xmlIO.ReadXmlNodeAttributes("CableSoft/NSTV", "SMS", "OPE_ID", IsEncryFile)
        '    Dim SMS As List(Of Object) = xmlIO.ReadXmlNodeValues("CableSoft/NSTV", "SMS", IsEncryFile)
        '    For i As Int32 = 0 To OPE_ID.Count - 1
        '        Dim data(1) As Object
        '        If String.IsNullOrEmpty(OPE_ID.Item(i).ToString) Then
        '            data(0) = "1"
        '        Else
        '            data(0) = OPE_ID.Item(i)
        '        End If

        '        If String.IsNullOrEmpty(SMS.Item(i).ToString) Then
        '            data(1) = "1"
        '        Else
        '            data(1) = SMS.Item(i)
        '        End If
        '        Dim rwNew As DataRow = tbSMS.NewRow
        '        rwNew.Item("OPE_ID") = data(0)
        '        rwNew.Item("SMS") = data(1)
        '        tbSMS.Rows.Add(rwNew)
        '        tbSMS.AcceptChanges()
        '        DataToTreList(TreLstSourceId, data, i)
        '    Next
        'Catch ex As Exception
        '    Throw
        'End Try
        Return True
    End Function
    Private Function ReadNstv() As Boolean
        Try
            Dim tbNagra As DataTable = xmlIO.ReadNagra(IsEncryFile)
            If tbNagra.Rows.Count = 1 Then
                With tbNagra.Rows(0)
                    spinChkStatus.Value = Int32.Parse(.Item("CheckTime"))
                    txtNagraIPAddress.Text = tbNagra.Rows(0).Item("IPAddress").ToString
                    spinNagraPort.Value = Decimal.Parse(tbNagra.Rows(0).Item("Port").ToString)

                    spinSendTimeout.Value = Int32.Parse(tbNagra.Rows(0).Item("SendTimeOut"))
                    spinReceiveTimeout.Value = Int32.Parse(tbNagra.Rows(0).Item("ReceiveTimeout"))
                    spinRetryCount.Value = Int32.Parse(tbNagra.Rows(0).Item("RetryProcCount"))

                    If Int32.Parse(tbNagra.Rows(0).Item("UTC")) = 1 Then
                        chkUseUTC.Checked = True
                    End If
                    'chkCanRecord.EditValue = Int32.Parse(tbNSTV.Rows(0).Item("Record"))
                    'chkUseUTC.EditValue = Int32.Parse(tbNSTV.Rows(0).Item("UTC"))
                    cmbEncoding.Text = tbNagra.Rows(0).Item("Encoding")
                    spinDisconnectRetry.Value = Int32.Parse(tbNagra.Rows(0).Item("DisconnectRetryTime"))

                    cobProcKind.SelectedIndex = Integer.Parse("0" & tbNagra.Rows(0).Item("ProcKind"))

                End With
            End If
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Private Function ReadErrorList() As Boolean
        Try
            Dim tbError As DataTable = xmlIO.ReadErrorList(IsEncryFile)
            For i As Int32 = 0 To tbError.Rows.Count - 1
                DataToTreList(TreLstErrCode, tbError.Rows(i).ItemArray, i)
            Next
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Private Function ReadSO() As Boolean
        Try
            Dim tbSO As DataTable = xmlIO.ReadSO(IsEncryFile)
            For i As Int32 = 0 To tbSO.Rows.Count - 1
                DataToTreList(TreLstDB, tbSO.Rows(i).ItemArray.Take(6).ToArray, i)
            Next
        Catch ex As Exception
            Throw
        End Try

        Return True
    End Function
    Private Function DataToTreList(ByVal aTree As TreeList,
                                   ByVal Data() As Object, ByVal Index As Int32) As Boolean
        Try
            aTree.AppendNode(Data, Index)
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Private Function CreateXMLFile() As Boolean
        'Dim stm As New FileStream(Application.StartupPath & "\" & xmlFileName, FileMode.Create, FileAccess.ReadWrite)
        'Dim stm2 As System.IO.TextWriter = New StreamWriter(stm)
        'XmlFile = New Xml.XmlTextWriter(stm2)
        'XmlFile.Formatting = Xml.Formatting.Indented
        'XmlFile.Indentation = 4
        'XmlFile.WriteStartDocument()

        'Return True
        Using mem As New MemoryStream
            Using XmlFile As New Xml.XmlTextWriter(mem, Encoding.UTF8)
                XmlFile.Formatting = Xml.Formatting.Indented
                XmlFile.Indentation = 4
                XmlFile.WriteStartDocument()
                XmlFile.WriteStartElement("CableSoft")
                XmlFile.WriteStartElement("ReadTime")
                XmlFile.WriteAttributeString("TEST", 2)
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
                XmlFile.WriteEndElement() 'CableSoft
                XmlFile.Flush()

                XmlFile.Close()

                Debug.Print(Convert.ToBase64String(mem.ToArray))
            End Using
        End Using
        '     using (var ms = new MemoryStream())
        'using (var x = new XmlTextWriter(ms, new UTF8Encoding(false)) 
        '  { Formatting = Formatting.Indented })
        '{
        '  // ...
        '  return Encoding.UTF8.GetString(ms.ToArray());
        '}


        Return True
    End Function

    Private Sub frmSetting_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        xmlIO.Dispose()
    End Sub

    Private Sub frmSetting_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Try
            cmbEncoding.Properties.Items.Add("UTF-8")
            cmbEncoding.Properties.Items.Add("UTF-32")
            cmbEncoding.Properties.Items.Add("Unicode")
            cmbEncoding.Properties.Items.Add("UTF-7")
            cmbEncoding.Properties.Items.Add("ASCII")
            cmbEncoding.Properties.Items.Add("Big5")
            cmbEncoding.SelectedIndex = 0
            'cmbEncoding.Properties.Items.AddRange(obj)
            xmlIO = New CableSoft.CardLess.XMLFileIO.XmlFileIO(Application.StartupPath & "\" & xmlFileName)
            If Not File.Exists(Application.StartupPath & "\" & xmlFileName) Then
                XtraMessageBox.Show("設定檔不存在，請儲存檔案", "訊息",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
                'xmlIO.CreateDefaultXMLFile(IsEncryFile)
            End If
            ReadData()
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString)
        End Try







        'ReadXmlNodeValue("CableSoft", "ReadTime")
    End Sub

    'Function ReadXmlNodeValue(ByVal ParentNode As String, ByVal NodeName As String) As Object
    '    Try
    '        If xmlDoc Is Nothing Then
    '            xmlDoc = New Xml.XmlDocument()
    '            Using stm As Stream = New FileStream(Application.StartupPath & "\" & xmlFileName, FileMode.Open, FileAccess.ReadWrite)
    '                Using rder As New System.IO.StreamReader(stm)

    '                    xmlDoc.LoadXml(rder.ReadToEnd())
    '                End Using

    '            End Using

    '        End If

    '    Catch ex As Exception
    '        MessageBox.Show(ex.ToString, "訊息", MessageBoxButtons.OK)
    '    End Try
    'End Function
    'Function readXmlnode(ByVal NodePath As String) As String





    '    xmlDoc.Load(Application.StartupPath & "\" & xmlFileName)

    '    'xmlDoc.ParentNode
    '    Dim mXmlNode As System.Xml.XmlElement

    '    'mXmlNode = doc.SelectNodes(NodePath)

    '    'Return mXmlNode.InnerText
    '    Return "empty"


    'End Function
    Private Sub btnAddConn_Click(sender As System.Object, e As System.EventArgs) Handles btnAddConn.Click
        Try
            TreLstDB.AppendNode(Nothing, -1, -1, -1, 0).StateImageIndex = 0

        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'Msg.Show(ex.ToString(), "訊息", MsgBtn.OK, MsgIco.Error)
            'Bs.Write2MdbLog(Nothing, Nothing, Nothing, ex)
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Application.Exit()
    End Sub

    Private Sub btnAddSourceId_Click(sender As System.Object, e As System.EventArgs) Handles btnAddSourceId.Click
        Try
            TreLstSourceId.AppendNode(Nothing, -1, -1, -1, 0).StateImageIndex = 2
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub



    Private Sub btnAddError_Click(sender As System.Object, e As System.EventArgs) Handles btnAddError.Click
        Try
            TreLstErrCode.AppendNode(Nothing, -1, -1, -1, 0).StateImageIndex = 1
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnAddLowCmd_Click(sender As System.Object, e As System.EventArgs) Handles btnAddLowCmd.Click
        Try
            TreLstLowCmd.AppendNode(Nothing, -1, -1, -1, 0).StateImageIndex = 1
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnAddHightCmd_Click(sender As System.Object, e As System.EventArgs) Handles btnAddHightCmd.Click
        Try
            TreLstHightCmd.AppendNode(Nothing, -1, -1, -1, 0).StateImageIndex = 1
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub btnRemoveConn_Click(sender As System.Object, e As System.EventArgs) Handles btnRemoveConn.Click
        Try
            TreLstDB.DeleteNode(TreLstDB.FocusedNode)
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub btnbtnRemoveSourceId_Click(sender As System.Object, e As System.EventArgs) Handles btnbtnRemoveSourceId.Click
        Try
            TreLstSourceId.DeleteNode(TreLstSourceId.FocusedNode)
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString)
        End Try

    End Sub


    Private Sub btnDelLowCmd_Click(sender As System.Object, e As System.EventArgs) Handles btnDelLowCmd.Click
        Try
            TreLstLowCmd.DeleteNode(TreLstLowCmd.FocusedNode)
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub btnDelHighCmd_Click(sender As System.Object, e As System.EventArgs) Handles btnDelHighCmd.Click
        Try
            TreLstHightCmd.DeleteNode(TreLstHightCmd.FocusedNode)
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub btnRemoveAuth_Click(sender As System.Object, e As System.EventArgs) Handles btnRemoveAuth.Click
        Try
            TreLstErrCode.DeleteNode(TreLstErrCode.FocusedNode)
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub btnTestConn_Click(sender As System.Object, e As System.EventArgs) Handles btnTestConn.Click
        Dim bduConnect As New OracleClient.OracleConnectionStringBuilder()
        bduConnect.Password = TreLstDB.FocusedNode.Item("DBPassword")
        bduConnect.UserID = TreLstDB.FocusedNode.Item("DBUser")
        bduConnect.DataSource = TreLstDB.FocusedNode.Item("DBSid")
        bduConnect.PersistSecurityInfo = True
        Using cn As New OracleClient.OracleConnection(bduConnect.ConnectionString)
            Try
                cn.Open()
                cn.Close()
                XtraMessageBox.Show("連接成功！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                XtraMessageBox.Show("連接失敗！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using

    End Sub


End Class