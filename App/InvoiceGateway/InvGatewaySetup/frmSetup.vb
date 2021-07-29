Imports DevExpress.XtraEditors
Imports System.IO
Imports System.Text
Imports DevExpress.XtraTreeList.Nodes
Imports DevExpress.XtraTreeList
Imports System.Xml

Public Class frmSetup
    Private Const xmlFileName As String = "InvoiceSetup.Set"
    Private GatewayIO As CableSoft.Invoice.Gateway.SetupFileIO.InvoiceSetupIO = Nothing
    Private tbSODB As DataTable = Nothing
    Private Const IsEncryFile As Boolean = True

    Private Sub btnAddConn_Click(sender As System.Object, e As System.EventArgs) Handles btnAddConn.Click
        Try
            Dim nod As TreeListNode = TreLstDB.AppendNode(Nothing, -1, -1, -1, 0)
            nod.StateImageIndex = 0
            nod.Item(colMinDBPool) = "0"
            nod.Item(colMaxDBPool) = "10"
            nod.Item(colDBPoolLifetime) = "0"
            nod.Item(colRunCmdTimer) = "1"
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
    Private Function ReadSystem() As Boolean
        Try
            Using tbSystem As DataTable = GatewayIO.ReadSystem(IsEncryFile)
                spinMaxThreads.Value = Convert.ToDecimal(tbSystem.Rows(0).Item("MaxThreads"))
                txtLoginUsreId.EditValue = tbSystem.Rows(0).Item("LoginUserId")
                txtLoginUserName.EditValue = tbSystem.Rows(0).Item("LoginUserName")
                txtDestroyReason.EditValue = tbSystem.Rows(0).Item("DestroyReason")
                chkLogSQL.Checked = Integer.Parse(tbSystem.Rows(0).Item("IsLogSQL")) = 1
                spinReserveLog.Value = Integer.Parse(tbSystem.Rows(0).Item("LogReserveTime"))
                spinReserveLog.Enabled = False
                If chkLogSQL.Checked Then
                    spinReserveLog.Enabled = True
                End If
            End Using
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Private Function ReadRunCommandType() As Boolean
        Try
            Using tbRunCommand As DataTable = GatewayIO.ReadRunCommand(IsEncryFile)
                For i As Integer = 0 To tbRunCommand.Rows.Count - 1
                    DataToTreList(TreLstSORunCmd, tbRunCommand.Rows(i).ItemArray, i)
                Next
            End Using
         
        Catch ex As Exception
            Throw
        End Try
        Return True

    End Function
    Private Function ReadCommand() As Boolean
        Try
            Using tbCommand As DataTable = GatewayIO.ReadCommand(IsEncryFile)
                For i As Int32 = 0 To tbCommand.Rows.Count - 1
                    DataToTreList(TreLstCmd, tbCommand.Rows(i).ItemArray, i)
                Next
            End Using
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Private Function ReadSO() As Boolean
        Try
            Using tbSO As DataTable = GatewayIO.ReadSO(IsEncryFile)
                For i As Int32 = 0 To tbSO.Rows.Count - 1
                    DataToTreList(TreLstDB, tbSO.Rows(i).ItemArray, i)
                Next
            End Using
            
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
    Private Function ReadData() As Boolean
        rCboCmdName.Items.Add(String.Format("{0} ( {1} )",
                                            "開立發票", [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                                                  CableSoft.B07.InvType.InvTypeEnum.InvFileType.C0401)))
        rCboCmdName.Items.Add(String.Format("{0} ( {1} )",
                                           "作廢發票", [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                                                 CableSoft.B07.InvType.InvTypeEnum.InvFileType.C0501)))
        rCboCmdName.Items.Add(String.Format("{0} ( {1} )",
                                           "折讓發票", [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                                                 CableSoft.B07.InvType.InvTypeEnum.InvFileType.D0401)))
        rCboCmdName.Items.Add(String.Format("{0} ( {1} )",
                                           "折讓作廢發票", [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                                                 CableSoft.B07.InvType.InvTypeEnum.InvFileType.D0501)))
        rCboCmdName.Items.Add(String.Format("{0} ( {1} )",
                                           "註銷發票", [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                                                 CableSoft.B07.InvType.InvTypeEnum.InvFileType.C0701)))
        rCboCmdName.Items.Add(String.Format("{0} ( {1} )",
                                           "註銷發票與重新上傳", [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                                                 CableSoft.B07.InvType.InvTypeEnum.InvFileType.X0401)))
        rCboCmdName.Items.Add(String.Format("{0} ( {1} )",
                                           "電子發票通知", [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                                                 CableSoft.B07.InvType.InvTypeEnum.InvFileType.B08)))

        Try
            ReadSO()
            ReadCommand()
            ReadRunCommandType()
            ReadSystem()
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK)
        End Try
        Return True

    End Function
    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        WriteData()
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
                    XmlFile.WriteStartElement("MaxThread")  'MaxThread
                    If spinMaxThreads.Value = 0 Then
                        spinMaxThreads.Value = 30
                    End If
                    XmlFile.WriteValue(Convert.ToInt32(spinMaxThreads.Value))
                    XmlFile.WriteEndElement()                           'MaxThread
                    XmlFile.WriteStartElement("LoginUserId")  'LoginUserId
                    If String.IsNullOrEmpty(txtLoginUsreId.Text) Then
                        XmlFile.WriteValue("Gateway")
                    Else
                        XmlFile.WriteValue(txtLoginUsreId.Text)
                    End If
                    XmlFile.WriteEndElement()                           'LoginUserId
                    XmlFile.WriteStartElement("LoginUserName")  'LoginUserName
                    If String.IsNullOrEmpty(txtLoginUserName.Text) Then
                        XmlFile.WriteValue("Gateway")
                    Else
                        XmlFile.WriteValue(txtLoginUserName.Text)
                    End If
                    XmlFile.WriteEndElement()                           'LoginUserName

                    XmlFile.WriteStartElement("DestroyReason")  'DestroyReason
                    XmlFile.WriteValue(txtDestroyReason.Text)
                    XmlFile.WriteEndElement()                           'DestroyReason

                    XmlFile.WriteStartElement("IsLogSQL")  'IsLogSQL
                    If chkLogSQL.Checked Then
                        XmlFile.WriteValue(1)
                    Else
                        XmlFile.WriteValue(0)
                    End If
                    XmlFile.WriteEndElement()                           'IsLogSQL

                    XmlFile.WriteStartElement("LogReserveTime")  'LogReserveTime
                    XmlFile.WriteValue(spinReserveLog.Value)                    
                    XmlFile.WriteEndElement()                           'LogReserveTime
                    'XmlFile.WriteStartElement("ReadTime")
                    'XmlFile.WriteValue(spinFreq.Value)
                    'XmlFile.WriteEndElement() 'ReadTime
                    'XmlFile.WriteStartElement("ProcCmdCount")
                    'XmlFile.WriteValue(spinProc.Value)
                    'XmlFile.WriteEndElement() 'ProcCmdCount
                    'XmlFile.WriteStartElement("MaxThread")
                    'XmlFile.WriteValue(spinMaxThreads.Value)
                    'XmlFile.WriteEndElement() 'MaxThread
                    'XmlFile.WriteStartElement("DBMinPool")
                    'XmlFile.WriteValue(spinMinP.Value)
                    'XmlFile.WriteEndElement() 'DBMinPool
                    'XmlFile.WriteStartElement("DBMaxPool")
                    'XmlFile.WriteValue(spinMaxP.Value)
                    'XmlFile.WriteEndElement() 'DBMaxPool
                    'XmlFile.WriteStartElement("DBPoolLiveTime")
                    'XmlFile.WriteValue(spinLifeP.Value)
                    'XmlFile.WriteEndElement() 'DBPoolLiveTime
                    XmlFile.WriteEndElement()   'System
                    '---------------------------------------------------------------------------------------------------------------------------
                    XmlFile.WriteStartElement("SODB")   'SODB
                    For i As Int32 = 0 To TreLstDB.Nodes.Count - 1
                        Dim Node As TreeListNode = TreLstDB.Nodes(i)
                        XmlFile.WriteStartElement("DB")     'DB
                        XmlFile.WriteAttributeString("SelectID", Node.Item("SelectID"))
                        XmlFile.WriteAttributeString("CompID", Node.Item("CompID"))
                        XmlFile.WriteAttributeString("DBSid", Node.Item("DBSid"))
                        XmlFile.WriteAttributeString("DBUser", Node.Item("DBUser"))
                        XmlFile.WriteAttributeString("DBPassword", Node.Item("DBPassword"))
                        XmlFile.WriteAttributeString("ComDBSid", Node.Item("ComDBSid"))
                        XmlFile.WriteAttributeString("ComDbUser", Node.Item("ComDbUser"))
                        XmlFile.WriteAttributeString("ComDBPassword", Node.Item("ComDBPassword"))
                        XmlFile.WriteAttributeString("MinDBPool", Node.Item(colMinDBPool))
                        XmlFile.WriteAttributeString("MaxDBPool", Node.Item(colMaxDBPool))
                        XmlFile.WriteAttributeString("DBPoolLifetime", Node.Item(colDBPoolLifetime))
                        XmlFile.WriteAttributeString("RunCmdTimer", Node.Item(colRunCmdTimer))
                        XmlFile.WriteAttributeString("InvoiceType", Node.Item(colInvoiceType))
                        XmlFile.WriteAttributeString("RunCommands", Node.Item(colRunCommands))
                        XmlFile.WriteAttributeString("LimtBeforeUpload", Node.Item(colLimtBeforeUpload))
                        XmlFile.WriteAttributeString("MisDBSid", Node.Item(colMisDBSid))
                        XmlFile.WriteAttributeString("MisDBUser", Node.Item(colMisDBUser))
                        XmlFile.WriteAttributeString("MisDBPassword", Node.Item(colMisDBPassword))
                        XmlFile.WriteAttributeString("MisOwner", Node.Item(colMisOwner))
                        XmlFile.WriteAttributeString("DBLink", Node.Item(colDBLink))
                        XmlFile.WriteAttributeString("InvBackupPath", Node.Item(colInvBackupPath))
                        XmlFile.WriteAttributeString("IsEmailNotify", Integer.Parse("0" & Node.Item(colIsEmailNotify)))
                        XmlFile.WriteAttributeString("IsSmsNotify", Integer.Parse("0" & Node.Item(colIsSmsNotify)))
                        XmlFile.WriteAttributeString("IsTvMailNotify", Integer.Parse("0" & Node.Item(colIsTvMailNotify)))
                        XmlFile.WriteAttributeString("IsCmNotify", Integer.Parse("0" & Node.Item(colIsCMNotify)))
                        XmlFile.WriteAttributeString("SendOrd", Node.Item(colSendOrd).ToString)
                        XmlFile.WriteAttributeString("AddSend", Node.Item(colAddSend))
                        XmlFile.WriteAttributeString("DonateMark", Node.Item(colDonateMark))
                        XmlFile.WriteAttributeString("ExportPath", Node.Item(colExportPath))
                        XmlFile.WriteAttributeString("AutoDo", Node.Item(colAutoDo))
                        XmlFile.WriteAttributeString("StartEmailNotify", Integer.Parse("0" & Node.Item(colStartEmailNotify)))
                        XmlFile.WriteAttributeString("StartSMSNotify", Integer.Parse("0" & Node.Item(colStartSMSNotify)))
                        XmlFile.WriteAttributeString("StartTVMailNotify", Integer.Parse("0" & Node.Item(colStartTVMailNotify)))
                        XmlFile.WriteAttributeString("StartCMNotify", Integer.Parse("0" & Node.Item(colStartCMNotify)))
                        XmlFile.WriteAttributeString("EndCMNotify", Integer.Parse("0" & Node.Item(colEndCMNotify)))
                        XmlFile.WriteAttributeString("CompType", Node.Item(colCompType))
                        XmlFile.WriteAttributeString("MaskInvNo", Integer.Parse("0" & Node.Item(colMaskInvNo)))
                        XmlFile.WriteAttributeString("StartNotifyPrize", Integer.Parse("0" & Node.Item(colStartNotifyPrize)))

                        ' tbSO.Columns.Add(New DataColumn("AddSend", GetType(String)))
                        'tbSO.Columns.Add(New DataColumn("DonateMark", GetType(String)))
                        'tbSO.Columns.Add(New DataColumn("ExportPath", GetType(String)))
                        'tbSO.Columns.Add(New DataColumn("AutoDo", GetType(String)))
                        XmlFile.WriteValue(Node.Item("CompName"))
                        XmlFile.WriteEndElement()            'DB
                    Next
                    XmlFile.WriteEndElement()                 'SODB
                    '------------------------------------------------------------------------------------------------------------------------------

                    XmlFile.WriteStartElement("CommandType") 'CommandType
                    For Each Node As TreeListNode In TreLstCmd.Nodes
                        XmlFile.WriteStartElement("Command")    'Command
                        XmlFile.WriteAttributeString("CodeNo", Node.Item("CodeNo"))
                        XmlFile.WriteValue(Node.Item("CmdName"))
                        XmlFile.WriteEndElement()                 'Command
                    Next
                    XmlFile.WriteEndElement()         'CommandType
                    '--------------------------------------------------------------------------------------------------------------------------------
                    XmlFile.WriteStartElement("RunCommand") 'RunCommand
                    For Each node As TreeListNode In TreLstSORunCmd.Nodes
                        XmlFile.WriteStartElement("RunType")        'RunType
                        XmlFile.WriteAttributeString("CompID", node.Item("CompID"))
                        XmlFile.WriteAttributeString("CompDescript", node.Item("CompDescript"))
                        XmlFile.WriteAttributeString("CmdType", node.Item("CmdType"))
                        XmlFile.WriteAttributeString("Runfrequency", node.Item("Runfrequency"))
                        XmlFile.WriteAttributeString("ExportElectronPath", node.Item(colExportElectronPath))
                        XmlFile.WriteValue(node.Item("CmdTypeName"))
                        XmlFile.WriteEndElement()                     'RunType
                    Next
                    XmlFile.WriteEndElement()             'RunCommand
                    '--------------------------------------------------------------------------------------------------------------------------------
                    XmlFile.WriteEndElement() 'CableSoft
                    XmlFile.WriteEndDocument()
                    XmlFile.Flush()
                    XmlFile.Close()

                    CableSoft.Invoice.Gateway.SetupFileIO.InvoiceSetupIO.WriteXmlFile(
                                                                                       String.Format("{0}\{1}",
                                                                                                     Environment.CurrentDirectory,
                                                                                                     xmlFileName),
                                                                                            mem, True)
                End Using
            End Using
            XtraMessageBox.Show("設定檔寫入完成！", "訊息", MessageBoxButtons.OK)
            Return True
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function

    Private Sub frmSetup_Disposed(sender As Object, e As System.EventArgs) Handles Me.Disposed
        Try
            If tbSODB IsNot Nothing Then
                tbSODB.Dispose()
            End If
            tbSODB = Nothing
        Catch ex As Exception

        End Try
    End Sub

    Private Sub frmSetup_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        GatewayIO = New CableSoft.Invoice.Gateway.SetupFileIO.InvoiceSetupIO(String.Format("{0}\{1}",
                                                                                         Environment.CurrentDirectory,
                                                                                         xmlFileName))

        ReadData()
       
    End Sub

    Private Sub btnAddLowCmd_Click(sender As System.Object, e As System.EventArgs) Handles btnAddLowCmd.Click
        Try
            TreLstCmd.AppendNode(Nothing, -1, -1, -1, 0).StateImageIndex = 0
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnAddHightCmd_Click(sender As System.Object, e As System.EventArgs) Handles btnAddHightCmd.Click
        Try
            TreLstSORunCmd.AppendNode(Nothing, -1, -1, -1, 0).StateImageIndex = 0
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub XtabCtl_SelectedPageChanged(sender As System.Object, e As DevExpress.XtraTab.TabPageChangedEventArgs) Handles XtabCtl.SelectedPageChanged
        If e.Page.Name.ToUpper = "XtabHighCmd".ToUpper Then
            riCboCmdType.Items.Clear()
            For Each nod As DevExpress.XtraTreeList.Nodes.TreeListNode In TreLstCmd.Nodes
                riCboCmdType.Items.Add(nod.Item("CodeNo"))
            Next
        End If
    End Sub

    Private Sub TreLstSORunCmd_CellValueChanging(sender As System.Object, e As DevExpress.XtraTreeList.CellValueChangedEventArgs) Handles TreLstSORunCmd.CellValueChanging
        Select Case e.Column.FieldName.ToUpper
            Case "CompId".ToUpper
                e.Node.Item("CompDescript") = Nothing
                For i As Integer = 0 To TreLstDB.Nodes.Count - 1
                    If Not String.IsNullOrEmpty(e.Value) Then
                        If TreLstDB.Nodes(i).Item("CompID") = e.Value Then
                            e.Node.Item("CompDescript") = TreLstDB.Nodes(i).Item("CompName")
                            Exit Sub
                        End If
                    End If
                Next
            Case "CmdType".ToUpper
                e.Node.Item("CmdTypeName") = Nothing
                For Each nod As DevExpress.XtraTreeList.Nodes.TreeListNode In TreLstCmd.Nodes
                    If Not String.IsNullOrEmpty(e.Value) Then
                        If nod.Item("CodeNo") = e.Value Then
                            e.Node.Item("CmdTypeName") = nod.Item("CmdName")
                            Exit Sub
                        End If
                    End If
                Next
        End Select
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Application.Exit()
    End Sub

    Private Sub btnTestConn_Click(sender As System.Object, e As System.EventArgs) Handles btnTestConn.Click
        Dim ConnectionBuilder As New OracleClient.OracleConnectionStringBuilder
        ConnectionBuilder.Unicode = True
        ConnectionBuilder.PersistSecurityInfo = True
        ConnectionBuilder.DataSource = TreLstDB.FocusedNode.Item(colDBSid).ToString
        ConnectionBuilder.UserID = TreLstDB.FocusedNode.Item(colDBUser).ToString
        ConnectionBuilder.Password = TreLstDB.FocusedNode.Item(colDBPassword).ToString
        Dim cn As OracleClient.OracleConnection = Nothing
        Try
            cn = New OracleClient.OracleConnection(ConnectionBuilder.ConnectionString)
            cn.Open()
            XtraMessageBox.Show("連接成功！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If cn IsNot Nothing Then
                cn.Close()
                cn.Dispose()
            End If
            cn = Nothing
        End Try        
    End Sub

    Private Sub riBtnEdtPath_ButtonClick(sender As System.Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles riBtnEdtPath.ButtonClick
        If FolderBrowserDialog1.ShowDialog() Then
            TreLstDB.FocusedNode.Item(TreLstDB.FocusedColumn) = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub ribtnElectronPath_ButtonClick(sender As System.Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles ribtnElectronPath.ButtonClick
        If FolderBrowserDialog1.ShowDialog() Then
            TreLstSORunCmd.FocusedNode.Item(TreLstSORunCmd.FocusedColumn) = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub btnDelHighCmd_Click(sender As System.Object, e As System.EventArgs) Handles btnDelHighCmd.Click
        Try
            TreLstSORunCmd.DeleteNode(TreLstSORunCmd.FocusedNode)
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub btnDelLowCmd_Click(sender As System.Object, e As System.EventArgs) Handles btnDelLowCmd.Click
        Try
            TreLstCmd.DeleteNode(TreLstCmd.FocusedNode)
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub chkLogSQL_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkLogSQL.CheckedChanged
        spinReserveLog.Enabled = False
        If chkLogSQL.Checked Then
            spinReserveLog.Enabled = True
        End If
    End Sub
End Class