Imports DevExpress.Skins
Imports DevExpress.XtraEditors
Imports CableSoft.CAS.CryptUtil
Imports CableSoft.GateWay.Common
Public Class xfrmSystemSet
    Private Const WM_NCLBUTTONDOWN As Int32 = &HA1S
    Private Const HTCAPTION As Int32 = 2
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        SaveSystemSet()
        SaveSOSet()
        SavePushServerSet()
        SaveErrCode()
        SaveUseComp()
        SaveServerUrl()
        SaveMailSet()
        SaveTVMailMsgSet()
        ProcAutoRun(chkAutoRun.Checked)

        Me.Close()
    End Sub
    Private Sub SaveUseComp()
        Try
            CableSoft.Gateway.Common.MDBCommon.SaveTreeListToMDB(fMDBFile, fMDBPassword, fUseCompanyTableName, TreLstDB)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub SaveServerUrl()
        Try
            CableSoft.Gateway.Common.MDBCommon.SaveTreeListToMDB(fMDBFile, fMDBPassword, fPushServerUrlTableName, TreLstServerUrl)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub SaveErrCode()
        Try
            CableSoft.GateWay.Common.MDBCommon.SaveTreeListToMDB(fMDBFile, fMDBPassword, fComparedErrorTableName, TreLstErrCode)

        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub SaveTVMailMsgSet()
        Try
            Dim row As DataRow = Nothing
            Dim blnNew As Boolean = False
            If (fTVmailSetTable Is Nothing) OrElse (fTVmailSetTable.Rows.Count = 0) Then
                row = fTVmailSetTable.NewRow
                blnNew = True
            Else
                row = fTVmailSetTable.Rows(0)
            End If
            With row
                .BeginEdit()
                If Not String.IsNullOrEmpty(edtTVMailMsg.Text) Then
                    .Item("TVMailMsg") = CryptUtil.Encrypt(edtTVMailMsg.Text)
                Else
                    .Item("TVMailMsg") = String.Empty
                End If

                .EndEdit()
            End With
            If blnNew Then
                fTVmailSetTable.Rows.Add(row)
            End If
            CableSoft.GateWay.Common.MDBCommon.SaveMDBDataTable(fMDBFile, fMDBPassword, fTVMailSetTableName, fTVmailSetTable, Nothing)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SaveSOSet()
        Try
            Dim row As DataRow = Nothing
            Dim blnNew As Boolean = False
            If (fSOTable Is Nothing) OrElse (fSOTable.Rows.Count = 0) Then
                row = fSOTable.NewRow
                blnNew = True
            Else
                row = fSOTable.Rows(0)
            End If
            'With row
            '    .BeginEdit()
            '    .Item("UserPassword") = CryptUtil.Encrypt(FSOTable.Rows(0).Item("UserPassword"))
            '    .EndEdit()
            '    row.AcceptChanges()
            'End With
            With row
                .BeginEdit()
                If Not String.IsNullOrEmpty(txtCOMSid.Text) Then
                    .Item("SID") = CryptUtil.Encrypt(txtCOMSid.Text)
                End If
                If Not String.IsNullOrEmpty(txtCOMUserId.Text) Then
                    .Item("UserId") = CryptUtil.Encrypt(txtCOMUserId.Text)
                End If
                If Not String.IsNullOrEmpty(txtCOMPassword.Text) Then
                    .Item("UserPassword") = CryptUtil.Encrypt(txtCOMPassword.Text)
                End If
                .EndEdit()
            End With
            If blnNew Then
                fSOTable.Rows.Add(row)
            End If
            Cablesoft.Gateway.Common.MDBCommon.SaveMDBDataTable(fMDBFile, fMDBPassword, fSOTableName, fSOTable, Nothing)
            'CableSoft.GateWay.Common.MDBCommon.SaveTreeListToMDB(FMDBFile, FMDBPassword, SOTableName, TreLstDB, New String() {"UserPassword"})
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Public Sub New()

        ' 此為 Windows Form 設計工具所需的呼叫。
        InitializeComponent()

        ' 在 InitializeComponent() 呼叫之後加入任何初始設定。

    End Sub
    Private Sub xfrmSystemSet_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        XtabCtl.SelectedTabPageIndex = 0
        InitialData()
    End Sub
    'Private Sub ShowNagraSet()
    '    Try
    '        If (FNagraParaTable IsNot Nothing) _
    '            AndAlso (FNagraParaTable.Rows.Count > 0) Then
    '            txtSndSourceId.Text = FSndSourceId
    '            txtSndMopPPId.Text = FSndMopPPId
    '            txtSndDestId.Text = FSndDestId
    '            txtRcvDestId.Text = FRcvDestId
    '            txtRcvMopPPId.Text = FRcvMopPPId
    '            txtRcvSourceId.Text = FRcvSourceId
    '            spinSndErrCntStop.Value = FSndErrStopCnt
    '            spinRcvErrCntStop.Value = FRcvErrStopCnt
    '            spinChkStatus.Value = FCheckStatusTime
    '            spinSndDelay.Value = FSndDelayTime
    '            txtNagraIPAddress.Text = FNagraServer
    '            spinNagraPort.Value = FNagraPort
    '            spinMaxProcing.Value = FNagraMaxProcing
    '            txtUPDEN.Text = FNagraUpdEn
    '        End If

    '    Catch ex As Exception
    '        XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub
    Private Sub ShowMailSet()

        Try
            spinWebErrCnt.Value = Int32.Parse(fReadWebErrCnt)
            txtSMTPHost.Text = fSmtpHost
            spinSMTPPort.Value = Int32.Parse(fSmtpPort)
            txtSMTPLoginID.Text = fSmtplUserId
            txtSMTPLoginPsw.Text = fSmtpPassword
            txtSender.Text = fMailSender
            txtDisplayName.Text = fMailDisplayName
            chkEnabledSSL.Checked = fEnabledSSL
            txtMailSubject.Text = fMailSubject
            txtMailBody.Text = fMailBody

            CableSoft.Gateway.Common.MDBCommon.ReadMDBDataToTreeList(MDBMailPara.GetDefaultMDBName, _
                                                                     fMDBMailPassword, _
                                                                     fMailToTableName, TreLstMailTo)
            CableSoft.Gateway.Common.MDBCommon.ReadMDBDataToTreeList(MDBMailPara.GetDefaultMDBName, _
                                                                     fMDBMailPassword, _
                                                                     fMailCCTableName, TreLstMailCC)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub InitialData()
        Try
            If fSysTable IsNot Nothing AndAlso fSysTable.Rows.Count > 0 Then
                txtTitle.Text = fAppTitle
                chkAutoRun.Checked = fAutoRunGW
                chkTray.Checked = fUseTray
                chkResource.Checked = fShowResource
                chkHotKey.Checked = fUseHotKey
                spinFreq.Value = fReadDataTime
                spinProc.Value = fProcessNumber
                spinRcd.Value = fShowDataCount
                spinBuff.Value = fClearDataCount
                spinMaxThreads.Value = fMaxThread
                spinMinP.Value = fDBPoolMinNumber
                spinMaxP.Value = fDBPoolMaxNumber
                spinLifeP.Value = fDBPoolLiveTime
                edtTVMailMsg.EditValue = fTVMailMDBMsgTxt

                '顯示Com區 SO資料
                If (fSOTable IsNot Nothing) AndAlso (fSOTable.Rows.Count > 0) Then
                    txtCOMSid.Text = GetFieldValue(fSOTable.Rows(0), "SID")
                    txtCOMUserId.Text = GetFieldValue(fSOTable.Rows(0), "UserId")
                    txtCOMPassword.Text = GetFieldValue(fSOTable.Rows(0), "UserPassword")
                End If

                '顯示Push Server的資料
                txtPushUrl.Text = fPushUrl
                spinPushDuration.Text = fPushDuration
                spinPushRepeat.Text = fPushRepeat
                spinPushWebTimeOut.Value = fPushUrlTimeout
                cboPushUrlMethod.Text = fPushUrlMethod
                spinPushRetryNum.Text = fPushServerRetryNum
                chkErrSndMail.Checked = fCmdErrSndMail
                chkSendTVMail.Checked = fSendTVMail
                spinResumeCnt.Value = fNoneResumeCnt

                '將MDB資料Show到Grid
                'CableSoft.GateWay.Common.MDBCommon.ReadMDBDataToTreeList(FMDBFile, FMDBPassword, SOTableName, TreLstDB, New String() {"UserPassword"})
                Cablesoft.Gateway.Common.MDBCommon.ReadMDBDataToTreeList(fMDBFile, fMDBPassword, fUseCompanyTableName, TreLstDB)
                CableSoft.GateWay.Common.MDBCommon.ReadMDBDataToTreeList(fMDBFile, fMDBPassword, fComparedErrorTableName, TreLstErrCode)
                CableSoft.Gateway.Common.MDBCommon.ReadMDBDataToTreeList(fMDBFile, fMDBPassword, fPushServerUrlTableName, TreLstServerUrl)
                If fIsUseMail Then
                    ShowMailSet()
                    For i As Int32 = 0 To TreLstMailTo.Nodes.Count - 1
                        TreLstMailTo.Nodes(i).StateImageIndex = 5
                    Next
                    For i As Int32 = 0 To TreLstMailCC.Nodes.Count - 1
                        TreLstMailCC.Nodes(i).StateImageIndex = 6
                    Next
                    chkErrSndMail.Visible = True
                    XtabCtl.HeaderAutoFill = DevExpress.Utils.DefaultBoolean.False
                Else
                    XtabCtl.TabPages(4).PageVisible = False
                    chkErrSndMail.Visible = False
                    XtabCtl.HeaderAutoFill = DevExpress.Utils.DefaultBoolean.True
                End If

                For i As Int32 = 0 To TreLstDB.Nodes.Count - 1
                    TreLstDB.Nodes(i).StateImageIndex = 0
                Next
                For i As Int32 = 0 To TreLstErrCode.Nodes.Count - 1
                    TreLstErrCode.Nodes(i).StateImageIndex = 1
                Next
                For i As Int32 = 0 To TreLstServerUrl.Nodes.Count - 1
                    TreLstServerUrl.Nodes(i).StateImageIndex = 2
                Next
                munCustomerField.Items.Clear()
                If (fLstCustomerVar IsNot Nothing) AndAlso (fLstCustomerVar.Count > 0) Then
                    For i As Int32 = 0 To fLstCustomerVar.Count - 1
                        AddHandler munCustomerField.Items.Add(fLstCustomerVar.Item(i)).Click,
                            AddressOf mItm_Click
                    Next
                Else
                    munCustomerField.Items.Add("無任何變數可使用")
                End If
                lblCore.Text = String.Format("( {0} 核心 )", Environment.ProcessorCount)
            End If
            'ShowNagraSet()
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub SaveMailSet()
        Dim row As DataRow = Nothing
        Dim row2 As DataRow = Nothing
        Dim blnNew As Boolean = False
        Try
            If (fMailInfoTable Is Nothing) OrElse (fMailInfoTable.Rows.Count = 0) Then
                row = fMailInfoTable.NewRow
                blnNew = True
            Else
                row = fMailInfoTable.Rows(0)
            End If
            If (fMailPSParaTable Is Nothing) OrElse (fMailPSParaTable.Rows.Count = 0) Then
                row2 = fMailPSParaTable.NewRow
            Else
                row2 = fMailPSParaTable.Rows(0)
            End If
            With row
                .BeginEdit()
                .Item("SmtpHost") = CryptUtil.Encrypt(txtSMTPHost.Text)
                .Item("SmtpUserId") = CryptUtil.Encrypt(txtSMTPLoginID.Text)
                .Item("SmtpPassword") = CryptUtil.Encrypt(txtSMTPLoginPsw.Text)
                If chkEnabledSSL.Checked Then
                    .Item("EnabledSSL") = CryptUtil.Encrypt("1")
                Else
                    .Item("EnabledSSL") = CryptUtil.Encrypt("0")
                End If
                .Item("MailDisplayNam") = CryptUtil.Encrypt(txtDisplayName.Text)
                .Item("MailSender") = CryptUtil.Encrypt(txtSender.Text)
                .Item("SmtpPort") = CryptUtil.Encrypt(spinSMTPPort.Value.ToString)
                .Item("MailSubject") = CryptUtil.Encrypt(txtMailSubject.Text)
                .Item("MailBody") = CryptUtil.Encrypt(txtMailBody.Text)
                .EndEdit()
            End With
            With row2
                .BeginEdit()
                .Item("ReadWebErrCnt") = CryptUtil.Encrypt(spinWebErrCnt.Value.ToString)
                .EndEdit()
            End With
            CableSoft.GateWay.Common.MDBCommon.SaveMDBDataTable(MDBMailPara.GetDefaultMDBName, _
                                                                fMDBMailPassword, fMailInfoTableName, fMailInfoTable, Nothing)
            CableSoft.GateWay.Common.MDBCommon.SaveMDBDataTable(MDBMailPara.GetDefaultMDBName, _
                                                                fMDBMailPassword, fMailPSParaTableName, fMailPSParaTable, Nothing)
            CableSoft.GateWay.Common.MDBCommon.SaveTreeListToMDB(MDBMailPara.GetDefaultMDBName, _
                                                                  fMDBMailPassword, fMailToTableName, TreLstMailTo)
            CableSoft.GateWay.Common.MDBCommon.SaveTreeListToMDB(MDBMailPara.GetDefaultMDBName, _
                                                                  fMDBMailPassword, fMailCCTableName, TreLstMailCC)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CableSoft.GateWay.csException.WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub


    Private Sub SavePushServerSet()
        Dim row As DataRow = Nothing
        Dim blnNew As Boolean = False
        Try
            If (fPushServerParaTable Is Nothing) OrElse (fPushServerParaTable.Rows.Count = 0) Then
                row = fPushServerParaTable.NewRow
                blnNew = True
            Else
                row = fPushServerParaTable.Rows(0)
            End If
            With row
                .BeginEdit()


                If Not String.IsNullOrEmpty(txtPushUrl.Text) Then
                    .Item("PushServerUrl") = txtPushUrl.Text
                End If
                .Item("PushServerUrlMethod") = cboPushUrlMethod.Text
                .Item("PushServerUrlTimeOut") = spinPushWebTimeOut.Text
                .Item("PushServerDuration") = spinPushDuration.Text
                .Item("PushServerRepeat") = spinPushRepeat.Text
                .Item("PushServerRetryNum") = spinPushRetryNum.Text
                .Item("PushServerSlowTime") = spinPushSlowTime.Text
                If chkErrSndMail.Checked Then
                    .Item("CmdErrSndMail") = "1"
                Else
                    .Item("CmdErrSndMail") = "0"
                End If
                .Item("SendTVMail") = "0"
                If chkSendTVMail.Checked Then
                    .Item("SendTVMail") = "1"
                End If

                .Item("NoneResumeCnt") = spinResumeCnt.Text
                .EndEdit()
            End With
            CableSoft.GateWay.Common.MDBCommon.SaveMDBDataTable(fMDBFile, fMDBPassword, fPushServerParaTableName, fPushServerParaTable)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CableSoft.GateWay.csException.WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub
    Private Sub SaveSystemSet()
        Dim row As DataRow = Nothing
        Dim blnNew As Boolean = False
        Try
            If fSysTable Is Nothing OrElse fSysTable.Rows.Count = 0 Then
                row = fSysTable.NewRow
                blnNew = True
            Else
                row = fSysTable.Rows(0)
            End If
            'FSysTable.Rows.Clear()
            With fSysTable.Rows(0)
                .BeginEdit()
                If Not String.IsNullOrEmpty(txtTitle.Text) Then
                    .Item("AppCaption") = CryptUtil.Encrypt(txtTitle.Text)
                End If
                .Item("AutoRunGw") = CryptUtil.Encrypt(Convert.ToInt32(chkAutoRun.Checked))
                .Item("UseTray") = CryptUtil.Encrypt(Convert.ToInt32(chkTray.Checked))
                .Item("ShowResource") = CryptUtil.Encrypt(Convert.ToInt32(chkResource.Checked))
                .Item("UseHotKey") = CryptUtil.Encrypt(Convert.ToInt32(chkHotKey.Checked))
                .Item("ReadDataTime") = CryptUtil.Encrypt(spinFreq.Value)
                .Item("ProcessNumber") = CryptUtil.Encrypt(spinProc.Value)
                .Item("ShowDataCount") = CryptUtil.Encrypt(spinRcd.Value)
                .Item("ClearDataCount") = CryptUtil.Encrypt(spinBuff.Value)
                .Item("MaxThread") = CryptUtil.Encrypt(spinMaxThreads.Value)
                .Item("DBPoolMinNumber") = CryptUtil.Encrypt(spinMinP.Value)
                .Item("DBPoolMaxNumber") = CryptUtil.Encrypt(spinMaxP.Value)
                .Item("DBPoolLiveTime") = CryptUtil.Encrypt(spinLifeP.Value)
                '.Item("TestSocketTime") = spinTimeout.Value
                '.Item("NagraIPAddress") = txtIPAddress.Text
                '.Item("NagraPort") = spinPort.Value
                .EndEdit()

            End With

            CableSoft.GateWay.Common.MDBCommon.SaveMDBData(fMDBFile, fMDBPassword, fSysTable, fSystemTableName)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CableSoft.GateWay.csException.WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub

    Private Sub btnAddConn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddConn.Click
        Try
            TreLstDB.AppendNode(Nothing, -1, -1, -1, 0).StateImageIndex = 0

        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'Msg.Show(ex.ToString(), "訊息", MsgBtn.OK, MsgIco.Error)
            'Bs.Write2MdbLog(Nothing, Nothing, Nothing, ex)
        End Try
    End Sub



    Private Sub btnRemoveConn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveConn.Click
        Try
            TreLstDB.DeleteNode(TreLstDB.FocusedNode)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnTestConn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTestConn.Click

        'If (Not TreLstDB.Focused) OrElse (TreLstDB.FocusedNode Is Nothing) Then
        '    Exit Sub
        'End If

        Try

            Dim aSidDb As String = TreLstDB.FocusedNode.Item("SidDBname")
            Dim aSidUserId As String = TreLstDB.FocusedNode.Item("SidUserId")
            Dim aSidPassword As String = TreLstDB.FocusedNode.Item("SidPassword")

            If String.IsNullOrEmpty(aSidDb) Then
                XtraMessageBox.Show("SID未設定", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Exit Sub
            End If
            If String.IsNullOrEmpty(aSidUserId) Then
                XtraMessageBox.Show("User Id 未設定", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Exit Sub
            End If
            If String.IsNullOrEmpty(aSidPassword) Then
                XtraMessageBox.Show("Password 未設定", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Exit Sub
            End If

            Dim aMsg As String = String.Empty
            Dim _ConnectString As String = "Data Source={0};Persist Security Info=True;" & _
                "User ID={1};Password={2};Unicode=True"
            _ConnectString = String.Format(_ConnectString, aSidDb, aSidUserId, aSidPassword)
            Using cn As New OracleClient.OracleConnection(_ConnectString)
                If IsConnectOK(cn, aMsg) Then
                    lblMsg.Text = "連線成功"
                Else
                    lblMsg.Text = aMsg
                End If
            End Using
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TreLstDB_FocusedNodeChanged(ByVal sender As System.Object, ByVal e As DevExpress.XtraTreeList.FocusedNodeChangedEventArgs) Handles TreLstDB.FocusedNodeChanged
        On Error Resume Next
        lblMsg.Text = ""
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub grpCtl_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grpCtl.MouseDown
        Try
            If e.Button = MouseButtons.Left Then
                Dim c As Control = DirectCast(sender, Control)
                c.Capture = False
                DefWndProc(Message.Create(ActiveForm.Handle, WM_NCLBUTTONDOWN, New IntPtr(HTCAPTION), IntPtr.Zero))
            End If
        Catch ex As Exception
            CableSoft.GateWay.csException.WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub



    Private Sub btnAddAuth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAuth.Click
        Try

            TreLstErrCode.AppendNode(Nothing, -1, -1, -1, 0).StateImageIndex = 1
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub btnRemoveAuth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveAuth.Click
        Try
            TreLstErrCode.DeleteNode(TreLstErrCode.FocusedNode)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub





    Private Sub xfrmSystemSet_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Static aDeveloper As String = String.Empty
        If e.Control Then
            If e.KeyCode <> Keys.ControlKey Then
                aDeveloper = aDeveloper & Convert.ToChar(e.KeyValue)
            Else
                aDeveloper = String.Empty
            End If

            If aDeveloper.ToUpper = "DEVELOPER" Then
                'Do Something
            End If
        Else
            aDeveloper = String.Empty
        End If
    End Sub



    Private Sub btnAddUrl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddUrl.Click
        Try
            TreLstServerUrl.AppendNode(Nothing, -1, -1, -1, 0).StateImageIndex = 2
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub btnRemoveUrl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveUrl.Click
        Try
            TreLstServerUrl.DeleteNode(TreLstServerUrl.FocusedNode)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnInsSMailTo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsSMailTo.Click
        Try
            TreLstMailTo.AppendNode(Nothing, -1, -1, -1, 0).StateImageIndex = 5
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelSMailTo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelSMailTo.Click
        Try
            TreLstMailTo.DeleteNode(TreLstMailTo.FocusedNode)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnInsMailCC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsMailCC.Click
        Try
            TreLstMailCC.AppendNode(Nothing, -1, -1, -1, 0).StateImageIndex = 6
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelMailCC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelMailCC.Click
        Try
            TreLstMailCC.DeleteNode(TreLstMailCC.FocusedNode)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub btnTestCom_Click(sender As System.Object, e As System.EventArgs) Handles btnTestCom.Click

        Try
            If String.IsNullOrEmpty(txtCOMSid.Text) Then
                XtraMessageBox.Show("COM區 SID未設定", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If txtCOMSid.CanFocus Then
                    txtCOMSid.Focus()
                End If
                Exit Sub
            End If
            If String.IsNullOrEmpty(txtCOMUserId.Text) Then
                XtraMessageBox.Show("COM區 User Id 未設定", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If txtCOMUserId.CanFocus Then
                    txtCOMUserId.Focus()
                End If
                Exit Sub
            End If
            If String.IsNullOrEmpty(txtCOMPassword.Text) Then
                XtraMessageBox.Show("COM區 Password 未設定", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If txtCOMPassword.CanFocus Then
                    txtCOMPassword.Focus()
                End If
                Exit Sub
            End If
            Dim aSID As String = txtCOMSid.Text
            Dim aUserId As String = txtCOMUserId.Text
            Dim aPassword As String = txtCOMPassword.Text
            Dim aMsg As String = String.Empty
            Dim _ConnectString As String = "Data Source={0};Persist Security Info=True;" & _
                "User ID={1};Password={2};Unicode=True"
            _ConnectString = String.Format(_ConnectString, aSID, aUserId, aPassword)
            Using cn As New OracleClient.OracleConnection(_ConnectString)
                If IsConnectOK(cn, aMsg) Then
                    lblComMsg.Text = "連線成功"
                Else
                    lblComMsg.Text = aMsg
                End If
            End Using
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub mItm_Click(sender As System.Object, e As System.EventArgs)
        Try
            Dim iPos As Integer = edtTVMailMsg.SelectionStart
            edtTVMailMsg.Text = edtTVMailMsg.Text.Substring(0, edtTVMailMsg.SelectionStart) &
                            "<" & CType(sender, ToolStripMenuItem).Text & ">" &
                            edtTVMailMsg.Text.Substring(edtTVMailMsg.SelectionStart)

            edtTVMailMsg.SelectionStart = iPos + CType(sender, ToolStripMenuItem).Text.Length + 2
        Catch ex As Exception
            CableSoft.GateWay.csException.WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub

    Private Sub item1_Click(sender As System.Object, e As System.EventArgs) Handles item1.Click

        Dim iPos As Integer = edtTVMailMsg.SelectionStart
        edtTVMailMsg.Text = edtTVMailMsg.Text.Substring(0, edtTVMailMsg.SelectionStart) &
                        item1.Text & edtTVMailMsg.Text.Substring(edtTVMailMsg.SelectionStart)

        edtTVMailMsg.SelectionStart = iPos + item1.Text.Length
    End Sub

    Private Sub XtabCtl_Click(sender As System.Object, e As System.EventArgs) Handles XtabCtl.Click
        Try
            If XtabCtl.SelectedTabPageIndex = 3 Then
                edtTVMailMsg.Focus()
                edtTVMailMsg.SelectionStart = edtTVMailMsg.Text.Length
                edtTVMailMsg.SelectionLength = 0
            End If
        Catch ex As Exception
            CableSoft.GateWay.csException.WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
        
    End Sub

    Private Sub chkUseRightKey_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkUseRightKey.CheckedChanged
        Try
            If chkUseRightKey.Checked Then
                edtTVMailMsg.Properties.ContextMenuStrip = munCustomerField
            Else
                edtTVMailMsg.Properties.ContextMenuStrip = Nothing
            End If
        Catch ex As Exception
            CableSoft.GateWay.csException.WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
        
    End Sub
End Class