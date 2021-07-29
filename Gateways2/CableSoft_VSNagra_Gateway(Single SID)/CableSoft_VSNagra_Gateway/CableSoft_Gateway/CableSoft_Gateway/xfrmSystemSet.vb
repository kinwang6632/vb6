Imports DevExpress.Skins
Imports DevExpress.XtraEditors
Imports CableSoft.CAS.CryptUtil
Imports System.ServiceProcess
Imports Microsoft.Win32
Imports System.Reflection
Imports CableSoft.Gateway.Common
'Imports System.Messaging

Public Class xfrmSystemSet
    Private Const WM_NCLBUTTONDOWN As Int32 = &HA1S
    Private Const HTCAPTION As Int32 = 2
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        SaveSystemSet()
        SaveSOSet()
        SaveNagraSet()
        SaveErrCode()
        SaveUseComp()
        SaveSocketSourceId()
        SaveMailSet()
        ProcAutoRun(chkAutoRun.Checked)
        Me.Close()
    End Sub
    Private Sub SaveSocketSourceId()
        Try
            CableSoft.GateWay.Common.MDBCommon.SaveTreeListToMDB(FMDBFile, FMDBPassword, fSourceIdTableName, TreLstSourceId)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub SaveUseComp()
        Try
            CableSoft.GateWay.Common.MDBCommon.SaveTreeListToMDB(FMDBFile, FMDBPassword, UseCompanyTableName, TreLstDB)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub SaveErrCode()
        Try
            CableSoft.GateWay.Common.MDBCommon.SaveTreeListToMDB(FMDBFile, FMDBPassword, ComparedErrorName, TreLstErrCode)

        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub SaveSOSet()
        Try
            Dim row As DataRow = Nothing
            Dim blnNew As Boolean = False
            If (FSOTable Is Nothing) OrElse (FSOTable.Rows.Count = 0) Then
                row = FSOTable.NewRow
                blnNew = True
            Else
                row = FSOTable.Rows(0)
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
                FSOTable.Rows.Add(row)
            End If
            CableSoft.GateWay.Common.MDBCommon.SaveMDBDataTable(FMDBFile, FMDBPassword, SOTableName, FSOTable, Nothing)
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
    Private Sub ShowMailSet()
        Try
            spinSocketErrCnt.Value = Int32.Parse(fSourceIdErrCnt)
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
                                                                     fMailToTableName, TreLstSMailTo)
            CableSoft.Gateway.Common.MDBCommon.ReadMDBDataToTreeList(MDBMailPara.GetDefaultMDBName, _
                                                                     fMDBMailPassword, _
                                                                     fMailCCTableName, TreLstSMailCC)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub ShowNagraSet()
        Try
            If (FNagraParaTable IsNot Nothing) _
                AndAlso (FNagraParaTable.Rows.Count > 0) Then
                'txtSndSourceId.Text = FSndSourceId(0)
                'txtSndMopPPId.Text = FSndMopPPId(0)
                'txtSndDestId.Text = FSndDestId(0)
                txtRcvDestId.Text = FRcvDestId
                txtRcvMopPPId.Text = FRcvMopPPId
                txtRcvSourceId.Text = FRcvSourceId
                spinSndErrCntStop.Value = FSndErrStopCnt
                spinRcvErrCntStop.Value = FRcvErrStopCnt
                spinChkStatus.Value = FCheckStatusTime
                spinSndDelay.Value = FSndDelayTime
                txtNagraIPAddress.Text = FNagraServer
                spinNagraPort.Value = FNagraPort
                spinMaxProcing.Value = FNagraMaxProcing
                txtUPDEN.Text = FNagraUpdEn
                spinResumeCnt.Value = Int32.Parse("0" & fResumeCnt)
                chkErrSndMail.Checked = fErrSndMail
                chkErrSndMail.Visible = fIsUseMail
             
            End If

        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub InitialData()
        Try
            If FSysTable IsNot Nothing AndAlso FSysTable.Rows.Count > 0 Then
                txtTitle.Text = FAppTitle
                chkAutoRun.Checked = FAutoRunGW
                chkTray.Checked = FUseTray
                chkResource.Checked = FShowResource
                chkHotKey.Checked = FUseHotKey
                spinFreq.Value = FReadDataTime
                spinProc.Value = FProcessNumber
                spinRcd.Value = FShowDataCount
                spinBuff.Value = FClearDataCount
                spinMaxThreads.Value = FMaxThread
                spinMinP.Value = FDBPoolMinNumber
                spinMaxP.Value = FDBPoolMaxNumber
                spinLifeP.Value = FDBPoolLiveTime
                spinTimeout.Value = FTestNagraSock
                txtIPAddress.Text = FNagraIPAddress.ToString
                spinPort.Value = FNagraPort
                If (FSOTable IsNot Nothing) AndAlso (FSOTable.Rows.Count > 0) Then
                    txtCOMSid.Text = GetFieldValue(FSOTable.Rows(0), "SID")
                    txtCOMUserId.Text = GetFieldValue(FSOTable.Rows(0), "UserId")
                    txtCOMPassword.Text = GetFieldValue(FSOTable.Rows(0), "UserPassword")
                End If

                '將MDB資料Show到Grid
                'CableSoft.GateWay.Common.MDBCommon.ReadMDBDataToTreeList(FMDBFile, FMDBPassword, SOTableName, TreLstDB, New String() {"UserPassword"})
                CableSoft.GateWay.Common.MDBCommon.ReadMDBDataToTreeList(FMDBFile, FMDBPassword, UseCompanyTableName, TreLstDB)
                CableSoft.GateWay.Common.MDBCommon.ReadMDBDataToTreeList(FMDBFile, FMDBPassword, ComparedErrorName, TreLstErrCode)
                CableSoft.GateWay.Common.MDBCommon.ReadMDBDataToTreeList(FMDBFile, FMDBPassword, fSourceIdTableName, TreLstSourceId)
                'PRM主機設定
                ShowNagraSet()
                'EMail設定
                If fIsUseMail Then
                    ShowMailSet()
                    For i As Int32 = 0 To TreLstSMailTo.Nodes.Count - 1
                        TreLstSMailTo.Nodes(i).StateImageIndex = 5
                    Next
                    For i As Int32 = 0 To TreLstSMailCC.Nodes.Count - 1
                        TreLstSMailCC.Nodes(i).StateImageIndex = 6
                    Next
                    XtabCtl.HeaderAutoFill = DevExpress.Utils.DefaultBoolean.False
                    
                Else
                    XtabCtl.TabPages(4).PageVisible = False
                    XtabCtl.HeaderAutoFill = DevExpress.Utils.DefaultBoolean.True
                End If

                For i As Int32 = 0 To TreLstDB.Nodes.Count - 1
                    TreLstDB.Nodes(i).StateImageIndex = 0
                Next
                For i As Int32 = 0 To TreLstErrCode.Nodes.Count - 1
                    TreLstErrCode.Nodes(i).StateImageIndex = 1
                Next
                For i As Int32 = 0 To TreLstSourceId.Nodes.Count - 1
                    TreLstSourceId.Nodes(i).StateImageIndex = 2
                Next
                
                lblCore.Text = String.Format("( {0} 核心 )", Environment.ProcessorCount)
            End If
            
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub SaveMailSet()
        Dim row As DataRow = Nothing
        Dim row2 As DataRow = Nothing
        Dim blnNew As Boolean = False
        Try
            If fMailInfoTable Is Nothing Then
                Exit Sub
            End If
            If (fMailInfoTable Is Nothing) OrElse (fMailInfoTable.Rows.Count = 0) Then
                row = fMailInfoTable.NewRow
                blnNew = True
            Else
                row = fMailInfoTable.Rows(0)
            End If
            If (fMailPRMParaTable Is Nothing) OrElse (fMailPRMParaTable.Rows.Count = 0) Then
                row2 = fMailPRMParaTable.NewRow
            Else
                row2 = fMailPRMParaTable.Rows(0)
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
                .Item("SourceIdErrCnt") = CryptUtil.Encrypt(spinSocketErrCnt.Value.ToString)
                .EndEdit()
            End With
            CableSoft.GateWay.Common.MDBCommon.SaveMDBDataTable(MDBMailPara.GetDefaultMDBName, _
                                                                fMDBMailPassword, fMailInfoTableName, fMailInfoTable, Nothing)
            CableSoft.GateWay.Common.MDBCommon.SaveMDBDataTable(MDBMailPara.GetDefaultMDBName, _
                                                                fMDBMailPassword, fMailPRMParaTableName, fMailPRMParaTable, Nothing)
            CableSoft.GateWay.Common.MDBCommon.SaveTreeListToMDB(MDBMailPara.GetDefaultMDBName, _
                                                                  fMDBMailPassword, fMailToTableName, TreLstSMailTo)
            CableSoft.GateWay.Common.MDBCommon.SaveTreeListToMDB(MDBMailPara.GetDefaultMDBName, _
                                                                  fMDBMailPassword, fMailCCTableName, TreLstSMailCC)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CableSoft.GateWay.csException.WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub

    Private Sub SaveNagraSet()
        Dim row As DataRow = Nothing
        Dim blnNew As Boolean = False
        Try
            If (FNagraParaTable Is Nothing) OrElse (FNagraParaTable.Rows.Count = 0) Then
                row = FNagraParaTable.NewRow
                blnNew = True
            Else
                row = FNagraParaTable.Rows(0)
            End If

            With row
                .BeginEdit()
                .Item("NagraIPAddress") = CryptUtil.Encrypt(txtNagraIPAddress.Text)
                .Item("NagraPort") = CryptUtil.Encrypt(spinNagraPort.Value)
                .Item("SndSourceId") = CryptUtil.Encrypt(txtSndSourceId.EditValue)
                .Item("SndDestId") = CryptUtil.Encrypt(txtSndDestId.EditValue)
                .Item("SndMopPPId") = CryptUtil.Encrypt(txtSndMopPPId.EditValue)
                .Item("RcvSourceId") = CryptUtil.Encrypt(txtRcvSourceId.EditValue)
                .Item("RcvDestId") = CryptUtil.Encrypt(txtRcvDestId.EditValue)
                .Item("RcvMopPPId") = CryptUtil.Encrypt(txtRcvMopPPId.EditValue)
                .Item("SenErrStopCnt") = CryptUtil.Encrypt(spinSndErrCntStop.Value)
                .Item("RcvErrStopCnt") = CryptUtil.Encrypt(spinRcvErrCntStop.Value)
                .Item("CheckStatusTime") = CryptUtil.Encrypt(spinChkStatus.Value)
                .Item("SndDelayTime") = CryptUtil.Encrypt(Convert.ToSingle(spinSndDelay.Value))
                .Item("SndTimeout") = CryptUtil.Encrypt("30")
                .Item("RcvTimeout") = CryptUtil.Encrypt("30")
                .Item("UPDEN") = CryptUtil.Encrypt(txtUPDEN.Text)
                .Item("MaxProcing") = CryptUtil.Encrypt(spinMaxProcing.Value)
                .Item("ResumeCmdCnt") = CryptUtil.Encrypt(spinResumeCnt.Value)
                .Item("CmdErrSndMail") = CryptUtil.Encrypt(Convert.ToInt32(chkErrSndMail.Checked))
                .EndEdit()
            End With

            CableSoft.Gateway.Common.MDBCommon.SaveMDBDataTable(FMDBFile, _
                                                                FMDBPassword, _
                                                                NagraParaName, _
                                                                FNagraParaTable, Nothing)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CableSoft.GateWay.csException.WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub
    Private Sub SaveSystemSet()
        Dim row As DataRow = Nothing
        Dim blnNew As Boolean = False
        Try
            If FSysTable Is Nothing OrElse FSysTable.Rows.Count = 0 Then
                row = FSysTable.NewRow
                blnNew = True
            Else
                row = FSysTable.Rows(0)
            End If
            'FSysTable.Rows.Clear()
            With FSysTable.Rows(0)
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

            CableSoft.GateWay.Common.MDBCommon.SaveMDBData(FMDBFile, FMDBPassword, FSysTable, SystemTableName)
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


    Private Sub btnAddSourceId_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddSourceId.Click
        Try
            TreLstSourceId.AppendNode(Nothing, -1, -1, -1, 0).StateImageIndex = 2
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub btnbtnRemoveSourceId_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnbtnRemoveSourceId.Click
        Try
            TreLstSourceId.DeleteNode(TreLstSourceId.FocusedNode)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnInsSMailTo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsSMailTo.Click
        Try
            TreLstSMailTo.AppendNode(Nothing, -1, -1, -1, 0).StateImageIndex = 5
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub btnDelSMailTo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelSMailTo.Click
        Try
            TreLstSMailTo.DeleteNode(TreLstSMailTo.FocusedNode)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnInsSMailCC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsSMailCC.Click
        Try
            TreLstSMailCC.AppendNode(Nothing, -1, -1, -1, 0).StateImageIndex = 6
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub btnDelSMailCC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelSMailCC.Click
        TreLstSMailCC.DeleteNode(TreLstSMailCC.FocusedNode)
    End Sub
End Class