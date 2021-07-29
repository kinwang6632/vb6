Imports DevExpress.Skins
Imports DevExpress.XtraEditors
Imports CableSoft.CAS.CryptUtil
Public Class xfrmSystemSet
    Private Const WM_NCLBUTTONDOWN As Int32 = &HA1S
    Private Const HTCAPTION As Int32 = 2
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        SaveSystemSet()
        SaveSOSet()
        SavePushServerSet()
        SaveErrCode()
        SaveUseComp()
        ProcAutoRun(FAutoRunGW)
        Me.Close()
    End Sub
    Private Sub SaveUseComp()
        Try
            Cablesoft.Gateway.Common.MDBCommon.SaveTreeListToMDB(FMDBFile, FMDBPassword, UseCompanyTableName, TreLstDB)
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
            Cablesoft.Gateway.Common.MDBCommon.SaveMDBDataTable(FMDBFile, FMDBPassword, SOTableName, FSOTable, Nothing)
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

                '顯示Com區 SO資料
                If (FSOTable IsNot Nothing) AndAlso (FSOTable.Rows.Count > 0) Then
                    txtCOMSid.Text = GetFieldValue(FSOTable.Rows(0), "SID")
                    txtCOMUserId.Text = GetFieldValue(FSOTable.Rows(0), "UserId")
                    txtCOMPassword.Text = GetFieldValue(FSOTable.Rows(0), "UserPassword")
                End If

                '顯示Push Server的資料
                txtPushUrl.Text = FPushUrl
                spinPushDuration.Text = FPushDuration
                spinPushRepeat.Text = FPushRepeat
                spinPushWebTimeOut.Value = FPushUrlTimeout
                cboPushUrlMethod.Text = FPushUrlMethod

                '將MDB資料Show到Grid
                'CableSoft.GateWay.Common.MDBCommon.ReadMDBDataToTreeList(FMDBFile, FMDBPassword, SOTableName, TreLstDB, New String() {"UserPassword"})
                Cablesoft.Gateway.Common.MDBCommon.ReadMDBDataToTreeList(FMDBFile, FMDBPassword, UseCompanyTableName, TreLstDB)
                CableSoft.GateWay.Common.MDBCommon.ReadMDBDataToTreeList(FMDBFile, FMDBPassword, ComparedErrorName, TreLstErrCode)
                For i As Int32 = 0 To TreLstDB.Nodes.Count - 1
                    TreLstDB.Nodes(i).StateImageIndex = 0
                Next
                For i As Int32 = 0 To TreLstErrCode.Nodes.Count - 1
                    TreLstErrCode.Nodes(i).StateImageIndex = 1
                Next
                lblCore.Text = String.Format("( {0} 核心 )", Environment.ProcessorCount)
            End If
            'ShowNagraSet()
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub SavePushServerSet()
        Dim row As DataRow = Nothing
        Dim blnNew As Boolean = False
        Try
            If (FPushServerParaTable Is Nothing) OrElse (FPushServerParaTable.Rows.Count = 0) Then
                row = FPushServerParaTable.NewRow
                blnNew = True
            Else
                row = FPushServerParaTable.Rows(0)
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
                .EndEdit()
            End With
            CableSoft.GateWay.Common.MDBCommon.SaveMDBDataTable(FMDBFile, FMDBPassword, PushServerParaTableName, FPushServerParaTable)
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CableSoft.Gateway.csException.WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
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
            
            Cablesoft.Gateway.Common.MDBCommon.SaveMDBData(FMDBFile, FMDBPassword, FSysTable, SystemTableName)
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

    Private Sub btnAddCMD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub SimpleButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

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


End Class