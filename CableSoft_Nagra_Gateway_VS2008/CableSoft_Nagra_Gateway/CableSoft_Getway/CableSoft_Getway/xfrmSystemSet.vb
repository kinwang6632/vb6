Imports DevExpress.Skins
Imports DevExpress.XtraEditors
Public Class xfrmSystemSet
    Private Const WM_NCLBUTTONDOWN As Int32 = &HA1S
    Private Const HTCAPTION As Int32 = 2
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        SaveSystemSet()
        SaveSOSet()
        SaveNagraSet()
        SaveErrCode()
        Me.Close()
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
            CableSoft.GateWay.Common.MDBCommon.SaveTreeListToMDB(FMDBFile, FMDBPassword, SOTableName, TreLstDB, New String() {"UserPassword"})
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
        InitialData()
    End Sub
    Private Sub ShowNagraSet()
        Try
            If FNagraParaTable IsNot Nothing AndAlso FNagraParaTable.Rows.Count > 0 Then
                txtSndSourceId.Text = FSndSourceId
                txtSndMopPPId.Text = FSndMopPPId
                txtSndDestId.Text = FSndDestId
                txtRcvDestId.Text = FRcvDestId
                txtRcvMopPPId.Text = FRcvMopPPId
                txtRcvSourceId.Text = FRcvSourceId
                spinSndErrCntStop.Value = FSenErrStopCnt
                spinRcvErrCntStop.Value = FRcvErrStopCnt
                spinChkStatus.Value = FCheckStatusTime
                spinSndDelay.Value = FSndDelayTime
                txtNagraIPAddress.Text = FNagraServer
                spinNagraPort.Value = FNagraPort
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
                '將MDB資料Show到Grid
                CableSoft.GateWay.Common.MDBCommon.ReadMDBDataToTreeList(FMDBFile, FMDBPassword, SOTableName, TreLstDB, New String() {"UserPassword"})
                CableSoft.GateWay.Common.MDBCommon.ReadMDBDataToTreeList(FMDBFile, FMDBPassword, ComparedErrorName, TreLstErrCode)
                For i As Int32 = 0 To TreLstDB.Nodes.Count - 1
                    TreLstDB.Nodes(i).StateImageIndex = 0
                Next
                For i As Int32 = 0 To TreLstErrCode.Nodes.Count - 1
                    TreLstErrCode.Nodes(i).StateImageIndex = 1
                Next
                lblCore.Text = String.Format("( {0} 核心 )", Environment.ProcessorCount)
            End If
            ShowNagraSet()
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
            With FNagraParaTable.Rows(0)
                .BeginEdit()
                .Item("NagraIPAddress") = txtNagraIPAddress.Text
                .Item("NagraPort") = spinNagraPort.Value
                .Item("SndSourceId") = txtSndSourceId.EditValue
                .Item("SndDestId") = txtSndDestId.EditValue
                .Item("SndMopPPId") = txtSndMopPPId.EditValue
                .Item("RcvSourceId") = txtRcvSourceId.EditValue
                .Item("RcvDestId") = txtRcvDestId.EditValue
                .Item("RcvMopPPId") = txtRcvMopPPId.EditValue
                .Item("SenErrStopCnt") = spinSndErrCntStop.Value
                .Item("RcvErrStopCnt") = spinRcvErrCntStop.Value
                .Item("CheckStatusTime") = spinChkStatus.Value
                .Item("SndDelayTime") = Convert.ToSingle(spinSndDelay.Value)
                .Item("SndTimeout") = 30
                .Item("RcvTimeout") = 30

                .EndEdit()
            End With
            CableSoft.Gateway.Common.MDBCommon.SaveMDBData(FMDBFile, FMDBPassword, FNagraParaTable, NagraParaName)
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
                .Item("AppCaption") = txtTitle.Text
                .Item("AutoRunGw") = Convert.ToInt32(chkAutoRun.Checked)
                .Item("UseTray") = Convert.ToInt32(chkTray.Checked)
                .Item("ShowResource") = Convert.ToInt32(chkResource.Checked)
                .Item("UseHotKey") = Convert.ToInt32(chkHotKey.Checked)
                .Item("ReadDataTime") = spinFreq.Value
                .Item("ProcessNumber") = spinProc.Value
                .Item("ShowDataCount") = spinRcd.Value
                .Item("ClearDataCount") = spinBuff.Value
                .Item("MaxThread") = spinMaxThreads.Value
                .Item("DBPoolMinNumber") = spinMinP.Value
                .Item("DBPoolMaxNumber") = spinMaxP.Value
                .Item("DBPoolLiveTime") = spinLifeP.Value
                .Item("TestSocketTime") = spinTimeout.Value
                .Item("NagraIPAddress") = txtIPAddress.Text
                .Item("NagraPort") = spinPort.Value
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
            Dim aSID As String = TreLstDB.FocusedNode.Item("SID")
            Dim aUserId As String = TreLstDB.FocusedNode.Item("UserId")
            Dim aPassword As String = TreLstDB.FocusedNode.Item("UserPassword")
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
End Class