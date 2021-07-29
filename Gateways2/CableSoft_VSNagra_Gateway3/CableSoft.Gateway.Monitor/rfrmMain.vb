Imports DevExpress.Skins
Imports DevExpress.XtraEditors
Imports System.IO
Imports System.Text
Imports DevExpress.XtraBars.Docking
Imports DevExpress.XtraBars.Ribbon
Imports DevExpress.XtraBars

Public Class rfrmMain

    Private brItem As DevExpress.XtraBars.BarItem = Nothing
    
    Private Const IsEncryFile As Boolean = True
    Private arySO() As SO = Nothing
    Private Enum ButtonlStates
        Inital = 0
        Start = 1
        StopRun = 2
    End Enum
    Private Sub ControlButtonEnable(ByVal State As ButtonlStates)
        bbiExec.Enabled = False
        bbiAbout.Enabled = False
        bbiOpt.Enabled = False
        bbiStop.Enabled = False
        bsiView.Enabled = False        
        bbiQuit.Enabled = False
        Select Case State
            Case ButtonlStates.Inital
                Exit Sub
            Case ButtonlStates.Start
                bbiStop.Enabled = True
                bsiView.Enabled = True
            Case ButtonlStates.StopRun
                bbiExec.Enabled = True
                bbiAbout.Enabled = True
                bbiOpt.Enabled = True                
                bsiView.Enabled = True
                bbiQuit.Enabled = True
        End Select
    End Sub
    Private Sub AddSkinName()
        Dim blnFindSkin As Boolean = False
        Try
            'DevExpress.Skins.SkinLoader.RegisterSkins(SkinLoader.DevExpressSkins.Bonus)
            'DevExpress.Skins.SkinLoader.RegisterSkins(SkinLoader.DevExpressSkins.Office)
            If String.IsNullOrEmpty(My.Settings.FormSkin) Then
                My.Settings.FormSkin = DevExpress.LookAndFeel.UserLookAndFeel.DefaultSkinName
            End If
            biSkin.ClearLinks()
            For Each objSkin As SkinContainer In DevExpress.Skins.SkinManager.Default.Skins

                Dim objBarbtn As New DevExpress.XtraBars.BarButtonItem()
                objBarbtn.Caption = objSkin.SkinName
                If objBarbtn.Caption = My.Settings.FormSkin Then
                    blnFindSkin = True
                    objBarbtn.ImageIndex = 47
                    brItem = objBarbtn
                End If
                biSkin.AddItem(objBarbtn)
                AddHandler objBarbtn.ItemClick, AddressOf brBtnClick
            Next
            If Not blnFindSkin Then
                My.Settings.FormSkin = DevExpress.LookAndFeel.UserLookAndFeel.DefaultSkinName

            End If
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub brBtnClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs)
        Try
            If brItem IsNot Nothing Then
                brItem.ImageIndex = -1
            End If
            My.Settings.FormSkin = e.Item.Caption
            My.Settings.Save()
            DefaultLookAndFeel1.LookAndFeel.SkinName = e.Item.Caption
            'DefaultLookAndFeel1.LookAndFeel.SetSkinStyle(e.Item.Caption)

            e.Item.ImageIndex = 47
            brItem = e.Item
        Catch ex As Exception
            MessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try

            '設定佈景主題
            AddSkinName()
            '設定視窗預設值
            If String.IsNullOrEmpty(My.Settings.DefLayout) Then
                Using stm As New MemoryStream()
                    dkmngr.SaveLayoutToStream(stm)
                    stm.Seek(0, SeekOrigin.Begin)
                    Using o As New StreamReader(stm)
                        My.Settings.DefLayout = o.ReadToEnd
                        My.Settings.Save()
                    End Using
                End Using
            End If
            '讀取設定值
            Common.tbSO = Common.ReadSO(IsEncryFile)
            Common.tbSystem = Common.ReadSystem(IsEncryFile)
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub rfrmMain_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            If Common.tbSO IsNot Nothing Then
                Common.tbSO.Dispose()
                Common.tbSO = Nothing
            End If
            If Common.tbSystem IsNot Nothing Then
                Common.tbSystem.Dispose()
                Common.tbSystem = Nothing
            End If
            If arySO IsNot Nothing AndAlso arySO.Count > 0 Then
                For Each SO As SO In arySO
                    If SO IsNot Nothing Then
                        SO.Dispose()
                        SO = Nothing
                    End If
                Next
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub rfrmMain_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If bbiExec.Enabled = False Then
            XtraMessageBox.Show("程式執行中！不允許關閉", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
            e.Cancel = True
            Exit Sub
        End If
        Common.CloseGateway()
    End Sub

    Private Sub rfrmMain_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ControlButtonEnable(ButtonlStates.Inital)
        Common.rfrmMain = Me
        BackgroundWorker1.RunWorkerAsync()
        Common.FLstSysMsg = New List(Of String)
    End Sub

    Private Sub bbiDefLayout_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiDefLayout.ItemClick
        Try
            If Not String.IsNullOrEmpty(My.Settings.DefLayout) Then
                Dim bty() As Byte = Encoding.UTF8.GetBytes(My.Settings.DefLayout)
                Using stm As New MemoryStream(bty)
                    dkmngr.RestoreLayoutFromStream(stm)
                End Using
            End If
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub bbiShowSO_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiShowSO.ItemClick, bbiSysMsg.ItemClick, bbiConnStatus.ItemClick
        'bbiSysMsg.ItemClick, bbiConnStatus.ItemClick
        'AddHandler bbiSysMsg.ItemClick, bbiConnStatus.ItemClick
        Try
            Dim dkPnl As DockPanel = Nothing
            Select Case e.Item.Name.ToUpper
                Case bbiShowSO.Name.ToUpper
                    dkPnl = DKpnlStatus
                    'Case bbiSysErr.Name.ToUpper
                    '    dkPnl = DKpnlSysErr
                Case bbiSysMsg.Name.ToUpper
                    dkPnl = DKpnlSysMsg
                Case bbiConnStatus.Name.ToUpper
                    dkPnl = DKpnlConnStatus
            End Select
            If dkPnl IsNot Nothing Then
                Select Case dkPnl.Visibility
                    Case DevExpress.XtraBars.Docking.DockVisibility.Hidden
                        dkPnl.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Visible
                    Case DevExpress.XtraBars.Docking.DockVisibility.Visible
                        dkPnl.Visibility = DevExpress.XtraBars.Docking.DockVisibility.AutoHide
                    Case Else
                        dkPnl.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Visible
                End Select
            End If
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub rfrmMain_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
        SuspendLayout()
        Try

            Dim LargWidt As Int32 = 0
            'If LargWidt = 0 Then LargWidt = (Width / RBPage.Groups.Count) - 12
            If LargWidt = 0 Then
                LargWidt = ((Me.Width - 40) / RBPage.Groups.Count) - (RBPage.Groups.Count * 2)
            End If

            Dim biCnt As Int32 = -1
            For Each grp As RibbonPageGroup In RBPage.Groups
                biCnt = grp.ItemLinks.Count
                For Each bil As BarItemLink In grp.ItemLinks
                    bil.Item.LargeWidth = LargWidt / biCnt
                Next
            Next
        Finally
            ResumeLayout()
        End Try

    End Sub

    Private Sub bbiQuit_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiQuit.ItemClick
        If XtraMessageBox.Show("是否要離開監控程式 ?", "詢問", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            Me.Close()
            Application.Exit()
        End If
    End Sub

    Private Sub bbiOpt_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiOpt.ItemClick
        If xfrmSetting.ShowDialog Then
            treLstSO.ClearNodes()
            TreLstConnStatus.ClearNodes()
            TreLstInstruct.ClearNodes()
            TreLstSysMsg.ClearNodes()
            BackgroundWorker1.RunWorkerAsync()
            xfrmSetting.Dispose()
        End If
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        Dim bduConnectString As New OracleClient.OracleConnectionStringBuilder
        Try
            If arySO IsNot Nothing Then
                Erase arySO
            End If    
            Dim rwAry As DataRow() = Common.tbSO.Select("IsSelect='1'")
            If rwAry IsNot Nothing AndAlso rwAry.Count > 0 Then
                Array.Resize(arySO, rwAry.Count)
                For i As Int32 = 0 To rwAry.Count - 1
                    bduConnectString.Clear()
                    bduConnectString.Pooling = True
                    bduConnectString.PersistSecurityInfo = True
                    bduConnectString.LoadBalanceTimeout = Common.tbSystem.Rows(0).Item("DBPoolLiveTime")
                    bduConnectString.MaxPoolSize = Integer.Parse("0" & Common.tbSystem.Rows(0).Item("DBMaxPool"))
                    bduConnectString.MinPoolSize = Integer.Parse("0" & Common.tbSystem.Rows(0).Item("DBMinPool"))
                    bduConnectString.DataSource = rwAry(i).Item("SID")
                    bduConnectString.UserID = rwAry(i).Item("User")
                    bduConnectString.Password = rwAry(i).Item("Password")
                    arySO(i) = New SO(rwAry(i).Item("CompCode"),
                                      rwAry(i).Item("CompName"),
                                      bduConnectString.ConnectionString,
                                      Common.tbSystem)

                    Common.UpdateShowSO(arySO(i), Common.SOStatus.Initial)

                Next

            End If
            Me.Text = Common.tbSystem.Rows(0).Item("Title")
            Common.SetSysMsgString()
            Common.UpdateSysMsgUI()
            If Boolean.Parse(Common.tbSystem.Rows(0).Item("StartGateway")) Then
                If Common.IsRegister Then
                    Common.StartGateway()
                Else
                    XtraMessageBox.Show("註冊不合法！無法啟動Gateway處理命令", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Else
                Common.CloseGateway()
            End If
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            ControlButtonEnable(ButtonlStates.StopRun)
        End Try
    End Sub

    Private Sub bbiExec_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiExec.ItemClick
        Try
            ControlButtonEnable(ButtonlStates.Start)
            If (arySO IsNot Nothing) AndAlso (arySO.Count > 0) Then
                For Each SO As SO In arySO
                    SO.IsStop = False
                    SO.Run()
                Next
            End If
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Common.UpdateSysErrUI(ex, "執行失敗")
        End Try

    End Sub

    Private Sub bbiStop_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiStop.ItemClick
        ControlButtonEnable(ButtonlStates.Inital)
        If arySO IsNot Nothing AndAlso arySO.Count > 0 Then
            treLstSO.ClearNodes()
            For Each SO As SO In arySO
                Common.UpdateShowSO(SO, Common.SOStatus.Initial)
                SO.IsStop = True
            Next
        End If

        ControlButtonEnable(ButtonlStates.StopRun)
    End Sub

    Private Sub txtCmd_Status_CustomDisplayText(sender As System.Object, e As DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs) Handles txtCmd_Status.CustomDisplayText
        If e.Value IsNot Nothing Then
            Select Case e.Value.ToString
                Case "W"
                    e.DisplayText = "待處理"
                Case "P", "P1"
                    e.DisplayText = "處理中"
                Case "T"
                    e.DisplayText = "逾時"
                Case "E", "E1", "E2"
                    e.DisplayText = "錯誤"
                Case Else
                    e.DisplayText = "成功"
            End Select
        End If

    End Sub

    Private Sub bbiAbout_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiAbout.ItemClick
        xfrmAbout.ShowDialog()
    End Sub
End Class