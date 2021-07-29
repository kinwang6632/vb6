Imports DevExpress.Skins
Imports System
Imports DevExpress.Utils
Imports DevExpress.XtraBars
Imports System.Threading
Imports System.IO
Imports System.Text
Imports System.Management
Imports DevExpress.XtraEditors.Repository
Imports System.ComponentModel
Imports System.Reflection
Imports DevExpress.XtraBars.Ribbon
Imports DevExpress.XtraEditors
Imports CableSoft.GateWay.csException
Imports DevExpress.XtraBars.Docking
Imports System.Collections
Imports System.Collections.Generic
Imports CableSoft_KeyPro
Imports CableSoft.GateWay.Common
Imports CableSoft.GateWay.SystemIO
Imports CableSoft_Nagra_BuilderCmd
Imports Raize.CodeSiteLogging
Imports System.Globalization

Public Class rfrmMain
    Private Structure POINTAPI
        Dim X As Long
        Dim Y As Long
    End Structure
    Private aMsg As String = String.Empty
    Private brItem As DevExpress.XtraBars.BarItem = Nothing
    Private brSetItem As DevExpress.XtraBars.BarItem = Nothing
    Private Shared cpu_core() As PerformanceCounter = Nothing
    Private Shared rItmPgb() As RepositoryItemProgressBar = Nothing
    Private Shared bEdtItm() As BarEditItem = Nothing
    Private Shared bStcItm() As BarStaticItem = Nothing
    Private Shared TotPhyMem As UInt64 = 0
    Private Shared AvPhyMem As UInt64 = 0
    Private Shared cptInfo As Devices.ComputerInfo = Nothing
    Private Shared blnInitial As Boolean = False
    Private ReadyGo As Boolean = False
    Private Shared cores As Int32 = -1
    Private tmr As Timer
    Private icoTray() As Icon
    Private FSetFileLst As Dictionary(Of String, String)
    Private Const FSearchFileName = "*.cfg"
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Integer
    Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Integer
    Private Declare Function GetTickCount Lib "kernel32" () As Integer
    Private Sub rfrmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try

            If Not bbiExec.Enabled Then
                XtraMessageBox.Show("GateWay正在執行中 ! 請先停止 GateWay", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                e.Cancel = True

                Exit Sub
            End If
            CableSoft.GateWay.csException.WriteErrTxtLog.WriteSysEventLog("使用者關閉 GateWay ! ")
        Catch ex As Exception

        End Try
    End Sub
    Private Sub rfrmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Thread.Sleep(100)
        fFrom = Me
        'Dim dest As CodeSiteDestination = New CodeSiteDestination

        'dest.LogFile.Active = True
        'dest.LogFile.FileName = My.Application.Info.AssemblyName   '  "MyFirstLog.csl"
        'dest.LogFile.FilePath = Environment.CurrentDirectory

        'CodeSite.Destination = dest
        'CodeSite.Clear()

        If Not String.IsNullOrEmpty(My.Settings.FormSkin) Then
            DefaultLookAndFeel1.LookAndFeel.SkinName = My.Settings.FormSkin
        End If

#If DEBUG Then
        FRegOK = True
        FInsDate = Date.Now
        FRegDate = New DateTime(2099, 12, 31)
        InitialSO()
#Else
        FRegOK = CableSoft_KeyPro.GetSystemInfo.IsRegister(ShowForm.Yes)
         If Not FRegOK Then
            XtraMessageBox.Show("註冊不合法！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
        Else

            Dim aRegInfo As String() = CableSoft_KeyPro.GetSystemInfo.DecryKeyPro.Split(Environment.NewLine)
            FInsDate = Date.Parse(aRegInfo(5), CultureInfo.CreateSpecificCulture("zh-TW")).Date
            FRegDate = Date.Parse(aRegInfo(6), CultureInfo.CreateSpecificCulture("zh-TW")).Date
            InitialSO()
        End If
#End If



       


    End Sub


    Private Sub InitialSO()
        Try


            DisableControl(RunStatus.Initial)
            If Not SetUseMDBFile() Then
                DisableControl(RunStatus.LoadComplete)
                'bbiExec.Enabled = False
                'biSets.Enabled = False
                Application.Exit()
            Else
                BackgroundWorker1.RunWorkerAsync(aMsg)
            End If

        Catch ex As Exception
            XtraMessageBox.Show(ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    '設定所選擇的設定MDB檔
    Private Sub biSetFile(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs)
        Try
            If brSetItem IsNot Nothing Then
                brSetItem.ImageIndex = -1
            End If
            My.Settings.CurrentMDB = e.Item.Caption
            e.Item.ImageIndex = 47
            brSetName.Caption = e.Item.Caption
            brSetItem = e.Item
            My.Settings.CurrentMDB = FSetFileLst(e.Item.Caption)
            My.Settings.Save()
            InitialSO()
        Catch ex As Exception

        End Try
    End Sub
    Private Function SetUseMDBFile() As Boolean
        Try
            Dim ary As ArrayList = CableSoft.GateWay.Common.csCommon.GetFindFiles(FSearchFileName)

            If ary Is Nothing Then
                XtraMessageBox.Show("找不到任何設定檔！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
                'Application.Exit()
            End If
            If FSetFileLst Is Nothing Then
                FSetFileLst = New Dictionary(Of String, String)
            End If
            FSetFileLst.Clear()
            biSets.ClearLinks()
            For i As Int32 = 0 To ary.Count - 1
                FSetFileLst.Add(Path.GetFileNameWithoutExtension(ary(i)), ary(i))
                Dim bi As New BarButtonItem()
                bi.Caption = Path.GetFileNameWithoutExtension(ary(i))
                If (Not String.IsNullOrEmpty(My.Settings.CurrentMDB)) AndAlso _
                   (My.Settings.CurrentMDB = ary(i)) Then
                    bi.ImageIndex = 47
                    brSetItem = bi
                End If
                AddHandler bi.ItemClick, AddressOf biSetFile
                biSets.AddItem(bi)

            Next
            'FMDBPassword = My.Settings("MDBPassword")
            If Not String.IsNullOrEmpty(My.Settings("CurrentMDB")) Then
                If ary.Contains(My.Settings("CurrentMDB")) Then
                    FMDBFile = My.Settings("CurrentMDB")
                Else
                    My.Settings("CurrentMDB") = ary.Item(0)
                    FMDBFile = ary.Item(0)
                    biSets.ItemLinks(0).ImageIndex = 47
                End If
            Else
                My.Settings("CurrentMDB") = ary.Item(0)
                FMDBFile = ary.Item(0)
                biSets.ItemLinks(0).ImageIndex = 47
            End If
            brSetName.Caption = Path.GetFileNameWithoutExtension(FMDBFile)
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try



            '設定環境變數
            SetSystem()
            '設定連線
            SetSO()
            SetUseComp()
            '設定錯誤對照表
            SetErrCode()
            '設定Nagra參數檔
            ReadNagraPara()
            SetSysMsgString()
            '設定Email相關參數
            ReadMailPara()
            '給予Nagra傳送Socket的初始值
            'NagraCommand.SourceId = FSndSourceId
            'NagraCommand.Dest_Id = FSndDestId
            'NagraCommand.MOP_PPID = FSndMopPPId

            'NagraCommand.IniPath = CableSoft.GateWay.Common.csCommon.GetAppPath
            'NagraCommand.ErrCodeLst = FErrCodeLst
            '設定Skin
            AddSkinName()
            '設定Form的抬頭
            UpdMainFormUI(Me)
            '設定Defaulayout
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

            '設定狀態列的顯示狀態
            If Not blnInitial Then
                blnInitial = True
                GetProcessor()
                CrtSysInfoStruct()
            End If
            ResrcDisp()

            '設定最小化的ICON
            Erase icoTray
            ReDim icoTray(2)
            For i As Int32 = 0 To icoTray.Length - 1
                Using bmp As New Bitmap(imgLstTray.Images(i))
                    icoTray(i) = Icon.FromHandle(bmp.GetHicon)
                End Using
            Next

            If tmr Is Nothing Then
                tmr = New Timer(AddressOf ShowResource, Nothing, 0, 1000)
            End If
            If (Not FUseTray) AndAlso (Not FShowResource) Then
                tmr.Change(Timeout.Infinite, Timeout.Infinite)
            Else
                tmr.Change(0, 1000)
            End If

        Catch ex As Exception
            e.Result = ex.Message
        End Try

    End Sub
    '建立狀態列資訊
    Private Sub CrtSysInfoStruct()
        Try

            If Me.InvokeRequired Then
                Dim act As New Action(AddressOf CrtSysInfoStruct)
                Me.Invoke(act)
            Else
                If cptInfo Is Nothing Then
                    cptInfo = New Devices.ComputerInfo()

                End If

                cores = Environment.ProcessorCount - 1

                ReDim cpu_core(cores)
                ReDim rItmPgb(cores)
                ReDim bEdtItm(cores)

                RBstb.SuspendLayout()

                For idx As Int32 = 0 To cores

                    cpu_core(idx) = New PerformanceCounter("Processor", "% Processor Time", idx.ToString)

                    rItmPgb(idx) = New RepositoryItemProgressBar
                    DirectCast(rItmPgb(idx), ISupportInitialize).BeginInit()
                    rItmPgb(idx).Name = String.Format("rItmPgb{0}", idx + 1)
                    rItmPgb(idx).ShowTitle = True
                    RBctl.RepositoryItems.Add(rItmPgb(idx))

                    bEdtItm(idx) = New BarEditItem
                    bEdtItm(idx).Visibility = BarItemVisibility.Never

                    bEdtItm(idx).Name = String.Format("bEdtItm{0}", idx + 1)
                    bEdtItm(idx).Caption = String.Format("核心#{0}", idx + 1)
                    bEdtItm(idx).Edit = rItmPgb(idx)
                    RBctl.Items.Add(bEdtItm(idx))
                    RBstb.ItemLinks.Add(bEdtItm(idx))

                    DirectCast(rItmPgb(idx), ISupportInitialize).EndInit()

                Next

                ReDim bStcItm(10)

                Dim capAry() As String = {"記憶體 ( 總共", "MB", ", 可用", "MB", ", 已用", "MB", _
                                          ")", "G/W使用", "MB", "線程", " "}
                For idx As Int32 = 0 To 10
                    bStcItm(idx) = New BarStaticItem
                    bStcItm(idx).Visibility = BarItemVisibility.Never
                    bStcItm(idx).Name = String.Format("bStcItm{0}", idx + 1)
                    bStcItm(idx).Caption = capAry(idx)
                    Select Case idx
                        Case 0
                            bStcItm(idx).ImageIndex = 23
                        Case 1, 3, 5, 8, 10
                            bStcItm(idx).Appearance.ForeColor = Color.Red
                        Case 7
                            bStcItm(idx).ImageIndex = 24
                        Case 9
                            bStcItm(idx).ImageIndex = 25
                        Case Else

                    End Select
                    RBctl.Items.Add(bStcItm(idx))
                    RBstb.ItemLinks.Add(bStcItm(idx))
                Next

                TotPhyMem = cptInfo.TotalPhysicalMemory
                bStcItm(1).Caption = KB2MB(TotPhyMem)

                RBstb.ResumeLayout()


                ReadyGo = True
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub
    Private Shared Function KB2MB(ByVal kb As Long) As String
        KB2MB = String.Empty
        Try
            Return String.Concat(Math.Round(kb / 2 ^ 20, 1).ToString("#,##0"), " MB")
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try

    End Function
    Private Function ComputerIsNotUse(ByVal aTime As Integer) As Boolean
        Try
            Static aMousePoint As POINTAPI
            Static aKeyStatus As Integer
            Dim i As Int32 = 0
            For i = 1 To 255
                aKeyStatus = GetAsyncKeyState(i)
                If aKeyStatus <> 0 Then aKeyStatus = 0
                GetCursorPos(aMousePoint)
                Exit For
            Next i

        Catch ex As Exception

        End Try
    End Function

    Private Sub ShowResource(ByVal state As Object)

        Try
            If FUseTray Then
                If Me.WindowState = FormWindowState.Minimized Then
                    If Not bbiExec.Enabled Then
                        If NotifyIco.Icon IsNot Nothing Then
                            NotifyIco.Icon.Dispose()
                        End If
                        If NotifyIco.Icon.Equals(icoTray(1)) Then
                            NotifyIco.Icon = icoTray(2)
                        Else
                            NotifyIco.Icon = icoTray(1)
                        End If
                    Else
                        NotifyIco.Icon = icoTray(0)
                    End If
                Else
                    If NotifyIco.Icon IsNot Nothing Then
                        NotifyIco.Icon.Dispose()
                    End If
                End If
            Else
                If NotifyIco.Icon IsNot Nothing Then
                    NotifyIco.Icon.Dispose()
                End If
            End If

            If (FShowResource) AndAlso (Me.WindowState <> FormWindowState.Minimized) Then
                UpdSysInfo()
            End If

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub

    Private Sub UpdSysInfo()
        Try
            If ReadyGo Then
                If Me.InvokeRequired Then
                    Dim act As New Action(AddressOf UpdSysInfo)
                    Me.Invoke(act)
                Else
                    SyncLock cptInfo
                        For idx As Int32 = 0 To cores
                            bEdtItm(idx).EditValue = cpu_core(idx).NextValue
                        Next
                        AvPhyMem = cptInfo.AvailablePhysicalMemory
                        bStcItm(3).Caption = KB2MB(AvPhyMem)
                        bStcItm(5).Caption = KB2MB(TotPhyMem - AvPhyMem)
                        bStcItm(8).Caption = KB2MB(Environment.WorkingSet)
                        bStcItm(10).Caption = Process.GetCurrentProcess.Threads.Count
                    End SyncLock
                End If
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub
    Private Sub ResrcDisp()
        Try
            If Me.InvokeRequired Then
                Dim act As New Action(AddressOf ResrcDisp)
                Me.Invoke(act)
            Else
                RBstb.SuspendLayout()

                If FShowResource Then
                    bsiCPU.Visibility = BarItemVisibility.Always

                    For i As Int32 = 0 To bEdtItm.Length - 1
                        bEdtItm(i).Visibility = BarItemVisibility.Always
                    Next
                    For i As Int32 = 0 To bStcItm.Length - 1
                        bStcItm(i).Visibility = BarItemVisibility.Always
                    Next



                Else
                    bsiCPU.Visibility = BarItemVisibility.Never
                    For i As Int32 = 0 To bEdtItm.Length - 1
                        bEdtItm(i).Visibility = BarItemVisibility.Never
                    Next
                    For i As Int32 = 0 To bStcItm.Length - 1
                        bStcItm(i).Visibility = BarItemVisibility.Never
                    Next

                End If

                RBctl.ResumeLayout()

            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub
    Private Sub GetProcessor()
        Try
            Static blnShow As Boolean
            If Not blnShow Then
                Dim SelQry As New SelectQuery("Win32_Processor")
                Using MngObjSch As New ManagementObjectSearcher(SelQry)
                    For Each MngObj As ManagementObject In MngObjSch.Get()
                        bliProcessor.Strings.Add(String.Format(" {0} ( MaxClock : {1} , L2Cache : {2} , Description : {3} ) ", _
                                                 MngObj("Name"), _
                                                 MngObj("MaxClockSpeed"), _
                                                 MngObj("L2CacheSize"), _
                                                 MngObj("Description")))

                    Next
                End Using
                blnShow = True
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub
    '取出所佈景主題,並設定佈景主題
    Private Sub AddSkinName()
        Dim blnFindSkin As Boolean = False
        Try
            DevExpress.Skins.SkinLoader.RegisterSkins(SkinLoader.DevExpressSkins.Bonus)
            DevExpress.Skins.SkinLoader.RegisterSkins(SkinLoader.DevExpressSkins.Office)
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
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try


    End Sub
    '選擇佈景主題並儲存佈景主題
    Private Sub brBtnClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs)
        Try
            If brItem IsNot Nothing Then
                brItem.ImageIndex = -1
            End If
            My.Settings.FormSkin = e.Item.Caption
            My.Settings.Save()
            DefaultLookAndFeel1.LookAndFeel.SkinName = e.Item.Caption
            e.Item.ImageIndex = 47
            brItem = e.Item
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try

    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        Try
            If Not String.IsNullOrEmpty(e.Result) Then
                XtraMessageBox.Show(e.Result, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
                xfrmSystemSet.ShowDialog()
                InitialSO()
            Else
                If e.Result Is Nothing Then
                    UpdateSOStatusUI(Nothing, SOStatus.Initial)

                    'UpdateSysMsgUI(TreLstSysMsg)
                    UpdateSysMsgUI()
                    DisableControl(RunStatus.LoadComplete)

                    Me.NotifyIco.Text = FAppTitle
                    Thread.Sleep(1000)
                    If FAutoRunGW Then
                        bbiExec.PerformClick()
                    End If

                End If
            End If
            'BackgroundWorker1.CancelAsync()
            'BackgroundWorker1.Dispose()
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)

        End Try

    End Sub

    Private Sub bbiExec_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiExec.ItemClick

        FThreadStop = False
        fFirstChkNagraStatus = False
        DisableControl(RunStatus.Execute)

        RunGateway()

    End Sub

    Private Sub bbiOpt_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiOpt.ItemClick
        DisableControl(RunStatus.Initial)
        If xfrmSystemSet.ShowDialog = Windows.Forms.DialogResult.OK Then
            InitialSO()
        Else
            DisableControl(RunStatus.LoadComplete)
        End If
    End Sub

    Private Sub bbiStop_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiStop.ItemClick
        Try

            FThreadStop = True
            fFirstChkNagraStatus = False
            tmrNagraStatus.Change(Timeout.Infinite, Timeout.Infinite)
            tmrNagraStatus.Dispose()
            tmrNagraStatus = Nothing
            'FNagraConnecting = False

            DisableControl(RunStatus.Initial)
            WriteErrTxtLog.WriteSysEventLog("使用者停止使用GateWay")
            '馬上把未執行完的Thread啟動
            evn.Set()
            'Interlocked.Exchange(Nagra_Socket.FUseSocket, 0)
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try

    End Sub

    Private Sub rfrmMain_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Application.Exit()
    End Sub


    Public Sub New()
        DevExpress.UserSkins.BonusSkins.Register()
        DevExpress.UserSkins.OfficeSkins.Register()
        DevExpress.Skins.SkinManager.EnableFormSkins()
        ' 此為 Windows Form 設計工具所需的呼叫。
        InitializeComponent()

        ' 在 InitializeComponent() 呼叫之後加入任何初始設定。

    End Sub

    Private Sub rfrmMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Try


            SuspendLayout()
            Dim LargWidt As Int32 = 0
            'If LargWidt = 0 Then LargWidt = (Width / RBPage.Groups.Count) - 12
            If LargWidt = 0 Then
                LargWidt = ((Me.Width - 20) / RBPage.Groups.Count) - (RBPage.Groups.Count * 2)
            End If


            Dim biCnt As Int32 = -1
            For Each grp As RibbonPageGroup In RBPage.Groups
                biCnt = grp.ItemLinks.Count
                For Each bil As BarItemLink In grp.ItemLinks
                    bil.Item.LargeWidth = LargWidt / biCnt
                Next

            Next
            ResumeLayout()
            If FUseTray Then
                If Me.WindowState = FormWindowState.Minimized Then
                    If NotifyIco Is Nothing OrElse icoTray Is Nothing Then
                        Me.Show()
                        Exit Sub
                    End If
                    NotifyIco.Icon = icoTray(0)
                    NotifyIco.Visible = True
                    Me.Hide()
                Else
                    NotifyIco.Visible = False
                    If Not Me.Visible Then
                        Me.Show()
                    End If
                End If
            End If

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub


    Private Sub mnuMain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMain.Click
        Try
            If Not Me.Visible Then
                Me.Show()
                Me.WindowState = FormWindowState.Maximized
            End If

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try

    End Sub

    Private Sub mnuExec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExec.Click
        bbiExec.PerformClick()
    End Sub

    Private Sub mnuStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStop.Click
        bbiStop.PerformClick()
    End Sub

    Private Sub mnuOpt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpt.Click
        bbiOpt.PerformClick()
    End Sub

    Private Sub NotifyIco_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIco.MouseDoubleClick

        Try
            If Not Me.Visible Then

                Me.Show()
                Me.WindowState = FormWindowState.Maximized
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try



    End Sub


    Private Sub bbiShowSO_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiShowSO.ItemClick, bbiSysErr.ItemClick, _
        bbiSysMsg.ItemClick, bbiConnStatus.ItemClick
        Try
            Dim dkPnl As DockPanel = Nothing
            Select Case e.Item.Name.ToUpper
                Case bbiShowSO.Name.ToUpper
                    dkPnl = DKpnlStatus
                Case bbiSysErr.Name.ToUpper
                    dkPnl = DKpnlSysErr
                Case bbiSysMsg.Name.ToUpper
                    dkPnl = DKpnlSysMsg
                Case bbiConnStatus.Name.ToUpper
                    dkPnl = DKpnlConnStatus
            End Select

            If dkPnl IsNot Nothing Then
                Select Case dkPnl.Visibility
                    Case Docking.DockVisibility.Hidden
                        dkPnl.Visibility = Docking.DockVisibility.Visible
                    Case Docking.DockVisibility.Visible
                        dkPnl.Visibility = Docking.DockVisibility.AutoHide
                    Case Else
                        dkPnl.Visibility = Docking.DockVisibility.Visible
                End Select
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try


    End Sub

    Private Sub bbiDefLayout_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiDefLayout.ItemClick
        Try
            If Not String.IsNullOrEmpty(My.Settings.DefLayout) Then
                Dim bty() As Byte = Encoding.UTF8.GetBytes(My.Settings.DefLayout)
                Using stm As New MemoryStream(bty)
                    dkmngr.RestoreLayoutFromStream(stm)
                End Using
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try

    End Sub

    Private Sub bbiQuit_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiQuit.ItemClick
        If XtraMessageBox.Show("是否要離開Gateway ?", "詢問", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            Me.Close()
        End If

    End Sub

    Private Sub mnuQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuQuit.Click
        bbiQuit.PerformClick()
    End Sub

    Private Sub bbiClearCmd_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiClearCmd.ItemClick
        ClearTreeLst()
    End Sub
    Private Sub ClearTreeLst()
        Try
            If Me.InvokeRequired Then
                Dim act As New Action(AddressOf ClearTreeLst)
                Me.Invoke(act)
            Else
                TreLstSysErr.ClearNodes()
                TreLstInstruct.ClearNodes()
                TreLstLowCmd.ClearNodes()
                TreLstConnStatus.ClearNodes()
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try

    End Sub

    Private Sub bbiAbout_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiAbout.ItemClick
        DisableControl(RunStatus.Initial)
        If Me.WindowState = FormWindowState.Minimized Then
            XfrmAbout.StartPosition = FormStartPosition.CenterScreen
            XfrmAbout.ShowDialog()
        Else
            XfrmAbout.StartPosition = FormStartPosition.CenterParent
            XfrmAbout.ShowDialog()
        End If
        DisableControl(RunStatus.LoadComplete)

    End Sub

    Private Sub mnuAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAbout.Click
        bbiAbout.PerformClick()
    End Sub

    'Private Sub TreLstInstruct_CustomNodeCellEdit(ByVal sender As System.Object, ByVal e As DevExpress.XtraTreeList.GetCustomNodeCellEditEventArgs) Handles TreLstInstruct.CustomNodeCellEdit
    '    Try
    '        If e.Column.FieldName.ToUpper = "VALIDITY_FLAG" Then
    '            Select Case e.Node.GetValue(e.Column.FieldName)

    '                Case "1"
    '                    e.Node.SetValue(e.Column.FieldName, "Relevant")
    '                Case "0", String.Empty
    '                    e.Node.SetValue(e.Column.FieldName, "Absolute")

    '            End Select

    '        End If
    '    Catch ex As Exception
    '        WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
    '    End Try

    'End Sub



End Class