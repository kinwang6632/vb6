<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class xfrmSetting
    Inherits DevExpress.XtraEditors.XtraForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(xfrmSetting))
        Dim SuperToolTip1 As DevExpress.Utils.SuperToolTip = New DevExpress.Utils.SuperToolTip()
        Dim ToolTipTitleItem1 As DevExpress.Utils.ToolTipTitleItem = New DevExpress.Utils.ToolTipTitleItem()
        Dim ToolTipItem1 As DevExpress.Utils.ToolTipItem = New DevExpress.Utils.ToolTipItem()
        Dim ToolTipSeparatorItem1 As DevExpress.Utils.ToolTipSeparatorItem = New DevExpress.Utils.ToolTipSeparatorItem()
        Dim ToolTipTitleItem2 As DevExpress.Utils.ToolTipTitleItem = New DevExpress.Utils.ToolTipTitleItem()
        Me.grpCtl = New DevExpress.XtraEditors.GroupControl()
        Me.btnCancel = New DevExpress.XtraEditors.SimpleButton()
        Me.btnOK = New DevExpress.XtraEditors.SimpleButton()
        Me.XtabCtl = New DevExpress.XtraTab.XtraTabControl()
        Me.XtabSysOpt = New DevExpress.XtraTab.XtraTabPage()
        Me.grpPooling = New DevExpress.XtraEditors.GroupControl()
        Me.spinLifeP = New DevExpress.XtraEditors.SpinEdit()
        Me.spinMaxP = New DevExpress.XtraEditors.SpinEdit()
        Me.spinMinP = New DevExpress.XtraEditors.SpinEdit()
        Me.lblPooling = New DevExpress.XtraEditors.LabelControl()
        Me.grpGW = New DevExpress.XtraEditors.GroupControl()
        Me.chkStartGateway = New DevExpress.XtraEditors.CheckEdit()
        Me.txtTitle = New DevExpress.XtraEditors.TextEdit()
        Me.lblTitle = New DevExpress.XtraEditors.LabelControl()
        Me.lblCore = New DevExpress.XtraEditors.LabelControl()
        Me.spinMaxThreads = New DevExpress.XtraEditors.SpinEdit()
        Me.lblMaxThreads = New DevExpress.XtraEditors.LabelControl()
        Me.spinFreq = New DevExpress.XtraEditors.SpinEdit()
        Me.spinRcd = New DevExpress.XtraEditors.SpinEdit()
        Me.spinProc = New DevExpress.XtraEditors.SpinEdit()
        Me.lblDisp = New DevExpress.XtraEditors.LabelControl()
        Me.lblDuration = New DevExpress.XtraEditors.LabelControl()
        Me.chkAutoRun = New DevExpress.XtraEditors.CheckEdit()
        Me.XtabDbOpt = New DevExpress.XtraTab.XtraTabPage()
        Me.TreLstDB = New DevExpress.XtraTreeList.TreeList()
        Me.colSelSO = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.rpstrItmChkEdt = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.colCompCode = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.rpstrItmSpnEdt = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.colCompName = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colSid = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colUserId = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colPassword = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.rpstrItmTxtEdtPwd = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.lblMsg = New DevExpress.XtraEditors.LabelControl()
        Me.btnTestConn = New DevExpress.XtraEditors.SimpleButton()
        Me.btnRemoveConn = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAddConn = New DevExpress.XtraEditors.SimpleButton()
        Me.XtabQtvWS = New DevExpress.XtraTab.XtraTabPage()
        Me.btnbtnRemoveSourceId = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAddSourceId = New DevExpress.XtraEditors.SimpleButton()
        Me.TreLstSourceId = New DevExpress.XtraTreeList.TreeList()
        Me.colSourceId = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.RepositoryItemTextEdit1 = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.colDestId = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colMoppId = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.RepositoryItemTextEdit2 = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.grpNagra = New DevExpress.XtraEditors.GroupControl()
        Me.chkErrSndMail = New DevExpress.XtraEditors.CheckEdit()
        Me.spinResumeCnt = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl34 = New DevExpress.XtraEditors.LabelControl()
        Me.spinMaxProcing = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl21 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl6 = New DevExpress.XtraEditors.LabelControl()
        Me.txtUPDEN = New DevExpress.XtraEditors.TextEdit()
        Me.spinSndDelay = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl5 = New DevExpress.XtraEditors.LabelControl()
        Me.spinChkStatus = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl4 = New DevExpress.XtraEditors.LabelControl()
        Me.spinRcvErrCntStop = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl18 = New DevExpress.XtraEditors.LabelControl()
        Me.spinSndErrCntStop = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl17 = New DevExpress.XtraEditors.LabelControl()
        Me.txtRcvMopPPId = New DevExpress.XtraEditors.TextEdit()
        Me.txtRcvDestId = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl15 = New DevExpress.XtraEditors.LabelControl()
        Me.txtRcvSourceId = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl14 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl13 = New DevExpress.XtraEditors.LabelControl()
        Me.txtSndMopPPId = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl12 = New DevExpress.XtraEditors.LabelControl()
        Me.txtSndDestId = New DevExpress.XtraEditors.TextEdit()
        Me.txtSndSourceId = New DevExpress.XtraEditors.TextEdit()
        Me.spinNagraPort = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl10 = New DevExpress.XtraEditors.LabelControl()
        Me.txtNagraIPAddress = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl3 = New DevExpress.XtraEditors.LabelControl()
        Me.XtabQtvAuth = New DevExpress.XtraTab.XtraTabPage()
        Me.TreLstErrCode = New DevExpress.XtraTreeList.TreeList()
        Me.colErrorCode = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colErrorName = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colResumeCmd = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.chkChoice = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.btnRemoveAuth = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAddAuth = New DevExpress.XtraEditors.SimpleButton()
        Me.XtabSMailInfo = New DevExpress.XtraTab.XtraTabPage()
        Me.btnDelSMailCC = New DevExpress.XtraEditors.SimpleButton()
        Me.btnInsSMailCC = New DevExpress.XtraEditors.SimpleButton()
        Me.btnDelSMailTo = New DevExpress.XtraEditors.SimpleButton()
        Me.btnInsSMailTo = New DevExpress.XtraEditors.SimpleButton()
        Me.GroupControl1 = New DevExpress.XtraEditors.GroupControl()
        Me.txtMailBody = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl29 = New DevExpress.XtraEditors.LabelControl()
        Me.txtMailSubject = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl28 = New DevExpress.XtraEditors.LabelControl()
        Me.txtDisplayName = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl27 = New DevExpress.XtraEditors.LabelControl()
        Me.chkEnabledSSL = New DevExpress.XtraEditors.CheckEdit()
        Me.txtSMTPLoginPsw = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl26 = New DevExpress.XtraEditors.LabelControl()
        Me.txtSMTPLoginID = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl25 = New DevExpress.XtraEditors.LabelControl()
        Me.txtSender = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl24 = New DevExpress.XtraEditors.LabelControl()
        Me.spinSMTPPort = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl23 = New DevExpress.XtraEditors.LabelControl()
        Me.txtSMTPHost = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl22 = New DevExpress.XtraEditors.LabelControl()
        Me.spinSocketErrCnt = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl20 = New DevExpress.XtraEditors.LabelControl()
        Me.TextEdit2 = New DevExpress.XtraEditors.TextEdit()
        Me.TextEdit3 = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl30 = New DevExpress.XtraEditors.LabelControl()
        Me.TextEdit4 = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl31 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl32 = New DevExpress.XtraEditors.LabelControl()
        Me.TextEdit5 = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl33 = New DevExpress.XtraEditors.LabelControl()
        Me.TreLstSMailCC = New DevExpress.XtraTreeList.TreeList()
        Me.colSMailCC = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.TreLstSMailTo = New DevExpress.XtraTreeList.TreeList()
        Me.colSMailTo = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        CType(Me.grpCtl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpCtl.SuspendLayout()
        CType(Me.XtabCtl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabCtl.SuspendLayout()
        Me.XtabSysOpt.SuspendLayout()
        CType(Me.grpPooling, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpPooling.SuspendLayout()
        CType(Me.spinLifeP.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinMaxP.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinMinP.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpGW, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpGW.SuspendLayout()
        CType(Me.chkStartGateway.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTitle.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinMaxThreads.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinFreq.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinRcd.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinProc.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkAutoRun.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabDbOpt.SuspendLayout()
        CType(Me.TreLstDB, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rpstrItmChkEdt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rpstrItmSpnEdt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rpstrItmTxtEdtPwd, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabQtvWS.SuspendLayout()
        CType(Me.TreLstSourceId, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemTextEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemTextEdit2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpNagra, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpNagra.SuspendLayout()
        CType(Me.chkErrSndMail.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinResumeCnt.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinMaxProcing.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtUPDEN.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinSndDelay.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinChkStatus.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinRcvErrCntStop.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinSndErrCntStop.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtRcvMopPPId.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtRcvDestId.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtRcvSourceId.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSndMopPPId.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSndDestId.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSndSourceId.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinNagraPort.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtNagraIPAddress.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabQtvAuth.SuspendLayout()
        CType(Me.TreLstErrCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkChoice, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabSMailInfo.SuspendLayout()
        CType(Me.GroupControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupControl1.SuspendLayout()
        CType(Me.txtMailBody.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtMailSubject.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDisplayName.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkEnabledSSL.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSMTPLoginPsw.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSMTPLoginID.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSender.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinSMTPPort.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSMTPHost.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinSocketErrCnt.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TextEdit2.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TextEdit3.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TextEdit4.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TextEdit5.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TreLstSMailCC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TreLstSMailTo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpCtl
        '
        Me.grpCtl.Appearance.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpCtl.Appearance.Options.UseFont = True
        Me.grpCtl.AppearanceCaption.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpCtl.AppearanceCaption.Options.UseFont = True
        Me.grpCtl.CaptionImage = CType(resources.GetObject("grpCtl.CaptionImage"), System.Drawing.Image)
        Me.grpCtl.Controls.Add(Me.btnCancel)
        Me.grpCtl.Controls.Add(Me.btnOK)
        Me.grpCtl.Controls.Add(Me.XtabCtl)
        Me.grpCtl.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpCtl.Location = New System.Drawing.Point(0, 0)
        Me.grpCtl.Name = "grpCtl"
        Me.grpCtl.Size = New System.Drawing.Size(815, 432)
        Me.grpCtl.TabIndex = 2
        Me.grpCtl.Text = " 設 定 "
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Image = CType(resources.GetObject("btnCancel.Image"), System.Drawing.Image)
        Me.btnCancel.Location = New System.Drawing.Point(465, 385)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(77, 27)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "取 消"
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnOK.Image = CType(resources.GetObject("btnOK.Image"), System.Drawing.Image)
        Me.btnOK.Location = New System.Drawing.Point(194, 385)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(77, 27)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "確 定"
        '
        'XtabCtl
        '
        Me.XtabCtl.Appearance.Options.UseTextOptions = True
        Me.XtabCtl.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.XtabCtl.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.XtabCtl.AppearancePage.Header.Options.UseTextOptions = True
        Me.XtabCtl.AppearancePage.Header.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.XtabCtl.HeaderAutoFill = DevExpress.Utils.DefaultBoolean.[False]
        Me.XtabCtl.Location = New System.Drawing.Point(15, 36)
        Me.XtabCtl.Name = "XtabCtl"
        Me.XtabCtl.SelectedTabPage = Me.XtabSysOpt
        Me.XtabCtl.ShowHeaderFocus = DevExpress.Utils.DefaultBoolean.[False]
        Me.XtabCtl.Size = New System.Drawing.Size(782, 343)
        Me.XtabCtl.TabIndex = 0
        Me.XtabCtl.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XtabSysOpt, Me.XtabDbOpt, Me.XtabQtvWS, Me.XtabQtvAuth, Me.XtabSMailInfo})
        '
        'XtabSysOpt
        '
        Me.XtabSysOpt.Controls.Add(Me.grpPooling)
        Me.XtabSysOpt.Controls.Add(Me.grpGW)
        Me.XtabSysOpt.Image = CType(resources.GetObject("XtabSysOpt.Image"), System.Drawing.Image)
        Me.XtabSysOpt.Name = "XtabSysOpt"
        Me.XtabSysOpt.Size = New System.Drawing.Size(776, 312)
        Me.XtabSysOpt.Text = " G/W 系 統 參 數 設 定 "
        '
        'grpPooling
        '
        Me.grpPooling.Controls.Add(Me.spinLifeP)
        Me.grpPooling.Controls.Add(Me.spinMaxP)
        Me.grpPooling.Controls.Add(Me.spinMinP)
        Me.grpPooling.Controls.Add(Me.lblPooling)
        Me.grpPooling.Location = New System.Drawing.Point(30, 211)
        Me.grpPooling.Name = "grpPooling"
        Me.grpPooling.Size = New System.Drawing.Size(702, 70)
        Me.grpPooling.TabIndex = 1
        Me.grpPooling.Text = " DB 集 區 設 定 "
        '
        'spinLifeP
        '
        Me.spinLifeP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinLifeP.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinLifeP.Location = New System.Drawing.Point(419, 35)
        Me.spinLifeP.Name = "spinLifeP"
        Me.spinLifeP.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinLifeP.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinLifeP.Properties.IsFloatValue = False
        Me.spinLifeP.Properties.Mask.EditMask = "N00"
        Me.spinLifeP.Properties.MaxValue = New Decimal(New Integer() {20000, 0, 0, 0})
        Me.spinLifeP.Size = New System.Drawing.Size(106, 22)
        Me.spinLifeP.TabIndex = 3
        '
        'spinMaxP
        '
        Me.spinMaxP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinMaxP.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinMaxP.Location = New System.Drawing.Point(271, 35)
        Me.spinMaxP.Name = "spinMaxP"
        Me.spinMaxP.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinMaxP.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinMaxP.Properties.IsFloatValue = False
        Me.spinMaxP.Properties.Mask.EditMask = "N00"
        Me.spinMaxP.Properties.MaxValue = New Decimal(New Integer() {999, 0, 0, 0})
        Me.spinMaxP.Size = New System.Drawing.Size(60, 22)
        Me.spinMaxP.TabIndex = 2
        '
        'spinMinP
        '
        Me.spinMinP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinMinP.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinMinP.Location = New System.Drawing.Point(143, 35)
        Me.spinMinP.Name = "spinMinP"
        Me.spinMinP.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinMinP.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinMinP.Properties.IsFloatValue = False
        Me.spinMinP.Properties.Mask.EditMask = "N00"
        Me.spinMinP.Properties.MaxValue = New Decimal(New Integer() {200, 0, 0, 0})
        Me.spinMinP.Size = New System.Drawing.Size(60, 22)
        Me.spinMinP.TabIndex = 1
        '
        'lblPooling
        '
        Me.lblPooling.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblPooling.Location = New System.Drawing.Point(86, 38)
        Me.lblPooling.Name = "lblPooling"
        Me.lblPooling.Size = New System.Drawing.Size(323, 14)
        Me.lblPooling.TabIndex = 0
        Me.lblPooling.Text = "最小集區                    最大集區                    集區Lifetime"
        Me.lblPooling.ToolTip = "資料庫連線集區 (Connection Pool) 可設定大於執行緒 (Max Threads)"
        '
        'grpGW
        '
        Me.grpGW.Controls.Add(Me.chkStartGateway)
        Me.grpGW.Controls.Add(Me.txtTitle)
        Me.grpGW.Controls.Add(Me.lblTitle)
        Me.grpGW.Controls.Add(Me.lblCore)
        Me.grpGW.Controls.Add(Me.spinMaxThreads)
        Me.grpGW.Controls.Add(Me.lblMaxThreads)
        Me.grpGW.Controls.Add(Me.spinFreq)
        Me.grpGW.Controls.Add(Me.spinRcd)
        Me.grpGW.Controls.Add(Me.spinProc)
        Me.grpGW.Controls.Add(Me.lblDisp)
        Me.grpGW.Controls.Add(Me.lblDuration)
        Me.grpGW.Controls.Add(Me.chkAutoRun)
        Me.grpGW.Location = New System.Drawing.Point(30, 14)
        Me.grpGW.Name = "grpGW"
        Me.grpGW.Size = New System.Drawing.Size(702, 182)
        Me.grpGW.TabIndex = 0
        Me.grpGW.Text = " G / W 設 定 "
        '
        'chkStartGateway
        '
        Me.chkStartGateway.Location = New System.Drawing.Point(96, 156)
        Me.chkStartGateway.Name = "chkStartGateway"
        Me.chkStartGateway.Properties.Caption = "啟動 Gateway 程式    (處理高階命令，請注意是否已有其它Gateway使用中)"
        Me.chkStartGateway.Size = New System.Drawing.Size(467, 19)
        Me.chkStartGateway.TabIndex = 19
        '
        'txtTitle
        '
        Me.txtTitle.EditValue = ""
        Me.txtTitle.Location = New System.Drawing.Point(161, 26)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtTitle.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtTitle.Properties.Appearance.Options.UseBackColor = True
        Me.txtTitle.Properties.Appearance.Options.UseForeColor = True
        Me.txtTitle.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtTitle.Size = New System.Drawing.Size(429, 22)
        Me.txtTitle.TabIndex = 1
        '
        'lblTitle
        '
        Me.lblTitle.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblTitle.Location = New System.Drawing.Point(98, 32)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(48, 14)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "抬頭顯示"
        Me.lblTitle.ToolTip = "Title"
        '
        'lblCore
        '
        Me.lblCore.Appearance.ForeColor = System.Drawing.Color.Red
        Me.lblCore.Location = New System.Drawing.Point(470, 130)
        Me.lblCore.Name = "lblCore"
        Me.lblCore.Size = New System.Drawing.Size(54, 14)
        Me.lblCore.TabIndex = 18
        Me.lblCore.Text = "( N 核心 )"
        '
        'spinMaxThreads
        '
        Me.spinMaxThreads.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinMaxThreads.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinMaxThreads.Location = New System.Drawing.Point(306, 127)
        Me.spinMaxThreads.Name = "spinMaxThreads"
        Me.spinMaxThreads.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinMaxThreads.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinMaxThreads.Properties.IsFloatValue = False
        Me.spinMaxThreads.Properties.Mask.EditMask = "N00"
        Me.spinMaxThreads.Properties.MaxValue = New Decimal(New Integer() {500, 0, 0, 0})
        Me.spinMaxThreads.Size = New System.Drawing.Size(60, 22)
        Me.spinMaxThreads.TabIndex = 17
        '
        'lblMaxThreads
        '
        Me.lblMaxThreads.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblMaxThreads.Location = New System.Drawing.Point(97, 131)
        Me.lblMaxThreads.Name = "lblMaxThreads"
        Me.lblMaxThreads.Size = New System.Drawing.Size(347, 14)
        ToolTipTitleItem1.Appearance.Image = CType(resources.GetObject("resource.Image"), System.Drawing.Image)
        ToolTipTitleItem1.Appearance.Options.UseImage = True
        ToolTipTitleItem1.Image = CType(resources.GetObject("ToolTipTitleItem1.Image"), System.Drawing.Image)
        ToolTipTitleItem1.Text = "參考值"
        ToolTipItem1.LeftIndent = 6
        ToolTipItem1.Text = "建議一個核心最高設25條執行緒" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "執行緒最小值則不得小於核心數"
        ToolTipTitleItem2.LeftIndent = 6
        ToolTipTitleItem2.Text = "系統會自行偵測核心數並" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "檢核最大值設定是否合宜"
        SuperToolTip1.Items.Add(ToolTipTitleItem1)
        SuperToolTip1.Items.Add(ToolTipItem1)
        SuperToolTip1.Items.Add(ToolTipSeparatorItem1)
        SuperToolTip1.Items.Add(ToolTipTitleItem2)
        Me.lblMaxThreads.SuperTip = SuperToolTip1
        Me.lblMaxThreads.TabIndex = 16
        Me.lblMaxThreads.Text = "最大命令處理線程 (最大執行序限制)                     Max Threads"
        '
        'spinFreq
        '
        Me.spinFreq.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinFreq.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinFreq.Location = New System.Drawing.Point(194, 62)
        Me.spinFreq.Name = "spinFreq"
        Me.spinFreq.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinFreq.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinFreq.Properties.IsFloatValue = False
        Me.spinFreq.Properties.Mask.EditMask = "N00"
        Me.spinFreq.Properties.MaxValue = New Decimal(New Integer() {1200, 0, 0, 0})
        Me.spinFreq.Size = New System.Drawing.Size(49, 22)
        Me.spinFreq.TabIndex = 11
        '
        'spinRcd
        '
        Me.spinRcd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinRcd.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinRcd.Location = New System.Drawing.Point(175, 95)
        Me.spinRcd.Name = "spinRcd"
        Me.spinRcd.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinRcd.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinRcd.Properties.IsFloatValue = False
        Me.spinRcd.Properties.Mask.EditMask = "N00"
        Me.spinRcd.Properties.MaxValue = New Decimal(New Integer() {5000, 0, 0, 0})
        Me.spinRcd.Size = New System.Drawing.Size(60, 22)
        Me.spinRcd.TabIndex = 14
        '
        'spinProc
        '
        Me.spinProc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinProc.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinProc.Location = New System.Drawing.Point(432, 62)
        Me.spinProc.Name = "spinProc"
        Me.spinProc.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinProc.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinProc.Properties.IsFloatValue = False
        Me.spinProc.Properties.Mask.EditMask = "N00"
        Me.spinProc.Properties.MaxValue = New Decimal(New Integer() {500, 0, 0, 0})
        Me.spinProc.Size = New System.Drawing.Size(49, 22)
        Me.spinProc.TabIndex = 12
        '
        'lblDisp
        '
        Me.lblDisp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblDisp.Location = New System.Drawing.Point(98, 99)
        Me.lblDisp.Name = "lblDisp"
        Me.lblDisp.Size = New System.Drawing.Size(192, 14)
        Me.lblDisp.TabIndex = 13
        Me.lblDisp.Text = "畫面資料顯示                    筆訊息 "
        '
        'lblDuration
        '
        Me.lblDuration.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblDuration.Location = New System.Drawing.Point(98, 66)
        Me.lblDuration.Name = "lblDuration"
        Me.lblDuration.Size = New System.Drawing.Size(453, 14)
        Me.lblDuration.TabIndex = 10
        Me.lblDuration.Text = "一般時段  -> 每                   秒讀取系統台資料 ,  每次處理                  筆高階指令"
        '
        'chkAutoRun
        '
        Me.chkAutoRun.Location = New System.Drawing.Point(302, 97)
        Me.chkAutoRun.Name = "chkAutoRun"
        Me.chkAutoRun.Properties.Caption = " Windows 開機後自動執行 Gateway"
        Me.chkAutoRun.Size = New System.Drawing.Size(235, 19)
        Me.chkAutoRun.TabIndex = 2
        '
        'XtabDbOpt
        '
        Me.XtabDbOpt.Controls.Add(Me.TreLstDB)
        Me.XtabDbOpt.Controls.Add(Me.lblMsg)
        Me.XtabDbOpt.Controls.Add(Me.btnTestConn)
        Me.XtabDbOpt.Controls.Add(Me.btnRemoveConn)
        Me.XtabDbOpt.Controls.Add(Me.btnAddConn)
        Me.XtabDbOpt.Image = CType(resources.GetObject("XtabDbOpt.Image"), System.Drawing.Image)
        Me.XtabDbOpt.Name = "XtabDbOpt"
        Me.XtabDbOpt.Size = New System.Drawing.Size(776, 312)
        Me.XtabDbOpt.Text = " 系 統 台 DB 連 線 設 定 "
        '
        'TreLstDB
        '
        Me.TreLstDB.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colSelSO, Me.colCompCode, Me.colCompName, Me.colSid, Me.colUserId, Me.colPassword})
        Me.TreLstDB.Location = New System.Drawing.Point(14, 54)
        Me.TreLstDB.Name = "TreLstDB"
        Me.TreLstDB.OptionsPrint.UsePrintStyles = True
        Me.TreLstDB.OptionsView.ShowButtons = False
        Me.TreLstDB.OptionsView.ShowRoot = False
        Me.TreLstDB.OptionsView.ShowSummaryFooter = True
        Me.TreLstDB.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.rpstrItmChkEdt, Me.rpstrItmSpnEdt, Me.rpstrItmTxtEdtPwd})
        Me.TreLstDB.Size = New System.Drawing.Size(746, 241)
        Me.TreLstDB.TabIndex = 19
        '
        'colSelSO
        '
        Me.colSelSO.AppearanceHeader.Options.UseTextOptions = True
        Me.colSelSO.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSelSO.Caption = "啟 動 作 業"
        Me.colSelSO.ColumnEdit = Me.rpstrItmChkEdt
        Me.colSelSO.FieldName = "SelectID"
        Me.colSelSO.MinWidth = 88
        Me.colSelSO.Name = "colSelSO"
        Me.colSelSO.SummaryFooter = DevExpress.XtraTreeList.SummaryItemType.Count
        Me.colSelSO.SummaryFooterStrFormat = "共 {0:N0} 個系統台"
        Me.colSelSO.Visible = True
        Me.colSelSO.VisibleIndex = 0
        Me.colSelSO.Width = 88
        '
        'rpstrItmChkEdt
        '
        Me.rpstrItmChkEdt.AutoHeight = False
        Me.rpstrItmChkEdt.DisplayValueChecked = "1"
        Me.rpstrItmChkEdt.DisplayValueUnchecked = "0"
        Me.rpstrItmChkEdt.Name = "rpstrItmChkEdt"
        Me.rpstrItmChkEdt.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.rpstrItmChkEdt.ValueChecked = "1"
        Me.rpstrItmChkEdt.ValueUnchecked = "0"
        '
        'colCompCode
        '
        Me.colCompCode.AppearanceCell.Options.UseTextOptions = True
        Me.colCompCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompCode.AppearanceHeader.Options.UseTextOptions = True
        Me.colCompCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompCode.Caption = "系 統 台 代 碼"
        Me.colCompCode.ColumnEdit = Me.rpstrItmSpnEdt
        Me.colCompCode.FieldName = "CompID"
        Me.colCompCode.Name = "colCompCode"
        Me.colCompCode.Visible = True
        Me.colCompCode.VisibleIndex = 1
        Me.colCompCode.Width = 86
        '
        'rpstrItmSpnEdt
        '
        Me.rpstrItmSpnEdt.AllowNullInput = DevExpress.Utils.DefaultBoolean.[False]
        Me.rpstrItmSpnEdt.Appearance.Options.UseTextOptions = True
        Me.rpstrItmSpnEdt.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.rpstrItmSpnEdt.AutoHeight = False
        Me.rpstrItmSpnEdt.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.rpstrItmSpnEdt.IsFloatValue = False
        Me.rpstrItmSpnEdt.Mask.EditMask = "N00"
        Me.rpstrItmSpnEdt.MaxLength = 2
        Me.rpstrItmSpnEdt.MaxValue = New Decimal(New Integer() {30, 0, 0, 0})
        Me.rpstrItmSpnEdt.Name = "rpstrItmSpnEdt"
        '
        'colCompName
        '
        Me.colCompName.AppearanceHeader.Options.UseTextOptions = True
        Me.colCompName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompName.Caption = "系 統 台 名 稱"
        Me.colCompName.FieldName = "CompName"
        Me.colCompName.Name = "colCompName"
        Me.colCompName.UnboundType = DevExpress.XtraTreeList.Data.UnboundColumnType.[String]
        Me.colCompName.Visible = True
        Me.colCompName.VisibleIndex = 2
        Me.colCompName.Width = 290
        '
        'colSid
        '
        Me.colSid.AppearanceHeader.Options.UseTextOptions = True
        Me.colSid.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSid.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colSid.Caption = "資 料 庫 名 稱"
        Me.colSid.FieldName = "DBSid"
        Me.colSid.Name = "colSid"
        Me.colSid.UnboundType = DevExpress.XtraTreeList.Data.UnboundColumnType.[String]
        Me.colSid.Visible = True
        Me.colSid.VisibleIndex = 3
        Me.colSid.Width = 94
        '
        'colUserId
        '
        Me.colUserId.AppearanceHeader.Options.UseTextOptions = True
        Me.colUserId.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colUserId.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colUserId.Caption = "使 用 者"
        Me.colUserId.FieldName = "DBUser"
        Me.colUserId.Name = "colUserId"
        Me.colUserId.UnboundType = DevExpress.XtraTreeList.Data.UnboundColumnType.[String]
        Me.colUserId.Visible = True
        Me.colUserId.VisibleIndex = 4
        Me.colUserId.Width = 77
        '
        'colPassword
        '
        Me.colPassword.AppearanceHeader.Options.UseTextOptions = True
        Me.colPassword.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colPassword.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colPassword.Caption = "密 碼"
        Me.colPassword.FieldName = "DBPassword"
        Me.colPassword.Name = "colPassword"
        Me.colPassword.Visible = True
        Me.colPassword.VisibleIndex = 5
        Me.colPassword.Width = 93
        '
        'rpstrItmTxtEdtPwd
        '
        Me.rpstrItmTxtEdtPwd.AutoHeight = False
        Me.rpstrItmTxtEdtPwd.Name = "rpstrItmTxtEdtPwd"
        Me.rpstrItmTxtEdtPwd.NullText = "未設定密碼"
        Me.rpstrItmTxtEdtPwd.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        '
        'lblMsg
        '
        Me.lblMsg.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblMsg.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.lblMsg.Location = New System.Drawing.Point(234, 8)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.Size = New System.Drawing.Size(515, 40)
        Me.lblMsg.TabIndex = 14
        '
        'btnTestConn
        '
        Me.btnTestConn.Image = CType(resources.GetObject("btnTestConn.Image"), System.Drawing.Image)
        Me.btnTestConn.Location = New System.Drawing.Point(156, 17)
        Me.btnTestConn.Name = "btnTestConn"
        Me.btnTestConn.Size = New System.Drawing.Size(65, 26)
        Me.btnTestConn.TabIndex = 18
        Me.btnTestConn.Text = "測 試"
        '
        'btnRemoveConn
        '
        Me.btnRemoveConn.Image = CType(resources.GetObject("btnRemoveConn.Image"), System.Drawing.Image)
        Me.btnRemoveConn.Location = New System.Drawing.Point(85, 17)
        Me.btnRemoveConn.Name = "btnRemoveConn"
        Me.btnRemoveConn.Size = New System.Drawing.Size(65, 26)
        Me.btnRemoveConn.TabIndex = 17
        Me.btnRemoveConn.Text = "刪 除"
        '
        'btnAddConn
        '
        Me.btnAddConn.Image = CType(resources.GetObject("btnAddConn.Image"), System.Drawing.Image)
        Me.btnAddConn.Location = New System.Drawing.Point(14, 17)
        Me.btnAddConn.Name = "btnAddConn"
        Me.btnAddConn.Size = New System.Drawing.Size(65, 26)
        Me.btnAddConn.TabIndex = 16
        Me.btnAddConn.Text = "新 增"
        '
        'XtabQtvWS
        '
        Me.XtabQtvWS.Controls.Add(Me.btnbtnRemoveSourceId)
        Me.XtabQtvWS.Controls.Add(Me.btnAddSourceId)
        Me.XtabQtvWS.Controls.Add(Me.TreLstSourceId)
        Me.XtabQtvWS.Controls.Add(Me.grpNagra)
        Me.XtabQtvWS.Image = CType(resources.GetObject("XtabQtvWS.Image"), System.Drawing.Image)
        Me.XtabQtvWS.Name = "XtabQtvWS"
        Me.XtabQtvWS.PageVisible = False
        Me.XtabQtvWS.Size = New System.Drawing.Size(776, 312)
        Me.XtabQtvWS.Text = "PRM 主 機 設 定 "
        '
        'btnbtnRemoveSourceId
        '
        Me.btnbtnRemoveSourceId.Image = CType(resources.GetObject("btnbtnRemoveSourceId.Image"), System.Drawing.Image)
        Me.btnbtnRemoveSourceId.Location = New System.Drawing.Point(97, 16)
        Me.btnbtnRemoveSourceId.Name = "btnbtnRemoveSourceId"
        Me.btnbtnRemoveSourceId.Size = New System.Drawing.Size(65, 26)
        Me.btnbtnRemoveSourceId.TabIndex = 25
        Me.btnbtnRemoveSourceId.Text = "刪 除"
        '
        'btnAddSourceId
        '
        Me.btnAddSourceId.Image = CType(resources.GetObject("btnAddSourceId.Image"), System.Drawing.Image)
        Me.btnAddSourceId.Location = New System.Drawing.Point(26, 16)
        Me.btnAddSourceId.Name = "btnAddSourceId"
        Me.btnAddSourceId.Size = New System.Drawing.Size(65, 26)
        Me.btnAddSourceId.TabIndex = 24
        Me.btnAddSourceId.Text = "新 增"
        '
        'TreLstSourceId
        '
        Me.TreLstSourceId.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colSourceId, Me.colDestId, Me.colMoppId})
        Me.TreLstSourceId.Location = New System.Drawing.Point(26, 48)
        Me.TreLstSourceId.Name = "TreLstSourceId"
        Me.TreLstSourceId.OptionsPrint.UsePrintStyles = True
        Me.TreLstSourceId.OptionsView.ShowButtons = False
        Me.TreLstSourceId.OptionsView.ShowRoot = False
        Me.TreLstSourceId.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.RepositoryItemTextEdit1, Me.RepositoryItemTextEdit2})
        Me.TreLstSourceId.Size = New System.Drawing.Size(702, 233)
        Me.TreLstSourceId.TabIndex = 8
        '
        'colSourceId
        '
        Me.colSourceId.Caption = "SourceId"
        Me.colSourceId.ColumnEdit = Me.RepositoryItemTextEdit1
        Me.colSourceId.FieldName = "SourceId"
        Me.colSourceId.MinWidth = 33
        Me.colSourceId.Name = "colSourceId"
        Me.colSourceId.Visible = True
        Me.colSourceId.VisibleIndex = 0
        '
        'RepositoryItemTextEdit1
        '
        Me.RepositoryItemTextEdit1.AutoHeight = False
        Me.RepositoryItemTextEdit1.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.RepositoryItemTextEdit1.Mask.EditMask = "\d{0,4}"
        Me.RepositoryItemTextEdit1.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.RegEx
        Me.RepositoryItemTextEdit1.Mask.PlaceHolder = Global.Microsoft.VisualBasic.ChrW(32)
        Me.RepositoryItemTextEdit1.MaxLength = 4
        Me.RepositoryItemTextEdit1.Name = "RepositoryItemTextEdit1"
        Me.RepositoryItemTextEdit1.NullText = "0000"
        '
        'colDestId
        '
        Me.colDestId.Caption = "Dest Id"
        Me.colDestId.ColumnEdit = Me.RepositoryItemTextEdit1
        Me.colDestId.FieldName = "DestId"
        Me.colDestId.Name = "colDestId"
        Me.colDestId.Visible = True
        Me.colDestId.VisibleIndex = 1
        '
        'colMoppId
        '
        Me.colMoppId.Caption = "Mopp Id"
        Me.colMoppId.ColumnEdit = Me.RepositoryItemTextEdit2
        Me.colMoppId.FieldName = "MoppId"
        Me.colMoppId.Name = "colMoppId"
        Me.colMoppId.Visible = True
        Me.colMoppId.VisibleIndex = 2
        '
        'RepositoryItemTextEdit2
        '
        Me.RepositoryItemTextEdit2.AutoHeight = False
        Me.RepositoryItemTextEdit2.Mask.EditMask = "\d{0,5}"
        Me.RepositoryItemTextEdit2.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.RegEx
        Me.RepositoryItemTextEdit2.MaxLength = 5
        Me.RepositoryItemTextEdit2.Name = "RepositoryItemTextEdit2"
        Me.RepositoryItemTextEdit2.NullText = "00000"
        '
        'grpNagra
        '
        Me.grpNagra.Controls.Add(Me.chkErrSndMail)
        Me.grpNagra.Controls.Add(Me.spinResumeCnt)
        Me.grpNagra.Controls.Add(Me.LabelControl34)
        Me.grpNagra.Controls.Add(Me.spinMaxProcing)
        Me.grpNagra.Controls.Add(Me.LabelControl21)
        Me.grpNagra.Controls.Add(Me.LabelControl6)
        Me.grpNagra.Controls.Add(Me.txtUPDEN)
        Me.grpNagra.Controls.Add(Me.spinSndDelay)
        Me.grpNagra.Controls.Add(Me.LabelControl5)
        Me.grpNagra.Controls.Add(Me.spinChkStatus)
        Me.grpNagra.Controls.Add(Me.LabelControl4)
        Me.grpNagra.Controls.Add(Me.spinRcvErrCntStop)
        Me.grpNagra.Controls.Add(Me.LabelControl18)
        Me.grpNagra.Controls.Add(Me.spinSndErrCntStop)
        Me.grpNagra.Controls.Add(Me.LabelControl17)
        Me.grpNagra.Controls.Add(Me.txtRcvMopPPId)
        Me.grpNagra.Controls.Add(Me.txtRcvDestId)
        Me.grpNagra.Controls.Add(Me.LabelControl15)
        Me.grpNagra.Controls.Add(Me.txtRcvSourceId)
        Me.grpNagra.Controls.Add(Me.LabelControl14)
        Me.grpNagra.Controls.Add(Me.LabelControl13)
        Me.grpNagra.Controls.Add(Me.txtSndMopPPId)
        Me.grpNagra.Controls.Add(Me.LabelControl12)
        Me.grpNagra.Controls.Add(Me.txtSndDestId)
        Me.grpNagra.Controls.Add(Me.txtSndSourceId)
        Me.grpNagra.Controls.Add(Me.spinNagraPort)
        Me.grpNagra.Controls.Add(Me.LabelControl10)
        Me.grpNagra.Controls.Add(Me.txtNagraIPAddress)
        Me.grpNagra.Controls.Add(Me.LabelControl3)
        Me.grpNagra.Location = New System.Drawing.Point(26, 287)
        Me.grpNagra.Name = "grpNagra"
        Me.grpNagra.Size = New System.Drawing.Size(702, 173)
        Me.grpNagra.TabIndex = 22
        Me.grpNagra.Text = "Nagra 設 定 "
        '
        'chkErrSndMail
        '
        Me.chkErrSndMail.Location = New System.Drawing.Point(557, 44)
        Me.chkErrSndMail.Name = "chkErrSndMail"
        Me.chkErrSndMail.Properties.Caption = "命令錯誤寄送郵件"
        Me.chkErrSndMail.Size = New System.Drawing.Size(125, 19)
        Me.chkErrSndMail.TabIndex = 52
        '
        'spinResumeCnt
        '
        Me.spinResumeCnt.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinResumeCnt.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinResumeCnt.Location = New System.Drawing.Point(453, 79)
        Me.spinResumeCnt.Name = "spinResumeCnt"
        Me.spinResumeCnt.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinResumeCnt.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinResumeCnt.Properties.IsFloatValue = False
        Me.spinResumeCnt.Properties.Mask.EditMask = "N00"
        Me.spinResumeCnt.Properties.MaxValue = New Decimal(New Integer() {999, 0, 0, 0})
        Me.spinResumeCnt.Properties.NullText = "0"
        Me.spinResumeCnt.Size = New System.Drawing.Size(49, 22)
        Me.spinResumeCnt.TabIndex = 51
        '
        'LabelControl34
        '
        Me.LabelControl34.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl34.Location = New System.Drawing.Point(346, 83)
        Me.LabelControl34.Name = "LabelControl34"
        Me.LabelControl34.Size = New System.Drawing.Size(228, 14)
        Me.LabelControl34.TabIndex = 50
        Me.LabelControl34.Text = "錯誤命令還原到達                 次,不再還原"
        '
        'spinMaxProcing
        '
        Me.spinMaxProcing.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinMaxProcing.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinMaxProcing.Location = New System.Drawing.Point(385, 41)
        Me.spinMaxProcing.Name = "spinMaxProcing"
        Me.spinMaxProcing.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinMaxProcing.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinMaxProcing.Properties.IsFloatValue = False
        Me.spinMaxProcing.Properties.Mask.EditMask = "N00"
        Me.spinMaxProcing.Properties.MaxValue = New Decimal(New Integer() {5000, 0, 0, 0})
        Me.spinMaxProcing.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinMaxProcing.Size = New System.Drawing.Size(67, 22)
        Me.spinMaxProcing.TabIndex = 5
        '
        'LabelControl21
        '
        Me.LabelControl21.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl21.Location = New System.Drawing.Point(295, 45)
        Me.LabelControl21.Name = "LabelControl21"
        Me.LabelControl21.Size = New System.Drawing.Size(256, 14)
        Me.LabelControl21.TabIndex = 49
        Me.LabelControl21.Text = "送出命令數量達                    筆, 暫時停止傳送"
        '
        'LabelControl6
        '
        Me.LabelControl6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl6.Location = New System.Drawing.Point(496, 121)
        Me.LabelControl6.Name = "LabelControl6"
        Me.LabelControl6.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl6.TabIndex = 46
        Me.LabelControl6.Text = "操作人員"
        Me.LabelControl6.Visible = False
        '
        'txtUPDEN
        '
        Me.txtUPDEN.EditValue = ""
        Me.txtUPDEN.Location = New System.Drawing.Point(550, 117)
        Me.txtUPDEN.Name = "txtUPDEN"
        Me.txtUPDEN.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtUPDEN.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtUPDEN.Properties.Appearance.Options.UseBackColor = True
        Me.txtUPDEN.Properties.Appearance.Options.UseForeColor = True
        Me.txtUPDEN.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtUPDEN.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtUPDEN.Size = New System.Drawing.Size(91, 22)
        Me.txtUPDEN.TabIndex = 2
        Me.txtUPDEN.Visible = False
        '
        'spinSndDelay
        '
        Me.spinSndDelay.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinSndDelay.EditValue = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.spinSndDelay.Location = New System.Drawing.Point(183, 150)
        Me.spinSndDelay.Name = "spinSndDelay"
        Me.spinSndDelay.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinSndDelay.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinSndDelay.Properties.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.spinSndDelay.Properties.MaxValue = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.spinSndDelay.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.spinSndDelay.Size = New System.Drawing.Size(49, 22)
        Me.spinSndDelay.TabIndex = 3
        Me.spinSndDelay.Visible = False
        '
        'LabelControl5
        '
        Me.LabelControl5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl5.Location = New System.Drawing.Point(27, 154)
        Me.LabelControl5.Name = "LabelControl5"
        Me.LabelControl5.Size = New System.Drawing.Size(220, 14)
        Me.LabelControl5.TabIndex = 43
        Me.LabelControl5.Text = "每傳送一筆低階指令後延遲                秒"
        Me.LabelControl5.Visible = False
        '
        'spinChkStatus
        '
        Me.spinChkStatus.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinChkStatus.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinChkStatus.Location = New System.Drawing.Point(47, 41)
        Me.spinChkStatus.Name = "spinChkStatus"
        Me.spinChkStatus.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinChkStatus.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinChkStatus.Properties.IsFloatValue = False
        Me.spinChkStatus.Properties.Mask.EditMask = "N00"
        Me.spinChkStatus.Properties.MaxValue = New Decimal(New Integer() {1200, 0, 0, 0})
        Me.spinChkStatus.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinChkStatus.Size = New System.Drawing.Size(49, 22)
        Me.spinChkStatus.TabIndex = 4
        '
        'LabelControl4
        '
        Me.LabelControl4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl4.Location = New System.Drawing.Point(27, 45)
        Me.LabelControl4.Name = "LabelControl4"
        Me.LabelControl4.Size = New System.Drawing.Size(262, 14)
        Me.LabelControl4.TabIndex = 41
        Me.LabelControl4.Text = "每                 秒傳送一次 Commnucation Check"
        '
        'spinRcvErrCntStop
        '
        Me.spinRcvErrCntStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinRcvErrCntStop.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinRcvErrCntStop.Location = New System.Drawing.Point(580, 150)
        Me.spinRcvErrCntStop.Name = "spinRcvErrCntStop"
        Me.spinRcvErrCntStop.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinRcvErrCntStop.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinRcvErrCntStop.Properties.IsFloatValue = False
        Me.spinRcvErrCntStop.Properties.Mask.EditMask = "N00"
        Me.spinRcvErrCntStop.Properties.MaxValue = New Decimal(New Integer() {1200, 0, 0, 0})
        Me.spinRcvErrCntStop.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinRcvErrCntStop.Size = New System.Drawing.Size(49, 22)
        Me.spinRcvErrCntStop.TabIndex = 1
        Me.spinRcvErrCntStop.Visible = False
        '
        'LabelControl18
        '
        Me.LabelControl18.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl18.Location = New System.Drawing.Point(499, 154)
        Me.LabelControl18.Name = "LabelControl18"
        Me.LabelControl18.Size = New System.Drawing.Size(240, 14)
        Me.LabelControl18.TabIndex = 39
        Me.LabelControl18.Text = "接收指令, 發生                次錯誤即停止執行"
        Me.LabelControl18.Visible = False
        '
        'spinSndErrCntStop
        '
        Me.spinSndErrCntStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinSndErrCntStop.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinSndErrCntStop.Location = New System.Drawing.Point(341, 150)
        Me.spinSndErrCntStop.Name = "spinSndErrCntStop"
        Me.spinSndErrCntStop.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinSndErrCntStop.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinSndErrCntStop.Properties.IsFloatValue = False
        Me.spinSndErrCntStop.Properties.Mask.EditMask = "N00"
        Me.spinSndErrCntStop.Properties.MaxValue = New Decimal(New Integer() {1200, 0, 0, 0})
        Me.spinSndErrCntStop.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinSndErrCntStop.Size = New System.Drawing.Size(49, 22)
        Me.spinSndErrCntStop.TabIndex = 0
        Me.spinSndErrCntStop.Visible = False
        '
        'LabelControl17
        '
        Me.LabelControl17.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl17.Location = New System.Drawing.Point(253, 154)
        Me.LabelControl17.Name = "LabelControl17"
        Me.LabelControl17.Size = New System.Drawing.Size(240, 14)
        Me.LabelControl17.TabIndex = 37
        Me.LabelControl17.Text = "傳送指令, 發生                次錯誤即停止執行"
        Me.LabelControl17.Visible = False
        '
        'txtRcvMopPPId
        '
        Me.txtRcvMopPPId.EditValue = ""
        Me.txtRcvMopPPId.Location = New System.Drawing.Point(813, 228)
        Me.txtRcvMopPPId.Name = "txtRcvMopPPId"
        Me.txtRcvMopPPId.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtRcvMopPPId.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtRcvMopPPId.Properties.Appearance.Options.UseBackColor = True
        Me.txtRcvMopPPId.Properties.Appearance.Options.UseForeColor = True
        Me.txtRcvMopPPId.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtRcvMopPPId.Properties.MaxLength = 5
        Me.txtRcvMopPPId.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtRcvMopPPId.Size = New System.Drawing.Size(78, 22)
        Me.txtRcvMopPPId.TabIndex = 36
        Me.txtRcvMopPPId.Visible = False
        '
        'txtRcvDestId
        '
        Me.txtRcvDestId.EditValue = ""
        Me.txtRcvDestId.Location = New System.Drawing.Point(649, 228)
        Me.txtRcvDestId.Name = "txtRcvDestId"
        Me.txtRcvDestId.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtRcvDestId.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtRcvDestId.Properties.Appearance.Options.UseBackColor = True
        Me.txtRcvDestId.Properties.Appearance.Options.UseForeColor = True
        Me.txtRcvDestId.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtRcvDestId.Properties.MaxLength = 4
        Me.txtRcvDestId.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtRcvDestId.Size = New System.Drawing.Size(78, 22)
        Me.txtRcvDestId.TabIndex = 34
        Me.txtRcvDestId.Visible = False
        '
        'LabelControl15
        '
        Me.LabelControl15.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl15.Location = New System.Drawing.Point(603, 174)
        Me.LabelControl15.Name = "LabelControl15"
        Me.LabelControl15.Size = New System.Drawing.Size(40, 14)
        Me.LabelControl15.TabIndex = 33
        Me.LabelControl15.Text = "Dest Id"
        Me.LabelControl15.Visible = False
        '
        'txtRcvSourceId
        '
        Me.txtRcvSourceId.EditValue = ""
        Me.txtRcvSourceId.Location = New System.Drawing.Point(496, 228)
        Me.txtRcvSourceId.Name = "txtRcvSourceId"
        Me.txtRcvSourceId.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtRcvSourceId.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtRcvSourceId.Properties.Appearance.Options.UseBackColor = True
        Me.txtRcvSourceId.Properties.Appearance.Options.UseForeColor = True
        Me.txtRcvSourceId.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtRcvSourceId.Properties.MaxLength = 4
        Me.txtRcvSourceId.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtRcvSourceId.Size = New System.Drawing.Size(78, 22)
        Me.txtRcvSourceId.TabIndex = 32
        Me.txtRcvSourceId.Visible = False
        '
        'LabelControl14
        '
        Me.LabelControl14.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl14.Location = New System.Drawing.Point(433, 174)
        Me.LabelControl14.Name = "LabelControl14"
        Me.LabelControl14.Size = New System.Drawing.Size(57, 14)
        Me.LabelControl14.TabIndex = 31
        Me.LabelControl14.Text = "Source  Id"
        Me.LabelControl14.Visible = False
        '
        'LabelControl13
        '
        Me.LabelControl13.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl13.Location = New System.Drawing.Point(368, 174)
        Me.LabelControl13.Name = "LabelControl13"
        Me.LabelControl13.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl13.TabIndex = 30
        Me.LabelControl13.Text = "回傳資料"
        Me.LabelControl13.Visible = False
        '
        'txtSndMopPPId
        '
        Me.txtSndMopPPId.EditValue = ""
        Me.txtSndMopPPId.Enabled = False
        Me.txtSndMopPPId.Location = New System.Drawing.Point(813, 198)
        Me.txtSndMopPPId.Name = "txtSndMopPPId"
        Me.txtSndMopPPId.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtSndMopPPId.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtSndMopPPId.Properties.Appearance.Options.UseBackColor = True
        Me.txtSndMopPPId.Properties.Appearance.Options.UseForeColor = True
        Me.txtSndMopPPId.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtSndMopPPId.Properties.MaxLength = 5
        Me.txtSndMopPPId.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtSndMopPPId.Size = New System.Drawing.Size(78, 22)
        Me.txtSndMopPPId.TabIndex = 29
        Me.txtSndMopPPId.Visible = False
        '
        'LabelControl12
        '
        Me.LabelControl12.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl12.Location = New System.Drawing.Point(755, 144)
        Me.LabelControl12.Name = "LabelControl12"
        Me.LabelControl12.Size = New System.Drawing.Size(52, 14)
        Me.LabelControl12.TabIndex = 28
        Me.LabelControl12.Text = "Mop PPId"
        Me.LabelControl12.Visible = False
        '
        'txtSndDestId
        '
        Me.txtSndDestId.EditValue = ""
        Me.txtSndDestId.Enabled = False
        Me.txtSndDestId.Location = New System.Drawing.Point(649, 198)
        Me.txtSndDestId.Name = "txtSndDestId"
        Me.txtSndDestId.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtSndDestId.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtSndDestId.Properties.Appearance.Options.UseBackColor = True
        Me.txtSndDestId.Properties.Appearance.Options.UseForeColor = True
        Me.txtSndDestId.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtSndDestId.Properties.MaxLength = 4
        Me.txtSndDestId.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtSndDestId.Size = New System.Drawing.Size(78, 22)
        Me.txtSndDestId.TabIndex = 27
        Me.txtSndDestId.Visible = False
        '
        'txtSndSourceId
        '
        Me.txtSndSourceId.EditValue = ""
        Me.txtSndSourceId.Enabled = False
        Me.txtSndSourceId.Location = New System.Drawing.Point(496, 198)
        Me.txtSndSourceId.Name = "txtSndSourceId"
        Me.txtSndSourceId.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtSndSourceId.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtSndSourceId.Properties.Appearance.Options.UseBackColor = True
        Me.txtSndSourceId.Properties.Appearance.Options.UseForeColor = True
        Me.txtSndSourceId.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtSndSourceId.Properties.MaxLength = 4
        Me.txtSndSourceId.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtSndSourceId.Size = New System.Drawing.Size(78, 22)
        Me.txtSndSourceId.TabIndex = 24
        Me.txtSndSourceId.Visible = False
        '
        'spinNagraPort
        '
        Me.spinNagraPort.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinNagraPort.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinNagraPort.Location = New System.Drawing.Point(268, 79)
        Me.spinNagraPort.Name = "spinNagraPort"
        Me.spinNagraPort.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinNagraPort.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinNagraPort.Properties.IsFloatValue = False
        Me.spinNagraPort.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.None
        Me.spinNagraPort.Size = New System.Drawing.Size(69, 22)
        Me.spinNagraPort.TabIndex = 7
        '
        'LabelControl10
        '
        Me.LabelControl10.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl10.Location = New System.Drawing.Point(178, 83)
        Me.LabelControl10.Name = "LabelControl10"
        Me.LabelControl10.Size = New System.Drawing.Size(84, 14)
        Me.LabelControl10.TabIndex = 21
        Me.LabelControl10.Text = "傳送指令通訊埠"
        '
        'txtNagraIPAddress
        '
        Me.txtNagraIPAddress.EditValue = ""
        Me.txtNagraIPAddress.Location = New System.Drawing.Point(47, 79)
        Me.txtNagraIPAddress.Name = "txtNagraIPAddress"
        Me.txtNagraIPAddress.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtNagraIPAddress.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtNagraIPAddress.Properties.Appearance.Options.UseBackColor = True
        Me.txtNagraIPAddress.Properties.Appearance.Options.UseForeColor = True
        Me.txtNagraIPAddress.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtNagraIPAddress.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtNagraIPAddress.Size = New System.Drawing.Size(115, 22)
        Me.txtNagraIPAddress.TabIndex = 6
        '
        'LabelControl3
        '
        Me.LabelControl3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl3.Location = New System.Drawing.Point(16, 83)
        Me.LabelControl3.Name = "LabelControl3"
        Me.LabelControl3.Size = New System.Drawing.Size(24, 14)
        Me.LabelControl3.TabIndex = 19
        Me.LabelControl3.Text = "主機"
        '
        'XtabQtvAuth
        '
        Me.XtabQtvAuth.Controls.Add(Me.TreLstErrCode)
        Me.XtabQtvAuth.Controls.Add(Me.btnRemoveAuth)
        Me.XtabQtvAuth.Controls.Add(Me.btnAddAuth)
        Me.XtabQtvAuth.Image = CType(resources.GetObject("XtabQtvAuth.Image"), System.Drawing.Image)
        Me.XtabQtvAuth.Name = "XtabQtvAuth"
        Me.XtabQtvAuth.PageVisible = False
        Me.XtabQtvAuth.Size = New System.Drawing.Size(776, 312)
        Me.XtabQtvAuth.Text = " 錯 誤 代 碼 表 設 定 "
        '
        'TreLstErrCode
        '
        Me.TreLstErrCode.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colErrorCode, Me.colErrorName, Me.colResumeCmd})
        Me.TreLstErrCode.Location = New System.Drawing.Point(14, 54)
        Me.TreLstErrCode.Name = "TreLstErrCode"
        Me.TreLstErrCode.OptionsPrint.UsePrintStyles = True
        Me.TreLstErrCode.OptionsView.ShowButtons = False
        Me.TreLstErrCode.OptionsView.ShowRoot = False
        Me.TreLstErrCode.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.chkChoice})
        Me.TreLstErrCode.Size = New System.Drawing.Size(735, 410)
        Me.TreLstErrCode.TabIndex = 22
        '
        'colErrorCode
        '
        Me.colErrorCode.Caption = "錯誤代碼"
        Me.colErrorCode.FieldName = "ErrorCode"
        Me.colErrorCode.Name = "colErrorCode"
        Me.colErrorCode.Visible = True
        Me.colErrorCode.VisibleIndex = 1
        Me.colErrorCode.Width = 65
        '
        'colErrorName
        '
        Me.colErrorName.Caption = "錯誤名稱"
        Me.colErrorName.FieldName = "ErrorName"
        Me.colErrorName.Name = "colErrorName"
        Me.colErrorName.Visible = True
        Me.colErrorName.VisibleIndex = 2
        Me.colErrorName.Width = 550
        '
        'colResumeCmd
        '
        Me.colResumeCmd.AppearanceCell.Options.UseTextOptions = True
        Me.colResumeCmd.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colResumeCmd.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colResumeCmd.AppearanceHeader.Options.UseTextOptions = True
        Me.colResumeCmd.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colResumeCmd.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colResumeCmd.Caption = "錯誤還原"
        Me.colResumeCmd.ColumnEdit = Me.chkChoice
        Me.colResumeCmd.FieldName = "Choice"
        Me.colResumeCmd.MinWidth = 33
        Me.colResumeCmd.Name = "colResumeCmd"
        Me.colResumeCmd.Visible = True
        Me.colResumeCmd.VisibleIndex = 0
        Me.colResumeCmd.Width = 99
        '
        'chkChoice
        '
        Me.chkChoice.AutoHeight = False
        Me.chkChoice.Name = "chkChoice"
        Me.chkChoice.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.chkChoice.ValueChecked = "1"
        Me.chkChoice.ValueUnchecked = "0"
        '
        'btnRemoveAuth
        '
        Me.btnRemoveAuth.Image = CType(resources.GetObject("btnRemoveAuth.Image"), System.Drawing.Image)
        Me.btnRemoveAuth.Location = New System.Drawing.Point(85, 17)
        Me.btnRemoveAuth.Name = "btnRemoveAuth"
        Me.btnRemoveAuth.Size = New System.Drawing.Size(65, 26)
        Me.btnRemoveAuth.TabIndex = 21
        Me.btnRemoveAuth.Text = "刪 除"
        '
        'btnAddAuth
        '
        Me.btnAddAuth.Image = CType(resources.GetObject("btnAddAuth.Image"), System.Drawing.Image)
        Me.btnAddAuth.Location = New System.Drawing.Point(14, 17)
        Me.btnAddAuth.Name = "btnAddAuth"
        Me.btnAddAuth.Size = New System.Drawing.Size(65, 26)
        Me.btnAddAuth.TabIndex = 20
        Me.btnAddAuth.Text = "新 增"
        '
        'XtabSMailInfo
        '
        Me.XtabSMailInfo.Controls.Add(Me.btnDelSMailCC)
        Me.XtabSMailInfo.Controls.Add(Me.btnInsSMailCC)
        Me.XtabSMailInfo.Controls.Add(Me.btnDelSMailTo)
        Me.XtabSMailInfo.Controls.Add(Me.btnInsSMailTo)
        Me.XtabSMailInfo.Controls.Add(Me.GroupControl1)
        Me.XtabSMailInfo.Controls.Add(Me.TreLstSMailCC)
        Me.XtabSMailInfo.Controls.Add(Me.TreLstSMailTo)
        Me.XtabSMailInfo.Image = CType(resources.GetObject("XtabSMailInfo.Image"), System.Drawing.Image)
        Me.XtabSMailInfo.Name = "XtabSMailInfo"
        Me.XtabSMailInfo.PageVisible = False
        Me.XtabSMailInfo.Size = New System.Drawing.Size(776, 312)
        Me.XtabSMailInfo.Text = " 發 送 Mail  警 告 設 定 "
        '
        'btnDelSMailCC
        '
        Me.btnDelSMailCC.Image = CType(resources.GetObject("btnDelSMailCC.Image"), System.Drawing.Image)
        Me.btnDelSMailCC.Location = New System.Drawing.Point(453, 17)
        Me.btnDelSMailCC.Name = "btnDelSMailCC"
        Me.btnDelSMailCC.Size = New System.Drawing.Size(65, 26)
        Me.btnDelSMailCC.TabIndex = 29
        Me.btnDelSMailCC.Text = "刪 除"
        '
        'btnInsSMailCC
        '
        Me.btnInsSMailCC.Image = CType(resources.GetObject("btnInsSMailCC.Image"), System.Drawing.Image)
        Me.btnInsSMailCC.Location = New System.Drawing.Point(382, 17)
        Me.btnInsSMailCC.Name = "btnInsSMailCC"
        Me.btnInsSMailCC.Size = New System.Drawing.Size(65, 26)
        Me.btnInsSMailCC.TabIndex = 28
        Me.btnInsSMailCC.Text = "新 增"
        '
        'btnDelSMailTo
        '
        Me.btnDelSMailTo.Image = CType(resources.GetObject("btnDelSMailTo.Image"), System.Drawing.Image)
        Me.btnDelSMailTo.Location = New System.Drawing.Point(86, 17)
        Me.btnDelSMailTo.Name = "btnDelSMailTo"
        Me.btnDelSMailTo.Size = New System.Drawing.Size(65, 26)
        Me.btnDelSMailTo.TabIndex = 27
        Me.btnDelSMailTo.Text = "刪 除"
        '
        'btnInsSMailTo
        '
        Me.btnInsSMailTo.Image = CType(resources.GetObject("btnInsSMailTo.Image"), System.Drawing.Image)
        Me.btnInsSMailTo.Location = New System.Drawing.Point(15, 17)
        Me.btnInsSMailTo.Name = "btnInsSMailTo"
        Me.btnInsSMailTo.Size = New System.Drawing.Size(65, 26)
        Me.btnInsSMailTo.TabIndex = 26
        Me.btnInsSMailTo.Text = "新 增"
        '
        'GroupControl1
        '
        Me.GroupControl1.Controls.Add(Me.txtMailBody)
        Me.GroupControl1.Controls.Add(Me.LabelControl29)
        Me.GroupControl1.Controls.Add(Me.txtMailSubject)
        Me.GroupControl1.Controls.Add(Me.LabelControl28)
        Me.GroupControl1.Controls.Add(Me.txtDisplayName)
        Me.GroupControl1.Controls.Add(Me.LabelControl27)
        Me.GroupControl1.Controls.Add(Me.chkEnabledSSL)
        Me.GroupControl1.Controls.Add(Me.txtSMTPLoginPsw)
        Me.GroupControl1.Controls.Add(Me.LabelControl26)
        Me.GroupControl1.Controls.Add(Me.txtSMTPLoginID)
        Me.GroupControl1.Controls.Add(Me.LabelControl25)
        Me.GroupControl1.Controls.Add(Me.txtSender)
        Me.GroupControl1.Controls.Add(Me.LabelControl24)
        Me.GroupControl1.Controls.Add(Me.spinSMTPPort)
        Me.GroupControl1.Controls.Add(Me.LabelControl23)
        Me.GroupControl1.Controls.Add(Me.txtSMTPHost)
        Me.GroupControl1.Controls.Add(Me.LabelControl22)
        Me.GroupControl1.Controls.Add(Me.spinSocketErrCnt)
        Me.GroupControl1.Controls.Add(Me.LabelControl20)
        Me.GroupControl1.Controls.Add(Me.TextEdit2)
        Me.GroupControl1.Controls.Add(Me.TextEdit3)
        Me.GroupControl1.Controls.Add(Me.LabelControl30)
        Me.GroupControl1.Controls.Add(Me.TextEdit4)
        Me.GroupControl1.Controls.Add(Me.LabelControl31)
        Me.GroupControl1.Controls.Add(Me.LabelControl32)
        Me.GroupControl1.Controls.Add(Me.TextEdit5)
        Me.GroupControl1.Controls.Add(Me.LabelControl33)
        Me.GroupControl1.Location = New System.Drawing.Point(14, 258)
        Me.GroupControl1.Name = "GroupControl1"
        Me.GroupControl1.Size = New System.Drawing.Size(732, 198)
        Me.GroupControl1.TabIndex = 25
        Me.GroupControl1.Text = "SMTP 設 定 "
        '
        'txtMailBody
        '
        Me.txtMailBody.EditValue = ""
        Me.txtMailBody.Location = New System.Drawing.Point(69, 131)
        Me.txtMailBody.Name = "txtMailBody"
        Me.txtMailBody.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtMailBody.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtMailBody.Properties.Appearance.Options.UseBackColor = True
        Me.txtMailBody.Properties.Appearance.Options.UseForeColor = True
        Me.txtMailBody.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtMailBody.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtMailBody.Size = New System.Drawing.Size(647, 22)
        Me.txtMailBody.TabIndex = 48
        '
        'LabelControl29
        '
        Me.LabelControl29.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl29.Location = New System.Drawing.Point(15, 135)
        Me.LabelControl29.Name = "LabelControl29"
        Me.LabelControl29.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl29.TabIndex = 56
        Me.LabelControl29.Text = "郵件內容"
        '
        'txtMailSubject
        '
        Me.txtMailSubject.EditValue = ""
        Me.txtMailSubject.Location = New System.Drawing.Point(177, 95)
        Me.txtMailSubject.Name = "txtMailSubject"
        Me.txtMailSubject.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtMailSubject.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtMailSubject.Properties.Appearance.Options.UseBackColor = True
        Me.txtMailSubject.Properties.Appearance.Options.UseForeColor = True
        Me.txtMailSubject.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtMailSubject.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtMailSubject.Size = New System.Drawing.Size(538, 22)
        Me.txtMailSubject.TabIndex = 47
        '
        'LabelControl28
        '
        Me.LabelControl28.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl28.Location = New System.Drawing.Point(123, 99)
        Me.LabelControl28.Name = "LabelControl28"
        Me.LabelControl28.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl28.TabIndex = 54
        Me.LabelControl28.Text = "郵件主旨"
        '
        'txtDisplayName
        '
        Me.txtDisplayName.EditValue = ""
        Me.txtDisplayName.Location = New System.Drawing.Point(597, 63)
        Me.txtDisplayName.Name = "txtDisplayName"
        Me.txtDisplayName.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtDisplayName.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtDisplayName.Properties.Appearance.Options.UseBackColor = True
        Me.txtDisplayName.Properties.Appearance.Options.UseForeColor = True
        Me.txtDisplayName.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtDisplayName.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtDisplayName.Size = New System.Drawing.Size(117, 22)
        Me.txtDisplayName.TabIndex = 45
        '
        'LabelControl27
        '
        Me.LabelControl27.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl27.Location = New System.Drawing.Point(543, 67)
        Me.LabelControl27.Name = "LabelControl27"
        Me.LabelControl27.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl27.TabIndex = 52
        Me.LabelControl27.Text = "顯示名稱"
        '
        'chkEnabledSSL
        '
        Me.chkEnabledSSL.Location = New System.Drawing.Point(12, 97)
        Me.chkEnabledSSL.Name = "chkEnabledSSL"
        Me.chkEnabledSSL.Properties.Caption = "啟用SSL安全性"
        Me.chkEnabledSSL.Size = New System.Drawing.Size(115, 19)
        Me.chkEnabledSSL.TabIndex = 46
        '
        'txtSMTPLoginPsw
        '
        Me.txtSMTPLoginPsw.EditValue = ""
        Me.txtSMTPLoginPsw.Location = New System.Drawing.Point(232, 63)
        Me.txtSMTPLoginPsw.Name = "txtSMTPLoginPsw"
        Me.txtSMTPLoginPsw.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtSMTPLoginPsw.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtSMTPLoginPsw.Properties.Appearance.Options.UseBackColor = True
        Me.txtSMTPLoginPsw.Properties.Appearance.Options.UseForeColor = True
        Me.txtSMTPLoginPsw.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtSMTPLoginPsw.Properties.NullValuePrompt = "請輸入密碼"
        Me.txtSMTPLoginPsw.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtSMTPLoginPsw.Properties.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtSMTPLoginPsw.Size = New System.Drawing.Size(91, 22)
        Me.txtSMTPLoginPsw.TabIndex = 43
        '
        'LabelControl26
        '
        Me.LabelControl26.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl26.Location = New System.Drawing.Point(178, 67)
        Me.LabelControl26.Name = "LabelControl26"
        Me.LabelControl26.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl26.TabIndex = 48
        Me.LabelControl26.Text = "登入密碼"
        '
        'txtSMTPLoginID
        '
        Me.txtSMTPLoginID.EditValue = ""
        Me.txtSMTPLoginID.Location = New System.Drawing.Point(80, 62)
        Me.txtSMTPLoginID.Name = "txtSMTPLoginID"
        Me.txtSMTPLoginID.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtSMTPLoginID.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtSMTPLoginID.Properties.Appearance.Options.UseBackColor = True
        Me.txtSMTPLoginID.Properties.Appearance.Options.UseForeColor = True
        Me.txtSMTPLoginID.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtSMTPLoginID.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtSMTPLoginID.Size = New System.Drawing.Size(91, 22)
        Me.txtSMTPLoginID.TabIndex = 42
        '
        'LabelControl25
        '
        Me.LabelControl25.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl25.Location = New System.Drawing.Point(15, 67)
        Me.LabelControl25.Name = "LabelControl25"
        Me.LabelControl25.Size = New System.Drawing.Size(60, 14)
        Me.LabelControl25.TabIndex = 46
        Me.LabelControl25.Text = "登入使用者"
        '
        'txtSender
        '
        Me.txtSender.EditValue = ""
        Me.txtSender.Location = New System.Drawing.Point(376, 63)
        Me.txtSender.Name = "txtSender"
        Me.txtSender.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtSender.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtSender.Properties.Appearance.Options.UseBackColor = True
        Me.txtSender.Properties.Appearance.Options.UseForeColor = True
        Me.txtSender.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtSender.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtSender.Size = New System.Drawing.Size(161, 22)
        Me.txtSender.TabIndex = 44
        '
        'LabelControl24
        '
        Me.LabelControl24.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl24.Location = New System.Drawing.Point(334, 66)
        Me.LabelControl24.Name = "LabelControl24"
        Me.LabelControl24.Size = New System.Drawing.Size(36, 14)
        Me.LabelControl24.TabIndex = 44
        Me.LabelControl24.Text = "寄件者"
        '
        'spinSMTPPort
        '
        Me.spinSMTPPort.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinSMTPPort.EditValue = New Decimal(New Integer() {25, 0, 0, 0})
        Me.spinSMTPPort.Location = New System.Drawing.Point(620, 28)
        Me.spinSMTPPort.Name = "spinSMTPPort"
        Me.spinSMTPPort.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinSMTPPort.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinSMTPPort.Properties.IsFloatValue = False
        Me.spinSMTPPort.Properties.Mask.EditMask = "N00"
        Me.spinSMTPPort.Properties.MaxValue = New Decimal(New Integer() {1200, 0, 0, 0})
        Me.spinSMTPPort.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinSMTPPort.Size = New System.Drawing.Size(56, 22)
        Me.spinSMTPPort.TabIndex = 41
        '
        'LabelControl23
        '
        Me.LabelControl23.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl23.Location = New System.Drawing.Point(578, 32)
        Me.LabelControl23.Name = "LabelControl23"
        Me.LabelControl23.Size = New System.Drawing.Size(36, 14)
        Me.LabelControl23.TabIndex = 42
        Me.LabelControl23.Text = "通訊埠"
        '
        'txtSMTPHost
        '
        Me.txtSMTPHost.EditValue = ""
        Me.txtSMTPHost.Location = New System.Drawing.Point(359, 28)
        Me.txtSMTPHost.Name = "txtSMTPHost"
        Me.txtSMTPHost.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtSMTPHost.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtSMTPHost.Properties.Appearance.Options.UseBackColor = True
        Me.txtSMTPHost.Properties.Appearance.Options.UseForeColor = True
        Me.txtSMTPHost.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtSMTPHost.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtSMTPHost.Size = New System.Drawing.Size(201, 22)
        Me.txtSMTPHost.TabIndex = 40
        '
        'LabelControl22
        '
        Me.LabelControl22.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl22.Location = New System.Drawing.Point(293, 32)
        Me.LabelControl22.Name = "LabelControl22"
        Me.LabelControl22.Size = New System.Drawing.Size(60, 14)
        Me.LabelControl22.TabIndex = 40
        Me.LabelControl22.Text = "郵件伺服器"
        '
        'spinSocketErrCnt
        '
        Me.spinSocketErrCnt.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinSocketErrCnt.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinSocketErrCnt.Location = New System.Drawing.Point(100, 28)
        Me.spinSocketErrCnt.Name = "spinSocketErrCnt"
        Me.spinSocketErrCnt.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinSocketErrCnt.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinSocketErrCnt.Properties.IsFloatValue = False
        Me.spinSocketErrCnt.Properties.Mask.EditMask = "N00"
        Me.spinSocketErrCnt.Properties.MaxValue = New Decimal(New Integer() {1200, 0, 0, 0})
        Me.spinSocketErrCnt.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinSocketErrCnt.Size = New System.Drawing.Size(49, 22)
        Me.spinSocketErrCnt.TabIndex = 39
        '
        'LabelControl20
        '
        Me.LabelControl20.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl20.Location = New System.Drawing.Point(15, 32)
        Me.LabelControl20.Name = "LabelControl20"
        Me.LabelControl20.Size = New System.Drawing.Size(261, 14)
        Me.LabelControl20.TabIndex = 38
        Me.LabelControl20.Text = "SOCKET 連線                   次失敗,寄發警告郵件"
        '
        'TextEdit2
        '
        Me.TextEdit2.EditValue = ""
        Me.TextEdit2.Location = New System.Drawing.Point(813, 228)
        Me.TextEdit2.Name = "TextEdit2"
        Me.TextEdit2.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.TextEdit2.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.TextEdit2.Properties.Appearance.Options.UseBackColor = True
        Me.TextEdit2.Properties.Appearance.Options.UseForeColor = True
        Me.TextEdit2.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.TextEdit2.Properties.MaxLength = 5
        Me.TextEdit2.Properties.NullValuePromptShowForEmptyValue = True
        Me.TextEdit2.Size = New System.Drawing.Size(78, 22)
        Me.TextEdit2.TabIndex = 36
        Me.TextEdit2.Visible = False
        '
        'TextEdit3
        '
        Me.TextEdit3.EditValue = ""
        Me.TextEdit3.Location = New System.Drawing.Point(649, 228)
        Me.TextEdit3.Name = "TextEdit3"
        Me.TextEdit3.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.TextEdit3.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.TextEdit3.Properties.Appearance.Options.UseBackColor = True
        Me.TextEdit3.Properties.Appearance.Options.UseForeColor = True
        Me.TextEdit3.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.TextEdit3.Properties.MaxLength = 4
        Me.TextEdit3.Properties.NullValuePromptShowForEmptyValue = True
        Me.TextEdit3.Size = New System.Drawing.Size(78, 22)
        Me.TextEdit3.TabIndex = 34
        Me.TextEdit3.Visible = False
        '
        'LabelControl30
        '
        Me.LabelControl30.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl30.Location = New System.Drawing.Point(603, 199)
        Me.LabelControl30.Name = "LabelControl30"
        Me.LabelControl30.Size = New System.Drawing.Size(40, 14)
        Me.LabelControl30.TabIndex = 33
        Me.LabelControl30.Text = "Dest Id"
        Me.LabelControl30.Visible = False
        '
        'TextEdit4
        '
        Me.TextEdit4.EditValue = ""
        Me.TextEdit4.Location = New System.Drawing.Point(496, 228)
        Me.TextEdit4.Name = "TextEdit4"
        Me.TextEdit4.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.TextEdit4.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.TextEdit4.Properties.Appearance.Options.UseBackColor = True
        Me.TextEdit4.Properties.Appearance.Options.UseForeColor = True
        Me.TextEdit4.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.TextEdit4.Properties.MaxLength = 4
        Me.TextEdit4.Properties.NullValuePromptShowForEmptyValue = True
        Me.TextEdit4.Size = New System.Drawing.Size(78, 22)
        Me.TextEdit4.TabIndex = 32
        Me.TextEdit4.Visible = False
        '
        'LabelControl31
        '
        Me.LabelControl31.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl31.Location = New System.Drawing.Point(433, 199)
        Me.LabelControl31.Name = "LabelControl31"
        Me.LabelControl31.Size = New System.Drawing.Size(57, 14)
        Me.LabelControl31.TabIndex = 31
        Me.LabelControl31.Text = "Source  Id"
        Me.LabelControl31.Visible = False
        '
        'LabelControl32
        '
        Me.LabelControl32.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl32.Location = New System.Drawing.Point(368, 199)
        Me.LabelControl32.Name = "LabelControl32"
        Me.LabelControl32.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl32.TabIndex = 30
        Me.LabelControl32.Text = "回傳資料"
        Me.LabelControl32.Visible = False
        '
        'TextEdit5
        '
        Me.TextEdit5.EditValue = ""
        Me.TextEdit5.Enabled = False
        Me.TextEdit5.Location = New System.Drawing.Point(813, 198)
        Me.TextEdit5.Name = "TextEdit5"
        Me.TextEdit5.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.TextEdit5.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.TextEdit5.Properties.Appearance.Options.UseBackColor = True
        Me.TextEdit5.Properties.Appearance.Options.UseForeColor = True
        Me.TextEdit5.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.TextEdit5.Properties.MaxLength = 5
        Me.TextEdit5.Properties.NullValuePromptShowForEmptyValue = True
        Me.TextEdit5.Size = New System.Drawing.Size(78, 22)
        Me.TextEdit5.TabIndex = 29
        Me.TextEdit5.Visible = False
        '
        'LabelControl33
        '
        Me.LabelControl33.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl33.Location = New System.Drawing.Point(755, 169)
        Me.LabelControl33.Name = "LabelControl33"
        Me.LabelControl33.Size = New System.Drawing.Size(52, 14)
        Me.LabelControl33.TabIndex = 28
        Me.LabelControl33.Text = "Mop PPId"
        Me.LabelControl33.Visible = False
        '
        'TreLstSMailCC
        '
        Me.TreLstSMailCC.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colSMailCC})
        Me.TreLstSMailCC.Location = New System.Drawing.Point(379, 54)
        Me.TreLstSMailCC.Name = "TreLstSMailCC"
        Me.TreLstSMailCC.OptionsPrint.UsePrintStyles = True
        Me.TreLstSMailCC.OptionsView.ShowButtons = False
        Me.TreLstSMailCC.OptionsView.ShowRoot = False
        Me.TreLstSMailCC.Size = New System.Drawing.Size(367, 186)
        Me.TreLstSMailCC.TabIndex = 24
        '
        'colSMailCC
        '
        Me.colSMailCC.AppearanceHeader.Options.UseTextOptions = True
        Me.colSMailCC.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSMailCC.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colSMailCC.Caption = "副 本"
        Me.colSMailCC.FieldName = "MailCC"
        Me.colSMailCC.MinWidth = 33
        Me.colSMailCC.Name = "colSMailCC"
        Me.colSMailCC.Visible = True
        Me.colSMailCC.VisibleIndex = 0
        Me.colSMailCC.Width = 123
        '
        'TreLstSMailTo
        '
        Me.TreLstSMailTo.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colSMailTo})
        Me.TreLstSMailTo.Location = New System.Drawing.Point(14, 54)
        Me.TreLstSMailTo.Name = "TreLstSMailTo"
        Me.TreLstSMailTo.OptionsPrint.UsePrintStyles = True
        Me.TreLstSMailTo.OptionsView.ShowButtons = False
        Me.TreLstSMailTo.OptionsView.ShowRoot = False
        Me.TreLstSMailTo.Size = New System.Drawing.Size(359, 186)
        Me.TreLstSMailTo.TabIndex = 23
        '
        'colSMailTo
        '
        Me.colSMailTo.AppearanceHeader.Options.UseTextOptions = True
        Me.colSMailTo.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSMailTo.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colSMailTo.Caption = "收 件 人 "
        Me.colSMailTo.FieldName = "MailTo"
        Me.colSMailTo.MinWidth = 33
        Me.colSMailTo.Name = "colSMailTo"
        Me.colSMailTo.Visible = True
        Me.colSMailTo.VisibleIndex = 0
        Me.colSMailTo.Width = 123
        '
        'xfrmSetting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(815, 432)
        Me.Controls.Add(Me.grpCtl)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "xfrmSetting"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "xfrmSetting"
        CType(Me.grpCtl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpCtl.ResumeLayout(False)
        CType(Me.XtabCtl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabCtl.ResumeLayout(False)
        Me.XtabSysOpt.ResumeLayout(False)
        CType(Me.grpPooling, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpPooling.ResumeLayout(False)
        Me.grpPooling.PerformLayout()
        CType(Me.spinLifeP.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinMaxP.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinMinP.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpGW, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpGW.ResumeLayout(False)
        Me.grpGW.PerformLayout()
        CType(Me.chkStartGateway.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTitle.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinMaxThreads.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinFreq.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinRcd.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinProc.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkAutoRun.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabDbOpt.ResumeLayout(False)
        CType(Me.TreLstDB, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rpstrItmChkEdt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rpstrItmSpnEdt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rpstrItmTxtEdtPwd, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabQtvWS.ResumeLayout(False)
        CType(Me.TreLstSourceId, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemTextEdit1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemTextEdit2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpNagra, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpNagra.ResumeLayout(False)
        Me.grpNagra.PerformLayout()
        CType(Me.chkErrSndMail.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinResumeCnt.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinMaxProcing.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtUPDEN.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinSndDelay.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinChkStatus.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinRcvErrCntStop.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinSndErrCntStop.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtRcvMopPPId.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtRcvDestId.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtRcvSourceId.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSndMopPPId.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSndDestId.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSndSourceId.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinNagraPort.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtNagraIPAddress.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabQtvAuth.ResumeLayout(False)
        CType(Me.TreLstErrCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkChoice, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabSMailInfo.ResumeLayout(False)
        CType(Me.GroupControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupControl1.ResumeLayout(False)
        Me.GroupControl1.PerformLayout()
        CType(Me.txtMailBody.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtMailSubject.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDisplayName.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkEnabledSSL.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSMTPLoginPsw.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSMTPLoginID.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSender.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinSMTPPort.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSMTPHost.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinSocketErrCnt.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TextEdit2.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TextEdit3.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TextEdit4.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TextEdit5.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TreLstSMailCC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TreLstSMailTo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grpCtl As DevExpress.XtraEditors.GroupControl
    Friend WithEvents XtabCtl As DevExpress.XtraTab.XtraTabControl
    Friend WithEvents XtabSysOpt As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents grpPooling As DevExpress.XtraEditors.GroupControl
    Friend WithEvents spinLifeP As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents spinMaxP As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents spinMinP As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents lblPooling As DevExpress.XtraEditors.LabelControl
    Friend WithEvents grpGW As DevExpress.XtraEditors.GroupControl
    Friend WithEvents txtTitle As DevExpress.XtraEditors.TextEdit
    Friend WithEvents lblTitle As DevExpress.XtraEditors.LabelControl
    Friend WithEvents lblCore As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinMaxThreads As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents lblMaxThreads As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinFreq As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents spinRcd As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents spinProc As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents lblDisp As DevExpress.XtraEditors.LabelControl
    Friend WithEvents lblDuration As DevExpress.XtraEditors.LabelControl
    Friend WithEvents chkAutoRun As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents XtabDbOpt As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents TreLstDB As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colSelSO As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents rpstrItmChkEdt As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents colCompCode As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents rpstrItmSpnEdt As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents colCompName As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents rpstrItmTxtEdtPwd As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents lblMsg As DevExpress.XtraEditors.LabelControl
    Friend WithEvents btnTestConn As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnRemoveConn As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAddConn As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents XtabQtvWS As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents btnbtnRemoveSourceId As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAddSourceId As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents TreLstSourceId As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colSourceId As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents RepositoryItemTextEdit1 As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents colDestId As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colMoppId As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents RepositoryItemTextEdit2 As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents grpNagra As DevExpress.XtraEditors.GroupControl
    Friend WithEvents chkErrSndMail As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents spinResumeCnt As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl34 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinMaxProcing As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl21 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl6 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtUPDEN As DevExpress.XtraEditors.TextEdit
    Friend WithEvents spinSndDelay As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl5 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinChkStatus As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl4 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinRcvErrCntStop As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl18 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinSndErrCntStop As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl17 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtRcvMopPPId As DevExpress.XtraEditors.TextEdit
    Friend WithEvents txtRcvDestId As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl15 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtRcvSourceId As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl14 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl13 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtSndMopPPId As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl12 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtSndDestId As DevExpress.XtraEditors.TextEdit
    Friend WithEvents txtSndSourceId As DevExpress.XtraEditors.TextEdit
    Friend WithEvents spinNagraPort As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl10 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtNagraIPAddress As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl3 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents XtabQtvAuth As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents TreLstErrCode As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colErrorCode As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colErrorName As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colResumeCmd As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents chkChoice As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents btnRemoveAuth As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAddAuth As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents XtabSMailInfo As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents btnDelSMailCC As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnInsSMailCC As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnDelSMailTo As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnInsSMailTo As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents GroupControl1 As DevExpress.XtraEditors.GroupControl
    Friend WithEvents txtMailBody As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl29 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtMailSubject As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl28 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtDisplayName As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl27 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents chkEnabledSSL As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents txtSMTPLoginPsw As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl26 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtSMTPLoginID As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl25 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtSender As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl24 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinSMTPPort As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl23 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtSMTPHost As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl22 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinSocketErrCnt As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl20 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents TextEdit2 As DevExpress.XtraEditors.TextEdit
    Friend WithEvents TextEdit3 As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl30 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents TextEdit4 As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl31 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl32 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents TextEdit5 As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl33 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents TreLstSMailCC As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colSMailCC As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents TreLstSMailTo As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colSMailTo As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents btnCancel As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnOK As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents colSid As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colUserId As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colPassword As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents chkStartGateway As DevExpress.XtraEditors.CheckEdit
End Class
