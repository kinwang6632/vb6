<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class xfrmSystemSet
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(xfrmSystemSet))
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
        Me.grpQtv = New DevExpress.XtraEditors.GroupControl()
        Me.txtPushUrl = New DevExpress.XtraEditors.TextEdit()
        Me.StyleController1 = New DevExpress.XtraEditors.StyleController(Me.components)
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.grpPooling = New DevExpress.XtraEditors.GroupControl()
        Me.spinLifeP = New DevExpress.XtraEditors.SpinEdit()
        Me.spinMaxP = New DevExpress.XtraEditors.SpinEdit()
        Me.spinMinP = New DevExpress.XtraEditors.SpinEdit()
        Me.lblPooling = New DevExpress.XtraEditors.LabelControl()
        Me.grpGW = New DevExpress.XtraEditors.GroupControl()
        Me.txtTitle = New DevExpress.XtraEditors.TextEdit()
        Me.lblTitle = New DevExpress.XtraEditors.LabelControl()
        Me.lblCore = New DevExpress.XtraEditors.LabelControl()
        Me.chkResource = New DevExpress.XtraEditors.CheckEdit()
        Me.ImgLst = New System.Windows.Forms.ImageList(Me.components)
        Me.spinMaxThreads = New DevExpress.XtraEditors.SpinEdit()
        Me.lblMaxThreads = New DevExpress.XtraEditors.LabelControl()
        Me.chkHotKey = New DevExpress.XtraEditors.CheckEdit()
        Me.chkTray = New DevExpress.XtraEditors.CheckEdit()
        Me.spinFreq = New DevExpress.XtraEditors.SpinEdit()
        Me.spinBuff = New DevExpress.XtraEditors.SpinEdit()
        Me.spinRcd = New DevExpress.XtraEditors.SpinEdit()
        Me.spinProc = New DevExpress.XtraEditors.SpinEdit()
        Me.lblDisp = New DevExpress.XtraEditors.LabelControl()
        Me.lblDuration = New DevExpress.XtraEditors.LabelControl()
        Me.chkAutoRun = New DevExpress.XtraEditors.CheckEdit()
        Me.XtabDbOpt = New DevExpress.XtraTab.XtraTabPage()
        Me.lblComMsg = New DevExpress.XtraEditors.LabelControl()
        Me.btnTestCom = New DevExpress.XtraEditors.SimpleButton()
        Me.grpComDB = New DevExpress.XtraEditors.GroupControl()
        Me.txtCOMPassword = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl19 = New DevExpress.XtraEditors.LabelControl()
        Me.txtCOMUserId = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl7 = New DevExpress.XtraEditors.LabelControl()
        Me.txtCOMSid = New DevExpress.XtraEditors.TextEdit()
        Me.lblComSid = New DevExpress.XtraEditors.LabelControl()
        Me.TreLstDB = New DevExpress.XtraTreeList.TreeList()
        Me.colSelSO = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.rpstrItmChkEdt = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.colCompCode = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.rpstrItmSpnEdt = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.colCompName = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colSidDBname = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colSidUserId = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colSidPassword = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.rpstrItmTxtEdtPwd = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.lblMsg = New DevExpress.XtraEditors.LabelControl()
        Me.btnTestConn = New DevExpress.XtraEditors.SimpleButton()
        Me.btnRemoveConn = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAddConn = New DevExpress.XtraEditors.SimpleButton()
        Me.XtabServerUrl = New DevExpress.XtraTab.XtraTabPage()
        Me.btnRemoveUrl = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAddUrl = New DevExpress.XtraEditors.SimpleButton()
        Me.TreLstServerUrl = New DevExpress.XtraTreeList.TreeList()
        Me.colPushServerUrl = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.grpNagra = New DevExpress.XtraEditors.GroupControl()
        Me.chkSendTVMail = New DevExpress.XtraEditors.CheckEdit()
        Me.chkErrSndMail = New DevExpress.XtraEditors.CheckEdit()
        Me.spinResumeCnt = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl34 = New DevExpress.XtraEditors.LabelControl()
        Me.spinPushSlowTime = New DevExpress.XtraEditors.SpinEdit()
        Me.spinPushRetryNum = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl3 = New DevExpress.XtraEditors.LabelControl()
        Me.cboPushUrlMethod = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.LabelControl22 = New DevExpress.XtraEditors.LabelControl()
        Me.spinPushWebTimeOut = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl21 = New DevExpress.XtraEditors.LabelControl()
        Me.spinPushRepeat = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl20 = New DevExpress.XtraEditors.LabelControl()
        Me.spinPushDuration = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl2 = New DevExpress.XtraEditors.LabelControl()
        Me.XtabTvMailMsg = New DevExpress.XtraTab.XtraTabPage()
        Me.edtTVMailMsg = New DevExpress.XtraEditors.MemoEdit()
        Me.munCustomerField = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.item1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.XtabQtvAuth = New DevExpress.XtraTab.XtraTabPage()
        Me.TreLstErrCode = New DevExpress.XtraTreeList.TreeList()
        Me.colErrorCode = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colErrorName = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colErrorRetry = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.RepositoryItemCheckEdit1 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.colErrorResume = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.btnRemoveAuth = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAddAuth = New DevExpress.XtraEditors.SimpleButton()
        Me.XtabMailInfo = New DevExpress.XtraTab.XtraTabPage()
        Me.btnDelMailCC = New DevExpress.XtraEditors.SimpleButton()
        Me.btnInsMailCC = New DevExpress.XtraEditors.SimpleButton()
        Me.TreLstMailCC = New DevExpress.XtraTreeList.TreeList()
        Me.colMailCC = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.btnDelSMailTo = New DevExpress.XtraEditors.SimpleButton()
        Me.btnInsSMailTo = New DevExpress.XtraEditors.SimpleButton()
        Me.TreLstMailTo = New DevExpress.XtraTreeList.TreeList()
        Me.colMailTo = New DevExpress.XtraTreeList.Columns.TreeListColumn()
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
        Me.LabelControl4 = New DevExpress.XtraEditors.LabelControl()
        Me.spinWebErrCnt = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl5 = New DevExpress.XtraEditors.LabelControl()
        Me.TextEdit2 = New DevExpress.XtraEditors.TextEdit()
        Me.TextEdit3 = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl30 = New DevExpress.XtraEditors.LabelControl()
        Me.TextEdit4 = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl31 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl32 = New DevExpress.XtraEditors.LabelControl()
        Me.TextEdit5 = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl33 = New DevExpress.XtraEditors.LabelControl()
        Me.rpstrItmTxtEdtPwd2 = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.chkUseRightKey = New DevExpress.XtraEditors.CheckEdit()
        CType(Me.grpCtl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpCtl.SuspendLayout()
        CType(Me.XtabCtl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabCtl.SuspendLayout()
        Me.XtabSysOpt.SuspendLayout()
        CType(Me.grpQtv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpQtv.SuspendLayout()
        CType(Me.txtPushUrl.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StyleController1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpPooling, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpPooling.SuspendLayout()
        CType(Me.spinLifeP.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinMaxP.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinMinP.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpGW, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpGW.SuspendLayout()
        CType(Me.txtTitle.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkResource.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinMaxThreads.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkHotKey.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkTray.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinFreq.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinBuff.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinRcd.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinProc.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkAutoRun.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabDbOpt.SuspendLayout()
        CType(Me.grpComDB, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpComDB.SuspendLayout()
        CType(Me.txtCOMPassword.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCOMUserId.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCOMSid.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TreLstDB, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rpstrItmChkEdt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rpstrItmSpnEdt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rpstrItmTxtEdtPwd, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabServerUrl.SuspendLayout()
        CType(Me.TreLstServerUrl, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpNagra, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpNagra.SuspendLayout()
        CType(Me.chkSendTVMail.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkErrSndMail.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinResumeCnt.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinPushSlowTime.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinPushRetryNum.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboPushUrlMethod.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinPushWebTimeOut.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinPushRepeat.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinPushDuration.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabTvMailMsg.SuspendLayout()
        CType(Me.edtTVMailMsg.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.munCustomerField.SuspendLayout()
        Me.XtabQtvAuth.SuspendLayout()
        CType(Me.TreLstErrCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabMailInfo.SuspendLayout()
        CType(Me.TreLstMailCC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TreLstMailTo, System.ComponentModel.ISupportInitialize).BeginInit()
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
        CType(Me.spinWebErrCnt.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TextEdit2.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TextEdit3.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TextEdit4.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TextEdit5.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rpstrItmTxtEdtPwd2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkUseRightKey.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.grpCtl.Size = New System.Drawing.Size(798, 602)
        Me.grpCtl.TabIndex = 1
        Me.grpCtl.Text = " 設 定 "
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Image = CType(resources.GetObject("btnCancel.Image"), System.Drawing.Image)
        Me.btnCancel.Location = New System.Drawing.Point(432, 562)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(77, 27)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "取 消"
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnOK.Image = CType(resources.GetObject("btnOK.Image"), System.Drawing.Image)
        Me.btnOK.Location = New System.Drawing.Point(287, 562)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(77, 27)
        Me.btnOK.TabIndex = 1
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
        Me.XtabCtl.Size = New System.Drawing.Size(771, 515)
        Me.XtabCtl.TabIndex = 0
        Me.XtabCtl.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XtabSysOpt, Me.XtabDbOpt, Me.XtabServerUrl, Me.XtabTvMailMsg, Me.XtabQtvAuth, Me.XtabMailInfo})
        '
        'XtabSysOpt
        '
        Me.XtabSysOpt.Controls.Add(Me.grpQtv)
        Me.XtabSysOpt.Controls.Add(Me.grpPooling)
        Me.XtabSysOpt.Controls.Add(Me.grpGW)
        Me.XtabSysOpt.Image = CType(resources.GetObject("XtabSysOpt.Image"), System.Drawing.Image)
        Me.XtabSysOpt.Name = "XtabSysOpt"
        Me.XtabSysOpt.Size = New System.Drawing.Size(764, 484)
        Me.XtabSysOpt.Text = " G/W 系 統 參 數 設 定 "
        '
        'grpQtv
        '
        Me.grpQtv.Controls.Add(Me.txtPushUrl)
        Me.grpQtv.Controls.Add(Me.LabelControl1)
        Me.grpQtv.Location = New System.Drawing.Point(30, 447)
        Me.grpQtv.Name = "grpQtv"
        Me.grpQtv.Size = New System.Drawing.Size(702, 123)
        Me.grpQtv.TabIndex = 2
        Me.grpQtv.Text = " Push Server 主 機 設 定 "
        Me.grpQtv.Visible = False
        '
        'txtPushUrl
        '
        Me.txtPushUrl.EditValue = ""
        Me.txtPushUrl.Location = New System.Drawing.Point(72, 36)
        Me.txtPushUrl.Name = "txtPushUrl"
        Me.txtPushUrl.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtPushUrl.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtPushUrl.Properties.Appearance.Options.UseBackColor = True
        Me.txtPushUrl.Properties.Appearance.Options.UseForeColor = True
        Me.txtPushUrl.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtPushUrl.Size = New System.Drawing.Size(541, 23)
        Me.txtPushUrl.StyleController = Me.StyleController1
        Me.txtPushUrl.TabIndex = 18
        '
        'StyleController1
        '
        Me.StyleController1.AppearanceFocused.BackColor = System.Drawing.Color.LemonChiffon
        Me.StyleController1.AppearanceFocused.BorderColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.StyleController1.AppearanceFocused.Options.UseBackColor = True
        Me.StyleController1.AppearanceFocused.Options.UseBorderColor = True
        Me.StyleController1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        '
        'LabelControl1
        '
        Me.LabelControl1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl1.Location = New System.Drawing.Point(18, 40)
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl1.TabIndex = 17
        Me.LabelControl1.Text = "主機位址"
        '
        'grpPooling
        '
        Me.grpPooling.Controls.Add(Me.spinLifeP)
        Me.grpPooling.Controls.Add(Me.spinMaxP)
        Me.grpPooling.Controls.Add(Me.spinMinP)
        Me.grpPooling.Controls.Add(Me.lblPooling)
        Me.grpPooling.Location = New System.Drawing.Point(30, 277)
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
        Me.spinLifeP.Size = New System.Drawing.Size(106, 23)
        Me.spinLifeP.StyleController = Me.StyleController1
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
        Me.spinMaxP.Size = New System.Drawing.Size(60, 23)
        Me.spinMaxP.StyleController = Me.StyleController1
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
        Me.spinMinP.Size = New System.Drawing.Size(60, 23)
        Me.spinMinP.StyleController = Me.StyleController1
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
        Me.grpGW.Controls.Add(Me.txtTitle)
        Me.grpGW.Controls.Add(Me.lblTitle)
        Me.grpGW.Controls.Add(Me.lblCore)
        Me.grpGW.Controls.Add(Me.chkResource)
        Me.grpGW.Controls.Add(Me.spinMaxThreads)
        Me.grpGW.Controls.Add(Me.lblMaxThreads)
        Me.grpGW.Controls.Add(Me.chkHotKey)
        Me.grpGW.Controls.Add(Me.chkTray)
        Me.grpGW.Controls.Add(Me.spinFreq)
        Me.grpGW.Controls.Add(Me.spinBuff)
        Me.grpGW.Controls.Add(Me.spinRcd)
        Me.grpGW.Controls.Add(Me.spinProc)
        Me.grpGW.Controls.Add(Me.lblDisp)
        Me.grpGW.Controls.Add(Me.lblDuration)
        Me.grpGW.Controls.Add(Me.chkAutoRun)
        Me.grpGW.Location = New System.Drawing.Point(30, 14)
        Me.grpGW.Name = "grpGW"
        Me.grpGW.Size = New System.Drawing.Size(702, 238)
        Me.grpGW.TabIndex = 0
        Me.grpGW.Text = " G / W 設 定 "
        '
        'txtTitle
        '
        Me.txtTitle.EditValue = ""
        Me.txtTitle.Location = New System.Drawing.Point(143, 36)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtTitle.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtTitle.Properties.Appearance.Options.UseBackColor = True
        Me.txtTitle.Properties.Appearance.Options.UseForeColor = True
        Me.txtTitle.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtTitle.Size = New System.Drawing.Size(429, 23)
        Me.txtTitle.StyleController = Me.StyleController1
        Me.txtTitle.TabIndex = 1
        '
        'lblTitle
        '
        Me.lblTitle.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblTitle.Location = New System.Drawing.Point(87, 40)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(48, 14)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "抬頭顯示"
        Me.lblTitle.ToolTip = "Title"
        '
        'lblCore
        '
        Me.lblCore.Appearance.ForeColor = System.Drawing.Color.Red
        Me.lblCore.Appearance.Options.UseForeColor = True
        Me.lblCore.Location = New System.Drawing.Point(438, 210)
        Me.lblCore.Name = "lblCore"
        Me.lblCore.Size = New System.Drawing.Size(54, 14)
        Me.lblCore.TabIndex = 18
        Me.lblCore.Text = "( N 核心 )"
        '
        'chkResource
        '
        Me.chkResource.Location = New System.Drawing.Point(86, 101)
        Me.chkResource.Name = "chkResource"
        Me.chkResource.Properties.Caption = " 顯示系統資源於狀態列 (Resource)"
        Me.chkResource.Properties.ImageIndexChecked = 0
        Me.chkResource.Properties.ImageIndexUnchecked = 1
        Me.chkResource.Properties.Images = Me.ImgLst
        Me.chkResource.Size = New System.Drawing.Size(229, 19)
        Me.chkResource.StyleController = Me.StyleController1
        Me.chkResource.TabIndex = 4
        '
        'ImgLst
        '
        Me.ImgLst.ImageStream = CType(resources.GetObject("ImgLst.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImgLst.TransparentColor = System.Drawing.Color.Transparent
        Me.ImgLst.Images.SetKeyName(0, "")
        Me.ImgLst.Images.SetKeyName(1, "")
        Me.ImgLst.Images.SetKeyName(2, "")
        Me.ImgLst.Images.SetKeyName(3, "")
        Me.ImgLst.Images.SetKeyName(4, "system-config-rootpassword.png")
        Me.ImgLst.Images.SetKeyName(5, "Sender.png")
        Me.ImgLst.Images.SetKeyName(6, "Sender2.png")
        '
        'spinMaxThreads
        '
        Me.spinMaxThreads.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinMaxThreads.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinMaxThreads.Location = New System.Drawing.Point(292, 206)
        Me.spinMaxThreads.Name = "spinMaxThreads"
        Me.spinMaxThreads.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinMaxThreads.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinMaxThreads.Properties.IsFloatValue = False
        Me.spinMaxThreads.Properties.Mask.EditMask = "N00"
        Me.spinMaxThreads.Properties.MaxValue = New Decimal(New Integer() {500, 0, 0, 0})
        Me.spinMaxThreads.Size = New System.Drawing.Size(60, 23)
        Me.spinMaxThreads.StyleController = Me.StyleController1
        Me.spinMaxThreads.TabIndex = 17
        '
        'lblMaxThreads
        '
        Me.lblMaxThreads.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblMaxThreads.Location = New System.Drawing.Point(85, 210)
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
        'chkHotKey
        '
        Me.chkHotKey.Enabled = False
        Me.chkHotKey.Location = New System.Drawing.Point(331, 101)
        Me.chkHotKey.Name = "chkHotKey"
        Me.chkHotKey.Properties.Caption = " 按下 [Ctrl] + [Alt] + [G] 自動呼叫開啟 Gateway"
        Me.chkHotKey.Size = New System.Drawing.Size(299, 19)
        Me.chkHotKey.StyleController = Me.StyleController1
        Me.chkHotKey.TabIndex = 5
        '
        'chkTray
        '
        Me.chkTray.Location = New System.Drawing.Point(331, 71)
        Me.chkTray.Name = "chkTray"
        Me.chkTray.Properties.Caption = " 最小化後常駐在系統工作列上Tray (TSR)"
        Me.chkTray.Size = New System.Drawing.Size(282, 19)
        Me.chkTray.StyleController = Me.StyleController1
        Me.chkTray.TabIndex = 3
        '
        'spinFreq
        '
        Me.spinFreq.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinFreq.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinFreq.Location = New System.Drawing.Point(185, 132)
        Me.spinFreq.Name = "spinFreq"
        Me.spinFreq.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinFreq.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinFreq.Properties.IsFloatValue = False
        Me.spinFreq.Properties.Mask.EditMask = "N00"
        Me.spinFreq.Properties.MaxValue = New Decimal(New Integer() {1200, 0, 0, 0})
        Me.spinFreq.Size = New System.Drawing.Size(49, 23)
        Me.spinFreq.StyleController = Me.StyleController1
        Me.spinFreq.TabIndex = 11
        '
        'spinBuff
        '
        Me.spinBuff.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinBuff.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinBuff.Location = New System.Drawing.Point(338, 169)
        Me.spinBuff.Name = "spinBuff"
        Me.spinBuff.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinBuff.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinBuff.Properties.IsFloatValue = False
        Me.spinBuff.Properties.Mask.EditMask = "N00"
        Me.spinBuff.Properties.MaxValue = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.spinBuff.Size = New System.Drawing.Size(60, 23)
        Me.spinBuff.StyleController = Me.StyleController1
        Me.spinBuff.TabIndex = 15
        '
        'spinRcd
        '
        Me.spinRcd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinRcd.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinRcd.Location = New System.Drawing.Point(174, 169)
        Me.spinRcd.Name = "spinRcd"
        Me.spinRcd.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinRcd.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinRcd.Properties.IsFloatValue = False
        Me.spinRcd.Properties.Mask.EditMask = "N00"
        Me.spinRcd.Properties.MaxValue = New Decimal(New Integer() {5000, 0, 0, 0})
        Me.spinRcd.Size = New System.Drawing.Size(60, 23)
        Me.spinRcd.StyleController = Me.StyleController1
        Me.spinRcd.TabIndex = 14
        '
        'spinProc
        '
        Me.spinProc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinProc.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinProc.Location = New System.Drawing.Point(419, 132)
        Me.spinProc.Name = "spinProc"
        Me.spinProc.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinProc.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinProc.Properties.IsFloatValue = False
        Me.spinProc.Properties.Mask.EditMask = "N00"
        Me.spinProc.Properties.MaxValue = New Decimal(New Integer() {500, 0, 0, 0})
        Me.spinProc.Size = New System.Drawing.Size(49, 23)
        Me.spinProc.StyleController = Me.StyleController1
        Me.spinProc.TabIndex = 12
        '
        'lblDisp
        '
        Me.lblDisp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblDisp.Location = New System.Drawing.Point(87, 173)
        Me.lblDisp.Name = "lblDisp"
        Me.lblDisp.Size = New System.Drawing.Size(480, 14)
        Me.lblDisp.TabIndex = 13
        Me.lblDisp.Text = "畫面資料顯示                    筆訊息 ,  每超過                    筆訊息 ,  則開始清除多餘訊息"
        '
        'lblDuration
        '
        Me.lblDuration.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblDuration.Location = New System.Drawing.Point(86, 136)
        Me.lblDuration.Name = "lblDuration"
        Me.lblDuration.Size = New System.Drawing.Size(453, 14)
        Me.lblDuration.TabIndex = 10
        Me.lblDuration.Text = "一般時段  -> 每                   秒讀取系統台資料 ,  每次處理                  筆高階指令"
        '
        'chkAutoRun
        '
        Me.chkAutoRun.Location = New System.Drawing.Point(86, 71)
        Me.chkAutoRun.Name = "chkAutoRun"
        Me.chkAutoRun.Properties.Caption = " Windows 開機後自動執行 Gateway"
        Me.chkAutoRun.Size = New System.Drawing.Size(235, 19)
        Me.chkAutoRun.StyleController = Me.StyleController1
        Me.chkAutoRun.TabIndex = 2
        '
        'XtabDbOpt
        '
        Me.XtabDbOpt.Controls.Add(Me.lblComMsg)
        Me.XtabDbOpt.Controls.Add(Me.btnTestCom)
        Me.XtabDbOpt.Controls.Add(Me.grpComDB)
        Me.XtabDbOpt.Controls.Add(Me.TreLstDB)
        Me.XtabDbOpt.Controls.Add(Me.lblMsg)
        Me.XtabDbOpt.Controls.Add(Me.btnTestConn)
        Me.XtabDbOpt.Controls.Add(Me.btnRemoveConn)
        Me.XtabDbOpt.Controls.Add(Me.btnAddConn)
        Me.XtabDbOpt.Image = CType(resources.GetObject("XtabDbOpt.Image"), System.Drawing.Image)
        Me.XtabDbOpt.Name = "XtabDbOpt"
        Me.XtabDbOpt.Size = New System.Drawing.Size(764, 484)
        Me.XtabDbOpt.Text = " 系 統 台 DB 連 線 設 定 "
        '
        'lblComMsg
        '
        Me.lblComMsg.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblComMsg.Appearance.Options.UseForeColor = True
        Me.lblComMsg.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.lblComMsg.Location = New System.Drawing.Point(85, 351)
        Me.lblComMsg.Name = "lblComMsg"
        Me.lblComMsg.Size = New System.Drawing.Size(664, 40)
        Me.lblComMsg.TabIndex = 28
        '
        'btnTestCom
        '
        Me.btnTestCom.Image = CType(resources.GetObject("btnTestCom.Image"), System.Drawing.Image)
        Me.btnTestCom.Location = New System.Drawing.Point(14, 358)
        Me.btnTestCom.Name = "btnTestCom"
        Me.btnTestCom.Size = New System.Drawing.Size(65, 26)
        Me.btnTestCom.TabIndex = 27
        Me.btnTestCom.Text = "測 試"
        '
        'grpComDB
        '
        Me.grpComDB.Controls.Add(Me.txtCOMPassword)
        Me.grpComDB.Controls.Add(Me.LabelControl19)
        Me.grpComDB.Controls.Add(Me.txtCOMUserId)
        Me.grpComDB.Controls.Add(Me.LabelControl7)
        Me.grpComDB.Controls.Add(Me.txtCOMSid)
        Me.grpComDB.Controls.Add(Me.lblComSid)
        Me.grpComDB.Location = New System.Drawing.Point(14, 396)
        Me.grpComDB.Name = "grpComDB"
        Me.grpComDB.Size = New System.Drawing.Size(735, 74)
        Me.grpComDB.TabIndex = 26
        Me.grpComDB.Text = "COM 區 設 定"
        '
        'txtCOMPassword
        '
        Me.txtCOMPassword.Location = New System.Drawing.Point(584, 37)
        Me.txtCOMPassword.Name = "txtCOMPassword"
        Me.txtCOMPassword.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtCOMPassword.Properties.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtCOMPassword.Size = New System.Drawing.Size(131, 23)
        Me.txtCOMPassword.StyleController = Me.StyleController1
        Me.txtCOMPassword.TabIndex = 31
        '
        'LabelControl19
        '
        Me.LabelControl19.Location = New System.Drawing.Point(481, 40)
        Me.LabelControl19.Name = "LabelControl19"
        Me.LabelControl19.Size = New System.Drawing.Size(97, 14)
        Me.LabelControl19.TabIndex = 30
        Me.LabelControl19.Text = "COM區資料庫密碼"
        '
        'txtCOMUserId
        '
        Me.txtCOMUserId.Location = New System.Drawing.Point(349, 37)
        Me.txtCOMUserId.Name = "txtCOMUserId"
        Me.txtCOMUserId.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtCOMUserId.Size = New System.Drawing.Size(110, 23)
        Me.txtCOMUserId.StyleController = Me.StyleController1
        Me.txtCOMUserId.TabIndex = 29
        '
        'LabelControl7
        '
        Me.LabelControl7.Location = New System.Drawing.Point(246, 40)
        Me.LabelControl7.Name = "LabelControl7"
        Me.LabelControl7.Size = New System.Drawing.Size(97, 14)
        Me.LabelControl7.TabIndex = 28
        Me.LabelControl7.Text = "COM區資料庫帳號"
        '
        'txtCOMSid
        '
        Me.txtCOMSid.Location = New System.Drawing.Point(117, 37)
        Me.txtCOMSid.Name = "txtCOMSid"
        Me.txtCOMSid.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtCOMSid.Size = New System.Drawing.Size(104, 23)
        Me.txtCOMSid.StyleController = Me.StyleController1
        Me.txtCOMSid.TabIndex = 27
        '
        'lblComSid
        '
        Me.lblComSid.Location = New System.Drawing.Point(14, 40)
        Me.lblComSid.Name = "lblComSid"
        Me.lblComSid.Size = New System.Drawing.Size(97, 14)
        Me.lblComSid.TabIndex = 26
        Me.lblComSid.Text = "COM區資料庫名稱"
        '
        'TreLstDB
        '
        Me.TreLstDB.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colSelSO, Me.colCompCode, Me.colCompName, Me.colSidDBname, Me.colSidUserId, Me.colSidPassword})
        Me.TreLstDB.Location = New System.Drawing.Point(14, 54)
        Me.TreLstDB.Name = "TreLstDB"
        Me.TreLstDB.OptionsView.AutoWidth = False
        Me.TreLstDB.OptionsView.ShowButtons = False
        Me.TreLstDB.OptionsView.ShowRoot = False
        Me.TreLstDB.OptionsView.ShowSummaryFooter = True
        Me.TreLstDB.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.rpstrItmChkEdt, Me.rpstrItmSpnEdt, Me.rpstrItmTxtEdtPwd})
        Me.TreLstDB.Size = New System.Drawing.Size(735, 291)
        Me.TreLstDB.StateImageList = Me.ImgLst
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
        Me.colCompCode.Width = 91
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
        Me.colCompName.Width = 267
        '
        'colSidDBname
        '
        Me.colSidDBname.Caption = "資料庫名稱"
        Me.colSidDBname.FieldName = "SidDBname"
        Me.colSidDBname.Name = "colSidDBname"
        Me.colSidDBname.UnboundType = DevExpress.XtraTreeList.Data.UnboundColumnType.[String]
        Me.colSidDBname.Visible = True
        Me.colSidDBname.VisibleIndex = 3
        Me.colSidDBname.Width = 114
        '
        'colSidUserId
        '
        Me.colSidUserId.Caption = "使用者帳號"
        Me.colSidUserId.FieldName = "SidUserId"
        Me.colSidUserId.Name = "colSidUserId"
        Me.colSidUserId.UnboundType = DevExpress.XtraTreeList.Data.UnboundColumnType.[String]
        Me.colSidUserId.Visible = True
        Me.colSidUserId.VisibleIndex = 4
        Me.colSidUserId.Width = 110
        '
        'colSidPassword
        '
        Me.colSidPassword.Caption = "使用者密碼"
        Me.colSidPassword.FieldName = "SidPassword"
        Me.colSidPassword.Name = "colSidPassword"
        Me.colSidPassword.UnboundType = DevExpress.XtraTreeList.Data.UnboundColumnType.[String]
        Me.colSidPassword.Visible = True
        Me.colSidPassword.VisibleIndex = 5
        Me.colSidPassword.Width = 186
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
        Me.lblMsg.Appearance.Options.UseForeColor = True
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
        'XtabServerUrl
        '
        Me.XtabServerUrl.Controls.Add(Me.btnRemoveUrl)
        Me.XtabServerUrl.Controls.Add(Me.btnAddUrl)
        Me.XtabServerUrl.Controls.Add(Me.TreLstServerUrl)
        Me.XtabServerUrl.Controls.Add(Me.grpNagra)
        Me.XtabServerUrl.Image = CType(resources.GetObject("XtabServerUrl.Image"), System.Drawing.Image)
        Me.XtabServerUrl.Name = "XtabServerUrl"
        Me.XtabServerUrl.Size = New System.Drawing.Size(764, 484)
        Me.XtabServerUrl.Text = "Push Server 主 機 設 定 "
        '
        'btnRemoveUrl
        '
        Me.btnRemoveUrl.Image = CType(resources.GetObject("btnRemoveUrl.Image"), System.Drawing.Image)
        Me.btnRemoveUrl.Location = New System.Drawing.Point(98, 19)
        Me.btnRemoveUrl.Name = "btnRemoveUrl"
        Me.btnRemoveUrl.Size = New System.Drawing.Size(65, 26)
        Me.btnRemoveUrl.TabIndex = 25
        Me.btnRemoveUrl.Text = "刪 除"
        '
        'btnAddUrl
        '
        Me.btnAddUrl.Image = CType(resources.GetObject("btnAddUrl.Image"), System.Drawing.Image)
        Me.btnAddUrl.Location = New System.Drawing.Point(27, 19)
        Me.btnAddUrl.Name = "btnAddUrl"
        Me.btnAddUrl.Size = New System.Drawing.Size(65, 26)
        Me.btnAddUrl.TabIndex = 24
        Me.btnAddUrl.Text = "新 增"
        '
        'TreLstServerUrl
        '
        Me.TreLstServerUrl.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colPushServerUrl})
        Me.TreLstServerUrl.Location = New System.Drawing.Point(27, 54)
        Me.TreLstServerUrl.Name = "TreLstServerUrl"
        Me.TreLstServerUrl.OptionsView.ShowButtons = False
        Me.TreLstServerUrl.OptionsView.ShowRoot = False
        Me.TreLstServerUrl.Size = New System.Drawing.Size(716, 237)
        Me.TreLstServerUrl.StateImageList = Me.ImgLst
        Me.TreLstServerUrl.TabIndex = 23
        '
        'colPushServerUrl
        '
        Me.colPushServerUrl.AppearanceHeader.Options.UseTextOptions = True
        Me.colPushServerUrl.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colPushServerUrl.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colPushServerUrl.Caption = "授 權 網 址"
        Me.colPushServerUrl.FieldName = "PushServerUrl"
        Me.colPushServerUrl.Name = "colPushServerUrl"
        Me.colPushServerUrl.Visible = True
        Me.colPushServerUrl.VisibleIndex = 0
        Me.colPushServerUrl.Width = 92
        '
        'grpNagra
        '
        Me.grpNagra.Controls.Add(Me.chkSendTVMail)
        Me.grpNagra.Controls.Add(Me.chkErrSndMail)
        Me.grpNagra.Controls.Add(Me.spinResumeCnt)
        Me.grpNagra.Controls.Add(Me.LabelControl34)
        Me.grpNagra.Controls.Add(Me.spinPushSlowTime)
        Me.grpNagra.Controls.Add(Me.spinPushRetryNum)
        Me.grpNagra.Controls.Add(Me.LabelControl3)
        Me.grpNagra.Controls.Add(Me.cboPushUrlMethod)
        Me.grpNagra.Controls.Add(Me.LabelControl22)
        Me.grpNagra.Controls.Add(Me.spinPushWebTimeOut)
        Me.grpNagra.Controls.Add(Me.LabelControl21)
        Me.grpNagra.Controls.Add(Me.spinPushRepeat)
        Me.grpNagra.Controls.Add(Me.LabelControl20)
        Me.grpNagra.Controls.Add(Me.spinPushDuration)
        Me.grpNagra.Controls.Add(Me.LabelControl2)
        Me.grpNagra.Location = New System.Drawing.Point(27, 314)
        Me.grpNagra.Name = "grpNagra"
        Me.grpNagra.Size = New System.Drawing.Size(716, 138)
        Me.grpNagra.TabIndex = 22
        Me.grpNagra.Text = "Push Server 參 數"
        '
        'chkSendTVMail
        '
        Me.chkSendTVMail.Location = New System.Drawing.Point(413, 92)
        Me.chkSendTVMail.Name = "chkSendTVMail"
        Me.chkSendTVMail.Properties.Caption = "發送TVMail"
        Me.chkSendTVMail.Size = New System.Drawing.Size(125, 19)
        Me.chkSendTVMail.StyleController = Me.StyleController1
        Me.chkSendTVMail.TabIndex = 54
        '
        'chkErrSndMail
        '
        Me.chkErrSndMail.Location = New System.Drawing.Point(265, 92)
        Me.chkErrSndMail.Name = "chkErrSndMail"
        Me.chkErrSndMail.Properties.Caption = "命令錯誤寄送郵件"
        Me.chkErrSndMail.Size = New System.Drawing.Size(125, 19)
        Me.chkErrSndMail.StyleController = Me.StyleController1
        Me.chkErrSndMail.TabIndex = 53
        '
        'spinResumeCnt
        '
        Me.spinResumeCnt.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinResumeCnt.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinResumeCnt.Location = New System.Drawing.Point(123, 90)
        Me.spinResumeCnt.Name = "spinResumeCnt"
        Me.spinResumeCnt.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinResumeCnt.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinResumeCnt.Properties.IsFloatValue = False
        Me.spinResumeCnt.Properties.Mask.EditMask = "N00"
        Me.spinResumeCnt.Properties.MaxValue = New Decimal(New Integer() {999, 0, 0, 0})
        Me.spinResumeCnt.Properties.NullText = "0"
        Me.spinResumeCnt.Size = New System.Drawing.Size(49, 23)
        Me.spinResumeCnt.StyleController = Me.StyleController1
        Me.spinResumeCnt.TabIndex = 52
        '
        'LabelControl34
        '
        Me.LabelControl34.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl34.Location = New System.Drawing.Point(15, 94)
        Me.LabelControl34.Name = "LabelControl34"
        Me.LabelControl34.Size = New System.Drawing.Size(228, 14)
        Me.LabelControl34.TabIndex = 51
        Me.LabelControl34.Text = "錯誤命令還原到達                 次,不再還原"
        '
        'spinPushSlowTime
        '
        Me.spinPushSlowTime.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinPushSlowTime.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinPushSlowTime.Location = New System.Drawing.Point(607, 91)
        Me.spinPushSlowTime.Name = "spinPushSlowTime"
        Me.spinPushSlowTime.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinPushSlowTime.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinPushSlowTime.Properties.IsFloatValue = False
        Me.spinPushSlowTime.Properties.Mask.EditMask = "N00"
        Me.spinPushSlowTime.Properties.MaxValue = New Decimal(New Integer() {999999, 0, 0, 0})
        Me.spinPushSlowTime.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinPushSlowTime.Size = New System.Drawing.Size(60, 23)
        Me.spinPushSlowTime.StyleController = Me.StyleController1
        Me.spinPushSlowTime.TabIndex = 37
        Me.spinPushSlowTime.Visible = False
        '
        'spinPushRetryNum
        '
        Me.spinPushRetryNum.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinPushRetryNum.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinPushRetryNum.Location = New System.Drawing.Point(629, 44)
        Me.spinPushRetryNum.Name = "spinPushRetryNum"
        Me.spinPushRetryNum.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinPushRetryNum.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinPushRetryNum.Properties.IsFloatValue = False
        Me.spinPushRetryNum.Properties.Mask.EditMask = "N00"
        Me.spinPushRetryNum.Properties.MaxValue = New Decimal(New Integer() {999999, 0, 0, 0})
        Me.spinPushRetryNum.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinPushRetryNum.Size = New System.Drawing.Size(60, 23)
        Me.spinPushRetryNum.StyleController = Me.StyleController1
        Me.spinPushRetryNum.TabIndex = 36
        '
        'LabelControl3
        '
        Me.LabelControl3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl3.Location = New System.Drawing.Point(551, 48)
        Me.LabelControl3.Name = "LabelControl3"
        Me.LabelControl3.Size = New System.Drawing.Size(72, 14)
        Me.LabelControl3.TabIndex = 35
        Me.LabelControl3.Text = "錯誤重試次數"
        '
        'cboPushUrlMethod
        '
        Me.cboPushUrlMethod.EditValue = "POST"
        Me.cboPushUrlMethod.Enabled = False
        Me.cboPushUrlMethod.Location = New System.Drawing.Point(471, 44)
        Me.cboPushUrlMethod.Name = "cboPushUrlMethod"
        Me.cboPushUrlMethod.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboPushUrlMethod.Properties.Items.AddRange(New Object() {"POST", "GET"})
        Me.cboPushUrlMethod.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        Me.cboPushUrlMethod.Size = New System.Drawing.Size(64, 21)
        Me.cboPushUrlMethod.TabIndex = 34
        '
        'LabelControl22
        '
        Me.LabelControl22.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl22.Location = New System.Drawing.Point(415, 48)
        Me.LabelControl22.Name = "LabelControl22"
        Me.LabelControl22.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl22.TabIndex = 33
        Me.LabelControl22.Text = "傳送方式"
        '
        'spinPushWebTimeOut
        '
        Me.spinPushWebTimeOut.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinPushWebTimeOut.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinPushWebTimeOut.Location = New System.Drawing.Point(320, 44)
        Me.spinPushWebTimeOut.Name = "spinPushWebTimeOut"
        Me.spinPushWebTimeOut.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinPushWebTimeOut.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinPushWebTimeOut.Properties.IsFloatValue = False
        Me.spinPushWebTimeOut.Properties.Mask.EditMask = "N00"
        Me.spinPushWebTimeOut.Properties.MaxValue = New Decimal(New Integer() {999999, 0, 0, 0})
        Me.spinPushWebTimeOut.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinPushWebTimeOut.Size = New System.Drawing.Size(60, 23)
        Me.spinPushWebTimeOut.StyleController = Me.StyleController1
        Me.spinPushWebTimeOut.TabIndex = 32
        '
        'LabelControl21
        '
        Me.LabelControl21.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl21.Location = New System.Drawing.Point(268, 48)
        Me.LabelControl21.Name = "LabelControl21"
        Me.LabelControl21.Size = New System.Drawing.Size(128, 14)
        Me.LabelControl21.TabIndex = 31
        Me.LabelControl21.Text = "連線逾時                 秒"
        '
        'spinPushRepeat
        '
        Me.spinPushRepeat.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinPushRepeat.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinPushRepeat.Location = New System.Drawing.Point(189, 44)
        Me.spinPushRepeat.Name = "spinPushRepeat"
        Me.spinPushRepeat.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinPushRepeat.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinPushRepeat.Properties.IsFloatValue = False
        Me.spinPushRepeat.Properties.Mask.EditMask = "N00"
        Me.spinPushRepeat.Properties.MaxValue = New Decimal(New Integer() {999999, 0, 0, 0})
        Me.spinPushRepeat.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinPushRepeat.Size = New System.Drawing.Size(60, 23)
        Me.spinPushRepeat.StyleController = Me.StyleController1
        Me.spinPushRepeat.TabIndex = 30
        '
        'LabelControl20
        '
        Me.LabelControl20.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl20.Location = New System.Drawing.Point(135, 48)
        Me.LabelControl20.Name = "LabelControl20"
        Me.LabelControl20.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl20.TabIndex = 29
        Me.LabelControl20.Text = "重複次數"
        '
        'spinPushDuration
        '
        Me.spinPushDuration.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinPushDuration.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinPushDuration.Location = New System.Drawing.Point(69, 44)
        Me.spinPushDuration.Name = "spinPushDuration"
        Me.spinPushDuration.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinPushDuration.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinPushDuration.Properties.IsFloatValue = False
        Me.spinPushDuration.Properties.Mask.EditMask = "N00"
        Me.spinPushDuration.Properties.MaxValue = New Decimal(New Integer() {999999, 0, 0, 0})
        Me.spinPushDuration.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinPushDuration.Size = New System.Drawing.Size(60, 23)
        Me.spinPushDuration.StyleController = Me.StyleController1
        Me.spinPushDuration.TabIndex = 28
        '
        'LabelControl2
        '
        Me.LabelControl2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl2.Location = New System.Drawing.Point(15, 48)
        Me.LabelControl2.Name = "LabelControl2"
        Me.LabelControl2.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl2.TabIndex = 27
        Me.LabelControl2.Text = "持續時間"
        '
        'XtabTvMailMsg
        '
        Me.XtabTvMailMsg.Controls.Add(Me.chkUseRightKey)
        Me.XtabTvMailMsg.Controls.Add(Me.edtTVMailMsg)
        Me.XtabTvMailMsg.Image = CType(resources.GetObject("XtabTvMailMsg.Image"), System.Drawing.Image)
        Me.XtabTvMailMsg.Name = "XtabTvMailMsg"
        Me.XtabTvMailMsg.Size = New System.Drawing.Size(764, 484)
        Me.XtabTvMailMsg.Text = "TV Mail 發送訊息設定"
        '
        'edtTVMailMsg
        '
        Me.edtTVMailMsg.EditValue = ""
        Me.edtTVMailMsg.Location = New System.Drawing.Point(15, 66)
        Me.edtTVMailMsg.Name = "edtTVMailMsg"
        Me.edtTVMailMsg.Properties.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.edtTVMailMsg.Properties.Appearance.Options.UseFont = True
        Me.edtTVMailMsg.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.edtTVMailMsg.Properties.ContextMenuStrip = Me.munCustomerField
        Me.edtTVMailMsg.Properties.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.edtTVMailMsg.Size = New System.Drawing.Size(733, 395)
        Me.edtTVMailMsg.StyleController = Me.StyleController1
        Me.edtTVMailMsg.TabIndex = 1
        '
        'munCustomerField
        '
        Me.munCustomerField.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.item1})
        Me.munCustomerField.Name = "ContextMenuStrip1"
        Me.munCustomerField.Size = New System.Drawing.Size(171, 26)
        '
        'item1
        '
        Me.item1.Name = "item1"
        Me.item1.Size = New System.Drawing.Size(170, 22)
        Me.item1.Text = "ToolStripMenuItem1"
        '
        'XtabQtvAuth
        '
        Me.XtabQtvAuth.Controls.Add(Me.TreLstErrCode)
        Me.XtabQtvAuth.Controls.Add(Me.btnRemoveAuth)
        Me.XtabQtvAuth.Controls.Add(Me.btnAddAuth)
        Me.XtabQtvAuth.Image = CType(resources.GetObject("XtabQtvAuth.Image"), System.Drawing.Image)
        Me.XtabQtvAuth.Name = "XtabQtvAuth"
        Me.XtabQtvAuth.Size = New System.Drawing.Size(764, 484)
        Me.XtabQtvAuth.Text = " 錯 誤 代 碼 表 設 定 "
        '
        'TreLstErrCode
        '
        Me.TreLstErrCode.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colErrorCode, Me.colErrorName, Me.colErrorRetry, Me.colErrorResume})
        Me.TreLstErrCode.Location = New System.Drawing.Point(14, 54)
        Me.TreLstErrCode.Name = "TreLstErrCode"
        Me.TreLstErrCode.OptionsView.ShowButtons = False
        Me.TreLstErrCode.OptionsView.ShowRoot = False
        Me.TreLstErrCode.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.RepositoryItemCheckEdit1})
        Me.TreLstErrCode.Size = New System.Drawing.Size(735, 410)
        Me.TreLstErrCode.StateImageList = Me.ImgLst
        Me.TreLstErrCode.TabIndex = 22
        '
        'colErrorCode
        '
        Me.colErrorCode.AppearanceCell.Options.UseTextOptions = True
        Me.colErrorCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colErrorCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colErrorCode.AppearanceHeader.Options.UseTextOptions = True
        Me.colErrorCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colErrorCode.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colErrorCode.Caption = "錯誤代碼"
        Me.colErrorCode.FieldName = "ErrorCode"
        Me.colErrorCode.Name = "colErrorCode"
        Me.colErrorCode.Visible = True
        Me.colErrorCode.VisibleIndex = 2
        Me.colErrorCode.Width = 96
        '
        'colErrorName
        '
        Me.colErrorName.Caption = "錯誤名稱"
        Me.colErrorName.FieldName = "ErrorName"
        Me.colErrorName.Name = "colErrorName"
        Me.colErrorName.Visible = True
        Me.colErrorName.VisibleIndex = 3
        Me.colErrorName.Width = 444
        '
        'colErrorRetry
        '
        Me.colErrorRetry.AppearanceCell.Options.UseTextOptions = True
        Me.colErrorRetry.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colErrorRetry.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colErrorRetry.AppearanceHeader.Options.UseTextOptions = True
        Me.colErrorRetry.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colErrorRetry.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colErrorRetry.Caption = "錯誤重試"
        Me.colErrorRetry.ColumnEdit = Me.RepositoryItemCheckEdit1
        Me.colErrorRetry.FieldName = "ErrorRetry"
        Me.colErrorRetry.Name = "colErrorRetry"
        Me.colErrorRetry.Visible = True
        Me.colErrorRetry.VisibleIndex = 0
        Me.colErrorRetry.Width = 80
        '
        'RepositoryItemCheckEdit1
        '
        Me.RepositoryItemCheckEdit1.AutoHeight = False
        Me.RepositoryItemCheckEdit1.DisplayValueChecked = "1"
        Me.RepositoryItemCheckEdit1.DisplayValueUnchecked = "0"
        Me.RepositoryItemCheckEdit1.Name = "RepositoryItemCheckEdit1"
        Me.RepositoryItemCheckEdit1.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.RepositoryItemCheckEdit1.NullText = "0"
        Me.RepositoryItemCheckEdit1.ValueChecked = "1"
        Me.RepositoryItemCheckEdit1.ValueUnchecked = "0"
        '
        'colErrorResume
        '
        Me.colErrorResume.AppearanceCell.Options.UseTextOptions = True
        Me.colErrorResume.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colErrorResume.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colErrorResume.AppearanceHeader.Options.UseTextOptions = True
        Me.colErrorResume.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colErrorResume.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colErrorResume.Caption = "錯誤還原"
        Me.colErrorResume.ColumnEdit = Me.RepositoryItemCheckEdit1
        Me.colErrorResume.FieldName = "ErrorResume"
        Me.colErrorResume.Name = "colErrorResume"
        Me.colErrorResume.Visible = True
        Me.colErrorResume.VisibleIndex = 1
        Me.colErrorResume.Width = 94
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
        'XtabMailInfo
        '
        Me.XtabMailInfo.Controls.Add(Me.btnDelMailCC)
        Me.XtabMailInfo.Controls.Add(Me.btnInsMailCC)
        Me.XtabMailInfo.Controls.Add(Me.TreLstMailCC)
        Me.XtabMailInfo.Controls.Add(Me.btnDelSMailTo)
        Me.XtabMailInfo.Controls.Add(Me.btnInsSMailTo)
        Me.XtabMailInfo.Controls.Add(Me.TreLstMailTo)
        Me.XtabMailInfo.Controls.Add(Me.GroupControl1)
        Me.XtabMailInfo.Image = CType(resources.GetObject("XtabMailInfo.Image"), System.Drawing.Image)
        Me.XtabMailInfo.Name = "XtabMailInfo"
        Me.XtabMailInfo.Size = New System.Drawing.Size(764, 484)
        Me.XtabMailInfo.Text = "發 送 Mail  警 告 設 定 "
        '
        'btnDelMailCC
        '
        Me.btnDelMailCC.Image = CType(resources.GetObject("btnDelMailCC.Image"), System.Drawing.Image)
        Me.btnDelMailCC.Location = New System.Drawing.Point(462, 17)
        Me.btnDelMailCC.Name = "btnDelMailCC"
        Me.btnDelMailCC.Size = New System.Drawing.Size(65, 26)
        Me.btnDelMailCC.TabIndex = 33
        Me.btnDelMailCC.Text = "刪 除"
        '
        'btnInsMailCC
        '
        Me.btnInsMailCC.Image = CType(resources.GetObject("btnInsMailCC.Image"), System.Drawing.Image)
        Me.btnInsMailCC.Location = New System.Drawing.Point(391, 17)
        Me.btnInsMailCC.Name = "btnInsMailCC"
        Me.btnInsMailCC.Size = New System.Drawing.Size(65, 26)
        Me.btnInsMailCC.TabIndex = 32
        Me.btnInsMailCC.Text = "新 增"
        '
        'TreLstMailCC
        '
        Me.TreLstMailCC.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colMailCC})
        Me.TreLstMailCC.Location = New System.Drawing.Point(380, 57)
        Me.TreLstMailCC.Name = "TreLstMailCC"
        Me.TreLstMailCC.OptionsView.ShowButtons = False
        Me.TreLstMailCC.OptionsView.ShowRoot = False
        Me.TreLstMailCC.Size = New System.Drawing.Size(367, 186)
        Me.TreLstMailCC.StateImageList = Me.ImgLst
        Me.TreLstMailCC.TabIndex = 31
        '
        'colMailCC
        '
        Me.colMailCC.AppearanceHeader.Options.UseTextOptions = True
        Me.colMailCC.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMailCC.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMailCC.Caption = "副 本"
        Me.colMailCC.FieldName = "MailCC"
        Me.colMailCC.Name = "colMailCC"
        Me.colMailCC.Visible = True
        Me.colMailCC.VisibleIndex = 0
        Me.colMailCC.Width = 123
        '
        'btnDelSMailTo
        '
        Me.btnDelSMailTo.Image = CType(resources.GetObject("btnDelSMailTo.Image"), System.Drawing.Image)
        Me.btnDelSMailTo.Location = New System.Drawing.Point(85, 17)
        Me.btnDelSMailTo.Name = "btnDelSMailTo"
        Me.btnDelSMailTo.Size = New System.Drawing.Size(65, 26)
        Me.btnDelSMailTo.TabIndex = 30
        Me.btnDelSMailTo.Text = "刪 除"
        '
        'btnInsSMailTo
        '
        Me.btnInsSMailTo.Image = CType(resources.GetObject("btnInsSMailTo.Image"), System.Drawing.Image)
        Me.btnInsSMailTo.Location = New System.Drawing.Point(14, 17)
        Me.btnInsSMailTo.Name = "btnInsSMailTo"
        Me.btnInsSMailTo.Size = New System.Drawing.Size(65, 26)
        Me.btnInsSMailTo.TabIndex = 29
        Me.btnInsSMailTo.Text = "新 增"
        '
        'TreLstMailTo
        '
        Me.TreLstMailTo.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colMailTo})
        Me.TreLstMailTo.Location = New System.Drawing.Point(14, 57)
        Me.TreLstMailTo.Name = "TreLstMailTo"
        Me.TreLstMailTo.OptionsView.ShowButtons = False
        Me.TreLstMailTo.OptionsView.ShowRoot = False
        Me.TreLstMailTo.Size = New System.Drawing.Size(359, 186)
        Me.TreLstMailTo.StateImageList = Me.ImgLst
        Me.TreLstMailTo.TabIndex = 28
        '
        'colMailTo
        '
        Me.colMailTo.AppearanceHeader.Options.UseTextOptions = True
        Me.colMailTo.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMailTo.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMailTo.Caption = "收 件 人 "
        Me.colMailTo.FieldName = "MailTo"
        Me.colMailTo.Name = "colMailTo"
        Me.colMailTo.Visible = True
        Me.colMailTo.VisibleIndex = 0
        Me.colMailTo.Width = 123
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
        Me.GroupControl1.Controls.Add(Me.LabelControl4)
        Me.GroupControl1.Controls.Add(Me.spinWebErrCnt)
        Me.GroupControl1.Controls.Add(Me.LabelControl5)
        Me.GroupControl1.Controls.Add(Me.TextEdit2)
        Me.GroupControl1.Controls.Add(Me.TextEdit3)
        Me.GroupControl1.Controls.Add(Me.LabelControl30)
        Me.GroupControl1.Controls.Add(Me.TextEdit4)
        Me.GroupControl1.Controls.Add(Me.LabelControl31)
        Me.GroupControl1.Controls.Add(Me.LabelControl32)
        Me.GroupControl1.Controls.Add(Me.TextEdit5)
        Me.GroupControl1.Controls.Add(Me.LabelControl33)
        Me.GroupControl1.Location = New System.Drawing.Point(15, 253)
        Me.GroupControl1.Name = "GroupControl1"
        Me.GroupControl1.Size = New System.Drawing.Size(732, 198)
        Me.GroupControl1.TabIndex = 26
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
        Me.txtMailBody.Size = New System.Drawing.Size(647, 23)
        Me.txtMailBody.StyleController = Me.StyleController1
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
        Me.txtMailSubject.Location = New System.Drawing.Point(178, 99)
        Me.txtMailSubject.Name = "txtMailSubject"
        Me.txtMailSubject.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtMailSubject.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtMailSubject.Properties.Appearance.Options.UseBackColor = True
        Me.txtMailSubject.Properties.Appearance.Options.UseForeColor = True
        Me.txtMailSubject.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtMailSubject.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtMailSubject.Size = New System.Drawing.Size(538, 23)
        Me.txtMailSubject.StyleController = Me.StyleController1
        Me.txtMailSubject.TabIndex = 47
        '
        'LabelControl28
        '
        Me.LabelControl28.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl28.Location = New System.Drawing.Point(124, 103)
        Me.LabelControl28.Name = "LabelControl28"
        Me.LabelControl28.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl28.TabIndex = 54
        Me.LabelControl28.Text = "郵件主旨"
        '
        'txtDisplayName
        '
        Me.txtDisplayName.EditValue = ""
        Me.txtDisplayName.Location = New System.Drawing.Point(595, 64)
        Me.txtDisplayName.Name = "txtDisplayName"
        Me.txtDisplayName.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtDisplayName.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtDisplayName.Properties.Appearance.Options.UseBackColor = True
        Me.txtDisplayName.Properties.Appearance.Options.UseForeColor = True
        Me.txtDisplayName.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtDisplayName.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtDisplayName.Size = New System.Drawing.Size(121, 23)
        Me.txtDisplayName.StyleController = Me.StyleController1
        Me.txtDisplayName.TabIndex = 45
        '
        'LabelControl27
        '
        Me.LabelControl27.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl27.Location = New System.Drawing.Point(541, 68)
        Me.LabelControl27.Name = "LabelControl27"
        Me.LabelControl27.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl27.TabIndex = 52
        Me.LabelControl27.Text = "顯示名稱"
        '
        'chkEnabledSSL
        '
        Me.chkEnabledSSL.Location = New System.Drawing.Point(13, 100)
        Me.chkEnabledSSL.Name = "chkEnabledSSL"
        Me.chkEnabledSSL.Properties.Caption = "啟用SSL安全性"
        Me.chkEnabledSSL.Size = New System.Drawing.Size(115, 19)
        Me.chkEnabledSSL.TabIndex = 46
        '
        'txtSMTPLoginPsw
        '
        Me.txtSMTPLoginPsw.EditValue = ""
        Me.txtSMTPLoginPsw.Location = New System.Drawing.Point(232, 65)
        Me.txtSMTPLoginPsw.Name = "txtSMTPLoginPsw"
        Me.txtSMTPLoginPsw.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtSMTPLoginPsw.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtSMTPLoginPsw.Properties.Appearance.Options.UseBackColor = True
        Me.txtSMTPLoginPsw.Properties.Appearance.Options.UseForeColor = True
        Me.txtSMTPLoginPsw.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtSMTPLoginPsw.Properties.NullValuePrompt = "請輸入密碼"
        Me.txtSMTPLoginPsw.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtSMTPLoginPsw.Properties.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtSMTPLoginPsw.Size = New System.Drawing.Size(91, 23)
        Me.txtSMTPLoginPsw.StyleController = Me.StyleController1
        Me.txtSMTPLoginPsw.TabIndex = 43
        '
        'LabelControl26
        '
        Me.LabelControl26.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl26.Location = New System.Drawing.Point(178, 69)
        Me.LabelControl26.Name = "LabelControl26"
        Me.LabelControl26.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl26.TabIndex = 48
        Me.LabelControl26.Text = "登入密碼"
        '
        'txtSMTPLoginID
        '
        Me.txtSMTPLoginID.EditValue = ""
        Me.txtSMTPLoginID.Location = New System.Drawing.Point(81, 64)
        Me.txtSMTPLoginID.Name = "txtSMTPLoginID"
        Me.txtSMTPLoginID.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtSMTPLoginID.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtSMTPLoginID.Properties.Appearance.Options.UseBackColor = True
        Me.txtSMTPLoginID.Properties.Appearance.Options.UseForeColor = True
        Me.txtSMTPLoginID.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtSMTPLoginID.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtSMTPLoginID.Size = New System.Drawing.Size(91, 23)
        Me.txtSMTPLoginID.StyleController = Me.StyleController1
        Me.txtSMTPLoginID.TabIndex = 42
        '
        'LabelControl25
        '
        Me.LabelControl25.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl25.Location = New System.Drawing.Point(15, 69)
        Me.LabelControl25.Name = "LabelControl25"
        Me.LabelControl25.Size = New System.Drawing.Size(60, 14)
        Me.LabelControl25.TabIndex = 46
        Me.LabelControl25.Text = "登入使用者"
        '
        'txtSender
        '
        Me.txtSender.EditValue = ""
        Me.txtSender.Location = New System.Drawing.Point(376, 64)
        Me.txtSender.Name = "txtSender"
        Me.txtSender.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtSender.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtSender.Properties.Appearance.Options.UseBackColor = True
        Me.txtSender.Properties.Appearance.Options.UseForeColor = True
        Me.txtSender.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtSender.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtSender.Size = New System.Drawing.Size(159, 23)
        Me.txtSender.StyleController = Me.StyleController1
        Me.txtSender.TabIndex = 44
        '
        'LabelControl24
        '
        Me.LabelControl24.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl24.Location = New System.Drawing.Point(334, 68)
        Me.LabelControl24.Name = "LabelControl24"
        Me.LabelControl24.Size = New System.Drawing.Size(36, 14)
        Me.LabelControl24.TabIndex = 44
        Me.LabelControl24.Text = "寄件者"
        '
        'spinSMTPPort
        '
        Me.spinSMTPPort.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinSMTPPort.EditValue = New Decimal(New Integer() {25, 0, 0, 0})
        Me.spinSMTPPort.Location = New System.Drawing.Point(595, 28)
        Me.spinSMTPPort.Name = "spinSMTPPort"
        Me.spinSMTPPort.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinSMTPPort.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinSMTPPort.Properties.IsFloatValue = False
        Me.spinSMTPPort.Properties.Mask.EditMask = "N00"
        Me.spinSMTPPort.Properties.MaxValue = New Decimal(New Integer() {1200, 0, 0, 0})
        Me.spinSMTPPort.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinSMTPPort.Size = New System.Drawing.Size(56, 23)
        Me.spinSMTPPort.StyleController = Me.StyleController1
        Me.spinSMTPPort.TabIndex = 41
        '
        'LabelControl23
        '
        Me.LabelControl23.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl23.Location = New System.Drawing.Point(553, 32)
        Me.LabelControl23.Name = "LabelControl23"
        Me.LabelControl23.Size = New System.Drawing.Size(36, 14)
        Me.LabelControl23.TabIndex = 42
        Me.LabelControl23.Text = "通訊埠"
        '
        'txtSMTPHost
        '
        Me.txtSMTPHost.EditValue = ""
        Me.txtSMTPHost.Location = New System.Drawing.Point(338, 28)
        Me.txtSMTPHost.Name = "txtSMTPHost"
        Me.txtSMTPHost.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtSMTPHost.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtSMTPHost.Properties.Appearance.Options.UseBackColor = True
        Me.txtSMTPHost.Properties.Appearance.Options.UseForeColor = True
        Me.txtSMTPHost.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtSMTPHost.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtSMTPHost.Size = New System.Drawing.Size(201, 23)
        Me.txtSMTPHost.StyleController = Me.StyleController1
        Me.txtSMTPHost.TabIndex = 40
        '
        'LabelControl4
        '
        Me.LabelControl4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl4.Location = New System.Drawing.Point(272, 32)
        Me.LabelControl4.Name = "LabelControl4"
        Me.LabelControl4.Size = New System.Drawing.Size(60, 14)
        Me.LabelControl4.TabIndex = 40
        Me.LabelControl4.Text = "郵件伺服器"
        '
        'spinWebErrCnt
        '
        Me.spinWebErrCnt.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinWebErrCnt.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinWebErrCnt.Location = New System.Drawing.Point(79, 28)
        Me.spinWebErrCnt.Name = "spinWebErrCnt"
        Me.spinWebErrCnt.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinWebErrCnt.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinWebErrCnt.Properties.IsFloatValue = False
        Me.spinWebErrCnt.Properties.Mask.EditMask = "N00"
        Me.spinWebErrCnt.Properties.MaxValue = New Decimal(New Integer() {1200, 0, 0, 0})
        Me.spinWebErrCnt.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinWebErrCnt.Size = New System.Drawing.Size(49, 23)
        Me.spinWebErrCnt.StyleController = Me.StyleController1
        Me.spinWebErrCnt.TabIndex = 39
        '
        'LabelControl5
        '
        Me.LabelControl5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl5.Location = New System.Drawing.Point(15, 32)
        Me.LabelControl5.Name = "LabelControl5"
        Me.LabelControl5.Size = New System.Drawing.Size(236, 14)
        Me.LabelControl5.TabIndex = 38
        Me.LabelControl5.Text = "網站發生                   次失敗,寄發警告郵件"
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
        Me.TextEdit2.Size = New System.Drawing.Size(78, 23)
        Me.TextEdit2.StyleController = Me.StyleController1
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
        Me.TextEdit3.Size = New System.Drawing.Size(78, 23)
        Me.TextEdit3.StyleController = Me.StyleController1
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
        Me.TextEdit4.Size = New System.Drawing.Size(78, 23)
        Me.TextEdit4.StyleController = Me.StyleController1
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
        Me.TextEdit5.Size = New System.Drawing.Size(78, 23)
        Me.TextEdit5.StyleController = Me.StyleController1
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
        'rpstrItmTxtEdtPwd2
        '
        Me.rpstrItmTxtEdtPwd2.AutoHeight = False
        Me.rpstrItmTxtEdtPwd2.Name = "rpstrItmTxtEdtPwd2"
        Me.rpstrItmTxtEdtPwd2.NullText = "未設定密碼"
        Me.rpstrItmTxtEdtPwd2.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        '
        'chkUseRightKey
        '
        Me.chkUseRightKey.EditValue = True
        Me.chkUseRightKey.Location = New System.Drawing.Point(13, 31)
        Me.chkUseRightKey.Name = "chkUseRightKey"
        Me.chkUseRightKey.Properties.Caption = "在訊息輸入框,使用滑鼠右鍵彈出可用變數"
        Me.chkUseRightKey.Size = New System.Drawing.Size(316, 19)
        Me.chkUseRightKey.StyleController = Me.StyleController1
        Me.chkUseRightKey.TabIndex = 55
        '
        'xfrmSystemSet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(798, 602)
        Me.Controls.Add(Me.grpCtl)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Name = "xfrmSystemSet"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "xfrmSystemSet"
        CType(Me.grpCtl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpCtl.ResumeLayout(False)
        CType(Me.XtabCtl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabCtl.ResumeLayout(False)
        Me.XtabSysOpt.ResumeLayout(False)
        CType(Me.grpQtv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpQtv.ResumeLayout(False)
        Me.grpQtv.PerformLayout()
        CType(Me.txtPushUrl.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StyleController1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpPooling, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpPooling.ResumeLayout(False)
        Me.grpPooling.PerformLayout()
        CType(Me.spinLifeP.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinMaxP.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinMinP.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpGW, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpGW.ResumeLayout(False)
        Me.grpGW.PerformLayout()
        CType(Me.txtTitle.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkResource.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinMaxThreads.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkHotKey.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkTray.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinFreq.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinBuff.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinRcd.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinProc.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkAutoRun.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabDbOpt.ResumeLayout(False)
        CType(Me.grpComDB, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpComDB.ResumeLayout(False)
        Me.grpComDB.PerformLayout()
        CType(Me.txtCOMPassword.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCOMUserId.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCOMSid.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TreLstDB, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rpstrItmChkEdt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rpstrItmSpnEdt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rpstrItmTxtEdtPwd, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabServerUrl.ResumeLayout(False)
        CType(Me.TreLstServerUrl, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpNagra, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpNagra.ResumeLayout(False)
        Me.grpNagra.PerformLayout()
        CType(Me.chkSendTVMail.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkErrSndMail.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinResumeCnt.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinPushSlowTime.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinPushRetryNum.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboPushUrlMethod.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinPushWebTimeOut.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinPushRepeat.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinPushDuration.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabTvMailMsg.ResumeLayout(False)
        CType(Me.edtTVMailMsg.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.munCustomerField.ResumeLayout(False)
        Me.XtabQtvAuth.ResumeLayout(False)
        CType(Me.TreLstErrCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabMailInfo.ResumeLayout(False)
        CType(Me.TreLstMailCC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TreLstMailTo, System.ComponentModel.ISupportInitialize).EndInit()
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
        CType(Me.spinWebErrCnt.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TextEdit2.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TextEdit3.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TextEdit4.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TextEdit5.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rpstrItmTxtEdtPwd2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkUseRightKey.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grpCtl As DevExpress.XtraEditors.GroupControl
    Friend WithEvents btnCancel As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnOK As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents XtabCtl As DevExpress.XtraTab.XtraTabControl
    Friend WithEvents XtabSysOpt As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents grpQtv As DevExpress.XtraEditors.GroupControl
    Friend WithEvents grpPooling As DevExpress.XtraEditors.GroupControl
    Friend WithEvents spinLifeP As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents spinMaxP As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents spinMinP As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents lblPooling As DevExpress.XtraEditors.LabelControl
    Friend WithEvents grpGW As DevExpress.XtraEditors.GroupControl
    Friend WithEvents txtTitle As DevExpress.XtraEditors.TextEdit
    Friend WithEvents lblTitle As DevExpress.XtraEditors.LabelControl
    Friend WithEvents lblCore As DevExpress.XtraEditors.LabelControl
    Friend WithEvents chkResource As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents spinMaxThreads As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents lblMaxThreads As DevExpress.XtraEditors.LabelControl
    Friend WithEvents chkHotKey As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents chkTray As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents spinFreq As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents spinBuff As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents spinRcd As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents spinProc As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents lblDisp As DevExpress.XtraEditors.LabelControl
    Friend WithEvents lblDuration As DevExpress.XtraEditors.LabelControl
    Friend WithEvents chkAutoRun As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents XtabDbOpt As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents lblMsg As DevExpress.XtraEditors.LabelControl
    Friend WithEvents btnTestConn As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnRemoveConn As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAddConn As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents XtabServerUrl As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents XtabQtvAuth As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents btnRemoveAuth As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAddAuth As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents rpstrItmTxtEdtPwd2 As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents TreLstDB As DevExpress.XtraTreeList.TreeList
    Friend WithEvents ImgLst As System.Windows.Forms.ImageList
    Friend WithEvents colSelSO As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents rpstrItmChkEdt As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents rpstrItmSpnEdt As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents colCompCode As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colCompName As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents rpstrItmTxtEdtPwd As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents StyleController1 As DevExpress.XtraEditors.StyleController
    Friend WithEvents grpNagra As DevExpress.XtraEditors.GroupControl
    Friend WithEvents TreLstErrCode As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colErrorCode As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colErrorName As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents grpComDB As DevExpress.XtraEditors.GroupControl
    Friend WithEvents txtCOMPassword As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl19 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtCOMUserId As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl7 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtCOMSid As DevExpress.XtraEditors.TextEdit
    Friend WithEvents lblComSid As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtPushUrl As DevExpress.XtraEditors.TextEdit
    Friend WithEvents TreLstServerUrl As DevExpress.XtraTreeList.TreeList
    Friend WithEvents btnRemoveUrl As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAddUrl As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents colPushServerUrl As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents spinPushRetryNum As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl3 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cboPushUrlMethod As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LabelControl22 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinPushWebTimeOut As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl21 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinPushRepeat As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl20 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinPushDuration As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl2 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinPushSlowTime As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents colErrorRetry As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents RepositoryItemCheckEdit1 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents XtabMailInfo As DevExpress.XtraTab.XtraTabPage
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
    Friend WithEvents LabelControl4 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinWebErrCnt As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl5 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents TextEdit2 As DevExpress.XtraEditors.TextEdit
    Friend WithEvents TextEdit3 As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl30 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents TextEdit4 As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl31 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl32 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents TextEdit5 As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl33 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents btnDelMailCC As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnInsMailCC As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents TreLstMailCC As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colMailCC As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents btnDelSMailTo As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnInsSMailTo As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents TreLstMailTo As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colMailTo As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents chkErrSndMail As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents spinResumeCnt As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl34 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents colErrorResume As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colSidUserId As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colSidPassword As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colSidDBname As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents btnTestCom As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents lblComMsg As DevExpress.XtraEditors.LabelControl
    Friend WithEvents XtabTvMailMsg As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents edtTVMailMsg As DevExpress.XtraEditors.MemoEdit
    Friend WithEvents chkSendTVMail As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents munCustomerField As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents item1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents chkUseRightKey As DevExpress.XtraEditors.CheckEdit
End Class
