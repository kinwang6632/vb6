<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSetting
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSetting))
        Dim SuperToolTip1 As DevExpress.Utils.SuperToolTip = New DevExpress.Utils.SuperToolTip()
        Dim ToolTipTitleItem1 As DevExpress.Utils.ToolTipTitleItem = New DevExpress.Utils.ToolTipTitleItem()
        Dim ToolTipItem1 As DevExpress.Utils.ToolTipItem = New DevExpress.Utils.ToolTipItem()
        Dim ToolTipSeparatorItem1 As DevExpress.Utils.ToolTipSeparatorItem = New DevExpress.Utils.ToolTipSeparatorItem()
        Dim ToolTipTitleItem2 As DevExpress.Utils.ToolTipTitleItem = New DevExpress.Utils.ToolTipTitleItem()
        Me.DefaultLookAndFeel1 = New DevExpress.LookAndFeel.DefaultLookAndFeel(Me.components)
        Me.StyleController1 = New DevExpress.XtraEditors.StyleController(Me.components)
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        Me.grpCtl = New DevExpress.XtraEditors.GroupControl()
        Me.btnCancel = New DevExpress.XtraEditors.SimpleButton()
        Me.btnOK = New DevExpress.XtraEditors.SimpleButton()
        Me.XtabCtl = New DevExpress.XtraTab.XtraTabControl()
        Me.XtabSysOpt = New DevExpress.XtraTab.XtraTabPage()
        Me.grpQtv = New DevExpress.XtraEditors.GroupControl()
        Me.spinPort = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl2 = New DevExpress.XtraEditors.LabelControl()
        Me.txtIPAddress = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.spinTimeout = New DevExpress.XtraEditors.SpinEdit()
        Me.txtCredPswd = New DevExpress.XtraEditors.TextEdit()
        Me.txtCredUser = New DevExpress.XtraEditors.TextEdit()
        Me.lblTimeout = New DevExpress.XtraEditors.LabelControl()
        Me.lblCred = New DevExpress.XtraEditors.LabelControl()
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
        Me.spinMaxThreads = New DevExpress.XtraEditors.SpinEdit()
        Me.lblMaxThreads = New DevExpress.XtraEditors.LabelControl()
        Me.chkHotKey = New DevExpress.XtraEditors.CheckEdit()
        Me.chkTray = New DevExpress.XtraEditors.CheckEdit()
        Me.spinFreq = New DevExpress.XtraEditors.SpinEdit()
        Me.spinBuff = New DevExpress.XtraEditors.SpinEdit()
        Me.spinRcd = New DevExpress.XtraEditors.SpinEdit()
        Me.spinProc = New DevExpress.XtraEditors.SpinEdit()
        Me.spinReCon = New DevExpress.XtraEditors.SpinEdit()
        Me.lblDisp = New DevExpress.XtraEditors.LabelControl()
        Me.lblDuration = New DevExpress.XtraEditors.LabelControl()
        Me.lblReConn = New DevExpress.XtraEditors.LabelControl()
        Me.txtGwPswd = New DevExpress.XtraEditors.TextEdit()
        Me.lblGwPwd = New DevExpress.XtraEditors.LabelControl()
        Me.chkAutoRun = New DevExpress.XtraEditors.CheckEdit()
        Me.XtabDbOpt = New DevExpress.XtraTab.XtraTabPage()
        Me.TreLstDB = New DevExpress.XtraTreeList.TreeList()
        Me.colSelSO = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.RepositoryItemCheckEdit4 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.colCompCode = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.RepositoryItemSpinEdit4 = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.colCompName = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colDBSid = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colDBUser = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colDBPassword = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.RepositoryItemTextEdit6 = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.grpComDB = New DevExpress.XtraEditors.GroupControl()
        Me.txtCOMPassword = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl19 = New DevExpress.XtraEditors.LabelControl()
        Me.txtCOMUserId = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl7 = New DevExpress.XtraEditors.LabelControl()
        Me.txtCOMSid = New DevExpress.XtraEditors.TextEdit()
        Me.lblComSid = New DevExpress.XtraEditors.LabelControl()
        Me.lblMsg = New DevExpress.XtraEditors.LabelControl()
        Me.btnTestConn = New DevExpress.XtraEditors.SimpleButton()
        Me.btnRemoveConn = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAddConn = New DevExpress.XtraEditors.SimpleButton()
        Me.XtabQtvWS = New DevExpress.XtraTab.XtraTabPage()
        Me.TreLstSourceId = New DevExpress.XtraTreeList.TreeList()
        Me.colServer = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.rpedtIP = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.colPort = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.rpedtPort = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.ColRootKey = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.rptedtNstvRootKey = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.btnbtnRemoveSourceId = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAddSourceId = New DevExpress.XtraEditors.SimpleButton()
        Me.grpNagra = New DevExpress.XtraEditors.GroupControl()
        Me.txtEmailTitle = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl35 = New DevExpress.XtraEditors.LabelControl()
        Me.cobProcKind = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.LabelControl34 = New DevExpress.XtraEditors.LabelControl()
        Me.spinRetryCount = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl21 = New DevExpress.XtraEditors.LabelControl()
        Me.chkUseUTC = New DevExpress.XtraEditors.CheckEdit()
        Me.spinDisconnectRetry = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl18 = New DevExpress.XtraEditors.LabelControl()
        Me.spinReceiveTimeout = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl17 = New DevExpress.XtraEditors.LabelControl()
        Me.spinSendTimeout = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl16 = New DevExpress.XtraEditors.LabelControl()
        Me.chkCanRecord = New DevExpress.XtraEditors.CheckEdit()
        Me.cmbEncoding = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.LabelControl12 = New DevExpress.XtraEditors.LabelControl()
        Me.cobKeyType = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.LabelControl11 = New DevExpress.XtraEditors.LabelControl()
        Me.spinCrypt_Ver = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl9 = New DevExpress.XtraEditors.LabelControl()
        Me.spinProtoVer = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl8 = New DevExpress.XtraEditors.LabelControl()
        Me.txtOPE_ID = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl6 = New DevExpress.XtraEditors.LabelControl()
        Me.spinSndDelay = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl5 = New DevExpress.XtraEditors.LabelControl()
        Me.spinChkStatus = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl4 = New DevExpress.XtraEditors.LabelControl()
        Me.txtRcvMopPPId = New DevExpress.XtraEditors.TextEdit()
        Me.txtRcvDestId = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl15 = New DevExpress.XtraEditors.LabelControl()
        Me.txtRcvSourceId = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl14 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl13 = New DevExpress.XtraEditors.LabelControl()
        Me.txtSndMopPPId = New DevExpress.XtraEditors.TextEdit()
        Me.txtSndDestId = New DevExpress.XtraEditors.TextEdit()
        Me.txtSndSourceId = New DevExpress.XtraEditors.TextEdit()
        Me.spinNstvPort = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl10 = New DevExpress.XtraEditors.LabelControl()
        Me.txtSMS_ID = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl3 = New DevExpress.XtraEditors.LabelControl()
        Me.XtabLowCmd = New DevExpress.XtraTab.XtraTabPage()
        Me.TreLstLowCmd = New DevExpress.XtraTreeList.TreeList()
        Me.colCode = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colValue = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colDescription = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colDefaultValue = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colRealField = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.RepositoryItemCheckEdit6 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.btnDelLowCmd = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAddLowCmd = New DevExpress.XtraEditors.SimpleButton()
        Me.XtabHighCmd = New DevExpress.XtraTab.XtraTabPage()
        Me.TreLstHightCmd = New DevExpress.XtraTreeList.TreeList()
        Me.colCodeNo = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colHightDesc = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colRunLowCmd = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colRunLowCmd2 = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.RepositoryItemCheckEdit7 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.btnDelHighCmd = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAddHightCmd = New DevExpress.XtraEditors.SimpleButton()
        Me.XtabQtvAuth = New DevExpress.XtraTab.XtraTabPage()
        Me.TreLstErrCode = New DevExpress.XtraTreeList.TreeList()
        Me.colErrorCode = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colErrorName = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.RepositoryItemCheckEdit5 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.btnRemoveAuth = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAddError = New DevExpress.XtraEditors.SimpleButton()
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
        Me.rpstrItmChkEdt = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.rpstrItmSpnEdt = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.rpstrItmTxtEdtPwd = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.RepositoryItemTextEdit1 = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.RepositoryItemTextEdit2 = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.chkChoice = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.RepositoryItemCheckEdit1 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.RepositoryItemSpinEdit1 = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.RepositoryItemTextEdit3 = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.RepositoryItemCheckEdit2 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.RepositoryItemSpinEdit2 = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.RepositoryItemTextEdit4 = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.RepositoryItemCheckEdit3 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.RepositoryItemSpinEdit3 = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.RepositoryItemTextEdit5 = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        CType(Me.StyleController1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        CType(Me.grpCtl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpCtl.SuspendLayout()
        CType(Me.XtabCtl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabCtl.SuspendLayout()
        Me.XtabSysOpt.SuspendLayout()
        CType(Me.grpQtv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpQtv.SuspendLayout()
        CType(Me.spinPort.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtIPAddress.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinTimeout.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCredPswd.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCredUser.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
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
        CType(Me.spinReCon.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtGwPswd.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkAutoRun.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabDbOpt.SuspendLayout()
        CType(Me.TreLstDB, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemSpinEdit4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemTextEdit6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpComDB, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpComDB.SuspendLayout()
        CType(Me.txtCOMPassword.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCOMUserId.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCOMSid.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabQtvWS.SuspendLayout()
        CType(Me.TreLstSourceId, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rpedtIP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rpedtPort, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rptedtNstvRootKey, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpNagra, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpNagra.SuspendLayout()
        CType(Me.txtEmailTitle.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cobProcKind.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinRetryCount.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkUseUTC.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinDisconnectRetry.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinReceiveTimeout.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinSendTimeout.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkCanRecord.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmbEncoding.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cobKeyType.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinCrypt_Ver.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinProtoVer.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtOPE_ID.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinSndDelay.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinChkStatus.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtRcvMopPPId.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtRcvDestId.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtRcvSourceId.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSndMopPPId.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSndDestId.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSndSourceId.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinNstvPort.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSMS_ID.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabLowCmd.SuspendLayout()
        CType(Me.TreLstLowCmd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit6, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabHighCmd.SuspendLayout()
        CType(Me.TreLstHightCmd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabQtvAuth.SuspendLayout()
        CType(Me.TreLstErrCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit5, System.ComponentModel.ISupportInitialize).BeginInit()
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
        CType(Me.rpstrItmChkEdt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rpstrItmSpnEdt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rpstrItmTxtEdtPwd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemTextEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemTextEdit2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkChoice, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemSpinEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemTextEdit3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemSpinEdit2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemTextEdit4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemSpinEdit3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemTextEdit5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DefaultLookAndFeel1
        '
        Me.DefaultLookAndFeel1.LookAndFeel.SkinName = "Office 2010 Silver"
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "1297823140_OK.png")
        Me.ImageList1.Images.SetKeyName(1, "1297823327_exit.png")
        '
        'PanelControl1
        '
        Me.PanelControl1.Controls.Add(Me.grpCtl)
        Me.PanelControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelControl1.Location = New System.Drawing.Point(0, 0)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(926, 610)
        Me.PanelControl1.TabIndex = 0
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
        Me.grpCtl.Location = New System.Drawing.Point(2, 2)
        Me.grpCtl.Name = "grpCtl"
        Me.grpCtl.Size = New System.Drawing.Size(922, 606)
        Me.grpCtl.TabIndex = 2
        Me.grpCtl.Text = " 設 定 "
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Image = CType(resources.GetObject("btnCancel.Image"), System.Drawing.Image)
        Me.btnCancel.Location = New System.Drawing.Point(556, 566)
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
        Me.btnOK.Location = New System.Drawing.Point(203, 566)
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
        Me.XtabCtl.Location = New System.Drawing.Point(10, 36)
        Me.XtabCtl.Name = "XtabCtl"
        Me.XtabCtl.SelectedTabPage = Me.XtabSysOpt
        Me.XtabCtl.ShowHeaderFocus = DevExpress.Utils.DefaultBoolean.[False]
        Me.XtabCtl.Size = New System.Drawing.Size(902, 512)
        Me.XtabCtl.TabIndex = 0
        Me.XtabCtl.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XtabSysOpt, Me.XtabDbOpt, Me.XtabQtvWS, Me.XtabLowCmd, Me.XtabHighCmd, Me.XtabQtvAuth, Me.XtabSMailInfo})
        '
        'XtabSysOpt
        '
        Me.XtabSysOpt.Controls.Add(Me.grpQtv)
        Me.XtabSysOpt.Controls.Add(Me.grpPooling)
        Me.XtabSysOpt.Controls.Add(Me.grpGW)
        Me.XtabSysOpt.Image = CType(resources.GetObject("XtabSysOpt.Image"), System.Drawing.Image)
        Me.XtabSysOpt.Name = "XtabSysOpt"
        Me.XtabSysOpt.Size = New System.Drawing.Size(896, 481)
        Me.XtabSysOpt.Text = " G/W 系 統 參 數 設 定 "
        '
        'grpQtv
        '
        Me.grpQtv.Controls.Add(Me.spinPort)
        Me.grpQtv.Controls.Add(Me.LabelControl2)
        Me.grpQtv.Controls.Add(Me.txtIPAddress)
        Me.grpQtv.Controls.Add(Me.LabelControl1)
        Me.grpQtv.Controls.Add(Me.spinTimeout)
        Me.grpQtv.Controls.Add(Me.txtCredPswd)
        Me.grpQtv.Controls.Add(Me.txtCredUser)
        Me.grpQtv.Controls.Add(Me.lblTimeout)
        Me.grpQtv.Controls.Add(Me.lblCred)
        Me.grpQtv.Location = New System.Drawing.Point(30, 399)
        Me.grpQtv.Name = "grpQtv"
        Me.grpQtv.Size = New System.Drawing.Size(843, 70)
        Me.grpQtv.TabIndex = 2
        Me.grpQtv.Text = " Nagra Socket 設 定 "
        Me.grpQtv.Visible = False
        '
        'spinPort
        '
        Me.spinPort.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinPort.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinPort.Location = New System.Drawing.Point(522, 34)
        Me.spinPort.Name = "spinPort"
        Me.spinPort.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinPort.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinPort.Properties.IsFloatValue = False
        Me.spinPort.Size = New System.Drawing.Size(60, 22)
        Me.spinPort.StyleController = Me.StyleController1
        Me.spinPort.TabIndex = 15
        '
        'LabelControl2
        '
        Me.LabelControl2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl2.Location = New System.Drawing.Point(493, 38)
        Me.LabelControl2.Name = "LabelControl2"
        Me.LabelControl2.Size = New System.Drawing.Size(23, 14)
        Me.LabelControl2.TabIndex = 9
        Me.LabelControl2.Text = "Port"
        '
        'txtIPAddress
        '
        Me.txtIPAddress.EditValue = ""
        Me.txtIPAddress.Location = New System.Drawing.Point(373, 34)
        Me.txtIPAddress.Name = "txtIPAddress"
        Me.txtIPAddress.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtIPAddress.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtIPAddress.Properties.Appearance.Options.UseBackColor = True
        Me.txtIPAddress.Properties.Appearance.Options.UseForeColor = True
        Me.txtIPAddress.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtIPAddress.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtIPAddress.Size = New System.Drawing.Size(106, 22)
        Me.txtIPAddress.StyleController = Me.StyleController1
        Me.txtIPAddress.TabIndex = 8
        '
        'LabelControl1
        '
        Me.LabelControl1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl1.Location = New System.Drawing.Point(309, 38)
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(58, 14)
        Me.LabelControl1.TabIndex = 5
        Me.LabelControl1.Text = "IP Address"
        '
        'spinTimeout
        '
        Me.spinTimeout.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinTimeout.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinTimeout.Location = New System.Drawing.Point(218, 34)
        Me.spinTimeout.Name = "spinTimeout"
        Me.spinTimeout.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinTimeout.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinTimeout.Properties.IsFloatValue = False
        Me.spinTimeout.Properties.Mask.EditMask = "N00"
        Me.spinTimeout.Properties.MaxValue = New Decimal(New Integer() {300, 0, 0, 0})
        Me.spinTimeout.Size = New System.Drawing.Size(49, 22)
        Me.spinTimeout.StyleController = Me.StyleController1
        Me.spinTimeout.TabIndex = 1
        '
        'txtCredPswd
        '
        Me.txtCredPswd.EditValue = ""
        Me.txtCredPswd.Location = New System.Drawing.Point(493, 75)
        Me.txtCredPswd.Name = "txtCredPswd"
        Me.txtCredPswd.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtCredPswd.Properties.Appearance.Options.UseForeColor = True
        Me.txtCredPswd.Properties.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtCredPswd.Size = New System.Drawing.Size(106, 20)
        Me.txtCredPswd.TabIndex = 4
        Me.txtCredPswd.TabStop = False
        Me.txtCredPswd.Visible = False
        '
        'txtCredUser
        '
        Me.txtCredUser.EditValue = ""
        Me.txtCredUser.Location = New System.Drawing.Point(339, 75)
        Me.txtCredUser.Name = "txtCredUser"
        Me.txtCredUser.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtCredUser.Properties.Appearance.Options.UseForeColor = True
        Me.txtCredUser.Properties.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtCredUser.Size = New System.Drawing.Size(106, 20)
        Me.txtCredUser.TabIndex = 3
        Me.txtCredUser.TabStop = False
        Me.txtCredUser.Visible = False
        '
        'lblTimeout
        '
        Me.lblTimeout.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblTimeout.Location = New System.Drawing.Point(19, 38)
        Me.lblTimeout.Name = "lblTimeout"
        Me.lblTimeout.Size = New System.Drawing.Size(275, 14)
        Me.lblTimeout.TabIndex = 0
        Me.lblTimeout.Text = "Gateway 測試 Nagra Socket(1002)                  秒"
        '
        'lblCred
        '
        Me.lblCred.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblCred.Location = New System.Drawing.Point(104, 67)
        Me.lblCred.Name = "lblCred"
        Me.lblCred.Size = New System.Drawing.Size(377, 14)
        Me.lblCred.TabIndex = 2
        Me.lblCred.Text = "Quative Web Service Credential ,  使用者                                密碼"
        Me.lblCred.Visible = False
        '
        'grpPooling
        '
        Me.grpPooling.Controls.Add(Me.spinLifeP)
        Me.grpPooling.Controls.Add(Me.spinMaxP)
        Me.grpPooling.Controls.Add(Me.spinMinP)
        Me.grpPooling.Controls.Add(Me.lblPooling)
        Me.grpPooling.Location = New System.Drawing.Point(30, 156)
        Me.grpPooling.Name = "grpPooling"
        Me.grpPooling.Size = New System.Drawing.Size(843, 70)
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
        Me.spinMaxP.Size = New System.Drawing.Size(60, 22)
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
        Me.spinMinP.Size = New System.Drawing.Size(60, 22)
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
        Me.grpGW.Controls.Add(Me.spinReCon)
        Me.grpGW.Controls.Add(Me.lblDisp)
        Me.grpGW.Controls.Add(Me.lblDuration)
        Me.grpGW.Controls.Add(Me.lblReConn)
        Me.grpGW.Controls.Add(Me.txtGwPswd)
        Me.grpGW.Controls.Add(Me.lblGwPwd)
        Me.grpGW.Controls.Add(Me.chkAutoRun)
        Me.grpGW.Location = New System.Drawing.Point(30, 14)
        Me.grpGW.Name = "grpGW"
        Me.grpGW.Size = New System.Drawing.Size(843, 133)
        Me.grpGW.TabIndex = 0
        Me.grpGW.Text = " G / W 設 定 "
        '
        'txtTitle
        '
        Me.txtTitle.EditValue = ""
        Me.txtTitle.Location = New System.Drawing.Point(268, 274)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtTitle.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtTitle.Properties.Appearance.Options.UseBackColor = True
        Me.txtTitle.Properties.Appearance.Options.UseForeColor = True
        Me.txtTitle.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtTitle.Size = New System.Drawing.Size(429, 22)
        Me.txtTitle.StyleController = Me.StyleController1
        Me.txtTitle.TabIndex = 1
        Me.txtTitle.Visible = False
        '
        'lblTitle
        '
        Me.lblTitle.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblTitle.Location = New System.Drawing.Point(214, 122)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(48, 14)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "抬頭顯示"
        Me.lblTitle.ToolTip = "Title"
        Me.lblTitle.Visible = False
        '
        'lblCore
        '
        Me.lblCore.Appearance.ForeColor = System.Drawing.Color.Red
        Me.lblCore.Location = New System.Drawing.Point(485, 206)
        Me.lblCore.Name = "lblCore"
        Me.lblCore.Size = New System.Drawing.Size(54, 14)
        Me.lblCore.TabIndex = 18
        Me.lblCore.Text = "( N 核心 )"
        Me.lblCore.Visible = False
        '
        'chkResource
        '
        Me.chkResource.Location = New System.Drawing.Point(217, 311)
        Me.chkResource.Name = "chkResource"
        Me.chkResource.Properties.Caption = " 顯示系統資源於狀態列 (Resource)"
        Me.chkResource.Properties.ImageIndexChecked = 0
        Me.chkResource.Properties.ImageIndexUnchecked = 1
        Me.chkResource.Size = New System.Drawing.Size(229, 19)
        Me.chkResource.StyleController = Me.StyleController1
        Me.chkResource.TabIndex = 4
        Me.chkResource.Visible = False
        '
        'spinMaxThreads
        '
        Me.spinMaxThreads.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinMaxThreads.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinMaxThreads.Location = New System.Drawing.Point(330, 71)
        Me.spinMaxThreads.Name = "spinMaxThreads"
        Me.spinMaxThreads.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinMaxThreads.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinMaxThreads.Properties.IsFloatValue = False
        Me.spinMaxThreads.Properties.Mask.EditMask = "N00"
        Me.spinMaxThreads.Properties.MaxValue = New Decimal(New Integer() {500, 0, 0, 0})
        Me.spinMaxThreads.Size = New System.Drawing.Size(60, 22)
        Me.spinMaxThreads.StyleController = Me.StyleController1
        Me.spinMaxThreads.TabIndex = 17
        '
        'lblMaxThreads
        '
        Me.lblMaxThreads.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblMaxThreads.Location = New System.Drawing.Point(117, 75)
        Me.lblMaxThreads.Name = "lblMaxThreads"
        Me.lblMaxThreads.Size = New System.Drawing.Size(347, 14)
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
        Me.chkHotKey.Location = New System.Drawing.Point(462, 311)
        Me.chkHotKey.Name = "chkHotKey"
        Me.chkHotKey.Properties.Caption = " 按下 [Ctrl] + [Alt] + [G] 自動呼叫開啟 Gateway"
        Me.chkHotKey.Size = New System.Drawing.Size(299, 19)
        Me.chkHotKey.StyleController = Me.StyleController1
        Me.chkHotKey.TabIndex = 5
        Me.chkHotKey.Visible = False
        '
        'chkTray
        '
        Me.chkTray.Location = New System.Drawing.Point(462, 283)
        Me.chkTray.Name = "chkTray"
        Me.chkTray.Properties.Caption = " 最小化後常駐在系統工作列上Tray (TSR)"
        Me.chkTray.Size = New System.Drawing.Size(282, 19)
        Me.chkTray.StyleController = Me.StyleController1
        Me.chkTray.TabIndex = 3
        Me.chkTray.Visible = False
        '
        'spinFreq
        '
        Me.spinFreq.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinFreq.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinFreq.Location = New System.Drawing.Point(217, 33)
        Me.spinFreq.Name = "spinFreq"
        Me.spinFreq.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinFreq.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinFreq.Properties.IsFloatValue = False
        Me.spinFreq.Properties.Mask.EditMask = "N00"
        Me.spinFreq.Properties.MaxValue = New Decimal(New Integer() {1200, 0, 0, 0})
        Me.spinFreq.Size = New System.Drawing.Size(49, 22)
        Me.spinFreq.StyleController = Me.StyleController1
        Me.spinFreq.TabIndex = 11
        '
        'spinBuff
        '
        Me.spinBuff.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinBuff.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinBuff.Location = New System.Drawing.Point(465, 278)
        Me.spinBuff.Name = "spinBuff"
        Me.spinBuff.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinBuff.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinBuff.Properties.IsFloatValue = False
        Me.spinBuff.Properties.Mask.EditMask = "N00"
        Me.spinBuff.Properties.MaxValue = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.spinBuff.Size = New System.Drawing.Size(60, 22)
        Me.spinBuff.StyleController = Me.StyleController1
        Me.spinBuff.TabIndex = 15
        Me.spinBuff.Visible = False
        '
        'spinRcd
        '
        Me.spinRcd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinRcd.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinRcd.Location = New System.Drawing.Point(298, 278)
        Me.spinRcd.Name = "spinRcd"
        Me.spinRcd.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinRcd.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinRcd.Properties.IsFloatValue = False
        Me.spinRcd.Properties.Mask.EditMask = "N00"
        Me.spinRcd.Properties.MaxValue = New Decimal(New Integer() {5000, 0, 0, 0})
        Me.spinRcd.Size = New System.Drawing.Size(60, 22)
        Me.spinRcd.StyleController = Me.StyleController1
        Me.spinRcd.TabIndex = 14
        Me.spinRcd.Visible = False
        '
        'spinProc
        '
        Me.spinProc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinProc.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinProc.Location = New System.Drawing.Point(449, 33)
        Me.spinProc.Name = "spinProc"
        Me.spinProc.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinProc.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinProc.Properties.IsFloatValue = False
        Me.spinProc.Properties.Mask.EditMask = "N00"
        Me.spinProc.Properties.MaxValue = New Decimal(New Integer() {500, 0, 0, 0})
        Me.spinProc.Size = New System.Drawing.Size(49, 22)
        Me.spinProc.StyleController = Me.StyleController1
        Me.spinProc.TabIndex = 12
        '
        'spinReCon
        '
        Me.spinReCon.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinReCon.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinReCon.Enabled = False
        Me.spinReCon.Location = New System.Drawing.Point(376, 213)
        Me.spinReCon.Name = "spinReCon"
        Me.spinReCon.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.UltraFlat
        Me.spinReCon.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinReCon.Properties.IsFloatValue = False
        Me.spinReCon.Properties.Mask.EditMask = "N00"
        Me.spinReCon.Properties.MaxValue = New Decimal(New Integer() {600, 0, 0, 0})
        Me.spinReCon.Size = New System.Drawing.Size(49, 20)
        Me.spinReCon.StyleController = Me.StyleController1
        Me.spinReCon.TabIndex = 9
        Me.spinReCon.Visible = False
        '
        'lblDisp
        '
        Me.lblDisp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblDisp.Location = New System.Drawing.Point(217, 281)
        Me.lblDisp.Name = "lblDisp"
        Me.lblDisp.Size = New System.Drawing.Size(480, 14)
        Me.lblDisp.TabIndex = 13
        Me.lblDisp.Text = "畫面資料顯示                    筆訊息 ,  每超過                    筆訊息 ,  則開始清除多餘訊息"
        Me.lblDisp.Visible = False
        '
        'lblDuration
        '
        Me.lblDuration.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblDuration.Location = New System.Drawing.Point(118, 36)
        Me.lblDuration.Name = "lblDuration"
        Me.lblDuration.Size = New System.Drawing.Size(453, 14)
        Me.lblDuration.TabIndex = 10
        Me.lblDuration.Text = "一般時段  -> 每                   秒讀取系統台資料 ,  每次處理                  筆高階指令"
        '
        'lblReConn
        '
        Me.lblReConn.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblReConn.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.lblReConn.Location = New System.Drawing.Point(217, 217)
        Me.lblReConn.Name = "lblReConn"
        Me.lblReConn.Size = New System.Drawing.Size(280, 14)
        Me.lblReConn.TabIndex = 8
        Me.lblReConn.Text = "系統台資料庫斷線時 ,  等候                  秒重試連線"
        Me.lblReConn.Visible = False
        '
        'txtGwPswd
        '
        Me.txtGwPswd.EditValue = ""
        Me.txtGwPswd.Enabled = False
        Me.txtGwPswd.Location = New System.Drawing.Point(326, 338)
        Me.txtGwPswd.Name = "txtGwPswd"
        Me.txtGwPswd.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtGwPswd.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtGwPswd.Properties.Appearance.Options.UseBackColor = True
        Me.txtGwPswd.Properties.Appearance.Options.UseForeColor = True
        Me.txtGwPswd.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtGwPswd.Properties.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtGwPswd.Size = New System.Drawing.Size(106, 22)
        Me.txtGwPswd.StyleController = Me.StyleController1
        Me.txtGwPswd.TabIndex = 7
        Me.txtGwPswd.Visible = False
        '
        'lblGwPwd
        '
        Me.lblGwPwd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblGwPwd.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.lblGwPwd.Location = New System.Drawing.Point(217, 185)
        Me.lblGwPwd.Name = "lblGwPwd"
        Me.lblGwPwd.Size = New System.Drawing.Size(322, 14)
        Me.lblGwPwd.TabIndex = 6
        Me.lblGwPwd.Text = "Gateway 啟動密碼                              ( 此功能暫不提供 )"
        Me.lblGwPwd.Visible = False
        '
        'chkAutoRun
        '
        Me.chkAutoRun.Location = New System.Drawing.Point(217, 283)
        Me.chkAutoRun.Name = "chkAutoRun"
        Me.chkAutoRun.Properties.Caption = " Windows 開機後自動執行 Gateway"
        Me.chkAutoRun.Size = New System.Drawing.Size(235, 19)
        Me.chkAutoRun.StyleController = Me.StyleController1
        Me.chkAutoRun.TabIndex = 2
        Me.chkAutoRun.Visible = False
        '
        'XtabDbOpt
        '
        Me.XtabDbOpt.Controls.Add(Me.TreLstDB)
        Me.XtabDbOpt.Controls.Add(Me.grpComDB)
        Me.XtabDbOpt.Controls.Add(Me.lblMsg)
        Me.XtabDbOpt.Controls.Add(Me.btnTestConn)
        Me.XtabDbOpt.Controls.Add(Me.btnRemoveConn)
        Me.XtabDbOpt.Controls.Add(Me.btnAddConn)
        Me.XtabDbOpt.Image = CType(resources.GetObject("XtabDbOpt.Image"), System.Drawing.Image)
        Me.XtabDbOpt.Name = "XtabDbOpt"
        Me.XtabDbOpt.Size = New System.Drawing.Size(896, 481)
        Me.XtabDbOpt.Text = " 系 統 台 DB 連 線 設 定 "
        '
        'TreLstDB
        '
        Me.TreLstDB.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colSelSO, Me.colCompCode, Me.colCompName, Me.colDBSid, Me.colDBUser, Me.colDBPassword})
        Me.TreLstDB.Location = New System.Drawing.Point(14, 49)
        Me.TreLstDB.Name = "TreLstDB"
        Me.TreLstDB.OptionsPrint.UsePrintStyles = True
        Me.TreLstDB.OptionsView.AutoWidth = False
        Me.TreLstDB.OptionsView.ShowButtons = False
        Me.TreLstDB.OptionsView.ShowRoot = False
        Me.TreLstDB.OptionsView.ShowSummaryFooter = True
        Me.TreLstDB.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.RepositoryItemCheckEdit4, Me.RepositoryItemSpinEdit4, Me.RepositoryItemTextEdit6})
        Me.TreLstDB.Size = New System.Drawing.Size(866, 424)
        Me.TreLstDB.TabIndex = 27
        '
        'colSelSO
        '
        Me.colSelSO.AppearanceHeader.Options.UseTextOptions = True
        Me.colSelSO.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSelSO.Caption = "啟 動 作 業"
        Me.colSelSO.ColumnEdit = Me.RepositoryItemCheckEdit4
        Me.colSelSO.FieldName = "SelectID"
        Me.colSelSO.MinWidth = 88
        Me.colSelSO.Name = "colSelSO"
        Me.colSelSO.SummaryFooter = DevExpress.XtraTreeList.SummaryItemType.Count
        Me.colSelSO.SummaryFooterStrFormat = "共 {0:N0} 個系統台"
        Me.colSelSO.Visible = True
        Me.colSelSO.VisibleIndex = 0
        Me.colSelSO.Width = 88
        '
        'RepositoryItemCheckEdit4
        '
        Me.RepositoryItemCheckEdit4.AutoHeight = False
        Me.RepositoryItemCheckEdit4.DisplayValueChecked = "1"
        Me.RepositoryItemCheckEdit4.DisplayValueUnchecked = "0"
        Me.RepositoryItemCheckEdit4.Name = "RepositoryItemCheckEdit4"
        Me.RepositoryItemCheckEdit4.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.RepositoryItemCheckEdit4.ValueChecked = "1"
        Me.RepositoryItemCheckEdit4.ValueUnchecked = "0"
        '
        'colCompCode
        '
        Me.colCompCode.AppearanceCell.Options.UseTextOptions = True
        Me.colCompCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompCode.AppearanceHeader.Options.UseTextOptions = True
        Me.colCompCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompCode.Caption = "系 統 台 代 碼"
        Me.colCompCode.ColumnEdit = Me.RepositoryItemSpinEdit4
        Me.colCompCode.FieldName = "CompID"
        Me.colCompCode.Name = "colCompCode"
        Me.colCompCode.Visible = True
        Me.colCompCode.VisibleIndex = 1
        Me.colCompCode.Width = 102
        '
        'RepositoryItemSpinEdit4
        '
        Me.RepositoryItemSpinEdit4.AllowNullInput = DevExpress.Utils.DefaultBoolean.[False]
        Me.RepositoryItemSpinEdit4.Appearance.Options.UseTextOptions = True
        Me.RepositoryItemSpinEdit4.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.RepositoryItemSpinEdit4.AutoHeight = False
        Me.RepositoryItemSpinEdit4.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.RepositoryItemSpinEdit4.IsFloatValue = False
        Me.RepositoryItemSpinEdit4.Mask.EditMask = "N00"
        Me.RepositoryItemSpinEdit4.MaxLength = 2
        Me.RepositoryItemSpinEdit4.MaxValue = New Decimal(New Integer() {99, 0, 0, 0})
        Me.RepositoryItemSpinEdit4.Name = "RepositoryItemSpinEdit4"
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
        Me.colCompName.Width = 276
        '
        'colDBSid
        '
        Me.colDBSid.Caption = "SID"
        Me.colDBSid.FieldName = "DBSid"
        Me.colDBSid.Name = "colDBSid"
        Me.colDBSid.Visible = True
        Me.colDBSid.VisibleIndex = 3
        Me.colDBSid.Width = 136
        '
        'colDBUser
        '
        Me.colDBUser.Caption = "使用者名稱"
        Me.colDBUser.FieldName = "DBUser"
        Me.colDBUser.Name = "colDBUser"
        Me.colDBUser.Visible = True
        Me.colDBUser.VisibleIndex = 4
        Me.colDBUser.Width = 130
        '
        'colDBPassword
        '
        Me.colDBPassword.Caption = "使用者密碼"
        Me.colDBPassword.ColumnEdit = Me.RepositoryItemTextEdit6
        Me.colDBPassword.FieldName = "DBPassword"
        Me.colDBPassword.Name = "colDBPassword"
        Me.colDBPassword.Visible = True
        Me.colDBPassword.VisibleIndex = 5
        Me.colDBPassword.Width = 150
        '
        'RepositoryItemTextEdit6
        '
        Me.RepositoryItemTextEdit6.AutoHeight = False
        Me.RepositoryItemTextEdit6.Name = "RepositoryItemTextEdit6"
        Me.RepositoryItemTextEdit6.NullText = "未設定密碼"
        Me.RepositoryItemTextEdit6.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        '
        'grpComDB
        '
        Me.grpComDB.Controls.Add(Me.txtCOMPassword)
        Me.grpComDB.Controls.Add(Me.LabelControl19)
        Me.grpComDB.Controls.Add(Me.txtCOMUserId)
        Me.grpComDB.Controls.Add(Me.LabelControl7)
        Me.grpComDB.Controls.Add(Me.txtCOMSid)
        Me.grpComDB.Controls.Add(Me.lblComSid)
        Me.grpComDB.Location = New System.Drawing.Point(14, 468)
        Me.grpComDB.Name = "grpComDB"
        Me.grpComDB.Size = New System.Drawing.Size(735, 74)
        Me.grpComDB.TabIndex = 26
        Me.grpComDB.Text = "COM 區 設 定"
        Me.grpComDB.Visible = False
        '
        'txtCOMPassword
        '
        Me.txtCOMPassword.Location = New System.Drawing.Point(584, 37)
        Me.txtCOMPassword.Name = "txtCOMPassword"
        Me.txtCOMPassword.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtCOMPassword.Properties.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtCOMPassword.Size = New System.Drawing.Size(131, 22)
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
        Me.txtCOMUserId.Size = New System.Drawing.Size(110, 22)
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
        Me.txtCOMSid.Size = New System.Drawing.Size(104, 22)
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
        'lblMsg
        '
        Me.lblMsg.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblMsg.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.lblMsg.Location = New System.Drawing.Point(234, 3)
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
        Me.XtabQtvWS.Controls.Add(Me.TreLstSourceId)
        Me.XtabQtvWS.Controls.Add(Me.btnbtnRemoveSourceId)
        Me.XtabQtvWS.Controls.Add(Me.btnAddSourceId)
        Me.XtabQtvWS.Controls.Add(Me.grpNagra)
        Me.XtabQtvWS.Image = CType(resources.GetObject("XtabQtvWS.Image"), System.Drawing.Image)
        Me.XtabQtvWS.Name = "XtabQtvWS"
        Me.XtabQtvWS.Size = New System.Drawing.Size(896, 481)
        Me.XtabQtvWS.Text = "NSTV 主 機 設 定 "
        '
        'TreLstSourceId
        '
        Me.TreLstSourceId.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colServer, Me.colPort, Me.ColRootKey})
        Me.TreLstSourceId.Location = New System.Drawing.Point(26, 48)
        Me.TreLstSourceId.Name = "TreLstSourceId"
        Me.TreLstSourceId.OptionsPrint.UsePrintStyles = True
        Me.TreLstSourceId.OptionsView.ShowButtons = False
        Me.TreLstSourceId.OptionsView.ShowRoot = False
        Me.TreLstSourceId.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.rpedtIP, Me.rpedtPort, Me.rptedtNstvRootKey})
        Me.TreLstSourceId.Size = New System.Drawing.Size(860, 189)
        Me.TreLstSourceId.TabIndex = 26
        '
        'colServer
        '
        Me.colServer.Caption = "IPAddress"
        Me.colServer.ColumnEdit = Me.rpedtIP
        Me.colServer.FieldName = "Server"
        Me.colServer.MinWidth = 33
        Me.colServer.Name = "colServer"
        Me.colServer.Visible = True
        Me.colServer.VisibleIndex = 0
        Me.colServer.Width = 261
        '
        'rpedtIP
        '
        Me.rpedtIP.AutoHeight = False
        Me.rpedtIP.Mask.PlaceHolder = Global.Microsoft.VisualBasic.ChrW(32)
        Me.rpedtIP.Name = "rpedtIP"
        '
        'colPort
        '
        Me.colPort.Caption = "Port"
        Me.colPort.ColumnEdit = Me.rpedtPort
        Me.colPort.FieldName = "Port"
        Me.colPort.Name = "colPort"
        Me.colPort.Visible = True
        Me.colPort.VisibleIndex = 1
        Me.colPort.Width = 231
        '
        'rpedtPort
        '
        Me.rpedtPort.AutoHeight = False
        Me.rpedtPort.Name = "rpedtPort"
        '
        'ColRootKey
        '
        Me.ColRootKey.Caption = "Root Key"
        Me.ColRootKey.ColumnEdit = Me.rptedtNstvRootKey
        Me.ColRootKey.FieldName = "RootKey"
        Me.ColRootKey.Name = "ColRootKey"
        Me.ColRootKey.Visible = True
        Me.ColRootKey.VisibleIndex = 2
        Me.ColRootKey.Width = 350
        '
        'rptedtNstvRootKey
        '
        Me.rptedtNstvRootKey.AutoHeight = False
        Me.rptedtNstvRootKey.Name = "rptedtNstvRootKey"
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
        'grpNagra
        '
        Me.grpNagra.Controls.Add(Me.txtEmailTitle)
        Me.grpNagra.Controls.Add(Me.LabelControl35)
        Me.grpNagra.Controls.Add(Me.cobProcKind)
        Me.grpNagra.Controls.Add(Me.LabelControl34)
        Me.grpNagra.Controls.Add(Me.spinRetryCount)
        Me.grpNagra.Controls.Add(Me.LabelControl21)
        Me.grpNagra.Controls.Add(Me.chkUseUTC)
        Me.grpNagra.Controls.Add(Me.spinDisconnectRetry)
        Me.grpNagra.Controls.Add(Me.LabelControl18)
        Me.grpNagra.Controls.Add(Me.spinReceiveTimeout)
        Me.grpNagra.Controls.Add(Me.LabelControl17)
        Me.grpNagra.Controls.Add(Me.spinSendTimeout)
        Me.grpNagra.Controls.Add(Me.LabelControl16)
        Me.grpNagra.Controls.Add(Me.chkCanRecord)
        Me.grpNagra.Controls.Add(Me.cmbEncoding)
        Me.grpNagra.Controls.Add(Me.LabelControl12)
        Me.grpNagra.Controls.Add(Me.cobKeyType)
        Me.grpNagra.Controls.Add(Me.LabelControl11)
        Me.grpNagra.Controls.Add(Me.spinCrypt_Ver)
        Me.grpNagra.Controls.Add(Me.LabelControl9)
        Me.grpNagra.Controls.Add(Me.spinProtoVer)
        Me.grpNagra.Controls.Add(Me.LabelControl8)
        Me.grpNagra.Controls.Add(Me.txtOPE_ID)
        Me.grpNagra.Controls.Add(Me.LabelControl6)
        Me.grpNagra.Controls.Add(Me.spinSndDelay)
        Me.grpNagra.Controls.Add(Me.LabelControl5)
        Me.grpNagra.Controls.Add(Me.spinChkStatus)
        Me.grpNagra.Controls.Add(Me.LabelControl4)
        Me.grpNagra.Controls.Add(Me.txtRcvMopPPId)
        Me.grpNagra.Controls.Add(Me.txtRcvDestId)
        Me.grpNagra.Controls.Add(Me.LabelControl15)
        Me.grpNagra.Controls.Add(Me.txtRcvSourceId)
        Me.grpNagra.Controls.Add(Me.LabelControl14)
        Me.grpNagra.Controls.Add(Me.LabelControl13)
        Me.grpNagra.Controls.Add(Me.txtSndMopPPId)
        Me.grpNagra.Controls.Add(Me.txtSndDestId)
        Me.grpNagra.Controls.Add(Me.txtSndSourceId)
        Me.grpNagra.Controls.Add(Me.spinNstvPort)
        Me.grpNagra.Controls.Add(Me.LabelControl10)
        Me.grpNagra.Controls.Add(Me.txtSMS_ID)
        Me.grpNagra.Controls.Add(Me.LabelControl3)
        Me.grpNagra.Location = New System.Drawing.Point(26, 252)
        Me.grpNagra.Name = "grpNagra"
        Me.grpNagra.Size = New System.Drawing.Size(855, 205)
        Me.grpNagra.TabIndex = 22
        Me.grpNagra.Text = "Nagra 設 定 "
        '
        'txtEmailTitle
        '
        Me.txtEmailTitle.EditValue = ""
        Me.txtEmailTitle.Location = New System.Drawing.Point(88, 158)
        Me.txtEmailTitle.Name = "txtEmailTitle"
        Me.txtEmailTitle.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtEmailTitle.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtEmailTitle.Properties.Appearance.Options.UseBackColor = True
        Me.txtEmailTitle.Properties.Appearance.Options.UseForeColor = True
        Me.txtEmailTitle.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtEmailTitle.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtEmailTitle.Size = New System.Drawing.Size(429, 22)
        Me.txtEmailTitle.StyleController = Me.StyleController1
        Me.txtEmailTitle.TabIndex = 67
        '
        'LabelControl35
        '
        Me.LabelControl35.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl35.Location = New System.Drawing.Point(27, 162)
        Me.LabelControl35.Name = "LabelControl35"
        Me.LabelControl35.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl35.TabIndex = 66
        Me.LabelControl35.Text = "郵件標題"
        '
        'cobProcKind
        '
        Me.cobProcKind.EditValue = "1.全部"
        Me.cobProcKind.Location = New System.Drawing.Point(607, 124)
        Me.cobProcKind.Name = "cobProcKind"
        Me.cobProcKind.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cobProcKind.Properties.Items.AddRange(New Object() {"1.全部", "2.即時命令", "3.預約命令"})
        Me.cobProcKind.Size = New System.Drawing.Size(91, 20)
        Me.cobProcKind.TabIndex = 65
        '
        'LabelControl34
        '
        Me.LabelControl34.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl34.Location = New System.Drawing.Point(523, 127)
        Me.LabelControl34.Name = "LabelControl34"
        Me.LabelControl34.Size = New System.Drawing.Size(72, 14)
        Me.LabelControl34.TabIndex = 64
        Me.LabelControl34.Text = "處理命令種類"
        '
        'spinRetryCount
        '
        Me.spinRetryCount.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinRetryCount.EditValue = New Decimal(New Integer() {3, 0, 0, 0})
        Me.spinRetryCount.Location = New System.Drawing.Point(448, 123)
        Me.spinRetryCount.Name = "spinRetryCount"
        Me.spinRetryCount.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinRetryCount.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinRetryCount.Properties.IsFloatValue = False
        Me.spinRetryCount.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.None
        Me.spinRetryCount.Properties.MaxValue = New Decimal(New Integer() {9, 0, 0, 0})
        Me.spinRetryCount.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinRetryCount.Size = New System.Drawing.Size(69, 22)
        Me.spinRetryCount.StyleController = Me.StyleController1
        Me.spinRetryCount.TabIndex = 63
        '
        'LabelControl21
        '
        Me.LabelControl21.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl21.Location = New System.Drawing.Point(368, 127)
        Me.LabelControl21.Name = "LabelControl21"
        Me.LabelControl21.Size = New System.Drawing.Size(72, 14)
        Me.LabelControl21.TabIndex = 62
        Me.LabelControl21.Text = "失敗重送次數"
        '
        'chkUseUTC
        '
        Me.chkUseUTC.EditValue = CType(0, Short)
        Me.chkUseUTC.Location = New System.Drawing.Point(746, 52)
        Me.chkUseUTC.Name = "chkUseUTC"
        Me.chkUseUTC.Properties.Caption = "使用 UTC 時間"
        Me.chkUseUTC.Properties.NullText = "0"
        Me.chkUseUTC.Properties.ValueChecked = CType(1, Short)
        Me.chkUseUTC.Properties.ValueGrayed = CType(0, Short)
        Me.chkUseUTC.Properties.ValueUnchecked = CType(0, Short)
        Me.chkUseUTC.Size = New System.Drawing.Size(104, 19)
        Me.chkUseUTC.TabIndex = 61
        '
        'spinDisconnectRetry
        '
        Me.spinDisconnectRetry.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinDisconnectRetry.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinDisconnectRetry.Location = New System.Drawing.Point(649, 87)
        Me.spinDisconnectRetry.Name = "spinDisconnectRetry"
        Me.spinDisconnectRetry.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinDisconnectRetry.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinDisconnectRetry.Properties.IsFloatValue = False
        Me.spinDisconnectRetry.Properties.Mask.EditMask = "N00"
        Me.spinDisconnectRetry.Properties.MaxValue = New Decimal(New Integer() {1200, 0, 0, 0})
        Me.spinDisconnectRetry.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinDisconnectRetry.Size = New System.Drawing.Size(49, 22)
        Me.spinDisconnectRetry.StyleController = Me.StyleController1
        Me.spinDisconnectRetry.TabIndex = 59
        '
        'LabelControl18
        '
        Me.LabelControl18.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl18.Location = New System.Drawing.Point(593, 91)
        Me.LabelControl18.Name = "LabelControl18"
        Me.LabelControl18.Size = New System.Drawing.Size(192, 14)
        Me.LabelControl18.TabIndex = 60
        Me.LabelControl18.Text = "斷線後每                  秒，重新連接"
        '
        'spinReceiveTimeout
        '
        Me.spinReceiveTimeout.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinReceiveTimeout.EditValue = New Decimal(New Integer() {30, 0, 0, 0})
        Me.spinReceiveTimeout.Location = New System.Drawing.Point(279, 123)
        Me.spinReceiveTimeout.Name = "spinReceiveTimeout"
        Me.spinReceiveTimeout.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinReceiveTimeout.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinReceiveTimeout.Properties.IsFloatValue = False
        Me.spinReceiveTimeout.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.None
        Me.spinReceiveTimeout.Properties.MaxValue = New Decimal(New Integer() {600, 0, 0, 0})
        Me.spinReceiveTimeout.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinReceiveTimeout.Size = New System.Drawing.Size(69, 22)
        Me.spinReceiveTimeout.StyleController = Me.StyleController1
        Me.spinReceiveTimeout.TabIndex = 58
        '
        'LabelControl17
        '
        Me.LabelControl17.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl17.Location = New System.Drawing.Point(185, 127)
        Me.LabelControl17.Name = "LabelControl17"
        Me.LabelControl17.Size = New System.Drawing.Size(88, 14)
        Me.LabelControl17.TabIndex = 57
        Me.LabelControl17.Text = "ReceiveTimeout"
        '
        'spinSendTimeout
        '
        Me.spinSendTimeout.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinSendTimeout.EditValue = New Decimal(New Integer() {30, 0, 0, 0})
        Me.spinSendTimeout.Location = New System.Drawing.Point(109, 123)
        Me.spinSendTimeout.Name = "spinSendTimeout"
        Me.spinSendTimeout.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinSendTimeout.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinSendTimeout.Properties.IsFloatValue = False
        Me.spinSendTimeout.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.None
        Me.spinSendTimeout.Properties.MaxValue = New Decimal(New Integer() {600, 0, 0, 0})
        Me.spinSendTimeout.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinSendTimeout.Size = New System.Drawing.Size(69, 22)
        Me.spinSendTimeout.StyleController = Me.StyleController1
        Me.spinSendTimeout.TabIndex = 56
        '
        'LabelControl16
        '
        Me.LabelControl16.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl16.Location = New System.Drawing.Point(27, 127)
        Me.LabelControl16.Name = "LabelControl16"
        Me.LabelControl16.Size = New System.Drawing.Size(76, 14)
        Me.LabelControl16.TabIndex = 55
        Me.LabelControl16.Text = "SendTimeOut"
        '
        'chkCanRecord
        '
        Me.chkCanRecord.EditValue = CType(1, Short)
        Me.chkCanRecord.Location = New System.Drawing.Point(494, 88)
        Me.chkCanRecord.Name = "chkCanRecord"
        Me.chkCanRecord.Properties.Caption = "是否可錄影"
        Me.chkCanRecord.Properties.DisplayValueChecked = "1"
        Me.chkCanRecord.Properties.DisplayValueUnchecked = "0"
        Me.chkCanRecord.Properties.NullText = "0"
        Me.chkCanRecord.Properties.ValueChecked = CType(1, Short)
        Me.chkCanRecord.Properties.ValueGrayed = CType(0, Short)
        Me.chkCanRecord.Properties.ValueUnchecked = CType(0, Short)
        Me.chkCanRecord.Size = New System.Drawing.Size(93, 19)
        Me.chkCanRecord.TabIndex = 54
        '
        'cmbEncoding
        '
        Me.cmbEncoding.Location = New System.Drawing.Point(387, 86)
        Me.cmbEncoding.Name = "cmbEncoding"
        Me.cmbEncoding.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cmbEncoding.Size = New System.Drawing.Size(103, 20)
        Me.cmbEncoding.TabIndex = 53
        '
        'LabelControl12
        '
        Me.LabelControl12.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl12.Location = New System.Drawing.Point(333, 89)
        Me.LabelControl12.Name = "LabelControl12"
        Me.LabelControl12.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl12.TabIndex = 52
        Me.LabelControl12.Text = "編碼方式"
        '
        'cobKeyType
        '
        Me.cobKeyType.EditValue = "不加密"
        Me.cobKeyType.Location = New System.Drawing.Point(649, 51)
        Me.cobKeyType.Name = "cobKeyType"
        Me.cobKeyType.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cobKeyType.Properties.Items.AddRange(New Object() {"不加密", "會話加密"})
        Me.cobKeyType.Size = New System.Drawing.Size(91, 20)
        Me.cobKeyType.TabIndex = 51
        '
        'LabelControl11
        '
        Me.LabelControl11.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl11.Location = New System.Drawing.Point(582, 54)
        Me.LabelControl11.Name = "LabelControl11"
        Me.LabelControl11.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl11.TabIndex = 50
        Me.LabelControl11.Text = "加密方式"
        '
        'spinCrypt_Ver
        '
        Me.spinCrypt_Ver.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinCrypt_Ver.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinCrypt_Ver.Location = New System.Drawing.Point(88, 85)
        Me.spinCrypt_Ver.Name = "spinCrypt_Ver"
        Me.spinCrypt_Ver.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinCrypt_Ver.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinCrypt_Ver.Properties.IsFloatValue = False
        Me.spinCrypt_Ver.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.None
        Me.spinCrypt_Ver.Properties.MaxValue = New Decimal(New Integer() {9, 0, 0, 0})
        Me.spinCrypt_Ver.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinCrypt_Ver.Size = New System.Drawing.Size(69, 22)
        Me.spinCrypt_Ver.StyleController = Me.StyleController1
        Me.spinCrypt_Ver.TabIndex = 49
        '
        'LabelControl9
        '
        Me.LabelControl9.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl9.Location = New System.Drawing.Point(27, 89)
        Me.LabelControl9.Name = "LabelControl9"
        Me.LabelControl9.Size = New System.Drawing.Size(55, 14)
        Me.LabelControl9.TabIndex = 48
        Me.LabelControl9.Text = "Crypt_Ver"
        '
        'spinProtoVer
        '
        Me.spinProtoVer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinProtoVer.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinProtoVer.Location = New System.Drawing.Point(239, 85)
        Me.spinProtoVer.Name = "spinProtoVer"
        Me.spinProtoVer.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinProtoVer.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinProtoVer.Properties.IsFloatValue = False
        Me.spinProtoVer.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.None
        Me.spinProtoVer.Properties.MaxValue = New Decimal(New Integer() {9, 0, 0, 0})
        Me.spinProtoVer.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinProtoVer.Size = New System.Drawing.Size(69, 22)
        Me.spinProtoVer.StyleController = Me.StyleController1
        Me.spinProtoVer.TabIndex = 47
        '
        'LabelControl8
        '
        Me.LabelControl8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl8.Location = New System.Drawing.Point(177, 89)
        Me.LabelControl8.Name = "LabelControl8"
        Me.LabelControl8.Size = New System.Drawing.Size(56, 14)
        Me.LabelControl8.TabIndex = 46
        Me.LabelControl8.Text = "Proto_Ver"
        '
        'txtOPE_ID
        '
        Me.txtOPE_ID.EditValue = ""
        Me.txtOPE_ID.Location = New System.Drawing.Point(508, 50)
        Me.txtOPE_ID.Name = "txtOPE_ID"
        Me.txtOPE_ID.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtOPE_ID.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtOPE_ID.Properties.Appearance.Options.UseBackColor = True
        Me.txtOPE_ID.Properties.Appearance.Options.UseForeColor = True
        Me.txtOPE_ID.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtOPE_ID.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtOPE_ID.Size = New System.Drawing.Size(57, 22)
        Me.txtOPE_ID.StyleController = Me.StyleController1
        Me.txtOPE_ID.TabIndex = 45
        '
        'LabelControl6
        '
        Me.LabelControl6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl6.Location = New System.Drawing.Point(448, 54)
        Me.LabelControl6.Name = "LabelControl6"
        Me.LabelControl6.Size = New System.Drawing.Size(42, 14)
        Me.LabelControl6.TabIndex = 44
        Me.LabelControl6.Text = "OPE_ID"
        '
        'spinSndDelay
        '
        Me.spinSndDelay.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinSndDelay.EditValue = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.spinSndDelay.Location = New System.Drawing.Point(612, 178)
        Me.spinSndDelay.Name = "spinSndDelay"
        Me.spinSndDelay.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinSndDelay.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinSndDelay.Properties.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.spinSndDelay.Properties.MaxValue = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.spinSndDelay.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.spinSndDelay.Size = New System.Drawing.Size(49, 22)
        Me.spinSndDelay.StyleController = Me.StyleController1
        Me.spinSndDelay.TabIndex = 3
        Me.spinSndDelay.Visible = False
        '
        'LabelControl5
        '
        Me.LabelControl5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl5.Location = New System.Drawing.Point(460, 182)
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
        Me.spinChkStatus.Location = New System.Drawing.Point(45, 50)
        Me.spinChkStatus.Name = "spinChkStatus"
        Me.spinChkStatus.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinChkStatus.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinChkStatus.Properties.IsFloatValue = False
        Me.spinChkStatus.Properties.Mask.EditMask = "N00"
        Me.spinChkStatus.Properties.MaxValue = New Decimal(New Integer() {1200, 0, 0, 0})
        Me.spinChkStatus.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinChkStatus.Size = New System.Drawing.Size(49, 22)
        Me.spinChkStatus.StyleController = Me.StyleController1
        Me.spinChkStatus.TabIndex = 4
        '
        'LabelControl4
        '
        Me.LabelControl4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl4.Location = New System.Drawing.Point(27, 54)
        Me.LabelControl4.Name = "LabelControl4"
        Me.LabelControl4.Size = New System.Drawing.Size(262, 14)
        Me.LabelControl4.TabIndex = 41
        Me.LabelControl4.Text = "每                 秒傳送一次 Commnucation Check"
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
        Me.txtRcvMopPPId.StyleController = Me.StyleController1
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
        Me.txtRcvDestId.StyleController = Me.StyleController1
        Me.txtRcvDestId.TabIndex = 34
        Me.txtRcvDestId.Visible = False
        '
        'LabelControl15
        '
        Me.LabelControl15.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl15.Location = New System.Drawing.Point(603, 206)
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
        Me.txtRcvSourceId.StyleController = Me.StyleController1
        Me.txtRcvSourceId.TabIndex = 32
        Me.txtRcvSourceId.Visible = False
        '
        'LabelControl14
        '
        Me.LabelControl14.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl14.Location = New System.Drawing.Point(433, 206)
        Me.LabelControl14.Name = "LabelControl14"
        Me.LabelControl14.Size = New System.Drawing.Size(57, 14)
        Me.LabelControl14.TabIndex = 31
        Me.LabelControl14.Text = "Source  Id"
        Me.LabelControl14.Visible = False
        '
        'LabelControl13
        '
        Me.LabelControl13.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl13.Location = New System.Drawing.Point(368, 206)
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
        Me.txtSndMopPPId.StyleController = Me.StyleController1
        Me.txtSndMopPPId.TabIndex = 29
        Me.txtSndMopPPId.Visible = False
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
        Me.txtSndDestId.StyleController = Me.StyleController1
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
        Me.txtSndSourceId.StyleController = Me.StyleController1
        Me.txtSndSourceId.TabIndex = 24
        Me.txtSndSourceId.Visible = False
        '
        'spinNstvPort
        '
        Me.spinNstvPort.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinNstvPort.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.spinNstvPort.Location = New System.Drawing.Point(786, 178)
        Me.spinNstvPort.Name = "spinNstvPort"
        Me.spinNstvPort.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinNstvPort.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinNstvPort.Properties.IsFloatValue = False
        Me.spinNstvPort.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.None
        Me.spinNstvPort.Size = New System.Drawing.Size(69, 22)
        Me.spinNstvPort.StyleController = Me.StyleController1
        Me.spinNstvPort.TabIndex = 7
        Me.spinNstvPort.Visible = False
        '
        'LabelControl10
        '
        Me.LabelControl10.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl10.Location = New System.Drawing.Point(715, 186)
        Me.LabelControl10.Name = "LabelControl10"
        Me.LabelControl10.Size = New System.Drawing.Size(84, 14)
        Me.LabelControl10.TabIndex = 21
        Me.LabelControl10.Text = "傳送指令通訊埠"
        Me.LabelControl10.Visible = False
        '
        'txtSMS_ID
        '
        Me.txtSMS_ID.EditValue = ""
        Me.txtSMS_ID.Location = New System.Drawing.Point(368, 50)
        Me.txtSMS_ID.Name = "txtSMS_ID"
        Me.txtSMS_ID.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtSMS_ID.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtSMS_ID.Properties.Appearance.Options.UseBackColor = True
        Me.txtSMS_ID.Properties.Appearance.Options.UseForeColor = True
        Me.txtSMS_ID.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtSMS_ID.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtSMS_ID.Size = New System.Drawing.Size(62, 22)
        Me.txtSMS_ID.StyleController = Me.StyleController1
        Me.txtSMS_ID.TabIndex = 6
        '
        'LabelControl3
        '
        Me.LabelControl3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl3.Location = New System.Drawing.Point(320, 54)
        Me.LabelControl3.Name = "LabelControl3"
        Me.LabelControl3.Size = New System.Drawing.Size(42, 14)
        Me.LabelControl3.TabIndex = 19
        Me.LabelControl3.Text = "SMS_ID"
        '
        'XtabLowCmd
        '
        Me.XtabLowCmd.Controls.Add(Me.TreLstLowCmd)
        Me.XtabLowCmd.Controls.Add(Me.btnDelLowCmd)
        Me.XtabLowCmd.Controls.Add(Me.btnAddLowCmd)
        Me.XtabLowCmd.Image = CType(resources.GetObject("XtabLowCmd.Image"), System.Drawing.Image)
        Me.XtabLowCmd.Name = "XtabLowCmd"
        Me.XtabLowCmd.Size = New System.Drawing.Size(896, 481)
        Me.XtabLowCmd.Text = "低 階 命 令 設 定 "
        '
        'TreLstLowCmd
        '
        Me.TreLstLowCmd.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colCode, Me.colValue, Me.colDescription, Me.colDefaultValue, Me.colRealField})
        Me.TreLstLowCmd.Location = New System.Drawing.Point(24, 52)
        Me.TreLstLowCmd.Name = "TreLstLowCmd"
        Me.TreLstLowCmd.OptionsPrint.UsePrintStyles = True
        Me.TreLstLowCmd.OptionsView.ShowButtons = False
        Me.TreLstLowCmd.OptionsView.ShowRoot = False
        Me.TreLstLowCmd.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.RepositoryItemCheckEdit6})
        Me.TreLstLowCmd.Size = New System.Drawing.Size(840, 410)
        Me.TreLstLowCmd.TabIndex = 24
        '
        'colCode
        '
        Me.colCode.Caption = "代碼"
        Me.colCode.FieldName = "Code"
        Me.colCode.MinWidth = 33
        Me.colCode.Name = "colCode"
        Me.colCode.Visible = True
        Me.colCode.VisibleIndex = 0
        Me.colCode.Width = 69
        '
        'colValue
        '
        Me.colValue.Caption = "數值"
        Me.colValue.FieldName = "Value"
        Me.colValue.Name = "colValue"
        Me.colValue.Visible = True
        Me.colValue.VisibleIndex = 1
        Me.colValue.Width = 99
        '
        'colDescription
        '
        Me.colDescription.Caption = "名稱"
        Me.colDescription.FieldName = "Description"
        Me.colDescription.Name = "colDescription"
        Me.colDescription.Visible = True
        Me.colDescription.VisibleIndex = 2
        Me.colDescription.Width = 338
        '
        'colDefaultValue
        '
        Me.colDefaultValue.Caption = "預設值"
        Me.colDefaultValue.FieldName = "DefaultValue"
        Me.colDefaultValue.Name = "colDefaultValue"
        Me.colDefaultValue.Visible = True
        Me.colDefaultValue.VisibleIndex = 3
        Me.colDefaultValue.Width = 114
        '
        'colRealField
        '
        Me.colRealField.Caption = "實際資料欄位"
        Me.colRealField.FieldName = "RealField"
        Me.colRealField.Name = "colRealField"
        Me.colRealField.Visible = True
        Me.colRealField.VisibleIndex = 4
        Me.colRealField.Width = 202
        '
        'RepositoryItemCheckEdit6
        '
        Me.RepositoryItemCheckEdit6.AutoHeight = False
        Me.RepositoryItemCheckEdit6.Name = "RepositoryItemCheckEdit6"
        Me.RepositoryItemCheckEdit6.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.RepositoryItemCheckEdit6.ValueChecked = "1"
        Me.RepositoryItemCheckEdit6.ValueUnchecked = "0"
        '
        'btnDelLowCmd
        '
        Me.btnDelLowCmd.Image = CType(resources.GetObject("btnDelLowCmd.Image"), System.Drawing.Image)
        Me.btnDelLowCmd.Location = New System.Drawing.Point(95, 20)
        Me.btnDelLowCmd.Name = "btnDelLowCmd"
        Me.btnDelLowCmd.Size = New System.Drawing.Size(65, 26)
        Me.btnDelLowCmd.TabIndex = 23
        Me.btnDelLowCmd.Text = "刪 除"
        '
        'btnAddLowCmd
        '
        Me.btnAddLowCmd.Image = CType(resources.GetObject("btnAddLowCmd.Image"), System.Drawing.Image)
        Me.btnAddLowCmd.Location = New System.Drawing.Point(24, 20)
        Me.btnAddLowCmd.Name = "btnAddLowCmd"
        Me.btnAddLowCmd.Size = New System.Drawing.Size(65, 26)
        Me.btnAddLowCmd.TabIndex = 22
        Me.btnAddLowCmd.Text = "新 增"
        '
        'XtabHighCmd
        '
        Me.XtabHighCmd.Controls.Add(Me.TreLstHightCmd)
        Me.XtabHighCmd.Controls.Add(Me.btnDelHighCmd)
        Me.XtabHighCmd.Controls.Add(Me.btnAddHightCmd)
        Me.XtabHighCmd.Image = CType(resources.GetObject("XtabHighCmd.Image"), System.Drawing.Image)
        Me.XtabHighCmd.Name = "XtabHighCmd"
        Me.XtabHighCmd.Size = New System.Drawing.Size(896, 481)
        Me.XtabHighCmd.Text = "高 階 命 令 設 定"
        '
        'TreLstHightCmd
        '
        Me.TreLstHightCmd.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colCodeNo, Me.colHightDesc, Me.colRunLowCmd, Me.colRunLowCmd2})
        Me.TreLstHightCmd.Location = New System.Drawing.Point(21, 51)
        Me.TreLstHightCmd.Name = "TreLstHightCmd"
        Me.TreLstHightCmd.OptionsPrint.UsePrintStyles = True
        Me.TreLstHightCmd.OptionsView.ShowButtons = False
        Me.TreLstHightCmd.OptionsView.ShowRoot = False
        Me.TreLstHightCmd.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.RepositoryItemCheckEdit7})
        Me.TreLstHightCmd.Size = New System.Drawing.Size(840, 410)
        Me.TreLstHightCmd.TabIndex = 26
        '
        'colCodeNo
        '
        Me.colCodeNo.Caption = "代碼"
        Me.colCodeNo.FieldName = "CodeNo"
        Me.colCodeNo.MinWidth = 33
        Me.colCodeNo.Name = "colCodeNo"
        Me.colCodeNo.Visible = True
        Me.colCodeNo.VisibleIndex = 0
        Me.colCodeNo.Width = 68
        '
        'colHightDesc
        '
        Me.colHightDesc.Caption = "名稱"
        Me.colHightDesc.FieldName = "HightDesc"
        Me.colHightDesc.Name = "colHightDesc"
        Me.colHightDesc.Visible = True
        Me.colHightDesc.VisibleIndex = 1
        Me.colHightDesc.Width = 248
        '
        'colRunLowCmd
        '
        Me.colRunLowCmd.Caption = "執行低階命令(第一組)"
        Me.colRunLowCmd.FieldName = "RunLowCmd"
        Me.colRunLowCmd.Name = "colRunLowCmd"
        Me.colRunLowCmd.Visible = True
        Me.colRunLowCmd.VisibleIndex = 2
        Me.colRunLowCmd.Width = 248
        '
        'colRunLowCmd2
        '
        Me.colRunLowCmd2.Caption = "執行低階命令(第二組)"
        Me.colRunLowCmd2.FieldName = "RunLowCmd2"
        Me.colRunLowCmd2.Name = "colRunLowCmd2"
        Me.colRunLowCmd2.Visible = True
        Me.colRunLowCmd2.VisibleIndex = 3
        Me.colRunLowCmd2.Width = 258
        '
        'RepositoryItemCheckEdit7
        '
        Me.RepositoryItemCheckEdit7.AutoHeight = False
        Me.RepositoryItemCheckEdit7.Name = "RepositoryItemCheckEdit7"
        Me.RepositoryItemCheckEdit7.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.RepositoryItemCheckEdit7.ValueChecked = "1"
        Me.RepositoryItemCheckEdit7.ValueUnchecked = "0"
        '
        'btnDelHighCmd
        '
        Me.btnDelHighCmd.Image = CType(resources.GetObject("btnDelHighCmd.Image"), System.Drawing.Image)
        Me.btnDelHighCmd.Location = New System.Drawing.Point(92, 19)
        Me.btnDelHighCmd.Name = "btnDelHighCmd"
        Me.btnDelHighCmd.Size = New System.Drawing.Size(65, 26)
        Me.btnDelHighCmd.TabIndex = 25
        Me.btnDelHighCmd.Text = "刪 除"
        '
        'btnAddHightCmd
        '
        Me.btnAddHightCmd.Image = CType(resources.GetObject("btnAddHightCmd.Image"), System.Drawing.Image)
        Me.btnAddHightCmd.Location = New System.Drawing.Point(21, 19)
        Me.btnAddHightCmd.Name = "btnAddHightCmd"
        Me.btnAddHightCmd.Size = New System.Drawing.Size(65, 26)
        Me.btnAddHightCmd.TabIndex = 24
        Me.btnAddHightCmd.Text = "新 增"
        '
        'XtabQtvAuth
        '
        Me.XtabQtvAuth.Controls.Add(Me.TreLstErrCode)
        Me.XtabQtvAuth.Controls.Add(Me.btnRemoveAuth)
        Me.XtabQtvAuth.Controls.Add(Me.btnAddError)
        Me.XtabQtvAuth.Image = CType(resources.GetObject("XtabQtvAuth.Image"), System.Drawing.Image)
        Me.XtabQtvAuth.Name = "XtabQtvAuth"
        Me.XtabQtvAuth.Size = New System.Drawing.Size(896, 481)
        Me.XtabQtvAuth.Text = " 錯 誤 代 碼 表 設 定 "
        '
        'TreLstErrCode
        '
        Me.TreLstErrCode.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colErrorCode, Me.colErrorName})
        Me.TreLstErrCode.Location = New System.Drawing.Point(14, 49)
        Me.TreLstErrCode.Name = "TreLstErrCode"
        Me.TreLstErrCode.OptionsPrint.UsePrintStyles = True
        Me.TreLstErrCode.OptionsView.ShowButtons = False
        Me.TreLstErrCode.OptionsView.ShowRoot = False
        Me.TreLstErrCode.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.RepositoryItemCheckEdit5})
        Me.TreLstErrCode.Size = New System.Drawing.Size(840, 410)
        Me.TreLstErrCode.TabIndex = 23
        '
        'colErrorCode
        '
        Me.colErrorCode.Caption = "錯誤代碼"
        Me.colErrorCode.FieldName = "ErrorCode"
        Me.colErrorCode.MinWidth = 33
        Me.colErrorCode.Name = "colErrorCode"
        Me.colErrorCode.Visible = True
        Me.colErrorCode.VisibleIndex = 0
        Me.colErrorCode.Width = 65
        '
        'colErrorName
        '
        Me.colErrorName.Caption = "錯誤名稱"
        Me.colErrorName.FieldName = "ErrorName"
        Me.colErrorName.Name = "colErrorName"
        Me.colErrorName.Visible = True
        Me.colErrorName.VisibleIndex = 1
        Me.colErrorName.Width = 550
        '
        'RepositoryItemCheckEdit5
        '
        Me.RepositoryItemCheckEdit5.AutoHeight = False
        Me.RepositoryItemCheckEdit5.Name = "RepositoryItemCheckEdit5"
        Me.RepositoryItemCheckEdit5.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.RepositoryItemCheckEdit5.ValueChecked = "1"
        Me.RepositoryItemCheckEdit5.ValueUnchecked = "0"
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
        'btnAddError
        '
        Me.btnAddError.Image = CType(resources.GetObject("btnAddError.Image"), System.Drawing.Image)
        Me.btnAddError.Location = New System.Drawing.Point(14, 17)
        Me.btnAddError.Name = "btnAddError"
        Me.btnAddError.Size = New System.Drawing.Size(65, 26)
        Me.btnAddError.TabIndex = 20
        Me.btnAddError.Text = "新 增"
        '
        'XtabSMailInfo
        '
        Me.XtabSMailInfo.Controls.Add(Me.btnDelSMailCC)
        Me.XtabSMailInfo.Controls.Add(Me.btnInsSMailCC)
        Me.XtabSMailInfo.Controls.Add(Me.btnDelSMailTo)
        Me.XtabSMailInfo.Controls.Add(Me.btnInsSMailTo)
        Me.XtabSMailInfo.Controls.Add(Me.GroupControl1)
        Me.XtabSMailInfo.Image = CType(resources.GetObject("XtabSMailInfo.Image"), System.Drawing.Image)
        Me.XtabSMailInfo.Name = "XtabSMailInfo"
        Me.XtabSMailInfo.PageVisible = False
        Me.XtabSMailInfo.Size = New System.Drawing.Size(896, 481)
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
        Me.txtMailSubject.Location = New System.Drawing.Point(177, 95)
        Me.txtMailSubject.Name = "txtMailSubject"
        Me.txtMailSubject.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.txtMailSubject.Properties.Appearance.ForeColor = System.Drawing.Color.Black
        Me.txtMailSubject.Properties.Appearance.Options.UseBackColor = True
        Me.txtMailSubject.Properties.Appearance.Options.UseForeColor = True
        Me.txtMailSubject.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtMailSubject.Properties.NullValuePromptShowForEmptyValue = True
        Me.txtMailSubject.Size = New System.Drawing.Size(538, 22)
        Me.txtMailSubject.StyleController = Me.StyleController1
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
        Me.txtDisplayName.StyleController = Me.StyleController1
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
        Me.txtSMTPLoginPsw.StyleController = Me.StyleController1
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
        Me.txtSMTPLoginID.StyleController = Me.StyleController1
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
        Me.txtSender.StyleController = Me.StyleController1
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
        Me.spinSMTPPort.StyleController = Me.StyleController1
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
        Me.txtSMTPHost.StyleController = Me.StyleController1
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
        Me.spinSocketErrCnt.StyleController = Me.StyleController1
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
        Me.TextEdit3.Size = New System.Drawing.Size(78, 22)
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
        Me.TextEdit4.Size = New System.Drawing.Size(78, 22)
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
        Me.TextEdit5.Size = New System.Drawing.Size(78, 22)
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
        'rpstrItmTxtEdtPwd
        '
        Me.rpstrItmTxtEdtPwd.AutoHeight = False
        Me.rpstrItmTxtEdtPwd.Name = "rpstrItmTxtEdtPwd"
        Me.rpstrItmTxtEdtPwd.NullText = "未設定密碼"
        Me.rpstrItmTxtEdtPwd.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
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
        'RepositoryItemTextEdit2
        '
        Me.RepositoryItemTextEdit2.AutoHeight = False
        Me.RepositoryItemTextEdit2.Mask.EditMask = "\d{0,5}"
        Me.RepositoryItemTextEdit2.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.RegEx
        Me.RepositoryItemTextEdit2.MaxLength = 5
        Me.RepositoryItemTextEdit2.Name = "RepositoryItemTextEdit2"
        Me.RepositoryItemTextEdit2.NullText = "00000"
        '
        'chkChoice
        '
        Me.chkChoice.AutoHeight = False
        Me.chkChoice.Name = "chkChoice"
        Me.chkChoice.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.chkChoice.ValueChecked = "1"
        Me.chkChoice.ValueUnchecked = "0"
        '
        'RepositoryItemCheckEdit1
        '
        Me.RepositoryItemCheckEdit1.AutoHeight = False
        Me.RepositoryItemCheckEdit1.DisplayValueChecked = "1"
        Me.RepositoryItemCheckEdit1.DisplayValueUnchecked = "0"
        Me.RepositoryItemCheckEdit1.Name = "RepositoryItemCheckEdit1"
        Me.RepositoryItemCheckEdit1.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.RepositoryItemCheckEdit1.ValueChecked = "1"
        Me.RepositoryItemCheckEdit1.ValueUnchecked = "0"
        '
        'RepositoryItemSpinEdit1
        '
        Me.RepositoryItemSpinEdit1.AllowNullInput = DevExpress.Utils.DefaultBoolean.[False]
        Me.RepositoryItemSpinEdit1.Appearance.Options.UseTextOptions = True
        Me.RepositoryItemSpinEdit1.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.RepositoryItemSpinEdit1.AutoHeight = False
        Me.RepositoryItemSpinEdit1.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.RepositoryItemSpinEdit1.IsFloatValue = False
        Me.RepositoryItemSpinEdit1.Mask.EditMask = "N00"
        Me.RepositoryItemSpinEdit1.MaxLength = 2
        Me.RepositoryItemSpinEdit1.MaxValue = New Decimal(New Integer() {30, 0, 0, 0})
        Me.RepositoryItemSpinEdit1.Name = "RepositoryItemSpinEdit1"
        '
        'RepositoryItemTextEdit3
        '
        Me.RepositoryItemTextEdit3.AutoHeight = False
        Me.RepositoryItemTextEdit3.Name = "RepositoryItemTextEdit3"
        Me.RepositoryItemTextEdit3.NullText = "未設定密碼"
        Me.RepositoryItemTextEdit3.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        '
        'RepositoryItemCheckEdit2
        '
        Me.RepositoryItemCheckEdit2.AutoHeight = False
        Me.RepositoryItemCheckEdit2.DisplayValueChecked = "1"
        Me.RepositoryItemCheckEdit2.DisplayValueUnchecked = "0"
        Me.RepositoryItemCheckEdit2.Name = "RepositoryItemCheckEdit2"
        Me.RepositoryItemCheckEdit2.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.RepositoryItemCheckEdit2.ValueChecked = "1"
        Me.RepositoryItemCheckEdit2.ValueUnchecked = "0"
        '
        'RepositoryItemSpinEdit2
        '
        Me.RepositoryItemSpinEdit2.AllowNullInput = DevExpress.Utils.DefaultBoolean.[False]
        Me.RepositoryItemSpinEdit2.Appearance.Options.UseTextOptions = True
        Me.RepositoryItemSpinEdit2.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.RepositoryItemSpinEdit2.AutoHeight = False
        Me.RepositoryItemSpinEdit2.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.RepositoryItemSpinEdit2.IsFloatValue = False
        Me.RepositoryItemSpinEdit2.Mask.EditMask = "N00"
        Me.RepositoryItemSpinEdit2.MaxLength = 2
        Me.RepositoryItemSpinEdit2.MaxValue = New Decimal(New Integer() {30, 0, 0, 0})
        Me.RepositoryItemSpinEdit2.Name = "RepositoryItemSpinEdit2"
        '
        'RepositoryItemTextEdit4
        '
        Me.RepositoryItemTextEdit4.AutoHeight = False
        Me.RepositoryItemTextEdit4.Name = "RepositoryItemTextEdit4"
        Me.RepositoryItemTextEdit4.NullText = "未設定密碼"
        Me.RepositoryItemTextEdit4.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        '
        'RepositoryItemCheckEdit3
        '
        Me.RepositoryItemCheckEdit3.AutoHeight = False
        Me.RepositoryItemCheckEdit3.DisplayValueChecked = "1"
        Me.RepositoryItemCheckEdit3.DisplayValueUnchecked = "0"
        Me.RepositoryItemCheckEdit3.Name = "RepositoryItemCheckEdit3"
        Me.RepositoryItemCheckEdit3.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.RepositoryItemCheckEdit3.ValueChecked = "1"
        Me.RepositoryItemCheckEdit3.ValueUnchecked = "0"
        '
        'RepositoryItemSpinEdit3
        '
        Me.RepositoryItemSpinEdit3.AllowNullInput = DevExpress.Utils.DefaultBoolean.[False]
        Me.RepositoryItemSpinEdit3.Appearance.Options.UseTextOptions = True
        Me.RepositoryItemSpinEdit3.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.RepositoryItemSpinEdit3.AutoHeight = False
        Me.RepositoryItemSpinEdit3.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.RepositoryItemSpinEdit3.IsFloatValue = False
        Me.RepositoryItemSpinEdit3.Mask.EditMask = "N00"
        Me.RepositoryItemSpinEdit3.MaxLength = 2
        Me.RepositoryItemSpinEdit3.MaxValue = New Decimal(New Integer() {30, 0, 0, 0})
        Me.RepositoryItemSpinEdit3.Name = "RepositoryItemSpinEdit3"
        '
        'RepositoryItemTextEdit5
        '
        Me.RepositoryItemTextEdit5.AutoHeight = False
        Me.RepositoryItemTextEdit5.Name = "RepositoryItemTextEdit5"
        Me.RepositoryItemTextEdit5.NullText = "未設定密碼"
        Me.RepositoryItemTextEdit5.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        '
        'frmSetting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(926, 610)
        Me.Controls.Add(Me.PanelControl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmSetting"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Gateway參數設定"
        CType(Me.StyleController1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        CType(Me.grpCtl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpCtl.ResumeLayout(False)
        CType(Me.XtabCtl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabCtl.ResumeLayout(False)
        Me.XtabSysOpt.ResumeLayout(False)
        CType(Me.grpQtv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpQtv.ResumeLayout(False)
        Me.grpQtv.PerformLayout()
        CType(Me.spinPort.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtIPAddress.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinTimeout.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCredPswd.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCredUser.Properties, System.ComponentModel.ISupportInitialize).EndInit()
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
        CType(Me.spinReCon.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtGwPswd.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkAutoRun.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabDbOpt.ResumeLayout(False)
        CType(Me.TreLstDB, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemSpinEdit4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemTextEdit6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpComDB, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpComDB.ResumeLayout(False)
        Me.grpComDB.PerformLayout()
        CType(Me.txtCOMPassword.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCOMUserId.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCOMSid.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabQtvWS.ResumeLayout(False)
        CType(Me.TreLstSourceId, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rpedtIP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rpedtPort, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rptedtNstvRootKey, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpNagra, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpNagra.ResumeLayout(False)
        Me.grpNagra.PerformLayout()
        CType(Me.txtEmailTitle.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cobProcKind.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinRetryCount.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkUseUTC.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinDisconnectRetry.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinReceiveTimeout.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinSendTimeout.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkCanRecord.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmbEncoding.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cobKeyType.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinCrypt_Ver.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinProtoVer.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtOPE_ID.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinSndDelay.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinChkStatus.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtRcvMopPPId.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtRcvDestId.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtRcvSourceId.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSndMopPPId.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSndDestId.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSndSourceId.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinNstvPort.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSMS_ID.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabLowCmd.ResumeLayout(False)
        CType(Me.TreLstLowCmd, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit6, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabHighCmd.ResumeLayout(False)
        CType(Me.TreLstHightCmd, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabQtvAuth.ResumeLayout(False)
        CType(Me.TreLstErrCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit5, System.ComponentModel.ISupportInitialize).EndInit()
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
        CType(Me.rpstrItmChkEdt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rpstrItmSpnEdt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rpstrItmTxtEdtPwd, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemTextEdit1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemTextEdit2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkChoice, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemSpinEdit1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemTextEdit3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemSpinEdit2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemTextEdit4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemSpinEdit3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemTextEdit5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DefaultLookAndFeel1 As DevExpress.LookAndFeel.DefaultLookAndFeel
    Friend WithEvents StyleController1 As DevExpress.XtraEditors.StyleController
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents grpCtl As DevExpress.XtraEditors.GroupControl
    Friend WithEvents btnCancel As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnOK As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents XtabCtl As DevExpress.XtraTab.XtraTabControl
    Friend WithEvents XtabSysOpt As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents grpQtv As DevExpress.XtraEditors.GroupControl
    Friend WithEvents spinPort As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl2 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtIPAddress As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinTimeout As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents txtCredPswd As DevExpress.XtraEditors.TextEdit
    Friend WithEvents txtCredUser As DevExpress.XtraEditors.TextEdit
    Friend WithEvents lblTimeout As DevExpress.XtraEditors.LabelControl
    Friend WithEvents lblCred As DevExpress.XtraEditors.LabelControl
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
    Friend WithEvents spinReCon As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents lblDisp As DevExpress.XtraEditors.LabelControl
    Friend WithEvents lblDuration As DevExpress.XtraEditors.LabelControl
    Friend WithEvents lblReConn As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtGwPswd As DevExpress.XtraEditors.TextEdit
    Friend WithEvents lblGwPwd As DevExpress.XtraEditors.LabelControl
    Friend WithEvents chkAutoRun As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents XtabDbOpt As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents grpComDB As DevExpress.XtraEditors.GroupControl
    Friend WithEvents txtCOMPassword As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl19 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtCOMUserId As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl7 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtCOMSid As DevExpress.XtraEditors.TextEdit
    Friend WithEvents lblComSid As DevExpress.XtraEditors.LabelControl
    Friend WithEvents lblMsg As DevExpress.XtraEditors.LabelControl
    Friend WithEvents btnTestConn As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnRemoveConn As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAddConn As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents XtabQtvWS As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents btnbtnRemoveSourceId As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAddSourceId As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents grpNagra As DevExpress.XtraEditors.GroupControl
    Friend WithEvents spinSndDelay As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl5 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinChkStatus As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl4 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtRcvMopPPId As DevExpress.XtraEditors.TextEdit
    Friend WithEvents txtRcvDestId As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl15 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtRcvSourceId As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl14 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl13 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtSndMopPPId As DevExpress.XtraEditors.TextEdit
    Friend WithEvents txtSndDestId As DevExpress.XtraEditors.TextEdit
    Friend WithEvents txtSndSourceId As DevExpress.XtraEditors.TextEdit
    Friend WithEvents spinNstvPort As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl10 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtSMS_ID As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl3 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents XtabQtvAuth As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents btnRemoveAuth As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAddError As DevExpress.XtraEditors.SimpleButton
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
    Friend WithEvents rpstrItmChkEdt As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents rpstrItmSpnEdt As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents rpstrItmTxtEdtPwd As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents RepositoryItemTextEdit1 As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents RepositoryItemTextEdit2 As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents chkChoice As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents TreLstDB As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colSelSO As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents RepositoryItemCheckEdit4 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents colCompCode As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents RepositoryItemSpinEdit4 As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents colCompName As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colDBSid As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colDBUser As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colDBPassword As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents RepositoryItemTextEdit6 As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents TreLstSourceId As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colServer As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents rpedtIP As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents colPort As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents rpedtPort As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents RepositoryItemCheckEdit1 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents RepositoryItemSpinEdit1 As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents RepositoryItemTextEdit3 As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents RepositoryItemCheckEdit2 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents RepositoryItemSpinEdit2 As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents RepositoryItemTextEdit4 As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents RepositoryItemCheckEdit3 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents RepositoryItemSpinEdit3 As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents RepositoryItemTextEdit5 As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents txtOPE_ID As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl6 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl8 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinProtoVer As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl9 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinCrypt_Ver As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl11 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cobKeyType As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents TreLstErrCode As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colErrorCode As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colErrorName As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents RepositoryItemCheckEdit5 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents XtabLowCmd As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents TreLstLowCmd As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colCode As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colValue As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colDescription As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents RepositoryItemCheckEdit6 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents btnDelLowCmd As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAddLowCmd As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents XtabHighCmd As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents TreLstHightCmd As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colCodeNo As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colHightDesc As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colRunLowCmd As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents RepositoryItemCheckEdit7 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents btnDelHighCmd As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAddHightCmd As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents ColRootKey As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents rptedtNstvRootKey As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents cmbEncoding As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LabelControl12 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinDisconnectRetry As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl18 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinReceiveTimeout As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl17 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinSendTimeout As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl16 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents chkCanRecord As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents chkUseUTC As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents spinRetryCount As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl21 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents colDefaultValue As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colRealField As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents cobProcKind As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LabelControl34 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents colRunLowCmd2 As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents txtEmailTitle As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl35 As DevExpress.XtraEditors.LabelControl
End Class
