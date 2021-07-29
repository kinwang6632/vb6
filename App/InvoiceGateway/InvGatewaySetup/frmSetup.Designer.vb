<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSetup
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSetup))
        Dim SuperToolTip1 As DevExpress.Utils.SuperToolTip = New DevExpress.Utils.SuperToolTip()
        Dim ToolTipTitleItem1 As DevExpress.Utils.ToolTipTitleItem = New DevExpress.Utils.ToolTipTitleItem()
        Dim ToolTipItem1 As DevExpress.Utils.ToolTipItem = New DevExpress.Utils.ToolTipItem()
        Dim ToolTipSeparatorItem1 As DevExpress.Utils.ToolTipSeparatorItem = New DevExpress.Utils.ToolTipSeparatorItem()
        Dim ToolTipTitleItem2 As DevExpress.Utils.ToolTipTitleItem = New DevExpress.Utils.ToolTipTitleItem()
        Me.richkCboSendOrd = New DevExpress.XtraEditors.Repository.RepositoryItemCheckedComboBoxEdit()
        Me.DefaultLookAndFeel1 = New DevExpress.LookAndFeel.DefaultLookAndFeel()
        Me.StyleController1 = New DevExpress.XtraEditors.StyleController()
        Me.ImageList1 = New System.Windows.Forms.ImageList()
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
        Me.spinReserveLog = New DevExpress.XtraEditors.SpinEdit()
        Me.chkLogSQL = New DevExpress.XtraEditors.CheckEdit()
        Me.txtDestroyReason = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl5 = New DevExpress.XtraEditors.LabelControl()
        Me.txtLoginUserName = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl4 = New DevExpress.XtraEditors.LabelControl()
        Me.txtLoginUsreId = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl3 = New DevExpress.XtraEditors.LabelControl()
        Me.spinMaxThreads = New DevExpress.XtraEditors.SpinEdit()
        Me.lblMaxThreads = New DevExpress.XtraEditors.LabelControl()
        Me.TreLstDB = New DevExpress.XtraTreeList.TreeList()
        Me.colSelSO = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.chkSelect = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.colCompCode = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.spnNum = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.colCompName = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colInvBackupPath = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.riBtnEdtPath = New DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit()
        Me.colDBSid = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colDBUser = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colDBPassword = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.riTxtPassword = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.colComDBSid = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colComDBUser = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colComDBPassword = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colMisDBSid = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colMisDBUser = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colMisDBPassword = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colMisOwner = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colDBLink = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colMinDBPool = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colMaxDBPool = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colDBPoolLifetime = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colRunCmdTimer = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colLimtBeforeUpload = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colInvoiceType = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.richkCboInvoiceType = New DevExpress.XtraEditors.Repository.RepositoryItemCheckedComboBoxEdit()
        Me.colRunCommands = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colIsEmailNotify = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colIsSmsNotify = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colIsTvMailNotify = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colIsCMNotify = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colSendOrd = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.ritxtSendOrd = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.colAddSend = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colDonateMark = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.ricboDonateMark = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox()
        Me.colExportPath = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colAutoDo = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.riCboAutoDo = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox()
        Me.colStartEmailNotify = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colStartSMSNotify = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colStartTVMailNotify = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colStartCMNotify = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colEndCMNotify = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colCompType = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.riCboCompType = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox()
        Me.colMaskInvNo = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colStartNotifyPrize = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.richkRunCmdCode = New DevExpress.XtraEditors.Repository.RepositoryItemCheckedComboBoxEdit()
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
        Me.XtabLowCmd = New DevExpress.XtraTab.XtraTabPage()
        Me.TreLstCmd = New DevExpress.XtraTreeList.TreeList()
        Me.colCodeNo = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colCmdName = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.rCboCmdName = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox()
        Me.btnDelLowCmd = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAddLowCmd = New DevExpress.XtraEditors.SimpleButton()
        Me.XtabHighCmd = New DevExpress.XtraTab.XtraTabPage()
        Me.TreLstSORunCmd = New DevExpress.XtraTreeList.TreeList()
        Me.colCompID = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.riSpnNum = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.colCompDescript = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.riTxtCompDescript = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.colCmdType = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.riCboCmdType = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox()
        Me.colCmdTypeName = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colRunfrequency = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colExportElectronPath = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.ribtnElectronPath = New DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit()
        Me.riCboElectronPath = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox()
        Me.btnDelHighCmd = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAddHightCmd = New DevExpress.XtraEditors.SimpleButton()
        Me.XtabQtvAuth = New DevExpress.XtraTab.XtraTabPage()
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
        Me.RepositoryItemSpinEdit3 = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.RepositoryItemCheckEdit4 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.RepositoryItemSpinEdit4 = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.RepositoryItemTextEdit6 = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.rpedtIP = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.rpedtPort = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.rptedtNstvRootKey = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.RepositoryItemCheckEdit6 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.RepositoryItemCheckEdit7 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.RepositoryItemCheckEdit5 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.RepositoryItemCheckEdit1 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.RepositoryItemSpinEdit1 = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.RepositoryItemTextEdit1 = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.RepositoryItemCheckEdit2 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.RepositoryItemSpinEdit2 = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.RepositoryItemTextEdit2 = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.RepositoryItemCheckEdit3 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.RepositoryItemTextEdit3 = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.RepositoryItemCheckEdit8 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.RepositoryItemSpinEdit5 = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.RepositoryItemTextEdit4 = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        CType(Me.richkCboSendOrd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StyleController1, System.ComponentModel.ISupportInitialize).BeginInit()
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
        CType(Me.spinReserveLog.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkLogSQL.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDestroyReason.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtLoginUserName.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtLoginUsreId.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinMaxThreads.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TreLstDB, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkSelect, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spnNum, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.riBtnEdtPath, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.riTxtPassword, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.richkCboInvoiceType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ritxtSendOrd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ricboDonateMark, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.riCboAutoDo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.riCboCompType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.richkRunCmdCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpComDB, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpComDB.SuspendLayout()
        CType(Me.txtCOMPassword.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCOMUserId.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCOMSid.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabLowCmd.SuspendLayout()
        CType(Me.TreLstCmd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rCboCmdName, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabHighCmd.SuspendLayout()
        CType(Me.TreLstSORunCmd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.riSpnNum, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.riTxtCompDescript, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.riCboCmdType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ribtnElectronPath, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.riCboElectronPath, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabQtvAuth.SuspendLayout()
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
        CType(Me.RepositoryItemSpinEdit3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemSpinEdit4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemTextEdit6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rpedtIP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rpedtPort, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rptedtNstvRootKey, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemSpinEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemTextEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemSpinEdit2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemTextEdit2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemTextEdit3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemSpinEdit5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemTextEdit4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'richkCboSendOrd
        '
        Me.richkCboSendOrd.AutoHeight = False
        Me.richkCboSendOrd.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.richkCboSendOrd.Items.AddRange(New DevExpress.XtraEditors.Controls.CheckedListBoxItem() {New DevExpress.XtraEditors.Controls.CheckedListBoxItem("1.電子郵件"), New DevExpress.XtraEditors.Controls.CheckedListBoxItem("2.手機號碼"), New DevExpress.XtraEditors.Controls.CheckedListBoxItem("3.TV Mail"), New DevExpress.XtraEditors.Controls.CheckedListBoxItem("4.CM 導流")})
        Me.richkCboSendOrd.Name = "richkCboSendOrd"
        '
        'DefaultLookAndFeel1
        '
        Me.DefaultLookAndFeel1.LookAndFeel.SkinName = "Office 2010 Blue"
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "1297823140_OK.png")
        Me.ImageList1.Images.SetKeyName(1, "1297823327_exit.png")
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
        Me.grpCtl.Size = New System.Drawing.Size(911, 592)
        Me.grpCtl.TabIndex = 3
        Me.grpCtl.Text = " 設 定 "
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Image = CType(resources.GetObject("btnCancel.Image"), System.Drawing.Image)
        Me.btnCancel.Location = New System.Drawing.Point(545, 552)
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
        Me.btnOK.Location = New System.Drawing.Point(192, 552)
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
        Me.XtabCtl.Size = New System.Drawing.Size(902, 510)
        Me.XtabCtl.TabIndex = 0
        Me.XtabCtl.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XtabDbOpt, Me.XtabLowCmd, Me.XtabHighCmd, Me.XtabSysOpt, Me.XtabQtvAuth, Me.XtabSMailInfo})
        '
        'XtabSysOpt
        '
        Me.XtabSysOpt.Controls.Add(Me.grpQtv)
        Me.XtabSysOpt.Controls.Add(Me.grpPooling)
        Me.XtabSysOpt.Controls.Add(Me.grpGW)
        Me.XtabSysOpt.Image = CType(resources.GetObject("XtabSysOpt.Image"), System.Drawing.Image)
        Me.XtabSysOpt.Name = "XtabSysOpt"
        Me.XtabSysOpt.PageVisible = False
        Me.XtabSysOpt.Size = New System.Drawing.Size(896, 479)
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
        Me.XtabDbOpt.Controls.Add(Me.spinReserveLog)
        Me.XtabDbOpt.Controls.Add(Me.chkLogSQL)
        Me.XtabDbOpt.Controls.Add(Me.txtDestroyReason)
        Me.XtabDbOpt.Controls.Add(Me.LabelControl5)
        Me.XtabDbOpt.Controls.Add(Me.txtLoginUserName)
        Me.XtabDbOpt.Controls.Add(Me.LabelControl4)
        Me.XtabDbOpt.Controls.Add(Me.txtLoginUsreId)
        Me.XtabDbOpt.Controls.Add(Me.LabelControl3)
        Me.XtabDbOpt.Controls.Add(Me.spinMaxThreads)
        Me.XtabDbOpt.Controls.Add(Me.lblMaxThreads)
        Me.XtabDbOpt.Controls.Add(Me.TreLstDB)
        Me.XtabDbOpt.Controls.Add(Me.grpComDB)
        Me.XtabDbOpt.Controls.Add(Me.lblMsg)
        Me.XtabDbOpt.Controls.Add(Me.btnTestConn)
        Me.XtabDbOpt.Controls.Add(Me.btnRemoveConn)
        Me.XtabDbOpt.Controls.Add(Me.btnAddConn)
        Me.XtabDbOpt.Image = CType(resources.GetObject("XtabDbOpt.Image"), System.Drawing.Image)
        Me.XtabDbOpt.Name = "XtabDbOpt"
        Me.XtabDbOpt.Size = New System.Drawing.Size(896, 479)
        Me.XtabDbOpt.Text = " 系 統 台 DB 連 線 設 定 "
        '
        'spinReserveLog
        '
        Me.spinReserveLog.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinReserveLog.EditValue = New Decimal(New Integer() {3, 0, 0, 0})
        Me.spinReserveLog.Location = New System.Drawing.Point(431, 412)
        Me.spinReserveLog.Name = "spinReserveLog"
        Me.spinReserveLog.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinReserveLog.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinReserveLog.Properties.IsFloatValue = False
        Me.spinReserveLog.Properties.Mask.EditMask = "N00"
        Me.spinReserveLog.Properties.MaxValue = New Decimal(New Integer() {12, 0, 0, 0})
        Me.spinReserveLog.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinReserveLog.Size = New System.Drawing.Size(60, 22)
        Me.spinReserveLog.StyleController = Me.StyleController1
        Me.spinReserveLog.TabIndex = 37
        '
        'chkLogSQL
        '
        Me.chkLogSQL.Location = New System.Drawing.Point(378, 413)
        Me.chkLogSQL.Name = "chkLogSQL"
        Me.chkLogSQL.Properties.Caption = "記錄                    月內執行語法"
        Me.chkLogSQL.Size = New System.Drawing.Size(214, 19)
        Me.chkLogSQL.TabIndex = 36
        '
        'txtDestroyReason
        '
        Me.txtDestroyReason.Location = New System.Drawing.Point(534, 440)
        Me.txtDestroyReason.Name = "txtDestroyReason"
        Me.txtDestroyReason.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtDestroyReason.Properties.MaxLength = 20
        Me.txtDestroyReason.Size = New System.Drawing.Size(354, 22)
        Me.txtDestroyReason.StyleController = Me.StyleController1
        Me.txtDestroyReason.TabIndex = 35
        '
        'LabelControl5
        '
        Me.LabelControl5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl5.Location = New System.Drawing.Point(479, 445)
        Me.LabelControl5.Name = "LabelControl5"
        Me.LabelControl5.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl5.TabIndex = 34
        Me.LabelControl5.Text = "註銷原因"
        '
        'txtLoginUserName
        '
        Me.txtLoginUserName.Location = New System.Drawing.Point(289, 441)
        Me.txtLoginUserName.Name = "txtLoginUserName"
        Me.txtLoginUserName.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtLoginUserName.Size = New System.Drawing.Size(184, 22)
        Me.txtLoginUserName.StyleController = Me.StyleController1
        Me.txtLoginUserName.TabIndex = 33
        '
        'LabelControl4
        '
        Me.LabelControl4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl4.Location = New System.Drawing.Point(219, 445)
        Me.LabelControl4.Name = "LabelControl4"
        Me.LabelControl4.Size = New System.Drawing.Size(60, 14)
        Me.LabelControl4.TabIndex = 32
        Me.LabelControl4.Text = "使用者名稱"
        '
        'txtLoginUsreId
        '
        Me.txtLoginUsreId.Location = New System.Drawing.Point(85, 441)
        Me.txtLoginUsreId.Name = "txtLoginUsreId"
        Me.txtLoginUsreId.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.txtLoginUsreId.Size = New System.Drawing.Size(128, 22)
        Me.txtLoginUsreId.StyleController = Me.StyleController1
        Me.txtLoginUsreId.TabIndex = 31
        '
        'LabelControl3
        '
        Me.LabelControl3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelControl3.Location = New System.Drawing.Point(14, 445)
        Me.LabelControl3.Name = "LabelControl3"
        Me.LabelControl3.Size = New System.Drawing.Size(60, 14)
        Me.LabelControl3.TabIndex = 30
        Me.LabelControl3.Text = "使用者代號"
        '
        'spinMaxThreads
        '
        Me.spinMaxThreads.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.spinMaxThreads.EditValue = New Decimal(New Integer() {50, 0, 0, 0})
        Me.spinMaxThreads.Location = New System.Drawing.Point(219, 411)
        Me.spinMaxThreads.Name = "spinMaxThreads"
        Me.spinMaxThreads.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
        Me.spinMaxThreads.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spinMaxThreads.Properties.IsFloatValue = False
        Me.spinMaxThreads.Properties.Mask.EditMask = "N00"
        Me.spinMaxThreads.Properties.MaxValue = New Decimal(New Integer() {500, 0, 0, 0})
        Me.spinMaxThreads.Size = New System.Drawing.Size(60, 22)
        Me.spinMaxThreads.StyleController = Me.StyleController1
        Me.spinMaxThreads.TabIndex = 29
        '
        'lblMaxThreads
        '
        Me.lblMaxThreads.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblMaxThreads.Location = New System.Drawing.Point(14, 415)
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
        Me.lblMaxThreads.TabIndex = 28
        Me.lblMaxThreads.Text = "最大命令處理線程 (最大執行序限制)                     Max Threads"
        '
        'TreLstDB
        '
        Me.TreLstDB.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colSelSO, Me.colCompCode, Me.colCompName, Me.colInvBackupPath, Me.colDBSid, Me.colDBUser, Me.colDBPassword, Me.colComDBSid, Me.colComDBUser, Me.colComDBPassword, Me.colMisDBSid, Me.colMisDBUser, Me.colMisDBPassword, Me.colMisOwner, Me.colDBLink, Me.colMinDBPool, Me.colMaxDBPool, Me.colDBPoolLifetime, Me.colRunCmdTimer, Me.colLimtBeforeUpload, Me.colInvoiceType, Me.colRunCommands, Me.colIsEmailNotify, Me.colIsSmsNotify, Me.colIsTvMailNotify, Me.colIsCMNotify, Me.colSendOrd, Me.colAddSend, Me.colDonateMark, Me.colExportPath, Me.colAutoDo, Me.colStartEmailNotify, Me.colStartSMSNotify, Me.colStartTVMailNotify, Me.colStartCMNotify, Me.colEndCMNotify, Me.colCompType, Me.colMaskInvNo, Me.colStartNotifyPrize})
        Me.TreLstDB.Location = New System.Drawing.Point(14, 49)
        Me.TreLstDB.Name = "TreLstDB"
        Me.TreLstDB.OptionsPrint.UsePrintStyles = True
        Me.TreLstDB.OptionsView.AutoWidth = False
        Me.TreLstDB.OptionsView.ShowButtons = False
        Me.TreLstDB.OptionsView.ShowRoot = False
        Me.TreLstDB.OptionsView.ShowRowFooterSummary = True
        Me.TreLstDB.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.chkSelect, Me.spnNum, Me.riTxtPassword, Me.richkCboInvoiceType, Me.richkRunCmdCode, Me.riBtnEdtPath, Me.ritxtSendOrd, Me.ricboDonateMark, Me.riCboAutoDo, Me.riCboCompType})
        Me.TreLstDB.Size = New System.Drawing.Size(874, 358)
        Me.TreLstDB.TabIndex = 27
        '
        'colSelSO
        '
        Me.colSelSO.AppearanceCell.Options.UseTextOptions = True
        Me.colSelSO.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSelSO.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colSelSO.AppearanceHeader.Options.UseTextOptions = True
        Me.colSelSO.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSelSO.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colSelSO.Caption = "啟 動 作 業"
        Me.colSelSO.ColumnEdit = Me.chkSelect
        Me.colSelSO.FieldName = "SelectID"
        Me.colSelSO.MinWidth = 88
        Me.colSelSO.Name = "colSelSO"
        Me.colSelSO.SummaryFooter = DevExpress.XtraTreeList.SummaryItemType.Count
        Me.colSelSO.SummaryFooterStrFormat = "共 {0:N0} 個系統台"
        Me.colSelSO.Visible = True
        Me.colSelSO.VisibleIndex = 0
        Me.colSelSO.Width = 88
        '
        'chkSelect
        '
        Me.chkSelect.AutoHeight = False
        Me.chkSelect.DisplayValueChecked = "1"
        Me.chkSelect.DisplayValueUnchecked = "0"
        Me.chkSelect.Name = "chkSelect"
        Me.chkSelect.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.chkSelect.ValueChecked = "1"
        Me.chkSelect.ValueUnchecked = "0"
        '
        'colCompCode
        '
        Me.colCompCode.AppearanceCell.Options.UseTextOptions = True
        Me.colCompCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompCode.AppearanceHeader.Options.UseTextOptions = True
        Me.colCompCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompCode.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompCode.Caption = "發票系統台代碼"
        Me.colCompCode.ColumnEdit = Me.spnNum
        Me.colCompCode.FieldName = "CompID"
        Me.colCompCode.Name = "colCompCode"
        Me.colCompCode.Visible = True
        Me.colCompCode.VisibleIndex = 1
        Me.colCompCode.Width = 92
        '
        'spnNum
        '
        Me.spnNum.AllowNullInput = DevExpress.Utils.DefaultBoolean.[True]
        Me.spnNum.AutoHeight = False
        Me.spnNum.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.spnNum.IsFloatValue = False
        Me.spnNum.Mask.EditMask = "N00"
        Me.spnNum.MaxValue = New Decimal(New Integer() {999, 0, 0, 0})
        Me.spnNum.Name = "spnNum"
        '
        'colCompName
        '
        Me.colCompName.AppearanceCell.Options.UseTextOptions = True
        Me.colCompName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompName.Caption = "發票系統台名稱"
        Me.colCompName.FieldName = "CompName"
        Me.colCompName.Name = "colCompName"
        Me.colCompName.UnboundType = DevExpress.XtraTreeList.Data.UnboundColumnType.[String]
        Me.colCompName.Visible = True
        Me.colCompName.VisibleIndex = 2
        Me.colCompName.Width = 214
        '
        'colInvBackupPath
        '
        Me.colInvBackupPath.AppearanceCell.Options.UseTextOptions = True
        Me.colInvBackupPath.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colInvBackupPath.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colInvBackupPath.AppearanceHeader.Options.UseTextOptions = True
        Me.colInvBackupPath.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colInvBackupPath.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colInvBackupPath.Caption = "上傳檔案訊息備份目錄"
        Me.colInvBackupPath.ColumnEdit = Me.riBtnEdtPath
        Me.colInvBackupPath.FieldName = "InvBackupPath"
        Me.colInvBackupPath.Name = "colInvBackupPath"
        Me.colInvBackupPath.Visible = True
        Me.colInvBackupPath.VisibleIndex = 3
        Me.colInvBackupPath.Width = 240
        '
        'riBtnEdtPath
        '
        Me.riBtnEdtPath.AutoHeight = False
        Me.riBtnEdtPath.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.riBtnEdtPath.Name = "riBtnEdtPath"
        '
        'colDBSid
        '
        Me.colDBSid.AppearanceCell.Options.UseTextOptions = True
        Me.colDBSid.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDBSid.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBSid.AppearanceHeader.Options.UseTextOptions = True
        Me.colDBSid.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDBSid.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBSid.Caption = "發票系統SID"
        Me.colDBSid.FieldName = "DBSid"
        Me.colDBSid.Name = "colDBSid"
        Me.colDBSid.Visible = True
        Me.colDBSid.VisibleIndex = 4
        Me.colDBSid.Width = 111
        '
        'colDBUser
        '
        Me.colDBUser.AppearanceCell.Options.UseTextOptions = True
        Me.colDBUser.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDBUser.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBUser.AppearanceHeader.Options.UseTextOptions = True
        Me.colDBUser.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDBUser.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBUser.Caption = "發票系統使用者名稱"
        Me.colDBUser.FieldName = "DBUser"
        Me.colDBUser.Name = "colDBUser"
        Me.colDBUser.Visible = True
        Me.colDBUser.VisibleIndex = 5
        Me.colDBUser.Width = 118
        '
        'colDBPassword
        '
        Me.colDBPassword.AppearanceCell.Options.UseTextOptions = True
        Me.colDBPassword.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDBPassword.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBPassword.AppearanceHeader.Options.UseTextOptions = True
        Me.colDBPassword.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDBPassword.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBPassword.Caption = "發票系統使用者密碼"
        Me.colDBPassword.ColumnEdit = Me.riTxtPassword
        Me.colDBPassword.FieldName = "DBPassword"
        Me.colDBPassword.Name = "colDBPassword"
        Me.colDBPassword.Visible = True
        Me.colDBPassword.VisibleIndex = 6
        Me.colDBPassword.Width = 165
        '
        'riTxtPassword
        '
        Me.riTxtPassword.AutoHeight = False
        Me.riTxtPassword.Name = "riTxtPassword"
        Me.riTxtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        '
        'colComDBSid
        '
        Me.colComDBSid.AppearanceCell.Options.UseTextOptions = True
        Me.colComDBSid.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colComDBSid.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colComDBSid.AppearanceHeader.Options.UseTextOptions = True
        Me.colComDBSid.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colComDBSid.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colComDBSid.Caption = "共用區SID"
        Me.colComDBSid.FieldName = "ComDBSid"
        Me.colComDBSid.Name = "colComDBSid"
        Me.colComDBSid.Visible = True
        Me.colComDBSid.VisibleIndex = 7
        Me.colComDBSid.Width = 108
        '
        'colComDBUser
        '
        Me.colComDBUser.AppearanceCell.Options.UseTextOptions = True
        Me.colComDBUser.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colComDBUser.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colComDBUser.AppearanceHeader.Options.UseTextOptions = True
        Me.colComDBUser.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colComDBUser.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colComDBUser.Caption = "共用區使用者名稱"
        Me.colComDBUser.FieldName = "ComDbUser"
        Me.colComDBUser.Name = "colComDBUser"
        Me.colComDBUser.Visible = True
        Me.colComDBUser.VisibleIndex = 8
        Me.colComDBUser.Width = 118
        '
        'colComDBPassword
        '
        Me.colComDBPassword.AppearanceCell.Options.UseTextOptions = True
        Me.colComDBPassword.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colComDBPassword.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colComDBPassword.AppearanceHeader.Options.UseTextOptions = True
        Me.colComDBPassword.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colComDBPassword.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colComDBPassword.Caption = "共用區使用者密碼"
        Me.colComDBPassword.ColumnEdit = Me.riTxtPassword
        Me.colComDBPassword.FieldName = "ComDBPassword"
        Me.colComDBPassword.Name = "colComDBPassword"
        Me.colComDBPassword.Visible = True
        Me.colComDBPassword.VisibleIndex = 9
        Me.colComDBPassword.Width = 139
        '
        'colMisDBSid
        '
        Me.colMisDBSid.AppearanceCell.Options.UseTextOptions = True
        Me.colMisDBSid.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMisDBSid.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMisDBSid.AppearanceHeader.Options.UseTextOptions = True
        Me.colMisDBSid.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMisDBSid.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMisDBSid.Caption = "系統台SID"
        Me.colMisDBSid.FieldName = "MisDBSid"
        Me.colMisDBSid.Name = "colMisDBSid"
        Me.colMisDBSid.Visible = True
        Me.colMisDBSid.VisibleIndex = 10
        Me.colMisDBSid.Width = 116
        '
        'colMisDBUser
        '
        Me.colMisDBUser.AppearanceCell.Options.UseTextOptions = True
        Me.colMisDBUser.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMisDBUser.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMisDBUser.AppearanceHeader.Options.UseTextOptions = True
        Me.colMisDBUser.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMisDBUser.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMisDBUser.Caption = "系統台使用者名稱"
        Me.colMisDBUser.FieldName = "MisDBUser"
        Me.colMisDBUser.Name = "colMisDBUser"
        Me.colMisDBUser.Visible = True
        Me.colMisDBUser.VisibleIndex = 11
        Me.colMisDBUser.Width = 144
        '
        'colMisDBPassword
        '
        Me.colMisDBPassword.AppearanceCell.Options.UseTextOptions = True
        Me.colMisDBPassword.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMisDBPassword.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMisDBPassword.AppearanceHeader.Options.UseTextOptions = True
        Me.colMisDBPassword.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMisDBPassword.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMisDBPassword.Caption = "系統台使用者密碼"
        Me.colMisDBPassword.ColumnEdit = Me.riTxtPassword
        Me.colMisDBPassword.FieldName = "MisDBPassword"
        Me.colMisDBPassword.Name = "colMisDBPassword"
        Me.colMisDBPassword.Visible = True
        Me.colMisDBPassword.VisibleIndex = 12
        Me.colMisDBPassword.Width = 169
        '
        'colMisOwner
        '
        Me.colMisOwner.AppearanceCell.Options.UseTextOptions = True
        Me.colMisOwner.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMisOwner.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMisOwner.AppearanceHeader.Options.UseTextOptions = True
        Me.colMisOwner.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMisOwner.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMisOwner.Caption = "系統台Owner"
        Me.colMisOwner.FieldName = "MisOwner"
        Me.colMisOwner.Name = "colMisOwner"
        Me.colMisOwner.Visible = True
        Me.colMisOwner.VisibleIndex = 13
        Me.colMisOwner.Width = 132
        '
        'colDBLink
        '
        Me.colDBLink.AppearanceCell.Options.UseTextOptions = True
        Me.colDBLink.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDBLink.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBLink.AppearanceHeader.Options.UseTextOptions = True
        Me.colDBLink.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDBLink.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBLink.Caption = "DBLink"
        Me.colDBLink.FieldName = "DBLink"
        Me.colDBLink.Name = "colDBLink"
        Me.colDBLink.Visible = True
        Me.colDBLink.VisibleIndex = 14
        Me.colDBLink.Width = 126
        '
        'colMinDBPool
        '
        Me.colMinDBPool.AppearanceCell.Options.UseTextOptions = True
        Me.colMinDBPool.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMinDBPool.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMinDBPool.AppearanceHeader.Options.UseTextOptions = True
        Me.colMinDBPool.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMinDBPool.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMinDBPool.Caption = "DB最小集區"
        Me.colMinDBPool.ColumnEdit = Me.spnNum
        Me.colMinDBPool.FieldName = "MinDBPool"
        Me.colMinDBPool.Name = "colMinDBPool"
        Me.colMinDBPool.Visible = True
        Me.colMinDBPool.VisibleIndex = 15
        Me.colMinDBPool.Width = 90
        '
        'colMaxDBPool
        '
        Me.colMaxDBPool.AppearanceCell.Options.UseTextOptions = True
        Me.colMaxDBPool.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMaxDBPool.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMaxDBPool.AppearanceHeader.Options.UseTextOptions = True
        Me.colMaxDBPool.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMaxDBPool.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMaxDBPool.Caption = "DB最大集區"
        Me.colMaxDBPool.ColumnEdit = Me.spnNum
        Me.colMaxDBPool.FieldName = "MaxDBPool"
        Me.colMaxDBPool.Name = "colMaxDBPool"
        Me.colMaxDBPool.Visible = True
        Me.colMaxDBPool.VisibleIndex = 16
        Me.colMaxDBPool.Width = 88
        '
        'colDBPoolLifetime
        '
        Me.colDBPoolLifetime.AppearanceCell.Options.UseTextOptions = True
        Me.colDBPoolLifetime.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDBPoolLifetime.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBPoolLifetime.AppearanceHeader.Options.UseTextOptions = True
        Me.colDBPoolLifetime.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDBPoolLifetime.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBPoolLifetime.Caption = "集區Lifetime"
        Me.colDBPoolLifetime.ColumnEdit = Me.spnNum
        Me.colDBPoolLifetime.FieldName = "DBPoolLifetime"
        Me.colDBPoolLifetime.Name = "colDBPoolLifetime"
        Me.colDBPoolLifetime.Visible = True
        Me.colDBPoolLifetime.VisibleIndex = 17
        Me.colDBPoolLifetime.Width = 92
        '
        'colRunCmdTimer
        '
        Me.colRunCmdTimer.AppearanceCell.Options.UseTextOptions = True
        Me.colRunCmdTimer.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colRunCmdTimer.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colRunCmdTimer.AppearanceHeader.Options.UseTextOptions = True
        Me.colRunCmdTimer.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colRunCmdTimer.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colRunCmdTimer.Caption = "監控頻率(分鐘)"
        Me.colRunCmdTimer.ColumnEdit = Me.spnNum
        Me.colRunCmdTimer.FieldName = "RunCmdTimer"
        Me.colRunCmdTimer.Name = "colRunCmdTimer"
        Me.colRunCmdTimer.Visible = True
        Me.colRunCmdTimer.VisibleIndex = 18
        Me.colRunCmdTimer.Width = 99
        '
        'colLimtBeforeUpload
        '
        Me.colLimtBeforeUpload.AppearanceCell.Options.UseTextOptions = True
        Me.colLimtBeforeUpload.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colLimtBeforeUpload.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colLimtBeforeUpload.AppearanceHeader.Options.UseTextOptions = True
        Me.colLimtBeforeUpload.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colLimtBeforeUpload.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colLimtBeforeUpload.Caption = "上傳限制時間(日)以前"
        Me.colLimtBeforeUpload.ColumnEdit = Me.spnNum
        Me.colLimtBeforeUpload.FieldName = "LimtBeforeUpload"
        Me.colLimtBeforeUpload.Name = "colLimtBeforeUpload"
        Me.colLimtBeforeUpload.Visible = True
        Me.colLimtBeforeUpload.VisibleIndex = 19
        Me.colLimtBeforeUpload.Width = 135
        '
        'colInvoiceType
        '
        Me.colInvoiceType.AppearanceCell.Options.UseTextOptions = True
        Me.colInvoiceType.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colInvoiceType.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colInvoiceType.AppearanceHeader.Options.UseTextOptions = True
        Me.colInvoiceType.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colInvoiceType.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colInvoiceType.Caption = "發票產生方式"
        Me.colInvoiceType.ColumnEdit = Me.richkCboInvoiceType
        Me.colInvoiceType.FieldName = "InvoiceType"
        Me.colInvoiceType.Name = "colInvoiceType"
        Me.colInvoiceType.Visible = True
        Me.colInvoiceType.VisibleIndex = 20
        Me.colInvoiceType.Width = 275
        '
        'richkCboInvoiceType
        '
        Me.richkCboInvoiceType.AutoHeight = False
        Me.richkCboInvoiceType.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.richkCboInvoiceType.Items.AddRange(New DevExpress.XtraEditors.Controls.CheckedListBoxItem() {New DevExpress.XtraEditors.Controls.CheckedListBoxItem("1.拋檔預開"), New DevExpress.XtraEditors.Controls.CheckedListBoxItem("2.拋檔後開"), New DevExpress.XtraEditors.Controls.CheckedListBoxItem("3.現場開立"), New DevExpress.XtraEditors.Controls.CheckedListBoxItem("4.一般開立")})
        Me.richkCboInvoiceType.Name = "richkCboInvoiceType"
        Me.richkCboInvoiceType.SelectAllItemCaption = "全選"
        '
        'colRunCommands
        '
        Me.colRunCommands.AppearanceCell.Options.UseTextOptions = True
        Me.colRunCommands.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colRunCommands.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colRunCommands.AppearanceHeader.Options.UseTextOptions = True
        Me.colRunCommands.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colRunCommands.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colRunCommands.Caption = "執行命令代碼(依循執行)"
        Me.colRunCommands.FieldName = "RunCommands"
        Me.colRunCommands.Name = "colRunCommands"
        Me.colRunCommands.Visible = True
        Me.colRunCommands.VisibleIndex = 21
        Me.colRunCommands.Width = 357
        '
        'colIsEmailNotify
        '
        Me.colIsEmailNotify.AppearanceCell.Options.UseTextOptions = True
        Me.colIsEmailNotify.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colIsEmailNotify.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colIsEmailNotify.AppearanceHeader.Options.UseTextOptions = True
        Me.colIsEmailNotify.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colIsEmailNotify.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colIsEmailNotify.Caption = "1.Email通知"
        Me.colIsEmailNotify.ColumnEdit = Me.chkSelect
        Me.colIsEmailNotify.FieldName = "IsEmailNotify"
        Me.colIsEmailNotify.Name = "colIsEmailNotify"
        Me.colIsEmailNotify.UnboundType = DevExpress.XtraTreeList.Data.UnboundColumnType.[String]
        Me.colIsEmailNotify.Visible = True
        Me.colIsEmailNotify.VisibleIndex = 22
        Me.colIsEmailNotify.Width = 79
        '
        'colIsSmsNotify
        '
        Me.colIsSmsNotify.AppearanceCell.Options.UseTextOptions = True
        Me.colIsSmsNotify.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colIsSmsNotify.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colIsSmsNotify.AppearanceHeader.Options.UseTextOptions = True
        Me.colIsSmsNotify.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colIsSmsNotify.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colIsSmsNotify.Caption = "2.簡訊通知"
        Me.colIsSmsNotify.ColumnEdit = Me.chkSelect
        Me.colIsSmsNotify.FieldName = "IsSmsNotify"
        Me.colIsSmsNotify.Name = "colIsSmsNotify"
        Me.colIsSmsNotify.Visible = True
        Me.colIsSmsNotify.VisibleIndex = 23
        '
        'colIsTvMailNotify
        '
        Me.colIsTvMailNotify.AppearanceCell.Options.UseTextOptions = True
        Me.colIsTvMailNotify.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colIsTvMailNotify.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colIsTvMailNotify.AppearanceHeader.Options.UseTextOptions = True
        Me.colIsTvMailNotify.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colIsTvMailNotify.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colIsTvMailNotify.Caption = "3.TV Mail 通知"
        Me.colIsTvMailNotify.ColumnEdit = Me.chkSelect
        Me.colIsTvMailNotify.FieldName = "IsTvMailNotify"
        Me.colIsTvMailNotify.Name = "colIsTvMailNotify"
        Me.colIsTvMailNotify.Visible = True
        Me.colIsTvMailNotify.VisibleIndex = 24
        Me.colIsTvMailNotify.Width = 93
        '
        'colIsCMNotify
        '
        Me.colIsCMNotify.AppearanceCell.Options.UseTextOptions = True
        Me.colIsCMNotify.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colIsCMNotify.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colIsCMNotify.AppearanceHeader.Options.UseTextOptions = True
        Me.colIsCMNotify.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colIsCMNotify.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colIsCMNotify.Caption = "4.CM 導流通知"
        Me.colIsCMNotify.ColumnEdit = Me.chkSelect
        Me.colIsCMNotify.FieldName = "IsCMNotify"
        Me.colIsCMNotify.Name = "colIsCMNotify"
        Me.colIsCMNotify.Visible = True
        Me.colIsCMNotify.VisibleIndex = 25
        Me.colIsCMNotify.Width = 101
        '
        'colSendOrd
        '
        Me.colSendOrd.AppearanceCell.Options.UseTextOptions = True
        Me.colSendOrd.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSendOrd.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colSendOrd.AppearanceHeader.Options.UseTextOptions = True
        Me.colSendOrd.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSendOrd.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colSendOrd.Caption = "發送優先順序"
        Me.colSendOrd.ColumnEdit = Me.ritxtSendOrd
        Me.colSendOrd.FieldName = "SendOrd"
        Me.colSendOrd.Name = "colSendOrd"
        Me.colSendOrd.Visible = True
        Me.colSendOrd.VisibleIndex = 26
        Me.colSendOrd.Width = 96
        '
        'ritxtSendOrd
        '
        Me.ritxtSendOrd.AllowNullInput = DevExpress.Utils.DefaultBoolean.[True]
        Me.ritxtSendOrd.AutoHeight = False
        Me.ritxtSendOrd.Mask.EditMask = "[1-4]{1,4}"
        Me.ritxtSendOrd.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.RegEx
        Me.ritxtSendOrd.Name = "ritxtSendOrd"
        '
        'colAddSend
        '
        Me.colAddSend.AppearanceCell.Options.UseTextOptions = True
        Me.colAddSend.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colAddSend.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colAddSend.AppearanceHeader.Options.UseTextOptions = True
        Me.colAddSend.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colAddSend.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colAddSend.Caption = "加送"
        Me.colAddSend.ColumnEdit = Me.ritxtSendOrd
        Me.colAddSend.FieldName = "AddSend"
        Me.colAddSend.Name = "colAddSend"
        Me.colAddSend.Visible = True
        Me.colAddSend.VisibleIndex = 27
        '
        'colDonateMark
        '
        Me.colDonateMark.AppearanceCell.Options.UseTextOptions = True
        Me.colDonateMark.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDonateMark.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDonateMark.AppearanceHeader.Options.UseTextOptions = True
        Me.colDonateMark.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDonateMark.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDonateMark.Caption = "捐贈方式"
        Me.colDonateMark.ColumnEdit = Me.ricboDonateMark
        Me.colDonateMark.FieldName = "DonateMark"
        Me.colDonateMark.Name = "colDonateMark"
        Me.colDonateMark.Visible = True
        Me.colDonateMark.VisibleIndex = 28
        Me.colDonateMark.Width = 79
        '
        'ricboDonateMark
        '
        Me.ricboDonateMark.AutoHeight = False
        Me.ricboDonateMark.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.ricboDonateMark.Items.AddRange(New Object() {"0.不捐贈", "1.捐贈", "2.全部"})
        Me.ricboDonateMark.Name = "ricboDonateMark"
        Me.ricboDonateMark.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        '
        'colExportPath
        '
        Me.colExportPath.AppearanceCell.Options.UseTextOptions = True
        Me.colExportPath.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colExportPath.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colExportPath.AppearanceHeader.Options.UseTextOptions = True
        Me.colExportPath.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colExportPath.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colExportPath.Caption = "B08產生清單路徑"
        Me.colExportPath.ColumnEdit = Me.riBtnEdtPath
        Me.colExportPath.FieldName = "ExportPath"
        Me.colExportPath.Name = "colExportPath"
        Me.colExportPath.Visible = True
        Me.colExportPath.VisibleIndex = 29
        Me.colExportPath.Width = 213
        '
        'colAutoDo
        '
        Me.colAutoDo.AppearanceCell.Options.UseTextOptions = True
        Me.colAutoDo.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colAutoDo.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colAutoDo.AppearanceHeader.Options.UseTextOptions = True
        Me.colAutoDo.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colAutoDo.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colAutoDo.Caption = "B08執行方式"
        Me.colAutoDo.ColumnEdit = Me.riCboAutoDo
        Me.colAutoDo.FieldName = "AutoDo"
        Me.colAutoDo.Name = "colAutoDo"
        Me.colAutoDo.Visible = True
        Me.colAutoDo.VisibleIndex = 30
        Me.colAutoDo.Width = 175
        '
        'riCboAutoDo
        '
        Me.riCboAutoDo.AutoHeight = False
        Me.riCboAutoDo.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.riCboAutoDo.Items.AddRange(New Object() {"0.B08單獨執行", "1.B07進入自動執行"})
        Me.riCboAutoDo.Name = "riCboAutoDo"
        Me.riCboAutoDo.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        '
        'colStartEmailNotify
        '
        Me.colStartEmailNotify.AppearanceCell.Options.UseTextOptions = True
        Me.colStartEmailNotify.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colStartEmailNotify.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colStartEmailNotify.AppearanceHeader.Options.UseTextOptions = True
        Me.colStartEmailNotify.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colStartEmailNotify.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colStartEmailNotify.Caption = "啟用 Email 通知排程"
        Me.colStartEmailNotify.ColumnEdit = Me.chkSelect
        Me.colStartEmailNotify.FieldName = "StartEmailNotify"
        Me.colStartEmailNotify.Name = "colStartEmailNotify"
        Me.colStartEmailNotify.Visible = True
        Me.colStartEmailNotify.VisibleIndex = 31
        Me.colStartEmailNotify.Width = 123
        '
        'colStartSMSNotify
        '
        Me.colStartSMSNotify.AppearanceCell.Options.UseTextOptions = True
        Me.colStartSMSNotify.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colStartSMSNotify.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colStartSMSNotify.AppearanceHeader.Options.UseTextOptions = True
        Me.colStartSMSNotify.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colStartSMSNotify.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colStartSMSNotify.Caption = "啟用簡訊通知排程"
        Me.colStartSMSNotify.ColumnEdit = Me.chkSelect
        Me.colStartSMSNotify.FieldName = "StartSMSNotify"
        Me.colStartSMSNotify.Name = "colStartSMSNotify"
        Me.colStartSMSNotify.Visible = True
        Me.colStartSMSNotify.VisibleIndex = 32
        Me.colStartSMSNotify.Width = 110
        '
        'colStartTVMailNotify
        '
        Me.colStartTVMailNotify.AppearanceCell.Options.UseTextOptions = True
        Me.colStartTVMailNotify.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colStartTVMailNotify.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colStartTVMailNotify.AppearanceHeader.Options.UseTextOptions = True
        Me.colStartTVMailNotify.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colStartTVMailNotify.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colStartTVMailNotify.Caption = "啟用 TVMail 排程"
        Me.colStartTVMailNotify.ColumnEdit = Me.chkSelect
        Me.colStartTVMailNotify.FieldName = "StartTVMailNotify"
        Me.colStartTVMailNotify.Name = "colStartTVMailNotify"
        Me.colStartTVMailNotify.Visible = True
        Me.colStartTVMailNotify.VisibleIndex = 33
        Me.colStartTVMailNotify.Width = 108
        '
        'colStartCMNotify
        '
        Me.colStartCMNotify.AppearanceCell.Options.UseTextOptions = True
        Me.colStartCMNotify.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colStartCMNotify.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colStartCMNotify.AppearanceHeader.Options.UseTextOptions = True
        Me.colStartCMNotify.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colStartCMNotify.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colStartCMNotify.Caption = "啟用 CM 導流通知"
        Me.colStartCMNotify.ColumnEdit = Me.chkSelect
        Me.colStartCMNotify.FieldName = "StartCMNotify"
        Me.colStartCMNotify.Name = "colStartCMNotify"
        Me.colStartCMNotify.Visible = True
        Me.colStartCMNotify.VisibleIndex = 34
        Me.colStartCMNotify.Width = 115
        '
        'colEndCMNotify
        '
        Me.colEndCMNotify.AppearanceCell.Options.UseTextOptions = True
        Me.colEndCMNotify.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colEndCMNotify.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colEndCMNotify.AppearanceHeader.Options.UseTextOptions = True
        Me.colEndCMNotify.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colEndCMNotify.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colEndCMNotify.Caption = "CM 導流通知結束排程"
        Me.colEndCMNotify.ColumnEdit = Me.chkSelect
        Me.colEndCMNotify.FieldName = "EndCMNotify"
        Me.colEndCMNotify.Name = "colEndCMNotify"
        Me.colEndCMNotify.Visible = True
        Me.colEndCMNotify.VisibleIndex = 35
        Me.colEndCMNotify.Width = 130
        '
        'colCompType
        '
        Me.colCompType.AppearanceCell.Options.UseTextOptions = True
        Me.colCompType.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompType.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompType.AppearanceHeader.Options.UseTextOptions = True
        Me.colCompType.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompType.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompType.Caption = "B08 執行公司別"
        Me.colCompType.ColumnEdit = Me.riCboCompType
        Me.colCompType.FieldName = "CompType"
        Me.colCompType.Name = "colCompType"
        Me.colCompType.Visible = True
        Me.colCompType.VisibleIndex = 36
        Me.colCompType.Width = 155
        '
        'riCboCompType
        '
        Me.riCboCompType.AutoHeight = False
        Me.riCboCompType.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.riCboCompType.Items.AddRange(New Object() {"0.個別公司", "1.全部公司"})
        Me.riCboCompType.Name = "riCboCompType"
        Me.riCboCompType.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        '
        'colMaskInvNo
        '
        Me.colMaskInvNo.AppearanceCell.Options.UseTextOptions = True
        Me.colMaskInvNo.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMaskInvNo.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMaskInvNo.AppearanceHeader.Options.UseTextOptions = True
        Me.colMaskInvNo.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMaskInvNo.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMaskInvNo.Caption = "遮罩捐贈發票號碼"
        Me.colMaskInvNo.ColumnEdit = Me.chkSelect
        Me.colMaskInvNo.FieldName = "MaskInvNo"
        Me.colMaskInvNo.Name = "colMaskInvNo"
        Me.colMaskInvNo.Visible = True
        Me.colMaskInvNo.VisibleIndex = 37
        Me.colMaskInvNo.Width = 123
        '
        'colStartNotifyPrize
        '
        Me.colStartNotifyPrize.AppearanceCell.Options.UseTextOptions = True
        Me.colStartNotifyPrize.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colStartNotifyPrize.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colStartNotifyPrize.AppearanceHeader.Options.UseTextOptions = True
        Me.colStartNotifyPrize.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colStartNotifyPrize.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colStartNotifyPrize.Caption = "中獎通知"
        Me.colStartNotifyPrize.ColumnEdit = Me.chkSelect
        Me.colStartNotifyPrize.FieldName = "StartNotifyPrize"
        Me.colStartNotifyPrize.Name = "colStartNotifyPrize"
        Me.colStartNotifyPrize.Visible = True
        Me.colStartNotifyPrize.VisibleIndex = 38
        Me.colStartNotifyPrize.Width = 91
        '
        'richkRunCmdCode
        '
        Me.richkRunCmdCode.AutoHeight = False
        Me.richkRunCmdCode.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.richkRunCmdCode.Name = "richkRunCmdCode"
        Me.richkRunCmdCode.SelectAllItemVisible = False
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
        'XtabLowCmd
        '
        Me.XtabLowCmd.Controls.Add(Me.TreLstCmd)
        Me.XtabLowCmd.Controls.Add(Me.btnDelLowCmd)
        Me.XtabLowCmd.Controls.Add(Me.btnAddLowCmd)
        Me.XtabLowCmd.Image = CType(resources.GetObject("XtabLowCmd.Image"), System.Drawing.Image)
        Me.XtabLowCmd.Name = "XtabLowCmd"
        Me.XtabLowCmd.Size = New System.Drawing.Size(896, 479)
        Me.XtabLowCmd.Text = "低 階 命 令 設 定 "
        '
        'TreLstCmd
        '
        Me.TreLstCmd.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colCodeNo, Me.colCmdName})
        Me.TreLstCmd.Location = New System.Drawing.Point(24, 52)
        Me.TreLstCmd.Name = "TreLstCmd"
        Me.TreLstCmd.OptionsPrint.UsePrintStyles = True
        Me.TreLstCmd.OptionsView.ShowButtons = False
        Me.TreLstCmd.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.rCboCmdName})
        Me.TreLstCmd.Size = New System.Drawing.Size(849, 407)
        Me.TreLstCmd.TabIndex = 24
        '
        'colCodeNo
        '
        Me.colCodeNo.AppearanceCell.Options.UseTextOptions = True
        Me.colCodeNo.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCodeNo.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCodeNo.AppearanceHeader.Options.UseTextOptions = True
        Me.colCodeNo.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCodeNo.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCodeNo.Caption = "代碼"
        Me.colCodeNo.FieldName = "CodeNo"
        Me.colCodeNo.Name = "colCodeNo"
        Me.colCodeNo.Visible = True
        Me.colCodeNo.VisibleIndex = 0
        Me.colCodeNo.Width = 193
        '
        'colCmdName
        '
        Me.colCmdName.AppearanceCell.Options.UseTextOptions = True
        Me.colCmdName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colCmdName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCmdName.AppearanceHeader.Options.UseTextOptions = True
        Me.colCmdName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colCmdName.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCmdName.Caption = "命令名稱"
        Me.colCmdName.ColumnEdit = Me.rCboCmdName
        Me.colCmdName.FieldName = "CmdName"
        Me.colCmdName.Name = "colCmdName"
        Me.colCmdName.Visible = True
        Me.colCmdName.VisibleIndex = 1
        Me.colCmdName.Width = 638
        '
        'rCboCmdName
        '
        Me.rCboCmdName.AutoHeight = False
        Me.rCboCmdName.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.rCboCmdName.Name = "rCboCmdName"
        Me.rCboCmdName.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
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
        Me.XtabHighCmd.Controls.Add(Me.TreLstSORunCmd)
        Me.XtabHighCmd.Controls.Add(Me.btnDelHighCmd)
        Me.XtabHighCmd.Controls.Add(Me.btnAddHightCmd)
        Me.XtabHighCmd.Image = CType(resources.GetObject("XtabHighCmd.Image"), System.Drawing.Image)
        Me.XtabHighCmd.Name = "XtabHighCmd"
        Me.XtabHighCmd.Size = New System.Drawing.Size(896, 479)
        Me.XtabHighCmd.Text = "高 階 命 令 設 定"
        '
        'TreLstSORunCmd
        '
        Me.TreLstSORunCmd.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colCompID, Me.colCompDescript, Me.colCmdType, Me.colCmdTypeName, Me.colRunfrequency, Me.colExportElectronPath})
        Me.TreLstSORunCmd.Location = New System.Drawing.Point(21, 51)
        Me.TreLstSORunCmd.Name = "TreLstSORunCmd"
        Me.TreLstSORunCmd.OptionsLayout.AddNewColumns = False
        Me.TreLstSORunCmd.OptionsPrint.UsePrintStyles = True
        Me.TreLstSORunCmd.OptionsView.AutoWidth = False
        Me.TreLstSORunCmd.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.riSpnNum, Me.riTxtCompDescript, Me.riCboCmdType, Me.riCboElectronPath, Me.ribtnElectronPath})
        Me.TreLstSORunCmd.Size = New System.Drawing.Size(867, 409)
        Me.TreLstSORunCmd.TabIndex = 26
        '
        'colCompID
        '
        Me.colCompID.AppearanceCell.Options.UseTextOptions = True
        Me.colCompID.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompID.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompID.AppearanceHeader.Options.UseTextOptions = True
        Me.colCompID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompID.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompID.Caption = "系統台代碼"
        Me.colCompID.ColumnEdit = Me.riSpnNum
        Me.colCompID.FieldName = "CompID"
        Me.colCompID.Name = "colCompID"
        Me.colCompID.Visible = True
        Me.colCompID.VisibleIndex = 0
        Me.colCompID.Width = 83
        '
        'riSpnNum
        '
        Me.riSpnNum.AutoHeight = False
        Me.riSpnNum.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.riSpnNum.IsFloatValue = False
        Me.riSpnNum.Mask.EditMask = "N00"
        Me.riSpnNum.MaxValue = New Decimal(New Integer() {9999, 0, 0, 0})
        Me.riSpnNum.Name = "riSpnNum"
        '
        'colCompDescript
        '
        Me.colCompDescript.AppearanceCell.Options.UseTextOptions = True
        Me.colCompDescript.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colCompDescript.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompDescript.AppearanceHeader.Options.UseTextOptions = True
        Me.colCompDescript.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colCompDescript.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompDescript.Caption = "公司名稱"
        Me.colCompDescript.ColumnEdit = Me.riTxtCompDescript
        Me.colCompDescript.FieldName = "CompDescript"
        Me.colCompDescript.Name = "colCompDescript"
        Me.colCompDescript.OptionsColumn.AllowEdit = False
        Me.colCompDescript.OptionsColumn.AllowFocus = False
        Me.colCompDescript.OptionsColumn.ReadOnly = True
        Me.colCompDescript.Visible = True
        Me.colCompDescript.VisibleIndex = 1
        Me.colCompDescript.Width = 267
        '
        'riTxtCompDescript
        '
        Me.riTxtCompDescript.AllowFocused = False
        Me.riTxtCompDescript.AutoHeight = False
        Me.riTxtCompDescript.Name = "riTxtCompDescript"
        Me.riTxtCompDescript.ReadOnly = True
        '
        'colCmdType
        '
        Me.colCmdType.AppearanceCell.Options.UseTextOptions = True
        Me.colCmdType.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colCmdType.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCmdType.AppearanceHeader.Options.UseTextOptions = True
        Me.colCmdType.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colCmdType.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCmdType.Caption = "命令種類"
        Me.colCmdType.ColumnEdit = Me.riCboCmdType
        Me.colCmdType.FieldName = "CmdType"
        Me.colCmdType.Name = "colCmdType"
        Me.colCmdType.Visible = True
        Me.colCmdType.VisibleIndex = 2
        Me.colCmdType.Width = 92
        '
        'riCboCmdType
        '
        Me.riCboCmdType.AutoHeight = False
        Me.riCboCmdType.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.riCboCmdType.Name = "riCboCmdType"
        Me.riCboCmdType.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        '
        'colCmdTypeName
        '
        Me.colCmdTypeName.AppearanceCell.Options.UseTextOptions = True
        Me.colCmdTypeName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colCmdTypeName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCmdTypeName.AppearanceHeader.Options.UseTextOptions = True
        Me.colCmdTypeName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colCmdTypeName.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCmdTypeName.Caption = "命令名稱"
        Me.colCmdTypeName.FieldName = "CmdTypeName"
        Me.colCmdTypeName.Name = "colCmdTypeName"
        Me.colCmdTypeName.OptionsColumn.AllowEdit = False
        Me.colCmdTypeName.OptionsColumn.AllowFocus = False
        Me.colCmdTypeName.OptionsColumn.ReadOnly = True
        Me.colCmdTypeName.Visible = True
        Me.colCmdTypeName.VisibleIndex = 3
        Me.colCmdTypeName.Width = 147
        '
        'colRunfrequency
        '
        Me.colRunfrequency.AppearanceCell.Options.UseTextOptions = True
        Me.colRunfrequency.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colRunfrequency.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colRunfrequency.AppearanceHeader.Options.UseTextOptions = True
        Me.colRunfrequency.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colRunfrequency.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colRunfrequency.Caption = "上傳頻率(分鐘)"
        Me.colRunfrequency.ColumnEdit = Me.riSpnNum
        Me.colRunfrequency.FieldName = "Runfrequency"
        Me.colRunfrequency.Name = "colRunfrequency"
        Me.colRunfrequency.Visible = True
        Me.colRunfrequency.VisibleIndex = 4
        Me.colRunfrequency.Width = 122
        '
        'colExportElectronPath
        '
        Me.colExportElectronPath.AppearanceCell.Options.UseTextOptions = True
        Me.colExportElectronPath.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colExportElectronPath.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colExportElectronPath.AppearanceHeader.Options.UseTextOptions = True
        Me.colExportElectronPath.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colExportElectronPath.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colExportElectronPath.Caption = "產生路徑"
        Me.colExportElectronPath.ColumnEdit = Me.ribtnElectronPath
        Me.colExportElectronPath.FieldName = "ExportElectronPath"
        Me.colExportElectronPath.Name = "colExportElectronPath"
        Me.colExportElectronPath.Visible = True
        Me.colExportElectronPath.VisibleIndex = 5
        Me.colExportElectronPath.Width = 253
        '
        'ribtnElectronPath
        '
        Me.ribtnElectronPath.AutoHeight = False
        Me.ribtnElectronPath.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.ribtnElectronPath.Name = "ribtnElectronPath"
        '
        'riCboElectronPath
        '
        Me.riCboElectronPath.AutoHeight = False
        Me.riCboElectronPath.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.riCboElectronPath.Name = "riCboElectronPath"
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
        Me.XtabQtvAuth.Controls.Add(Me.btnRemoveAuth)
        Me.XtabQtvAuth.Controls.Add(Me.btnAddError)
        Me.XtabQtvAuth.Image = CType(resources.GetObject("XtabQtvAuth.Image"), System.Drawing.Image)
        Me.XtabQtvAuth.Name = "XtabQtvAuth"
        Me.XtabQtvAuth.PageVisible = False
        Me.XtabQtvAuth.Size = New System.Drawing.Size(896, 479)
        Me.XtabQtvAuth.Text = " 錯 誤 代 碼 表 設 定 "
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
        Me.XtabSMailInfo.Size = New System.Drawing.Size(896, 479)
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
        Me.RepositoryItemSpinEdit3.MaxValue = New Decimal(New Integer() {99, 0, 0, 0})
        Me.RepositoryItemSpinEdit3.Name = "RepositoryItemSpinEdit3"
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
        'RepositoryItemTextEdit6
        '
        Me.RepositoryItemTextEdit6.AutoHeight = False
        Me.RepositoryItemTextEdit6.Name = "RepositoryItemTextEdit6"
        Me.RepositoryItemTextEdit6.NullText = "未設定密碼"
        Me.RepositoryItemTextEdit6.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        '
        'rpedtIP
        '
        Me.rpedtIP.AutoHeight = False
        Me.rpedtIP.Mask.PlaceHolder = Global.Microsoft.VisualBasic.ChrW(32)
        Me.rpedtIP.Name = "rpedtIP"
        '
        'rpedtPort
        '
        Me.rpedtPort.AutoHeight = False
        Me.rpedtPort.Name = "rpedtPort"
        '
        'rptedtNstvRootKey
        '
        Me.rptedtNstvRootKey.AutoHeight = False
        Me.rptedtNstvRootKey.Name = "rptedtNstvRootKey"
        '
        'RepositoryItemCheckEdit6
        '
        Me.RepositoryItemCheckEdit6.AutoHeight = False
        Me.RepositoryItemCheckEdit6.Name = "RepositoryItemCheckEdit6"
        Me.RepositoryItemCheckEdit6.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.RepositoryItemCheckEdit6.ValueChecked = "1"
        Me.RepositoryItemCheckEdit6.ValueUnchecked = "0"
        '
        'RepositoryItemCheckEdit7
        '
        Me.RepositoryItemCheckEdit7.AutoHeight = False
        Me.RepositoryItemCheckEdit7.Name = "RepositoryItemCheckEdit7"
        Me.RepositoryItemCheckEdit7.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.RepositoryItemCheckEdit7.ValueChecked = "1"
        Me.RepositoryItemCheckEdit7.ValueUnchecked = "0"
        '
        'RepositoryItemCheckEdit5
        '
        Me.RepositoryItemCheckEdit5.AutoHeight = False
        Me.RepositoryItemCheckEdit5.Name = "RepositoryItemCheckEdit5"
        Me.RepositoryItemCheckEdit5.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.RepositoryItemCheckEdit5.ValueChecked = "1"
        Me.RepositoryItemCheckEdit5.ValueUnchecked = "0"
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
        Me.RepositoryItemSpinEdit1.MaxValue = New Decimal(New Integer() {99, 0, 0, 0})
        Me.RepositoryItemSpinEdit1.Name = "RepositoryItemSpinEdit1"
        '
        'RepositoryItemTextEdit1
        '
        Me.RepositoryItemTextEdit1.AutoHeight = False
        Me.RepositoryItemTextEdit1.Name = "RepositoryItemTextEdit1"
        Me.RepositoryItemTextEdit1.NullText = "未設定密碼"
        Me.RepositoryItemTextEdit1.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
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
        Me.RepositoryItemSpinEdit2.MaxValue = New Decimal(New Integer() {99, 0, 0, 0})
        Me.RepositoryItemSpinEdit2.Name = "RepositoryItemSpinEdit2"
        '
        'RepositoryItemTextEdit2
        '
        Me.RepositoryItemTextEdit2.AutoHeight = False
        Me.RepositoryItemTextEdit2.Name = "RepositoryItemTextEdit2"
        Me.RepositoryItemTextEdit2.NullText = "未設定密碼"
        Me.RepositoryItemTextEdit2.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
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
        'RepositoryItemTextEdit3
        '
        Me.RepositoryItemTextEdit3.AutoHeight = False
        Me.RepositoryItemTextEdit3.Name = "RepositoryItemTextEdit3"
        Me.RepositoryItemTextEdit3.NullText = "未設定密碼"
        Me.RepositoryItemTextEdit3.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        '
        'RepositoryItemCheckEdit8
        '
        Me.RepositoryItemCheckEdit8.AutoHeight = False
        Me.RepositoryItemCheckEdit8.DisplayValueChecked = "1"
        Me.RepositoryItemCheckEdit8.DisplayValueUnchecked = "0"
        Me.RepositoryItemCheckEdit8.Name = "RepositoryItemCheckEdit8"
        Me.RepositoryItemCheckEdit8.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.RepositoryItemCheckEdit8.ValueChecked = "1"
        Me.RepositoryItemCheckEdit8.ValueUnchecked = "0"
        '
        'RepositoryItemSpinEdit5
        '
        Me.RepositoryItemSpinEdit5.AllowNullInput = DevExpress.Utils.DefaultBoolean.[False]
        Me.RepositoryItemSpinEdit5.Appearance.Options.UseTextOptions = True
        Me.RepositoryItemSpinEdit5.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.RepositoryItemSpinEdit5.AutoHeight = False
        Me.RepositoryItemSpinEdit5.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.RepositoryItemSpinEdit5.IsFloatValue = False
        Me.RepositoryItemSpinEdit5.Mask.EditMask = "N00"
        Me.RepositoryItemSpinEdit5.MaxLength = 2
        Me.RepositoryItemSpinEdit5.MaxValue = New Decimal(New Integer() {99, 0, 0, 0})
        Me.RepositoryItemSpinEdit5.Name = "RepositoryItemSpinEdit5"
        '
        'RepositoryItemTextEdit4
        '
        Me.RepositoryItemTextEdit4.AutoHeight = False
        Me.RepositoryItemTextEdit4.Name = "RepositoryItemTextEdit4"
        Me.RepositoryItemTextEdit4.NullText = "未設定密碼"
        Me.RepositoryItemTextEdit4.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        '
        'frmSetup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(911, 592)
        Me.Controls.Add(Me.grpCtl)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmSetup"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "發票系統 Gateway 設定檔"
        CType(Me.richkCboSendOrd, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StyleController1, System.ComponentModel.ISupportInitialize).EndInit()
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
        Me.XtabDbOpt.PerformLayout()
        CType(Me.spinReserveLog.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkLogSQL.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDestroyReason.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtLoginUserName.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtLoginUsreId.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinMaxThreads.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TreLstDB, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkSelect, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spnNum, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.riBtnEdtPath, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.riTxtPassword, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.richkCboInvoiceType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ritxtSendOrd, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ricboDonateMark, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.riCboAutoDo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.riCboCompType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.richkRunCmdCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpComDB, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpComDB.ResumeLayout(False)
        Me.grpComDB.PerformLayout()
        CType(Me.txtCOMPassword.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCOMUserId.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCOMSid.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabLowCmd.ResumeLayout(False)
        CType(Me.TreLstCmd, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rCboCmdName, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabHighCmd.ResumeLayout(False)
        CType(Me.TreLstSORunCmd, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.riSpnNum, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.riTxtCompDescript, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.riCboCmdType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ribtnElectronPath, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.riCboElectronPath, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabQtvAuth.ResumeLayout(False)
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
        CType(Me.RepositoryItemSpinEdit3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemSpinEdit4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemTextEdit6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rpedtIP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rpedtPort, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rptedtNstvRootKey, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemSpinEdit1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemTextEdit1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemSpinEdit2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemTextEdit2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemTextEdit3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemSpinEdit5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemTextEdit4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DefaultLookAndFeel1 As DevExpress.LookAndFeel.DefaultLookAndFeel
    Friend WithEvents StyleController1 As DevExpress.XtraEditors.StyleController
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
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
    Friend WithEvents XtabLowCmd As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents btnDelLowCmd As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAddLowCmd As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents XtabHighCmd As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents btnDelHighCmd As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAddHightCmd As DevExpress.XtraEditors.SimpleButton
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
    Friend WithEvents RepositoryItemCheckEdit4 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents RepositoryItemSpinEdit4 As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents RepositoryItemTextEdit6 As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents rpedtIP As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents rpedtPort As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents rptedtNstvRootKey As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents RepositoryItemCheckEdit6 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents RepositoryItemCheckEdit7 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents RepositoryItemCheckEdit5 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents TreLstDB As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colSelSO As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents chkSelect As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents colCompCode As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents spnNum As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents colCompName As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colDBSid As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colDBUser As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colDBPassword As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents riTxtPassword As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents colComDBSid As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colComDBUser As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colComDBPassword As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents RepositoryItemCheckEdit1 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents RepositoryItemSpinEdit1 As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents RepositoryItemTextEdit1 As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents RepositoryItemCheckEdit2 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents RepositoryItemSpinEdit2 As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents RepositoryItemTextEdit2 As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents RepositoryItemCheckEdit3 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents RepositoryItemSpinEdit3 As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents RepositoryItemTextEdit3 As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents RepositoryItemCheckEdit8 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents RepositoryItemSpinEdit5 As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents RepositoryItemTextEdit4 As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents colRunCmdTimer As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents TreLstCmd As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colCodeNo As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colCmdName As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents rCboCmdName As DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Friend WithEvents TreLstSORunCmd As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colCompID As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents riSpnNum As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents colCompDescript As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents riTxtCompDescript As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents colCmdType As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents riCboCmdType As DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Friend WithEvents spinMaxThreads As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents lblMaxThreads As DevExpress.XtraEditors.LabelControl
    Friend WithEvents colRunfrequency As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colInvoiceType As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colCmdTypeName As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents richkCboInvoiceType As DevExpress.XtraEditors.Repository.RepositoryItemCheckedComboBoxEdit
    Friend WithEvents colMinDBPool As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colMaxDBPool As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colDBPoolLifetime As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colRunCommands As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents richkRunCmdCode As DevExpress.XtraEditors.Repository.RepositoryItemCheckedComboBoxEdit
    Friend WithEvents colLimtBeforeUpload As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colMisDBSid As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colMisDBUser As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colMisDBPassword As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colMisOwner As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colDBLink As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents txtLoginUserName As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl4 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtLoginUsreId As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl3 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtDestroyReason As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl5 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents colInvBackupPath As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents riBtnEdtPath As DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents colIsEmailNotify As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colIsSmsNotify As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colIsTvMailNotify As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colIsCMNotify As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colSendOrd As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents richkCboSendOrd As DevExpress.XtraEditors.Repository.RepositoryItemCheckedComboBoxEdit
    Friend WithEvents ritxtSendOrd As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents colAddSend As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colDonateMark As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents ricboDonateMark As DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Friend WithEvents colExportPath As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colExportElectronPath As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents riCboElectronPath As DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Friend WithEvents ribtnElectronPath As DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit
    Friend WithEvents colAutoDo As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents riCboAutoDo As DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Friend WithEvents colStartEmailNotify As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colStartSMSNotify As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colStartTVMailNotify As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colStartCMNotify As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colEndCMNotify As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colCompType As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents riCboCompType As DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Friend WithEvents colMaskInvNo As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colStartNotifyPrize As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents spinReserveLog As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents chkLogSQL As DevExpress.XtraEditors.CheckEdit
End Class
