<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOverseeSetting
    ' Inherits System.Windows.Forms.Form
    Inherits DevExpress.XtraEditors.XtraForm
    'Form 覆寫 Dispose 以清除元件清單。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請不要使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOverseeSetting))
        Me.grpCtl = New DevExpress.XtraEditors.GroupControl()
        Me.btnExit = New DevExpress.XtraEditors.SimpleButton()
        Me.spinLogHistorical = New DevExpress.XtraEditors.SpinEdit()
        Me.StyleController1 = New DevExpress.XtraEditors.StyleController()
        Me.LabelControl2 = New DevExpress.XtraEditors.LabelControl()
        Me.spinMaxThreads = New DevExpress.XtraEditors.SpinEdit()
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.btnSave = New DevExpress.XtraEditors.SimpleButton()
        Me.XtabCtl = New DevExpress.XtraTab.XtraTabControl()
        Me.XtabCommonOpt = New DevExpress.XtraTab.XtraTabPage()
        Me.PanelControl2 = New DevExpress.XtraEditors.PanelControl()
        Me.TreLstDB = New DevExpress.XtraTreeList.TreeList()
        Me.colSelSO = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.chkSelect = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.colSid = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.txtSid = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.colDBUserName = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.txtDBUser = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.colDBPassword = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.txtDBPassword = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.colMonitorType = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.cmbMonitorType = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox()
        Me.colDBMinPool = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.spnNum = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.colDBMaxPool = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colDBPoolLifetime = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colPrgName = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.txtPrgName = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.colRunTimer = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colSeqSQL = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.mdxCommon = New DevExpress.XtraEditors.Repository.RepositoryItemMemoExEdit()
        Me.ColInsSql = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colTimeout = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colGetSql = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colCompCode = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colCompName = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.txtCommon = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.colWebServiceName = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colWebSrvMethod = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colWebServicePara = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colUrl = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colWebGetPara = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colWebPostPara = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.txtGatewayName = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.meErrorTxt = New DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit()
        Me.cmbMonitorType1 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckedComboBoxEdit()
        Me.mdCommon = New DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit()
        Me.rgMonitorType = New DevExpress.XtraEditors.Repository.RepositoryItemRadioGroup()
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        Me.btnTest = New DevExpress.XtraEditors.SimpleButton()
        Me.btnDel = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAdd = New DevExpress.XtraEditors.SimpleButton()
        Me.DefaultLookAndFeel1 = New DevExpress.LookAndFeel.DefaultLookAndFeel()
        CType(Me.grpCtl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpCtl.SuspendLayout()
        CType(Me.spinLogHistorical.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StyleController1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spinMaxThreads.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.XtabCtl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtabCtl.SuspendLayout()
        Me.XtabCommonOpt.SuspendLayout()
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl2.SuspendLayout()
        CType(Me.TreLstDB, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkSelect, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSid, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDBUser, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDBPassword, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmbMonitorType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spnNum, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtPrgName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mdxCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtGatewayName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.meErrorTxt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmbMonitorType1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mdCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rgMonitorType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpCtl
        '
        Me.grpCtl.Controls.Add(Me.btnExit)
        Me.grpCtl.Controls.Add(Me.spinLogHistorical)
        Me.grpCtl.Controls.Add(Me.LabelControl2)
        Me.grpCtl.Controls.Add(Me.spinMaxThreads)
        Me.grpCtl.Controls.Add(Me.LabelControl1)
        Me.grpCtl.Controls.Add(Me.btnSave)
        Me.grpCtl.Controls.Add(Me.XtabCtl)
        Me.grpCtl.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpCtl.Location = New System.Drawing.Point(0, 0)
        Me.grpCtl.Name = "grpCtl"
        Me.grpCtl.Size = New System.Drawing.Size(882, 565)
        Me.grpCtl.TabIndex = 0
        Me.grpCtl.Text = " 設 定 "
        '
        'btnExit
        '
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.Location = New System.Drawing.Point(795, 522)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 31)
        Me.btnExit.TabIndex = 10
        Me.btnExit.Text = "離開"
        '
        'spinLogHistorical
        '
        Me.spinLogHistorical.EditValue = New Decimal(New Integer() {3, 0, 0, 0})
        Me.spinLogHistorical.Location = New System.Drawing.Point(68, 477)
        Me.spinLogHistorical.Name = "spinLogHistorical"
        Me.spinLogHistorical.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.spinLogHistorical.Properties.IsFloatValue = False
        Me.spinLogHistorical.Properties.Mask.EditMask = "N00"
        Me.spinLogHistorical.Properties.MaxValue = New Decimal(New Integer() {99, 0, 0, 0})
        Me.spinLogHistorical.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinLogHistorical.Size = New System.Drawing.Size(68, 20)
        Me.spinLogHistorical.StyleController = Me.StyleController1
        Me.spinLogHistorical.TabIndex = 9
        '
        'StyleController1
        '
        Me.StyleController1.AppearanceFocused.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.StyleController1.AppearanceFocused.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.StyleController1.AppearanceFocused.BorderColor = System.Drawing.Color.Blue
        Me.StyleController1.AppearanceFocused.Options.UseBackColor = True
        Me.StyleController1.AppearanceFocused.Options.UseBorderColor = True
        '
        'LabelControl2
        '
        Me.LabelControl2.Location = New System.Drawing.Point(12, 480)
        Me.LabelControl2.Name = "LabelControl2"
        Me.LabelControl2.Size = New System.Drawing.Size(156, 14)
        Me.LabelControl2.TabIndex = 8
        Me.LabelControl2.Text = "訊息記錄                     個月"
        '
        'spinMaxThreads
        '
        Me.spinMaxThreads.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinMaxThreads.Location = New System.Drawing.Point(215, 445)
        Me.spinMaxThreads.Name = "spinMaxThreads"
        Me.spinMaxThreads.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.spinMaxThreads.Properties.IsFloatValue = False
        Me.spinMaxThreads.Properties.Mask.EditMask = "N00"
        Me.spinMaxThreads.Properties.MaxValue = New Decimal(New Integer() {999, 0, 0, 0})
        Me.spinMaxThreads.Properties.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spinMaxThreads.Size = New System.Drawing.Size(68, 20)
        Me.spinMaxThreads.StyleController = Me.StyleController1
        Me.spinMaxThreads.TabIndex = 7
        '
        'LabelControl1
        '
        Me.LabelControl1.Location = New System.Drawing.Point(12, 448)
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(347, 14)
        Me.LabelControl1.TabIndex = 6
        Me.LabelControl1.Text = "最大命令處理線程 (最大執行序限制)                     Max Threads"
        '
        'btnSave
        '
        Me.btnSave.Image = CType(resources.GetObject("btnSave.Image"), System.Drawing.Image)
        Me.btnSave.Location = New System.Drawing.Point(5, 522)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(80, 31)
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "儲存"
        '
        'XtabCtl
        '
        Me.XtabCtl.Dock = System.Windows.Forms.DockStyle.Top
        Me.XtabCtl.Location = New System.Drawing.Point(2, 22)
        Me.XtabCtl.Name = "XtabCtl"
        Me.XtabCtl.SelectedTabPage = Me.XtabCommonOpt
        Me.XtabCtl.Size = New System.Drawing.Size(878, 420)
        Me.XtabCtl.TabIndex = 0
        Me.XtabCtl.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XtabCommonOpt})
        '
        'XtabCommonOpt
        '
        Me.XtabCommonOpt.Controls.Add(Me.PanelControl2)
        Me.XtabCommonOpt.Controls.Add(Me.PanelControl1)
        Me.XtabCommonOpt.Name = "XtabCommonOpt"
        Me.XtabCommonOpt.Size = New System.Drawing.Size(872, 391)
        Me.XtabCommonOpt.Text = "一般設定"
        '
        'PanelControl2
        '
        Me.PanelControl2.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl2.Controls.Add(Me.TreLstDB)
        Me.PanelControl2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelControl2.Location = New System.Drawing.Point(0, 43)
        Me.PanelControl2.Name = "PanelControl2"
        Me.PanelControl2.Size = New System.Drawing.Size(872, 348)
        Me.PanelControl2.TabIndex = 8
        '
        'TreLstDB
        '
        Me.TreLstDB.Appearance.FocusedCell.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.TreLstDB.Appearance.FocusedCell.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.TreLstDB.Appearance.FocusedCell.BorderColor = System.Drawing.Color.Blue
        Me.TreLstDB.Appearance.FocusedCell.Options.UseBackColor = True
        Me.TreLstDB.Appearance.FocusedCell.Options.UseBorderColor = True
        Me.TreLstDB.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colSelSO, Me.colSid, Me.colDBUserName, Me.colDBPassword, Me.colMonitorType, Me.colDBMinPool, Me.colDBMaxPool, Me.colDBPoolLifetime, Me.colPrgName, Me.colRunTimer, Me.colSeqSQL, Me.ColInsSql, Me.colTimeout, Me.colGetSql, Me.colCompCode, Me.colCompName, Me.colWebServiceName, Me.colWebSrvMethod, Me.colWebServicePara, Me.colUrl, Me.colWebGetPara, Me.colWebPostPara})
        Me.TreLstDB.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TreLstDB.Location = New System.Drawing.Point(0, 0)
        Me.TreLstDB.Name = "TreLstDB"
        Me.TreLstDB.OptionsView.AutoWidth = False
        Me.TreLstDB.OptionsView.ShowRoot = False
        Me.TreLstDB.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.chkSelect, Me.txtGatewayName, Me.meErrorTxt, Me.txtSid, Me.txtDBUser, Me.txtDBPassword, Me.cmbMonitorType1, Me.spnNum, Me.txtPrgName, Me.txtCommon, Me.mdCommon, Me.mdxCommon, Me.rgMonitorType, Me.cmbMonitorType})
        Me.TreLstDB.Size = New System.Drawing.Size(872, 348)
        Me.TreLstDB.TabIndex = 14
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
        Me.colSelSO.Name = "colSelSO"
        Me.colSelSO.Visible = True
        Me.colSelSO.VisibleIndex = 0
        Me.colSelSO.Width = 78
        '
        'chkSelect
        '
        Me.chkSelect.AutoHeight = False
        Me.chkSelect.Caption = "Check"
        Me.chkSelect.DisplayValueChecked = "1"
        Me.chkSelect.DisplayValueUnchecked = "0"
        Me.chkSelect.Name = "chkSelect"
        Me.chkSelect.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked
        Me.chkSelect.ValueChecked = 1
        Me.chkSelect.ValueUnchecked = 0
        '
        'colSid
        '
        Me.colSid.AppearanceCell.Options.UseTextOptions = True
        Me.colSid.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSid.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colSid.AppearanceHeader.Options.UseTextOptions = True
        Me.colSid.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSid.Caption = "SID"
        Me.colSid.ColumnEdit = Me.txtSid
        Me.colSid.FieldName = "Sid"
        Me.colSid.Name = "colSid"
        Me.colSid.Visible = True
        Me.colSid.VisibleIndex = 4
        Me.colSid.Width = 107
        '
        'txtSid
        '
        Me.txtSid.AutoHeight = False
        Me.txtSid.Name = "txtSid"
        '
        'colDBUserName
        '
        Me.colDBUserName.AppearanceCell.Options.UseTextOptions = True
        Me.colDBUserName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colDBUserName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBUserName.AppearanceHeader.Options.UseTextOptions = True
        Me.colDBUserName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colDBUserName.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBUserName.Caption = "使用者名稱"
        Me.colDBUserName.ColumnEdit = Me.txtDBUser
        Me.colDBUserName.FieldName = "DBUser"
        Me.colDBUserName.Name = "colDBUserName"
        Me.colDBUserName.Visible = True
        Me.colDBUserName.VisibleIndex = 5
        Me.colDBUserName.Width = 103
        '
        'txtDBUser
        '
        Me.txtDBUser.Appearance.Options.UseTextOptions = True
        Me.txtDBUser.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.txtDBUser.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.txtDBUser.AutoHeight = False
        Me.txtDBUser.Name = "txtDBUser"
        '
        'colDBPassword
        '
        Me.colDBPassword.AppearanceCell.Options.UseTextOptions = True
        Me.colDBPassword.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colDBPassword.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBPassword.AppearanceHeader.Options.UseTextOptions = True
        Me.colDBPassword.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colDBPassword.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBPassword.Caption = "資料庫密碼"
        Me.colDBPassword.ColumnEdit = Me.txtDBPassword
        Me.colDBPassword.FieldName = "DBPassword"
        Me.colDBPassword.Name = "colDBPassword"
        Me.colDBPassword.Visible = True
        Me.colDBPassword.VisibleIndex = 6
        Me.colDBPassword.Width = 100
        '
        'txtDBPassword
        '
        Me.txtDBPassword.AutoHeight = False
        Me.txtDBPassword.Name = "txtDBPassword"
        Me.txtDBPassword.NullValuePrompt = "資料庫密碼"
        Me.txtDBPassword.NullValuePromptShowForEmptyValue = True
        Me.txtDBPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        '
        'colMonitorType
        '
        Me.colMonitorType.Caption = "監控方式"
        Me.colMonitorType.ColumnEdit = Me.cmbMonitorType
        Me.colMonitorType.FieldName = "MonitorType"
        Me.colMonitorType.Name = "colMonitorType"
        Me.colMonitorType.Visible = True
        Me.colMonitorType.VisibleIndex = 10
        Me.colMonitorType.Width = 112
        '
        'cmbMonitorType
        '
        Me.cmbMonitorType.Appearance.Options.UseTextOptions = True
        Me.cmbMonitorType.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.cmbMonitorType.AutoHeight = False
        Me.cmbMonitorType.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cmbMonitorType.Items.AddRange(New Object() {"Web Service", "Web Post", "Web Get", "DB Access"})
        Me.cmbMonitorType.Name = "cmbMonitorType"
        Me.cmbMonitorType.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        '
        'colDBMinPool
        '
        Me.colDBMinPool.AppearanceCell.Options.UseTextOptions = True
        Me.colDBMinPool.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDBMinPool.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBMinPool.AppearanceHeader.Options.UseTextOptions = True
        Me.colDBMinPool.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDBMinPool.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBMinPool.Caption = "DB 最小集區"
        Me.colDBMinPool.ColumnEdit = Me.spnNum
        Me.colDBMinPool.FieldName = "DBMinPool"
        Me.colDBMinPool.Name = "colDBMinPool"
        Me.colDBMinPool.Visible = True
        Me.colDBMinPool.VisibleIndex = 7
        Me.colDBMinPool.Width = 81
        '
        'spnNum
        '
        Me.spnNum.AutoHeight = False
        Me.spnNum.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.spnNum.IsFloatValue = False
        Me.spnNum.Mask.EditMask = "N00"
        Me.spnNum.MaxValue = New Decimal(New Integer() {9999, 0, 0, 0})
        Me.spnNum.MinValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.spnNum.Name = "spnNum"
        Me.spnNum.NullText = "0"
        '
        'colDBMaxPool
        '
        Me.colDBMaxPool.AppearanceCell.Options.UseTextOptions = True
        Me.colDBMaxPool.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDBMaxPool.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBMaxPool.AppearanceHeader.Options.UseTextOptions = True
        Me.colDBMaxPool.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDBMaxPool.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBMaxPool.Caption = "DB 最大集區"
        Me.colDBMaxPool.ColumnEdit = Me.spnNum
        Me.colDBMaxPool.FieldName = "DBMaxPool"
        Me.colDBMaxPool.Name = "colDBMaxPool"
        Me.colDBMaxPool.Visible = True
        Me.colDBMaxPool.VisibleIndex = 8
        Me.colDBMaxPool.Width = 74
        '
        'colDBPoolLifetime
        '
        Me.colDBPoolLifetime.AppearanceCell.Options.UseTextOptions = True
        Me.colDBPoolLifetime.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDBPoolLifetime.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBPoolLifetime.AppearanceHeader.Options.UseTextOptions = True
        Me.colDBPoolLifetime.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colDBPoolLifetime.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colDBPoolLifetime.Caption = "集區 Lifetime"
        Me.colDBPoolLifetime.ColumnEdit = Me.spnNum
        Me.colDBPoolLifetime.FieldName = "DBLifetime"
        Me.colDBPoolLifetime.Name = "colDBPoolLifetime"
        Me.colDBPoolLifetime.Visible = True
        Me.colDBPoolLifetime.VisibleIndex = 9
        Me.colDBPoolLifetime.Width = 91
        '
        'colPrgName
        '
        Me.colPrgName.AppearanceCell.Options.UseTextOptions = True
        Me.colPrgName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colPrgName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colPrgName.AppearanceHeader.Options.UseTextOptions = True
        Me.colPrgName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colPrgName.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colPrgName.Caption = "監控名稱"
        Me.colPrgName.ColumnEdit = Me.txtPrgName
        Me.colPrgName.FieldName = "prgName"
        Me.colPrgName.Name = "colPrgName"
        Me.colPrgName.Visible = True
        Me.colPrgName.VisibleIndex = 3
        Me.colPrgName.Width = 179
        '
        'txtPrgName
        '
        Me.txtPrgName.AutoHeight = False
        Me.txtPrgName.Name = "txtPrgName"
        '
        'colRunTimer
        '
        Me.colRunTimer.AppearanceCell.Options.UseTextOptions = True
        Me.colRunTimer.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colRunTimer.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colRunTimer.AppearanceHeader.Options.UseTextOptions = True
        Me.colRunTimer.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colRunTimer.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colRunTimer.Caption = "監控頻率( 秒 )"
        Me.colRunTimer.ColumnEdit = Me.spnNum
        Me.colRunTimer.FieldName = "Runtimer"
        Me.colRunTimer.Name = "colRunTimer"
        Me.colRunTimer.Visible = True
        Me.colRunTimer.VisibleIndex = 11
        Me.colRunTimer.Width = 90
        '
        'colSeqSQL
        '
        Me.colSeqSQL.AppearanceCell.Options.UseTextOptions = True
        Me.colSeqSQL.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSeqSQL.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colSeqSQL.AppearanceHeader.Options.UseTextOptions = True
        Me.colSeqSQL.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSeqSQL.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colSeqSQL.Caption = "DB Sequence 語法"
        Me.colSeqSQL.ColumnEdit = Me.mdxCommon
        Me.colSeqSQL.FieldName = "SeqSql"
        Me.colSeqSQL.Name = "colSeqSQL"
        Me.colSeqSQL.Visible = True
        Me.colSeqSQL.VisibleIndex = 13
        Me.colSeqSQL.Width = 136
        '
        'mdxCommon
        '
        Me.mdxCommon.AutoHeight = False
        Me.mdxCommon.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.mdxCommon.Name = "mdxCommon"
        '
        'ColInsSql
        '
        Me.ColInsSql.AppearanceCell.Options.UseTextOptions = True
        Me.ColInsSql.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.ColInsSql.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.ColInsSql.AppearanceHeader.Options.UseTextOptions = True
        Me.ColInsSql.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.ColInsSql.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.ColInsSql.Caption = "Insert DB SQL 語法"
        Me.ColInsSql.ColumnEdit = Me.mdxCommon
        Me.ColInsSql.FieldName = "InsSql"
        Me.ColInsSql.Name = "ColInsSql"
        Me.ColInsSql.Visible = True
        Me.ColInsSql.VisibleIndex = 14
        Me.ColInsSql.Width = 126
        '
        'colTimeout
        '
        Me.colTimeout.AppearanceCell.Options.UseTextOptions = True
        Me.colTimeout.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colTimeout.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colTimeout.AppearanceHeader.Options.UseTextOptions = True
        Me.colTimeout.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colTimeout.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colTimeout.Caption = "判斷 Timeout 時間( 秒 )"
        Me.colTimeout.ColumnEdit = Me.spnNum
        Me.colTimeout.FieldName = "Timeout"
        Me.colTimeout.Name = "colTimeout"
        Me.colTimeout.Visible = True
        Me.colTimeout.VisibleIndex = 12
        Me.colTimeout.Width = 147
        '
        'colGetSql
        '
        Me.colGetSql.AppearanceCell.Options.UseTextOptions = True
        Me.colGetSql.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colGetSql.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colGetSql.AppearanceHeader.Options.UseTextOptions = True
        Me.colGetSql.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colGetSql.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colGetSql.Caption = "判斷狀態語法"
        Me.colGetSql.ColumnEdit = Me.mdxCommon
        Me.colGetSql.FieldName = "GetSql"
        Me.colGetSql.Name = "colGetSql"
        Me.colGetSql.Visible = True
        Me.colGetSql.VisibleIndex = 15
        Me.colGetSql.Width = 107
        '
        'colCompCode
        '
        Me.colCompCode.AppearanceCell.Options.UseTextOptions = True
        Me.colCompCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompCode.AppearanceHeader.Options.UseTextOptions = True
        Me.colCompCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompCode.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompCode.Caption = "公司別"
        Me.colCompCode.ColumnEdit = Me.spnNum
        Me.colCompCode.FieldName = "CompCode"
        Me.colCompCode.Name = "colCompCode"
        Me.colCompCode.Visible = True
        Me.colCompCode.VisibleIndex = 1
        '
        'colCompName
        '
        Me.colCompName.AppearanceCell.Options.UseTextOptions = True
        Me.colCompName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colCompName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompName.AppearanceHeader.Options.UseTextOptions = True
        Me.colCompName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colCompName.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompName.Caption = "公司名稱"
        Me.colCompName.ColumnEdit = Me.txtCommon
        Me.colCompName.FieldName = "CompName"
        Me.colCompName.Name = "colCompName"
        Me.colCompName.Visible = True
        Me.colCompName.VisibleIndex = 2
        Me.colCompName.Width = 197
        '
        'txtCommon
        '
        Me.txtCommon.AutoHeight = False
        Me.txtCommon.Name = "txtCommon"
        '
        'colWebServiceName
        '
        Me.colWebServiceName.AppearanceCell.Options.UseTextOptions = True
        Me.colWebServiceName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colWebServiceName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colWebServiceName.AppearanceHeader.Options.UseTextOptions = True
        Me.colWebServiceName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colWebServiceName.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colWebServiceName.Caption = "Web ServiceName"
        Me.colWebServiceName.ColumnEdit = Me.txtCommon
        Me.colWebServiceName.FieldName = "WebServiceName"
        Me.colWebServiceName.Name = "colWebServiceName"
        Me.colWebServiceName.Visible = True
        Me.colWebServiceName.VisibleIndex = 16
        Me.colWebServiceName.Width = 124
        '
        'colWebSrvMethod
        '
        Me.colWebSrvMethod.AppearanceCell.Options.UseTextOptions = True
        Me.colWebSrvMethod.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colWebSrvMethod.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colWebSrvMethod.AppearanceHeader.Options.UseTextOptions = True
        Me.colWebSrvMethod.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colWebSrvMethod.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colWebSrvMethod.Caption = "Web Service 功能"
        Me.colWebSrvMethod.ColumnEdit = Me.txtCommon
        Me.colWebSrvMethod.FieldName = "WebSrvMethod"
        Me.colWebSrvMethod.Name = "colWebSrvMethod"
        Me.colWebSrvMethod.Visible = True
        Me.colWebSrvMethod.VisibleIndex = 17
        Me.colWebSrvMethod.Width = 137
        '
        'colWebServicePara
        '
        Me.colWebServicePara.AppearanceCell.Options.UseTextOptions = True
        Me.colWebServicePara.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colWebServicePara.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colWebServicePara.AppearanceHeader.Options.UseTextOptions = True
        Me.colWebServicePara.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colWebServicePara.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colWebServicePara.Caption = "Web Service 參數"
        Me.colWebServicePara.ColumnEdit = Me.txtCommon
        Me.colWebServicePara.FieldName = "WebServicePara"
        Me.colWebServicePara.Name = "colWebServicePara"
        Me.colWebServicePara.Visible = True
        Me.colWebServicePara.VisibleIndex = 18
        Me.colWebServicePara.Width = 159
        '
        'colUrl
        '
        Me.colUrl.AppearanceCell.Options.UseTextOptions = True
        Me.colUrl.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colUrl.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colUrl.AppearanceHeader.Options.UseTextOptions = True
        Me.colUrl.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colUrl.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colUrl.Caption = "網址"
        Me.colUrl.ColumnEdit = Me.txtCommon
        Me.colUrl.FieldName = "url"
        Me.colUrl.Name = "colUrl"
        Me.colUrl.Visible = True
        Me.colUrl.VisibleIndex = 19
        Me.colUrl.Width = 413
        '
        'colWebGetPara
        '
        Me.colWebGetPara.AppearanceCell.Options.UseTextOptions = True
        Me.colWebGetPara.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colWebGetPara.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colWebGetPara.AppearanceHeader.Options.UseTextOptions = True
        Me.colWebGetPara.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colWebGetPara.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colWebGetPara.Caption = "Web Get 參數"
        Me.colWebGetPara.ColumnEdit = Me.txtCommon
        Me.colWebGetPara.FieldName = "WebGetPara"
        Me.colWebGetPara.Name = "colWebGetPara"
        Me.colWebGetPara.Visible = True
        Me.colWebGetPara.VisibleIndex = 20
        Me.colWebGetPara.Width = 186
        '
        'colWebPostPara
        '
        Me.colWebPostPara.AppearanceCell.Options.UseTextOptions = True
        Me.colWebPostPara.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colWebPostPara.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colWebPostPara.AppearanceHeader.Options.UseTextOptions = True
        Me.colWebPostPara.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colWebPostPara.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colWebPostPara.Caption = "Web Post 參數"
        Me.colWebPostPara.ColumnEdit = Me.txtCommon
        Me.colWebPostPara.FieldName = "WebPostPara"
        Me.colWebPostPara.Name = "colWebPostPara"
        Me.colWebPostPara.Visible = True
        Me.colWebPostPara.VisibleIndex = 21
        Me.colWebPostPara.Width = 178
        '
        'txtGatewayName
        '
        Me.txtGatewayName.AutoHeight = False
        Me.txtGatewayName.Name = "txtGatewayName"
        '
        'meErrorTxt
        '
        Me.meErrorTxt.Name = "meErrorTxt"
        '
        'cmbMonitorType1
        '
        Me.cmbMonitorType1.AutoHeight = False
        Me.cmbMonitorType1.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cmbMonitorType1.Items.AddRange(New DevExpress.XtraEditors.Controls.CheckedListBoxItem() {New DevExpress.XtraEditors.Controls.CheckedListBoxItem(0, "Web Service"), New DevExpress.XtraEditors.Controls.CheckedListBoxItem(1, "Web Post"), New DevExpress.XtraEditors.Controls.CheckedListBoxItem(2, "Web Get"), New DevExpress.XtraEditors.Controls.CheckedListBoxItem(3, "DB Access")})
        Me.cmbMonitorType1.Name = "cmbMonitorType1"
        Me.cmbMonitorType1.SelectAllItemCaption = "(全選)"
        Me.cmbMonitorType1.SelectAllItemVisible = False
        '
        'mdCommon
        '
        Me.mdCommon.Name = "mdCommon"
        '
        'rgMonitorType
        '
        Me.rgMonitorType.Items.AddRange(New DevExpress.XtraEditors.Controls.RadioGroupItem() {New DevExpress.XtraEditors.Controls.RadioGroupItem(0, "1235"), New DevExpress.XtraEditors.Controls.RadioGroupItem(1, "eeee")})
        Me.rgMonitorType.Name = "rgMonitorType"
        '
        'PanelControl1
        '
        Me.PanelControl1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl1.Controls.Add(Me.btnTest)
        Me.PanelControl1.Controls.Add(Me.btnDel)
        Me.PanelControl1.Controls.Add(Me.btnAdd)
        Me.PanelControl1.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelControl1.Location = New System.Drawing.Point(0, 0)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(872, 43)
        Me.PanelControl1.TabIndex = 13
        '
        'btnTest
        '
        Me.btnTest.Image = CType(resources.GetObject("btnTest.Image"), System.Drawing.Image)
        Me.btnTest.Location = New System.Drawing.Point(163, 5)
        Me.btnTest.Name = "btnTest"
        Me.btnTest.Size = New System.Drawing.Size(73, 31)
        Me.btnTest.TabIndex = 8
        Me.btnTest.Text = "測試"
        '
        'btnDel
        '
        Me.btnDel.Image = CType(resources.GetObject("btnDel.Image"), System.Drawing.Image)
        Me.btnDel.Location = New System.Drawing.Point(84, 5)
        Me.btnDel.Name = "btnDel"
        Me.btnDel.Size = New System.Drawing.Size(73, 31)
        Me.btnDel.TabIndex = 7
        Me.btnDel.Text = "刪除"
        '
        'btnAdd
        '
        Me.btnAdd.Image = CType(resources.GetObject("btnAdd.Image"), System.Drawing.Image)
        Me.btnAdd.Location = New System.Drawing.Point(5, 5)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(73, 31)
        Me.btnAdd.TabIndex = 2
        Me.btnAdd.Text = "新增"
        '
        'DefaultLookAndFeel1
        '
        Me.DefaultLookAndFeel1.LookAndFeel.SkinName = "Office 2010 Silver"
        '
        'frmOverseeSetting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(882, 565)
        Me.Controls.Add(Me.grpCtl)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.LookAndFeel.SkinName = "Office 2010 Black"
        Me.MaximizeBox = False
        Me.Name = "frmOverseeSetting"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Getway 監控程式設定"
        CType(Me.grpCtl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpCtl.ResumeLayout(False)
        Me.grpCtl.PerformLayout()
        CType(Me.spinLogHistorical.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StyleController1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spinMaxThreads.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.XtabCtl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtabCtl.ResumeLayout(False)
        Me.XtabCommonOpt.ResumeLayout(False)
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl2.ResumeLayout(False)
        CType(Me.TreLstDB, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkSelect, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSid, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDBUser, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDBPassword, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmbMonitorType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.spnNum, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtPrgName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mdxCommon, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCommon, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtGatewayName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.meErrorTxt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmbMonitorType1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mdCommon, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rgMonitorType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grpCtl As DevExpress.XtraEditors.GroupControl
    Friend WithEvents XtabCtl As DevExpress.XtraTab.XtraTabControl
    Friend WithEvents DefaultLookAndFeel1 As DevExpress.LookAndFeel.DefaultLookAndFeel
    Friend WithEvents StyleController1 As DevExpress.XtraEditors.StyleController
    Friend WithEvents btnSave As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents spinLogHistorical As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl2 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents spinMaxThreads As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents XtabCommonOpt As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents PanelControl2 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents TreLstDB As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colSelSO As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents chkSelect As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents colSid As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents txtSid As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents colDBUserName As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents txtDBUser As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents colDBPassword As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents txtDBPassword As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents colMonitorType As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents cmbMonitorType As DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Friend WithEvents colDBMinPool As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents spnNum As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents colDBMaxPool As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colDBPoolLifetime As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colPrgName As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents txtPrgName As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents colRunTimer As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colSeqSQL As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents mdxCommon As DevExpress.XtraEditors.Repository.RepositoryItemMemoExEdit
    Friend WithEvents ColInsSql As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colTimeout As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colGetSql As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colCompCode As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colCompName As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents txtCommon As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents colWebServiceName As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colWebSrvMethod As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colWebServicePara As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colUrl As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colWebGetPara As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colWebPostPara As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents txtGatewayName As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents meErrorTxt As DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit
    Friend WithEvents cmbMonitorType1 As DevExpress.XtraEditors.Repository.RepositoryItemCheckedComboBoxEdit
    Friend WithEvents mdCommon As DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit
    Friend WithEvents rgMonitorType As DevExpress.XtraEditors.Repository.RepositoryItemRadioGroup
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents btnDel As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAdd As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnTest As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnExit As DevExpress.XtraEditors.SimpleButton
End Class
