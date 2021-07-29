<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class fmOversee
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
        Me.components = New System.ComponentModel.Container()
        Dim SerializableAppearanceObject1 As DevExpress.Utils.SerializableAppearanceObject = New DevExpress.Utils.SerializableAppearanceObject()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(fmOversee))
        Dim SuperToolTip1 As DevExpress.Utils.SuperToolTip = New DevExpress.Utils.SuperToolTip()
        Dim ToolTipTitleItem1 As DevExpress.Utils.ToolTipTitleItem = New DevExpress.Utils.ToolTipTitleItem()
        Dim ToolTipItem1 As DevExpress.Utils.ToolTipItem = New DevExpress.Utils.ToolTipItem()
        Dim SuperToolTip2 As DevExpress.Utils.SuperToolTip = New DevExpress.Utils.SuperToolTip()
        Dim ToolTipTitleItem2 As DevExpress.Utils.ToolTipTitleItem = New DevExpress.Utils.ToolTipTitleItem()
        Dim ToolTipItem2 As DevExpress.Utils.ToolTipItem = New DevExpress.Utils.ToolTipItem()
        Dim SuperToolTip3 As DevExpress.Utils.SuperToolTip = New DevExpress.Utils.SuperToolTip()
        Dim ToolTipTitleItem3 As DevExpress.Utils.ToolTipTitleItem = New DevExpress.Utils.ToolTipTitleItem()
        Dim ToolTipItem3 As DevExpress.Utils.ToolTipItem = New DevExpress.Utils.ToolTipItem()
        Dim SuperToolTip4 As DevExpress.Utils.SuperToolTip = New DevExpress.Utils.SuperToolTip()
        Dim ToolTipTitleItem4 As DevExpress.Utils.ToolTipTitleItem = New DevExpress.Utils.ToolTipTitleItem()
        Dim ToolTipItem4 As DevExpress.Utils.ToolTipItem = New DevExpress.Utils.ToolTipItem()
        Dim SuperToolTip5 As DevExpress.Utils.SuperToolTip = New DevExpress.Utils.SuperToolTip()
        Dim ToolTipTitleItem5 As DevExpress.Utils.ToolTipTitleItem = New DevExpress.Utils.ToolTipTitleItem()
        Dim ToolTipItem5 As DevExpress.Utils.ToolTipItem = New DevExpress.Utils.ToolTipItem()
        Dim SuperToolTip6 As DevExpress.Utils.SuperToolTip = New DevExpress.Utils.SuperToolTip()
        Dim ToolTipTitleItem6 As DevExpress.Utils.ToolTipTitleItem = New DevExpress.Utils.ToolTipTitleItem()
        Dim ToolTipItem6 As DevExpress.Utils.ToolTipItem = New DevExpress.Utils.ToolTipItem()
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        Me.treSO = New DevExpress.XtraTreeList.TreeList()
        Me.colCompId = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.txtComm = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.colCompName = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colprgName = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colStatus = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.icbStatus = New DevExpress.XtraEditors.Repository.RepositoryItemImageComboBox()
        Me.imgStatus = New System.Windows.Forms.ImageList(Me.components)
        Me.colSequence = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colErrmsg = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colOpenLog = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.btnOpenlog = New DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit()
        Me.edtEx = New DevExpress.XtraEditors.Repository.RepositoryItemMemoExEdit()
        Me.treld = New DevExpress.XtraEditors.Repository.RepositoryItemTreeListLookUpEdit()
        Me.lktrelst = New DevExpress.XtraTreeList.TreeList()
        Me.colMsgTime = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colMsg = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.cbMsg = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox()
        Me.imgCheck = New System.Windows.Forms.ImageList(Me.components)
        Me.bmMain = New DevExpress.XtraBars.BarManager(Me.components)
        Me.brButton = New DevExpress.XtraBars.Bar()
        Me.btnPlay = New DevExpress.XtraBars.BarButtonItem()
        Me.btnStop = New DevExpress.XtraBars.BarButtonItem()
        Me.btnSetting = New DevExpress.XtraBars.BarButtonItem()
        Me.btnRefresh = New DevExpress.XtraBars.BarButtonItem()
        Me.btnTopMost = New DevExpress.XtraBars.BarButtonItem()
        Me.btnExit = New DevExpress.XtraBars.BarButtonItem()
        Me.Bar3 = New DevExpress.XtraBars.Bar()
        Me.barDockControlTop = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlBottom = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlLeft = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlRight = New DevExpress.XtraBars.BarDockControl()
        Me.BarButtonItem1 = New DevExpress.XtraBars.BarButtonItem()
        Me.DefaultLookAndFeel1 = New DevExpress.LookAndFeel.DefaultLookAndFeel(Me.components)
        Me.bkWork = New System.ComponentModel.BackgroundWorker()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        CType(Me.treSO, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtComm, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.icbStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnOpenlog, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.edtEx, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.treld, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lktrelst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cbMsg, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.bmMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PanelControl1
        '
        Me.PanelControl1.Controls.Add(Me.treSO)
        Me.PanelControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelControl1.Location = New System.Drawing.Point(0, 56)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(718, 175)
        Me.PanelControl1.TabIndex = 0
        '
        'treSO
        '
        Me.treSO.Appearance.FocusedCell.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.treSO.Appearance.FocusedCell.BorderColor = System.Drawing.Color.Blue
        Me.treSO.Appearance.FocusedCell.Options.UseBackColor = True
        Me.treSO.Appearance.FocusedCell.Options.UseBorderColor = True
        Me.treSO.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colCompId, Me.colCompName, Me.colprgName, Me.colStatus, Me.colSequence, Me.colErrmsg, Me.colOpenLog})
        Me.treSO.Dock = System.Windows.Forms.DockStyle.Fill
        Me.treSO.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.treSO.KeyFieldName = ""
        Me.treSO.Location = New System.Drawing.Point(2, 2)
        Me.treSO.Name = "treSO"
        Me.treSO.OptionsBehavior.AllowQuickHideColumns = False
        Me.treSO.OptionsFilter.AllowColumnMRUFilterList = False
        Me.treSO.OptionsFilter.AllowFilterEditor = False
        Me.treSO.OptionsLayout.AddNewColumns = False
        Me.treSO.OptionsMenu.EnableFooterMenu = False
        Me.treSO.OptionsMenu.ShowAutoFilterRowItem = False
        Me.treSO.OptionsView.AutoWidth = False
        Me.treSO.OptionsView.ShowRoot = False
        Me.treSO.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.txtComm, Me.icbStatus, Me.edtEx, Me.btnOpenlog, Me.treld, Me.cbMsg})
        Me.treSO.ShowButtonMode = DevExpress.XtraTreeList.ShowButtonModeEnum.ShowForFocusedRow
        Me.treSO.Size = New System.Drawing.Size(714, 171)
        Me.treSO.StateImageList = Me.imgCheck
        Me.treSO.TabIndex = 0
        '
        'colCompId
        '
        Me.colCompId.AppearanceCell.Options.UseTextOptions = True
        Me.colCompId.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompId.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompId.AppearanceHeader.Options.UseTextOptions = True
        Me.colCompId.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompId.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompId.Caption = "公司別"
        Me.colCompId.ColumnEdit = Me.txtComm
        Me.colCompId.FieldName = "compId"
        Me.colCompId.MinWidth = 49
        Me.colCompId.Name = "colCompId"
        Me.colCompId.OptionsColumn.AllowEdit = False
        Me.colCompId.OptionsColumn.ReadOnly = True
        Me.colCompId.Visible = True
        Me.colCompId.VisibleIndex = 0
        Me.colCompId.Width = 76
        '
        'txtComm
        '
        Me.txtComm.AutoHeight = False
        Me.txtComm.Name = "txtComm"
        Me.txtComm.ReadOnly = True
        '
        'colCompName
        '
        Me.colCompName.AppearanceCell.Options.UseTextOptions = True
        Me.colCompName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colCompName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompName.AppearanceHeader.Options.UseTextOptions = True
        Me.colCompName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colCompName.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompName.Caption = "系統台名稱"
        Me.colCompName.ColumnEdit = Me.txtComm
        Me.colCompName.FieldName = "CompName"
        Me.colCompName.Name = "colCompName"
        Me.colCompName.OptionsColumn.AllowEdit = False
        Me.colCompName.OptionsColumn.ReadOnly = True
        Me.colCompName.Visible = True
        Me.colCompName.VisibleIndex = 1
        Me.colCompName.Width = 156
        '
        'colprgName
        '
        Me.colprgName.AppearanceCell.Options.UseTextOptions = True
        Me.colprgName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colprgName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colprgName.AppearanceHeader.Options.UseTextOptions = True
        Me.colprgName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colprgName.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colprgName.Caption = "監控名稱"
        Me.colprgName.ColumnEdit = Me.txtComm
        Me.colprgName.FieldName = "prgName"
        Me.colprgName.Name = "colprgName"
        Me.colprgName.OptionsColumn.AllowEdit = False
        Me.colprgName.OptionsColumn.ReadOnly = True
        Me.colprgName.Visible = True
        Me.colprgName.VisibleIndex = 2
        Me.colprgName.Width = 236
        '
        'colStatus
        '
        Me.colStatus.AppearanceCell.Options.UseTextOptions = True
        Me.colStatus.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colStatus.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colStatus.AppearanceHeader.Options.UseTextOptions = True
        Me.colStatus.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colStatus.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colStatus.Caption = "狀態"
        Me.colStatus.ColumnEdit = Me.icbStatus
        Me.colStatus.FieldName = "Status"
        Me.colStatus.ImageAlignment = System.Drawing.StringAlignment.Center
        Me.colStatus.Name = "colStatus"
        Me.colStatus.OptionsColumn.AllowEdit = False
        Me.colStatus.OptionsColumn.AllowFocus = False
        Me.colStatus.OptionsColumn.AllowSort = False
        Me.colStatus.OptionsFilter.AllowAutoFilter = False
        Me.colStatus.OptionsFilter.AllowFilter = False
        Me.colStatus.Visible = True
        Me.colStatus.VisibleIndex = 3
        Me.colStatus.Width = 54
        '
        'icbStatus
        '
        Me.icbStatus.Appearance.Options.UseTextOptions = True
        Me.icbStatus.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.icbStatus.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.icbStatus.AppearanceReadOnly.Options.UseTextOptions = True
        Me.icbStatus.AppearanceReadOnly.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.icbStatus.AppearanceReadOnly.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.icbStatus.AutoHeight = False
        Me.icbStatus.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo, "", -1, True, False, False, DevExpress.XtraEditors.ImageLocation.MiddleCenter, Nothing, New DevExpress.Utils.KeyShortcut(System.Windows.Forms.Keys.None), SerializableAppearanceObject1, "", Nothing, Nothing, True)})
        Me.icbStatus.GlyphAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.icbStatus.ImmediatePopup = True
        Me.icbStatus.Items.AddRange(New DevExpress.XtraEditors.Controls.ImageComboBoxItem() {New DevExpress.XtraEditors.Controls.ImageComboBoxItem("", 0, 0), New DevExpress.XtraEditors.Controls.ImageComboBoxItem("", 1, 1), New DevExpress.XtraEditors.Controls.ImageComboBoxItem("", 2, 2), New DevExpress.XtraEditors.Controls.ImageComboBoxItem("", 3, 3)})
        Me.icbStatus.LargeImages = Me.imgStatus
        Me.icbStatus.Name = "icbStatus"
        Me.icbStatus.ReadOnly = True
        Me.icbStatus.SmallImages = Me.imgStatus
        '
        'imgStatus
        '
        Me.imgStatus.ImageStream = CType(resources.GetObject("imgStatus.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgStatus.TransparentColor = System.Drawing.Color.Transparent
        Me.imgStatus.Images.SetKeyName(0, "green")
        Me.imgStatus.Images.SetKeyName(1, "red")
        Me.imgStatus.Images.SetKeyName(2, "yellow")
        Me.imgStatus.Images.SetKeyName(3, "1458215612_Pause.png")
        '
        'colSequence
        '
        Me.colSequence.Caption = "序號"
        Me.colSequence.FieldName = "Sequence"
        Me.colSequence.Name = "colSequence"
        Me.colSequence.OptionsColumn.AllowEdit = False
        Me.colSequence.OptionsColumn.AllowFocus = False
        Me.colSequence.OptionsColumn.ReadOnly = True
        '
        'colErrmsg
        '
        Me.colErrmsg.AllowIncrementalSearch = False
        Me.colErrmsg.AppearanceCell.Options.UseTextOptions = True
        Me.colErrmsg.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colErrmsg.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colErrmsg.AppearanceHeader.Options.UseTextOptions = True
        Me.colErrmsg.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colErrmsg.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colErrmsg.Caption = "錯誤訊息"
        Me.colErrmsg.FieldName = "errMsg"
        Me.colErrmsg.Name = "colErrmsg"
        Me.colErrmsg.OptionsColumn.AllowSort = False
        Me.colErrmsg.OptionsColumn.ReadOnly = True
        Me.colErrmsg.OptionsFilter.AllowAutoFilter = False
        Me.colErrmsg.OptionsFilter.AllowFilter = False
        Me.colErrmsg.Visible = True
        Me.colErrmsg.VisibleIndex = 4
        Me.colErrmsg.Width = 330
        '
        'colOpenLog
        '
        Me.colOpenLog.AppearanceCell.Options.UseTextOptions = True
        Me.colOpenLog.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colOpenLog.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colOpenLog.AppearanceHeader.Options.UseTextOptions = True
        Me.colOpenLog.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colOpenLog.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colOpenLog.Caption = "開啟記錄檔目錄"
        Me.colOpenLog.ColumnEdit = Me.btnOpenlog
        Me.colOpenLog.FieldName = "openLog"
        Me.colOpenLog.Name = "colOpenLog"
        Me.colOpenLog.Visible = True
        Me.colOpenLog.VisibleIndex = 5
        Me.colOpenLog.Width = 109
        '
        'btnOpenlog
        '
        Me.btnOpenlog.AutoHeight = False
        Me.btnOpenlog.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.btnOpenlog.Name = "btnOpenlog"
        Me.btnOpenlog.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.HideTextEditor
        '
        'edtEx
        '
        Me.edtEx.AutoHeight = False
        Me.edtEx.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.edtEx.Name = "edtEx"
        Me.edtEx.ReadOnly = True
        '
        'treld
        '
        Me.treld.AutoHeight = False
        Me.treld.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.treld.Name = "treld"
        Me.treld.NullText = ""
        Me.treld.TreeList = Me.lktrelst
        '
        'lktrelst
        '
        Me.lktrelst.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colMsgTime, Me.colMsg})
        Me.lktrelst.Location = New System.Drawing.Point(0, 0)
        Me.lktrelst.Name = "lktrelst"
        Me.lktrelst.OptionsBehavior.AllowQuickHideColumns = False
        Me.lktrelst.OptionsBehavior.AutoNodeHeight = False
        Me.lktrelst.OptionsBehavior.AutoSelectAllInEditor = False
        Me.lktrelst.OptionsBehavior.EnableFiltering = True
        Me.lktrelst.OptionsBehavior.PopulateServiceColumns = True
        Me.lktrelst.OptionsFilter.AllowFilterEditor = False
        Me.lktrelst.OptionsFind.ShowClearButton = False
        Me.lktrelst.OptionsFind.ShowCloseButton = False
        Me.lktrelst.OptionsFind.ShowFindButton = False
        Me.lktrelst.OptionsLayout.AddNewColumns = False
        Me.lktrelst.OptionsMenu.EnableColumnMenu = False
        Me.lktrelst.OptionsMenu.EnableFooterMenu = False
        Me.lktrelst.OptionsMenu.ShowAutoFilterRowItem = False
        Me.lktrelst.OptionsView.AutoWidth = False
        Me.lktrelst.OptionsView.ShowRoot = False
        Me.lktrelst.Size = New System.Drawing.Size(400, 200)
        Me.lktrelst.TabIndex = 0
        '
        'colMsgTime
        '
        Me.colMsgTime.AppearanceCell.Options.UseTextOptions = True
        Me.colMsgTime.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colMsgTime.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMsgTime.AppearanceHeader.Options.UseTextOptions = True
        Me.colMsgTime.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colMsgTime.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMsgTime.Caption = "訊息時間"
        Me.colMsgTime.FieldName = "MsgTime"
        Me.colMsgTime.Name = "colMsgTime"
        Me.colMsgTime.OptionsColumn.AllowEdit = False
        Me.colMsgTime.OptionsColumn.AllowFocus = False
        Me.colMsgTime.OptionsColumn.AllowMove = False
        Me.colMsgTime.OptionsColumn.AllowSort = False
        Me.colMsgTime.OptionsColumn.ReadOnly = True
        Me.colMsgTime.OptionsFilter.AllowAutoFilter = False
        Me.colMsgTime.OptionsFilter.AllowFilter = False
        Me.colMsgTime.OptionsFilter.ImmediateUpdateAutoFilter = False
        Me.colMsgTime.Visible = True
        Me.colMsgTime.VisibleIndex = 0
        Me.colMsgTime.Width = 150
        '
        'colMsg
        '
        Me.colMsg.AppearanceCell.Options.UseTextOptions = True
        Me.colMsg.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colMsg.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMsg.AppearanceHeader.Options.UseTextOptions = True
        Me.colMsg.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.colMsg.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colMsg.Caption = "訊息"
        Me.colMsg.FieldName = "Msg"
        Me.colMsg.Name = "colMsg"
        Me.colMsg.OptionsColumn.AllowEdit = False
        Me.colMsg.OptionsColumn.AllowMove = False
        Me.colMsg.OptionsColumn.AllowSort = False
        Me.colMsg.OptionsColumn.ReadOnly = True
        Me.colMsg.OptionsFilter.AllowAutoFilter = False
        Me.colMsg.OptionsFilter.AllowFilter = False
        Me.colMsg.Visible = True
        Me.colMsg.VisibleIndex = 1
        Me.colMsg.Width = 999
        '
        'cbMsg
        '
        Me.cbMsg.AutoHeight = False
        Me.cbMsg.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cbMsg.Name = "cbMsg"
        '
        'imgCheck
        '
        Me.imgCheck.ImageStream = CType(resources.GetObject("imgCheck.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgCheck.TransparentColor = System.Drawing.Color.Transparent
        Me.imgCheck.Images.SetKeyName(0, "1458657203_player_stop.ico")
        Me.imgCheck.Images.SetKeyName(1, "1458656990_clean.ico")
        Me.imgCheck.Images.SetKeyName(2, "1458656919_warning.ico")
        '
        'bmMain
        '
        Me.bmMain.Bars.AddRange(New DevExpress.XtraBars.Bar() {Me.brButton, Me.Bar3})
        Me.bmMain.DockControls.Add(Me.barDockControlTop)
        Me.bmMain.DockControls.Add(Me.barDockControlBottom)
        Me.bmMain.DockControls.Add(Me.barDockControlLeft)
        Me.bmMain.DockControls.Add(Me.barDockControlRight)
        Me.bmMain.Form = Me
        Me.bmMain.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.btnPlay, Me.btnStop, Me.btnSetting, Me.BarButtonItem1, Me.btnExit, Me.btnRefresh, Me.btnTopMost})
        Me.bmMain.MainMenu = Me.brButton
        Me.bmMain.MaxItemId = 8
        Me.bmMain.StatusBar = Me.Bar3
        '
        'brButton
        '
        Me.brButton.BarName = "Main menu"
        Me.brButton.DockCol = 0
        Me.brButton.DockRow = 0
        Me.brButton.DockStyle = DevExpress.XtraBars.BarDockStyle.Top
        Me.brButton.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, Me.btnPlay, DevExpress.XtraBars.BarItemPaintStyle.CaptionInMenu), New DevExpress.XtraBars.LinkPersistInfo(Me.btnStop), New DevExpress.XtraBars.LinkPersistInfo(Me.btnSetting), New DevExpress.XtraBars.LinkPersistInfo(Me.btnRefresh), New DevExpress.XtraBars.LinkPersistInfo(Me.btnTopMost), New DevExpress.XtraBars.LinkPersistInfo(Me.btnExit)})
        Me.brButton.OptionsBar.AllowQuickCustomization = False
        Me.brButton.OptionsBar.DrawDragBorder = False
        Me.brButton.OptionsBar.MultiLine = True
        Me.brButton.OptionsBar.UseWholeRow = True
        Me.brButton.Text = "Main menu"
        '
        'btnPlay
        '
        Me.btnPlay.Caption = "啟動"
        Me.btnPlay.Glyph = CType(resources.GetObject("btnPlay.Glyph"), System.Drawing.Image)
        Me.btnPlay.Id = 1
        Me.btnPlay.LargeGlyph = CType(resources.GetObject("btnPlay.LargeGlyph"), System.Drawing.Image)
        Me.btnPlay.Name = "btnPlay"
        ToolTipTitleItem1.Text = "啟動"
        ToolTipItem1.LeftIndent = 6
        ToolTipItem1.Text = "啟動監控程式"
        SuperToolTip1.Items.Add(ToolTipTitleItem1)
        SuperToolTip1.Items.Add(ToolTipItem1)
        Me.btnPlay.SuperTip = SuperToolTip1
        '
        'btnStop
        '
        Me.btnStop.Caption = "停止"
        Me.btnStop.Enabled = False
        Me.btnStop.Glyph = CType(resources.GetObject("btnStop.Glyph"), System.Drawing.Image)
        Me.btnStop.Id = 2
        Me.btnStop.LargeGlyph = CType(resources.GetObject("btnStop.LargeGlyph"), System.Drawing.Image)
        Me.btnStop.Name = "btnStop"
        ToolTipTitleItem2.Text = "停止"
        ToolTipItem2.LeftIndent = 6
        ToolTipItem2.Text = "停止監控"
        SuperToolTip2.Items.Add(ToolTipTitleItem2)
        SuperToolTip2.Items.Add(ToolTipItem2)
        Me.btnStop.SuperTip = SuperToolTip2
        '
        'btnSetting
        '
        Me.btnSetting.Caption = " 設定"
        Me.btnSetting.Glyph = CType(resources.GetObject("btnSetting.Glyph"), System.Drawing.Image)
        Me.btnSetting.Id = 3
        Me.btnSetting.LargeGlyph = CType(resources.GetObject("btnSetting.LargeGlyph"), System.Drawing.Image)
        Me.btnSetting.Name = "btnSetting"
        ToolTipTitleItem3.Text = "設定"
        ToolTipItem3.LeftIndent = 6
        ToolTipItem3.Text = "設定監控參數"
        SuperToolTip3.Items.Add(ToolTipTitleItem3)
        SuperToolTip3.Items.Add(ToolTipItem3)
        Me.btnSetting.SuperTip = SuperToolTip3
        '
        'btnRefresh
        '
        Me.btnRefresh.Caption = "重新讀取設定"
        Me.btnRefresh.Glyph = CType(resources.GetObject("btnRefresh.Glyph"), System.Drawing.Image)
        Me.btnRefresh.Id = 6
        Me.btnRefresh.Name = "btnRefresh"
        ToolTipTitleItem4.Text = "更新"
        ToolTipItem4.LeftIndent = 6
        ToolTipItem4.Text = "重新讀取設定檔"
        SuperToolTip4.Items.Add(ToolTipTitleItem4)
        SuperToolTip4.Items.Add(ToolTipItem4)
        Me.btnRefresh.SuperTip = SuperToolTip4
        Me.btnRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        '
        'btnTopMost
        '
        Me.btnTopMost.Caption = "最上層顯示"
        Me.btnTopMost.Glyph = CType(resources.GetObject("btnTopMost.Glyph"), System.Drawing.Image)
        Me.btnTopMost.Id = 7
        Me.btnTopMost.Name = "btnTopMost"
        ToolTipTitleItem5.Text = "最上層顯示"
        ToolTipItem5.LeftIndent = 6
        ToolTipItem5.Text = "一直顯示在最上層"
        SuperToolTip5.Items.Add(ToolTipTitleItem5)
        SuperToolTip5.Items.Add(ToolTipItem5)
        Me.btnTopMost.SuperTip = SuperToolTip5
        '
        'btnExit
        '
        Me.btnExit.Caption = "離開"
        Me.btnExit.Glyph = CType(resources.GetObject("btnExit.Glyph"), System.Drawing.Image)
        Me.btnExit.Id = 5
        Me.btnExit.LargeGlyph = CType(resources.GetObject("btnExit.LargeGlyph"), System.Drawing.Image)
        Me.btnExit.Name = "btnExit"
        ToolTipTitleItem6.Text = "離開"
        ToolTipItem6.LeftIndent = 6
        ToolTipItem6.Text = "離開監控程式"
        SuperToolTip6.Items.Add(ToolTipTitleItem6)
        SuperToolTip6.Items.Add(ToolTipItem6)
        Me.btnExit.SuperTip = SuperToolTip6
        '
        'Bar3
        '
        Me.Bar3.BarName = "Status bar"
        Me.Bar3.CanDockStyle = DevExpress.XtraBars.BarCanDockStyle.Bottom
        Me.Bar3.DockCol = 0
        Me.Bar3.DockRow = 0
        Me.Bar3.DockStyle = DevExpress.XtraBars.BarDockStyle.Bottom
        Me.Bar3.OptionsBar.AllowQuickCustomization = False
        Me.Bar3.OptionsBar.DrawDragBorder = False
        Me.Bar3.OptionsBar.UseWholeRow = True
        Me.Bar3.Text = "Status bar"
        Me.Bar3.Visible = False
        '
        'barDockControlTop
        '
        Me.barDockControlTop.CausesValidation = False
        Me.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.barDockControlTop.Location = New System.Drawing.Point(0, 0)
        Me.barDockControlTop.Size = New System.Drawing.Size(718, 56)
        '
        'barDockControlBottom
        '
        Me.barDockControlBottom.CausesValidation = False
        Me.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.barDockControlBottom.Location = New System.Drawing.Point(0, 231)
        Me.barDockControlBottom.Size = New System.Drawing.Size(718, 21)
        '
        'barDockControlLeft
        '
        Me.barDockControlLeft.CausesValidation = False
        Me.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.barDockControlLeft.Location = New System.Drawing.Point(0, 56)
        Me.barDockControlLeft.Size = New System.Drawing.Size(0, 175)
        '
        'barDockControlRight
        '
        Me.barDockControlRight.CausesValidation = False
        Me.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.barDockControlRight.Location = New System.Drawing.Point(718, 56)
        Me.barDockControlRight.Size = New System.Drawing.Size(0, 175)
        '
        'BarButtonItem1
        '
        Me.BarButtonItem1.Caption = "1235"
        Me.BarButtonItem1.Id = 4
        Me.BarButtonItem1.Name = "BarButtonItem1"
        '
        'DefaultLookAndFeel1
        '
        Me.DefaultLookAndFeel1.LookAndFeel.SkinName = "Office 2010 Silver"
        '
        'bkWork
        '
        '
        'fmOversee
        '
        Me.Appearance.Options.UseTextOptions = True
        Me.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(718, 252)
        Me.Controls.Add(Me.PanelControl1)
        Me.Controls.Add(Me.barDockControlLeft)
        Me.Controls.Add(Me.barDockControlRight)
        Me.Controls.Add(Me.barDockControlBottom)
        Me.Controls.Add(Me.barDockControlTop)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "fmOversee"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Gateway Monitor Application"
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        CType(Me.treSO, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtComm, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.icbStatus, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnOpenlog, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.edtEx, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.treld, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lktrelst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cbMsg, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.bmMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents bmMain As DevExpress.XtraBars.BarManager
    Friend WithEvents brButton As DevExpress.XtraBars.Bar
    Friend WithEvents Bar3 As DevExpress.XtraBars.Bar
    Friend WithEvents barDockControlTop As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlBottom As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlLeft As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlRight As DevExpress.XtraBars.BarDockControl
    Friend WithEvents btnPlay As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnStop As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnSetting As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents treSO As DevExpress.XtraTreeList.TreeList
    Friend WithEvents BarButtonItem1 As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents DefaultLookAndFeel1 As DevExpress.LookAndFeel.DefaultLookAndFeel
    Friend WithEvents colCompId As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colCompName As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colprgName As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents imgCheck As System.Windows.Forms.ImageList
    Friend WithEvents colStatus As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents btnExit As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents txtComm As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents btnRefresh As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bkWork As System.ComponentModel.BackgroundWorker
    Friend WithEvents imgStatus As System.Windows.Forms.ImageList
    Friend WithEvents colSequence As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colErrmsg As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents icbStatus As DevExpress.XtraEditors.Repository.RepositoryItemImageComboBox
    Friend WithEvents edtEx As DevExpress.XtraEditors.Repository.RepositoryItemMemoExEdit
    Friend WithEvents colOpenLog As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents btnOpenlog As DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit
    Friend WithEvents treld As DevExpress.XtraEditors.Repository.RepositoryItemTreeListLookUpEdit
    Friend WithEvents lktrelst As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colMsgTime As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colMsg As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents cbMsg As DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Friend WithEvents btnTopMost As DevExpress.XtraBars.BarButtonItem

End Class
