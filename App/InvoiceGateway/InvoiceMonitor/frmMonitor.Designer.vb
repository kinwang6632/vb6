<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMonitor
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMonitor))
        Me.imgCheck = New System.Windows.Forms.ImageList(Me.components)
        Me.DefaultLookAndFeel1 = New DevExpress.LookAndFeel.DefaultLookAndFeel(Me.components)
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        Me.trelstRunSO = New DevExpress.XtraTreeList.TreeList()
        Me.colCompCode = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colCompName = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colC0401 = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.RepositoryItemCheckEdit1 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.colC0501 = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colD0401 = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colD0501 = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colC0701 = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colX0401 = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.colB08 = New DevExpress.XtraTreeList.Columns.TreeListColumn()
        Me.BarManager1 = New DevExpress.XtraBars.BarManager(Me.components)
        Me.barMain = New DevExpress.XtraBars.Bar()
        Me.barbtnPlay = New DevExpress.XtraBars.BarButtonItem()
        Me.barbtnStop = New DevExpress.XtraBars.BarButtonItem()
        Me.barbtnSetting = New DevExpress.XtraBars.BarButtonItem()
        Me.barbtnExit = New DevExpress.XtraBars.BarButtonItem()
        Me.barStatus = New DevExpress.XtraBars.Bar()
        Me.barstcPath = New DevExpress.XtraBars.BarStaticItem()
        Me.BarStaticItem1 = New DevExpress.XtraBars.BarStaticItem()
        Me.barSpinReadTime = New DevExpress.XtraBars.BarEditItem()
        Me.RepositoryItemSpinEdit1 = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit()
        Me.BarStaticItem2 = New DevExpress.XtraBars.BarStaticItem()
        Me.barDockControlTop = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlBottom = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlLeft = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlRight = New DevExpress.XtraBars.BarDockControl()
        Me.BarEditItem1 = New DevExpress.XtraBars.BarEditItem()
        Me.RepositoryItemCheckEdit2 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.imgLarge = New DevExpress.Utils.ImageCollection(Me.components)
        Me.RepositoryItemTextEdit1 = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.RepositoryItemTextEdit2 = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        CType(Me.trelstRunSO, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BarManager1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemSpinEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgLarge, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemTextEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemTextEdit2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'imgCheck
        '
        Me.imgCheck.ImageStream = CType(resources.GetObject("imgCheck.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgCheck.TransparentColor = System.Drawing.Color.Transparent
        Me.imgCheck.Images.SetKeyName(0, "green")
        Me.imgCheck.Images.SetKeyName(1, "red")
        Me.imgCheck.Images.SetKeyName(2, "yellow")
        '
        'DefaultLookAndFeel1
        '
        Me.DefaultLookAndFeel1.LookAndFeel.SkinName = "Office 2010 Blue"
        '
        'PanelControl1
        '
        Me.PanelControl1.Controls.Add(Me.trelstRunSO)
        Me.PanelControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelControl1.Location = New System.Drawing.Point(0, 40)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(873, 196)
        Me.PanelControl1.TabIndex = 1
        '
        'trelstRunSO
        '
        Me.trelstRunSO.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.colCompCode, Me.colCompName, Me.colC0401, Me.colC0501, Me.colD0401, Me.colD0501, Me.colC0701, Me.colX0401, Me.colB08})
        Me.trelstRunSO.Dock = System.Windows.Forms.DockStyle.Fill
        Me.trelstRunSO.Location = New System.Drawing.Point(2, 2)
        Me.trelstRunSO.Name = "trelstRunSO"
        Me.trelstRunSO.OptionsBehavior.Editable = False
        Me.trelstRunSO.OptionsBehavior.ImmediateEditor = False
        Me.trelstRunSO.OptionsFilter.AllowFilterEditor = False
        Me.trelstRunSO.OptionsPrint.AutoRowHeight = False
        Me.trelstRunSO.OptionsPrint.AutoWidth = False
        Me.trelstRunSO.OptionsPrint.UsePrintStyles = True
        Me.trelstRunSO.OptionsView.AutoWidth = False
        Me.trelstRunSO.OptionsView.ShowFilterPanelMode = DevExpress.XtraTreeList.ShowFilterPanelMode.Never
        Me.trelstRunSO.OptionsView.ShowIndicator = False
        Me.trelstRunSO.OptionsView.ShowRoot = False
        Me.trelstRunSO.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.RepositoryItemCheckEdit1})
        Me.trelstRunSO.Size = New System.Drawing.Size(869, 192)
        Me.trelstRunSO.TabIndex = 1
        '
        'colCompCode
        '
        Me.colCompCode.AppearanceCell.Options.UseTextOptions = True
        Me.colCompCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompCode.AppearanceHeader.Options.UseTextOptions = True
        Me.colCompCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCompCode.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colCompCode.Caption = "系統台代號"
        Me.colCompCode.FieldName = "CompCode"
        Me.colCompCode.Name = "colCompCode"
        Me.colCompCode.OptionsColumn.AllowEdit = False
        Me.colCompCode.Visible = True
        Me.colCompCode.VisibleIndex = 0
        Me.colCompCode.Width = 74
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
        Me.colCompName.FieldName = "CompName"
        Me.colCompName.Name = "colCompName"
        Me.colCompName.Visible = True
        Me.colCompName.VisibleIndex = 1
        Me.colCompName.Width = 229
        '
        'colC0401
        '
        Me.colC0401.AppearanceCell.Options.UseTextOptions = True
        Me.colC0401.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colC0401.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colC0401.AppearanceHeader.Options.UseTextOptions = True
        Me.colC0401.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colC0401.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colC0401.Caption = "一般發票"
        Me.colC0401.ColumnEdit = Me.RepositoryItemCheckEdit1
        Me.colC0401.FieldName = "C0401"
        Me.colC0401.Name = "colC0401"
        Me.colC0401.Visible = True
        Me.colC0401.VisibleIndex = 2
        Me.colC0401.Width = 79
        '
        'RepositoryItemCheckEdit1
        '
        Me.RepositoryItemCheckEdit1.AllowGrayed = True
        Me.RepositoryItemCheckEdit1.AutoHeight = False
        Me.RepositoryItemCheckEdit1.CheckStyle = DevExpress.XtraEditors.Controls.CheckStyles.UserDefined
        Me.RepositoryItemCheckEdit1.ImageIndexChecked = 1
        Me.RepositoryItemCheckEdit1.ImageIndexGrayed = 2
        Me.RepositoryItemCheckEdit1.ImageIndexUnchecked = 0
        Me.RepositoryItemCheckEdit1.Images = Me.imgCheck
        Me.RepositoryItemCheckEdit1.LookAndFeel.SkinName = "Office 2010 Blue"
        Me.RepositoryItemCheckEdit1.Name = "RepositoryItemCheckEdit1"
        Me.RepositoryItemCheckEdit1.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Inactive
        Me.RepositoryItemCheckEdit1.ValueChecked = "1"
        Me.RepositoryItemCheckEdit1.ValueGrayed = ""
        Me.RepositoryItemCheckEdit1.ValueUnchecked = "0"
        '
        'colC0501
        '
        Me.colC0501.AppearanceCell.Options.UseTextOptions = True
        Me.colC0501.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colC0501.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colC0501.AppearanceHeader.Options.UseTextOptions = True
        Me.colC0501.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colC0501.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colC0501.Caption = "作廢發票"
        Me.colC0501.ColumnEdit = Me.RepositoryItemCheckEdit1
        Me.colC0501.FieldName = "C0501"
        Me.colC0501.Name = "colC0501"
        Me.colC0501.Visible = True
        Me.colC0501.VisibleIndex = 3
        '
        'colD0401
        '
        Me.colD0401.AppearanceCell.Options.UseTextOptions = True
        Me.colD0401.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colD0401.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colD0401.AppearanceHeader.Options.UseTextOptions = True
        Me.colD0401.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colD0401.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colD0401.Caption = "折讓發票"
        Me.colD0401.ColumnEdit = Me.RepositoryItemCheckEdit1
        Me.colD0401.FieldName = "D0401"
        Me.colD0401.Name = "colD0401"
        Me.colD0401.Visible = True
        Me.colD0401.VisibleIndex = 4
        '
        'colD0501
        '
        Me.colD0501.AppearanceCell.Options.UseTextOptions = True
        Me.colD0501.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colD0501.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colD0501.AppearanceHeader.Options.UseTextOptions = True
        Me.colD0501.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colD0501.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colD0501.Caption = "折讓作廢發票"
        Me.colD0501.ColumnEdit = Me.RepositoryItemCheckEdit1
        Me.colD0501.FieldName = "D0501"
        Me.colD0501.Name = "colD0501"
        Me.colD0501.Visible = True
        Me.colD0501.VisibleIndex = 5
        Me.colD0501.Width = 93
        '
        'colC0701
        '
        Me.colC0701.AppearanceCell.Options.UseTextOptions = True
        Me.colC0701.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colC0701.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colC0701.AppearanceHeader.Options.UseTextOptions = True
        Me.colC0701.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colC0701.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colC0701.Caption = "註銷發票"
        Me.colC0701.ColumnEdit = Me.RepositoryItemCheckEdit1
        Me.colC0701.FieldName = "C0701"
        Me.colC0701.Name = "colC0701"
        Me.colC0701.Visible = True
        Me.colC0701.VisibleIndex = 6
        '
        'colX0401
        '
        Me.colX0401.AppearanceCell.Options.UseTextOptions = True
        Me.colX0401.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colX0401.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colX0401.AppearanceHeader.Options.UseTextOptions = True
        Me.colX0401.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colX0401.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colX0401.Caption = "註銷重傳發票"
        Me.colX0401.ColumnEdit = Me.RepositoryItemCheckEdit1
        Me.colX0401.FieldName = "X0401"
        Me.colX0401.Name = "colX0401"
        Me.colX0401.Visible = True
        Me.colX0401.VisibleIndex = 7
        Me.colX0401.Width = 90
        '
        'colB08
        '
        Me.colB08.AppearanceCell.Options.UseTextOptions = True
        Me.colB08.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colB08.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colB08.AppearanceHeader.Options.UseTextOptions = True
        Me.colB08.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colB08.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.colB08.Caption = "發送簡訊"
        Me.colB08.ColumnEdit = Me.RepositoryItemCheckEdit1
        Me.colB08.FieldName = "B08"
        Me.colB08.Name = "colB08"
        Me.colB08.Visible = True
        Me.colB08.VisibleIndex = 8
        '
        'BarManager1
        '
        Me.BarManager1.Bars.AddRange(New DevExpress.XtraBars.Bar() {Me.barMain, Me.barStatus})
        Me.BarManager1.DockControls.Add(Me.barDockControlTop)
        Me.BarManager1.DockControls.Add(Me.barDockControlBottom)
        Me.BarManager1.DockControls.Add(Me.barDockControlLeft)
        Me.BarManager1.DockControls.Add(Me.barDockControlRight)
        Me.BarManager1.Form = Me
        Me.BarManager1.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.BarEditItem1, Me.barbtnPlay, Me.barbtnStop, Me.barbtnSetting, Me.barbtnExit, Me.barstcPath, Me.barSpinReadTime, Me.BarStaticItem1, Me.BarStaticItem2})
        Me.BarManager1.LargeImages = Me.imgLarge
        Me.BarManager1.MainMenu = Me.barMain
        Me.BarManager1.MaxItemId = 19
        Me.BarManager1.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.RepositoryItemCheckEdit2, Me.RepositoryItemTextEdit1, Me.RepositoryItemTextEdit2, Me.RepositoryItemSpinEdit1})
        Me.BarManager1.StatusBar = Me.barStatus
        '
        'barMain
        '
        Me.barMain.BarName = "Main menu"
        Me.barMain.DockCol = 0
        Me.barMain.DockRow = 0
        Me.barMain.DockStyle = DevExpress.XtraBars.BarDockStyle.Top
        Me.barMain.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.barbtnPlay, True), New DevExpress.XtraBars.LinkPersistInfo(Me.barbtnStop, True), New DevExpress.XtraBars.LinkPersistInfo(Me.barbtnSetting, True), New DevExpress.XtraBars.LinkPersistInfo(Me.barbtnExit, True)})
        Me.barMain.OptionsBar.AllowQuickCustomization = False
        Me.barMain.OptionsBar.MultiLine = True
        Me.barMain.OptionsBar.UseWholeRow = True
        Me.barMain.Text = "Main menu"
        '
        'barbtnPlay
        '
        Me.barbtnPlay.Glyph = CType(resources.GetObject("barbtnPlay.Glyph"), System.Drawing.Image)
        Me.barbtnPlay.Id = 5
        Me.barbtnPlay.Name = "barbtnPlay"
        '
        'barbtnStop
        '
        Me.barbtnStop.Glyph = CType(resources.GetObject("barbtnStop.Glyph"), System.Drawing.Image)
        Me.barbtnStop.Id = 7
        Me.barbtnStop.Name = "barbtnStop"
        '
        'barbtnSetting
        '
        Me.barbtnSetting.Glyph = CType(resources.GetObject("barbtnSetting.Glyph"), System.Drawing.Image)
        Me.barbtnSetting.Id = 8
        Me.barbtnSetting.Name = "barbtnSetting"
        '
        'barbtnExit
        '
        Me.barbtnExit.Glyph = CType(resources.GetObject("barbtnExit.Glyph"), System.Drawing.Image)
        Me.barbtnExit.Id = 9
        Me.barbtnExit.Name = "barbtnExit"
        '
        'barStatus
        '
        Me.barStatus.BarName = "Status bar"
        Me.barStatus.CanDockStyle = DevExpress.XtraBars.BarCanDockStyle.Bottom
        Me.barStatus.DockCol = 0
        Me.barStatus.DockRow = 0
        Me.barStatus.DockStyle = DevExpress.XtraBars.BarDockStyle.Bottom
        Me.barStatus.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.barstcPath), New DevExpress.XtraBars.LinkPersistInfo(Me.BarStaticItem1), New DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.Caption, Me.barSpinReadTime, "BarEditItem3"), New DevExpress.XtraBars.LinkPersistInfo(Me.BarStaticItem2)})
        Me.barStatus.OptionsBar.AllowQuickCustomization = False
        Me.barStatus.OptionsBar.DrawDragBorder = False
        Me.barStatus.OptionsBar.UseWholeRow = True
        Me.barStatus.Text = "Status bar"
        '
        'barstcPath
        '
        Me.barstcPath.Id = 10
        Me.barstcPath.Name = "barstcPath"
        Me.barstcPath.TextAlignment = System.Drawing.StringAlignment.Near
        '
        'BarStaticItem1
        '
        Me.BarStaticItem1.Alignment = DevExpress.XtraBars.BarItemLinkAlignment.Right
        Me.BarStaticItem1.Caption = "每"
        Me.BarStaticItem1.Id = 17
        Me.BarStaticItem1.Name = "BarStaticItem1"
        Me.BarStaticItem1.TextAlignment = System.Drawing.StringAlignment.Near
        '
        'barSpinReadTime
        '
        Me.barSpinReadTime.Alignment = DevExpress.XtraBars.BarItemLinkAlignment.Right
        Me.barSpinReadTime.Edit = Me.RepositoryItemSpinEdit1
        Me.barSpinReadTime.EditValue = "1"
        Me.barSpinReadTime.Id = 13
        Me.barSpinReadTime.Name = "barSpinReadTime"
        Me.barSpinReadTime.Width = 48
        '
        'RepositoryItemSpinEdit1
        '
        Me.RepositoryItemSpinEdit1.AutoHeight = False
        Me.RepositoryItemSpinEdit1.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.RepositoryItemSpinEdit1.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.RepositoryItemSpinEdit1.MaxValue = New Decimal(New Integer() {900, 0, 0, 0})
        Me.RepositoryItemSpinEdit1.MinValue = New Decimal(New Integer() {5, 0, 0, 65536})
        Me.RepositoryItemSpinEdit1.Name = "RepositoryItemSpinEdit1"
        '
        'BarStaticItem2
        '
        Me.BarStaticItem2.Caption = "秒監控一次"
        Me.BarStaticItem2.Id = 18
        Me.BarStaticItem2.Name = "BarStaticItem2"
        Me.BarStaticItem2.TextAlignment = System.Drawing.StringAlignment.Near
        '
        'barDockControlTop
        '
        Me.barDockControlTop.CausesValidation = False
        Me.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.barDockControlTop.Location = New System.Drawing.Point(0, 0)
        Me.barDockControlTop.Size = New System.Drawing.Size(873, 40)
        '
        'barDockControlBottom
        '
        Me.barDockControlBottom.CausesValidation = False
        Me.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.barDockControlBottom.Location = New System.Drawing.Point(0, 236)
        Me.barDockControlBottom.Size = New System.Drawing.Size(873, 26)
        '
        'barDockControlLeft
        '
        Me.barDockControlLeft.CausesValidation = False
        Me.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.barDockControlLeft.Location = New System.Drawing.Point(0, 40)
        Me.barDockControlLeft.Size = New System.Drawing.Size(0, 196)
        '
        'barDockControlRight
        '
        Me.barDockControlRight.CausesValidation = False
        Me.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.barDockControlRight.Location = New System.Drawing.Point(873, 40)
        Me.barDockControlRight.Size = New System.Drawing.Size(0, 196)
        '
        'BarEditItem1
        '
        Me.BarEditItem1.Edit = Me.RepositoryItemCheckEdit2
        Me.BarEditItem1.Id = 2
        Me.BarEditItem1.Name = "BarEditItem1"
        '
        'RepositoryItemCheckEdit2
        '
        Me.RepositoryItemCheckEdit2.AutoHeight = False
        Me.RepositoryItemCheckEdit2.Name = "RepositoryItemCheckEdit2"
        '
        'imgLarge
        '
        Me.imgLarge.ImageStream = CType(resources.GetObject("imgLarge.ImageStream"), DevExpress.Utils.ImageCollectionStreamer)
        Me.imgLarge.Images.SetKeyName(0, "player_play.png")
        '
        'RepositoryItemTextEdit1
        '
        Me.RepositoryItemTextEdit1.AutoHeight = False
        Me.RepositoryItemTextEdit1.Name = "RepositoryItemTextEdit1"
        '
        'RepositoryItemTextEdit2
        '
        Me.RepositoryItemTextEdit2.AutoHeight = False
        Me.RepositoryItemTextEdit2.Name = "RepositoryItemTextEdit2"
        '
        'BackgroundWorker1
        '
        '
        'frmMonitor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(873, 262)
        Me.Controls.Add(Me.PanelControl1)
        Me.Controls.Add(Me.barDockControlLeft)
        Me.Controls.Add(Me.barDockControlRight)
        Me.Controls.Add(Me.barDockControlBottom)
        Me.Controls.Add(Me.barDockControlTop)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMonitor"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "發票系統自動上傳監控程式"
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        CType(Me.trelstRunSO, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BarManager1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemSpinEdit1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgLarge, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemTextEdit1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemTextEdit2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents imgCheck As System.Windows.Forms.ImageList
    Friend WithEvents DefaultLookAndFeel1 As DevExpress.LookAndFeel.DefaultLookAndFeel
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents trelstRunSO As DevExpress.XtraTreeList.TreeList
    Friend WithEvents colCompCode As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colCompName As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colC0401 As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents RepositoryItemCheckEdit1 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents colC0501 As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colD0401 As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colD0501 As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colC0701 As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colX0401 As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents colB08 As DevExpress.XtraTreeList.Columns.TreeListColumn
    Friend WithEvents BarManager1 As DevExpress.XtraBars.BarManager
    Friend WithEvents barMain As DevExpress.XtraBars.Bar
    Friend WithEvents barStatus As DevExpress.XtraBars.Bar
    Friend WithEvents barDockControlTop As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlBottom As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlLeft As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlRight As DevExpress.XtraBars.BarDockControl
    Friend WithEvents BarEditItem1 As DevExpress.XtraBars.BarEditItem
    Friend WithEvents RepositoryItemCheckEdit2 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents RepositoryItemTextEdit1 As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents barbtnPlay As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents imgLarge As DevExpress.Utils.ImageCollection
    Friend WithEvents barbtnStop As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents RepositoryItemTextEdit2 As DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    Friend WithEvents barbtnSetting As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents barbtnExit As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents barstcPath As DevExpress.XtraBars.BarStaticItem
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents barSpinReadTime As DevExpress.XtraBars.BarEditItem
    Friend WithEvents RepositoryItemSpinEdit1 As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents BarStaticItem1 As DevExpress.XtraBars.BarStaticItem
    Friend WithEvents BarStaticItem2 As DevExpress.XtraBars.BarStaticItem
End Class
