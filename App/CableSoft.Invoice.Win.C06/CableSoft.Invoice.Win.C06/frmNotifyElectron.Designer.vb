<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNotifyElectron
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNotifyElectron))
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl
        Me.PanelControl2 = New DevExpress.XtraEditors.PanelControl
        Me.btnExit = New DevExpress.XtraEditors.SimpleButton
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.btnDesign = New DevExpress.XtraEditors.SimpleButton
        Me.btnOK = New DevExpress.XtraEditors.SimpleButton
        Me.DefaultLookAndFeel1 = New DevExpress.LookAndFeel.DefaultLookAndFeel(Me.components)
        Me.btnShowSQL = New DevExpress.XtraEditors.SimpleButton
        Me.grpOption = New DevExpress.XtraEditors.GroupControl
        Me.edtInvId2 = New DevExpress.XtraEditors.TextEdit
        Me.LabelControl6 = New DevExpress.XtraEditors.LabelControl
        Me.LabelControl5 = New DevExpress.XtraEditors.LabelControl
        Me.grpComp = New DevExpress.XtraEditors.GroupControl
        Me.chkAllComp = New DevExpress.XtraEditors.CheckEdit
        Me.chkSingleComp = New DevExpress.XtraEditors.CheckEdit
        Me.LabelControl4 = New DevExpress.XtraEditors.LabelControl
        Me.LabelControl3 = New DevExpress.XtraEditors.LabelControl
        Me.GroupControl1 = New DevExpress.XtraEditors.GroupControl
        Me.chkNotifyCnt = New DevExpress.XtraEditors.CheckEdit
        Me.chkNotifyDetail = New DevExpress.XtraEditors.CheckEdit
        Me.dtInv2 = New DevExpress.XtraEditors.DateEdit
        Me.LabelControl2 = New DevExpress.XtraEditors.LabelControl
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl
        Me.dtInv1 = New DevExpress.XtraEditors.DateEdit
        Me.edtInvId1 = New DevExpress.XtraEditors.TextEdit
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl2.SuspendLayout()
        CType(Me.grpOption, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpOption.SuspendLayout()
        CType(Me.edtInvId2.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpComp, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpComp.SuspendLayout()
        CType(Me.chkAllComp.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkSingleComp.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GroupControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupControl1.SuspendLayout()
        CType(Me.chkNotifyCnt.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkNotifyDetail.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtInv2.Properties.VistaTimeProperties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtInv2.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtInv1.Properties.VistaTimeProperties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtInv1.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.edtInvId1.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PanelControl1
        '
        Me.PanelControl1.Controls.Add(Me.grpOption)
        Me.PanelControl1.Controls.Add(Me.PanelControl2)
        Me.PanelControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelControl1.Location = New System.Drawing.Point(0, 0)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(329, 222)
        Me.PanelControl1.TabIndex = 5
        '
        'PanelControl2
        '
        Me.PanelControl2.Controls.Add(Me.btnShowSQL)
        Me.PanelControl2.Controls.Add(Me.btnExit)
        Me.PanelControl2.Controls.Add(Me.btnDesign)
        Me.PanelControl2.Controls.Add(Me.btnOK)
        Me.PanelControl2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelControl2.Location = New System.Drawing.Point(2, 176)
        Me.PanelControl2.Name = "PanelControl2"
        Me.PanelControl2.Size = New System.Drawing.Size(325, 44)
        Me.PanelControl2.TabIndex = 7
        '
        'btnExit
        '
        Me.btnExit.ImageIndex = 0
        Me.btnExit.ImageList = Me.ImageList1
        Me.btnExit.Location = New System.Drawing.Point(87, 10)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(74, 27)
        Me.btnExit.TabIndex = 10
        Me.btnExit.Text = "離開(&X)"
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "1297823327_exit.png")
        Me.ImageList1.Images.SetKeyName(1, "1297823140_OK.png")
        '
        'btnDesign
        '
        Me.btnDesign.Location = New System.Drawing.Point(255, 10)
        Me.btnDesign.Name = "btnDesign"
        Me.btnDesign.Size = New System.Drawing.Size(63, 27)
        Me.btnDesign.TabIndex = 99
        Me.btnDesign.Text = "報表設計"
        Me.btnDesign.Visible = False
        '
        'btnOK
        '
        Me.btnOK.ImageIndex = 1
        Me.btnOK.ImageList = Me.ImageList1
        Me.btnOK.Location = New System.Drawing.Point(7, 10)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(74, 27)
        Me.btnOK.TabIndex = 9
        Me.btnOK.Text = "F2.列印"
        '
        'btnShowSQL
        '
        Me.btnShowSQL.Location = New System.Drawing.Point(190, 10)
        Me.btnShowSQL.Name = "btnShowSQL"
        Me.btnShowSQL.Size = New System.Drawing.Size(59, 27)
        Me.btnShowSQL.TabIndex = 100
        Me.btnShowSQL.Text = "顯示SQL"
        Me.btnShowSQL.Visible = False
        '
        'grpOption
        '
        Me.grpOption.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003
        Me.grpOption.CaptionImage = Global.CableSoft.Invoice.Win.C06.My.Resources.Resources._Option
        Me.grpOption.Controls.Add(Me.edtInvId2)
        Me.grpOption.Controls.Add(Me.LabelControl6)
        Me.grpOption.Controls.Add(Me.LabelControl5)
        Me.grpOption.Controls.Add(Me.grpComp)
        Me.grpOption.Controls.Add(Me.LabelControl4)
        Me.grpOption.Controls.Add(Me.LabelControl3)
        Me.grpOption.Controls.Add(Me.GroupControl1)
        Me.grpOption.Controls.Add(Me.dtInv2)
        Me.grpOption.Controls.Add(Me.LabelControl2)
        Me.grpOption.Controls.Add(Me.LabelControl1)
        Me.grpOption.Controls.Add(Me.dtInv1)
        Me.grpOption.Controls.Add(Me.edtInvId1)
        Me.grpOption.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpOption.Location = New System.Drawing.Point(2, 2)
        Me.grpOption.Name = "grpOption"
        Me.grpOption.Size = New System.Drawing.Size(325, 174)
        Me.grpOption.TabIndex = 0
        Me.grpOption.Text = "  選項設定"
        '
        'edtInvId2
        '
        Me.edtInvId2.Location = New System.Drawing.Point(216, 62)
        Me.edtInvId2.Name = "edtInvId2"
        Me.edtInvId2.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom
        Me.edtInvId2.Properties.Mask.EditMask = "\p{Lu}\p{Lu}\d{8}"
        Me.edtInvId2.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.RegEx
        Me.edtInvId2.Properties.Mask.ShowPlaceHolders = False
        Me.edtInvId2.Size = New System.Drawing.Size(100, 21)
        Me.edtInvId2.TabIndex = 4
        '
        'LabelControl6
        '
        Me.LabelControl6.Location = New System.Drawing.Point(201, 65)
        Me.LabelControl6.Name = "LabelControl6"
        Me.LabelControl6.Size = New System.Drawing.Size(9, 14)
        Me.LabelControl6.TabIndex = 10
        Me.LabelControl6.Text = "~"
        '
        'LabelControl5
        '
        Me.LabelControl5.Location = New System.Drawing.Point(15, 65)
        Me.LabelControl5.Name = "LabelControl5"
        Me.LabelControl5.Size = New System.Drawing.Size(72, 14)
        Me.LabelControl5.TabIndex = 8
        Me.LabelControl5.Text = "發票號碼起訖"
        '
        'grpComp
        '
        Me.grpComp.Controls.Add(Me.chkAllComp)
        Me.grpComp.Controls.Add(Me.chkSingleComp)
        Me.grpComp.Location = New System.Drawing.Point(70, 133)
        Me.grpComp.Name = "grpComp"
        Me.grpComp.ShowCaption = False
        Me.grpComp.Size = New System.Drawing.Size(246, 33)
        Me.grpComp.TabIndex = 7
        Me.grpComp.Text = "GroupControl2"
        '
        'chkAllComp
        '
        Me.chkAllComp.Location = New System.Drawing.Point(129, 6)
        Me.chkAllComp.Name = "chkAllComp"
        Me.chkAllComp.Properties.Caption = "所有公司"
        Me.chkAllComp.Properties.CheckStyle = DevExpress.XtraEditors.Controls.CheckStyles.Style1
        Me.chkAllComp.Properties.RadioGroupIndex = 1
        Me.chkAllComp.Size = New System.Drawing.Size(102, 22)
        Me.chkAllComp.TabIndex = 8
        Me.chkAllComp.TabStop = False
        '
        'chkSingleComp
        '
        Me.chkSingleComp.EditValue = True
        Me.chkSingleComp.Location = New System.Drawing.Point(3, 6)
        Me.chkSingleComp.Name = "chkSingleComp"
        Me.chkSingleComp.Properties.Caption = "個別公司"
        Me.chkSingleComp.Properties.CheckStyle = DevExpress.XtraEditors.Controls.CheckStyles.Style1
        Me.chkSingleComp.Properties.RadioGroupIndex = 1
        Me.chkSingleComp.Size = New System.Drawing.Size(96, 22)
        Me.chkSingleComp.TabIndex = 7
        '
        'LabelControl4
        '
        Me.LabelControl4.Location = New System.Drawing.Point(17, 143)
        Me.LabelControl4.Name = "LabelControl4"
        Me.LabelControl4.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl4.TabIndex = 6
        Me.LabelControl4.Text = "公司選項"
        '
        'LabelControl3
        '
        Me.LabelControl3.Location = New System.Drawing.Point(17, 103)
        Me.LabelControl3.Name = "LabelControl3"
        Me.LabelControl3.Size = New System.Drawing.Size(48, 14)
        Me.LabelControl3.TabIndex = 5
        Me.LabelControl3.Text = "報表格式"
        '
        'GroupControl1
        '
        Me.GroupControl1.Controls.Add(Me.chkNotifyCnt)
        Me.GroupControl1.Controls.Add(Me.chkNotifyDetail)
        Me.GroupControl1.Location = New System.Drawing.Point(70, 92)
        Me.GroupControl1.Name = "GroupControl1"
        Me.GroupControl1.ShowCaption = False
        Me.GroupControl1.Size = New System.Drawing.Size(246, 33)
        Me.GroupControl1.TabIndex = 4
        Me.GroupControl1.Text = "GroupControl1"
        '
        'chkNotifyCnt
        '
        Me.chkNotifyCnt.Location = New System.Drawing.Point(129, 6)
        Me.chkNotifyCnt.Name = "chkNotifyCnt"
        Me.chkNotifyCnt.Properties.Caption = "發票筆數統計"
        Me.chkNotifyCnt.Properties.CheckStyle = DevExpress.XtraEditors.Controls.CheckStyles.Style1
        Me.chkNotifyCnt.Properties.RadioGroupIndex = 1
        Me.chkNotifyCnt.Size = New System.Drawing.Size(102, 22)
        Me.chkNotifyCnt.TabIndex = 6
        Me.chkNotifyCnt.TabStop = False
        '
        'chkNotifyDetail
        '
        Me.chkNotifyDetail.EditValue = True
        Me.chkNotifyDetail.Location = New System.Drawing.Point(3, 6)
        Me.chkNotifyDetail.Name = "chkNotifyDetail"
        Me.chkNotifyDetail.Properties.Caption = "電子發票通知明細"
        Me.chkNotifyDetail.Properties.CheckStyle = DevExpress.XtraEditors.Controls.CheckStyles.Style1
        Me.chkNotifyDetail.Properties.RadioGroupIndex = 1
        Me.chkNotifyDetail.Size = New System.Drawing.Size(128, 22)
        Me.chkNotifyDetail.TabIndex = 5
        '
        'dtInv2
        '
        Me.dtInv2.EditValue = Nothing
        Me.dtInv2.Location = New System.Drawing.Point(216, 31)
        Me.dtInv2.Name = "dtInv2"
        Me.dtInv2.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.dtInv2.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTimeAdvancingCaret
        Me.dtInv2.Properties.VistaTimeProperties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton})
        Me.dtInv2.Size = New System.Drawing.Size(100, 21)
        Me.dtInv2.TabIndex = 2
        '
        'LabelControl2
        '
        Me.LabelControl2.Location = New System.Drawing.Point(201, 33)
        Me.LabelControl2.Name = "LabelControl2"
        Me.LabelControl2.Size = New System.Drawing.Size(9, 14)
        Me.LabelControl2.TabIndex = 2
        Me.LabelControl2.Text = "~"
        '
        'LabelControl1
        '
        Me.LabelControl1.Location = New System.Drawing.Point(15, 35)
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(72, 14)
        Me.LabelControl1.TabIndex = 1
        Me.LabelControl1.Text = "發票日期起訖"
        '
        'dtInv1
        '
        Me.dtInv1.EditValue = Nothing
        Me.dtInv1.Location = New System.Drawing.Point(95, 31)
        Me.dtInv1.Name = "dtInv1"
        Me.dtInv1.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.dtInv1.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTimeAdvancingCaret
        Me.dtInv1.Properties.VistaTimeProperties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton})
        Me.dtInv1.Size = New System.Drawing.Size(100, 21)
        Me.dtInv1.TabIndex = 1
        '
        'edtInvId1
        '
        Me.edtInvId1.Location = New System.Drawing.Point(95, 62)
        Me.edtInvId1.Name = "edtInvId1"
        Me.edtInvId1.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom
        Me.edtInvId1.Properties.Mask.EditMask = "\p{Lu}\p{Lu}\d{8}"
        Me.edtInvId1.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.RegEx
        Me.edtInvId1.Properties.Mask.ShowPlaceHolders = False
        Me.edtInvId1.Size = New System.Drawing.Size(100, 21)
        Me.edtInvId1.TabIndex = 3
        '
        'frmNotifyElectron
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(329, 222)
        Me.Controls.Add(Me.PanelControl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmNotifyElectron"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "C06-電子發票通知統計表"
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl2.ResumeLayout(False)
        CType(Me.grpOption, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpOption.ResumeLayout(False)
        Me.grpOption.PerformLayout()
        CType(Me.edtInvId2.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpComp, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpComp.ResumeLayout(False)
        CType(Me.chkAllComp.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkSingleComp.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GroupControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupControl1.ResumeLayout(False)
        CType(Me.chkNotifyCnt.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkNotifyDetail.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtInv2.Properties.VistaTimeProperties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtInv2.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtInv1.Properties.VistaTimeProperties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtInv1.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.edtInvId1.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents grpOption As DevExpress.XtraEditors.GroupControl
    Friend WithEvents dtInv1 As DevExpress.XtraEditors.DateEdit
    Friend WithEvents PanelControl2 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents btnOK As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl3 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents GroupControl1 As DevExpress.XtraEditors.GroupControl
    Friend WithEvents chkNotifyDetail As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents dtInv2 As DevExpress.XtraEditors.DateEdit
    Friend WithEvents LabelControl2 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents chkNotifyCnt As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents btnDesign As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LabelControl4 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents grpComp As DevExpress.XtraEditors.GroupControl
    Friend WithEvents chkAllComp As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents chkSingleComp As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents btnExit As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents DefaultLookAndFeel1 As DevExpress.LookAndFeel.DefaultLookAndFeel
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents LabelControl5 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents edtInvId1 As DevExpress.XtraEditors.TextEdit
    Friend WithEvents edtInvId2 As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl6 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents btnShowSQL As DevExpress.XtraEditors.SimpleButton

End Class
