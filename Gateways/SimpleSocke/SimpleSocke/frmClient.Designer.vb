<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmClient
    Inherits System.Windows.Forms.Form

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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmClient))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.txtUrlDesc = New System.Windows.Forms.TextBox
        Me.Label31 = New System.Windows.Forms.Label
        Me.txtUrlCode = New System.Windows.Forms.TextBox
        Me.Label32 = New System.Windows.Forms.Label
        Me.btnGetWebStatus = New System.Windows.Forms.Button
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.btnDisconnect = New System.Windows.Forms.Button
        Me.Label23 = New System.Windows.Forms.Label
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.txtReturnLines = New System.Windows.Forms.TextBox
        Me.txtErrName = New System.Windows.Forms.TextBox
        Me.txtErrCode = New System.Windows.Forms.TextBox
        Me.txtStatus = New System.Windows.Forms.TextBox
        Me.回應狀態 = New System.Windows.Forms.Label
        Me.txtSendMsg = New System.Windows.Forms.TextBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.btnCMD0851 = New System.Windows.Forms.Button
        Me.btnCMD1002 = New System.Windows.Forms.Button
        Me.btnConnect = New System.Windows.Forms.Button
        Me.txtResult = New System.Windows.Forms.TextBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnSetOK = New System.Windows.Forms.Button
        Me.txtcontent_id = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.txtProductId = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.dtpexpiration_date = New System.Windows.Forms.DateTimePicker
        Me.Label24 = New System.Windows.Forms.Label
        Me.txtMOP_PPID = New System.Windows.Forms.TextBox
        Me.txtdest_id = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.dest_id = New System.Windows.Forms.Label
        Me.txtSource_id = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.txtviewing_duration = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.dtpOrderDate = New System.Windows.Forms.DateTimePicker
        Me.Label15 = New System.Windows.Forms.Label
        Me.txtSTBNo = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.dtpDMM_creation_date = New System.Windows.Forms.DateTimePicker
        Me.Label13 = New System.Windows.Forms.Label
        Me.txtDMM_creation_fate_flag = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtadditional_address_flag = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtchipset_type_flag = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtnb_of_R_VOD_ENT = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtValidity_Flag = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtEncrypted_asset_flag = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtLowCmd = New System.Windows.Forms.TextBox
        Me.numRcvTimeOut = New System.Windows.Forms.NumericUpDown
        Me.Label7 = New System.Windows.Forms.Label
        Me.numSndTimeOut = New System.Windows.Forms.NumericUpDown
        Me.Label8 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.numPort = New System.Windows.Forms.NumericUpDown
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtIPAddress = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btnSaveWebSet = New System.Windows.Forms.Button
        Me.numUrlRepeat = New System.Windows.Forms.NumericUpDown
        Me.Label29 = New System.Windows.Forms.Label
        Me.numUrlDuration = New System.Windows.Forms.NumericUpDown
        Me.Label30 = New System.Windows.Forms.Label
        Me.numUrlTimeout = New System.Windows.Forms.NumericUpDown
        Me.Label27 = New System.Windows.Forms.Label
        Me.Label28 = New System.Windows.Forms.Label
        Me.cboUrlMetho = New System.Windows.Forms.ComboBox
        Me.txtUrl = New System.Windows.Forms.TextBox
        Me.lblUrl = New System.Windows.Forms.Label
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.Panel1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.numRcvTimeOut, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numSndTimeOut, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numPort, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.numUrlRepeat, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numUrlDuration, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numUrlTimeout, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.TabControl1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(623, 391)
        Me.Panel1.TabIndex = 0
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(623, 391)
        Me.TabControl1.TabIndex = 10
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.txtUrlDesc)
        Me.TabPage1.Controls.Add(Me.Label31)
        Me.TabPage1.Controls.Add(Me.txtUrlCode)
        Me.TabPage1.Controls.Add(Me.Label32)
        Me.TabPage1.Controls.Add(Me.btnGetWebStatus)
        Me.TabPage1.Controls.Add(Me.btnDisconnect)
        Me.TabPage1.Controls.Add(Me.Label23)
        Me.TabPage1.Controls.Add(Me.Label22)
        Me.TabPage1.Controls.Add(Me.Label21)
        Me.TabPage1.Controls.Add(Me.txtReturnLines)
        Me.TabPage1.Controls.Add(Me.txtErrName)
        Me.TabPage1.Controls.Add(Me.txtErrCode)
        Me.TabPage1.Controls.Add(Me.txtStatus)
        Me.TabPage1.Controls.Add(Me.回應狀態)
        Me.TabPage1.Controls.Add(Me.txtSendMsg)
        Me.TabPage1.Controls.Add(Me.Label20)
        Me.TabPage1.Controls.Add(Me.Label19)
        Me.TabPage1.Controls.Add(Me.btnCMD0851)
        Me.TabPage1.Controls.Add(Me.btnCMD1002)
        Me.TabPage1.Controls.Add(Me.btnConnect)
        Me.TabPage1.Controls.Add(Me.txtResult)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(615, 362)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "測試傳送"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'txtUrlDesc
        '
        Me.txtUrlDesc.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUrlDesc.Location = New System.Drawing.Point(227, 260)
        Me.txtUrlDesc.Name = "txtUrlDesc"
        Me.txtUrlDesc.ReadOnly = True
        Me.txtUrlDesc.Size = New System.Drawing.Size(384, 23)
        Me.txtUrlDesc.TabIndex = 26
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label31.Location = New System.Drawing.Point(153, 263)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(76, 16)
        Me.Label31.TabIndex = 28
        Me.Label31.Text = "Description:"
        '
        'txtUrlCode
        '
        Me.txtUrlCode.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUrlCode.Location = New System.Drawing.Point(77, 260)
        Me.txtUrlCode.Name = "txtUrlCode"
        Me.txtUrlCode.ReadOnly = True
        Me.txtUrlCode.Size = New System.Drawing.Size(50, 23)
        Me.txtUrlCode.TabIndex = 25
        Me.txtUrlCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label32.Location = New System.Drawing.Point(29, 263)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(42, 16)
        Me.Label32.TabIndex = 27
        Me.Label32.Text = "Code:"
        '
        'btnGetWebStatus
        '
        Me.btnGetWebStatus.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnGetWebStatus.ImageIndex = 5
        Me.btnGetWebStatus.ImageList = Me.ImageList1
        Me.btnGetWebStatus.Location = New System.Drawing.Point(394, 311)
        Me.btnGetWebStatus.Name = "btnGetWebStatus"
        Me.btnGetWebStatus.Size = New System.Drawing.Size(134, 24)
        Me.btnGetWebStatus.TabIndex = 24
        Me.btnGetWebStatus.Text = "取得授權狀態"
        Me.btnGetWebStatus.UseVisualStyleBackColor = True
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "btnConnect.png")
        Me.ImageList1.Images.SetKeyName(1, "disconnect.png")
        Me.ImageList1.Images.SetKeyName(2, "cmd0851.png")
        Me.ImageList1.Images.SetKeyName(3, "cmd1002.png")
        Me.ImageList1.Images.SetKeyName(4, "OK.png")
        Me.ImageList1.Images.SetKeyName(5, "Appicon.ico")
        '
        'btnDisconnect
        '
        Me.btnDisconnect.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDisconnect.ImageIndex = 1
        Me.btnDisconnect.ImageList = Me.ImageList1
        Me.btnDisconnect.Location = New System.Drawing.Point(530, 311)
        Me.btnDisconnect.Name = "btnDisconnect"
        Me.btnDisconnect.Size = New System.Drawing.Size(77, 24)
        Me.btnDisconnect.TabIndex = 23
        Me.btnDisconnect.Text = "斷線"
        Me.btnDisconnect.UseVisualStyleBackColor = True
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(438, 229)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(68, 16)
        Me.Label23.TabIndex = 22
        Me.Label23.Text = "回應授權"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(273, 229)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(68, 16)
        Me.Label22.TabIndex = 21
        Me.Label22.Text = "錯誤名稱"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(133, 229)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(68, 16)
        Me.Label21.TabIndex = 20
        Me.Label21.Text = "錯誤代碼"
        '
        'txtReturnLines
        '
        Me.txtReturnLines.Location = New System.Drawing.Point(507, 226)
        Me.txtReturnLines.Name = "txtReturnLines"
        Me.txtReturnLines.ReadOnly = True
        Me.txtReturnLines.Size = New System.Drawing.Size(104, 23)
        Me.txtReturnLines.TabIndex = 19
        '
        'txtErrName
        '
        Me.txtErrName.Location = New System.Drawing.Point(347, 226)
        Me.txtErrName.Name = "txtErrName"
        Me.txtErrName.ReadOnly = True
        Me.txtErrName.Size = New System.Drawing.Size(85, 23)
        Me.txtErrName.TabIndex = 18
        '
        'txtErrCode
        '
        Me.txtErrCode.Location = New System.Drawing.Point(204, 226)
        Me.txtErrCode.Name = "txtErrCode"
        Me.txtErrCode.ReadOnly = True
        Me.txtErrCode.Size = New System.Drawing.Size(63, 23)
        Me.txtErrCode.TabIndex = 17
        '
        'txtStatus
        '
        Me.txtStatus.Location = New System.Drawing.Point(80, 226)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.ReadOnly = True
        Me.txtStatus.Size = New System.Drawing.Size(47, 23)
        Me.txtStatus.TabIndex = 16
        '
        '回應狀態
        '
        Me.回應狀態.AutoSize = True
        Me.回應狀態.Location = New System.Drawing.Point(12, 229)
        Me.回應狀態.Name = "回應狀態"
        Me.回應狀態.Size = New System.Drawing.Size(68, 16)
        Me.回應狀態.TabIndex = 15
        Me.回應狀態.Text = "回應狀態"
        '
        'txtSendMsg
        '
        Me.txtSendMsg.Location = New System.Drawing.Point(76, 16)
        Me.txtSendMsg.Multiline = True
        Me.txtSendMsg.Name = "txtSendMsg"
        Me.txtSendMsg.ReadOnly = True
        Me.txtSendMsg.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSendMsg.Size = New System.Drawing.Size(531, 66)
        Me.txtSendMsg.TabIndex = 14
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(8, 42)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(68, 16)
        Me.Label20.TabIndex = 13
        Me.Label20.Text = "傳送訊息"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(8, 145)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(68, 16)
        Me.Label19.TabIndex = 12
        Me.Label19.Text = "接收訊息"
        '
        'btnCMD0851
        '
        Me.btnCMD0851.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCMD0851.ImageIndex = 2
        Me.btnCMD0851.ImageList = Me.ImageList1
        Me.btnCMD0851.Location = New System.Drawing.Point(259, 311)
        Me.btnCMD0851.Name = "btnCMD0851"
        Me.btnCMD0851.Size = New System.Drawing.Size(129, 24)
        Me.btnCMD0851.TabIndex = 11
        Me.btnCMD0851.Text = "測試CMD0851"
        Me.btnCMD0851.UseVisualStyleBackColor = True
        '
        'btnCMD1002
        '
        Me.btnCMD1002.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCMD1002.ImageIndex = 3
        Me.btnCMD1002.ImageList = Me.ImageList1
        Me.btnCMD1002.Location = New System.Drawing.Point(117, 311)
        Me.btnCMD1002.Name = "btnCMD1002"
        Me.btnCMD1002.Size = New System.Drawing.Size(136, 24)
        Me.btnCMD1002.TabIndex = 10
        Me.btnCMD1002.Text = "測試CMD1002"
        Me.btnCMD1002.UseVisualStyleBackColor = True
        '
        'btnConnect
        '
        Me.btnConnect.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnConnect.ImageIndex = 0
        Me.btnConnect.ImageList = Me.ImageList1
        Me.btnConnect.Location = New System.Drawing.Point(8, 311)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.Size = New System.Drawing.Size(103, 24)
        Me.btnConnect.TabIndex = 9
        Me.btnConnect.Text = "測試連線"
        Me.btnConnect.UseVisualStyleBackColor = True
        '
        'txtResult
        '
        Me.txtResult.Location = New System.Drawing.Point(76, 99)
        Me.txtResult.Multiline = True
        Me.txtResult.Name = "txtResult"
        Me.txtResult.ReadOnly = True
        Me.txtResult.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtResult.Size = New System.Drawing.Size(531, 102)
        Me.txtResult.TabIndex = 8
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.GroupBox1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(615, 362)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Nagra參數設定"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnSetOK)
        Me.GroupBox1.Controls.Add(Me.txtcontent_id)
        Me.GroupBox1.Controls.Add(Me.Label26)
        Me.GroupBox1.Controls.Add(Me.txtProductId)
        Me.GroupBox1.Controls.Add(Me.Label25)
        Me.GroupBox1.Controls.Add(Me.dtpexpiration_date)
        Me.GroupBox1.Controls.Add(Me.Label24)
        Me.GroupBox1.Controls.Add(Me.txtMOP_PPID)
        Me.GroupBox1.Controls.Add(Me.txtdest_id)
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.Controls.Add(Me.dest_id)
        Me.GroupBox1.Controls.Add(Me.txtSource_id)
        Me.GroupBox1.Controls.Add(Me.Label17)
        Me.GroupBox1.Controls.Add(Me.txtviewing_duration)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.dtpOrderDate)
        Me.GroupBox1.Controls.Add(Me.Label15)
        Me.GroupBox1.Controls.Add(Me.txtSTBNo)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.dtpDMM_creation_date)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.txtDMM_creation_fate_flag)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.txtadditional_address_flag)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.txtchipset_type_flag)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtnb_of_R_VOD_ENT)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtValidity_Flag)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtEncrypted_asset_flag)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtLowCmd)
        Me.GroupBox1.Controls.Add(Me.numRcvTimeOut)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.numSndTimeOut)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.numPort)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.txtIPAddress)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(609, 352)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "傳送參數"
        '
        'btnSetOK
        '
        Me.btnSetOK.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSetOK.ImageIndex = 4
        Me.btnSetOK.ImageList = Me.ImageList1
        Me.btnSetOK.Location = New System.Drawing.Point(246, 318)
        Me.btnSetOK.Name = "btnSetOK"
        Me.btnSetOK.Size = New System.Drawing.Size(85, 23)
        Me.btnSetOK.TabIndex = 23
        Me.btnSetOK.Text = "確定"
        Me.btnSetOK.UseVisualStyleBackColor = True
        '
        'txtcontent_id
        '
        Me.txtcontent_id.Location = New System.Drawing.Point(286, 181)
        Me.txtcontent_id.MaxLength = 10
        Me.txtcontent_id.Name = "txtcontent_id"
        Me.txtcontent_id.Size = New System.Drawing.Size(114, 23)
        Me.txtcontent_id.TabIndex = 14
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(213, 185)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(67, 16)
        Me.Label26.TabIndex = 40
        Me.Label26.Text = "content_id"
        '
        'txtProductId
        '
        Me.txtProductId.Location = New System.Drawing.Point(82, 182)
        Me.txtProductId.Name = "txtProductId"
        Me.txtProductId.Size = New System.Drawing.Size(114, 23)
        Me.txtProductId.TabIndex = 13
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(13, 185)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(62, 16)
        Me.Label25.TabIndex = 38
        Me.Label25.Text = "ProductId"
        '
        'dtpexpiration_date
        '
        Me.dtpexpiration_date.CustomFormat = "yyyy/MM/dd hh:mm:ss"
        Me.dtpexpiration_date.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpexpiration_date.Location = New System.Drawing.Point(356, 103)
        Me.dtpexpiration_date.Name = "dtpexpiration_date"
        Me.dtpexpiration_date.Size = New System.Drawing.Size(150, 23)
        Me.dtpexpiration_date.TabIndex = 10
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(252, 107)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(96, 16)
        Me.Label24.TabIndex = 36
        Me.Label24.Text = "expiration_date"
        '
        'txtMOP_PPID
        '
        Me.txtMOP_PPID.Location = New System.Drawing.Point(337, 61)
        Me.txtMOP_PPID.MaxLength = 5
        Me.txtMOP_PPID.Name = "txtMOP_PPID"
        Me.txtMOP_PPID.Size = New System.Drawing.Size(55, 23)
        Me.txtMOP_PPID.TabIndex = 7
        '
        'txtdest_id
        '
        Me.txtdest_id.Location = New System.Drawing.Point(202, 61)
        Me.txtdest_id.MaxLength = 4
        Me.txtdest_id.Name = "txtdest_id"
        Me.txtdest_id.Size = New System.Drawing.Size(55, 23)
        Me.txtdest_id.TabIndex = 6
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(266, 64)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(67, 16)
        Me.Label18.TabIndex = 35
        Me.Label18.Text = "MOP_PPID"
        '
        'dest_id
        '
        Me.dest_id.AutoSize = True
        Me.dest_id.Location = New System.Drawing.Point(147, 64)
        Me.dest_id.Name = "dest_id"
        Me.dest_id.Size = New System.Drawing.Size(49, 16)
        Me.dest_id.TabIndex = 34
        Me.dest_id.Text = "dest_id"
        '
        'txtSource_id
        '
        Me.txtSource_id.Location = New System.Drawing.Point(83, 61)
        Me.txtSource_id.MaxLength = 4
        Me.txtSource_id.Name = "txtSource_id"
        Me.txtSource_id.Size = New System.Drawing.Size(55, 23)
        Me.txtSource_id.TabIndex = 5
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(14, 64)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(63, 16)
        Me.Label17.TabIndex = 32
        Me.Label17.Text = "source_id"
        '
        'txtviewing_duration
        '
        Me.txtviewing_duration.Location = New System.Drawing.Point(524, 181)
        Me.txtviewing_duration.Name = "txtviewing_duration"
        Me.txtviewing_duration.Size = New System.Drawing.Size(60, 23)
        Me.txtviewing_duration.TabIndex = 15
        Me.txtviewing_duration.Text = "172800"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(420, 185)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(98, 16)
        Me.Label16.TabIndex = 29
        Me.Label16.Text = "授權觀賞期間"
        '
        'dtpOrderDate
        '
        Me.dtpOrderDate.CustomFormat = "yyyy/MM/dd hh:mm:ss"
        Me.dtpOrderDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpOrderDate.Location = New System.Drawing.Point(83, 103)
        Me.dtpOrderDate.Name = "dtpOrderDate"
        Me.dtpOrderDate.Size = New System.Drawing.Size(150, 23)
        Me.dtpOrderDate.TabIndex = 9
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(15, 107)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(68, 16)
        Me.Label15.TabIndex = 27
        Me.Label15.Text = "訂購日期"
        '
        'txtSTBNo
        '
        Me.txtSTBNo.Location = New System.Drawing.Point(470, 61)
        Me.txtSTBNo.MaxLength = 10
        Me.txtSTBNo.Name = "txtSTBNo"
        Me.txtSTBNo.Size = New System.Drawing.Size(114, 23)
        Me.txtSTBNo.TabIndex = 8
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(407, 64)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(61, 16)
        Me.Label14.TabIndex = 25
        Me.Label14.Text = "STB序號"
        '
        'dtpDMM_creation_date
        '
        Me.dtpDMM_creation_date.CustomFormat = "yyyy/MM/dd hh:mm:ss"
        Me.dtpDMM_creation_date.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDMM_creation_date.Location = New System.Drawing.Point(340, 143)
        Me.dtpDMM_creation_date.Name = "dtpDMM_creation_date"
        Me.dtpDMM_creation_date.Size = New System.Drawing.Size(165, 23)
        Me.dtpDMM_creation_date.TabIndex = 12
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(213, 146)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(121, 16)
        Me.Label13.TabIndex = 23
        Me.Label13.Text = "DMM_creation_date"
        '
        'txtDMM_creation_fate_flag
        '
        Me.txtDMM_creation_fate_flag.Location = New System.Drawing.Point(165, 143)
        Me.txtDMM_creation_fate_flag.MaxLength = 1
        Me.txtDMM_creation_fate_flag.Name = "txtDMM_creation_fate_flag"
        Me.txtDMM_creation_fate_flag.Size = New System.Drawing.Size(31, 23)
        Me.txtDMM_creation_fate_flag.TabIndex = 11
        Me.txtDMM_creation_fate_flag.Text = "0"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(13, 146)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(146, 16)
        Me.Label12.TabIndex = 21
        Me.Label12.Text = "DMM_creation_fate_flag"
        '
        'txtadditional_address_flag
        '
        Me.txtadditional_address_flag.Location = New System.Drawing.Point(497, 249)
        Me.txtadditional_address_flag.Name = "txtadditional_address_flag"
        Me.txtadditional_address_flag.ReadOnly = True
        Me.txtadditional_address_flag.Size = New System.Drawing.Size(31, 23)
        Me.txtadditional_address_flag.TabIndex = 19
        Me.txtadditional_address_flag.Text = "0"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(348, 252)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(143, 16)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "additional_address_flag"
        '
        'txtchipset_type_flag
        '
        Me.txtchipset_type_flag.Location = New System.Drawing.Point(300, 249)
        Me.txtchipset_type_flag.Name = "txtchipset_type_flag"
        Me.txtchipset_type_flag.ReadOnly = True
        Me.txtchipset_type_flag.Size = New System.Drawing.Size(31, 23)
        Me.txtchipset_type_flag.TabIndex = 18
        Me.txtchipset_type_flag.Text = "0"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(188, 252)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(107, 16)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "chipset_type_flag"
        '
        'txtnb_of_R_VOD_ENT
        '
        Me.txtnb_of_R_VOD_ENT.Location = New System.Drawing.Point(140, 249)
        Me.txtnb_of_R_VOD_ENT.Name = "txtnb_of_R_VOD_ENT"
        Me.txtnb_of_R_VOD_ENT.ReadOnly = True
        Me.txtnb_of_R_VOD_ENT.Size = New System.Drawing.Size(31, 23)
        Me.txtnb_of_R_VOD_ENT.TabIndex = 17
        Me.txtnb_of_R_VOD_ENT.Text = "0"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 252)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(117, 16)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "nb_of_R_VOD_ENT"
        '
        'txtValidity_Flag
        '
        Me.txtValidity_Flag.Location = New System.Drawing.Point(140, 217)
        Me.txtValidity_Flag.MaxLength = 1
        Me.txtValidity_Flag.Name = "txtValidity_Flag"
        Me.txtValidity_Flag.Size = New System.Drawing.Size(31, 23)
        Me.txtValidity_Flag.TabIndex = 16
        Me.txtValidity_Flag.Text = "0"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 220)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(316, 16)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "觀賞期間起迄方式            (absolute=0,relevant=1)"
        '
        'txtEncrypted_asset_flag
        '
        Me.txtEncrypted_asset_flag.Location = New System.Drawing.Point(215, 284)
        Me.txtEncrypted_asset_flag.Name = "txtEncrypted_asset_flag"
        Me.txtEncrypted_asset_flag.ReadOnly = True
        Me.txtEncrypted_asset_flag.Size = New System.Drawing.Size(31, 23)
        Me.txtEncrypted_asset_flag.TabIndex = 21
        Me.txtEncrypted_asset_flag.Text = "1"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(141, 287)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(68, 16)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "使用加密"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 287)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 16)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "低階命令"
        '
        'txtLowCmd
        '
        Me.txtLowCmd.Location = New System.Drawing.Point(81, 284)
        Me.txtLowCmd.Name = "txtLowCmd"
        Me.txtLowCmd.ReadOnly = True
        Me.txtLowCmd.Size = New System.Drawing.Size(54, 23)
        Me.txtLowCmd.TabIndex = 20
        Me.txtLowCmd.Text = "0851"
        '
        'numRcvTimeOut
        '
        Me.numRcvTimeOut.Location = New System.Drawing.Point(546, 26)
        Me.numRcvTimeOut.Maximum = New Decimal(New Integer() {99999, 0, 0, 0})
        Me.numRcvTimeOut.Name = "numRcvTimeOut"
        Me.numRcvTimeOut.Size = New System.Drawing.Size(51, 23)
        Me.numRcvTimeOut.TabIndex = 4
        Me.numRcvTimeOut.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.numRcvTimeOut.Value = New Decimal(New Integer() {30, 0, 0, 0})
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(454, 29)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(93, 16)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "接收逾時(秒)"
        '
        'numSndTimeOut
        '
        Me.numSndTimeOut.Location = New System.Drawing.Point(397, 26)
        Me.numSndTimeOut.Maximum = New Decimal(New Integer() {99999, 0, 0, 0})
        Me.numSndTimeOut.Name = "numSndTimeOut"
        Me.numSndTimeOut.Size = New System.Drawing.Size(51, 23)
        Me.numSndTimeOut.TabIndex = 3
        Me.numSndTimeOut.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.numSndTimeOut.Value = New Decimal(New Integer() {30, 0, 0, 0})
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(299, 29)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(93, 16)
        Me.Label8.TabIndex = 6
        Me.Label8.Text = "傳送逾時(秒)"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(93, 357)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(513, 23)
        Me.TextBox1.TabIndex = 6
        Me.TextBox1.Visible = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(22, 363)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(68, 16)
        Me.Label9.TabIndex = 4
        Me.Label9.Text = "傳送參數"
        Me.Label9.Visible = False
        '
        'numPort
        '
        Me.numPort.Location = New System.Drawing.Point(231, 26)
        Me.numPort.Maximum = New Decimal(New Integer() {99999, 0, 0, 0})
        Me.numPort.Name = "numPort"
        Me.numPort.Size = New System.Drawing.Size(57, 23)
        Me.numPort.TabIndex = 2
        Me.numPort.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(201, 29)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(31, 16)
        Me.Label10.TabIndex = 2
        Me.Label10.Text = "Port"
        '
        'txtIPAddress
        '
        Me.txtIPAddress.Location = New System.Drawing.Point(82, 26)
        Me.txtIPAddress.Name = "txtIPAddress"
        Me.txtIPAddress.Size = New System.Drawing.Size(114, 23)
        Me.txtIPAddress.TabIndex = 1
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(14, 29)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(65, 16)
        Me.Label11.TabIndex = 0
        Me.Label11.Text = "IPAddress"
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.GroupBox2)
        Me.TabPage3.Location = New System.Drawing.Point(4, 25)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(615, 362)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Push Server 參數設定"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnSaveWebSet)
        Me.GroupBox2.Controls.Add(Me.numUrlRepeat)
        Me.GroupBox2.Controls.Add(Me.Label29)
        Me.GroupBox2.Controls.Add(Me.numUrlDuration)
        Me.GroupBox2.Controls.Add(Me.Label30)
        Me.GroupBox2.Controls.Add(Me.numUrlTimeout)
        Me.GroupBox2.Controls.Add(Me.Label27)
        Me.GroupBox2.Controls.Add(Me.Label28)
        Me.GroupBox2.Controls.Add(Me.cboUrlMetho)
        Me.GroupBox2.Controls.Add(Me.txtUrl)
        Me.GroupBox2.Controls.Add(Me.lblUrl)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(615, 362)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Push Server參數"
        '
        'btnSaveWebSet
        '
        Me.btnSaveWebSet.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSaveWebSet.ImageIndex = 4
        Me.btnSaveWebSet.ImageList = Me.ImageList1
        Me.btnSaveWebSet.Location = New System.Drawing.Point(244, 320)
        Me.btnSaveWebSet.Name = "btnSaveWebSet"
        Me.btnSaveWebSet.Size = New System.Drawing.Size(85, 23)
        Me.btnSaveWebSet.TabIndex = 24
        Me.btnSaveWebSet.Text = "確定"
        Me.btnSaveWebSet.UseVisualStyleBackColor = True
        '
        'numUrlRepeat
        '
        Me.numUrlRepeat.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.numUrlRepeat.Location = New System.Drawing.Point(212, 80)
        Me.numUrlRepeat.Name = "numUrlRepeat"
        Me.numUrlRepeat.Size = New System.Drawing.Size(59, 23)
        Me.numUrlRepeat.TabIndex = 16
        Me.numUrlRepeat.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.Location = New System.Drawing.Point(153, 82)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(53, 16)
        Me.Label29.TabIndex = 18
        Me.Label29.Text = "Repeat:"
        '
        'numUrlDuration
        '
        Me.numUrlDuration.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.numUrlDuration.Location = New System.Drawing.Point(71, 80)
        Me.numUrlDuration.Name = "numUrlDuration"
        Me.numUrlDuration.Size = New System.Drawing.Size(59, 23)
        Me.numUrlDuration.TabIndex = 15
        Me.numUrlDuration.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label30.Location = New System.Drawing.Point(8, 82)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(61, 16)
        Me.Label30.TabIndex = 17
        Me.Label30.Text = "Duration:"
        '
        'numUrlTimeout
        '
        Me.numUrlTimeout.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.numUrlTimeout.Location = New System.Drawing.Point(370, 80)
        Me.numUrlTimeout.Name = "numUrlTimeout"
        Me.numUrlTimeout.Size = New System.Drawing.Size(46, 23)
        Me.numUrlTimeout.TabIndex = 11
        Me.numUrlTimeout.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.numUrlTimeout.Value = New Decimal(New Integer() {30, 0, 0, 0})
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.Location = New System.Drawing.Point(299, 82)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(62, 16)
        Me.Label27.TabIndex = 14
        Me.Label27.Text = "TimeOut:"
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.Location = New System.Drawing.Point(437, 82)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(73, 16)
        Me.Label28.TabIndex = 13
        Me.Label28.Text = "傳送方式:"
        '
        'cboUrlMetho
        '
        Me.cboUrlMetho.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboUrlMetho.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboUrlMetho.FormattingEnabled = True
        Me.cboUrlMetho.Items.AddRange(New Object() {"POST", "GET"})
        Me.cboUrlMetho.Location = New System.Drawing.Point(516, 79)
        Me.cboUrlMetho.Name = "cboUrlMetho"
        Me.cboUrlMetho.Size = New System.Drawing.Size(68, 24)
        Me.cboUrlMetho.TabIndex = 12
        '
        'txtUrl
        '
        Me.txtUrl.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUrl.Location = New System.Drawing.Point(40, 36)
        Me.txtUrl.Name = "txtUrl"
        Me.txtUrl.Size = New System.Drawing.Size(544, 23)
        Me.txtUrl.TabIndex = 10
        '
        'lblUrl
        '
        Me.lblUrl.AutoSize = True
        Me.lblUrl.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUrl.Location = New System.Drawing.Point(8, 39)
        Me.lblUrl.Name = "lblUrl"
        Me.lblUrl.Size = New System.Drawing.Size(35, 16)
        Me.lblUrl.TabIndex = 9
        Me.lblUrl.Text = "URL:"
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'frmClient
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(623, 391)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmClient"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "簡易Nagra Socket (0851)"
        Me.Panel1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.numRcvTimeOut, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numSndTimeOut, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numPort, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.numUrlRepeat, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numUrlDuration, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numUrlTimeout, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents btnConnect As System.Windows.Forms.Button
    Friend WithEvents txtResult As System.Windows.Forms.TextBox
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtLowCmd As System.Windows.Forms.TextBox
    Friend WithEvents numRcvTimeOut As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents numSndTimeOut As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents numPort As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtIPAddress As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtValidity_Flag As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtEncrypted_asset_flag As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtnb_of_R_VOD_ENT As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtadditional_address_flag As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtchipset_type_flag As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtDMM_creation_fate_flag As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents dtpDMM_creation_date As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtviewing_duration As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents dtpOrderDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtSTBNo As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents btnCMD0851 As System.Windows.Forms.Button
    Friend WithEvents btnCMD1002 As System.Windows.Forms.Button
    Friend WithEvents txtMOP_PPID As System.Windows.Forms.TextBox
    Friend WithEvents txtdest_id As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents dest_id As System.Windows.Forms.Label
    Friend WithEvents txtSource_id As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtSendMsg As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtStatus As System.Windows.Forms.TextBox
    Friend WithEvents 回應狀態 As System.Windows.Forms.Label
    Friend WithEvents txtErrName As System.Windows.Forms.TextBox
    Friend WithEvents txtErrCode As System.Windows.Forms.TextBox
    Friend WithEvents txtReturnLines As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents txtProductId As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents dtpexpiration_date As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtcontent_id As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents btnDisconnect As System.Windows.Forms.Button
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents numUrlRepeat As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents numUrlDuration As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents numUrlTimeout As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents cboUrlMetho As System.Windows.Forms.ComboBox
    Friend WithEvents txtUrl As System.Windows.Forms.TextBox
    Friend WithEvents lblUrl As System.Windows.Forms.Label
    Friend WithEvents btnGetWebStatus As System.Windows.Forms.Button
    Friend WithEvents txtUrlDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents txtUrlCode As System.Windows.Forms.TextBox
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents btnSetOK As System.Windows.Forms.Button
    Friend WithEvents btnSaveWebSet As System.Windows.Forms.Button

End Class
