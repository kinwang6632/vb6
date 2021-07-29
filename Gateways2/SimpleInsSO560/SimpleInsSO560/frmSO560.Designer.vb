<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtSid = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtUserId = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.numCompCode = New System.Windows.Forms.NumericUpDown
        Me.Label13 = New System.Windows.Forms.Label
        Me.dtexpiration_date = New System.Windows.Forms.DateTimePicker
        Me.Label12 = New System.Windows.Forms.Label
        Me.dtOrderDate = New System.Windows.Forms.DateTimePicker
        Me.Label11 = New System.Windows.Forms.Label
        Me.dtDMM_creation_date = New System.Windows.Forms.DateTimePicker
        Me.Label10 = New System.Windows.Forms.Label
        Me.numViewing_duration = New System.Windows.Forms.NumericUpDown
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtAsset = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtProductId = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtSTBNo = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.numDMM_creation_fate_flag = New System.Windows.Forms.NumericUpDown
        Me.Label5 = New System.Windows.Forms.Label
        Me.numValidity_Flag = New System.Windows.Forms.NumericUpDown
        Me.Label4 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnIns = New System.Windows.Forms.Button
        Me.btnBatchSingle = New System.Windows.Forms.Button
        Me.numBatch = New System.Windows.Forms.NumericUpDown
        Me.btnInsBatch = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.numCompCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numViewing_duration, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numDMM_creation_fate_flag, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numValidity_Flag, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        CType(Me.numBatch, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtSid)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtPassword)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtUserId)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(684, 56)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "連線資訊"
        '
        'txtSid
        '
        Me.txtSid.Location = New System.Drawing.Point(384, 21)
        Me.txtSid.Name = "txtSid"
        Me.txtSid.Size = New System.Drawing.Size(100, 22)
        Me.txtSid.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(352, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(26, 16)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Sid"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(239, 21)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(100, 22)
        Me.txtPassword.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(173, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Password"
        '
        'txtUserId
        '
        Me.txtUserId.Location = New System.Drawing.Point(63, 21)
        Me.txtUserId.Name = "txtUserId"
        Me.txtUserId.Size = New System.Drawing.Size(100, 22)
        Me.txtUserId.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "UserId"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.numCompCode)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.dtexpiration_date)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.dtOrderDate)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.dtDMM_creation_date)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.numViewing_duration)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.txtAsset)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.txtProductId)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.txtSTBNo)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.numDMM_creation_fate_flag)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.numValidity_Flag)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 56)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(684, 150)
        Me.Panel1.TabIndex = 1
        '
        'numCompCode
        '
        Me.numCompCode.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.numCompCode.Location = New System.Drawing.Point(89, 119)
        Me.numCompCode.Maximum = New Decimal(New Integer() {99, 0, 0, 0})
        Me.numCompCode.Name = "numCompCode"
        Me.numCompCode.Size = New System.Drawing.Size(52, 23)
        Me.numCompCode.TabIndex = 20
        Me.numCompCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(13, 121)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(70, 16)
        Me.Label13.TabIndex = 19
        Me.Label13.Text = "CompCode"
        '
        'dtexpiration_date
        '
        Me.dtexpiration_date.CustomFormat = "yyyy/mm/dd hh:mm:ss"
        Me.dtexpiration_date.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtexpiration_date.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtexpiration_date.Location = New System.Drawing.Point(528, 87)
        Me.dtexpiration_date.Name = "dtexpiration_date"
        Me.dtexpiration_date.Size = New System.Drawing.Size(148, 23)
        Me.dtexpiration_date.TabIndex = 18
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(426, 90)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(96, 16)
        Me.Label12.TabIndex = 17
        Me.Label12.Text = "expiration_date"
        '
        'dtOrderDate
        '
        Me.dtOrderDate.CustomFormat = "yyyy/mm/dd hh:mm:ss"
        Me.dtOrderDate.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtOrderDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtOrderDate.Location = New System.Drawing.Point(270, 86)
        Me.dtOrderDate.Name = "dtOrderDate"
        Me.dtOrderDate.Size = New System.Drawing.Size(150, 23)
        Me.dtOrderDate.TabIndex = 16
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(196, 91)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(68, 16)
        Me.Label11.TabIndex = 15
        Me.Label11.Text = "訂購時間"
        '
        'dtDMM_creation_date
        '
        Me.dtDMM_creation_date.CustomFormat = "yyyy/mm/dd hh:mm:ss"
        Me.dtDMM_creation_date.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtDMM_creation_date.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtDMM_creation_date.Location = New System.Drawing.Point(477, 13)
        Me.dtDMM_creation_date.Name = "dtDMM_creation_date"
        Me.dtDMM_creation_date.Size = New System.Drawing.Size(148, 23)
        Me.dtDMM_creation_date.TabIndex = 14
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(352, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(121, 16)
        Me.Label10.TabIndex = 13
        Me.Label10.Text = "DMM_creation_date"
        '
        'numViewing_duration
        '
        Me.numViewing_duration.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.numViewing_duration.Location = New System.Drawing.Point(113, 87)
        Me.numViewing_duration.Maximum = New Decimal(New Integer() {86400, 0, 0, 0})
        Me.numViewing_duration.Name = "numViewing_duration"
        Me.numViewing_duration.Size = New System.Drawing.Size(73, 23)
        Me.numViewing_duration.TabIndex = 12
        Me.numViewing_duration.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(12, 91)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(98, 16)
        Me.Label9.TabIndex = 11
        Me.Label9.Text = "授權觀賞期間"
        '
        'txtAsset
        '
        Me.txtAsset.Location = New System.Drawing.Point(436, 52)
        Me.txtAsset.Name = "txtAsset"
        Me.txtAsset.Size = New System.Drawing.Size(103, 22)
        Me.txtAsset.TabIndex = 10
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(391, 54)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(39, 16)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "Asset"
        '
        'txtProductId
        '
        Me.txtProductId.Location = New System.Drawing.Point(270, 52)
        Me.txtProductId.Name = "txtProductId"
        Me.txtProductId.Size = New System.Drawing.Size(108, 22)
        Me.txtProductId.TabIndex = 8
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(202, 54)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(62, 16)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "ProductId"
        '
        'txtSTBNo
        '
        Me.txtSTBNo.Location = New System.Drawing.Point(60, 52)
        Me.txtSTBNo.MaxLength = 10
        Me.txtSTBNo.Name = "txtSTBNo"
        Me.txtSTBNo.Size = New System.Drawing.Size(126, 22)
        Me.txtSTBNo.TabIndex = 6
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(12, 54)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(46, 16)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "STBNo"
        '
        'numDMM_creation_fate_flag
        '
        Me.numDMM_creation_fate_flag.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.numDMM_creation_fate_flag.Location = New System.Drawing.Point(294, 12)
        Me.numDMM_creation_fate_flag.Maximum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numDMM_creation_fate_flag.Name = "numDMM_creation_fate_flag"
        Me.numDMM_creation_fate_flag.Size = New System.Drawing.Size(52, 23)
        Me.numDMM_creation_fate_flag.TabIndex = 4
        Me.numDMM_creation_fate_flag.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(148, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(146, 16)
        Me.Label5.TabIndex = 3
        Me.Label5.Text = "DMM_creation_fate_flag"
        '
        'numValidity_Flag
        '
        Me.numValidity_Flag.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.numValidity_Flag.Location = New System.Drawing.Point(90, 12)
        Me.numValidity_Flag.Maximum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numValidity_Flag.Name = "numValidity_Flag"
        Me.numValidity_Flag.Size = New System.Drawing.Size(52, 23)
        Me.numValidity_Flag.TabIndex = 2
        Me.numValidity_Flag.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(12, 14)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 16)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Validity_Flag"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnOK)
        Me.Panel2.Controls.Add(Me.btnIns)
        Me.Panel2.Controls.Add(Me.btnBatchSingle)
        Me.Panel2.Controls.Add(Me.numBatch)
        Me.Panel2.Controls.Add(Me.btnInsBatch)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 206)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(684, 61)
        Me.Panel2.TabIndex = 2
        '
        'btnOK
        '
        Me.btnOK.Enabled = False
        Me.btnOK.Location = New System.Drawing.Point(373, 23)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(67, 23)
        Me.btnOK.TabIndex = 8
        Me.btnOK.Text = "完成"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnIns
        '
        Me.btnIns.Enabled = False
        Me.btnIns.Location = New System.Drawing.Point(300, 23)
        Me.btnIns.Name = "btnIns"
        Me.btnIns.Size = New System.Drawing.Size(67, 23)
        Me.btnIns.TabIndex = 7
        Me.btnIns.Text = "新增"
        Me.btnIns.UseVisualStyleBackColor = True
        '
        'btnBatchSingle
        '
        Me.btnBatchSingle.Location = New System.Drawing.Point(189, 23)
        Me.btnBatchSingle.Name = "btnBatchSingle"
        Me.btnBatchSingle.Size = New System.Drawing.Size(105, 23)
        Me.btnBatchSingle.TabIndex = 6
        Me.btnBatchSingle.Text = "新增一筆命令"
        Me.btnBatchSingle.UseVisualStyleBackColor = True
        '
        'numBatch
        '
        Me.numBatch.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.numBatch.Location = New System.Drawing.Point(98, 23)
        Me.numBatch.Maximum = New Decimal(New Integer() {32767, 0, 0, 0})
        Me.numBatch.Name = "numBatch"
        Me.numBatch.Size = New System.Drawing.Size(52, 23)
        Me.numBatch.TabIndex = 5
        Me.numBatch.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.numBatch.Value = New Decimal(New Integer() {1000, 0, 0, 0})
        '
        'btnInsBatch
        '
        Me.btnInsBatch.Location = New System.Drawing.Point(17, 23)
        Me.btnInsBatch.Name = "btnInsBatch"
        Me.btnInsBatch.Size = New System.Drawing.Size(75, 23)
        Me.btnInsBatch.TabIndex = 0
        Me.btnInsBatch.Text = "大量新增"
        Me.btnInsBatch.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(684, 267)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SO560"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.numCompCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numViewing_duration, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numDMM_creation_fate_flag, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numValidity_Flag, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        CType(Me.numBatch, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtSid As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtUserId As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents numValidity_Flag As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtSTBNo As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents numDMM_creation_fate_flag As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtAsset As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtProductId As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents numViewing_duration As System.Windows.Forms.NumericUpDown
    Friend WithEvents dtDMM_creation_date As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtOrderDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents numCompCode As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents dtexpiration_date As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnIns As System.Windows.Forms.Button
    Friend WithEvents btnBatchSingle As System.Windows.Forms.Button
    Friend WithEvents numBatch As System.Windows.Forms.NumericUpDown
    Friend WithEvents btnInsBatch As System.Windows.Forms.Button

End Class
