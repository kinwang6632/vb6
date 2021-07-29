<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmKey
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmKey))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtRegSN = New System.Windows.Forms.TextBox()
        Me.dtpUseDate = New System.Windows.Forms.DateTimePicker()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.dtpInsDate = New System.Windows.Forms.DateTimePicker()
        Me.lblInsDate = New System.Windows.Forms.Label()
        Me.txtApp = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtMac = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtMB = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtHD = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtCpu = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtSoftSN = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCopy = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(521, 65)
        Me.Panel1.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(173, 23)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(203, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "輸入軟體序號可產生註冊號碼"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(105, 10)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(48, 48)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox1.TabIndex = 3
        Me.PictureBox1.TabStop = False
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 65)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Padding = New System.Windows.Forms.Padding(5)
        Me.Panel2.Size = New System.Drawing.Size(521, 224)
        Me.Panel2.TabIndex = 4
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.txtRegSN)
        Me.GroupBox1.Controls.Add(Me.dtpUseDate)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.dtpInsDate)
        Me.GroupBox1.Controls.Add(Me.lblInsDate)
        Me.GroupBox1.Controls.Add(Me.txtApp)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtMac)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.txtMB)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtHD)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtCpu)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtSoftSN)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox1.Location = New System.Drawing.Point(5, 5)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(511, 216)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "授權資訊"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(19, 186)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(68, 16)
        Me.Label9.TabIndex = 18
        Me.Label9.Text = "註冊序號"
        '
        'txtRegSN
        '
        Me.txtRegSN.Location = New System.Drawing.Point(93, 183)
        Me.txtRegSN.Name = "txtRegSN"
        Me.txtRegSN.ReadOnly = True
        Me.txtRegSN.Size = New System.Drawing.Size(410, 23)
        Me.txtRegSN.TabIndex = 17
        '
        'dtpUseDate
        '
        Me.dtpUseDate.CustomFormat = "yyyy/MM/dd"
        Me.dtpUseDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpUseDate.Location = New System.Drawing.Point(352, 150)
        Me.dtpUseDate.Name = "dtpUseDate"
        Me.dtpUseDate.Size = New System.Drawing.Size(151, 23)
        Me.dtpUseDate.TabIndex = 16
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(266, 153)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(83, 16)
        Me.Label8.TabIndex = 15
        Me.Label8.Text = "使用到期日"
        '
        'dtpInsDate
        '
        Me.dtpInsDate.CustomFormat = "yyyy/MM/dd"
        Me.dtpInsDate.Enabled = False
        Me.dtpInsDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpInsDate.Location = New System.Drawing.Point(93, 150)
        Me.dtpInsDate.Name = "dtpInsDate"
        Me.dtpInsDate.Size = New System.Drawing.Size(151, 23)
        Me.dtpInsDate.TabIndex = 14
        '
        'lblInsDate
        '
        Me.lblInsDate.AutoSize = True
        Me.lblInsDate.Location = New System.Drawing.Point(23, 153)
        Me.lblInsDate.Name = "lblInsDate"
        Me.lblInsDate.Size = New System.Drawing.Size(68, 16)
        Me.lblInsDate.TabIndex = 13
        Me.lblInsDate.Text = "安裝日期"
        '
        'txtApp
        '
        Me.txtApp.Location = New System.Drawing.Point(93, 116)
        Me.txtApp.Name = "txtApp"
        Me.txtApp.ReadOnly = True
        Me.txtApp.Size = New System.Drawing.Size(410, 23)
        Me.txtApp.TabIndex = 12
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(22, 121)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(68, 16)
        Me.Label7.TabIndex = 11
        Me.Label7.Text = "應用程式"
        '
        'txtMac
        '
        Me.txtMac.Location = New System.Drawing.Point(352, 85)
        Me.txtMac.Name = "txtMac"
        Me.txtMac.ReadOnly = True
        Me.txtMac.Size = New System.Drawing.Size(151, 23)
        Me.txtMac.TabIndex = 10
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(285, 89)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 16)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "MAC序號"
        '
        'txtMB
        '
        Me.txtMB.Location = New System.Drawing.Point(93, 85)
        Me.txtMB.Name = "txtMB"
        Me.txtMB.ReadOnly = True
        Me.txtMB.Size = New System.Drawing.Size(151, 23)
        Me.txtMB.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(7, 89)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(83, 16)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "主機板序號"
        '
        'txtHD
        '
        Me.txtHD.Location = New System.Drawing.Point(352, 55)
        Me.txtHD.Name = "txtHD"
        Me.txtHD.ReadOnly = True
        Me.txtHD.Size = New System.Drawing.Size(151, 23)
        Me.txtHD.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(282, 58)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(68, 16)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "硬碟序號"
        '
        'txtCpu
        '
        Me.txtCpu.Location = New System.Drawing.Point(93, 55)
        Me.txtCpu.Name = "txtCpu"
        Me.txtCpu.ReadOnly = True
        Me.txtCpu.Size = New System.Drawing.Size(151, 23)
        Me.txtCpu.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(30, 58)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 16)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "CPU序號"
        '
        'txtSoftSN
        '
        Me.txtSoftSN.Location = New System.Drawing.Point(93, 24)
        Me.txtSoftSN.Name = "txtSoftSN"
        Me.txtSoftSN.Size = New System.Drawing.Size(410, 23)
        Me.txtSoftSN.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(23, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "軟體序號"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(428, 290)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 27)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "GenKey"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCopy
        '
        Me.btnCopy.Enabled = False
        Me.btnCopy.Location = New System.Drawing.Point(312, 290)
        Me.btnCopy.Name = "btnCopy"
        Me.btnCopy.Size = New System.Drawing.Size(99, 27)
        Me.btnCopy.TabIndex = 6
        Me.btnCopy.Text = "CopyRegKey"
        Me.btnCopy.UseVisualStyleBackColor = True
        '
        'frmKey
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(521, 323)
        Me.Controls.Add(Me.btnCopy)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmKey"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "軟體授權"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtSoftSN As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtCpu As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtHD As System.Windows.Forms.TextBox
    Friend WithEvents txtMac As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtMB As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtApp As System.Windows.Forms.TextBox
    Friend WithEvents dtpInsDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblInsDate As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtRegSN As System.Windows.Forms.TextBox
    Friend WithEvents dtpUseDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents btnCopy As System.Windows.Forms.Button

End Class
