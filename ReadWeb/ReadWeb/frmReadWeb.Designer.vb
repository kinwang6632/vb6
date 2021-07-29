<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReadWeb
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmReadWeb))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtCode = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.txtResult = New System.Windows.Forms.TextBox
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.txtLicenese = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.numRepeat = New System.Windows.Forms.NumericUpDown
        Me.Label8 = New System.Windows.Forms.Label
        Me.numDuration = New System.Windows.Forms.NumericUpDown
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtSTBNo = New System.Windows.Forms.TextBox
        Me.numTimeout = New System.Windows.Forms.NumericUpDown
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboMetho = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtPara = New System.Windows.Forms.TextBox
        Me.txtUrl = New System.Windows.Forms.TextBox
        Me.lblUrl = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.numRepeat, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numDuration, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numTimeout, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.txtDesc)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.txtCode)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.btnOK)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(681, 371)
        Me.Panel1.TabIndex = 0
        '
        'txtDesc
        '
        Me.txtDesc.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDesc.Location = New System.Drawing.Point(294, 340)
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReadOnly = True
        Me.txtDesc.Size = New System.Drawing.Size(384, 23)
        Me.txtDesc.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(220, 343)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(76, 16)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Description:"
        '
        'txtCode
        '
        Me.txtCode.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCode.Location = New System.Drawing.Point(165, 340)
        Me.txtCode.Name = "txtCode"
        Me.txtCode.ReadOnly = True
        Me.txtCode.Size = New System.Drawing.Size(52, 23)
        Me.txtCode.TabIndex = 1
        Me.txtCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(126, 344)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(42, 16)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Code:"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(21, 340)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(92, 23)
        Me.btnOK.TabIndex = 0
        Me.btnOK.Text = "讀取網頁內容"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.txtResult)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 109)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(681, 210)
        Me.Panel3.TabIndex = 1
        '
        'txtResult
        '
        Me.txtResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtResult.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtResult.Location = New System.Drawing.Point(0, 0)
        Me.txtResult.Multiline = True
        Me.txtResult.Name = "txtResult"
        Me.txtResult.ReadOnly = True
        Me.txtResult.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtResult.Size = New System.Drawing.Size(681, 210)
        Me.txtResult.TabIndex = 0
        Me.txtResult.WordWrap = False
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.txtLicenese)
        Me.Panel2.Controls.Add(Me.Label9)
        Me.Panel2.Controls.Add(Me.numRepeat)
        Me.Panel2.Controls.Add(Me.Label8)
        Me.Panel2.Controls.Add(Me.numDuration)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.txtSTBNo)
        Me.Panel2.Controls.Add(Me.numTimeout)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.cboMetho)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.txtPara)
        Me.Panel2.Controls.Add(Me.txtUrl)
        Me.Panel2.Controls.Add(Me.lblUrl)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(681, 109)
        Me.Panel2.TabIndex = 0
        '
        'txtLicenese
        '
        Me.txtLicenese.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLicenese.Location = New System.Drawing.Point(497, 37)
        Me.txtLicenese.Name = "txtLicenese"
        Me.txtLicenese.Size = New System.Drawing.Size(172, 23)
        Me.txtLicenese.TabIndex = 7
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(433, 40)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(58, 16)
        Me.Label9.TabIndex = 16
        Me.Label9.Text = "授權碼:"
        '
        'numRepeat
        '
        Me.numRepeat.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.numRepeat.Location = New System.Drawing.Point(368, 38)
        Me.numRepeat.Name = "numRepeat"
        Me.numRepeat.Size = New System.Drawing.Size(59, 23)
        Me.numRepeat.TabIndex = 6
        Me.numRepeat.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(309, 40)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(53, 16)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Repeat:"
        '
        'numDuration
        '
        Me.numDuration.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.numDuration.Location = New System.Drawing.Point(233, 38)
        Me.numDuration.Name = "numDuration"
        Me.numDuration.Size = New System.Drawing.Size(59, 23)
        Me.numDuration.TabIndex = 5
        Me.numDuration.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(170, 40)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(61, 16)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "Duration:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(10, 40)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(51, 16)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "STBNo:"
        '
        'txtSTBNo
        '
        Me.txtSTBNo.Location = New System.Drawing.Point(64, 38)
        Me.txtSTBNo.MaxLength = 10
        Me.txtSTBNo.Name = "txtSTBNo"
        Me.txtSTBNo.Size = New System.Drawing.Size(100, 22)
        Me.txtSTBNo.TabIndex = 4
        '
        'numTimeout
        '
        Me.numTimeout.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.numTimeout.Location = New System.Drawing.Point(471, 9)
        Me.numTimeout.Name = "numTimeout"
        Me.numTimeout.Size = New System.Drawing.Size(55, 23)
        Me.numTimeout.TabIndex = 2
        Me.numTimeout.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.numTimeout.Value = New Decimal(New Integer() {30, 0, 0, 0})
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(406, 12)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(62, 16)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "TimeOut:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(532, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 16)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "傳送方式:"
        '
        'cboMetho
        '
        Me.cboMetho.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMetho.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboMetho.FormattingEnabled = True
        Me.cboMetho.Items.AddRange(New Object() {"POST", "GET"})
        Me.cboMetho.Location = New System.Drawing.Point(611, 9)
        Me.cboMetho.Name = "cboMetho"
        Me.cboMetho.Size = New System.Drawing.Size(57, 24)
        Me.cboMetho.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 70)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "傳送字串:"
        '
        'txtPara
        '
        Me.txtPara.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPara.Location = New System.Drawing.Point(82, 67)
        Me.txtPara.Name = "txtPara"
        Me.txtPara.ReadOnly = True
        Me.txtPara.Size = New System.Drawing.Size(587, 23)
        Me.txtPara.TabIndex = 8
        '
        'txtUrl
        '
        Me.txtUrl.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUrl.Location = New System.Drawing.Point(64, 9)
        Me.txtUrl.Name = "txtUrl"
        Me.txtUrl.Size = New System.Drawing.Size(340, 23)
        Me.txtUrl.TabIndex = 1
        '
        'lblUrl
        '
        Me.lblUrl.AutoSize = True
        Me.lblUrl.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUrl.Location = New System.Drawing.Point(26, 12)
        Me.lblUrl.Name = "lblUrl"
        Me.lblUrl.Size = New System.Drawing.Size(35, 16)
        Me.lblUrl.TabIndex = 0
        Me.lblUrl.Text = "URL:"
        '
        'frmReadWeb
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(681, 371)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmReadWeb"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "讀取網頁內容"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.numRepeat, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numDuration, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numTimeout, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents txtUrl As System.Windows.Forms.TextBox
    Friend WithEvents lblUrl As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents txtResult As System.Windows.Forms.TextBox
    Friend WithEvents txtPara As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtCode As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboMetho As System.Windows.Forms.ComboBox
    Friend WithEvents numTimeout As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtSTBNo As System.Windows.Forms.TextBox
    Friend WithEvents numDuration As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtLicenese As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents numRepeat As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label8 As System.Windows.Forms.Label

End Class
