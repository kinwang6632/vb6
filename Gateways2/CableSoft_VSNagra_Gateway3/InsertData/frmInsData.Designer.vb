<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInsData
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
        Me.btnOK = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSid = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtUserId = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtHIGH_LEVEL_CMD_ID = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtICC_NO = New System.Windows.Forms.TextBox()
        Me.txtSTB_NO = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtNOTES = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtCompCode = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtRcd = New System.Windows.Forms.TextBox()
        Me.txt_ZipCode = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.dtpSubscription_begin_date = New System.Windows.Forms.DateTimePicker()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.dtpSubscription_end_date = New System.Windows.Forms.DateTimePicker()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtMIS_IRD_CMD_DATA = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtProcessingDate = New System.Windows.Forms.MaskedTextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtGatewayType = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(12, 247)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(98, 24)
        Me.btnOK.TabIndex = 0
        Me.btnOK.Text = "確定"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(23, 12)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "SID"
        '
        'txtSid
        '
        Me.txtSid.Location = New System.Drawing.Point(37, 19)
        Me.txtSid.Name = "txtSid"
        Me.txtSid.Size = New System.Drawing.Size(111, 22)
        Me.txtSid.TabIndex = 2
        Me.txtSid.Text = "RDKNET"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(154, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(36, 12)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "UserId"
        '
        'txtUserId
        '
        Me.txtUserId.Location = New System.Drawing.Point(196, 19)
        Me.txtUserId.Name = "txtUserId"
        Me.txtUserId.Size = New System.Drawing.Size(83, 22)
        Me.txtUserId.TabIndex = 4
        Me.txtUserId.Text = "tbcsh"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(285, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 12)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Password"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(339, 19)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(83, 22)
        Me.txtPassword.TabIndex = 6
        Me.txtPassword.Text = "tbcsh"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(438, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(125, 12)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "HIGH_LEVEL_CMD_ID"
        '
        'txtHIGH_LEVEL_CMD_ID
        '
        Me.txtHIGH_LEVEL_CMD_ID.Location = New System.Drawing.Point(569, 19)
        Me.txtHIGH_LEVEL_CMD_ID.Name = "txtHIGH_LEVEL_CMD_ID"
        Me.txtHIGH_LEVEL_CMD_ID.Size = New System.Drawing.Size(83, 22)
        Me.txtHIGH_LEVEL_CMD_ID.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(10, 61)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(47, 12)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "ICC_NO"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(179, 61)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(48, 12)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "STB_NO"
        '
        'txtICC_NO
        '
        Me.txtICC_NO.Location = New System.Drawing.Point(62, 55)
        Me.txtICC_NO.Name = "txtICC_NO"
        Me.txtICC_NO.Size = New System.Drawing.Size(111, 22)
        Me.txtICC_NO.TabIndex = 11
        Me.txtICC_NO.Text = "1234567890123456"
        '
        'txtSTB_NO
        '
        Me.txtSTB_NO.Location = New System.Drawing.Point(233, 55)
        Me.txtSTB_NO.Name = "txtSTB_NO"
        Me.txtSTB_NO.Size = New System.Drawing.Size(111, 22)
        Me.txtSTB_NO.TabIndex = 12
        Me.txtSTB_NO.Text = "9999999999999999"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(12, 98)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(41, 12)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "NOTES"
        '
        'txtNOTES
        '
        Me.txtNOTES.Location = New System.Drawing.Point(56, 92)
        Me.txtNOTES.Name = "txtNOTES"
        Me.txtNOTES.Size = New System.Drawing.Size(651, 22)
        Me.txtNOTES.TabIndex = 14
        Me.txtNOTES.Text = "A~1,A~2,A~3,B~5~20020201~20140205"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(350, 58)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(59, 12)
        Me.Label8.TabIndex = 15
        Me.Label8.Text = "CompCode"
        '
        'txtCompCode
        '
        Me.txtCompCode.Location = New System.Drawing.Point(415, 55)
        Me.txtCompCode.Name = "txtCompCode"
        Me.txtCompCode.Size = New System.Drawing.Size(60, 22)
        Me.txtCompCode.TabIndex = 16
        Me.txtCompCode.Text = "3"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(136, 253)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(53, 12)
        Me.Label9.TabIndex = 17
        Me.Label9.Text = "新增筆數"
        '
        'txtRcd
        '
        Me.txtRcd.Location = New System.Drawing.Point(198, 249)
        Me.txtRcd.Name = "txtRcd"
        Me.txtRcd.Size = New System.Drawing.Size(89, 22)
        Me.txtRcd.TabIndex = 18
        Me.txtRcd.Text = "1000"
        '
        'txt_ZipCode
        '
        Me.txt_ZipCode.Location = New System.Drawing.Point(569, 55)
        Me.txt_ZipCode.Name = "txt_ZipCode"
        Me.txt_ZipCode.Size = New System.Drawing.Size(60, 22)
        Me.txt_ZipCode.TabIndex = 19
        Me.txt_ZipCode.Text = "12"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(493, 58)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(52, 12)
        Me.Label10.TabIndex = 20
        Me.Label10.Text = "Zip_Code"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(12, 133)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(118, 12)
        Me.Label11.TabIndex = 21
        Me.Label11.Text = "subscription_begin_date"
        '
        'dtpSubscription_begin_date
        '
        Me.dtpSubscription_begin_date.Location = New System.Drawing.Point(136, 126)
        Me.dtpSubscription_begin_date.Name = "dtpSubscription_begin_date"
        Me.dtpSubscription_begin_date.Size = New System.Drawing.Size(133, 22)
        Me.dtpSubscription_begin_date.TabIndex = 22
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(275, 133)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(109, 12)
        Me.Label12.TabIndex = 23
        Me.Label12.Text = "subscription_end_date"
        '
        'dtpSubscription_end_date
        '
        Me.dtpSubscription_end_date.Location = New System.Drawing.Point(399, 126)
        Me.dtpSubscription_end_date.Name = "dtpSubscription_end_date"
        Me.dtpSubscription_end_date.Size = New System.Drawing.Size(133, 22)
        Me.dtpSubscription_end_date.TabIndex = 24
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(12, 175)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(120, 12)
        Me.Label13.TabIndex = 25
        Me.Label13.Text = "MIS_IRD_CMD_DATA"
        '
        'txtMIS_IRD_CMD_DATA
        '
        Me.txtMIS_IRD_CMD_DATA.Location = New System.Drawing.Point(136, 172)
        Me.txtMIS_IRD_CMD_DATA.Name = "txtMIS_IRD_CMD_DATA"
        Me.txtMIS_IRD_CMD_DATA.Size = New System.Drawing.Size(143, 22)
        Me.txtMIS_IRD_CMD_DATA.TabIndex = 26
        Me.txtMIS_IRD_CMD_DATA.Text = "7~2~22"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(285, 175)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(53, 12)
        Me.Label14.TabIndex = 27
        Me.Label14.Text = "預約時間"
        '
        'txtProcessingDate
        '
        Me.txtProcessingDate.Location = New System.Drawing.Point(339, 172)
        Me.txtProcessingDate.Mask = "0000/00/00 00:00:00"
        Me.txtProcessingDate.Name = "txtProcessingDate"
        Me.txtProcessingDate.Size = New System.Drawing.Size(175, 22)
        Me.txtProcessingDate.TabIndex = 29
        Me.txtProcessingDate.TextMaskFormat = System.Windows.Forms.MaskFormat.ExcludePromptAndLiterals
        Me.txtProcessingDate.ValidatingType = GetType(Date)
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(12, 205)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(72, 12)
        Me.Label15.TabIndex = 30
        Me.Label15.Text = "GateWayType"
        '
        'txtGatewayType
        '
        Me.txtGatewayType.Location = New System.Drawing.Point(90, 202)
        Me.txtGatewayType.Name = "txtGatewayType"
        Me.txtGatewayType.Size = New System.Drawing.Size(20, 22)
        Me.txtGatewayType.TabIndex = 31
        Me.txtGatewayType.Text = "0"
        '
        'frmInsData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(719, 283)
        Me.Controls.Add(Me.txtGatewayType)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.txtProcessingDate)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.txtMIS_IRD_CMD_DATA)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.dtpSubscription_end_date)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.dtpSubscription_begin_date)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.txt_ZipCode)
        Me.Controls.Add(Me.txtRcd)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.txtCompCode)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtNOTES)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtSTB_NO)
        Me.Controls.Add(Me.txtICC_NO)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtHIGH_LEVEL_CMD_ID)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtUserId)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtSid)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnOK)
        Me.Name = "frmInsData"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmInsData"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtSid As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtUserId As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtHIGH_LEVEL_CMD_ID As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtICC_NO As System.Windows.Forms.TextBox
    Friend WithEvents txtSTB_NO As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtNOTES As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtCompCode As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtRcd As System.Windows.Forms.TextBox
    Friend WithEvents txt_ZipCode As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents dtpSubscription_begin_date As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents dtpSubscription_end_date As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtMIS_IRD_CMD_DATA As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtProcessingDate As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtGatewayType As System.Windows.Forms.TextBox

End Class
