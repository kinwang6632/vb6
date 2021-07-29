<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLowCmd
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
        Me.tabLowCmd = New System.Windows.Forms.TabControl()
        Me.tabSetting = New System.Windows.Forms.TabPage()
        Me.btnSaveSet = New System.Windows.Forms.Button()
        Me.txtPort = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtIPAddress = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtRootKey = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtSMS_ID = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtOPE_ID = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtKey_Type = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtCrypt_Ver = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtProto_Ver = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbRun = New System.Windows.Forms.TabPage()
        Me.txtCha = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.txtNo = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.txtFeed_Interval_Hour = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.txtParent_Card_SN = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.btnSavePara = New System.Windows.Forms.Button()
        Me.txtENTITLEEXT = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtSTBID = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtCard_SN = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtMsg = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtRecvData = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtSendData = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnDisconnect = New System.Windows.Forms.Button()
        Me.btnCheck = New System.Windows.Forms.Button()
        Me.btnSend = New System.Windows.Forms.Button()
        Me.btnConnect = New System.Windows.Forms.Button()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.cboMsgId = New System.Windows.Forms.ComboBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.tabLowCmd.SuspendLayout()
        Me.tabSetting.SuspendLayout()
        Me.tbRun.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabLowCmd
        '
        Me.tabLowCmd.Controls.Add(Me.tabSetting)
        Me.tabLowCmd.Controls.Add(Me.tbRun)
        Me.tabLowCmd.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabLowCmd.Location = New System.Drawing.Point(0, 0)
        Me.tabLowCmd.Margin = New System.Windows.Forms.Padding(4)
        Me.tabLowCmd.Name = "tabLowCmd"
        Me.tabLowCmd.SelectedIndex = 0
        Me.tabLowCmd.Size = New System.Drawing.Size(1036, 404)
        Me.tabLowCmd.TabIndex = 0
        '
        'tabSetting
        '
        Me.tabSetting.Controls.Add(Me.btnSaveSet)
        Me.tabSetting.Controls.Add(Me.txtPort)
        Me.tabSetting.Controls.Add(Me.Label8)
        Me.tabSetting.Controls.Add(Me.txtIPAddress)
        Me.tabSetting.Controls.Add(Me.Label7)
        Me.tabSetting.Controls.Add(Me.txtRootKey)
        Me.tabSetting.Controls.Add(Me.Label6)
        Me.tabSetting.Controls.Add(Me.txtSMS_ID)
        Me.tabSetting.Controls.Add(Me.Label5)
        Me.tabSetting.Controls.Add(Me.txtOPE_ID)
        Me.tabSetting.Controls.Add(Me.Label4)
        Me.tabSetting.Controls.Add(Me.txtKey_Type)
        Me.tabSetting.Controls.Add(Me.Label3)
        Me.tabSetting.Controls.Add(Me.txtCrypt_Ver)
        Me.tabSetting.Controls.Add(Me.Label2)
        Me.tabSetting.Controls.Add(Me.txtProto_Ver)
        Me.tabSetting.Controls.Add(Me.Label1)
        Me.tabSetting.Location = New System.Drawing.Point(4, 26)
        Me.tabSetting.Margin = New System.Windows.Forms.Padding(4)
        Me.tabSetting.Name = "tabSetting"
        Me.tabSetting.Padding = New System.Windows.Forms.Padding(4)
        Me.tabSetting.Size = New System.Drawing.Size(1028, 358)
        Me.tabSetting.TabIndex = 0
        Me.tabSetting.Text = "參數設定"
        Me.tabSetting.UseVisualStyleBackColor = True
        '
        'btnSaveSet
        '
        Me.btnSaveSet.Location = New System.Drawing.Point(15, 178)
        Me.btnSaveSet.Name = "btnSaveSet"
        Me.btnSaveSet.Size = New System.Drawing.Size(108, 28)
        Me.btnSaveSet.TabIndex = 16
        Me.btnSaveSet.Text = "儲存設定"
        Me.btnSaveSet.UseVisualStyleBackColor = True
        '
        'txtPort
        '
        Me.txtPort.Location = New System.Drawing.Point(318, 117)
        Me.txtPort.Name = "txtPort"
        Me.txtPort.Size = New System.Drawing.Size(170, 27)
        Me.txtPort.TabIndex = 15
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(278, 120)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(33, 16)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Port"
        '
        'txtIPAddress
        '
        Me.txtIPAddress.Location = New System.Drawing.Point(91, 117)
        Me.txtIPAddress.Name = "txtIPAddress"
        Me.txtIPAddress.Size = New System.Drawing.Size(170, 27)
        Me.txtIPAddress.TabIndex = 13
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(12, 120)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(76, 16)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "IP Address"
        '
        'txtRootKey
        '
        Me.txtRootKey.Location = New System.Drawing.Point(481, 74)
        Me.txtRootKey.Name = "txtRootKey"
        Me.txtRootKey.Size = New System.Drawing.Size(176, 27)
        Me.txtRootKey.TabIndex = 11
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(400, 77)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(68, 16)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Root Key"
        '
        'txtSMS_ID
        '
        Me.txtSMS_ID.Location = New System.Drawing.Point(281, 74)
        Me.txtSMS_ID.Name = "txtSMS_ID"
        Me.txtSMS_ID.Size = New System.Drawing.Size(100, 27)
        Me.txtSMS_ID.TabIndex = 9
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(200, 77)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(61, 16)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "SMS_ID"
        '
        'txtOPE_ID
        '
        Me.txtOPE_ID.Location = New System.Drawing.Point(91, 74)
        Me.txtOPE_ID.Name = "txtOPE_ID"
        Me.txtOPE_ID.Size = New System.Drawing.Size(100, 27)
        Me.txtOPE_ID.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 77)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(60, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "OPE_ID"
        '
        'txtKey_Type
        '
        Me.txtKey_Type.Location = New System.Drawing.Point(481, 24)
        Me.txtKey_Type.Name = "txtKey_Type"
        Me.txtKey_Type.Size = New System.Drawing.Size(100, 27)
        Me.txtKey_Type.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(400, 27)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 16)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Key_Type"
        '
        'txtCrypt_Ver
        '
        Me.txtCrypt_Ver.Location = New System.Drawing.Point(281, 24)
        Me.txtCrypt_Ver.Name = "txtCrypt_Ver"
        Me.txtCrypt_Ver.Size = New System.Drawing.Size(100, 27)
        Me.txtCrypt_Ver.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(200, 27)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(74, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Crypt_Ver"
        '
        'txtProto_Ver
        '
        Me.txtProto_Ver.Location = New System.Drawing.Point(91, 24)
        Me.txtProto_Ver.Name = "txtProto_Ver"
        Me.txtProto_Ver.Size = New System.Drawing.Size(100, 27)
        Me.txtProto_Ver.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 27)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Proto_Ver"
        '
        'tbRun
        '
        Me.tbRun.Controls.Add(Me.Label20)
        Me.tbRun.Controls.Add(Me.txtCha)
        Me.tbRun.Controls.Add(Me.Label19)
        Me.tbRun.Controls.Add(Me.txtNo)
        Me.tbRun.Controls.Add(Me.Label18)
        Me.tbRun.Controls.Add(Me.txtFeed_Interval_Hour)
        Me.tbRun.Controls.Add(Me.Label17)
        Me.tbRun.Controls.Add(Me.txtParent_Card_SN)
        Me.tbRun.Controls.Add(Me.Label16)
        Me.tbRun.Controls.Add(Me.btnSavePara)
        Me.tbRun.Controls.Add(Me.txtENTITLEEXT)
        Me.tbRun.Controls.Add(Me.Label15)
        Me.tbRun.Controls.Add(Me.txtSTBID)
        Me.tbRun.Controls.Add(Me.Label14)
        Me.tbRun.Controls.Add(Me.txtCard_SN)
        Me.tbRun.Controls.Add(Me.Label13)
        Me.tbRun.Controls.Add(Me.GroupBox2)
        Me.tbRun.Controls.Add(Me.GroupBox1)
        Me.tbRun.Location = New System.Drawing.Point(4, 26)
        Me.tbRun.Margin = New System.Windows.Forms.Padding(4)
        Me.tbRun.Name = "tbRun"
        Me.tbRun.Padding = New System.Windows.Forms.Padding(4)
        Me.tbRun.Size = New System.Drawing.Size(1028, 374)
        Me.tbRun.TabIndex = 1
        Me.tbRun.Text = "執行命令"
        Me.tbRun.UseVisualStyleBackColor = True
        '
        'txtCha
        '
        Me.txtCha.Location = New System.Drawing.Point(312, 303)
        Me.txtCha.MaxLength = 0
        Me.txtCha.Name = "txtCha"
        Me.txtCha.Size = New System.Drawing.Size(59, 27)
        Me.txtCha.TabIndex = 22
        Me.txtCha.Text = "aaaa"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(273, 306)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(33, 16)
        Me.Label19.TabIndex = 21
        Me.Label19.Text = "Cha"
        '
        'txtNo
        '
        Me.txtNo.Location = New System.Drawing.Point(236, 303)
        Me.txtNo.MaxLength = 1
        Me.txtNo.Name = "txtNo"
        Me.txtNo.Size = New System.Drawing.Size(20, 27)
        Me.txtNo.TabIndex = 20
        Me.txtNo.Text = "9"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(203, 306)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(27, 16)
        Me.Label18.TabIndex = 19
        Me.Label18.Text = "No"
        '
        'txtFeed_Interval_Hour
        '
        Me.txtFeed_Interval_Hour.Location = New System.Drawing.Point(150, 303)
        Me.txtFeed_Interval_Hour.Name = "txtFeed_Interval_Hour"
        Me.txtFeed_Interval_Hour.Size = New System.Drawing.Size(35, 27)
        Me.txtFeed_Interval_Hour.TabIndex = 18
        Me.txtFeed_Interval_Hour.Text = "15"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(10, 306)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(134, 16)
        Me.Label17.TabIndex = 17
        Me.Label17.Text = "Feed_Interval_Hour"
        '
        'txtParent_Card_SN
        '
        Me.txtParent_Card_SN.Location = New System.Drawing.Point(673, 171)
        Me.txtParent_Card_SN.Name = "txtParent_Card_SN"
        Me.txtParent_Card_SN.Size = New System.Drawing.Size(169, 27)
        Me.txtParent_Card_SN.TabIndex = 16
        Me.txtParent_Card_SN.Text = "1234567891234567"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(562, 174)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(112, 16)
        Me.Label16.TabIndex = 15
        Me.Label16.Text = "Parent_Card_SN"
        '
        'btnSavePara
        '
        Me.btnSavePara.Location = New System.Drawing.Point(0, 340)
        Me.btnSavePara.Name = "btnSavePara"
        Me.btnSavePara.Size = New System.Drawing.Size(109, 26)
        Me.btnSavePara.TabIndex = 14
        Me.btnSavePara.Text = "儲存參數"
        Me.btnSavePara.UseVisualStyleBackColor = True
        '
        'txtENTITLEEXT
        '
        Me.txtENTITLEEXT.Location = New System.Drawing.Point(85, 218)
        Me.txtENTITLEEXT.Name = "txtENTITLEEXT"
        Me.txtENTITLEEXT.Size = New System.Drawing.Size(911, 27)
        Me.txtENTITLEEXT.TabIndex = 13
        Me.txtENTITLEEXT.Text = "1,1023,0,2005/3/8 15:01:40,2030/12/31 00:00:00;1,1023,0,2005/3/8 15:01:40,2030/12" & _
    "/31 00:00:00"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(8, 218)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(72, 16)
        Me.Label15.TabIndex = 12
        Me.Label15.Text = "授權資料"
        '
        'txtSTBID
        '
        Me.txtSTBID.Location = New System.Drawing.Point(312, 172)
        Me.txtSTBID.MaxLength = 30
        Me.txtSTBID.Name = "txtSTBID"
        Me.txtSTBID.Size = New System.Drawing.Size(244, 27)
        Me.txtSTBID.TabIndex = 11
        Me.txtSTBID.Text = "11111122222233333344444455555"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(260, 174)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(51, 16)
        Me.Label14.TabIndex = 10
        Me.Label14.Text = "STBID"
        '
        'txtCard_SN
        '
        Me.txtCard_SN.Location = New System.Drawing.Point(85, 171)
        Me.txtCard_SN.Name = "txtCard_SN"
        Me.txtCard_SN.Size = New System.Drawing.Size(169, 27)
        Me.txtCard_SN.TabIndex = 8
        Me.txtCard_SN.Text = "1234567891234567"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(8, 175)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(77, 16)
        Me.Label13.TabIndex = 7
        Me.Label13.Text = "CARD_SN"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtMsg)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.txtRecvData)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Controls.Add(Me.txtSendData)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Location = New System.Drawing.Point(501, 19)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(519, 138)
        Me.GroupBox2.TabIndex = 6
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "回應訊息"
        '
        'txtMsg
        '
        Me.txtMsg.Location = New System.Drawing.Point(84, 98)
        Me.txtMsg.Name = "txtMsg"
        Me.txtMsg.ReadOnly = True
        Me.txtMsg.Size = New System.Drawing.Size(429, 27)
        Me.txtMsg.TabIndex = 5
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(6, 101)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(72, 16)
        Me.Label12.TabIndex = 4
        Me.Label12.Text = "詳細訊息"
        '
        'txtRecvData
        '
        Me.txtRecvData.Location = New System.Drawing.Point(84, 60)
        Me.txtRecvData.Name = "txtRecvData"
        Me.txtRecvData.ReadOnly = True
        Me.txtRecvData.Size = New System.Drawing.Size(429, 27)
        Me.txtRecvData.TabIndex = 3
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(6, 68)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(72, 16)
        Me.Label11.TabIndex = 2
        Me.Label11.Text = "接收數據"
        '
        'txtSendData
        '
        Me.txtSendData.Location = New System.Drawing.Point(84, 22)
        Me.txtSendData.Name = "txtSendData"
        Me.txtSendData.ReadOnly = True
        Me.txtSendData.Size = New System.Drawing.Size(429, 27)
        Me.txtSendData.TabIndex = 1
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(6, 28)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(72, 16)
        Me.Label10.TabIndex = 0
        Me.Label10.Text = "送出數據"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnDisconnect)
        Me.GroupBox1.Controls.Add(Me.btnCheck)
        Me.GroupBox1.Controls.Add(Me.btnSend)
        Me.GroupBox1.Controls.Add(Me.btnConnect)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.cboMsgId)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 7)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(487, 150)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        '
        'btnDisconnect
        '
        Me.btnDisconnect.Location = New System.Drawing.Point(337, 80)
        Me.btnDisconnect.Name = "btnDisconnect"
        Me.btnDisconnect.Size = New System.Drawing.Size(105, 30)
        Me.btnDisconnect.TabIndex = 10
        Me.btnDisconnect.Text = "斷線"
        Me.btnDisconnect.UseVisualStyleBackColor = True
        '
        'btnCheck
        '
        Me.btnCheck.Location = New System.Drawing.Point(118, 80)
        Me.btnCheck.Name = "btnCheck"
        Me.btnCheck.Size = New System.Drawing.Size(102, 29)
        Me.btnCheck.TabIndex = 9
        Me.btnCheck.Text = "定時命令"
        Me.btnCheck.UseVisualStyleBackColor = True
        '
        'btnSend
        '
        Me.btnSend.Location = New System.Drawing.Point(226, 80)
        Me.btnSend.Name = "btnSend"
        Me.btnSend.Size = New System.Drawing.Size(105, 30)
        Me.btnSend.TabIndex = 8
        Me.btnSend.Text = "送出命令"
        Me.btnSend.UseVisualStyleBackColor = True
        '
        'btnConnect
        '
        Me.btnConnect.Location = New System.Drawing.Point(22, 80)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.Size = New System.Drawing.Size(90, 29)
        Me.btnConnect.TabIndex = 7
        Me.btnConnect.Text = "建立連線"
        Me.btnConnect.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(19, 40)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(72, 16)
        Me.Label9.TabIndex = 6
        Me.Label9.Text = "命令種類"
        '
        'cboMsgId
        '
        Me.cboMsgId.FormattingEnabled = True
        Me.cboMsgId.Location = New System.Drawing.Point(98, 37)
        Me.cboMsgId.Name = "cboMsgId"
        Me.cboMsgId.Size = New System.Drawing.Size(374, 24)
        Me.cboMsgId.TabIndex = 5
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(82, 258)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(601, 16)
        Me.Label20.TabIndex = 23
        Me.Label20.Text = "授權種類(0=反授權,1=授權),ProductId,錄影功能(0=是,1=否),觀賞時間(起),觀賞時間(迄);"
        '
        'frmLowCmd
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1036, 404)
        Me.Controls.Add(Me.tabLowCmd)
        Me.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmLowCmd"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "測試指令"
        Me.tabLowCmd.ResumeLayout(False)
        Me.tabSetting.ResumeLayout(False)
        Me.tabSetting.PerformLayout()
        Me.tbRun.ResumeLayout(False)
        Me.tbRun.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tabLowCmd As System.Windows.Forms.TabControl
    Friend WithEvents tabSetting As System.Windows.Forms.TabPage
    Friend WithEvents tbRun As System.Windows.Forms.TabPage
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtProto_Ver As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtCrypt_Ver As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtKey_Type As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtOPE_ID As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtRootKey As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtSMS_ID As System.Windows.Forms.TextBox
    Friend WithEvents txtPort As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtIPAddress As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnCheck As System.Windows.Forms.Button
    Friend WithEvents btnSend As System.Windows.Forms.Button
    Friend WithEvents btnConnect As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cboMsgId As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtMsg As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtRecvData As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtSendData As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtCard_SN As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtSTBID As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtENTITLEEXT As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents btnSaveSet As System.Windows.Forms.Button
    Friend WithEvents btnSavePara As System.Windows.Forms.Button
    Friend WithEvents txtFeed_Interval_Hour As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtParent_Card_SN As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtCha As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtNo As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents btnDisconnect As System.Windows.Forms.Button
    Friend WithEvents Label20 As System.Windows.Forms.Label

End Class
