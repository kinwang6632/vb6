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
        Me.btnConnect = New System.Windows.Forms.Button()
        Me.btnSendData = New System.Windows.Forms.Button()
        Me.btnRecv = New System.Windows.Forms.Button()
        Me.btnStopSendData = New System.Windows.Forms.Button()
        Me.btnBeginSend = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnConnect
        '
        Me.btnConnect.Location = New System.Drawing.Point(31, 12)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.Size = New System.Drawing.Size(75, 23)
        Me.btnConnect.TabIndex = 0
        Me.btnConnect.Text = "Connect"
        Me.btnConnect.UseVisualStyleBackColor = True
        '
        'btnSendData
        '
        Me.btnSendData.Location = New System.Drawing.Point(112, 12)
        Me.btnSendData.Name = "btnSendData"
        Me.btnSendData.Size = New System.Drawing.Size(75, 23)
        Me.btnSendData.TabIndex = 1
        Me.btnSendData.Text = "SendData"
        Me.btnSendData.UseVisualStyleBackColor = True
        '
        'btnRecv
        '
        Me.btnRecv.Location = New System.Drawing.Point(193, 12)
        Me.btnRecv.Name = "btnRecv"
        Me.btnRecv.Size = New System.Drawing.Size(65, 20)
        Me.btnRecv.TabIndex = 2
        Me.btnRecv.Text = "RecvData"
        Me.btnRecv.UseVisualStyleBackColor = True
        '
        'btnStopSendData
        '
        Me.btnStopSendData.Location = New System.Drawing.Point(264, 12)
        Me.btnStopSendData.Name = "btnStopSendData"
        Me.btnStopSendData.Size = New System.Drawing.Size(75, 23)
        Me.btnStopSendData.TabIndex = 3
        Me.btnStopSendData.Text = "StopSend"
        Me.btnStopSendData.UseVisualStyleBackColor = True
        '
        'btnBeginSend
        '
        Me.btnBeginSend.Location = New System.Drawing.Point(112, 41)
        Me.btnBeginSend.Name = "btnBeginSend"
        Me.btnBeginSend.Size = New System.Drawing.Size(75, 23)
        Me.btnBeginSend.TabIndex = 4
        Me.btnBeginSend.Text = "Begin Send"
        Me.btnBeginSend.UseVisualStyleBackColor = True
        '
        'frmClient
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(443, 94)
        Me.Controls.Add(Me.btnBeginSend)
        Me.Controls.Add(Me.btnStopSendData)
        Me.Controls.Add(Me.btnRecv)
        Me.Controls.Add(Me.btnSendData)
        Me.Controls.Add(Me.btnConnect)
        Me.Name = "frmClient"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Client"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnConnect As System.Windows.Forms.Button
    Friend WithEvents btnSendData As System.Windows.Forms.Button
    Friend WithEvents btnRecv As System.Windows.Forms.Button
    Friend WithEvents btnStopSendData As System.Windows.Forms.Button
    Friend WithEvents btnBeginSend As System.Windows.Forms.Button

End Class
