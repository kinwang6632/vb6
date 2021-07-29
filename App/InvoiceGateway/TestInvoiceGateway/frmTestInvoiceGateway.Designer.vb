<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTestInvoiceGateway
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
        Me.btnExecute = New System.Windows.Forms.Button()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnExecute
        '
        Me.btnExecute.Location = New System.Drawing.Point(30, 12)
        Me.btnExecute.Name = "btnExecute"
        Me.btnExecute.Size = New System.Drawing.Size(139, 23)
        Me.btnExecute.TabIndex = 0
        Me.btnExecute.Text = "執行"
        Me.btnExecute.UseVisualStyleBackColor = True
        '
        'btnStop
        '
        Me.btnStop.Enabled = False
        Me.btnStop.Location = New System.Drawing.Point(30, 54)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(139, 23)
        Me.btnStop.TabIndex = 1
        Me.btnStop.Text = "停止"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'frmTestInvoiceGateway
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(204, 92)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.btnExecute)
        Me.Name = "frmTestInvoiceGateway"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "測試執行"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnExecute As System.Windows.Forms.Button
    Friend WithEvents btnStop As System.Windows.Forms.Button

End Class
