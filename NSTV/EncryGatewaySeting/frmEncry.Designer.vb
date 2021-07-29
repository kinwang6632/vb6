<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEncry
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
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtFileName = New System.Windows.Forms.TextBox()
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.btnDescry = New System.Windows.Forms.Button()
        Me.btnEncry = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "Set"
        Me.OpenFileDialog1.FileName = "GatewaySetting.Set"
        Me.OpenFileDialog1.Filter = "Gateway設定檔|*.Set"
        Me.OpenFileDialog1.Title = "選擇設定檔"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 39)
        Me.Label1.Margin = New System.Windows.Forms.Padding(5, 0, 5, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(154, 20)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "選擇Gateway設定檔"
        '
        'txtFileName
        '
        Me.txtFileName.Location = New System.Drawing.Point(167, 36)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(449, 29)
        Me.txtFileName.TabIndex = 1
        '
        'btnSelect
        '
        Me.btnSelect.Location = New System.Drawing.Point(622, 36)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(54, 30)
        Me.btnSelect.TabIndex = 2
        Me.btnSelect.Text = "......"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'btnDescry
        '
        Me.btnDescry.Location = New System.Drawing.Point(18, 87)
        Me.btnDescry.Name = "btnDescry"
        Me.btnDescry.Size = New System.Drawing.Size(75, 33)
        Me.btnDescry.TabIndex = 3
        Me.btnDescry.Text = "解密"
        Me.btnDescry.UseVisualStyleBackColor = True
        '
        'btnEncry
        '
        Me.btnEncry.Location = New System.Drawing.Point(601, 88)
        Me.btnEncry.Name = "btnEncry"
        Me.btnEncry.Size = New System.Drawing.Size(75, 33)
        Me.btnEncry.TabIndex = 4
        Me.btnEncry.Text = "加密"
        Me.btnEncry.UseVisualStyleBackColor = True
        '
        'frmEncry
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(688, 133)
        Me.Controls.Add(Me.btnEncry)
        Me.Controls.Add(Me.btnDescry)
        Me.Controls.Add(Me.btnSelect)
        Me.Controls.Add(Me.txtFileName)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Margin = New System.Windows.Forms.Padding(5)
        Me.Name = "frmEncry"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "加解密Gateway設定檔"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtFileName As System.Windows.Forms.TextBox
    Friend WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents btnDescry As System.Windows.Forms.Button
    Friend WithEvents btnEncry As System.Windows.Forms.Button

End Class
