<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTest_NstvProcess
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
        Me.btnNstvProcess = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnNstvProcess
        '
        Me.btnNstvProcess.Location = New System.Drawing.Point(58, 24)
        Me.btnNstvProcess.Name = "btnNstvProcess"
        Me.btnNstvProcess.Size = New System.Drawing.Size(122, 23)
        Me.btnNstvProcess.TabIndex = 1
        Me.btnNstvProcess.Text = "執行Gateway"
        Me.btnNstvProcess.UseVisualStyleBackColor = True
        '
        'frmTest_NstvProcess
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(254, 72)
        Me.Controls.Add(Me.btnNstvProcess)
        Me.Name = "frmTest_NstvProcess"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "測試Gateway程式"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnNstvProcess As System.Windows.Forms.Button

End Class
