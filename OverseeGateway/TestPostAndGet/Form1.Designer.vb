<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPosAndGet
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
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnPost = New System.Windows.Forms.Button()
        Me.btnGet = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnPost
        '
        Me.btnPost.Location = New System.Drawing.Point(0, 0)
        Me.btnPost.Name = "btnPost"
        Me.btnPost.Size = New System.Drawing.Size(75, 23)
        Me.btnPost.TabIndex = 0
        Me.btnPost.Text = "Post"
        Me.btnPost.UseVisualStyleBackColor = True
        '
        'btnGet
        '
        Me.btnGet.Location = New System.Drawing.Point(0, 29)
        Me.btnGet.Name = "btnGet"
        Me.btnGet.Size = New System.Drawing.Size(75, 23)
        Me.btnGet.TabIndex = 1
        Me.btnGet.Text = "Get"
        Me.btnGet.UseVisualStyleBackColor = True
        '
        'frmPosAndGet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(282, 253)
        Me.Controls.Add(Me.btnGet)
        Me.Controls.Add(Me.btnPost)
        Me.Name = "frmPosAndGet"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnPost As Button
    Friend WithEvents btnGet As Button
End Class
