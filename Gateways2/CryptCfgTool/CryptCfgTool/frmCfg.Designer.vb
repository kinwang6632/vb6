<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCryptCfg
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
        Me.btnPath = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.txtFileName = New System.Windows.Forms.TextBox()
        Me.lsvTable = New System.Windows.Forms.ListView()
        Me.btnDecry = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnPath
        '
        Me.btnPath.Location = New System.Drawing.Point(327, 12)
        Me.btnPath.Name = "btnPath"
        Me.btnPath.Size = New System.Drawing.Size(44, 23)
        Me.btnPath.TabIndex = 0
        Me.btnPath.Text = "......"
        Me.btnPath.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'txtFileName
        '
        Me.txtFileName.Location = New System.Drawing.Point(12, 12)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.ReadOnly = True
        Me.txtFileName.Size = New System.Drawing.Size(309, 22)
        Me.txtFileName.TabIndex = 1
        '
        'lsvTable
        '
        Me.lsvTable.CheckBoxes = True
        Me.lsvTable.Location = New System.Drawing.Point(12, 41)
        Me.lsvTable.Name = "lsvTable"
        Me.lsvTable.Size = New System.Drawing.Size(359, 166)
        Me.lsvTable.TabIndex = 2
        Me.lsvTable.UseCompatibleStateImageBehavior = False
        Me.lsvTable.View = System.Windows.Forms.View.List
        '
        'btnDecry
        '
        Me.btnDecry.Location = New System.Drawing.Point(12, 218)
        Me.btnDecry.Name = "btnDecry"
        Me.btnDecry.Size = New System.Drawing.Size(75, 23)
        Me.btnDecry.TabIndex = 3
        Me.btnDecry.Text = "解密"
        Me.btnDecry.UseVisualStyleBackColor = True
        '
        'frmCryptCfg
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(383, 253)
        Me.Controls.Add(Me.btnDecry)
        Me.Controls.Add(Me.lsvTable)
        Me.Controls.Add(Me.txtFileName)
        Me.Controls.Add(Me.btnPath)
        Me.Name = "frmCryptCfg"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CFG加解密工具"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnPath As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtFileName As System.Windows.Forms.TextBox
    Friend WithEvents lsvTable As System.Windows.Forms.ListView
    Friend WithEvents btnDecry As System.Windows.Forms.Button

End Class
