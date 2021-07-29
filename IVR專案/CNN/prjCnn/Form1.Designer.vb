<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConnPara
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請不要使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnOpenFile = New System.Windows.Forms.Button
        Me.dtgConn = New System.Windows.Forms.DataGridView
        Me.CompCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ConnectString = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnSave = New System.Windows.Forms.Button
        CType(Me.dtgConn, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOpenFile
        '
        Me.btnOpenFile.Location = New System.Drawing.Point(12, 270)
        Me.btnOpenFile.Name = "btnOpenFile"
        Me.btnOpenFile.Size = New System.Drawing.Size(75, 23)
        Me.btnOpenFile.TabIndex = 0
        Me.btnOpenFile.Text = "開啟檔案"
        Me.btnOpenFile.UseVisualStyleBackColor = True
        '
        'dtgConn
        '
        Me.dtgConn.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dtgConn.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CompCode, Me.ConnectString})
        Me.dtgConn.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnF2
        Me.dtgConn.Location = New System.Drawing.Point(12, 3)
        Me.dtgConn.MultiSelect = False
        Me.dtgConn.Name = "dtgConn"
        Me.dtgConn.RowTemplate.Height = 24
        Me.dtgConn.Size = New System.Drawing.Size(774, 261)
        Me.dtgConn.TabIndex = 1
        '
        'CompCode
        '
        Me.CompCode.Frozen = True
        Me.CompCode.HeaderText = "公司別"
        Me.CompCode.MaxInputLength = 2
        Me.CompCode.Name = "CompCode"
        Me.CompCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CompCode.Width = 70
        '
        'ConnectString
        '
        Me.ConnectString.HeaderText = "連線字串"
        Me.ConnectString.Name = "ConnectString"
        Me.ConnectString.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ConnectString.Width = 800
        '
        'btnSave
        '
        Me.btnSave.Enabled = False
        Me.btnSave.Location = New System.Drawing.Point(115, 270)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "儲存檔案"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'frmConnPara
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(798, 299)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.dtgConn)
        Me.Controls.Add(Me.btnOpenFile)
        Me.Name = "frmConnPara"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "連線字串設定"
        CType(Me.dtgConn, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnOpenFile As System.Windows.Forms.Button
    Friend WithEvents dtgConn As System.Windows.Forms.DataGridView
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents CompCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ConnectString As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
