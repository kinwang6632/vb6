<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmShowSQL
    Inherits DevExpress.XtraEditors.XtraForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmShowSQL))
        Me.edtShowSQL = New DevExpress.XtraEditors.MemoEdit
        CType(Me.edtShowSQL.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'edtShowSQL
        '
        Me.edtShowSQL.Dock = System.Windows.Forms.DockStyle.Fill
        Me.edtShowSQL.Location = New System.Drawing.Point(0, 0)
        Me.edtShowSQL.Name = "edtShowSQL"
        Me.edtShowSQL.Properties.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.edtShowSQL.Properties.Appearance.Options.UseFont = True
        Me.edtShowSQL.Properties.ReadOnly = True
        Me.edtShowSQL.Properties.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.edtShowSQL.Size = New System.Drawing.Size(808, 334)
        Me.edtShowSQL.TabIndex = 0
        '
        'frmShowSQL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(808, 334)
        Me.Controls.Add(Me.edtShowSQL)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmShowSQL"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SQL語法"
        CType(Me.edtShowSQL.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents edtShowSQL As DevExpress.XtraEditors.MemoEdit
End Class
