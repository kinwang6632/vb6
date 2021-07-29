<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class XfrmAbout
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(XfrmAbout))
        Me.grpCtl = New DevExpress.XtraEditors.GroupControl
        Me.lblCtlTitle = New DevExpress.XtraEditors.LabelControl
        Me.hpLnkCS = New DevExpress.XtraEditors.HyperLinkEdit
        Me.lblCopyright = New DevExpress.XtraEditors.LabelControl
        Me.picEdtCS = New DevExpress.XtraEditors.PictureEdit
        CType(Me.grpCtl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpCtl.SuspendLayout()
        CType(Me.hpLnkCS.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picEdtCS.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpCtl
        '
        Me.grpCtl.Appearance.ForeColor = System.Drawing.Color.Blue
        Me.grpCtl.Appearance.Options.UseForeColor = True
        Me.grpCtl.AppearanceCaption.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpCtl.AppearanceCaption.Options.UseFont = True
        Me.grpCtl.CaptionImage = CType(resources.GetObject("grpCtl.CaptionImage"), System.Drawing.Image)
        Me.grpCtl.Controls.Add(Me.lblCtlTitle)
        Me.grpCtl.Controls.Add(Me.hpLnkCS)
        Me.grpCtl.Controls.Add(Me.lblCopyright)
        Me.grpCtl.Controls.Add(Me.picEdtCS)
        Me.grpCtl.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpCtl.Location = New System.Drawing.Point(0, 0)
        Me.grpCtl.Name = "grpCtl"
        Me.grpCtl.Size = New System.Drawing.Size(427, 200)
        Me.grpCtl.TabIndex = 1
        Me.grpCtl.Text = " Cablesoft System "
        '
        'lblCtlTitle
        '
        Me.lblCtlTitle.Appearance.Font = New System.Drawing.Font("Tahoma", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCtlTitle.Appearance.Options.UseFont = True
        Me.lblCtlTitle.Location = New System.Drawing.Point(188, 57)
        Me.lblCtlTitle.Name = "lblCtlTitle"
        Me.lblCtlTitle.Size = New System.Drawing.Size(213, 38)
        Me.lblCtlTitle.TabIndex = 9
        Me.lblCtlTitle.Text = "PVOD - Socket" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "               Gateway System"
        '
        'hpLnkCS
        '
        Me.hpLnkCS.EditValue = "www.cablesoft.com.tw"
        Me.hpLnkCS.Location = New System.Drawing.Point(21, 118)
        Me.hpLnkCS.Name = "hpLnkCS"
        Me.hpLnkCS.Properties.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.hpLnkCS.Properties.Appearance.Options.UseFont = True
        Me.hpLnkCS.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.hpLnkCS.Properties.Caption = "Cablesoft Tech Inc. On the Web"
        Me.hpLnkCS.Properties.Image = CType(resources.GetObject("hpLnkCS.Properties.Image"), System.Drawing.Image)
        Me.hpLnkCS.Size = New System.Drawing.Size(257, 22)
        Me.hpLnkCS.TabIndex = 8
        Me.hpLnkCS.ToolTip = " 連 結 到 開 博 科 技 網 站 "
        '
        'lblCopyright
        '
        Me.lblCopyright.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblCopyright.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCopyright.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblCopyright.Appearance.Options.UseFont = True
        Me.lblCopyright.Appearance.Options.UseForeColor = True
        Me.lblCopyright.Location = New System.Drawing.Point(23, 142)
        Me.lblCopyright.Name = "lblCopyright"
        Me.lblCopyright.Size = New System.Drawing.Size(378, 16)
        Me.lblCopyright.TabIndex = 7
        Me.lblCopyright.Text = "Copyright 1992 - 2007 Cablesoft Tech Inc. ALL RIGHTS RESERVED"
        '
        'picEdtCS
        '
        Me.picEdtCS.EditValue = CType(resources.GetObject("picEdtCS.EditValue"), Object)
        Me.picEdtCS.Location = New System.Drawing.Point(21, 29)
        Me.picEdtCS.Name = "picEdtCS"
        Me.picEdtCS.Properties.SizeMode = DevExpress.XtraEditors.Controls.PictureSizeMode.Stretch
        Me.picEdtCS.Size = New System.Drawing.Size(160, 84)
        Me.picEdtCS.TabIndex = 6
        '
        'XfrmAbout
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(427, 200)
        Me.ControlBox = False
        Me.Controls.Add(Me.grpCtl)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Name = "XfrmAbout"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "XfrmAbout"
        CType(Me.grpCtl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpCtl.ResumeLayout(False)
        Me.grpCtl.PerformLayout()
        CType(Me.hpLnkCS.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picEdtCS.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grpCtl As DevExpress.XtraEditors.GroupControl
    Private WithEvents lblCtlTitle As DevExpress.XtraEditors.LabelControl
    Private WithEvents hpLnkCS As DevExpress.XtraEditors.HyperLinkEdit
    Private WithEvents lblCopyright As DevExpress.XtraEditors.LabelControl
    Private WithEvents picEdtCS As DevExpress.XtraEditors.PictureEdit
End Class
