Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form �]�p�u�㲣�ͪ��{���X "

    Public Sub New()
        MyBase.New()

        '���� Windows Form �]�p�u��һݪ��I�s�C
        InitializeComponent()

        '�b InitializeComponent() �I�s����[�J�Ҧ�����l�]�w

    End Sub

    'Form �мg Dispose �H�M������M��C
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    '�� Windows Form �]�p�u�㪺���n��
    Private components As System.ComponentModel.IContainer

    '�`�N: �H�U�� Windows Form �]�p�u��һݪ��{��
    '�z�i�H�ϥ� Windows Form �]�p�u��i��ק�C
    '�ФŨϥε{���X�s�边�ӭק�o�ǵ{�ǡC
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(286, 394)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Button1"
        '
        'DataGrid1
        '
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(6, 4)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(644, 380)
        Me.DataGrid1.TabIndex = 1
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 15)
        Me.ClientSize = New System.Drawing.Size(654, 427)
        Me.Controls.Add(Me.DataGrid1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ds As New DataSet
        ds.ReadXml("C:\Gird\V400\Hammer\CM CT �M��\CS_Web_Command_Matrix\XML Data\ClearCPE.xml")
        DataGrid1.DataSource = ds.Tables(0)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

    End Sub
End Class
