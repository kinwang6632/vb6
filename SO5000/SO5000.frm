VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSO5000 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '�S���ؽu
   Caption         =   "����޲z[SO5000]"
   ClientHeight    =   7185
   ClientLeft      =   75
   ClientTop       =   1590
   ClientWidth     =   11895
   Icon            =   "SO5000.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   11895
   WindowState     =   2  '�̤j��
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '������U��
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   6780
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
            Text            =   "����޲z[SO5000]"
            TextSave        =   "����޲z[SO5000]"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgPic 
      Height          =   4665
      Left            =   0
      Picture         =   "SO5000.frx":0442
      Top             =   0
      Width           =   6270
   End
   Begin VB.Menu mnuSO5100 
      Caption         =   "&1.�H���Z��"
      Begin VB.Menu mnuSO5110 
         Caption         =   "&1.�q�ܷ~���Z�Ĳέp"
      End
      Begin VB.Menu mnuSO5120 
         Caption         =   "&2.�޳N�H���Z�Ĳέp"
      End
      Begin VB.Menu mnuSO5130 
         Caption         =   "&3.�A�ȰϬ��u�Z�Ĳέp"
      End
      Begin VB.Menu mnuSO5140 
         Caption         =   "&4.���O�H�����p���R"
         Begin VB.Menu mnuSO5141 
            Caption         =   "&1.���O���X/�^��d��"
         End
         Begin VB.Menu mnuSO5142 
            Caption         =   "&2.���O�������d��"
         End
         Begin VB.Menu mnuSO5143 
            Caption         =   "&3.���O��������d��"
         End
         Begin VB.Menu mnuSO5144 
            Caption         =   "&4.���O�H���Z�Ĳέp"
         End
      End
      Begin VB.Menu mnuSO5150 
         Caption         =   "&5.�P���I�Z�Ĳέp"
      End
      Begin VB.Menu mnuSO5160 
         Caption         =   "&6.�~�ȤH���Z�Ĳέp"
      End
      Begin VB.Menu mnuSO5170 
         Caption         =   "&7.�Ȥᤶ���Z�Ĳέp"
      End
   End
   Begin VB.Menu mnuSO5200 
      Caption         =   "&2.���O���"
      Begin VB.Menu mnuSO5210 
         Caption         =   "&1.�C�를ú�O�Ȥ���Ӫ�"
      End
      Begin VB.Menu mnuSO5220 
         Caption         =   "&2.�C��U�����O�[�`�έp"
      End
      Begin VB.Menu mnuSO5230 
         Caption         =   "&3.�C�������P�ꦬ���`��"
      End
      Begin VB.Menu mnuSO5240 
         Caption         =   "&4.���O���B�δ��Ʋ��`��"
      End
      Begin VB.Menu mnuSO5250 
         Caption         =   "&5.�C�리�O�槹���v�έp"
      End
      Begin VB.Menu mnuSO5260 
         Caption         =   "&6.���O����Ȥ�Ʋέp��"
      End
      Begin VB.Menu mnuSO5270 
         Caption         =   "&7.�����b���Ӫ�(��D)"
      End
      Begin VB.Menu mnuSO5280 
         Caption         =   "&8.�����b���Ӫ�(�Ȥ�s��)"
      End
      Begin VB.Menu mnuSO5290 
         Caption         =   "&9.���O�沧�`��"
      End
   End
   Begin VB.Menu mnuSO5300 
      Caption         =   "&3.�����W�D"
      Begin VB.Menu mnuSO5310 
         Caption         =   "&1.�w�}�t�γ]�w�����@����"
      End
      Begin VB.Menu mnuSO5320 
         Caption         =   "&2.��X�W�D��Ʋέp"
      End
      Begin VB.Menu mnuSO5330 
         Caption         =   "&3.��X�W�D�Ȥ�@����"
      End
   End
   Begin VB.Menu mnuSO5400 
      Caption         =   "&4.�Ȥ�]��"
      Begin VB.Menu mnuSO5410 
         Caption         =   "&1.�Ȥ�]�ƥ[�ˤ@����"
      End
      Begin VB.Menu mnuSO5420 
         Caption         =   "&2.���ɨ���Ȥ�@����"
      End
      Begin VB.Menu mnuSO5430 
         Caption         =   "&3.�]�Ʃ�Ȥ�@����"
      End
   End
   Begin VB.Menu mnuSO5500 
      Caption         =   "&5.�Ȥ���"
      Begin VB.Menu mnuSO5510 
         Caption         =   "&1.�{���U���Ȥ�Ʋέp"
      End
      Begin VB.Menu mnuSO5520 
         Caption         =   "&2.�U��D�Ȥ᪬�A�έp"
      End
      Begin VB.Menu mnuSO5530 
         Caption         =   "&3.�Ȥ�򥻸�Ƥ@����"
      End
      Begin VB.Menu mnuSO5540 
         Caption         =   "&4.���P����Ȥ��Ƥ@����"
      End
      Begin VB.Menu mnuSO5550 
         Caption         =   "&5.�s�˾��Ȥ�έp��"
      End
      Begin VB.Menu mnuSO5560 
         Caption         =   "&6.�U�A�Ȱ�/��F�ϫȤ�W��Ʋέp"
      End
      Begin VB.Menu mnuSO5570 
         Caption         =   "&7.�L�g���ʦ��O���ثȤ����"
      End
      Begin VB.Menu mnuSO5580 
         Caption         =   "&8.�j�Ӥ�W��Ʋέp"
      End
      Begin VB.Menu mnuSO5590 
         Caption         =   "&9.�l�H�����ɮצC�L"
      End
   End
   Begin VB.Menu mnuSO5600 
      Caption         =   "&6.���u��"
      Begin VB.Menu mnuSO5610 
         Caption         =   "&1.�˾��έp����"
      End
      Begin VB.Menu mnuSO5620 
         Caption         =   "&2.���ײέp����"
      End
      Begin VB.Menu mnuSO5630 
         Caption         =   "&3.������έp����"
      End
      Begin VB.Menu mnuSO5640 
         Caption         =   "&4.�˾����z(�w��)�ε��ץ������έp"
      End
      Begin VB.Menu mnuSO5650 
         Caption         =   "&5.���ר��z(�w��)�ε��ץ������έp"
      End
      Begin VB.Menu mnuSO5660 
         Caption         =   "&6.��������z(�w��)�ε��ץ������έp"
      End
      Begin VB.Menu mnuSO5670 
         Caption         =   "&7.�U�����u�槹�u�`��"
      End
      Begin VB.Menu mnuSO5680 
         Caption         =   "&8.�C����z���u�����`��"
      End
   End
   Begin VB.Menu mnuSO5700 
      Caption         =   "&7.�U����]"
      Begin VB.Menu mnuSO5710 
         Caption         =   "&1.�˾����I�u��]�έp"
      End
      Begin VB.Menu mnuSO5720 
         Caption         =   "&2.���ץ��I�u��]�έp"
      End
      Begin VB.Menu mnuSO5730 
         Caption         =   "&3.��������I�u��]�έp"
      End
      Begin VB.Menu mnuSO5740 
         Caption         =   "&4.���ץӧi�άG�٭�]�έp"
      End
      Begin VB.Menu mnuSO5750 
         Caption         =   "&5.�������]�έp"
      End
      Begin VB.Menu mnuSO5760 
         Caption         =   "&6.��ú�O��]�έp"
      End
      Begin VB.Menu mnuSO5770 
         Caption         =   "&7.�C�餶�����O�έp"
      End
      Begin VB.Menu mnuSO5780 
         Caption         =   "&8.�Ȥ�q�ܥӧi��]�έp"
      End
      Begin VB.Menu mnuSO5790 
         Caption         =   "&9.�Ȥ᦬�O�u����]�έp"
      End
   End
   Begin VB.Menu mnuSO5800 
      Caption         =   "&8.�Ȥ�A��"
      Begin VB.Menu mnuSO5810 
         Caption         =   "&1.�Ȥᬣ�u�����έp"
      End
      Begin VB.Menu mnuSO5820 
         Caption         =   "&2.�Ȥ���׬����έp"
      End
   End
   Begin VB.Menu mnuSO5900 
      Caption         =   "&9.�l������"
      Begin VB.Menu mnuSO5910 
         Caption         =   "&1.��ú�O�Ȥ���ҦC�L"
      End
      Begin VB.Menu mnuSO5920 
         Caption         =   "&2.�{���Ȥ���ҦC�L"
      End
      Begin VB.Menu mnuSO5930 
         Caption         =   "&3.��@�Ȥ���ҦC�L"
      End
      Begin VB.Menu mnuSO5940 
         Caption         =   "&4.�w���O�Ȥ���ҦC�L"
      End
      Begin VB.Menu mnuSO5950 
         Caption         =   "&5.�P���I���ҦC�L"
      End
      Begin VB.Menu mnuSO5960 
         Caption         =   "&6.�Ȥᤶ�ЫȤ���ҦC�L"
      End
   End
   Begin VB.Menu mnuSO5A00 
      Caption         =   "&A.���U���R"
      Begin VB.Menu mnuSO5A10 
         Caption         =   "&1.�w�����J����"
      End
      Begin VB.Menu mnuSO5A20 
         Caption         =   "&2.�C��˾����u���O���B�έp"
      End
      Begin VB.Menu mnuSO5A30 
         Caption         =   "&3.�C�를ú�O���B�έp"
      End
      Begin VB.Menu mnuSO5A40 
         Caption         =   "&4.�C�리�O���B�έp"
      End
      Begin VB.Menu mnuSO5A50 
         Caption         =   "&5.�C�리�O���B�����k�ݲέp"
      End
      Begin VB.Menu mnuSO5A60 
         Caption         =   "&6.�~�צ��J���B�έp����"
      End
      Begin VB.Menu mnuSO5A70 
         Caption         =   "&7.�������J�έp����"
      End
   End
   Begin VB.Menu mnuSO5B00 
      Caption         =   "&B.����ˮ�"
   End
   Begin VB.Menu mnuSO5C00 
      Caption         =   "&C.PPV"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&0.����"
   End
End
Attribute VB_Name = "frmSO5000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GB As Object

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
   Call FunctionKey(KeyCode, Shift)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub FunctionKey(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Select Case KeyCode
           Case vbKeyEscape '   Esc
                If mnuExit.Enabled = True Then mnuExit_Click
    End Select
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub

Private Sub subSetPerv()
  On Error GoTo ChkErr
    'False  ������������
    mnuSO5110.Enabled = GetUserPriv("SO5100", ("SO5110"))
    mnuSO5120.Enabled = GetUserPriv("SO5100", ("SO5120"))
    mnuSO5130.Enabled = GetUserPriv("SO5100", ("SO5130"))
    mnuSO5140.Enabled = GetUserPriv("SO5100", ("SO5140"))
    mnuSO5141.Enabled = GetUserPriv("SO5140", ("SO5141"))
    mnuSO5142.Enabled = GetUserPriv("SO5140", ("SO5142"))
    mnuSO5143.Enabled = GetUserPriv("SO5140", ("SO5143"))
    mnuSO5144.Enabled = GetUserPriv("SO5140", ("SO5144"))
    mnuSO5150.Enabled = GetUserPriv("SO5100", ("SO5150")) And False
    mnuSO5160.Enabled = GetUserPriv("SO5100", ("SO5160")) And False
    mnuSO5170.Enabled = GetUserPriv("SO5100", ("SO5170")) And False
'    mnuSO5180.Enabled = GetUserPriv("SO5100", ("SO5180"))
'    mnuSO5190.Enabled = GetUserPriv("SO5100", ("SO5190"))
    
    mnuSO5210.Enabled = GetUserPriv("SO5200", ("SO5210"))
    mnuSO5220.Enabled = GetUserPriv("SO5200", ("SO5220"))
    mnuSO5230.Enabled = GetUserPriv("SO5200", ("SO5230"))
    mnuSO5240.Enabled = GetUserPriv("SO5200", ("SO5240"))
    mnuSO5250.Enabled = GetUserPriv("SO5200", ("SO5250")) And False
    mnuSO5260.Enabled = GetUserPriv("SO5200", ("SO5260")) And False
    mnuSO5270.Enabled = GetUserPriv("SO5200", ("SO5270")) And False
    mnuSO5280.Enabled = GetUserPriv("SO5200", ("SO5280")) And False
    mnuSO5290.Enabled = GetUserPriv("SO5200", ("SO5290")) And False
    
    mnuSO5310.Enabled = GetUserPriv("SO5300", ("SO5310"))
    mnuSO5320.Enabled = GetUserPriv("SO5300", ("SO5320"))
    mnuSO5330.Enabled = GetUserPriv("SO5300", ("SO5330"))
'    mnuSO5340.Enabled = GetUserPriv("SO5300", ("SO5340"))
'    mnuSO5350.Enabled = GetUserPriv("SO5300", ("SO5350"))
'    mnuSO5360.Enabled = GetUserPriv("SO5300", ("SO5360"))
'    mnuSO5370.Enabled = GetUserPriv("SO5300", ("SO5370"))
'    mnuSO5380.Enabled = GetUserPriv("SO5300", ("SO5380"))
'    mnuSO5390.Enabled = GetUserPriv("SO5300", ("SO5390"))
    
    mnuSO5410.Enabled = GetUserPriv("SO5400", ("SO5410"))
    mnuSO5420.Enabled = GetUserPriv("SO5400", ("SO5420"))
    mnuSO5430.Enabled = GetUserPriv("SO5400", ("SO5430"))
'    mnuSO5440.Enabled = GetUserPriv("SO5400", ("SO5440"))
'    mnuSO5450.Enabled = GetUserPriv("SO5400", ("SO5450"))
'    mnuSO5460.Enabled = GetUserPriv("SO5400", ("SO5460"))
'    mnuSO5470.Enabled = GetUserPriv("SO5400", ("SO5470"))
'    mnuSO5480.Enabled = GetUserPriv("SO5400", ("SO5480"))
'    mnuSO5490.Enabled = GetUserPriv("SO5400", ("SO5490"))
    
    mnuSO5510.Enabled = GetUserPriv("SO5500", ("SO5510")) And False
    mnuSO5520.Enabled = GetUserPriv("SO5500", ("SO5520"))
    mnuSO5530.Enabled = GetUserPriv("SO5500", ("SO5530"))
    mnuSO5540.Enabled = GetUserPriv("SO5500", ("SO5540"))
    mnuSO5550.Enabled = GetUserPriv("SO5500", ("SO5550"))
    mnuSO5560.Enabled = GetUserPriv("SO5500", ("SO5560")) And False
    mnuSO5570.Enabled = GetUserPriv("SO5500", ("SO5570")) And False
    mnuSO5580.Enabled = GetUserPriv("SO5500", ("SO5580")) And False
    mnuSO5590.Enabled = GetUserPriv("SO5500", ("SO5590")) And False
    
    mnuSO5610.Enabled = GetUserPriv("SO5600", ("SO5610"))
    mnuSO5620.Enabled = GetUserPriv("SO5600", ("SO5620"))
    mnuSO5630.Enabled = GetUserPriv("SO5600", ("SO5630"))
    mnuSO5640.Enabled = GetUserPriv("SO5600", ("SO5640"))
    mnuSO5650.Enabled = GetUserPriv("SO5600", ("SO5650"))
    mnuSO5660.Enabled = GetUserPriv("SO5600", ("SO5660"))
    mnuSO5670.Enabled = GetUserPriv("SO5600", ("SO5670"))
    mnuSO5680.Enabled = GetUserPriv("SO5600", ("SO5680"))
'    mnuSO5690.Enabled = GetUserPriv("SO5600", ("SO5690"))
    
    mnuSO5710.Enabled = GetUserPriv("SO5700", ("SO5710"))
    mnuSO5720.Enabled = GetUserPriv("SO5700", ("SO5720"))
    mnuSO5730.Enabled = GetUserPriv("SO5700", ("SO5730"))
    mnuSO5740.Enabled = GetUserPriv("SO5700", ("SO5740"))
    mnuSO5750.Enabled = GetUserPriv("SO5700", ("SO5750"))
    mnuSO5760.Enabled = GetUserPriv("SO5700", ("SO5760"))
    mnuSO5770.Enabled = GetUserPriv("SO5700", ("SO5770"))
    mnuSO5780.Enabled = GetUserPriv("SO5700", ("SO5780"))
    mnuSO5790.Enabled = GetUserPriv("SO5700", ("SO5790"))
    
    mnuSO5810.Enabled = GetUserPriv("SO5800", ("SO5810"))
    mnuSO5820.Enabled = GetUserPriv("SO5800", ("SO5820"))
'    mnuSO5830.Enabled = GetUserPriv("SO5800", ("SO5830"))
'    mnuSO5840.Enabled = GetUserPriv("SO5800", ("SO5840"))
'    mnuSO5850.Enabled = GetUserPriv("SO5800", ("SO5850"))
'    mnuSO5860.Enabled = GetUserPriv("SO5800", ("SO5860"))
'    mnuSO5870.Enabled = GetUserPriv("SO5800", ("SO5870"))
'    mnuSO5880.Enabled = GetUserPriv("SO5800", ("SO5880"))
'    mnuSO5890.Enabled = GetUserPriv("SO5800", ("SO5890"))
    
    mnuSO5910.Enabled = GetUserPriv("SO5900", ("SO5910")) And False
    mnuSO5920.Enabled = GetUserPriv("SO5900", ("SO5920")) And False
    mnuSO5930.Enabled = GetUserPriv("SO5900", ("SO5930")) And False
    mnuSO5940.Enabled = GetUserPriv("SO5900", ("SO5940")) And False
    mnuSO5950.Enabled = GetUserPriv("SO5900", ("SO5950")) And False
    mnuSO5960.Enabled = GetUserPriv("SO5900", ("SO5960")) And False
'    mnuSO5970.Enabled = GetUserPriv("SO5900", ("SO5970"))
'    mnuSO5980.Enabled = GetUserPriv("SO5900", ("SO5980"))
'    mnuSO5990.Enabled = GetUserPriv("SO5900", ("SO5990"))
    
    mnuSO5A10.Enabled = GetUserPriv("SO5A00", ("SO5A10")) And False
    mnuSO5A20.Enabled = GetUserPriv("SO5A00", ("SO5A20")) And False
    mnuSO5A30.Enabled = GetUserPriv("SO5A00", ("SO5A30")) And False
    mnuSO5A40.Enabled = GetUserPriv("SO5A00", ("SO5A40")) And False
    mnuSO5A50.Enabled = GetUserPriv("SO5A00", ("SO5A50")) And False
    mnuSO5A60.Enabled = GetUserPriv("SO5A00", ("SO5A60")) And False
    mnuSO5A70.Enabled = GetUserPriv("SO5A00", ("SO5A70")) And False
'    mnuSO5A80.Enabled = GetUserPriv("SO5A00", ("SO5A80"))
'    mnuSO5A90.Enabled = GetUserPriv("SO5A00", ("SO5A90"))
    
'    mnuSO5B10.Enabled = GetUserPriv("SO5B00", ("SO5B10"))
'    mnuSO5B20.Enabled = GetUserPriv("SO5B00", ("SO5B20"))
'    mnuSO5B30.Enabled = GetUserPriv("SO5B00", ("SO5B30"))
'    mnuSO5B40.Enabled = GetUserPriv("SO5B00", ("SO5B40"))
'    mnuSO5B50.Enabled = GetUserPriv("SO5B00", ("SO5B50"))
'    mnuSO5B60.Enabled = GetUserPriv("SO5B00", ("SO5B60"))
'    mnuSO5B70.Enabled = GetUserPriv("SO5B00", ("SO5B70"))
'    mnuSO5B80.Enabled = GetUserPriv("SO5B00", ("SO5B80"))
'    mnuSO5B90.Enabled = GetUserPriv("SO5B00", ("SO5B90"))
    
'    mnuSO5C10.Enabled = GetUserPriv("SO5C00", ("SO5C10"))
'    mnuSO5C20.Enabled = GetUserPriv("SO5C00", ("SO5C20"))
'    mnuSO5C30.Enabled = GetUserPriv("SO5C00", ("SO5C30"))
'    mnuSO5C40.Enabled = GetUserPriv("SO5C00", ("SO5C40"))
'    mnuSO5C50.Enabled = GetUserPriv("SO5C00", ("SO5C50"))
'    mnuSO5C60.Enabled = GetUserPriv("SO5C00", ("SO5C60"))
'    mnuSO5C70.Enabled = GetUserPriv("SO5C00", ("SO5C70"))
'    mnuSO5C80.Enabled = GetUserPriv("SO5C00", ("SO5C80"))
'    mnuSO5C90.Enabled = GetUserPriv("SO5C00", ("SO5C90"))
    
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subSetPerv")
End Sub

Private Sub Form_Load()
  On Error Resume Next
  Set GB = CreateObject("GICMIS_BOJ.OBJ")
  GB.FormGicmis.Hide
  Set GB.Form5000 = Me
  On Error GoTo ChkErr
   Call GetGlobal
   Call subSetPerv
   Set cnn = GetTmpMdbCn
   If garyGi(4) = 0 Then Exit Sub
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_Resize()
  On Error Resume Next
   imgPic.Move (Me.Width - imgPic.Width) / 2, (Me.Height - imgPic.Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
   Set GB.Form5000 = Nothing
   GB.FormGicmis.Show
   Set cnn = Nothing
End Sub

Private Sub mnuExit_Click()
  On Error Resume Next
   Unload Me
End Sub

Private Sub mnuSO5110_Click()
  On Error Resume Next
    frmSO5110A.Show vbModal
End Sub

Private Sub mnuSO5120_Click()
  On Error Resume Next
    frmSO5120A.Show vbModal
End Sub

Private Sub mnuSO5130_Click()
  On Error Resume Next
    frmSO5130A.Show vbModal
End Sub

Private Sub mnuSO5141_Click()
  On Error Resume Next
    frmSO5141A.Show vbModal
End Sub

Private Sub mnuSO5142_Click()
  On Error Resume Next
    frmSO5142A.Show vbModal
End Sub

Private Sub mnuSO5143_Click()
  On Error Resume Next
    frmSO5143A.Show vbModal
End Sub

Private Sub mnuSO5144_Click()
  On Error Resume Next
   frmSO5144A.Show vbModal
End Sub

Private Sub mnuSO5210_Click()
  On Error Resume Next
   frmSO5210A.Show vbModal
End Sub

Private Sub mnuSO5220_Click()
  On Error Resume Next
   frmSO5220A.Show vbModal
End Sub

Private Sub mnuSO5230_Click()
  On Error Resume Next
   frmSO5230A.Show vbModal
End Sub

Private Sub mnuSO5240_Click()
  On Error Resume Next
   frmSO5240A.Show vbModal
End Sub

Private Sub mnuSO5310_Click()
  On Error Resume Next
   frmSO5310A.Show vbModal
End Sub

Private Sub mnuSO5320_Click()
  On Error Resume Next
   frmSO5320A.Show vbModal
End Sub

Private Sub mnuSO5330_Click()
  On Error Resume Next
   frmSO5330A.Show vbModal
End Sub

Private Sub mnuSO5410_Click()
  On Error Resume Next
   frmSO5410A.Show vbModal
End Sub

Private Sub mnuSO5420_Click()
  On Error Resume Next
   frmSO5420A.Show vbModal
End Sub

Private Sub mnuSO5430_Click()
  On Error Resume Next
   frmSO5430A.Show vbModal
End Sub

Private Sub mnuSO5520_Click()
  On Error Resume Next
   frmSO5520A.Show vbModal
End Sub

Private Sub mnuSO5530_Click()
  On Error Resume Next
   frmSO5530A.Show vbModal
End Sub

Private Sub mnuSO5540_Click()
  On Error Resume Next
   frmSO5540A.Show vbModal
End Sub

Private Sub mnuSO5550_Click()
  On Error Resume Next
   frmSO5550A.Show vbModal
End Sub

Private Sub mnuSO5610_Click()
  On Error Resume Next
   frmSO5610A.Show vbModal
End Sub

Private Sub mnuSO5620_Click()
  On Error Resume Next
   frmSO5620A.Show vbModal
End Sub

Private Sub mnuSO5630_Click()
  On Error Resume Next
   frmSO5630A.Show vbModal
End Sub

Private Sub mnuSO5640_Click()
  On Error Resume Next
   frmSO5640A.Show vbModal
End Sub

Private Sub mnuSO5650_Click()
  On Error Resume Next
   frmSO5650A.Show vbModal
End Sub

Private Sub mnuSO5660_Click()
  On Error Resume Next
   frmSO5660A.Show vbModal
End Sub

Private Sub mnuSO5710_Click()
  On Error Resume Next
   frmSO5710A.Show vbModal
End Sub

Private Sub mnuSO5720_Click()
  On Error Resume Next
   frmSO5720A.Show vbModal
End Sub

Private Sub mnuSO5730_Click()
  On Error Resume Next
   frmSO5730A.Show vbModal
End Sub

Private Sub mnuSO5740_Click()
  On Error Resume Next
   frmSO5740A.Show vbModal
End Sub

Private Sub mnuSO5750_Click()
  On Error Resume Next
   frmSO5750A.Show vbModal
End Sub

Private Sub mnuSO5760_Click()
  On Error Resume Next
   frmSO5760A.Show vbModal
End Sub

Private Sub mnuSO5770_Click()
  On Error Resume Next
   frmSO5770A.Show vbModal
End Sub

Private Sub mnuSO5780_Click()
  On Error Resume Next
   frmSO5780A.Show vbModal
End Sub

Private Sub mnuSO5790_Click()
  On Error Resume Next
   frmSO5790A.Show vbModal
End Sub

Private Sub mnuSO5810_Click()
  On Error Resume Next
   frmSO5810A.Show vbModal
End Sub

Private Sub mnuSO5820_Click()
  On Error Resume Next
   frmSO5820A.Show vbModal
End Sub

Private Sub mnuSO5B00_Click()
  On Error Resume Next
   MsgBox "�ȵL�W��"
End Sub

Private Sub mnuSO5C00_Click()
  On Error Resume Next
   MsgBox "�ȵL�W��"

End Sub
