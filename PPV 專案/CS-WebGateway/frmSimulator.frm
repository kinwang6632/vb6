VERSION 5.00
Object = "{898CF0BC-AF45-45AC-9831-A27FC475455D}#1.0#0"; "WinXPctl.ocx"
Begin VB.Form frmSimulator 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   " SMS Web Gateway Simulator  .."
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSimulator.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   10290
   StartUpPosition =   2  '�ù�����
   Begin VB.ComboBox cboOpt2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmSimulator.frx":0442
      Left            =   2070
      List            =   "frmSimulator.frx":044C
      TabIndex        =   24
      Text            =   "�q�� PPV �{��"
      Top             =   645
      Width           =   1935
   End
   Begin WinXP_Engine.WinXPctl xpc 
      Left            =   3510
      Top             =   -420
      _ExtentX        =   3995
      _ExtentY        =   873
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit (&X)"
      Height          =   375
      Left            =   6420
      TabIndex        =   22
      Top             =   8895
      Width           =   1335
   End
   Begin VB.Frame fraOpt 
      BackColor       =   &H00E0E0E0&
      Caption         =   " �� �e �R �O "
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8655
      Left            =   150
      TabIndex        =   23
      Top             =   150
      Width           =   9795
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 28.ESERVICE�^���פu��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   225
         Index           =   27
         Left            =   5550
         TabIndex        =   34
         Top             =   7650
         Width           =   2925
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 27.ESERVICE�^�������"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   225
         Index           =   26
         Left            =   5550
         TabIndex        =   33
         Top             =   7380
         Width           =   2925
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 26.ESERVICE�^�������"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   225
         Index           =   25
         Left            =   5550
         TabIndex        =   32
         Top             =   7110
         Width           =   2805
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 25.ESERVICE�^�˾���"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   225
         Index           =   24
         Left            =   5550
         TabIndex        =   31
         Top             =   6840
         Width           =   2295
      End
      Begin VB.ComboBox cboOpt24 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmSimulator.frx":0471
         Left            =   3060
         List            =   "frmSimulator.frx":047B
         TabIndex        =   30
         Text            =   "�]�ƺ���=1.MAC"
         Top             =   7590
         Width           =   1935
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 24. ETHOME-CM������T�d��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   225
         Index           =   23
         Left            =   150
         TabIndex        =   29
         Top             =   7650
         Width           =   4965
      End
      Begin VB.ComboBox cboOpt22 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmSimulator.frx":04A0
         Left            =   3060
         List            =   "frmSimulator.frx":04AA
         TabIndex        =   28
         Text            =   "�d�b�ȸ�T"
         Top             =   6870
         Width           =   1935
      End
      Begin VB.ComboBox cboOpt23 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmSimulator.frx":04C6
         Left            =   3060
         List            =   "frmSimulator.frx":04D0
         TabIndex        =   27
         Text            =   "�d�b�ȸ�T"
         Top             =   7230
         Width           =   1935
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 23. CMCP������T�d��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   22
         Left            =   150
         TabIndex        =   26
         Top             =   7324
         Width           =   5025
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 22. CATV&DSTB������T�d��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   21
         Left            =   150
         TabIndex        =   25
         Top             =   7002
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 21. �g��ϴM��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   20
         Left            =   150
         TabIndex        =   20
         Top             =   6680
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 20. ��O�Ȥ�W��d��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   19
         Left            =   150
         TabIndex        =   19
         Top             =   6358
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 19. �Ǧ^SMS�|����� ( For web update .. )"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   18
         Left            =   150
         TabIndex        =   18
         Top             =   6036
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 13. �p���I�O�`�ب���(�ݸg�LSMS�q�ʱK�X�{��,�å]�ACA�t�έq�ʱ��v�@�~�γB�z���G)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   12
         Left            =   135
         TabIndex        =   12
         Top             =   4104
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 14. �p���I�O�`�ب�����]���"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   13
         Left            =   135
         TabIndex        =   13
         Top             =   4426
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00C0C0C0&
         Caption         =   " 15. ���u�q���Ȥ�򥻸�Ƭd��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   225
         Index           =   14
         Left            =   135
         TabIndex        =   14
         Top             =   4748
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00C0C0C0&
         Caption         =   " 16. ���u�Ȥ�򥻸�ƺ��@(�ּư򥻸�����)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   225
         Index           =   15
         Left            =   135
         TabIndex        =   15
         ToolTipText     =   " �d��(ppv,ippv) [���ѭӤH�P�޲z��] "
         Top             =   5070
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 17. �Ȥ�w��STB�]�Ƹ�Ƭd��(�w�w��STB����,�s���Φ�m)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   16
         Left            =   135
         TabIndex        =   16
         Top             =   5392
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 18. Ū���I�Ʀ��O���"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   17
         Left            =   135
         TabIndex        =   17
         Top             =   5714
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 7. �Ȥᦳ�u�q�������O��b�@�~/�p���I�O�w�I���B��b�@�~(�H�Υd�p�B�I�ھ���)--�u�W�I�ڱ��v"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   6
         Left            =   135
         TabIndex        =   6
         Top             =   2172
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 8. �I�ƽu�W�q��--�u�W�I�ڱ��v"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   7
         Left            =   135
         TabIndex        =   7
         Top             =   2494
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 9. �I�ƽu�W�q��--�U���I��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   8
         Left            =   135
         TabIndex        =   8
         Top             =   2816
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 10. �Ȥ��ܧ�STB�|�X�ƦrParental PIN Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   9
         Left            =   135
         TabIndex        =   9
         ToolTipText     =   " �d��(ppv,ippv) [���ѭӤH�P�޲z��] "
         Top             =   3138
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 11. �Ȥ� iPPV�q��PIN Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   10
         Left            =   135
         TabIndex        =   10
         Top             =   3460
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 12. �p���I�O�`�حq��(�ݸg�LSMS�q�ʱK�X�{��,�å]�ACA�t�έq�ʱ��v�@�~�γB�z���G)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   11
         Left            =   135
         TabIndex        =   11
         Top             =   3782
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 6. �Ȥᦳ�u�q�������O��b�@�~/�p���I�O�w�I���B��b�@�~(�H�Υd�p�B�I�ھ���)--�u�W�I�ڱ��v--���d��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   5
         Left            =   135
         TabIndex        =   5
         Top             =   1850
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 5. �Ȥ�b�ȸ�Ƭd��(����϶����o�����b���ú�ڰO��)(�ӤH)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   4
         Left            =   135
         TabIndex        =   4
         Top             =   1528
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 4. �Ȥ�q�ʭp���I�O�`�ة��Ӭd��(ppv,ippv) [���ѭӤH�P�޲z��] "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   3
         Left            =   135
         TabIndex        =   3
         ToolTipText     =   " �d��(ppv,ippv) [���ѭӤH�P�޲z��] "
         Top             =   1206
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00C0C0FF&
         Caption         =   " 3. �p���I�O�`�ؤ��e��� "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   2
         Top             =   884
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 2. �q�� PPV �{�� "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   1
         Top             =   562
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 1. ���u�q���Ȥᨭ�� SMS �b���K�X�{�ҥӽ� "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   9500
      End
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Testing (&E)"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   21
      Top             =   8895
      Width           =   1335
   End
End
Attribute VB_Name = "frmSimulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboOpt2_GotFocus()
    opt(1).Value = True
End Sub

Private Sub cboOpt23_GotFocus()
    opt(22).Value = True
End Sub

Private Sub Form_Load()
    With xpc
        .TextControl = False
        .OptionControl = False
        .ColorScheme = System
        .InitSubClassing
    End With
    cboOpt2.ListIndex = 0
    cboOpt22.ListIndex = 0
    cboOpt23.ListIndex = 0
End Sub

Public Function StartUp(ByVal strPara As String) As String
    StartUp = CreateObject("CS_Web.Gateway").Executing(strPara)
End Function

Private Sub cmdTest_Click()
    
    '1,2,19
    
    Dim bytFuncID As Byte
    Dim strResult As String
    bytFuncID = Val(GetOpt) + 1
    strResult = CallByName(Me, "Func" & bytFuncID, VbMethod)
    Debug.Print strResult
    MsgBox strResult, vbInformation, "Message"
    
End Sub

Public Function Func1() As String ' ���u�q���Ȥᨭ��SMS�b���K�X�{�ҥӽ� ( �R�O 1 )
'   1�B���q�O�B�Ȥ�s���BSTB�Ǹ�
'   �|���N���B�|���W�١B�����Ҧr���B�X�ͦ~���B�p���q�ܡB
'   ������X�B�ʧO�B�p���a�}�BEmail�BPPV��I�q�ʱK�X�B�n�J�K�X

'   1�B���q�O�B�Ȥ�s���BSTB�Ǹ�
'   �s���Ѽ� : �ȪA�����|���N���B�|���W�١B�����Ҧr���B�X�ͤ��-�~�B�X�ͤ��-��B�X�ͤ��-��B
'   �p���q��(�v) �B�p���q��(��)�B��ʹq�ܡB�ʧO�B�q�T�a�}-�����B�q�T�a�}-�m���ϡB�q�T�a�}-���q��ѧ˫Ѹ��ӡB
'   Email�B�ȪA�����n�J�b���BPPV�q�ʱK�X�B�ȪA�����|���K�X�B�K�X���ܰ��D�B�K�X���ܵ��סB�Ǿ��B
'   �@�N���즳�u�q���A�Ȥ��W�D�`�ص������T���B�B�ê��p�B�P��H�ơB�P�����B¾�~�O�B�ӤH�~���J�B
'   �����O�B�b�����A�B�ӽФ���ɶ��B�ҥΤ���B�̫�n�J����B�̫Ყ�ʤ���B�t�Υx�N�X

'1�B���q�O�B�Ȥ�s��(���)
'C�B�Τ�m�W�B�q��(���)
'D�BDSTB�]�ƧǸ��B���z�d�d��(���)
'I�B�ӽФH�B�����ҡB�ͤ�(���)
'�K�K

'1�B���q�O�B�Ȥ�s���B�����B�{��ID (�����=0�ɻ{��ID������)(���)
'C�B�Τ�m�W�B�q��(���)
'D�BDSTB�]�ƧǸ��B���z�d�d��(���)
'I�B�ӽФH�B�����ҡB�ͤ�(���)
'�K�K

'    Func1 = StartUp("1,1,12,0,0010000006" & vbCrLf & _
                                "C,�Τ�m�W,�q��(���)" & vbCrLf & _
                                "D,DSTB�]�ƧǸ�,���z�d�d��(���)" & vbCrLf & _
                                "I,�ӽФH,������,�ͤ�(���)")

'    Func1 = StartUp("1,1,12,0,0010000006,200605080002,test,,2005,08,04,,,,�k,�x����,NO,��90��86��,winnie@eland.com.tw,null,,test,test,test,,,,,,��T/�����,30�U�H�U,1,Y,20050804,20050804,20050804,20050804,16,,,," & vbCrLf & _
                                "C,�Ȥ�W��,4254389" & vbCrLf & _
                                "D,38383388,22660066" & vbCrLf & _
                                "I,�S���H�ӽ�,V111111111,19750922")
                                
'    Func1 = StartUp("1,1,12,1,,200605080002,test,A123456789,2005,08,04,,,,�k,�x����,NO,��90��86��,winnie@123,null,888,test,test,test,,,,,,��T/�����,30�U�H�U,1,Y,20060509,20060509,20060509,20060509,16,,,," & vbCrLf & _
                                "D,119119374855,116021893636")
                                
'    Func1 = StartUp("1,1,12,1,,200605080002,���ߵ�,,2006,05,17,,,,�k,,NO,A����90��86��,winnie@eland.com.tw,,winnie,winnie,winnie,winnie,,,,,,��T/�����,30�U�H�U,1,Y,20060517,20060517,20060517,20060517,C,�o�����Y,�o������,�o���a�},�Τ@�s��" & vbCrLf & _
                                "D,119120020205,116025519411" & vbCrLf & _
                                "C,���ߵ�,987654321")
                                
'    Func1 = StartUp("1,16,1980,1,,200605080003,GAGA,A123456789,2005,08,04,,,,�k,�x����,NO,��90��86��,winnie@eland.com.tw,null,888,test,test,test,,,,,,��T/�����,30�U�H�U,1,Y,20060509,20060509,20060509,20060509,16,,,," & vbCrLf & _
                                "C,�Ȥ�W��,1980" & vbCrLf & _
                                "D,200411250066126,119119589500,116019114162")

    Func1 = StartUp("1,1,12,0,0010000006," & _
                                "MemberID,Hammer,V111111111," & _
                                "1975,09,22,05-3838338,07-5561926,0938123456,�k," & _
                                "������,�����,�շR�G��198��11�Ӥ�1," & _
                                "PowerHammer@msn.com,LoginID," & _
                                "PPVOrderPwd,LoginPassword,HintName,HintAnswer," & _
                                "���X��,1,���B,200,�@��H,�ۥѷ~,200�U,�S����,����," & _
                                "2005/05/01,2005/05/06,2005/05/20,2005/05/20,1," & _
                                "�o�����Y,�o������,�o���a�},�Τ@�s��,'000CE5DDF72C'" & _
                                "C,�Ȥ�W��,4254389" & vbCrLf & _
                                "D,38383388,22660066" & vbCrLf & _
                                "I,�S���H�ӽ�,V111111111,19750922,3388")

                                
'    Func1 = StartUp("1,1,12,000CE5DDF72C," & _
                            "MemberID,Hammer,V111111111," & _
                            "1975,09,22,05-3838338,07-5561926,0938123456,�k," & _
                            "������,�����,�շR�G��198��11�Ӥ�1," & _
                            "PowerHammer@msn.com,LoginID," & _
                            "PPVOrderPwd,LoginPassword,HintName,HintAnswer," & _
                            "���X��,1,���B,200,�@��H,�ۥѷ~,200�U,�S����,����," & _
                            "2005/05/01,2005/05/06,2005/05/20,2005/05/20,1")

'    Func1 = StartUp("1,16,84,119120026164," & _
                                "0000001,�|���@,A123456789," & _
                                "2005,07,20,02-1234556,02-1235679,0931234890,M," & _
                                "A��,,�Ĥ@��182��4��2��3��," & _
                                "abcd@test.com.tw,callcenter," & _
                                "ppvpwd,custpwd,�K�X���D,�K�X���ܵ���," & _
                                ",,,,,,,1,Y,20050720,20050720,20050720,20050720,16")
    
'    Func1 = StartUp("1,1,12,000CE5DDF72C," & _
                            "HaHa66,Hammer,V000000000," & _
                            "1975/09/22,07-5561926,0938123456,1," & _
                            "����������ϳշR�G��198��11�Ӥ�1," & _
                            "PowerHammer@msn.com,3388,2266")
    
'    Func1 = CreateObject("CS_MT.Birdge").StartUp("1,1,12,000CE5DDF72C," & _
'                            "HaHa66,Hammer,V000000000," & _
'                            "1975/09/22,07-5561926,0938123456,1," & _
'                            "����������ϳշR�G��198��11�Ӥ�1," & _
'                            "PowerHammer@msn.com,3388,2266")
    
'    Func1 = CreateObject("CS_Web.GateWay").Executing("1,1,12,000CE5DDF72C," & _
                            "HaHa66,Hammer,V000000000," & _
                            "1975/09/22,07-5561926,0938123456,1," & _
                            "����������ϳշR�G��198��11�Ӥ�1," & _
                            "PowerHammer@msn.com,3388,2266")
End Function

Public Function Func2() As String ' �q��PPV�{�� ( �R�O 2 )
'   2�B�|���N���B�{��ID�BPPV��I�q�ʱK�X

    '   �°Ѽ� : �|���W�١B�����Ҧr���B�X�ͦ~���B�p���q�ܡB������X�B�ʧO�B�p���a�}�BEmail�BPPV��I�q�ʱK�X�B�n�J�K�X
    '   �s�Ѽ� : �ȪA�����|���N���B�|���W�١B�����Ҧr���B�X�ͤ��-�~�B�X�ͤ��-��B�X�ͤ��-��B
    '   �p���q��(�v) �B�p���q��(��)�B��ʹq�ܡB�ʧO�B�q�T�a�}-�����B�q�T�a�}-�m���ϡB�q�T�a�}-���q��ѧ˫Ѹ��ӡB
    '   Email�B�ȪA�����n�J�b���BPPV�q�ʱK�X�B�ȪA�����|���K�X�B�K�X���ܰ��D�B�K�X���ܵ��סB
    '   �Ǿ��B�@�N���즳�u�q���A�Ȥ��W�D�`�ص������T���B�B�ê��p�B�P��H�ơB�P�����B¾�~�O�B
    '   �ӤH�~���J�B�����O�B�b�����A�B�ӽФ���ɶ��B�ҥΤ���B�̫�n�J����B�̫Ყ�ʤ���B
    '   �t�Υx�N�X(Web���|�����ƳQ���ʮɻݶ�)
    '   �o�����Y�B�o������(�G�p�ΤT�p)�B�o���a�}�B�Τ@�s��(��o���������T�p�ɶ�������) �B�]�Ƭy����
'2,200605080001,0160000036,aaa,200605080001,test,,2005,08,04,,,,�k,�x����,NO,��90��86��,winnie@eland.com.tw,null,,test,test,test,,,,,,��T/�����,30�U�H�U,1,Y,20050804,20050804,20050804,20050804,16,,,,,

    If cboOpt2.ListIndex = 0 Then
        Func2 = StartUp("2,MemberID,0010000006,PPVOrderPwd")
    Else
        Func2 = StartUp("2,MemberID,0010000006,PPVOrderPwd," & _
                                    "MemberID,Hammer,V111111111," & _
                                    "1975,09,22,05-3838338,07-5561926,0938123456,�k," & _
                                    "������,�����,�շR�G��198��11�Ӥ�1," & _
                                    "PowerHammer@msn.com,LoginID," & _
                                    "PPVOrderPwd,LoginPassword,HintName,HintAnswer," & _
                                    "���X��,1,���B,200,�@��H,�ۥѷ~,200�U,�S����,����," & _
                                    "2005/05/01,2005/05/06,2005/05/20,2005/05/20,1," & _
                                    "�o�����Y,,�o���a�},�Τ@�s��',000CE5DDF72C'")
    End If
End Function

Public Function Func3() As String
End Function

Public Function Func4() As String ' �Ȥ�q�ʭp���I�O�`�ة��Ӭd��(ppv,ippv) [���ѭӤH�P�޲z��] ( �R�O 4 )
'   4�B�|���N���B�{��ID�B�q�ʰ_��B�q�ʨ���
        Func4 = StartUp("4,HaHa66,0010000062,20050410,20050412")
'   0:  ���\
'   -5 : �d�L�q�ʸ��
End Function

Public Function Func5() As String ' �Ȥ�b�ȸ�Ƭd��(����϶����o�����b���ú�ڰO��)(�ӤH) ( �R�O 5 )
'   5�B�|���N���B�{��ID�B����_��B�������
        Func5 = StartUp("5,HaHa66,0010000022,20050101,20050412")
'   0:  ���\
'   -6 : �d�L�b�ȸ��
End Function

Public Function Func6() As String ' �Ȥᦳ�u�q�������O��b�@�~/�p���I�O�w�I���B��b�@�~(�H�Υd�p�B�I�ھ���)--�u�W�I�ڱ��v--���d�� ( �R�O 6 )
'   6�B�|���N���B�{��ID
        Func6 = StartUp("6,HaHa66,0010000022")
'   0:  ���\
'   -6 : �d�L�b�ȸ��
End Function

Public Function Func7() As String ' �Ȥᦳ�u�q�������O��b�@�~/�p���I�O�w�I���B��b�@�~(�H�Υd�p�B�I�ھ���)--�u�W�I�ڱ��v ( �R�O 7 )
'   7�B�|���N���B�{��ID�B�d���B�H�Υd���Ĵ����B�H�Υd�I�����v�X�B�H�Υd�O(���)
'   ��ڽs���B����(���)
'   ��ڽs���B����(���)
        Func7 = StartUp("7,HaHa66,0010000002,1111111111111111,122007,078038,1" & vbCrLf & _
                                    "200509ID0506570,1" & vbCrLf & _
                                    "200509ID0506570,2")
'        Func7 = StartUp("7,HaHa66,0010000002,1111111111111111,122007,,1")
'   0:  ���\
'   -7 : �H�α��v���~
End Function

Public Function Func8() As String ' �I�ƽu�W�q��--�u�W�I�ڱ��v ( �R�O 8 )
'   8�B�|���N���B�{��ID�B�d���B�H�Υd���Ĵ����B�H�Υd�I�����v�X�B�H�Υd�O(���)
'   ���O���إN�� (���)
'   ���O���إN�� (���)
        Func8 = StartUp("8,HaHa66,0010000002,1111111111111111,122007,078038,1" & vbCrLf & _
                                    "235" & vbCrLf & _
                                    "301")
'        Func8 = StartUp("8,HaHa66,0010000002,1111111111111111,122007,,1")
'   0:  ���\
'   -7 : �H�α��v���~
End Function

Public Function Func9() As String ' �I�ƽu�W�q��--�U���I�� ( �R�O 9 )
'    9�B�|���N���B�{��ID(���)
'    �]�Ƭy�����BSTB�Ǹ��BICC�Ǹ��B�I��(���)
'    �]�Ƭy�����BSTB�Ǹ��BICC�Ǹ��B�I��(���)
'    �K
'        Func9 = StartUp("9,HaHa66,0010000022" & vbCrLf & "200406140057792,000CE5DDF72C,ICCsno,0")
        Func9 = StartUp("9,HaHa66,0010000022" & vbCrLf & "200406140057792,000CE5DDF72C,ICCsno,0" & vbCrLf)
'        Func9 = StartUp("9,HaHa66,0010000022" & vbCrLf & "200406140057792,000CE5DDF72C,ICCsno,0" & vbCrLf)
'        Func9 = StartUp("9,1126250439693,0160000028" & vbCrLf & "200502180068625,119121267688,116019066076,999")

'   0:  ���\
'   -8 : �I�ƤU�����~
End Function

Public Function Func10() As String ' �Ȥ��ܧ�STB�|�X�ƦrParental PIN Code ( �R�O 10 )
'    10�B�|���N���B�{��ID�B�]�Ƭy�����BSTB�Ǹ��BICC�Ǹ�
        Func10 = StartUp("10,HaHa66,0010000022,200406140057792,000CE5DDF72C,ICCsno")
'   0:  ���\
'   -9 : �K�X�ܧ󥢱�
End Function

Public Function Func11() As String ' �Ȥ� iPPV�q��PIN Code ( �R�O 11 )
'   11�B�|���N���B�{��ID�B�]�Ƭy�����BSTB�Ǹ��BICC�Ǹ��BiPPV PIN Code(4�X)
        Func11 = StartUp("11,HaHa66,0010000022,200406140057792,000CE5DDF72C,ICCsno,ABCD")
'   0:  ���\
'   -9 : �K�X�ܧ󥢱�
End Function

'cable_liga: API-12�W�沧�ʡG 1�D�W�[[����]�Ѽ� 2�D�p�p���B�B�p�p�I�ƥH�N�ũҶǰѹL�Ӫ��^�s�C
'cable_liga: 5�D �p�p���B=�`�ػ���+����O [�N�ŶǹL�Ӫ�����]�B �p�p�I��=EPG�W���I��+�I�� [�N�ŶǹL�Ӫ��I��]
'cable_liga: �]�Ƭy�����BSTB�Ǹ��BICC�Ǹ��B���~�N�X�B�`�ئW�١B�I�ơB����(���)

Public Function Func12() As String ' �p���I�O�`�حq��(�ݸg�LSMS�q�ʱK�X�{��,�å]�ACA�t�έq�ʱ��v�@�~�γB�z���G) ( �R�O 12 )
'   12�B�|���N���B�{��ID�BPPV��I�q�ʱK�X�B2(��ӷ�����) (���)
'   �]�Ƭy�����BSTB�Ǹ��BICC�Ǹ��B���~�N�X�B�`�ئW�١B����(���)
'   �]�Ƭy�����BSTB�Ǹ��BICC�Ǹ��B���~�N�X�B�`�ئW�١B����(���)
'   �K�K..

'   �`�حq�ʩ��Ӫ���6�ӰѼơA���I�Ƨ�Ǭ�����C�ѼƭӼƤ���
        Func12 = StartUp("12,HaHa66,0010000002,PPVOrderPwd,2" & vbCrLf & _
                                        "200509160058579,00111AC86A30,ICCsno,1,HaHa,200,2006/06/20 12:12:12" & vbCrLf & _
                                        "200512270058636,0002CF011E10,ICCsno,1,HaHa,400,2006/06/20 12:12:12")
'   0:  ���\
'   -2:  �q��PPV�{�ҿ��~
'   -4 : ���ҥ�PPV�q���v
'   -11 : �p���I�O�`�حq�ʥ���
End Function

Public Function Func13() As String ' �p���I�O�`�ب���(�ݸg�LSMS�q�ʱK�X�{��,�å]�ACA�t�έq�ʱ��v�@�~�γB�z���G) ( �R�O 13 )
'   13�B�|���N���B�{��ID�BPPV��I�q�ʱK�X�B2(��ӷ�����)�B�����ɶ�(YYYYMMDD HH24MI)�B�����H���N���B��������(�����H��)�B�q�������]�N�X�B�q�������](���)
'   �q��渹�B�q�涵��(���)
'   �q��渹�B�q�涵��(���)
'   �K
'        Func13 = StartUp("13,HaHa66,0010000002,PPVOrderPwd,2,20060119 2325,WEB,��������,38,���n����" & _
                                        "200000071,1" & vbCrLf & _
                                        "200000071,2")

        Func13 = StartUp("13,HaHa66,0010000002,PPVOrderPwd,2,20060119 2325,WEB,��������,38,���n����" & vbCrLf & _
                                        "2,1" & vbCrLf & _
                                        "3,2")

'   PS�K�����H���N��='WEB'  ,   �����H��='��������'

'   0:  ���\
'   -4 : ���ҥ�PPV�q���v
'   -12 : �p���I�O�`�ب����q�ʥ���

End Function

Public Function Func14() As String ' �p���I�O�`�ب�����]��� ( �R�O 14 )
'   14
        Func14 = StartUp("14,HaHa66,0010000022")
'   0:  ���\
'   -13 : ������]�N�XŪ������
End Function

Public Function Func17() As String ' �Ȥ�w��STB�]�Ƹ�Ƭd��(�w�w��STB����,�s���Φ�m) ( �R�O 17 )
'   17�B�|���N���B�{��ID
        Func17 = StartUp("17,HaHa66,0010000012")
'   0 : ���\
'   -14 : �d�L�w��STB�]�Ƹ��
End Function

Public Function Func18() As String ' Ū���I�Ʀ��O��� ( �R�O 18 )
    '18�B�|���N���B�{��ID
        Func18 = StartUp("18,HaHa66,0010000022")
'   0:  ���\
'   -15 : �d�L�I�Ʀ��O���
End Function

Public Function Func19() As String ' �Ǧ^SMS�|����� For web update ( �R�O 19 )
'   19�B�|���N���B�{��ID
        Func19 = StartUp("19,HaHa66,0010000062")
'   0:  ���\
'   -16 : �d�L�Ǧ^SMS�|�����
End Function

Public Function Func20() As String ' ��O�Ȥ�W��d�� ( �R�O 20 )
'   20�B�|���N���B�{��ID
        Func20 = StartUp("20,HaHa66,0010000022")
'   0:  ���\
'   -17 : �d�L��O���
End Function

Public Function Func21() As String ' �g��ϴM�� ( �R�O 21 )
'   21�B(�t�Υx�O�ƿ�)�B�����B�m���B�����ܮH�����L�B�F�B�q�B�ѡB�ˡB���B���B��L
'        Func21 = StartUp("21,1,,���c��,�N�L��,1�F,2��,�~��1��,,2��,6�Ӥ�1")
'        Func21 = StartUp("21,1,�z�Ѯv��,���c��,�N�L��,1�F,2��,�~��1��,,2��,6�Ӥ�1")
'        Func21 = StartUp("21,1,,��饫,�j�����,,1�q,,,70��,")
        ' ��饫�j�����1�q70��
        Func21 = StartUp("21,2" & vbTab & "1,,���c��,�N�L��,1�F,2��,�~��1��,,2��,6�Ӥ�1")
'        Func21 = StartUp("21,1" & vbTab & "1,�z�Ѯv��,���c��,�N�L��,1�F,2��,�~��1��,,2��,6�Ӥ�1")
'   0:  ���\
'   -18 : �d�L�g��ϸ��! �i�ण�b�F�˸g��ϩΥثe���t�u�ܸӦa�}
End Function

Public Function Func22() As String ' �Ȥ�q�ʼƦ��W�D�d�� ( �R�O 22 )
'   �� :22, ���q�O,�ӽФH, �������r��, �q��, ����(1=�d�b�ȸ�T�B2=�d�]�Ƹ�T),�_�l���,�I����]
'   22�B�|���N���B�{��ID�B����(1=�d�b�ȸ�T�B2=�d�]�Ƹ�T)�B�_�l����B�I����
    If cboOpt22.ListIndex <= 0 Then
        Func22 = StartUp("22,HaHa66,0050000002,1,20000101,20060922")
    Else
        Func22 = StartUp("22,HaHa66,0010000012,2,20000101,20060922")
    End If
'   0:  ���\
'   -6 : �d�L�b�ȸ��
'   -14, �d�L�w��STB�]�Ƹ��
End Function

Public Function Func23() As String ' CMCP ������T�d�� ( �R�O 23 )
'   23�B���q�O�B�ӽФH�B�������r���B�ͤ�B����(1=�d�b�ȸ�T�B2=�d�]�Ƹ�T)�B�_�l����B�I����
    If cboOpt23.ListIndex <= 0 Then
        Func23 = StartUp("23,1,���B��,A100000001,1977/01/03,1,2005/09/22,2006/09/22")
    Else
        Func23 = StartUp("23,1,���p��,V111111111,1978/09/03,2")
'        Func23 = StartUp("23,1,���p��,V111111111,1978/09/03,2,2005/09/22,2006/09/22")
    End If
'        Func23 = StartUp("23,1,���M�D,84177282,1949/08/18,1,2005/09/22,2006/09/22")
'        Func23 = StartUp("23,1,Hammer,V111111111,1975/09/22,2,2005/09/22,2006/09/22")
'   0:  ���\
'   -6 : �d�L�b�ȸ��
'   -21, �d�L�w��CMCP�]�Ƹ��
End Function

Public Function Func24() As String ' ETHOME-CM������T�d�� ( �R�O 24 )
'   24�B�ӽФH�m�W�B���������B�νs�B�]�ƺ����BCM MAC

'   �������� ( 1:������;  2:�Τ@�s��)
'   �����Ҧr�� �� �Τ@�s��,
'   �]�ƺ��� (1: MAC; 2: �b��)
'   CM MAC �� ACS �b��

    If cboOpt24.ListIndex <= 0 Then
        ' �]�ƺ��� MAC
        Func24 = StartUp("24,�H�w��,1,03-2530903,2,1111111111111111")
    Else
        ' �]�ƺ��� �b��
        Func24 = StartUp("23,1,���p��,V111111111,1978/09/03,2")
'        Func24 = StartUp("23,1,���p��,V111111111,1978/09/03,2,2005/09/22,2006/09/22")
    End If
'   0:  ���\
'   -99 : �ѼƤ���
'   -20, �d�L�ŦX���󤧸��
End Function


Public Function Func25() As String
'���u������
    Func25 = StartUp("25,5,9,200708II0032900,20070815 153050,,20070815 153050,0363,0363,20070815,")
'�h�檺����
    'Func25 = StartUp("25,5,17,200708II0032899,,401,20070815 153050,0363,0363,20070815,")
End Function


Public Function Func26() As String
'���u������
    Func26 = StartUp("26,5,2410,200708PI0012041,20070815 153050,,20070815 153050,0363,0363,20070815,0,363")
'�h�檺����
     'Func26 = StartUp("26,5,17,200708PI0012035,,101,20070815 153050,0363,0363,20070815,0,365")
End Function


Public Function Func28() As String
'���u������
    Func28 = StartUp("28,5,2410,200708PI0012041,20070815 153050,,20070815 153050,0363,0363,20070815,0,363")
'�h�檺����
     'Func28 = StartUp("28,5,17,200708PI0012035,,101,20070815 153050,0363,0363,20070815,0,365")
End Function


'�Ȥ� iPPV�q��PIN Code
Private Function GetOpt() As Byte
    Dim ctl As Control
    For Each ctl In Me
        If TypeOf ctl Is OptionButton Then
            If ctl.Value Then
                GetOpt = ctl.Index
                Exit For
            End If
        End If
    Next
    On Error Resume Next
    Set ctl = Nothing
End Function

Private Sub cmdExit_Click()
  On Error Resume Next
    Unload frmSimulator
    End
End Sub
Public Function Func27() As String ' �Ȥᦳ�u�q�������O��b�@�~/�p���I�O�w�I���B��b�@�~(�H�Υd�p�B�I�ھ���)--�u�W�I�ڱ��v ( �R�O 7 )
'   7�B�|���N���B�{��ID�B�d���B�H�Υd���Ĵ����B�H�Υd�I�����v�X�B�H�Υd�O(���)
'   ��ڽs���B����(���)
'   ��ڽs���B����(���)
        Func27 = StartUp("27,5,2838,06090470900")
'        Func7 = StartUp("7,HaHa66,0010000002,1111111111111111,122007,,1")
'   0:  ���\
'   -7 : �H�α��v���~
End Function

