VERSION 5.00
Object = "{818C3FA5-C15E-4B09-B383-B06546ED97C7}#1.0#0"; "PHbtn.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   " �t �� �n �J"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.TextBox txtPwd 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  '�Ȥ�
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   " �п�J�K�X ! "
      Top             =   1080
      Width           =   2715
   End
   Begin PHbtn.PHbutton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2070
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "�T �w (&Y)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      FCOL            =   8388608
   End
   Begin PHbtn.PHbutton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2940
      TabIndex        =   2
      Top             =   2070
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "�� �� (&N)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      FCOL            =   8388608
   End
   Begin VB.Label lblPwd 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BackStyle       =   0  '�z��
      Caption         =   "�K �X :"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   630
      TabIndex        =   4
      Top             =   1185
      Width           =   495
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "�� �� �J �K �X �n ��"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   780
      TabIndex        =   3
      Top             =   210
      Width           =   3015
   End
   Begin VB.Shape Shp 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  '���z��
      BorderColor     =   &H00808080&
      FillColor       =   &H00400000&
      Height          =   1035
      Left            =   360
      Shape           =   4  '�ꨤ�x��
      Top             =   765
      Width           =   3855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PR : Hammer
' Start date : 2004/12/23
' Last Modify : 2004/12/23

Option Explicit

Private bytCount As Byte

Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
    End
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    With txtPwd
        If Len(Trim(.Text)) = 0 Then
            MsgBox "�п�J�n�J�K�X !", vbInformation, "�T��"
            On Error Resume Next
           .SetFocus
        Else
            If StrComp(.Text, strPassword) = 0 Then
                Hide
                frmMain.Show
                Unload Me
            Else
                If bytCount = 2 Then
                    MsgBox "�Ъ`�N , �z�w�n�J�T�����~ , �N�L�k�ϥΥ��t�� !", vbInformation, "�T��"
                    Unload Me
                    End
                Else
                    bytCount = bytCount + 1
                    If StrComp(.Text, strPassword, vbTextCompare) = 0 Then
                        MsgBox "�K�X���~ , �Э��s��J ! ( �`�N : �K�X�Ϥ��j�p�g ) ", vbInformation, "�T��"
                    Else
                        MsgBox "�K�X���~ , �Э��s��J !", vbInformation, "�T��"
                    End If
                    On Error Resume Next
                   .Text = ""
                   .SetFocus
                End If
            End If
        End If
    End With
  Exit Sub
ChkErr:
    ErrHandle Name, "Click Event : cmdOK_Click"
End Sub

Private Sub Form_Load()
    bytCount = 0
End Sub
