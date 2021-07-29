VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Begin VB.Form frmLock 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�{���[�K"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6165
   StartUpPosition =   2  '�ù�����
   Begin VB.TextBox txtMacAddress 
      Height          =   405
      Left            =   3540
      TabIndex        =   7
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "�ƨ�[�K��r"
      Height          =   465
      Left            =   1380
      TabIndex        =   4
      Top             =   1260
      Width           =   1245
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "�[�K"
      Height          =   465
      Left            =   90
      TabIndex        =   3
      Top             =   1260
      Width           =   1245
   End
   Begin VB.TextBox txtLock 
      Height          =   435
      Left            =   120
      ScrollBars      =   1  '�������b
      TabIndex        =   2
      Top             =   690
      Width           =   5955
   End
   Begin Gi_Date.GiDate gadEnd 
      Height          =   345
      Left            =   1230
      TabIndex        =   1
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "MacAddress"
      Height          =   225
      Left            =   2550
      TabIndex        =   6
      Top             =   300
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "�[�K��ƥi�঳�S��r���I�̦n���ƨ�K��INI"
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   2730
      TabIndex        =   5
      Top             =   1290
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "�{�������"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Width           =   1005
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function IsDataOK() As Boolean
  On Error Resume Next
    If Not IsDate(gadEnd.GetValue(True)) Then
        MsgBox "��J����������T�I", vbInformation, "ĵ�i"
        gadEnd.SetFocus
        Exit Function
    End If
    IsDataOK = True
End Function

Private Sub cmdCopy_Click()
  On Error Resume Next
  Clipboard.Clear
  Clipboard.SetText txtLock.Text
End Sub
Private Sub cmdLock_Click()
  On Error GoTo ChkErr
    Dim strEncoding
    If Not IsDataOK Then Exit Sub
    Dim obj As Object
    Set obj = CreateObject("csLockProcess.EnableProcess")
    Dim strMac As String
    strMac = txtMacAddress.Text
    strEncoding = obj.Encryption(gadEnd.GetValue(True), strMac)
    If strEncoding = "" Then GoTo ChkErr
    txtLock.Text = strEncoding
    
    'cmdCopy.Enabled = True
    'MsgBox CreateObject("Encryption.Password").Decrypt(strEncoding, "CS")
    Exit Sub
    'Xdecrypt = CreateObject("Encryption.Password").Decrypt(EncryptString, EncryptKey)
ChkErr:
  MsgBox "�[�K���~�I", vbInformation, "ĵ�i"
  End
End Sub

Private Sub Form_Load()
  On Error Resume Next
  cmdCopy.Enabled = False
End Sub

Private Sub txtLock_Change()
  On Error Resume Next
    If txtLock.Text = "" Then
        cmdCopy.Enabled = False
    Else
        cmdCopy.Enabled = True
    End If
End Sub

Private Sub txtLock_GotFocus()
  On Error Resume Next
    txtLock.SelLength = 9999
End Sub

