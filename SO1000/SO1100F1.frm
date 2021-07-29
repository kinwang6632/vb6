VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1100F1 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "����s�� [SO1100F1]"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7590
   Icon            =   "SO1100F1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7590
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "����(&X)"
      Height          =   405
      Left            =   4680
      TabIndex        =   5
      Top             =   4620
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�T�w"
      Height          =   405
      Left            =   1860
      TabIndex        =   3
      Top             =   4650
      Width           =   915
   End
   Begin VB.Frame fraSO033 
      Caption         =   "���O���"
      Height          =   2145
      Left            =   150
      TabIndex        =   1
      Top             =   2400
      Width           =   7305
      Begin prjGiGridR.GiGridR ggrData2 
         Height          =   1755
         Left            =   90
         TabIndex        =   4
         Top             =   270
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3096
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraSO003 
      Caption         =   "�g���ʸ��"
      Height          =   2145
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   7305
      Begin prjGiGridR.GiGridR ggrData 
         Height          =   1755
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3096
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSO1100F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#3929 �R���o����Ʈɭn�q�X�s���e�� By Kin 2008/06/02
Option Explicit
Private rsSO003 As New ADODB.Recordset
Private rsSO033 As New ADODB.Recordset
Private Sub subGrid()
    On Error Resume Next
    Dim mFlds As New GiGridFlds
    Dim mFlds2 As New GiGridFlds
    '#3929 ���դ�OK,�b���n�M��,�_�W��S�����T��� By Kin 2008/07/24
    With mFlds
        '.Add "AccountNo", , , , , "�b��" & Space(18)
        .Add "CitemName", , , , , "���O���ئW��" & Space(15)
        .Add "StartDate", , , , , "���İ_�l��"
        .Add "StopDate", , , , , "���ĺI���"
        .Add "FaciSNO", , , , , "�]�ƧǸ�" & Space(10)
    End With
    With mFlds2
        '.Add "AccountNo", , , , , "�b��" & Space(18)
        .Add "CitemName", , , , , "���O���ئW��" & Space(15)
        .Add "RealStartDate", , , , , "���İ_�l��"
        .Add "RealStopDate", , , , , "���ĺI���"
        .Add "FaciSNO", , , , , "�]�ƧǸ�" & Space(10)
    End With
    ggrData.AllFields = mFlds
    ggrData.SetHead
    Set ggrData.Recordset = rsSO003
    ggrData.Refresh
    ggrData2.AllFields = mFlds2
    ggrData2.SetHead
    Set ggrData2.Recordset = rsSO033
    ggrData2.Refresh
End Sub
Public Property Let uRs1(ByRef vData As ADODB.Recordset)
  On Error Resume Next
  Set rsSO003 = vData
End Property
Public Property Let uRs2(ByRef vData As ADODB.Recordset)
  On Error Resume Next
  Set rsSO033 = vData
End Property

Private Sub cmdCancel_Click()
  On Error Resume Next
    frmSO1100F.uDelete = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
  On Error Resume Next
    frmSO1100F.uDelete = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Select Case KeyCode
        Case vbKeyF2
            cmdOK.Value = True
        Case vbKeyEscape
            cmdCancel.Value = True
    End Select
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
    Call subGrid
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  Call Close3Recordset(rsSO003)
  Call Close3Recordset(rsSO033)
End Sub
