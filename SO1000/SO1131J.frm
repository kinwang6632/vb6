VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO1131J 
   BorderStyle     =   1  '單線固定
   Caption         =   "收費項目置換 [SO1131J]"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "SO1131J.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   3750
      TabIndex        =   3
      Top             =   720
      Width           =   825
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   825
   End
   Begin prjGiList.GiList gilCitemCode 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   210
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth2       =   2400
      TableName       =   "CD019"
      F2Corresponding =   -1  'True
   End
   Begin VB.Label lblCitemCode 
      AutoSize        =   -1  'True
      Caption         =   "收費項目"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   270
      Width           =   795
   End
End
Attribute VB_Name = "frmSO1131J"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#4150 增加收費項目置換功能 By Kin 2008/10/17
Option Explicit
Private strCitemCode As String
Private Sub cmdCancel_Click()
  On Error GoTo ChkErr
    frmSO1131C.uChgCitem = Empty
    Unload Me
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdCancel_Click")
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    If gilCitemCode.GetCodeNo & "" <> "" Then
        frmSO1131C.uChgCitem = gilCitemCode.GetCodeNo
        Unload Me
    End If
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGil
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub
Private Sub subGil()
  On Error GoTo ChkErr
    SetgiList gilCitemCode, "CodeNO", "Description", "CD019", , , , , , , "Where PeriodFlag=1 ", True
    Call GiListFilter(gilCitemCode, gServiceType)
    gilCitemCode.Filter = gilCitemCode.Filter & " And PeriodFlag=1"
    If strCitemCode <> Empty Then
        gilCitemCode.SetCodeNo strCitemCode
        gilCitemCode.Query_Description
    End If
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub
Public Property Let uCitemCode(ByVal vData As String)
  On Error Resume Next
  strCitemCode = vData
End Property
