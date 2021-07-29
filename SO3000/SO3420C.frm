VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO3420C 
   BorderStyle     =   1  '單線固定
   Caption         =   "瀏覽畫面 [SO3420C] "
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   11760
   Icon            =   "SO3420C.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10650
      TabIndex        =   2
      Top             =   4980
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      Height          =   375
      Left            =   9480
      TabIndex        =   1
      Top             =   4980
      Width           =   990
   End
   Begin prjGiGridR.GiGridR ggrView 
      Height          =   4875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   8599
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSO3420C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#3474 增加一個瀏覽晝面,可讓User先看看資料是否正確 By Kin 2008/05/14
Option Explicit
Private rsTmp As New ADODB.Recordset
Private blnOK As Boolean

Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
End Sub

Private Sub cmdOK_Click()
  On Error Resume Next
    frmSO3420A.uViewOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub subGrd()
  On Error GoTo ChkErr
    Dim mflds As New prjGiGridR.GiGridFlds
        With mflds
            .Add "CustId", , , , False, "客戶編號", vbLeftJustify
            .Add "BillNo", , , , False, "單據編號" & Space(10), vbLeftJustify
            .Add "CitemCode", , , , False, "收費項目代碼", vbLeftJustify
            .Add "CitemName", , , , False, "收費項目名稱" & Space(10), vbLeftJustify
            .Add "ShouldDate", giControlTypeDate, , , False, "應收日期   ", vbLeftJustify
            .Add "RealStartDate", giControlTypeDate, , , False, "收費起日   ", vbLeftJustify
            .Add "RealStopDate", giControlTypeDate, , , False, "收費迄日   ", vbLeftJustify
            .Add "ShouldAmt", , , , False, "應收金額   ", vbRightJustify
            .Add "RealAmt", , , , False, "實收金額   ", vbRightJustify
            .Add "RealDate", giControlTypeDate, , , False, "實收日期    ", vbLeftJustify
            .Add "UCCode", , , , , "未收原因代碼", vbLeftJustify
            .Add "UCName", , , , False, "未收原因名稱", vbLeftJustify
            .Add "PrtSNo", , , , False, "印單序號" & Space(10), vbLeftJustify
            .Add "MediaBillNo", , , , False, "媒體單號" & Space(10), vbLeftJustify
        End With
    ggrView.AllFields = mflds
    ggrView.SetHead
    If rsTmp.EOF = False Then rsTmp.MoveFirst
    Set ggrView.Recordset = rsTmp
    ggrView.Refresh
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subGrid"
End Sub

Public Property Let uRS(ByRef vData As Recordset)
    Set rsTmp = vData
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = vbKeyF2 Then If cmdOK.Enabled Then cmdOK.Value = True
    If KeyCode = vbKeyEscape Then Unload Me
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub
Private Sub Form_Load()
    blnOK = False
    Call subGrd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  Call CloseRecordset(rsTmp)
End Sub
