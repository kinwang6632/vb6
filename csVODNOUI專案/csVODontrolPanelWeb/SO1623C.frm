VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1623C 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "瀏覽畫面"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   ClipControls    =   0   'False
   Icon            =   "SO1623C.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   10110
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5.列印"
      Default         =   -1  'True
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
      Left            =   7230
      TabIndex        =   2
      Top             =   5010
      Visible         =   0   'False
      Width           =   1365
   End
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
      Left            =   8700
      TabIndex        =   1
      Top             =   5010
      Width           =   1365
   End
   Begin prjGiGridR.GiGridR ggrView 
      Height          =   4935
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   8705
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSO1623C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsData As New ADODB.Recordset

Private Sub cmdCancel_Click()
    On Error Resume Next
        Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error Resume Next
'        ReadyGoPrint
'        PrintRpt GetPrinterName(5), "SO3220A.rpt", , "收費單給號地址調整一覽表 [SO3220A]", "SELECT * From tmp007", , , True, , , , GiPaperLandscape
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ChkErr
        If KeyCode = vbKeyEscape Then Unload Me
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr
        Call InitData
        Call subGrd
        Call OpenData
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Sub InitData()
    On Error Resume Next
        gLogInID = GetSystemParaItem("LogInID", gCompCode, "", "SO041", , True) & ""
        If gLogInID <> "" Then gLogInID = gLogInID & "."

End Sub

Private Sub OpenData()
    On Error GoTo ChkErr
        If Not GetRS(rsData, "Select SeqNo, CompCode, HIGH_LEVEL_CMD_ID, ICC_NO, STB_NO, SUBSCRIPTION_BEGIN_DATE, " & _
            "SUBSCRIPTION_END_DATE, CMD_STATUS, ProcessingDate From " & gLogInID & "SEND_NAGRA Where CompCode = " & gCompCode) Then Exit Sub
        Set ggrView.Recordset = rsData
        ggrView.Refresh
    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub

Private Sub subGrd()
    On Error GoTo ChkErr
        Dim mFlds As New prjGiGridR.GiGridFlds
            With mFlds
                .Add "CompCode", , , , False, "公司別", vbLeftJustify
                .Add "CMD_STATUS", , , , False, "命命狀態", vbLeftJustify
                .Add "SeqNo", , , , False, "序號   ", vbLeftJustify
                .Add "ICC_NO", , , , False, "智慧卡序號", vbLeftJustify
                .Add "STB_NO", , , , False, "STB 序號  ", vbLeftJustify
                .Add "HIGH_LEVEL_CMD_ID", , , , False, "指令名稱     ", vbLeftJustify
                .Add "SUBSCRIPTION_BEGIN_DATE", giControlTypeDate, , , False, "收視起始日", vbLeftJustify
                .Add "SUBSCRIPTION_END_DATE", giControlTypeDate, , , False, "收視截止日", vbLeftJustify
                .Add "ProcessingDate", giControlTypeTime, , , False, "預約時間        ", vbLeftJustify
            End With
        ggrView.AllFields = mFlds
        ggrView.SetHead
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subGrd"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ChkErr
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_QueryUnload"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub ggrView_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error GoTo ChkErr
        Select Case UCase(Fld.Name)
            Case "HIGH_LEVEL_CMD_ID"
                Select Case Value
                    Case "A1"
                        Value = "開機"
                    Case "A2"
                        Value = "關機"
                    Case "A3"
                        Value = "停機"
                    Case "A4"
                        Value = "復機"
                    Case "A5"
                        Value = "註銷"
                    Case "B1"
                        Value = "開頻道"
                    Case "B2"
                        Value = "關頻道(個別)"
                    Case "B3"
                        Value = "關頻道(全部)"
                    Case "B4"
                        Value = "暫停頻道"
                    Case "B5"
                        Value = "恢復頻道"
                    Case "B6"
                        Value = "期限延展"
                    Case "B7"
                        Value = "收視權限複製"
                    Case "E1"
                        Value = "設定親子密碼"
                    Case "E2"
                        Value = "傳送訊息"
                End Select
            Case "CMD_STATUS"
                Select Case Value
                    Case "C"
                        Value = "完成"
                    Case "W"
                        Value = "等候處理"
                    Case "P"
                        Value = "處理中"
                    Case "E"
                        Value = "錯誤"
                End Select
        End Select
        
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ggrCustRate_ShowCellDate")

End Sub
