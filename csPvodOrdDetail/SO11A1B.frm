VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO11A1B 
   BorderStyle     =   1  '單線固定
   Caption         =   "SO11A1B [短收]"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   6975
   Icon            =   "SO11A1B.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdExit 
      Caption         =   "離開(&X)"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   2460
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   2460
      Width           =   945
   End
   Begin VB.TextBox txtMemo 
      Height          =   1485
      Left            =   1170
      TabIndex        =   2
      Top             =   720
      Width           =   5655
   End
   Begin prjGiList.GiList gilSTCode 
      Height          =   345
      Left            =   1170
      TabIndex        =   0
      Top             =   270
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   540
      TabIndex        =   5
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "短收原因"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   1
      Top             =   330
      Width           =   885
   End
End
Attribute VB_Name = "frmSO11A1B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsCancel As ADODB.Recordset
Private Sub subGil()
  On Error GoTo ChkErr
    SetgiList gilSTCode, "CodeNo", "Description", "CD016", , , , , 3, 12, " Where RefNo = 3 ", True
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub cmdExit_Click()
  On Error Resume Next
  Unload Me
End Sub
Private Function UpdData(ByRef aCnt As Long) As Boolean
  On Error GoTo ChkErr
  Dim aQryWhere As String
  Dim aUpdField As String
  Dim aNow As String
  rsCancel.MoveFirst
  aNow = Format(RightNow, "yyyymmdd")
  Do While Not rsCancel.EOF
    DoEvents
    aQryWhere = Empty
    aUpdField = Empty
    If Len(rsCancel("BillNo") & "") > 0 Then
        aQryWhere = "SO033.BillNo = '" & rsCancel("BillNo") & "'"
        If Val(rsCancel("OrderWayCode") & "") = 4 And Len(rsCancel("Item") & "") > 0 Then
            aQryWhere = aQryWhere & " AND SO033.ITEM = " & rsCancel("ITEM")
        End If
        aUpdField = " UCCODE = NULL,UCName=NULL,RealDate = To_DATE('" & aNow & "','yyyymmdd'), " & _
                "STCode = " & gilSTCode.GetCodeNo & ",STName ='" & gilSTCode.GetDescription & "'"
        If txtMemo.Text <> Empty Then
            aUpdField = aUpdField & ",Note='" & txtMemo.Text & "'"
        End If
        gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & aUpdField & " WHERE " & aQryWhere
    End If
    cnCom.Execute "UPDATE SO033PVOD SET CANCELFLAG = 1 WHERE SEQNO = " & rsCancel("SEQNO")
    aCnt = aCnt + 1
    rsCancel.MoveNext
  Loop
  UpdData = True
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "UpdData")
End Function
Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    Dim aCnt As Long
    If Not IsDataOK Then Exit Sub
    aCnt = 0
    Screen.MousePointer = vbHourglass
    gcnGi.BeginTrans
    cnCom.BeginTrans
    If UpdData(aCnt) Then
        gcnGi.CommitTrans
        cnCom.CommitTrans
        MsgBox "取消成功！" & vbCrLf & "筆數：" & aCnt & " 筆", vbInformation, "訊息"
        Unload Me
    Else
        gcnGi.RollbackTrans
        cnCom.RollbackTrans
        MsgBox "更新資料失敗！", vbCritical, "訊息"
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
 Call ErrSub(Me.Name, "cmdOK_Click")
End Sub
Private Function IsDataOK() As Boolean
  On Error GoTo ChkErr
    If Len(gilSTCode.GetCodeNo & "") = 0 Then
        MsgBox "請選擇短收原因", vbInformation, "訊息"
        Exit Function
    End If
    
    IsDataOK = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOK")
End Function

Private Sub Form_Activate()
  On Error Resume Next
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF2 Then
        cmdOK.Value = True
    End If
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGil
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub
Public Property Set uRs(ByRef vData As ADODB.Recordset)
    Set rsCancel = vData
End Property
