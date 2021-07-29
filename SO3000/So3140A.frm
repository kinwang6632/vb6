VERSION 5.00
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Begin VB.Form frmSO3140A 
   Caption         =   "減免出單金額 [SO3140A]"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "So3140A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2745
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin prjNumber.GiNumber ginDisAmt 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   150
      Width           =   1005
      _ExtentX        =   1773
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
      ForeColor       =   0
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   435
      Left            =   3270
      TabIndex        =   2
      Top             =   2160
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "F5. 確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      TabIndex        =   1
      Top             =   2190
      Width           =   1245
   End
   Begin VB.Label lblDisAmt 
      AutoSize        =   -1  'True
      Caption         =   "減免金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   240
      Width           =   780
   End
End
Attribute VB_Name = "frmSO3140A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTmp As New ADODB.Recordset

Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ChkErr
        Dim strSQL As String
        Dim rsSo032 As New ADODB.Recordset
        Dim intCount As Integer
        
        
        If ginDisAmt.Text = "" Then MsgBox "此為必要欄位,需有值!", vbExclamation, "警告": ginDisAmt.SetFocus: Exit Sub
        gcnGi.BeginTrans
        Screen.MousePointer = vbHourglass
                        
        gcnGi.Execute "Update " & GetOwner & "So032 Set ShouldAmt=ShouldAmt-" & ginDisAmt.Value & " Where CustId In (Select Custid From " & GetOwner & "So032Tmp Where Flag = 0) And CitemCode = (Select CodeNo From " & GetOwner & "CD019 Where RefNO = 1)", intCount
        gcnGi.Execute "Update " & GetOwner & "So032Tmp Set Flag = 1,UpdTime ='" & GetDTString(RightNow) & "' Where CustId In (Select A.CustId From " & GetOwner & "So032 A," & GetOwner & "CD019 B Where A.CitemCode=B.CodeNo And B.RefNo=1)"
        MsgBox "執行完畢!共更新 " & intCount & " 筆資料", vbInformation, "訊息"
        gcnGi.CommitTrans
        Call CloseObj
        Screen.MousePointer = vbDefault
        
    Exit Sub
ChkErr:
    gcnGi.RollbackTrans
    If Err.Number = -2147217865 Then
        MsgBox "查無資料表So032Tmp,請確認!!", vbCritical, gimsgWarning
        Screen.MousePointer = vbDefault
    Else
        Call ErrSub(Me.Name, "cmdSave_Click")
    End If
End Sub

Private Sub CloseObj()
  On Error Resume Next
     Screen.MousePointer = vbDefault
     rsTmp.Close
     Set rsTmp = Nothing
End Sub

Private Sub Form_Activate()
  On Error Resume Next
     Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Not ChkGiList(KeyCode) Then Exit Sub
    If KeyCode = vbKeyEscape Then Unload Me: KeyCode = 0
    
    If KeyCode = vbKeyF5 Then cmdSave.Value = True: KeyCode = 0
End Sub
