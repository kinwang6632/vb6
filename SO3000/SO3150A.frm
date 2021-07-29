VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Begin VB.Form frmSO3150A 
   BorderStyle     =   1  '單線固定
   Caption         =   "合併帳單[SO3150]"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3150A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdSave 
      Caption         =   "F2. 確定"
      Height          =   435
      Left            =   270
      TabIndex        =   2
      Top             =   1815
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&X)"
      Height          =   435
      Left            =   3360
      TabIndex        =   3
      Top             =   1815
      Width           =   1245
   End
   Begin Gi_YM.GiYM gymCBill 
      Height          =   315
      Left            =   1170
      TabIndex        =   1
      Top             =   555
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   90
      Width           =   3000
      _ExtentX        =   5292
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
      FldWidth1       =   800
      F2Corresponding =   -1  'True
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "公司別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   300
      TabIndex        =   5
      Top             =   150
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "合併年月"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   300
      TabIndex        =   4
      Top             =   615
      Width           =   780
   End
End
Attribute VB_Name = "frmSO3150A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ChkErr
    Dim strCombineDate As String
    Dim strMessage As String
        If Not IsDataOk Then Exit Sub
        If CallSF_CombineBill(gymCBill.GetValue, gilCompCode.GetCodeNo, strMessage) < 0 Then MsgBox strMessage, vbCritical, gimsgPrompt
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Function CallSF_CombineBill(ByVal strCombineDate As String, ByVal p_CompCode As Long, _
    ByRef strMessage As String) As Long
    On Error GoTo ChkErr
    Dim Cmd_CombineBill As New ADODB.Command
        ' 以下執行後端程式 SF_CombineBill('200208',:msg);
        With Cmd_CombineBill
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 10)
            .Parameters.Append .CreateParameter("p_CombineYM", adVarChar, adParamInput, 8, strCombineDate)
            .Parameters.Append .CreateParameter("P_COMPCODE", adVarNumeric, adParamInput, , p_CompCode)
            .Parameters.Append .CreateParameter("p_RetMsg", adVarChar, adParamOutput, 2000)
            Set .ActiveConnection = gcnGi
            .CommandText = "sf_CombineBill"
            .CommandType = adCmdStoredProc
            .Execute
            CallSF_CombineBill = Val(.Parameters(0) & "")
            strMessage = .Parameters(2) & ""
        End With
        Set Cmd_CombineBill = Nothing
        CallSF_CombineBill = True
    Exit Function
ChkErr:
    CallSF_CombineBill = -98
    strMessage = "程式錯誤!!"
    ErrSub Me.Name, "CallSF_CombineBill"
End Function

Private Function IsDataOk() As Boolean
    On Error Resume Next
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        If Not MustExist(gymCBill, 1, "合併年月") Then Exit Function
        IsDataOk = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "IsDataOk"
End Function

Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF2 Then cmdSave.Value = True: KeyCode = 0
        If KeyCode = vbKeyEscape Then Unload Me: KeyCode = 0
End Sub

Private Sub Form_Load()
    On Error Resume Next
        Call subGil
End Sub

Private Sub subGil()
    On Error GoTo ChkErr
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3100", "SO3150") Then Exit Sub
        Call subGil
End Sub
