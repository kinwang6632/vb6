VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO3210A 
   BorderStyle     =   1  '單線固定
   Caption         =   "收費單過入 [SO3210A]"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3210A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.PictureBox pic1 
      Height          =   1005
      Left            =   570
      ScaleHeight     =   945
      ScaleWidth      =   5085
      TabIndex        =   7
      Top             =   990
      Width           =   5145
      Begin VB.TextBox txtComPaswrd 
         ForeColor       =   &H000040C0&
         Height          =   315
         IMEMode         =   3  '暫止
         Left            =   1350
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   540
         Width           =   3135
      End
      Begin VB.TextBox txtPassword 
         ForeColor       =   &H00004080&
         Height          =   315
         IMEMode         =   3  '暫止
         Left            =   1350
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   150
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "確認密碼"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   600
         Width           =   780
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         Caption         =   "請輸入密碼"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   210
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   4305
      TabIndex        =   5
      Top             =   2925
      Width           =   1425
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2. 確認"
      Default         =   -1  'True
      Height          =   375
      Left            =   525
      TabIndex        =   4
      Top             =   2925
      Width           =   1425
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1530
      TabIndex        =   0
      Top             =   90
      Width           =   2940
      _ExtentX        =   5186
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
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   1530
      TabIndex        =   1
      Top             =   540
      Width           =   2940
      _ExtentX        =   5186
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
      F2Corresponding =   -1  'True
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "公司別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   630
      TabIndex        =   11
      Top             =   180
      Width           =   585
   End
   Begin VB.Label lblServiceType 
      AutoSize        =   -1  'True
      Caption         =   "服務類別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   630
      TabIndex        =   10
      Top             =   600
      Width           =   780
   End
   Begin VB.Label lbl1 
      Caption         =   "本功能將先前於出單管理中的""本期收費資料產生""所產生的暫存收費資料轉入正式的收費資料檔中, 請確定已查核過這些資料內容"
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   570
      TabIndex        =   6
      Top             =   2160
      Width           =   5115
   End
End
Attribute VB_Name = "frmSO3210A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intPara24 As Integer
Private blnGetPara As Boolean
Private blnIsOK As Boolean
Public rsM As New ADODB.Recordset
Public rsD As New ADODB.Recordset
Private blnAutoExec As Boolean
Private strReturnMsg As String

Public Property Let uAutoExec(vData As Boolean)
    blnAutoExec = vData
End Property

Public Property Let uGetPara(vData As Boolean)
    blnGetPara = vData
End Property

Public Property Get uIsOk() As Boolean
    uIsOk = blnIsOK
End Property

Public Property Get uReturnMsg() As String
    uReturnMsg = strReturnMsg
End Property

Private Sub cmdCancel_Click()
        Unload Me
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ChkErr
    Dim strmsg As String
    Dim lngRetCode As Long
        If Not IsDataOk Then Exit Sub
        '呼叫取參數 2011/11/30 Jacky 6153 MinChen
        If blnGetPara Then
            Call GetPara
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        gcnGi.BeginTrans '#4317 Transcation改由我這邊控制 By Kin 2009/01/20
        lngRetCode = SF_TmpToC1(gcnGi, garyGi(1), gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "", strmsg)
        If lngRetCode < 0 Then
            gcnGi.RollbackTrans
        Else
            gcnGi.CommitTrans
        End If
        Screen.MousePointer = vbDefault
        If Not blnAutoExec Then
            MsgBox strmsg, vbInformation, "訊息！"
        Else
            strReturnMsg = strmsg
            blnIsOK = True
            Unload Me
        End If
        Unload Me
    Exit Sub
ChkErr:
    gcnGi.RollbackTrans
    Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Function GetPara() As Boolean
    If Not BatchCharge.GetBatchChargePara(Me, rsM, rsD) Then Exit Function
    blnIsOK = True
    GetPara = True
    Unload Me
End Function

Private Function SetPara() As Boolean
    If Not BatchCharge.SetBatchChargePara(Me, rsM, rsD) Then Exit Function
    If blnAutoExec Then
        Me.txtComPaswrd.Text = garyGi(12)
        Me.txtPassword.Text = garyGi(12)
    End If
    SetPara = True
End Function

Private Function IsDataOk() As Boolean
    On Error GoTo ChkErr
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        If intPara24 = 0 Then If Not MustExist(gilServiceType, 2, "服務類別") Then Exit Function
        If txtPassword.Text = "" Then txtPassword.SetFocus: MsgBox "請輸入密碼", vbExclamation, gimsgPrompt: Exit Function
        If txtComPaswrd.Text = "" Then txtComPaswrd.SetFocus: MsgBox "請確認密碼", vbExclamation, gimsgPrompt: Exit Function
        If txtPassword.Text <> txtComPaswrd Then txtComPaswrd.SetFocus: MsgBox "確認密碼與輸入密碼不符", vbExclamation, gimsgPrompt: Exit Function
        If txtPassword.Text <> garyGi(12) Then txtPassword.SetFocus: MsgBox "輸入密碼不正確!!", vbExclamation, gimsgPrompt: Exit Function
        If txtComPaswrd.Text <> garyGi(12) Then txtComPaswrd.SetFocus: MsgBox "確認密碼不正確!!", vbExclamation, gimsgPrompt: Exit Function
        
        IsDataOk = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "IsDataOk"
End Function

Public Function SF_TmpToC1(ByVal cnConn As ADODB.Connection, ByVal p_UserId As String, _
    p_CompCode As Long, p_ServiceType As String, p_FileName As String, ByRef P_RETMSG As String) As Long
On Error GoTo ChkErr
        Dim ComSF_TMPTOC1 As New ADODB.Command
        If cnConn Is Nothing Then
                'MsgBox vmsgSendConn, vbCritical, vmsgPrompt
                Exit Function
        End If
        'function SF_TmpToC1(p_UserId IN VARCHAR2, p_CompCode IN NUMBER, p_ServiceType IN VARCHAR2, p_ShouldDate1 IN VARCHAR2, p_ShouldDate2 IN VARCHAR2, p_FileName IN VARCHAR2, p_IsProdCharge IN NUMBER, p_RetMsg OUT VARCHAR2) RETURN NUMBER
        With ComSF_TMPTOC1
                .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
                .Parameters.Append .CreateParameter("P_USERID", adVarChar, adParamInput, 2000, p_UserId)
                .Parameters.Append .CreateParameter("P_COMPCODE", adVarNumeric, adParamInput, , p_CompCode)
                .Parameters.Append .CreateParameter("P_SERVICETYPE", adVarChar, adParamInput, 2000, p_ServiceType)
                .Parameters.Append .CreateParameter("p_ShouldDate1", adVarChar, adParamInput, 2000, "")
                .Parameters.Append .CreateParameter("p_ShouldDate2", adVarChar, adParamInput, 2000, "")
                .Parameters.Append .CreateParameter("p_FileName", adVarChar, adParamInput, 2000, p_FileName)
                .Parameters.Append .CreateParameter("p_IsProdCharge", adVarNumeric, adParamInput, , 0)
                .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, P_RETMSG)
                Set .ActiveConnection = cnConn
                .CommandText = "SF_TMPTOC1"
                .CommandType = adCmdStoredProc
                .Execute
                P_RETMSG = .Parameters("P_RETMSG").Value & ""
                SF_TmpToC1 = Val(.Parameters("Return_value").Value & "")
        End With
        Exit Function
ChkErr:
        ErrSub "clsStoreFunction", "SF_TMPTOC1"
End Function

Private Sub Form_Activate()
    On Error Resume Next
        SetPara
        Screen.MousePointer = vbDefault
        If blnAutoExec Then
            cmdOk.Value = True
        End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call cmdOk_Click: Exit Sub
End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr
        If Alfa2 = True Then
            Call GetGlobal
        End If
        Call subGil
        Call InitData
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub InitData()
    On Error Resume Next
        intPara24 = Val(GetSystemParaItem("Para24", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043") & "")
        If intPara24 = 1 Then
            gilServiceType.Clear
            lblServiceType.Visible = False
            gilServiceType.Visible = False
        Else
            lblServiceType.Visible = True
            gilServiceType.Visible = True
        End If
End Sub

Private Sub subGil()
    On Error Resume Next
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode)
        SetgiList gilServiceType, "CodeNo", "Description", "CD046"
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3200", "SO3210") Then Exit Sub
        subGil
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.SetCodeNo ""
        gilServiceType.Query_Description
        gilServiceType.ListIndex = 1
        Call InitData
End Sub

Private Sub gilServiceType_Change()
    Call InitData
End Sub
