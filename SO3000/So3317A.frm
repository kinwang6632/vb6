VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSO3317A 
   BorderStyle     =   1  '單線固定
   Caption         =   "ATM轉帳資料過帳作業 [SO3317A]"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "So3317A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5820
      TabIndex        =   6
      Top             =   2760
      Width           =   435
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2805
      TabIndex        =   5
      Top             =   2760
      Width           =   2925
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1395
      TabIndex        =   7
      Top             =   3600
      Width           =   1410
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
      Height          =   465
      Left            =   4320
      TabIndex        =   8
      Top             =   3600
      Width           =   1410
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1950
      TabIndex        =   0
      Top             =   360
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
   Begin prjGiList.GiList gilBankCode 
      Height          =   315
      Left            =   1950
      TabIndex        =   2
      Top             =   1260
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
      Left            =   1950
      TabIndex        =   1
      Top             =   825
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
   Begin prjGiList.GiList gilPTCode 
      Height          =   315
      Left            =   1950
      TabIndex        =   3
      Top             =   1695
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
   Begin Gi_Date.GiDate gdaRealDate 
      Height          =   405
      Left            =   1950
      TabIndex        =   4
      Top             =   2265
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
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
   Begin MSComDlg.CommonDialog comdPath 
      Left            =   165
      Top             =   3420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblBankCode 
      AutoSize        =   -1  'True
      Caption         =   "轉帳代收銀行"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   435
      TabIndex        =   15
      Top             =   1320
      Width           =   1170
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   435
      TabIndex        =   14
      Top             =   420
      Width           =   585
   End
   Begin VB.Label lblServiceType 
      AutoSize        =   -1  'True
      Caption         =   "服務類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   435
      TabIndex        =   13
      Top             =   885
      Width           =   780
   End
   Begin VB.Label lblPTCode 
      AutoSize        =   -1  'True
      Caption         =   "付款種類"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   435
      TabIndex        =   12
      Top             =   1755
      Width           =   780
   End
   Begin VB.Label lblDate1 
      AutoSize        =   -1  'True
      Caption         =   "收款入帳日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   435
      TabIndex        =   11
      Top             =   2355
      Width           =   1365
   End
   Begin VB.Label lblAccount 
      AutoSize        =   -1  'True
      Caption         =   "（若空白，則以扣帳日為入帳日）"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3180
      TabIndex        =   10
      Top             =   2355
      Width           =   2925
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "資料來源（路徑+檔名）："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   435
      TabIndex        =   9
      Top             =   2850
      Width           =   2250
   End
End
Attribute VB_Name = "frmSO3317A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strPrgName As String        '轉帳程式名稱
Private rsSo3317 As New ADODB.Recordset
Private strFilePath As String       '記錄路徑名稱
Private intState As Integer
Dim intPara24 As Integer

Private Sub cmdCancel_Click()
On Error Resume Next
   Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ChkErr
  Dim objATMIn As Object
  Dim objAction As Object
  Dim TmpPath As String
  Dim ErrFilePath As String
  
   If Not ChkDTok Then Exit Sub
   If Not IsDataOk Then Exit Sub
   Set objATMIn = CreateObject("ATMIn4.clsInterface")
   With objATMIn
        .FilePath = strFilePath                          'INI檔案路徑
        .PrgName = strPrgName                            '轉帳程式名稱
        .SourcePath = txtPath.Text                       '原始檔案路徑
        .RealDate = gdaRealDate.GetValue                 '入帳日期
        .UpdEn = garyGi(1)                               '異動人員
        .PTCode = gilPTCode.GetCodeNo & ""              '付款種類
        .PTName = gilPTCode.GetDescription & ""
        .ServiceType = gilServiceType.GetCodeNo & ""    '服務類別
        .COMPCODE = gilCompCode.GetCodeNo & ""
        .BankName = gilBankCode.GetDescription & ""
        .BankId = gilBankCode.GetCodeNo
        .uGetOwner = GetOwner
        Set .ugcnGi = gcnGi
        If Len(strPrgName & "") = 0 Then
            MsgBox "設定程式名稱無設或無使用權限！！", vbExclamation, "提示"
        Else
            'MsgBox ReadGICMIS1("ErrLogPath"), , "frmSO3273A"
            Set objAction = .InitObject(strPrgName & "")
            objAction.Action
            Set objAction = Nothing
        End If
   End With
   '若更新失敗，則關閉程式
   If objATMIn.UpdState = False Then Exit Sub
   Call DefineRs
   Call ScrToRcd
   Set objATMIn = Nothing
   
  Exit Sub
ChkErr:
    If Err.Number = 91 Then
       MsgBox "轉帳程式名稱錯誤或者該銀行沒有轉帳資料產生功能!", vbExclamation, "警告..."
       Exit Sub
    End If

  ErrSub Me.Name, "cmdOK_Click"
End Sub

Private Sub cmdPath_Click()
On Error GoTo ChkErr
    With comdPath
       .FileName = txtPath.Text
       .Filter = "文字檔(*.txt;*.dat)|*.txt;*.dat|所有檔案(*.*)|*.*"
       .InitDir = ReadGICMIS1("ErrLogPath")
       .Action = 1
       txtPath.Text = .FileName
    End With
 Exit Sub
ChkErr:
  ErrSub Me.Name, "cmdPath_Click"
End Sub

Private Sub ScrToRcd()
On Error GoTo ChkErr
    With rsSo3317

       .AddNew
       .Fields("RealDate") = gdaRealDate.GetValue
       .Fields("FilePath") = txtPath.Text
       If Dir(strFilePath & "\" & strPrgName & "In1.log") <> "" Then
          Kill strFilePath & "\" & strPrgName & "In1.log"
       End If
          .Save strFilePath & "\" & strPrgName & "In1.log"
    End With
  Exit Sub
ChkErr:
  ErrSub Me.Name, "ScrToRcd"
End Sub

Private Sub DefineRs()
On Error GoTo ChkErr
  If rsSo3317.State = adStateOpen Then rsSo3317.Close
    With rsSo3317
       .LockType = adLockOptimistic
       .Fields.Append "RealDate", adBSTR, 8
       .Fields.Append "FilePath", adBSTR, 50
       .Open
    End With
  Exit Sub
ChkErr:
  ErrSub Me.Name, "DefineRS"
End Sub
Private Sub DefaultValue()
On Error GoTo ChkErr
ChkErr:
    ErrSub Me.Name, "DefaultValue"
End Sub
Private Sub Form_Activate()
On Error GoTo ChkErr
  Dim strSQL As String
  Dim rsTmp As New ADODB.Recordset
    gilCompCode.SetFocus
    strSQL = "Select PrgName From " & GetOwner & "CD018 Where UPPER(SUBSTR(PrgName,1,3)) = '" & "ATM" & "'"
    If rsTmp.State = adStateOpen Then rsTmp.Close
'    rsTmp.Open strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
    If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then Exit Sub
    If rsTmp.EOF Then
       MsgBox "請設定銀行代碼中轉帳程式名稱!", vbExclamation, "警告"
       Unload Me
    End If
    If rsTmp.State = adStateOpen Then
       rsTmp.Close
       Set rsTmp = Nothing
    End If
  Exit Sub
ChkErr:
  ErrSub Me.Name, "Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If Not ChkGiList(KeyCode) Then Exit Sub
    If KeyCode = vbKeyEscape Then Unload Me
    If KeyCode = vbKeyF2 Then
        If Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "") Then
            gdaRealDate.SetFocus
            Exit Sub
        End If
        Call cmdOK_Click
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ChkErr
   Screen.MousePointer = vbDefault
   Call subGil
    intPara24 = GetSystemParaItem("Para24", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043")
     If intPara24 = 1 Then
            gilServiceType.SetCodeNo ""
            gilServiceType.Query_Description
            gilServiceType.Enabled = False
            lblServiceType.Enabled = False
     End If

 Exit Sub
ChkErr:
  ErrSub Me.Name, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   If rsSo3317.State = adStateOpen Then
      rsSo3317.Close
      Set rsSo3317 = Nothing
   End If
   Call ReleaseCOM(Me)
End Sub

Private Sub subGil()
On Error GoTo ChkErr
    SetgiList gilBankCode, "CodeNo", "Description", "CD018", , , , , 3, 24, " Where UPPER(SUBSTR(PrgName,1,3)) = '" & "ATM" & "'"
    SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
    SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 3, 12
    SetgiList gilPTCode, "CodeNo", "Description", "CD032", , , , , 3, 12
  Exit Sub
ChkErr:
  ErrSub Me.Name, "subGil"
End Sub

Private Sub GetInitData()
On Error GoTo ChkErr
   
    strFilePath = ReadGICMIS1("ErrLogPath")
    If Dir(strFilePath & "\" & GetPrgName & "In1.log") = "" Then
       gdaRealDate.SetValue ""
       txtPath.Text = ""
    Else
       If rsSo3317.State = adStateOpen Then rsSo3317.Close
       rsSo3317.Open strFilePath & "\" & strPrgName & "In1.log"
       gdaRealDate.SetValue rsSo3317("RealDate") & ""
       txtPath.Text = rsSo3317("FilePath") & ""
    End If
    
 Exit Sub
ChkErr:
  ErrSub Me.Name, "GetInitData"
End Sub

Private Function GetPrgName() As String
On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
  Dim strSQL As String
  
    strSQL = "Select PrgName From " & GetOwner & "CD018 Where CodeNo = " & gilBankCode.GetCodeNo
    Call OpenRecordset(rsTmp, strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False)
    
    If Len(rsTmp("PrgName") & "") = 0 Then
       strPrgName = ""
       MsgBox "請設定轉帳銀行代碼之程式名稱!", vbCritical, "錯誤"
    Else
       strPrgName = rsTmp("PrgName")
       GetPrgName = strPrgName
    End If
    If rsTmp.State = adStateOpen Then
       rsTmp.Close
       Set rsTmp = Nothing
    End If
  Exit Function
ChkErr:
  ErrSub Me.Name, "GetPrgName"
End Function

Private Sub gdaRealDate_Validate(Cancel As Boolean)
    On Error Resume Next
        '#3411 如果未輸入則不檢查 By Kin 2008/06/02
        If gdaRealDate.GetValue = Empty Then Exit Sub
        Cancel = Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "")
End Sub

Private Sub gilBankCode_Change()
On Error GoTo ChkErr
   If Len(gilBankCode.GetCodeNo & "") <> 0 And Len(gilBankCode.GetDescription & "") <> 0 Then
       Call GetInitData
   End If
  Exit Sub
ChkErr:
  ErrSub Me.Name, "gilBankCode_Change"
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo ChkErr
    Dim strErrMsg As String
    
        If Len(gilCompCode.GetCodeNo & "") = 0 Then
           strErrMsg = "公司別"
           gilCompCode.SetFocus
           GoTo 66
        End If
        If intPara24 = 0 And Len(gilServiceType.GetCodeNo & "") = 0 Then
           strErrMsg = "服務類別"
           gilServiceType.SetFocus
           GoTo 66
        End If
        If gilBankCode.GetCodeNo & "" = "" Then
           strErrMsg = "轉帳代收銀行"
           gilBankCode.SetFocus
           GoTo 66
        End If
        If Len(txtPath) = 0 Then MsgBox "請輸入檔名", vbExclamation, "提示": txtPath.SetFocus: Exit Function
        If Dir(Trim(txtPath.Text)) = "" Then MsgBox "代收資料不存在，請重新選取！", vbExclamation, "訊息！": Exit Function
        IsDataOk = True
    Exit Function
66:
     Call MsgMustBe(strErrMsg)
   Exit Function
ChkErr:
  ErrSub Me.Name, "IsDataOk"
End Function

Private Sub gilCompCode_Change()
On Error GoTo ChkErr
    garyGi(16) = gilCompCode.GetOwner
'    MsgBox garyGi(16)
    If Len(gilCompCode.GetCodeNo & "") = 0 Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 99
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    Call subGil
  Exit Sub
99:
  MsgMustBe (strErrFile)
  Exit Sub
ChkErr:
  ErrSub Me.Name, "gilCompCode_Change"
End Sub

Private Sub gilServiceType_Change()
On Error GoTo ChkErr
Dim strWhere As String
   
'    If Len(gilServiceType.GetCodeNo & "") = 0 Then
'       strWhere = ""
'    Else
'       strWhere = " Where ServiceType = '" & gilServiceType.GetCodeNo & "'"
'    End If
'    gilPTCode.Filter = strWhere
   Call GiListFilter(gilPTCode, gilServiceType.GetCodeNo, gilCompCode.GetCodeNo)
  Exit Sub
ChkErr:
  ErrSub Me.Name, "gilServiceType_Change"
End Sub
