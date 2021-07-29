VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO3313B 
   BorderStyle     =   1  '單線固定
   Caption         =   "信用卡自動扣款[SO3313B]"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3313B.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7185
   StartUpPosition =   1  '所屬視窗中央
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   315
      Left            =   120
      TabIndex        =   24
      Top             =   2730
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   556
      ButtonCaption   =   "收費項目"
   End
   Begin VB.ComboBox cboBillType 
      ForeColor       =   &H000000C0&
      Height          =   315
      ItemData        =   "SO3313B.frx":0442
      Left            =   2160
      List            =   "SO3313B.frx":044C
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   1236
      Width           =   2085
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   465
      Left            =   4290
      TabIndex        =   12
      Top             =   4080
      Width           =   1410
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      Height          =   465
      Left            =   1365
      TabIndex        =   11
      Top             =   4080
      Width           =   1410
   End
   Begin VB.TextBox txtPath 
      Height          =   360
      Left            =   2775
      TabIndex        =   9
      Top             =   3570
      Width           =   2925
   End
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
      Left            =   5775
      TabIndex        =   10
      Top             =   3570
      Width           =   435
   End
   Begin VB.ComboBox cboBillHeadFmt 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2160
      Style           =   2  '單純下拉式
      TabIndex        =   1
      Top             =   525
      Visible         =   0   'False
      Width           =   3555
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   150
      Width           =   3570
      _ExtentX        =   6297
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
      FldWidth2       =   2500
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilBankCode 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   870
      Width           =   3570
      _ExtentX        =   6297
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
      FldWidth2       =   2500
      F2Corresponding =   -1  'True
   End
   Begin MSComDlg.CommonDialog comdPath 
      Left            =   345
      Top             =   3450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   510
      Width           =   3570
      _ExtentX        =   6297
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
      FldWidth2       =   2500
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilPTCode 
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Top             =   1605
      Width           =   3570
      _ExtentX        =   6297
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
      FldWidth2       =   2500
      F2Corresponding =   -1  'True
   End
   Begin Gi_Date.GiDate gdaRealDate 
      Height          =   405
      Left            =   1935
      TabIndex        =   8
      Top             =   3105
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
   Begin prjGiList.GiList gilClctEn 
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Top             =   1965
      Width           =   3585
      _ExtentX        =   6324
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
      FldWidth1       =   1000
      FldWidth2       =   2250
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilCMCode 
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      Top             =   2325
      Width           =   3570
      _ExtentX        =   6297
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
      FldWidth2       =   2500
      F2Corresponding =   -1  'True
   End
   Begin VB.Label lblBankCode 
      AutoSize        =   -1  'True
      Caption         =   "信用卡扣款銀行"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   390
      TabIndex        =   23
      Top             =   944
      Width           =   1365
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "公司別"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   390
      TabIndex        =   22
      Top             =   210
      Width           =   585
   End
   Begin VB.Label lblServiceType 
      AutoSize        =   -1  'True
      Caption         =   "服務類別"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   390
      TabIndex        =   21
      Top             =   577
      Width           =   780
   End
   Begin VB.Label lblPTCode 
      AutoSize        =   -1  'True
      Caption         =   "付款種類"
      Height          =   195
      Left            =   390
      TabIndex        =   20
      Top             =   1678
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "入帳格式"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   390
      TabIndex        =   19
      Top             =   1311
      Width           =   780
   End
   Begin VB.Label lblCMCode 
      AutoSize        =   -1  'True
      Caption         =   "收費方式"
      Height          =   195
      Left            =   390
      TabIndex        =   18
      Top             =   2415
      Width           =   780
   End
   Begin VB.Label lblClctEn 
      AutoSize        =   -1  'True
      Caption         =   "收費人員"
      Height          =   195
      Left            =   390
      TabIndex        =   17
      Top             =   2045
      Width           =   780
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "資料來源（路徑+檔名）："
      Height          =   195
      Left            =   405
      TabIndex        =   16
      Top             =   3660
      Width           =   2250
   End
   Begin VB.Label lblAccount 
      AutoSize        =   -1  'True
      Caption         =   "（若空白，則以扣帳日為入帳日）"
      Height          =   195
      Left            =   3150
      TabIndex        =   15
      Top             =   3195
      Width           =   2925
   End
   Begin VB.Label lblDate1 
      AutoSize        =   -1  'True
      Caption         =   "收款入帳日期："
      Height          =   195
      Left            =   405
      TabIndex        =   14
      Top             =   3195
      Width           =   1365
   End
   Begin VB.Label lblBillHeadFmt 
      AutoSize        =   -1  'True
      Caption         =   "多帳戶產生依據設定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   390
      TabIndex        =   13
      Top             =   577
      Width           =   1620
   End
End
Attribute VB_Name = "frmSO3313B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strPrgName As String        '轉帳程式名稱
Private rsData As New ADODB.Recordset
Private strFilePath As String       '記錄路徑名稱
Private objCreditCardIn As Object
Private objAction As Object
Dim rsBillHeadFmt As New ADODB.Recordset
Dim intPara24 As Integer

Private Sub CboAddItem()
    On Error Resume Next
        'If cboBillHeadFmt.ListCount > 0 Then Exit Sub
        If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr From " & GetOwner & "CD068") Then Exit Sub
        Do While Not rsBillHeadFmt.EOF
            cboBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
            rsBillHeadFmt.MoveNext
        Loop
End Sub
Private Sub cboBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
    Dim strCD019SQL As String
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
        '#7049 改用CD068A.CitemCode By Kin 2015/07/17
        strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                                    " And CD068.BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1"
        If Not GetRS(rsCD019, strCD019SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        'If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        gimCitemCode.Filter = gimCitemCode.Filter
        gimCitemCode.SetQueryCode strCitemCode
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
        Unload Me
    Exit Sub
End Sub
Private Sub subGim()
    On Error Resume Next
        '收費項目
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱", , True
End Sub

Private Sub cmdOK_Click()
    On Error GoTo chkErr
 
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    Set objCreditCardIn = CreateObject("CreditCardIn.clsInterface")
    With objCreditCardIn
        .FilePath = strFilePath                          'INI檔案路徑
        .prgName = strPrgName                            '轉帳程式名稱
        .SourcePath = txtPath.Text                       '原始檔案路徑
        .RealDate = gdaRealDate.GetValue                 '入帳日期
        .UpdEn = garyGi(0)                               '異動人員
        .Updname = garyGi(1)                             '異動人員
        .PTCode = gilPTCode.GetCodeNo & ""               '付款種類
        .PTName = gilPTCode.GetDescription & ""
        .ServiceType = gilServiceType.GetCodeNo & ""     '服務類別
        .CompCode = gilCompCode.GetCodeNo & ""
        .BankId = gilBankCode.GetCodeNo & ""
        .BankName = gilBankCode.GetDescription & ""
        .pOwnerName = GetOwner
        '#7049 改用CD068A.CitemCode By Kin 2015/07/17
        If cboBillHeadFmt.Visible Then
             .uCitemQty = " IN (Select CitemCode From " & GetOwner & "CD068A Where BillHeadFmt='" & cboBillHeadFmt.Text & "') "
        Else
            .uCitemQty = gimCitemCode.GetQueryCode           '收費項目
        End If
        .uClctEn = gilClctEn.GetCodeNo
        .uClctName = gilClctEn.GetDescription
        .uCMCode = gilCMCode.GetCodeNo
        .uCMName = gilCMCode.GetDescription
        If cboBillType.ListIndex = 1 Then
            .ExportFormat = 1
        End If
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
    
    If objCreditCardIn.UpdState = False Then Exit Sub
    Call DefineRs
    Call ScrToRcd
    Set objCreditCardIn = Nothing
        
   Exit Sub
chkErr:
    If Err.Number = 91 Then
       MsgBox "轉帳程式名稱錯誤或者該銀行沒有轉帳資料產生功能!", vbExclamation, "警告..."
       Exit Sub
    End If
   Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Sub ScrToRcd()
On Error GoTo chkErr
    With rsData
       .AddNew
       .Fields("RealDate") = gdaRealDate.GetValue
       .Fields("FilePath") = txtPath.Text
       If Dir(strFilePath & "\" & strPrgName & "In1.log") <> "" Then
          Kill strFilePath & "\" & strPrgName & "In1.log"
       End If
          .Save strFilePath & "\" & strPrgName & "In1.log"
    End With
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub

Private Sub cmdPath_Click()
    On Error GoTo chkErr
        With comdPath
                .FileName = txtPath.Text
                If cboBillType.ListIndex = 1 Then
                    .Filter = "回覆檔(*.*)|*.*"
                Else
                    .Filter = "回覆檔 (*.aut)|*.aut|所有檔案 (*.*)|*.*"
                End If
                .InitDir = ReadGICMIS1("ErrLogPath")
                .Action = 1
                txtPath.Text = .FileName
        End With
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdPath_Click")
End Sub

Private Function IsDataOk() As Boolean
    On Error GoTo chkErr
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
           strErrMsg = "信用卡扣款銀行"
           gilBankCode.SetFocus
           GoTo 66
        End If
        If Len(txtPath) = 0 Then MsgBox "請輸入檔名", vbExclamation, "提示": txtPath.SetFocus: Exit Function
        If Dir(Trim(txtPath.Text)) = "" Then MsgBox "扣款資料不存在，請重新選取！", vbExclamation, "訊息！": Exit Function
        IsDataOk = True
    Exit Function
66:
     Call MsgMustBe(strErrMsg)
   Exit Function
chkErr:
    Call ErrSub(Me.Name, "IsdataOK")
End Function


Private Sub Form_Activate()
On Error GoTo chkErr
  Dim strSQL As String
  Dim rsTmp As New ADODB.Recordset
    gilCompCode.SetFocus
    strSQL = "Select PrgName From " & GetOwner & "CD018 Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'"
    If Not GetRS(rsTmp, strSQL) Then Exit Sub
    If rsTmp.EOF Then
       MsgBox "請設定銀行代碼中轉帳程式名稱!", vbExclamation, "警告"
       Unload Me
    End If
  Exit Sub
chkErr:
  ErrSub Me.Name, "Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyEscape Then Unload Me
        If (KeyCode = vbKeyF2) Or (KeyCode = vbKeyReturn) Then
            If Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "") Then
                gdaRealDate.SetFocus
                Exit Sub
            End If
            If cmdOK.Enabled = True Then cmdOK.Value = True
        End If
End Sub

Private Sub Form_Load()
    On Error GoTo chkErr
        Screen.MousePointer = vbDefault
        Call subGil
        Call subGim
        Call getDefault
        CboAddItem
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub DefineRs()
On Error GoTo chkErr
  If rsData.State = adStateOpen Then rsData.Close
    With rsData
       .LockType = adLockOptimistic
       .Fields.Append "RealDate", adBSTR, 8
       .Fields.Append "FilePath", adBSTR, 50
       .Open
    End With
  Exit Sub
chkErr:
  ErrSub Me.Name, "DefineRS"
End Sub

Private Sub GetInitData()
    On Error GoTo chkErr
   
        strFilePath = ReadGICMIS1("ErrLogPath")
        If Dir(strFilePath & "\" & GetPrgName & "In1.log") = "" Then
            gdaRealDate.SetValue ""
            txtPath.Text = ""
        Else
            If rsData.State = adStateOpen Then rsData.Close
            rsData.Open strFilePath & "\" & strPrgName & "In1.log"
            gdaRealDate.SetValue rsData("RealDate").Value & ""
            txtPath.Text = rsData("FilePath").Value & ""
        End If
   Exit Sub
chkErr:
    Call ErrSub(Me.Name, "GetInitData")
End Sub

Private Function GetPrgName() As String
  On Error GoTo chkErr
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
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "GetPrgName")
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   If rsData.State = adStateOpen Then
      rsData.Close
      Set rsData = Nothing
   End If
   Call ReleaseCOM(Me)
End Sub

Private Sub gdaRealDate_Validate(Cancel As Boolean)
    On Error Resume Next
        '#3411 如果未輸入則不檢查 By Kin 2008/06/02
        If gdaRealDate.GetValue = "" Then Exit Sub
        Cancel = Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "")
End Sub

Private Sub gilBankCode_Change()
  On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
       If Len(gilBankCode.GetCodeNo) > 0 And Len(gilBankCode.GetDescription) > 0 Then
           GetInitData
       End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "gilBankCode_Change")
End Sub

Private Sub subGil()
    On Error GoTo chkErr
        SetgiList gilBankCode, "CodeNo", "Description", "CD018", , , , , 3, 24, "Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'"
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
        SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 3, 12
        SetgiList gilPTCode, "CodeNo", "Description", "CD032", , , , , 3, 12
        SetgiList gilCMCode, "CodeNo", "Description", "CD031", , , , , 3, 20, , True
        SetgiList gilClctEn, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGil"
End Sub

Private Sub gilCompCode_Change()
On Error GoTo chkErr
    garyGi(16) = gilCompCode.GetOwner
    'MsgBox garyGi(16)
    'Call subGil
    If Len(gilCompCode.GetCodeNo & "") = 0 Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 99
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    
    GiListFilter gilBankCode, , gilCompCode.GetCodeNo
    gilBankCode.Filter = gilBankCode.Filter & IIf(gilBankCode.Filter = "", " Where ", " And ") & " UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'"
    Call subGil

  Exit Sub
99:
  MsgMustBe (strErrFile)
  Exit Sub
chkErr:
  ErrSub Me.Name, "gilCompCode_Change"
End Sub

Private Sub gilServiceType_Change()
On Error GoTo chkErr
Dim strWhere As String
           
'    If Len(gilServiceType.GetCodeNo & "") = 0 Then
'       strWhere = ""
'    Else
'       strWhere = " Where ServiceType = '" & gilServiceType.GetCodeNo & "'"
'    End If
'    gilPTCode.Filter = strWhere
    Call GiListFilter(gilPTCode, gilServiceType.GetCodeNo, gilCompCode.GetCodeNo)
  Exit Sub
chkErr:
  ErrSub Me.Name, "gilServiceType_Change"
End Sub

Private Sub getDefault()
  On Error GoTo chkErr
     Dim intPara23 As Integer
     gilCompCode.SetCodeNo garyGi(9)
     gilCompCode.Query_Description
        intPara24 = GetSystemParaItem("Para24", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043")
'        intpara24 = 0
        intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043") & "")
        If intPara24 = 1 Then
            lblBillHeadFmt.Visible = True
            cboBillHeadFmt.Visible = True
            '  用媒體入帳才Lock
'            If intPara23 = 2 Then gimCitemCode.Enabled = False
            If intPara23 = 2 Then gimCitemCode.IsReadOnly = True
            gilServiceType.SetCodeNo ""
            gilServiceType.Query_Description
            gilServiceType.Enabled = False
            gilServiceType.Visible = False
            lblServiceType.Visible = False
            '#7049 如果有多帳戶產生依據設定 則要將收費項目鎖定 By Kin 2015/07/17
            'gimCitemCode.Enabled = False
            gimCitemCode.IsReadOnly = True
        Else
            lblBillHeadFmt.Visible = False
            gimCitemCode.Enabled = True
        End If
        cboBillType.ListIndex = 0
  Exit Sub
chkErr:
  ErrSub Me.Name, "getDefault"
End Sub


Private Sub gimCitemCodeX_Click()

End Sub
