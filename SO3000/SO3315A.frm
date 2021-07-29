VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSO3315A 
   Caption         =   "金融格式轉帳資料登錄 [SO3315A]"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3315A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5355
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   375
      Left            =   180
      TabIndex        =   27
      Top             =   2130
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      ButtonCaption   =   "收費項目"
   End
   Begin VB.CheckBox chkZero 
      Caption         =   "收費項目金額<=0是否產生"
      Height          =   495
      Left            =   5340
      TabIndex        =   6
      Top             =   1560
      Value           =   1  '核取
      Width           =   1875
   End
   Begin VB.CheckBox chkNT 
      Caption         =   "南資專用"
      Height          =   315
      Left            =   5340
      TabIndex        =   5
      Top             =   1305
      Width           =   1290
   End
   Begin VB.ComboBox cboBillHeadFmt 
      Height          =   315
      Left            =   2025
      Style           =   2  '單純下拉式
      TabIndex        =   25
      Top             =   480
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.CheckBox chkBeOne 
      Caption         =   "為聯合中心"
      Height          =   315
      Left            =   5340
      TabIndex        =   3
      Top             =   960
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      Height          =   1125
      Left            =   330
      ScaleHeight     =   1065
      ScaleWidth      =   6255
      TabIndex        =   19
      Top             =   3480
      Width           =   6315
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "（若收費方式空白，則以原收費方式）"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   720
         TabIndex        =   23
         Top             =   150
         Width           =   3315
      End
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         Caption         =   "說明:"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   30
         TabIndex        =   22
         Top             =   60
         Width           =   540
      End
      Begin VB.Label lblAccount 
         AutoSize        =   -1  'True
         Caption         =   "（若收款入帳日期空白，則以扣帳日為入帳日）"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   720
         TabIndex        =   21
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "（若收費人員空白，則以原收費人員）"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   720
         TabIndex        =   20
         Top             =   810
         Width           =   3315
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   435
      Left            =   4215
      TabIndex        =   12
      Top             =   4800
      Width           =   1410
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      Default         =   -1  'True
      Height          =   465
      Left            =   1290
      TabIndex        =   11
      Top             =   4800
      Width           =   1410
   End
   Begin VB.TextBox txtPath 
      Height          =   360
      Left            =   3060
      TabIndex        =   9
      Top             =   3000
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
      Left            =   6060
      TabIndex        =   10
      Top             =   3000
      Width           =   435
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   2025
      TabIndex        =   0
      Top             =   90
      Width           =   3270
      _ExtentX        =   5768
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
      FldWidth2       =   2200
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilBankCode 
      Height          =   315
      Left            =   2025
      TabIndex        =   2
      Top             =   900
      Width           =   3270
      _ExtentX        =   5768
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
      FldWidth2       =   2200
      F2Corresponding =   -1  'True
   End
   Begin MSComDlg.CommonDialog comdPath 
      Left            =   270
      Top             =   4770
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   2025
      TabIndex        =   1
      Top             =   495
      Width           =   3270
      _ExtentX        =   5768
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
      FldWidth2       =   2200
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilClctEn 
      Height          =   315
      Left            =   2025
      TabIndex        =   4
      Top             =   1305
      Width           =   3255
      _ExtentX        =   5741
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
      FldWidth2       =   1930
      F2Corresponding =   -1  'True
   End
   Begin Gi_Date.GiDate gdaRealDate 
      Height          =   375
      Left            =   2235
      TabIndex        =   8
      Top             =   2540
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
   Begin prjGiList.GiList gilCMCode 
      Height          =   315
      Left            =   2025
      TabIndex        =   7
      Top             =   1725
      Width           =   3270
      _ExtentX        =   5768
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
      FldWidth2       =   2200
      F2Corresponding =   -1  'True
   End
   Begin VB.Label lblBillHeadFmt 
      AutoSize        =   -1  'True
      Caption         =   "多帳戶產生依據設定"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   26
      Top             =   560
      Width           =   1755
   End
   Begin VB.Label lblCMCode 
      AutoSize        =   -1  'True
      Caption         =   "收費方式"
      Height          =   195
      Left            =   210
      TabIndex        =   24
      Top             =   1800
      Width           =   780
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "資料來源（路徑+檔名）："
      Height          =   195
      Left            =   690
      TabIndex        =   18
      Top             =   3120
      Width           =   2250
   End
   Begin VB.Label lblDate1 
      AutoSize        =   -1  'True
      Caption         =   "收款入帳日期："
      Height          =   195
      Left            =   690
      TabIndex        =   17
      Top             =   2760
      Width           =   1365
   End
   Begin VB.Label lblClctEn 
      AutoSize        =   -1  'True
      Caption         =   "收費人員"
      Height          =   195
      Left            =   210
      TabIndex        =   16
      Top             =   1320
      Width           =   780
   End
   Begin VB.Label lblServiceType 
      AutoSize        =   -1  'True
      Caption         =   "服務類別"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   555
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "公司別"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   150
      Width           =   585
   End
   Begin VB.Label lblBankCode 
      AutoSize        =   -1  'True
      Caption         =   "轉帳代收銀行"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   210
      TabIndex        =   13
      Top             =   960
      Width           =   1170
   End
End
Attribute VB_Name = "frmSO3315A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strPrgName As String        '轉帳程式名稱
Private rsSo3315 As New ADODB.Recordset
Private strFilePath As String       '記錄路徑名稱
Private objAgencyIn As Object
Private objAction As Object
Private strChoose As String
Dim rsBillHeadFmt As New ADODB.Recordset
Dim intPara24 As Integer
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
Private Sub CboAddItem()
    On Error Resume Next
        cboBillHeadFmt.Clear
        If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr From " & GetOwner & "CD068") Then Exit Sub
        Do While Not rsBillHeadFmt.EOF
            cboBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
            rsBillHeadFmt.MoveNext
        Loop
End Sub

Private Sub chkBeOne_Click()
    On Error Resume Next
        If chkBeOne.Value = 1 Then
            chkNT.Value = 0
            gilClctEn.Locked = True
            '問題集1295 20050103 for minchen 三冠王要求選聯合中心時收費方式可以提供選擇
'            gilCMCode.Locked = True
        Else
            gilClctEn.Locked = False
'            gilCMCode.Locked = False
        End If
End Sub

Private Sub chkNT_Click()
    If chkNT.Value = 1 Then
        chkBeOne.Value = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
        Unload Me
    Exit Sub
End Sub

Private Sub cmdOk_Click()
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim objAction As Object
    Dim objInterface As Object
        If IsDataOk = False Then Exit Sub
        Set objInterface = New clsInterface
        Screen.MousePointer = vbHourglass
        With objInterface
            .uBeOne = IIf(chkBeOne.Value = 1, True, False)
            .uCompCode = gilCompCode.GetCodeNo
            .uServiceType = gilServiceType.GetCodeNo
            .uErrPath = ReadGICMIS1("ErrLogPath")
            Set .ugcnGi = gcnGi
            .uBankSQL = " = '" & gilBankCode.GetCodeNo & "'"
            .uBankId = gilBankCode.GetCodeNo & ""
            .uPTCode = ""
            .uPTName = ""
            .uUpdEn = garyGi(0)
            .uUpdName = garyGi(1)
            .uCMCode = gilCMCode.GetCodeNo
            .uCMName = gilCMCode.GetDescription
            .uClctEn = gilClctEn.GetCodeNo
            .uClctName = gilClctEn.GetDescription
            .uRealDate = gdaRealDate.GetValue
            '#7049 改用CD068A.CitemCode By Kin 2015/07/17
            If cboBillHeadFmt.Visible Then
                .uCitemQty = " Select CitemCode From " & GetOwner & "CD068A Where BillHeadFmt = '" & cboBillHeadFmt.Text & "' "
            Else
                .uCitemQty = gimCitemCode.GetQueryCode
            End If
            .uFilePath = txtPath.Text
            .uGetOwner = GetOwner
            .FlowId = garyGi(17)
            .uChkNT = chkNT.Value
            .uCkZero = IIf(chkZero.Value = 1, True, False)
            Set objAction = New clsBankCenterIn
            Call objAction.Action
        End With
        Set objAction = Nothing
        Set objInterface = Nothing
        Call ScrToRcd
        Screen.MousePointer = vbDefault
        Unload Me
    Exit Sub
chkErr:
    If Err.Number = 20 Then Resume Next
    If Err.Number = 91 Then
       MsgBox "轉帳程式名稱錯誤或者該銀行沒有轉帳資料產生功能!", vbExclamation, "警告..."
       Exit Sub
    End If
    Call ErrSub(Me.Name, "cmdOk_Click")
End Sub

Private Sub ScrToRcd()
    On Error GoTo chkErr
        With rsSo3315
            If .EOF Or .BOF Then
                .AddNew
            End If
            .Fields("RealDate") = gdaRealDate.GetValue
            .Fields("FilePath") = txtPath.Text
            If Dir(GetLogFileName(strFilePath)) <> "" Then
                Kill GetLogFileName(strFilePath)
            End If
            .Save GetLogFileName(strFilePath)
        End With
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub

Private Sub cmdPath_Click()
    On Error GoTo chkErr
        With comdPath
                .FileName = txtPath.Text
                .Filter = "文字檔 (*.txt)|*.txt|所有檔案 (*.*)|*.*"
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
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        If Not MustExist(gilBankCode, 2, "轉帳代收銀行") Then Exit Function
        If Not MustExist(txtPath, , "資料來源") Then Exit Function
        If Not ChkDirExist(Trim(txtPath.Text)) Then MsgBox "代收資料不存在，請重新選取！", vbExclamation, "訊息！": Exit Function
        
        IsDataOk = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "IsdataOK")
End Function


Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyEscape Then Unload Me
        If KeyCode = vbKeyF2 Then If cmdOK.Enabled = True Then cmdOK.Value = True
End Sub

Private Sub Form_Load()
    On Error GoTo chkErr
        Call DefineRS
        Call subGil
        Call subGim
        Call DefaultValue
        CboAddItem
        Call OpenData
        
        
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Function DefaultValue()
    On Error Resume Next
    Dim intPara23 As Integer
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        intPara24 = GetSystemParaItem("Para24", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043")
        intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043") & "")
        If intPara24 = 1 Then
            lblBillHeadFmt.Visible = True
            cboBillHeadFmt.Visible = True
            '  用媒體入帳才Lock
            'If intPara23 = 2 Then gimCitemCode.Enabled = False
            If intPara23 = 2 Then gimCitemCode.IsReadOnly = True
            gilServiceType.SetCodeNo ""
            gilServiceType.Query_Description
            gilServiceType.Enabled = False
            gilServiceType.Visible = False
            lblServiceType.Visible = False
            '#7049
            'gimCitemCode.Enabled = False
            gimCitemCode.IsReadOnly = True
        Else
            lblBillHeadFmt.Visible = False
            gimCitemCode.Enabled = True
        End If
        
        Call GetInitData
End Function

Private Function OpenData() As Boolean
    On Error GoTo chkErr
    Exit Function
chkErr:
    ErrSub Me.Name, "OpenData"
End Function

Private Sub DefineRS()
On Error GoTo chkErr
  If rsSo3315.State = adStateOpen Then rsSo3315.Close
    With rsSo3315
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
        If Trim(GetLogFileName(strFilePath)) = "" Then Exit Sub
        If Not ChkDirExist(GetLogFileName(strFilePath)) Then
            gdaRealDate.SetValue ""
            txtPath.Text = ""
        Else
            If rsSo3315.State = adStateOpen Then rsSo3315.Close
            rsSo3315.Open GetLogFileName(strFilePath)
            gdaRealDate.SetValue rsSo3315("RealDate").Value & ""
            txtPath.Text = rsSo3315("FilePath").Value & ""
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
        If Not GetRS(rsTmp, strSQL, gcnGi) Then Exit Function
        
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
   If rsSo3315.State = adStateOpen Then
      rsSo3315.Close
      Set rsSo3315 = Nothing
   End If
   Call ReleaseCOM(Me)
End Sub

Private Sub gdaRealDate_Validate(Cancel As Boolean)
    On Error Resume Next
        '#3411 如果沒有輸入要能離開 By Kin 2008/06/20
        If gdaRealDate.GetValue & "" = "" Then Exit Sub
        Cancel = Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "")
End Sub

Private Sub gilBankCode_Change()
  On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strEmpNo As String
        If Len(gilBankCode.GetCodeNo) > 0 And Len(gilBankCode.GetDescription) > 0 Then
            GetInitData
            strEmpNo = GetRsValue("Select EmpNo From " & GetOwner & "CD018 Where CodeNo =" & gilBankCode.GetCodeNo) & ""
            If strEmpNo <> "" Then
                gilClctEn.SetCodeNo strEmpNo
                gilClctEn.Query_Description
            End If
            '#3517 如果CD018.REFNO=13,則南資的選項要隱藏 By Kin 2007/10/29
            If GetRsValue("Select Count(*) Cnt From " & GetOwner & "CD018 Where CodeNo=" & _
                          gilBankCode.GetCodeNo & " And RefNo=13") > 0 Then
                chkNT.Visible = False
            Else
                chkNT.Visible = True
            End If
        End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "gilBankCode_Change")
End Sub

Private Sub subGil()
    On Error GoTo chkErr
        SetgiList gilBankCode, "CodeNo", "Description", "CD018", , , , , , , " Where UPPER(PrgName) = 'BANKCENTER'", True
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode)
        SetgiList gilServiceType, "CodeNo", "Description", "CD046"
        SetgiList gilCMCode, "CodeNo", "Description", "CD031", , , , , 3, 20, , True
        SetgiList gilClctEn, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True
        GiListFilter gilCMCode, gilServiceType.GetCodeNo
        
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGil"
End Sub

Private Sub gilCompCode_Change()
    On Error GoTo chkErr
    garyGi(16) = gilCompCode.GetOwner
    'MsgBox garyGi(16)
    Dim strErrFile As String
        If Len(gilCompCode.GetCodeNo & "") = 0 Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 99
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.ListIndex = 1
        GiListFilter gilBankCode, , gilCompCode.GetCodeNo
        gilBankCode.Filter = gilBankCode.Filter & IIf(gilBankCode.Filter = "", " Where ", " And ") & " UPPER(PrgName) = 'BANKCENTER'"
        Call subGil
        Call subGim
        CboAddItem
        cboBillHeadFmt.ListIndex = -1
        gimCitemCode.Clear
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
        If Len(gilServiceType.GetCodeNo & "") = 0 Then Exit Sub
        If intPara24 = 0 Then
            GiListFilter gilCMCode, "Where servicetype in " & gilServiceType.GetCodeNo
        End If
'        GiListFilter gilCMCode, gilServiceType.GetCodeNo
  Exit Sub
chkErr:
  ErrSub Me.Name, "gilServiceType_Change"
End Sub

Private Function GetLogFileName(strPath As String) As String
    On Error Resume Next
        GetLogFileName = strPath & IIf(Right(strPath, 1) = "\", "", "\") & "BankCenterIn.log"
End Function

Public Function GetOwner(Optional strOwner As String) As String
  On Error Resume Next
    If strOwner = "" Then strOwner = garyGi(16)
    If strOwner <> "" Then GetOwner = strOwner & "."
End Function

Private Sub subGim()
    On Error Resume Next
        '收費項目
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱", , True
End Sub

Public Function ChkCloseDate(strCloseDate As String, _
    strCompCode As String, strServiceType As String) As Boolean
    On Error GoTo chkErr
    Dim intDayCut As Integer
    Dim strTranDate As String
    Dim intPara6 As Integer
        intDayCut = Val(GetSystemParaItem("DayCut", strCompCode, "", "SO041") & "")
        strTranDate = GetRsValue("Select tranDate from " & GetOwner & "So062 Where CompCode= " & strCompCode & " And Type=1 Order By TranDate Desc") & ""
        '收費參數檔（SO043)，取收費日登錄期限<Para6>，供後續使用！
        intPara6 = Val(GetSystemParaItem("Para6", strCompCode, strServiceType, "SO043") & "")
    
        If strCloseDate = "" Then If vbNo = MsgBox("本欄位是否為空值！", vbExclamation, "訊息！") Then Exit Function
        
        If Not IsDate(strCloseDate) Then
            MsgBox "日期不合法！", vbExclamation, "訊息！"
            Exit Function
        End If
        
        If CDate(strCloseDate) > RightDate Then MsgBox "此日期超過今天日期！", vbExclamation, "訊息！": Exit Function
        
        If (RightDate - CDate(strCloseDate)) > intPara6 Then
            MsgBox "此日期已超過系統設定的安全期限！", vbExclamation, "訊息！":  Exit Function
        End If
        If CDate(strCloseDate) <= IIf(strTranDate <> "", CDate(strTranDate), 0) Then
            If intDayCut = 1 And CDate(strCloseDate) = IIf(strTranDate <> "", CDate(strTranDate), 0) Then ChkCloseDate = True: Exit Function
            MsgBox strTranDate & "之前已做過日結，不可使用之前日期入帳", vbExclamation, "訊息！"
            Exit Function
        End If
        ChkCloseDate = True
    Exit Function
chkErr:
    ErrSub Me.Name, "ChkCloseDate"
End Function


