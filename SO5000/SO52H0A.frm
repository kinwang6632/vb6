VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.17#0"; "csMulti.ocx"
Begin VB.Form frmSO52H0A 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶分期收費繳款表 [SO52H0A]"
   ClientHeight    =   5355
   ClientLeft      =   1950
   ClientTop       =   1530
   ClientWidth     =   7845
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO52H0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form15"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox txtBudgetPeriod2 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1710
      MaxLength       =   2
      TabIndex        =   3
      Top             =   510
      Width           =   405
   End
   Begin VB.TextBox txtBudgetPeriod1 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1020
      MaxLength       =   2
      TabIndex        =   2
      Top             =   510
      Width           =   405
   End
   Begin VB.Frame Frame1 
      Caption         =   "客戶編號(以,分隔 以-為區間)"
      ForeColor       =   &H00C00000&
      Height          =   945
      Left            =   30
      TabIndex        =   24
      Top             =   2430
      Width           =   5115
      Begin VB.TextBox txtCustId 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4875
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "*每個區間以不超過200為限"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   660
         Width           =   2310
      End
   End
   Begin VB.Frame fraType 
      BackColor       =   &H00E0E0E0&
      Caption         =   "報表格式"
      ForeColor       =   &H00004080&
      Height          =   945
      Left            =   6510
      TabIndex        =   23
      Top             =   2430
      Width           =   1305
      Begin VB.OptionButton optDail 
         BackColor       =   &H00E0E0E0&
         Caption         =   "明細表"
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   270
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton optSummaries 
         BackColor       =   &H00E0E0E0&
         Caption         =   "彙總表"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   630
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      Height          =   525
      Left            =   240
      TabIndex        =   15
      Top             =   4650
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   525
      Left            =   6300
      TabIndex        =   17
      Top             =   4650
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   525
      Left            =   1890
      TabIndex        =   16
      Top             =   4650
      Width           =   1395
   End
   Begin VB.Frame fraPage 
      BackColor       =   &H00E0E0E0&
      Caption         =   "分頁方式"
      ForeColor       =   &H00FF0000&
      Height          =   945
      Left            =   5190
      TabIndex        =   20
      Top             =   2430
      Width           =   1245
      Begin VB.OptionButton optNothing 
         BackColor       =   &H00E0E0E0&
         Caption         =   "無"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   630
         Width           =   495
      End
      Begin VB.OptionButton optBankCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "銀行別"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin Gi_Date.GiDate gdaCreateTime2 
      Height          =   375
      Left            =   2430
      TabIndex        =   1
      Top             =   60
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
   Begin Gi_Date.GiDate gdaCreateTime1 
      Height          =   375
      Left            =   1020
      TabIndex        =   0
      Top             =   60
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
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   4830
      TabIndex        =   4
      Top             =   60
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
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   4830
      TabIndex        =   5
      Top             =   495
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
      F5Corresponding =   -1  'True
   End
   Begin Gi_Multi.GiMulti gimCMCode 
      Height          =   345
      Left            =   30
      TabIndex        =   7
      Top             =   1260
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   609
      ButtonCaption   =   "收  費  方  式"
   End
   Begin Gi_Multi.GiMulti gimBankCode 
      Height          =   345
      Left            =   30
      TabIndex        =   9
      Top             =   1980
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   609
      ButtonCaption   =   "銀  行  類  別"
   End
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   315
      Left            =   30
      TabIndex        =   6
      Top             =   930
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   556
      ButtonCaption   =   "收  費  項  目"
   End
   Begin Gi_Multi.GiMulti gimPTCode 
      Height          =   345
      Left            =   30
      TabIndex        =   8
      Top             =   1620
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   609
      ButtonCaption   =   "付  款  種  類"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "[註]: 出單日期(等於工單的受理日期)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   60
      TabIndex        =   28
      Top             =   3570
      Width           =   3360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1560
      TabIndex        =   27
      Top             =   570
      Width           =   105
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "分期期數"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   180
      TabIndex        =   26
      Top             =   570
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3990
      TabIndex        =   22
      Top             =   150
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "服務類別"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3990
      TabIndex        =   21
      Top             =   585
      Width           =   810
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2220
      TabIndex        =   19
      Top             =   150
      Width           =   105
   End
   Begin VB.Label lblShouldDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "出單日期"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   150
      TabIndex        =   18
      Top             =   150
      Width           =   780
   End
End
Attribute VB_Name = "frmSO52H0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table:SO033 or SO034 A(,SO001 B,SO004 C,SO105B D)
'條件：Budget=1
Option Explicit
Dim blnSO001 As Boolean
Dim blnSO105B As Boolean

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      If optDail Then
          Call PreviousRpt(GetPrinterName(5), RptName("SO52H0", 1), "客戶分期收費繳款明細表 [SO52H0A1]")
      Else
          Call PreviousRpt(GetPrinterName(5), RptName("SO52H0", 2), "客戶分期收費繳款彙總表 [SO52H0A2]")
      End If
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo ChkErr
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    Screen.MousePointer = vbHourglass
      cmdExit.SetFocus
      Me.Enabled = False
        ReadyGoPrint
        If Not subChoose Then Exit Sub
        If Not subCreateView Then Exit Sub
        If Not subPrint Then Exit Sub
      Me.Enabled = True
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: MsgMustBe ("公司別"): Exit Function
    If Right(txtCustId, 1) = "," Then txtCustId.SetFocus: GoTo 88
    If Right(txtCustId, 1) = "-" Then txtCustId.SetFocus: GoTo 88
    IsDataOk = True
  Exit Function
88:
    MsgBox "客戶編號," & Chr(10) + Chr(13) & "最後一個字元,需為數字", vbExclamation, "訊息.."
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Function subChoose() As Boolean
  On Error GoTo ChkErr
    Dim strpagetype As String
    strChoose = "A.Budget=1"
    blnSO001 = False
    blnSO105B = False
    

    '日期
      If gdaCreateTime1.GetValue <> "" Then Call subAnd("A.CreateTime >= To_Date('" & gdaCreateTime1.GetValue & "','YYYYMMDD')")
      If gdaCreateTime2.GetValue <> "" Then Call subAnd("A.CreateTime < To_Date('" & gdaCreateTime2.GetValue & "','YYYYMMDD')+1")

    '客戶編號
      If txtCustId.Text <> "" Then
        If TimetxtCustId(txtCustId, "B") = False Then Exit Function
        blnSO001 = True
      End If
      
    '分期期數(SO105B)
      If Trim(txtBudgetPeriod1) <> "" Then Call subAnd("D.BudgetPeriod >=" & Val(txtBudgetPeriod1)): blnSO105B = True
      If Trim(txtBudgetPeriod2) <> "" Then Call subAnd("D.BudgetPeriod <=" & Val(txtBudgetPeriod2)): blnSO105B = True
      
    'GiMulti
      If gimCitemCode.GetQryStr <> "" Then Call subAnd("A.CitemCode " & gimCitemCode.GetQryStr)
      If gimCMCode.GetQryStr <> "" Then Call subAnd("A.CMCode " & gimCMCode.GetQryStr)
      If gimPTCode.GetQryStr <> "" Then Call subAnd("A.ptCode " & gimPTCode.GetQryStr)
      'If gimBankCode.GetQryStr <> "" Then Call subAnd("A.BankCode " & gimBankCode.GetQryStr)
    '銀行類別
      If gimBankCode.GetQryStr <> "" Then
          Call subAnd("A.BankCode " & gimBankCode.GetQryStr)
      Else
         If gimBankCode.GetDispStr <> "" Then Call subAnd("A.BankCode is not null")
      End If
      
    'GiList
      If gilCompCode.GetCodeNo <> "" Then Call subAnd("A.CompCode=" & gilCompCode.GetCodeNo)
      If gilServiceType.GetCodeNo <> "" Then Call subAnd("A.ServiceType='" & gilServiceType.GetCodeNo & "'")

    '分頁方式
      Select Case True
             Case optBankCode.Value
                  strGroupName = "ReportType=True;GroupCode={V.BankCode};GroupName={V.BankName}"
                  strpagetype = "銀行別"
             Case optNothing.Value
                  strGroupName = "ReportType=False;GroupCode={V.Custid};GroupName={V.CustName}"
                  strpagetype = "無"
      End Select

      strChooseString = "出單日期: " & subSpace(gdaCreateTime1.GetValue(True)) & "~" & subSpace(gdaCreateTime2.GetValue(True)) & ";" & _
                        "收費期數: " & subSpace(txtBudgetPeriod1.Text) & "~" & subSpace(txtBudgetPeriod2.Text) & ";" & _
                        "公司別　: " & subSpace(gilCompCode.GetDescription) & ";" & _
                        "服務類別: " & subSpace(gilServiceType.GetDescription) & ";" & _
                        "收費項目: " & subSpace(gimCitemCode.GetDispStr) & ";" & _
                        "收費方式: " & subSpace(gimCMCode.GetDispStr) & ";" & _
                        "付款種類: " & subSpace(gimPTCode.GetDispStr) & ";" & _
                        "銀行類別: " & subSpace(gimBankCode.GetDispStr) & ";" & _
                        "客戶編號: " & subSpace(Trim(txtCustId.Text)) & ";" & _
                        IIf(optSummaries.Value = True, "", ";" & "分頁方式: " & strpagetype) & ";" & _
                        "報表格式:" & IIf(optDail.Value = True, "明細表", "彙總表")
      subChoose = True
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Function

Private Function subPrint() As Boolean
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    If rsTmp.State = 1 Then rsTmp.Close
'    rsTmp.CursorLocation = adUseClient
'    rsTmp.Open "SELECT Count(*) intCount FROM " & strViewName & " Where RowNum=1", gcnGi, adOpenForwardOnly, adLockReadOnly
    strsql = "SELECT Count(*) intCount FROM " & strViewName & " Where RowNum=1"
    If Not GetRS(rsTmp, strsql, gcnGi) Then Exit Function
    If rsTmp("intCount") = 0 Then
        MsgNoRcd
        SendSQL , , True
    Else
        strsql = "Select * From " & strViewName & " V "
        If optDail Then
            Call PrintRpt(GetPrinterName(5), RptName("SO52H0", 1), , Me.Caption, strsql, strChooseString, , True, , , strGroupName, GiPaperLandscape)
        Else
            Call PrintRpt(GetPrinterName(5), RptName("SO52H0", 2), , Me.Caption, strsql, strChooseString, , True)
        End If
    End If
    CloseRecordset rsTmp
    subPrint = True
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subPrint")
End Function

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGim
    Call subGil
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub cmdExit_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr

    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱")
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱")
    Call SetgiMulti(gimPTCode, "CodeNo", "Description", "CD032", "付款種類代碼", "付款種類名稱")
    Call SetgiMulti(gimBankCode, "CodeNo", "Description", "CD018", "銀行別代碼", "銀行別名稱")
    
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    'Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO52H0A)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  'Call DropTMPVIEW(strViewName, strSubViewName())
  gcnGi.Execute "Drop View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO52H0A
End Sub

Private Sub gdaCreateTime1_GotFocus()
  On Error GoTo ChkErr
    If gdaCreateTime1.GetValue = "" Then gdaCreateTime1.SetValue (RightDate)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaCreateTime1_GotFocus")
End Sub

Private Sub gdaCreateTime2_GotFocus()
  On Error GoTo ChkErr
    If gdaCreateTime1.GetValue = "" Or gdaCreateTime2.GetValue = "" Then gdaCreateTime2.SetValue (gdaCreateTime1.GetValue)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaCreateTime2_GotFocus")
End Sub

Private Sub gdaCreateTime2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaCreateTime1, gdaCreateTime2)
  gdaCreateTime2_GotFocus
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    GiMultiFilter gimBankCode, , gilCompCode.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Change")
End Sub

Private Sub gilCompCode_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then
        MsgMustBe ("公司別")
        Cancel = True
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Sub gilServiceType_Change()
  On Error GoTo ChkErr
    GiMultiFilter gimCitemCode, gilServiceType.GetCodeNo
    GiMultiFilter gimCMCode, gilServiceType.GetCodeNo
    GiMultiFilter gimPTCode, gilServiceType.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilServiceType_Change")
End Sub

Private Sub optDail_Click()
  On Error Resume Next
    fraPage.Enabled = True
End Sub

Private Sub optSummaries_Click()
  On Error Resume Next
    fraPage.Enabled = False
End Sub

Private Sub txtCustId_KeyPress(KeyAscii As Integer)
  On Error GoTo ChkErr
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 44 Or KeyAscii = 8 Or KeyAscii = 45 Then
        If KeyAscii = 44 Or KeyAscii = 45 Then
            If Asc(Right(" " & txtCustId.Text, 1)) = 44 Or Asc(Right(" " & txtCustId.Text, 1)) = 45 Or txtCustId = "" Then KeyAscii = 9
        End If
        If Not ChkMaxLengthOK(txtCustId) Then
           If Not (KeyAscii = 44 Or KeyAscii = 8 Or KeyAscii = 45) Then KeyAscii = 9
        End If
    Else
        KeyAscii = 9
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtCustId_KeyPress")
End Sub

Private Function subCreateView() As Boolean
  Dim strView As String
  Dim strTable As String
  Dim strField As String
  Dim strWhere As String
  On Error Resume Next
    gcnGi.Execute "Drop View " & strViewName
  On Error GoTo ChkErr
    strViewName = GetTmpViewName
    strField = ""
    strTable = ""
    strWhere = ""
    
    '客編，客戶姓名，收費方式，信用卡卡號，STB序號，ICC卡號，應收金額，實收金額)，報表內容依客編排序
    '串SQL語法
      If blnSO105B Then
          strTable = ",SO105B D "
          strWhere = strWhere & " A.OrderNo=D.OrderNo And A.CitemCode=D.CitemCode And "
      End If
      If optDail Then
          strField = "Select A.Custid,B.CustName,A.OrderNo,A.CitemCode,A.CitemName,A.BillNo,A.CMCode,A.CMName,A.ShouldAmt,A.RealAmt,A.BankCode,A.BankName,A.AccountNo,C.FaciSNo,C.SmartCardNo"
          strWhere = " Where A.Custid=B.Custid And A.FaciSeqNo=C.SEQNo And " & strWhere & strChoose
          strView = strField & " From SO033 A,SO001 B,SO004 C" & strTable & strWhere & " Union All " & _
                    strField & " From SO034 A,SO001 B,SO004 C" & strTable & strWhere
      Else
          'Decode(Decode(Nvl(Length(SubStr(UCCode,1,1)),0),0,0,1),0,Decode(Nvl(Length(SubStr(RealDate,1,1)),0),0,0,Decode(Nvl(CancelFlag,0),1,0,Sum(RealAmt))),0) RealAmt
          strField = "Select A.Custid,A.OrderNo,A.ShouldAmt," & _
                     "Decode(Decode(Nvl(Length(SubStr(A.UCCode,1,1)),0),0,0,1),0,Decode(Nvl(Length(SubStr(A.RealDate,1,1)),0),0,0,Decode(Nvl(A.CancelFlag,0),1,0,1)),0) RealCounts," & _
                     "Decode(Decode(Nvl(Length(SubStr(A.UCCode,1,1)),0),0,0,1),0,Decode(Nvl(Length(SubStr(A.RealDate,1,1)),0),0,0,Decode(Nvl(A.CancelFlag,0),1,0,A.RealAmt)),0) RealAmt"
          If blnSO001 Then
              strTable = strTable & ",SO001 B"
              strWhere = strWhere & " A.Custid=B.Custid And "
          End If
          strView = "Select Custid,OrderNo,Count(OrderNo) Counts,Sum(ShouldAmt) SumAmt,Sum(RealCounts) RealCounts,Sum(RealAmt) RealAmt,(Count(OrderNo)-Sum(RealCounts)) RemainCount,(Sum(ShouldAmt)-Sum(RealAmt)) RemainAmt From(" & _
                    strField & " From SO033 A" & strTable & " Where " & strWhere & strChoose & " Union All " & _
                    strField & " From SO034 A" & strTable & " Where " & strWhere & strChoose & ") " & _
                    "Group By Custid,OrderNo"
      End If
      
    strView = "Create View " & strViewName & " as (" & strView & ")"
    gcnGi.Execute strView
    SendSQL strView, True
    subCreateView = True
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function
