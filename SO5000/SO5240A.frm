VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.16#0"; "csMulti.ocx"
Begin VB.Form frmSO5240A 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '單線固定
   Caption         =   "收費金額及期數異常表 [SO5240A]"
   ClientHeight    =   4830
   ClientLeft      =   1950
   ClientTop       =   1530
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO5240A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form18"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin prjNumber.GiNumber ginAmount 
      Height          =   345
      Left            =   5580
      TabIndex        =   1
      Top             =   150
      Width           =   1065
      _ExtentX        =   1879
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
      ForeColor       =   0
      MaxLength       =   10
   End
   Begin MSMask.MaskEdBox mskPeriod 
      Height          =   315
      Left            =   5580
      TabIndex        =   2
      Top             =   630
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   2
      Mask            =   "99"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      Height          =   525
      Left            =   90
      TabIndex        =   7
      Top             =   4200
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   525
      Left            =   5670
      TabIndex        =   9
      Top             =   4200
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   525
      Left            =   1770
      TabIndex        =   8
      Top             =   4200
      Width           =   1395
   End
   Begin Gi_Multi.GiMulti gimCustStatus 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   661
      ButtonCaption   =   "客  戶  狀  態"
   End
   Begin prjGiList.GiList gilCitemCode 
      Height          =   315
      Left            =   1050
      TabIndex        =   0
      Top             =   120
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
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1050
      TabIndex        =   3
      Top             =   570
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
      Left            =   1050
      TabIndex        =   4
      Top             =   1020
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
   Begin CS_Multi.CSmulti gimClassCode 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   661
      ButtonCaption   =   "客  戶  類  別"
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   630
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "服務類別"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "2.比較週期性收費之金額是否異常"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   750
      TabIndex        =   15
      Top             =   3150
      Width           =   2865
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "1.收費項目,正常費用,正常每期月數為必要欄位"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   750
      TabIndex        =   14
      Top             =   2760
      Width           =   3930
   End
   Begin VB.Label lblHelp 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "說明:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   90
      TabIndex        =   13
      Top             =   2490
      Width           =   510
   End
   Begin VB.Label lblPeriod 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "正常每期月數"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4320
      TabIndex        =   12
      Top             =   690
      Width           =   1170
   End
   Begin VB.Label lblAmt 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  '透明
      Caption         =   "正常費用"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4320
      TabIndex        =   11
      Top             =   210
      Width           =   780
   End
   Begin VB.Label lblCitemCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收費項目"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   180
      Width           =   780
   End
End
Attribute VB_Name = "frmSO5240A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'使用Table: SO001,SO002,SO003
Option Explicit
Dim StrTableName As String

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      Call PreviousRpt(GetPrinterName(5), RptName("SO5240"), "收費金額及期數異常表 [SO5240A]")
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
        Call subChoose
        Call subPrint
      Me.Enabled = True
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 66
    If gilCitemCode.GetCodeNo = "" Then gilCitemCode.SetFocus: strErrFile = "收費項目": GoTo 66
    If ginAmount.Text = "" Then ginAmount.SetFocus: strErrFile = "正常費用": GoTo 66
    If mskPeriod.Text = "" Then mskPeriod.SetFocus: strErrFile = "正常每期月數": GoTo 66
    IsDataOk = True
  Exit Function
66:
  MsgMustBe (strErrFile)
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub subChoose()
  On Error GoTo ChkErr
    strChoose = " WHERE SO003.CUSTID=SO001.CUSTID And SO003.CUSTID=SO002.CUSTID And SO003.SERVICETYPE=SO002.SERVICETYPE"
    strChooseString = ""
    
  '正常費用及正常每期月數
     Call subAnd("(SO003.Amount<>" & Val(ginAmount.Value) & " or SO003.Period<>" & Val(mskPeriod.Text) & ")")
  
  'GiMulti
    If gimCustStatus.GetQryStr <> "" Then Call subAnd("SO002.CustStatusCode " & gimCustStatus.GetQryStr)
    If gimClassCode.GetQryStr <> "" Then Call subAnd("SO001.ClassCode1 " & gimClassCode.GetQryStr)
  'GiList
    If gilCitemCode.GetCodeNo <> "" Then Call subAnd("SO003.CitemCode=" & gilCitemCode.GetCodeNo)
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO003.CompCode=" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO003.ServiceType='" & gilServiceType.GetCodeNo & "'")
    
    strChooseString = "收費項目: " & subSpace(gilCitemCode.GetDescription) & ";" & _
                      "正常金額: " & subSpace(ginAmount.Value) & ";" & _
                      "正常期數: " & subSpace(mskPeriod.Text) & ";" & _
                      "公司別　: " & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別: " & subSpace(gilServiceType.GetDescription) & ";" & _
                      "客戶類別: " & subSpace(gimClassCode.GetDispStr) & ";" & _
                      "客戶狀態: " & subSpace(gimCustStatus.GetDispStr)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub subPrint()
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    If rsTmp.State = 1 Then rsTmp.Close
'    rsTmp.CursorLocation = adUseClient
'      rsTmp.Open "SELECT Count(*) as intCount FROM SO001,SO002,SO003 " & IIf(strChoose = "", " Where RowNum=1", strChoose & " And RowNum=1"), gcnGi, adOpenForwardOnly, adLockReadOnly
    strsql = "SELECT Count(*) as intCount FROM SO001,SO002,SO003 " & IIf(strChoose = "", " Where RowNum=1", strChoose & " And RowNum=1")
    If Not GetRS(rsTmp, strsql, gcnGi) Then Exit Sub
      If rsTmp("intCount") = 0 Then
         MsgNoRcd
         SendSQL , , True
      Else
         strsql = "SELECT SO003.CUSTID,SO003.PERIOD,SO003.AMOUNT,SO003.STARTDATE,SO003.STOPDATE," & _
               "SO003.CLCTDATE,SO001.CUSTNAME,SO001.TEL1,SO001.INSTADDRESS,SO002.CUSTSTATUSNAME," & _
               "SO001.CLASSNAME1,SO001.VIEWLEVEL,SO002.INSTCOUNT,SO002.INSTTIME " & _
               "From SO001,SO002,SO003 " & strChoose & " Order By CustId"
         Call PrintRpt(GetPrinterName(5), RptName("SO5240"), , "收費金額及期數異常表 [SO5240A]", strsql, strChooseString, , True, , , , GiPaperLandscape)
      End If
      CloseRecordset rsTmp
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subInsertMDB")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGim
    Call subGil
'    gilCompCode.ListIndex = 1
'    gilCompCode.Query_Description
'    gilCompCode_Change
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
    Call SetgiMulti(gimCustStatus, "CodeNo", "Description", "CD035", "客戶狀態代碼", "客戶狀態名稱")
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    gimCustStatus.SetQueryCode "1,2"
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    Call SetgiList(gilCitemCode, "CodeNo", "Description", "CD019")
    'Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub
  
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5240A)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5240A
End Sub

Private Sub gilCitemCode_LostFocus()
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    If gilCitemCode.GetCodeNo = "" Then Exit Sub
    Set rsTmp = gcnGi.Execute("SELECT Amount,Period From CD019 WHERE CodeNo=" & gilCitemCode.GetCodeNo)
    ginAmount.Text = IIf(varType(rsTmp("Amount")) = vbNull, "", rsTmp("Amount"))
    mskPeriod.Text = IIf(varType(rsTmp("Period")) = vbNull, "", IIf(rsTmp("Period") = 0, "", rsTmp("Period")))
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCitemCode_LostFocus")
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr

    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
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
    GiMultiFilter gimClassCode, gilServiceType.GetCodeNo
    GiListFilter gilCitemCode, gilServiceType.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilServiceType_Change")
End Sub

Private Sub mskPeriod_GotFocus()
  On Error Resume Next
    mskPeriod.SelStart = 0
    mskPeriod.SelLength = mskPeriod.SelLength
End Sub
