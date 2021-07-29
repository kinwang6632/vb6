VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO5590A 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶信用卡資料統計表 [SO5590A]"
   ClientHeight    =   7260
   ClientLeft      =   1890
   ClientTop       =   435
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO5590A.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin MSMask.MaskEdBox gymCardExpDate1 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   90
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin VB.CheckBox chkAllStop 
      Caption         =   "是否包含統計停用"
      Height          =   345
      Left            =   4830
      TabIndex        =   22
      Top             =   5460
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "地址依據"
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   3900
      TabIndex        =   38
      Top             =   840
      Width           =   3825
      Begin VB.OptionButton optInstAddress 
         Caption         =   "裝機地址"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optChargeAddress 
         Caption         =   "收費地址"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1290
         TabIndex        =   9
         Top             =   270
         Width           =   1125
      End
      Begin VB.OptionButton optMailAddress 
         Caption         =   "郵寄地址"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2460
         TabIndex        =   10
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "網路編號"
      Height          =   735
      Left            =   15
      TabIndex        =   32
      Top             =   5745
      Width           =   7935
      Begin VB.OptionButton optCircuitNo 
         Caption         =   "只印空白網路編號"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   4830
         TabIndex        =   24
         Top             =   240
         Width           =   1995
      End
      Begin VB.OptionButton optAll 
         Caption         =   "全部"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7050
         TabIndex        =   25
         Top             =   270
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.TextBox mskCircuitNo 
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   90
         TabIndex        =   23
         Top             =   240
         Width           =   4605
      End
   End
   Begin CS_Multi.CSmulti gimStrtCode 
      Height          =   405
      Left            =   0
      TabIndex        =   18
      Top             =   3900
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   714
      ButtonCaption   =   "街  道  範  圍"
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   525
      Left            =   1770
      TabIndex        =   27
      Top             =   6600
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   525
      Left            =   6600
      TabIndex        =   28
      Top             =   6600
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      Height          =   525
      Left            =   90
      TabIndex        =   26
      Top             =   6600
      Width           =   1245
   End
   Begin Gi_Multi.GiMulti gimBankCode 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   1455
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      ButtonCaption   =   "發  卡  銀  行"
   End
   Begin Gi_Multi.GiMulti gimStatusCode 
      Height          =   345
      Left            =   0
      TabIndex        =   13
      Top             =   2175
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   609
      ButtonCaption   =   "客  戶  狀  態"
   End
   Begin Gi_Multi.GiMulti gimCMCode 
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   1815
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      ButtonCaption   =   "收  費  方  式"
   End
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   3555
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      ButtonCaption   =   "行     政     區"
   End
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   3225
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      ButtonCaption   =   "服     務     區"
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   4770
      TabIndex        =   6
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
      Left            =   4770
      TabIndex        =   7
      Top             =   465
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
   Begin Gi_Multi.GiMulti gimOrder 
      Height          =   405
      Left            =   0
      TabIndex        =   21
      Top             =   5040
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   714
      ButtonCaption   =   "排  序  方  式"
      DataType        =   2
      ColumnOrder     =   -1  'True
   End
   Begin CS_Multi.CSmulti gimMduId 
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   4290
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      ButtonCaption   =   "大  樓  編  號"
   End
   Begin CS_Multi.CSmulti gimClassCode 
      Height          =   345
      Left            =   0
      TabIndex        =   14
      Top             =   2520
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   609
      ButtonCaption   =   "客  戶  類  別"
   End
   Begin CS_Multi.CSmulti gimUpdEn 
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   2880
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      ButtonCaption   =   "異  動  人  員"
   End
   Begin Gi_Date.GiDate gdaClctDate2 
      Height          =   375
      Left            =   2610
      TabIndex        =   5
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ForeColor       =   12582912
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
   Begin Gi_Date.GiDate gdaClctDate1 
      Height          =   375
      Left            =   1260
      TabIndex        =   4
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ForeColor       =   12582912
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
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   4650
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      ButtonCaption   =   "收  費  項  目"
   End
   Begin Gi_Date.GiDate gdaUpdTime2 
      Height          =   345
      Left            =   2610
      TabIndex        =   3
      Top             =   510
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   8388608
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
   Begin Gi_Date.GiDate gdaUpdTime1 
      Height          =   345
      Left            =   1260
      TabIndex        =   2
      Top             =   510
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   8388608
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
   Begin MSMask.MaskEdBox gymCardExpDate2 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   90
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "---"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2760
      TabIndex        =   31
      Top             =   165
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "信用卡有效期限(西元)"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   90
      TabIndex        =   37
      Top             =   180
      Width           =   1875
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "---"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2400
      TabIndex        =   36
      Top             =   600
      Width           =   180
   End
   Begin VB.Label lblUpdTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "異動日期"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   90
      TabIndex        =   35
      Top             =   630
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "---"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2400
      TabIndex        =   34
      Top             =   1050
      Width           =   180
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "次收費日"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   90
      TabIndex        =   33
      Top             =   1050
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "服務類別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3900
      TabIndex        =   30
      Top             =   540
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公 司 別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3900
      TabIndex        =   29
      Top             =   150
      Width           =   675
   End
End
Attribute VB_Name = "frmSO5590A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table:SO001,SO002,SO003(,SO014)
Option Explicit
Dim blnSO014 As Boolean

Private Sub cmdExit_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      Call PreviousRpt(GetPrinterName(5), RptName("SO5590"), Me.Caption)
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
        Call subPrint
      Me.Enabled = True
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Sub subPrint()
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
        If Not subCreateView(strChoose) Then Exit Sub
        rsTmp.CursorLocation = adUseClient
        Set rsTmp = gcnGi.Execute("SELECT  Count(*) as intCount FROM " & strViewName & " Where RowNum=1")
        If rsTmp("intCount") = 0 Then
            MsgNoRcd
        Else
            strSQL = "Select * From " & strViewName & " V"
            Call PrintRpt2(GetPrinterName(5), RptName("SO5590"), , Me.Caption, strSQL, strChooseString, , True, , , strGroupName, GiPaperLandscape)
        End If
    CloseRecordset rsTmp
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subPrint")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 66
    If gilServiceType.GetCodeNo = "" Then gilServiceType.SetFocus: strErrFile = "服務類別": GoTo 66
'    If gymCardExpDate1.GetValue = "" Then gymCardExpDate1.SetFocus: strErrFile = "信用卡有效期限": GoTo 66
'    If gymCardExpDate2.GetValue = "" Then gymCardExpDate2.SetFocus: strErrFile = "信用卡有效期限": GoTo 66
    IsDataOk = True
  Exit Function
66:
  MsgMustBe (strErrFile)
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Function subChoose() As Boolean
  On Error GoTo ChkErr
    Dim strFrom As String
    Dim strField As String
    Dim strCardExpDate As String
    Dim strAddress As String
    Dim strQry As String
    strCardExpDate = "DeCode(length(SO002A.CardExpDate),5,substr(SO002A.CardExpDate,2,4) || '0' || substr(SO002A.CardExpDate,1,1),substr(SO002A.CardExpDate,3,4) || substr(SO002A.CardExpDate,1,2))"
    strChoose = "SO001.Custid=SO002.Custid And SO001.Custid=SO002A.Custid AND SO002A.Custid=SO003.Custid(+) AND SO002A.AccountNo=SO003.AccountNo(+) AND SO002A.AccountNo=SO106.AccountID AND SO002A.custid=SO106.custid(+)"
    strChooseString = ""
    strField = "Select SO001.Custid,SO001.CustName,SO001.Tel1,SO002.CustStatusName,SO001.ClassName1,SO002A.BankCode,SO002A.BankName,SO002A.AccountNo CreditCardNo,SO002A.CardExpDate,SO001.ServCode,SO001.ServArea,SO003.ClctDate," & strCardExpDate & " CardExpDateSort,SO002A.Note,SO106.AccountName"
    strFrom = "SO001,SO002,SO002A,SO003,SO106"
    strGroupName = ""
    blnSO014 = False

  '日期
    If gymCardExpDate1.ClipText <> "" Then Call subAnd(strCardExpDate & " >= " & gymCardExpDate1.Text)
    If gymCardExpDate2.ClipText <> "" Then Call subAnd(strCardExpDate & " <= " & gymCardExpDate2.Text)
    '改用GiPackage Function By Kin 2011/10/18
    If gdaUpdTime1.GetValue <> "" Then Call subAnd("GiPackage.GetDate1(so106.UpdTime) >= " & GetNullString(gdaUpdTime1.GetValue(True), giDateV, giOracle, True) & "")
    If gdaUpdTime2.GetValue <> "" Then Call subAnd("GiPackage.GetDate1(so106.UpdTime) <= " & GetNullString(gdaUpdTime2.GetValue(True), giDateV, giOracle, True) & "")
'    If gdaUpdTime1.GetValue <> "" Then Call subAnd("substr(so106.UpdTime,1,8) >='" & gdaUpdTime1.GetOriginalValue & "'")
'    If gdaUpdTime2.GetValue <> "" Then Call subAnd("substr(so106.UpdTime,1,8) <='" & gdaUpdTime2.GetOriginalValue & "'")
    If gdaClctDate1.GetValue <> "" Then Call subAnd("SO003.ClctDate >= To_Date('" & gdaClctDate1.GetValue & "','YYYYMMDD')")
    If gdaClctDate2.GetValue <> "" Then Call subAnd("SO003.ClctDate < To_Date('" & gdaClctDate2.GetValue & "','YYYYMMDD')+1")

  'GiList
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO002.ServiceType='" & gilServiceType.GetCodeNo & "'")
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("instr(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "')>0")
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO003.ServiceType='" & gilServiceType.GetCodeNo & "'")
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO001.CompCode=" & gilCompCode.GetCodeNo)
  
  'GiMulti
    If gimBankCode.GetQryStr <> "" Then Call subAnd("SO002A.BankCode " & gimBankCode.GetQryStr)
    If gimCMCode.GetQryStr <> "" Then Call subAnd("SO003.CMCode " & gimCMCode.GetQryStr)
    If gimStatusCode.GetQryStr <> "" Then Call subAnd("SO002.CustStatusCode " & gimStatusCode.GetQryStr)
    If gimClassCode.GetQryStr <> "" Then Call subAnd("SO001.ClassCode1 " & gimClassCode.GetQryStr)
    If gimUpdEn.GetQryStr <> "" Then Call subAnd("SO001.UpdEn " & gimUpdEn.GetDescStr)
    If gimServCode.GetQryStr <> "" Then Call subAnd("SO001.ServCode " & gimServCode.GetQryStr)
    If gimAreaCode.GetQryStr <> "" Then Call subAnd("SO014.AreaCode " & gimAreaCode.GetQryStr): blnSO014 = True
    If gimStrtCode.GetQryStr <> "" Then Call subAnd("SO014.StrtCode " & gimStrtCode.GetQryStr): blnSO014 = True
    If gimCitemCode.GetQryStr <> "" Then Call subAnd("SO003.CitemCode " & gimCitemCode.GetQryStr)
      
  '大樓編號
    If gimMduId.GetQryStr <> "" Then
        Call subAnd("SO001.MduId " & gimMduId.GetQryStr)
    Else
        If gimMduId.GetDispStr <> "" Then Call subAnd("SO001.MduId is Not Null")
    End If
    
  '網路編號
    If mskCircuitNo.Text <> "" Then
       Call subAnd("SO014.CircuitNo='" & mskCircuitNo.Text & "'")
       blnSO014 = True
    Else
      If optCircuitNo.Value Then Call subAnd("SO014.CircuitNo is null "): blnSO014 = True
    End If
    
  '是否包含統計停用-沒有勾選是過濾停用,勾選時不過濾
    If chkAllStop.Value = 0 Then Call subAnd("(SO002A.Stopflag=0 AND SO106.Stopflag=0)")
    
    
  '排序方式
    SetGiMultiOrder
    
  '串SO014
    If blnSO014 Then
        strFrom = strFrom & ",SO014"
        strField = strField & ",SO014.AreaCode,SO014.AddrSort"
    Else
        strField = strField & ",'' AreaCode,'' AddrSort"
    End If
    
  '裝機依據
    Select Case True
           Case optInstAddress.Value
                strChoose = strField & ",SO001.InstAddress Address From " & strFrom & " Where " & IIf(blnSO014, "SO001.InstAddrNo=SO014.AddrNo And ", "") & strChoose
                strAddress = "裝機地址"
           Case optChargeAddress.Value
                '#5892 測試不OK,收費地址也要依據SO138 By Kin 2011/06/15
                strQry = strField & ",SO001.ChargeAddress Address From " & strFrom & " Where " & IIf(blnSO014, "SO001.ChargeAddrNo=SO014.AddrNo And ", "") & strChoose & " AND SO106.INVSEQNO IS NULL "
                strQry = strQry & " UNION ALL " & strField & ",SO138.ChargeAddress FROM " & strFrom & ",SO138 " & _
                        " Where " & IIf(blnSO014, "SO138.ChargeAddrNo=SO014.AddrNo And ", "") & strChoose & " AND SO106.INVSEQNO IS NOT NULL AND SO106.INVSEQNO = SO138.InvSeqNo "
                strChoose = strQry
                'strChoose = strField & ",SO001.ChargeAddress Address From " & strFrom & " Where " & IIf(blnSO014, "SO001.ChargeAddrNo=SO014.AddrNo And ", "") & strChoose
                strAddress = "收費地址"
           Case optMailAddress.Value
                '#5892 郵寄地址如果SO138.INVSEQNO有值，要以SO138的為主 By Kin 2011/05/23
                strQry = strField & ",SO001.MailAddress Address From " & strFrom & " Where " & IIf(blnSO014, "SO001.ChargeAddrNo=SO014.AddrNo And ", "") & strChoose & " AND SO106.INVSEQNO IS  NULL "
                strQry = strQry & " UNION ALL " & strField & ",SO138.MailAddress FROM " & strFrom & ",SO138 " & _
                        " Where " & IIf(blnSO014, "SO138.MailAddrNo=SO014.AddrNo And ", "") & strChoose & " AND SO106.INVSEQNO IS NOT NULL AND SO106.INVSEQNO = SO138.InvSeqNo "
                strChoose = strQry
                'strChoose = strField & ",SO001.MailAddress Address From " & strFrom & " Where " & IIf(blnSO014, "SO001.ChargeAddrNo=SO014.AddrNo And ", "") & strChoose
                strAddress = "郵寄地址"
    End Select
    
    
    strChooseString = "信用卡期限:" & subSpace(gymCardExpDate1.Text) & "~" & subSpace(gymCardExpDate2.Text) & ";" & _
                      "異動日期:" & subSpace(gdaUpdTime1.GetValue(True)) & "~" & subSpace(gdaUpdTime2.GetValue(True)) & ";" & _
                      "次收費日:" & subSpace(gdaClctDate1.GetValue(True)) & "~" & subSpace(gdaClctDate1.GetValue(True)) & ";" & _
                      "公司別　:" & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別:" & subSpace(gilServiceType.GetDescription) & ";" & _
                      "地址依據:" & strAddress & ";" & _
                      "發卡銀行:" & subSpace(gimBankCode.GetDispStr) & ";" & _
                      "收費方式:" & subSpace(gimCMCode.GetDispStr) & ";" & _
                      "客戶狀態:" & subSpace(gimStatusCode.GetDispStr) & ";" & _
                      "客戶類別:" & subSpace(gimClassCode.GetDispStr) & ";" & _
                      "異動人員:" & subSpace(gimUpdEn.GetDispStr) & ";" & _
                      "服務區　:" & subSpace(gimServCode.GetDispStr) & ";" & _
                      "行政區　:" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                      "街道範圍:" & subSpace(gimStrtCode.GetDispStr) & ";" & _
                      "大樓名稱:" & subSpace(gimMduId.GetDispStr) & ";" & _
                      "收費項目:" & subSpace(gimCitemCode.GetDispStr) & ";" & _
                      "排序方式:" & subSpace(gimOrder.GetColumnOrderDspStr) & ";" & _
                      "網路編號:" & subSpace(mskCircuitNo.Text)
    subChoose = True
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5590A)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

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

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimBankCode, "CodeNo", "Description", "CD018", "銀行代碼", "銀行名稱")
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱")
    Call SetgiMulti(gimStatusCode, "CodeNo", "Description", "CD035", "客戶狀態代碼", "客戶狀態名稱")
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    Call SetgiMulti(gimUpdEn, "EmpNo", "EmpName", "CM003", "異動人員代碼", "異動人員名稱")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "街道代碼", "街道名稱")
    Call SetgiMulti(gimMduId, "MduId", "Name", "SO017", "大樓代號", "大樓名稱")
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱")
    Call SetgiMultiAddItem(gimOrder, "A,B,C,D,E,F,G", "客戶編號,信用卡號,銀行代碼,到期年月,服務區,行政區,住址")

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

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    ReleaseCOM frmSO5590A
End Sub

Private Sub gymCardExpDate1_GotFocus()
  On Error Resume Next
    If gymCardExpDate1.ClipText = "" Then gymCardExpDate1.Text = Left(RightDate, 7)
End Sub

Private Sub gymCardExpDate2_GotFocus()
  On Error Resume Next
    If gymCardExpDate1.ClipText = "" Or gymCardExpDate2.ClipText = "" Then gymCardExpDate2.Text = gymCardExpDate1.Text
End Sub

'Private Sub gymCardExpDate2_Validate(Cancel As Boolean)
'  On Error Resume Next
'    Cancel = ChkDate2(gymCardExpDate1, gymCardExpDate2)
'End Sub

Private Sub gdaUpdTime1_GotFocus()
  On Error Resume Next
    If gdaUpdTime1.GetValue = "" Then gdaUpdTime1.SetValue (RightDate)
End Sub

Private Sub gdaUpdTime2_GotFocus()
  On Error Resume Next
    If gdaUpdTime1.GetValue = "" Or gdaUpdTime2.GetValue = "" Then gdaUpdTime2.SetValue (gdaUpdTime1.GetValue)
End Sub

Private Sub gdaUpdTime2_Validate(Cancel As Boolean)
  On Error Resume Next
    Cancel = ChkDate2(gdaUpdTime1, gdaUpdTime2)
End Sub

Private Sub gdaClctDate1_GotFocus()
  On Error Resume Next
    If gdaClctDate1.GetValue = "" Then gdaClctDate1.SetValue (RightDate)
End Sub

Private Sub gdaClctDate2_GotFocus()
  On Error Resume Next
    If gdaClctDate1.GetValue = "" Or gdaClctDate2.GetValue = "" Then gdaClctDate2.SetValue (gdaClctDate1.GetValue)
End Sub

Private Sub gdaClctDate2_Validate(Cancel As Boolean)
  On Error Resume Next
    Cancel = ChkDate2(gdaClctDate1, gdaClctDate2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    GiMultiFilter gimBankCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimUpdEn, , gilCompCode.GetCodeNo
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimMduId, , gilCompCode.GetCodeNo
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
    GiMultiFilter gimCMCode, gilServiceType.GetCodeNo
    GiMultiFilter gimClassCode, gilServiceType.GetCodeNo
    GiMultiFilter gimCitemCode, gilServiceType.GetCodeNo
    'gimCitemCode.Filter = gimCitemCode.Filter & IIf(gimCitemCode.Filter = "", " WHERE ", " AND ") & " PeriodFlag = 1"
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilServiceType_Change")
End Sub

Public Sub SetGiMultiOrder()
  On Error GoTo ChkErr
    Dim intSort As Integer
    Dim arrOrder() As String
    Dim varOrder As Variant
      intSort = 0
      arrOrder = Split(gimOrder.GetColumnOrderCode, ",")
      For Each varOrder In arrOrder
        Select Case arrOrder(intSort)
               Case "A"
                    strGroupName = strGroupName & ";Sort" & intSort & "={V.Custid}"
               Case "B"
                    strGroupName = strGroupName & ";Sort" & intSort & "={V.CreditCardNo}"
               Case "C"
                    strGroupName = strGroupName & ";Sort" & intSort & "={V.BankCode}"
               Case "D"
                    strGroupName = strGroupName & ";Sort" & intSort & "={V.CardExpDateSort}"
               Case "E"
                    strGroupName = strGroupName & ";Sort" & intSort & "={V.ServCode}"
               Case "F"
                    strGroupName = strGroupName & ";Sort" & intSort & "={V.AreaCode}"
                    blnSO014 = True
               Case "G"
                    strGroupName = strGroupName & ";Sort" & intSort & "={V.AddrSort}"
                    blnSO014 = True
        End Select
        intSort = intSort + 1
      Next
  Exit Sub
ChkErr:
    ErrSub Me.Name, "SetGiMultiOrder"
End Sub

Private Function subCreateView(ByVal strViewSql As String) As Boolean
  On Error Resume Next
    gcnGi.Execute "Drop View " & strViewName
  On Error GoTo ChkErr
    strViewName = GetTmpViewName
    strViewSql = "Create View " & strViewName & " as (" & strViewSql & ")"
    gcnGi.Execute strViewSql
    SendSQL strViewSql, True
  subCreateView = True
  Exit Function
ChkErr:
    ErrSub Me.Name, "subCreateView"
End Function

