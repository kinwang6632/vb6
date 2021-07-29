VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO52E0A 
   BorderStyle     =   1  '單線固定
   Caption         =   "應收報表查詢 [SO52E0A]"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "SO52E0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin Gi_Multi.GiMulti gimPTCode 
      Height          =   345
      Left            =   0
      TabIndex        =   19
      Top             =   4800
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   609
      ButtonCaption   =   "付  款  種  類"
   End
   Begin CS_Multi.CSmulti gimStrtCode 
      Height          =   345
      Left            =   0
      TabIndex        =   16
      Top             =   3750
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   609
      ButtonCaption   =   "街  道  編  號"
   End
   Begin VB.CheckBox chkBad 
      Caption         =   "是否包含呆帳資料統計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3870
      TabIndex        =   8
      Top             =   870
      Value           =   1  '核取
      Width           =   3285
   End
   Begin VB.Frame fraReportType 
      Caption         =   "報表格式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   705
      Left            =   5430
      TabIndex        =   36
      Top             =   5580
      Width           =   2355
      Begin VB.OptionButton optDetail 
         Caption         =   "明細表"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1320
         TabIndex        =   27
         Top             =   300
         Width           =   915
      End
      Begin VB.OptionButton optSummaric 
         Caption         =   "彙總表"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   270
         TabIndex        =   26
         Top             =   300
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Frame fraPage 
      Caption         =   "分頁方式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   705
      Left            =   0
      TabIndex        =   31
      Top             =   5580
      Width           =   5325
      Begin VB.OptionButton optClctAreaCode 
         Caption         =   "收費區"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2430
         TabIndex        =   23
         Top             =   330
         Width           =   855
      End
      Begin VB.OptionButton optServCode 
         Caption         =   "服務區"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   270
         TabIndex        =   21
         Top             =   330
         Width           =   975
      End
      Begin VB.OptionButton optAreaCode 
         Caption         =   "行政區"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1350
         TabIndex        =   22
         Top             =   330
         Width           =   975
      End
      Begin VB.OptionButton optClctEn 
         Caption         =   "收費人員"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3390
         TabIndex        =   24
         Top             =   330
         Width           =   1125
      End
      Begin VB.OptionButton optNothing 
         Caption         =   "無"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4620
         TabIndex        =   25
         Top             =   330
         Value           =   -1  'True
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1830
      TabIndex        =   29
      Top             =   6600
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6090
      TabIndex        =   30
      Top             =   6570
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   150
      TabIndex        =   28
      Top             =   6570
      Width           =   1245
   End
   Begin Gi_Multi.GiMulti gimUCCode 
      Height          =   345
      Left            =   0
      TabIndex        =   18
      Top             =   4455
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   609
      ButtonCaption   =   "未  收  原  因"
   End
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   345
      Left            =   0
      TabIndex        =   15
      Top             =   3405
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   609
      ButtonCaption   =   "行     政     區"
   End
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   345
      Left            =   0
      TabIndex        =   14
      Top             =   3075
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   609
      ButtonCaption   =   "服     務     區"
   End
   Begin Gi_Date.GiDate gdaShouldDate2 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   30
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
   Begin Gi_Date.GiDate gdaShouldDate1 
      Height          =   375
      Left            =   990
      TabIndex        =   0
      Top             =   30
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
   Begin Gi_Multi.GiMulti gimBillType 
      Height          =   345
      Left            =   0
      TabIndex        =   10
      Top             =   1725
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   609
      ButtonCaption   =   "單  據  類  別"
      DataType        =   2
      DIY             =   -1  'True
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   4650
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
      Left            =   4650
      TabIndex        =   7
      Top             =   435
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
      Height          =   345
      Left            =   0
      TabIndex        =   20
      Top             =   5145
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   609
      ButtonCaption   =   "排  序  方  式"
      DataType        =   2
      ColumnOrder     =   -1  'True
   End
   Begin Gi_Multi.GiMulti gimCustStatus 
      Height          =   345
      Left            =   0
      TabIndex        =   13
      Top             =   2730
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   609
      ButtonCaption   =   "客  戶  狀  態"
   End
   Begin Gi_Date.GiDate gdaRealDate2 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   480
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
   Begin Gi_Date.GiDate gdaRealDate1 
      Height          =   375
      Left            =   990
      TabIndex        =   2
      Top             =   480
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
   Begin CS_Multi.CSmulti gimMduId 
      Height          =   345
      Left            =   0
      TabIndex        =   17
      Top             =   4110
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   609
      ButtonCaption   =   "大  樓  編  號"
   End
   Begin Gi_Date.GiDate gdaCreateTime2 
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   930
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
      Left            =   990
      TabIndex        =   4
      Top             =   930
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
   Begin CS_Multi.CSmulti gimClctEn 
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   2070
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   609
      ButtonCaption   =   "收  費  人   員"
   End
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   1410
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   556
      ButtonCaption   =   "收  費  項  目"
   End
   Begin CS_Multi.CSmulti gimClassCode 
      Height          =   345
      Left            =   0
      TabIndex        =   12
      Top             =   2400
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   609
      ButtonCaption   =   "客  戶  類  別"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "出單日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   40
      Top             =   1020
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2190
      TabIndex        =   39
      Top             =   1020
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2190
      TabIndex        =   38
      Top             =   570
      Width           =   105
   End
   Begin VB.Label lblRealDay 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "實收日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   37
      Top             =   570
      Width           =   780
   End
   Begin VB.Label lblShouldDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "應收日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   150
      Width           =   780
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2190
      TabIndex        =   34
      Top             =   120
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3810
      TabIndex        =   33
      Top             =   495
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
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
      Left            =   3810
      TabIndex        =   32
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "frmSO52E0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table:(SO033 or SO034 or SO035 or SO036 ) A,SO001,SO002,SO014,CD004
Option Explicit
Dim strST As String 'RefNo<>1
Dim strST2 As String 'RefNo=1 呆帳
Dim GroupCode1 As String
Dim GroupName1 As String
Dim GroupCode2 As String
Dim GroupName2 As String
Dim blnSO002 As Boolean

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      If optDetail Then '明細
          Call PreviousRpt(GetPrinterName(5), RptName("SO52E0", 2), "應收明細表 [SO52E0A]")
      Else
          Call PreviousRpt(GetPrinterName(5), RptName("SO52E0", 1), "應收彙總表 [SO52E0A]")
      End If
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo ChkErr
    If Not ChkDTok Then Exit Sub
      Screen.MousePointer = vbHourglass
        cmdExit.SetFocus
        Me.Enabled = False
          ReadyGoPrint
          Call subChoose
          If Not subCreateView Then Exit Sub
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
    Dim strpagetype As String
    Dim strBadCode As String
    Dim arrOrder() As String
    Dim varOrder As Variant
    Dim intSort As Integer
      'strChoose = " A.CancelFlag = 0"
      strChoose = ""
      strChooseString = ""
      strGroupName = ""
      blnSO002 = False
    
      'strBadCode = GetRsValue("Select CodeNo From CD016 Where RefNo = 1") & ""
      'If strBadCode = "" Then strBadCode = "-1"
    '日期
      If gdaShouldDate1.GetValue <> "" Then Call subAnd("A.ShouldDate >= To_Date('" & gdaShouldDate1.GetValue & "','YYYYMMDD')")
      If gdaShouldDate2.GetValue <> "" Then Call subAnd("A.ShouldDate < To_Date('" & gdaShouldDate2.GetValue & "','YYYYMMDD')+1")
      If gdaRealDate1.GetValue <> "" Then subAnd ("A.RealDate >= To_Date('" & gdaRealDate1.GetValue & "','YYYYMMDD')")
      If gdaRealDate2.GetValue <> "" Then subAnd ("A.RealDate < To_Date('" & gdaRealDate2.GetValue & "','YYYYMMDD')+1")
      If gdaRealDate1.GetValue <> "" Then strGroupName = "RealDate1=#" & gdaRealDate1.GetValue(True) & "#;"
      If gdaRealDate2.GetValue <> "" Then strGroupName = strGroupName & "RealDate2=#" & gdaRealDate2.GetValue(True) & "#;"
      If gdaCreateTime1.GetValue <> "" Then Call subAnd("A.CreateTime >= To_Date('" & gdaCreateTime1.GetValue & "','YYYYMMDD')")
      If gdaCreateTime2.GetValue <> "" Then Call subAnd("A.CreateTime < To_Date('" & gdaCreateTime2.GetValue & "','YYYYMMDD')+1")
      
    '是否包含呆帳資料統計
      If chkBad.Value = 0 Then Call subAnd(IIf(strST = "", "", " nvl(A.STCode,-1) in(" & strST & ",-1)"))
      
    'GiMulti
      If gimAreaCode.GetQryStr <> "" Then Call subAnd("A.AreaCode " & gimAreaCode.GetQryStr)
      If gimCitemCode.GetQryStr <> "" Then Call subAnd("A.CitemCode " & gimCitemCode.GetQryStr)
      If gimClassCode.GetQryStr <> "" Then Call subAnd("A.ClassCode " & gimClassCode.GetQryStr)
      If gimCustStatus.GetQryStr <> "" Then Call subAnd("SO002.CustStatusCode " & gimCustStatus.GetQryStr): blnSO002 = True
      If gimMduId.GetQryStr <> "" Then
         Call subAnd("A.MduId " & gimMduId.GetQryStr)
      Else
         If gimMduId.GetDispStr <> "" Then subAnd ("A.MduId is Not Null")
      End If
      If gimClctEn.GetQryStr <> "" Then Call subAnd("A.ClctEn " & gimClctEn.GetQryStr)
      If gimServCode.GetQryStr <> "" Then Call subAnd("A.ServCode " & gimServCode.GetQryStr)
      If gimStrtCode.GetQryStr <> "" Then Call subAnd("A.StrtCode " & gimStrtCode.GetQryStr)
      If gimUCCode.GetQryStr <> "" Then Call subAnd("A.UCCode " & gimUCCode.GetQryStr)
      If gimBillType.GetQryStr <> "" Then Call subAnd("SubStr(A.BillNo,7,1) " & gimBillType.GetQryStr)
      If gimPTCode.GetQryStr <> "" Then Call subAnd("A.ptCode " & gimPTCode.GetQryStr)
      
    'GiList
      If gilCompCode.GetCodeNo <> "" Then Call subAnd("A.CompCode=" & gilCompCode.GetCodeNo)
      If gilServiceType.GetCodeNo <> "" Then Call subAnd("A.ServiceType='" & gilServiceType.GetCodeNo & "'")
      'If gilServiceType.GetCodeNo <> "" Then Call subAnd("Instr(Nvl(A.ServiceType,'" & gilServiceType.GetCodeNo & "'),'" & gilServiceType.GetCodeNo & "') > 0")
      
    '分頁方式
      If optDetail Then
          Select Case True
                 Case optServCode.Value
                      strGroupName = "ReportType=True;GroupName={V.ServCode};GroupName2={V.ServName}"
                      strpagetype = "服務區"
                 Case optAreaCode.Value
                      strGroupName = "ReportType=True;GroupName={V.AreaCode};GroupName2={V.AreaName}"
                      strpagetype = "行政區"
                 Case optClctAreaCode.Value
                      strGroupName = "ReportType=True;GroupName={V.ClctAreaCode};GroupName2={V.ClctAreaName}"
                      strpagetype = "收費區"
                 Case optClctEn.Value
                      strGroupName = "ReportType=True;GroupName={V.ClctEn};GroupName2={V.ClctName}"
                      strpagetype = "收費人員"
                 Case optNothing.Value
                      Select Case Left(gimOrder.GetColumnOrderCode, 1)
                             Case "A"
                                  strGroupName = "ReportType=false;GroupName={V.NODENO};GroupName2={V.NODENO}"
                             Case "B"
                                  strGroupName = "ReportType=false;GroupName={V.AreaCode};GroupName2={V.AreaCode}"
                             Case "C"
                                  strGroupName = "ReportType=false;GroupName={V.ServCode};GroupName2={V.ServCode}"
                             Case "D"
                                  strGroupName = "ReportType=false;GroupName={V.CustId};GroupName2={V.CustId}"
                             Case "E"
                                  strGroupName = "ReportType=false;GroupName={V.StrtCode};GroupName2={V.StrtCode}"
                             Case "F"
                                  strGroupName = "ReportType=false;GroupName={V.REALPERIOD};GroupName2={V.REALPERIOD}"
                             Case Else
                                  strGroupName = "ReportType=false;GroupName={V.CustId};GroupName2={V.CustId}"
                      End Select
                     strpagetype = "無"
          End Select
      Else
          Select Case True
                 Case optServCode.Value
                      GroupCode1 = "ServCode"
                      GroupName1 = "ServName"
                      GroupCode2 = "CitemCode"
                      GroupName2 = "CitemName"
                      strpagetype = "服務區"
                 Case optAreaCode.Value
                      GroupCode1 = "AreaCode"
                      GroupName1 = "AreaName"
                      GroupCode2 = "CitemCode"
                      GroupName2 = "CitemName"
                      strpagetype = "行政區"
                 Case optClctAreaCode.Value
                      GroupCode1 = "ClctAreaCode"
                      GroupName1 = "ClctAreaName"
                      GroupCode2 = "CitemCode"
                      GroupName2 = "CitemName"
                      strpagetype = "收費區"
                 Case optClctEn.Value
                      GroupCode1 = "ClctEn"
                      GroupName1 = "ClctName"
                      GroupCode2 = "CitemCode"
                      GroupName2 = "CitemName"
                      strpagetype = "收費人員"
                 Case optNothing.Value
                      GroupCode1 = "0"
                      GroupName1 = "'無'"
                      GroupCode2 = "CitemCode"
                      GroupName2 = "CitemName"
                      strpagetype = "無"
          End Select
      End If
    
    '排序方式
      If gimOrder.GetColumnOrderCode <> "" Then
        arrOrder = Split(gimOrder.GetColumnOrderCode, ",")
        intSort = 0
        For Each varOrder In arrOrder
          Select Case arrOrder(intSort)
                 Case "A"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.NODENO}"
                 Case "B"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.AreaCode}"
                 Case "C"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.ServCode}"
                 Case "D"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.CustId}"
                 Case "E"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.StrtCode}"
          End Select
          intSort = intSort + 1
        Next
      End If
      strGroupName = strGroupName & IIf(strGroupName = "", "", ";") & "BadCode = '" & strST2 & "'"
      
      strChooseString = "應收日期: " & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & _
                        "實收日期: " & subSpace(gdaRealDate1.GetValue(True)) & "~" & subSpace(gdaRealDate2.GetValue(True)) & ";" & _
                        "出單日期: " & subSpace(gdaCreateTime1.GetValue(True)) & "~" & subSpace(gdaCreateTime2.GetValue(True)) & ";" & _
                        "是否包含呆帳資料統計:" & IIf(chkBad.Value = 1, " 是", " 否") & ";" & _
                        "公司別　: " & subSpace(gilCompCode.GetDescription) & ";" & _
                        "服務類別: " & subSpace(gilServiceType.GetDescription) & ";" & _
                        "收費項目: " & subSpace(gimCitemCode.GetDispStr) & ";" & _
                        "單據類別: " & subSpace(gimBillType.GetDispStr) & ";" & _
                        "收費人員: " & subSpace(gimClctEn.GetDispStr) & ";" & _
                        "客戶類別: " & subSpace(gimClassCode.GetDispStr) & ";" & _
                        "客戶狀態: " & subSpace(gimCustStatus.GetDispStr) & ";" & _
                        "服務區　: " & subSpace(gimServCode.GetDispStr) & ";" & _
                        "行政區　: " & subSpace(gimAreaCode.GetDispStr) & ";" & _
                        "街道名稱: " & subSpace(gimStrtCode.GetDispStr) & ";" & _
                        "大樓名稱: " & subSpace(gimMduId.GetDispStr) & ";" & _
                        "未收原因: " & subSpace(gimUCCode.GetDispStr) & ";" & _
                        "付款種類: " & subSpace(gimPTCode.GetDispStr) & ";" & _
                        "分頁方式: " & strpagetype & ";" & _
                        IIf(optDetail.Value = True, "排序方式: " & subSpace(gimOrder.GetColumnOrderDspStr), "")
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub subPrint()
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim CountCust As Integer
    Dim CountBill As Integer
    Dim strGroupName1 As String
    Dim lngAffected As Long
        If Not GetRS(rsTmp, "Select count(*) intCount From " & strViewName & " Where RowNum = 1", gcnGi) Then Exit Sub
        If rsTmp("IntCount") = 0 Then
            MsgNoRcd
            SendSQL , , True
        Else
            If optDetail Then '明細
                strsql = "SELECT * From " & strViewName & " V "
                Call PrintRpt2(GetPrinterName(5), RptName("SO52E0", 2), , "應收明細表 [SO52E0A]", strsql, strChooseString, , True, , , strGroupName, GiPaperLandscape)
            Else
                strsql = GetComplexStr(strViewName, GroupCode1, giStringV, GroupName1, giStringV, GroupCode2, giLongV, GroupName2, , chkBad = 1)
                SendSQL strsql, True
                If Not CreateMDBTable(strsql, "So52E0A1", lngAffected) Then Exit Sub
                'Dim rsTmp2 As New adodb.Recordset
                'strGroupName1 = "BadCode = '" & strST & "'"
                'If Not GetRS(rsTmp2, "Select count(distinct CustId) CustCount,count(distinct BillNo) BillCount From " & strViewName) Then Exit Sub
                'CountCust = Val(rsTmp2("CustCount") & "")
                'CountBill = Val(rsTmp2("BillCount") & "")
                strsql = "Select * From So52E0A1"
                Call PrintRpt(GetPrinterName(5), RptName("SO52E0", 1), , "應收彙總表 [SO52E0A]", strsql, strChooseString, , True, "TMP0000.MDB", , , GiPaperLandscape)
            End If
        End If
        CloseRecordset rsTmp
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subPrint")
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

Private Sub cmdExit_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱")
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱")
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    Call SetgiMulti(gimCustStatus, "CodeNo", "Description", "CD035", "客戶狀態代碼", "客戶狀態名稱")
    Call SetgiMulti(gimMduId, "MduId", "Name", "SO017", "大樓編號", "大樓名稱")
    Call SetgiMulti(gimClctEn, "EmpNo", "EmpName", "CM003", "收費員代號", "收費員名稱")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "街道代碼", "街道名稱")
    Call SetgiMulti(gimUCCode, "CodeNo", "Description", "CD013", "未繳費原因代碼", "未繳費原因名稱")
    Call SetgiMultiAddItem(gimBillType, "B,I,P,M,T", "收費單,裝機單,停拆移機單,維修單,臨時收費單")
    Call SetgiMulti(gimPTCode, "CodeNo", "Description", "CD032", "付款種類代碼", "付款種類名稱")
    Call SetgiMultiAddItem(gimOrder, "A,B,C,D,E,F", "Node,行政區,服務區,客戶編號,街道編號,實收期數")
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
    Call FunctionKey(KeyCode, Shift, frmSO52E0A)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  gcnGi.Execute "Drop View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO52E0A
End Sub

Private Sub gdaShouldDate1_GotFocus()
  On Error GoTo ChkErr
    If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetValue (Format(RightDate, "YYYYMM") & "01")
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaShouldDate1_GotFocus")
End Sub

Private Sub gdaShouldDate2_GotFocus()
  On Error GoTo ChkErr
    If gdaShouldDate1.GetValue = "" Or gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue GetLastDate(gdaShouldDate1.GetValue(True))
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaShouldDate2_GotFocus")
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaShouldDate1, gdaShouldDate2)
  gdaShouldDate2_GotFocus
End Sub

Private Sub gdaRealDate1_GotFocus()
  On Error GoTo ChkErr
    If gdaRealDate1.GetValue = "" Then gdaRealDate1.SetValue (Format(RightDate, "YYYYMM") & "01")
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaRealDate1_GotFocus")
End Sub

Private Sub gdaRealDate2_GotFocus()
  On Error GoTo ChkErr
    If gdaRealDate1.GetValue = "" Or gdaRealDate2.GetValue = "" Then gdaRealDate2.SetValue GetLastDate(gdaRealDate1.GetValue(True))
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaRealDate2_GotFocus")
End Sub

Private Sub gdaRealDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaRealDate1, gdaRealDate2)
  gdaRealDate2_GotFocus
End Sub

Private Sub gdaCreateTime1_GotFocus()
  On Error GoTo ChkErr
    If gdaCreateTime1.GetValue = "" Then gdaCreateTime1.SetValue (Format(RightDate, "YYYYMM") & "01")
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaCreateTime1_GotFocus")
End Sub

Private Sub gdaCreateTime2_GotFocus()
  On Error GoTo ChkErr
    If gdaCreateTime1.GetValue = "" Or gdaCreateTime2.GetValue = "" Then gdaCreateTime2.SetValue GetLastDate(gdaCreateTime1.GetValue(True))
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
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimClctEn, , gilCompCode.GetCodeNo
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

Private Function subCreateView() As Boolean
  On Error Resume Next
    gcnGi.Execute "Drop View " & strViewName
  On Error GoTo ChkErr
  Dim strField As String
  Dim strWhere As String
  Dim strView As String
    strViewName = GetTmpViewName
    If optSummaric = True Then
        Select Case True
               Case optServCode.Value
                    strField = " A.CustId,A.BillNo,A.RealDate,A.CancelFlag,A.UCCode,A.ShouldAmt,A.RealAmt,A.RealPeriod,A.CitemCode,A.CitemName,A.STCode,A.AreaCode,'-' as AreaName,A.ClctEn,A.ClctName,A.ServCode,CD002.Description ServName,A.ClctAreaCode,'-' as ClctAreaName "
                    strWhere = ",CD002 Where CD002.CodeNo(+)=A.ServCode And "
               Case optAreaCode.Value
                    strField = " A.CustId,A.BillNo,A.RealDate,A.CancelFlag,A.UCCode,A.ShouldAmt,A.RealAmt,A.RealPeriod,A.CitemCode,A.CitemName,A.STCode,A.AreaCode,CD001.Description as AreaName,A.ClctEn,A.ClctName,A.ServCode,'-' as  ServName,A.ClctAreaCode,'-' as ClctAreaName "
                    strWhere = ",CD001 Where CD001.CodeNo(+)=A.AreaCode And "
               Case optClctAreaCode.Value
                    strField = " A.CustId,A.BillNo,A.RealDate,A.CancelFlag,A.UCCode,A.ShouldAmt,A.RealAmt,A.RealPeriod,A.CitemCode,A.CitemName,A.STCode,A.AreaCode,'-' as AreaName,A.ClctEn,A.ClctName,A.ServCode,'-' as  ServName,A.ClctAreaCode,CD040.Description as ClctAreaName "
                    strWhere = ",CD040 Where CD040.CodeNo(+)=A.ClctAreaCode And "
               Case Else
                    strField = " A.CustId,A.BillNo,A.RealDate,A.CancelFlag,A.UCCode,A.ShouldAmt,A.RealAmt,A.RealPeriod,A.CitemCode,A.CitemName,A.STCode,A.AreaCode,'-' as AreaName,A.ClctEn,A.ClctName,A.ServCode,'-' as  ServName,A.ClctAreaCode,'-' as ClctAreaName "
                    strWhere = " Where "
        End Select
        If blnSO002 Then strWhere = ",SO002" & strWhere & "A.Custid=SO002.Custid And A.ServiceType=SO002.ServiceType And "
        strView = "Create View " & strViewName & " as (" & _
                  "SELECT " & strField & "From SO033 A" & strWhere & strChoose & " Union All " & _
                  "SELECT " & strField & "From SO034 A" & strWhere & strChoose & " Union All " & _
                  "SELECT " & strField & "From SO035 A" & strWhere & strChoose & " Union All " & _
                  "SELECT " & strField & "From SO036 A" & strWhere & strChoose & " ) "
    Else
        strField = "SELECT Nvl(A.STCode,-1) STCode,Nvl(A.CancelFlag,0) CancelFlag,A.UCCode,A.CUSTID,A.BILLNO,A.CITEMCODE,A.CITEMNAME," & _
                   "A.REALSTARTDATE,A.REALSTOPDATE,A.SHOULDDATE,Nvl(A.SHOULDAMT,0) as ShouldAmt,A.REALDATE," & _
                   "A.REALAMT,A.OLDAMT,A.REALPERIOD,A.SERVCODE,SO014.SERVName,SO001.CUSTNAME,SO001.TEL1," & _
                   "SO002.CUSTSTATUSNAME,SO014.ADDRESS,SO014.NODENO,A.AREACODE,SO014.AREAName,A.ClctEn,A.ClctName,A.StrtCode,A.ClctAreaCode,SO001.ClctAreaName "
        strWhere = "Where A.CUSTID = SO001.CUSTID And A.ADDRNO = SO014.ADDRNO And A.Custid=SO002.Custid And A.ServiceType=SO002.ServiceType And " & strChoose
        strView = "Create View " & strViewName & " as (" & _
                  strField & "From SO001,SO002,SO014,SO033 A " & strWhere & " Union All " & _
                  strField & "From SO001,SO002,SO014,SO034 A " & strWhere & " Union All " & _
                  strField & "From SO001,SO002,SO014,SO035 A " & strWhere & " Union All " & _
                  strField & "From SO001,SO002,SO014,SO036 A " & strWhere & ")"
    End If
    gcnGi.Execute strView
    SendSQL strView
    subCreateView = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function

Private Sub gilServiceType_Change()
  On Error GoTo ChkErr
    GiMultiFilter gimCitemCode, gilServiceType.GetCodeNo
    GiMultiFilter gimClassCode, gilServiceType.GetCodeNo
    GiMultiFilter gimUCCode, gilServiceType.GetCodeNo
    GiMultiFilter gimPTCode, gilServiceType.GetCodeNo
    strST = GetCode("CD016", "<>1", gilServiceType.GetCodeNo, , , False)
    strST2 = GetCode("CD016", "=1", gilServiceType.GetCodeNo, , , False)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Public Function GetComplexStr(strsql As String, _
    Optional strFld1 As String = "0", Optional strFld1Type As giVType = giLongV, _
    Optional strFld2 As String = "0", Optional strFld2Type As giVType = giStringV, _
    Optional strFld3 As String = "0", Optional strFld3Type As giVType = giStringV, _
    Optional strFld4 As String = "0", Optional strFld4Type As giVType = giStringV, _
    Optional blnIncludeBad As Boolean = False) As String
    On Error Resume Next
    Dim strComplex As String
    Dim strField1 As String, strField2 As String, strField3 As String, strField4 As String
    Dim strTotalField1 As String, strTotalField2 As String, strTotalField3 As String, strTotalField4 As String
        strField1 = strFld1
        strField2 = strFld2
        strField3 = strFld3
        strField4 = strFld4
        strTotalField1 = strFld1
        strTotalField2 = strFld2
        strTotalField3 = "'99999999' " & strFld3
        strTotalField4 = "'小計:' " & strFld4
        
        If strFld1 <> "0" And strFld1Type = giLongV Then strField1 = " to_char(" & strFld1 & ") " & strFld1
        If strFld2 <> "0" And strFld2Type = giLongV Then strField2 = " to_char(" & strFld2 & ") " & strFld2
        If strFld3 <> "0" And strFld3Type = giLongV Then strField3 = " to_char(" & strFld3 & ") " & strFld3
        If strFld4 <> "0" And strFld4Type = giLongV Then strField4 = " to_char(" & strFld4 & ") " & strFld4
        
        If strFld1 = "0" Then strField1 = " '0'  Type1 ": strTotalField1 = strField1
        If strFld2 = "0" Then strField2 = " '0'  Type2 ": strTotalField2 = strField2
        If strFld3 = "0" Then strField3 = " '0'  Type3 ": strTotalField3 = " '9999999' Type3 "
        If strFld4 = "0" Then strField4 = " '0'  Type4 ": strTotalField4 = " '小計:' Type4 "
        
        strComplex = "Select " & strField1 & "," & strField2 & "," & IIf(strFld1Type = giLongV, " to_number(" & strFld3 & ")", strFld3) & "," & strField4 & _
            ",Sum(TotalAmt) TotalAmt,Sum(TotalCount) TotalCount,Sum(CountBillNo) CountBillNo,Sum(CountCustId) CountCustId,Sum(RealAmt) RealAmt " & _
            ",Sum(RealCount) RealCount,Sum(RealPeriod) RealPeriod,Sum(UCAmt) UCAmt,Sum(UCCount) UCCount " & _
            ",Sum(CancelAmt) CancelAmt,Sum(CancelCount) CancelCount,Sum(BadAmt) BadAmt,Sum(BadCount) BadCount From (" & _
            "select " & strField1 & "," & strField2 & "," & strField3 & "," & strField4 & _
            ",sum(ShouldAmt) TotalAmt,Count(*) TotalCount,0 CountBillNo,0 CountCustId," & _
            "Decode(Decode(Nvl(Length(SubStr(UCCode,1,1)),0),0,0,1),0,Decode(Nvl(Length(SubStr(RealDate,1,1)),0),0,0,Decode(Nvl(CancelFlag,0),1,0,Sum(RealAmt))),0) RealAmt," & _
            "Decode(Decode(Nvl(Length(SubStr(UCCode,1,1)),0),0,0,1),0,Decode(Nvl(Length(SubStr(RealDate,1,1)),0),0,0,Decode(Nvl(CancelFlag,0),1,0,Count(Distinct BillNo))),0) RealCount," & _
            "Decode(Decode(Nvl(Length(SubStr(UCCode,1,1)),0),0,0,1),0,Decode(Nvl(Length(SubStr(RealDate,1,1)),0),0,0,Decode(Nvl(CancelFlag,0),1,0,Count(RealPeriod))),0) RealPeriod," & _
            "Decode(Decode(Nvl(Length(SubStr(UCCode,1,1)),0),1,0,1),0,Decode(Nvl(Length(SubStr(RealDate,1,1)),0),1,Decode(Nvl(CancelFlag,0),0,Sum(ShouldAmt),0),Sum(ShouldAmt)),0) UCAmt," & _
            "Decode(Decode(Nvl(Length(SubStr(UCCode,1,1)),0),1,0,1),0,Decode(Nvl(Length(SubStr(RealDate,1,1)),0),1,Decode(Nvl(CancelFlag,0),0,Count(Distinct BillNo),0),Count(Distinct BillNo)),0) UCCount, " & _
            "Decode(Nvl(CancelFlag,0),1,Sum(ShouldAmt),0) CancelAmt," & _
            "Decode(Nvl(CancelFlag,0),1,Count(*),0) CancelCount  ,0 BadAmt,0 BadCount" & _
            " From " & strsql & _
            " Group by " & strFld1 & "," & strFld2 & "," & strFld3 & "," & strFld4 & ",Length(SubStr(UCCode,1,1)),Length(SubStr(RealDate,1,1)),Nvl(CancelFlag,0)" & _
            " Union All " & _
            " Select " & strField1 & "," & strField2 & "," & strField3 & "," & strField4 & ", 0 TotalAmt,0 TotalCount, Count(Distinct BillNo) CountBillNo,Count(Distinct CustId) CountCustId, " & _
            " 0 RealAmt,0 RealCount,0 RealPeriod, 0 UCAmt,0 UCCount,0 CancelAmt, 0 CancelCount ,0 BadAmt,0 BadCount " & _
            " From " & strsql & " Group By " & strFld1 & "," & strFld2 & "," & strFld3 & "," & strFld4


        '合計
'        strComplex = strComplex & _
'            " Union All " & _
'            " Select " & strTotalField1 & "," & strTotalField2 & "," & strTotalField3 & "," & strTotalField4 & _
'            ",sum(ShouldAmt) TotalAmt,Count(*) TotalCount, 0 CountBillNo,0 CountCustId," & _
'            "Decode(Decode(Nvl(Length(SubStr(UCCode,1,1)),0),0,0,1),0,Decode(Nvl(Length(SubStr(RealDate,1,1)),0),0,0,Decode(Nvl(CancelFlag,0),1,0,Sum(RealAmt))),0) RealAmt," & _
'            "Decode(Decode(Nvl(Length(SubStr(UCCode,1,1)),0),0,0,1),0,Decode(Nvl(Length(SubStr(RealDate,1,1)),0),0,0,Decode(Nvl(CancelFlag,0),1,0,Count(*))),0) RealCount," & _
'            "Decode(Decode(Nvl(Length(SubStr(UCCode,1,1)),0),0,0,1),0,Decode(Nvl(Length(SubStr(RealDate,1,1)),0),0,0,Decode(Nvl(CancelFlag,0),1,0,Count(RealPeriod))),0) RealPeriod," & _
'            "Decode(Decode(Nvl(Length(SubStr(UCCode,1,1)),0),1,0,1),0,Decode(Nvl(Length(SubStr(RealDate,1,1)),0),1,0,Sum(ShouldAmt)),0) UCAmt," & _
'            "Decode(Decode(Nvl(Length(SubStr(UCCode,1,1)),0),1,0,1),0,Decode(Nvl(Length(SubStr(RealDate,1,1)),0),1,0,Count(*)),0) UCCount, " & _
'            "Decode(Nvl(CancelFlag,0),1,Sum(ShouldAmt),0) CancelAmt," & _
'            "Decode(Nvl(CancelFlag,0),1,Count(*),0) CancelCount ,0 BadAmt,0 BadCount " & _
'            " From " & strSql & _
'            " Group By Length(SubStr(UCCode,1,1)),Length(SubStr(RealDate,1,1)),Nvl(CancelFlag,0)" & _
'            " Union All " & _
'            " Select " & strTotalField1 & "," & strTotalField2 & "," & strTotalField3 & "," & strTotalField4 & _
'            ", 0 TotalAmt,0 TotalCount, Count(Distinct BillNo) CountBillNo,Count(Distinct CustId) CountCustId, " & _
'            " 0 RealAmt,0 RealCount,0 RealPeriod, 0 UCAmt,0 UCCount,0 CancelAmt, 0 CancelCount ,0 BadAmt,0 BadCount " & _
'            " From " & strSql
            
        strComplex = strComplex & _
            " Union All " & _
            " Select " & strTotalField1 & "," & strTotalField2 & "," & strTotalField3 & "," & strTotalField4 & _
            ",sum(ShouldAmt) TotalAmt,Count(*) TotalCount, 0 CountBillNo,0 CountCustId," & _
            "Decode(Decode(Nvl(Length(SubStr(UCCode,1,1)),0),0,0,1),0,Decode(Nvl(Length(SubStr(RealDate,1,1)),0),0,0,Decode(Nvl(CancelFlag,0),1,0,Sum(RealAmt))),0) RealAmt," & _
            "Decode(Decode(Nvl(Length(SubStr(UCCode,1,1)),0),0,0,1),0,Decode(Nvl(Length(SubStr(RealDate,1,1)),0),0,0,Decode(Nvl(CancelFlag,0),1,0,Count(Distinct BillNo))),0) RealCount," & _
            "Decode(Decode(Nvl(Length(SubStr(UCCode,1,1)),0),0,0,1),0,Decode(Nvl(Length(SubStr(RealDate,1,1)),0),0,0,Decode(Nvl(CancelFlag,0),1,0,Count(RealPeriod))),0) RealPeriod," & _
            "Decode(Decode(Nvl(Length(SubStr(UCCode,1,1)),0),1,0,1),0,Decode(Nvl(Length(SubStr(RealDate,1,1)),0),1,Decode(Nvl(CancelFlag,0),0,Sum(ShouldAmt),0),Sum(ShouldAmt)),0) UCAmt," & _
            "Decode(Decode(Nvl(Length(SubStr(UCCode,1,1)),0),1,0,1),0,Decode(Nvl(Length(SubStr(RealDate,1,1)),0),1,Decode(Nvl(CancelFlag,0),0,Count(Distinct BillNo),0),Count(Distinct BillNo)),0) UCCount, " & _
            "Decode(Nvl(CancelFlag,0),1,Sum(ShouldAmt),0) CancelAmt," & _
            "Decode(Nvl(CancelFlag,0),1,Count(*),0) CancelCount ,0 BadAmt,0 BadCount " & _
            " From " & strsql & _
            " Group By " & strFld1 & "," & strFld2 & ",Length(SubStr(UCCode,1,1)),Length(SubStr(RealDate,1,1)),Nvl(CancelFlag,0)" & _
            " Union All " & _
            " Select " & strTotalField1 & "," & strTotalField2 & "," & strTotalField3 & "," & strTotalField4 & _
            ", 0 TotalAmt,0 TotalCount, Count(Distinct BillNo) CountBillNo,Count(Distinct CustId) CountCustId, " & _
            " 0 RealAmt,0 RealCount,0 RealPeriod, 0 UCAmt,0 UCCount,0 CancelAmt, 0 CancelCount ,0 BadAmt,0 BadCount " & _
            " From " & strsql & " Group By " & strFld1 & "," & strFld2
            
            
        If blnIncludeBad Then
'            strComplex = strComplex & _
'            " Union All Select " & strField1 & "," & strField2 & "," & strField3 & "," & strField4 & ", 0 TotalAmt,0 TotalCount, 0 CountBillNo,0 CountCustId, " & _
'            " 0 RealAmt,0 RealCount,0 RealPeriod, 0 UCAmt,0 UCCount,0 CancelAmt, 0 CancelCount ,Sum(ShouldAmt) BadAmt,Count(*) BadCount " & _
'            " From " & strSql & IIf(InStr(UCase(strSql), "WHERE") > 0, " And ", " Where ") & _
'            " STCode In (Select CodeNo From CD016 Where RefNo=1) " & _
'            "  Group By " & strFld1 & "," & strFld2 & "," & strFld3 & "," & strFld4 & _
'            " Union All " & _
'            " Select " & strTotalField1 & "," & strTotalField2 & "," & strTotalField3 & "," & strTotalField4 & _
'            ", 0 TotalAmt,0 TotalCount, 0 CountBillNo,0 CountCustId, " & _
'            " 0 RealAmt,0 RealCount,0 RealPeriod, 0 UCAmt,0 UCCount,0 CancelAmt, 0 CancelCount ,Sum(ShouldAmt) BadAmt,Count(*) BadCount " & _
'            " From " & strSql & IIf(InStr(UCase(strSql), "WHERE") > 0, " And ", " Where ") & _
'            " STCode In (Select CodeNo From CD016 Where RefNo=1) "

            strComplex = strComplex & _
            " Union All Select " & strField1 & "," & strField2 & "," & strField3 & "," & strField4 & ", 0 TotalAmt,0 TotalCount, 0 CountBillNo,0 CountCustId, " & _
            " 0 RealAmt,0 RealCount,0 RealPeriod, 0 UCAmt,0 UCCount,0 CancelAmt, 0 CancelCount ,Sum(ShouldAmt) BadAmt,Count(Distinct BillNo) BadCount " & _
            " From " & strsql & IIf(InStr(UCase(strsql), "WHERE") > 0, " And ", " Where ") & _
            " STCode In (Select CodeNo From CD016 Where RefNo=1) " & _
            "  Group By " & strFld1 & "," & strFld2 & "," & strFld3 & "," & strFld4 & _
            " Union All " & _
            " Select " & strTotalField1 & "," & strTotalField2 & "," & strTotalField3 & "," & strTotalField4 & _
            ", 0 TotalAmt,0 TotalCount, 0 CountBillNo,0 CountCustId, " & _
            " 0 RealAmt,0 RealCount,0 RealPeriod, 0 UCAmt,0 UCCount,0 CancelAmt, 0 CancelCount ,Sum(ShouldAmt) BadAmt,Count(Distinct BillNo) BadCount " & _
            " From " & strsql & IIf(InStr(UCase(strsql), "WHERE") > 0, " And ", " Where ") & _
            " STCode In (Select CodeNo From CD016 Where RefNo=1) Group By " & strFld1 & "," & strFld2
        End If
        strComplex = strComplex & " ) Group By " & IIf(strFld1Type = giLongV, " to_number(" & strFld1 & ")", strFld1) & "," & strFld2 & "," & strFld3 & "," & strFld4
        GetComplexStr = strComplex
End Function

Public Function CreateMDBTable(strsql As String, StrTableName As String, Optional ByRef lngAffected As Long) As Boolean
  On Error Resume Next
      cnn.Execute "Drop Table " & StrTableName
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
      cnn.Execute "Create Table " & StrTableName & " (GroupField1 Text(40),GroupField2 Text(40)," & _
              " GroupField3 Long,GroupField4 Text(40)," & _
              "TotalAmt Double,TotalCount Double,CountBillNo Double,CountCustId Single," & _
              "RealAmt Double,RealCount Double,RealPeriod Double,UCAmt Double,UCCount Double,CancelAmt Double,CancelCount Double,BadAmt Double,BadCount Double ) "
      If Not GetRS(rsTmp, strsql, gcnGi) Then Exit Function
      lngAffected = rsTmp.RecordCount
      If Not InsertToSummaric(rsTmp, StrTableName) Then Exit Function
      
      strsql = "Select chr(255) & '99' as GroupField1,'   ' as GroupField2,99999999 as GroupField3,'總計:' as GroupField4,Sum(TotalAmt),Sum(TotalCount),Sum(CountBillNo),Sum(CountCustId),Sum(RealAmt),Sum(RealCount),Sum(RealPeriod),Sum(UCAmt),Sum(UCCount),Sum(CancelAmt),Sum(CancelCount),Sum(BadAmt),Sum(BadCount) " & _
               "From " & StrTableName & " Where GroupField3=99999999"

      If Not GetRS(rsTmp, strsql, cnn) Then Exit Function
      If Not InsertToSummaric(rsTmp, StrTableName) Then Exit Function
      
      
      CreateMDBTable = True
  Exit Function
ChkErr:
  lngAffected = 0
  Call ErrSub(Me.Name, "CreateMDBTable")
End Function

Public Function InsertToSummaric(rs As ADODB.Recordset, StrTableName As String) As Boolean
  On Error GoTo ChkErr
      cnn.BeginTrans
      Do While Not rs.EOF
          cnn.Execute "Insert Into " & StrTableName & " (GroupField1,GroupField2,GroupField3,GroupField4," & _
                  " TotalAmt,TotalCount,CountBillNo,CountCustId," & _
                  "RealAmt,RealCount,RealPeriod,UCAmt,UCCount,CancelAmt,CancelCount,BadAmt,BadCount) Values (" & _
                    GetNullString(rs(0)) & "," & GetNullString(rs(1)) & "," & _
                    GetNullString(rs(2)) & "," & GetNullString(rs(3)) & "," & _
                    GetNullString(rs(4)) & "," & GetNullString(rs(5)) & "," & _
                    GetNullString(rs(6)) & "," & GetNullString(rs(7)) & "," & _
                    GetNullString(rs(8)) & "," & GetNullString(rs(9)) & "," & _
                    GetNullString(rs(10)) & "," & GetNullString(rs(11)) & "," & _
                    GetNullString(rs(12)) & "," & GetNullString(rs(13)) & "," & _
                    GetNullString(rs(14)) & "," & GetNullString(rs(15)) & "," & _
                    GetNullString(rs(16)) & ")"
          rs.MoveNext
      Loop
      cnn.CommitTrans
      InsertToSummaric = True
    Exit Function
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "InsertToSummaric")
End Function

'Public Function InsertToSummaric(strSql As String, StrTableName As String, Optional ByRef lngAffected As Long) As Boolean
'    On Error Resume Next
'        cnn.Execute "Drop Table " & StrTableName
'    On Error GoTo ChkErr
'    Dim rsTmp As New ADODB.Recordset
'        cnn.Execute "Create Table " & StrTableName & " (GroupField1 Text(40),GroupField2 Text(40)," & _
'                " GroupField3 Long,GroupField4 Text(40)," & _
'                "TotalAmt Double,TotalCount Double,CountBillNo Double,CountCustId Single," & _
'                "RealAmt Double,RealCount Double,RealPeriod Double,UCAmt Double,UCCount Double,CancelAmt Double,CancelCount Double,BadAmt Double,BadCount Double ) "
'
'        If Not GetRS(rsTmp, strSql, gcnGi) Then Exit Function
'        lngAffected = rsTmp.RecordCount
'        Do While Not rsTmp.EOF
'            cnn.Execute "Insert Into " & StrTableName & " (GroupField1,GroupField2,GroupField3,GroupField4," & _
'                    " TotalAmt,TotalCount,CountBillNo,CountCustId," & _
'                    "RealAmt,RealCount,RealPeriod,UCAmt,UCCount,CancelAmt,CancelCount,BadAmt,BadCount) Values (" & _
'                      GetNullString(rsTmp(0)) & "," & GetNullString(rsTmp(1)) & "," & _
'                      GetNullString(rsTmp(2)) & "," & GetNullString(rsTmp(3)) & "," & _
'                      GetNullString(rsTmp(4)) & "," & GetNullString(rsTmp(5)) & "," & _
'                      GetNullString(rsTmp(6)) & "," & GetNullString(rsTmp(7)) & "," & _
'                      GetNullString(rsTmp(8)) & "," & GetNullString(rsTmp(9)) & "," & _
'                      GetNullString(rsTmp(10)) & "," & GetNullString(rsTmp(11)) & "," & _
'                      GetNullString(rsTmp(12)) & "," & GetNullString(rsTmp(13)) & "," & _
'                      GetNullString(rsTmp(14)) & "," & GetNullString(rsTmp(15)) & "," & _
'                      GetNullString(rsTmp(16)) & ")"
'            rsTmp.MoveNext
'        Loop
'        InsertToSummaric = True
'    Exit Function
'ChkErr:
'    lngAffected = 0
'    ErrSub Me.Name, "InsertToSummaric"
'End Function
