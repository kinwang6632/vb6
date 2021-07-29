VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.17#0"; "csMulti.ocx"
Begin VB.Form frmSO5260A 
   BorderStyle     =   1  '單線固定
   Caption         =   "預估未來一年收費客戶及金額統計[ SO5260A ]"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   Icon            =   "SO5260A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
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
      Left            =   1710
      TabIndex        =   37
      Top             =   6060
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
      Left            =   6840
      TabIndex        =   36
      Top             =   6060
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
      Left            =   120
      TabIndex        =   35
      Top             =   6060
      Width           =   1245
   End
   Begin VB.CommandButton cmdToExcel 
      Caption         =   "匯成Excel"
      Height          =   525
      Left            =   3480
      TabIndex        =   34
      Top             =   6060
      Width           =   1305
   End
   Begin Gi_YM.GiYM gymBehind1 
      Height          =   345
      Left            =   1410
      TabIndex        =   3
      Top             =   510
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   609
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
   Begin VB.Frame Frame4 
      Caption         =   "網路編號"
      Height          =   735
      Left            =   0
      TabIndex        =   31
      Top             =   4050
      Width           =   8115
      Begin VB.OptionButton optCircuitNo 
         Caption         =   "只印空白網路編號"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   4830
         TabIndex        =   14
         Top             =   240
         Width           =   1845
      End
      Begin VB.OptionButton optAll 
         Caption         =   "全部"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6660
         TabIndex        =   15
         Top             =   270
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.TextBox mskCircuitNo 
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   90
         TabIndex        =   13
         Top             =   240
         Width           =   4605
      End
   End
   Begin VB.PictureBox pic1 
      BackColor       =   &H00FFFFFF&
      Height          =   990
      Left            =   1500
      ScaleHeight     =   930
      ScaleWidth      =   5055
      TabIndex        =   28
      Top             =   1980
      Visible         =   0   'False
      Width           =   5115
      Begin MSComctlLib.ProgressBar pgbBar 
         Height          =   375
         Left            =   165
         TabIndex        =   29
         Top             =   390
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblPmt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "程序處理中，請稍後 ..."
         Height          =   180
         Left            =   210
         TabIndex        =   30
         Top             =   135
         Width           =   1800
      End
   End
   Begin CS_Multi.CSmulti gimStrtCode 
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   3270
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   609
      ButtonCaption   =   "街  道  範  圍"
   End
   Begin VB.Frame fraPage 
      BackColor       =   &H00E0E0E0&
      Caption         =   "小計依據"
      ForeColor       =   &H00808000&
      Height          =   855
      Left            =   0
      TabIndex        =   25
      Top             =   4830
      Width           =   8115
      Begin VB.OptionButton optPeriod 
         BackColor       =   &H00E0E0E0&
         Caption         =   "期數+金額累計"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton optClctAreaCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "收費區"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   4530
         TabIndex        =   20
         Top             =   330
         Width           =   840
      End
      Begin VB.OptionButton optClassCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "客戶類別"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   6540
         TabIndex        =   22
         Top             =   330
         Width           =   1035
      End
      Begin VB.OptionButton optClctDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "下次收費年月"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   330
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.OptionButton optStrtCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "街道編號"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   5460
         TabIndex        =   21
         Top             =   330
         Width           =   1065
      End
      Begin VB.OptionButton optClctEn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "收費人員"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1560
         TabIndex        =   17
         Top             =   330
         Width           =   1035
      End
      Begin VB.OptionButton optAreaCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "行政區"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   3585
         TabIndex        =   19
         Top             =   330
         Width           =   840
      End
      Begin VB.OptionButton optServCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "服務區"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   2670
         TabIndex        =   18
         Top             =   330
         Width           =   840
      End
   End
   Begin Gi_Date.GiDate gdaClctDate2 
      Height          =   345
      Left            =   2670
      TabIndex        =   1
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   255
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
      Height          =   345
      Left            =   1410
      TabIndex        =   0
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   255
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
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   345
      Left            =   0
      TabIndex        =   10
      Top             =   2880
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   609
      ButtonCaption   =   "行     政     區"
   End
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   345
      Left            =   0
      TabIndex        =   9
      Top             =   2490
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   609
      ButtonCaption   =   "服     務     區"
   End
   Begin Gi_Multi.GiMulti gimCMCode 
      Height          =   345
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   609
      ButtonCaption   =   "收  費  方  式"
   End
   Begin Gi_Multi.GiMulti gimStatusCode 
      Height          =   345
      Left            =   0
      TabIndex        =   8
      Top             =   2100
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   609
      ButtonCaption   =   "客  戶  狀  態"
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   4680
      TabIndex        =   2
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
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   4680
      TabIndex        =   4
      Top             =   510
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
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   609
      ButtonCaption   =   "收  費  項  目"
   End
   Begin CS_Multi.CSmulti gimClassCode 
      Height          =   345
      Left            =   0
      TabIndex        =   7
      Top             =   1710
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   609
      ButtonCaption   =   "客  戶  類  別"
   End
   Begin CS_Multi.CSmulti gimOldClctEn 
      Height          =   345
      Left            =   0
      TabIndex        =   12
      Top             =   3660
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   609
      ButtonCaption   =   "原 收 費 人 員"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "回溯年月"
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
      Left            =   600
      TabIndex        =   32
      Top             =   600
      Width           =   780
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3870
      TabIndex        =   27
      Top             =   570
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
      Left            =   3870
      TabIndex        =   26
      Top             =   150
      Width           =   765
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   2550
      TabIndex        =   24
      Top             =   210
      Width           =   105
   End
   Begin VB.Label lblShouldDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收費日預估範圍"
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
      Left            =   30
      TabIndex        =   23
      Top             =   180
      Width           =   1365
   End
End
Attribute VB_Name = "frmSO5260A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table=SO003,SO001,SO002,SO014
'DataBase=MDB [SO5260A1][SO5260A2]
Option Explicit
Dim MdbChoose As String
Dim strTables As String
Dim strWhere As String
Dim GroupField As String
Dim strGroupBy As String
Dim strBehindWhere As String
Dim rsTemp As New ADODB.Recordset
Dim blnExcel As Boolean
Dim strPage As String
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
    If Not optPeriod Then
         Call PreviousRpt(GetPrinterName(5), RptName("SO5260"), Me.Caption)
    Else
         Call PreviousRpt(GetPrinterName(5), RptName("SO5260", 2), Me.Caption)
    End If
   
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
  Dim strSubQry1 As String
  Dim strSubQry2 As String
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    Screen.MousePointer = vbHourglass
      cmdExit.SetFocus
      Me.Enabled = False
        ReadyGoPrint
        pic1.Visible = True
          Call subChoose
          '由2801問題集，改為動態產生Table
          Call CreateTable
          subInsertSO5260A1
          '問題集 2801增加回塑年月
          If strBehindWhere <> "" Then
            Call InsertTemp
          End If
          '新增到SO5260A1的CountCust
          subInsertCountCust
          subInsertSO5260A2
        pic1.Visible = False
        Set rsTmp = cnn.Execute("SELECT Count(*) as intCount From SO5260A1")
        If rsTmp("intCount") = 0 Then
           MsgNoRcd
           SendSQL , , True
        Else
           InsertDataSO5260A4
           If optPeriod Then
                strsql = "Select * from SO5260A4"
           Else
                strsql = "Select * from SO5260A1"
           End If
           'strsql = "Select * From SO5260A1 M"
           strSubQry1 = "SELECT ClassCode1,ClassName,CustId FROM SO5260A2 M"
           'strSubQry2 = "Select * from SO5260A4 M order by Period"
           If blnExcel Then
               Call toExcel
               
           Else
              If Not optPeriod Then
                   Call PrintRpt(GetPrinterName(5), RptName("SO5260"), , Me.Caption, strsql, strChooseString, , True, "Tmp0000.Mdb", , strGroupName, GiPaperLandscape, , strSubQry1)
              Else
                   Call PrintRpt(GetPrinterName(5), RptName("SO5260", 2), , Me.Caption, strsql, strChooseString, , True, "Tmp0000.Mdb", , strGroupName, GiPaperLandscape, , strSubQry1)
              End If
           End If
           
        End If
      Me.Enabled = True
    Screen.MousePointer = vbDefault
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 66
    If gilServiceType.GetCodeNo = "" Then gilServiceType.SetFocus: strErrFile = "服務類別": GoTo 66
    If gdaClctDate1.GetValue = "" Then gdaClctDate1.SetFocus: strErrFile = "收費日預估起始日": GoTo 66
    If gdaClctDate2.GetValue = "" Then gdaClctDate2.SetFocus: strErrFile = "收費日預估截止日": GoTo 66
    
    '收費日預估範圍必須為未來日期
    If CDate(gdaClctDate1.GetValue(True)) < RightDate Then
        MsgBox "收費日預估範圍必須為未來日期", , "錯誤訊息"
        gdaClctDate1.SetValue ""
        gdaClctDate1_GotFocus
        Exit Function
    End If
    
    '收費日預估範圍不能超過一年
    If gdaClctDate2.GetValue(True) >= Format(DateAdd("yyyy", 1, gdaClctDate1.Text), "YYYY/MM/DD") Then
        MsgBox "收費日預估範圍不能超過一年", , "錯誤訊息"
        gdaClctDate2_GotFocus
        Exit Function
    End If
    '回溯年月不得超過一年
    If gymBehind1.GetValue <> "" Then
        If DateAdd("YYYY", 1, gymBehind1.GetValue(True) & "/01") < gdaClctDate1.GetValue(True) Then
            MsgBox "回溯年月不能超過一年", vbInformation, "警告訊息"
            gymBehind1.Text = ""
            gymBehind1.SetFocus
            Exit Function
        End If
    End If
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
    Dim blnSO014 As Boolean
    Dim blnSO002 As Boolean
    'strChoose = "SO002.CustStatusCode='1'"
    strChoose = ""
    strChooseString = ""
    GroupField = ""
    strGroupBy = ""
    strTables = "SO003,SO001"
    strWhere = "SO003.Custid=SO001.Custid And SO003.CompCode = SO001.CompCode "
    blnSO014 = False
    blnSO002 = False
         
'  '日期
'    If gdaClctDate1.GetValue <> "" Then Call subAnd("SO003.ClctDate >= To_Date('" & gdaClctDate1.GetValue & "','YYYYMMDD')")
'    If gdaClctDate2.GetValue <> "" Then Call subAnd("SO003.ClctDate < To_Date('" & gdaClctDate2.GetValue & "','YYYYMMDD')+1")
    
  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO001.CompCode=" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("instr(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "')>0")
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("Instr(Nvl(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "'),'" & gilServiceType.GetCodeNo & "') > 0")
    
  'GiMulti
    If gimCitemCode.GetQueryCode <> "" Then Call subAnd("SO003.CitemCode IN (" & gimCitemCode.GetQueryCode & ")")
    'If gimCitemCode.GetQryStr <> "" Then Call subAnd("SO003.CitemCode " & gimCitemCode.GetQryStr)
    If gimCMCode.GetQryStr <> "" Then Call subAnd("SO002.CMCode " & gimCMCode.GetQryStr): blnSO002 = True
    If gimClassCode.GetQryStr <> "" Then Call subAnd("SO001.ClassCode1 " & gimClassCode.GetQryStr)
    If gimStatusCode.GetQryStr <> "" Then Call subAnd("SO002.CustStatusCode " & gimStatusCode.GetQryStr): blnSO002 = True
    If gimServCode.GetQryStr <> "" Then Call subAnd("SO001.ServCode " & gimServCode.GetQryStr)
    If gimAreaCode.GetQryStr <> "" Then Call subAnd("SO014.AreaCode " & gimAreaCode.GetQryStr): blnSO014 = True
    If gimStrtCode.GetQryStr <> "" Then Call subAnd("SO014.StrtCode " & gimStrtCode.GetQryStr): blnSO014 = True
    If gimOldClctEn.GetQryStr <> "" Then Call subAnd("SO014.ClctEn " & gimOldClctEn.GetQryStr): blnSO014 = True
    
  '網路編號
    If mskCircuitNo.Text <> "" Then
        Call subAnd("SO014.CircuitNo='" & mskCircuitNo.Text & "'")
        blnSO014 = True
    Else
       If optCircuitNo.Value Then Call subAnd("SO014.CircuitNo is null "): blnSO014 = True
    End If

    '小計依據
    Select Case True
           Case optClctDate.Value
                strGroupName = "Type=1"
                GroupField = "To_Char(SO003.ClctDate,'YYYYMM') as GroupCode,To_Char(SO003.ClctDate,'YYYY/MM') as GroupName"
                strGroupBy = "To_Char(SO003.ClctDate,'YYYYMM'),To_Char(SO003.ClctDate,'YYYY/MM')"
                strPage = "下次收費年月"

           Case optClctEn.Value
                strGroupName = "Type=2"
                GroupField = "nvl(SO014.ClctEn,'XXXX') as GroupCode,nvl(SO014.ClctName,'XXXX') as GroupName"
                strGroupBy = "SO014.ClctEn,SO014.ClctName"
                strPage = "收費人員"
                blnSO014 = True
               
           Case optServCode.Value
                strGroupName = "Type=3"
                GroupField = "nvl(SO001.ServCode,'XXXX') as GroupCode,nvl(SO001.ServArea,'XXXX') as GroupName"
                strGroupBy = "SO001.ServCode,SO001.ServArea"
                strPage = "服務區"

           Case optAreaCode.Value
                strGroupName = "Type=4"
                GroupField = "Nvl(SO014.AreaCode,'XXXX') as GroupCode,nvl(SO014.AreaName,'XXXX') as GroupName"
                strGroupBy = "SO014.AreaCode,SO014.AreaName"
                strPage = "行政區"
                blnSO014 = True
               
           Case optClctAreaCode.Value
                strGroupName = "Type=7"
                GroupField = "Nvl(SO001.ClctAreaCode,'XXXX') as GroupCode,nvl(SO001.ClctAreaName,'XXXX') as GroupName"
                strGroupBy = "SO001.ClctAreaCode,SO001.ClctAreaName"
                strPage = "收費區"
               
           Case optStrtCode.Value
                strGroupName = "Type=5"
                GroupField = "nvl(Trim(To_Char(SO014.StrtCode)),'XXXX') as GroupCode,nvl(SO014.StrtName,'XXXX') as GroupName"
                strGroupBy = "SO014.StrtCode,SO014.StrtName"
                strPage = "街道編號"
                blnSO014 = True
               
           Case optClassCode.Value
                strGroupName = "Type=6"
                GroupField = "Nvl(Trim(To_Char(SO001.ClassCode1)),'XXXX') as GroupCode,Nvl(SO001.ClassName1,'XXXX') as GroupName"
                strGroupBy = "SO001.ClassCode1,SO001.ClassName1"
                strPage = "客戶類別"
           Case optPeriod.Value
                strGroupName = "Type=8"
                GroupField = "NVL(SO003.Period,0) as GroupCode,NVL(SO003.Period,0) as GroupName"
                strGroupBy = "SO003.Period"
                strPage = "期數+累計金額"
    End Select
   
    strGroupName = strGroupName & ";1thMonth=Date(" & Year(gdaClctDate1.GetValue(True)) & "," & Month(gdaClctDate1.GetValue(True)) & ",01)"
   
  '串SO002
    If blnSO002 Then
        strTables = strTables & ",SO002"
        strWhere = strWhere & " And SO003.Custid=SO002.Custid And SO003.CompCode = SO002.CompCode And SO003.ServiceType = SO002.ServiceType "
        If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO002.ServiceType='" & gilServiceType.GetCodeNo & "'")
    End If
    
  '串SO014
    If blnSO014 Then
        strTables = strTables & ",SO014"
        strWhere = strWhere & " And SO001.InstAddrNo=SO014.AddrNo And SO001.CompCode = SO014.CompCode "
    End If
    '問題集2801 增加尋找前期金額
    If gymBehind1.GetValue <> "" Then
       strBehindWhere = " And " & strChoose & " And SO003.ClctDate>=To_Date('" & Format(gymBehind1.GetValue(True) & "/01", "YYYYMMDD") & "','YYYYMMDD')" & _
                        " And SO003.ClctDate<To_Date('" & gdaClctDate1.GetValue & "','YYYYMMDD')"
              
    Else
      strBehindWhere = ""
    End If
    
     '日期
    If gdaClctDate1.GetValue <> "" Then Call subAnd("SO003.ClctDate >= To_Date('" & gdaClctDate1.GetValue & "','YYYYMMDD')")
    If gdaClctDate2.GetValue <> "" Then Call subAnd("SO003.ClctDate < To_Date('" & gdaClctDate2.GetValue & "','YYYYMMDD')+1")

    If strChoose <> "" Then
        strChoose = " And " & strChoose
    End If


    strChooseString = "收費日預估範圍: " & subSpace(gdaClctDate1.GetOriginalValue) & "~" & subSpace(gdaClctDate2.GetOriginalValue) & ";" & _
                      "公司別　: " & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別: " & subSpace(gilServiceType.GetDescription) & ";" & _
                      "收費項目: " & subSpace(gimCitemCode.GetDispStr) & ";" & _
                      "收費方式: " & subSpace(gimCMCode.GetDispStr) & ";" & _
                      "客戶類別: " & subSpace(gimClassCode.GetDispStr) & ";" & _
                      "客戶狀態: " & subSpace(gimStatusCode.GetDispStr) & ";" & _
                      "服務區　: " & subSpace(gimServCode.GetDispStr) & ";" & _
                      "行政區　: " & subSpace(gimAreaCode.GetDispStr) & ";" & _
                      "街道範圍: " & subSpace(gimStrtCode.GetDispStr) & ";" & _
                      "綱路編號: " & subSpace(mskCircuitNo.Text) & ";" & _
                      "小計依據: " & subSpace(strPage)

  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub cmdToExcel_Click()
   blnExcel = True
   Call cmdPrint_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5260A)
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
    cmdToExcel.Enabled = False
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱")
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱")
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    Call SetgiMulti(gimStatusCode, "CodeNo", "Description", "CD035", "客戶狀態代碼", "客戶狀態名稱")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "街道代碼", "街道名稱")
    Call SetgiMulti(gimOldClctEn, "EmpNO", "EmpName", "CM003", "代碼", "名稱")
    gimStatusCode.SetQueryCode "1"
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

Private Sub subInsertSO5260A1()
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strMDB As String
      cnn.Execute "DELETE * FROM SO5260A1"
      If cnn.State = adStateClosed Then Set cnn = GetTmpMdbCn
      If rsTmp.State = 1 Then rsTmp.Close
        pgbBar = 1
            strMDB = "SELECT " & GroupField & ",SO003.ClctDate,SO003.Period,SO003.Amount,Sum(SO003.Amount) as SumAmount " & _
                     " From " & strTables & " Where " & strWhere & strChoose & " Group By " & strGroupBy & ",SO003.ClctDate,SO003.Period,SO003.Amount"
        Set rsTmp = gcnGi.Execute(strMDB)
        SendSQL strMDB
        
        While Not rsTmp.EOF
            pgbBar = rsTmp.AbsolutePosition / rsTmp.RecordCount * 90
 
            If ChkDataExist(rsTmp) = True Then
                Call UpdateData(rsTmp)
            Else
                Call InsertData(rsTmp)
            End If
            rsTmp.MoveNext
            DoEvents
        Wend
        CloseRecordset rsTmp
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subInsertSO5260A1")
End Sub

Private Sub subInsertCountCust()
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strMDB As String
    Dim strTempMDB As String
      If cnn.State = adStateClosed Then Set cnn = GetTmpMdbCn
      If rsTmp.State = 1 Then rsTmp.Close
      '將得到的客戶數加到SO5260A1
        strMDB = "SELECT " & GroupField & ",SO003.Period,SO003.Amount,Count(Distinct SO001.CustId) as CountCust " & _
                 " From " & strTables & " Where " & strWhere & strChoose & " Group By " & strGroupBy & ",SO003.Period,SO003.Amount"
                    
        Set rsTmp = gcnGi.Execute(strMDB)
        SendSQL strMDB
        
        While Not rsTmp.EOF
            pgbBar = 90 + (rsTmp.AbsolutePosition / rsTmp.RecordCount * 8)
            Call UpdateCountCust(rsTmp)
            rsTmp.MoveNext
            DoEvents
        Wend
      '問題集 2801 將得到的客戶數加到SO5260A3
       If strBehindWhere <> "" Then
          If cnn.State = adStateClosed Then Set cnn = GetTmpMdbCn
          If rsTemp.State = 1 Then rsTemp.Close
          strTempMDB = "SELECT " & GroupField & ",SO003.Period,SO003.Amount,Count(Distinct SO001.CustId) as CountCust,SO003.ClctDate " & _
                       " From " & strTables & " Where " & strWhere & strBehindWhere & " Group By " & strGroupBy & ",SO003.Period,SO003.Amount,SO003.ClctDate"
          Set rsTemp = gcnGi.Execute(strTempMDB)
          Do While Not rsTemp.EOF
             Call UpdateCountCust(rsTemp, True)
             rsTemp.MoveNext
             DoEvents
          Loop
       End If
        CloseRecordset rsTmp
        CloseRecordset rsTemp
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subInsertSO5260A1")
End Sub

Private Sub subInsertSO5260A2()
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strMDB As String
      
      cnn.Execute "DELETE * FROM SO5260A2"
        
        If rsTmp.State = 1 Then rsTmp.Close
        strMDB = "SELECT Distinct Count(SO001.Custid) as intCust,SO001.ClassCode1,SO001.ClassName1 " & _
                 " From " & strTables & " Where " & strWhere & strChoose & " Group by SO001.ClassCode1,SO001.ClassName1"
        Set rsTmp = gcnGi.Execute(strMDB)
        SendSQL strMDB, True
        
        cnn.BeginTrans
        While Not rsTmp.EOF
              pgbBar = 98 + (rsTmp.AbsolutePosition / rsTmp.RecordCount * 2)
              cnn.Execute "Insert into SO5260A2 (ClassCode,ClassName,intCust) Values (" & _
                          GetNullString(rsTmp("ClassCode1")) & "," & _
                          GetNullString(rsTmp("ClassName1")) & "," & _
                          GetNullString(rsTmp("intCust")) & ")"
              rsTmp.MoveNext
        Wend
        cnn.CommitTrans
        CloseRecordset rsTmp
        If strBehindWhere <> "" Then Call CalcuSO5260A2
    Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "subInsertSO5260A2")
End Sub

Private Function ChkDataExist(ByVal rsCHK As Object) As Boolean
  On Error GoTo ChkErr
    Dim rsData As New ADODB.Recordset
    ChkDataExist = True
    
    Set rsData = cnn.Execute("Select Count(*) as intCount From SO5260A1 " & _
                           "Where Groupcode='" & rsCHK("GroupCode") & "' And GroupName='" & rsCHK("GroupName") & _
                           "' And Period=" & GetNullString(rsCHK("Period"), giLongV) & " And Amount=" & GetNullString(rsCHK("Amount"), giLongV))
    If rsData("intCount") = 0 Then ChkDataExist = False
    CloseRecordset rsData
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "ChkDataExist")
End Function

Private Sub InsertData(ByVal rsTmp As Object)
  On Error GoTo ChkErr
    Dim strField As String
    Dim strFieldValues As String
    Call GetFieldName(rsTmp, True, strField, strFieldValues)
    
    cnn.BeginTrans
    cnn.Execute "Insert into SO5260A1(GroupCode,GroupName,Period,Amount" & strField & ",BehindAmound,CountCust ) Values (" & _
                GetNullString(rsTmp("GroupCode")) & "," & _
                GetNullString(rsTmp("GroupName")) & "," & _
                Val(rsTmp("Period")) & "," & _
                GetNullString(rsTmp("Amount"), giLongV) & _
                strFieldValues & ",0,0)"
   '將SO5260A1月份值為NULL UPDATE成0，以免做四則運算時會為成NULL
'    cnn.Execute "Update SO5260A1 set Month01=(IIF(Month01>0,Month01,0))," & _
'                "Month02=(IIF(Month02>0,Month02,0))," & _
'                "Month03=(IIF(Month03>0,Month03,0))," & _
'                "Month04=(IIF(Month04>0,Month04,0))," & _
'                "Month05=(IIF(Month05>0,Month05,0))," & _
'                "Month06=(IIF(Month06>0,Month06,0))," & _
'                "Month07=(IIF(Month07>0,Month07,0))," & _
'                "Month08=(IIF(Month08>0,Month08,0))," & _
'                "Month09=(IIF(Month09>0,Month09,0))," & _
'                "Month10=(IIF(Month10>0,Month10,0))," & _
'                "Month11=(IIF(Month11>0,Month11,0))," & _
'                "Month12=(IIF(Month12>0,Month12,0))"
'    cnn.Execute "Update SO5260A1 set Month01=(IIF(Month01>0,Month01,0))," & _
'                "Month02=(IIF(Month02>0,Month02,0)),Month03=(IIF(Month03>0,Month03,0)),Month04=(IIF(Month04>0,Month04,0))," & _
'                "Month05=(IIF(Month05>0,Month05,0)),Month06=(IIF(Month06>0,Month06,0)),Month07=(IIF(Month07>0,Month07,0))," & _
'                "Month08=(IIF(Month08>0,Month08,0)),Month09=(IIF(Month09>0,Month09,0)),Month10=(IIF(Month10>0,Month10,0))," & _
'                "Month11=(IIF(Month11>0,Month11,0)),Month12=(IIF(Month12>0,Month12,0))," & _
'                "CountCust01=(IIF(CountCust01>0,CountCust01,0)),CountCust02=(IIF(CountCust02>0,CountCust02,0)),CountCust03=(IIF(CountCust03>0,CountCust03,0))," & _
'                "CountCust04=(IIF(CountCust04>0,CountCust04,0)),CountCust05=(IIF(CountCust05>0,CountCust05,0)),CountCust06=(IIF(CountCust06>0,CountCust06,0))," & _
'                "CountCust07=(IIF(CountCust07>0,CountCust07,0)),CountCust08=(IIF(CountCust08>0,CountCust08,0)),CountCust09=(IIF(CountCust09>0,CountCust09,0))," & _
'                "CountCust10=(IIF(CountCust10>0,CountCust10,0)),CountCust11=(IIF(CountCust11>0,CountCust11,0)),CountCust12=(IIF(CountCust12>0,CountCust12,0))"
    cnn.CommitTrans
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "InsertData")
End Sub

Private Sub UpdateData(ByVal rsUpdate As Object)
  On Error GoTo ChkErr
    Dim strField As String
    Dim strFieldValues As String
    Call GetFieldName(rsUpdate, False, strField, strFieldValues)
    
    cnn.BeginTrans
    cnn.Execute "Update SO5260A1 set " & Mid(strField, 2) & _
                " Where Groupcode=" & GetNullString(rsUpdate("GroupCode")) & " And GroupName=" & GetNullString(rsUpdate("GroupName")) & _
                " And Period=" & rsUpdate("Period") & " And Amount=" & rsUpdate("Amount")
    cnn.CommitTrans
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "UpdateData")
End Sub

Private Sub UpdateCountCust(ByVal rsUpdate As Object, Optional ByVal blnBehind As Boolean = False)
  On Error GoTo ChkErr
   Dim strDatePeriod As String
   Dim strField As String
   Dim strgroupcode As String
   If blnBehind Then
      strDatePeriod = Format(DateAdd("m", rsTemp("Period"), rsTemp("ClctDate")), "yyyymm")
      Do
         If Val(strDatePeriod) >= Val(Format(gdaClctDate1.GetValue(True), "yyyymm")) Then
            strField = GetUpdFieldName(strDatePeriod, rsTemp("CountCust"), rsTemp("Period"), True)
            Exit Do
         Else
            strDatePeriod = Format(DateAdd("m", rsTemp("period"), Left(strDatePeriod, 4) & "/" & Right(strDatePeriod, 2)), "yyyymm")
         End If
      Loop
      If optClctDate Then
         strgroupcode = strDatePeriod
      Else
         strgroupcode = rsUpdate("GroupCode")
      End If
      cnn.BeginTrans
      '更新SO5260A3的客戶數與總數
      cnn.Execute "Update SO5260A3 set CountCust=CountCust +" & rsUpdate("CountCust") & _
                   " Where Groupcode=" & GetNullString(rsUpdate("GroupCode")) & " And GroupName=" & GetNullString(rsUpdate("GroupName")) & _
                   " And Period=" & rsUpdate("Period") & " And Amount=" & rsUpdate("Amount")
      '以SO5260A3 CountCust的值，累加進去SO5260A1
'      cnn.Execute "Update SO5260A1 set CountCust=CountCust+" & rsUpdate("CountCust") & strField & _
'                  " Where GroupCode='" & strgroupcode & "' And Period=" & rsUpdate("period") & _
'                  " And Amount=" & rsUpdate("Amount")
        cnn.Execute "Update SO5260A1 set CountCust=CountCust+" & rsUpdate("CountCust") & _
             " Where GroupCode='" & strgroupcode & "' And Period=" & rsUpdate("period") & _
             " And Amount=" & rsUpdate("Amount")
        

'      cnn.Execute "Update SO5260A1 set CountCust=CountCust+" & rsUpdate("CountCust") & _
'                  ",CountCust01=(IIF(CountCust01>0,CountCust01+" & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust02=(IIF(CountCust02>0,CountCust02+" & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust03=(IIF(CountCust03>0,CountCust03+" & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust04=(IIF(CountCust04>0,CountCust04+" & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust05=(IIF(CountCust05>0,CountCust05+" & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust06=(IIF(CountCust06>0,CountCust06+" & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust07=(IIF(CountCust07>0,CountCust07+" & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust08=(IIF(CountCust08>0,CountCust08+" & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust09=(IIF(CountCust09>0,CountCust09+" & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust10=(IIF(CountCust10>0,CountCust10+" & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust11=(IIF(CountCust11>0,CountCust11+" & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust12=(IIF(CountCust12>0,CountCust12+" & rsUpdate("CountCust") & ",0))" & _
'                  " Where Groupcode=" & GetNullString(rsUpdate("GroupCode")) & " And GroupName=" & GetNullString(rsUpdate("GroupName")) & _
'                  " And Period=" & rsUpdate("Period") & " And Amount=" & rsUpdate("Amount")
'
      cnn.CommitTrans
   Else
      '挑選要Update的欄位
      cnn.BeginTrans
'      cnn.Execute "Update SO5260A1 set CountCust=CountCust +" & rsUpdate("CountCust") & _
'                  ",CountCust01=(IIF(Month01>0," & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust02=(IIF(Month02>0," & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust03=(IIF(Month03>0," & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust04=(IIF(Month04>0," & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust05=(IIF(Month05>0," & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust06=(IIF(Month06>0," & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust07=(IIF(Month07>0," & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust08=(IIF(Month08>0," & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust09=(IIF(Month09>0," & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust10=(IIF(Month10>0," & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust11=(IIF(Month11>0," & rsUpdate("CountCust") & ",0))" & _
'                  ",CountCust12=(IIF(Month12>0," & rsUpdate("CountCust") & ",0))" & _
'                  " Where Groupcode=" & GetNullString(rsUpdate("GroupCode")) & " And GroupName=" & GetNullString(rsUpdate("GroupName")) & _
'                  " And Period=" & rsUpdate("Period") & " And Amount=" & rsUpdate("Amount")
        cnn.Execute "Update SO5260A1 set CountCust=CountCust +" & rsUpdate("CountCust") & _
            ",CountCust01=(IIF(Month01>0,Month01/Amount,0))" & _
            ",CountCust02=(IIF(Month02>0,Month02/Amount,0))" & _
            ",CountCust03=(IIF(Month03>0,Month03/Amount,0))" & _
            ",CountCust04=(IIF(Month04>0,Month04/Amount,0))" & _
            ",CountCust05=(IIF(Month05>0,Month05/Amount,0))" & _
            ",CountCust06=(IIF(Month06>0,Month06/Amount,0))" & _
            ",CountCust07=(IIF(Month07>0,Month07/Amount,0))" & _
            ",CountCust08=(IIF(Month08>0,Month08/Amount,0))" & _
            ",CountCust09=(IIF(Month09>0,Month09/Amount,0))" & _
            ",CountCust10=(IIF(Month10>0,Month10/Amount,0))" & _
            ",CountCust11=(IIF(Month11>0,Month11/Amount,0))" & _
            ",CountCust12=(IIF(Month12>0,Month12/Amount,0))" & _
            " Where Groupcode=" & GetNullString(rsUpdate("GroupCode")) & " And GroupName=" & GetNullString(rsUpdate("GroupName")) & _
            " And Period=" & rsUpdate("Period") & " And Amount=" & rsUpdate("Amount")

      cnn.CommitTrans
   End If
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "UpdateData")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5260A
End Sub

Private Sub gdaClctDate1_GotFocus()
  On Error Resume Next
  If gdaClctDate1.GetValue = "" Then gdaClctDate1.SetValue (Format(DateAdd("M", 1, RightDate), "YYYYMM") & "01")
End Sub

Private Sub gdaClctDate2_GotFocus()
  On Error Resume Next
  If gdaClctDate1.GetValue <> "" Then
      gdaClctDate2.SetValue GetLastDate(DateAdd("M", 11, gdaClctDate1.GetValue(True)))
  Else
      gdaClctDate2.SetValue ""
  End If
End Sub

Private Sub gdaClctDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaClctDate1, gdaClctDate2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
    
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
    
    Call GiMultiFilter(gimCitemCode, gilServiceType.GetCodeNo)
    gimCitemCode.Filter = gimCitemCode.Filter & IIf(gimCitemCode.Filter = "", " WHERE ", " AND ") & " PeriodFlag = 1"
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilServiceType_Change")
End Sub

Private Function GetMonthName(ByVal strMonthName As String) As String
  On Error GoTo ChkErr
    Dim Date1 As Date
    Dim Date2 As Date
    Dim i As Integer
    Date1 = CDate(Format(gdaClctDate1.GetValue(True), "YYYY/MM") & "/01")
    Date2 = CDate(strMonthName & "/01")
    For i = 0 To 11
        If DateAdd("M", i, Date1) = Date2 Then Exit For
    Next
    GetMonthName = "Month" & Format(i + 1, "00")
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "GetMonthName")
End Function

Private Function GetFieldName(ByVal rsField As Object, ByVal blnInsert As Boolean, strField As String, strFieldValues As String, Optional ByVal blnBehind As Boolean = False) As String
  On Error GoTo ChkErr
    Dim i As Integer
    Dim days As String
      i = 1
      If blnBehind Then
         days = Format(DateAdd("m", Val(rsTemp("Period")), Left(rsTemp("GroupCode"), 2) & "/" & Right(rsTemp("GroupCode"), 2)), "yyyy/mm")
      Else
         days = Format(rsField("ClctDate"), "YYYY/MM") '下次收費[年/月]
      End If
      Do Until days > Format(gdaClctDate2.GetValue(True), "YYYY/MM") Or rsField("period") = 0
          If blnInsert = True Then
              strField = strField & "," & GetMonthName(days)
              strFieldValues = strFieldValues & "," & rsField("SumAmount")
          Else
              strField = strField & "," & GetMonthName(days) & "=" & GetMonthName(days) & "+" & rsField("SumAmounT")
          End If
          '下次收費年月依期數累加以不超出輸入預估日期範圍
          If blnBehind Then
            days = Format(DateAdd("M", rsField("Period") * i, days), "YYYY/MM")
          Else
          
            days = Format(DateAdd("M", rsField("PERIOD") * i, rsField("CLCTDATE")), "YYYY/MM")
          End If
          i = i + 1
      Loop
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "GetFieldName")
End Function

Private Sub gymBehind1_Validate(Cancel As Boolean)
 On Error GoTo ChkErr
    Dim dtaBehind As Date
    If IsDate(gymBehind1.GetOriginalValue) Then
        dtaBehind = Format(gymBehind1.GetOriginalValue & "/01", "YYYY/MM/DD")
        If dtaBehind >= gdaClctDate1.GetOriginalValue Then
            MsgBox "回溯年月不可大於預估起始日", , "警告訊息！！"
            gymBehind1.Text = ""
            Cancel = True
        End If
    End If
    Exit Sub
ChkErr:
 Call ErrSub(Me.Name, "gymBehind1_Validate")
End Sub
Private Sub InsertTemp()
  On Error GoTo ChkErr
    Dim strField As String
    Dim strFieldValues As String
    Dim strTemp As String
    Dim strDayPeriod As String
'    Call CreateTemp
'問題集 2801 將回塑年的資料加到SO5260A3 MDB
        strTemp = "SELECT " & GroupField & ",SO003.ClctDate,SO003.Period,SO003.Amount,Sum(SO003.Amount) as SumAmount" & _
                 " From " & strTables & " Where " & strWhere & strBehindWhere & " Group By " & strGroupBy & ",SO003.ClctDate,SO003.Period,SO003.Amount"
        Set rsTemp = gcnGi.Execute(strTemp)
        cnn.BeginTrans
        Do While Not rsTemp.EOF
            cnn.Execute "Insert into SO5260A3 (GroupCode,GroupName,ClctDate,Period,Amount,CountCust,SumAmount) Values (" & _
                        GetNullString(rsTemp("GroupCode")) & "," & GetNullString(rsTemp("GroupName")) & "," & _
                        GetNullString(rsTemp("ClctDate")) & "," & GetNullString(rsTemp("Period")) & "," & _
                        rsTemp("Amount") & ",0," & rsTemp("SumAmount") & ")"
'            '假如是下次收費年月，要特別處理
'            If optClctDate Then
'               strDayPeriod = Format(DateAdd("m", rsTemp("Period"), Left(rsTemp("GroupCode"), 4) & "/" & Right(rsTemp("GroupCode"), 2)), "YYYYMM")
'               Do
'                  If Val(strDayPeriod) >= Val(Format(gdaClctDate1.GetValue(True), "yyyymm")) Then
'                     cnn.Execute "Update SO5260A3 set GroupCode='" & strDayPeriod & "'" & _
'                                 " Where GroupCode='" & rsTemp("GroupCode") & "' And Period=" & _
'                                 rsTemp("Period") & " And Amount=" & rsTemp("Amount") & " And " & _
'                                 "ClctDate='" & rsTemp("ClctDate") & "'"
'                     Exit Do
'                  Else
'                     strDayPeriod = Format(DateAdd("m", rsTemp("period"), Left(strDayPeriod, 4) & "/" & Right(strDayPeriod, 2)), "yyyymm")
'                  End If
'               Loop
'            End If
            rsTemp.MoveNext
            DoEvents
        Loop
        cnn.CommitTrans
'問題集 2801 將回塑資料Appen到 SO5260A1
      Call UpdateTemp

        Set rsTemp = Nothing
    Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "InsertTemp")
End Sub
Private Sub UpdateTemp()
  On Error GoTo ChkErr
      Dim strUpdSQL, strUpdField As String
      Dim strgroupcode As String
      Dim rsMDB As New ADODB.Recordset
      Dim strField As String
      Dim strDayPeriod As String
      '將得到回塑資料的SumAmountk加到SO5260A1
      rsTemp.MoveFirst
      cnn.BeginTrans
      Do While Not rsTemp.EOF
         If optClctDate Then
            strDayPeriod = Format(DateAdd("m", rsTemp("Period"), Left(rsTemp("GroupCode"), 4) & "/" & Right(rsTemp("GroupCode"), 2)), "YYYYMM")
            Do
               If Val(strDayPeriod) >= Val(Format(gdaClctDate1.GetValue(True), "yyyymm")) Then
                    strField = GetUpdFieldName(strDayPeriod, rsTemp("SumAmount"), rsTemp("Period"))
                  Exit Do
               Else
                  strDayPeriod = Format(DateAdd("m", rsTemp("period"), Left(strDayPeriod, 4) & "/" & Right(strDayPeriod, 2)), "yyyymm")
               End If
            Loop
            'strGroupCode = Format(DateAdd("m", Val(rsTemp("Period")), Left(rsTemp("GroupCode"), 4) & "/" & Right(rsTemp("GroupCode"), 2)), "yyyymm")
            strgroupcode = strDayPeriod
         Else
            strDayPeriod = Format(DateAdd("m", rsTemp("Period"), rsTemp("ClctDate")), "yyyymm")
            Do
               If Val(strDayPeriod) >= Val(Format(gdaClctDate1.GetValue(True), "yyyymm")) Then
                    strField = GetUpdFieldName(strDayPeriod, rsTemp("SumAmount"), rsTemp("Period"))
                  Exit Do
               Else
                  strDayPeriod = Format(DateAdd("m", rsTemp("period"), Left(strDayPeriod, 4) & "/" & Right(strDayPeriod, 2)), "yyyymm")
               End If
            Loop
            strgroupcode = rsTemp("GroupCode")
         End If
        '更新的月份資料
       If Trim(strField) <> "" Then
            strUpdSQL = "Update SO5260A1 SET BehindAmound=BehindAmound+" & rsTemp("SumAmount") & strField & " Where GroupCode='" & _
                         strgroupcode & "' And Period=" & rsTemp("Period") & " And Amount=" & rsTemp("Amount")
            Set rsMDB = cnn.Execute(strUpdSQL)
       End If
         rsTemp.MoveNext
      Loop
      cnn.CommitTrans
      Set rsTemp = Nothing
  Exit Sub
ChkErr:
   cnn.RollbackTrans
 Call ErrSub(Me.Name, "UpdateTemp")
End Sub
Sub CreateTable()
  On Error Resume Next
    cnn.Execute "Drop Table SO5260A1"
    cnn.Execute "Drop Table SO5260A2"
    cnn.Execute "Drop Table SO5260A3"
    cnn.Execute "Drop Table SO5260A4"
    'cnn.Execute "Drop Table So5270A2"
  On Error GoTo ChkErr
  'SO5260A1 問題集2801 增加各個月分的客戶數
    cnn.Execute "Create Table SO5260A1 (GroupCode text(50),GroupName Text(50)," & _
                "Period Long default 0,Amount Long default 0,CountCust Long default 0,Month01 Long default 0," & _
                "Month02 Long default 0,Month03 Long default 0,Month04 Long default 0,Month05 Long default 0," & _
                "Month06 Long default 0,Month07 Long default 0,Month08 Long default 0,Month09 Long default 0," & _
                "Month10 Long default 0,Month11 Long default 0,Month12 Long default 0," & _
                "BehindAmound Long default 0,CountCust01 Long default 0,CountCust02 Long default 0,CountCust03 Long default 0," & _
                "CountCust04 Long default 0,CountCust05 Long default 0,CountCust06 Long default 0,CountCust07 Long default 0, " & _
                "CountCust08 Long default 0,CountCust09 Long default 0,CountCust10 Long default 0,CountCust11 Long default 0, " & _
                "CountCust12 Long default 0)"
  'SO5260A2
    cnn.Execute "Create Table SO5260A2 (ClassCode long,ClassName Text(50)," & _
                "intCust Long default 0)"
  'SO5260A3 回塑日期
    cnn.Execute "Create Table SO5260A3 (GroupCode text(50),GroupName Text(50)," & _
                "ClctDate text(20),Period Long default 0,Amount Long default 0,CountCust Long default 0,SumAmount long default 0)"
  'SO5260A4 各期統計
    cnn.Execute "Create Table SO5260A4 (" & _
                "Period Long default 0,Amount Long default 0,CountCust Long default 0,Month01 Long default 0," & _
                "Month02 Long default 0,Month03 Long default 0,Month04 Long default 0,Month05 Long default 0," & _
                "Month06 Long default 0,Month07 Long default 0,Month08 Long default 0,Month09 Long default 0," & _
                "Month10 Long default 0,Month11 Long default 0,Month12 Long default 0," & _
                "BehindAmound Long default 0,CountCust01 Long default 0,CountCust02 Long default 0,CountCust03 Long default 0," & _
                "CountCust04 Long default 0,CountCust05 Long default 0,CountCust06 Long default 0,CountCust07 Long default 0, " & _
                "CountCust08 Long default 0,CountCust09 Long default 0,CountCust10 Long default 0,CountCust11 Long default 0, " & _
                "CountCust12 Long default 0)"
   Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "CreateTable")
End Sub
Private Sub toExcel(Optional ByVal strsql As String = "", Optional ByVal strSubQry1 As String = "")
  On Error GoTo ChkErr
    Dim rsExcel As New ADODB.Recordset
    Dim rsSubExcel1 As New ADODB.Recordset
    Dim i As Long
    Dim strMonth(11) As String
    Dim strExcelSQL As String
    Dim strSubExcelSQL As String
    Dim strShowField As String
    Dim strChangeYear As String
    '取得匯出Excel的表頭年月
    For i = 0 To 11
      strChangeYear = Format(DateAdd("m", i, gdaClctDate1.GetValue(True)), "YYYYMM")
      strMonth(i) = Left(Trim(str(Val(strChangeYear) - 191100)), 2) & "/" & Right(Trim(str(Val(strChangeYear)) - 191100), 2)
      strShowField = strShowField & ",客戶數,   " & strMonth(i) & "  "
    Next i
    strExcelSQL = "Select Period as 期數,AMOUNT AS 金額" & _
                  ",CountCust01 as CountCust01" & _
                  ",Month01 as Month01" & _
                  ",CountCust02 as CountCust02" & _
                  ",Month02 as Month02" & _
                  ",CountCust03 as CountCust03" & _
                  ",Month03 as Month03" & _
                  ",CountCust04 as CountCust04" & _
                  ",Month04 as Month04" & _
                  ",CountCust05 as CountCust05" & _
                  ",Month05 as Month05" & _
                  ",CountCust06 as CountCust06" & _
                  ",Month06 as Month06" & _
                  ",CountCust07 as CountCust07" & _
                  ",Month07 as Month07" & _
                  ",CountCust08 as CountCust08" & _
                  ",Month08 as Month08" & _
                  ",CountCust09 as CountCust09" & _
                  ",Month09 as Month09" & _
                  ",CountCust10 as CountCust10" & _
                  ",Month10 as Month10" & _
                  ",CountCust11 as CountCust11" & _
                  ",Month11 as Month11" & _
                  ",CountCust12 as CountCust12" & _
                  ",Month12 as Month12 From SO5260A4 Order by Period,Amount"
    strSubExcelSQL = "Select Period as 期數, sum((Month01+Month02+Month03+Month04+ " & _
                     "Month05+Month06+Month07+Month08+Month09+Month10+Month11+Month12)) " & _
                     "as 總金額         from SO5260A4 GROUP BY Period"
    strShowField = "期數,   金額   " & strShowField
    GetRS rsExcel, strExcelSQL, cnn, adUseClient, adOpenKeyset
    GetRS rsSubExcel1, strSubExcelSQL, cnn, adUseClient, adOpenKeyset
    Call UseProperty(rsExcel, "預收資料一覽表", "第一頁", strShowField, , rsSubExcel1)
    
    blnExcel = False
    CloseRecordset rsExcel
    CloseRecordset rsSubExcel1
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "toExcel")
End Sub
Sub InsertDataSO5260A4()
  On Error GoTo ChkErr:
  '將SO5260A1 相同期數與相同金額加總，以呈現一列方式
    Dim rsA1 As New ADODB.Recordset
    Dim strA1SQL As String
    strA1SQL = "Insert into SO5260A4 Select Period,Amount,Sum(CountCust) as CountCust," & _
               "Sum(Month01) as Month01,Sum(Month02) as Month02,Sum(Month03) as Month03," & _
               "Sum(Month04) as Month04,Sum(Month05) as Month05,Sum(Month06) as Month06," & _
               "Sum(Month07) as Month07,Sum(Month08) as Month08,Sum(Month09) as Month09," & _
               "Sum(Month10) as Month10,Sum(Month11) as Month11,Sum(Month12) as Month12," & _
               "Sum(CountCust01) as CountCust01,Sum(CountCust02) as CountCust02,Sum(CountCust03) as CountCust03," & _
               "Sum(CountCust04) as CountCust04,Sum(CountCust05) as CountCust05,Sum(CountCust06) as CountCust06," & _
               "Sum(CountCust07) as CountCust07,Sum(CountCust08) as CountCust08,Sum(CountCust09) as CountCust09," & _
               "Sum(CountCust10) as CountCust10,Sum(CountCust11) as CountCust11,Sum(CountCust12) as CountCust12" & _
               " from SO5260A1 Group By Period,Amount"
    cnn.BeginTrans
        '將得來群組化Null值的資料填0
        Set rsA1 = cnn.Execute(strA1SQL)
'        strA1SQL = "Update SO5260A4 SET COUNTCUST01=(IIF(COUNTCUST01>0,COUNTCUST01,0))," & _
'                   "COUNTCUST02=(IIF(COUNTCUST02>0,COUNTCUST02,0))," & _
'                   "COUNTCUST03=(IIF(COUNTCUST03>0,COUNTCUST03,0))," & _
'                   "COUNTCUST04=(IIF(COUNTCUST04>0,COUNTCUST04,0))," & _
'                   "COUNTCUST05=(IIF(COUNTCUST05>0,COUNTCUST05,0))," & _
'                   "COUNTCUST06=(IIF(COUNTCUST06>0,COUNTCUST06,0))," & _
'                   "COUNTCUST07=(IIF(COUNTCUST07>0,COUNTCUST07,0))," & _
'                   "COUNTCUST08=(IIF(COUNTCUST08>0,COUNTCUST08,0))," & _
'                   "COUNTCUST09=(IIF(COUNTCUST09>0,COUNTCUST09,0))," & _
'                   "COUNTCUST10=(IIF(COUNTCUST10>0,COUNTCUST10,0))," & _
'                   "COUNTCUST11=(IIF(COUNTCUST11>0,COUNTCUST11,0))," & _
'                   "COUNTCUST12=(IIF(COUNTCUST12>0,COUNTCUST12,0))," & _
'                   "Month01=(IIF(Month01>0,Month01,0)),Month02=(IIF(Month02>0,Month02,0)),Month03=(IIF(Month03>0,Month03,0))," & _
'                   "Month04=(IIF(Month04>0,Month04,0)),Month05=(IIF(Month05>0,Month05,0)),Month06=(IIF(Month06>0,Month06,0))," & _
'                   "Month07=(IIF(Month07>0,Month07,0)),Month08=(IIF(Month08>0,Month08,0)),Month09=(IIF(Month09>0,Month09,0))," & _
'                   "Month10=(IIF(Month10>0,Month10,0)),Month11=(IIF(Month11>0,Month11,0)),Month12=(IIF(Month12>0,Month12,0))"
'        Set rsA1 = cnn.Execute(strA1SQL)
    cnn.CommitTrans
    Set rsA1 = Nothing
    Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "InsertDataSO5260A4")
End Sub
Private Function GetUpdFieldName(ByVal strDay As String, ByVal strUpdVaule As String, ByVal lngPeriod As Long, Optional ByVal blnCust As Boolean = False) As String
  On Error GoTo ChkErr
    Dim i As Integer
    Dim days As String
    Dim strField As String
    days = Left(strDay, 4) & "/" & Right(strDay, 2)
      Do Until days > Format(gdaClctDate2.GetValue(True), "YYYY/MM")
          i = i + 1
          If blnCust Then
             strField = strField & "," & GetCustName(days) & "=" & GetCustName(days) & "+" & strUpdVaule
          Else
             strField = strField & "," & GetMonthName(days) & "=" & GetMonthName(days) & "+" & strUpdVaule
          End If
          '下次收費年月依期數累加以不超出輸入預估日期範圍
          days = Format(DateAdd("M", lngPeriod, days), "YYYY/MM")
      Loop
      GetUpdFieldName = strField
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "GetUpdFieldName")
End Function

Private Function GetCustName(ByVal strCustName As String) As String
  On Error GoTo ChkErr
    Dim Date1 As Date
    Dim Date2 As Date
    Dim i As Integer
    Date1 = CDate(Format(gdaClctDate1.GetValue(True), "YYYY/MM") & "/01")
    Date2 = CDate(strCustName & "/01")
    For i = 0 To 11
        If DateAdd("M", i, Date1) = Date2 Then Exit For
    Next
    GetCustName = "CountCust" & Format(i + 1, "00")
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "GetCustName")
End Function
Private Sub CalcuSO5260A2()
  On Error GoTo ChkErr:
    Dim rsSO5260A1 As New ADODB.Recordset
    Dim rsSO5260A2 As New ADODB.Recordset
    Dim rsCacuA3 As New ADODB.Recordset
    Dim strMDB As String
    Dim strDatePeriod As String
    strMDB = "SELECT " & GroupField & ",SO003.ClctDate,SO003.Period,SO003.Amount,Sum(SO003.Amount) as SumAmount, " & _
             "SO001.ClassCode1,SO001.ClassName1,count(distinct SO001.Custid) as intCount " & _
            " From " & strTables & " Where " & strWhere & strBehindWhere & " Group By " & strGroupBy & ",SO003.ClctDate,SO003.Period,SO003.Amount," & _
            "SO001.ClassCode1,SO001.ClassName1"
    Set rsCacuA3 = gcnGi.Execute(strMDB)
    '將回塑的資料全部挑選出來，一一比對是否要加到SO5260A2
        Do While Not rsCacuA3.EOF
        '假如是選擇下收日期，要特別處理
            If optClctDate Then
                strDatePeriod = Format(DateAdd("m", rsCacuA3("Period"), Left(rsCacuA3("GroupCode"), 4) & "/" & Right(rsCacuA3("GroupCode"), 2)), "yyyymm")
                Do
                    If Val(strDatePeriod) >= Val(Left(gdaClctDate1.GetValue, 6)) Then
                        Exit Do
                    Else
                        strDatePeriod = Format(DateAdd("m", rsCacuA3("Period"), Left(strDatePeriod, 4) & "/" & Right(strDatePeriod, 2)), "YYYYMM")
                    End If
                Loop
                strDatePeriod = Left(strDatePeriod, 6)
            Else
                strDatePeriod = rsCacuA3("GroupCode")
            End If
            '查詢回塑的資料是否落在SO5260A1，如果是將SO5260A2資料Append或Update
            Set rsSO5260A1 = cnn.Execute("Select count(*) from SO5260A1 Where GroupCode='" & strDatePeriod & "' " & _
                                       " And Period=" & rsCacuA3("Period") & " And Amount=" & rsCacuA3("Amount"))
                                      
            If rsSO5260A1(0) >= 1 Then
                Set rsSO5260A2 = cnn.Execute("Select Count(*) from SO5260A2 Where ClassCode=" & rsCacuA3("ClassCode1"))
                If rsSO5260A2(0) >= 1 Then
                    cnn.BeginTrans
                    Set rsSO5260A2 = cnn.Execute("Update SO5260A2 Set intCust=intCust+" & rsCacuA3("intCount") & " Where ClassCode=" & rsCacuA3("ClassCode1"))
                    cnn.CommitTrans
                Else
                    cnn.BeginTrans
                    Set rsSO5260A2 = cnn.Execute("Insert into SO5260A2 (ClassCode,ClassName,intCust) Values (" & _
                                                 rsCacuA3("ClassCode1") & ",'" & rsCacuA3("ClassName1") & "'," & rsCacuA3("intCount") & ")")
                    cnn.CommitTrans
                End If
            End If
            rsCacuA3.MoveNext
            DoEvents
        Loop
    Set rsSO5260A1 = Nothing
    Set rsSO5260A2 = Nothing
    Set rsCacuA3 = Nothing
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "CalcuSO5260A2")
End Sub

Private Sub optAreaCode_Click()
    cmdToExcel.Enabled = False
End Sub

Private Sub optClassCode_Click()
    cmdToExcel.Enabled = False
End Sub

Private Sub optClctAreaCode_Click()
    cmdToExcel.Enabled = False
End Sub

Private Sub optClctDate_Click()
    cmdToExcel.Enabled = False
End Sub

Private Sub optClctEn_Click()
    cmdToExcel.Enabled = False
End Sub

Private Sub optPeriod_Click()
    cmdToExcel.Enabled = True
End Sub

Private Sub optServCode_Click()
    cmdToExcel.Enabled = False
End Sub

Private Sub optStrtCode_Click()
    cmdToExcel.Enabled = False
End Sub
