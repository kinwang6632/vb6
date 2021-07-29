VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO52J0A 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '單線固定
   Caption         =   "IVR回單稽核報表 [SO52J0A]"
   ClientHeight    =   6525
   ClientLeft      =   1950
   ClientTop       =   1530
   ClientWidth     =   8580
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO52J0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form17"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CheckBox chkDispBillWip 
      Caption         =   "不顯示無收費之工單"
      Height          =   315
      Left            =   4770
      TabIndex        =   50
      Top             =   960
      Width           =   2085
   End
   Begin VB.Frame fraSimple 
      Caption         =   "工單回單簡表"
      Height          =   645
      Left            =   1620
      TabIndex        =   46
      Top             =   5220
      Width           =   4275
      Begin VB.OptionButton optWorker 
         Caption         =   "工程人員"
         Height          =   285
         Left            =   180
         TabIndex        =   43
         Top             =   240
         Width           =   1275
      End
      Begin VB.OptionButton optServ 
         Caption         =   "服務區"
         Height          =   315
         Left            =   1620
         TabIndex        =   44
         Top             =   210
         Width           =   1035
      End
      Begin VB.OptionButton optTotal 
         Caption         =   "派工總量"
         Height          =   315
         Left            =   2790
         TabIndex        =   45
         Top             =   210
         Width           =   1275
      End
   End
   Begin VB.Frame fraPrint 
      Height          =   585
      Left            =   150
      TabIndex        =   41
      Top             =   5250
      Width           =   1395
      Begin VB.OptionButton optDetail 
         Caption         =   "明細表"
         Height          =   405
         Left            =   180
         TabIndex        =   42
         Top             =   150
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.OptionButton optSimple 
      Caption         =   "工單回收簡表"
      Height          =   405
      Left            =   7440
      TabIndex        =   40
      Top             =   5910
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CheckBox chkHistory 
      Caption         =   "統計歷史資料"
      Height          =   345
      Left            =   6840
      TabIndex        =   7
      Top             =   570
      Width           =   1575
   End
   Begin VB.CheckBox chkSO033 
      Caption         =   "是否顯示收費資料"
      Height          =   315
      Left            =   4800
      TabIndex        =   6
      Top             =   600
      Value           =   1  '核取
      Width           =   1875
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "匯成EXCEL"
      Height          =   525
      Left            =   3240
      TabIndex        =   30
      Top             =   5940
      Width           =   1155
   End
   Begin VB.Frame fraStatus 
      Caption         =   "派工狀態"
      Height          =   675
      Left            =   2970
      TabIndex        =   38
      Top             =   1350
      Width           =   3975
      Begin VB.OptionButton optStausAll 
         Caption         =   "全部"
         Height          =   315
         Left            =   3030
         TabIndex        =   15
         Top             =   270
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton optNotAns 
         Caption         =   "未回報"
         Height          =   315
         Left            =   3840
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.OptionButton optNotFin 
         Caption         =   "未完成"
         Height          =   285
         Left            =   1980
         TabIndex        =   14
         Top             =   270
         Width           =   975
      End
      Begin VB.OptionButton optRet 
         Caption         =   "退單"
         Height          =   315
         Left            =   1170
         TabIndex        =   13
         Top             =   270
         Width           =   765
      End
      Begin VB.OptionButton optFin 
         Caption         =   "已完成"
         Height          =   315
         Left            =   180
         TabIndex        =   12
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.Frame fraIVR 
      Caption         =   "回單來源"
      Height          =   675
      Left            =   120
      TabIndex        =   37
      Top             =   1350
      Width           =   2655
      Begin VB.OptionButton optIVRAll 
         Caption         =   "全部"
         Height          =   315
         Left            =   1740
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optNoIVR 
         Caption         =   "非IVR"
         Height          =   315
         Left            =   810
         TabIndex        =   10
         Top             =   240
         Width           =   885
      End
      Begin VB.OptionButton optIVR 
         Caption         =   "IVR"
         Height          =   315
         Left            =   30
         TabIndex        =   9
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   525
      Left            =   1590
      TabIndex        =   29
      Top             =   5940
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   525
      Left            =   4680
      TabIndex        =   31
      Top             =   5940
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      Height          =   525
      Left            =   90
      TabIndex        =   28
      Top             =   5940
      Width           =   1245
   End
   Begin VB.Frame fraPage 
      BackColor       =   &H00E0E0E0&
      Caption         =   "分頁方式"
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   7200
      TabIndex        =   34
      Top             =   5160
      Visible         =   0   'False
      Width           =   1035
      Begin VB.OptionButton optClctEn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "工程人員"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3390
         TabIndex        =   27
         Top             =   330
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optAreaCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "行政區"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1770
         TabIndex        =   26
         Top             =   330
         Width           =   885
      End
      Begin VB.OptionButton optServCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "服務區"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   360
         Width           =   885
      End
   End
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   345
      Left            =   90
      TabIndex        =   21
      Top             =   3960
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   609
      ButtonCaption   =   "行     政     區"
      FontSize        =   9
      FontName        =   "新細明體"
   End
   Begin Gi_Date.GiDate gdaRealDate1 
      Height          =   345
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   16711680
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
   Begin Gi_Date.GiDate gdaRealDate2 
      Height          =   345
      Left            =   7110
      TabIndex        =   3
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   16711680
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
      Left            =   930
      TabIndex        =   8
      Top             =   900
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
      Left            =   6240
      TabIndex        =   22
      Top             =   5940
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   609
      ButtonCaption   =   "排  序  方  式"
      DataType        =   2
      ColumnOrder     =   -1  'True
      FontSize        =   9
      FontName        =   "新細明體"
   End
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   345
      Left            =   90
      TabIndex        =   20
      Top             =   3600
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   609
      ButtonCaption   =   "服     務     區"
      FontSize        =   9
      FontName        =   "新細明體"
   End
   Begin Gi_Multi.GiMulti gimBillType 
      Height          =   375
      Left            =   90
      TabIndex        =   23
      Top             =   4350
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   661
      ButtonCaption   =   "單  據  類  別"
      DataType        =   2
      DIY             =   -1  'True
      FontSize        =   9
      FontName        =   "新細明體"
   End
   Begin Gi_Time.GiTime gdtPRTime2 
      Height          =   315
      Left            =   2910
      TabIndex        =   1
      Top             =   150
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      ForeColor       =   16711680
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
   Begin Gi_Multi.GiMulti gimServiceType 
      Height          =   345
      Left            =   90
      TabIndex        =   16
      Top             =   2130
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   609
      ButtonCaption   =   "服  務  類  別"
      DataType        =   2
      FontSize        =   9
      FontName        =   "新細明體"
   End
   Begin CS_Multi.CSmulti gimGroupCode 
      Height          =   375
      Left            =   90
      TabIndex        =   19
      Top             =   3210
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   661
      ButtonCaption   =   "工   程   人   員"
   End
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   345
      Left            =   90
      TabIndex        =   17
      Top             =   2490
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   609
      ButtonCaption   =   "收  費  項  目"
   End
   Begin CS_Multi.CSmulti gimCheckEn 
      Height          =   375
      Left            =   90
      TabIndex        =   24
      Top             =   4740
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   661
      ButtonCaption   =   "稽   核   人   員"
   End
   Begin Gi_Time.GiTime gdtFinTime1 
      Height          =   315
      Left            =   930
      TabIndex        =   4
      Top             =   540
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      ForeColor       =   16711680
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
   Begin Gi_Time.GiTime gdtFinTime2 
      Height          =   315
      Left            =   2910
      TabIndex        =   5
      Top             =   540
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      ForeColor       =   16711680
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
   Begin CS_Multi.CSmulti gimWorkGroup 
      Height          =   375
      Left            =   90
      TabIndex        =   18
      Top             =   2820
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   661
      ButtonCaption   =   "工   程   組   別"
   End
   Begin Gi_Time.GiTime gdtPRTime1 
      Height          =   315
      Left            =   930
      TabIndex        =   0
      Top             =   150
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      ForeColor       =   16711680
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2760
      TabIndex        =   49
      Top             =   630
      Width           =   105
   End
   Begin VB.Label lblPRT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2760
      TabIndex        =   48
      Top             =   240
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "完工日期"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   60
      TabIndex        =   47
      Top             =   630
      Width           =   780
   End
   Begin VB.Label lblResvTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "預約日期"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   60
      TabIndex        =   36
      Top             =   210
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   60
      TabIndex        =   35
      Top             =   990
      Width           =   765
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6930
      TabIndex        =   33
      Top             =   210
      Width           =   105
   End
   Begin VB.Label lblRealDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "實收日期"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4860
      TabIndex        =   32
      Top             =   210
      Width           =   780
   End
End
Attribute VB_Name = "frmSO52J0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'使用Table: SO033 or SO034 or SO035 A,SO001,SO014
'#5039 原本是做在SO5230A的功能,因為要多增加明細表,所以將功能獨立出來做在這裡 By Kin 2009/04/21
Option Explicit
Private strSO00XWhere As String
Private blnUseIVR As Boolean
Private strStatusWhere As String
Private blnExcel As Boolean
Private strCM003Where As String
Private strCD071 As String
Private Sub cmdExcel_Click()
  On Error GoTo ChkErr
    blnExcel = True
    cmdPrint_Click
    blnExcel = False
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExcel_Click")
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
    If optDetail.Value Then
        Call PreviousRpt(GetPrinterName(5), RptName("SO52J0"), "IVR回單稽核報表 [SO52J0A]")
    Else
        Select Case True
            Case optWorker.Value
                Call PreviousRpt(GetPrinterName(5), RptName("SO52J0", "2"), "工單回收簡表 [SO52J0A2]")
            Case optServ.Value
                Call PreviousRpt(GetPrinterName(5), RptName("SO52J0", "3"), "工單回收簡表 [SO52J0A2]")
            Case optTotal.Value
                Call PreviousRpt(GetPrinterName(5), RptName("SO52J0", "4"), "工單回收簡表 [SO52J0A2]")
        End Select
    End If
    
'    If optSimple.Value Then
'        Call PreviousRpt(GetPrinterName(5), RptName("SO52J0", "2"), "工單回收簡表 [SO52J0A2]")
'    Else
'        Call PreviousRpt(GetPrinterName(5), RptName("SO52J0"), "IVR回單稽核報表 [SO52J0A]")
'    End If
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
          blnExcel = False
        Me.Enabled = True
      Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    '#5593 完工時間或預約時間2者其中一個要有值 By Kin 2010/03/19
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 66
    If (gdtPRTime1.GetValue = "") And (gdtPRTime2.GetValue = "") Then
        If (gdtFinTime1.GetValue = "") And (gdtFinTime2.GetValue = "") Then
            gdtPRTime1.SetFocus
            strErrFile = "預約時間或完工時間"
            GoTo 66
        End If
    End If
    If gimGroupCode.GetDispStr = "" Then
        gimGroupCode.SetFocus
        strErrFile = "工程人員"
        GoTo 66
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
    Dim strpagetype As String
    Dim arrOrder() As String
    Dim varOrder As Variant
    Dim intSort As Integer
    Dim strChooseIVRStatus As String
    strSO00XWhere = Empty
    strStatusWhere = Empty
    strChoose = " WHERE A.CUSTID=SO001.CUSTID And A.CompCode = SO001.CompCode AND A.ADDRNO=SO014.ADDRNO And A.CompCode = SO014.CompCode "
   
    strChooseString = ""
    strpagetype = ""
    
  '日期
    If gdaRealDate1.GetValue <> "" Then Call subAnd("A.RealDate >= To_Date('" & gdaRealDate1.GetValue & "','YYYYMMDD')")
    If gdaRealDate2.GetValue <> "" Then Call subAnd("A.RealDate < To_Date('" & gdaRealDate2.GetValue & "','YYYYMMDD')+1")
    
    If gdtPRTime1.GetValue <> "" Then
        Call subAnd2(strSO00XWhere, "D.ResvTime>=To_Date('" & gdtPRTime1.GetValue & "','yyyymmddhh24mi')")
    End If
    If gdtPRTime2.GetValue <> "" Then
        Call subAnd2(strSO00XWhere, "D.ResvTime<=To_Date('" & gdtPRTime2.GetValue & "','yyyymmddhh24mi')")
    End If
   
  'GiMulti
    'If gimAreaCode.GetQryStr <> "" Then Call subAnd("A.AreaCode " & gimAreaCode.GetQryStr)
    If gimAreaCode.GetQryStr <> "" Then Call subAnd2(strSO00XWhere, "SO014.AreaCode " & gimAreaCode.GetQryStr)
    If gimCitemCode.GetQryStr <> "" Then Call subAnd("A.CitemCode " & gimCitemCode.GetQryStr)
    If gimServiceType.GetQryStr <> "" Then
        Call subAnd("A.ServiceType " & gimServiceType.GetQryStr)
        Call subAnd2(strSO00XWhere, " D.ServiceType " & gimServiceType.GetQryStr)
    End If
    If gimGroupCode.GetQryStr <> "" Then
        Call subAnd2(strSO00XWhere, " D.WorkerEn1 " & gimGroupCode.GetQryStr)
    End If
    '#5411 增加稽核人員條件 By Kin 2009/12/10
    If gimCheckEn.GetQryStr <> "" Then
        Call subAnd2(strSO00XWhere, "D.CheckEn " & gimCheckEn.GetQryStr)
    End If
    'If gimServCode.GetQryStr <> "" Then Call subAnd("A.ServCode " & gimServCode.GetQryStr)
    If gimServCode.GetQryStr <> "" Then Call subAnd2(strSO00XWhere, "D.ServCode " & gimServCode.GetQryStr)
    If gimBillType.GetQryStr <> "" Then
        Call subAnd("SubStr(A.BillNo,7,1) " & gimBillType.GetQryStr)
        Call subAnd2(strSO00XWhere, " SubStr(D.SNO,7,1) " & gimBillType.GetQryStr)
    Else
        Call subAnd("SubStr(A.BillNo,7,1) In('I','M','P')")
    End If
    '#5593 增加完工時間 By Kin 2010/03/19
    If gdtFinTime1.GetValue <> "" Then
        Call subAnd2(strSO00XWhere, "D.FinTime>=To_Date('" & gdtFinTime1.GetValue & "','yyyymmddhh24mi')")
    End If
    If gdtFinTime2.GetValue <> "" Then
        Call subAnd2(strSO00XWhere, "D.FinTime<=To_Date('" & gdtFinTime2.GetValue & "','yyyymmddhh24mi')")
    End If
    
    '#5593 增加工程組別 By Kin 2010/03/19
    If gimWorkGroup.GetQryStr <> "" Then
        strCD071 = ",CM003,CD071"
        strCM003Where = " AND D.WorkerEn1=CM003.EmpNo AND CM003.WorkClass=CD071.CodeNo "
    Else
        strCD071 = Empty
        strCM003Where = Empty
    End If
    
  'GiList
    If gilCompCode.GetCodeNo <> "" Then
        Call subAnd("A.CompCode =" & gilCompCode.GetCodeNo)
        Call subAnd2(strSO00XWhere, "D.CompCode=" & gilCompCode.GetCodeNo)
    End If
    
    
    If optIVR.Value Then Call subAnd2(strSO00XWhere, " D.IVRFLAG > 0 ")
    If optNoIVR.Value Then Call subAnd2(strSO00XWhere, " D.IVRFLAG = 0 ")
    
    '***************************************************************************
    '#5195 將Group By 條件拿掉 By Kin 2009/07/21
'    Select Case True
'           Case optAreaCode.Value '行政區
'                strGroupName = "GroupName={V.AreaCode};strGroupName={V.AreaName}"
'                strpagetype = "行政區"
'           Case optClctEn.Value '工程人員
'                strGroupName = "GroupName={V.WorkerEn1};strGroupName={V.WorkerName1}"
'                strpagetype = "工程人員"
'           Case optServCode.Value '服務區
'                strGroupName = "GroupName={V.ServCode};strGroupName={V.ServName}"
'                strpagetype = "服務區"
'    End Select
    '***************************************************************************
'    If optSimple.Value Then
'        strGroupName = "GroupName={V.ServCode};strGroupName={V.ServName}"
'    Else
'        strGroupName = "GroupName={V.SNO};strGroupName={V.SNO}" '明細表強迫用SNO By Kin 2009/07/21
'    End If
    If optDetail.Value Then
        strGroupName = "GroupName={V.SNO};strGroupName={V.SNO}" '明細表強迫用SNO By Kin 2009/07/21
    Else
        strGroupName = "GroupName={V.ServCode};strGroupName={V.ServName}"
    End If
    strGroupName = strGroupName & ";Visable=" & chkSO033.Value
    strGroupName = strGroupName & ";blnIVR=1"
'    If optIVR.Value Then
'        strGroupName = strGroupName & ";blnIVR=1"
'    Else
'        strGroupName = strGroupName & ";blnIVR=0"
'    End If
    '#5090 明細表增加派工狀態 By Kin 2009/05/06
    'If optDetail.Value Then
        If optFin.Value Then
            Call subAnd2(strSO00XWhere, "D.FINTIME IS NOT NULL")
        ElseIf optRet.Value Then
            Call subAnd2(strSO00XWhere, "D.Returncode is not null")
        ElseIf optNotFin.Value Then
            Call subAnd2(strSO00XWhere, "D.RETURNCODE IS  NULL AND D.FINTIME IS NULL")
        ElseIf optNotAns.Value Then
            Call subAnd2(strSO00XWhere, "D.CALLOKTIME IS NULL")
        End If
    'End If
  '排序方式
    strGroupName = strGroupName & ";Sort0={V.ServiceType};Sort1={@Status};Sort2={@InstKind}"
    If gimOrder.GetColumnOrderCode <> "" Then
      arrOrder = Split(gimOrder.GetColumnOrderCode, ",")
      intSort = 0
      '#5195 排序條件這3個要寫死 By Kin 2009/07/21
      For Each varOrder In arrOrder
          Select Case arrOrder(intSort)
                 Case "A"
                      strGroupName = strGroupName & ";Sort" & intSort + 3 & "={V.CustId}"
                 Case "B"
                      strGroupName = strGroupName & ";Sort" & intSort + 3 & "={V.WorkerEn1}"
                 Case "C"
                      strGroupName = strGroupName & ";Sort" & intSort + 3 & "={V.SNo}"
                 Case "D"
                      strGroupName = strGroupName & ";Sort" & intSort + 3 & "={V.RealDate}"
                 Case "E"
                      strGroupName = strGroupName & ";Sort" & intSort + 3 & "={V.AddrSort}"
                 Case "F"
                      strGroupName = strGroupName & ";Sort" & intSort + 3 & "={V.ClassCode}"
          End Select
          intSort = intSort + 1
      Next
    End If
    
    strChooseIVRStatus = "派工狀態:"
    If optFin.Value Then
        strChooseIVRStatus = strChooseIVRStatus & "已完成"
    ElseIf optRet.Value Then
        strChooseIVRStatus = strChooseIVRStatus & "退單"
    ElseIf optNotFin.Value Then
        strChooseIVRStatus = strChooseIVRStatus & "未完成"
    ElseIf optNotAns.Value Then
        strChooseIVRStatus = strChooseIVRStatus & "未回報"
    Else
        strChooseIVRStatus = strChooseIVRStatus & "全部"
    End If
    
    strChooseString = "預約日期:" & gdtPRTime1.GetValue(True) & "~" & gdtPRTime2.GetValue(True) & ";" & _
                      "完工日期:" & gdtFinTime1.GetValue(True) & "~" & gdtFinTime2.GetValue(True) & ";" & _
                      "實收日期: " & subSpace(gdaRealDate1.GetValue(True)) & "~" & subSpace(gdaRealDate2.GetValue(True)) & ";" & _
                      "工程人員: " & subSpace(gimGroupCode.GetDispStr) & ";" & _
                      "稽核人員: " & subSpace(gimCheckEn.GetDispStr) & ";" & _
                      "公司別  : " & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別: " & subSpace(gimServiceType.GetDispStr) & ";" & _
                      "收費項目: " & subSpace(gimCitemCode.GetDispStr) & ";" & _
                      "服務區　: " & subSpace(gimServCode.GetDispStr) & ";" & _
                      "行政區　: " & subSpace(gimAreaCode.GetDispStr) & ";" & _
                      "單據類別: " & subSpace(gimBillType.GetDispStr) & ";" & _
                      "工程組別:" & subSpace(gimWorkGroup.GetDispStr) & ";" & _
                      "統計歷史資料:" & IIf(chkHistory.Value, "是", "否") & ";" & _
                      "不顯示無收費之工單:" & IIf(chkDispBillWip.Value, "是", "否") & ";" & _
                      IIf(optDetail.Value, strChooseIVRStatus, "") & ";" & _
                      IIf(optDetail.Value, IIf(chkSO033.Value, "是否顯示收費資料:是", "是否顯示收費資料:否"), "") & ";" & _
                      "回單來源:" & IIf(optIVR.Value, "IVR", IIf(optNoIVR.Value, "非IVR", "全部")) & ";" & _
                      "排序方式: " & subSpace(gimOrder.GetColumnOrderDspStr) & ";"
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub subPrint()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    '建立View
    Call subCreateView
  On Error GoTo ChkErr
    Set rsTmp = gcnGi.Execute("Select Count(*) as intCount From " & strViewName & " Where RowNum=1")
    If rsTmp("intCount") = 0 Then
        MsgNoRcd
        SendSQL , , True
    Else
        strSQL = ""
        strSQL = "SELECT * FROM " & strViewName & " V"
        If blnExcel Then
            Call toExcel(strSQL)
        Else
            If optDetail.Value Then
                Call PrintRpt2(GetPrinterName(5), "SO52J0A.rpt", , "IVR回單稽核報表 [SO52J0A]", strSQL, strChooseString, , True, , , strGroupName, GiPaperLandscape)
            Else
                Select Case True
                    Case optWorker.Value
                        Call PrintRpt2(GetPrinterName(5), "SO52J0A2.rpt", , "工單回收簡表 [SO52J0A2]", strSQL, strChooseString, , True, , , , GiPaperLandscape)
                    Case optServ.Value
                        Call PrintRpt2(GetPrinterName(5), "SO52J0A3.rpt", , "工單回收簡表 [SO52J0A2]", strSQL, strChooseString, , True, , , , GiPaperLandscape)
                    Case optTotal.Value
                        Call PrintRpt2(GetPrinterName(5), "SO52J0A4.rpt", , "工單回收簡表 [SO52J0A2]", strSQL, strChooseString, , True, , , , GiPaperLandscape)
                End Select
            End If
        
'            If optSimple Then
'                Call PrintRpt2(GetPrinterName(5), "SO52J0A2.rpt", , "工單回收簡表 [SO52J0A2]", strsql, strChooseString, , True, , , , GiPaperLandscape)
'            Else
'                Call PrintRpt2(GetPrinterName(5), "SO52J0A.rpt", , "IVR回單稽核報表 [SO52J0A]", strsql, strChooseString, , True, , , strGroupName, GiPaperLandscape)
'            End If
        End If
    End If
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subPrint")
End Sub

Private Sub Form_Initialize()
  On Error GoTo ChkErr
    blnUseIVR = True
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Initialize")
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
    Call SetgiMulti(gimServiceType, "CodeNo", "Description", "CD046", "服務類別代碼", "服務類別名稱")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱")
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱")
    Call SetgiMulti(gimGroupCode, "EmpNo", "EmpName", "CM003", "工程人員代號", "工程人員名稱")
    '#5411 增加稽核人員 By Kin 2009/12/10
    Call SetgiMulti(gimCheckEn, "EmpNo", "EmpName", "CM003", "稽核人員代號", "稽核人員名稱")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱")
    Call SetgiMultiAddItem(gimOrder, "A,B,C,D,E,F", "客戶編號,工程人員代號,工單單號,收費日期,地址,客戶類別")
    '#5593 增加工程組別
    Call SetgiMulti(gimWorkGroup, "CodeNo", "Description", "CD071", "工作類別代碼", "工作類別名稱")
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
    Call SetgiMultiAddItem(gimBillType, "I,P,M", "裝機單,停拆移機單,維修單")
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub
  
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO52J0A)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  gcnGi.Execute "Droup View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO52J0A
End Sub

Private Sub gdaRealDate1_GotFocus()
  On Error GoTo ChkErr
    If gdaRealDate1.GetValue = "" Then gdaRealDate1.SetValue (RightDate)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaRealDate1_GotFocus")
End Sub

Private Sub gdaRealDate2_GotFocus()
  On Error GoTo ChkErr
    If gdaRealDate1.GetValue = "" Or gdaRealDate2.GetValue = "" Then gdaRealDate2.SetValue (gdaRealDate1.GetValue)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaRealDate2_GotFocus")
End Sub

Private Sub gdaRealDate2_Validate(Cancel As Boolean)
  On Error Resume Next
    If gdaRealDate2.GetValue < gdaRealDate1.GetValue Then MsgBox "截止日期必須大於起始日期": gdaRealDate2.SetFocus
End Sub



Private Function subCreateView() As Boolean
  Dim strView As String
  Dim i, j As Integer
  Dim strQryCnt As String
  Dim strQrySQL As String
  Dim strWhere As String
  Dim strAddrWhere As String
  Dim strCD007 As String
  Dim strJoinAll As String
  Dim strblnIVR As String
  Dim strQrySQL1 As String
  Dim strQrySQL2 As String
  Dim strQrySQL3 As String
  Dim intHistory As Integer
  
    On Error Resume Next
      gcnGi.Execute "Drop View " & strViewName
    On Error GoTo ChkErr
    
    strViewName = GetTmpViewName
    strQryCnt = Empty
    strQrySQL = Empty
    strJoinAll = Empty
    strJoinAll = "(+)"
    strblnIVR = Empty
    strblnIVR = " D.IVRDATAMATCH<>1 And D.FinTime Is Not Null  "
    If chkHistory.Value = 1 Then
        intHistory = 2
    Else
        intHistory = 0
    End If


    '***************************************************************************
    '#5039 測試不OK,如果有實收日期或收費項目就不要ourjoin By Kin 2009/04/27
    If gdaRealDate1.GetValue <> "" Or gdaRealDate2.GetValue <> "" Then
        strJoinAll = Empty
    End If
    If gimCitemCode.GetQryStr <> "" Then
        strJoinAll = Empty
    End If
    '***************************************************************************
    '#5195 因為簡表大幅改版,所以不管簡表或明細都改用明細表的資料 By Kin 2009/07/21
    '#5195 增加IVR簡碼欄位IVRShortSNo By Kin 2009/07/28
    If (optDetail.Value) Or (True) Then
        For i = 0 To intHistory
            For j = 0 To 2
                If j <> 2 Then
                    strAddrWhere = " And D.AddrNo = SO014.AddrNo"
                    strCD007 = Empty
                    If j = 0 Then
                        '2012.01.31 #6208 增加判斷SO041.StartWebAPI By Miggie
                        Dim blnStartWebAPI As Boolean
                        blnStartWebAPI = Val(GetRsValue("Select StartWebAPI From SO041 Where Sysid=" & gCompCode, gcnGi) & "") = 1
                        If blnStartWebAPI Then
                            strAddrWhere = strAddrWhere & " And ((Nvl(D.CreateByAPI,0)=0 AND Nvl(D.APICheckOk,0)=0) or Nvl(D.APICheckOk,0)=1) "
                        End If
                    End If
                Else
                    'SO009的AddrNO要特別處理,裝機類別的參考號不同對應的地址資料也不一樣
                    strCD007 = ",CD007"
                    strAddrWhere = " And D.PRCODE=CD007.CodeNo(+) And " & _
                            " DeCode(Nvl(CD007.RefNo,0),3,D.REINSTADDRNO,D.OLDADDRNO)=SO014.AddrNO "
                End If
                If strQrySQL <> Empty Then strQrySQL = strQrySQL & " Union All "
                '#5593 增加工程組別(strCM003Where、strCD071) By Kin 2010/03/19
                strQrySQL = strQrySQL & " Select D.CUSTID,D.SNO,A.CITEMNAME,A.REALSTARTDATE,A.REALSTOPDATE," & _
                            "A.SHOULDDATE,Nvl(A.SHOULDAMT,0) SHOULDAMT,A.REALDATE,NVL(A.REALAMT,0) REALAMT," & _
                            "SO001.CUSTNAME,SO001.TEL1,D.IVRShortSNo," & _
                            "SO014.ADDRESS,SO014.AREACODE,SO014.AreaName,D.WorkerEN1,D.WorkerName1," & _
                            "D.ServCode,CD002.Description ServName,SO014.AddrNo,A.ClassCode,SO014.AddrSort," & _
                            "D.ServiceType,D.IVRDATAMATCH,D.FINTIME,D.RETURNCODE,D.CALLOKTIME,A.CLCTEN,A.CLCTNAME " & _
                            " From (Select A.* From SO001,SO014,SO03" & i + 3 & " A " & strChoose & _
                            " And Nvl(A.CancelFlag,0)=0 ) A,SO00" & j + 7 & " D,SO001,CD002,SO014 " & _
                            strCD007 & strCD071 & " Where A.BillNo" & strJoinAll & "=D.SNO" & _
                            " And " & strSO00XWhere & _
                            " And D.CustId=SO001.CustId " & _
                            " And D.ServCode=CD002.CodeNo" & _
                            strAddrWhere & strCM003Where
                            
            Next j
        Next i
        If chkHistory.Value Then
            strQrySQL1 = "Select Distinct A.* From (" & strQrySQL & ") A Where CitemName is Not Null"   '找出有收費料的工單
            strQrySQL2 = "Select Distinct A.* From (" & strQrySQL & ") A Where CitemName is  Null"      '找出沒有收費料的工單
            strQrySQL3 = "Select A.SNO From (" & strQrySQL & ") A Where CitemName is Not Null"
            strQrySQL = strQrySQL1 & " Union " & strQrySQL2 & " And SNO Not In(" & strQrySQL3 & ")"     '將有收費資料的工單 Union 沒有收費資料的工單,但有出現過收費資料的工單要過濾掉
        End If
        '#5892 增加不顯示無收費之工單條件，如果打勾，一定要有收費資料才可顯示 By Kin 2011/05/23
        If chkDispBillWip Then
            strQrySQL = "SELECT * FROM (" & strQrySQL & ") WHERE CITEMNAME IS NOT NULL "
        End If
        strView = "Create View " & strViewName & " As(" & strQrySQL & ")"
    '#5090 增加退單與未回報 By Kin 2009/05/08
    Else
        For i = 1 To 2
            For j = 0 To 2
                If strQryCnt <> Empty Then strQryCnt = strQryCnt & " Union All "
                strQryCnt = strQryCnt & " Select  D.SNO FSNO,Nvl(Sum(A.ShouldAmt),0) FAMT,NULL NSNO,0 NAMT,NULL ESNO,0 EAMT,D.WORKEREN1,D.WORKERNAME1 " & _
                    ",Null RSNO,0 RAMT,NULL CSNO,0 CAMT,NULL XSNO,0 XAMT " & _
                    " From ( Select A.* From SO001,SO014,SO03" & i + 3 & " A " & strChoose & _
                    " ) A,SO00" & j + 7 & " D Where A.BillNo" & strJoinAll & "=D.SNO And " & strSO00XWhere & _
                    " GROUP BY D.WORKEREN1,D.WORKERNAME1,D.SNO " & _
                    " UNION  " & _
                    " Select NULL FSNO,0 FAMT, D.SNO NSNO,NVL(SUM(A.SHOULDAMT),0) NAMT,NULL ESNO,0 EAMT,D.WORKEREN1,D.WORKERNAME1 " & _
                    ",Null RSNO,0 RAMT,NULL CSNO,0 CAMT,NULL XSNO,0 XAMT " & _
                    " From ( Select A.* From SO001,SO014,SO03" & i + 3 & " A " & strChoose & _
                    " ) A,SO00" & j + 7 & " D Where A.BillNo" & strJoinAll & "=D.SNO AND" & strblnIVR & " And " & strSO00XWhere & _
                    " GROUP BY D.WORKEREN1,D.WORKERNAME1,D.SNO " & _
                    " UNION  " & _
                    " Select NULL FSNO,0 FAMT,NULL NSNO,0 NAMT,D.SNO ESNO,NVL(SUM(A.SHOULDAMT),0) EAMT,D.WORKEREN1,D.WORKERNAME1 " & _
                    ",Null RSNO,0 RAMT,NULL CSNO,0 CAMT,NULL XSNO,0 XAMT " & _
                    " From ( Select A.* From SO001,SO014,SO03" & i + 3 & " A " & strChoose & _
                    " ) A,SO00" & j + 7 & " D Where A.BillNo" & strJoinAll & "=D.SNO AND D.IVRDataMatch=1 " & _
                    " And D.FinTime Is Not Null And  " & strSO00XWhere & _
                    " GROUP BY D.WORKEREN1,D.WORKERNAME1,D.SNO" & _
                    " UNION " & _
                    " Select NULL FSNO,0 FAMT, NULL NSNO,0 NAMT,NULL ESNO,0 EAMT,D.WORKEREN1,D.WORKERNAME1 " & _
                    ", D.SNO RSNO, NVL(SUM(A.SHOULDAMT),0) RAMT,NULL CSNO,0 CAMT,NULL XSNO,0 XAMT " & _
                    " From ( Select A.* From SO001,SO014,SO03" & i + 3 & " A " & strChoose & _
                    " ) A,SO00" & j + 7 & " D Where A.BillNo" & strJoinAll & "=D.SNO AND D.Returncode is not null  And " & strSO00XWhere & _
                    " GROUP BY D.WORKEREN1,D.WORKERNAME1,D.SNO "
                strQryCnt = strQryCnt & " UNION " & _
                    " Select NULL FSNO,0 FAMT, NULL NSNO,0 NAMT,NULL ESNO,0 EAMT,D.WORKEREN1,D.WORKERNAME1 " & _
                    ", NULL RSNO, 0 RAMT,D.SNO CSNO,NVL(SUM(A.SHOULDAMT),0) CSUM,NULL XSNO,0 XAMT " & _
                    " From ( Select A.* From SO001,SO014,SO03" & i + 3 & " A " & strChoose & _
                    " ) A,SO00" & j + 7 & " D Where A.BillNo" & strJoinAll & "=D.SNO AND D.CALLOKTIME IS NULL And " & strSO00XWhere & _
                    " GROUP BY D.WORKEREN1,D.WORKERNAME1,D.SNO "
                strQryCnt = strQryCnt & " UNION " & _
                    " Select NULL FSNO,0 FAMT, NULL NSNO,0 NAMT,NULL ESNO,0 EAMT,D.WORKEREN1,D.WORKERNAME1 " & _
                    ", NULL RSNO, 0 RAMT,NULL CSNO,0 CSUM,D.SNO XSNO,NVL(SUM(A.SHOULDAMT),0) XAMT " & _
                    " From ( Select A.* From SO001,SO014,SO03" & i + 3 & " A " & strChoose & _
                    " ) A,SO00" & j + 7 & " D Where A.BillNo" & strJoinAll & "=D.SNO AND D.RETURNCODE IS NULL And D.FINTIME IS NULL And " & strSO00XWhere & _
                    " GROUP BY D.WORKEREN1,D.WORKERNAME1,D.SNO "
                    
            Next
        Next i
        strQryCnt = "Select COUNT(DISTINCT FSNO) CNT,SUM(FAMT) FAMT ," & _
                    "COUNT(DISTINCT NSNO) NSNO,SUM(NAMT) NAMT ," & _
                    " COUNT(DISTINCT ESNO) ESNO,SUM(EAMT) EAMT ," & _
                    " COUNT(DISTINCT RSNO) RSNO,SUM(RAMT) RAMT ," & _
                    " COUNT(DISTINCT CSNO) CSNO,SUM(CAMT) CAMT ," & _
                    " COUNT(DISTINCT XSNO) XSNO,SUM(XAMT) XAMT ," & _
                    "WORKEREN1,WORKERNAME1 FROM (" & _
                strQryCnt & ") Group By WORKEREN1,WORKERNAME1 "
        strView = "Create View " & strViewName & " As(" & strQryCnt & ")"

    End If
    gcnGi.Execute strView
    
    SendSQL strView, True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function

Private Sub gdtFinTime1_Validate(Cancel As Boolean)
  On Error Resume Next
    If gdtFinTime1.GetValue & "" <> "" Then
        If gdtFinTime2.GetValue & "" = "" Then
            gdtFinTime2.SetValue gdtFinTime1.GetDate & "2359"
        End If
    End If
End Sub

Private Sub gdtFinTime2_Validate(Cancel As Boolean)
  On Error Resume Next
    Cancel = ChkDate2(gdtFinTime1, gdtFinTime2)
End Sub

Private Sub gdtPRTime1_Validate(Cancel As Boolean)
    On Error Resume Next
    If gdtPRTime1.GetValue & "" <> "" Then
        If gdtPRTime2.GetValue & "" = "" Then
            gdtPRTime2.SetValue gdtPRTime1.GetDate & "2359"
        End If
    End If
    
End Sub

Private Sub gdtPRTime2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdtPRTime1, gdtPRTime2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gimServiceType)
    gimServiceType.Clear
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimGroupCode, , gilCompCode.GetCodeNo
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
Private Sub gimServiceType_Change()
  On Error GoTo ChkErr
    GiMultiFilter gimCitemCode, gimServiceType.GetQryStr, , True
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gimServiceType_Change")
End Sub

Private Sub gimWorkGroup_SelectChange()
  On Error GoTo ChkErr
    gimGroupCode.Clear
    If gimWorkGroup.GetQryStr <> "" Then
        Call SetgiMulti(gimGroupCode, "EmpNo", "EmpName", "CM003", "工程人員代號", "工程人員名稱", " Where WorkClass " & gimWorkGroup.GetQryStr)
    Else
        gimGroupCode.Filter = ""
        Call SetgiMulti(gimGroupCode, "EmpNo", "EmpName", "CM003", "工程人員代號", "工程人員名稱")
    End If
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gimWorkGroup_SelectChange")
End Sub

Private Sub optDetail_Click()
  On Error Resume Next
    gimOrder.Enabled = True
 '   fraStatus.Enabled = True
    chkSO033.Enabled = True
    cmdExcel.Enabled = True
    optWorker.Value = False
    optServ.Value = False
    optTotal.Value = False
End Sub

Private Sub optServ_Click()
  On Error Resume Next
    optDetail.Value = False
    cmdExcel.Enabled = False
End Sub

Private Sub optSimple_Click()
  On Error Resume Next
    gimOrder.Clear
    gimOrder.Enabled = False
'    fraStatus.Enabled = False
    chkSO033.Enabled = False
    cmdExcel.Enabled = False
End Sub
Private Sub toExcel(ByVal strSQL As String)
  On Error GoTo ChkErr
    Dim rsExcel As New ADODB.Recordset
      RptToTxt RptName("SO52J0", "E"), , strSQL, , _
               , , strGroupName, , , garyGi(19) & "\SO52J0", False
      If Not Get_RS_From_Txt(garyGi(19), "SO52J0.txt", rsExcel) Then blnExcel = False: Exit Sub
       '#7074 To Add a choice-condition parameter into excel print function by Kin 2015/09/09
      Call UseProperty(rsExcel, "IVR回單稽核報表", "第一頁", _
                    , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , strChooseString)
      blnExcel = False
      On Error Resume Next
      CloseRecordset rsExcel
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "toExcel")
End Sub

Private Sub optTotal_Click()
 On Error Resume Next
    optDetail.Value = False
    cmdExcel.Enabled = False
End Sub

Private Sub optWorker_Click()
  On Error Resume Next
    optDetail.Value = False
    cmdExcel.Enabled = False
End Sub
