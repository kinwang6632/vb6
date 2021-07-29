VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.17#0"; "csMulti.ocx"
Begin VB.Form frmSO5510A 
   BorderStyle     =   1  '單線固定
   Caption         =   "現有各類客戶數統計[SO5510A]"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "SO5510A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form48"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.Frame fraCount 
      Caption         =   "統計對象"
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
      Left            =   60
      TabIndex        =   41
      Top             =   6840
      Width           =   4935
      Begin VB.OptionButton optCountNotMduid 
         Caption         =   "非大樓戶"
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
         Height          =   345
         Left            =   3450
         TabIndex        =   44
         Top             =   270
         Width           =   1185
      End
      Begin VB.OptionButton optCountMduid 
         Caption         =   "大樓戶"
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
         Height          =   345
         Left            =   1830
         TabIndex        =   43
         Top             =   270
         Width           =   915
      End
      Begin VB.OptionButton optCountAll 
         Caption         =   "全部"
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
         Height          =   345
         Left            =   240
         TabIndex        =   42
         Top             =   270
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.PictureBox pic2 
      Height          =   4005
      Left            =   0
      ScaleHeight     =   3945
      ScaleWidth      =   7935
      TabIndex        =   25
      Top             =   1920
      Width           =   7995
      Begin VB.VScrollBar vsl1 
         Height          =   3945
         LargeChange     =   100
         Left            =   7650
         Max             =   100
         SmallChange     =   100
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   0
         Width           =   285
      End
      Begin VB.Frame fraMula 
         BorderStyle     =   0  '沒有框線
         Height          =   4755
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   8115
         Begin Gi_Multi.GiMulti gimStatusCode 
            Height          =   375
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   661
            ButtonCaption   =   "客  戶  狀  態"
         End
         Begin Gi_Multi.GiMulti gimWipCode3 
            Height          =   375
            Left            =   0
            TabIndex        =   28
            Top             =   360
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   661
            ButtonCaption   =   "非停拆移機中"
         End
         Begin CS_Multi.CSmulti gimClassCode 
            Height          =   375
            Left            =   0
            TabIndex        =   29
            Top             =   720
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   661
            ButtonCaption   =   "客  戶  類  別"
         End
         Begin CS_Multi.CSmulti gimClassCode2 
            Height          =   375
            Left            =   0
            TabIndex        =   30
            Top             =   1080
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   661
            ButtonCaption   =   "客戶類別(二)"
         End
         Begin CS_Multi.CSmulti gimClassCode3 
            Height          =   375
            Left            =   0
            TabIndex        =   31
            Top             =   1440
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   661
            ButtonCaption   =   "客戶類別(三)"
         End
         Begin Gi_Multi.GiMulti gimServCode 
            Height          =   375
            Left            =   0
            TabIndex        =   32
            Top             =   2160
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   661
            ButtonCaption   =   "服     務     區"
         End
         Begin Gi_Multi.GiMulti gimAreaCode 
            Height          =   375
            Left            =   0
            TabIndex        =   33
            Top             =   2520
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   661
            ButtonCaption   =   "行     政     區"
         End
         Begin Gi_Multi.GiMulti gimClctAreaCode 
            Height          =   375
            Left            =   0
            TabIndex        =   34
            Top             =   1800
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   661
            ButtonCaption   =   "收     費     區"
         End
         Begin CS_Multi.CSmulti gimStrtRange 
            Height          =   375
            Left            =   0
            TabIndex        =   35
            Top             =   2880
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   661
            ButtonCaption   =   "街  道  範  圍"
         End
         Begin Gi_Multi.GiMulti gimSales 
            Height          =   375
            Left            =   0
            TabIndex        =   36
            Top             =   3240
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   661
            ButtonCaption   =   "業     務     區"
         End
         Begin Gi_Multi.GiMulti gimFaciCode 
            Height          =   375
            Left            =   0
            TabIndex        =   37
            Top             =   3600
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   661
            ButtonCaption   =   "設  備  項  目"
         End
         Begin Gi_Multi.GiMulti gimNodeNo 
            Height          =   375
            Left            =   0
            TabIndex        =   38
            Top             =   3960
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   661
            ButtonCaption   =   "NodeNo"
            OneColumn       =   -1  'True
         End
         Begin Gi_Multi.GiMulti gimJobPeriod 
            Height          =   375
            Left            =   0
            TabIndex        =   39
            Top             =   4320
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   661
            ButtonCaption   =   "工  期  類  別"
         End
      End
   End
   Begin VB.CommandButton cmdNightRun 
      Caption         =   "Night-run報表[&N]"
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
      Left            =   4080
      TabIndex        =   24
      Top             =   7710
      Width           =   1695
   End
   Begin VB.CheckBox chkMduId 
      Caption         =   "同一大樓戶視為同一客戶"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3870
      TabIndex        =   5
      Top             =   600
      Width           =   3735
   End
   Begin VB.Frame Frame2 
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
      Height          =   495
      Left            =   3840
      TabIndex        =   20
      Top             =   0
      Width           =   4155
      Begin VB.OptionButton optClctAreaCode 
         Caption         =   "收費區"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2820
         TabIndex        =   4
         Top             =   210
         Width           =   855
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
         Height          =   315
         Left            =   1890
         TabIndex        =   3
         Top             =   150
         Width           =   885
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
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   150
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "小計依據:"
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
         Left            =   90
         TabIndex        =   21
         Top             =   210
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   0
      TabIndex        =   17
      Top             =   870
      Width           =   7995
      Begin VB.CheckBox chkYes 
         Caption         =   "大樓統收子戶是否列入統計"
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
         Left            =   4230
         TabIndex        =   9
         Top             =   660
         Width           =   3270
      End
      Begin Gi_Date.GiDate gdaClctDate2 
         Height          =   375
         Left            =   2625
         TabIndex        =   8
         Top             =   555
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
         Left            =   1275
         TabIndex        =   7
         Top             =   555
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
      Begin Gi_Multi.GiMulti gimCitemCode 
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Top             =   165
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   635
         ButtonCaption   =   "收  費  項  目"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "下次收費日"
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
         Left            =   135
         TabIndex        =   19
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "---"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2415
         TabIndex        =   18
         Top             =   645
         Width           =   180
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "網路編號"
      Height          =   735
      Left            =   30
      TabIndex        =   16
      Top             =   6030
      Width           =   7965
      Begin VB.TextBox mskCircuitNo 
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   90
         TabIndex        =   10
         Top             =   240
         Width           =   4605
      End
      Begin VB.OptionButton optAll 
         Caption         =   "全部"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6660
         TabIndex        =   12
         Top             =   270
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton optCircuitNo 
         Caption         =   "只印空白網路編號"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   4830
         TabIndex        =   11
         Top             =   240
         Width           =   1845
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
      Left            =   1890
      TabIndex        =   14
      Top             =   7710
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
      Left            =   6780
      TabIndex        =   15
      Top             =   7710
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
      Left            =   30
      TabIndex        =   13
      Top             =   7710
      Width           =   1245
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   525
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
      Left            =   840
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
      F5Corresponding =   -1  'True
   End
   Begin VB.Label Label4 
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
      Left            =   45
      TabIndex        =   23
      Top             =   585
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
      Left            =   30
      TabIndex        =   22
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "frmSO5510A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table:CD035,SO001,SO014(,SO002)
Option Explicit
Dim strAreaName As String
Dim GroupStr As String
Dim strServArea As String
Dim strServArea1 As String
Dim strServiceType As String
Dim strTitleChoose As String
Dim rsTmp1 As New ADODB.Recordset
Dim strChooseSO003 As String
Dim strChooseNot As String
Dim blnSO004 As Boolean
Dim blnSO014 As Boolean

Private Sub chkMduId_Click()
    If optCountNotMduid.Value Then
        chkMduId.Value = 0
    End If
End Sub

Private Sub chkYes_Click()
  On Error Resume Next
    If optCountNotMduid Then
        chkYes.Value = 0
    End If
End Sub

Private Sub cmdExit_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdNightRun_Click()
On Error GoTo ChkErr
  Dim strCodeNo As Variant
  
  strCodeNo = gilServiceType.GetCodeNo
  With frmSO5510B
    Set .uParentForm = Me
    .uServiceType = strCodeNo
    .Show vbModal
  End With
ChkErr:
  Call ErrSub(Me.Name, "cmdNightRun_Click")
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
    Call PreviousRpt(GetPrinterName(5), RptName("SO5510"), "現有各類客戶數統計 [SO5510A]")
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
  Dim strSubQry3 As String
  Dim strSubQry4 As String
  Dim strSubQry5 As String
  Dim strSubQry6 As String
  Dim strSubQry7 As String
  Dim strSubQry8 As String
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    Screen.MousePointer = vbHourglass
      cmdExit.SetFocus
      Me.Enabled = False
        ReadyGoPrint
        Call subChoose
        If Not subPrint Then Me.Enabled = True: Screen.MousePointer = vbDefault: Exit Sub
        Set rsTmp = cnn.Execute("SELECT ServCode,ServArea FROM SO5510A1")
        If rsTmp.EOF Then
           MsgBox "查無資料!!!", , "提示"
        Else
          If Not Sample(rsTmp.Fields("ServCode") & "", rsTmp.Fields("ServArea") & "") Then Exit Sub
          sample2
          strsql = "SELECT * FROM SO5510A1"
          strSubQry1 = "Select * From SO5510A2 A"
          strSubQry2 = "Select * From SO5510A3 A"
          strSubQry3 = "Select * From SO5510A4 A"
          strSubQry4 = "Select * From SO5510A5 A"
          strSubQry5 = "Select * From SO5510A6 A"
          strSubQry6 = "Select * From SO5510A7 A"
          strSubQry7 = "Select * from SO5510A8 A"
          strSubQry8 = "Select * from SO5510A9 A"
          If gimStatusCode.GetDispStr <> "" Then
              If gimStatusCode.GetDispStr = "(全選)" Then
                  GroupStr = "strGroup='正常,停機,拆機,註銷,促銷,拆機中';CustStatusCode='1,2,3,4,5,6'"
              Else
                  GroupStr = "strGroup='" & gimStatusCode.GetDispStr & "';CustStatusCode='" & Replace(gimStatusCode.GetQueryCode, "'", "") & "'"
              End If
          Else
              GroupStr = "strGroup='正常,停機,拆機,註銷,促銷,拆機中';CustStatusCode=''"
          End If
          Call PrintRpt(, RptName("SO5510"), , "現有各類客戶數統計 [SO5510A]", strsql, strChooseString, , True, "Tmp0000.MDB", , GroupStr, GiPaperLandscape, , strSubQry1, strSubQry2, strSubQry3, strSubQry4, strSubQry5, , , , , , , strSubQry6, strSubQry7, strSubQry8)
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
    Dim strGroupString As String
    Dim strCountString As String
    Dim strFrom As String
    Dim strFrom2 As String
    strFrom = " SO001,SO002 "
    strChoose = ""
    strChooseString = ""
    strServiceType = ""
    strTitleChoose = ""
    'strChooseSO003 = " Where SO001.CUSTID = SO002.CUSTID"
    strChooseSO003 = ""
    strChooseNot = " Where SO001.CUSTID = SO002.CUSTID"
    blnSO004 = False
    blnSO014 = False
    
  '日期
    If gdaClctDate1.GetValue <> "" Then Call subAnd2(strChooseSO003, "SO003.ClctDate >= To_Date('" & gdaClctDate1.GetValue & "','YYYYMMDD')")
    If gdaClctDate2.GetValue <> "" Then Call subAnd2(strChooseSO003, "SO003.ClctDate < To_Date('" & gdaClctDate2.GetValue & "','YYYYMMDD')+1")
    
  'GiMulti
    If gimStatusCode.GetQryStr <> "" Then strTitleChoose = "SO002.CustStatusCode " & gimStatusCode.GetQryStr
    If gimWipCode3.GetQryStr <> "" Then Call subAnd2(strChooseNot, "NOT SO002.WipCode3 " & gimWipCode3.GetQryStr)
    If gimClassCode.GetQryStr <> "" Then Call subAnd2(strChooseNot, "SO001.ClassCode1 " & gimClassCode.GetQryStr)
    '問題集2760 增加客戶類別2、3
    If gimClassCode2.GetQryStr <> "" Then Call subAnd2(strChooseNot, "SO001.ClassCode2 " & gimClassCode2.GetQryStr)
    If gimClassCode3.GetQryStr <> "" Then Call subAnd2(strChooseNot, "SO001.ClassCode3 " & gimClassCode3.GetQryStr)
    If gimServCode.GetQryStr <> "" Then Call subAnd2(strChooseNot, "SO001.ServCode " & gimServCode.GetQryStr)
    If gimClctAreaCode.GetQryStr <> "" Then Call subAnd2(strChooseNot, "SO001.ClctAreaCode " & gimClctAreaCode.GetQryStr)
    If gimAreaCode.GetQryStr <> "" Then Call subAnd2(strChooseNot, "SO014.AreaCode " & gimAreaCode.GetQryStr): blnSO014 = True
    If gimStrtRange.GetQryStr <> "" Then Call subAnd2(strChooseNot, "SO014.StrtCode " & gimStrtRange.GetQryStr): blnSO014 = True
    If gimSales.GetQryStr <> "" Then Call subAnd2(strChooseNot, "SO014.SalesCode " & gimSales.GetQryStr): blnSO014 = True
    If gimFaciCode.GetQryStr <> "" Then Call subAnd2(strChooseNot, "SO004.FaciCode " & gimFaciCode.GetQryStr): blnSO004 = True
    If gimNodeNo.GetQryStr <> "" Then Call subAnd2(strChooseNot, "SO014.NodeNo " & gimNodeNo.GetQryStr): blnSO014 = True
    If gimJobPeriod.GetQryStr <> "" Then Call subAnd2(strChooseNot, "SO014.NodeNo in (Select NodeNo From CD065A Where CodeNo " & gimJobPeriod.GetQryStr & ")"): blnSO014 = True
    If gimCitemCode.GetQryStr <> "" Then Call subAnd2(strChooseSO003, "SO003.CitemCode " & gimCitemCode.GetQryStr)
    
  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd2(strChooseNot, "SO001.CompCode=" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then
        Call subAnd2(strChooseNot, "instr(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "')>0")
        Call subAnd2(strChooseNot, "SO002.ServiceType='" & gilServiceType.GetCodeNo & "'")
        If blnSO004 Then Call subAnd2(strChooseNot, "SO004.ServiceType='" & gilServiceType.GetCodeNo & "'")
    End If
  
  '網路編號
    If mskCircuitNo.Text <> "" Then
        Call subAnd2(strChooseNot, "SO014.CircuitNo='" & mskCircuitNo.Text & "'")
        blnSO014 = True
    Else
        If optCircuitNo.Value Then Call subAnd2(strChooseNot, "SO014.CircuitNo is null "): blnSO014 = True
    End If
    
    If blnSO014 = True Or optAreaCode Then
        strFrom = strFrom & ",SO014"
        strChooseNot = subAnd2(strChooseNot, "SO001.InstAddrNo = SO014.AddrNo And SO001.CompCode = SO014.CompCode ")
        blnSO014 = True
    End If
    
    If blnSO004 Then
        strFrom = strFrom & ",SO004"
        strChooseNot = subAnd2(strChooseNot, " SO004.CustId = SO001.CustId And SO004.CompCode = SO001.CompCode And SO004.InstDate is Not Null And SO004.PRDate is Null ")
    End If
    '問題集2760 增加判斷大樓戶的判斷
    Select Case True
        Case optCountMduid.Value
          '大樓用戶
            Call subAnd2(strChooseNot, "SO001.Mduid is not null")
            strCountString = "大樓戶"
        Case optCountNotMduid.Value
            Call subAnd2(strChooseNot, "SO001.MduId Is Null")
            strCountString = "非大樓戶"
        Case optCountAll.Value
            strCountString = "全部"
    End Select
    strChoose = strChooseNot
    'strChooseNot = strFrom & strChooseNot
    If strChooseSO003 <> "" Then
        strFrom = strFrom & ",SO003"
        If gilServiceType.GetCodeNo <> "" Then Call subAnd2(strChooseSO003, " SO003.ServiceType = '" & gilServiceType.GetCodeNo & "'")
        strChoose = strChoose & " And SO002.CustId = SO003.CustId And SO002.CompCode = SO003.CompCode And SO002.ServiceType = SO003.ServiceType And " & strChooseSO003
    End If
    
    strChoose = strFrom & strChoose
    '問題集2121 修正為 原上方註解移到這裡(原字串少串到SO003) --Crystal
    strChooseNot = strFrom & strChooseNot
    '小計依據
    Select Case True
           Case optServCode.Value
                strGroupString = "服務區"
           Case optAreaCode.Value
                strGroupString = "行政區"
           Case optClctAreaCode.Value
                strGroupString = "收費區"
    End Select
    
    strChooseString = "公司別　  :" & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別: " & subSpace(gilServiceType.GetDescription) & ";" & _
                      "收費項目  :" & subSpace(gimCitemCode.GetDispStr) & ";" & _
                      "下次收費日:" & subSpace(gdaClctDate1.GetOriginalValue) & "~" & subSpace(gdaClctDate2.GetOriginalValue) & ";" & _
                      "大樓統收子戶是否列入統計:" & IIf(chkYes.Value = 1, "是", "否") & ";" & _
                      "網路編號  :" & subSpace(mskCircuitNo.Text) & ";" & _
                      "同一大樓戶視為同一客戶:" & subSpace(IIf(chkMduId.Value = 1, "是", "否")) & ";" & _
                      "客戶狀態  :" & subSpace(gimStatusCode.GetDispStr) & ";" & _
                      "停拆移機類狀態  :" & subSpace(gimWipCode3.GetDispStr) & ";" & _
                      "客戶類別  :" & subSpace(gimClassCode.GetDispStr) & ";" & _
                      "收費區    : " & subSpace(gimClctAreaCode.GetDispStr) & ";" & _
                      "服務區　  :" & subSpace(gimServCode.GetDispStr) & ";" & _
                      "行政區　  :" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                      "街道範圍  :" & subSpace(gimStrtRange.GetDispStr) & ";" & _
                      "業務區    :" & subSpace(gimSales.GetDispStr) & ";" & _
                      "設備項目    :" & subSpace(gimFaciCode.GetDispStr) & ";" & _
                      "NodeNo  :" & subSpace(gimNodeNo.GetDispStr) & ";" & _
                      "工期類別:" & subSpace(gimJobPeriod.GetDispStr) & ";" & _
                      "分區方式  :" & subSpace(strGroupString) & ";" & _
                      "統計對象  :" & subSpace(strCountString)

    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Private Function Sample(strAreaCode As String, strAreaName As String) As Boolean
  On Error GoTo ChkErr
    Dim rsCD035 As New ADODB.Recordset
      Set rsCD035 = gcnGi.Execute("SELECT CodeNo,Description From CD035")
      While Not rsCD035.EOF
          cnn.Execute "Insert Into SO5510A1 (CustStatusCode,CustStatusName,ServCode,ServArea,intCount) Values (" & _
                      rsCD035("CodeNo") & ",'" & _
                      rsCD035("Description") & "'," & _
                      GetNullString(strAreaCode) & "," & _
                      GetNullString(strAreaName) & ",0)"
          rsCD035.MoveNext
      Wend
      'cnn.Execute ("Delete From SO5510A1 Where len(ServArea)=0 ")
      CloseRecordset rsCD035
      Sample = True
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "Sample")
End Function

Private Sub sample2()
  On Error GoTo ChkErr
    Dim rsSO5510A1 As New ADODB.Recordset
      Set rsSO5510A1 = cnn.Execute("SELECT ServCode,ServArea,count(*) From SO5510A1 Group By ServCode,ServArea")
      While Not rsSO5510A1.EOF
        If Trim(rsSO5510A1("ServArea")) <> "" Then
            cnn.Execute "Insert Into SO5510A2 (ServCode,ServArea) Values ('" & _
                       rsSO5510A1("ServCode") & "','" & rsSO5510A1("ServArea") & "')"
        End If
            rsSO5510A1.MoveNext
      Wend
      CloseRecordset rsSO5510A1
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "sample2")
End Sub

Private Function subPrint() As Boolean
  On Error GoTo ChkErr
  Dim strMDB As String
  'Dim rsSO075 As New ADODB.Recordset
  Dim strMduid  As String
  Dim strGroup As String
  Dim strField As String
  Dim strGroupField As String
  Dim strFaciWhere1 As String
  Dim strFaciWhere2 As String
  Dim strCMFrom As String
  Dim strCMSQL As String
  '建立資料表
  Call CreateTable
    Select Case True
           Case optServCode.Value '服務區
                strField = " SO001.ServCode ASCode,SO001.ServArea ASName "
                strGroupField = " SO001.ServCode,SO001.ServArea "
                Set rsTmp1 = gcnGi.Execute("Select CodeNO,Description From CD002 Where CompCode=" & gilCompCode.GetCodeNo)
           Case optAreaCode.Value '行政區
                strField = " SO014.AreaCode ASCode,So014.AreaName ASName "
                strGroupField = " SO014.AreaCode,So014.AreaName "
                Set rsTmp1 = gcnGi.Execute("Select CodeNO,Description from CD001 Where CompCode=" & gilCompCode.GetCodeNo)
           Case optClctAreaCode.Value '收費區
                strField = " SO001.ClctAreaCode ASCode,SO001.ClctAreaName ASName "
                strGroupField = " SO001.ClctAreaCode,SO001.ClctAreaName "
                Set rsTmp1 = gcnGi.Execute("Select CodeNO,Description From CD040 Where CompCode=" & gilCompCode.GetCodeNo)
    End Select
    
    If chkYes.Value = 1 Then
        strMDB = "SELECT A.CustStatusCode,A.CustStatusName,A.ASCode,A.ASName,A.Mduid,Count(Distinct A.CustId) as intCount From (" & _
                 "SELECT So002.CustStatusCode CustStatusCode,So002.CustStatusName CustStatusName," & strField & ",SO001.Mduid,SO001.CustId From " & strChoose & " And SO001.ChargeType in (1,2) Union All " & _
                 "SELECT So002.CustStatusCode CustStatusCode,So002.CustStatusName CustStatusName," & strField & ",SO001.Mduid,SO001.CustId From " & strChooseNot & " And SO001.MduId in (Select SO001.MduId From " & strChoose & " And SO001.ChargeType=3)) A " & _
                 " Group By A.CustStatusCode,A.CustStatusName,A.ASCode,A.ASName,A.Mduid"
    Else
        strMDB = "SELECT So002.CustStatusCode CustStatusCode,So002.CustStatusName CustStatusName," & strField & _
                ",SO001.Mduid,Count(Distinct SO001.CustId) as intCount From " & _
                 strChoose & " Group By So002.CustStatusCode,So002.CustStatusName," & strGroupField & ",SO001.Mduid"
    End If
             
    If Not subInsertMDB(strMDB, strServArea) Then Exit Function
    SendSQL strMDB
    
    If strTitleChoose <> "" Then strTitleChoose = " And " & strTitleChoose
    strChoose = strChoose & strTitleChoose
    '收費方式
    If chkMduId.Value = True Then
      strMduid = ",SO001.Mduid"
      strGroup = ",SO001.Mduid"
    Else
      strMduid = ",'' as mduid"
      strGroup = ""
    End If
    
    If strChooseSO003 = "" Then
        strCMFrom = "SO003,"
        strCMSQL = "And SO002.CustId = SO003.CustId(+) And SO002.CompCode = SO003.CompCode And SO002.ServiceType = SO003.ServiceType And SO003.CitemCode in (Select CodeNo From CD019 Where RefNo = 1)"
        
    End If
    
    If chkYes.Value = 1 Then
        strMDB = "Select A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid,Count(Distinct A.CustId) As intCount  From(" & _
                 "Select SO003.CMCode TypeCode,SO003.CMName TypeName," & strField & strMduid & " ,SO001.CustId From " & strCMFrom & strChoose & strCMSQL & " And SO001.ChargeType in (1,2) Union All " & _
                 "Select SO003.CMCode TypeCode,SO003.CMName TypeName," & strField & strMduid & " ,SO001.CustId From " & strCMFrom & strChooseNot & strTitleChoose & strCMSQL & " And SO001.MduId in (Select SO001.MduId From " & strChoose & " And SO001.ChargeType=3)) A " & _
                 " Group by A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid"
    Else
        strMDB = "select SO003.CMCode TypeCode,SO003.CMName TypeName," & strField & _
                  strMduid & " ,Count(Distinct SO001.CustId) As intCount  From " & strCMFrom & strChoose & _
                  strCMSQL & " Group by SO003.CMCode,SO003.CMName, " & strGroupField & strGroup
    End If
    
    Call subInsertMDB2(strMDB, strServArea, "SO5510A2")
    SendSQL strMDB
    
    '客戶類別
    If chkYes.Value = 1 Then
        strMDB = "Select A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid,Count(Distinct A.CustId) As intCount  From(" & _
                 "Select SO001.ClassCode1 TypeCode,SO001.ClassName1 TypeName," & strField & strMduid & " ,SO001.CustId From  " & strChoose & " And SO001.ChargeType in (1,2) Union All " & _
                 "Select SO001.ClassCode1 TypeCode,SO001.ClassName1 TypeName," & strField & strMduid & " ,SO001.CustId From  " & strChooseNot & strTitleChoose & " And SO001.MduId in (Select SO001.MduId From " & strChoose & " And SO001.ChargeType=3)) A " & _
                 " Group by A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid "

    Else
        strMDB = "Select SO001.ClassCode1 TypeCode,SO001.ClassName1 TypeName," & strField & _
                  strMduid & " ,Count(Distinct SO001.CustId) As intCount  From  " & strChoose & _
                  " Group by SO001.ClassCode1,SO001.ClassName1, " & strGroupField & strGroup
    End If
    '問題集 2760增加子報表 客戶類別2、3
'    If chkYes.Value = 1 Then
'        strMDB = "Select A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid,Count(Distinct A.CustId) As intCount  From(" & _
'                 "Select SO001.ClassCode2 TypeCode,SO001.ClassName2 TypeName," & strField & strMduid & " ,SO001.CustId From  " & strChoose & " And SO001.ChargeType in (1,2) Union All " & _
'                 "Select SO001.ClassCode2 TypeCode,SO001.ClassName2 TypeName," & strField & strMduid & " ,SO001.CustId From  " & strChooseNot & strTitleChoose & " And SO001.MduId in (Select SO001.MduId From " & strChoose & " And SO001.ChargeType=3)) A " & _
'                 " Group by A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid "
'
'    Else
'        strMDB = "Select SO001.ClassCode2 TypeCode,SO001.ClassName2 TypeName," & strField & _
'                  strMduid & " ,Count(Distinct SO001.CustId) As intCount  From  " & strChoose & _
'                  " Group by SO001.ClassCode2,SO001.ClassName2, " & strGroupField & strGroup
'    End If
    Call subInsertMDB2(strMDB, strServArea, "SO5510A3")
    SendSQL strMDB
    
'    '設備項目
'    If chkYes.Value = 1 Then
'        strMDB = "Select A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid,Count(A.TypeCode) As intCount  From(" & _
'                 "Select SO004.FaciCode TypeCode,SO004.FaciName TypeName," & strField & strMduid & " ,SO001.CustId From SO004, " & strChoose & " And So004.CustId = So001.CustId " & IIf(gilServiceType.GetCodeNo = "", "", " And So004.ServiceType ='" & gilServiceType.GetCodeNo & "' ") & " And SO001.ChargeType in (1,2) And SO004.InstDate is Not Null And SO004.PRDate is Null Union All " & _
'                 "Select SO004.FaciCode TypeCode,SO004.FaciName TypeName," & strField & strMduid & " ,SO001.CustId From SO004, " & strChooseNot & strTitleChoose & " And So004.CustId = So001.CustId " & IIf(gilServiceType.GetCodeNo = "", "", " And So004.ServiceType ='" & gilServiceType.GetCodeNo & "' ") & " And SO001.MduId in (Select SO001.MduId From " & strChoose & " And SO001.ChargeType=3) And SO004.InstDate is Not Null And SO004.PRDate is Null) A " & _
'                 " Group by A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid"
'    Else
'        strMDB = "Select SO004.FaciCode TypeCode,SO004.FaciName TypeName," & strField & _
'                  strMduid & " ,Count(SO004.FaciCode) As intCount  From SO004, " & _
'                  strChoose & " And So004.CustId = So001.CustId " & _
'                  IIf(gilServiceType.GetCodeNo = "", "", " And So004.ServiceType ='" & gilServiceType.GetCodeNo & "' ") & _
'                  " And SO004.InstDate is Not Null And SO004.PRDate is Null" & _
'                  " Group by SO004.FaciCode,SO004.FaciName, " & strGroupField & strGroup
'    End If
'
'    Call subInsertMDB2(strMDB, strServArea, "SO5510A4")
'    SendSQL strMDB, True
'設備項目
    If blnSO004 Then
        strFaciWhere1 = strChoose
        strFaciWhere2 = strChooseNot & strTitleChoose & " And SO001.MduId in (Select SO001.MduId From " & strChoose & " And SO001.ChargeType=3)) A "
    Else
        strFaciWhere1 = "SO004, " & strChoose & " And SO004.CustId = SO001.CustId " & IIf(gilServiceType.GetCodeNo = "", "", " And SO004.ServiceType ='" & gilServiceType.GetCodeNo & "' ")
        strFaciWhere2 = "SO004, " & strChooseNot & strTitleChoose & " And SO004.CustId = SO001.CustId " & IIf(gilServiceType.GetCodeNo = "", "", " And SO004.ServiceType ='" & gilServiceType.GetCodeNo & "' ") & " And SO001.MduId in (Select SO001.MduId From " & strChoose & " And SO001.ChargeType=3) And SO004.InstDate is Not Null And SO004.PRDate is Null ) A "
    End If
    If chkYes.Value = 1 Then
        strMDB = "Select A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid,Count(A.TypeCode) As intCount  From (" & _
                 "Select SO004.FaciCode TypeCode,SO004.FaciName TypeName," & strField & strMduid & " ,SO001.CustId From " & strFaciWhere1 & " And SO001.ChargeType in (1,2) And SO004.InstDate is Not Null And SO004.PRDate is Null Union All " & _
                 "Select SO004.FaciCode TypeCode,SO004.FaciName TypeName," & strField & strMduid & " ,SO001.CustId From " & strFaciWhere2 & _
                 " Group by A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid"

    Else
        strMDB = "Select SO004.FaciCode TypeCode,SO004.FaciName TypeName," & strField & strMduid & " ,Count(SO004.FaciCode) As intCount  From " & strFaciWhere1 & " And SO004.InstDate is Not Null And SO004.PRDate is Null" & _
                  " Group by SO004.FaciCode,SO004.FaciName, " & strGroupField & strGroup
    End If
    Call subInsertMDB2(strMDB, strServArea, "SO5510A4")
    SendSQL strMDB

  'NodeNo
    If blnSO014 Then
        strFaciWhere1 = strChoose
        strFaciWhere2 = strChooseNot & strTitleChoose & " And SO001.MduId in (Select SO001.MduId From " & strChoose & " And SO001.ChargeType=3)) A "
    Else
        strFaciWhere1 = "SO014, " & strChoose & " And SO001.InstAddrNo = SO014.AddrNo And SO001.CompCode = SO014.CompCode "
        strFaciWhere2 = "SO014, " & strChooseNot & strTitleChoose & " And SO001.InstAddrNo = SO014.AddrNo And SO001.CompCode = SO014.CompCode And SO001.MduId in (Select SO001.MduId From " & strChoose & " And SO001.ChargeType=3)) A "
    End If
    If chkYes.Value = 1 Then
        strMDB = "Select A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid,Count(Distinct A.CustId) As intCount  From (" & _
                 "Select SO014.NodeNo TypeCode,SO014.NodeNo TypeName," & strField & strMduid & " ,SO001.CustId From  " & strFaciWhere1 & " And SO001.ChargeType in (1,2) Union All " & _
                 "Select SO014.NodeNo TypeCode,SO014.NodeNo TypeName," & strField & strMduid & " ,SO001.CustId From  " & strFaciWhere2 & _
                 " Group by A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid"
    Else
        strMDB = "Select SO014.NodeNo TypeCode,SO014.NodeNo TypeName," & strField & strMduid & " ,Count(Distinct SO001.CustId) As intCount  From  " & strFaciWhere1 & _
                  " Group by SO014.NodeNo,SO014.NodeNo, " & strGroupField & strGroup
    End If

    Call subInsertMDB2(strMDB, strServArea, "SO5510A5")
    SendSQL strMDB, True
              
  '工期類別
    If chkYes.Value = 1 Then
        strMDB = "Select CD065A.CodeNo TypeCode,CD065A.Description TypeName,A.ASCode,A.ASName,A.Mduid,Count(Distinct A.CustId) As intCount  From (" & _
                 "Select SO014.NodeNo," & strField & strMduid & " ,SO001.CustId From  " & strFaciWhere1 & " And SO001.ChargeType in (1,2) Union All " & _
                 "Select SO014.NodeNo," & strField & strMduid & " ,SO001.CustId From  " & strFaciWhere2 & _
                 ",CD065A Where A.NodeNo=CD065A.NodeNo(+) Group by CD065A.CodeNo,CD065A.Description,A.ASCode,A.ASName,A.Mduid"
    Else
        strMDB = "Select CD065A.CodeNo TypeCode,CD065A.Description TypeName," & strField & strMduid & " ,Count(Distinct SO001.CustId) As intCount From CD065A," & strFaciWhere1 & _
                  " And SO014.NodeNo=CD065A.NodeNo(+) Group by CD065A.CodeNo,CD065A.Description, " & strGroupField & strGroup
    End If
    Call subInsertMDB2(strMDB, strServArea, "SO5510A6")
    SendSQL strMDB
    '收費項目 問題集1014 20040805 Edit by Crystal for Debby 20040729轉RD
    If chkMduId.Value = True Then
      strMduid = ",SO001.Mduid"
      strGroup = ",SO001.Mduid"
    Else
      strMduid = ",'' as mduid"
      strGroup = ""
    End If
    
    If chkYes.Value = 1 Then
        strMDB = "Select A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid,Count(Distinct A.CustId) As intCount  From(" & _
                 "Select SO003.CitemCode TypeCode,SO003.CitemName TypeName," & strField & strMduid & " ,SO001.CustId From " & Left(strChoose, InStr(strChoose, "W") - 1) & IIf(InStr(strChoose, "SO003") = 0, ",SO003 ", "") & _
                 Right(strChoose, Len(strChoose) + 1 - InStr(strChoose, "W")) & " AND SO001.CustId=SO003.CustId  And SO001.ChargeType in (1,2) Union All " & _
                 "Select SO003.CitemCode TypeCode,SO003.CitemName TypeName," & strField & strMduid & " ,SO001.CustId From " & Left(strChoose, InStr(strChoose, "W") - 1) & IIf(InStr(strChoose, "SO003") = 0, ",SO003 ", "") & _
                 Right(strChoose, Len(strChoose) + 1 - InStr(strChoose, "W")) & " AND SO001.CustId=SO003.CustId  And SO001.MduId in (Select SO001.MduId From " & strChoose & " And SO001.ChargeType=3)) A " & _
                 " Group by A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid"
    Else
        strMDB = "select SO003.CitemCode TypeCode,SO003.CitemName TypeName," & strField & _
                  strMduid & " ,Count(Distinct SO001.CustId) As intCount  From " & Left(strChoose, InStr(strChoose, "W") - 1) & IIf(InStr(strChoose, "SO003") = 0, ",SO003 ", "") & _
                  Right(strChoose, Len(strChoose) + 1 - InStr(strChoose, "W")) & " AND SO001.CustId=SO003.CustId " & " Group by SO003.CitemCode,SO003.CitemName, " & strGroupField & strGroup
    End If
    
    Call subInsertMDB2(strMDB, strServArea, "SO5510A7")
    '客戶類別2
    If chkYes.Value = 1 Then
        strMDB = "Select A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid,Count(Distinct A.CustId) As intCount  From(" & _
                 "Select SO001.ClassCode2 TypeCode,SO001.ClassName2 TypeName," & strField & strMduid & " ,SO001.CustId From  " & strChoose & " And SO001.ChargeType in (1,2) And ClassCode2 is not null Union All " & _
                 "Select SO001.ClassCode2 TypeCode,SO001.ClassName2 TypeName," & strField & strMduid & " ,SO001.CustId From  " & strChooseNot & strTitleChoose & " And SO001.MduId in (Select SO001.MduId From " & strChoose & " And SO001.ChargeType=3 And ClassCode2 is not null)) A " & _
                 " Group by A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid "

    Else
        strMDB = "Select SO001.ClassCode2 TypeCode,SO001.ClassName2 TypeName," & strField & _
                  strMduid & " ,Count(Distinct SO001.CustId) As intCount  From  " & strChoose & _
                  " And ClassCode2 is not null Group by SO001.ClassCode2,SO001.ClassName2, " & strGroupField & strGroup
    End If
    Call subInsertMDB2(strMDB, strServArea, "SO5510A8")
    SendSQL strMDB, True
    '客戶類別3
    If chkYes.Value = 1 Then
        strMDB = "Select A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid,Count(Distinct A.CustId) As intCount  From(" & _
                 "Select SO001.ClassCode3 TypeCode,SO001.ClassName3 TypeName," & strField & strMduid & " ,SO001.CustId From  " & strChoose & " And SO001.ChargeType in (1,2) And ClassCode3 is not null Union All " & _
                 "Select SO001.ClassCode3 TypeCode,SO001.ClassName3 TypeName," & strField & strMduid & " ,SO001.CustId From  " & strChooseNot & strTitleChoose & " And SO001.MduId in (Select SO001.MduId From " & strChoose & " And SO001.ChargeType=3 And ClassCode3 is null)) A " & _
                 " Group by A.TypeCode,A.TypeName,A.ASCode,A.ASName,A.Mduid "

    Else
        strMDB = "Select SO001.ClassCode3 TypeCode,SO001.ClassName3 TypeName," & strField & _
                  strMduid & " ,Count(Distinct SO001.CustId) As intCount  From  " & strChoose & _
                  " And ClassCode3 is not null Group by SO001.ClassCode3,SO001.ClassName3, " & strGroupField & strGroup
    End If
    Call subInsertMDB2(strMDB, strServArea, "SO5510A9")
    SendSQL strMDB, True
    subPrint = True
              
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subPrint")
End Function

Private Function subInsertMDB(strSQLQry As String, strServArea As String) As Boolean
    On Error GoTo ChkErr
    Dim strStatusName As String
    Dim rsTmp As New ADODB.Recordset
        cnn.Execute "DELETE FROM SO5510A1"
        If Not GetRS(rsTmp, strSQLQry) Then Exit Function
    On Error GoTo ErrTrans
        cnn.BeginTrans
        While Not rsTmp.EOF
            If InStr(gimStatusCode.GetQueryCode, rsTmp("CustStatusCode")) > 0 Then
                strStatusName = "正常客戶數"
            Else
                strStatusName = "非正常客戶數"
            End If
            cnn.Execute "Insert Into SO5510A1 (CustStatusCode,CustStatusName,ServCode,ServArea,intCount,MduId,StatusName) Values (" & _
                        GetNullString(rsTmp("CustStatusCode")) & "," & _
                        GetNullString(rsTmp("CustStatusName")) & "," & _
                        GetNullString(rsTmp("ASCode")) & "," & _
                        GetNullString(rsTmp("ASName")) & "," & _
                        GetNullString(rsTmp("intCount")) & "," & _
                        GetNullString(rsTmp("MduId")) & ",'" & _
                        strStatusName & "')"
            rsTmp.MoveNext
            DoEvents
        Wend
        If rsTmp.RecordCount > 0 Then
            If chkMduId.Value = 1 Then
               cnn.Execute "Update SO5510A1 set intCount=1 Where Mduid <>''"
            End If
            
            If rsTmp1.RecordCount > 0 And rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                rsTmp1.MoveFirst
                While Not rsTmp1.EOF
                    cnn.Execute "Insert Into SO5510A1 (CustStatusCode,CustStatusName,ServCode,ServArea,intCount,MduId) Values (" & _
                                GetNullString(rsTmp("CustStatusCode")) & "," & _
                                GetNullString(rsTmp("CustStatusName")) & "," & _
                                GetNullString(rsTmp1("CodeNO")) & "," & _
                                GetNullString(rsTmp1("Description")) & "," & _
                                0 & "," & _
                                0 & ")"
                    rsTmp1.MoveNext
                    DoEvents
                Wend
            End If
        End If
        cnn.CommitTrans
    CloseRecordset rsTmp
    subInsertMDB = True
    Exit Function
ErrTrans:
    cnn.RollbackTrans
ChkErr:
    Call ErrSub(Me.Name, "subInsertMDB")
End Function

Private Function subInsertMDB2(strSQLQry As String, strServArea As String, strTable As String) As Boolean
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
        cnn.Execute "DELETE FROM " & strTable
        If Not GetRS(rsTmp, strSQLQry) Then Exit Function
        cnn.BeginTrans
        While Not rsTmp.EOF
            cnn.Execute "Insert Into " & strTable & " (CitemCode,CitemName,ServCode,ServArea,intCount,MduId) Values (" & _
                        GetNullString(rsTmp("TypeCode")) & "," & _
                        GetNullString(rsTmp("TypeName")) & "," & _
                        GetNullString(rsTmp("ASCode")) & "," & _
                        GetNullString(rsTmp("ASName")) & "," & _
                        GetNullString(rsTmp("intCount"), giLongV) & "," & _
                        GetNullString(rsTmp("MduId")) & ")"
            rsTmp.MoveNext
            DoEvents
        Wend
        If chkMduId.Value = 1 Then
           cnn.Execute "Update SO5510A2 set intCount=1 Where Mduid is not null"
        End If
        
        If rsTmp.RecordCount > 0 And rsTmp1.RecordCount > 0 Then
          rsTmp.MoveFirst
          rsTmp1.MoveFirst
          While Not rsTmp1.EOF
            cnn.Execute "Insert Into " & strTable & " (CitemCode,CitemName,ServCode,ServArea,intCount,MduId) Values (" & _
                          GetNullString(rsTmp("TypeCode")) & "," & _
                          GetNullString(rsTmp("TypeName")) & "," & _
                          GetNullString(rsTmp1("CodeNO")) & "," & _
                          GetNullString(rsTmp1("Description")) & "," & _
                          0 & "," & 0 & ")"
            rsTmp1.MoveNext
            DoEvents
          Wend
        End If
        cnn.CommitTrans
        CloseRecordset rsTmp
        subInsertMDB2 = True
    Exit Function
ChkErr:
    cnn.RollbackTrans
    Call ErrSub(Me.Name, "subInsertMDB2-" & strTable)
End Function

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5510A)
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

Private Sub subGil()
  On Error GoTo ChkErr
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimStatusCode, "CodeNo", "Description", "CD035", "客戶狀態代碼", "客戶狀態名稱")
    Call SetgiMulti(gimWipCode3, "CodeNo", "Description", "CD036", "派工狀態代碼", "派工狀態名稱")
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    '問題集2760 增加客戶類別2、3
    Call SetgiMulti(gimClassCode2, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    Call SetgiMulti(gimClassCode3, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    Call SetgiMulti(gimClctAreaCode, "CodeNo", "Description", "CD040", "收費區代碼", "收費區名稱")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱")
    Call SetgiMulti(gimStrtRange, "CodeNo", "Description", "CD017", "街道代碼", "街道名稱")
    Call SetgiMulti(gimSales, "CodeNo", "Description", "CD050", "業務區代碼", "業務區名稱")
    Call SetgiMulti(gimFaciCode, "CodeNo", "Description", "CD022", "設備項目代碼", "設備項目名稱")
    Call SetgiMulti(gimNodeNo, "CompCode", "CodeNo", "CD047", "CompCode", "NodeNo")
    Call SetgiMulti(gimJobPeriod, "CodeNo", "Description", "CD065", "工期代碼", "工期名稱")
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱")
    gimStatusCode.SetQueryCode "1"
    gimWipCode3.SetQueryCode "13"
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  CloseRecordset rsTmp1
  ReleaseCOM frmSO5510A
End Sub

Private Sub gdaClctDate1_GotFocus()
  On Error Resume Next
  If gdaClctDate1.GetValue = "" Then gdaClctDate1.SetValue (RightDate)
End Sub

Private Sub gdaClctDate2_GotFocus()
  On Error Resume Next
  If gdaClctDate1.GetValue = "" Or gdaClctDate2.GetValue = "" Then gdaClctDate2.SetValue gdaClctDate1.GetValue
End Sub

Private Sub gdaClctDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaClctDate1, gdaClctDate2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    GiMultiFilter gimClctAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtRange, , gilCompCode.GetCodeNo
    GiMultiFilter gimSales, , gilCompCode.GetCodeNo
    GiMultiFilter gimNodeNo, gilServiceType.GetCodeNo, gilCompCode.GetCodeNo
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
    GiMultiFilter gimFaciCode, gilServiceType.GetCodeNo
    GiMultiFilter gimNodeNo, gilServiceType.GetCodeNo, gilCompCode.GetCodeNo
    GiMultiFilter gimJobPeriod, gilServiceType.GetCodeNo
    GiMultiFilter gimCitemCode, gilServiceType.GetCodeNo
    gimCitemCode.Filter = gimCitemCode.Filter & IIf(gimCitemCode.Filter = "", " WHERE ", " AND ") & " PeriodFlag = 1"
    gimCitemCode.SetQueryCode RPxx(GetRsValue("SELECT CodeNo FROM CD019 WHERE ServiceType='" & gilServiceType.GetCodeNo & "' AND RefNo=1") & "")
 
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilServiceType_Change")
End Sub

Private Sub gimJobPeriod_GotFocus()
  On Error Resume Next
    vsl1.Value = 100

End Sub

Private Sub gimStatusCode_GotFocus()
        On Error Resume Next
        vsl1.Value = 0
End Sub

Private Sub mskCircuitNo_Validate(Cancel As Boolean)
  If mskCircuitNo.Text = "" Then
    optCircuitNo.Enabled = True
    optAll.Enabled = True
  Else
    optCircuitNo.Enabled = False
    optAll.Enabled = False
  End If
End Sub

Private Sub optCountNotMduid_Click()
  On Error Resume Next
    If optCountNotMduid.Value Then
        chkMduId.Value = 0
        chkYes.Value = 0
    End If
  
End Sub

Private Sub vsl1_Change()
  On Error Resume Next
    If vsl1.Value = 0 Then
        fraMula.Top = 20
    ElseIf vsl1.Value = 100 Then
        fraMula.Top = -(pic2.Height) + 480
    End If
End Sub
Sub CreateTable()
  On Error Resume Next
    cnn.Execute "Drop Table SO5510A8"
    cnn.Execute "Drop Table SO5510A9"
  On Error GoTo ChkErr
    cnn.Execute "Create Table SO5510A8 (CitemCode long,CitemName Text(50)," & _
                "ServArea text(50),intCount Long,MduId Text(50),ServCode Text(50))"
    cnn.Execute "Create Table SO5510A9 (CitemCode long,CitemName Text(50)," & _
                "ServArea text(50),intCount Long,MduId Text(50),ServCode Text(50))"
   Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "CreateTable")
End Sub
