VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO5270A 
   BorderStyle     =   1  '單線固定
   Caption         =   "每月收費單完成率統計[ SO5270A ]"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "SO5270A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.PictureBox pic2 
      Height          =   2955
      Left            =   60
      ScaleHeight     =   2895
      ScaleWidth      =   7515
      TabIndex        =   36
      Top             =   1770
      Width           =   7575
      Begin VB.VScrollBar vsl1 
         Height          =   2835
         LargeChange     =   100
         Left            =   7230
         Max             =   100
         SmallChange     =   100
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.Frame fraMulti 
         BorderStyle     =   0  '沒有框線
         Height          =   3975
         Left            =   30
         TabIndex        =   37
         Top             =   30
         Width           =   7125
         Begin CS_Multi.CSmulti gimCitemCode 
            Height          =   345
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   609
            ButtonCaption   =   "收  費  項  目"
         End
         Begin Gi_Multi.GiMulti gimCMCode 
            Height          =   345
            Left            =   0
            TabIndex        =   40
            Top             =   330
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   609
            ButtonCaption   =   "收  費  方  式"
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin Gi_Multi.GiMulti gimBillType 
            Height          =   345
            Left            =   0
            TabIndex        =   41
            Top             =   690
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   609
            ButtonCaption   =   "單  據  類  別"
            DataType        =   2
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin Gi_Multi.GiMulti gimStatusCode 
            Height          =   345
            Left            =   0
            TabIndex        =   42
            Top             =   1050
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   609
            ButtonCaption   =   "客  戶  狀  態"
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin Gi_Multi.GiMulti gimServCode 
            Height          =   345
            Left            =   0
            TabIndex        =   43
            Top             =   1410
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   609
            ButtonCaption   =   "服     務     區"
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin Gi_Multi.GiMulti gimAreaCode 
            Height          =   345
            Left            =   0
            TabIndex        =   44
            Top             =   1770
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   609
            ButtonCaption   =   "行     政     區"
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin CS_Multi.CSmulti gimStrtCode 
            Height          =   345
            Left            =   0
            TabIndex        =   45
            Top             =   2130
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   609
            ButtonCaption   =   "街  道  範  圍"
         End
         Begin Gi_Multi.GiMulti gimCancel 
            Height          =   345
            Left            =   0
            TabIndex        =   46
            Top             =   2490
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   609
            ButtonCaption   =   "作  廢  原  因"
            Enabled         =   0   'False
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin CS_Multi.CSmulti gimClassCode 
            Height          =   345
            Left            =   30
            TabIndex        =   47
            Top             =   2850
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   609
            ButtonCaption   =   "客  戶  類  別"
         End
         Begin CS_Multi.CSmulti gimMduId 
            Height          =   345
            Left            =   30
            TabIndex        =   48
            Top             =   3210
            Width           =   10035
            _ExtentX        =   17701
            _ExtentY        =   609
            ButtonCaption   =   "大  樓  編  號"
         End
         Begin Gi_Multi.GiMulti gimPayType 
            Height          =   390
            Left            =   30
            TabIndex        =   49
            Top             =   3570
            Width           =   10035
            _ExtentX        =   17701
            _ExtentY        =   688
            ButtonCaption   =   "繳  付  類  別"
            FontSize        =   9
            FontName        =   "新細明體"
         End
      End
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "匯成Excel"
      Height          =   525
      Left            =   3570
      TabIndex        =   35
      Top             =   6510
      Width           =   1245
   End
   Begin VB.CheckBox chkFirst 
      Caption         =   "是否統計首期資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3795
      TabIndex        =   10
      Top             =   1485
      Value           =   1  '核取
      Width           =   3735
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
      Left            =   3795
      TabIndex        =   9
      Top             =   1215
      Value           =   1  '核取
      Width           =   3735
   End
   Begin VB.CheckBox chkCancel 
      Caption         =   "作廢單據是否一併統計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   135
      TabIndex        =   6
      Top             =   1365
      Width           =   3285
   End
   Begin VB.Frame Frame4 
      Caption         =   "網路編號"
      Height          =   735
      Left            =   30
      TabIndex        =   30
      Top             =   4770
      Width           =   7635
      Begin VB.OptionButton optCircuitNo 
         Caption         =   "只印空白網路編號"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   4830
         TabIndex        =   12
         Top             =   240
         Width           =   1845
      End
      Begin VB.OptionButton optAll 
         Caption         =   "全部"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6660
         TabIndex        =   13
         Top             =   270
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.TextBox mskCircuitNo 
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   90
         TabIndex        =   11
         Top             =   240
         Width           =   4605
      End
   End
   Begin Gi_Date.GiDate gdaShouldDate2 
      Height          =   345
      Left            =   2340
      TabIndex        =   1
      Top             =   90
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
   Begin Gi_Date.GiDate gdaShouldDate1 
      Height          =   345
      Left            =   930
      TabIndex        =   0
      Top             =   90
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
      Left            =   90
      TabIndex        =   22
      Top             =   6510
      Width           =   1245
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
      Left            =   6240
      TabIndex        =   24
      Top             =   6510
      Width           =   1245
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
      Left            =   1770
      TabIndex        =   23
      Top             =   6510
      Width           =   1395
   End
   Begin VB.Frame fraPage 
      Caption         =   "分區方式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   855
      Left            =   30
      TabIndex        =   25
      Top             =   5550
      Width           =   7635
      Begin VB.OptionButton optClassCode 
         Caption         =   "客戶類別"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   3480
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optOldClctEn 
         Caption         =   "原收費人員"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   4800
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optClctAreaCode 
         Caption         =   "收費區"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   2160
         TabIndex        =   16
         Top             =   240
         Width           =   915
      End
      Begin VB.OptionButton optClctEn 
         BackColor       =   &H80000004&
         Caption         =   "收費人員"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   3480
         TabIndex        =   18
         Top             =   240
         Width           =   1035
      End
      Begin VB.OptionButton optNo 
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
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   4800
         TabIndex        =   21
         Top             =   600
         Value           =   -1  'True
         Width           =   525
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
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   960
         TabIndex        =   14
         Top             =   240
         Width           =   915
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
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   960
         TabIndex        =   15
         Top             =   600
         Width           =   915
      End
      Begin VB.OptionButton optStrtCode 
         Caption         =   "街道編號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   2160
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   4590
      TabIndex        =   8
      Top             =   450
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
      Left            =   4590
      TabIndex        =   7
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
   Begin Gi_Date.GiDate gdaCreateTime2 
      Height          =   375
      Left            =   2340
      TabIndex        =   3
      Top             =   510
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
   Begin Gi_Date.GiDate gdaCreateTime1 
      Height          =   375
      Left            =   930
      TabIndex        =   2
      Top             =   510
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
      Height          =   375
      Left            =   2340
      TabIndex        =   5
      Top             =   945
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
   Begin Gi_Date.GiDate gdaRealDate1 
      Height          =   375
      Left            =   930
      TabIndex        =   4
      Top             =   930
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
      Left            =   2130
      TabIndex        =   34
      Top             =   1005
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "實收日期"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   90
      TabIndex        =   33
      Top             =   1005
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "出單日期"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   32
      Top             =   600
      Width           =   780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2130
      TabIndex        =   31
      Top             =   600
      Width           =   105
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
      Left            =   3750
      TabIndex        =   29
      Top             =   150
      Width           =   765
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
      Left            =   3780
      TabIndex        =   28
      Top             =   510
      Width           =   780
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   2130
      TabIndex        =   27
      Top             =   150
      Width           =   105
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
      TabIndex        =   26
      Top             =   180
      Width           =   780
   End
End
Attribute VB_Name = "frmSO5270A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table=SO033orSO034 A,SO014 B(,SO002 C)
Option Explicit
Dim strFrom As String
'Dim strST As String
Dim strgroupcode As String
Dim strGroup As String
Dim blnSO014 As Boolean
Dim strUseTable As String
Dim blnExcel As Boolean
Private Sub chkCancel_Click()
  On Error GoTo ChkErr
    If chkCancel.Value = 1 Then
        gimCancel.Enabled = True
    Else
        gimCancel.Enabled = False
        gimCancel.Clear
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "chkCancel_Click")
End Sub

Private Sub cmdExcel_Click()
  On Error GoTo ChkErr
    blnExcel = True
    cmdPrint_Click
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExcel_Click")
End Sub

Private Sub cmdExit_Click() '結束
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdPrevRpt_Click() '上次統計結果
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      Call PreviousRpt(GetPrinterName(5), RptName("SO5270"), Me.Caption)
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click() '列印
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim rsExcel As New ADODB.Recordset
      If Not ChkDTok Then Exit Sub
      If Not IsDataOk Then Exit Sub
      Screen.MousePointer = vbHourglass
        cmdExit.SetFocus
        Me.Enabled = False
          ReadyGoPrint
          Call subChoose
          subCreateView
          CreateTable
          Set rsTmp = cnn.Execute("SELECT Count(*) as intCount From SO5270A2")
          If rsTmp("intCount") = 0 Then
              MsgNoRcd
              SendSQL , , True
          Else
              strSQL = "Select * From SO5270A2"
            If blnExcel Then
                RptToTxt "SO5270AE.rpt", , strSQL, strChooseString, , "Tmp0000.MDB", strGroupName, , , garyGi(19) & "\SO5270"
                If Not Get_RS_From_Txt(garyGi(19), "SO5270.txt", rsExcel) Then blnExcel = False: Exit Sub
'                Dim objCVT As Object
'                Set objCVT = CreateObject("Converter.txt2xls")
'                objCVT.Convert_TXT_2_XLS Environ("Temp") & "\SO5270.txt", , , , , , , True
'                Set objCVT = Nothing
                '#7074 To Add a choice-condition parameter into excel print function by Kin 2015/09/09
                Call UseProperty(rsExcel, "每月未繳費客戶明細表", "第一頁", _
                        , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , strChooseString)
                blnExcel = False
            Else

              Call PrintRpt(GetPrinterName(5), RptName("SO5270"), , Me.Caption, strSQL, strChooseString, , True, "Tmp0000.MDB", , strGroupName, GiPaperLandscape)
            End If
          End If
        Me.Enabled = True
      Screen.MousePointer = vbDefault
      CloseRecordset rsTmp
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean '檢查必要欄位是否有值
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

Private Sub subChoose() '篩選條件
  On Error GoTo ChkErr
  Dim strpagetype As String
  Dim strGroupString As String
  Dim blnSO002 As Boolean
    'strChoose = IIf(strST = "", "", " And nvl(A.STCode,-1) in(" & strST & ",-1)")
    If chkBad.Value = 0 Then
        strChoose = "Nvl(A.STCode,-1) NOT IN (Select CodeNo From CD016 Where RefNo =1) "
    Else
        strChoose = ""
    End If
    strChooseString = ""
    strFrom = ""
    strUseTable = ""
    blnSO002 = False
    blnSO014 = False

  '日期
    If gdaShouldDate1.GetValue <> "" Then Call subAnd("A.ShouldDate >= To_Date('" & gdaShouldDate1.GetValue & "','YYYYMMDD')")
    If gdaShouldDate2.GetValue <> "" Then Call subAnd("A.ShouldDate < To_Date('" & gdaShouldDate2.GetValue & "','YYYYMMDD')+1")
    If gdaCreateTime1.GetValue <> "" Then Call subAnd("A.CreateTime >= To_Date('" & gdaCreateTime1.GetValue & "','YYYYMMDD')")
    If gdaCreateTime2.GetValue <> "" Then Call subAnd("A.CreateTime < To_Date('" & gdaCreateTime2.GetValue & "','YYYYMMDD')+1")

  'GiList
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("A.ServiceType='" & gilServiceType.GetCodeNo & "'")
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("A.CompCode=" & gilCompCode.GetCodeNo)
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("Instr(Nvl(C.ServiceType,'" & gilServiceType.GetCodeNo & "'),'" & gilServiceType.GetCodeNo & "') > 0")
    
  'GiMulti
    'If gimCompCode.GetQryStr <> "" Then Call subAnd("C.CompCode" & gimCompCode.GetQryStr)
    If gimCitemCode.GetQryStr <> "" Then Call subAnd("A.CitemCode " & gimCitemCode.GetQryStr)
    If gimCMCode.GetQryStr <> "" Then Call subAnd("A.CMCode " & gimCMCode.GetQryStr)
    If gimBillType.GetQryStr <> "" Then Call subAnd("SubStr(A.BillNo,7,1) " & gimBillType.GetQryStr)
    If gimStatusCode.GetQryStr <> "" Then Call subAnd("C.CustStatusCode " & gimStatusCode.GetQryStr): blnSO002 = True
    If gimServCode.GetQryStr <> "" Then Call subAnd("A.ServCode " & gimServCode.GetQryStr)
    If gimAreaCode.GetQryStr <> "" Then Call subAnd("A.AreaCode " & gimAreaCode.GetQryStr)
    If gimStrtCode.GetQryStr <> "" Then Call subAnd("A.StrtCode " & gimStrtCode.GetQryStr)
    If gimCancel.GetQryStr <> "" Then Call subAnd("A.CancelCode " & gimCancel.GetQryStr)
    If gimCancel.GetDispStr = "(全選)" Then Call subAnd("A.CancelCode is not Null")
    If chkCancel.Value = 0 Then subAnd ("A.cancelFlag=0")
    '問題2494 2006/06/01 大樓編號移至下方,問題2536修正
    If gimClassCode.GetQryStr <> "" Then Call subAnd("A.ClassCode " & gimClassCode.GetQryStr)
'    If gimMduId.GetQryStr <> "" Then Call subAnd("A.MduId " & gimMduId.GetQryStr)
    '#5683
    If gimPayType.GetQryStr <> "" Then Call subAnd("A.PayKind " & gimPayType.GetQryStr)
    
  '是否統計首期資料
    If chkFirst.Value = 0 Then Call subAnd("A.FirstFlag is Null")
    
  '網路編號
    If mskCircuitNo.Text <> "" Then
        Call subAnd("B.CircuitNo='" & mskCircuitNo.Text & "'")
        blnSO014 = True
    Else
      If optCircuitNo.Value Then
          Call subAnd("B.CircuitNo is null ")
          blnSO014 = True
      End If
    End If
  
  '分區方式
      Select Case True
         Case optServCode.Value
              strGroupName = "Type=1"
              strgroupcode = "CD002.CodeNo"
              strGroup = "CD002.Description"
              strUseTable = strUseTable & ",CD002"
              strChoose = "A.ServCode=CD002.CodeNo(+) And " & strChoose
              strpagetype = "服務區"
              
         Case optAreaCode.Value
              strGroupName = "Type=2"
              strgroupcode = "CD001.CodeNo"
              strGroup = "CD001.Description"
              strUseTable = strUseTable & ",CD001"
              strChoose = "A.AreaCode=CD001.CodeNo(+) And " & strChoose
              strpagetype = "行政區"
              
         Case optClctAreaCode.Value
              strGroupName = "Type=3"
              strgroupcode = "CD040.CodeNo"
              strGroup = "CD040.Description"
              strUseTable = strUseTable & ",CD040"
              strChoose = "A.ClctAreaCode=CD040.CodeNo(+) And " & strChoose
              strpagetype = "收費區"
              
         Case optStrtCode.Value
              strGroupName = "Type=4"
              strgroupcode = "CD017.CodeNo"
              strGroup = "CD017.Description"
              strUseTable = strUseTable & ",CD017"
              strChoose = "A.StrtCode=CD017.CodeNo(+) And " & strChoose
              strpagetype = "街道編號"
              
         Case optClctEn.Value
              strGroupName = "Type=5"
              strgroupcode = "A.ClctEn"
              strGroup = "A.ClctName"
              strpagetype = "收費人員"
         
         Case optOldClctEn.Value
              strGroupName = "Type=7"
              strgroupcode = "A.OldClctEn"
              strGroup = "A.OldClctName"
              strpagetype = "原收費人員"
         
         Case optClassCode.Value
              strGroupName = "Type=8"
              strgroupcode = "CD004.CodeNo"
              strGroup = "CD004.Description"
              strUseTable = strUseTable & ",CD004"
              strChoose = "A.ClassCode=CD004.CodeNo(+) And " & strChoose
              strpagetype = "客戶類別"
         Case optNo.Value
              strGroupName = "Type=6"
              strgroupcode = "''"
              strGroup = "''"
              strpagetype = "無"
      End Select
    
    
   '串 SO014 語法
    If blnSO014 Then
        strUseTable = strUseTable & ",SO014 B"
        strChoose = "A.AddrNo=B.AddrNo And " & strChoose
    End If
      
   '串 SO002 語法
    If blnSO002 Then
        strUseTable = strUseTable & ",SO002 C"
        strChoose = "A.Custid=C.Custid And A.ServiceType=C.ServiceType And " & strChoose
    End If
    
    '問題集2536 Casper 提 2006/06/14 Crystal Edit
    If gimMduId.GetQryStr <> "" Then
        strUseTable = strUseTable & ",SO001"
        Call subAnd("SO001.MduId " & gimMduId.GetQryStr)
        '問題集???? SO033.ServiceType不能等於SO001.ServiceType,因為SO001.ServiceType可以包含CIP For Corey 2006/11/17
        'strChoose = "A.Custid=SO001.Custid And A.ServiceType=SO001.ServiceType And A.CompCode=SO001.CompCode AND " & strChoose
        strChoose = " A.Custid=SO001.Custid And instr(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "')>0 And A.CompCode=SO001.CompCode And " & strChoose
    End If

  '選取條件
    strChooseString = "應收日期: " & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & _
                      "出單日期: " & subSpace(gdaCreateTime1.GetValue(True)) & "~" & subSpace(gdaCreateTime2.GetValue(True)) & ";" & _
                      "服務類別: " & subSpace(gilServiceType.GetDescription) & ";" & _
                      "公司別　: " & subSpace(gilCompCode.GetDescription) & ";" & _
                      "作廢單據是否一併統計:" & IIf(chkCancel.Value = 1, "是", "否") & ";" & _
                      "是否包含呆帳資料統計:" & IIf(chkFirst.Value = 1, "是", "否") & ";" & _
                      "是否統計首期資料:" & IIf(chkBad.Value = 1, "是", "否") & ";" & _
                      "收費項目: " & subSpace(gimCitemCode.GetDispStr) & ";" & _
                      "收費方式: " & subSpace(gimCMCode.GetDispStr) & ";" & _
                      "單據編號: " & subSpace(gimBillType.GetDispStr) & ";" & _
                      "客戶狀態: " & subSpace(gimStatusCode.GetDispStr) & ";" & _
                      "服務區　: " & subSpace(gimServCode.GetDispStr) & ";" & _
                      "行政區　: " & subSpace(gimAreaCode.GetDispStr) & ";" & _
                      "街道範圍: " & subSpace(gimStrtCode.GetDispStr) & ";" & _
                      "作廢原因: " & subSpace(gimCancel.GetDispStr) & ";" & _
                      "綱路編號: " & subSpace(mskCircuitNo.Text) & ";" & _
                      "客戶類別: " & subSpace(gimClassCode.GetDispStr) & ";" & _
                      "大樓名稱: " & subSpace(gimMduId.GetDispStr) & ";" & _
                      "分區方式: " & strpagetype
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) '快速鍵
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5270A)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    If Alfa2 Then Call GetGlobal
    Call subGim
    Call subGil
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description

  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub subGim() 'GiMulti
  On Error GoTo ChkErr
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱")
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱")
    Call SetgiMultiAddItem(gimBillType, "B,I,P,M,T", "收費單,裝機單,停拆移機單,維修單,臨時收費單")
    Call SetgiMulti(gimStatusCode, "CodeNo", "Description", "CD035", "客戶狀態代碼", "客戶狀態名稱")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "街道代碼", "街道名稱")
    Call SetgiMulti(gimCancel, "CodeNo", "Description", "CD051", "作廢原因代碼", "作廢原因名稱")
    '問題 2497 2006/06/01
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    Call SetgiMulti(gimMduId, "MduId", "Name", "SO017", "大樓編號", "大樓名稱")
    Call SetgiMulti(gimPayType, "CodeNo", "Description", "CD112", "代碼", "名稱")
    'Call SetgiMultiAddItem(gimPayType, "0,1", "預付制,現付制", "代碼", "名稱")
    If GetPaynowFlag Then
'        With gimPayType
'         '   .SetDispStr "預付制,現付制"
'            .SetQueryCode "0"
'        End With
        '#5925 繳費類別增加預設值 By Kin 2011/04/12
        With gimPayType
            Select Case Val(GetRsValue("SELECT NVL(PayKindDefault,0) FROM " & GetOwner & "SO041", gcnGi) & "")
                Case 0
                    .SelectAll
                Case 1
                    .SetQueryCode "0"
                Case 2
                    .SetQueryCode "1"
            End Select
        End With
    Else
    
        gimPayType.Clear
        gimPayType.Enabled = False
    End If
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

Private Function subCreateView() As Boolean '建一個 SO033,SO034 View
  On Error Resume Next
      gcnGi.Execute "Drop View " & strViewName
    On Error GoTo ChkErr
    Dim strViewSql As String
        strViewName = GetTmpViewName
        
        strViewSql = "Create View " & strViewName & " as (" & _
                        "SELECT " & strgroupcode & " GroupCode," & strGroup & " GroupName,A.UCCode,A.Custid,A.BillNo,A.ShouldAmt,A.RealAmt,A.RealDate,A.STCode,A.CancelFlag,A.ClctEn,A.ClctName,A.ClassCode " & _
                            "From SO033 A" & strUseTable & " Where " & strChoose & _
                      " Union All " & _
                        "SELECT " & strgroupcode & " GroupCode," & strGroup & " GroupName,A.UCCode,A.Custid,A.BillNo,A.ShouldAmt,A.RealAmt,A.RealDate,A.STCode,A.CancelFlag,A.ClctEn,A.ClctName,A.ClassCode " & _
                            "From SO034 A" & strUseTable & " Where " & strChoose
        If chkCancel.Value = 0 Then
            strViewSql = strViewSql & ")"
        Else
            strViewSql = strViewSql & " Union All " & _
                        "SELECT " & strgroupcode & " GroupCode," & strGroup & " GroupName,A.UCCode,A.Custid,A.BillNo,A.ShouldAmt,A.RealAmt,A.RealDate,A.STCode,A.CancelFlag,A.ClctEn,A.ClctName,A.ClassCode " & _
                            "From SO035 A" & strUseTable & " Where " & strChoose & ")"
        End If
'        strViewSql = "Create View " & strViewName & " as (" & _
'                        "SELECT A.ServCode," & strServName & " ServName,A.AreaCode," & strAreaName & " AreaName,A.StrtCode," & strStrtName & " StrtName,A.UCCode,A.Custid,A.BillNo,A.ShouldAmt,A.RealAmt,A.RealDate,A.STCode,A.CancelFlag,A.ClctEn,A.ClctName " & _
'                            "From SO033 A" & strUseTable & strChoose & _
'                      " Union All " & _
'                        "SELECT A.ServCode," & strServName & " ServName,A.AreaCode," & strAreaName & " AreaName,A.StrtCode," & strStrtName & " StrtName,A.UCCode,A.Custid,A.BillNo,A.ShouldAmt,A.RealAmt,A.RealDate,A.STCode,A.CancelFlag,A.ClctEn,A.ClctName " & _
'                            "From SO034 A" & strUseTable & strChoose & ")"
        gcnGi.Execute strViewSql
        SendSQL strViewSql
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  gcnGi.Execute "Drop View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5270A
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
  Cancel = ChkDate2(gdaRealDate1, gdaRealDate2)
  gdaRealDate2_GotFocus
End Sub

Private Sub gdaShouldDate1_GotFocus()
  On Error Resume Next
  If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetValue (RightDate)
End Sub

Private Sub gdaShouldDate2_GotFocus()
  On Error Resume Next
    If gdaShouldDate1.GetValue = "" Or gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue gdaShouldDate1.GetValue
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaShouldDate1, gdaShouldDate2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimCancel, , gilCompCode.GetCodeNo
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
    'strST = GetCode("CD016", "<>1", gilServiceType.GetCodeNo, , , False)
    Call GiMultiFilter(gimCitemCode, gilServiceType.GetCodeNo)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Sub mskCircuitNo_Validate(Cancel As Boolean) '網路編號
  If mskCircuitNo.Text = "" Then
    optCircuitNo.Enabled = True
    optAll.Enabled = True
  Else
    optCircuitNo.Enabled = False
    optAll.Enabled = False
  End If
End Sub

Private Sub CreateTable()
  On Error Resume Next
    cnn.Execute "Drop Table SO5270A1"
    cnn.Execute "Drop Table So5270A2"
  On Error GoTo ChkErr

    cnn.Execute "Create Table SO5270A1 (GroupCode text(20),GroupName Text(40)," & _
                "ShouldCust Long,ShouldBillNo Long,ShouldAmt Long,ShouldCount Long," & _
                "RealCust Long,RealBillNo Long,RealAmt Long,RealCount Long," & _
                "NeverCust Long,NeverBillNo Long,NeverAmt Long,NeverCount Long," & _
                "WaitCust Long,WaitBillNo Long,WaitAmt Long,WaitCount Long," & _
                "STCust Long,STBillNo Long,STAmt Long,STCount Long," & _
                "CancelCust Long,CancelBillNo Long,CancelAmt Long,CancelCount Long," & _
                "BadCust Long,BadBillNo Long,BadAmt Long,BadCount Long)"
    '應收
    Call subInsertMDB("Select GroupCode,GroupName,Count(Distinct Custid) as CustCount,Count(Distinct BillNo) as BillCount,Sum(ShouldAmt) as SumAmt,Count(*) as intCounts " & _
                          "From " & strViewName & " Group by GroupCode,GroupName", "Should")
    '實收
    strSQL = "Select GroupCode,GroupName,Count(Distinct Custid) as CustCount,Count(Distinct BillNo) as BillCount,Sum(RealAmt) as SumAmt,Count(*) as intCounts " & _
             "From " & strViewName & " Where UCCode is Null And RealDate is Not Null Group by GroupCode,GroupName"
             
'     If gdaRealDate1.GetValue <> "" And gdaRealDate2.GetValue <> "" Then
'        strsql = strsql + " AND RealDate >= To_Date('" & gdaRealDate1.GetValue & "','YYYYMMDD')" & _
'                 " AND RealDate < To_Date('" & gdaRealDate2.GetValue & "','YYYYMMDD')+1 Group by GroupCode,GroupName"
'     Else
'        strsql = strsql + " And RealDate is Not Null Group by GroupCode,GroupName"
'     End If
'    If gdaRealDate2.GetValue <> "" Then Call subAnd("RealDate < To_Date('" & gdaRealDate2.GetValue & "','YYYYMMDD')+1")

    Call subInsertMDB(strSQL, "Real")

    '未收
    Call subInsertMDB("Select GroupCode,GroupName,Count(Distinct Custid) as CustCount,Count(Distinct BillNo) as BillCount,Sum(ShouldAmt) as SumAmt,Count(*) as intCounts " & _
                          "From " & strViewName & _
                      " Where UCCode not in (Select CodeNo From CD013 Where Refno=1) Group by GroupCode,GroupName", "Never")
    
    '待收
    Call subInsertMDB("Select GroupCode,GroupName,Count(Distinct Custid) as CustCount,Count(Distinct BillNo) as BillCount,Sum(ShouldAmt) as SumAmt,Count(*) as intCounts " & _
                          "From " & strViewName & _
                      " Where UCCode in (Select CodeNo From CD013 Where RefNo=1) Group by GroupCode,GroupName", "Wait")
    '短收
    Call subInsertMDB("Select GroupCode,GroupName,Count(Distinct Custid) as CustCount,Count(Distinct BillNo) as BillCount,Sum(RealAmt) as SumAmt,Count(*) as intCounts " & _
                          "From " & strViewName & _
                      " Where STCode in (Select CodeNo From CD016 Where Nvl(RefNo,0) <>1) Group by GroupCode,GroupName", "ST")
    '作廢
    Call subInsertMDB("Select GroupCode,GroupName,Count(Distinct Custid) as CustCount,Count(Distinct BillNo) as BillCount,Sum(ShouldAmt) as SumAmt,Count(*) as intCounts " & _
                          "From " & strViewName & _
                      " Where CancelFlag=1  Group by GroupCode,GroupName", "Cancel")
    '呆帳
    Call subInsertMDB("Select GroupCode,GroupName,Count(Distinct Custid) as CustCount,Count(Distinct BillNo) as BillCount,Sum(ShouldAmt) as SumAmt,Count(*) as intCounts " & _
                          "From " & strViewName & " A,(Select CodeNo From CD016 Where RefNo=1) B " & _
                      "Where A.STCode=B.CodeNo Group by GroupCode,GroupName", "Bad")

    SumMDB
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "CreateTabel")
End Sub

Private Sub subInsertMDB(strSQL As String, strType As String)
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
    Set rsTmp = gcnGi.Execute(strSQL)
    SendSQL strSQL
    cnn.BeginTrans
      While Not rsTmp.EOF
        cnn.Execute "Insert Into SO5270A1 (GroupCode,GroupName," & strType & "Cust," & strType & "BillNo," & strType & "Amt," & strType & "Count " & _
                    ") Values(" & _
                    GetNullString(rsTmp("GroupCode")) & "," & _
                    GetNullString(rsTmp("GroupName")) & "," & _
                    GetNullString(rsTmp("CustCount")) & "," & _
                    GetNullString(rsTmp("BillCount")) & "," & _
                    GetNullString(rsTmp("SumAmt")) & "," & _
                    GetNullString(rsTmp("intCounts")) & ")"
       rsTmp.MoveNext
      Wend
    cnn.CommitTrans
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "subInsertMDB")
End Sub

Private Sub SumMDB()
  On Error GoTo ChkErr
  Dim strSQL As String
    cnn.BeginTrans
        strSQL = "Select GroupCode,GroupName," & _
                 "iif(isnull(sum(ShouldCust)),0,sum(ShouldCust)) as ShouldCust,iif(isnull(sum(ShouldBillNo)) ,0,sum(ShouldBillNo)) as ShouldBillNo,iif(isnull(sum(ShouldAmt)),0,sum(ShouldAmt)) as ShouldAmt,iif(isnull(sum(ShouldCount)),0,sum(ShouldCount)) as ShouldCount," & _
                  "iif(isnull(sum(RealCust)),0,sum(RealCust)) as RealCust,iif(isnull(sum(RealBillNo)),0,sum(RealBillNo)) as RealBillNo,iif(isnull(sum(RealAmt)),0,sum(RealAmt)) as RealAmt,iif(isnull(sum(RealCount)),0,sum(RealCount)) as RealCount," & _
                  "iif(isnull(sum(NeverCust)),0,sum(NeverCust)) as NeverCust,iif(isnull(sum(NeverBillNo)),0,sum(NeverBillNo)) as NeverBillNo,iif(isnull(sum(NeverAmt)),0,sum(NeverAmt)) as NeverAmt,iif(isnull(sum(NeverCount)),0,sum(NeverCount)) as NeverCount," & _
                  "iif(isnull(sum(WaitCust)),0,sum(WaitCust)) as WaitCust,iif(isnull(sum(WaitBillNo)),0,sum(WaitBillNo)) as WaitBillNo,iif(isnull(sum(WaitAmt)),0,sum(WaitAmt)) as WaitAmt,iif(isnull(sum(WaitCount)),0,sum(WaitCount)) as WaitCount," & _
                  "iif(isnull(sum(STCust)),0,sum(STCust)) as STCust,iif(isnull(sum(STBillNo)),0,sum(STBillNo)) as STBillNo,iif(isnull(sum(STAmt)),0,sum(STAmt)) as STAmt,iif(isnull(sum(STCount)),0,sum(STCount)) as STCount," & _
                  "iif(isnull(sum(CancelCust)),0,sum(CancelCust)) as CancelCust,iif(isnull(sum(CancelBillNo)),0,sum(CancelBillNo)) as CancelBillNo,iif(isnull(sum(CancelAmt)),0,sum(CancelAmt)) as CancelAmt,iif(isnull(sum(CancelCount)),0,sum(CancelCount)) as CancelCount," & _
                  "iif(isnull(sum(BadCust)),0,sum(BadCust)) as BadCust,iif(isnull(sum(BadBillNo)),0,sum(BadBillNo)) as BadBillNo,iif(isnull(sum(BadAmt)),0,sum(BadAmt)) as BadAmt,iif(isnull(sum(BadCount)),0,sum(BadCount)) as BadCount " & _
                  "Into SO5270A2 From SO5270A1 Group by GroupCode,GroupName"
        cnn.Execute strSQL
        SendSQL strSQL, True
    cnn.CommitTrans
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "SumMDB")
End Sub




Private Sub vsl1_Change()
  On Error Resume Next
    If vsl1.Value = 0 Then
        fraMulti.Top = 20
    ElseIf vsl1.Value = 100 Then
        fraMulti.Top = -(pic2.Height) + 480
    End If
End Sub
