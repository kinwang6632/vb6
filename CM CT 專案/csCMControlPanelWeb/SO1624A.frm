VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.0#0"; "GiTime.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.16#0"; "csMulti.ocx"
Begin VB.Form frmSO1624A 
   BorderStyle     =   1  '單線固定
   Caption         =   "數位機上盒整批作業 [SO1624A]"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO1624A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11310
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CheckBox chkResvTime 
      Caption         =   "預約時間"
      Height          =   195
      Left            =   90
      TabIndex        =   17
      Top             =   2910
      Width           =   1095
   End
   Begin VB.Frame frmResvTime 
      Enabled         =   0   'False
      Height          =   615
      Left            =   90
      TabIndex        =   83
      Top             =   3030
      Width           =   5325
      Begin MSMask.MaskEdBox mskResvDay 
         Height          =   315
         Left            =   4500
         TabIndex        =   21
         Top             =   210
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "999"
         PromptChar      =   " "
      End
      Begin VB.OptionButton optResvDay 
         Caption         =   "應收日期後           日"
         Height          =   195
         Left            =   3210
         TabIndex        =   20
         Top             =   270
         Width           =   1965
      End
      Begin VB.OptionButton optResvTime 
         Caption         =   "固定時間"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   270
         Value           =   -1  'True
         Width           =   1095
      End
      Begin Gi_Time.GiTime gdtResvTime 
         Height          =   315
         Left            =   1290
         TabIndex        =   19
         Top             =   210
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "F3. 尋找"
      Height          =   360
      Left            =   120
      TabIndex        =   82
      Top             =   7200
      Width           =   1125
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5. 列印..."
      Enabled         =   0   'False
      Height          =   360
      Left            =   1230
      TabIndex        =   81
      Top             =   7200
      Width           =   1125
   End
   Begin VB.Frame fraPPV 
      Height          =   3165
      Left            =   5130
      TabIndex        =   68
      Top             =   4500
      Visible         =   0   'False
      Width           =   5715
      Begin VB.ComboBox cboPPVChoose 
         Height          =   315
         ItemData        =   "SO1624A.frx":0442
         Left            =   1680
         List            =   "SO1624A.frx":044F
         Style           =   2  '單純下拉式
         TabIndex        =   47
         Top             =   2220
         Width           =   1125
      End
      Begin VB.ComboBox cboIPPVChoose 
         Height          =   315
         ItemData        =   "SO1624A.frx":046F
         Left            =   1680
         List            =   "SO1624A.frx":047C
         Style           =   2  '單純下拉式
         TabIndex        =   46
         Top             =   2070
         Width           =   1125
      End
      Begin Gi_Multi.GiMulti gimPPVCustStatusCode 
         Height          =   375
         Left            =   150
         TabIndex        =   42
         Top             =   600
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   661
         ButtonCaption   =   "客戶狀態"
      End
      Begin Gi_Multi.GiMulti gimPPVClassCode 
         Height          =   375
         Left            =   150
         TabIndex        =   43
         Top             =   960
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   661
         ButtonCaption   =   "客戶類別"
      End
      Begin Gi_Multi.GiMulti gimPPVCitemCode 
         Height          =   375
         Left            =   150
         TabIndex        =   44
         Top             =   1320
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   661
         ButtonCaption   =   "收費項目"
      End
      Begin Gi_Multi.GiMulti gimPPVChCode 
         Height          =   375
         Left            =   150
         TabIndex        =   45
         Top             =   1680
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   661
         ButtonCaption   =   "頻道編號"
         Enabled         =   0   'False
      End
      Begin Gi_Date.GiDate gdaPPVShouldDate1 
         Height          =   315
         Left            =   1110
         TabIndex        =   40
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Date.GiDate gdaPPVShouldDate2 
         Height          =   315
         Left            =   2760
         TabIndex        =   41
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblPPV 
         AutoSize        =   -1  'True
         Caption         =   "PPV訂購權選項"
         Height          =   195
         Left            =   210
         TabIndex        =   80
         Top             =   2280
         Width           =   1320
      End
      Begin VB.Label lblIPPV 
         AutoSize        =   -1  'True
         Caption         =   "IPPV訂購權選項"
         Height          =   195
         Left            =   210
         TabIndex        =   71
         Top             =   2130
         Width           =   1380
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2490
         TabIndex        =   70
         Top             =   300
         Width           =   195
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "應收日期"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   210
         TabIndex        =   69
         Top             =   300
         Width           =   780
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   5070
      TabIndex        =   67
      Top             =   3570
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txtSQL 
      Height          =   525
      Left            =   5070
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   64
      Top             =   3570
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.PictureBox pic1 
      Height          =   855
      Left            =   3330
      ScaleHeight     =   795
      ScaleWidth      =   4935
      TabIndex        =   61
      Top             =   3870
      Visible         =   0   'False
      Width           =   5000
      Begin MSComctlLib.ProgressBar pbr1 
         Height          =   375
         Left            =   150
         TabIndex        =   62
         Top             =   330
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblProcess 
         AutoSize        =   -1  'True
         Caption         =   "開機中,請稍侯!!"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   180
         TabIndex        =   63
         Top             =   90
         Width           =   1335
      End
   End
   Begin VB.Frame fraOther 
      ForeColor       =   &H00FF0000&
      Height          =   3165
      Left            =   5460
      TabIndex        =   51
      Top             =   420
      Width           =   5715
      Begin CS_Multi.CSmulti gimStrtCode 
         Height          =   345
         Left            =   120
         TabIndex        =   30
         Top             =   2340
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   609
         ButtonCaption   =   "街道範圍"
      End
      Begin Gi_Multi.GiMulti gimServCode 
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1635
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   661
         ButtonCaption   =   "服務區"
      End
      Begin MSMask.MaskEdBox mskSTB1 
         Height          =   315
         Left            =   1260
         TabIndex        =   22
         Top             =   210
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "AAAAAAAAAAAAAA"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskSTB2 
         Height          =   315
         Left            =   3540
         TabIndex        =   23
         Top             =   210
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "AAAAAAAAAAAAAA"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskSmartCardNo1 
         Height          =   315
         Left            =   1260
         TabIndex        =   24
         Top             =   570
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "AAAAAAAAAAAAAA"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskSmartCardNo2 
         Height          =   315
         Left            =   3540
         TabIndex        =   25
         Top             =   570
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "AAAAAAAAAAAAAA"
         PromptChar      =   "_"
      End
      Begin Gi_Multi.GiMulti gimAreaCode 
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   1995
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   661
         ButtonCaption   =   "行政區"
      End
      Begin CS_Multi.CSmulti gimMduId 
         Height          =   345
         Left            =   120
         TabIndex        =   31
         Top             =   2670
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   609
         ButtonCaption   =   "大樓編號"
      End
      Begin Gi_Multi.GiMulti gimCustStatusCode 
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   930
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   661
         ButtonCaption   =   "客戶狀態"
      End
      Begin Gi_Multi.GiMulti gimClassCode 
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1290
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   661
         ButtonCaption   =   "客戶類別"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "至"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3330
         TabIndex        =   55
         Top             =   660
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "智慧卡"
         Height          =   195
         Left            =   270
         TabIndex        =   54
         Top             =   630
         Width           =   585
      End
      Begin VB.Label lblTo1 
         AutoSize        =   -1  'True
         Caption         =   "至"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3330
         TabIndex        =   53
         Top             =   300
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "STB序號"
         Height          =   195
         Left            =   270
         TabIndex        =   52
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Frame fraStopCh 
      ForeColor       =   &H00FF0000&
      Height          =   2625
      Left            =   5580
      TabIndex        =   57
      Top             =   4500
      Visible         =   0   'False
      Width           =   5805
      Begin Gi_Multi.GiMulti gimUCCode 
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   1785
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   661
         ButtonCaption   =   "未收原因"
      End
      Begin Gi_Multi.GiMulti gimCMCode 
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   2145
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   661
         ButtonCaption   =   "收費方式"
      End
      Begin Gi_Multi.GiMulti gimCitemCode 
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   1020
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   661
         ButtonCaption   =   "收費項目"
      End
      Begin Gi_Multi.GiMulti gimChCode 
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   1410
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   661
         ButtonCaption   =   "頻道編號"
         Enabled         =   0   'False
      End
      Begin Gi_Date.GiDate gdaShouldDate1 
         Height          =   315
         Left            =   1110
         TabIndex        =   32
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Date.GiDate gdaShouldDate2 
         Height          =   315
         Left            =   2760
         TabIndex        =   33
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Date.GiDate gdaRealDate1 
         Height          =   315
         Left            =   1110
         TabIndex        =   34
         Top             =   630
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Date.GiDate gdaRealDate2 
         Height          =   315
         Left            =   2760
         TabIndex        =   35
         Top             =   630
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2490
         TabIndex        =   66
         Top             =   690
         Width           =   195
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "實收日期"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   210
         TabIndex        =   65
         Top             =   690
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "應收日期"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   210
         TabIndex        =   60
         Top             =   300
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2490
         TabIndex        =   59
         Top             =   300
         Width           =   195
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   9960
      TabIndex        =   50
      Top             =   7230
      Width           =   1245
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "執行(&F2)"
      Height          =   360
      Left            =   2640
      TabIndex        =   49
      Top             =   7200
      Width           =   1125
   End
   Begin VB.Frame fraActivate 
      ForeColor       =   &H000000FF&
      Height          =   2370
      Left            =   90
      TabIndex        =   56
      Top             =   420
      Width           =   5325
      Begin VB.ComboBox cboOpenChMode 
         Height          =   315
         ItemData        =   "SO1624A.frx":049C
         Left            =   960
         List            =   "SO1624A.frx":04A9
         Style           =   2  '單純下拉式
         TabIndex        =   10
         Top             =   2190
         Visible         =   0   'False
         Width           =   1905
      End
      Begin prjGiList.GiList gilPgNo 
         Height          =   315
         Left            =   2700
         TabIndex        =   9
         Top             =   1860
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         FldWidth1       =   600
         FldWidth2       =   1600
      End
      Begin VB.CheckBox chkPgNo 
         Caption         =   "一併設定限制級節目欄位"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   1920
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.ComboBox cboCreditMode 
         Height          =   315
         ItemData        =   "SO1624A.frx":04C9
         Left            =   960
         List            =   "SO1624A.frx":04D0
         Style           =   2  '單純下拉式
         TabIndex        =   4
         Top             =   900
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtSendMsg 
         Height          =   735
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   16
         Top             =   3540
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.TextBox txtRecallTel 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   1560
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.ComboBox cboCommand 
         Height          =   315
         ItemData        =   "SO1624A.frx":04E2
         Left            =   960
         List            =   "SO1624A.frx":04E4
         Style           =   2  '單純下拉式
         TabIndex        =   1
         Top             =   240
         Width           =   2265
      End
      Begin prjNumber.GiNumber giniPPVthresh 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   1230
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         MaxLength       =   3
         AutoSelect      =   -1  'True
      End
      Begin Gi_Multi.GiMulti gimOpenTmpCh 
         Height          =   345
         Left            =   120
         TabIndex        =   14
         Top             =   2850
         Visible         =   0   'False
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   609
         ButtonCaption   =   "開試看頻道"
      End
      Begin Gi_Multi.GiMulti gimCloseCh 
         Height          =   345
         Left            =   120
         TabIndex        =   15
         Top             =   3180
         Visible         =   0   'False
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   609
         ButtonCaption   =   "關頻道"
      End
      Begin Gi_Date.GiDate gdaExtendRange 
         Height          =   315
         Left            =   3690
         TabIndex        =   3
         Top             =   570
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjGiList.GiList gilChCode 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   570
         Visible         =   0   'False
         Width           =   2730
         _ExtentX        =   4815
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
         FldWidth1       =   800
         FldWidth2       =   1600
      End
      Begin Gi_Multi.GiMulti gimOpenCh 
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Visible         =   0   'False
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   609
         ButtonCaption   =   "開頻道"
      End
      Begin prjNumber.GiNumber ginCredit 
         Height          =   315
         Left            =   3270
         TabIndex        =   5
         Top             =   900
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         MaxLength       =   3
         AutoSelect      =   -1  'True
      End
      Begin Gi_Date.GiDate gdaOpenChStopDate 
         Height          =   315
         Left            =   4170
         TabIndex        =   12
         Top             =   2190
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Date.GiDate gdaOpenChStartDate 
         Height          =   315
         Left            =   2880
         TabIndex        =   11
         Top             =   2190
         Visible         =   0   'False
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblOpenChMode2 
         AutoSize        =   -1  'True
         Caption         =   "至"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3960
         TabIndex        =   79
         Top             =   2250
         Width           =   195
      End
      Begin VB.Label lblOpenChMode 
         AutoSize        =   -1  'True
         Caption         =   "設定模式"
         Height          =   195
         Left            =   150
         TabIndex        =   78
         Top             =   2220
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblCreditMode 
         AutoSize        =   -1  'True
         Caption         =   "設定模式"
         Height          =   195
         Left            =   150
         TabIndex        =   77
         Top             =   930
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lbliPPVthresh 
         AutoSize        =   -1  'True
         Caption         =   "臨界點數"
         Height          =   195
         Left            =   150
         TabIndex        =   76
         Top             =   1260
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblRecallTel 
         AutoSize        =   -1  'True
         Caption         =   "回撥電話"
         Height          =   195
         Left            =   150
         TabIndex        =   75
         Top             =   1590
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblIPPVCredit 
         AutoSize        =   -1  'True
         Caption         =   "IPPV點數"
         Height          =   195
         Left            =   2430
         TabIndex        =   74
         Top             =   990
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblChCode 
         AutoSize        =   -1  'True
         Caption         =   "頻道編號"
         Height          =   195
         Left            =   150
         TabIndex        =   73
         Top             =   600
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblCommand 
         AutoSize        =   -1  'True
         Caption         =   "執行動作"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   150
         TabIndex        =   72
         Top             =   300
         Width           =   780
      End
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   3465
      Left            =   90
      TabIndex        =   48
      Top             =   3660
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   6112
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   285
      Left            =   1170
      TabIndex        =   0
      Top             =   90
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth1       =   650
      FldWidth2       =   1600
      F5Corresponding =   -1  'True
   End
   Begin VB.Label lblComp 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   180
      TabIndex        =   58
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "frmSO1624A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim strChoose As String
Dim strChooseString As String
Dim strForm As String
Dim txtErrFile As TextStream
Dim strResvTime As String
Dim strChStopDate As String
Dim strUserPremission As String
Dim strViewName As String

Private Sub ExecuteGo()
    On Error GoTo ChkErr
        'If Not OpenData Then Exit Sub
        If Not txtErrFile Is Nothing Then txtErrFile.Close
        If Not OpenFile(txtErrFile, ErrFilePath, True) Then Exit Sub
        pic1.Visible = True
        lblProcess.Caption = "中,請稍候!!"
        strResvTime = gdtResvTime.GetValue(True)
        Select Case Left(cboCommand.Text, 1)
            Case "1"
                lblProcess.Caption = "延展收視期限" & lblProcess.Caption
                Call ExtendRange
            Case 99
                lblProcess.Caption = "ForceTune" & lblProcess.Caption
                'Call ForceTune
            Case 99
                lblProcess.Caption = "傳送訊息" & lblProcess.Caption
                Call SendMsg
            Case "2"
                lblProcess.Caption = "暫停頻道" & lblProcess.Caption
                Call StopCh
            Case Else
                Call ExecuteGo2
        End Select
        pic1.Visible = False

    Exit Sub
ChkErr:
    ErrSub Me.Name, "ExecuteGo"
End Sub

Private Sub ExitGo()
    Unload Me
End Sub

Private Function subChoose() As Boolean
    On Error GoTo ChkErr
    'A: SO001,B.SO004,C:SO014
        strForm = GetOwner & "SO001 A," & GetOwner & "SO002 A2," & GetOwner & "SO004 B," & GetOwner & "CD056 D," & GetOwner & "CD029 E"
        strChoose = "A.CustId = A2.CustId And A.CompCode = A2.CompCode And " & _
            "A2.CustId = B.CustId And A2.CompCode = B.CompCode And A2.ServiceType = B.ServiceType And B.InitPlaceNo = D.CodeNo(+) And B.PgNo = E.CodeNo(+)"
        '公司別
        If gilCompCode.GetCodeNo <> "" Then Call GetJoinSQL("A.CompCode = " & gilCompCode.GetCodeNo, strChoose)
        '機上盒
        If mskSTB1.Text <> "" Then Call GetJoinSQL("B.FaciSNo >= '" & mskSTB1.Text & "'", strChoose)
        If mskSTB2.Text <> "" Then Call GetJoinSQL("B.FaciSNo <= '" & mskSTB2.Text & "'", strChoose)
        '智慧卡
        If mskSmartCardNo1.Text <> "" Then Call GetJoinSQL("B.SmartCardNo >= '" & mskSmartCardNo1.Text & "'", strChoose)
        If mskSmartCardNo2.Text <> "" Then Call GetJoinSQL("B.SmartCardNo <= '" & mskSmartCardNo2.Text & "'", strChoose)
        '客戶狀態
        If gimCustStatusCode.GetQryStr <> "" Then Call GetJoinSQL("A2.CustStatusCode " & gimCustStatusCode.GetQryStr, strChoose)
        '客戶類別
        If gimClassCode.GetQryStr <> "" Then Call GetJoinSQL("A.ClassCode1 " & gimClassCode.GetQryStr, strChoose)
        '服務區
        If gimServCode.GetQryStr <> "" Then Call GetJoinSQL("A.ServCode " & gimServCode.GetQryStr, strChoose)
        '行政區
        If gimAreaCode.GetQryStr <> "" Then Call GetJoinSQL("C.AreaCode " & gimAreaCode.GetQryStr, strChoose)
        '街道
        If gimStrtCode.GetQryStr <> "" Then Call GetJoinSQL("C.StrtCode " & gimStrtCode.GetQryStr, strChoose)
        If gimMduId.GetQryStr <> "" Then
            Call GetJoinSQL("A.MduId " & gimMduId.GetQryStr, strChoose)
        '全部大樓
        ElseIf gimMduId.GetDispStr <> "" Then
            Call GetJoinSQL("A.MduId Is Not Null", strChoose)
        End If
        If strChoose = "" Then
            MsgBox "無選取條件,請選取", vbInformation, gimsgPrompt
            mskSTB1.SetFocus
            Exit Function
        End If
        Call GetJoinSQL("B.PRDate Is Null And B.InstDate Is Not Null And SmartCardNo Is Not Null", strChoose)
        
        If InStr(strChoose, "C.") > 0 Then
            strForm = strForm & "," & GetOwner & "SO014 C"
            Call GetJoinSQL("A.InstAddrNo = C.AddrNo And A.CompCode = C.CompCode ", strChoose, giLeft)
        End If
        strChoose = " Where " & strChoose
        strChooseString = "公司別:" & subSpace(gilCompCode.GetDescription) & ";" & _
                              "STB序號 :" & subSpace(mskSTB1.Text) & "~" & subSpace(mskSTB2.Text) & ";" & _
                              "智慧卡序號 :" & subSpace(mskSmartCardNo1.Text) & "~" & subSpace(mskSmartCardNo1.Text) & ";" & _
                              "客戶類別:" & subSpace(gimClassCode.GetDispStr) & ";" & _
                              "客戶狀態:" & subSpace(gimCustStatusCode.GetDispStr) & ";" & _
                              "服務區　:" & subSpace(gimServCode.GetDispStr) & ";" & _
                              "行政區　:" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                              "街道範圍:" & subSpace(gimStrtCode.GetDispStr) & ";" & _
                              "大樓名稱:" & subSpace(gimMduId.GetDispStr)
        subChoose = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "subChoose"
End Function

Private Function subChoose2() As Boolean
    On Error GoTo ChkErr
        strChoose = " Where A.CustId = B.CustId And A.CustId = C.CustId  "
        strForm = " " & GetOwner & "SO033 A," & GetOwner & "SO001 B," & GetOwner & "SO004 C "
        If gilCompCode.GetCodeNo <> "" Then Call GetJoinSQL("A.CompCode = " & gilCompCode.GetCodeNo, strChoose)
        If gdaShouldDate1.GetValue <> "" Then GetJoinSQL "A.ShouldDate >=" & GetNullString(gdaShouldDate1.GetValue(True), giDateV), strChoose
        If gdaShouldDate2.GetValue <> "" Then GetJoinSQL "A.ShouldDate <" & GetNullString(gdaShouldDate2.GetValue(True), giDateV) & "+1", strChoose
        '93/11/29 加入實收日期的欄位
        If gdaRealDate1.GetValue <> "" Then GetJoinSQL "A.RealDate >=" & GetNullString(gdaRealDate1.GetValue(True), giDateV), strChoose
        If gdaRealDate2.GetValue <> "" Then GetJoinSQL "A.RealDate <" & GetNullString(gdaRealDate2.GetValue(True), giDateV) & "+1", strChoose
        If gimCitemCode.GetQryStr <> "" Then GetJoinSQL "A.CitemCode " & gimCitemCode.GetQryStr, strChoose
        If gimUCCode.GetDispStr <> "" Then
            If gimUCCode.GetQryStr = "" Then
                GetJoinSQL "A.UCCode is not null ", strChoose
            Else
                GetJoinSQL "A.UCCode " & gimUCCode.GetQryStr, strChoose
            End If
        End If
        If gimCMCode.GetQryStr <> "" Then GetJoinSQL "A.CMCode " & gimCMCode.GetQryStr, strChoose
        
        GetJoinSQL " A.UCCode is not null And ShouldAmt>0 And C.FaciCode In (Select CodeNo From " & GetOwner & "CD022 Where RefNo = 3) And C.PRDate Is Null And SmartCardNo is Not Null ", strChoose
        
        strChooseString = "公司別:" & subSpace(gilCompCode.GetDescription) & ";" & _
                              "應收日期:" & subSpace(gdaShouldDate1.GetOriginalValue) & "~" & subSpace(gdaShouldDate2.GetOriginalValue) & ";" & _
                              "實收日期:" & subSpace(gdaRealDate1.GetOriginalValue) & "~" & subSpace(gdaRealDate2.GetOriginalValue) & ";" & _
                              "收費項目:" & subSpace(gimCitemCode.GetDispStr) & ";" & _
                              "頻道編號:" & subSpace(gimChCode.GetDispStr) & ";" & _
                              "未收原因:" & subSpace(gimUCCode.GetDispStr) & ";" & _
                              "收費方式:" & subSpace(gimCMCode.GetDispStr)

        subChoose2 = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "subChoose2"
End Function

Private Function subChoose3() As Boolean
    On Error GoTo ChkErr
        strChoose = " Where A.CustId = B.CustId And A.CompCode = B.CompCode And A.CustId = C.CustId "
        strForm = " " & GetOwner & "SO033 A," & GetOwner & "SO001 B," & GetOwner & "SO004 C "
        If gimPPVCustStatusCode.GetQryStr <> "" Then
            strForm = strForm & ",SO002 D "
            GetJoinSQL "A.CustId= D.CustId And A.CompCode=D.CompCode And A.ServiceType=D.ServiceType", strChoose
            GetJoinSQL "D.CustStatusCode " & gimPPVCustStatusCode.GetQryStr, strChoose
        End If
        If gilCompCode.GetCodeNo <> "" Then Call GetJoinSQL("A.CompCode = " & gilCompCode.GetCodeNo, strChoose)
        If gdaPPVShouldDate1.GetValue <> "" Then GetJoinSQL "A.ShouldDate >=" & GetNullString(gdaPPVShouldDate1.GetValue(True), giDateV), strChoose
        If gdaPPVShouldDate2.GetValue <> "" Then GetJoinSQL "A.ShouldDate <" & GetNullString(gdaPPVShouldDate2.GetValue(True), giDateV) & "+1", strChoose
        If gimPPVClassCode.GetQryStr <> "" Then GetJoinSQL "A.ClassCode " & gimPPVClassCode.GetQryStr, strChoose
        If gimPPVCitemCode.GetQryStr <> "" Then GetJoinSQL "A.CitemCode " & gimPPVCitemCode.GetQryStr, strChoose
        If gimCMCode.GetQryStr <> "" Then GetJoinSQL "A.CMCode " & gimCMCode.GetQryStr, strChoose
        If cboPPVChoose.ListIndex = 1 Then GetJoinSQL "C.PPVright = 1", strChoose
        If cboPPVChoose.ListIndex = 2 Then GetJoinSQL "C.PPVright = 0", strChoose
        If cboIPPVChoose.ListIndex = 1 Then GetJoinSQL "C.IPPVright = 1", strChoose
        If cboIPPVChoose.ListIndex = 2 Then GetJoinSQL "C.IPPVright = 0", strChoose
        If Left(cboCommand.Text, 1) = "E" Or Left(cboCommand.Text, 1) = "F" Then
            GetJoinSQL "C.MasterSlaveCtrlCode is not null ", strChoose
        End If
        GetJoinSQL " A.UCCode is not null And ShouldAmt>0 And C.FaciCode In (Select CodeNo From " & GetOwner & "CD022 Where RefNo = 3) And C.PRDate Is Null And SmartCardNo is Not Null ", strChoose
        
        strChooseString = "公司別:" & subSpace(gilCompCode.GetDescription) & ";" & _
                              "應收日期:" & subSpace(gdaPPVShouldDate1.GetOriginalValue) & "~" & subSpace(gdaPPVShouldDate2.GetOriginalValue) & ";" & _
                              "客戶狀態:" & subSpace(gimPPVCustStatusCode.GetDispStr) & ";" & _
                              "客戶類別:" & subSpace(gimPPVClassCode.GetDispStr) & ";" & _
                              "收費項目:" & subSpace(gimPPVCitemCode.GetDispStr) & ";" & _
                              "頻道編號:" & subSpace(gimPPVChCode.GetDispStr)
        If cboIPPVChoose.Visible Then strChooseString = strChooseString & ";" & "IPPV訂購權選項:" & subSpace(cboIPPVChoose.Text)
        If cboPPVChoose.Visible Then strChooseString = strChooseString & ";" & "PPV訂購權選項:" & subSpace(cboPPVChoose.Text)
        subChoose3 = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "subChoose3"
End Function

Private Function OpenData(Optional blnClear As Boolean = False) As Boolean
    On Error GoTo ChkErr
    Dim strsql As String
        strChoose = ""
        strChooseString = ""
        If Left(cboCommand.Text, 1) = 1 Then
            If Not subChoose Then Exit Function
            If blnClear Then GetJoinSQL " A.CustID = -1 And A.CompCode = -1", strChoose
            strsql = "Select A.CompCode,A.CompName,A.CustId,A.CustName,A.ClassName1,A2.CustStatusName,A.InstAddress,B.SmartCardNo,B.FaciSNo," & _
                "D.Description InitPlaceName,B.ModelName,B.IPPVright,B.PPVright,E.Description PgName,B.IPPVthresh,B.ReFaciSNo,B.STBAutoCB,B.CallBack From " & strForm & strChoose & " Order By A.CustId "
        ElseIf Left(cboCommand.Text, 1) = 2 Then
            If Not subChoose2 Then Exit Function
            If blnClear Then GetJoinSQL " A.BillNo = '' And A.Item = -1 ", strChoose
            strsql = "Select A.CompCode ,A.ShouldDate ,A.BillNO ,A.ShouldAmt ,A.CitemCode , A.CitemName , A.CustId ,B.CustName,B.Tel1,B.InstAddress,Count(*) STBCount From " & strForm & strChoose & " Group By A.CompCode ,A.ShouldDate ,A.BillNO ,A.ShouldAmt ,A.CitemCode, A.CitemName ,A.CustId ,B.CustName,B.Tel1,B.InstAddress"
        Else
            If Not subChoose3 Then Exit Function
            If blnClear Then GetJoinSQL " A.BillNo = '' And A.Item = -1 ", strChoose
            strsql = "Select A.CompCode ,A.ShouldDate ,A.BillNO ,A.ShouldAmt ,A.CitemCode , A.CitemName , A.CustId ,B.CustName,B.Tel1,B.InstAddress,Count(*) STBCount From " & strForm & strChoose & " Group By A.CompCode ,A.ShouldDate ,A.BillNO ,A.ShouldAmt ,A.CitemCode, A.CitemName ,A.CustId ,B.CustName,B.Tel1,B.InstAddress"
        End If
        If blnClear Then
            If Not GetRS(rs, strsql, gcnGi) Then Exit Function
        Else
            strViewName = GetTmpViewName
            gcnGi.Execute "Create View " & GetOwner & strViewName & " As " & strsql
            If Not GetRS(rs, "Select * From " & GetOwner & strViewName, gcnGi) Then Exit Function
        End If
        
        txtSQL = strsql
        Set ggrData.Recordset = rs
        ggrData.Refresh
        If rs.RecordCount = 0 Then
            cmdExecute.Enabled = False
            If Not blnClear Then MsgBox gimsgNoRcd: Exit Function
        Else
            cmdExecute.Enabled = True
        End If
        cmdPrint.Enabled = cmdExecute.Enabled
        OpenData = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "OpenData"
End Function

Private Function IsDataOk() As Boolean
    On Error GoTo ChkErr
        Select Case Left(cboCommand.Text, 1)
            Case 1
            Case 2
                If Not MustExist(gdaShouldDate1, 1, "應收日期起始日") Then Exit Function
                If Not MustExist(gdaShouldDate2, 1, "應收日期截止日") Then Exit Function
                If Not MustExist(gimChCode, 5, "頻道編號") Then Exit Function
            Case Else
                If Not MustExist(gdaPPVShouldDate1, 1, "應收日期起始日") Then Exit Function
                If Not MustExist(gdaPPVShouldDate2, 1, "應收日期截止日") Then Exit Function
        End Select
        IsDataOk = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "IsDataOk"
End Function

Private Function IsDataOk2() As Boolean
    On Error GoTo ChkErr
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        If chkResvTime.Value = 1 Then
            If optResvTime Then
                If Not MustExist(gdtResvTime, 1, "預約時間") Then Exit Function
                If CDate(gdtResvTime.GetValue(True)) <= RightNow Then
                    MsgBox "預約時間需大於現在時間", vbCritical, gimsgPrompt
                    gdtResvTime.SetValue DateAdd("h", 1, RightNow)
                    gdtResvTime.SetFocus
                    Exit Function
                End If
            Else
                If Not MustExist(mskResvDay, , "應收日期後幾天") Then Exit Function
            End If
            If MsgBox("有設定預約時間,程式將不會等待CA Command Gateway 回應,確定使用預約功能??", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then Exit Function
        End If
        
        Select Case Left(cboCommand.Text, 1)
            Case 1
                If Not MustExist(gilChCode, 2, "頻道編號") Then Exit Function
                If Not MustExist(gdaExtendRange, 1, "收視期限日") Then Exit Function
            Case 99
                'Debug.Print optForceTune
            Case 99
                If Not MustExist(txtSendMsg, , "訊息內容") Then Exit Function
            Case 2
                If Not MustExist(gdaShouldDate1, 1, "應收日期起始日") Then Exit Function
                If Not MustExist(gdaShouldDate2, 1, "應收日期截止日") Then Exit Function
                If Not MustExist(gimChCode, 5, "頻道編號") Then Exit Function
            Case Else
                If Not MustExist(gdaPPVShouldDate1, 1, "應收日期起始日") Then Exit Function
                If Not MustExist(gdaPPVShouldDate2, 1, "應收日期截止日") Then Exit Function
                Select Case Left(cboCommand.Text, 1)
                    Case "G"
                        If Not MustExist(gimOpenCh, 5, "開頻道") Then Exit Function
                        Select Case cboOpenChMode.ListIndex
                            Case 1
                                If Not MustExist(gdaOpenChStopDate, 1, "頻道截止日") Then Exit Function
                                If CDate(gdaOpenChStopDate.GetValue(True)) < RightDate Then
                                    gdaOpenChStopDate.SetValue RightDate
                                    MsgBox "頻道截止日需大於今天", vbExclamation, gimsgPrompt
                                    Exit Function
                                End If
                            Case 2
                                If Not MustExist(gdaOpenChStartDate, 1, "頻道起始日") Then Exit Function
                                If Not MustExist(gdaOpenChStopDate, 1, "頻道截止日") Then Exit Function
                                If CDate(gdaOpenChStartDate.GetValue(True)) < RightDate Then
                                    gdaOpenChStartDate.SetValue RightDate
                                    MsgBox "頻道起始日需大於今天", vbExclamation, gimsgPrompt
                                    Exit Function
                                End If
                                
                                If CDate(gdaOpenChStartDate.GetValue(True)) > CDate(gdaOpenChStopDate.GetValue(True)) Then
                                    gdaOpenChStopDate.SetValue gdaOpenChStartDate.GetValue(True)
                                    MsgBox "頻道截止日需大於頻道起始日", vbExclamation, gimsgPrompt
                                    Exit Function
                                End If
                        End Select
                    Case "H"
                        If Not MustExist(gimOpenTmpCh, 5, "開試看頻道") Then Exit Function
                    Case "I"
                        If Not MustExist(gimCloseCh, 5, "關頻道") Then Exit Function
                    Case 3
                        If Not MustExist(ginCredit, , "IPPV 點數") Then Exit Function
                    Case 6
                        If Not MustExist(giniPPVthresh, , "IPPV 臨界點數") Then Exit Function
                    Case 8
                        If Not MustExist(txtRecallTel, , "STB 回撥電話") Then Exit Function
                End Select
        End Select
        IsDataOk2 = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "IsDataOk2"
End Function

Private Sub cboCommand_Click()
    On Error Resume Next
    Dim obj As New Collection
    Dim intLoop As Integer
        'fraActivate.Visible = False
        With obj
            .Add lblChCode
            .Add gilChCode
            .Add gdaExtendRange
            .Add lblCreditMode
            .Add cboCreditMode
            .Add lblIPPVCredit
            .Add ginCredit
            .Add lbliPPVthresh
            .Add giniPPVthresh
            .Add lblRecallTel
            .Add txtRecallTel
            .Add chkPgNo
            .Add gilPgNo
            .Add lblOpenChMode
            .Add cboOpenChMode
            .Add gdaOpenChStartDate
            .Add lblOpenChMode2
            .Add gdaOpenChStopDate
            .Add gimOpenCh
            .Add gimOpenTmpCh
            .Add gimCloseCh
        End With
        For intLoop = 1 To obj.Count
            obj.Item(intLoop).Visible = False
        Next
        mskResvDay.Enabled = True
        optResvDay.Enabled = True
        
        Select Case Left(cboCommand.Text, 1)
            Case "1"
                FramePosition fraOther
                Call OpenData(True)
                Call subGrd
                lblChCode.Visible = True
                gilChCode.Visible = True
                gdaExtendRange.Visible = True
                mskResvDay.Enabled = False
                optResvDay.Enabled = False
            Case "2"
                FramePosition fraStopCh
                Call OpenData(True)
                Call subGrd2
            Case "3"
                lblIPPVCredit.Visible = True
                ginCredit.Visible = True
                lblCreditMode.Visible = True
                cboCreditMode.Visible = True
                cboCreditMode.ListIndex = 0
                ClickPPVChoose
            Case "6"
                giniPPVthresh.Visible = True
                lbliPPVthresh.Visible = True
                ClickPPVChoose
            Case "8"
                txtRecallTel.Visible = True
                lblRecallTel.Visible = True
                ClickPPVChoose
            Case "C", "D"
                gilPgNo.Visible = True
                chkPgNo.Visible = True
                ClickPPVChoose
            Case "G"
                lblOpenChMode.Visible = True
                cboOpenChMode.Visible = True
                gdaOpenChStartDate.Visible = True
                lblOpenChMode2.Visible = True
                gdaOpenChStopDate.Visible = True
                gimOpenCh.Visible = True
                cboOpenChMode.ListIndex = 0
                ClickPPVChoose
            Case "H"
                gimOpenTmpCh.Visible = True
                ClickPPVChoose
            Case "I"
                gimCloseCh.Visible = True
                ClickPPVChoose
            Case Else
                ClickPPVChoose
        End Select
        For intLoop = 1 To obj.Count
            If obj(intLoop).Visible Then
                obj(intLoop).Top = cboCommand.Top + cboCommand.Height + 90
                If TypeOf obj(intLoop) Is Label Or TypeOf obj(intLoop) Is CheckBox Then obj(intLoop).Top = obj(intLoop).Top + 30
                If obj(intLoop).Name = "gimOpenCh" Then gimOpenCh.Top = cboOpenChMode.Top + cboOpenChMode.Height + 90
            End If
            'DoEvents
        Next
        Set obj = Nothing
        'fraActivate.Visible = True
End Sub

Private Sub cboOpenChMode_Click()
    On Error Resume Next
        gdaOpenChStartDate.SetValue ""
        gdaOpenChStopDate.SetValue ""
        Select Case cboOpenChMode.ListIndex
            Case 0
                gdaOpenChStartDate.Enabled = False
                gdaOpenChStopDate.Enabled = False
            Case 1
                gdaOpenChStartDate.Enabled = False
                gdaOpenChStopDate.Enabled = True
            Case 2
                gdaOpenChStartDate.Enabled = True
                gdaOpenChStopDate.Enabled = True
                gdaOpenChStartDate.SetValue RightDate
                gdaOpenChStopDate.SetValue strChStopDate
        End Select
End Sub

Private Sub chkPgNo_Click()
    gilPgNo.Enabled = chkPgNo = 1
End Sub

Private Sub chkResvTime_Click()
    On Error Resume Next
        frmResvTime.Enabled = (chkResvTime.Value = 1)
End Sub

Private Sub cmdExecute_Click()
    On Error Resume Next
        If Not IsDataOk2 Then Exit Sub
        If MsgBox("確定要執行批次 [" & Mid(cboCommand.Text, 3) & "] ??", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then Exit Sub
        Me.Enabled = False
        If gcnGi2.State = adStateClosed Then gcnGi2.Open gcnGi.ConnectionString
        Call ExecuteGo
        Me.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ChkErr
        SendSQL txtSQL, True
        ReadyGoPrint
        Select Case cboCommand.ListIndex
            Case 0
                PrintRpt2 , "SO1624A1.rpt", , "STB 資料一覽表(設備) [SO1624A1]", "SELECT * From " & GetOwner & strViewName & " V", strChooseString, , True
            Case Else
                PrintRpt2 , "SO1624A2.rpt", , "STB 資料一覽表(收費) [SO1624A2]", "SELECT * From " & GetOwner & strViewName & " V", strChooseString, , True
        End Select
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdPrint_Click"
End Sub

Private Sub cmdQuery_Click()
    On Error Resume Next
        If Not IsDataOk Then Exit Sub
        Call OpenData
End Sub

Private Sub Command1_Click()
    Dim obj As Object
    For Each obj In Me
        If TypeOf obj Is OptionButton Then
            Select Case Mid(obj.Caption, 2, 1)
                Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J"
                    'Debug.Print obj.Name & " " & obj.Caption
                    Debug.Print Mid(obj.Caption, 2)
            End Select
        End If
    Next
End Sub

Private Sub Form_Activate()
    On Error Resume Next
'        Me.Visible = False
        If cboCommand.ListCount = 0 Then
            MsgBox "無可執行的命令!!", vbExclamation, gimsgPrompt
            Unload Me: Exit Sub
        End If
        Me.Visible = True
End Sub

Private Sub Form_Load()
    On Error Resume Next
        Call InitData
        Call subCommandList
        Call subGrd
        Call subGil
        Call subGim
        Call DefaultValue
        Me.Visible = False
End Sub

Private Sub InitData()
    On Error Resume Next
        'If Not OpenFile(txtErrFile, ErrFilePath, True) Then Exit Sub
        pic1.Move (Me.Width - pic1.Width) / 2, (Me.Height - pic1.Height) / 2
        giniPPVthresh.Text = GetSystemParaItem("iPPVthresh", gilCompCode.GetCodeNo, "", "SO042")
        fraActivate.Height = 2500
End Sub

Private Sub DefaultValue()
    On Error Resume Next
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        cboCommand.ListIndex = 0

End Sub

Private Function subCommandList() As Boolean
    On Error GoTo ChkErr
        With cboCommand
            If GetPremission("SO16541") Then .AddItem "1.延展收視期限"
            If GetPremission("SO16542") Then .AddItem "2.暫停頻道"
            If GetPremission("SO16543") Then .AddItem "3.下載/設定IPPV點數"
            If GetPremission("SO16544") Then .AddItem "4.取消 IPPV 訂購權"
            If GetPremission("SO16545") Then .AddItem "5.啟動 IPPV 訂購權"
            If GetPremission("SO16546") Then .AddItem "6.設定臨界點數"
            If GetPremission("SO16547") Then .AddItem "7.清除 IPPV 收視記錄"
            If GetPremission("SO16548") Then .AddItem "8.設定 STB 回撥電話"
            If GetPremission("SO16549") Then .AddItem "9.取消 STB 自動回撥"
            'If GetPremission("SO1654A") Then .AddItem "A.啟動 STB 自動回撥"
            If GetPremission("SO1654B") Then .AddItem "B.STB 立即回撥"
            If GetPremission("SO1654C") Then .AddItem "C.取消 PPV 訂購權"
            If GetPremission("SO1654D") Then .AddItem "D.啟動 PPV 訂購權"
            If GetPremission("SO1654E") Then .AddItem "E.臨時驗證子機"
            If GetPremission("SO1654F") Then .AddItem "F.取消母子機驗證"
            If GetPremission("SO1654G") Then .AddItem "G.開頻道"
            If GetPremission("SO1654H") Then .AddItem "H.開試看頻道"
            If GetPremission("SO1654I") Then .AddItem "I.關頻道"
            '恢復頻道,Force Tune,傳送訊息
        End With
        
        subCommandList = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "subCommandList"
End Function

Private Function GetPremission(strMID As String) As Boolean
    On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
        If Val(garyGi(4)) = 0 Then GetPremission = True: Exit Function
        If strUserPremission = "" Then
            If Not GetRS(rs, "Select Mid From " & GetOwner & "So029 Where Mid like 'SO1654%' And Group" & garyGi(4) & " = 1") Then Exit Function
            If Not rs.EOF Then strUserPremission = UCase(rs.GetString(, , ",", ",", ""))
        End If
        GetPremission = InStr(1, strUserPremission, UCase(strMID)) > 0
    Exit Function
ChkErr:
    ErrSub Me.Name, "GetPremission"
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ChkErr
        If Shift = 0 Then
            Select Case KeyCode
                Case vbKeyF2
                    If cmdExecute.Enabled Then cmdExecute.Value = True
                Case vbKeyF3
                    If cmdQuery.Enabled Then cmdQuery.Value = True
                Case vbKeyEscape
                    If cmdExit.Enabled Then cmdExit.Value = True
            End Select
        ElseIf Shift = 2 Then
            If KeyCode = vbKeyF Then
                txtSQL.Move 0, 0, Me.Width, Me.Height / 2
                txtSQL.Visible = Not txtSQL.Visible
            End If
        End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub subGil()
    On Error GoTo ChkErr
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 20, GetCompCodeFilter(gilCompCode, garyGi(15))
        SetgiList gilChCode, "CodeNo", "Description", "CD024"
        SetgiList gilPgNo, "CodeNO", "Description", "CD029"
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub subGim()
    On Error GoTo ChkErr
        SetgiMulti gimCustStatusCode, "CodeNo", "Description", "CD035", "代碼", "名稱"
        SetgiMulti gimClassCode, "CodeNo", "Description", "CD004", "代碼", "名稱"
        SetgiMulti gimServCode, "CodeNo", "Description", "CD002", "代碼", "名稱"
        SetgiMulti gimAreaCode, "CodeNo", "Description", "CD001", "代碼", "名稱"
        SetgiMulti gimStrtCode, "CodeNo", "Description", "CD017", "代碼", "名稱"
        SetgiMulti gimMduId, "MduId", "Name", "SO017", "代碼", "名稱"
        
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱", , True
        SetgiMulti gimChCode, "CodeNo", "Description", "CD024", "代碼", "名稱", , True
        SetgiMulti gimUCCode, "CodeNo", "Description", "CD013", "代碼", "名稱", , True
        SetgiMulti gimCMCode, "CodeNo", "Description", "CD031", "代碼", "名稱", , True
        
        SetgiMulti gimOpenCh, "CodeNo", "Description", "CD024", "代碼", "名稱", , True
        SetgiMulti gimOpenTmpCh, "CodeNo", "Description", "CD024", "代碼", "名稱", "Where ChanceDays>0 ", True
        SetgiMulti gimCloseCh, "CodeNo", "Description", "CD024", "代碼", "名稱", , True
    
        SetgiMulti gimPPVCustStatusCode, "CodeNo", "Description", "CD035", "代碼", "名稱"
        SetgiMulti gimPPVClassCode, "CodeNo", "Description", "CD004", "代碼", "名稱"
        SetgiMulti gimPPVCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱", , True
        SetgiMulti gimPPVChCode, "CodeNo", "Description", "CD024", "代碼", "名稱", , True
    
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGrd()
    On Error Resume Next
    Dim mFlds As New prjGiGridR.GiGridFlds
        With mFlds
            .Add "CompName", , , , , "公司別"
            .Add "CustId", , , , , "客戶編號"
            .Add "CustName", , , , , "客戶名稱      "
            .Add "CustStatusName", , , , , "客戶狀態"
            '.Add "ClassName1", , , , , "客戶類別"
            .Add "SmartCardNo", , , , , "智慧卡序號"
            .Add "FaciSNo", , , , , "機上盒序號     "
            .Add "InstAddress", , , , , "裝機地址" & Space(40)
            .Add "InitPlaceName", , , , , "裝置點"
            .Add "ModelName", , , , , "設備型號"
            .Add "IPPVright", , , , , "IPPV訂購權"
            .Add "PPVright", , , , , "PPV訂購權"
            .Add "PgName", , , , , "節目等級"
            .Add "IPPVthresh", , , , , "IPPV臨界點數", vbRightJustify
            .Add "ReFaciSNo", , , , , "母機序號    "
            .Add "STBAutoCB", , , , , "自動回撥"
            .Add "CallBack", , , , , "回撥頻率"
        End With
        ggrData.AllFields = mFlds
        ggrData.SetHead
End Sub

Private Sub subGrd2()
    On Error Resume Next
    Dim mFlds As New prjGiGridR.GiGridFlds
        With mFlds
            .Add "ShouldDate", giControlTypeDate, , , , "應收日期", vbLeftJustify
            .Add "BillNO", , , , , "單據編號           ", vbLeftJustify
            .Add "ShouldAmt", , , , , "出單金額", vbRightJustify
            .Add "CitemName", , , , , "收費項目       ", vbLeftJustify
            .Add "CustId", , , , , "客戶編號      ", vbLeftJustify
            .Add "CustName", , , , , "客戶名稱", vbLeftJustify
            .Add "Tel1", , , , , "電話            ", vbLeftJustify
            .Add "InstAddress", , , , , "裝機地址" & Space(40), vbLeftJustify
        End With
        ggrData.AllFields = mFlds
        ggrData.SetHead
End Sub

'其他的命令
Private Function ExecuteGo2() As Boolean
    On Error GoTo ChkErr
    Dim strcmdStatus As String, strErrorCode As String, strErrorMsg As String, strRowId As String
    Dim strUpdTime As String
    Dim lngCount As Long, lngErrCount As Long
    Dim rsTmp As New ADODB.Recordset
    Dim blnFlag As Boolean
    Dim strCommandCaption As String
    Dim obj As Object
    Dim strChStr As String
    Dim strExtendRangeStopDate As String
    Dim rsTmp2 As New ADODB.Recordset
    Dim strMsgBack As String
    Dim rsChoose As New ADODB.Recordset
    Dim strStart As String, strStop As String
    Dim strResvTime As String
        strUpdTime = Format(RightNow, "yyyy/mm/dd hh:mm:ss")
        strCommandCaption = Mid(cboCommand.Text, 3)
        Set NagraCommand.uProgressBar = pbr1
        lblProcess.Caption = strCommandCaption & lblProcess.Caption
        rs.MoveFirst
        Do While Not rs.EOF
            If Not GetRS(rsTmp, "Select SeqNO,FaciSNo ,SmartCardNo,MasterSlaveCtrlCode From " & GetOwner & "SO004 Where CustId =" & rs("CustId") & " And FaciCode In (Select CodeNo From " & GetOwner & "CD022 Where RefNo = 3) And PRDate Is Null And SmartCardNo is Not Null") Then Exit Function
            blnFlag = False
            strResvTime = GetResvTime(rs)
            gcnGi.BeginTrans
            Select Case Left(cboCommand.Text, 1)
                Case "G"     '開頻道
                    blnFlag = GetOpenChChoose(rs("CustId"), rsTmp("FaciSNo"), rsChoose, strStart, strStop)
                    If blnFlag And rsChoose.RecordCount > 0 Then blnFlag = SetOpenCh(strRowId, rs("CompCode"), rs("CustId"), rsTmp("SmartCardNo") & "", strResvTime, rsTmp("FaciSNo") & "", strCommandCaption, "", rsTmp("FaciSNO"), strStart, strStop, True, rsChoose, "", strErrorCode, strErrorMsg)
                Case "H"     '開試看頻道
                    blnFlag = GetOpenTmpChChoose(rs("CustId"), rsTmp("FaciSNo"), rsChoose, strStart, strStop)
                    If blnFlag And rsChoose.RecordCount > 0 Then blnFlag = SetOpenCh(strRowId, rs("CompCode"), rs("CustId"), rsTmp("SmartCardNo") & "", strResvTime, rsTmp("FaciSNo") & "", strCommandCaption, "", rsTmp("FaciSNO"), strStart, strStop, True, rsChoose, "", strErrorCode, strErrorMsg, , True)
                Case "I"     '關頻道
                    blnFlag = SetCloseCh(strRowId, rs("CompCode"), rs("CustId"), rsTmp("SmartCardNo"), strResvTime, rsTmp("FaciSNo"), strCommandCaption, rsTmp("FaciSNo"), Replace(gimCloseCh.GetQueryCode, "'", ""), True, "", strErrorCode, strErrorMsg, strMsgBack)
                Case "3"      '下載/設定IPPV點數
                    blnFlag = STBCreDold(strRowId, rs("CompCode"), rs("CustId"), rsTmp("SmartCardNo"), strResvTime, rsTmp("FaciSNo"), strCommandCaption, strErrorCode, strErrorMsg, IIf(cboCreditMode.ListIndex = 0, "04", "01"), ginCredit.Value, , Format(RightDate, "YYYY/MM/DD"), IPPVCallBack(rsTmp("SeqNo")))
                Case "4"      '取消IPPV訂購權
                    blnFlag = SetDisiPPVright(strRowId, rs("CompCode"), rs("CustId"), rsTmp("SmartCardNo"), strResvTime, rsTmp("FaciSNo"), strCommandCaption, rsTmp("FaciSNo"), strErrorCode, strErrorMsg, strMsgBack)
                Case "5"        '啟動IPPV訂購權
                    blnFlag = EnIPPVright(strRowId, rs("CompCode"), rs("CustId"), rsTmp("SmartCardNo"), strResvTime, rsTmp("FaciSNo"), strCommandCaption, rsTmp("FaciSNo"), strErrorCode, strErrorMsg, strMsgBack)
                Case "6"       '設定 IPPV 臨界點數
                    blnFlag = SetCreditThrd(strRowId, rs("CompCode"), rs("CustId"), rsTmp("SmartCardNo"), strResvTime, rsTmp("FaciSNo"), strCommandCaption, strErrorCode, strErrorMsg, giniPPVthresh.Value, rsTmp("FaciSNO"), strMsgBack)
                Case "7"       '清除 IPPV 收視記錄
                    blnFlag = SetiPPVPrgClear(strRowId, rs("CompCode"), rs("CustId"), rsTmp("SmartCardNo"), strResvTime, rsTmp("FaciSNo"), strCommandCaption, strErrorCode, strErrorMsg, Format(strUpdTime, "yyyy/mm/dd"), Format(strUpdTime, "yyyy/mm/dd"), strMsgBack)
                Case "8"     '設定 STB 回撥電話
                    blnFlag = SetAuthenticTel(strRowId, rs("CompCode"), rs("CustId"), rsTmp("SmartCardNo"), strResvTime, rsTmp("FaciSNo"), strCommandCaption, strErrorCode, strErrorMsg, Trim(txtRecallTel.Text), , , rsTmp("FaciSNO"), strMsgBack)
                Case "9"  '取消 STB 自動回撥
                    blnFlag = SetSTBDisAutoCall(strRowId, rs("CompCode"), rs("CustId"), rsTmp("SmartCardNo"), strResvTime, rsTmp("FaciSNo"), strCommandCaption, rsTmp("FaciSNO"), strErrorCode, strErrorMsg, strMsgBack)
                Case "A"    '啟動 STB 自動回撥(  不用)
                    'blnFlag = SetSTBEnAutoCall(strRowId, rs("CompCode"), rs("CustId"), rsTmp("SmartCardNo"), strresvtime, rsTmp("FaciSNo"), strCommandCaption, strErrorCode, strErrorMsg, , strMsgBack)
                Case "B"     'STB 立即回撥
                    blnFlag = RightCall(strRowId, rs("CompCode"), rs("CustId"), rsTmp("SmartCardNo"), strResvTime, rsTmp("FaciSNo"), strCommandCaption, strErrorCode, strErrorMsg, strMsgBack)
                Case "C"      '取消 PPV 訂購權
                    blnFlag = SetDisPPVright(strRowId, rs("CompCode"), rs("CustId"), rsTmp("SmartCardNo"), strResvTime, rsTmp("FaciSNo"), strCommandCaption, strErrorCode, strErrorMsg, rsTmp("FaciSNO"), strMsgBack)
                    '一併設定限制級節目欄位
                    If blnFlag And chkPgNo.Value = 1 Then
                        blnFlag = ExecuteCommand("Update " & GetOwner & "SO004 SET PgNo=" & GetNullString(gilPgNo.GetCodeNo) & " Where SeqNo = '" & rsTmp("SeqNo") & "'")
                    End If
                Case "D"         '啟動 PPV 訂購權
                    blnFlag = EnPPVright(strRowId, rs("CompCode"), rs("CustId"), rsTmp("SmartCardNo"), strResvTime, rsTmp("FaciSNo"), strCommandCaption, rsTmp("FaciSNO"), gilPgNo.GetCodeNo, strErrorCode, strErrorMsg, strMsgBack)
                    '一併設定限制級節目欄位
                Case "E"     '臨時驗證子機
                    blnFlag = GetRS(rsTmp2, "Select ValidateDuration2 From " & GetOwner & "CD066 Where CodeNo = " & rsTmp("MasterSlaveCtrlCode"))
                    If blnFlag Then blnFlag = SetMDSetOnTime(strRowId, rs("CompCode"), rs("CustId"), rsTmp("SmartCardNo"), strResvTime, rsTmp("FaciSNo"), strCommandCaption, rsTmp2("ValidateDuration2"), strErrorCode, strErrorMsg, strMsgBack)
                Case "F"        '取消母子機驗證
                    blnFlag = GetRS(rsTmp2, "Select ValidateDuration2 From " & GetOwner & "CD066 Where CodeNo = " & rsTmp("MasterSlaveCtrlCode"))
                    If blnFlag Then blnFlag = SetMDSetCancel(strRowId, rs("CompCode"), rs("CustId"), rsTmp("SmartCardNo"), strResvTime, rsTmp("FaciSNo"), strCommandCaption, rsTmp("FaciNO"), strErrorCode, strErrorMsg, strMsgBack)
                Case 99         '傳送訊息
                    blnFlag = SetSendMsg(rs("CompCode"), rs("CustId"), rs("FaciSNo"), rs("SmartCardNo"), _
                            txtSendMsg, strUpdTime, strcmdStatus, strRowId, strErrorCode, strErrorMsg, strResvTime)
                Case "1"     '延展期限
                    strChStr = gilChCode.GetCodeNo
                    strExtendRangeStopDate = gdaExtendRange.GetValue(True)
                    blnFlag = SetExtendRange(rs("CompCode"), rs("CustId"), rs("FaciSNo"), rs("SmartCardNo"), _
                        strExtendRangeStopDate, strChStr, strExtendRangeStopDate, strUpdTime, strcmdStatus, _
                        strRowId, strErrorCode, strErrorMsg, strResvTime, , pbr1)
                Case "2"
                    StopCh
            End Select
            If blnFlag Then
                If strcmdStatus = "C" Or strcmdStatus = "" Then
                    '正常
                    lngCount = lngCount + 1
                    If strRowId <> "" Then Call ExecuteCommand("Delete From " & gLogInID & "Send_Nagra Where RowId = '" & strRowId & "'", gcnGi2)
                    gcnGi.CommitTrans
                ElseIf strcmdStatus = "E" Then
                    '錯誤
                    Call InsertToErr2(rs, strErrorCode, strErrorMsg)
                    lngErrCount = lngErrCount + 1
                    gcnGi.RollbackTrans
                End If
            Else
                '錯誤
                Call InsertToErr2(rs, strErrorCode, strErrorMsg)
                lngErrCount = lngErrCount + 1
                gcnGi.RollbackTrans
            End If
            DoEvents
            rs.MoveNext
        Loop
        ExecuteGo2 = True
        Set NagraCommand.uProgressBar = Nothing
        '如有錯誤筆數, 呼叫notepad 顯示
        MsgBox "批次" & strCommandCaption & "完成 !! " & vbCrLf & "" & _
                    "   成功:" & lngCount & " 筆!!" & vbCrLf & _
                    "   失敗: " & lngErrCount & " 筆!!"
        If lngErrCount > 0 Then Call Shell("NotePad " & ErrFilePath, vbNormalFocus)
        Exit Function
    Exit Function
ChkErr:
    ErrSub Me.Name, "ExecuteGo2"
End Function

Private Function GetResvTime(rs As ADODB.Recordset) As String
    On Error Resume Next
        If chkResvTime.Value = 0 Then Exit Function
        If optResvTime Then
            GetResvTime = gdtResvTime.GetValue(True)
        Else
            GetResvTime = rs("ShouldDate") + mskResvDay
            If Err.Number <> 0 Then GetResvTime = ""
        End If

End Function

'暫停頻道 93/08/18
Private Function StopCh() As Boolean
    On Error GoTo ChkErr
    Dim strcmdStatus As String, strErrorCode As String, strErrorMsg As String, strRowId As String
    Dim strUpdTime As String
    Dim lngCount As Long, lngErrCount As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strChCode As String
        strUpdTime = Format(RightNow, "yyyy/mm/dd hh:mm:ss")
        rs.MoveFirst
        Do While Not rs.EOF
            If rs("STBCount") = 1 Then
                If Not GetRS(rsTmp, "Select CodeNo From " & GetOwner & "CD019A Where CitemCode= " & rs("CitemCode")) Then Exit Function
                If Not rsTmp.EOF Then
                    strChCode = rsTmp.GetString(, , , "','")
                    strChCode = "'" & Left(strChCode, Len(strChCode) - 2)
                    If Not GetRS(rsTmp, "Select FaciSNo ,SmartCardNo From " & GetOwner & "SO004 Where CustId =" & rs("CustId") & " And FaciCode In (Select CodeNo From " & GetOwner & "CD022 Where RefNo = 3) And PRDate Is Null And SmartCardNo is Not Null") Then Exit Function
                    strResvTime = GetResvTime(rs)
                    
                    If SetStopCh(rs("CompCode"), rs("CustId"), rsTmp("FaciSNo"), rsTmp("SmartCardNo"), _
                         "", strChCode, "", strUpdTime, strcmdStatus, _
                        strRowId, strErrorCode, strErrorMsg, strResvTime, , pbr1, rs("BillNo") & "") Then
                        
                        If strcmdStatus = "C" Then
                            '正常
                            lngCount = lngCount + 1
                            Call ExecuteCommand("Delete From " & gLogInID & "Send_Nagra Where RowId = '" & strRowId & "'", gcnGi2)
                        ElseIf strcmdStatus = "E" Then
                            '錯誤
                            Call InsertToErr2(rs, strErrorCode, strErrorMsg)
                            lngErrCount = lngErrCount + 1
                        End If
                    Else
                        '錯誤
                        Call InsertToErr2(rs, strErrorCode, strErrorMsg)
                        lngErrCount = lngErrCount + 1
                    End If
                End If
            Else
                Call InsertToErr2(rs, strErrorCode, strErrorMsg, True)
                lngErrCount = lngErrCount + 1
            End If
            DoEvents
            rs.MoveNext
        Loop
        StopCh = True
        '如有錯誤筆數, 呼叫notepad 顯示
        MsgBox "批次暫停頻道完成 !! " & vbCrLf & "" & _
                    "   成功:" & lngCount & " 筆!!" & vbCrLf & _
                    "   失敗: " & lngErrCount & " 筆!!"
        If lngErrCount > 0 Then Call Shell("NotePad " & ErrFilePath, vbNormalFocus)
        Exit Function
    Exit Function
ChkErr:
    ErrSub Me.Name, "STOPCh"
End Function

'延展期限
Private Function ExtendRange() As Boolean
    On Error GoTo ChkErr
    Dim strcmdStatus As String, strErrorCode As String, strErrorMsg As String, strRowId As String
    Dim strUpdTime As String
    Dim strChStr As String, strExtendRangeStopDate As String
    Dim lngCount As Long, lngErrCount As Long

        strChStr = gilChCode.GetCodeNo
        strExtendRangeStopDate = gdaExtendRange.GetValue(True)
        strUpdTime = RightNow
        rs.MoveFirst
        Do While Not rs.EOF
            If SetExtendRange(rs("CompCode"), rs("CustId"), rs("FaciSNo"), rs("SmartCardNo"), _
                strExtendRangeStopDate, strChStr, strExtendRangeStopDate, strUpdTime, strcmdStatus, _
                strRowId, strErrorCode, strErrorMsg, strResvTime, , pbr1) Then
                If strcmdStatus = "C" Then
                    '正常
                    lngCount = lngCount + 1
                    Call ExecuteCommand("Delete From " & gLogInID & "Send_Nagra Where RowId = '" & strRowId & "'", gcnGi2)
                ElseIf strcmdStatus = "E" Then
                    '錯誤
                    Call InsertToErr(rs("CustId"), rs("CustName"), rs("SmartCardNo"), rs("FaciSNo"), strErrorCode, strErrorMsg)
                    lngErrCount = lngErrCount + 1
                End If
            Else
                '錯誤
                Call InsertToErr(rs("CustId"), rs("CustName"), rs("SmartCardNo"), rs("FaciSNo"), strErrorCode, strErrorMsg)
                lngErrCount = lngErrCount + 1
            End If
            DoEvents
            rs.MoveNext
        Loop
        ExtendRange = True
        '如有錯誤筆數, 呼叫notepad 顯示
        MsgBox "批次期限延展完成 !! " & vbCrLf & "" & _
                    "   成功:" & lngCount & " 筆!!" & vbCrLf & _
                    "   失敗: " & lngErrCount & " 筆!!"
        If lngErrCount > 0 Then Call Shell("NotePad " & ErrFilePath, vbNormalFocus)
        Exit Function
    Exit Function
ChkErr:
    ErrSub Me.Name, "ExtendRange"
End Function

'傳送訊息
Private Function SendMsg() As Boolean
    On Error GoTo ChkErr
    Dim strcmdStatus As String, strErrorCode As String, strErrorMsg As String, strRowId As String
    Dim strUpdTime As String
    Dim lngCount As Long, lngErrCount As Long

        strUpdTime = Format(RightNow, "yyyy/mm/dd hh:mm:ss")
        rs.MoveFirst
        Do While Not rs.EOF
            If SetSendMsg(rs("CompCode"), rs("CustId"), rs("FaciSNo"), rs("SmartCardNo"), _
                txtSendMsg, strUpdTime, strcmdStatus, strRowId, strErrorCode, strErrorMsg, strResvTime) Then
                If strcmdStatus = "C" Then
                    '正常
                    lngCount = lngCount + 1
                    Call ExecuteCommand("Delete From " & gLogInID & "Send_Nagra Where RowId = '" & strRowId & "'", gcnGi2)
                ElseIf strcmdStatus = "E" Then
                    '錯誤
                    Call InsertToErr(rs("CustId"), rs("CustName"), rs("SmartCardNo"), rs("FaciSNo"), strErrorCode, strErrorMsg)
                    lngErrCount = lngErrCount + 1
                End If
            Else
                '錯誤
                Call InsertToErr(rs("CustId"), rs("CustName"), rs("SmartCardNo"), rs("FaciSNo"), strErrorCode, strErrorMsg)
                lngErrCount = lngErrCount + 1
            End If
            DoEvents
            rs.MoveNext
        Loop
        SendMsg = True
        '如有錯誤筆數, 呼叫notepad 顯示
        MsgBox "批次傳送訊息完成 !! " & vbCrLf & "" & _
                    "   成功:" & lngCount & " 筆!!" & vbCrLf & _
                    "   失敗: " & lngErrCount & " 筆!!"
        If lngErrCount > 0 Then Call Shell("NotePad " & ErrFilePath, vbNormalFocus)
        Exit Function
    Exit Function
ChkErr:
    ErrSub Me.Name, "SendMsg"
End Function

'ForceTune
Private Function ForceTune() As Boolean
    On Error GoTo ChkErr
    Dim strcmdStatus As String, strErrorCode As String, strErrorMsg As String, strRowId As String
    Dim strUpdTime As String
    Dim lngCount As Long, lngErrCount As Long

        strUpdTime = Format(RightNow, "yyyy/mm/dd hh:mm:ss")
        rs.MoveFirst
        Do While Not rs.EOF
            If SetForceTune(rs("CompCode"), rs("CustId"), rs("FaciSNo"), rs("SmartCardNo"), _
                txtSendMsg, strUpdTime, strcmdStatus, strRowId, strErrorCode, strErrorMsg, strResvTime) Then
                If strcmdStatus = "C" Then
                    '正常
                    lngCount = lngCount + 1
                    Call ExecuteCommand("Delete From " & gLogInID & "Send_Nagra Where RowId = '" & strRowId & "'", gcnGi2)
                ElseIf strcmdStatus = "E" Then
                    '錯誤
                    Call InsertToErr(rs("CustId"), rs("CustName"), rs("SmartCardNo"), rs("FaciSNo"), strErrorCode, strErrorMsg)
                    lngErrCount = lngErrCount + 1
                End If
            Else
                '錯誤
                Call InsertToErr(rs("CustId"), rs("CustName"), rs("SmartCardNo"), rs("FaciSNo"), strErrorCode, strErrorMsg)
                lngErrCount = lngErrCount + 1
            End If
            DoEvents
            rs.MoveNext
        Loop
        ForceTune = True
        '如有錯誤筆數, 呼叫notepad 顯示
        MsgBox "批次Force Tune完成 !! " & vbCrLf & "" & _
                    "   成功:" & lngCount & " 筆!!" & vbCrLf & _
                    "   失敗: " & lngErrCount & " 筆!!"
        If lngErrCount > 0 Then Call Shell("NotePad " & ErrFilePath, vbNormalFocus)
        Exit Function
    Exit Function
ChkErr:
    ErrSub Me.Name, "ForceTune"
End Function

Private Sub InsertToErr(lngCustId As Long, strCustName As String, _
    strICCNo As String, strSTB As String, _
    strErrorCode As String, strErrorMsg As String)
    On Error Resume Next
        
        txtErrFile.WriteLine "客戶編號:" & lngCustId & ",客戶名稱:" & strCustName & ",智慧卡編號:" & _
                strICCNo & ",數位機上盒:" & strSTB & ",錯誤代碼:" & strErrorCode & ", 錯誤名稱:" & strErrorMsg
End Sub

Private Sub InsertToErr2(rs As ADODB.Recordset, _
    strErrorCode As String, strErrorMsg As String, Optional bln2Count As Boolean)
    On Error Resume Next
        txtErrFile.WriteLine "應收日期:" & rs("ShouldDate") & ",單據編號:" & rs("BillNO") & ",出單金額:" & _
                 rs("ShouldAmt") & ",客戶項目:" & rs("CitemName") & ",客戶編號:" & rs("CustId") & ",客戶名稱:" & rs("CustName") & ",電話:" & _
                 rs("Tel1") & ",裝機地址:" & rs("InstAddress") & ", 錯誤代碼:" & strErrorCode & ", 錯誤名稱:" & IIf(bln2Count, "該客戶有" & rs("STBCount") & "台數位機上盒,請自行處理!!", strErrorMsg)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        gcnGi.Execute "Drop View " & GetOwner & strViewName
        txtErrFile.Close
        Set txtErrFile = Nothing
        Call CloseRecordset(rs)
End Sub

Private Function ErrFilePath() As String
    On Error Resume Next
        ErrFilePath = IIf(Right(gErrLogPath, 1) = "\", "", "\") & "SO1624A_BatchErr.txt"
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        ReleaseCOM Me

End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error Resume Next
        Select Case giFld.FieldName
            Case "IPPVright", "PPVright"
                If Value = "1" Then
                    Value = "有"
                Else
                    Value = "無"
                End If
            Case "STBAutoCB"
                If Value = "1" Then
                    Value = "是"
                Else
                    Value = "否"
                End If
        End Select

End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If gilCompCode.GetCodeNo = "" Then Exit Sub
        garyGi(16) = gilCompCode.GetOwner
        garyGi(9) = gilCompCode.GetCodeNo
        'garyGi(16) = gilCompCode.GetCodeNo
        gLogInID = GetSystemParaItem("LogInID", gilCompCode.GetCodeNo, "", "SO041", , True) & ""
        If gLogInID <> "" Then gLogInID = gLogInID & "."
        strChStopDate = GetRsValue("Select CADefDate From " & GetOwner & "So041 Where CompCode = " & gilCompCode.GetCodeNo) & ""
        Call subGil
        Call subGim
        Call GiMultiFilter(gimServCode, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimAreaCode, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimStrtCode, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimMduId, , gilCompCode.GetCodeNo)
    
End Sub

Private Sub gimCitemCode_Change()
    Call SetChannel(gimCitemCode, gimChCode)
End Sub

Private Sub SetChannel(objCitem As Object, objCh As Object)
    Dim rsCD019A As New ADODB.Recordset
    Dim strQuery As String
    Dim strsql As String
    If objCitem.GetQueryCode <> "" Then
        strsql = "Select DISTINCT CodeNo,Description From " & GetOwner & "CD019A Where CitemCode in (" & objCitem.GetQueryCode & ")"
        If Not GetRS(rsCD019A, strsql, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub

        Do While Not rsCD019A.EOF
            strQuery = strQuery & "," & rsCD019A("CodeNo") & ""
            rsCD019A.MoveNext
        Loop
        strQuery = Mid(strQuery, 2)
        objCh.SetQueryCode (strQuery)
    Else
        objCh.Clear
    End If

End Sub

Private Sub gimPPVCitemCode_Change()
    Call SetChannel(gimPPVCitemCode, gimPPVChCode)

End Sub

Private Sub ClickPPVChoose()
        FramePosition fraPPV
        lblIPPV.Visible = False
        cboIPPVChoose.Visible = False
        lblPPV.Visible = False
        cboPPVChoose.Visible = False
        Select Case Left(cboCommand.Text, 1)
            Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B"
                lblIPPV.Visible = True
                cboIPPVChoose.Visible = True
            Case "C", "D", "E", "F"
                lblPPV.Visible = True
                cboPPVChoose.Visible = True
                lblPPV.Top = lblIPPV.Top
                cboPPVChoose.Top = cboIPPVChoose.Top
        End Select
        cboPPVChoose.ListIndex = 0
        cboIPPVChoose.ListIndex = 0
        Call OpenData(True)
        Call subGrd2
End Sub

Private Sub FramePosition(objFrame As Object)
    On Error Resume Next
        fraOther.Visible = False
        fraStopCh.Visible = False
        fraPPV.Visible = False
        objFrame.Visible = True
        objFrame.Move fraActivate.Left + fraActivate.Width + 30, fraActivate.Top, fraOther.Width, fraOther.Height
        
End Sub

Private Function GetOpenChChoose(ByVal lngCustId As Long, ByVal strFaciSNo As String, _
    ByRef rsChoose As ADODB.Recordset, ByRef strStart As String, ByRef strStop As String) As Boolean
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim varX As Variant
    Dim intLoop As Integer
        strStart = ""
        strStop = ""

        Select Case cboOpenChMode.ListIndex
            Case 0
                If Not GetRS(rs, "Select A.ChCode,B.Description ChName,To_Date(To_Char(A.SetTime,'YYYYMMDD'),'YYYYMMDD') AuthorStartDate,To_Date(To_Char(A.DueDate,'YYYYMMDD'),'YYYYMMDD') AuthorStopDate From " & GetOwner & "SO005 A," & GetOwner & "CD024 B Where A.ChCode=B.CodeNo And A.CustId = " & lngCustId & " And A.CvtSNo= '" & strFaciSNo & "' And A.ChCode In (" & gimOpenCh.GetQueryCode & ")") Then Exit Function
                strStart = Format(rs("AuthorStartDate") & "", "yyyy/mm/dd")
                strStop = Format(rs("AuthorStopDate") & "", "yyyy/mm/dd")
            Case 1
                strStart = RightDate
                strStop = gdaOpenChStopDate.GetValue(True)
                If Not GetRS(rs, "Select A.ChCode,B.Description ChName,To_Date('" & strStart & "','yyyy/mm/dd') AuthorStartDate, To_Date('" & strStop & "','yyyy/mm/dd') AuthorStopDate From " & GetOwner & "SO005 A," & GetOwner & "CD024 B Where A.ChCode=B.CodeNo And A.CustId = " & lngCustId & " And A.CvtSNo= '" & strFaciSNo & "' And A.ChCode In (" & gimOpenCh.GetQueryCode & ")") Then Exit Function
            Case 2
                strStart = gdaOpenChStartDate.GetValue(True)
                strStop = gdaOpenChStopDate.GetValue(True)
                If Not GetRS(rs, "Select CodeNo ChCode,Description ChName,To_Date('" & strStart & "','yyyy/mm/dd') AuthorStartDate, To_Date('" & strStop & "','yyyy/mm/dd') AuthorStopDate From " & GetOwner & "CD024 Where CodeNo In (" & gimOpenCh.GetQueryCode & ")") Then Exit Function
        End Select
        
        Set rsChoose = rs.Clone
        Call CloseRecordset(rs)
        GetOpenChChoose = True
End Function

Private Function GetOpenTmpChChoose(ByVal lngCustId As Long, ByVal strFaciSNo As String, _
    ByRef rsChoose As ADODB.Recordset, ByRef strStart As String, ByRef strStop As String) As Boolean
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim varX As Variant
    Dim intLoop As Integer
        strStart = Format(RightDate, "yyyy/mm/dd")
        If Not GetRS(rs, "Select A.CodeNo ChCode,A.Description ChName,To_Date('" & strStart & "','yyyy/mm/dd') AuthorStartDate, To_Date('" & strStart & "','yyyy/mm/dd') + A.ChanceDays AuthorStopDate From " & GetOwner & "CD024 A Where A.CodeNo In (" & gimOpenTmpCh.GetQueryCode & ") And A.CodeNo Not In (Select ChCode From " & GetOwner & "SO005 Where CustId = " & lngCustId & " And CvtSNo = '" & strFaciSNo & "')") Then Exit Function
        If rs.RecordCount > 0 Then strStop = Format(rs("AuthorStopDate") & "", "yyyy/mm/dd")
        
        Set rsChoose = rs.Clone
        Call CloseRecordset(rs)
        GetOpenTmpChChoose = True
End Function

' 判斷 有/無 回傳機制
Private Function IPPVCallBack(strSeqNO As String) As Boolean
  On Error GoTo ChkErr
    Dim lngModelCode As Long
    Dim blnIPPVCallBack As Boolean
    
    blnIPPVCallBack = GetSystemParaItem("IPPVCallBack", gilCompCode.GetCodeNo, "", "SO041", , True) = 1
    
    If blnIPPVCallBack Then
        lngModelCode = Val(GetRsValue("SELECT MODELCODE FROM " & GetOwner & "SO004 WHERE SEQNO='" & strSeqNO & "'") & "")
        
        If lngModelCode > 0 Then
            blnIPPVCallBack = GetRsValue("SELECT FUNCFLAG01 FROM " & GetOwner & "CD043 WHERE MODELCODE=" & lngModelCode) = 1
        Else
            blnIPPVCallBack = False
        End If
    End If
    IPPVCallBack = blnIPPVCallBack
  Exit Function
ChkErr:
    ErrSub Me.Name, "IPPVCallBack"
End Function

Private Sub mskSmartCardNo1_GotFocus()
    Call objSelectAll(mskSmartCardNo1)

End Sub

Private Sub mskSmartCardNo2_GotFocus()
    Call objSelectAll(mskSmartCardNo2)

End Sub

Private Sub mskSTB1_GotFocus()
    Call objSelectAll(mskSTB1)
End Sub

Private Sub mskSTB2_GotFocus()
    Call objSelectAll(mskSTB2)

End Sub
