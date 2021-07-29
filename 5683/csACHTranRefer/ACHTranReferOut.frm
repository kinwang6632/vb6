VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   9135
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   7500
      TabIndex        =   25
      Top             =   4110
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.開始"
      Height          =   375
      Left            =   150
      TabIndex        =   24
      Top             =   4125
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.Frame frmData 
         Caption         =   "資料檔位置"
         Height          =   1275
         HelpContextID   =   2
         Left            =   285
         TabIndex        =   8
         Top             =   2475
         Width           =   8160
         Begin VB.TextBox txtErrPath 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3015
            TabIndex        =   12
            Top             =   765
            Width           =   4095
         End
         Begin VB.CommandButton cmdDataPath 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   14.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7320
            TabIndex        =   11
            Top             =   315
            Width           =   375
         End
         Begin VB.CommandButton cmdErrPath 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   14.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7335
            TabIndex        =   10
            Top             =   780
            Width           =   375
         End
         Begin VB.TextBox txtDataPath 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3015
            TabIndex        =   9
            ToolTipText     =   "請輸入字檔之路徑及檔名！"
            Top             =   300
            Width           =   4095
         End
         Begin VB.Label lblErrpath 
            AutoSize        =   -1  'True
            Caption         =   "問題參考檔位置（路徑 + 名稱）"
            Height          =   195
            Left            =   330
            TabIndex        =   14
            Top             =   810
            Width           =   2640
         End
         Begin VB.Label lblDatapath 
            AutoSize        =   -1  'True
            Caption         =   "資料檔位置（路徑 )"
            Height          =   195
            Left            =   1290
            TabIndex        =   13
            Top             =   375
            Width           =   1665
         End
      End
      Begin VB.ComboBox cboSet 
         Height          =   300
         ItemData        =   "ACHTranReferOut.frx":0000
         Left            =   2025
         List            =   "ACHTranReferOut.frx":0010
         TabIndex        =   7
         Top             =   780
         Width           =   2460
      End
      Begin VB.TextBox txtPayDate 
         Height          =   345
         Left            =   6075
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "00"
         Top             =   1140
         Width           =   330
      End
      Begin VB.TextBox txtDiscountRate 
         Height          =   345
         Left            =   2025
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "0"
         Top             =   1140
         Width           =   900
      End
      Begin VB.TextBox txtStatement 
         Height          =   345
         Left            =   2025
         TabIndex        =   4
         Top             =   1530
         Width           =   6465
      End
      Begin VB.TextBox txtMerchantName 
         Height          =   300
         Left            =   6060
         MaxLength       =   9
         TabIndex        =   3
         Text            =   "TBGB"
         Top             =   390
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.TextBox txtSpcNO 
         Height          =   300
         Left            =   2025
         MaxLength       =   9
         TabIndex        =   2
         Top             =   435
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.CheckBox chkDuteDate 
         Caption         =   "信用卡過期資料一併產生"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2040
         TabIndex        =   1
         Top             =   1995
         Width           =   2535
      End
      Begin Gi_Date.GiDate gdaSystemDate 
         Height          =   345
         Left            =   6060
         TabIndex        =   15
         Top             =   720
         Width           =   1125
         _ExtentX        =   1984
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
      Begin VB.Label Label1 
         Caption         =   "結帳種類"
         Height          =   285
         Left            =   1140
         TabIndex        =   23
         Top             =   870
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "請款日期"
         Height          =   210
         Left            =   5190
         TabIndex        =   22
         Top             =   1185
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   1  '靠右對齊
         Caption         =   "折扣率"
         Height          =   210
         Left            =   1215
         TabIndex        =   21
         Top             =   1245
         Width           =   720
      End
      Begin VB.Label Label4 
         Caption         =   "客戶帳單內容"
         Height          =   210
         Left            =   750
         TabIndex        =   20
         Top             =   1620
         Width           =   1185
      End
      Begin VB.Label Label8 
         Caption         =   "信用卡扣帳特店名稱"
         Height          =   285
         Left            =   4215
         TabIndex        =   19
         Top             =   435
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label9 
         Caption         =   "(DD)"
         Height          =   240
         Left            =   6465
         TabIndex        =   18
         Top             =   1170
         Width           =   1080
      End
      Begin VB.Label Label5 
         Alignment       =   1  '靠右對齊
         Caption         =   "系統日期"
         Height          =   285
         Left            =   4215
         TabIndex        =   17
         Top             =   780
         Width           =   1755
      End
      Begin VB.Label Label7 
         Alignment       =   1  '靠右對齊
         Caption         =   "信用卡扣帳特店代碼"
         Height          =   285
         Left            =   45
         TabIndex        =   16
         Top             =   510
         Visible         =   0   'False
         Width           =   1890
      End
   End
   Begin MSComDlg.CommonDialog cnd1 
      Left            =   7755
      Top             =   570
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
