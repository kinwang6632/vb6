VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSO1131C 
   Caption         =   "週期性收費項目編輯 [SO1131C]"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   4050
   ClientWidth     =   11940
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO1131C.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab sstDetail 
      Height          =   4065
      Left            =   5010
      TabIndex        =   50
      Top             =   3120
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   7170
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "優惠組合"
      TabPicture(0)   =   "SO1131C.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBPCode"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdContract"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdBPCode"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdOpen"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdStop"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdViewChange"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboPenalType"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboExpireType"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtOrderNo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtContNo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkBundle"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkPenal"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "發票抬頭"
      TabPicture(1)   =   "SO1131C.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraInvTitle"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "OTT 資訊"
      TabPicture(2)   =   "SO1131C.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label21"
      Tab(2).Control(1)=   "Label22"
      Tab(2).Control(2)=   "Label23"
      Tab(2).Control(3)=   "txtCardBillNo"
      Tab(2).Control(4)=   "txtFonsOrderNo"
      Tab(2).Control(5)=   "txtIOSOrderNo"
      Tab(2).ControlCount=   6
      Begin VB.TextBox txtIOSOrderNo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73200
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   123
         Top             =   1530
         Width           =   3015
      End
      Begin VB.TextBox txtFonsOrderNo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73200
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   122
         Top             =   1020
         Width           =   3015
      End
      Begin VB.TextBox txtCardBillNo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73200
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   117
         Top             =   510
         Width           =   3015
      End
      Begin VB.CheckBox chkPenal 
         Caption         =   "是否違約計算"
         Height          =   195
         Left            =   1500
         TabIndex        =   93
         Top             =   690
         Width           =   1545
      End
      Begin VB.CheckBox chkBundle 
         Caption         =   "是否綁約"
         Height          =   195
         Left            =   360
         TabIndex        =   91
         Top             =   690
         Width           =   1215
      End
      Begin VB.TextBox txtContNo 
         Height          =   315
         Left            =   3210
         TabIndex        =   90
         Top             =   930
         Width           =   1605
      End
      Begin VB.TextBox txtOrderNo 
         Height          =   315
         Left            =   1230
         TabIndex        =   89
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cboExpireType 
         Height          =   315
         ItemData        =   "SO1131C.frx":0496
         Left            =   1980
         List            =   "SO1131C.frx":0498
         TabIndex        =   102
         Top             =   3390
         Width           =   2715
      End
      Begin VB.ComboBox cboPenalType 
         Height          =   315
         ItemData        =   "SO1131C.frx":049A
         Left            =   1980
         List            =   "SO1131C.frx":049C
         TabIndex        =   101
         Top             =   3060
         Width           =   2715
      End
      Begin VB.CommandButton cmdViewChange 
         Caption         =   "資料異動記錄檔"
         Height          =   315
         Left            =   4860
         TabIndex        =   80
         Top             =   1620
         Width           =   1635
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "關閉優惠"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4860
         TabIndex        =   79
         Top             =   1290
         Width           =   1635
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "重新開啟優惠"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4860
         TabIndex        =   78
         Top             =   960
         Width           =   1635
      End
      Begin VB.CommandButton cmdBPCode 
         Caption         =   "顯示階段式優惠"
         Height          =   315
         Left            =   4860
         TabIndex        =   77
         Top             =   600
         Width           =   1635
      End
      Begin VB.CommandButton cmdContract 
         Caption         =   "合約歷程"
         Height          =   315
         Left            =   4860
         TabIndex        =   76
         Top             =   2010
         Width           =   1635
      End
      Begin VB.Frame fraBPCode 
         Caption         =   "優惠組合"
         Enabled         =   0   'False
         Height          =   3615
         Left            =   150
         TabIndex        =   66
         Top             =   390
         Width           =   6495
         Begin VB.TextBox txtDiscountAmt 
            Alignment       =   1  '靠右對齊
            Height          =   315
            Left            =   1080
            TabIndex        =   99
            Top             =   1980
            Width           =   1095
         End
         Begin VB.TextBox txtBundleMon 
            Alignment       =   1  '靠右對齊
            Height          =   315
            Left            =   5490
            MaxLength       =   2
            TabIndex        =   67
            Top             =   2310
            Width           =   405
         End
         Begin VB.CheckBox chkMergePrint 
            Caption         =   "是否合併列印"
            Height          =   195
            Left            =   3060
            TabIndex        =   96
            Top             =   300
            Width           =   1545
         End
         Begin Gi_Date.GiDate gdaContStopDate 
            Height          =   315
            Left            =   2790
            TabIndex        =   94
            Top             =   900
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
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
         Begin Gi_Date.GiDate gdaContStartDate 
            Height          =   315
            Left            =   1080
            TabIndex        =   92
            Top             =   900
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
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
         Begin prjGiList.GiList gilBPCode 
            Height          =   315
            Left            =   1080
            TabIndex        =   95
            Top             =   1230
            Width           =   3465
            _ExtentX        =   6112
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
            FldWidth2       =   2400
            TableName       =   "CD019"
            F2Corresponding =   -1  'True
         End
         Begin prjGiList.GiList gilPromCode 
            Height          =   315
            Left            =   1080
            TabIndex        =   100
            Top             =   2310
            Width           =   3465
            _ExtentX        =   6112
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
            FldWidth2       =   2400
            TableName       =   "CD019"
            F2Corresponding =   -1  'True
         End
         Begin Gi_Date.GiDate gdaDiscountDate3 
            Height          =   315
            Left            =   1080
            TabIndex        =   97
            Top             =   1620
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin Gi_Date.GiDate gdaDiscountDate4 
            Height          =   315
            Left            =   2790
            TabIndex        =   98
            Top             =   1620
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin VB.Label lblOrderNo 
            AutoSize        =   -1  'True
            Caption         =   "訂單單號"
            Height          =   195
            Left            =   210
            TabIndex        =   88
            Top             =   630
            Width           =   780
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "優待金額"
            Height          =   195
            Left            =   210
            TabIndex        =   85
            Top             =   2040
            Width           =   780
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   195
            Left            =   2430
            TabIndex        =   82
            Top             =   1650
            Width           =   195
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "優待階段"
            Height          =   195
            Left            =   210
            TabIndex        =   81
            Top             =   1680
            Width           =   780
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "合約"
            Height          =   195
            Left            =   2670
            TabIndex        =   75
            Top             =   630
            Width           =   390
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "合約期限"
            Height          =   195
            Left            =   210
            TabIndex        =   74
            Top             =   960
            Width           =   780
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   195
            Left            =   2430
            TabIndex        =   73
            Top             =   975
            Width           =   195
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "優惠組合"
            Height          =   195
            Left            =   210
            TabIndex        =   72
            Top             =   1290
            Width           =   780
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "促銷方案"
            Height          =   195
            Left            =   210
            TabIndex        =   71
            Top             =   2370
            Width           =   780
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "綁約月數"
            Height          =   195
            Left            =   4620
            TabIndex        =   70
            Top             =   2370
            Width           =   780
         End
         Begin VB.Label lblPenalType 
            AutoSize        =   -1  'True
            Caption         =   "違約時之計價依據"
            Height          =   195
            Left            =   210
            TabIndex        =   69
            Top             =   2730
            Width           =   1560
         End
         Begin VB.Label lblExpireType 
            AutoSize        =   -1  'True
            Caption         =   "優惠到期計價依據"
            Height          =   195
            Left            =   210
            TabIndex        =   68
            Top             =   3060
            Width           =   1560
         End
      End
      Begin VB.Frame fraInvTitle 
         Height          =   3495
         Left            =   -74940
         TabIndex        =   51
         Top             =   390
         Width           =   6585
         Begin VB.Label lblA_CarrierId2 
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1860
            TabIndex        =   111
            Top             =   3120
            Width           =   4575
         End
         Begin VB.Label lblA_CarrierId1 
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1860
            TabIndex        =   110
            Top             =   2820
            Width           =   4575
         End
         Begin VB.Label Label19 
            Caption         =   "會員載具隱碼"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   540
            TabIndex        =   109
            Top             =   3150
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "會員載具顯碼"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   540
            TabIndex        =   108
            Top             =   2820
            Width           =   1215
         End
         Begin VB.Label lblInvoiceKind 
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1860
            TabIndex        =   107
            Top             =   2520
            Width           =   4575
         End
         Begin VB.Label Label13 
            Caption         =   "發票開立種類"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   540
            TabIndex        =   106
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label lblMailAddressValue 
            Caption         =   "lblChargeAddressValue"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1860
            TabIndex        =   84
            Top             =   2220
            Width           =   4575
         End
         Begin VB.Label lblMailAddress 
            Caption         =   "郵  寄  地  址"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   660
            TabIndex        =   83
            Top             =   2220
            Width           =   1125
         End
         Begin VB.Label lblInvSeqNo 
            Caption         =   "發票抬頭流水號"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   330
            TabIndex        =   65
            Top             =   210
            Width           =   1395
         End
         Begin VB.Label lblInvSeqNoValue 
            Caption         =   "1234569"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1860
            TabIndex        =   64
            Top             =   210
            Width           =   4575
         End
         Begin VB.Label lblChargeTitle 
            Caption         =   "收  件  人"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   930
            TabIndex        =   63
            Top             =   495
            Width           =   795
         End
         Begin VB.Label lblChargeTitleValue 
            Caption         =   "收件人值"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1860
            TabIndex        =   62
            Top             =   495
            Width           =   4575
         End
         Begin VB.Label lblInvoiceType 
            Caption         =   "發  票  種  類"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   660
            TabIndex        =   61
            Top             =   765
            Width           =   1095
         End
         Begin VB.Label lblInvoiceTypeValue 
            Caption         =   "發  票  種  類值"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1860
            TabIndex        =   60
            Top             =   750
            Width           =   4575
         End
         Begin VB.Label lblInvNo 
            Caption         =   "統  一  編  號"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   660
            TabIndex        =   59
            Top             =   1050
            Width           =   1125
         End
         Begin VB.Label lblInvNoValue 
            Caption         =   "統  一  編  號值"
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   1860
            TabIndex        =   58
            Top             =   1050
            Width           =   4575
         End
         Begin VB.Label lblInvTitle 
            Caption         =   "發  票  抬  頭"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   660
            TabIndex        =   57
            Top             =   1335
            Width           =   1125
         End
         Begin VB.Label lblInvTitleValue 
            Caption         =   "發  票  抬  投  值"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1860
            TabIndex        =   56
            Top             =   1335
            Width           =   4575
         End
         Begin VB.Label lblInvAddress 
            Caption         =   "發  票  地  址"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   660
            TabIndex        =   55
            Top             =   1635
            Width           =   1125
         End
         Begin VB.Label lblInvAddressValue 
            Caption         =   "發  票  地  址 值"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1860
            TabIndex        =   54
            Top             =   1620
            Width           =   4575
         End
         Begin VB.Label lblChargeAddress 
            Caption         =   "收  費  地  址"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   660
            TabIndex        =   53
            Top             =   1920
            Width           =   1125
         End
         Begin VB.Label lblChargeAddressValue 
            Caption         =   "lblChargeAddressValue"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1860
            TabIndex        =   52
            Top             =   1920
            Width           =   4575
         End
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "iOS訂單單號"
         Height          =   195
         Left            =   -74910
         TabIndex        =   121
         Top             =   1620
         Width           =   1065
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "FonsView訂單單號"
         Height          =   195
         Left            =   -74910
         TabIndex        =   120
         Top             =   1080
         Width           =   1530
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "信用卡訂單編號"
         Height          =   195
         Left            =   -74910
         TabIndex        =   118
         Top             =   570
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdHouse 
      Caption         =   "查詢設備明細"
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdShowRate 
      Caption         =   "F7. 顯示費率表 >>"
      Height          =   315
      Left            =   5310
      TabIndex        =   24
      Top             =   210
      Width           =   1905
   End
   Begin VB.Frame fraRate 
      BorderStyle     =   0  '沒有框線
      Height          =   2385
      Left            =   5130
      TabIndex        =   38
      Top             =   540
      Visible         =   0   'False
      Width           =   6465
      Begin prjGiGridR.GiGridR ggrCustRate 
         Height          =   2055
         Left            =   150
         TabIndex        =   25
         Top             =   360
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   3625
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
      Begin prjGiGridR.GiGridR ggrMduRate 
         Height          =   2055
         Left            =   3390
         TabIndex        =   26
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   3625
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
      Begin VB.Label lblMduName 
         Height          =   195
         Left            =   4380
         TabIndex        =   42
         Top             =   120
         Width           =   2115
      End
      Begin VB.Label lblCustClass 
         Height          =   195
         Left            =   1620
         TabIndex        =   41
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblMduRate 
         AutoSize        =   -1  'True
         Caption         =   "大樓費率表"
         Height          =   195
         Left            =   3330
         TabIndex        =   40
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblCustRate 
         AutoSize        =   -1  'True
         Caption         =   "客戶類別費率表"
         Height          =   195
         Left            =   180
         TabIndex        =   39
         Top             =   120
         Width           =   1365
      End
   End
   Begin VB.Frame fraData 
      Height          =   7305
      Left            =   60
      TabIndex        =   27
      Top             =   90
      Width           =   11805
      Begin VB.CheckBox chkPeriodFlag 
         Caption         =   "指定繳別出單"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2490
         TabIndex        =   119
         Top             =   2280
         Width           =   2205
      End
      Begin VB.TextBox txtSTBSeqNo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   116
         Top             =   4890
         Width           =   3045
      End
      Begin VB.CheckBox chkLongPay 
         Caption         =   "長繳別註記"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2490
         TabIndex        =   112
         Top             =   2610
         Width           =   2205
      End
      Begin VB.TextBox txtAccountNo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   6630
         Width           =   3210
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "置換"
         Height          =   315
         Left            =   4650
         TabIndex        =   103
         Top             =   150
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.CommandButton cmdInvSeqNo 
         Caption         =   "選擇發票抬頭"
         Height          =   315
         Left            =   1560
         TabIndex        =   18
         Top             =   5250
         Width           =   1365
      End
      Begin VB.CheckBox chkStopFlag 
         Caption         =   "停用"
         Height          =   195
         Left            =   1200
         TabIndex        =   2
         Top             =   990
         Width           =   1455
      End
      Begin VB.TextBox txtFaciSNo 
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   0
         Top             =   540
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.CheckBox chkCustAllot 
         Caption         =   "客戶指定"
         Height          =   195
         Left            =   2910
         TabIndex        =   7
         Top             =   1665
         Width           =   1215
      End
      Begin prjNumber.GiNumber ginBudgetPeriod 
         Height          =   315
         Left            =   2910
         TabIndex        =   5
         Top             =   1260
         Width           =   405
         _ExtentX        =   714
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
      Begin VB.CommandButton cmdChooseAcc 
         Caption         =   "選擇付款帳號"
         Height          =   315
         Left            =   150
         TabIndex        =   17
         Top             =   5250
         Width           =   1365
      End
      Begin VB.ComboBox cboAccountNo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   6630
         Visible         =   0   'False
         Width           =   3465
      End
      Begin VB.Frame fraDiscount 
         Caption         =   "優待辦法"
         Enabled         =   0   'False
         Height          =   975
         Left            =   90
         TabIndex        =   34
         Top             =   2880
         Width           =   4815
         Begin prjNumber.GiNumber ginDiscountAmt 
            Height          =   315
            Left            =   1110
            TabIndex        =   13
            Top             =   630
            Width           =   1155
            _ExtentX        =   2037
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
            MaxLength       =   10
            AutoSelect      =   -1  'True
            AllowMinus      =   -1  'True
         End
         Begin prjGiList.GiList gilDiscount 
            Height          =   315
            Left            =   1110
            TabIndex        =   12
            Top             =   270
            Width           =   3465
            _ExtentX        =   6112
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
            FldWidth2       =   2400
            TableName       =   "CD019"
            F2Corresponding =   -1  'True
         End
         Begin VB.Label lblDiscount 
            AutoSize        =   -1  'True
            Caption         =   "優待辦法"
            Height          =   195
            Left            =   240
            TabIndex        =   36
            Top             =   330
            Width           =   795
         End
         Begin VB.Label lblDiscountAmt 
            AutoSize        =   -1  'True
            Caption         =   "優待金額"
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   690
            Width           =   795
         End
      End
      Begin prjNumber.GiNumber ginAmount 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   1590
         Width           =   1155
         _ExtentX        =   2037
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
         MaxLength       =   10
         AutoSelect      =   -1  'True
         AllowMinus      =   -1  'True
      End
      Begin VB.TextBox txtPeriod 
         Alignment       =   1  '靠右對齊
         Height          =   315
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1260
         Width           =   405
      End
      Begin Gi_Date.GiDate gdaStartDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   1920
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
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
      Begin Gi_Date.GiDate gdaStopDate 
         Height          =   315
         Left            =   2910
         TabIndex        =   9
         Top             =   1920
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
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
      Begin Gi_Date.GiDate gdaClctDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   2250
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
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
      Begin prjGiList.GiList gilPackageNo 
         Height          =   315
         Left            =   1200
         TabIndex        =   16
         Top             =   4200
         Width           =   3465
         _ExtentX        =   6112
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
         FldWidth2       =   2400
         TableName       =   "CD019"
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilCMCode 
         Height          =   315
         Left            =   1200
         TabIndex        =   20
         Top             =   5940
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth2       =   2400
         TableName       =   "CD019"
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilBankCode 
         Height          =   315
         Left            =   1200
         TabIndex        =   21
         Top             =   6270
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth2       =   2400
         TableName       =   "CD019"
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilPTCode 
         Height          =   315
         Left            =   1200
         TabIndex        =   19
         Top             =   5610
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth2       =   2400
         TableName       =   "CD019"
         F2Corresponding =   -1  'True
      End
      Begin Gi_Date.GiDate gdaCeaseDate 
         Height          =   315
         Left            =   3540
         TabIndex        =   3
         Top             =   930
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin Gi_Date.GiDate gdaDiscountDate2 
         Height          =   315
         Left            =   2910
         TabIndex        =   15
         Top             =   3870
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
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
      Begin Gi_Date.GiDate gdaDiscountDate1 
         Height          =   315
         Left            =   1200
         TabIndex        =   14
         Top             =   3870
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
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
      Begin prjGiList.GiList gilCitemCode 
         Height          =   315
         Left            =   1200
         TabIndex        =   104
         Top             =   150
         Width           =   3465
         _ExtentX        =   6112
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
         FldWidth2       =   2400
         TableName       =   "CD019"
         F2Corresponding =   -1  'True
      End
      Begin Gi_Date.GiDate gdaClctStopDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   2580
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin prjGiList.GiList gilSalePoint 
         Height          =   315
         Left            =   1200
         TabIndex        =   114
         Top             =   4530
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth2       =   2400
         TableName       =   "CD019"
         F2Corresponding =   -1  'True
      End
      Begin VB.Label lblSTBSeqNo 
         AutoSize        =   -1  'True
         Caption         =   "STB設備流水號"
         Height          =   195
         Left            =   240
         TabIndex        =   115
         Top             =   4920
         Width           =   1305
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "行銷辦法"
         Height          =   195
         Left            =   240
         TabIndex        =   113
         Top             =   4590
         Width           =   780
      End
      Begin VB.Label lblClctStopDate 
         AutoSize        =   -1  'True
         Caption         =   "出單到期日"
         Height          =   195
         Left            =   60
         TabIndex        =   105
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lblDiscountDate 
         AutoSize        =   -1  'True
         Caption         =   "總優惠期限"
         Height          =   195
         Left            =   150
         TabIndex        =   87
         Top             =   3930
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2550
         TabIndex        =   86
         Top             =   3945
         Width           =   195
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "停用日期"
         Height          =   195
         Left            =   2715
         TabIndex        =   49
         Top             =   990
         Width           =   780
      End
      Begin VB.Label lblFaciSNo 
         AutoSize        =   -1  'True
         Caption         =   "設備序號"
         Height          =   195
         Left            =   240
         TabIndex        =   48
         Top             =   660
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "分期剩餘期數"
         Height          =   195
         Left            =   1710
         TabIndex        =   47
         Top             =   1350
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "付款種類"
         Height          =   195
         Left            =   240
         TabIndex        =   46
         Top             =   5670
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "扣帳帳號"
         Height          =   195
         Left            =   240
         TabIndex        =   45
         Top             =   6690
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "收費方式"
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   6000
         Width           =   780
      End
      Begin VB.Label lblBank 
         AutoSize        =   -1  'True
         Caption         =   "銀行名稱"
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   6330
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "套餐名稱"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   4260
         Width           =   780
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         Caption         =   "有效期限"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   1995
         Width           =   795
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2550
         TabIndex        =   32
         Top             =   1995
         Width           =   195
      End
      Begin VB.Label lblCitemCode 
         AutoSize        =   -1  'True
         Caption         =   "收費項目"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   270
         Width           =   795
      End
      Begin VB.Label lblClctDate 
         AutoSize        =   -1  'True
         Caption         =   "次收費日"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   2325
         Width           =   795
      End
      Begin VB.Label lblAmount 
         AutoSize        =   -1  'True
         Caption         =   "收費金額"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   1665
         Width           =   795
      End
      Begin VB.Label lblPeriod 
         AutoSize        =   -1  'True
         Caption         =   "每期月數"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   1350
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmSO1131C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Program Name :     SO1131C
' Version      :     1.0.0(Alpha)
' Developer    :     Hammer
' Written Date :     1999/11/20
' Last Modify  :     1999/12/16

Option Explicit
Private rsCust As New ADODB.Recordset
Private rsMdu As New ADODB.Recordset
Private rs As New ADODB.Recordset
Private lngEditMode As giEditModeEnu
Dim rsSO043 As New ADODB.Recordset
Dim rsTmp As New ADODB.Recordset
Dim rsSO138 As New ADODB.Recordset
Dim strInvSeqNo As String
Dim intClassCode As Long
Dim strMduId As String
Dim lngCompCode As Long
Dim strCustName As String
Dim rsDiscount As New Recordset
Private strChgCitem As String
Private blnChgCitem As Boolean
Private rsFROM As New ADODB.Recordset
Private strServiceType As String
Private blnShowInv As Boolean
Private varC() As String
Private blnSO1131FSave As Boolean

Public Sub UserPermissionGo()
    On Error GoTo ChkErr
    Dim blnFlag As Boolean
        blnFlag = Not frmSO1100BMDI.rsSO003.EOF
        ' 根據權限組別, 取相關操作權限
        MenuEnabled GetUserPriv("SO1131", "(SO11311)"), _
                            GetUserPriv("SO1131", "(SO11312)") And blnFlag, _
                            GetUserPriv("SO1131", "(SO11313)") And blnFlag, _
                            , GetUserPriv("SO1131", "(SO11315)") And blnFlag And False
                            
    Exit Sub
ChkErr:
    ErrSub Me.Name, "UserPermissionGo"
End Sub

Public Sub FirstGo()
    On Error GoTo ChkErr
        Call ErrorActiveFrmHandle("frmSO1131C")
        
    Exit Sub
ChkErr:
    ErrSub Me.Name, "FirstGo"
End Sub

Public Sub NextGo()
    On Error GoTo ChkErr
        Call ErrorActiveFrmHandle("frmSO1131C")
    Exit Sub
ChkErr:
    ErrSub Me.Name, "NextGo"
End Sub

Public Sub PreviousGo()
    On Error GoTo ChkErr
        Call ErrorActiveFrmHandle("frmSO1131C")
    Exit Sub
ChkErr:
    ErrSub Me.Name, "PreviousGo"
End Sub

Public Sub LastGo()
    On Error GoTo ChkErr
        Call ErrorActiveFrmHandle("frmSO1131C")
    Exit Sub
ChkErr:
    ErrSub Me.Name, "LastGo"
End Sub

Public Sub AddNewGo()
  On Error GoTo ChkErr
    If frmSO1100BMDI.rsSo002("CustStatusCode") <> 1 Then
        If MsgBox("該客戶非正常收視中客戶,是否確認新增?", vbYesNo, "確認") = vbNo Then
            'Call FormQueryUnload
            Exit Sub
        End If
    End If
    With frmSO1131C
        .EditMode = giEditModeInsert
        frmSO1100BMDI.Pic1.Enabled = False
        .Show , frmSO1100BMDI
    End With
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "AddNewGo")
End Sub

Public Sub EditGo()
  On Error GoTo ChkErr
    With frmSO1131C
        .EditMode = giEditModeEdit
        frmSO1100BMDI.Pic1.Enabled = False
        .uServiceType = frmSO1100BMDI.rsSO003("ServiceType") & ""
        .Show , frmSO1100BMDI
    End With
    '#4030 如果BpCode
    If rs("BpCode") & "" <> "" Then
        fraDiscount.Enabled = False
        gilPackageNo.Enabled = False
    Else
        fraDiscount.Enabled = True
        gilPackageNo.Enabled = True
    End If
    '********************************************************
    '如果StopType>0收費項目不允許被修改 By Kin 2008/12/11
    If Val(rs("StopType") & "") > 0 Then
        gilCitemCode.Enabled = False
    Else
        gilCitemCode.Enabled = True
    End If
    '********************************************************
        If GetUserPriv("SO1131", "(SO1131B)") Then
            gdaDiscountDate1.Enabled = True
            gdaDiscountDate1.Enabled = True
        Else
            gdaDiscountDate1.Enabled = False
            gdaDiscountDate1.Enabled = False
        End If
    '#7225
    If GetUserPriv("SO1131", "(SO1131K)") Then
        txtPeriod.Enabled = True
    Else
        txtPeriod.Enabled = False
    End If
    If GetUserPriv("SO1131", "(SO1131L)") Then
        ginAmount.Enabled = True
    Else
        ginAmount.Enabled = False
    End If
    If GetUserPriv("SO1131", "(SO1131M)") Then
        gdaClctDate.Enabled = True
    Else
        gdaClctDate.Enabled = False
    End If
     If GetUserPriv("SO1131", "(SO1131N)") Then
        gdaStartDate.Enabled = True
        gdaStopDate.Enabled = True
    Else
        gdaStartDate.Enabled = False
        gdaStopDate.Enabled = False
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "EditGo")
End Sub

Public Sub ViewGo()
  On Error GoTo ChkErr
    With frmSO1131C
        .EditMode = giEditModeView
        .Show vbModal
    End With
 Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ViewGo")
End Sub

Public Sub DeleteGo()
    On Error Resume Next
        '#3374 改變MsgBox內容-> 確定刪除此收費項目資料嗎 By Kin 2008/03/12
        If MsgBox("確定作廢此週期資料嗎 ?", vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub
    On Error GoTo ChkErr
    Dim strCancelDate As String
    Dim strCancelCode As String
    Dim strCancelName As String
        If (rsFROM Is Nothing) Or (rsFROM.State = adStateClosed) Then
            Set rs = frmSO1100BMDI.rsSO003
            '#5812 從外層進來按刪除,ServiceType會是空值,導致SO003A會沒被刪除,
            '所以外層的ServiceType強迫指定為 gServiceType By Kin 2010/10/21
            strServiceType = gServiceType
        Else
            Set rs = rsFROM
        End If
        With frmSO1110E
            Set .uParentForm = frmSO1100BMDI
            .uType = 4  '週期性收費資料
            .Show vbModal
            If Not .uSaveOk Then Exit Sub
            strCancelDate = .uCancelDate
            strCancelCode = .uCancelCode
            strCancelName = .uCancelName
        End With
        
        gcnGi.BeginTrans
        Me.uCustName = frmSO1100BMDI.txtCustName.Text
        '*******************************************************************************************************
        '#4448 如果是阿貴呼叫,則要用阿貴的Recordset By Kin 2009/04/13
        If Not InsertToLog(rs, giEditModeDelete) Then
        
            gcnGi.RollbackTrans
            Exit Sub
        End If
        gcnGi.Execute "Update " & GetOwner & "SO003 Set CancelDate = " & GetNullString(strCancelDate, giDateV) & ",CancelCode = " & GetNullString(strCancelCode) & ",CancelName = " & GetNullString(strCancelName) & " Where RowId = '" & rs("RowId") & "'"
        
        gcnGi.Execute "Delete From " & GetOwner & "SO003A Where CustId = " & rs("CustId") & _
            " And CompCode = " & gCompCode & " And ServiceType = '" & strServiceType & "' And CitemCode = " & rs("CitemCode") & _
            " And ContNo = '" & rs("ContNo") & "' And FaciSeqNo = '" & rs("FaciSeqNo") & "'"
        
        gcnGi.Execute "Delete From " & GetOwner & "SO003 Where RowId = '" & rs("RowId") & "'"
        gcnGi.CommitTrans
        MsgBox "作廢成功！", vbExclamation, "訊息"
        Call RefreshSo1100BNothing
        '*******************************************************************************************************
    Exit Sub
ChkErr:
    gcnGi.RollbackTrans
    Call ErrSub(Me.Name, "DeleteGo")
End Sub

Public Sub UpdateGo()
  On Error GoTo ChkErr
    Dim StarTrans As Boolean
    StarTrans = False
    If IsDataOk() = False Then
        Exit Sub 'Is data ok?
    End If
    gcnGi.BeginTrans
    StarTrans = True
   
    If Not frmSO1100BMDI Is Nothing Then frmSO1100BMDI.Enabled = False
    Me.Enabled = False
    If ScrToRcd = False Then
        Me.Enabled = True
        If Not frmSO1100BMDI Is Nothing Then frmSO1100BMDI.Enabled = True
        gcnGi.RollbackTrans
        rs.Requery
        Exit Sub            'Save record
    End If
     '#6841 判斷是否為長繳別,要更新SO003A By Kin 2014/08/19
     '#6865 測試不OK,不要判斷是否有優惠組合，由後端自行判斷 By Kin 2014/09/30
     '#7070 修正有設定長繳別無法停用 By Kin 2015/08/05
    If chkStopFlag.Value = 0 Then
        If Val(rs("LongPayFlag") & "") = 1 Then
            Dim RetMsg As String
            Dim strFaciSeqNo As String
            If GetByHouse(gilCitemCode.GetCodeNo) = 1 Then
                strFaciSeqNo = gCustId
            Else
                strFaciSeqNo = txtFaciSNo.Tag
            End If
            If Not ExecuteSingleTransfer(garyGi(1), gCustId, gilCitemCode.GetCodeNo, strFaciSeqNo, RetMsg) Then
                Me.Enabled = True
                 If Not frmSO1100BMDI Is Nothing Then frmSO1100BMDI.Enabled = True
                gcnGi.RollbackTrans
                rs.Requery
                Exit Sub
            End If
        End If
    End If
    If Not frmSO1100BMDI Is Nothing Then frmSO1100BMDI.Enabled = True
    Me.Enabled = True
    Me.Tag = "Save"
    strChgCitem = Empty
    gcnGi.CommitTrans
    Unload Me
  Exit Sub
ChkErr:
    If StarTrans Then gcnGi.RollbackTrans
    Call ErrSub(Me.Name, "UpdateGo")
End Sub

Public Sub PrintGo()

End Sub

Public Sub CancelGo()
    On Error Resume Next
        strChgCitem = Empty
        Unload Me
End Sub

' 自訂屬性：EditMode
' 記錄目前在編輯、新增或檢視模式
' giEditModeEnu(自訂列舉值，設定於Sys_Lib)
Public Property Get EditMode() As giEditModeEnu ' 取目前編輯模式
  On Error GoTo ChkErr
    EditMode = lngEditMode
  Exit Property
ChkErr:
   Call ErrSub(Me.Name, "Get EditMode")
End Property

Public Property Let EditMode(ByVal vNewValue As giEditModeEnu) '設定編輯模式
  On Error GoTo ChkErr
    lngEditMode = vNewValue
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Let EditMode")
End Property

'根據傳來之編輯模式, 設定各物件屬性
Public Sub ChangeMode(ByVal lngMode As giEditModeEnu)
  On Error GoTo ChkErr
    Dim blnFlag As Boolean  ' 記錄是否在資料瀏覽模式，
    lngEditMode = lngMode
    Select Case lngMode
           Case giEditModeView
                blnFlag = True
                Call RcdToScr
           Case giEditModeEdit
                blnFlag = False
                Call RcdToScr
           Case giEditModeInsert
                blnFlag = False
                Call NewRcd
    End Select
    'txtAccountNo.Visible = Not GetUserPriv("SO1100G", "(SO1100G9)")
    txtAccountNo.Enabled = GetUserPriv("SO1100G", "(SO1100G9)")
    cboAccountNo.Enabled = GetUserPriv("SO1100G", "(SO1100G9)")
    
    SetStatusBar lngEditMode
    ginBudgetPeriod.Enabled = GetUserPriv("SO1131", "(SO11316)")
    fraDiscount.Enabled = GetUserPriv("SO1131", "(SO11317)")
    '5996 增加權限控管銀行與帳號 By Kin 2011/06/03
    cmdChooseAcc.Enabled = GetUserPriv("SO1131", "(SO1131J)")
    'cboAccountNo.Visible = cmdChooseAcc.Enabled
    gilBankCode.Enabled = cmdChooseAcc.Enabled
    '93/06/11
    chkCustAllot.Enabled = GetUserPriv("SO1131", "(SO11318)")
    cmdBPCode.Enabled = GetUserPriv("SO1131", "(SO11319)")
    chkStopFlag.Enabled = GetUserPriv("SO1131", "(SO1131A)")
    If GetUserPriv("SO1131", "(SO1131B)") Then
        gdaDiscountDate1.Enabled = True
        gdaDiscountDate2.Enabled = True
    Else
        gdaDiscountDate1.Enabled = False
        gdaDiscountDate2.Enabled = False
    End If
    '#4133 SO1131C中之【違約時計價依據】及【優惠到期計價依據】各自增加二個權限出來可做修改 By Kin 2008/10/08
    If blnFlag Then
        cboPenalType.Enabled = False
        cboExpireType.Enabled = False
    Else
        cboPenalType.Enabled = GetUserPriv("SO1131", "(SO1131C)")
        cboExpireType.Enabled = GetUserPriv("SO1131", "(SO1131D)")
    End If
    '#4143 請增加SO1131C　欄位權限。『訂單單號』、『合約編號』、『是否綁約』及『是否違約計算』四個欄位權限 By Kin 2008/10/13
    If blnFlag Then
        txtOrderNo.Enabled = False
        txtContNo.Enabled = False
        chkBundle.Enabled = False
        chkPenal.Enabled = False
    Else
        txtOrderNo.Enabled = GetUserPriv("SO1131", "(SO1131E)")
        txtContNo.Enabled = GetUserPriv("SO1131", "(SO1131F)")
        chkBundle.Enabled = GetUserPriv("SO1131", "(SO1131G)")
        chkPenal.Enabled = GetUserPriv("SO1131", "(SO1131H)")
    End If
    '*************************************************************************************
    '#4150 控制是否要同步更新SO003A,阿貴說也要控制收費項目是否可以修改 By Kin 2008/10/22
    '#4287 增加同步設備流水號 By Kin 2008/12/12
    blnChgCitem = GetUserPriv("SO1131", "(SO1131I)")
    gilCitemCode.Enabled = blnChgCitem
    cmdHouse.Enabled = blnChgCitem
    '*************************************************************************************
    'txtPeriod.Enabled = GetUserPriv("SO1131", "(SO11318)")
    'fraDiscount.Enabled = Not blnFlag
    fraData.Enabled = Not blnFlag
    'fraBPCode.Enabled = Not blnFlag
    sstDetail.Enabled = True
     '#7225
    If Not blnFlag Then
        If GetUserPriv("SO1131", "(SO1131K)") Then
            txtPeriod.Enabled = True
        Else
            txtPeriod.Enabled = False
        End If
        If GetUserPriv("SO1131", "(SO1131L)") Then
            ginAmount.Enabled = True
        Else
            ginAmount.Enabled = False
        End If
        If GetUserPriv("SO1131", "(SO1131M)") Then
            gdaClctDate.Enabled = True
        Else
            gdaClctDate.Enabled = False
        End If
         If GetUserPriv("SO1131", "(SO1131N)") Then
            gdaStartDate.Enabled = True
            gdaStopDate.Enabled = True
        Else
            gdaStartDate.Enabled = False
            gdaStopDate.Enabled = False
        End If
    End If
   
    
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "ChangeMode")
End Sub

Private Sub cboAccountNo_Change()
    Call cboAccountNo_Click
End Sub
Private Sub EnabledCtl()
  On Error GoTo ChkErr
    Dim aSQL As String
    aSQL = "SELECT NVL(CustomerByFaci,0) FROM " & GetOwner & "SO041 "
    blnShowInv = gcnGi.Execute(aSQL)(0) = 0
    cmdChooseAcc.Visible = False
    cmdInvSeqNo.Visible = False
    cboAccountNo.Visible = False
    cboAccountNo.Enabled = False
    gilPTCode.Enabled = False
    gilCMCode.Enabled = False
    gilBankCode.Enabled = False
    txtAccountNo.Enabled = False
    If blnShowInv Then
        cmdChooseAcc.Visible = True
        cmdInvSeqNo.Visible = True
        cboAccountNo.Visible = True
        cboAccountNo.Enabled = True
        gilPTCode.Enabled = True
        gilCMCode.Enabled = True
        gilBankCode.Enabled = True
        txtAccountNo.Enabled = True
        '#5996 增加權限控管 By Kin 2011/05/30
        cmdChooseAcc.Enabled = GetUserPriv("SO1131", "(SO1131J)")
        'cboAccountNo.Visible = cmdChooseAcc.Enabled
        gilBankCode.Enabled = cmdChooseAcc.Enabled
        
    End If
    '後付制專案需求,判斷期數是否可以修改 By Kin 2010/09/16
    '#7225
    If txtPeriod.Enabled Then
        aSQL = "SELECT NVL(PayNowNoChangePeriod,0) PayNowNoChangePeriod FROM " & GetOwner & "SO041"
        If gcnGi.Execute(aSQL)(0) = 1 Then
            If Not rs.EOF Then
                If Val(rs("PayKind") & "") = 1 Then
                    txtPeriod.Enabled = False
                End If
            End If
        End If
    End If
     '#8724
    If Len(cboAccountNo.Text & "") > 0 Or Len(txtAccountNo.Text & "") > 0 Then
        cmdChooseAcc.Enabled = False
        gilPTCode.Enabled = False
        gilCMCode.Enabled = False
        gilBankCode.Enabled = False
        cboAccountNo.Enabled = False
        txtAccountNo.Enabled = False
    End If
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "EnabledCtl")
End Sub
Private Sub cboAccountNo_Click()
    On Error Resume Next
        If Not GetUserPriv("SO1100G", "(SO1100G9)") Then
            txtAccountNo = GetMaskAccountNo(cboAccountNo)
        Else
            txtAccountNo = cboAccountNo.Text
        End If
                
        If Not GetRS(rsTmp, "Select CMCode,CMName From " & GetOwner & "SO106 Where CustId= " & gCustId & " And AccountId = '" & cboAccountNo & "' And StopFlag =0") Then Exit Sub
        If Not rsTmp.EOF Then
            '********************************************************************************************
            '#3975 每次一進去就會顯示SO106的收費方式,判斷如果SO003有值就不以SO106的值 By Kin 2008/06/16
            If rs.RecordCount > 0 Then
                If IsNull(rs("CMCode")) Then
                    gilCMCode.SetCodeNo rsTmp("CMCode") & ""
                    gilCMCode.SetDescription rsTmp("CMName") & ""
                End If
            Else
                gilCMCode.SetCodeNo rsTmp("CMCode") & ""
                gilCMCode.SetDescription rsTmp("CMName") & ""
            End If
            '********************************************************************************************
        End If
        If txtAccountNo & "" <> Empty Then
            cmdInvSeqNo.Enabled = True
        Else
            cmdInvSeqNo.Enabled = False
        End If
        'txtAccountNo.Visible = True
End Sub

Private Sub cboAccountNo_DropDown()
    On Error GoTo ChkErr
    Dim rs2A As New ADODB.Recordset
    
        If gilBankCode.GetCodeNo <> "" Then
            cboAccountNo.Clear
            If Not GetRS(rs2A, "Select AccountNo From " & GetOwner & "SO002A Where CustId = " & gCustId & " And BankCode = " & gilBankCode.GetCodeNo & " And StopFlag =0") Then Exit Sub
            Do While Not rs2A.EOF
                cboAccountNo.AddItem rs2A("AccountNo") & ""
                rs2A.MoveNext
            Loop
        End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cboAccountNo_DropDown"
End Sub

Private Sub cboAccountNo_Validate(Cancel As Boolean)
    If cboAccountNo.Text <> Empty Then
        cmdInvSeqNo.Enabled = True
        
    Else
        cmdInvSeqNo.Enabled = False
    End If
End Sub

Private Sub chkCustAllot_Click()
    On Error Resume Next
        If chkCustAllot = 1 Then
            txtPeriod.Enabled = False
        Else
            txtPeriod.Enabled = chkCustAllot.Enabled
        End If
End Sub

Private Sub chkStopFlag_Click()
    If chkStopFlag.Value = 1 Then
        gdaCeaseDate.SetValue RightDate
        gdaCeaseDate.Enabled = True
    Else
        gdaCeaseDate.SetValue ""
        gdaCeaseDate.Enabled = False
    End If
End Sub

Private Sub cmdBPCode_Click()
    On Error Resume Next
'        If gilCitemCode.GetCodeNo = "" Or txtContNo = "" Then Exit Sub
        If EditMode = giEditModeInsert Then
            If gilCitemCode.GetCodeNo = Empty Then
                MsgBox "請先指定收費項目", vbInformation, "警告"
                Exit Sub
            End If
            If txtFaciSNo.Tag = "" Then
                MsgBox "請先指定設備", vbInformation, "警告"
                Exit Sub
            End If
        End If
        
        With frmSO1131F
            blnSO1131FSave = False
            .uParentForm = Me
            .uCitemCode = gilCitemCode.GetCodeNo
            .uFaciSeqNo = txtFaciSNo.Tag
            .uContNo = txtContNo
            .uClctDate = gdaClctDate.GetValue(True)
            '#6908 2014/10/28 Jacky 調整
            If rs.RecordCount > 0 Then
                .uPenalType = IIf(IsNull(rs("PenalType")), Empty, rs("PenalType"))
                .uExpiretype = IIf(IsNull(rs("ExpireType")), Empty, rs("ExpireType"))
                .uServiceType = rs("ServiceType")   '#5072 依傳來的為主RecordSet
                .uCustId = rs("CustId")             '#5072 依傳來的為主RecordSet
            Else
                .uServiceType = strServiceType
                .uCustId = gCustId
            End If
            '#4165 要多傳FaciSNo,以便SO003A可以儲存 By Kin 2008/10/23
            .uFaciSNo = IIf(txtFaciSNo.Text = Empty, Empty, txtFaciSNo.Text)
            .Show vbModal
        End With
        If blnSO1131FSave Then
            '#4150 將畫面同步 By Kin 2008/10/17
            '#4338 如果是新約則帶次收費日 By Kin 2009/03/12
            '#5996 測試不OK,要記錄位置,不然位置會跑掉,存檔會出現收費項目已存在 By Kin 2011/06/29
            Dim aPos As Variant
            aPos = rs.Bookmark
            If CDate(rs("StartDate") & "") > CDate(rs("StopDate") & "") Then
                Call ShowDiscountDate(rs("ClctDate") & "")
            Else
                Call ShowDiscountDate(rs("StopDate") & "")
            End If
            rs.Bookmark = aPos
        End If
        blnSO1131FSave = False
End Sub

Private Sub cmdChange_Click()
  On Error GoTo ChkErr
'    frmSO1131J.uCitemCode = gilCitemCode.GetCodeNo
'    frmSO1131J.Show 1, Me
'    If strChgCitem <> Empty Then
'        gilCitemCode.SetCodeNo strChgCitem
'        gilCitemCode.Query_Description
'    End If
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdChange_Click")
End Sub

Private Sub cmdChooseAcc_Click()
    On Error Resume Next
        Dim rs2AD As New ADODB.Recordset
        With frmSO1131D
            .uParentForm = Me
            .uCustId = Val(gCustId)
            .Show vbModal
            '#6316 測試不ok,如果有指定帳唬要自動把InvseqNo帶入 By Kin 2012/10/02
            If Len(.uAccountNo) > 0 Then
                Dim aSQL As String
                cboAccountNo.Text = .uAccountNo
                aSQL = "SELECT InvSeqNo FROM " & GetOwner & "SO002AD " & _
                        " WHERE CUSTID = " & gCustId & _
                        " AND COMPCODE = " & gCompCode & _
                        " AND AccountNo ='" & .uAccountNo & "'"
                If Not GetRS(rs2AD, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                If Not rs2AD.EOF Then
                    strInvSeqNo = rs2AD("InvSeqNo")
                    OpenInvTitle
                End If
            End If
        End With
        Call CloseRecordset(rs2AD)
End Sub

Private Sub cmdContract_Click()
    On Error GoTo ChkErr
      '#3732 新增合約歷程 By Kin 2008/06/16
        Screen.MousePointer = vbHourglass
        With frmSO1131I
            '.uServiceType = strServiceType
            .Show 1
        End With
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdContract_Click")
End Sub

Private Sub cmdHouse_Click()
'    On Error Resume Next
'        With frmSO1131E
'            .uParentForm = Me
'            .uCustId = Val(rs("CustID"))
'            .uEditMode = lngEditMode
'            .uDefFaciSeqNo = txtFaciSNo.Tag
'            .uPRCanClick = True
'            .uServiceType = rs("ServiceType")   '#5072 依RecordSet的ServiceType為主 By Kin 2009/05/06
'            .Show vbModal
'            If lngEditMode <> giEditModeView Then
'                If .uFaciSeqNo <> "" Then
'                    txtFaciSNo.Tag = .uFaciSeqNo
'                    txtFaciSNo.Text = .uFaciSNo
'                    txtFaciSNo.ToolTipText = "設備流水號:" & txtFaciSNo.Tag
'                    '#5608 如果有指定設備要以設備的資訊為主
'                    Call GetSO004Data
'                End If
'            End If
'        End With

    On Error Resume Next
    Dim strRefFaci As String
        With frmSO1131E
            .uParentForm = Me
            .uCustId = Val(rs("CustID"))
            .uEditMode = lngEditMode
            .uDefFaciSeqNo = txtFaciSNo.Tag
            .uPRCanClick = True
            .uServiceType = rs("ServiceType")   '#5072 依RecordSet的ServiceType為主 By Kin 2009/05/06
            strRefFaci = GetRefFaci(gilCitemCode.GetCodeNo)
            '2016/09/02 Jacky 7285
            If strRefFaci <> "" Then
                .uFilter = "A.FaciCode in (" & strRefFaci & ")"
            End If
            .Show vbModal
            If lngEditMode <> giEditModeView Then
                If .uFaciSeqNo <> "" Then
                    txtFaciSNo.Tag = .uFaciSeqNo
                    txtFaciSNo.Text = .uFaciSNo
                    txtFaciSNo.ToolTipText = "設備流水號:" & txtFaciSNo.Tag
                    '#5608 如果有指定設備要以設備的資訊為主
                    Call GetSO004Data
                End If
            End If
        End With

End Sub
'#3436 進入選擇發票抬頭 By Kin 2007/11/27
Private Sub cmdInvSeqNo_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
    With frmSO1131G
        .uAccountNo = cboAccountNo.Text
        .uCompCode = gCompCode
        .uCustId = gCustId
        .uParentForm = Me
        .uInvSeqNo = IIf(strInvSeqNo & "" = "", "", strInvSeqNo)
        .Show 1
    End With
    
    strInvSeqNo = frmSO1131G.uInvSeqNo
    Call OpenInvTitle
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdInvSeqNo_Click"
End Sub

Private Sub cmdOpen_Click()
    On Error GoTo ChkErr
    Dim lngAffected As Long
        
        If gilCitemCode.GetCodeNo = "" Or txtContNo = "" Then Exit Sub
        If MsgBox("是否重新開啟多階優惠??", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then Exit Sub
        gcnGi.Execute "Update " & GetOwner & "SO003A Set StopFlag = 0,StopDate = null Where CustId = " & _
                gCustId & " And CompCode = " & gCompCode & " And ServiceType = '" & _
                strServiceType & "' And CitemCode = " & gilCitemCode.GetCodeNo & " And ContNo = '" & txtContNo & "' And FaciSeqNo = '" & txtFaciSNo.Tag & "' And StopFlag = 1", lngAffected
        MsgBox "開啟多階優惠完畢,成功共" & lngAffected & "筆!!", vbInformation, gimsgPrompt
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdOpen_Click"
End Sub

Private Sub cmdShowRate_Click()
  On Error GoTo ChkErr
    If fraRate.Visible Then
        fraRate.Visible = False
        cmdShowRate.Caption = "F7. 顯示費率表 >>"
        Exit Sub
    Else
        fraRate.Visible = True
        cmdShowRate.Caption = "<< F7. 隱藏費率表"
    End If
    If strMduId <> "" Then
        If Not GetRS(rsTmp, "SELECT Name FROM " & GetOwner & "SO017 WHERE MduId='" & strMduId & "'") Then Exit Sub
        lblMduName.Caption = "( " & rsTmp("Name") & " )"
        subGrd1 strMduId
        subGrd3
    Else
        If CStr(intClassCode) = "" Then
            MsgBox "客戶資料檔中此客戶編號無類別!!", , "訊息.."
            Exit Sub
        Else
            If Not GetRS(rsTmp, "SELECT Description FROM " & GetOwner & "CD004 WHERE CodeNo=" & intClassCode) Then Exit Sub
            If rsTmp("Description") <> "" Then lblCustClass.Caption = "( " & rsTmp("Description") & " )" & ""
            If Not GetRS(rsTmp, "SELECT CD019.Description AS ClassName,CD019CD004.Period,CD019CD004.Amount,CD019CD004.ClassCode,CD019CD004.CitemCode FROM " & GetOwner & "CD019CD004," & GetOwner & "CD019 WHERE CD019CD004.ClassCode=" & intClassCode & " AND CD019CD004.CitemCode=CD019.CodeNo And CD019.ServiceType In (Null,'" & strServiceType & "')") Then Exit Sub
            If Not rsTmp.EOF Then
               subGrd2 intClassCode '有資料
               subGrd1 strMduId
            Else
               subGrd3 '無資料
               subGrd1 strMduId
            End If
        End If
    End If
    
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdShowRate_Click")
End Sub


Private Function subGrd1(ByVal strMduId As String)
  On Error GoTo ChkErr
  Dim mFlds As New GiGridFlds
    If rsMdu.State = adStateOpen Then rsMdu.Close
    rsMdu.CursorLocation = adUseClient
    rsMdu.Open "SELECT CD019.Description,CD019SO017.Period,CD019SO017.Amount,CD019SO017.CitemCode FROM " & GetOwner & "CD019SO017," & GetOwner & "CD019 WHERE MduId='" & strMduId & "' AND CD019SO017.CitemCode=CD019.CodeNo", gcnGi, adOpenForwardOnly, adLockReadOnly
    mFlds.Add "Description", , , , , "收費項目  ", vbLeftJustify
    mFlds.Add "Period", , , , , "期數 ", vbRightJustify
    mFlds.Add "Amount", , , , , "金額         ", vbRightJustify
    With ggrMduRate
         .AllFields = mFlds
         Set .Recordset = rsMdu
         .Refresh
    End With

    rsMdu.Close
    Set rsMdu = Nothing
  Exit Function
ChkErr:
  ErrSub Me.Name, "subGrd1"
End Function

Private Sub subGrd2(intClassCode As Long)
  On Error GoTo ChkErr
  Dim mFlds As New GiGridFlds
      If rsCust.State = adStateOpen Then rsCust.Close
      rsCust.CursorLocation = adUseClient
       rsCust.Open "SELECT CD019.Description ,CD019CD004.Period,CD019CD004.Amount,CD019CD004.ClassCode,CD019CD004.CitemCode FROM " & GetOwner & "CD019CD004," & GetOwner & "CD019 WHERE CD019CD004.ClassCode=" & intClassCode & " AND CD019CD004.CitemCode=CD019.CodeNo", gcnGi, adOpenForwardOnly, adLockReadOnly

      mFlds.Add "Description", , , , , "收費項目  ", vbLeftJustify
      mFlds.Add "Period", , , , , "期數 ", vbRightJustify
      mFlds.Add "Amount", , , , , "金額         ", vbRightJustify
      With ggrCustRate
           .AllFields = mFlds
           Set .Recordset = rsCust
           .Refresh
      End With
      rsCust.Close
      Set rsCust = Nothing
  Exit Sub
ChkErr:
  ErrSub Me.Name, "subGrd2"
End Sub

Private Sub subGrd3()
  On Error GoTo ChkErr
  Dim mFlds As New GiGridFlds
      If rsCust.State = adStateOpen Then rsCust.Close
      rsCust.CursorLocation = adUseClient
      rsCust.Open "SELECT Description,Period,Amount FROM " & GetOwner & "CD019 WHERE PeriodFlag=1 And ServiceType in (Null,'" & strServiceType & "')", gcnGi, adOpenForwardOnly, adLockReadOnly

      mFlds.Add "Description", , , , , "收費項目  ", vbLeftJustify
      mFlds.Add "Period", , , , , "期數 ", vbRightJustify
      mFlds.Add "Amount", , , , , "金額         ", vbRightJustify
      With ggrCustRate
           .AllFields = mFlds
           Set .Recordset = rsCust
           .Refresh
      End With
      rsCust.Close
      Set rsCust = Nothing
  Exit Sub
ChkErr:
  ErrSub Me.Name, "subGrd3"
End Sub

Private Sub cmdStop_Click()
    On Error GoTo ChkErr
    Dim lngAffected As Long
        If gilCitemCode.GetCodeNo = "" Or txtContNo = "" Then Exit Sub
        If MsgBox("是否關閉多階優惠??", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then Exit Sub
        gcnGi.Execute "Update " & GetOwner & "SO003A Set StopFlag = 1,StopDate = " & GetNullString(RightDate, giDateV) & " Where CustId = " & _
                gCustId & " And CompCode = " & gCompCode & " And ServiceType = '" & _
                strServiceType & "' And CitemCode = " & gilCitemCode.GetCodeNo & " And ContNo = '" & txtContNo & "' And FaciSeqNo = '" & txtFaciSNo.Tag & "' And StopFlag = 0", lngAffected
        MsgBox "關閉多階優惠完畢,成功共" & lngAffected & "筆!!", vbInformation, gimsgPrompt
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdStop_Click"
End Sub

Private Sub cmdViewChange_Click()
    On Error Resume Next
        With frmSO1132J
            '.uServiceType = strServiceType
            .uType = 2
            .Show vbModal
        End With
End Sub

Private Sub Form_Activate()
    On Error GoTo ChkErr
    Dim intCount As Integer
        'If Me.Tag = "" Then Call FormKeepPosition(Me, varC): Me.Tag = "Loaded"
        If blnChgCitem Then
            If Not rs.EOF Then
                If fraData.Enabled = True And Val(rs("StopType") & "") = 0 Then gilCitemCode.SetFocus
            End If
        End If
        MenuEnabled , , , lngEditMode <> giEditModeView, , True
        Set ObjActiveForm = Me
        strActFrmName = "frmSO1131C"
        Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
    If Err.Number = 0 Then
        If intCount < 10 Then
            Resume 0
            intCount = intCount + 1
        Else
            Resume Next
        End If
    Else
        ErrSub Me.Name, "FormActivate"
    End If
End Sub

Private Sub OpenData()
    On Error GoTo ChkErr
        Dim aSQL As String
        
        If rsTmp.State = adStateOpen Then rsTmp.Close
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open "SELECT CustName,ClassCode1,MduId,CompCode FROM " & GetOwner & "SO001 WHERE CustId=" & gCustId, gcnGi, adOpenForwardOnly, adLockReadOnly
        intClassCode = rsTmp("ClassCode1") & "" '客戶類別代碼
        strMduId = rsTmp("MduId") & ""  '大樓編號
        strCustName = rsTmp("CustName") & ""
        lngCompCode = rsTmp("CompCode")
        
        '#4448 判斷是否由阿貴傳來的Recordset,如果是以傳來的為主 By Kin 2009/04/08
        If (rsFROM Is Nothing) Or (rsFROM.State = adStateClosed) Then
            Dim aRowId As String
            aRowId = ""
            If frmSO1100BMDI.rsSO003.RecordCount > 0 Then
                aRowId = frmSO1100BMDI.rsSO003("RowId")
            End If
            Dim aTableName As String
            aTableName = "SO003"
            If aRowId <> "" Then
                If Val(frmSO1100BMDI.rsSO003("VOId") & "") = 1 Then
                    aTableName = "SO003Log"
                End If
            End If
            
            aSQL = "Select A.Rowid,A.* From " & GetOwner & aTableName & " A Where RowId = '" & aRowId & "' "
            If Not GetRS(rs, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
            'Set rs = frmSO1100BMDI.rsSO003
        Else
            Set rs = rsFROM
        End If
        
        '#3436 顯示發票抬頭資料 By Kin 2007/11/27
        If Not rs.EOF Then
            If rs("InvSeqNo") & "" <> "" Then
                If Not GetRS(rsSO138, "Select * From " & GetOwner & "SO138 Where InvSeqNo=" & rs("InvSeqNo") & "" & " And Nvl(StopFlag, 0) = 0") Then Exit Sub
                If rsSO138.EOF Then
                    strInvSeqNo = Empty
                Else
                    strInvSeqNo = rsSO138("InvSeqNo") & ""
                End If
            Else
                strInvSeqNo = Empty
            End If
        Else
            strInvSeqNo = Empty
        End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub
Private Function chkCitemCodeOK(ByVal aCode As String, ByVal aUpdEn As String) As Boolean
  On Error GoTo ChkErr
    Dim aSQL  As String
    aSQL = " Select Count(*) from " & GetOwner & "CD019D " & _
                " Where CitemCode = " & aCode & _
                " And UserId = '" & aUpdEn & "' "
    If Val(GetRsValue(aSQL, gcnGi) & "") > 0 Then
        chkCitemCodeOK = False
    Else
        chkCitemCodeOK = True
    End If
  Exit Function
ChkErr:
  ErrSub Me.Name, "chkCitemCodeOK"
End Function
Private Function IsDataOk() As Boolean
    On Error GoTo ChkErr
    '檢查必要欄位是否有值, 若為必要欄位且無值, 則顯示"欄位必需有值", 且focus移到該欄位上
        IsDataOk = False
        
        If Not ChkDTok Then Exit Function
        If Not MustExist(gilCitemCode, 2, "收費項目") Then Exit Function
'        If Not chkCitemCodeOK(gilCitemCode.GetCodeNo, garyGi(0)) Then
'            MsgBox "該使用者無法選取此收費項目！", vbCritical, "警告"
'            gilCitemCode.SetFocus
'            Exit Function
'        End If
        If Not MustExist(txtPeriod, , "每月期數") Then Exit Function
        If Val(txtPeriod.Text) <= 0 Then MsgBox "每月期數須大於0": txtPeriod.SetFocus: Exit Function
        
        If Not MustExist(ginAmount, , "收費金額") Then Exit Function
        If Val(ginAmount.Text) = 0 Then
            If MsgBox("收費金額為0,是否要存檔??", vbYesNo + vbDefaultButton2 + vbQuestion, gimsgPrompt) = vbNo Then ginAmount.SetFocus: Exit Function
        End If
        
        If Not MustExist(gdaStartDate, , "有效期限起始日") Then Exit Function
        If Not MustExist(gdaStopDate, , "有效期限截止日") Then Exit Function
        
        If Not MustExist(gdaClctDate, , "次收費日") Then Exit Function
        
        If gdaStopDate.GetValue <= gdaStartDate.GetValue Then MsgBox "有效期限截止日需大於有效期限起始日": gdaStopDate.SetFocus: Exit Function
        If gdaClctDate.GetValue < gdaStartDate.GetValue Then MsgBox "次收費日需大於等有效期限截止日": gdaStopDate.SetFocus: Exit Function
        If GetByHouse(gilCitemCode.GetCodeNo) = 1 Then
            txtFaciSNo.Tag = gCustId
        Else
            If txtFaciSNo = "" Then
                MsgBox "該收費項目為 BY機收費項目,請選擇設備!!", vbInformation, gimsgPrompt
                cmdHouse.Value = True
                Exit Function
            End If
        End If
        Dim strWhere As String
        strWhere = "CustId = " & gCustId & " And CitemCode = " & gilCitemCode.GetCodeNo & " And FaciSeqNo = '" & txtFaciSNo.Tag & "' And CompCode=" & gCompCode & " And ServiceType = '" & strServiceType & "'"
        If lngEditMode = giEditModeEdit Then strWhere = strWhere & " And RowId <> '" & rs("RowId") & "'"
        
        If ChkDupData(strWhere, "SO003") Then
            MsgBox "該收費項目已存在!!", vbCritical, gimsgPrompt
            Exit Function
        End If
        If gilBankCode.GetCodeNo = "" And cboAccountNo.Text <> "" Then
            MsgBox "銀行代碼無值,帳號資料將會清除!!", vbInformation, gimsgPrompt
            cboAccountNo = ""
        End If
        '優待辦法有值,其他優待欄位需有值 93/05/18
        If gilDiscount.GetCodeNo <> "" Then
            If Not MustExist(ginDiscountAmt, , "優待辦法有值,優待金額") Then Exit Function
            If Not MustExist(gdaDiscountDate1, 1, "優待辦法有值,優待期限起始日") Then Exit Function
            If Not MustExist(gdaDiscountDate2, 1, "優待辦法有值,優待期限截止日") Then Exit Function
        End If
        If gdaClctDate.Text <> Format(CStr(CDate(gdaStopDate.Text) + Val(rsSO043.Fields("para5"))), "yyyy/mm/dd") Then
          MsgBox "截止日與次收費日不符合[SO043收費參數檔]中之定義", , "訊息.."
          Exit Function
        End If
        '若該收費方式的欄位沒有設定任何資料請直接預設帶入 "其他參數設定"[SO044. CMCode]
        If Not MustExist(gilCMCode, 2, "收費方式") Then
            If Not GetRS(rsTmp, "Select CMCode,CMName From " & GetOwner & "SO002 Where Custid = " & gCustId & " And CompCode = " & gCompCode & " And ServiceType = '" & strServiceType & "'") Then Exit Function
            gilCMCode.SetCodeNo rsTmp("CMCode") & ""
            gilCMCode.Query_Description
            gilCMCode.SetFocus
            Exit Function
        End If
        If chkStopFlag.Value = 1 Then If Not MustExist(gdaCeaseDate, 1, "停用日期") Then Exit Function
        If Not ChkSign Then Exit Function
        
        
        
        '******************************************************************
        '#3436 如果有帳號資料一定要有發票抬頭資料 By Kin 2007/11/27
'        If cboAccountNo.Text <> "" Then
'            If strInvSeqNo = Empty Then
'                MsgBox "請選擇發票抬頭資料", vbInformation, "訊息"
'                cmdInvSeqNo.SetFocus
'                Exit Function
'            Else
'                Dim strQry As String
'                strQry = "Select Count(*) Cnt From " & GetOwner & "SO002AD" & _
'                         " Where AccountNo='" & cboAccountNo.Text & "'" & _
'                         " And CompCode=" & gCompCode & _
'                         " And CustId=" & gCustId & _
'                         " And InvSeqNo=" & strInvSeqNo
'
'                If gcnGi.Execute(strQry)(0) = 0 Then
'                    MsgBox "發票抬頭資料不正確!!請重新選擇發票抬頭資料", vbInformation, "訊息"
'                    cmdInvSeqNo.SetFocus
'                    Exit Function
'                End If
'            End If
'        Else
'            '#3852 清空帳號時要連InvSeqNO一併清除 By Kin 2008/04/09
'            strInvSeqNo = Empty
'        End If
        '******************************************************************
        IsDataOk = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Function ChkSign() As Boolean
    On Error GoTo ChkErr
    Dim strMsg As String
    Dim strSign As String
  
        strSign = GetRsValue("Select Sign From " & GetOwner & "CD019 Where CodeNo = " & gilCitemCode.GetCodeNo) & ""
        strMsg = "借貸符號為 '" & strSign & "',只能輸入'" & strSign & "' 金額"
        
        If Val(ginAmount.Value) <> Abs(Val(ginAmount.Value)) * Val(strSign & "1") Then
            ginAmount.Text = Abs(Val(ginAmount.Value)) * Val(strSign & "1")
            ginAmount.SetFocus
            MsgBox strMsg, vbExclamation, gimsgPrompt: Exit Function
        End If
        If Val(ginDiscountAmt.Value) <> Abs(Val(ginDiscountAmt.Value)) * Val(strSign & "1") Then
            ginDiscountAmt.Text = Abs(Val(ginDiscountAmt.Value)) * Val(strSign & "1")
            ginDiscountAmt.SetFocus
            MsgBox strMsg, vbExclamation, gimsgPrompt: Exit Function
        End If
        
    
    ChkSign = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "ChkSign"
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ChkErr
        If Not ChkGiList(KeyCode) Then Exit Sub
        Select Case KeyCode
               Case vbKeyEscape
                    Unload Me
               Case vbKeyF2
                    If lngEditMode <> giEditModeView Then Call UpdateGo
    '                If cmdSave.Enabled = True Then cmdSave.Value = True
               Case vbKeyF7
                    If cmdShowRate.Enabled = True Then cmdShowRate.Value = True
        End Select
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
'    If Crm Then
'        PolyFrmFunction Me
'        Me.Move frmCrm.trv.Width + 66, frmCrm.lvwItem.Height + 3010
'        frmCrm.picService.Enabled = False
'    Else
'
'        If Not800600 Then
'            Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
'        End If
'    End If
    'Me.Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2)
    'Me.Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2)
    If Crm Then
        PolyFrmFunction Me
        Me.Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2)
    Else

        'If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 + 300
        If Not 800600 Then Me.Move (Screen.Width - Me.Width) / 2, frmSO1100BMDI.Top + frmSO1100BMDI.tbrMenu.Height + 700
    End If
    If strServiceType = "" Then strServiceType = gServiceType
    If gServiceType <> strServiceType Then gServiceType = strServiceType
    sstDetail.Tab = 0
    Call InitData
    Call SetCombol
    
    Call cboAccountNo_DropDown
    Call OpenData
    Call SubGil
    Call ChangeMode(lngEditMode)
    Call cboAccountNo_Click
    txtOrderNo.MaxLength = rs.Fields("OrderNo").DefinedSize
    txtContNo.MaxLength = rs.Fields("ContNo").DefinedSize
    sstDetail.TabVisible(2) = GetRsValue("Select Nvl(OTTCompid,-1) From " & GetOwner & "SO041") > -1
    '#5608 參數控制是否顯示 By Kin 2010/04/26
    EnabledCtl
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub SetCombol()
    On Error Resume Next
        cboPenalType.AddItem "1.回溯牌價"
        cboPenalType.AddItem "2.優惠價"
        cboPenalType.AddItem "3.優惠平均價"
        '#6865 增加4.應收金額 By Kin 2014/09/09
        cboPenalType.AddItem "4.應收/實收金額"
        cboExpireType.AddItem "1.恢復最新公告牌價"
        cboExpireType.AddItem "2.以最後一階繼續優惠"
        '************************************************
        '#3732 多增加以下兩個選項 By Kin 2008/06/09
        '#4187 增加5.約滿不跨階 By Kin 2008/10/30
        cboExpireType.AddItem "3.到期不出"
        cboExpireType.AddItem "4.接新約"
        cboExpireType.AddItem "5.約滿不跨階"
        '************************************************
    
End Sub

Private Sub InitData()
    On Error Resume Next
    Dim intStatusBarH As Long
    Dim objParentForm As Object
        If Not GetSystemPara(rsSO043, gCompCode, strServiceType, "SO043") Then Exit Sub
        If Not frmSO1100BMDI Is Nothing Then
            Set objParentForm = frmSO1100BMDI
            If objParentForm.Name = "frmSO1100BMDI" Then intStatusBarH = 390
        End If
        'Me.Move objParentForm.Left, objParentForm.Top + objParentForm.Height - Me.Height - intStatusBarH
End Sub

Private Sub SubGil()
    On Error GoTo ChkErr
        SetgiList gilCitemCode, "CodeNO", "Description", "CD019", , , , , , , "Where PeriodFlag=1 ", True
        Call GiListFilter(gilCitemCode, strServiceType)
        'CD019.RefNo=19要被過濾掉 By Kin 2008/12/11
        gilCitemCode.Filter = gilCitemCode.Filter & " And PeriodFlag=1 And Nvl(RefNo,0)<>19"
        '# 7334 to filter what citemcode can't be chosen by kin 2016/12/15
        Dim containCode As String
        containCode = "-1"
        If EditMode = giEditModeEdit Then
             If rs.State <> adStateClosed Then
                If rs.RecordCount > 0 Then
                    If Len(rs("CitemCode") & "") > 0 Then
                        containCode = rs("CitemCode")
                    End If
                End If
            End If
        End If
        gilCitemCode.Filter = gilCitemCode.Filter & "  And (CodeNo Not in (Select CitemCode From " & GetOwner & "CD019D Where UserId = '" & garyGi(0) & "' )  Or CodeNo = " & containCode & " )"
       
        Call SetgiList(gilDiscount, "CodeNo", "Description", "CD048", , , , , 3, 30, , True)
        Call SetgiList(gilPackageNo, "CodeNo", "Description", "CD027", , , , , 3, 30, , True)
        Call SetgiList(gilCMCode, "CodeNo", "Description", "CD031", , , , , 3, 30, "Where ServiceType='" & strServiceType & "' or ServiceType is null", True)
        Call SetgiList(gilBankCode, "BankCode", "BankName", "SO002A", , , , , 3, 30, "Where CustId = " & gCustId & " And CompCode =  " & gCompCode & " And " & GetBankCodeRowId, True)
        Call SetgiList(gilPTCode, "CodeNo", "Description", "CD032", , , , , 10, 20, "Where ServiceType='" & strServiceType & "' or ServiceType is null", True)
        Call SetgiList(gilBPCode, "CodeNo", "Description", "CD078", , , , , 10, 20, "Where ServiceType='" & strServiceType & "' or ServiceType is null", True)
        Call SetgiList(gilPromCode, "CodeNo", "Description", "CD042", , , , , 10, 20, "Where ServiceType='" & strServiceType & "' or ServiceType is null", True)
        Call SetgiList(gilSalePoint, "CodeNo", "Description", "CD107", , , , , 3, 30, , True)
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ini_gilCitemCode")
End Sub

Private Function GetBankCodeRowId() As String
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim strRowId As String
    Dim strBankCode As String
        If Not GetRS(rs, "Select RowId,BankCode From " & GetOwner & "SO002A Where CustId = " & gCustId & " And CompCode =  " & gCompCode & " And StopFlag =0 Order By BankCode") Then Exit Function
        Do While Not rs.EOF
            If strBankCode <> rs("BankCode") & "" And rs("BankCode") & "" <> "" Then
                strRowId = strRowId & ",'" & rs("RowId") & "'"
            End If
            strBankCode = rs("BankCode") & ""
            rs.MoveNext
        Loop
        If strRowId = "" Then
            strRowId = " 1 = 1 "
        Else
            strRowId = " RowId In (" & Mid(strRowId, 2) & ")"
        End If
        GetBankCodeRowId = strRowId
End Function

Private Sub NewRcd()
  On Error Resume Next
    Call gilCitemCode.SetCodeNo("")
    Call gilCitemCode.SetDescription("")
    gilPTCode.Clear
    gilPackageNo.Clear
    txtPeriod.Text = ""
    ginAmount.Text = ""
    gdaStartDate.Text = ""
    gdaStopDate.Text = ""
    gdaClctDate.Text = ""
    gilDiscount.Clear
    ginDiscountAmt.Text = 0
    gdaDiscountDate1.SetValue ""
    gdaDiscountDate2.SetValue ""
    ginBudgetPeriod.Text = 0
    '#7460 Clear the field of CardBillNo By Kin 2017/06/29
    txtCardBillNo.Text = ""
    '#7626
    txtFonsOrderNo.Text = ""
    txtiOSOrderNo.Text = ""
    'If gilBankCode.GetListCount = 0 Then
        ' 93/12/24 調整問題集 1357
        If Not GetRS(rsTmp, "Select CMCode From " & GetOwner & "SO002 Where Custid = " & gCustId & " And CompCode = " & gCompCode & " And ServiceType = '" & strServiceType & "'") Then Exit Sub
        If Not rsTmp.EOF Then
            gilCMCode.SetCodeNo rsTmp("CMCode") & ""
            gilCMCode.Query_Description
        End If
    'End If
    chkCustAllot = 0
    gilPTCode.ListIndex = 1
    'gilCMCode.Clear
    'gilBankCode.Clear
    chkBundle = 0
    chkPenal = 0
    chkMergePrint = 0
    txtContNo = ""
    '#4056 增加訂單單號 By Kin 2008/08/20
    txtOrderNo.Text = ""
    gdaContStartDate.SetValue ""
    gdaContStopDate.SetValue ""
    gilBPCode.Clear
    gilPromCode.Clear
    txtBundleMon = ""
    cboPenalType = ""
    cboExpireType = ""
    chkStopFlag.Value = ""
    gdaCeaseDate.SetValue ""
    gilCitemCode.SetFocus
    
    '#3436 By Kin 2007/11/27
    strInvSeqNo = ""
    If txtFaciSNo.Tag <> "" Then
        Call GetSO004Data
    End If
    Call OpenInvTitle
End Sub

Public Sub OpenRS()
  On Error GoTo ChkErr
    Set rs = frmSO1100BMDI.rsSO003
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "OpenRS")
End Sub

Private Sub RcdToScr()
  On Error GoTo ChkErr
    Call gilCitemCode.SetCodeNo(rs("CitemCode") & "")
    Call gilCitemCode.SetDescription(rs("CitemName") & "")
    txtPeriod.Text = rs("Period") & ""
    ginAmount.Text = rs("Amount") & ""
    gdaStartDate.Text = rs("StartDate") & ""
    gdaStopDate.Text = rs("StopDate") & ""
    gdaClctDate.Text = rs("ClctDate") & ""
    gilDiscount.SetCodeNo rs("DiscountCode") & ""
    gilDiscount.SetDescription rs("DiscountName") & ""
    gdaDiscountDate1.SetValue rs("DiscountDate1") & ""
    gdaDiscountDate2.SetValue rs("DiscountDate2") & ""
    ginDiscountAmt.Text = rs("DiscountAmt") & ""
    '套餐 91/09/14
    gilPackageNo.SetCodeNo rs("PackageNo") & ""
    gilPackageNo.SetDescription rs("PackageName") & ""
    gilBankCode.SetCodeNo rs("BankCode") & ""
    gilBankCode.SetDescription rs("BankName") & ""
    gilPTCode.SetCodeNo rs("PTCode") & ""
    gilPTCode.SetDescription rs("PTName") & ""
    cboAccountNo = rs("AccountNo") & ""
    ginBudgetPeriod.Text = rs("BudgetPeriod") & ""
    chkCustAllot = rs("CustAllot")
    txtFaciSNo.Tag = rs("FaciSeqNo") & ""
    txtFaciSNo.Text = rs("FaciSNo") & ""
    txtFaciSNo.ToolTipText = "設備流水號:" & txtFaciSNo.Tag
    gilCMCode.SetCodeNo rs("CMCode") & ""
    gilCMCode.SetDescription rs("CMName") & ""
    
    txtSTBSeqNo.Text = rs("STBSeqNo") & ""
    gilSalePoint.SetCodeNo rs("SalePointcode") & ""
    gilSalePoint.SetDescription rs("SalePointname") & ""
    '#7460 display the field of  CardBillNo  By Kin 2017/06/29
    txtCardBillNo.Text = rs("CardBillNo") & ""
    '#7626
    txtFonsOrderNo.Text = rs("FVOrderNo") & ""
    txtiOSOrderNo.Text = rs("iOSOrderNo") & ""
    '#4306 增加出單到期日 By Kin 2009/03/25
    If Not IsNull(rs("ClctStopDate")) Then
        gdaClctStopDate.SetValue rs("ClctStopDate") & ""
    End If
    Call gilCitemCode_Change
    
    chkBundle = rs("Bundle") & ""
    chkPenal = rs("Penal") & ""
    chkMergePrint = rs("MergePrint") & ""
    txtContNo = rs("ContNo") & ""
    '原本只顯示停用,現在改為判斷StopType By Kin 2008/12/11
    Select Case Val(rs("StopType") & "")
        Case 0
            chkStopFlag.Caption = "停用"
        Case 1
            chkStopFlag.Caption = "暫停停用"
        Case 2
            chkStopFlag.Caption = "結清停用"
        Case 3
            chkStopFlag.Caption = "促變中停用"
        Case Else
            chkStopFlag.Caption = "停用"
    End Select
    
    '#4056 增加訂單單號
    '#4121 改抓SO003A.OrderNo By Kin 2008/10/01
    'txtOrderNo.Text = rs("OrderNo") & ""
    'txtDiscountAmt = rs("DiscountAmt") & ""
    '#4121 改抓SO003A的資料 By Kin 2008/10/01
    'gdaContStartDate.SetValue rs("ContStartDate") & ""
    'gdaContStopDate.SetValue rs("ContStopDate") & ""
    '#4121 改抓SO003A的資料 By Kin 2008/10/01
    'gilBPCode.SetCodeNo rs("BPCode") & ""
    'gilBPCode.SetDescription rs("BPName") & ""
    '#4121 改抓SO003A的資料 By Kin 2008/10/01
'    gilPromCode.SetCodeNo rs("PromCode") & ""
'    gilPromCode.SetDescription rs("PromName") & ""
    '#4121 改抓SO003A的資料 By Kin 2008/10/01
    'txtBundleMon = rs("BundleMon") & ""
    '#6971 如果是空值要給空白 By Kin 2015/01/05
    If Len(rs("PenalType") & "") = 0 Then
        cboPenalType.Text = ""
    Else
        cboPenalType.ListIndex = Val(rs("PenalType") & "") - 1
    End If
    If Len(rs("ExpireType") & "") = 0 Then
        cboExpireType.Text = ""
    Else
        cboExpireType.ListIndex = Val(rs("ExpireType") & "") - 1
    End If
    chkStopFlag.Value = Val(rs("StopFlag") & "")
    gdaCeaseDate.SetValue rs("CeaseDate") & ""
    'Call ShowDiscountDate(rs("ClctDate") & "")
    '#4338 如果是新約則帶入次收費日 By Kin 2009/03/12
    If CDate(rs("StartDate") & "") > CDate(rs("StopDate") & "") Then
        Call ShowDiscountDate(rs("ClctDate") & "")
    Else
        Call ShowDiscountDate(rs("StopDate") & "")
    End If
    gdaDiscountDate2.SetValue rs("DiscountDate2") & ""
    '#6841 增加長繳別註記 By Kin 2014/08/19
    chkLongPay.Value = 0
    If Val(rs("LONGPAYflag") & "") = 1 Then
        chkLongPay.Value = 1
    End If
    '#7605 By Kin 2017/09/26
    chkPeriodFlag.Value = 0
    If Val(rs("PeriodFlag") & "") = 1 Then
        chkPeriodFlag.Value = 1
    End If
    Call OpenInvTitle
    '***********************************************************
    '#4150 增加權限控管置換功能是否顯示 By Kin 2008/10/17
    
    'If Not blnChgCitem Then gilCitemCode.Enabled = False
    'cmdChange.Visible = False
    'cmdChange.Visible = GetUserPriv("SO1131", "(SO1131I)")
'    If GetUserPriv("SO1131", "(SO1131I)") Then
'        If Not IsNull(rs("BpCode")) Then
'            cmdChange.Visible = True
'        End If
'    End If
    '***********************************************************
    'Call SetCombolListIndex(cboAccountNo, rs("AccountNo") & "")
  Exit Sub
ChkErr:
    
    Call ErrSub(Me.Name, "RcdToScr")
    
End Sub
Private Sub GetSO004Data()
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim rs004 As New ADODB.Recordset
    gilCMCode.SetCodeNo ""
    gilCMCode.Query_Description
    gilPTCode.SetCodeNo ""
    gilPTCode.Query_Description
    gilBankCode.SetCodeNo ""
    gilBankCode.Query_Description
    txtAccountNo.Text = ""
    strInvSeqNo = Empty
    aSQL = "SELECT * FROM " & GetOwner & "SO004 " & _
        " WHERE SEQNO='" & txtFaciSNo.Tag & "'"
    If Not GetRS(rs004, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If Not rs004.EOF Then
        With rs004
            If .Fields("ChCMCode") & "" <> "" Then
                gilCMCode.SetCodeNo .Fields("ChCMCode") & ""
                gilCMCode.Query_Description
            End If
            If .Fields("ChPTCode") & "" <> "" Then
                gilPTCode.SetCodeNo .Fields("ChPTCode") & ""
                gilPTCode.Query_Description
            End If
            If .Fields("AccountNo") & "" <> "" Then
                txtAccountNo.Text = .Fields("AccountNo") & ""
            End If
            If .Fields("BankCode") & "" <> "" Then
                gilBankCode.SetCodeNo .Fields("BankCode") & ""
                gilBankCode.Query_Description
            End If
            strInvSeqNo = .Fields("InvSeqNo") & ""
        End With
    End If
    On Error Resume Next
    Call CloseRecordset(rs004)
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "GetSO004Data")
End Sub
Private Function ExecuteSingleTransfer(ByVal strUpdEn As String, ByVal strCustid As String, _
                        ByVal strCitemCode As String, ByVal strFaciSeqNo As String, ByRef RetMsg As String) As Boolean
    On Error GoTo ChkErr

    Dim P_RETMSG As String
    Dim ComSingleTransfer As New ADODB.Command
    With ComSingleTransfer
        
       .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
       .Parameters.Append .CreateParameter("P_UPDEN", adVarChar, adParamInput, 2000, strUpdEn)
       .Parameters.Append .CreateParameter("P_CUSTID", adVarNumeric, adParamInput, , Val(strCustid))
        .Parameters.Append .CreateParameter("P_CITEMCODE", adVarNumeric, adParamInput, , Val(strCitemCode))
        .Parameters.Append .CreateParameter("P_FACISEQNO", adVarChar, adParamInput, 2000, strFaciSeqNo)
       .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, P_RETMSG)
         Set .ActiveConnection = gcnGi
        .CommandText = "PK_LongPayTransfer.ExecuteSingleTransfer"
        .CommandType = adCmdStoredProc
        .Execute
        P_RETMSG = .Parameters("P_RETMSG").Value & ""
        If Val(.Parameters("Return_value").Value & "") < 0 Then
            MsgBox P_RETMSG, vbCritical, "錯誤"
            ExecuteSingleTransfer = False
            Exit Function
        Else
            '#6865 如果有回傳訊息，Show出回傳訊息 By Kin 2014/09/12
            If Len(P_RETMSG & "") > 0 Then
                MsgBox P_RETMSG, vbInformation, "訊息"
            End If
           ExecuteSingleTransfer = True
        End If
    End With
    ExecuteSingleTransfer = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "ExecuteSingleTransfer"
End Function
Private Function ScrToRcd() As Boolean
    On Error GoTo ChkErr
        '將螢幕上物件內容轉換成資料檔對應欄位內容
        'If strChgCitem <> Empty Then blnChgCitem = True: strChgCitem = Empty
        Dim rsOther As New ADODB.Recordset
        Dim aSQL As String
        Dim strSign As String
        Dim aAmount As Long
        With rs
            If EditMode = giEditModeInsert Then
                .AddNew
                
                .Fields("SeqNo") = GetInvoiceNo3("SO003")
                '組合產品 94/09/01 Jacky 改 Janice 提
'                .Fields("Bundle") = chkBundle
'                .Fields("Penal") = chkPenal
'                .Fields("MergePrint") = chkMergePrint
'                .Fields("ContNo") = NoZero(txtContNo)
'                .Fields("ContStartDate") = NoZero(gdaContStartDate.GetValue(True))
'                .Fields("ContStopDate") = gdaContStopDate.GetValue(True)
'                .Fields("BPCode") = gilBPCode.GetCodeNo2
'                .Fields("BPName") = gilBPCode.GetDescription2
'                .Fields("PromCode") = gilPromCode.GetCodeNo2
'                .Fields("PromCode") = gilPromCode.GetDescription2
'                .Fields("BundleMon") = NoZero(txtBundleMon)
'                .Fields("PenalType") = NoZero(cboPenalType)
'                .Fields("ExpireType") = NoZero(cboExpireType)
            
                .Fields("CustId") = gCustId
                .Fields("CompCode") = gCompCode
                .Fields("ServiceType") = strServiceType
            End If
            '**************************************************************************************************************************
            '#4143  請增加SO1131C 欄位權限。『訂單單號』、『合約編號』、『是否綁約』及『是否違約計算』四個欄位權限 By Kin 2008/10/13
            .Fields("Bundle") = chkBundle.Value
            .Fields("Penal") = chkPenal.Value
            If Trim(txtOrderNo.Text) <> "" Then .Fields("OrderNo") = txtOrderNo.Text
            If Trim(txtContNo.Text) <> "" Then .Fields("ContNo") = txtContNo.Text
            '***************************************************************************************************************************
            .Fields("CitemCode") = gilCitemCode.GetCodeNo
            .Fields("CitemName") = gilCitemCode.GetDescription
            .Fields("Period") = NoZero(txtPeriod.Text, True)
            .Fields("Amount") = NoZero(ginAmount.Value & "", True)
            .Fields("StartDate") = gdaStartDate.Text
            .Fields("StopDate") = gdaStopDate.Text
            .Fields("ClctDate") = gdaClctDate.Text
            '優待辦法
            .Fields("DiscountCode") = gilDiscount.GetCodeNo2
            .Fields("DiscountName") = gilDiscount.GetDescription2
            .Fields("DiscountDate1") = NoZero(gdaDiscountDate1.GetValue(True))
            .Fields("DiscountDate2") = NoZero(gdaDiscountDate2.GetValue(True))
            .Fields("DiscountAmt") = ginDiscountAmt.Value
            '套餐 91/09/14
            .Fields("PackageNo") = gilPackageNo.GetCodeNo2
            .Fields("PackageName") = gilPackageNo.GetDescription2
            .Fields("UpdTime") = GetDTString(RightNow)
            .Fields("UpdEn") = garyGi(1)
            .Fields("BankCode") = gilBankCode.GetCodeNo2
            .Fields("BankName") = gilBankCode.GetDescription2
            .Fields("CMCode") = gilCMCode.GetCodeNo2
            .Fields("CMName") = gilCMCode.GetDescription2
            .Fields("PTCode") = gilPTCode.GetCodeNo2
            .Fields("PTName") = gilPTCode.GetDescription2
            .Fields("AccountNo") = NoZero(cboAccountNo.Text)
            .Fields("BudgetPeriod") = NoZero(ginBudgetPeriod.Value)
            .Fields("CustAllot") = NoZero(chkCustAllot)
            .Fields("StopFlag") = NoZero(chkStopFlag.Value, True)
            .Fields("CeaseDate") = NoZero(gdaCeaseDate.GetValue(True))
            '***********************************************************************************************************
            '#4133 SO1131C中之【違約時計價依據】及【優惠到期計價依據】各自增加二個權限出來可做修改 By Kin 2008/10/08
            '#6971 如果是清空，資料庫也要跟著清空 By Kin 2015/01/05
            If cboPenalType.Text <> Empty Then
                .Fields("PenalType") = cboPenalType.ListIndex + 1
            Else
                .Fields("PenalType") = Null
            End If
            If cboExpireType.Text <> "" Then
                .Fields("ExpireType") = cboExpireType.ListIndex + 1
            Else
                .Fields("ExpireType") = Null
            End If
            '************************************************************************************************************
            '*******************************************************************
            '#3852 清空帳號要連InvSeqNo一併清除 By Kin 2008/04/16
            If cboAccountNo.Text = Empty Then
                strInvSeqNo = Empty
            End If
            '*******************************************************************
            
        
            
            '#3436 儲存發票流水編號 By Kin 2007/11/27
            .Fields("InvSeqNo") = IIf(strInvSeqNo = Empty, Null, strInvSeqNo)
            If GetByHouse(gilCitemCode.GetCodeNo) = 1 Then
                rs("FaciSeqNo") = gCustId
                rs("FaciSNo") = ""
            Else
                rs("FaciSeqNo") = txtFaciSNo.Tag
                rs("FaciSNo") = txtFaciSNo
            End If
            '#4150 連動SO003A.CitemCode By Kin 2008/10/17
            If blnChgCitem Then
                If Not Chg3ACitem Then
                    
                    Exit Function
                End If
            End If
            If EditMode = giEditModeEdit Then
                If ChkDataModify(rs) Then
                    If Not InsertToLog(rs, giEditModeEdit) Then
                        
                        Exit Function
                    End If
                End If
            End If
            '#4306 增加出單到期日 By Kin 2009/03/25
            If gdaClctStopDate.GetValue & "" <> "" Then
                .Fields("ClctStopDate") = gdaClctStopDate.GetValue(True)
            End If
            '#7228
            '#7645 取消只要有客戶指定就不update By Kin 2017/11/10
            If EditMode = giEditModeEdit Then
                If rs("Period").OriginalValue & "" <> txtPeriod.Text & "" Then
                    If Len(rs("FaciSNo") & "") > 0 Then
'                        aSQL = "Select count(*) from " & GetOwner & "SO003 A" & _
'                                " where rowid <> '" & rs("RowId") & "' " & _
'                                " And CustAllot = 1 " & _
'                                " And A.FaciSNo = '" & rs("FaciSNo") & "' "

                        aSQL = "Select count(*) from " & GetOwner & "SO003 A" & _
                                " where rowid <> '" & rs("RowId") & "' " & _
                                " And Nvl(CustAllot,0) = 0 " & _
                                " And A.FaciSNo = '" & rs("FaciSNo") & "' "
                        '#8627 有指定繳別也不update by kin 2020/07/27
                        If Val(gcnGi.Execute(aSQL)(0)) > 0 Or 1 = 1 Then
                            aSQL = "Select rowid,A.CitemCode,bpcode,Nvl(A.Amount,0) Amount ,Nvl(A.Period,0) Period from " & GetOwner & "SO003 A" & _
                                    " where rowid <> '" & rs("RowId") & "' " & _
                                    " And Nvl(CustAllot,0) = 0 " & _
                                    " And A.FaciSNo = '" & rs("FaciSNo") & "' " & _
                                    " And Nvl(A.PeriodFlag,0) = 0 " & _
                                    " union all " & _
                                    " Select rowid,A.CitemCode,A.bpcode,Nvl(A.Amount,0) Amount ,Nvl(A.Period,0) Period " & _
                                    " FROM " & GetOwner & "SO003 A " & _
                                    " Where faciseqno in (select stbsno from " & GetOwner & "so004 where seqno = '" & rs("FaciSeqNo") & "')" & _
                                    " And rowid <> '" & rs("RowId") & "' " & _
                                    " And Nvl(CustAllot,0) = 0 " & _
                                    " And Nvl(A.PeriodFlag,0) = 0 " & _
                                    " union all " & _
                                     " Select rowid,A.CitemCode,A.bpcode,Nvl(A.Amount,0) Amount ,Nvl(A.Period,0) Period " & _
                                    " FROM " & GetOwner & "SO003 A " & _
                                    " Where faciseqno in (select seqno from " & GetOwner & "so004 where Stbsno = '" & rs("FaciSeqNo") & "')" & _
                                    " And rowid <> '" & rs("RowId") & "' " & _
                                    " And Nvl(CustAllot,0) = 0 " & _
                                    " And Nvl(A.PeriodFlag,0) = 0 "
                                    
                            If Not GetRS(rsOther, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
                            If rsOther.RecordCount > 0 Then
                                rsOther.MoveFirst
                                Do While Not rsOther.EOF
                                    strSign = GetRsValue("Select Sign From " & GetOwner & "CD019 Where CodeNo =" & rsOther("CitemCode")) & ""
                                    If Len(rsOther("BpCode") & "") <> 0 Then
                                        aAmount = Abs((rsOther("Amount") / rsOther("Period")) * Val(txtPeriod.Text & "")) * Val(strSign & "1")
                                    Else
                                        aAmount = Abs(CalAmount(rsOther("CitemCode") & "", Len(rsOther("BpCode") & "") <> 0)) * Val(strSign & "1")
                                    End If
                                    
                                    aSQL = "Update " & GetOwner & "SO003 set Period=" & txtPeriod.Text & _
                                            ",Amount =" & aAmount & _
                                            " Where RowId = '" & rsOther("RowId") & "' "
                                    gcnGi.Execute aSQL
                                    rsOther.MoveNext
                                Loop
                            End If
                            
                        Else
                            'MsgBox "其它收費項目有包含客戶指定！" & vbCrLf & "不同步其它收費項目期數", vbInformation, "訊息"
                        End If
                    End If
                End If
            End If
            .Update
        End With
        ScrToRcd = True
        On Error Resume Next
        CloseRecordset rsOther
    Exit Function
ChkErr:
    ScrToRcd = False
    
    'If Err.Number = "-2147217873" Then
    '   MsgBox "收費項目或期數重覆!!", , "訊息.."
    'Else
        rs.CancelUpdate
        
       Call ErrSub(Me.Name, "ScrToRcd")
    'End If
End Function
'#4150 更改SO003A的收費項目 By Kin 2008/10/17
Private Function Chg3ACitem() As Boolean
  On Error GoTo ChkErr
    Dim strUpdSql As String
    '#4287 設備也要跟著連動 By Kin 2008/12/12
    If rs("FaciSeqNo").OriginalValue & "" <> txtFaciSNo.Tag Then
        strUpdSql = "UpDate " & GetOwner & "SO003A Set FaciSNo='" & txtFaciSNo & "'," & _
                    "FaciSeqNo='" & txtFaciSNo.Tag & "'" & _
                    " Where CustId=" & gCustId & " And CompCode=" & gCompCode & _
                    IIf(EditMode = giEditModeEdit, " And CitemCode=" & rs("CitemCode").OriginalValue, " And CitemCode=" & gilCitemCode.GetCodeNo) & _
                    IIf(IsNull(rs("FaciSeqNo").OriginalValue), " And FaciSeqNo is Null ", " And FaciSeqNo='" & rs("FaciSeqNo").OriginalValue & "'")
        gcnGi.Execute strUpdSql
    End If
    strUpdSql = Empty
    '#4150 阿貴說要多串設備流水號 By Kin 2008/10/22
    If rs("CitemCode").OriginalValue <> gilCitemCode.GetCodeNo Then
        strUpdSql = "Update " & GetOwner & "SO003A Set CitemCode=" & gilCitemCode.GetCodeNo & _
                    " Where CustId=" & gCustId & " And CompCode=" & gCompCode & _
                    IIf(EditMode = giEditModeEdit, " And CitemCode=" & rs("CitemCode").OriginalValue, " And CitemCode=" & gilCitemCode.GetCodeNo) & _
                    IIf(IsNull(rs("FaciSeqNo")), " And FaciSeqNo is Null ", " And FaciSeqNo='" & rs("FaciSeqNo") & "'")
                    
        gcnGi.Execute strUpdSql
    End If
    '#6856 同步更新SO003A的違約時計價依據 By Kin 2014/09/30
    If (rs("PenalType").OriginalValue & "") <> (cboPenalType.ListIndex + 1) Then
        If cboPenalType.Text & "" = "" Then
            strUpdSql = "Update " & GetOwner & "SO003A Set PenalType = Null " & _
                " Where CustId = " & gCustId & " And CompCode = " & gCompCode & _
                  IIf(EditMode = giEditModeEdit, " And CitemCode=" & rs("CitemCode").OriginalValue, " And CitemCode=" & gilCitemCode.GetCodeNo) & _
                    IIf(IsNull(rs("FaciSeqNo")), " And FaciSeqNo is Null ", " And FaciSeqNo='" & rs("FaciSeqNo") & "'")
        Else
            strUpdSql = "Update " & GetOwner & "SO003A Set PenalType = " & cboPenalType.ListIndex + 1 & _
                " Where CustId = " & gCustId & " And CompCode = " & gCompCode & _
                  IIf(EditMode = giEditModeEdit, " And CitemCode=" & rs("CitemCode").OriginalValue, " And CitemCode=" & gilCitemCode.GetCodeNo) & _
                    IIf(IsNull(rs("FaciSeqNo")), " And FaciSeqNo is Null ", " And FaciSeqNo='" & rs("FaciSeqNo") & "'")
        End If
        gcnGi.Execute strUpdSql
    End If
    Chg3ACitem = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "Chg3ACitem")
End Function
Private Sub ShowDiscountDate(ByVal sDate As String)
  On Error GoTo ChkErr
    Dim strQry As String
    Dim strQry2 As String
    Dim strQry3 As String
    Dim strSQLCD078A As String
    Dim strSign As String
    Dim rsCD078a As New ADODB.Recordset
    Dim strCitemCode As String
    Dim blnBeen As Boolean
    blnBeen = True
    '#4121 增加BpCode,BpName,OrderNo,ContStartDate,ContStopDate,PromCode,PromName,ContMonth By Kin 2008/10/01
    '#4448 如果是由阿貴呼叫,要以阿貴的Recordset為主 By Kin 2009/04/13
    strQry = "Select DiscountDate1,DiscountDate2,DiscountAmt,BpCode,BpName,OrderNo" & _
             ",ContStartDate,ContStopDate,PromCode,PromName,ContMonth From " & _
             GetOwner & "SO003A" & _
            " Where Custid=" & rs("CustId") & _
            " And CompCode=" & gCompCode & _
            " And ServiceType='" & rs("ServiceType") & "'" & _
            " And FaciSeqNo='" & rs("FaciSeqNo") & "'" & _
            " And ContNo = '" & rs("ContNo") & "'" & _
            " And CitemCode = " & rs("CitemCode") & _
            " And DiscountDate1<=to_Date('" & sDate & "','yyyy/mm/dd')" & _
            " And DiscountDate2>=to_Date('" & sDate & "','yyyy/mm/dd')" & _
            " And StopFlag<>1"
    strQry2 = "Select Min(DiscountDate1) DiscountDate1,Min(DiscountDate2) DiscountDate2 From " & _
             GetOwner & "SO003A" & _
            " Where Custid=" & rs("CustId") & _
            " And CompCode=" & gCompCode & _
            " And ServiceType='" & rs("ServiceType") & "'" & _
            " And FaciSeqNo='" & rs("FaciSeqNo") & "'" & _
            " And ContNo = '" & rs("ContNo") & "'" & _
            " And CitemCode = " & rs("CitemCode") & _
            " And StopFlag<>1" & _
            " Having Min(DiscountDate1)>=to_Date('" & sDate & "','yyyy/mm/dd')"
    strQry2 = "Select A.DiscountDate1,A.DiscountDate2,A.DiscountAmt,A.BpCode,A.BpName,A.OrderNo" & _
             ",A.ContStartDate,A.ContStopDate,A.PromCode,A.PromName,A.ContMonth From " & _
             GetOwner & "SO003A A,(" & strQry2 & ") B " & _
            " Where A.Custid=" & rs("CustId") & _
            " And A.CompCode=" & gCompCode & _
            " And A.ServiceType='" & rs("ServiceType") & "'" & _
            " And A.FaciSeqNo='" & rs("FaciSeqNo") & "'" & _
            " And A.ContNo = '" & rs("ContNo") & "'" & _
            " And A.CitemCode = " & rs("CitemCode") & _
            " And A.DiscountDate1=B.DiscountDate1" & _
            " And A.DiscountDate2=B.DiscountDate2" & _
            " And A.StopFlag<>1"
    strQry3 = "Select Max(DiscountDate1) DiscountDate1,Max(DiscountDate2) DisCountDate2 From " & _
             GetOwner & "SO003A" & _
            " Where Custid=" & rs("CustId") & _
            " And CompCode=" & gCompCode & _
            " And ServiceType='" & rs("ServiceType") & "'" & _
            " And FaciSeqNo='" & rs("FaciSeqNo") & "'" & _
            " And ContNo = '" & rs("ContNo") & "'" & _
            " And CitemCode = " & rs("CitemCode") & _
            " And StopFlag<>1" & _
            " Having Max(DiscountDate1)<=to_Date('" & sDate & "','yyyy/mm/dd')"
    strQry3 = "Select A.DiscountDate1,A.DiscountDate2,A.DiscountAmt,A.BpCode,A.BpName,A.OrderNo" & _
             ",A.ContStartDate,A.ContStopDate,A.PromCode,A.PromName,A.ContMonth From " & _
             GetOwner & "SO003A A,(" & strQry3 & ") B " & _
            " Where A.Custid=" & rs("CustId") & _
            " And A.CompCode=" & gCompCode & _
            " And A.ServiceType='" & rs("ServiceType") & "'" & _
            " And A.FaciSeqNo='" & rs("FaciSeqNo") & "'" & _
            " And A.ContNo = '" & rs("ContNo") & "'" & _
            " And A.CitemCode = " & rs("CitemCode") & _
            " And A.DiscountDate1=B.DiscountDate1" & _
            " And A.DiscountDate2=B.DiscountDate2" & _
            " And A.StopFlag<>1"
    If Not GetRS(rsDiscount, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If rsDiscount.EOF Then
        If Not GetRS(rsDiscount, strQry2, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If rsDiscount.EOF Then
            If Not GetRS(rsDiscount, strQry3, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        End If
    Else
        blnBeen = False
    End If
    If Not rsDiscount.EOF Then
        If Not IsNull(rsDiscount("DiscountDate1")) Then
            gdaDiscountDate3.SetValue rsDiscount("DiscountDate1") & ""
        End If
        If Not IsNull(rsDiscount("DiscountDate2")) Then
            gdaDiscountDate4.SetValue rsDiscount("DiscountDate2") & ""
        End If
        If Not IsNull(rsDiscount("DisCountAmt")) Then
            txtDiscountAmt.Text = rsDiscount("DiscountAmt") & ""
        End If
        If Not IsNull(rsDiscount("BpCode")) Then
            gilBPCode.SetCodeNo rsDiscount("BPCode") & ""
            gilBPCode.SetDescription rsDiscount("BPName") & ""
            '#4213 如果在編輯模式要過濾符合CD078A的收費項目 By Kin 2008/11/07
            'CD019.RefNo=19不能被選擇 By Kin 2008/12/11
            '#4305 過期的話不用過濾收費項目 By Kin 2008/12/25
            If EditMode = giEditModeEdit And Not blnBeen Then
                strCitemCode = gilCitemCode.GetCodeNo
                strSign = GetRsValue("Select Sign From " & GetOwner & "CD019 Where CodeNo = " & gilCitemCode.GetCodeNo) & ""
                strSQLCD078A = " Where CodeNo In(Select CitemCode From " & GetOwner & "CD078A" & _
                            " Where BpCode='" & rsDiscount("BpCode") & "' And ServiceType='" & strServiceType & "')" & _
                            " And Sign='" & strSign & "' And PeriodFlag=1 And StopFlag<>1 AND NVL(RefNo,0)<>19"
                SetgiList gilCitemCode, "CodeNO", "Description", "CD019", , , , , , , strSQLCD078A, True
                gilCitemCode.SetCodeNo strCitemCode
                gilCitemCode.Query_Description
                If gilCitemCode.GetDescription = "" Then gilCitemCode.SetCodeNo ""
            End If

        End If
        If Not IsNull(rsDiscount("OrderNo")) Then
            txtOrderNo.Text = rsDiscount("OrderNo") & ""
        End If
        '#5996 將Max(SO003A.DiscountDate2) Upd至SO003.DiscountDate2 By Kin 2011/05/24
        If Len(rsDiscount("DiscountDate2") & "") > 0 Then
            If blnSO1131FSave Then          'FrmSO1131F有按存檔才需要更新資料
                Dim aUpdSQL As String
                Dim aUpdCnt As Long
                Dim aMaxDisDate2 As String
                '#5996 尋找SO003A.Max(DiscountDate2) By Kin 2011/05/24
                aMaxDisDate2 = GetRsValue("SELECT MAX(DISCOUNTDATE2) FROM " & GetOwner & "SO003A " & _
                                    " WHERE SERVICETYPE = '" & rs("SERVICETYPE") & "'" & _
                                    " AND FACISEQNO = '" & rs("FACISEQNO") & "' " & _
                                    " AND CITEMCODE = " & rs("CITEMCODE") & _
                                    " AND COMPCODE = " & gCompCode & _
                                    " AND CUSTID = " & rs("CUSTID") & _
                                    " AND ContNo = '" & rs("ContNo") & "'" & _
                                    " AND NVL(STOPFLAG,0) = 0 ", gcnGi) & ""
                                                    
                If Len(aMaxDisDate2) > 0 Then
                    If IsDate(aMaxDisDate2) Then
                        '#5996 將Max(DiscountDate2)回填至SO003 By Kin 2011/05/24
                        aUpdSQL = "UPDATE " & GetOwner & "SO003 " & _
                                " SET DISCOUNTDATE2 = TO_DATE('" & Format(aMaxDisDate2, "yyyymmdd") & "','yyyymmdd') " & _
                                " WHERE CUSTID = " & rs("CUSTID") & _
                                " AND SERVICETYPE = '" & rs("SERVICETYPE") & "' " & _
                                " AND FACISEQNO = '" & rs("FACISEQNO") & "' " & _
                                " AND CITEMCODE = " & rs("CITEMCODE") & _
                                " AND COMPCODE = " & gCompCode
                                
                        gcnGi.Execute aUpdSQL, aUpdCnt
                    End If
                End If
                If aUpdCnt > 0 Then
                    gdaDiscountDate2.SetValue aMaxDisDate2 & ""
                    rs.Requery
                End If
            End If

        End If
        gdaContStartDate.SetValue rsDiscount("ContStartDate") & ""
        gdaContStopDate.SetValue rsDiscount("ContStopDate") & ""
        gilPromCode.SetCodeNo rsDiscount("PromCode") & ""
        gilPromCode.SetDescription rsDiscount("PromName") & ""
        txtBundleMon.Text = rsDiscount("ContMonth") & ""
        
    End If
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "ShowDiscountDate")
End Sub
Private Function ChkDataModify(rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim intLoop As Integer
        ChkDataModify = False
        For intLoop = 0 To rs.Fields.Count - 1
            If UCase(rs.Fields(intLoop).Name) <> "UPDEN" And UCase(rs.Fields(intLoop).Name) <> "UPDTIME" Then
                If rs.Fields(intLoop) & "" <> rs.Fields(intLoop).OriginalValue & "" Then
                    ChkDataModify = True
                    Exit For
                End If
            End If
        Next
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "ChkDataModify")
End Function

Private Function InsertToLog(rs As ADODB.Recordset, lngMode As giEditModeEnu) As Boolean
    On Error GoTo ChkErr
    Dim lngLoop As Long
    Dim rsSo061 As New ADODB.Recordset
    '#7206 Don't record the so061 table by kin 2016/03/21
    InsertToLog = True
    On Error Resume Next
    CloseRecordset rsSo061
    Exit Function
    Dim obj As Object
        Set obj = CreateObject("csLogCom.clsLogTABLE")
        InsertToLog = obj.InsertToLog061(rs, lngMode, GetOwner, gcnGi)
        Set obj = Nothing
        Exit Function
        If Not GetRS(rsSo061, "Select * From " & GetOwner & "So061 Where 1=0 ", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        With rsSo061
            .AddNew
            .Fields("CustID") = NoZero(rs("CustID"))
            .Fields("Custname") = GetRsValue("Select CustName From " & GetOwner & "So001 Where CustId = " & rs("CustId") & " And CompCode=" & gCompCode)
            .Fields("CompCode") = NoZero(rs("CompCode"))
            .Fields("ServiceType") = NoZero(rs("ServiceType"))
            .Fields("UPDEN") = NoZero(rs("UpdEn"))
            .Fields("UPDTIME") = NoZero(rs("UPDTIME"))
            
            .Fields("CITEMCODE") = NoZero(rs("CITEMCODE").OriginalValue)
            .Fields("CITEMNAME") = NoZero(rs("CITEMNAME").OriginalValue)
            .Fields("PERIOD") = NoZero(rs("PERIOD").OriginalValue, True)
            .Fields("AMOUNT") = NoZero(rs("AMOUNT").OriginalValue, True)
            .Fields("STARTDATE") = NoZero(rs("STARTDATE").OriginalValue)
            .Fields("STOPDATE") = NoZero(rs("STOPDATE").OriginalValue)
            .Fields("CLCTDATE") = NoZero(rs("CLCTDATE").OriginalValue)
            .Fields("DISCOUNTCODE") = NoZero(rs("DISCOUNTCODE").OriginalValue)
            .Fields("DISCOUNTNAME") = NoZero(rs("DISCOUNTNAME").OriginalValue)
            .Fields("DISCOUNTDATE1") = NoZero(rs("DISCOUNTDATE1").OriginalValue)
            .Fields("DISCOUNTDATE2") = NoZero(rs("DISCOUNTDATE2").OriginalValue)
            .Fields("DISCOUNTAMT") = NoZero(rs("DISCOUNTAMT").OriginalValue)
            .Fields("AccountNo") = NoZero(rs("AccountNo").OriginalValue)
            .Fields("BankCode") = NoZero(rs("BankCode").OriginalValue)
            .Fields("BankName") = NoZero(rs("BankName").OriginalValue)
            .Fields("CMCode") = NoZero(rs("CMCode").OriginalValue)
            .Fields("CMName") = NoZero(rs("CMName").OriginalValue)
            .Fields("BudgetPeriod") = NoZero(rs("BudgetPeriod").OriginalValue)
            .Fields("CustAllot") = NoZero(rs("CustAllot").OriginalValue)
            
            If lngMode = giEditModeInsert Then
                .Fields("FuncType") = 3      '新增
            ElseIf lngMode = giEditModeDelete Then
                .Fields("FuncType") = 2      '刪除
                .Fields("UPDEN") = garyGi(1)
                .Fields("UPDTIME") = GetDTString(RightNow)
            Else
                .Fields("FuncType") = 1      '修改
                
                .Fields("CITEMCODEB") = NoZero(rs("CITEMCODE").Value)
                .Fields("CITEMNAMEB") = NoZero(rs("CITEMNAME").Value)
                .Fields("PERIODB") = NoZero(rs("PERIOD").Value, True)
                .Fields("AMOUNTB") = NoZero(rs("AMOUNT").Value, True)
                .Fields("STARTDATEB") = NoZero(rs("STARTDATE").Value)
                .Fields("STOPDATEB") = NoZero(rs("STOPDATE").Value)
                .Fields("CLCTDATEB") = NoZero(rs("CLCTDATE").Value)
                .Fields("DISCOUNTCODEB") = NoZero(rs("DISCOUNTCODE").Value)
                .Fields("DISCOUNTNAMEB") = NoZero(rs("DISCOUNTNAME").Value)
                .Fields("DISCOUNTDATE1B") = NoZero(rs("DISCOUNTDATE1").Value)
                .Fields("DISCOUNTDATE2B") = NoZero(rs("DISCOUNTDATE2").Value)
                .Fields("DISCOUNTAMTB") = NoZero(rs("DISCOUNTAMT").Value, True)
                .Fields("AccountNoB") = NoZero(rs("AccountNo"))
                .Fields("BankCodeB") = NoZero(rs("BankCode"))
                .Fields("BankNameB") = NoZero(rs("BankName"))
                .Fields("CMCodeB") = NoZero(rs("CMCode"))
                .Fields("CMNameB") = NoZero(rs("CMName"))
                .Fields("BudgetPeriodB") = NoZero(rs("BudgetPeriod"))
                .Fields("CustAllotB") = NoZero(rs("CustAllot"))
            End If
            .Update
        End With
        InsertToLog = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "InsertToLog")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ChkErr
        If Not lngEditMode = giEditModeView And Me.Tag = "" Then
            If giMsgNotSave = vbNo Then Cancel = True: Exit Sub
        End If
        Call CloseRecordset(rsCust)
        Call CloseRecordset(rsMdu)
        Call CloseRecordset(rs)
        Call CloseRecordset(rsSO043)
        Call CloseRecordset(rsTmp)
        '#3436 新增SO138的Recordset By Kin 2007/11/27
        Call CloseRecordset(rsSO138)
        Call CloseRecordset(rsDiscount)
        Call CloseRecordset(rsFROM)
        Call FormQueryUnload
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_QueryUnload"
End Sub

Private Sub Form_Resize()
        'Call FormResizePosition(Me, varC)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        ReleaseCOM frmSO1131C

End Sub

Private Sub gdaStartDate_Validate(Cancel As Boolean)
  If lngEditMode = giEditModeInsert Then Call AutoDate
End Sub


Private Sub gdaStopDate_Validate(Cancel As Boolean)
  On Error Resume Next
  If gdaStopDate.Text = "" Then Exit Sub
  gdaClctDate.SetValue CStr(CDate(gdaStopDate.Text) + Val(rsSO043.Fields("para5")))
End Sub

Private Sub ggrCustRate_DblClick()
  On Error GoTo ChkErr
  If lngEditMode = giEditModeInsert Then
     With ggrCustRate
          gilCitemCode.SetDescription .TextMatrix(.Row, 1)
          gilCitemCode.GotoDescription
          gilCitemCode.GotoCodeNo
          txtPeriod.Text = .TextMatrix(.Row, 2)
          ginAmount.Text = .TextMatrix(.Row, 3)
          gdaStartDate.SetFocus
     End With
  End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ggrCustRate_DblClick")
End Sub

Private Sub ggrMduRate_DblClick()
  On Error GoTo ChkErr
  If lngEditMode = giEditModeInsert Then
     With ggrMduRate
          gilCitemCode.SetDescription .TextMatrix(.Row, 1)
          gilCitemCode.GotoDescription
          gilCitemCode.GotoCodeNo
          txtPeriod.Text = .TextMatrix(.Row, 2)
          ginAmount.Text = .TextMatrix(.Row, 3)
          gdaStartDate.SetFocus
     End With
  End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ggrMduRate_DblClick")
End Sub

Private Sub gilBankCode_Change()
     If gilBankCode.GetCodeNo = "" Then
        cboAccountNo.ListIndex = -1
        cboAccountNo.Text = ""
        cboAccountNo.Enabled = False
     Else
        'cboAccountNo.Enabled = True
        Call cboAccountNo_DropDown
        If cboAccountNo.ListCount = 1 Then cboAccountNo.ListIndex = 0
     End If
     
End Sub

Private Sub gilCitemCode_Change()
    On Error Resume Next
        If GetByHouse(gilCitemCode.GetCodeNo) = 1 Then
            lblFaciSno.Visible = False
            txtFaciSNo.Visible = False
            cmdHouse.Visible = False
        Else
            lblFaciSno.Visible = True
            txtFaciSNo.Visible = True
            cmdHouse.Visible = True
        End If
End Sub

Private Sub txtAccountNo_Change()
  On Error Resume Next
    Dim strQry As String
    Dim rsTemp As New ADODB.Recordset
    
    If txtAccountNo.Text <> Empty Then
        cmdInvSeqNo.Enabled = True
        '*****************************************************************************************************
        '#3929 對應的帳號如果只有一筆發票則要自動帶入 By Kin 2008/06/02
'        strQry = "Select InvSeqNo From " & GetOwner & "SO002AD" & _
'                " Where CustId=" & gCustId & _
'                " And CompCode=" & gCompCode & _
'                " And AccountNo='" & IIf(IsNumeric(txtAccountNo.Text), txtAccountNo.Text, cboAccountNo.Text) & "'"
'        If Not GetRS(rsTemp, strQry, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Sub
'        If rsTemp.RecordCount = 1 Then
'            '#5418 如果只有一筆，就直接以SO003.InvSeqNo的資料為主 By Kin 2009/12/03
'            'strInvSeqNo = rsTemp("InvSeqNo") & ""
'            strInvSeqNo = rs("InvSeqNo") & ""
'            OpenInvTitle
'        End If
        
        strQry = "Select InvSeqNo From " & GetOwner & "SO138" & _
                " Where CustId=" & gCustId
        If Not GetRS(rsTemp, strQry, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Sub
        If rsTemp.RecordCount = 1 Then
            '#5418 如果只有一筆，就直接以SO003.InvSeqNo的資料為主 By Kin 2009/12/03
            'strInvSeqNo = rsTemp("InvSeqNo") & ""
            strInvSeqNo = rs("InvSeqNo") & ""
            OpenInvTitle
        End If

        '*******************************************************************************************************
    Else
        cmdInvSeqNo.Enabled = False
    End If
    Close3Recordset rsTemp
End Sub
Private Function CalAmount(ByVal strCitemCode As String, Optional ByVal blnBpCode As Boolean = False) As String
  On Error GoTo ChkErr
      Dim rsAmount As New ADODB.Recordset
        If txtPeriod = "" Or gilCitemCode.GetCodeNo = "" Then Exit Function
        If lngEditMode = giEditModeInsert Then Call AutoDate
        '#7228 if bpcode existed then using amount/period by kin 2016/07/21
'        If blnBpCode Then
'
'            If rsAmount.State = adStateOpen Then rsAmount.Close
'                rsAmount.CursorLocation = adUseClient
'                rsAmount.Open "Select CD019.Amount,CD019.Period From " & GetOwner & "CD019 " & _
'                        " Where CD019.CodeNo=" & strCitemCode & " And CD019.Period>0 " _
'                        , gcnGi, adOpenForwardOnly, adLockReadOnly
'                 If rsAmount.EOF Or rsAmount.BOF Then
'                    'ginAmount.Text = 0
'                    CalAmount = 0
'                Else
'                    CalAmount = (Val(rsAmount("Amount") & "") / Val(rsAmount("Period") & "")) * Val(txtPeriod.Text)
'                End If
'                Exit Function
'        End If
        
        If strMduId <> "" Then ' 為大樓住戶
            If rsAmount.State = adStateOpen Then rsAmount.Close
            rsAmount.Open "SELECT CD019.CodeNo,CD019SO017.Period,CD019SO017.Amount,CD019SO017.CitemCode " & _
                            " FROM " & GetOwner & "CD019SO017," & GetOwner & "CD019 " & _
                            " WHERE MduId='" & strMduId & "' AND CD019SO017.CitemCode=CD019.CodeNo" & _
                        " And CD019.CodeNo=" & strCitemCode & " And CD019SO017.Period=" & Me.txtPeriod.Text, _
                        gcnGi, adOpenForwardOnly, adLockReadOnly
          '************************************************************
          '#5073 如果找不到,以一般客戶角度下去抓 By Kin 2009/05/20
            If rsAmount.BOF Or rsAmount.EOF Then
                If rsAmount.State = adStateOpen Then rsAmount.Close
                rsAmount.CursorLocation = adUseClient
                rsAmount.Open "SELECT CD019.CodeNo ,CD019CD004.Period,CD019CD004.Amount,CD019CD004.ClassCode,CD019CD004.CitemCode " & _
                                    " FROM " & GetOwner & "CD019CD004," & GetOwner & "CD019 " & _
                                    " WHERE CD019CD004.ClassCode=" & intClassCode & " AND CD019CD004.CitemCode=CD019.CodeNo and " & _
                                    "CD019.CodeNo=" & strCitemCode & " And CD019CD004.Period=" & Me.txtPeriod.Text, _
                                    gcnGi, adOpenForwardOnly, adLockReadOnly
            End If
          '*************************************************************
        Else '一般客戶
            If rsAmount.State = adStateOpen Then rsAmount.Close
            rsAmount.CursorLocation = adUseClient
            rsAmount.Open "SELECT CD019.CodeNo ,CD019CD004.Period,CD019CD004.Amount," & _
                                    " CD019CD004.ClassCode,CD019CD004.CitemCode " & _
                                    " FROM " & GetOwner & "CD019CD004," & GetOwner & "CD019 " & _
                                    " WHERE CD019CD004.ClassCode=" & intClassCode & " AND CD019CD004.CitemCode=CD019.CodeNo and " & _
                                        "CD019.CodeNo=" & strCitemCode & " And CD019CD004.Period=" & Me.txtPeriod.Text, _
                                    gcnGi, adOpenForwardOnly, adLockReadOnly
    
        End If
        '***********************************************************************************************************
        '#5073 如果都抓不到,則抓取CD019CD004空的客戶類別,再抓不到就用預設金額 By Kin 2009/05/20
        If rsAmount.EOF Or rsAmount.BOF Then
            If rsAmount.State = adStateOpen Then rsAmount.Close
            rsAmount.CursorLocation = adUseClient
            rsAmount.Open "SELECT CD019.CodeNo ,CD019CD004.Period,CD019CD004.Amount," & _
                        "CD019CD004.ClassCode,CD019CD004.CitemCode FROM " & _
                        GetOwner & "CD019CD004," & GetOwner & "CD019 " & _
                        " WHERE  CD019CD004.CitemCode=CD019.CodeNo  " & _
                        " And CD019.CodeNo=" & strCitemCode & _
                        " And CD019CD004.ClassCode Is Null " & _
                        " And CD019CD004.Period=" & Me.txtPeriod.Text, _
                        gcnGi, adOpenForwardOnly, adLockReadOnly
            '#5073 測試不OK,金額應該是為Amount/Period(單月) By Kin 2009/06/02
            If rsAmount.EOF Or rsAmount.BOF Then
                If rsAmount.State = adStateOpen Then rsAmount.Close
                rsAmount.CursorLocation = adUseClient
                rsAmount.Open "Select CD019.Amount,CD019.Period From " & GetOwner & "CD019 " & _
                        " Where CD019.CodeNo=" & strCitemCode & " And CD019.Period>0 " _
                        , gcnGi, adOpenForwardOnly, adLockReadOnly
            End If
        End If
        '***************************************************************************************************************
        If rsAmount.EOF Or rsAmount.BOF Then
            'ginAmount.Text = 0
            CalAmount = 0
        Else
            CalAmount = (Val(rsAmount("Amount") & "") / Val(rsAmount("Period") & "")) * Val(txtPeriod.Text)
            'ginAmount.Text = (Val(rsAmount("Amount") & "") / Val(rsAmount("Period") & "")) * Val(txtPeriod.Text)
        End If
        '#6908 2014/10/28 Jacky 調整
        If rs.RecordCount > 0 Then
            '#6841 增加長繳別判斷並修改金額 By Kin 2014/08/18
            If Val(rs("LongPayFlag") & "") = 1 And Len(gilBPCode.GetCodeNo & "") > 0 Then
                If Not LongPayTypeAmount Then
                    
               
                End If
            End If
        End If
        On Error Resume Next
        Call CloseRecordset(rsAmount)

  
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "CalAmount")
End Function


Private Sub txtPeriod_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If EditMode <> giEditModeInsert Then
        If Len(rs("BpCode") & "") <> 0 Then
            ginAmount.Text = (rs("Amount").OriginalValue / rs("Period").OriginalValue) * Val(txtPeriod.Text & "")
        Else
            ginAmount.Text = CalAmount(gilCitemCode.GetCodeNo & "", (Len(rs("BpCode") & "") <> 0))
        End If
    Else
        ginAmount.Text = CalAmount(gilCitemCode.GetCodeNo & "", False)
    End If
  Exit Sub
  Dim rsAmount As New ADODB.Recordset
    If txtPeriod = "" Or gilCitemCode.GetCodeNo = "" Then Exit Sub
    If lngEditMode = giEditModeInsert Then Call AutoDate
   
    If strMduId <> "" Then ' 為大樓住戶
        If rsAmount.State = adStateOpen Then rsAmount.Close
        rsAmount.Open "SELECT CD019.CodeNo,CD019SO017.Period,CD019SO017.Amount,CD019SO017.CitemCode FROM " & _
                            GetOwner & "CD019SO017," & GetOwner & "CD019 " & _
                            " WHERE MduId='" & strMduId & "' AND CD019SO017.CitemCode=CD019.CodeNo" & _
                            " And CD019.CodeNo=" & gilCitemCode.GetCodeNo & " And CD019SO017.Period=" & Me.txtPeriod.Text, _
                            gcnGi, adOpenForwardOnly, adLockReadOnly
      '************************************************************
      '#5073 如果找不到,以一般客戶角度下去抓 By Kin 2009/05/20
        If rsAmount.BOF Or rsAmount.EOF Then
'        ginAmount.Text = ""
'        Exit Sub
            If rsAmount.State = adStateOpen Then rsAmount.Close
            rsAmount.CursorLocation = adUseClient
            rsAmount.Open "SELECT CD019.CodeNo ,CD019CD004.Period,CD019CD004.Amount,CD019CD004.ClassCode,CD019CD004.CitemCode FROM " & GetOwner & "CD019CD004," & GetOwner & "CD019 WHERE CD019CD004.ClassCode=" & intClassCode & " AND CD019CD004.CitemCode=CD019.CodeNo and " & _
                        "CD019.CodeNo=" & gilCitemCode.GetCodeNo & " And CD019CD004.Period=" & Me.txtPeriod.Text, gcnGi, adOpenForwardOnly, adLockReadOnly
        End If
      '*************************************************************
    Else '一般客戶
        If rsAmount.State = adStateOpen Then rsAmount.Close
        rsAmount.CursorLocation = adUseClient
        rsAmount.Open "SELECT CD019.CodeNo ,CD019CD004.Period,CD019CD004.Amount,CD019CD004.ClassCode,CD019CD004.CitemCode FROM " & GetOwner & "CD019CD004," & GetOwner & "CD019 WHERE CD019CD004.ClassCode=" & intClassCode & " AND CD019CD004.CitemCode=CD019.CodeNo and " & _
                      "CD019.CodeNo=" & gilCitemCode.GetCodeNo & " And CD019CD004.Period=" & Me.txtPeriod.Text, gcnGi, adOpenForwardOnly, adLockReadOnly
   
    End If
    '***********************************************************************************************************
    '#5073 如果都抓不到,則抓取CD019CD004空的客戶類別,再抓不到就用預設金額 By Kin 2009/05/20
    If rsAmount.EOF Or rsAmount.BOF Then
        If rsAmount.State = adStateOpen Then rsAmount.Close
        rsAmount.CursorLocation = adUseClient
        rsAmount.Open "SELECT CD019.CodeNo ,CD019CD004.Period,CD019CD004.Amount," & _
                    "CD019CD004.ClassCode,CD019CD004.CitemCode FROM " & _
                    GetOwner & "CD019CD004," & GetOwner & "CD019 " & _
                    " WHERE  CD019CD004.CitemCode=CD019.CodeNo  " & _
                    " And CD019.CodeNo=" & gilCitemCode.GetCodeNo & _
                    " And CD019CD004.ClassCode Is Null " & _
                    " And CD019CD004.Period=" & Me.txtPeriod.Text, _
                    gcnGi, adOpenForwardOnly, adLockReadOnly
        '#5073 測試不OK,金額應該是為Amount/Period(單月) By Kin 2009/06/02
        If rsAmount.EOF Or rsAmount.BOF Then
            If rsAmount.State = adStateOpen Then rsAmount.Close
            rsAmount.CursorLocation = adUseClient
            rsAmount.Open "Select CD019.Amount,CD019.Period From " & GetOwner & "CD019 " & _
                    " Where CD019.CodeNo=" & gilCitemCode.GetCodeNo & " And CD019.Period>0 " _
                    , gcnGi, adOpenForwardOnly, adLockReadOnly
        End If
    End If
    '***************************************************************************************************************
    If rsAmount.EOF Or rsAmount.BOF Then
        ginAmount.Text = 0
    Else
        ginAmount.Text = (Val(rsAmount("Amount") & "") / Val(rsAmount("Period") & "")) * Val(txtPeriod.Text)
    End If
    '#6908 2014/10/28 Jacky 調整
    If rs.RecordCount > 0 Then
        '#6841 增加長繳別判斷並修改金額 By Kin 2014/08/18
        If Val(rs("LongPayFlag") & "") = 1 And Len(gilBPCode.GetCodeNo & "") > 0 Then
            If Not LongPayTypeAmount Then
                
           
            End If
        End If
    End If
    On Error Resume Next
    Call CloseRecordset(rsAmount)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "txtPeriod_Validate")
End Sub

Private Sub txtPeriod_GotFocus()
  On Error GoTo ChkErr
    txtPeriod.SelStart = 0
    txtPeriod.SelLength = Len(txtPeriod)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "txtPeriod_GotFocus")
End Sub
Private Function LongPayTypeAmount() As Boolean
  On Error GoTo ChkErr
    Dim rsCD078A1 As New ADODB.Recordset
    Dim aSQL As String
    If Val(rs("LongPayFlag") & "") = 1 And Len(gilBPCode.GetCodeNo & "") > 0 Then
        aSQL = "SELECT MONTHAMT1 FROM " & GetOwner & "CD078A1 " & _
            " WHERE BPCODE = '" & gilBPCode.GetCodeNo & "'" & _
            " AND CITEMCODE = " & gilCitemCode.GetCodeNo & _
            " AND PERIOD1 = " & txtPeriod.Text
        If Not GetRS(rsCD078A1, aSQL, gcnGi, adUseClient, adOpenKeyset) Then Exit Function
        If (rsCD078A1.RecordCount = 0) Or (rsCD078A1.RecordCount > 1) Then
            MsgBox "對應不到優惠組合期數，無法調整對應多階金額資料！", vbInformation, "訊息"
            LongPayTypeAmount = False
            Exit Function
        Else
            ginAmount.Text = Val(txtPeriod.Text) * Val(rsCD078A1("MONTHAMT1")) & ""
        End If
    End If
    LongPayTypeAmount = True
    On Error Resume Next
    Call CloseRecordset(rsCD078A1)
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "LongPayTypeAmount")
End Function
Private Sub AutoDate()
    On Error GoTo ChkErr
    Dim strStopDate As String, strClctDate As String
        If Not ChkDTok Then Exit Sub
        GetAutoStopDate Val(txtPeriod.Text), gdaStartDate.GetValue(True), strStopDate, strClctDate
        gdaStopDate.SetValue strStopDate
        gdaClctDate.SetValue strClctDate
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "AutoDate")
End Sub

'#3436 顯示發票資料 By Kin 2007/11/27
Private Sub OpenInvTitle()
  On Error GoTo ChkErr
    lblInvSeqNoValue = ""
    lblChargeAddressValue = ""
    lblInvoiceTypeValue = ""
    lblInvNoValue = ""
    lblInvTitleValue = ""
    lblInvAddressValue = ""
    lblChargeTitleValue = ""
    lblMailAddressValue = ""
    lblInvoiceKind = ""
    
    If strInvSeqNo <> "" Then
        If Not GetRS(rsSO138, "Select * From " & GetOwner & "SO138 Where InvSeqNo=" & strInvSeqNo & "") Then Exit Sub
        If Not rsSO138.EOF Then
            With rsSO138
                lblInvSeqNoValue = .Fields("InvSeqNo") & ""
                lblChargeAddressValue = .Fields("ChargeAddress") & ""
                lblInvoiceTypeValue = IIf(.Fields("InvoiceType") & "" = 0, "免開", IIf(.Fields("InvoiceType") & "" = "2", "二聯式", "三聯式"))
                lblInvNoValue = .Fields("InvNo") & ""
                lblInvTitleValue = .Fields("InvTitle") & ""
                lblInvAddressValue = .Fields("InvAddress") & ""
                lblChargeTitleValue = .Fields("ChargeTitle") & ""
                '#3982 增加郵寄地址顯示 By Kin 2008/07/16
                lblMailAddressValue = .Fields("MailAddress") & ""
                If Val(.Fields("InvoiceKind") & "") = 0 Then
'                    lblInvoiceKind = "電子計算機發票"
                    lblInvoiceKind = "電子發票證明聯"
                Else
                    lblInvoiceKind = "電子發票"
                End If
                '#6736 增加顯示會員載具顯碼與隱碼 By Kin 2014/03/12
                lblA_CarrierId1.Caption = ""
                lblA_CarrierId2.Caption = ""
                If Len(.Fields("A_CarrierId1") & "") > 0 Then
                    lblA_CarrierId1.Caption = .Fields("A_CarrierId1") & ""
                End If
                If Len(.Fields("A_CarrierId2") & "") > 0 Then
                    lblA_CarrierId2.Caption = .Fields("A_CarrierId2") & ""
                End If
            End With
        End If
    End If

  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "OpenInvTitle")
End Sub


Public Property Let uServiceType(ByVal vNewValue As String)
    strServiceType = vNewValue
End Property

Public Property Get uServiceType() As String
    uServiceType = strServiceType
End Property

Public Property Let uCustName(ByVal vNewValue As String)
    strCustName = vNewValue
End Property
Public Property Let uSO1131FSave(ByVal vData As Boolean)
    blnSO1131FSave = vData
End Property
Public Function GetInvoiceNo3(StrTableName As String) As String
    On Error Resume Next
    Dim strInv As String
    Dim strSeq As String
        strSeq = GetOwner & "S_" & StrTableName
        strInv = GetRsValue("SELECT " & strSeq & ".NextVal FROM Dual", gcnGi) & ""
        GetInvoiceNo3 = strInv
End Function
'#4150 增加置換收費項目 By Kin 2008/10/17
Public Property Let uChgCitem(ByVal vData As String)
  On Error GoTo ChkErr
    strChgCitem = vData
    Exit Property
ChkErr:
  Call ErrSub(Me.Name, "uChgCitem")
End Property
Property Set uRS(ByRef vData As ADODB.Recordset)
  On Error GoTo ChkErr
    Set rsFROM = vData
    Exit Property
ChkErr:
  Call ErrSub(Me.Name, "uRS")
End Property
