VERSION 5.00
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO3320A 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "收款報表查詢"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   975
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3320A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.PictureBox picFilePath 
      BackColor       =   &H00C0FFC0&
      Height          =   1305
      Left            =   3400
      ScaleHeight     =   1245
      ScaleWidth      =   5055
      TabIndex        =   94
      Top             =   3195
      Visible         =   0   'False
      Width           =   5115
      Begin VB.CommandButton cmdPath 
         Caption         =   "..."
         Height          =   375
         Left            =   4650
         TabIndex        =   98
         Top             =   150
         Width           =   315
      End
      Begin VB.TextBox txtFilePath 
         Height          =   360
         Left            =   1665
         TabIndex        =   97
         Top             =   165
         Width           =   2970
      End
      Begin VB.CommandButton cmdExeclOk 
         Caption         =   "F2 確定"
         Height          =   375
         Left            =   150
         TabIndex        =   96
         Top             =   735
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&X)"
         Height          =   375
         Left            =   3660
         TabIndex        =   95
         Top             =   735
         Width           =   1215
      End
      Begin VB.Label lblEmailList 
         AutoSize        =   -1  'True
         Caption         =   "輸出路徑(含檔名)"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   150
         TabIndex        =   99
         Top             =   255
         Width           =   1485
      End
   End
   Begin VB.PictureBox pic2 
      Height          =   4035
      Left            =   120
      ScaleHeight     =   3975
      ScaleWidth      =   11535
      TabIndex        =   87
      Top             =   1950
      Width           =   11595
      Begin VB.VScrollBar vsl1 
         Height          =   3915
         LargeChange     =   50
         Left            =   11220
         Max             =   100
         SmallChange     =   20
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.Frame fraMulti 
         BorderStyle     =   0  '沒有框線
         Height          =   8295
         Left            =   -30
         TabIndex        =   88
         Top             =   0
         Width           =   11445
         Begin CS_Multi.CSmulti gimUpdEn 
            Height          =   345
            Left            =   60
            TabIndex        =   42
            Top             =   5700
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   609
            ButtonCaption   =   "異動人員"
         End
         Begin CS_Multi.CSmulti gimClctEn 
            Height          =   375
            Left            =   60
            TabIndex        =   38
            Top             =   4290
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "收費人員"
         End
         Begin CS_Multi.CSmulti gimMduId 
            Height          =   345
            Left            =   60
            TabIndex        =   36
            Top             =   3630
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   609
            ButtonCaption   =   "大樓編號"
         End
         Begin CS_Multi.CSmulti gimStrtCode 
            Height          =   375
            Left            =   60
            TabIndex        =   40
            Top             =   5010
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "街道範圍"
         End
         Begin Gi_Multi.GiMulti gimBillType 
            Height          =   375
            Left            =   60
            TabIndex        =   27
            Top             =   390
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "單據類別"
            DataType        =   2
            FldCaption1     =   "單據類別代碼"
            FldCaption2     =   "單據類別名稱"
            DIY             =   -1  'True
            Exception       =   -1  'True
         End
         Begin Gi_Multi.GiMulti gimCMCode 
            Height          =   375
            Left            =   60
            TabIndex        =   28
            Top             =   750
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "收費方式"
         End
         Begin Gi_Multi.GiMulti gimCustStatus 
            Height          =   375
            Left            =   60
            TabIndex        =   32
            Top             =   2190
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "客戶狀態"
         End
         Begin Gi_Multi.GiMulti gimServCode 
            Height          =   375
            Left            =   60
            TabIndex        =   33
            Top             =   2550
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "服  務  區"
         End
         Begin Gi_Multi.GiMulti gimAreaCode 
            Height          =   375
            Left            =   60
            TabIndex        =   34
            Top             =   2910
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "行  政  區"
         End
         Begin Gi_Multi.GiMulti gimSTCode 
            Height          =   345
            Left            =   60
            TabIndex        =   41
            Top             =   5370
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   609
            ButtonCaption   =   "短收原因"
         End
         Begin Gi_Multi.GiMulti gimOrder 
            Height          =   375
            Left            =   60
            TabIndex        =   46
            Top             =   7110
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "排序方式"
            ColumnOrder     =   -1  'True
         End
         Begin Gi_Multi.GiMulti gimBankCode 
            Height          =   375
            Left            =   60
            TabIndex        =   43
            Top             =   6030
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "銀行名稱"
         End
         Begin Gi_Multi.GiMulti gimWipCode 
            Height          =   375
            Left            =   60
            TabIndex        =   45
            Top             =   6750
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "派工類別"
         End
         Begin CS_Multi.CSmulti gimClassCode 
            Height          =   375
            Left            =   60
            TabIndex        =   31
            Top             =   1830
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "客戶類別"
         End
         Begin CS_Multi.CSmulti gimCitemCode 
            Height          =   375
            Left            =   60
            TabIndex        =   29
            Top             =   1110
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "收費項目"
         End
         Begin Gi_Multi.GiMulti gimPTCode 
            Height          =   375
            Left            =   60
            TabIndex        =   30
            Top             =   1470
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "付款種類"
         End
         Begin Gi_Multi.GiMulti gimCardCode 
            Height          =   375
            Left            =   60
            TabIndex        =   44
            Top             =   6390
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "信用卡種類"
         End
         Begin Gi_Multi.GiMulti gimServiceType 
            Height          =   375
            Left            =   60
            TabIndex        =   26
            Top             =   30
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "服務類別"
            DataType        =   2
            FldCaption1     =   "單據類別代碼"
            FldCaption2     =   "單據類別名稱"
         End
         Begin Gi_Multi.GiMulti gimWorkClass 
            Height          =   345
            Left            =   60
            TabIndex        =   37
            Top             =   3960
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   609
            ButtonCaption   =   "工作類別"
         End
         Begin Gi_Multi.GiMulti gimClctAreaCode 
            Height          =   375
            Left            =   60
            TabIndex        =   35
            Top             =   3270
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "收  費  區"
         End
         Begin CS_Multi.CSmulti gimOldClctEn 
            Height          =   375
            Left            =   60
            TabIndex        =   39
            Top             =   4650
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "原收費人員"
         End
         Begin CS_Multi.CSmulti gimCM001 
            Height          =   375
            Left            =   60
            TabIndex        =   107
            Top             =   7500
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   661
            ButtonCaption   =   "會計科目"
         End
         Begin Gi_Multi.GiMulti gimPayType 
            Height          =   420
            Left            =   90
            TabIndex        =   108
            Top             =   7860
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   741
            ButtonCaption   =   "繳付類別"
         End
      End
   End
   Begin VB.TextBox txtRealAmt2 
      Alignment       =   1  '靠右對齊
      Height          =   285
      Left            =   8220
      MaxLength       =   8
      TabIndex        =   13
      Top             =   450
      Width           =   735
   End
   Begin VB.TextBox txtRealAmt1 
      Alignment       =   1  '靠右對齊
      Height          =   285
      Left            =   7320
      MaxLength       =   8
      TabIndex        =   12
      Top             =   450
      Width           =   765
   End
   Begin VB.TextBox txtBillNo2 
      Height          =   315
      Left            =   3000
      MaxLength       =   15
      TabIndex        =   25
      Top             =   1590
      Width           =   1485
   End
   Begin VB.ComboBox cboBillType 
      Height          =   315
      ItemData        =   "SO3320A.frx":0442
      Left            =   120
      List            =   "SO3320A.frx":044F
      Style           =   2  '單純下拉式
      TabIndex        =   23
      Top             =   1590
      Width           =   1185
   End
   Begin VB.TextBox txtBillNo1 
      Height          =   315
      Left            =   1290
      MaxLength       =   15
      TabIndex        =   24
      Top             =   1590
      Width           =   1485
   End
   Begin VB.ComboBox cboFirstFlag 
      Height          =   315
      ItemData        =   "SO3320A.frx":0471
      Left            =   10020
      List            =   "SO3320A.frx":047E
      Style           =   2  '單純下拉式
      TabIndex        =   22
      Top             =   1170
      Width           =   1695
   End
   Begin VB.CommandButton cmdExeclDetail 
      Caption         =   "將明細匯成Excel"
      Height          =   375
      Left            =   4650
      TabIndex        =   100
      Top             =   7410
      Width           =   1905
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1110
      Left            =   3400
      ScaleHeight     =   1050
      ScaleWidth      =   5055
      TabIndex        =   76
      Top             =   3360
      Visible         =   0   'False
      Width           =   5115
      Begin MSComctlLib.ProgressBar pgbBar 
         Height          =   375
         Left            =   165
         TabIndex        =   77
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
         TabIndex        =   78
         Top             =   135
         Width           =   1800
      End
   End
   Begin prjGiList.GiList gilIntroId 
      Height          =   315
      Left            =   9090
      TabIndex        =   51
      Top             =   6030
      Width           =   2625
      _ExtentX        =   4630
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
      FldWidth1       =   700
      FldWidth2       =   1600
      F5Corresponding =   -1  'True
   End
   Begin VB.TextBox txtCustId 
      Height          =   315
      Left            =   4650
      TabIndex        =   52
      Top             =   6450
      Width           =   3345
   End
   Begin Gi_Time.GiTime gdtUpdTime2 
      Height          =   315
      Left            =   6780
      TabIndex        =   15
      Top             =   825
      Width           =   2175
      _ExtentX        =   3836
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
   Begin Gi_Time.GiTime gdtUpdTime1 
      Height          =   315
      Left            =   4380
      TabIndex        =   14
      Top             =   825
      Width           =   2085
      _ExtentX        =   3678
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
   Begin VB.Frame fraCircuitNo 
      Height          =   465
      Left            =   6210
      TabIndex        =   84
      Top             =   5940
      Width           =   1785
      Begin VB.OptionButton optCircuitNoAll 
         Caption         =   "全部"
         Height          =   195
         Left            =   900
         TabIndex        =   50
         Top             =   180
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optCircuitIsNull 
         Caption         =   "空白"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   180
         Width           =   735
      End
   End
   Begin VB.TextBox txtCircuitNo 
      Height          =   315
      Left            =   4650
      MaxLength       =   20
      TabIndex        =   48
      Top             =   6060
      Width           =   1485
   End
   Begin VB.Frame fraInvoiceType 
      Caption         =   "發票種類"
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   3810
      TabIndex        =   75
      Top             =   6810
      Width           =   4185
      Begin VB.OptionButton optInvoiceBack 
         Caption         =   "後開"
         Height          =   195
         Left            =   1350
         TabIndex        =   58
         Top             =   210
         Width           =   765
      End
      Begin VB.OptionButton optInvoiceForward 
         Caption         =   "預開"
         Height          =   195
         Left            =   330
         TabIndex        =   57
         Top             =   210
         Width           =   765
      End
      Begin VB.OptionButton optInvoiceAll 
         Caption         =   "全部"
         Height          =   195
         Left            =   2340
         TabIndex        =   59
         Top             =   210
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin VB.CheckBox chkInvoice 
      Caption         =   "是否已拋檔"
      Height          =   195
      Left            =   9180
      TabIndex        =   19
      Top             =   330
      Width           =   2325
   End
   Begin VB.CheckBox chkGUINo 
      Caption         =   "是否已開發票"
      Height          =   195
      Left            =   9180
      TabIndex        =   18
      Top             =   105
      Width           =   2325
   End
   Begin VB.CheckBox chkBad 
      Caption         =   "是否包含呆帳資料統計"
      Height          =   195
      Left            =   9180
      TabIndex        =   20
      Top             =   570
      Value           =   1  '核取
      Width           =   2325
   End
   Begin VB.Frame fraClose 
      Caption         =   "選擇條件"
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   150
      TabIndex        =   73
      Top             =   6810
      Width           =   3525
      Begin VB.OptionButton optClose_All 
         Caption         =   "全部"
         Height          =   195
         Left            =   2400
         TabIndex        =   56
         Top             =   210
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optNoClose 
         Caption         =   "未日結"
         Height          =   195
         Left            =   1350
         TabIndex        =   55
         Top             =   210
         Width           =   1065
      End
      Begin VB.OptionButton optClose 
         Caption         =   "已日結"
         Height          =   195
         Left            =   360
         TabIndex        =   54
         Top             =   210
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      Height          =   375
      Left            =   180
      TabIndex        =   62
      Top             =   7440
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   10380
      TabIndex        =   64
      Top             =   7410
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   375
      Left            =   1920
      TabIndex        =   63
      Top             =   7410
      Width           =   1455
   End
   Begin Gi_Date.GiDate gdaRealDate1 
      Height          =   315
      Left            =   1020
      TabIndex        =   2
      Top             =   460
      Width           =   1065
      _ExtentX        =   1879
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
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1020
      TabIndex        =   0
      Top             =   90
      Width           =   2475
      _ExtentX        =   4366
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
      FldWidth2       =   1550
      F5Corresponding =   -1  'True
   End
   Begin VB.Frame fraRptOpt 
      Caption         =   "報表格式"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8130
      TabIndex        =   65
      Top             =   6810
      Width           =   3585
      Begin VB.OptionButton optSimple 
         Caption         =   "簡表"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1590
         TabIndex        =   61
         Top             =   210
         Width           =   885
      End
      Begin VB.OptionButton optDetail 
         Caption         =   "明細表"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   240
         TabIndex        =   60
         Top             =   240
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin Gi_Date.GiDate gdaRealDate2 
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   460
      Width           =   1065
      _ExtentX        =   1879
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
   Begin Gi_Date.GiDate gdaInstDay1 
      Height          =   315
      Left            =   1020
      TabIndex        =   6
      Top             =   830
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      ForeColor       =   12583104
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
   Begin Gi_Date.GiDate gdaShouldDay1 
      Height          =   315
      Left            =   1020
      TabIndex        =   8
      Top             =   1200
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      ForeColor       =   8388736
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
   Begin Gi_Date.GiDate gdaInstDay2 
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   830
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      ForeColor       =   12583104
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
   Begin Gi_Date.GiDate gdaShouldDay2 
      Height          =   315
      Left            =   2400
      TabIndex        =   9
      Top             =   1200
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      ForeColor       =   8388736
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
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   4380
      TabIndex        =   1
      Top             =   90
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      FldWidth2       =   1500
      F5Corresponding =   -1  'True
   End
   Begin VB.ComboBox cboAddress 
      Height          =   315
      ItemData        =   "SO3320A.frx":04A2
      Left            =   1020
      List            =   "SO3320A.frx":04AF
      Style           =   2  '單純下拉式
      TabIndex        =   47
      Top             =   6060
      Width           =   2565
   End
   Begin VB.CheckBox chkUseOldClctEn 
      Caption         =   "列印顯示原收費人員"
      Height          =   195
      Left            =   9180
      TabIndex        =   21
      Top             =   900
      Width           =   2325
   End
   Begin prjNumber.GiNumber ginRealPeriod2 
      Height          =   285
      Left            =   8220
      TabIndex        =   11
      Top             =   90
      Width           =   735
      _ExtentX        =   1296
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
      ForeColor       =   0
      MaxLength       =   3
      AutoSelect      =   -1  'True
      AllowZero       =   0   'False
   End
   Begin prjNumber.GiNumber ginRealPeriod1 
      Height          =   285
      Left            =   7320
      TabIndex        =   10
      Top             =   90
      Width           =   765
      _ExtentX        =   1349
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
      ForeColor       =   0
      MaxLength       =   3
      AutoSelect      =   -1  'True
      AllowZero       =   0   'False
   End
   Begin Gi_Time.GiTime gdtCreateTime2 
      Height          =   315
      Left            =   6780
      TabIndex        =   17
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
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
   Begin Gi_Time.GiTime gdtCreateTime1 
      Height          =   315
      Left            =   4380
      TabIndex        =   16
      Top             =   1200
      Width           =   2085
      _ExtentX        =   3678
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
   Begin VB.ComboBox cboGroup 
      Height          =   315
      ItemData        =   "SO3320A.frx":04DF
      Left            =   9090
      List            =   "SO3320A.frx":04FB
      Style           =   2  '單純下拉式
      TabIndex        =   53
      Top             =   6450
      Width           =   2625
   End
   Begin MSComDlg.CommonDialog comdPath 
      Left            =   90
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Gi_Date.GiDate gdaClctDate1 
      Height          =   315
      Left            =   4380
      TabIndex        =   4
      Top             =   465
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      ForeColor       =   12583104
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
   Begin Gi_Date.GiDate gdaClctDate2 
      Height          =   315
      Left            =   5730
      TabIndex        =   5
      Top             =   460
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      ForeColor       =   12583104
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "至"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   5490
      TabIndex        =   106
      Top             =   540
      Width           =   195
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "下  收  日"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   3570
      TabIndex        =   105
      Top             =   540
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "至"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   2820
      TabIndex        =   104
      Top             =   1680
      Width           =   195
   End
   Begin VB.Label lblFirstFlag 
      AutoSize        =   -1  'True
      Caption         =   "首期條件"
      Height          =   195
      Left            =   9180
      TabIndex        =   103
      Top             =   1260
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "至"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   6540
      TabIndex        =   102
      Top             =   1275
      Width           =   195
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "產生時間"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   3570
      TabIndex        =   101
      Top             =   1275
      Width           =   780
   End
   Begin VB.Label lblGroup 
      Caption         =   "分頁方式"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   8160
      TabIndex        =   93
      Top             =   6480
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "~"
      Height          =   195
      Left            =   8100
      TabIndex        =   92
      Top             =   480
      Width           =   105
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "~"
      Height          =   195
      Left            =   8100
      TabIndex        =   91
      Top             =   150
      Width           =   105
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      Caption         =   "地址依據"
      Height          =   195
      Left            =   150
      TabIndex        =   90
      Top             =   6120
      Width           =   780
   End
   Begin VB.Label lblIntroId 
      AutoSize        =   -1  'True
      Caption         =   "銷售點"
      Height          =   195
      Left            =   8160
      TabIndex        =   86
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label lblCustId 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號(以,相隔或以-取範圍)，例如: 1,2,3,5,11,20-33"
      Height          =   195
      Left            =   150
      TabIndex        =   85
      Top             =   6510
      Width           =   4425
   End
   Begin VB.Label lblCircuitNo 
      AutoSize        =   -1  'True
      Caption         =   "網路編號"
      Height          =   195
      Left            =   3795
      TabIndex        =   83
      Top             =   6150
      Width           =   780
   End
   Begin VB.Label lblRealAmt 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "金額"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   6870
      TabIndex        =   82
      Top             =   480
      Width           =   390
   End
   Begin VB.Label lblPeriod 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "期數"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   6870
      TabIndex        =   81
      Top             =   150
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "異動時間"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   3570
      TabIndex        =   80
      Top             =   900
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "至"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   6540
      TabIndex        =   79
      Top             =   900
      Width           =   195
   End
   Begin VB.Label lblServiceType 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "服務類別"
      Height          =   195
      Left            =   3570
      TabIndex        =   74
      Top             =   150
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label lblComp 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   150
      TabIndex        =   72
      Top             =   150
      Width           =   765
   End
   Begin VB.Label lblD3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "至"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   2160
      TabIndex        =   71
      Top             =   1290
      Width           =   195
   End
   Begin VB.Label lblShouldDay 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "應收日期"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   150
      TabIndex        =   70
      Top             =   1260
      Width           =   780
   End
   Begin VB.Label lblD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "至"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2160
      TabIndex        =   69
      Top             =   510
      Width           =   195
   End
   Begin VB.Label lblRealDay 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "實收日期"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   150
      TabIndex        =   68
      Top             =   510
      Width           =   780
   End
   Begin VB.Label lblD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "至"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   2160
      TabIndex        =   67
      Top             =   900
      Width           =   195
   End
   Begin VB.Label lblInstDay 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "裝機日期"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   150
      TabIndex        =   66
      Top             =   900
      Width           =   780
   End
End
Attribute VB_Name = "frmSO3320A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SimCount = 12
Private strForm As String
Private strGroupName As String
Private lngAffected As Long
Private strChooseSO002A As String
Dim strMDBOrder As String

Private Sub cboBillType_Click()
    Dim lngLength As Long
    Select Case cboBillType.ListIndex
        Case 0
            lngLength = 15
        Case 1
            lngLength = 12
        Case 2
            lngLength = 11
    End Select
    txtBillNo1.MaxLength = lngLength
    txtBillNo2.MaxLength = lngLength
End Sub

Private Sub cmdCancel_Click()
    Dim obj As Object
'    For Each obj In Me
'        On Error Resume Next
'        obj.Enabled = True
'    Next
    cmdPrint.Enabled = True
    picFilePath.Visible = False
    cmdPrevRpt.Enabled = True
End Sub

Private Sub cmdExeclDetail_Click()
    On Error Resume Next
    Dim obj As Object
'    For Each obj In Me
'        On Error Resume Next
'        If Not TypeOf obj Is Label Then
'            obj.Enabled = False
'            Select Case obj.Name
'                Case "cmdExeclOk", "cmdCancel", "cmdPath", "txtFilePath", "picFilePath"
'                    obj.Enabled = True
'            End Select
'        End If
'        DoEvents
'    Next
    cmdPrint.Enabled = False
    picFilePath.Visible = True
    picFilePath.ZOrder 0
    cmdPrevRpt.Enabled = False
End Sub

Private Sub cmdExeclOk_Click()
    On Error GoTo chkErr
    Dim lngAffected As Long
    Dim objExecl As Object
    Dim obj As Object
    Dim strHead As String
        If Not MustExist(txtFilePath, , "輸出路徑") Then Exit Sub
        Me.Enabled = False
        Screen.MousePointer = vbHourglass
        picFilePath.Visible = False
        Call subChoose
        Picture1.ZOrder 0
        Call DetialReport(lngAffected, False)
        '#4289 要增加BpCode、BpName By Kin 2008/12/18
        '#5683 增加繳付類別 By Kin 2010/08/16
        If lngAffected > 0 Then
            strHead = "BillNo as 單據編號,CustId as 客戶編號,CustName as 客戶姓名,Tel1 as 電話1,Tel3 as 電話3,CitemCode as 收費項目代碼,CitemName as 收費項目, " & _
                             "RealDate as 實收日期,RealPeriod as 實收期數,RealAmt as 實收金額," & _
                            "ClctName as 收費人員,Address as 地址,RealStartDate as 實收起始日,RealStopDate as 實收截止日," & _
                            "GUINo as 發票號碼,MduName as 大樓名稱,ServName as 服務區,AreaName as 行政區,CMCode as 收費方式代碼,CMName as 收費方式," & _
                            "ClctEn as 收費人員代碼,NoTaxAmt as 未稅金額,CustStatusName as 客戶狀態,ClctDate as 下收日,AddrSort as 地址排序字串," & _
                            "WipCode1 as 派工狀態代碼1,WipName1 as 派工狀態1,OldClctName as 舊客戶狀態,Note as 備註, " & _
                            "OldClctEn as 原收費人員代碼,AreaCode as 行政區代碼,ServCode as 服務區代碼,InvNo as 統一編號,ClctAreaCode as 收費區代碼,ClctAreaName as 收費區," & _
                            "UpdTime as 異動時間,MduCount as 大樓有效戶,STName as 短收原因,ClctEnWorkClass as 收費員工作類別,CardType as 信用卡別,ManualNo as 手開單號,AccountNo as 帳號,MediaBillNo as 媒體單號,ClassName as 客戶類別," & _
                            "BpCode as 優惠組合代碼,BpName as 優惠組合名稱,iif(PayKind=0,'預付制','現付制') as 繳付類別"
            
            'strHead = ",單據編號,客戶姓名,電話1,收費項目,實收日期,實收期數,實收金額,收費人員,地址,實收起始日,截止日," & _
            "發票號碼,大樓名稱,服務區,行政區,收費方式代碼,收費方式名稱,收費人員代碼,收費項目代碼,未稅金額,客戶狀態,下收日," & _
            "地址排序字串,派工狀態1,派工狀態代碼1,舊客戶狀態,備註,原收費人員代碼,行政區代碼,服務區代碼,服務類別,統一編號,收費區代碼," & _
            "異動時間,大樓有效戶,短收原因,大樓編號,收費員工作類別,帳號,序號,信用卡別,手開單號,收費區"
            If strMDBOrder <> "" Then cnn.Execute ("Create Index I_SO3320A1_XX On SO3320A1(" & Replace(UCase(strMDBOrder), "SO3320A1.", "") & ")")
            If ExporttoExcel("Select " & strHead & " From SO3320A1 " & IIf(strMDBOrder = "", "", "Order By ") & strMDBOrder, , txtFilePath, "收款明細表", True, cnn) Then MsgBox "匯出成功 !!", vbInformation, gimsgPrompt
            'If ExporttoExcel("Select * From SO3320A1", , txtFilePath, "收款明細表", True, cnn) Then MsgBox "匯出成功 !!", vbInformation, gimsgPrompt
        End If
'        For Each obj In Me
'            On Error Resume Next
'            obj.Enabled = True
'        Next
        Me.Enabled = True
        cmdPrint.Enabled = True
        cmdPrevRpt.Enabled = True
        Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdExeclOk_Click"
End Sub

Private Sub cmdExit_Click()
  On Error GoTo chkErr
    Unload Me
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdPath_Click()
    On Error GoTo chkErr
    Dim strFilePath As String
        With comdPath
            .DialogTitle = "選擇輸出路徑"
            .Filter = "Excel(*.xls)|*.xls|全部(*.*)|*.*"
            .Action = 1
            If .FileName <> "" Then txtFilePath.Text = .FileName
            
        End With
    Exit Sub
chkErr:
   ErrSub Me.Name, "cmdPath_Click"
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo chkErr
    If optDetail.Value = True Then
        Call PreviousRpt(GetPrinterName(5), RptName("SO3320", 1), "收款明細表 [SO3320A1]")
    Else
        Call PreviousRpt(GetPrinterName(5), RptName("SO3320", 2), "收款明細表 [SO3320A2]")
    End If
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo chkErr
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strSubSQL1 As String
    Dim strSubSQL2 As String
    Dim strOrder As String
    Dim strSubOrder As String
        Screen.MousePointer = vbHourglass
        If Not IsDataOk Then Exit Sub
        ReadyGoPrint
        Me.Picture1.ZOrder 0
        Me.Enabled = False
        lngAffected = 0
        subChoose
        If optSimple.Value Then
            Call SimpleReport
        Else
            Call DetialReport
        End If
        Me.Enabled = True
        Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function subInsertMDB1(strSQL As String) As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim blnFlag As Boolean, lngCITEMCODE As Long
    Dim lngRate1 As Long, strClctDate As Variant
    Dim blnSO002 As Boolean, strInvNo As String, lngMduCount As Long
    Dim lngCustId As Long, strClctEnWorkClass As String
    Dim rsCM003 As New ADODB.Recordset
    Dim strInstrMDB As String
    Dim strCardType As String
    Dim strSO003sql As String
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        '工作類別(收費人員) 取資料
        'If Not GetRS(rsCM003, "Select EmpNo, WorkClass From " & GetOwner & "CM003 ") Then Exit Function
        If Not GetRS(rsCM003, "Select A.EmpNo, A.WorkClass,B.Description WorkClassName From " & GetOwner & "CM003 A," & GetOwner & "CD071 B Where A.WorkClass=B.CodeNo(+) ") Then Exit Function
        
        If Not CreateMDBTable(rsTmp, "SO3320A1", cnn) Then Exit Function
        If rsTmp.RecordCount > 0 Then Picture1.Visible = True
        SendSQL rsTmp.Source
        blnSO002 = InStr(1, UCase(strSQL), "SO002") > 0
        cnn.BeginTrans
        While Not rsTmp.EOF
            lngMduCount = 0
            If rsTmp("MduName") & "" <> "" Then lngMduCount = Val(GetRsValue("Select InstCnt From " & GetOwner & "SO017 Where MduId = '" & rsTmp("MduId") & "' And CompCode = " & gilCompCode.GetCodeNo) & "")
            
            If lngCITEMCODE <> Val(rsTmp("CitemCode") & "") Then
                lngCITEMCODE = Val(rsTmp("CitemCode") & "")
                lngRate1 = Val(GetRsValue("Select Nvl(Rate1,0) Rate1 From " & GetOwner & "CD019," & GetOwner & "CD033 Where " & _
                                " CD019.CodeNo = " & lngCITEMCODE & " and CD019.TaxCode = CD033.CodeNo") & "")
            End If
            If blnSO002 = False Then strInvNo = GetRsValue("Select InvNo From " & GetOwner & "So002 Where CustId = " & rsTmp("CustId") & " And ServiceType = '" & rsTmp("ServiceType") & "'") & ""
            '工作類別(收費人員)
            strClctEnWorkClass = ""
            If rsTmp("ClctEn") & "" <> "" Then
                rsCM003.Filter = "EmpNo = '" & rsTmp("ClctEn") & "'"
                'If Not rsCM003.EOF Then strClctEnWorkClass = GetWorkclass(rsCM003("WorkClass") & "")
                If Not rsCM003.EOF Then strClctEnWorkClass = rsCM003("WorkClassName") & ""
            End If
            strCardType = ""
            If rsTmp("AccountNo") & "" <> "" Then
                If rsTmp("CardType") & "" <> "" Then
                    strCardType = rsTmp("CardType")
                Else
                    strCardType = GetRsValue("Select Distinct CardName From " & GetOwner & "SO106 " & _
                                        " Where CustId = " & rsTmp("CustId") & " And AccountID = '" & rsTmp("AccountNo") & "'" & _
                                        " AND StopFlag<> 1" & _
                                        IIf(strChooseSO002A <> "", "And ", "") & strChooseSO002A) & ""
                End If
            End If
            
            'If lngCustId <> rsTmp("CustId") Or lngCITEMCODE <> Val(rsTmp("CitemCode") & "") Then
'                strSO003sql = "Select ClctDate From " & GetOwner & "So003 where CustId =" & rsTmp("CustId") & " And CitemCode = " & rsTmp("CitemCode") & " And Servicetype ='" & rsTmp("ServiceType") & "'"
'                'If rsTmp("SeqNo") <> "" Then strSO003sql = strSO003sql & " And SeqNo = " & rsTmp("SeqNo") & ""
'                If rsTmp("FaciSeqNo") <> "" Then strSO003sql = strSO003sql & " And FaciSeqNo = '" & rsTmp("FaciSeqNo") & "'"
'                strClctDate = GetRsValue(strSO003sql)
            'End If
            '#4289要增加BpCode、BpName By Kin 2008/12/18
            '#5683 增加PayKind By Kin 2010/08/16
            strInstrMDB = "Insert Into SO3320A1 (CustId,BillNo,CustName,Tel1,CitemName,RealDate,RealPeriod,RealAmt," & _
                            "ClctName,Address,RealStartDate,RealStopDate,GUINo,MduName,ServName,AreaName,CMCode,CMName," & _
                            "ClctEn,CitemCode,NoTaxAmt,CustStatusName,ClctDate,AddrSort,WipName1,WipCode1,OldClctName,[Note], " & _
                            "OldClctEn,AreaCode,ServCode,InvNo,ClctAreaCode,ClctAreaName,UpdTime,MduCount,STName,ClctEnWorkClass," & _
                            "CardType,ManualNo,AccountNo,FaciSeqNo,DiscountCode,DiscountName,DiscountDate1,DiscountDate2,MediaBillNo," & _
                            "Tel3,ClassCode,ClassName,PTCode,PTName,BpCode,BpName,PayKind) Values (" & _
                            GetNullString(rsTmp("CustId")) & "," & GetNullString(rsTmp("BillNo")) & "," & _
                            GetNullString(rsTmp("CustName")) & "," & GetNullString(rsTmp(3)) & "," & _
                            GetNullString(rsTmp(4)) & "," & GetNullString(rsTmp(5)) & "," & GetNullString(rsTmp(6)) & "," & GetNullString(rsTmp(7)) & "," & _
                            GetNullString(rsTmp(8)) & "," & GetNullString(rsTmp(9)) & "," & GetNullString(rsTmp(10)) & "," & GetNullString(rsTmp(11)) & "," & _
                            GetNullString(rsTmp(12)) & "," & GetNullString(rsTmp(13)) & "," & GetNullString(rsTmp(14)) & "," & GetNullString(rsTmp(15)) & "," & _
                            GetNullString(rsTmp(16)) & "," & GetNullString(rsTmp(17)) & "," & GetNullString(rsTmp(18)) & "," & GetNullString(rsTmp(19)) & "," & _
                            Round(Val(rsTmp("RealAmt") & "") / (1 + lngRate1 / 100)) & ","
                            
            strInstrMDB = strInstrMDB & GetNullString(rsTmp("CustStatusName")) & "," & GetNullString(rsTmp("ClctDate"), giDateV, giAccessDb) & "," & _
                            GetNullString(rsTmp("AddrSort")) & "," & GetNullString(rsTmp("WipName1")) & "," & _
                            GetNullString(rsTmp("WipCode1")) & "," & GetNullString(rsTmp("OldClctName")) & "," & _
                            GetNullString(rsTmp("Note")) & "," & GetNullString(rsTmp("OldClctEn")) & "," & _
                            GetNullString(rsTmp("AreaCode")) & "," & GetNullString(rsTmp("ServCode")) & "," & _
                            GetNullString(strInvNo) & "," & GetNullString(rsTmp("ClctAreaCode")) & "," & _
                            GetNullString(rsTmp("ClctAreaName")) & "," & GetNullString(rsTmp("UpdTime")) & "," & _
                            lngMduCount & "," & GetNullString(rsTmp("STName")) & "," & GetNullString(strClctEnWorkClass) & "," & _
                            GetNullString(strCardType) & "," & GetNullString(rsTmp("ManualNo")) & "," & _
                            GetNullString(rsTmp("AccountNo")) & "," & GetNullString(rsTmp("FaciSeqNo")) & "," & _
                            GetNullString(rsTmp("DiscountCode")) & "," & GetNullString(rsTmp("DiscountName")) & "," & _
                            GetNullString(rsTmp("DiscountDate1")) & "," & GetNullString(rsTmp("DiscountDate2")) & "," & _
                            GetNullString(rsTmp("MediaBillNo")) & "," & GetNullString(rsTmp("Tel3")) & "," & _
                            GetNullString(rsTmp("ClassCode")) & "," & GetNullString(rsTmp("ClassName")) & "," & _
                            GetNullString(rsTmp("PTCode")) & "," & GetNullString(rsTmp("PTName")) & "," & _
                            GetNullString(rsTmp("BpCode")) & "," & GetNullString(rsTmp("BpName")) & "," & _
                            GetNullString(rsTmp("PayKind")) & ")"
            cnn.Execute Replace(strInstrMDB, Chr(0), "")
            
            
'            cnn.Execute Replace("Insert Into SO3320A1 (CustId,BillNo,CustName,Tel1,CitemName,RealDate,RealPeriod,RealAmt," & _
'                            "ClctName,Address,RealStartDate,RealStopDate,GUINo,MduName,ServName,AreaName,CMCode,CMName," & _
'                            "ClctEn,CitemCode,NoTaxAmt,CustStatusName,ClctDate,AddrSort,WipName1,WipCode1,OldClctName,[Note], " & _
'                            "OldClctEn,AreaCode,ServCode,InvNo,ClctAreaCode,ClctAreaName,UpdTime,MduCount,STName,ClctEnWorkClass," & _
'                            "CardType,ManualNo,AccountNo,FaciSeqNo,DiscountCode,DiscountName,DiscountDate1,DiscountDate2,MediaBillNo,Tel3,ClassCode,ClassName,PTCode,PTName) Values (" & _
'                            GetNullString(rsTmp("CustId")) & "," & GetNullString(rsTmp("BillNo")) & "," & _
'                            GetNullString(rsTmp("CustName")) & "," & GetNullString(rsTmp(3)) & "," & _
'                            GetNullString(rsTmp(4)) & "," & GetNullString(rsTmp(5)) & "," & GetNullString(rsTmp(6)) & "," & GetNullString(rsTmp(7)) & "," & _
'                            GetNullString(rsTmp(8)) & "," & GetNullString(rsTmp(9)) & "," & GetNullString(rsTmp(10)) & "," & GetNullString(rsTmp(11)) & "," & _
'                            GetNullString(rsTmp(12)) & "," & GetNullString(rsTmp(13)) & "," & GetNullString(rsTmp(14)) & "," & GetNullString(rsTmp(15)) & "," & _
'                            GetNullString(rsTmp(16)) & "," & GetNullString(rsTmp(17)) & "," & GetNullString(rsTmp(18)) & "," & GetNullString(rsTmp(19)) & "," & _
'                            Round(Val(rsTmp("RealAmt") & "") / (1 + lngRate1 / 100)) & "," & _
'                            GetNullString(rsTmp("CustStatusName")) & "," & GetNullString(rsTmp("ClctDate"), giDateV, giAccessDb) & "," & _
'                            GetNullString(rsTmp("AddrSort")) & "," & GetNullString(rsTmp("WipName1")) & "," & _
'                            GetNullString(rsTmp("WipCode1")) & "," & GetNullString(rsTmp("OldClctName")) & "," & _
'                            GetNullString(rsTmp("Note")) & "," & GetNullString(rsTmp("OldClctEn")) & "," & _
'                            GetNullString(rsTmp("AreaCode")) & "," & GetNullString(rsTmp("ServCode")) & "," & _
'                            GetNullString(strInvNo) & "," & GetNullString(rsTmp("ClctAreaCode")) & "," & _
'                            GetNullString(rsTmp("ClctAreaName")) & "," & GetNullString(rsTmp("UpdTime")) & "," & _
'                            lngMduCount & "," & GetNullString(rsTmp("STName")) & "," & GetNullString(strClctEnWorkClass) & "," & _
'                            GetNullString(strCardType) & "," & GetNullString(rsTmp("ManualNo")) & "," & _
'                            GetNullString(rsTmp("AccountNo")) & "," & GetNullString(rsTmp("FaciSeqNo")) & "," & _
'                            GetNullString(rsTmp("DiscountCode")) & "," & GetNullString(rsTmp("DiscountName")) & "," & _
'                            GetNullString(rsTmp("DiscountDate1")) & "," & GetNullString(rsTmp("DiscountDate2")) & "," & _
'                            GetNullString(rsTmp("MediaBillNo")) & "," & GetNullString(rsTmp("Tel3")) & "," & GetNullString(rsTmp("ClassCode")) & "," & GetNullString(rsTmp("ClassName")) & "," & GetNullString(rsTmp("PTCode")) & "," & GetNullString(rsTmp("PTName")) & ")", Chr(0), "")
            lngCustId = rsTmp("CustId")
            If (rsTmp.AbsolutePosition / rsTmp.RecordCount * 100) Mod 1 = 0 Then
                pgbBar.Value = 100 * rsTmp.AbsolutePosition / rsTmp.RecordCount
                DoEvents
            End If
            rsTmp.MoveNext
        Wend
        cnn.CommitTrans
        lngAffected = rsTmp.RecordCount
        Picture1.Visible = False
        Set rsTmp = Nothing
        subInsertMDB1 = True
    Exit Function
chkErr:
    cnn.RollbackTrans
    Call ErrSub(Me.Name, "subInsertMDB1")
End Function

Private Function GetWorkclass(strWorkClass As String) As String
    On Error Resume Next
        Select Case strWorkClass
            Case "A"
                GetWorkclass = "A(客服)"
            Case "B"
                GetWorkclass = "B(工程)"
            Case "C"
                GetWorkclass = "C(收費)"
            Case "D"
                GetWorkclass = "D(業務)"
            Case "E"
                GetWorkclass = "E(其他)"
        End Select
End Function
'#4015 簡表增加收費付款小計 By Kin 2008/07/18
Private Function subInsertMDB2B() As Boolean
  On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim blnFlag As Boolean, strSQL As String
        subInsertMDB2B = False
        On Error GoTo chkErr
        If optNoClose Or optClose_All Then
            If Not GetStrSQL2B(strSQL, "SO033") Then Exit Function
        End If
        If optClose Or optClose_All Then
            If Not GetStrSQL2B(strSQL, "SO034") Then Exit Function
            If Not GetStrSQL2B(strSQL, "SO035") Then Exit Function
        End If
        '#4030 增加實收金額,統計單數,統計客戶數 By Kin 2008/07/31
        strSQL = "Select CitemCode,CitemName,PTCode,PTName,Count(*) RcdNo,Sum(RealAmt) TotalAmt" & _
                 ",Count(Distinct CustId) CountCustId,Count(Distinct BillNo) BillCount From (" & _
                 strSQL & ") A Group By CitemCode,CitemName,PTCode,PTName"
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, "SO3320A2B", cnn) Then Exit Function
        
        SendSQL strSQL, True
        blnFlag = False
        While Not rsTmp.EOF
            cnn.Execute Replace("Insert Into SO3320A2B (CitemCode,CitemName,PTCode,PTName,RcdNo,TotalAmt,CountCustId,BillCount ) Values (" & _
                            GetNullString(rsTmp(0)) & "," & _
                            GetNullString(rsTmp(1)) & "," & GetNullString(rsTmp(2)) & "," & _
                            GetNullString(rsTmp(3)) & "," & GetNullString(rsTmp(4)) & "," & _
                            GetNullString(rsTmp(5)) & "," & GetNullString(rsTmp(6)) & "," & _
                            GetNullString(rsTmp(7)) & ")", Chr(0), "")
            'pgbBar.Value = (100 / SimCount) * 9 + (100 / SimCount) * rsTmp.AbsolutePosition / rsTmp.RecordCount
            pgbBar.Value = (100 / SimCount) * 10 + (100 / SimCount) * rsTmp.AbsolutePosition / rsTmp.RecordCount
            rsTmp.MoveNext
            DoEvents
        Wend
        lngAffected = rsTmp.RecordCount
        Set rsTmp = Nothing
        subInsertMDB2B = True
    Exit Function
chkErr:
  ErrSub Me.Name, "subInsertMDB2B"
End Function
'#5102 增加會計科目小計 By Kin 2009/06/10
Private Function subInsertMDB2C() As Boolean
  On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim blnFlag As Boolean, strSQL As String
        subInsertMDB2C = False
        On Error GoTo chkErr
        If optNoClose Or optClose_All Then
            If Not GetStrSQL2C(strSQL, "SO033") Then Exit Function
        End If
        If optClose Or optClose_All Then
            If Not GetStrSQL2C(strSQL, "SO034") Then Exit Function
            If Not GetStrSQL2C(strSQL, "SO035") Then Exit Function
        End If
        'strsql = "Select CodeNo,Description ,Sum(RealAmt) TotalAmt,Sum(Round(RealAmt / (1+(Nvl(Rate1,0)/100) ))) Tax,Count(*) RcdNo From (" & _
        'strsql & ") A Group By CodeNo,Description "
        strSQL = "Select CODENO,Description ,Sum(RealAmt) TotalAmt, " & _
        "Count(Distinct BillNo) BillCount,Count(Distinct CustId) CustCount,Count(*) IntCount From (" & _
        strSQL & ") A Group By CodeNo,Description Union All Select 'XXX' CodeNo,'合計' Description ,Sum(RealAmt) TotalAmt, " & _
        "Count(Distinct BillNo) BillCount,Count(Distinct CustId) CustCount,Count(*) IntCount From (" & _
        strSQL & ") A "
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, "SO3320A2C", cnn) Then Exit Function
        'If rsTmp.RecordCount > 0 Then Picture1.Visible = True
        SendSQL strSQL, True
        blnFlag = False
        While Not rsTmp.EOF
            cnn.Execute GetTmpMDBExecuteStr(rsTmp, "SO3320A2C")
'            cnn.Execute Replace("Insert Into SO3320A2C (CodeNo,Description,TotalAmt,Tax,RcdNo ) Values (" & _
'                            GetNullString(rsTmp(0)) & "," & _
'                            GetNullString(rsTmp(1)) & "," & GetNullString(rsTmp(2)) & "," & _
'                            GetNullString(rsTmp(3)) & "," & GetNullString(rsTmp(4)) & ")", Chr(0), "")
            'pgbBar.Value = (100 / SimCount) * rsTmp.AbsolutePosition / rsTmp.RecordCount
            pgbBar.Value = (100 / SimCount) * 11 + (100 / SimCount) * rsTmp.AbsolutePosition / rsTmp.RecordCount
            rsTmp.MoveNext
            DoEvents
        Wend
        'lngAffected = rsTmp.RecordCount
        Set rsTmp = Nothing
        subInsertMDB2C = True
    Exit Function
chkErr:
  Call ErrSub(Me.Name, "subInsertMDB2C")
End Function

Private Function subInsertMDB21() As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim blnFlag As Boolean, strSQL As String
        subInsertMDB21 = False
        On Error GoTo chkErr
        If optNoClose Or optClose_All Then
            If Not GetStrSQL2(strSQL, "SO033") Then Exit Function
        End If
        If optClose Or optClose_All Then
            If Not GetStrSQL2(strSQL, "SO034") Then Exit Function
            If Not GetStrSQL2(strSQL, "SO035") Then Exit Function
        End If
        strSQL = "Select CitemCode,CitemName,Sum(RealAmt) TotalAmt,Sum(Round(RealAmt / (1+(Nvl(Rate1,0)/100) ))) Tax,Count(*) RcdNo From (" & _
        strSQL & ") A Group By CitemCode,CitemName"
        
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, "SO3320A2", cnn) Then Exit Function
        If rsTmp.RecordCount > 0 Then Picture1.Visible = True
        SendSQL strSQL, True
        blnFlag = False
        While Not rsTmp.EOF
            cnn.Execute Replace("Insert Into SO3320A2 (CitemCode,CitemName,TotalAmt,Tax,RcdNo ) Values (" & _
                            GetNullString(rsTmp(0)) & "," & _
                            GetNullString(rsTmp(1)) & "," & GetNullString(rsTmp(2)) & "," & _
                            GetNullString(rsTmp(3)) & "," & GetNullString(rsTmp(4)) & ")", Chr(0), "")
            pgbBar.Value = (100 / SimCount) * rsTmp.AbsolutePosition / rsTmp.RecordCount
            rsTmp.MoveNext
            DoEvents
        Wend
        lngAffected = rsTmp.RecordCount
        Set rsTmp = Nothing
        subInsertMDB21 = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "subInsertMDB21")
End Function

Private Function subInsertMDB22() As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        If optNoClose Or optClose_All Then
            If Not GetStrSQL22(strSQL, "SO033") Then Exit Function
        End If
        If optClose Or optClose_All Then
            If Not GetStrSQL22(strSQL, "SO034") Then Exit Function
            If Not GetStrSQL22(strSQL, "SO035") Then Exit Function
        End If
        strSQL = "Select ServCode,ServName,Sum(RealAmt) TotalAmt, " & _
        "Count(Distinct BillNo) BillCount,Count(Distinct CustId) CustCount,Count(*) IntCount From (" & _
        strSQL & ") A Group By ServCode,ServName"
        
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, "So3320A22", cnn) Then Exit Function
        SendSQL strSQL
        While Not rsTmp.EOF
            cnn.Execute Replace("Insert Into SO3320A22 (ServCode,ServName,TotalAmt,BillCount,CustCount,IntCount ) Values (" & _
                            GetNullString(rsTmp(0)) & "," & GetNullString(rsTmp(1)) & "," & _
                            GetNullString(rsTmp(2)) & "," & GetNullString(rsTmp(3)) & "," & _
                            GetNullString(rsTmp(4)) & "," & GetNullString(rsTmp(5)) & ")", Chr(0), "")
            pgbBar.Value = (100 / SimCount) + (100 / SimCount) * rsTmp.AbsolutePosition / rsTmp.RecordCount
            rsTmp.MoveNext
            DoEvents
        Wend
        Set rsTmp = Nothing
        subInsertMDB22 = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "subInsertMDB22")
End Function

Private Function subInsertMDB23() As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        On Error GoTo chkErr
        If optNoClose Or optClose_All Then
            If Not GetStrSQL23(strSQL, "SO033") Then Exit Function
        End If
        If optClose Or optClose_All Then
            If Not GetStrSQL23(strSQL, "SO034") Then Exit Function
            If Not GetStrSQL23(strSQL, "SO035") Then Exit Function
        End If
        
        strSQL = "Select CitemCode,CitemName,Nvl(RealAmt,0) RealAmt,Nvl(RealPeriod,0) RealPeriod,Sum(RealAmt) TotalAmt,Count(*) IntCount,Sum(NoTaxAmount) NoTaxAmount " & _
        " From (" & strSQL & ") A Group By CitemCode,CitemName,RealAmt,RealPeriod "
        
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, "SO3320A23", cnn) Then Exit Function
        SendSQL strSQL
        Do While Not rsTmp.EOF
            cnn.Execute GetTmpMDBExecuteStr(rsTmp, "SO3320A23")
            pgbBar.Value = (100 / SimCount) * 2 + (100 / SimCount) * rsTmp.AbsolutePosition / rsTmp.RecordCount
            rsTmp.MoveNext
            DoEvents
        Loop
        Set rsTmp = Nothing
        subInsertMDB23 = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "subInsertMDB23")
End Function

Private Function subInsertMDB24() As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        On Error GoTo chkErr
        If optNoClose Or optClose_All Then
            If Not GetStrSQL24(strSQL, "SO033") Then Exit Function
        End If
        If optClose Or optClose_All Then
            If Not GetStrSQL24(strSQL, "SO034") Then Exit Function
            If Not GetStrSQL24(strSQL, "SO035") Then Exit Function
        End If
        
        strSQL = "Select AreaCode,AreaName,Sum(RealAmt) TotalAmt, " & _
        "Count(Distinct BillNo) BillCount,Count(Distinct CustId) CustCount,Count(*) IntCount From (" & _
        strSQL & ") A Group By AreaCode,AreaName"
        
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, "SO3320A24", cnn) Then Exit Function
        SendSQL strSQL
        Do While Not rsTmp.EOF
            cnn.Execute GetTmpMDBExecuteStr(rsTmp, "SO3320A24")
            pgbBar.Value = (100 / SimCount) * 3 + (100 / SimCount) * rsTmp.AbsolutePosition / rsTmp.RecordCount
            rsTmp.MoveNext
            DoEvents
        Loop
        Set rsTmp = Nothing
        subInsertMDB24 = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "subInsertMDB24")
End Function

Private Function subInsertMDB25() As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        On Error GoTo chkErr
        If optNoClose Or optClose_All Then
            If Not GetStrSQL25(strSQL, "SO033") Then Exit Function
        End If
        If optClose Or optClose_All Then
            If Not GetStrSQL25(strSQL, "SO034") Then Exit Function
            If Not GetStrSQL25(strSQL, "SO035") Then Exit Function
        End If
        
        strSQL = "Select CMCode,CMName,Sum(RealAmt) TotalAmt, " & _
        "Count(Distinct BillNo) BillCount,Count(Distinct CustId) CustCount,Count(*) IntCount From (" & _
        strSQL & ") A Group By CMCode,CMName Union All Select 999 CMCode,'合計' CMName,Sum(RealAmt) TotalAmt, " & _
        "Count(Distinct BillNo) BillCount,Count(Distinct CustId) CustCount,Count(*) IntCount From (" & _
        strSQL & ") A "
        
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, "SO3320A25", cnn) Then Exit Function
        SendSQL strSQL, True
        Do While Not rsTmp.EOF
            cnn.Execute GetTmpMDBExecuteStr(rsTmp, "SO3320A25")
            pgbBar.Value = (100 / SimCount) * 4 + (100 / SimCount) * rsTmp.AbsolutePosition / rsTmp.RecordCount
            rsTmp.MoveNext
            DoEvents
        Loop
        Set rsTmp = Nothing
        subInsertMDB25 = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "subInsertMDB25")
End Function

Private Function subInsertMDB26() As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        On Error GoTo chkErr
        If optNoClose Or optClose_All Then
            If Not GetStrSQL26(strSQL, "SO033") Then Exit Function
        End If
        If optClose Or optClose_All Then
            If Not GetStrSQL26(strSQL, "SO034") Then Exit Function
            If Not GetStrSQL26(strSQL, "SO035") Then Exit Function
        End If
        
        strSQL = "Select MduId,MduName,Sum(RealAmt) TotalAmt, " & _
        "Count(Distinct BillNo) BillCount,Count(Distinct CustId) CustCount,Count(*) IntCount From (" & _
        strSQL & ") A Group By MduId,MduName "
        
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, "SO3320A26", cnn) Then Exit Function
        SendSQL strSQL, True
        Do While Not rsTmp.EOF
            cnn.Execute GetTmpMDBExecuteStr(rsTmp, "SO3320A26")
            pgbBar.Value = (100 / SimCount) * 5 + (100 / SimCount) * rsTmp.AbsolutePosition / rsTmp.RecordCount
            rsTmp.MoveNext
            DoEvents
        Loop
        Set rsTmp = Nothing
        subInsertMDB26 = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "subInsertMDB26")
End Function

Private Function subInsertMDB27() As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        On Error GoTo chkErr
        If optNoClose Or optClose_All Then
            If Not GetStrSQL27(strSQL, "SO033") Then Exit Function
        End If
        If optClose Or optClose_All Then
            If Not GetStrSQL27(strSQL, "SO034") Then Exit Function
            If Not GetStrSQL27(strSQL, "SO035") Then Exit Function
        End If
        
        strSQL = "Select OldClctEn,OldClctName,Sum(RealAmt) TotalAmt, " & _
        "Count(Distinct BillNo) BillCount,Count(Distinct CustId) CustCount,Count(*) IntCount From (" & _
        strSQL & ") A Group By OldClctEn,OldClctName "
        
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, "SO3320A27", cnn) Then Exit Function
        SendSQL strSQL, True
        Do While Not rsTmp.EOF
            cnn.Execute GetTmpMDBExecuteStr(rsTmp, "SO3320A27")
            pgbBar.Value = (100 / SimCount) * 6 + (100 / SimCount) * rsTmp.AbsolutePosition / rsTmp.RecordCount
            rsTmp.MoveNext
            DoEvents
        Loop
        Set rsTmp = Nothing
        subInsertMDB27 = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "subInsertMDB27")
End Function

Private Function subInsertMDB28() As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        On Error GoTo chkErr
        If optNoClose Or optClose_All Then
            If Not GetStrSQL28(strSQL, "SO033") Then Exit Function
        End If
        If optClose Or optClose_All Then
            If Not GetStrSQL28(strSQL, "SO034") Then Exit Function
            If Not GetStrSQL28(strSQL, "SO035") Then Exit Function
        End If
        
        strSQL = "Select ClctEn,ClctName,Sum(RealAmt) TotalAmt, " & _
        "Count(Distinct BillNo) BillCount,Count(Distinct CustId) CustCount,Count(*) IntCount From (" & _
        strSQL & ") A Group By ClctEn,ClctName "
        
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, "SO3320A28", cnn) Then Exit Function
        SendSQL strSQL, True
        Do While Not rsTmp.EOF
            cnn.Execute GetTmpMDBExecuteStr(rsTmp, "SO3320A28")
            pgbBar.Value = (100 / SimCount) * 7 + (100 / SimCount) * rsTmp.AbsolutePosition / rsTmp.RecordCount
            rsTmp.MoveNext
            DoEvents
        Loop
        Set rsTmp = Nothing
        subInsertMDB28 = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "subInsertMDB28")
End Function

'95/05/15 Jacky 2412 Karen
Private Function subInsertMDB29() As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        On Error GoTo chkErr
        If optNoClose Or optClose_All Then
            If Not GetStrSQL22(strSQL, "SO033") Then Exit Function
        End If
        If optClose Or optClose_All Then
            If Not GetStrSQL22(strSQL, "SO034") Then Exit Function
            If Not GetStrSQL22(strSQL, "SO035") Then Exit Function
        End If
        
        strSQL = "Select SubStr(BillNO,7,1) BillTypeCode,Sum(RealAmt) TotalAmt, " & _
        "Count(Distinct BillNo) BillCount,Count(Distinct CustId) CustCount,Count(*) IntCount From (" & _
        strSQL & ") A Group By SubStr(BillNO,7,1)"
        
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, "SO3320A29", cnn) Then Exit Function
        SendSQL strSQL, True
        Do While Not rsTmp.EOF
            cnn.Execute GetTmpMDBExecuteStr(rsTmp, "SO3320A29")
            pgbBar.Value = (100 / SimCount) * 8 + (100 / SimCount) * rsTmp.AbsolutePosition / rsTmp.RecordCount
            rsTmp.MoveNext
            DoEvents
        Loop
        Set rsTmp = Nothing
        subInsertMDB29 = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "subInsertMDB29")
End Function

Private Function subInsertMDB2A() As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        On Error GoTo chkErr
        If optNoClose Or optClose_All Then
            If Not GetStrSQL2A(strSQL, "SO033") Then Exit Function
        End If
        If optClose Or optClose_All Then
            If Not GetStrSQL2A(strSQL, "SO034") Then Exit Function
            If Not GetStrSQL2A(strSQL, "SO035") Then Exit Function
        End If
        
        strSQL = "Select PTCode,PTName,Sum(RealAmt) TotalAmt, " & _
        "Count(Distinct BillNo) BillCount,Count(Distinct CustId) CustCount,Count(*) IntCount From (" & _
        strSQL & ") A Group By PTCode,PTName "
        
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, "SO3320A2A", cnn) Then Exit Function
        SendSQL strSQL, True
        Do While Not rsTmp.EOF
            cnn.Execute GetTmpMDBExecuteStr(rsTmp, "SO3320A2A")
            pgbBar.Value = (100 / SimCount) * 9 + (100 / SimCount) * rsTmp.AbsolutePosition / rsTmp.RecordCount
            rsTmp.MoveNext
            DoEvents
        Loop
        Set rsTmp = Nothing
        subInsertMDB2A = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "subInsertMDB2A")
End Function

Private Function GetStrSQL1(ByRef strSQL As String, ByVal StrTableName As String) As Boolean
    On Error GoTo chkErr
    Dim strUseIndexStr As String
    Dim strWhere As String
    Dim strField As String
    Dim strSQL2 As String, strTableName2 As String
    Dim strAlias As String
    Dim strCustName As String
    Dim strTel1 As String
    
        If Len(Trim(strSQL)) > 0 Then
            strSQL = strSQL & " Union All "
        End If
        '沒用了
        If gdaRealDate1.GetValue <> "" Or gdaRealDate1.GetValue <> "" Then
            strUseIndexStr = GetUseIndStr(StrTableName, "RealDate")
        End If
        '沒用了
        If InStr(1, UCase(strChoose), "CUSTID") > 0 Then
            strUseIndexStr = GetUseIndStr(StrTableName, "CustId")
        End If
        '#5267 原本取SO002A,現在改用SO106 By Kin 2010/04/21
        If strChooseSO002A <> "" Then
            strWhere = "And SO033.CustId=SO106.CustId And SO033.AccountNo=SO106.AccountID And " & strChooseSO002A
            strTableName2 = "," & GetOwner & "SO106"
        End If
        '#2810 如果該設備有申請人姓名,則要Show出SO004裡的申請人姓名與電話 By Kin 2007/05/14
        strCustName = "Decode(SO004.DeclarantName,Null,SO001.CustName,SO004.DeclarantName) CustName"
        strTel1 = "Decode(SO004.DeclarantName,Null,SO001.TEL1,SO004.ContTel) TEL1"
        
        '檢查地址依據
        Select Case cboAddress.ListIndex
            Case 0
                strWhere = strWhere & " And So033.AddrNo=SO014.AddrNo And SO033.CompCode=SO014.CompCode "
            Case 1
                strWhere = strWhere & " And So001.InstAddrNo=SO014.AddrNo And SO001.CompCode=SO014.CompCode "
            Case 2
                strWhere = strWhere & " And So001.ChargeAddrNo=SO014.AddrNo And SO001.CompCode=SO014.CompCode "
            Case Else
                strWhere = strWhere & " And So033.AddrNo=SO014.AddrNo And SO033.CompCode=SO014.CompCode "
        End Select
        '#2810 如果該設備有申請人姓名,則要Show出SO004裡的申請人姓名與電話,所以要改變語法 By Kin 2007/05/14
        '#4289 要增加BpCode、BpName By Kin 2008/12/18
        '#5102 增加會計科目(CM001)條件,所以要多串CM001與CD019 By Kin 2009/06/10
        '#5267 原本取SO002A.CardName,改用SO106,所以加一個Distinct
        '#5683 增加繳付類別 By Kin 2010/08/17
        strSQL2 = "Select Distinct" & strUseIndexStr & "  SO033.CustId,SO033.BillNo," & _
                 strCustName & "," & strTel1 & "," & _
                 "SO033.CitemName,SO033.RealDate,Nvl(SO033.RealPeriod,0) RealPeriod," & _
                 "Nvl(SO033.RealAmt,0) RealAmt,SO033.ClctName,SO014.Address,SO033.RealStartDate," & _
                 "SO033.RealStopDate,SO033.GUINo,SO014.MduName,CD002.Description ServName," & _
                 "Cd001.Description AreaName,SO033.CMCode,SO033.CMName,SO033.ClctEn, So033.CitemCode,0 as NoTaxAmt, " & _
                 "So002.CustStatusName,SO003.ClctDate,So014.AddrSort,So002.WipName1,So002.WipCode1," & _
                 "So033.OldClctName,So033.Note,So033.OldClctEn,So033.AreaCode,So033.ServCode,SO033.ServiceType,SO002.InvNo InvNo" & _
                 ",SO033.ClctAreaCode,SO033.UpdTime,to_number(null) as MduCount,SO033.STName,SO014.MduId,to_char(null) as ClctEnWorkClass," & _
                 "SO033.AccountNo,SO033.SeqNo, " & _
                 IIf(strChooseSO002A <> "", "SO106.CardName", "to_char(null)") & " as CardType," & _
                 "SO033.ManualNo,SO033.FaciSeqNo,CD040.Description ClctAreaName,SO033.DiscountCode," & _
                 "SO033.DiscountName,SO033.DiscountDate1,SO033.DiscountDate2," & _
                 "SO033.MediaBillNo,SO001.Tel3,SO033.ClassCode,CD004.Description ClassName," & _
                 "SO033.PTCode,SO033.PTName,SO033.BpCode,SO033.BPName,Nvl(SO033.PayKind,0) PayKind From " & _
                 GetOwner & StrTableName & " So033," & GetOwner & "So001," & GetOwner & "So002," & GetOwner & "So014," & GetOwner & "Cd002," & GetOwner & "Cd001," & GetOwner & "CD040 CD040," & GetOwner & "SO003 SO003," & GetOwner & "CD004 CD004 " & strTableName2 & "," & GetOwner & "SO004 SO004 " & _
                 "," & GetOwner & "CM001," & GetOwner & "CD019 " & _
                 " Where So033.CustId=SO001.CustID And SO033.CompCode=SO001.CompCode And SO033.CustId=SO002.CustId And SO033.CompCode=SO002.CompCode And SO033.ServiceType=SO002.ServiceType " & _
                 " And SO033.CustId = SO003.CustId(+) And SO033.CompCode = SO003.CompCode(+) And SO033.ServiceType = SO003.ServiceType(+) And SO033.CitemCode = SO003.CitemCode(+) And SO033.FaciSeqNo = SO003.FaciSeqNo(+) " & strWhere & " And So033.ServCode=Cd002.CodeNo And So033.AreaCode=Cd001.CodeNo And So033.ClassCode=Cd004.CodeNo And SO033.ClctAreaCode = CD040.CodeNo(+) And " & _
                 " SO033.FaciSeqNO=SO004.SEQNO(+) And SO033.CustId=SO004.CustId(+) And " & strChoose & _
                 IIf(strChooseSO002A <> "", " AND SO106.StopFlag<>1 And SnactionDate Is Not Null ", "")
        
'        strSQL2 = "Select " & strUseIndexStr & "  SO033.CustId,SO033.BillNo," & _
'                 "SO001.CustName,SO001.TEL1," & _
'                 "SO033.CitemName,SO033.RealDate,Nvl(SO033.RealPeriod,0) RealPeriod," & _
'                 "Nvl(SO033.RealAmt,0) RealAmt,SO033.ClctName,SO014.Address,SO033.RealStartDate," & _
'                 "SO033.RealStopDate,SO033.GUINo,SO014.MduName,CD002.Description ServName," & _
'                 "Cd001.Description AreaName,SO033.CMCode,SO033.CMName,SO033.ClctEn, So033.CitemCode,0 as NoTaxAmt, " & _
'                 "So002.CustStatusName,SO003.ClctDate,So014.AddrSort,So002.WipName1,So002.WipCode1," & _
'                 "So033.OldClctName,So033.Note,So033.OldClctEn,So033.AreaCode,So033.ServCode,SO033.ServiceType,SO002.InvNo InvNo" & _
'                 ",SO033.ClctAreaCode,SO033.UpdTime,to_number(null) as MduCount,SO033.STName,SO014.MduId,to_char(null) as ClctEnWorkClass,SO033.AccountNo,SO033.SeqNo, " & _
'                 IIf(strChooseSO002A <> "", "SO002A.CardName", "to_char(null)") & " as CardType,SO033.ManualNo,SO033.FaciSeqNo,CD040.Description ClctAreaName,SO033.DiscountCode,SO033.DiscountName,SO033.DiscountDate1,SO033.DiscountDate2,SO033.MediaBillNo,SO001.Tel3,SO033.ClassCode,CD004.Description ClassName,SO033.PTCode,SO033.PTName From " & _
'                 GetOwner & StrTableName & " So033," & GetOwner & "So001," & GetOwner & "So002," & GetOwner & "So014," & GetOwner & "Cd002," & GetOwner & "Cd001," & GetOwner & "CD040 CD040," & GetOwner & "SO003 SO003," & GetOwner & "CD004 CD004 " & strTableName2 & "," & GetOwner & "SO004 SO004 " & _
'                 " Where So033.CustId=SO001.CustID And SO033.CompCode=SO001.CompCode And SO033.CustId=SO002.CustId And SO033.CompCode=SO002.CompCode And SO033.ServiceType=SO002.ServiceType " & _
'                 " And SO033.CustId = SO003.CustId(+) And SO033.CompCode = SO003.CompCode(+) And SO033.ServiceType = SO003.ServiceType(+) And SO033.CitemCode = SO003.CitemCode(+) And SO033.FaciSeqNo = SO003.FaciSeqNo(+) " & strWhere & " And So033.ServCode=Cd002.CodeNo And So033.AreaCode=Cd001.CodeNo And So033.ClassCode=Cd004.CodeNo And SO033.ClctAreaCode = CD040.CodeNo(+) And " & _
'                 " SO033.FaciSeqNO=SO004.SEQNO(+) And " & strChoose
        strAlias = "A" & Right(StrTableName, 1)
'        If cboGroup.ListIndex = 3 Then
'           strSQL2 = "Select " & strAlias & ".* ,B.Description ClctAreaName From (" & strSQL2 & ") " & strAlias & "," & GetOwner & "CD040 B Where A.ClctAreaCode = B.CodeNo(+)"
'        Else
'            strSQL2 = "Select " & strAlias & ".* ,To_char(Null) as ClctAreaName From (" & strSQL2 & ") " & strAlias
'        End If
        strSQL = strSQL & strSQL2
        GetStrSQL1 = True
    Exit Function
chkErr:
    ErrSub Me.Name, "GetStrSQL1"
End Function
'#5102 增加會計科目小計 By Kin 2009/06/10
Private Function GetStrSQL2C(ByRef strSQL As String, StrTableName As String) As Boolean
  On Error GoTo chkErr
    Dim strAllTable As String
    Dim strAllChoose As String
    
    If Len(Trim(strSQL)) > 0 Then
        strSQL = strSQL & " Union All "
    End If
    Call GetTableAndWhere(strAllTable, strAllChoose)
     
'    strSQL = strSQL & "Select CM001.CodeNo,CM001.Description ,SO033.CustId," & _
'                    "SO033.RealAmt,SO033.BillNo " & _
'                    "From " & GetOwner & StrTableName & " SO033 " & strAllTable & ", " & GetOwner & "CD019 CD019, " & GetOwner & "CD033 CD033 " & _
'                    "," & GetOwner & "CM001 CM001" & _
'                    " WHERE  SO033.CitemCode=CD019.CodeNo " & strAllChoose & _
'                    " And CD019.TaxCode=CD033.CodeNo(+) And " & strChoose & _
'                    " And CD019.AccIdP1=CM001.CodeNo(+)"

    strSQL = strSQL & "Select CM001.CodeNo,CM001.Description ,SO033.CustId," & _
                    "SO033.RealAmt,SO033.BillNo " & _
                    "From " & GetOwner & StrTableName & " SO033 " & strAllTable & _
                    " WHERE  1=1  " & strAllChoose & _
                    " And " & strChoose
                    

    GetStrSQL2C = True
    Exit Function
chkErr:
  ErrSub Me.Name, "GetStrSQL2C"
End Function
Private Function GetStrSQL2(ByRef strSQL As String, StrTableName As String) As Boolean
    On Error GoTo chkErr
    Dim strAllTable As String
    Dim strAllChoose As String
    Dim strUseIndexStr As String
        If Len(Trim(strSQL)) > 0 Then
            strSQL = strSQL & " Union All "
        End If
        
        strUseIndexStr = GetUseIndStr(StrTableName, "RealDate")
        Call GetTableAndWhere(strAllTable, strAllChoose)
        '#5102 所有的語法都強迫串上CM001與CD019了,所以將原本語法的CD019拿掉 By Kin 2009/06/10
'        strSQL = strSQL & "Select " & strUseIndexStr & " SO033.CitemCode,SO033.CitemName,SO033.CustId," & _
'                    "SO033.RealAmt,CD033.Rate1 " & _
'                    "From " & GetOwner & StrTableName & " SO033 " & strAllTable & ", " & GetOwner & "CD019 CD019, " & GetOwner & "CD033 CD033 " & _
'                    " WHERE  SO033.CitemCode=CD019.CodeNo " & strAllChoose & _
'                    " And CD019.TaxCode=CD033.CodeNo(+) And " & strChoose
                    
                    
        strSQL = strSQL & "Select " & strUseIndexStr & " SO033.CitemCode,SO033.CitemName,SO033.CustId," & _
                    "SO033.RealAmt,CD033.Rate1 " & _
                    "From " & GetOwner & StrTableName & " SO033 " & strAllTable & ", " & GetOwner & "CD033 CD033 " & _
                    " WHERE  1=1 " & strAllChoose & _
                    " And CD019.TaxCode=CD033.CodeNo(+) And " & strChoose

        GetStrSQL2 = True
    Exit Function
chkErr:
    ErrSub Me.Name, "GetStrSQL2"
End Function
'#4015 簡表增加增加收費付款小計 By Kin 2008/07/18
Private Function GetStrSQL2B(ByRef strSQL As String, StrTableName As String) As Boolean
  On Error GoTo chkErr
    Dim strAllTable As String
    Dim strAllChoose As String
    Dim strUseIndexStr As String
    If Len(Trim(strSQL)) > 0 Then
        strSQL = strSQL & " Union All "
    End If
    Call GetTableAndWhere(strAllTable, strAllChoose)
    strSQL = strSQL & "Select " & "SO033.CitemCode,SO033.CitemName,SO033.PTCode,SO033.PTName,SO033.CustId,SO033.BillNo,SO033.RealAmt " & _
                " From " & GetOwner & StrTableName & " SO033 " & strAllTable & _
                " Where 1 = 1 " & strAllChoose & _
                    " And " & strChoose
    GetStrSQL2B = True
    Exit Function
chkErr:
  ErrSub Me.Name, "GetStrSQL2B"
End Function
Private Function GetStrSQL22(ByRef strSQL As String, StrTableName As String) As Boolean
    On Error GoTo chkErr
    Dim strAllTable As String
    Dim strAllChoose As String
    Dim strUseIndexStr As String
        If Len(Trim(strSQL)) > 0 Then
            strSQL = strSQL & " Union All "
        End If
        strUseIndexStr = GetUseIndStr(StrTableName, "RealDate")
        Call GetTableAndWhere(strAllTable, strAllChoose)
        '#5102 所有的語法都強迫串上CM001與CD019了,所以將原本語法的CD019拿掉 By Kin 2009/06/10
'        strSQL = strSQL & "Select " & strUseIndexStr & " SO033.ServCode,CD002.Description ServName," & _
'                    "SO033.RealAmt,So033.CustId,So033.BillNo " & _
'                    "From " & GetOwner & StrTableName & " SO033 " & strAllTable & ", " & GetOwner & "CD002 CD002 " & _
'                    " WHERE  SO033.ServCode=CD002.CodeNo " & strAllChoose & _
'                    " And " & strChoose

        strSQL = strSQL & "Select " & strUseIndexStr & " SO033.ServCode,CD002.Description ServName," & _
                    "SO033.RealAmt,So033.CustId,So033.BillNo " & _
                    "From " & GetOwner & StrTableName & " SO033 " & strAllTable & ", " & GetOwner & "CD002 CD002 " & _
                    " WHERE  SO033.ServCode=CD002.CodeNo " & strAllChoose & _
                    " And " & strChoose

        GetStrSQL22 = True
    Exit Function
chkErr:
    ErrSub Me.Name, "GetStrSQL22"
End Function

Private Function GetStrSQL23(ByRef strSQL As String, StrTableName As String) As Boolean
    On Error GoTo chkErr
    Dim strAllTable As String
    Dim strAllChoose As String
    Dim strUseIndexStr As String
        If Len(Trim(strSQL)) > 0 Then
            strSQL = strSQL & " Union All "
        End If
        strUseIndexStr = GetUseIndStr(StrTableName, "RealDate")
        Call GetTableAndWhere(strAllTable, strAllChoose)
        '#5102 所有的語法都強迫串上CM001與CD019了,所以將原本語法的CD019拿掉 By Kin 2009/06/10
'        strSQL = strSQL & "Select " & strUseIndexStr & " SO033.CitemCode,So033.CitemName," & _
'                    "SO033.RealAmt,So033.RealPeriod,So033.CustId,So033.BillNo,Round((SO033.RealAmt/(1+ Nvl(CD033.Rate1,0)/100))) NoTaxAmount " & _
'                    "From " & GetOwner & StrTableName & " SO033 " & strAllTable & ", " & GetOwner & "CD002 CD002," & GetOwner & "CD019," & GetOwner & "CD033 CD033 " & _
'                    " WHERE  SO033.ServCode=CD002.CodeNo And SO033.CitemCode=CD019.CodeNo And CD019.TaxCode=CD033.CodeNo(+) " & strAllChoose & _
'                    " And " & strChoose

        strSQL = strSQL & "Select " & strUseIndexStr & " SO033.CitemCode,So033.CitemName," & _
                    "SO033.RealAmt,So033.RealPeriod,So033.CustId,So033.BillNo," & _
                    " Round((SO033.RealAmt/(1+ Nvl(CD033.Rate1,0)/100))) NoTaxAmount " & _
                    "From " & GetOwner & StrTableName & " SO033 " & strAllTable & ", " & _
                    GetOwner & "CD002 CD002," & GetOwner & "CD033 CD033 " & _
                    " WHERE  SO033.ServCode=CD002.CodeNo And SO033.CitemCode=CD019.CodeNo And CD019.TaxCode=CD033.CodeNo(+) " & strAllChoose & _
                    " And " & strChoose

        GetStrSQL23 = True
    Exit Function
chkErr:
    ErrSub Me.Name, "GetStrSQL23"
End Function

Private Function GetStrSQL24(ByRef strSQL As String, StrTableName As String) As Boolean
    On Error GoTo chkErr
    Dim strAllTable As String
    Dim strAllChoose As String
    Dim strUseIndexStr As String
        If Len(Trim(strSQL)) > 0 Then
            strSQL = strSQL & " Union All "
        End If
        strUseIndexStr = GetUseIndStr(StrTableName, "RealDate")
        Call GetTableAndWhere(strAllTable, strAllChoose)
        '#5102 所有的語法都強迫串上CM001與CD019了,所以將原本語法的CD019拿掉 By Kin 2009/06/10
'        strSQL = strSQL & "Select " & strUseIndexStr & " SO033.AreaCode,CD001.Description AreaName," & _
'                    "SO033.RealAmt,So033.CustId,So033.BillNo " & _
'                    "From " & GetOwner & StrTableName & " SO033 " & strAllTable & ", " & GetOwner & "CD001 CD001 " & _
'                    " WHERE  SO033.AreaCode=CD001.CodeNo " & strAllChoose & _
'                    " And " & strChoose

        strSQL = strSQL & "Select " & strUseIndexStr & " SO033.AreaCode,CD001.Description AreaName," & _
                    "SO033.RealAmt,So033.CustId,So033.BillNo " & _
                    "From " & GetOwner & StrTableName & " SO033 " & strAllTable & ", " & _
                    GetOwner & "CD001 CD001 " & _
                    " WHERE  SO033.AreaCode=CD001.CodeNo " & strAllChoose & _
                    " And " & strChoose

        GetStrSQL24 = True
    Exit Function
chkErr:
    ErrSub Me.Name, "GetStrSQL24"
End Function

Private Function GetStrSQL25(ByRef strSQL As String, StrTableName As String) As Boolean
    On Error GoTo chkErr
    Dim strAllTable As String
    Dim strAllChoose As String
    Dim strUseIndexStr As String
        If Len(Trim(strSQL)) > 0 Then
            strSQL = strSQL & " Union All "
        End If
        strUseIndexStr = GetUseIndStr(StrTableName, "RealDate")
        Call GetTableAndWhere(strAllTable, strAllChoose)
        
        strSQL = strSQL & "Select " & strUseIndexStr & " SO033.CMCode,SO033.CMName," & _
                    "SO033.RealAmt,So033.CustId,So033.BillNo " & _
                    "From " & GetOwner & StrTableName & " SO033 " & strAllTable & _
                    " Where 1 = 1 " & strAllChoose & _
                    " And " & strChoose
        GetStrSQL25 = True
    Exit Function
chkErr:
    ErrSub Me.Name, "GetStrSQL25"
End Function

Private Function GetStrSQL26(ByRef strSQL As String, StrTableName As String) As Boolean
    On Error GoTo chkErr
    Dim strAllTable As String
    Dim strAllChoose As String
    Dim strUseIndexStr As String
        If Len(Trim(strSQL)) > 0 Then
            strSQL = strSQL & " Union All "
        End If
        strUseIndexStr = GetUseIndStr(StrTableName, "RealDate")
        Call GetTableAndWhere(strAllTable, strAllChoose)
        
        strSQL = strSQL & "Select " & strUseIndexStr & " SO033.MduId,SO017.Name MduName," & _
                    "SO033.RealAmt,So033.CustId,So033.BillNo " & _
                    "From " & GetOwner & StrTableName & " SO033," & GetOwner & "SO017 " & strAllTable & " Where SO033.MduId = SO017.MduId And SO033.MduId is not null " & strAllChoose & _
                    " And " & strChoose
        GetStrSQL26 = True
    Exit Function
chkErr:
    ErrSub Me.Name, "GetStrSQL26"
End Function

Private Function GetStrSQL27(ByRef strSQL As String, StrTableName As String) As Boolean
    On Error GoTo chkErr
    Dim strAllTable As String
    Dim strAllChoose As String
    Dim strUseIndexStr As String
        If Len(Trim(strSQL)) > 0 Then
            strSQL = strSQL & " Union All "
        End If
        strUseIndexStr = GetUseIndStr(StrTableName, "RealDate")
        Call GetTableAndWhere(strAllTable, strAllChoose)
        
        strSQL = strSQL & "Select " & strUseIndexStr & " SO033.OldClctEn,SO033.OldClctName," & _
                    "SO033.RealAmt,So033.CustId,So033.BillNo " & _
                    "From " & GetOwner & StrTableName & " SO033 " & strAllTable & " Where 1 = 1 " & strAllChoose & _
                    " And " & strChoose
        GetStrSQL27 = True
    Exit Function
chkErr:
    ErrSub Me.Name, "GetStrSQL27"
End Function

Private Function GetStrSQL28(ByRef strSQL As String, StrTableName As String) As Boolean
    On Error GoTo chkErr
    Dim strAllTable As String
    Dim strAllChoose As String
    Dim strUseIndexStr As String
        If Len(Trim(strSQL)) > 0 Then
            strSQL = strSQL & " Union All "
        End If
        strUseIndexStr = GetUseIndStr(StrTableName, "RealDate")
        Call GetTableAndWhere(strAllTable, strAllChoose)
        
        strSQL = strSQL & "Select " & strUseIndexStr & " SO033.ClctEn,SO033.ClctName," & _
                    "SO033.RealAmt,So033.CustId,So033.BillNo " & _
                    "From " & GetOwner & StrTableName & " SO033 " & strAllTable & " Where 1 = 1 " & strAllChoose & _
                    " And " & strChoose
        GetStrSQL28 = True
    Exit Function
chkErr:
    ErrSub Me.Name, "GetStrSQL28"
End Function

Private Function GetStrSQL2A(ByRef strSQL As String, StrTableName As String) As Boolean
    On Error GoTo chkErr
    Dim strAllTable As String
    Dim strAllChoose As String
    Dim strUseIndexStr As String
        If Len(Trim(strSQL)) > 0 Then
            strSQL = strSQL & " Union All "
        End If
        strUseIndexStr = GetUseIndStr(StrTableName, "RealDate")
        Call GetTableAndWhere(strAllTable, strAllChoose)
        
        strSQL = strSQL & "Select " & strUseIndexStr & " SO033.PTCode,SO033.PTName," & _
                    "SO033.RealAmt,So033.CustId,So033.BillNo " & _
                    "From " & GetOwner & StrTableName & " SO033 " & strAllTable & " Where 1 = 1 " & strAllChoose & _
                    " And " & strChoose
        GetStrSQL2A = True
    Exit Function
chkErr:
    ErrSub Me.Name, "GetStrSQL2A"
End Function

Private Function GetTableAndWhere(ByRef strTable As String, _
        ByRef strWhere As String) As Boolean
    On Error GoTo chkErr
    '#2810 測試時發現strTable少串了之前的Table,會導致SQL錯誤現在加上去strTable=strTable & Where條件 By Kin 2007/05/14
        If InStr(UCase(strChoose), "SO001") Then
            strTable = strTable & "," & GetOwner & "SO001 SO001"
            strWhere = " And SO033.CustId=SO001.CustID"
        End If
        If InStr(UCase(strChoose), "SO014") Then
            strTable = strTable & "," & GetOwner & "SO014 SO014"
            strWhere = " And SO033.AddrNo=SO014.AddrNo And SO033.CompCode=SO014.CompCode "
        End If
        If InStr(UCase(strChoose), "SO002") Then
            strTable = strTable & "," & GetOwner & "SO002 SO002"
            strWhere = " And SO033.CustId=SO002.CustId And SO033.ServiceType=SO002.ServiceType"
        End If
        
        '#2810 增加下收日條件(簡表) By Kin 2007/05/14
        If InStr(UCase(strChoose), "SO003") Then
            strTable = strTable & "," & GetOwner & "SO003 SO003"
            strWhere = strWhere & " And SO003.CustId=SO033.CustId And SO003.CompCode=SO033.CompCode And SO003.ServiceType=SO033.ServiceType " & _
                                  " And SO003.CitemCode=SO033.CitemCode And SO003.FaciSeqNo=SO033.FaciSeqNo"
                       
        End If
        '#5102 增加CM001(會計科目)條件,所以所有的語法要多串CD019與CM001 By Kin 2009/06/10
        strTable = strTable & "," & GetOwner & "CD019 CD019"
        strTable = strTable & "," & GetOwner & "CM001 CM001"
        
'        If InStr(UCase(strChoose), "ACCIDP1") > 0 Then
'            If InStr(UCase(strTable), "CM001") = 0 Then
'                strTable = strTable & "," & GetOwner & "CM001 CM001"
'                strWhere = strWhere & " And CD019.AccIdP1=CM001.CodeNo(+) "
'            End If
'            If InStr(UCase(strTable), "CD019") = 0 And blnCD019 Then
'                strTable = strTable & "," & GetOwner & "CD019 CD019"
'                strWhere = strWhere & " And SO033.CitemCode=CD019.CodeNo "
'            End If
'        End If
    Exit Function
chkErr:
    ErrSub Me.Name, "GetTableName"
End Function

Private Sub SimpleReport()
    On Error GoTo chkErr
    Dim strSQL As String
    Dim strsubsql(11) As String
    Dim intLoop As Integer
        If cnn.State = adStateClosed Then Set cnn = GetTmpMdbCn
'        Picture1.Visible = True
        cnn.BeginTrans
        If Not subInsertMDB21 Then cnn.RollbackTrans: Exit Sub
        If Not subInsertMDB22 Then cnn.RollbackTrans: Exit Sub
        If Not subInsertMDB23 Then cnn.RollbackTrans: Exit Sub
        If Not subInsertMDB24 Then cnn.RollbackTrans: Exit Sub
        If Not subInsertMDB25 Then cnn.RollbackTrans: Exit Sub
        If Not subInsertMDB26 Then cnn.RollbackTrans: Exit Sub
        If Not subInsertMDB27 Then cnn.RollbackTrans: Exit Sub
        If Not subInsertMDB28 Then cnn.RollbackTrans: Exit Sub
        If Not subInsertMDB29 Then cnn.RollbackTrans: Exit Sub
        If Not subInsertMDB2A Then cnn.RollbackTrans: Exit Sub
        '#4015 簡表增加收費付款小計 By Kin 2008/07/18
        If Not subInsertMDB2B Then cnn.RollbackTrans: Exit Sub
        '#5102 增加會計科目小計 By Kin 2009/06/10
        If Not subInsertMDB2C Then cnn.RollbackTrans: Exit Sub
        cnn.CommitTrans
        Picture1.Visible = False
        strsubsql(1) = "SELECT * From SO3320A2"
        For intLoop = 1 To UBound(strsubsql)
            strsubsql(intLoop) = "Select * from So3320A2" & IIf(intLoop >= 10, Chr(65 + intLoop - 10), intLoop + 1)
        Next
        If lngAffected > 0 Then
            PrintRpt2 GetPrinterName(5), RptName("SO3320", 2), "SO3320", "收款報表查詢 (簡表) [SO3320A2]", strsubsql(1), GetQryStr, , True, "Tmp0000.MDB", , , GiPaperPortrait, , strsubsql(1), strsubsql(2), strsubsql(3), strsubsql(4), strsubsql(5), , , , , , , strsubsql(6), strsubsql(7), strsubsql(8), strsubsql(9), strsubsql(10), strsubsql(11)
        Else
            SendSQL , , True
            MsgNoRcd
        End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Sub SimpleReport")
End Sub

Private Sub DetialReport(Optional lngAffected As Long = 0, Optional blnPrintReport As Boolean = True)
     On Error GoTo chkErr
    Dim strSQL As String
    Dim strSQL1 As String
    Dim strSQL2 As String
    Dim strSQL3 As String
    Dim strSQL4 As String
    Dim strField As String
    Dim strSo033 As String
    Dim rsTmp As New ADODB.Recordset
        If cnn.State = adStateClosed Then Set cnn = GetTmpMdbCn
        
        On Error GoTo ErrTrans
        cnn.BeginTrans
        '未日結或全部
        If optNoClose Or optClose_All Then
            If Not GetStrSQL1(strSQL, "SO033") Then Exit Sub
        End If
        '已日結或全部
        If optClose Or optClose_All Then
            If Not GetStrSQL1(strSQL, "SO034") Then Exit Sub
            If Not GetStrSQL1(strSQL, "SO035") Then Exit Sub
        End If
        If Not subInsertMDB1(strSQL) Then Exit Sub
        cnn.CommitTrans
        SendSQL "End", True
        If optSimple Then Exit Sub
        On Error GoTo chkErr
        lngAffected = GetRsValue("Select Count(*) as intCount from SO3320A1", cnn)
        If lngAffected = 0 Then
            SendSQL , , True
            MsgNoRcd
        Else
           strSQL = "Select * From SO3320A1"
           If blnPrintReport Then PrintRpt2 GetPrinterName(5), RptName("SO3320", 1), "SO3320", "收款明細表 [SO3320A1]", strSQL, GetQryStr, , True, "Tmp0000.MDB", , strGroupName, GiPaperLandscape, , strSQL, strSQL, strSQL, strSQL, strSQL, , , , , , , strSQL, strSQL, strSQL, strSQL, strSQL
        End If
        Set rsTmp = Nothing
    Exit Sub
ErrTrans:
    cnn.RollbackTrans
chkErr:
    ErrSub Me.Name, "DetailReport"
End Sub

Private Sub subChoose()
    On Error GoTo chkErr
    Dim strField As String
        strChooseSO002A = ""
        strField = " So033 "
        strChoose = ""
        strMDBOrder = ""
        
        '實收日期
        If gdaRealDate1.GetValue = "" And gdaRealDate2.GetValue = "" Then
            subAnd "RealDate Is Not Null"
        Else
            If gdaRealDate1.GetValue <> "" Then subAnd GetCompareDate("SO033.RealDate ", gdaRealDate1.GetValue, True)
            If gdaRealDate2.GetValue <> "" Then subAnd GetCompareDate("SO033.RealDate ", gdaRealDate2.GetValue, False)
        End If
        
        '印單序號
        If txtBillNo1 <> "" Then subAnd GetBillnoStr(cboBillType.ListIndex, "'" & txtBillNo1 & "'", True)
        If txtBillNo2 <> "" Then subAnd GetBillnoStr(cboBillType.ListIndex, "'" & txtBillNo2 & "'", False)
        
        If gilCompCode.GetCodeNo <> "" Then subAnd gilCompCode.GetCodeNo & " =SO033.CompCode"
        '應收日期
        If gdaShouldDay1.GetValue <> "" Then subAnd GetCompareDate("SO033.ShouldDate ", gdaShouldDay1.GetValue, True)
        If gdaShouldDay2.GetValue <> "" Then subAnd GetCompareDate("SO033.ShouldDate", gdaShouldDay2.GetValue, False)
        '92/06/26
        If gdtCreateTime1.GetValue <> "" Then subAnd "SO033.CreateTime >=" & GetNullString(gdtCreateTime1.GetValue(True), giDateV, , True)
        If gdtCreateTime2.GetValue <> "" Then subAnd "SO033.CreateTime <=" & GetNullString(gdtCreateTime2.GetValue(True), giDateV, , True)
        '#2810 增加下收日 By Kin 2007/05/14
        If gdaClctDate1.GetValue & "" <> "" Then subAnd GetCompareDate("SO003.ClctDate ", gdaClctDate1.GetValue, True)
        If gdaClctDate1.GetValue & "" <> "" Then subAnd GetCompareDate("SO003.ClctDate ", gdaClctDate2.GetValue, False)
        'If gilServiceType.GetCodeNo <> "" Then subAnd "So033.ServiceType = '" & gilServiceType.GetCodeNo & "'"
        '服務類別
        If gimServiceType.GetQryStr <> "" Then subAnd "SO033.ServiceType " & gimServiceType.GetQryStr
        
        
        
        If gimBillType.GetQryStr <> "" Then subAnd "substr(SO033.BillNo,7,1) " & gimBillType.GetQryStr
        '收費方式
        If gimCMCode.GetQryStr <> "" Then subAnd "SO033.CMCode " & gimCMCode.GetQryStr
        If gimCitemCode.GetQryStr <> "" Then subAnd "SO033.CitemCode " & gimCitemCode.GetQryStr
        '#5102 增加會計科目 By Kin 2009/06/10
        If gimCM001.GetQryStr <> "" Then
            subAnd "CD019.AccIdP1 " & gimCM001.GetQryStr
        End If
        '#5683 增加繳付類別 By Kin 2010/08/16
        If gimPayType.GetQryStr <> "" Then
            Call subAnd("SO033.PayKind " & gimPayType.GetQryStr)
        End If
        '客戶類別
        If gimClassCode.GetQryStr <> "" Then subAnd "SO033.ClassCode " & gimClassCode.GetQryStr
        '客戶狀態
        If gimCustStatus.GetQryStr <> "" Then subAnd "SO002.CustStatusCode " & gimCustStatus.GetQryStr
        '服務區
        If gimServCode.GetQryStr <> "" Then subAnd "SO033.ServCode " & gimServCode.GetQryStr
        '行政區
        If gimAreaCode.GetQryStr <> "" Then subAnd "SO033.AreaCode " & gimAreaCode.GetQryStr
        '2006/07/19 Jacky 1596 Kobe
        '收費區
        If gimClctAreaCode.GetQryStr <> "" Then subAnd "SO033.ClctAreaCode " & gimClctAreaCode.GetQryStr
        '付款種類
        If gimPTCode.GetQryStr <> "" Then subAnd "SO033.PTCode " & gimPTCode.GetQryStr
        '大樓編號
        If gimMduId.GetQryStr <> "" Then
            subAnd "SO033.MduId " & gimMduId.GetQryStr
        Else
            If gimMduId.GetDispStr <> "" Then subAnd "So033.MduId Is Not Null"
        End If
        '工作類別
        If gimWorkClass.GetQryStr <> "" Then Call subAnd("SO033.ClctEn in (Select EmpNo From CM003 Where WorkClass " & gimWorkClass.GetQryStr & ")")
        '收費人員
        If gimClctEn.GetQryStr <> "" Then subAnd "SO033.ClctEn " & gimClctEn.GetQryStr
        '2006/07/19 Jacky 1596 Kobe
        '原收費人員
        If gimOldClctEn.GetQryStr <> "" Then subAnd "SO033.OldClctEn " & gimOldClctEn.GetQryStr
        '街道範圍
        If gimStrtCode.GetQryStr <> "" Then subAnd "SO033.StrtCode " & gimStrtCode.GetQryStr
        '短收原因
        If gimSTCode.GetQryStr <> "" Then
            subAnd "SO033.STCode " & gimSTCode.GetQryStr
        Else
            If chkBad = 0 Then subAnd "(SO033.STCode In (Select CodeNo From " & GetOwner & "CD016 Where nvl(RefNo,0) <>  1)  or So033.STCode is Null )"
        End If
        '銀行名稱
        If gimBankCode.GetQryStr <> "" Then subAnd "SO033.BankCode " & gimBankCode.GetQryStr
        '發票種類(預開)
        If optInvoiceForward Then
            Call subAnd("SO002.PreInvoice = 1")
        ElseIf Me.optInvoiceBack Then
            Call subAnd("SO002.PreInvoice = 0")
        End If
        '2002/08/31
        If ginRealPeriod1.Text <> "" Then Call subAnd("SO033.RealPeriod >=" & ginRealPeriod1.Value)
        If ginRealPeriod2.Text <> "" Then Call subAnd("SO033.RealPeriod <=" & ginRealPeriod2.Value)
        '2002/08/31
        If txtRealAmt1.Text <> "" Then Call subAnd("SO033.RealAmt >= " & txtRealAmt1.Text)
        If txtRealAmt2.Text <> "" Then Call subAnd("SO033.RealAmt <= " & txtRealAmt2.Text)
        '2001/06/15
        If gdtUpdTime1.GetValue <> "" Then Call subAnd("SO033.UpdTime >= '" & GetDTString(gdtUpdTime1.GetValue(True)) & "'")
        If gdtUpdTime2.GetValue <> "" Then Call subAnd("SO033.UpdTime <= '" & GetDTString(gdtUpdTime2.GetValue(True)) & "'")
        
        '是否已拋檔
        If chkInvoice.Value = 1 Then Call subAnd("SO033.InvoiceTime is not null")
        '是否已開發票
        If chkGUINo.Value = 1 Then Call subAnd("SO033.GUINo is not null")
        Select Case cboFirstFlag.ListIndex
            Case 0  '只印首期
                Call subAnd("1=Nvl(SO033.FirstFlag,0)")
            Case 1  '不印首期
                Call subAnd("0=Nvl(SO033.FirstFlag,0)")
        End Select
        '2001/08/06
        '網路編號
        If txtCircuitNo.Text = "" Then
            If optCircuitIsNull.Value = True Then Call subAnd("So014.CircuitNo Is Null")
        Else
            Call subAnd("'" & txtCircuitNo.Text & "'=SO014.CircuitNo")
        End If
        '2001/08/06
        '客戶編號
        If txtCustId.Text <> "" Then
            Call TimetxtCustId(txtCustId, "So033")
        End If
        '信用卡種類
        '#5267 原本是串SO002A,現在改SO106 By Kin 2010/04/21
        If gimCardCode.GetQryStr <> "" Then Call subAnd2(strChooseSO002A, "SO106.CardName " & gimCardCode.GetDescStr)
      
        '異動人員
        If gimUpdEn.GetQryStr <> "" Then subAnd "SO033.UpdEn " & gimUpdEn.GetDescStr
        '裝機日期
        If gdaInstDay1.GetValue <> "" Then subAnd GetCompareDate("SO002.InstTime", gdaInstDay1.GetValue, True)
        If gdaInstDay2.GetValue <> "" Then subAnd GetCompareDate("SO002.InstTime", gdaInstDay2.GetValue, False)
        '銷售點
        If gilIntroId.GetCodeNo <> "" Then subAnd "SO002.IntroId = '" & gilIntroId.GetCodeNo & "'"
        '派工類別
        If gimWipCode.GetQueryCode <> "" Then
            Dim varString
            Dim varS
            Dim strWip0  As String, strWip1 As String, strWip2 As String, strWip3 As String
            Dim strWipAll As String
            varString = Split(gimWipCode.GetQueryCode, ",")
            For Each varS In varString
                If varS = 0 Then
                    strWip0 = " (SO002.WipCode1 = 0 And SO002.WipCode2 = 0 And SO002.WipCode3 = 0 ) OR (SO002.WipCode1 = 0 And SO002.WipCode2 = 0 And SO002.WipCode3 = 0) "
                ElseIf varS > 0 And varS < 10 Then
                    strWip1 = strWip1 & "," & varS
                ElseIf varS > 20 Then
                    strWip2 = strWip2 & "," & varS
                ElseIf varS > 10 And varS < 20 Then
                    strWip3 = strWip3 & "," & varS
                End If
            Next
            If strWip0 <> "" Then strWipAll = " OR (" & strWip0 & ")"
            If strWip1 <> "" Then strWipAll = strWipAll & " OR (SO002.WipCode1 In (" & Mid(strWip1, 2) & "))"
            If strWip2 <> "" Then strWipAll = strWipAll & " OR (SO002.WipCode2 In (" & Mid(strWip2, 2) & "))"
            If strWip3 <> "" Then strWipAll = strWipAll & " OR (SO002.WipCode3 In (" & Mid(strWip3, 2) & "))"
            strWipAll = "( " & Mid(strWipAll, 4) & " )"
            subAnd strWipAll
        End If
        
        subAnd "0= SO033.CancelFlag and SO033.UCCode is Null"
        '#5102 增加CM001 By Kin 2009/06/10
        subAnd "CD019.AccIdP1=CM001.CodeNo(+) And SO033.CitemCode=CD019.CodeNo"
        If optDetail.Value = True Then
            Select Case cboGroup.ListIndex
                   Case 1
                        strGroupName = "GroupName={So3320A1.ServName};GroupNameStr=if isnull({So3320A1.ServName}) then '服務區: 無' Else '服務區:'+{So3320A1.ServName}"
                   Case 2
                        strGroupName = "GroupName={So3320A1.AreaName};GroupNameStr=if isnull({So3320A1.AreaName}) then '行政區: 無' Else '行政區:'+{So3320A1.AREANAME}"
                   Case 3
                        strGroupName = "GroupName={So3320A1.ClctAreaName};GroupNameStr=if isnull({So3320A1.ClctAreaName}) then '收費區: 無' Else '收費區:'+{So3320A1.ClctAreaName}"
                   Case 4
                        strGroupName = "GroupName={SO3320A1.ClctName};GroupNameStr=if isnull({So3320A1.ClctName}) then '收費人員: 無' Else '收費人員:'+{So3320A1.ClctName}"
                   Case 5
                        strGroupName = "GroupName={SO3320A1.MduName};GroupNameStr=if isnull({So3320A1.MduName}) then '大樓編號: 無' Else '大樓編號:'+{So3320A1.MduName}"
                   Case 6
                        strGroupName = "GroupName={SO3320A1.ClctEnWorkClass};GroupNameStr=if isnull({So3320A1.ClctEnWorkClass}) then '工作類別(收費人員): 無' Else '工作類別(收費人員):'+{So3320A1.ClctEnWorkClass}"
                   Case 7
                        strGroupName = "GroupName={SO3320A1.OldClctName};GroupNameStr=if isnull({So3320A1.OldClctName}) then '原收費人員: 無' Else '原收費人員:'+{So3320A1.OldClctName}"
                   Case Else
                        strGroupName = ""
            End Select
            If chkUseOldClctEn Then strGroupName = strGroupName & IIf(strGroupName <> "", ";", "") & "UseOldClctEn=True"
            If gimOrder.GetQryStr <> "" Then Call SetGiMultiOrder(strGroupName, gimOrder, , strMDBOrder)
        End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Private Function GetBillnoStr(intIndex As Integer, strValue As String, _
    Optional blnStart As Boolean = False) As String
    On Error Resume Next
    Dim strFormula As String
        If blnStart Then
            strFormula = " >= "
        Else
            strFormula = " <= "
        End If
        Select Case intIndex
            Case 0
                GetBillnoStr = "BillNo " & strFormula & strValue
            Case 1
                GetBillnoStr = "PrtSNo " & strFormula & strValue
            Case 2
                GetBillnoStr = "MediaBillNo " & strFormula & strValue
        End Select
End Function

Private Function subInsertCD002()
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
        Set rsTmp = gcnGi.Execute("Select CodeNo,Description From " & GetOwner & "CD002 ")
        While Not rsTmp.EOF
            cnn.Execute "Insert Into CD002 (CodeNo,Description) Values (" & _
                            GetNullString(rsTmp(0)) & "," & GetNullString(rsTmp(1)) & ")"
            rsTmp.MoveNext
            DoEvents
        Wend
        Set rsTmp = Nothing
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "InsertCD002")
End Function

Private Function GetSort() As String
    On Error GoTo chkErr
        Select Case cboGroup.ListIndex
            Case 1
                GetSort = "服務區": Exit Function
            Case 2
                GetSort = "行政區": Exit Function
            Case 3
                GetSort = "收費區": Exit Function
            Case 4
                GetSort = "收費人員": Exit Function
            Case 5
                GetSort = "大樓編號": Exit Function
            Case 6
                GetSort = "工作類別(收費人員)": Exit Function
        End Select
    GetSort = "無"
  Exit Function
chkErr:
    Call ErrSub(Me.Name, "Function GetSort")
End Function

Private Function GetQryStr() As String
    On Error GoTo chkErr
    Dim str1 As String
        strChooseString = ""
        If chkBad = 1 Then
            str1 = str1 & "是否包含呆帳資料統計:是 , "
        Else
            str1 = str1 & "是否包含呆帳資料統計:否 , "
        End If
        If chkGUINo.Value = 1 Then
            str1 = str1 & "是否已開發票:是 , "
        Else
            str1 = str1 & "是否已開發票:否 , "
        End If
        If chkInvoice.Value = 1 Then
            str1 = str1 & "是否已拋檔:是;"
        Else
            str1 = str1 & "是否已拋檔:否;"
        End If
        str1 = str1 & "發票種類:"
        Select Case True
            Case optInvoiceAll
                str1 = str1 & "全部 , "
            Case optInvoiceForward
                str1 = str1 & "預開 , "
            Case optInvoiceBack
                str1 = str1 & "後開 , "
        End Select
        str1 = str1 & "選擇條件: "
        If optClose Then
            str1 = str1 & "日結 , "
        ElseIf optNoClose Then
            str1 = str1 & "未日結 , "
        Else
            str1 = str1 & "全部 , "
        End If
        str1 = str1 & "首期條件: " & Mid(cboFirstFlag, 3) & ","
        str1 = str1 & "地址依據: " & Mid(cboAddress, 3) & ","
        str1 = str1 & "分頁方式:" & GetSort
        str1 = str1 & " , 報表格式:" & IIf(optDetail, "明細表", "簡表")
        
        
        strChooseString = "公司別:" & subSpace(gilCompCode.GetDescription) & ";" & _
                              "服務類別:" & subSpace(gimServiceType.GetDispStr) & ";" & _
                              "實收日期:" & subSpace(gdaRealDate1.GetValue(True)) & "~" & subSpace(gdaRealDate2.GetValue(True)) & ";" & _
                              "裝機日期:" & subSpace(gdaInstDay1.GetValue(True)) & "~" & subSpace(gdaInstDay2.GetValue(True)) & ";" & _
                              "應收日期:" & subSpace(gdaShouldDay1.GetValue(True)) & "~" & subSpace(gdaShouldDay2.GetValue(True)) & ";" & _
                              "異動時間:" & subSpace(gdtUpdTime1.GetValue(True)) & "~" & subSpace(gdtUpdTime2.GetValue(True)) & ";" & _
                              "單據類別:" & subSpace(gimBillType.GetDispStr) & ";" & _
                              "收費方式:" & subSpace(gimCMCode.GetDispStr) & ";" & _
                              "收費項目:" & subSpace(gimCitemCode.GetDispStr) & ";" & _
                              "客戶類別:" & subSpace(gimClassCode.GetDispStr) & ";" & _
                              "客戶狀態:" & subSpace(gimCustStatus.GetDispStr) & ";" & _
                              "服務區　:" & subSpace(gimServCode.GetDispStr) & ";"
        strChooseString = strChooseString & "行政區　:" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                              "收費區:" & subSpace(gimClctAreaCode.GetDispStr) & ";" & _
                              "大樓名稱:" & subSpace(gimMduId.GetDispStr) & ";" & _
                              "收費人員:" & subSpace(gimClctEn.GetDispStr) & ";" & _
                              "原收費人員:" & subSpace(gimOldClctEn.GetDispStr) & ";" & _
                              "街道範圍:" & subSpace(gimStrtCode.GetDispStr) & ";" & _
                              "短收原因:" & subSpace(gimSTCode.GetDispStr) & ";" & _
                              "異動人員:" & subSpace(gimUpdEn.GetDispStr) & ";" & _
                              "銀行名稱:" & subSpace(gimBankCode.GetDispStr) & ";" & _
                              "排序方式:" & subSpace(gimOrder.GetDispStr) & ";" & _
                              "網路編號:" & subSpace(txtCircuitNo) & ";" & _
                              "地址依據:" & subSpace(cboAddress.Text) & ",  " & _
                              IIf(chkUseOldClctEn.Value = 1, "列印顯示原收人員, ", "") & _
                              "客戶編號:" & subSpace(txtCustId) & ", " & _
                              "銷售點:" & subSpace(gilIntroId.Text) & ";" & _
                              "期數 " & subSpace(ginRealPeriod1.Text & "~" & ginRealPeriod2.Text) & ", 金額 " & subSpace(txtRealAmt1.Text & "~" & txtRealAmt2.Text) & ";" & _
                              "會計科目:" & subSpace(gimCM001.GetDispStr) & ";" & str1
        
                        
        GetQryStr = strChooseString
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "Function GetQryStr")
End Function

Private Sub Form_Activate()
    On Error Resume Next
    Screen.MousePointer = vbDefault
    'PrcList Me
End Sub

Private Sub Form_Load()
    On Error GoTo chkErr
        If Alfa2 Then GetGlobal
        frmSO3320A.Caption = frmSO3320A.Caption & " [" & uForm & "]"
        subGil
        subGim
        DefaultValue
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub DefaultValue()
    On Error Resume Next
    Dim intPara23 As Integer
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        gdaRealDate1.SetValue RightDate
        cboAddress.ListIndex = 0
        cboGroup.ListIndex = 0
        cboFirstFlag.ListIndex = 2
        cboBillType.ListIndex = 1
        cboBillType.Enabled = False
'        intPara23 = GetSystemParaItem("Para23", gilCompCode.GetCodeNo, "", "SO043")
'        If intPara23 = 1 Then
'            cboBillType.ListIndex = 1
'        ElseIf intPara23 = 2 Then
'            cboBillType.ListIndex = 2
'        Else
'            cboBillType.ListIndex = 0
'        End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo chkErr
        If Not ChkGiList(KeyCode) Then Exit Sub
        Select Case KeyCode
               Case vbKeyEscape
                    If cmdExit.Enabled = True Then cmdExit.Value = True
               Case vbKeyF5
                    If cmdPrint.Enabled Then cmdPrint.Value = True
               Case vbKeyF2
                    If cmdExeclOk.Enabled Then cmdExeclOk.Value = True
        End Select
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub subGim()
    On Error GoTo chkErr
        SetgiMulti gimServiceType, "CodeNo", "Description", "CD046", "服務類別代碼", "服務類別名稱"
        SetgiMulti gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱"
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱"
        SetgiMulti gimPTCode, "CodeNo", "Description", "CD032", "付款種類代碼", "付款種類名稱"
        SetgiMulti gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱"
        SetgiMulti gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱"
        SetgiMulti gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱"
        SetgiMulti gimMduId, "MduId", "Name", "SO017", "大樓編號", "大樓名稱"
        SetgiMulti gimClctEn, "EmpNo", "EmpName", "CM003", "收費員代碼", "收費員名稱"
        SetgiMulti gimCustStatus, "CodeNo", "Description", "CD035", "客戶狀態代碼", "客戶狀態名稱"
        SetgiMulti gimStrtCode, "CodeNo", "Description", "CD017", "街道代碼", "街道名稱"
        SetgiMulti gimSTCode, "CodeNo", "Description", "CD016", "短收原因代碼", "短收原因名稱"
        SetgiMulti gimBankCode, "CodeNo", "Description", "CD018", "銀行代碼", "銀行名稱"
        SetgiMulti gimUpdEn, "EmpNo", "EmpName", "CM003", "異動人員代碼", "異動人員名稱"
        SetgiMulti gimCardCode, "CodeNo", "Description", "CD037", "信用卡代碼", "信用卡名稱"
        SetgiMulti gimWipCode, "CodeNo", "Description", "CD036", "代碼", "名稱"
        SetgiMultiAddItem gimBillType, "B,I,P,M,T", "收費單,裝機單,停拆移機單,維修單,臨時收費單"
        SetgiMulti gimWorkClass, "CodeNo", "Description", "CD071", "代碼", "名稱"
        SetgiMulti gimClctAreaCode, "CodeNo", "Description", "CD040", "代碼", "名稱"
        SetgiMulti gimOldClctEn, "EmpNo", "EmpName", "CM003", "代碼", "名稱"
        '#5102 增加會計科目 By Kin 2009/06/10
        SetgiMulti gimCM001, "CodeNo", "Description ", "CM001", "會計科目代碼", "會計科目名稱"
        'Call SetgiMultiAddItem(gimWorkClass, "A,B,C,D,E", "客服,工程,收費,業務,其他", "工作類別代碼", "工作類別名稱")
        SetgiMultiAddItem gimOrder, "A,B,C,D,E,F,G,H,I,J", "收費人員,客戶編號,實收日期,地址,服務區,實收金額,實收期數,原收費人員,異動時間,收費方式"
        '#5683 增加繳付類別 By Kin 2010/08/17
        'Call SetgiMultiAddItem(gimPayType, "0,1", "預付制,現付制", "代碼", "名稱")
        SetgiMulti gimPayType, "CodeNo", "Description", "CD112", "代碼", "名稱"
        If GetPaynowFlag Then
            With gimPayType
                '.SetDispStr "預付制"
                .SetQueryCode "0"
            End With
        Else
            gimPayType.Clear
            gimPayType.Enabled = False
        End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGil()
  On Error GoTo chkErr
    SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
    SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
    SetgiList gilIntroId, "IntroId", "NameP", "So013", , , , , 10, 20
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Function IsDataOk() As Boolean
    On Error GoTo chkErr
        If Not ChkDTok Then Exit Function
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        'If Not MustExist(gilServiceType, 2, "服務類別") Then Exit Function
        If gimServiceType.GetQueryCode = "" Then MsgMustBe: Exit Function
        IsDataOk = True
        Exit Function
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)
End Sub

Private Sub gdaInstDay2_GotFocus()
    On Error Resume Next
        If Len(gdaInstDay1.Text) > 0 Then gdaInstDay2.SetValue gdaInstDay1.GetValue
End Sub

Private Sub gdaRealDate2_GotFocus()
    On Error Resume Next
        If Len(gdaRealDate1.Text) > 0 Then gdaRealDate2.SetValue gdaRealDate1.GetValue
End Sub

Private Sub gdaRealDate2_Validate(Cancel As Boolean)
    On Error Resume Next
        If gdaRealDate1.Text <> "" And gdaRealDate2.Text < gdaRealDate1.Text Then
            Cancel = True
            MsgDateRangeX
        End If
End Sub

Public Property Get uForm() As String
    On Error GoTo chkErr
        uForm = strForm
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uForm"
End Property

Public Property Let uForm(ByVal vForm As String)
    On Error GoTo chkErr
        strForm = vForm
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uForm"
End Property

Private Sub gdaShouldDay2_GotFocus()
    On Error Resume Next
        If Len(gdaShouldDay1.GetValue) > 0 Then gdaShouldDay2.SetValue gdaShouldDay1.GetValue
End Sub

Private Sub gdaShouldDay2_Validate(Cancel As Boolean)
    On Error Resume Next
        If gdaShouldDay1.Text <> "" And gdaShouldDay2.Text < gdaShouldDay1.Text Then
            Cancel = True
            MsgDateRangeX
        End If

End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If strForm = "SO3320A" Then
            If Not ChgComp(gilCompCode, "SO3300", "SO3320") Then Exit Sub
        Else
            If Not ChgComp(gilCompCode, "SO5200", "SO5250") Then Exit Sub
        End If
        Call subGil
        Call subGim
        'Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gimServiceType.SelectAll
        'gilServiceType.SetCodeNo ""
        'gilServiceType.Query_Description
        'gilServiceType.ListIndex = 1
    '    Call GiMultiFilter(gimNodeNo, gilServiceType.GetCodeNo, gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimServCode, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimAreaCode, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimStrtCode, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimMduId, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimClctEn, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimUpdEn, , gilCompCode.GetCodeNo)
End Sub

Private Sub gilServiceType_Change()
    On Error Resume Next
'        Call GiMultiFilter(gimNodeNo, gilServiceType.GetCodeNo, gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimCMCode, gilServiceType.GetCodeNo)
        Call GiMultiFilter(gimCitemCode, gilServiceType.GetCodeNo)
        Call GiMultiFilter(gimClassCode, gilServiceType.GetCodeNo)
        Call GiMultiFilter(gimSTCode, gilServiceType.GetCodeNo)
End Sub

Private Sub gimBillType_GotFocus()
    On Error Resume Next
'        vsl1.Value = 0
        
End Sub

Private Sub gimClctEn_GotFocus()
'    vsl1.Value = 100

End Sub

Private Sub gimOrder_GotFocus()
    On Error Resume Next
'        vsl1.Value = 100
End Sub

Private Sub gimServiceType_Change()
    On Error Resume Next
        GiMultiFilter gimCitemCode, gimServiceType.GetQryStr, , True
        GiMultiFilter gimCMCode, gimServiceType.GetQryStr, , True
        GiMultiFilter gimClassCode, gimServiceType.GetQryStr, , True
        GiMultiFilter gimSTCode, gimServiceType.GetQryStr, , True
    

End Sub

Private Sub gimWorkClass_Change()
  On Error GoTo chkErr
    If gimWorkClass.GetQryStr <> "" Then
        gimClctEn.Filter = " Where WorkClass " & gimWorkClass.GetQryStr
    Else
        gimClctEn.Filter = ""
    End If
  Exit Sub
chkErr:
    ErrSub Me.Name, "gimWorkClass_Change"
End Sub

Private Sub gimWorkClass_GotFocus()
'    If vsl1.Value = 0 Then
'        vsl1.Value = 100
'    Else
'        vsl1.Value = 0
'    End If
End Sub

Private Sub txtRealAmt1_Change()
    On Error Resume Next
        txtRealAmt2.Text = txtRealAmt1.Text
End Sub

Private Sub ginRealPeriod1_Change()
    On Error Resume Next
        ginRealPeriod2.Text = ginRealPeriod1.Value
End Sub

Private Sub optDetail_Click()
    On Error Resume Next
        cboGroup.Enabled = True
        cmdExeclDetail.Enabled = True
'        gimCM001.Clear
'        gimCM001.Enabled = False
End Sub

Private Sub optSimple_Click()
    On Error Resume Next
        cboGroup.Enabled = False
        cmdExeclDetail.Enabled = False
'        gimCM001.Enabled = True
End Sub

Private Function GetCompareDate(strField As String, strCompareDate As String, Optional blnFrist As Boolean = True) As String
    On Error GoTo chkErr
        If blnFrist Then
            GetCompareDate = strField & " > = To_Date(" & strCompareDate & ",'yyyymmdd') "
        Else
            GetCompareDate = strField & " < To_Date(" & strCompareDate & ",'yyyymmdd') + 1 "
        End If
    Exit Function
chkErr:
    ErrSub Me.Name, "GetCompareDate"
End Function

Public Function GetSubTotal(strSQL As String, _
        strFld1 As String, Optional strFld2 As String, _
        Optional strFld3 As String, Optional strFld4 As String) As String
    On Error Resume Next
    Dim strGroupSQL As String
    Dim strGroupField As String
    Dim strSummField As String
        strGroupField = strFld1
        strSummField = " Null " & strFld1
        If Len(strFld2) > 0 Then
            strGroupField = strGroupField & "," & strFld2
            strSummField = strSummField & ", Null " & strFld2
        Else
            strGroupField = strGroupField & ",Null Type2 "
            strSummField = strSummField & ", Null  Type2"
        End If
        If Len(strFld3) > 0 Then
            strGroupField = strGroupField & "," & strFld3
            strSummField = strSummField & ", Null " & strFld3
        Else
            strGroupField = strGroupField & ",Null Type3 "
            strSummField = strSummField & ", Null  Type3"
        End If
        If Len(strFld4) > 0 Then
            strGroupField = strGroupField & "," & strFld4
            strSummField = strSummField & ", Null " & strFld4
        Else
            strGroupField = strGroupField & ",Null Type4 "
            strSummField = strSummField & ", Null  Type4"
        End If
        
        strGroupSQL = "Select " & strGroupField & _
            " ,Sum(RealAmt) RealAmt ,Sum(RealPeriod) RealPeriod, " & _
            " Count(Distinct CustId) CountCustId,Count(Distinct BillNo) CountBillNo,Count(*) TotalCount " & _
            " From " & strSQL & " Group By " & strGroupField & _
            " Union All Select " & strSummField & ", Sum(RealAmt) RealAmt,Sum(RealPeriod)," & _
            " Count(Distinct CustId) CountCustId,Count(Distinct BillNo) CountBillNo,Count(*) TotalCount " & _
            " From " & strSQL
        GetSubTotal = strGroupSQL
End Function

Private Function InsertSubTotal(strSQL As String, StrTableName As String, Optional lngAffected As Long) As Boolean
    On Error Resume Next
        cnn.Execute "Drop Table " & StrTableName
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
        cnn.Execute "Create Table " & StrTableName & _
                " (GroupField1 Text(40),GroupField2 Text(40),GroupField3 Text(40),GroupField4 Text(40)," & _
                " RealAmt Single,RealPeriod Single,CountCustId Single,CountBillNo Single,TotalCount Single) "
        If Not GetRS(rsTmp, strSQL, gcnGi) Then Exit Function
        On Error GoTo ErrTrans
        cnn.BeginTrans
        Do While Not rsTmp.EOF
            cnn.Execute "Insert Into " & StrTableName & " (GroupField1,GroupField2,GroupField3,GroupField4," & _
                "RealAmt,RealPeriod,CountCustId,CountBillNo,TotalCount) Values (" & _
                GetNullString(rsTmp(0)) & "," & GetNullString(rsTmp(1)) & "," & _
                GetNullString(rsTmp(2)) & "," & GetNullString(rsTmp(3)) & "," & _
                GetNullString(rsTmp(4)) & "," & GetNullString(rsTmp(5)) & "," & _
                GetNullString(rsTmp(6)) & "," & GetNullString(rsTmp(7)) & "," & _
                GetNullString(rsTmp(8)) & ")"
            rsTmp.MoveNext
        Loop
        lngAffected = rsTmp.RecordCount
        cnn.CommitTrans
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
        InsertSubTotal = True
    Exit Function
ErrTrans:
    cnn.RollbackTrans
chkErr:
    lngAffected = 0
    ErrSub Me.Name, "InsertSubTotoal"
End Function

Private Sub txtCircuitNo_Change()
    On Error Resume Next
        If txtCircuitNo = "" Then
            fraCircuitNo.Visible = True
        Else
            fraCircuitNo.Visible = False
        End If
End Sub

Private Sub txtCustId_KeyPress(KeyAscii As Integer)
    On Error Resume Next
        If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 44 Or KeyAscii = 8 Or KeyAscii = 45 Then
            If KeyAscii = 44 Or KeyAscii = 45 Then
                If Asc(Right(" " & txtCustId.Text, 1)) = 44 Or Asc(Right(" " & txtCustId.Text, 1)) = 45 Or txtCustId = "" Then KeyAscii = 9
            End If
            If Not ChkMaxLengthOK(txtCustId) Then
               If Not (KeyAscii = 44 Or KeyAscii = 8 Or KeyAscii = 45) Then KeyAscii = 9
            End If
        Else
            KeyAscii = 9
        End If
End Sub

Private Sub SetGiMultiOrder(ByRef strFormulaName As String, _
        gimOrder As Object, Optional ByRef strGroupName As String, _
        Optional ByRef strMDBOrder As String)
    On Error GoTo chkErr
      '排序方式
    Dim varCollect As Variant
    Dim varSplit As Variant
    Dim strOrderName As String
    Dim lngLoop As Long
        varCollect = Split(Replace(gimOrder.GetColumnOrderCode, "'", ""), ",")
        lngLoop = 0
        For Each varSplit In varCollect
            Select Case Trim(varSplit)
                Case "A"
                    strOrderName = "{So3320A1.ClctEn}"
                Case "B"
                    strOrderName = "{So3320A1.CustId}"
                Case "C"
                    strOrderName = "{So3320A1.RealDate}"
                Case "D"
                    strOrderName = "{So3320A1.AddrSort}"
                Case "E"
                    strOrderName = "{So3320A1.ServName}"
                Case "F"
                    strOrderName = "{So3320A1.RealAmt}"
                Case "G"
                    strOrderName = "{So3320A1.RealPeriod}"
                Case "H"
                    strOrderName = "{So3320A1.OldClctEn}"
                Case "I"
                    strOrderName = "{So3320A1.UpdTime}"
                Case "J"
                    strOrderName = "{So3320A1.CMCode}"
            End Select
            If cboGroup.ListIndex = 0 And lngLoop = 0 Then strFormulaName = strFormulaName & IIf(strFormulaName = "", "", ";") & "GroupName=" & strOrderName
            Call ReturnAndString(strFormulaName, "Sort" & Chr(Asc("A") + lngLoop) & "=" & strOrderName, ";")
            
            strMDBOrder = strMDBOrder & "," & Mid(strOrderName, 2, Len(strOrderName) - 2)
            
            lngLoop = lngLoop + 1
        Next
        Call ReturnAndString(strFormulaName, "Sort" & Chr(Asc("A") + lngLoop) & "={So3320A1.CustId}", ";")
        strMDBOrder = Mid(strMDBOrder, 2)
    Exit Sub
chkErr:
    ErrSub Me.Name, "SetGiMultiOrder"
End Sub

Private Function GetUseIndStr(StrTableName As String, strColumnName As String) As String
    On Error Resume Next
        'GetUseIndStr = GetUseIndexStr(StrTableName, strColumnName)
End Function

Private Sub txtRealAmt1_KeyPress(KeyAscii As Integer)
    EnableKeyNumber KeyAscii

End Sub

Private Sub txtRealAmt2_KeyPress(KeyAscii As Integer)
    EnableKeyNumber KeyAscii
End Sub

Private Sub vsl1_Change()
    On Error Resume Next
        fraMulti.Top = 20 - fraMulti.Height * (vsl1.Value / 20) / 6
End Sub

Public Sub EnableKeyNumber(ByRef KeyAscii As Integer, Optional ByVal blnMinus As Boolean = True)
    On Error Resume Next
        If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeySpace Or KeyAscii = vbKeyBack Then Exit Sub
        If blnMinus And Chr(KeyAscii) = "-" Then Exit Sub
        KeyAscii = 0
End Sub
