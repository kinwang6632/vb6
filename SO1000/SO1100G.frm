VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSO1100G 
   BorderStyle     =   1  '單線固定
   Caption         =   "申請轉帳記錄 [SO1100G]"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   1425
   ClientWidth     =   11880
   Icon            =   "SO1100G.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNote 
      Height          =   450
      Left            =   750
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   73
      Top             =   3450
      Width           =   8235
   End
   Begin VB.CommandButton cmdACHDetail 
      Caption         =   "ACH授權細項查詢"
      Height          =   315
      Left            =   7740
      TabIndex        =   68
      Top             =   1800
      Width           =   1905
   End
   Begin VB.Frame fraData 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      Left            =   120
      TabIndex        =   35
      Top             =   -30
      Width           =   11625
      Begin VB.TextBox txtAchSN 
         Height          =   315
         Left            =   8640
         MaxLength       =   20
         TabIndex        =   71
         Top             =   2310
         Width           =   1905
      End
      Begin VB.ComboBox cboProposer 
         Height          =   300
         ItemData        =   "SO1100G.frx":0442
         Left            =   1020
         List            =   "SO1100G.frx":0444
         TabIndex        =   0
         Top             =   150
         Width           =   2355
      End
      Begin VB.CommandButton cmdCreateINV 
         BackColor       =   &H00C0FFFF&
         Caption         =   "發票抬頭維護"
         Enabled         =   0   'False
         Height          =   315
         Left            =   210
         Style           =   1  '圖片外觀
         TabIndex        =   66
         Top             =   3090
         Width           =   1485
      End
      Begin VB.CommandButton cmdInvSeqNo 
         Caption         =   "選擇發票抬頭"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5610
         TabIndex        =   65
         Top             =   3090
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CheckBox chkDeAut 
         Caption         =   "取消授權"
         Height          =   210
         Left            =   3060
         TabIndex        =   27
         Top             =   2460
         Width           =   1065
      End
      Begin VB.Frame FraACH 
         Height          =   1440
         Left            =   7560
         TabIndex        =   60
         Top             =   830
         Width           =   3915
         Begin VB.CheckBox chkToACH 
            Caption         =   "轉ACH授權"
            Height          =   210
            Left            =   3210
            TabIndex        =   23
            Top             =   90
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.TextBox txtACHcustID 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1230
            MaxLength       =   20
            TabIndex        =   22
            Top             =   230
            Width           =   1665
         End
         Begin CS_Multi.CSmulti csmACH 
            Height          =   375
            Left            =   60
            TabIndex        =   24
            Top             =   570
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   661
            ButtonCaption   =   "ACH授權交易別"
         End
         Begin VB.Label lblACHid 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "ACH用戶號碼"
            Height          =   180
            Left            =   60
            TabIndex        =   63
            Top             =   300
            Width           =   1110
         End
         Begin VB.Label lblState 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "狀態 : "
            Height          =   180
            Left            =   2010
            TabIndex        =   62
            Top             =   1110
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "未授權"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   2640
            TabIndex        =   61
            Top             =   1110
            Visible         =   0   'False
            Width           =   780
         End
      End
      Begin VB.CheckBox chkAddAcc 
         Caption         =   "待收為新週期,下次出帳以此帳戶繳費"
         Height          =   255
         Left            =   8250
         TabIndex        =   31
         Top             =   2745
         Width           =   3225
      End
      Begin Gi_Multi.GiMulti gimCitem 
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   4290
         Visible         =   0   'False
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   661
         ButtonCaption   =   "收 費 項 目"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin VB.TextBox txtCVC2 
         Height          =   315
         Left            =   10950
         MaxLength       =   3
         TabIndex        =   7
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox chkFore2 
         Caption         =   "外籍人士"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   6480
         TabIndex        =   13
         Top             =   1200
         Width           =   1035
      End
      Begin VB.CheckBox chkFore 
         Caption         =   "外籍人士"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5820
         TabIndex        =   2
         Top             =   210
         Width           =   1065
      End
      Begin VB.TextBox txtAccountNo 
         Height          =   315
         Left            =   7920
         MaxLength       =   16
         TabIndex        =   6
         Top             =   480
         Width           =   2370
      End
      Begin VB.TextBox txtHide 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  '暫止
         Left            =   6120
         Locked          =   -1  'True
         PasswordChar    =   "*"
         TabIndex        =   10
         Text            =   "*******"
         Top             =   810
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtIntroId 
         Height          =   285
         Left            =   4380
         MaxLength       =   10
         TabIndex        =   15
         Top             =   1470
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Appearance      =   0  '平面
         Height          =   285
         Left            =   7230
         Picture         =   "SO1100G.frx":0446
         Style           =   1  '圖片外觀
         TabIndex        =   16
         Top             =   1470
         Width           =   285
      End
      Begin VB.TextBox txtAccountOwner 
         Height          =   315
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1140
         Width           =   2355
      End
      Begin VB.TextBox txtProposer 
         Height          =   315
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   67
         Top             =   150
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.TextBox txtAccOwnerID 
         Height          =   315
         IMEMode         =   2  '關閉
         Left            =   5130
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1140
         Width           =   1215
      End
      Begin VB.CheckBox chkStop 
         Caption         =   "停用"
         Height          =   210
         Left            =   240
         TabIndex        =   25
         Top             =   2460
         Width           =   675
      End
      Begin VB.TextBox txtID 
         Height          =   315
         IMEMode         =   2  '關閉
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   1
         Top             =   150
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskCardExpDate 
         Height          =   315
         Left            =   6120
         TabIndex        =   9
         Tag             =   "OK"
         Top             =   810
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin Gi_Time.GiTime gdtAcceptTime 
         Height          =   315
         Left            =   1020
         TabIndex        =   20
         Top             =   2070
         Width           =   1935
         _ExtentX        =   3413
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
      Begin prjGiList.GiList gilAcceptName 
         Height          =   315
         Left            =   3840
         TabIndex        =   21
         Top             =   2070
         Width           =   3030
         _ExtentX        =   5345
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
         FldWidth1       =   880
         FldWidth2       =   1820
         F2Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin prjGiList.GiList gilMediaCode 
         Height          =   285
         Left            =   1020
         TabIndex        =   14
         Top             =   1470
         Width           =   2355
         _ExtentX        =   4154
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
         FldWidth1       =   405
         FldWidth2       =   1620
         F2Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin Gi_Date.GiDate gdaContiDate 
         Height          =   285
         Left            =   3840
         TabIndex        =   18
         Top             =   1770
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ReplaceText     =   -1  'True
      End
      Begin Gi_Date.GiDate gdaSendDate 
         Height          =   285
         Left            =   1020
         TabIndex        =   17
         Top             =   1770
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ReplaceText     =   -1  'True
      End
      Begin Gi_Date.GiDate gdaSnactionDate 
         Height          =   285
         Left            =   6120
         TabIndex        =   19
         Top             =   1770
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ReplaceText     =   -1  'True
      End
      Begin Gi_Date.GiDate gdaStopDate 
         Height          =   285
         Left            =   1860
         TabIndex        =   26
         Top             =   2400
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ReplaceText     =   -1  'True
      End
      Begin prjGiList.GiList gilBankCode 
         Height          =   315
         Left            =   4290
         TabIndex        =   5
         Top             =   480
         Width           =   2580
         _ExtentX        =   4551
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
         FldWidth1       =   405
         FldWidth2       =   1850
         F2Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin prjGiList.GiList gilCardName 
         Height          =   315
         Left            =   1020
         TabIndex        =   8
         Top             =   810
         Width           =   2355
         _ExtentX        =   4154
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
         FldWidth1       =   405
         FldWidth2       =   1620
         F2Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin prjGiList.GiList gilCMCode 
         Height          =   315
         Left            =   7920
         TabIndex        =   3
         Top             =   150
         Width           =   2370
         _ExtentX        =   4180
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
         FldWidth1       =   420
         FldWidth2       =   1620
         F2Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin CS_Multi.CSmulti csmCitem2 
         Height          =   375
         Left            =   4200
         TabIndex        =   30
         Top             =   2700
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   661
         ButtonCaption   =   "待收 收費項目"
      End
      Begin CS_Multi.CSmulti csmCitem1 
         Height          =   375
         Left            =   180
         TabIndex        =   29
         Top             =   2700
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   661
         ButtonCaption   =   "週期 收費項目"
      End
      Begin prjGiList.GiList gilPTCode 
         Height          =   315
         Left            =   1020
         TabIndex        =   4
         Top             =   480
         Width           =   2355
         _ExtentX        =   4154
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
         FldWidth1       =   405
         FldWidth2       =   1620
         TableName       =   "CD019"
         F2Corresponding =   -1  'True
      End
      Begin Gi_Date.GiDate gdaAuthStopDate 
         Height          =   285
         Left            =   5770
         TabIndex        =   28
         Top             =   2400
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ReplaceText     =   -1  'True
      End
      Begin VB.Label lblApplication 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "申請書單號"
         Height          =   180
         Left            =   7620
         TabIndex        =   72
         Top             =   2370
         Width           =   900
      End
      Begin VB.Label lblInvSeqNo 
         Height          =   195
         Left            =   2820
         TabIndex        =   70
         Top             =   3150
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "發票流水號:"
         Height          =   255
         Left            =   1770
         TabIndex        =   69
         Top             =   3150
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "授權截止日"
         Height          =   180
         Left            =   4830
         TabIndex        =   64
         Top             =   2460
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "付款種類"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   180
         TabIndex        =   59
         Top             =   540
         Width           =   780
      End
      Begin VB.Label lblCVC2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "檢查碼"
         Height          =   180
         Left            =   10350
         TabIndex        =   58
         Top             =   540
         Width           =   540
      End
      Begin VB.Label lblUpdTime 
         AutoSize        =   -1  'True
         Caption         =   "yyyy/mm/dd hh:mm:ss"
         Height          =   195
         Left            =   9810
         TabIndex        =   57
         Top             =   3510
         Width           =   1605
      End
      Begin VB.Label lblUpdEn 
         AutoSize        =   -1  'True
         Caption         =   "XXXXXXXX"
         Height          =   195
         Left            =   9810
         TabIndex        =   56
         Top             =   3750
         Width           =   990
      End
      Begin VB.Label lblUTime 
         AutoSize        =   -1  'True
         Caption         =   "異動時間:"
         Height          =   195
         Left            =   8970
         TabIndex        =   55
         Top             =   3510
         Width           =   795
      End
      Begin VB.Label lblUEn 
         AutoSize        =   -1  'True
         Caption         =   "異動人員:"
         Height          =   195
         Left            =   8970
         TabIndex        =   54
         Top             =   3750
         Width           =   825
      End
      Begin VB.Label lblCMCode 
         Caption         =   "收費方式"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   7110
         TabIndex        =   53
         Top             =   210
         Width           =   750
      End
      Begin VB.Label lblAccountNo 
         AutoSize        =   -1  'True
         Caption         =   "帳號/卡號"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   7095
         TabIndex        =   52
         Top             =   540
         Width           =   765
      End
      Begin VB.Label lblCardExpDate 
         Caption         =   "信用卡有效期限 (西曆MM/YYYY)"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3450
         TabIndex        =   51
         Top             =   870
         Width           =   2655
      End
      Begin VB.Label lblAcceptName 
         AutoSize        =   -1  'True
         Caption         =   "受理人員"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3060
         TabIndex        =   50
         Top             =   2130
         Width           =   795
      End
      Begin VB.Label lblAcceptTime 
         AutoSize        =   -1  'True
         Caption         =   "受理時間"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   180
         TabIndex        =   49
         Top             =   2130
         Width           =   795
      End
      Begin VB.Label lblMediaCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "介紹媒介"
         Height          =   180
         Left            =   180
         TabIndex        =   48
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lblIntro 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "介紹人代碼"
         Height          =   180
         Left            =   3450
         TabIndex        =   47
         Top             =   1530
         Width           =   900
      End
      Begin VB.Label lblIntroName 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '單線固定
         Height          =   285
         Left            =   5475
         TabIndex        =   32
         Top             =   1470
         Width           =   1755
      End
      Begin VB.Label lblAccountOwner 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "帳號所有人"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   90
         TabIndex        =   46
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label lblProposer 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "申請人"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   330
         TabIndex        =   45
         Top             =   210
         Width           =   555
      End
      Begin VB.Label lblAccNameID 
         Caption         =   "帳號所有人身分證號"
         Height          =   195
         Left            =   3450
         TabIndex        =   44
         Top             =   1200
         Width           =   1650
      End
      Begin VB.Label lblDelDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "送件日期"
         Height          =   180
         Left            =   180
         TabIndex        =   43
         Top             =   1830
         Width           =   720
      End
      Begin VB.Label lblContiDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "銀行核印日期"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   2680
         TabIndex        =   42
         Top             =   1830
         Width           =   1080
      End
      Begin VB.Label lblSnDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "核准日期"
         ForeColor       =   &H00C000C0&
         Height          =   180
         Left            =   5310
         TabIndex        =   41
         Top             =   1830
         Width           =   720
      End
      Begin VB.Label lblStopDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "停用日期"
         Height          =   180
         Left            =   1050
         TabIndex        =   40
         Top             =   2460
         Width           =   720
      End
      Begin VB.Label lblID 
         Caption         =   "身分證號"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3450
         TabIndex        =   39
         Top             =   210
         Width           =   795
      End
      Begin VB.Label lblBankCode 
         AutoSize        =   -1  'True
         Caption         =   "銀行"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3795
         TabIndex        =   38
         Top             =   540
         Width           =   360
      End
      Begin VB.Label lblCardName 
         AutoSize        =   -1  'True
         Caption         =   "信用卡別"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   870
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "備註"
         Height          =   180
         Left            =   240
         TabIndex        =   36
         Top             =   3630
         Width           =   375
      End
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   3420
      Left            =   60
      TabIndex        =   33
      Top             =   4080
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   6033
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
   Begin CS_Multi.CSmulti csmTest 
      Height          =   375
      Left            =   150
      TabIndex        =   74
      Top             =   7650
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   661
      ButtonCaption   =   "ACH授權交易別"
   End
End
Attribute VB_Name = "frmSO1100G"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmer: Power Hammer
' Last Modify: 2002/06/27

Option Explicit
Private lngEditMode As giEditModeEnu ' 記錄目前在編輯、新增或檢視模式
Private WithEvents rsSO106 As ADODB.Recordset
Attribute rsSO106.VB_VarHelpID = -1
Private intCD031RefNo As Integer
Private blnVAcc As Boolean
Private blnACHAccountNoCanEdit As Boolean
Private strRealAccNo As String
Private strAchChange As String
Private blnAchChgClick As Boolean
Private blnAddFlag As Boolean
Private blnACHCustID As Boolean
Private strInvSeqNo As String
Private blnCallSO106A As Boolean
Private strACHID As String
Private strACHDesc As String
Private strFirstACHID As String
Private strFirstACHDesc As String
Private blnHaveUpd As Boolean
Private blnNoShowMsg As Boolean
Private blnInsSO138 As Boolean
Private blnEnableStop As Boolean
Private blnUCCodeIsNull As Boolean
Private strAccountName As String
Private posAchContition As String
Private AchTypeCondition As String
Private blnStartPost As Boolean
Private Sub chkFore_Click()
  On Error Resume Next
    If chkFore.Value Then
        txtID.MaxLength = 30
    Else
        txtID.MaxLength = 10
        txtID = Left(txtID, 10)
    End If
End Sub
'#3982 檢查SO138是否有帳號所有人 By Kin 2008/07/16
Private Function ChkSO138() As Boolean
  On Error GoTo ChkErr
    Dim strQry As String
    strQry = Empty
    If txtAccountOwner.Text & "" <> "" Then
        strQry = "Select Count(*) From " & GetOwner & "SO138" & _
                " Where ChargeTitle='" & Trim(txtAccountOwner.Text) & "'"
            
        If gcnGi.Execute(strQry)(0) > 0 Then
        
            ChkSO138 = True
        Else
            ChkSO138 = False
        End If
    
    End If
    Exit Function
ChkErr:
  ErrSub Me.Name, "ChkSO138"
End Function
'#3982 如果無資料則新增一筆資料至138並自動代入H項中 By Kin 2008/07/16
Private Sub InsSO138()
  On Error GoTo ChkErr
    Dim rsSo001 As New ADODB.Recordset
    Dim rsSo002 As New ADODB.Recordset
    
    Dim rsSO138Tmp As New ADODB.Recordset
    Dim strSo001Sql As String
    Dim strSo002Sql As String
    Dim aMsg As String
    Dim blnElectron As Boolean
    blnElectron = False
    '#5945 將此行拿掉,改用SO002的資訊
    'blnElectron = SetInvKind(aMsg)
    If aMsg <> Empty Then
        If MsgBox(aMsg, vbQuestion + vbYesNo, "詢問") = vbYes Then
            blnElectron = True
        End If
    End If
    strSo001Sql = "Select A.CustName,A.InstAddrNo,A.InstAddress,A.ChargeAddrNo,B.InvTitle,B.InvAddress," & _
                    "A.ChargeAddress,A.MailAddrNo,A.MailAddress,B.InvoiceType,B.InvNo, " & _
                    "B.ApplyInvDate,B.InvPurposeCode,B.InvPurposeName,B.DenRecCode,B.DenRecName,B.DenRecDate " & _
                    " From " & GetOwner & "SO001 A," & GetOwner & "SO002 B " & _
                    " Where A.Custid=" & gCustId & " And A.CompCode=" & gCompCode & _
                    " And B.CustId=A.CustId And B.CompCode=A.CompCode" & _
                    " And B.ServiceType='" & gServiceType & "'"
    Set rsSo001 = gcnGi.Execute(strSo001Sql)
    strSo002Sql = "SELECT * FROM " & GetOwner & "SO002 " & _
                " WHERE CUSTID = " & gCustId & _
                " AND COMPCODE = " & gCompCode & _
                " AND SERVICETYPE = '" & gServiceType & "'"
    If Not GetRS(rsSo002, strSo002Sql, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    
    If Not GetRS(rsSO138Tmp, "Select * From " & GetOwner & "SO138 Where 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If rsSo001.EOF Or rsSo001.BOF Then Exit Sub
    With rsSO138Tmp
        .AddNew
        .Fields("InvSeqNo") = Get138SeqNo
        
        .Fields("ChargeTitle") = txtAccountOwner.Text
        .Fields("InvTitle") = txtAccountOwner.Text
        .Fields("InvAddress") = rsSo001("InvAddress") & ""
        '*********************************************************************************
        '#4210 如果是虛擬帳號要帶入SO001的裝機地址 By Kin 2008/11/06
        If intCD031RefNo = 2 Or intCD031RefNo = 4 Or intCD031RefNo = 5 Then
            If rsSo001("ChargeAddrNo") > 0 And Not IsNull(rsSo001("ChargeAddrNo")) Then
                .Fields("ChargeAddrNo") = rsSo001("ChargeAddrNo")
                .Fields("ChargeAddress") = rsSo001("ChargeAddress") & ""
            End If
            If rsSo001("MailAddrNo") > 0 And Not IsNull(rsSo001("MailAddrNo")) Then
                .Fields("MailAddrNo") = rsSo001("MailAddrNo")
                .Fields("MailAddress") = rsSo001("MailAddress")
            End If
        Else
            If rsSo001("InstAddrNo") > 0 And Not IsNull(rsSo001("InstAddress")) Then
                .Fields("ChargeAddrNo") = rsSo001("InstAddrNo")
                .Fields("ChargeAddress") = rsSo001("InstAddress") & ""
                .Fields("MailAddrNo") = rsSo001("InstAddrNo")
                .Fields("MailAddress") = rsSo001("InstAddress") & ""
            End If
        End If
        '*********************************************************************************
        If Not IsNull(rsSo001("InvoiceType")) Then
            Select Case rsSo001("InvoiceType") & ""
                Case "0"
                    .Fields("InvoiceType") = 0
                Case "2"
                    .Fields("InvoiceType") = 2
                '**************************************************************************************************
                '#4210 如果是3聯式發票,發票抬頭要帶SO002.InvTitle By Kin 2008/11/06
                Case "3"
                    .Fields("InvoiceType") = 3
                    .Fields("InvTitle") = IIf(IsNull(rsSo001("InvTitle")), txtAccountOwner.Text, rsSo001("InvTitle") & "")
                '***************************************************************************************************
                Case Else
                    .Fields("InvoiceType") = 2
            End Select
        End If
        If Not IsNull(rsSo001("InvNo")) Then
            .Fields("InvNo") = rsSo001("InvNo") & ""
            '如果統一編號有值一定要設成 電子計算機發票 By Kin 2010/08/06
            blnElectron = False
        End If
        .Fields("StopFlag") = 0
        .Fields("Updtime") = GetDTString(RightNow)
        .Fields("NEWUPDTIME") = RightNow
        .Fields("UpdEn") = garyGi(1)
        strInvSeqNo = .Fields("InvSeqNo")
        If blnElectron Then
            .Fields("INVOICEKIND") = 1
        Else
            .Fields("INVOICEKIND") = 0
        End If
        '#5945 發票用途、發票種類、發票開立種類要用SO002的資訊 By Kin 2011/03/24
        '#5945 測試不OK,受贈單位、受贈日期也要Copy By Kin 2011/04/20
        If Not rsSo002.EOF Then
            If Len(rsSo002("InvoiceType") & "") > 0 Then
                .Fields("InvoiceType") = rsSo002("InvoiceType") & ""
            End If
            
            If Len(rsSo002("InvoiceKind") & "") > 0 Then
                .Fields("InvoiceKind") = rsSo002("InvoiceKind") & ""
            End If
            
             '#6049 發票用途不要更新 By Kin 2011/06/17
'            If Len(rsSo002("InvPurposeCode") & "") > 0 Then
'                .Fields("InvPurposeCode") = rsSo002("InvPurposeCode") & ""
'                .Fields("InvPurposeName") = rsSo002("InvPurposeName") & ""
'            End If
            
            If Len(rsSo002("InvNo") & "") > 0 Then
                .Fields("InvoiceKind") = 0
            End If
            
            '#6049 受贈單位不要更新 By Kin 2011/06/17
'            If Len(rsSo002("DenRecCode") & "") > 0 Then
'                .Fields("DenRecCode") = rsSo002("DenRecCode") & ""
'            End If
'
'            If Len(rsSo002("DenRecName") & "") > 0 Then
'                .Fields("DenRecName") = rsSo002("DenRecName") & ""
'            End If
            '#6049 受贈日期不要更新 By Kin 2011/06/17
'            If Len(rsSo002("DenRecDate") & "") > 0 Then
'                .Fields("DenRecDate") = rsSo002("DenRecDate") & ""
'            End If
'            #6049 發票抬頭以上面畫面輸入的為主 By Kin 2011/06/17
'            If Len(rsSo002("InvTitle") & "") > 0 Then
'                .Fields("InvTitle") = rsSo002("InvTitle") & ""
'            End If
        End If
         '#6438 自動帶入(1) 發票用途、(2) 受贈單位、(3) 受贈日期、(4) 申請電子記算機發票日期 By Kin 2013/03/19
         '#6438 判斷申請人與SO001是否相同,相同才帶入資料 By Kin 2013/04/22
        If cboProposer.Text = rsSo001("CustName") Then
            If Len(rsSo002("InvPurposeCode") & "") > 0 Then
                .Fields("InvPurposeCode") = rsSo002("InvPurposeCode") & ""
                .Fields("InvPurposeName") = rsSo002("InvPurposeName") & ""
            End If
            If Len(rsSo002("DenRecCode") & "") > 0 Then
               .Fields("DenRecCode") = rsSo002("DenRecCode") & ""
               .Fields("DenRecName") = rsSo002("DenRecName") & ""
            End If
            If Len(rsSo002("DenRecDate") & "") > 0 Then
                .Fields("DenRecDate") = rsSo002("DenRecDate") & ""
            End If
            If Len(rsSo002("ApplyInvDate") & "") > 0 Then
               .Fields("ApplyInvDate") = rsSo002("ApplyInvDate") & ""
            End If
        End If
        
        .Update
    End With
    
    On Error Resume Next
    Call CloseRecordset(rsSo001)
    Call CloseRecordset(rsSO138Tmp)
    Call CloseRecordset(rsSo002)
    Exit Sub
ChkErr:
  ErrSub Me.Name, "InsSO138"
End Sub
'#5668 選擇週期性項目判斷是否要設定為電子發票 By Kin 2010/07/19
Private Function SetInvKind(ByRef aMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim aCitemStr As String
    Dim aryCitemStr() As String
    Dim i As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim aTmp As String
    Dim aSQL As String
    Dim iCntRct As Long
    Dim iType As Integer
    Dim strTmpCitem As String
    Dim blnExitFor As Boolean
    aCitemStr = csmCitem1.GetQryFld(1)
    If aCitemStr = "" Then Exit Function
    aryCitemStr = Split(aCitemStr, ",")
    iCntRct = 0
    strTmpCitem = Empty
    For i = LBound(aryCitemStr) To UBound(aryCitemStr)
        aTmp = aryCitemStr(i)
        iType = GetByHouse(aTmp)
        If iType = 0 Then
            aSQL = "SELECT Contmobile,Email   FROM " & GetOwner & "SO004 " & _
                " WHERE (Contmobile IS NOT NULL OR EMail IS NOT NULL )" & _
                " AND CUSTID = " & gCustId & " AND COMPCODE = " & gCompCode & _
                " AND FACISNO IN ( SELECT FACISNO FROM " & GetOwner & "SO003 " & _
                " WHERE CITEMCODE = " & aTmp & " AND FACISNO IS NOT NULL " & _
                " AND CUSTID = " & gCustId & ")"
            
        Else
            aSQL = " SELECT  Null Contmobile,Email  FROM " & GetOwner & "SO002 " & _
                " WHERE CUSTID = " & gCustId & " AND COMPCODE = " & gCompCode & _
                " AND SERVICETYPE IN ( SELECT SERVICETYPE FROM " & GetOwner & "SO003 " & _
                " WHERE CUSTID = " & gCustId & " AND COMPCODE = " & gCompCode & _
                " AND CITEMCODE = " & aTmp & ") " & _
                " AND EMAIL IS NOT NULL " & _
                " UNION ALL " & _
                " SELECT Tel3 Contmobile,Null Email  FROM " & GetOwner & "SO001 " & _
                " WHERE CUSTID = " & gCustId & " AND COMPCODE = " & gCompCode & _
                " AND Tel3 IS NOT NULL "
            aSQL = "Select Distinct * From ( " & aSQL & ") "
        End If
        If Not GetRS(rsTmp, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        Do While Not rsTmp.EOF
             
            If ChkMobileOk(rsTmp("Contmobile") & "") Or ChkEmailOk(rsTmp("Email") & "") Then
                blnExitFor = True
                Exit Do
            End If
            rsTmp.MoveNext
        Loop
        If blnExitFor Then
            iCntRct = UBound(aryCitemStr) + 1
            Exit For
        End If
'        If Val(GetRsValue(aSQL, gcnGi) & "") > 0 Then
'            iCntRct = iCntRct + 1
'        Else
'            If strTmpCitem = Empty Then
'                strTmpCitem = aTmp
'            Else
'                strTmpCitem = strTmpCitem & "，" & aTmp
'            End If
'        End If
        
    Next i
    If UBound(aryCitemStr) = 0 Then
        SetInvKind = iCntRct = 1
        aMsg = Empty
    Else
        If UBound(aryCitemStr) + 1 <> iCntRct Then
            aMsg = "收費項目：" & strTmpCitem & vbCrLf & _
                " 無手機、電子郵件資訊，是否設定為電子發票？"
            SetInvKind = False
        Else
            aMsg = Empty
            SetInvKind = True
        End If
        
    End If
    aMsg = Empty
    On Error Resume Next
    Call Close3Recordset(rsTmp)
    Exit Function
ChkErr:
  ErrSub Me.Name, "SetInvKind"
End Function
Private Function Get138SeqNo() As String
  On Error GoTo ChkErr
    Get138SeqNo = RPxx(gcnGi.Execute("SELECT " & GetOwner & "S_SO138_InvSeqNo.NEXTVAL FROM DUAL").GetString & "")
  Exit Function
ChkErr:
    ErrSub Me.Name, "GetSeqNo"
End Function

Private Sub chkFore2_Click()
  On Error Resume Next
    If chkFore2.Value Then
        txtAccOwnerID.MaxLength = 30
    Else
        txtAccOwnerID.MaxLength = 10
        txtAccOwnerID = Left(txtAccOwnerID, 10)
    End If
End Sub

Private Sub chkStop_Click()
  On Error Resume Next
    lblStopDate.ForeColor = IIf(chkStop.Value = 1, vbRed, vbBlack)
    If chkStop.Value <> 1 Then
        gdaStopDate.Text = ""
    Else
        gdaStopDate.Text = Date
        gdaStopDate.SetFocus
    End If
End Sub

Private Sub cmdACHDetail_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
    '*****************************************************************************************
    '#3946 MasterRowId 改為 MasterId By Kin 2008/05/30
    If EditMode = giEditModeInsert Then
        frmSO1100G1.uMasterRowId = 0
    Else
        frmSO1100G1.uMasterRowId = IIf(IsNull(rsSO106("MasterId")), 0, rsSO106("MasterId"))
    End If
    '*****************************************************************************************
    If blnCallSO106A Then
        frmSO1100G1.EditMode = giEditModeEdit
    Else
        frmSO1100G1.EditMode = giEditModeView
    End If
    frmSO1100G1.uInAchId = strACHID
    frmSO1100G1.uInAchDesc = strACHDesc
    frmSO1100G1.uFirstACHDesc = strFirstACHDesc
    frmSO1100G1.uFirstAchId = strFirstACHID
    frmSO1100G1.Show 1
    strACHID = Empty
    strACHDesc = Empty
    strFirstACHID = Empty
    strFirstACHDesc = Empty
    blnCallSO106A = False
    Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdACHDetail"
End Sub

Private Sub cmdCreateINV_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
    With frmSO114FA
        .OldEditMode = Me.EditMode
        .uCustId = gCustId
        If EditMode = giEditModeEdit Then
            .uAccountNo = Trim(txtAccountNo.Text)
            If Not IsNull(rsSO106("InvSeqNo")) Then .uInvSeqNo = rsSO106("InvSeqNo") & ""
            If Not IsNull(rsSO106("InvSeqNo")) Then .uShowAbs = rsSO106("InvSeqNo") & ""
        Else
            If txtAccountNo.Text <> "" Then
                .uAccountNo = Trim(txtAccountNo.Text)
            End If
            .uInvSeqNo = Empty
        End If
        .uAccountOwner = txtAccountOwner.Text
        .uMutilChoice = False
        .uOwnerForm = Me
        .Show , Me
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
  ErrSub Me.Name, "cmdCreateINV"
End Sub

Private Sub cmdFind_Click()
  On Error GoTo ChkErr
    ServiceCommon.FindIntroduce Me
  Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdFind_Click"
End Sub

Private Sub cmdInvSeqNo_Click()
  On Error GoTo ChkErr
    With frmSO1131H
        .uParentForm = Me
        .uMutilChoice = False
        .Show 1
    End With
    strInvSeqNo = Empty
    strInvSeqNo = frmSO1131H.uNewInvSeqNo
   
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdInvSeqNo"
End Sub

'#3236 ACH改變時如果收費項目有值需Show出警告訊息,所以要一個變數記錄原來的值 By Kin 2007/05/29
Private Sub csmACH_ButtonClick()
    On Error Resume Next
    strAchChange = csmACH.GetDispStr
    blnAchChgClick = True
End Sub

Private Sub csmACH_Change()
    On Error GoTo ChkErr
    '#3322 一開始新增時會出現警告訊息，多加一個Flag做判斷，如果一進來是新增先不判斷，直接跳出，但要馬上設為False，因為後續還是要判斷 By Kin 2007/07/04
    If blnAddFlag Then
        blnAddFlag = False
        Exit Sub
    End If
    If Not blnNoShowMsg Then
        If (csmACH.GetDispStr & "" = "") And (csmCitem1.GetDispStr & "" <> "") And (EditMode <> giEditModeView) Then
            MsgBox "ACH授權交易別已改變!!請重新設定收費項目", vbInformation, "警告訊息"
            csmCitem1.Clear
        End If
    End If
    blnNoShowMsg = False
    Exit Sub
ChkErr:
     ErrSub Name, "csmACH_Change"
End Sub

'#3236 ACH改變時如果收費項目有值需Show出警告訊息By Kin 2007/05/29
Private Sub csmACH_SelectChange()
    On Error GoTo ChkErr
    '有經過ButtonClick事件
    If blnAchChgClick Then
        
        If (csmACH.GetDispStr <> strAchChange) And (csmCitem1.GetDispStr <> "") Then
            MsgBox "ACH授權交易別已改變!!請重新設定收費項目", vbInformation, "警告訊息"
            csmCitem1.Clear
        End If
    Else 'User直接按下Clear按鈕
'        If lngEditMode <> giEditModeView Then
'            If (csmACH.GetDispStr & "" = "") And (csmCitem1.GetDispStr & "" <> "") And (rsSO106("ACHTNo") & "" <> "") Then
'                MsgBox "ACH授權交易別已改變!!請重新設定收費項目", vbInformation, "警告訊息"
'                csmCitem1.Clear
'            End If
'        End If
    End If
    blnAchChgClick = False
    blnNoShowMsg = False
    Exit Sub
ChkErr:
    ErrSub Name, "csmACH_SelectChange"
End Sub

Private Sub csmCitem1_ButtonClick()
  On Error GoTo ChkErr
    Dim strTmp As String
    strTmp = RPxx(Replace(csmACH.GetQryFld(3), "'", "", 1))
    '#7049 將CitemCode串到CD068A By Kin 2015/07/08
    strTmp = "Select CD068A.CitemCode From " & GetOwner & "CD068A Where BillHeadFmt In (" & csmACH.GetQryFld(4) & ")"
    '#3454 多串SO003.FaciSNo、SO004.DeclarantName、SO004.FaciName By Kin 2007/12/18
    If csmACH.GetDispStr <> Empty Then
        '#6577 增加StopFlag By Kin 2013/09/04
        SetMsQry csmCitem1, "SELECT " & _
                            " A.CITEMCODE,A.CITEMNAME,  Decode(Nvl(A.StopFlag,0),0,'否','是') StopFlag, " & _
                            " A.PERIOD,A.AMOUNT,A.ACCOUNTNO,A.CMNAME,A.STARTDATE," & _
                            "A.STOPDATE,A.SEQNO,A.FaciSNo,B.DeclarantName,B.FaciName FROM " & _
                             GetOwner & "SO003 A," & GetOwner & "SO004 B  WHERE A.CUSTID = " & _
                             gCustId & " AND A.COMPCODE=" & gCompCode & _
                             " AND A.CITEMCODE IN (" & strTmp & ")" & _
                             " AND A.CustId=B.CustId(+) And A.CompCode=B.CompCode(+)" & _
                             " And A.FaciSeqNo=B.SeqNo(+)" & _
                             " ORDER BY A.CLCTDATE,A.CITEMCODE", _
                             , , "代碼,收費項目,停用,期數,金額,扣帳帳號,收費方式,起始日,截止日,序號,物品序號,申請人姓名,項目名稱", _
                             "1000,1800,500,460,560,1800,850,850,850,770,850,1600,1600", True, , , , , _
                             , , , , , , , , , , , , , , , , , , , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD", , , , 10
    Else
        '#6577 增加StopFlag By Kin 2013/09/04
        SetMsQry csmCitem1, "SELECT " & _
                            "A.CITEMCODE,A.CITEMNAME,Decode(Nvl(A.StopFlag,0),0,'否','是') StopFlag," & _
                            " A.PERIOD,A.AMOUNT,A.ACCOUNTNO,A.CMNAME,A.STARTDATE," & _
                            "A.STOPDATE,A.SEQNO,A.FaciSNo,B.DeclarantName,B.FaciName FROM " & _
                             GetOwner & "SO003 A," & GetOwner & "SO004 B  WHERE A.CUSTID = " & _
                             gCustId & " AND A.COMPCODE=" & gCompCode & _
                             " AND A.CustId=B.CustId(+) And A.CompCode=B.CompCode(+)" & _
                             " And A.FaciSeqNo=B.SeqNo(+)" & _
                             " ORDER BY A.CLCTDATE,A.CITEMCODE", _
                             , , "代碼,收費項目,停用,期數,金額,扣帳帳號,收費方式,起始日,截止日,序號,物品序號,申請人姓名,項目名稱", _
                             "1000,1800,500,460,560,1800,850,850,850,770,850,1600,1600", True, , , , , _
                             , , , , , , , , , , , , , , , , , , , , , , "###,###,###", , "EE/MM/DD", "EE/MM/DD", , , , , 10



    End If
  Exit Sub
ChkErr:
    ErrSub Name, "csmCitem1_ButtonClick"
End Sub

Private Sub csmCitem2_ButtonClick()
  On Error GoTo ChkErr
    Dim strTmp As String
    Dim strQry As String
    Dim strOrder As String
    
    '#3454 多串SO003.FaciSNo、SO004.DeclarantName、SO004.FaciName By Kin 2007/12/10
    '#5045 用參數判斷是否打挑選已收費的資料 By Kin 2009/05/04
    
    strOrder = " ORDER BY A.REALDATE DESC,A.BILLNO DESC,A.SEQNO DESC "
    strTmp = RPxx(Replace(csmACH.GetQryFld(3), "'", "", 1))
    '#7049 增加BillHeadFmt欄位,可以串到CD068A By Kin 2015/07/08
    strTmp = "Select CD068A.CitemCode From " & GetOwner & "CD068A Where BillHeadFmt In (" & csmACH.GetQryFld(4) & ")"
    If csmACH.GetDispStr <> Empty Then
        strQry = "SELECT DISTINCT CITEMCODE,CITEMNAME,REALPERIOD,SHOULDAMT,ACCOUNTNO,CMNAME,REALSTARTDATE," & _
                            "REALSTOPDATE,SEQNO,PK,RowID,FaciSNo,DeclarantName,FaciName FROM " & _
                             "(SELECT A.CITEMCODE,A.CITEMNAME,A.REALPERIOD,A.SHOULDAMT,A.ACCOUNTNO,A.CMNAME,A.REALSTARTDATE," & _
                                "A.REALSTOPDATE,A.SEQNO,A.BILLNO || A.ITEM PK,A.ROWID,A.FaciSNo,B.DeclarantName,B.FaciName FROM " & _
                             GetOwner & "SO033 A," & GetOwner & "SO004 B WHERE A.CUSTID = " & gCustId & _
                             " AND A.COMPCODE=" & gCompCode & _
                             " AND A.CANCELFLAG <> 1  AND A.CHEVEN <> 1 " & _
                             " AND A.CITEMCODE IN (" & strTmp & ") " & _
                             " AND A.CustId=B.CustId(+) And A.CompCode=B.CompCode(+)" & _
                             " And A.FaciSeqNo=B.SeqNo(+) "
                             
        '#5915 增加過濾 PAYOK=1 與 REFNO IN (3,7) By Kin 2011/03/08
        '#5999 HItemAutoUpd就算開啟,取出已收部份也不要再過濾PayOK=1 By Kin 2011/04/28
        strQry = strQry & " AND A.UCCODE NOT IN(SELECT CODENO FROM " & GetOwner & "CD013 " & _
                    " WHERE PAYOK = 1 OR REFNO IN (3,7)) "
        
        strQry = strQry & " And A.UCCode Is Not Null " & strOrder & ")" & _
               IIf(blnUCCodeIsNull, " Union " & strQry & " And " & _
            " A.UCCode Is Null  " & _
            " And A.INVSEQNO IS NULL " & strOrder & " )", "")

        SetMsQry csmCitem2, strQry, _
                             , , "代碼,收費項目,期數,金額,扣帳帳號,收費方式,起始日,截止日,序號,PK,RowID,物品序號,申請人姓名,項目名稱", _
                             "500,1800,460,560,1800,850,850,850,770,1,1,850,1600,1600", True, , , , , _
                             , , , , , , , , , , , , , , , , , , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD", , , , , 10

    Else
        strQry = "SELECT DISTINCT CITEMCODE,CITEMNAME,REALPERIOD,SHOULDAMT,ACCOUNTNO,CMNAME,REALSTARTDATE," & _
                            "REALSTOPDATE,SEQNO,PK,RowID,FaciSNo,DeclarantName,FaciName FROM " & _
                             "(SELECT A.CITEMCODE,A.CITEMNAME,A.REALPERIOD,A.SHOULDAMT,A.ACCOUNTNO,A.CMNAME,A.REALSTARTDATE," & _
                                "A.REALSTOPDATE,A.SEQNO,A.BILLNO || A.ITEM PK,A.ROWID,A.FaciSNo,B.DeclarantName,B.FaciName FROM " & _
                             GetOwner & "SO033 A," & GetOwner & "SO004 B " & _
                             " WHERE A.CUSTID = " & gCustId & _
                             " AND A.COMPCODE=" & gCompCode & _
                             " AND A.CANCELFLAG <> 1  AND A.CHEVEN <> 1 " & _
                             " AND A.CustId=B.CustId(+) And A.CompCode=B.CompCode(+)" & _
                             " And A.FaciSeqNo=B.SeqNo(+) "
                             
        '#5915 增加過濾 PAYOK=1 與 REFNO IN (3,7) By Kin 2011/03/08
        '#5999 HItemAutoUpd就算開啟,取出已收部份也不要再過濾PayOK=1 By Kin 2011/04/28
        Dim strQry2 As String
        strQry2 = strQry & " AND A.UCCODE NOT IN(SELECT CODENO FROM " & GetOwner & "CD013 " & _
                    " WHERE PAYOK = 1 OR REFNO IN (3,7)) "
        
                                             
        strQry = strQry2 & " And A.UCCode Is Not Null " & strOrder & ")" & _
                       IIf(blnUCCodeIsNull, " Union " & strQry & " And " & _
                    " A.UCCode Is Null  " & _
                    " And A.INVSEQNO IS NULL " & strOrder & " )", "")
        
    
        SetMsQry csmCitem2, strQry, _
                             , , "代碼,收費項目,期數,金額,扣帳帳號,收費方式,起始日,截止日,序號,PK,RowID,物品序號,申請人姓名,項目名稱", _
                             "500,1800,460,560,1800,850,850,850,770,1,1,850,1600,1600", True, , , , , _
                             , , , , , , , , , , , , , , , , , , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD", , , , , 10
                
    End If
  Exit Sub
ChkErr:
    ErrSub Name, "csmCitem2_ButtonClick"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    ReleaseCOM Me
End Sub

Private Sub RcdToScr()
  On Error GoTo ChkErr
    strACHID = Empty
    strACHDesc = Empty
    With rsSO106
         gdtAcceptTime.SetValue .Fields("AcceptTime") & ""
         '#6215 使用 Query_CodeNo會不正確改成 Query_Description By Kin 2012/02/22
         gilAcceptName.SetCodeNo .Fields("AcceptEn") & ""
         gilAcceptName.Query_Description
'         gilAcceptName.SetDescription .Fields("AcceptName") & ""
'         gilAcceptName.Query_CodeNo
         'txtProposer.Text = .Fields("Proposer") & ""
         '#3454 申請人改用ComboBox By Kin 2007/12/18
         cboProposer.Text = .Fields("Proposer") & ""
         If GetUserPriv("SO1100G", "(SO1100G0)") Then
             txtID.Text = .Fields("ID") & ""
             txtAccOwnerID = .Fields("AccountNameID") & ""
         Else
             txtID.Tag = .Fields("ID") & ""
             If txtID.Tag = Empty Then
                 txtID.Text = ""
             Else
                 txtID.Text = Left(txtID.Tag, 3) & "XXXX" & Right(txtID.Tag, 3)
             End If
             txtAccOwnerID.Tag = .Fields("AccountNameID") & ""
             If txtAccOwnerID.Tag = Empty Then
                 txtAccOwnerID.Text = ""
             Else
                 txtAccOwnerID.Text = Left(txtAccOwnerID.Tag, 3) & "XXXX" & Right(txtAccOwnerID.Tag, 3)
             End If
         End If
         
         gilBankCode.SetCodeNo .Fields("BankCode") & ""
         gilBankCode.SetDescription .Fields("BankName") & ""
         gilCardName.SetDescription .Fields("CardName") & ""
         gilCardName.Query_CodeNo
         mskCardExpDate.Text = .Fields("StopYM") & ""
         If Len(.Fields("StopYM")) > 1 Then
            mskCardExpDate.Text = Right("0" & (.Fields("StopYM") & ""), 6)
            txtHide.Text = mskCardExpDate.Text
         Else
            mskCardExpDate.Text = ""
            txtHide.Text = ""
         End If
         
         If GetUserPriv("SO1100G", "(SO1100G9)") Then
            txtAccountNo.Text = .Fields("AccountID") & ""
         Else
            txtAccountNo.Tag = .Fields("AccountID") & ""
            If txtAccountNo.Tag = Empty Then
                txtAccountNo.Text = ""
            Else
                txtAccountNo.Text = Left(txtAccountNo.Tag, 5) & "XXX" & Mid(txtAccountNo.Tag, 9, 2) & "XX" & Right(txtAccountNo.Tag, 4)
            End If
         End If
          If (Len(.Fields("BankCode") & "") > 0) Then
                  If GetRsValue("Select count(*) from cd018 Where CD018.PRGNAME like '%POST%' and CodeNo = " & .Fields("BankCode"), gcnGi) = 0 Then
                      AchTypeCondition = " In ( 1 ) "
                  Else
                      AchTypeCondition = " In ( 2 ) "
                  End If
         End If
         If Not blnStartPost Then
            AchTypeCondition = " In ( 1 ) "
         End If
            
        '#3236 ACHTNO代碼沒有唯一,顯示出來都會變成第一個代碼,SO106新增ACHTDESC欄位,SetQueryCode改用ACHDESC欄位 By Kin 2007/05/28
         If (Not IsNull(.Fields("ACHTNO"))) And (Not IsNull(.Fields("ACHTDESC"))) Then
             'csmACH.SetQueryCode .Fields("ACHTNO") & ""
             '********************************************************************************************************************************************************
             '#3728 只Show出有挑選的ACH代碼,如果為空值全部Show出 By Kin 2008/03/03
             '#3946 送件日期為空值的話ACH代碼還是要能選擇 By Kin 2008/05/23
             '#4030 不過濾停用
             '#4106 增加判斷ACHType參數 By Kin 2008/09/22
             '#7049 增加BillHeadFmt欄位,可以串到CD068A By Kin 2015/07/08
           
            If Not IsNull(.Fields("SendDate")) Then
               
                SetMsQry csmACH, "SELECT ACHTNO,ACHTDESC,CITEMCODESTR,BillHeadFmt FROM " & GetOwner & "CD068  " & _
                                            " Where ACHTDESC in(" & .Fields("ACHTDESC") & ") And ACHTDESC is NOT NULL And ACHType " & AchTypeCondition, _
                                       "ACH交易代碼", "ACH交易說明", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
            Else
                SetMsQry csmACH, "SELECT ACHTNO,ACHTDESC,CITEMCODESTR,BillHeadFmt FROM " & GetOwner & "CD068  " & _
                                            " Where  ACHTNO Is Not Null And ACHTDESC is NOT NULL And ACHType  " & AchTypeCondition, _
                                    "ACH交易代碼", "ACH交易說明", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
            End If
            
            csmACH.SetQueryCode .Fields("ACHTDESC") & ""
            
            
         Else
             SetMsQry csmACH, "SELECT ACHTNO,ACHTDESC,CITEMCODESTR,BillHeadFmt FROM " & GetOwner & "CD068  " & _
                                        " Where  ACHTNO is Not Null And ACHTDESC is Not NULL And ACHType " & AchTypeCondition, _
                                    "ACH交易代碼", "ACH交易說明", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
             csmACH.Clear
         End If
             '*********************************************************************************************************************************************************
         txtCVC2 = .Fields("CVC2") & ""
         
         If Not IsNull(.Fields("CITEMSTR")) Then
            csmCitem1_ButtonClick
            '#3236 將 csmCitem1.DispStr清空,防止無資料時會停留上一次未清除的畫面 By Kin 2007/05/28
             csmCitem1.SetDispStr ""
             csmCitem1.SetQueryCode .Fields("CITEMSTR") & ""
             csmCitem1.SetQueryCode .Fields("CITEMSTR") & ""
         Else
             csmCitem1.Clear
         End If
         If Not IsNull(.Fields("CITEMSTR2")) Then
            csmCitem2_ButtonClick
            csmCitem2.SetDispStr ""
            csmCitem2.SetQueryCode .Fields("CITEMSTR2") & ""
            csmCitem2.SetQueryCode .Fields("CITEMSTR2") & ""
         Else
             csmCitem2.Clear
         End If
         
         txtAccountOwner = .Fields("AccountName") & ""
         gdaContiDate.Text = .Fields("PropDate") & ""
         gdaSendDate.Text = .Fields("SendDate") & ""
         
         gdaAuthStopDate.Text = .Fields("AuthorizeStopDate") & ""
         
         gdaSnactionDate.Text = .Fields("SnactionDate") & ""
         chkStop.Value = Val(.Fields("StopFlag") & "")
         gdaStopDate.Text = .Fields("StopDate") & ""
         gilMediaCode.SetCodeNo .Fields("MediaCode") & ""
         gilMediaCode.SetDescription .Fields("MediaName") & ""
         txtIntroId.Text = .Fields("IntroID") & ""
         lblIntroName.Caption = .Fields("IntroName") & ""
         gilCMCode.SetCodeNo .Fields("CMcode") & ""
         gilCMCode.SetDescription .Fields("CMname") & ""
         gilPTCode.SetCodeNo .Fields("PTcode") & ""
         gilPTCode.SetDescription .Fields("PTname") & ""
         lblUpdTime.Caption = .Fields("Updtime").Value & ""
         lblUpdEn.Caption = .Fields("UpdEn").Value & ""
         txtNote.Text = .Fields("Note") & ""
         chkFore.Value = IIf(IsNull(.Fields("Alien")), 0, .Fields("Alien"))
         chkFore2.Value = IIf(IsNull(.Fields("AccountAlien")), 0, .Fields("AccountAlien"))
         chkAddAcc.Value = IIf(IsNull(.Fields("AddCitemAccount")), 0, .Fields("AddCitemAccount"))
         ' ACH 部份 93/03/02 Hammer
         Select Case .Fields("AuthorizeStatus") & ""
                Case 1
                      lblStatus = "已授權成功"
                Case 2
                      lblStatus = "已取消授權"
                Case Else
                      lblStatus = ""
         End Select
         txtACHcustID = .Fields("ACHCustID") & ""
         chkDeAut.Value = IIf(IsNull(.Fields("DeAuthorize")), 0, .Fields("DeAuthorize"))
         
         chkToACH.Value = IIf(IsNull(.Fields("ToACH")), 0, .Fields("ToACH"))
         txtAchSN = .Fields("ACHSN") & ""
         strInvSeqNo = .Fields("InvSeqNo") & ""
         lblInvSeqNo = strInvSeqNo
         
    End With
    '個資法
    If Len(strAccountName) > 0 Then
        If strAccountName <> rsSO106.Fields("Proposer") Then
            cboProposer.Text = "XXXXX"
            gilCMCode.SetCodeNo "X"
            gilCMCode.SetDescription "XXX"
            gilPTCode.SetCodeNo "X"
            gilPTCode.SetDescription "XXX"
            gilBankCode.SetCodeNo "X"
            gilBankCode.SetDescription "XXX"
            txtAccountNo.PasswordChar = "X"
            txtAccountOwner.PasswordChar = "X"
            txtHide.PasswordChar = "X"
            
        End If
    End If

  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScr")
End Sub

Public Sub DeleteGo()
  On Error GoTo ChkErr
    '#3728 刪除SO106時一併刪除SO106A對應的資料 By Kin 2008/03/10
    
   
    With frmSO1100BMDI
        If Screen.ActiveForm.Name = "frmSO1100G" Then
            If MsgBox("確定刪除 [帳號/卡號] 為 " & rsSO106!ACCOUNTID & " 的信用卡轉帳資料嗎 ?", vbQuestion + vbYesNo, "詢問") = vbYes Then
                '#5946 增加記錄Log檔 By Kin 2011/03/16
                EditMode = giEditModeDelete
                Call InsSO106Log
                EditMode = giEditModeView
                StopGo CStr(rsSO106!ACCOUNTID)
                gcnGi.Execute "DELETE FROM " & GetOwner & "SO106 WHERE ROWID='" & .rsSO106!RowId & "'"
                '#3946 改用MasterId為PK By Kin 2008/05/30
                gcnGi.Execute "DELETE FROM " & GetOwner & "SO106A WHERE MasterId='" & .rsSO106("MasterId") & "'"
                OpenData
               .RefreshFormData
            End If
            If rsSO106.RecordCount = 0 Then NewRcd: UserPermissionGo
        Else
            If MsgBox("確定刪除 [帳號/卡號] 為 " & .rsSO106!ACCOUNTID & " 的信用卡轉帳資料嗎 ?", vbQuestion + vbYesNo, "詢問") = vbYes Then
                '#5946 增加記錄Log檔 By Kin 2011/03/16
                EditMode = giEditModeDelete
                Call InsSO106Log
                EditMode = giEditModeView
                StopGo CStr(.rsSO106!ACCOUNTID)
                gcnGi.Execute "DELETE FROM " & GetOwner & "SO106 WHERE ROWID='" & .rsSO106!RowId & "'"
                '#3946 改用MasterId為PK By Kin 2008/05/30
                gcnGi.Execute "DELETE FROM " & GetOwner & "SO106A WHERE MasterId='" & .rsSO106("MasterId") & "'"
               .RefreshFormData
            End If
            If .rsSO106.RecordCount = 0 Then UserPermissionGo
            Unload Me
        End If
    End With
    
  Exit Sub
ChkErr:
    ErrSub Me.Name, "DeleteGo"
End Sub

Public Function SmartAdd() As Boolean
  On Error GoTo ChkErr
  Exit Function
ChkErr:
    ErrSub Name, "SmartAdd"
End Function
'#3585 測試不OK(RA規格重新修訂的) 取Ach最大值 By Kin 2007/10/30
'#3861 改取SO041.ACHHeadCode的值,為了避免User沒設定,所以原本取值的部份不拿掉,如果SO041沒設定的話,一樣取ACH的最大值 By Kin 2008/04/24
'#4006 編碼原則要取最大值 By Kin 2008/07/21
Private Function GetMaxAchNo(ByVal strMax As String) As String
  On Error GoTo ChkErr
    Dim arrVar As Variant
    Dim strAchMax As String
    Dim i As Long
    Dim rsACHCode As New ADODB.Recordset
    If blnACHCustID Then
        If Not GetRS(rsACHCode, "Select ACHHeadCode From " & GetOwner & "SO041 Where SysId=" & gCompCode, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        If (Not IsNull(rsACHCode("ACHHeadCode"))) And (rsACHCode("ACHHeadCode") & "" <> "") Then
            GetMaxAchNo = rsACHCode("ACHHeadCode") & ""
        Else
            If Len(strMax) = 0 Then
                GetMaxAchNo = Empty
            Else
                arrVar = Split(Replace(strMax, "'", ""), ",")
                strAchMax = arrVar(0)
                For i = 1 To UBound(arrVar)
                    If Val(arrVar(i)) > Val(strAchMax) Then strAchMax = arrVar(i)
                Next i
                GetMaxAchNo = strAchMax
            End If
        End If
    Else
        If Len(strMax) = 0 Then
             GetMaxAchNo = Empty
        Else
            arrVar = Split(Replace(strMax, "'", ""), ",")
            strAchMax = arrVar(0)
            For i = 1 To UBound(arrVar)
                If Val(arrVar(i)) > Val(strAchMax) Then strAchMax = arrVar(i)
            Next i
            GetMaxAchNo = strAchMax
             
        End If
    End If
    On Error Resume Next
    Call CloseRecordset(rsACHCode)
  Exit Function
ChkErr:
    ErrSub Name, "GetMaxAchNo"
End Function
Private Function ScrToRcd() As Boolean
  On Error GoTo ChkErr
    Dim blnUpd As Boolean
    Dim strSqlCid As String, strAchCid As String
    Dim rsAchCid As New ADODB.Recordset
    Dim arySO106A() As String
    blnCallSO106A = False
    strACHID = Empty
    strACHDesc = Empty
    strFirstACHDesc = Empty
    strFirstACHID = Empty
    Dim i As Long
    Dim rsSO106A As New ADODB.Recordset
    blnUpd = False
    
    '*******************************************************************
    '#4130 在此增加檢核,當InvSeqNo=0時出現警告訊息 By Kin 2008/10/21
    If Val(strInvSeqNo) <= 0 Then
        MsgBox "發票流水號不正確！", vbInformation, "警告"
        Exit Function
    End If
    '*******************************************************************
    With rsSO106
         If EditMode = giEditModeInsert Then
            .AddNew
            .Fields("InheritFlag") = 0
            .Fields("InheritKey") = ""
            
            '#3436 儲存InvSeqNo資訊 By Kin 2007/12/13
            '.Fields("InvSeqNo") = strInvSeqNo
         Else
            
            If chkStop.Value = 1 Then
                .Fields("InheritFlag") = 0
                .Fields("InheritKey") = ""
                
                '********************************************************************
                '#3728 如果停用,則將產生取消授權資料 By Kin 2008/03/18
                If blnACHCustID Then
                    If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018" & _
                                    " WHERE CODENO=" & gilBankCode.GetCodeNo & _
                                    " AND ( PRGNAME LIKE 'ACH%' Or " & posAchContition & " ) " & _
                                    " AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
                        blnNoShowMsg = True
                        csmACH.Clear
                        csmCitem1.Clear
                        
                    End If
                End If
                '*********************************************************************
            Else
                If .Fields("InheritFlag") = 0 And .Fields("InheritKey") <> "" Then blnUpd = True
                
                '#3436 儲存InvSeqNo資訊 By Kin 2007/12/13
'                If strInvSeqNo <> "" Then
'                    .Fields("InvSeqNo") = strInvSeqNo
'                End If
            End If
         End If
         '*****************************************************
         '#3946 增加此欄位,要與SO106A對應 By Kin 2008/05/30
         If Len(.Fields("MasterId") & "") = 0 Then
            .Fields("MasterId") = Get_MasterID_Seq
         End If
         '******************************************************
        .Fields("InvSeqNo") = strInvSeqNo
        .Fields("COMPCODE") = gCompCode & ""
        .Fields("CustID") = gCustId
        .Fields("AcceptName") = gilAcceptName.GetDescription2
        .Fields("AcceptEn") = gilAcceptName.GetCodeNo2
        .Fields("AcceptTime") = SolveNull(gdtAcceptTime.Text)
        '.Fields("Proposer") = txtProposer.Text
        '#3454 申請人改用ComboBox By Kin 2007/12/18
        .Fields("Proposer") = cboProposer.Text
        .Fields("BankCode") = gilBankCode.GetCodeNo2
        .Fields("BankName") = gilBankCode.GetDescription2
        .Fields("CardCode") = gilCardName.GetCodeNo2
        .Fields("CardName") = gilCardName.GetDescription2
        .Fields("StopYM") = SolveNull(mskCardExpDate.Text & "")
         
         If GetUserPriv("SO1100G", "(SO1100G0)") Then
            .Fields("ID") = UCase(txtID.Text)
            .Fields("AccountNameID") = UCase(txtAccOwnerID.Text)
         Else
            If txtID.Tag = Empty And txtID <> Empty Then txtID.Tag = txtID
            .Fields("ID") = UCase(txtID.Tag)
             If txtAccOwnerID.Tag = Empty And txtAccOwnerID <> Empty Then txtAccOwnerID.Tag = txtAccOwnerID
            .Fields("AccountNameID") = UCase(txtAccOwnerID.Tag)
         End If
         
'         If GetUserPriv("SO1100G", "(SO1100G9)") Then
'            .Fields("AccountID") = SolveNull(txtAccountNo & "")
'         Else
'             If txtAccountNo.Tag = Empty And txtAccountNo <> Empty Then
'                txtAccountNo.Tag = txtAccountNo
'            End If
'            .Fields("AccountID") = SolveNull(txtAccountNo.Tag)
'         End If
        Dim aWhereIn As String
        Dim aMaxACHNo As String
        Dim aACHTotal As String
        Dim isAch As Boolean
        
        If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018" & _
                            " WHERE CODENO=" & gilBankCode.GetCodeNo & _
                            " AND ( PRGNAME LIKE 'ACH%' Or " & posAchContition & " )" & _
                            " AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
                                       
            isAch = True
        End If
                        
       If isAch Then
            aMaxACHNo = GetMaxAchNo(csmACH.GetQryFld(1) & "")
            
            For i = 1 To 99
                If aWhereIn = "" Then
                    aWhereIn = aMaxACHNo & GetString(txtAccountNo & "", 6, giRight, True) & _
                            Right("00" & i, 2)
                    aWhereIn = "'" & aWhereIn & "'"
                Else
                    aWhereIn = aWhereIn & ",'" & aMaxACHNo & GetString(txtAccountNo & "", 6, giRight, True) & _
                            Right("00" & i, 2) & "'"
                End If
                 
            Next
        End If
        
        If EditMode = giEditModeEdit Then
             If IsNumeric(txtAccountNo) Then
                .Fields("AccountID") = SolveNull(txtAccountNo & "")
             Else
                .Fields("AccountID") = strRealAccNo
             End If
             '#3585 測試不OK,編輯狀態也要修改(RA文件重新異動)
             If .Fields("AchCustId") & "" <> "" And gdaSendDate.GetValue = Empty Then
                '#5354 修改ACH的H項，若該筆資料沒有送件日期，且修改前、修改後，其帳號相同，就不要重算ACH用戶識別碼 By Kin 2012/09/03
                If .Fields("ACCOUNTID").OriginalValue <> txtAccountNo.Text Then
                    If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018" & _
                                        " WHERE CODENO=" & gilBankCode.GetCodeNo & _
                                        " AND (PRGNAME LIKE 'ACH%' Or " & posAchContition & " ) " & _
                                        " AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
                        '計算時需扣掉自己本身
                        '#5354 計算時,只要取帳號後6碼就好 SUBSTR(LPAD(ACCOUNTID,30,'0'),25,6) By Kin 2012/09/03
                        '#5354 ACHCUSTID後面2碼從01-99找出沒用到的 By Kin 2012/09/05
'                        strSqlCid = "SELECT COUNT(*) AS COUNTS FROM " & GetOwner & " SO106 " & _
'                                        " Where ACCOUNTID='" & txtAccountNo.Text & "'" & _
'                                        " AND SUBSTR(LPAD(ACCOUNTID,30,'0'),25,6)='" & Right(txtAccountNo.Text, 6) & "'" & _
'                                        " AND RowID<>'" & .Fields("RowID") & "" & "'"
                        strSqlCid = "SELECT ACHCUSTID FROM " & GetOwner & " SO106 " & _
                                        " Where SUBSTR(LPAD(ACCOUNTID,30,'0'),25,6)='" & Right(txtAccountNo.Text, 6) & "'" & _
                                        " AND RowID<>'" & .Fields("RowID") & "" & "'" & _
                                        " AND ACHCUSTID IN (" & aWhereIn & ") ORDER BY ACHCUSTID "
                        
                        If Not GetRS(rsAchCid, strSqlCid, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
                        aACHTotal = GetAchNoTotal(rsAchCid)
                        
                        If csmACH.GetQryFld(1) & "" <> "" And isAch Then
                            If blnACHCustID Then
'                                strAchCid = GetMaxAchNo(csmACH.GetQryFld(1) & "") & _
'                                            GetString(txtAccountNo & "", 6, giRight, True) & _
'                                                Right("00" & CStr(Trim(rsAchCid("Counts")) + 1), 2)
                                strAchCid = aMaxACHNo & _
                                            GetString(txtAccountNo & "", 6, giRight, True) & _
                                                Right("00" & aACHTotal, 2)
                                .Fields("ACHCustid") = strAchCid
                            Else
                                '#4006 舊格式編碼原則也要與新格式一樣 By Kin 2008/07/21
                                strAchCid = GetMaxAchNo(csmACH.GetQryFld(1) & "") & Format(gCustId & "", "00000000")
                                .Fields("ACHCustid") = strAchCid
                            End If
                        End If
    
                    End If
                End If
             End If
        ElseIf EditMode = giEditModeInsert Then
            .Fields("AccountID") = SolveNull(txtAccountNo & "")
            
            '****************************************************************************************************************************************************************************************
            '#3585 ACHCustID=1時，需要填入ACH用戶號碼 By Kin 2007/10/26
            '#3585 測試不OK,本來要判斷核準日期(gdaSnactionDate),但RA又修改成要判斷送件日期(gdaSendDate) By Kin 2007/10/30
            '#3585 測試不OK(RA文件重新異動) SO041.AchCustId=0時也要自動產生 By Kin 2007/10/30
            If gdaSendDate.GetValue = Empty Then
                If blnACHCustID Then
                    If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018" & _
                                    " WHERE CODENO=" & gilBankCode.GetCodeNo & _
                                    " AND ( PRGNAME LIKE 'ACH%' Or " & posAchContition & " ) " & _
                                    " AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
                        '#5354 計算時,只要取帳號後6碼就好 SUBSTR(LPAD(ACCOUNTID,30,'0'),25,6) By Kin 2012/09/03
                        '#5354 ACHCUSTID後面2碼從01-99找出沒用到的 By Kin 2012/09/05
'                        strSqlCid = "SELECT COUNT(*) AS COUNTS FROM " & GetOwner & " SO106 " & _
'                                                                    " Where SUBSTR(LPAD(ACCOUNTID,30,'0'),25,6)='" & Right(txtAccountNo.Text, 6) & "'" & _
'                                                                    " AND SUBSTR(ACHCUSTID,4,6) ='" & Right(txtAccountNo & "", 6) & "'"
                        strSqlCid = "SELECT ACHCUSTID FROM " & GetOwner & " SO106 " & _
                                                                    " Where SUBSTR(LPAD(ACCOUNTID,30,'0'),25,6)='" & Right(txtAccountNo.Text, 6) & "'" & _
                                                                    " AND SUBSTR(ACHCUSTID,4,6) ='" & Right(txtAccountNo & "", 6) & "'" & _
                                                                    " AND ACHCUSTID IN (" & aWhereIn & ")"
                        If Not GetRS(rsAchCid, strSqlCid, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
                        aACHTotal = GetAchNoTotal(rsAchCid)
    '                    strAchCid = Replace(csmACH.GetQryFld(1), "'", "") & _
    '                                GetString(txtAccountNo & "", 6, giRight, True) & _
    '                                Right("00" & CStr(Trim(rsAchCid("Counts")) + 1), 2)
                        If csmACH.GetQryFld(1) & "" <> "" Then
                            strAchCid = aMaxACHNo & _
                                        GetString(txtAccountNo & "", 6, giRight, True) & _
                                        Right("00" & aACHTotal, 2)
                            .Fields("ACHCustid") = strAchCid
                        End If
                    End If
                Else
                    If csmACH.GetQryFld(1) & "" <> "" And isAch Then
                        strAchCid = GetMaxAchNo(csmACH.GetQryFld(1) & "") & Format(gCustId & "", "00000000")
                        .Fields("ACHCustid") = strAchCid
                    End If
                End If
            End If
            '**********************************************************************************************************************************************************************************************
        End If
        
'        .Fields("CITEMSTR") = gimCitem.GetQueryCode
        .Fields("CITEMSTR") = IIf(csmCitem1.GetQueryCode <> Empty, csmCitem1.GetQueryCode(True), Null)
        
        '*************************************************************************************************
        '#3791 因為CitemStr2會被存成兩個'',所以會導致比對SO033的資料錯誤 By Kin 2008/04/10
        .Fields("CITEMSTR2") = IIf(csmCitem2.GetQueryCode <> Empty, csmCitem2.GetQueryCode(True), Null)
        '*************************************************************************************************
        
        '.Fields("ACHTNO") = csmACH.GetQueryCode
        If csmACH.GetDispStr & "" <> "" Then
            '#3236 因為SetQueryCode方法改變,儲存時改用GetQryFld方法 By Kin 2007/05/28
            .Fields("ACHTNO") = IIf(csmACH.GetQueryCode <> Empty, csmACH.GetQryFld(1), Null)
            '#3236 多增加儲存"ACHTDESC"欄位 By Kin 2007/05/28
            .Fields("ACHTDESC") = IIf(csmACH.GetQueryCode <> Empty, csmACH.GetQryFld(2), Null)
        Else
            .Fields("ACHTNO") = Null
            .Fields("ACHTDESC") = Null
        End If
        .Fields("CVC2") = SolveNull(txtCVC2 & "")
        .Fields("AccountName") = txtAccountOwner.Text & ""
        .Fields("PropDate") = SolveNull(gdaContiDate.Text)
        .Fields("SendDate") = SolveNull(gdaSendDate.Text)
        
        .Fields("AuthorizeStopDate") = SolveNull(gdaAuthStopDate.Text)
        
        .Fields("SnactionDate") = SolveNull(gdaSnactionDate.Text)
        .Fields("StopFlag") = chkStop.Value
        .Fields("StopDate") = SolveNull(gdaStopDate.Text)
        .Fields("MediaCode") = gilMediaCode.GetCodeNo2
        .Fields("MediaName") = gilMediaCode.GetDescription2
        .Fields("IntroID") = SolveNull(txtIntroId.Text & "")
        .Fields("IntroName") = SolveNull(lblIntroName.Caption & "")
        .Fields("Updtime") = GetDTString(RightNow)
        .Fields("UpdEn") = garyGi(1)
        .Fields("Note") = SolveNull(txtNote.Text & "")
        .Fields("CMcode") = gilCMCode.GetCodeNo2
        .Fields("CMname") = gilCMCode.GetDescription2
        .Fields("PTcode") = gilPTCode.GetCodeNo2
        .Fields("PTname") = gilPTCode.GetDescription2
        .Fields("Alien") = chkFore.Value
        .Fields("AccountAlien") = chkFore2.Value
        .Fields("AddCitemAccount") = chkAddAcc.Value
        'ACH 部份 93/03/02 Hammer
'         Select Case lblStatus
'                Case "已授權成功"
'                    .Fields("Status") = 1
'                Case "已取消授權"
'                    .Fields("Status") = 2
'                Case Else
'                    .Fields("Status") = 0
'         End Select
'        .Fields("ACHCustID") = txtACHcustID
        .Fields("DeAuthorize") = chkDeAut.Value
        .Fields("ToACH") = chkToACH.Value
        .Fields("ACHSN") = txtAchSN
        '*********************************************************************************************************************************************************
        '#3728 判斷ACH那些有異動 By Kin 2008/03/05
        If EditMode = giEditModeInsert Then
            
            
            Call GetCitemCode(IIf(csmACH.GetDispStr <> Empty, csmACH.GetQryFld(1), Empty), IIf(csmACH.GetDispStr <> Empty, csmACH.GetQryFld(2), Empty), _
                              IIf(csmCitem1.GetDispStr <> Empty, csmCitem1.GetQryFld(1), Empty), _
                              IIf(csmCitem1.GetDispStr <> Empty, csmCitem1.GetQryFld(2), Empty), True, arySO106A)
        Else
            Call GetCitemCode(IIf(csmACH.GetDispStr <> Empty, csmACH.GetQryFld(1), Empty), _
                              IIf(csmACH.GetDispStr <> Empty, csmACH.GetQryFld(2), Empty), _
                              IIf(csmCitem1.GetDispStr <> Empty, csmCitem1.GetQryFld(1), Empty), _
                              IIf(csmCitem1.GetDispStr <> Empty, csmCitem1.GetQryFld(2), Empty), False, arySO106A)
        End If
        '*********************************************************************************************************************************************************
        '#5946 增加記錄Log檔
        Call InsSO106Log
        .Update
        
        
        '************************************************************************************************************************************************
        '#3728 如果ACH有進行異動則開始比較 並新增或更新SO106A By Kin 2008/03/10
        Dim strUpdTime As String
        Dim strCreateTime As String
        Dim strStopDate As String
        strUpdTime = GetDTString(RightNow)
        strCreateTime = Format(RightNow, "yyyy/mm/dd hh:mm:ss")
        strStopDate = Format(RightNow, "yyyy/mm/dd hh:mm:ss")
        For i = LBound(arySO106A) To UBound(arySO106A)
        
            Dim strUpdSO106A As String
            strUpdSO106A = Empty
            Select Case arySO106A(i, 2)
                '產生一筆授權資料
                Case "0"
                    '#3946 新增時將AuthorizeStatus原本為0改成4,代表待提出
                    strUpdSO106A = "Insert into " & GetOwner & "SO106A " & _
                                "(MasterRowID,ACHTNO,CitemCodeStr,CitemNameStr," & _
                                "UpdEn,UpdTime,CreateTime,CreateEn,RecordType," & _
                                "AuthorizeStatus,ACHDesc,MasterId)" & _
                                " Values ('" & .Fields("RowId") & "'," & _
                                          "'" & arySO106A(i, 0) & "'," & _
                                          IIf(arySO106A(i, 3) = Empty, "Null,", "'" & arySO106A(i, 3) & "',") & _
                                          IIf(arySO106A(i, 4) = Empty, "Null,", "'" & arySO106A(i, 4) & "',") & _
                                          "'" & garyGi(1) & "'," & _
                                          "'" & strUpdTime & "'," & _
                                          "TO_DATE('" & strCreateTime & "','YYYY/MM/DD HH24:MI:SS')," & _
                                          "'" & garyGi(1) & "'," & _
                                          "0,4,'" & arySO106A(i, 1) & "'," & .Fields("MasterId") & ")"
                    gcnGi.Execute strUpdSO106A
                '產生一筆取消授權資料,如果是授權中的資料，則不允許產生一筆資料
                '#3946 原本是以MasterRowId做對應,現在改為以MasterId做對應 By Kin 2008/05/30
                Case "1"
                    '#3946 將AuthorizeStatus狀態改為4 By Kin 2008/05/30
                    strUpdSO106A = "Select * From " & GetOwner & "SO106A" & _
                                    " Where MasterId='" & .Fields("MasterId") & "'" & _
                                    " And ACHTNO='" & arySO106A(i, 0) & "'" & _
                                    " And ACHDesc='" & arySO106A(i, 1) & "'" & _
                                    " And " & IIf(arySO106A(i, 3) = "", " CitemCodeStr is Null", " CitemCodeStr='" & arySO106A(i, 3) & "'") & _
                                    " And " & IIf(arySO106A(i, 4) = "", "CitemNameStr is Null", " CitemNameStr='" & arySO106A(i, 4) & "'") & _
                                    " And AuthorizeStatus=4"
                    Set rsSO106A = gcnGi.Execute(strUpdSO106A)
                    '3946 有異動的話,不取消停用,直接刪除 By Kin 2008/05/23
                    If Not rsSO106A.EOF Then
                        If IsNull(rsSO106A("AuthorizeStatus")) Or Val(rsSO106A("AuthorizeStatus") & "") = 3 Or Val(rsSO106A("AuthorizeStatus") & "") = 4 Then
'                            strUpdSO106A = "Update " & GetOwner & "SO106A" & _
'                                         " Set StopFlag=1," & _
'                                         "StopDate=TO_DATE('" & strStopDate & "','YYYY/MM/DD HH24:MI:SS')," & _
'                                         "Notes=Notes" & IIf(IsNull(rsSO106A("AuthorizeStatus")), " || '授權中取消該筆資料'", " ||'授權失敗取消該筆'") & _
'                                         " Where MasterRowId='" & .Fields("RowId") & "'" & _
'                                         " And ACHTNO='" & arySO106A(i, 0) & "'" & _
'                                         " And ACHDesc='" & arySO106A(i, 1) & "'" & _
'                                         " And " & IIf(arySO106A(i, 3) = "", "CitemCodeStr is Null", "CitemCodeStr='" & arySO106A(i, 3) & "'") & _
'                                         " And " & IIf(arySO106A(i, 4) = "", "CitemNameStr is Null", "CitemNameStr='" & arySO106A(i, 4) & "'") & _
'                                         " And AuthorizeStatus is Null"
                            '#3946 將AuthorizeStatus狀態改為4 By Kin 2008/05/30
                            strUpdSO106A = "Delete From " & GetOwner & "SO106A " & _
                                            " Where MasterId='" & .Fields("MasterId") & "'" & _
                                            " And ACHTNO='" & arySO106A(i, 0) & "'" & _
                                            " And ACHDesc='" & arySO106A(i, 1) & "'" & _
                                            " And " & IIf(arySO106A(i, 3) = "", "CitemCodeStr is Null", "CitemCodeStr='" & arySO106A(i, 3) & "'") & _
                                            " And " & IIf(arySO106A(i, 4) = "", "CitemNameStr is Null", "CitemNameStr='" & arySO106A(i, 4) & "'") & _
                                            " And  AuthorizeStatus=4"
                            blnCallSO106A = False
                        Else
                            strUpdSO106A = "Insert into " & GetOwner & "SO106A " & _
                                        "(MasterRowID, ACHTNO, CitemCodeStr, CitemNameStr," & _
                                        " UpdEn,UpdTime,CreateTime,CreateEn,RecordType,StopFlag," & _
                                          "StopDate,AuthorizeStatus,ACHDesc,MasterId) " & _
                                        " Values ('" & .Fields("RowId") & "'," & _
                                                  "'" & arySO106A(i, 0) & "'," & _
                                                  IIf(arySO106A(i, 3) = Empty, "Null,", "'" & arySO106A(i, 3) & "',") & _
                                                  IIf(arySO106A(i, 4) = Empty, "Null,", "'" & arySO106A(i, 4) & "',") & _
                                                  "'" & garyGi(1) & "'," & _
                                                  "'" & strUpdTime & "'," & _
                                                  "TO_DATE('" & strCreateTime & "','YYYY/MM/DD HH24:MI:SS')," & _
                                                  "'" & garyGi(1) & "'," & _
                                                  "1,1," & _
                                                  "TO_DATE('" & strStopDate & "','YYYY/MM/DD HH24:MI:SS')," & _
                                                  "Null,'" & arySO106A(i, 1) & "'," & .Fields("MasterId") & ")"
                            If strFirstACHID = Empty Then strFirstACHID = arySO106A(i, 0)
                            If strFirstACHDesc = Empty Then strFirstACHDesc = arySO106A(i, 1)
                            blnCallSO106A = True
                            strACHID = IIf(strACHID = Empty, "'" & arySO106A(i, 0) & "'", strACHID & ",'" & arySO106A(i, 0) & "'")
                            strACHDesc = IIf(strACHDesc = Empty, "'" & arySO106A(i, 1) & "'", strACHDesc & ",'" & arySO106A(i, 1) & "'")
                        End If
                    Else
                        '#3946 將AuthorizeStatus狀態改為4 By Kin 2008/05/30
                        strUpdSO106A = "Insert into " & GetOwner & "SO106A " & _
                                        "(MasterRowID, ACHTNO, CitemCodeStr, CitemNameStr," & _
                                        " UpdEn,UpdTime,CreateTime,CreateEn,RecordType,StopFlag," & _
                                          "StopDate,AuthorizeStatus,ACHDesc,MasterId) " & _
                                        " Values ('" & .Fields("RowId") & "'," & _
                                                  "'" & arySO106A(i, 0) & "'," & _
                                                  IIf(arySO106A(i, 3) = Empty, "Null,", "'" & arySO106A(i, 3) & "',") & _
                                                  IIf(arySO106A(i, 4) = Empty, "Null,", "'" & arySO106A(i, 4) & "',") & _
                                                  "'" & garyGi(1) & "'," & _
                                                  "'" & strUpdTime & "'," & _
                                                  "TO_DATE('" & strCreateTime & "','YYYY/MM/DD HH24:MI:SS')," & _
                                                  "'" & garyGi(1) & "'," & _
                                                  "1,1," & _
                                                  "TO_DATE('" & strStopDate & "','YYYY/MM/DD HH24:MI:SS')," & _
                                                  "4,'" & arySO106A(i, 1) & "'," & .Fields("MasterId") & ")"
                        If strFirstACHID = Empty Then strFirstACHID = arySO106A(i, 0)
                        If strFirstACHDesc = Empty Then strFirstACHDesc = arySO106A(i, 1)
                        blnCallSO106A = True
                        strACHID = IIf(strACHID = Empty, "'" & arySO106A(i, 0) & "'", strACHID & ",'" & arySO106A(i, 0) & "'")
                        strACHDesc = IIf(strACHDesc = Empty, "'" & arySO106A(i, 1) & "'", strACHDesc & ",'" & arySO106A(i, 1) & "'")
                    
                    End If
                    gcnGi.Execute strUpdSO106A
                    '多筆取消時，呼叫SO1100G1用，判斷SO1100G1要停留在那一筆
'                    If strFirstACHID = Empty Then strFirstACHID = arySO106A(i, 0)
'                    If strFirstACHDesc = Empty Then strFirstACHDesc = arySO106A(i, 1)
'                    blnCallSO106A = True
'                    strACHID = IIf(strACHID = Empty, "'" & arySO106A(i, 0) & "'", strACHID & ",'" & arySO106A(i, 0) & "'")
'                    strACHDesc = IIf(strACHDesc = Empty, "'" & arySO106A(i, 1) & "'", strACHDesc & ",'" & arySO106A(i, 1) & "'")
                '更新ACH裡的收費項目
                Case "2"
                    strUpdSO106A = "Select * From " & GetOwner & "SO106A Where " & _
                                 "ACHTNO='" & arySO106A(i, 0) & "' " & _
                                 "And MasterID='" & .Fields("MasterId") & "' " & _
                                 "And ACHDesc='" & arySO106A(i, 1) & "' " & _
                                 "And StopFlag<>1"
                    Set rsSO106A = gcnGi.Execute(strUpdSO106A)
                    If Not rsSO106A.EOF Then
                        If rsSO106A("CitemCodestr") & "" <> arySO106A(i, 3) Then
                            strUpdSO106A = "Update " & GetOwner & "SO106A Set " & _
                                                    "CitemCodeStr=" & IIf(arySO106A(i, 3) <> Empty, "'" & arySO106A(i, 3) & "',", "NULL,") & _
                                                    "CitemNameStr=" & IIf(arySO106A(i, 3) <> Empty, "'" & arySO106A(i, 4) & "',", "NULL,") & _
                                                    "UpdEn='" & garyGi(1) & "'," & _
                                                    "UpdTime='" & strUpdTime & "' " & _
                                           "Where MasterId='" & .Fields("MasterId") & "'" & _
                                           " And ACHTNO='" & arySO106A(i, 0) & "'" & _
                                           " And ACHDesc='" & arySO106A(i, 1) & "'" & _
                                           " And StopFlag<>1"
                            gcnGi.Execute (strUpdSO106A)
                        End If
                    '#3946 測試不OK,如果只有一筆的時候,要直接Update By Kin 2008/07/25
                    Else
                        strUpdSO106A = "Update " & GetOwner & "SO106A Set " & _
                                                    "CitemCodeStr=" & IIf(arySO106A(i, 3) <> Empty, "'" & arySO106A(i, 3) & "',", "NULL,") & _
                                                    "CitemNameStr=" & IIf(arySO106A(i, 3) <> Empty, "'" & arySO106A(i, 4) & "',", "NULL,") & _
                                                    "ACHTNO=" & IIf(arySO106A(i, 0) <> Empty, "'" & arySO106A(i, 0) & "',", "NULL") & _
                                                    "ACHDesc=" & IIf(arySO106A(i, 1) <> Empty, "'" & arySO106A(i, 1) & "',", "NULL") & _
                                                    "UpdEn='" & garyGi(1) & "'," & _
                                                    "UpdTime='" & strUpdTime & "' " & _
                                           "Where MasterId='" & .Fields("MasterId") & "'"
                        gcnGi.Execute (strUpdSO106A)
                    End If

            End Select
        Next i
        '***************************************************************************************************************************************************************************************
         If blnUpd Then
            Dim strAllChild As String
            Dim varAry, varElement
            Dim lngAft As Long, lngChildID As Long
            strAllChild = GetChild
            varAry = Split(strAllChild, ",")
            For Each varElement In varAry
                lngChildID = Val(varElement & "")
                If lngChildID > 0 Then
                    .Fields("CustID") = lngChildID
                    .Fields("InheritFlag") = 1
                    .Fields("ACHSN") = GetRsValue("SELECT ACHSN FROM " & GetOwner & "SO106" & _
                                                                        " WHERE CUSTID=" & lngChildID & _
                                                                        " AND INHERITKEY='" & rsSO106("InheritKey") & "'" & _
                                                                        " AND INHERITFLAG=1") & ""
                    Debug.Print InsertToOracle("SO106", rsSO106, _
                                                                " WHERE CUSTID=" & lngChildID & _
                                                                " AND INHERITKEY='" & rsSO106("InheritKey") & "'" & _
                                                                " AND INHERITFLAG=1", lngAft, False)
                    .CancelUpdate
                End If
            Next
         End If
    End With
    
    ScrToRcd = True
    'csmACH.Clear
    On Error Resume Next
    Call CloseRecordset(rsSO106A)
  Exit Function
ChkErr:
    
    ScrToRcd = False
    gcnGi.RollbackTrans
    Call ErrSub(Me.Name, "ScrToRcd")
End Function
'取出ACH後面2碼沒被用的值
Private Function GetAchNoTotal(ByRef rsAchCid As Recordset) As String
  On Error GoTo ChkErr
    Dim aACHTotal As String
    Dim aFind As Boolean
    Dim i As Integer
    aACHTotal = "01"
    aFind = False
    If rsAchCid.RecordCount > 0 Then
        For i = 1 To 99
            If aFind Then Exit For
            aACHTotal = Right("00" & i, 2)
            rsAchCid.MoveFirst
            Do While Not rsAchCid.EOF
                aFind = True
                If Right(rsAchCid("ACHCUSTID"), 2) = aACHTotal Then
                    aFind = False
                    Exit Do
                End If
                rsAchCid.MoveNext
            Loop
        Next
    End If
    GetAchNoTotal = aACHTotal
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "GetAchNoTotal")
End Function
'#5946 增加記錄Log檔
Private Sub InsSO106Log()
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim rsLog As New ADODB.Recordset
    Dim blnChg As Boolean
    Dim i As Integer
    aSQL = "SELECT * FROM SO106LOG WHERE 1=0 "
    If Not GetRS(rsLog, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    
    Select Case lngEditMode
        Case giEditModeEdit
            '判斷有沒有異動資料
            For i = 0 To rsSO106.Fields.Count - 1
                If UCase(rsSO106.Fields(i).Name) <> "UPDTIME" Then
                    If (rsSO106.Fields(i).Value <> rsSO106.Fields(i).OriginalValue) Or _
                      (Len(rsSO106.Fields(i).Value & "") <> Len(rsSO106.Fields(i).OriginalValue & "")) Then
                        blnChg = True
                        Exit For
                    End If
                End If
            Next i
            '有異動資料
            If blnChg Then
                rsLog.AddNew
                rsLog("FuncType") = 1
                For i = 0 To rsSO106.Fields.Count - 1
                    If UCase(rsSO106.Fields(i).Name) <> "ROWID" Then
                        If Len(rsSO106.Fields(i).OriginalValue & "") = 0 Then
                            rsLog.Fields(rsSO106.Fields(i).Name) = Null
                        Else
                            rsLog.Fields(rsSO106.Fields(i).Name) = rsSO106.Fields(i).OriginalValue & ""
                        End If
                    End If
                Next i
                
                For i = 0 To rsSO106.Fields.Count - 1
                    If (rsSO106.Fields(i).Value <> rsSO106.Fields(i).OriginalValue) Or _
                      (Len(rsSO106.Fields(i).Value & "") <> Len(rsSO106.Fields(i).OriginalValue & "")) Then
                        Select Case UCase(rsSO106.Fields(i).Name)
                            Case "ROWID"
                            Case "UPDTIME"
                                rsLog("UPDTIME") = GetDTString(RightNow)
                            Case "UPDEN"
                                rsLog("UPDEN") = garyGi(1)
                            Case "COMPCODE"
                            Case "CUSTID"
                            Case "MASTERID"
                            Case "NEWUPDTIME"
                                rsLog("NEWUPDTIME") = rsSO106.Fields("NEWUPDTIME")
                            Case Else
                                If Len(rsSO106.Fields(i).Value & "") = 0 Then
                                    rsLog.Fields(rsSO106.Fields(i).Name & "B").Value = Null
                                Else
                                    rsLog.Fields(rsSO106.Fields(i).Name & "B").Value = rsSO106.Fields(i).Value
                                End If
                        End Select
                      
                    End If
                Next i
                rsLog.Update
            End If
        Case giEditModeInsert
            rsLog.AddNew
            rsLog("FuncType") = 3
            For i = 0 To rsSO106.Fields.Count - 1
                Select Case UCase(rsSO106.Fields(i).Name)
                    Case "ROWID"
                    Case "UPDTIME"
                        rsLog.Fields("UPDTIME") = GetDTString(RightNow)
                    Case "UPDEN"
                        rsLog.Fields("UPDEN").Value = garyGi(1)
                    Case "COMPCODE"
                    Case "CUSTID"
                    Case "MASTERID"
                    Case "NEWUPDTIME"
                                rsLog("NEWUPDTIME") = rsSO106.Fields("NEWUPDTIME")
                    Case Else
                        If Len(rsSO106.Fields(i).Value & "") = 0 Then
                            rsLog.Fields(rsSO106.Fields(i).Name & "B").Value = Null
                        Else
                            rsLog.Fields(rsSO106.Fields(i).Name & "B").Value = rsSO106.Fields(i).Value
                        End If
                End Select
            Next i
            rsLog.Update
        Case giEditModeDelete
            rsLog.AddNew
            rsLog("FuncType") = 2
            If UCase(Screen.ActiveForm.Name) = "FRMSO1100G" Then
                For i = 0 To rsSO106.Fields.Count - 1
                    If UCase(rsSO106.Fields(i).Name) <> "ROWID" Then
                        If Len(rsSO106.Fields(i).Value & "") = 0 Then
                            rsLog.Fields(rsSO106.Fields(i).Name) = Null
                        Else
                            rsLog.Fields(rsSO106.Fields(i).Name) = rsSO106.Fields(i).Value & ""
                        End If
                    End If
                    
                Next i
            Else
                With frmSO1100BMDI
                    For i = 0 To .rsSO106.Fields.Count - 1
                        If UCase(.rsSO106.Fields(i).Name) <> "ROWID" Then
                            If Len(.rsSO106.Fields(i).Value & "") = 0 Then
                                rsLog.Fields(.rsSO106.Fields(i).Name) = Null
                            Else
                                rsLog.Fields(.rsSO106.Fields(i).Name) = .rsSO106.Fields(i).Value & ""
                            End If
                        End If
                        
                    Next i
                End With
            End If
            rsLog.Fields("UPDTIME") = GetDTString(RightNow)
            rsLog.Fields("UPDEN") = garyGi(1)
            rsLog.Update
    End Select

    On Error Resume Next
    Call CloseRecordset(rsLog)
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "InsSO106Log")
  
End Sub
Public Sub FunctionKey(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Select Case KeyCode
           '----------------------------------------------------
           Case vbKeyF2  '   F3:存檔, 相當於按下cmdsave
           '#3337 判斷是否在瀏覽模式按F2，如果是瀏覽模式不能存檔 By Kin 2007/07/23
                If lngEditMode = giEditModeView Then Exit Sub
                If Not ChkGiList(KeyCode) Then Exit Sub
                UpdateGo
           '----------------------------------------------------
           Case vbKeyEscape
                CancelGo
    End Select
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    IsDataOk = False
    blnInsSO138 = False
   
    
    If txtAchSN <> Empty Then
        If EditMode = giEditModeInsert Then
            If Val(RPxx(gcnGi.Execute("SELECT COUNT(*) FROM " & _
                                                        GetOwner & "SO106 WHERE ACHSN='" & txtAchSN & "'").GetString)) > 0 Then
                MsgBox "[申請書單號] 資料重複 ! 請確認 !", vbInformation, "訊息"
                On Error Resume Next
                txtAchSN.SetFocus
                Exit Function
            End If
        ElseIf EditMode = giEditModeEdit Then
            Dim strRID As String
            On Error Resume Next
            strRID = RPxx(gcnGi.Execute("SELECT ROWID FROM " & GetOwner & "SO106 WHERE ACHSN='" & txtAchSN & "'").GetString & "")
            If Err.Number = 0 Then
                If strRID <> rsSO106("ROWID") Then
                    MsgBox "[申請書單號] 資料重複 ! 請確認 !", vbInformation, "訊息"
                    On Error Resume Next
                    txtAchSN.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    If gdtAcceptTime.Text = "" Then gdtAcceptTime.SetFocus: GoTo 66
    If gilAcceptName.GetCodeNo = "" Then gilAcceptName.SetFocus: GoTo 66
    '#3454 申請人改用ComboBox By Kin 2007/12/18
    If cboProposer.Text = "" Then cboProposer.SetFocus: GoTo 66
'    If txtProposer = "" Then txtProposer.SetFocus: GoTo 66
'    If lblID.ForeColor = vbRed Then
'        If txtID = "" Then txtID.SetFocus: GoTo 66
'    End If
    If gdaContiDate.Text = "" Then gdaContiDate.SetFocus: GoTo 66
    If chkStop.Value = 1 Then
        If gdaStopDate.Text = "" Then gdaStopDate.SetFocus: GoTo 66
    End If
    If gilCMCode.GetCodeNo = "" Then gilCMCode.SetFocus: GoTo 66
    If gilPTCode.GetCodeNo = Empty Then gilPTCode.SetFocus: GoTo 66
    If gilBankCode.GetCodeNo = "" Then gilBankCode.SetFocus: GoTo 66
    If lblAccountOwner.ForeColor = vbRed Then
        If txtAccountOwner.Text = Empty Then txtAccountOwner.SetFocus: GoTo 66
    End If
    If lblAccNameID.ForeColor = vbRed Then
        If txtAccOwnerID.Text = Empty Then txtAccOwnerID.SetFocus: GoTo 66
    End If
    '#8724
    If intCD031RefNo = 4 Or intCD031RefNo = 5 Then
        If gilCardName.GetCodeNo = Empty Then gilCardName.SetFocus: GoTo 66
        If mskCardExpDate.Text = Empty Then mskCardExpDate.SetFocus: GoTo 66
    End If
    
    If lblAccountNo.ForeColor = vbRed Then
        If Trim(txtAccountNo) = Empty Then txtAccountNo.SetFocus: GoTo 66
    End If
    '*************************************************************************************************************************
    '#3395 收費方式為信用卡時(CD031.RefNo=4)，使用信用卡檢核方式，收費方式為銀行時規則不變(CD031.RefNo=2) By Kin 2007/08/20
    '#8724
    If intCD031RefNo = 2 Or intCD031RefNo = 4 Or intCD031RefNo = 5 Then
        If intCD031RefNo = 2 Then
            If ServiceCommon.AccountNo_Validate(Me) Then
                txtAccountNo.SetFocus
                Exit Function
            End If
        Else
            If Not ChkCreditCard(gilCardName.GetCodeNo, txtAccountNo.Text) Then
                txtAccountNo.SetFocus
                Exit Function
            End If
        End If
    End If
    '**************************************************************************************************************************
    If txtID <> Empty And chkFore.Value = 0 Then
        If GetUserPriv("SO1100G", "(SO1100G0)") Then
            IDIsOk txtID.Text, True
        Else
            IDIsOk txtID.Tag, True
        End If
'        If Not IDIsOk(txtID, True) Then
'    '       On Error Resume Next
'    '       txtID.SetFocus
'    '       Exit Function
'        End If
    End If
    If txtAccOwnerID <> Empty And chkFore2.Value = 0 Then
        If GetUserPriv("SO1100G", "(SO1100G0)") Then
            IDIsOk txtAccOwnerID.Text, True
        Else
            IDIsOk txtAccOwnerID.Tag, True
        End If
'        If Not IDIsOk(txtAccOwnerID, True) Then
'    '       On Error Resume Next
'    '       txtID.SetFocus
'    '       Exit Function
'        End If
    End If
    
    If csmACH.GetDispStr = Empty Then
        
        If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018" & _
                                " WHERE CODENO=" & gilBankCode.GetCodeNo & _
                                " AND (PRGNAME LIKE 'ACH%' Or " & posAchContition & " ) " & _
                                " AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
            If Not blnACHCustID Then
                On Error Resume Next
                csmACH.SetFocus
                GoTo 66
                Exit Function
            Else
                If GetRsValue("Select Count(*) From " & GetOwner & "SO106A" & _
                        " Where MasterId=" & IIf(rsSO106.EOF, 0, rsSO106("MasterId")) & " And AuthorizeStatus = 1") = 0 Then
                    On Error Resume Next
                    csmACH.SetFocus
                    GoTo 66
                End If
            End If
        End If
    End If
    '*******************************************************************************
    '#3436 判斷是否有發票資料 By Kin 2007/12/13
    
    '#3982 檢查SO138是否有帳號所有人的資料,如果沒有直接新增SO138,如果有讓User自行選擇 By Kin 2008/07/16
    If EditMode = giEditModeInsert Then
'        If Not ChkInvData Then
            '***********************************************************
            '#4210 不檢查是否存在了,全部改為直接新增 By Kin 2008/11/06
'            If strInvSeqNo = Empty Or Val(strInvSeqNo) <= 0 Then
'                If ChkSO138 Then
'                    MsgBox "請選擇一筆發票資料 !!", vbInformation, "訊息"
'                    If cmdCreateINV.Enabled Then cmdCreateINV.Value = True
'                    Exit Function
'                Else
'                    blnInsSO138 = True
'                    'Call InsSO138
'                End If
'            End If
            blnInsSO138 = True
            '*************************************************************
 '       End If
    ElseIf EditMode = giEditModeEdit Then
        If chkStop.Value <> 1 And gdaSnactionDate.Text <> Empty Then
            If rsSO106("InvSeqNo") & "" = "" Or Val(rsSO106("InvSeqNO") & "") <= 0 Then
                '***************************************************************
                '#4210 不檢查是不存在了,全部改為直接新增 By Kin 2008/11/06
'                If strInvSeqNo = "" Or Val(strInvSeqNo) <= 0 Then
'                    If ChkSO138 Then
'                        MsgBox "請選擇一筆發票資料 !!", vbInformation, "訊息"
'                        If cmdCreateINV.Enabled Then cmdCreateINV.Value = True
'                        Exit Function
'                    Else
'                        'Call InsSO138
'                        blnInsSO138 = True
'                    End If
'                End If
                blnInsSO138 = True
                '*****************************************************************
            Else
                strInvSeqNo = rsSO106("InvSeqNo") & ""
            End If
        End If

    End If
    '******************************************************************************
    IsDataOk = True
  Exit Function
66:
    MsgBox " 此為必要欄位,須有值 !! ", vbExclamation, "訊息.."
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub NewRcd()
  On Error Resume Next
    Dim ctl As Control
    '#3322 多增一個Flag，以判別是否為新增狀態(因為不能Show出ACH狀態改變的訊息) By Kin 2007/07/04
    blnAddFlag = True
    For Each ctl In Me
        If TypeOf ctl Is TextBox Then ctl = ""
        If TypeOf ctl Is GiDate Then Call ctl.SetValue("")
        If TypeOf ctl Is GiTime Then ctl.Text = ""
        If TypeOf ctl Is GiList Then ctl.Clear
        If TypeOf ctl Is CheckBox Then ctl.Value = 0
        If TypeOf ctl Is MaskEdBox Then ctl.Text = ""
        If TypeOf ctl Is CSmulti Then ctl.Clear
    Next
    chkAddAcc.Value = 1
    lblUpdTime.Caption = GetDTString(RightNow)
    lblUpdEn.Caption = garyGi(1)
    cmdInvSeqNo.Enabled = True
    cmdCreateINV.Enabled = True
    strInvSeqNo = Empty
    lblInvSeqNo = ""
    '新增資料時預設讓停用能使用 By Kin 2008/09/19
    chkStop.Enabled = True
    gdaStopDate.Enabled = True
    '#3728 之前沒有過濾停用,在此問題集發現,將停用過濾掉 By Kin 2008/03/10
    '#3946 多增加 ACHTNO is Not Null 條件 By Kin 2008/05/23
    '#4030 不過濾停用 By Kin 2008/08/26
    '#4106 增加判斷ACHType參數 By Kin 2008/09/22
    '#7049 增加BillHeadFmt欄位,可以串到CD068A By Kin 2015/07/08
    If blnStartPost Then
        AchTypeCondition = " In (1,2 ) "
    Else
        AchTypeCondition = " In ( 1 ) "
    End If
    
    SetMsQry csmACH, "SELECT ACHTNO,ACHTDESC,CITEMCODESTR,BillHeadFmt FROM " & GetOwner & "CD068 Where ACHTNO Is Not Null And ACHTDESC is Not NULL And ACHType " & AchTypeCondition, _
                           "ACH交易代碼", "ACH交易說明", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
    csmACH.Clear
    csmCitem1.Clear
    csmCitem2.Clear
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "NewRcd")
End Sub

Public Sub UpdateGo()
  On Error GoTo ChkErr
    Dim strSQL As String
    Dim lngChargeAddrNo As Long
    Dim lngMailAddrNo As Long
    Dim blnReturn As Boolean
    Dim blnAdjust As Boolean
    Dim lngPos As Long
    
    If IsDataOk() = False Then Exit Sub
    
    If EditMode = giEditModeEdit Then lngPos = rsSO106.AbsolutePosition
    
    gcnGi.BeginTrans
    '#4068 將自動新增SO138改來這裡,不然strInvseqNo會是空值,造成SO003沒被Update到 By Kin 2008/09/02
    '#4130 增加判斷strInvSeqNo是否有值,如果有值要以有值的為準 By Kin 2008/11/10
    If blnInsSO138 Then
        If strInvSeqNo = Empty Then
            Call InsSO138
        Else
            blnInsSO138 = False
        End If
    End If
    If EditMode = giEditModeEdit Then
        If Not IsNull(rsSO106.Fields("ACHTNO").OriginalValue) And csmACH.GetDispStr = Empty Then
            chkStop.Value = 1
        End If
    
    End If
    If GetUserPriv("SO1100G", "(SO1100G9)") Then
       strRealAccNo = SolveNull(txtAccountNo & "")
    Else
        If txtAccountNo.Tag = Empty And txtAccountNo <> Empty Then txtAccountNo.Tag = txtAccountNo
        If IsNumeric(txtAccountNo) Then
            strRealAccNo = SolveNull(txtAccountNo)
        Else
            strRealAccNo = SolveNull(txtAccountNo.Tag)
        End If
    End If
    
    If chkStop.Value <> 1 And gdaSnactionDate.Text <> Empty Then
    Else
        StopSO002A
        If StrComp(strRealAccNo, frmSO1100BMDI.rsSo002!AccountNo & "", vbTextCompare) = 0 Then
            blnReturn = MsgBox("是否清除客戶資料主檔的客戶基本資料 (二) 的扣款資料 ? 並選擇新的收費方式 !", vbQuestion + vbYesNo, "詢問") = vbYes
        End If
    End If
    
    If chkStop.Value <> 1 And gdaSnactionDate.Text <> Empty Then
        ChkSO002A
        If chkAddAcc.Value <> 1 Then StopSO002A
        ChkSO003
        ChkSO033
    End If
    
    If chkStop.Value = 1 And gdaStopDate.Text <> Empty Then StopGo
    '#4130 測試不OK,如果流水號為0,要Rollback,不然會Trans二次 By Kin 2008/11/10
    If Not ScrToRcd Then gcnGi.RollbackTrans: Exit Sub            ' 把控制項內的值，存到資料庫裡
    gcnGi.CommitTrans
    If blnCallSO106A And blnACHCustID Then
        cmdACHDetail.Value = True
    End If
    With rsSO106
        
        If .State = adStateOpen Then
    '        rsSO106.Requery
            strSQL = .Source
            .Close
            .CursorLocation = adUseClient
            .Open strSQL, gcnGi, adOpenKeyset, adLockOptimistic
            Set ggrData.Recordset = rsSO106
            
        End If
'        Set frmSO1100BMDI.rsSO106 = rsSO106
    End With
    
    With frmSO1100BMDI
        .rsSO106.Requery
        Set .ggrData.Recordset = .rsSO106
        .ggrData.Refresh
        '***********************************************
        '#3929 發現的問題,如果在編輯模式,要重新指定frmSO1100BMDI.rsSo106.AbsolutePosititon不然在按編輯時會跳到第一筆 By Kin 2008/06/09
        If lngEditMode = giEditModeEdit Then
            .rsSO106.AbsolutePosition = lngPos
        End If
        '***********************************************
    End With
    
    ggrData.Refresh
    
    ggrData.Blank = False
    If blnReturn Then
        ChangeMode giEditModeView
'        Me.Visible = False
'        DoEvents
        Me.Hide
        Unload Me
        frmSO1100BMDI.ProcTrans
        Exit Sub
    End If
'    If blnAdjust Then
'        Me.Hide
'        ChangeMode giEditModeView
'        Unload Me
'        frmSO1100BMDI.AdjustCM
'        Exit Sub
'    End If
'     If EditMode = giEditModeInsert Then '繼續新增
'        Call ChangeMode(giEditModeInsert)
'        AddNewGo
'    Else '進入顯示模式
        If EditMode = giEditModeEdit Then rsSO106.AbsolutePosition = lngPos
        Call ChangeMode(giEditModeView)
        RcdToScr
        If blnInsSO138 Then MsgBox "系統已自動產生發票抬頭資料", vbInformation, "訊息"
        blnInsSO138 = False
'    End If
   Exit Sub
ChkErr:
    gcnGi.RollbackTrans
    Call ErrSub(Me.Name, "UpdateGo")
End Sub

Private Sub StopGo(Optional strAccNo As String = "")
  On Error GoTo ChkErr
    Dim strBankCode As String
    Dim strCMCode As String, strCMName As String
    Dim strPTCode As String, strPTName As String
    Dim rsTmp As New ADODB.Recordset
    Dim RetVal As Long
    Dim strChkSO003 As String
    Dim varSeqNo() As String
    Dim strCitemCode As String
    Dim strCitemSQL As String
    Dim strBillNos As String
    Dim strBillNoSQL As String
    Dim strUCCode As String
    Dim strUCName As String
    Dim strServiceType As String
    Dim rsPara As New ADODB.Recordset
    Dim rsBill As New ADODB.Recordset
    Dim strUCCodeSQL As String
    Dim i As Long
    strCMCode = frmSO1100BMDI.rsSO044!CMcode
    strCMName = GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [收費方式代碼檔] 查無代碼為 " & frmSO1100BMDI.rsSO044!CMcode & " 的資料 !")
    
    If strCMName = Empty Then Exit Sub
    
    strBankCode = gilBankCode.GetCodeNo
    If strAccNo = Empty Then strAccNo = strRealAccNo
    '#5623 原本是RowNum=1,現在改為CodeNo=1
    If GetRS(rsTmp, "SELECT CODENO,DESCRIPTION FROM " & GetOwner & "CD032 WHERE CodeNo=1") Then
        If Not rsTmp.EOF Then
            strPTCode = rsTmp("CODENO") & ""
            strPTName = rsTmp("DESCRIPTION") & ""
            If strPTCode = Empty Then strPTCode = "NULL"
        End If
    End If
    
    
    If csmCitem1.GetQueryCode = Empty Then
        ReDim varSeqNo(0)
        varSeqNo(0) = ""
    Else
        varSeqNo = Split(csmCitem1.GetQueryCode, ",")
    End If
    
    '#3573 測試不OK,改用陣列,取得每一個的值 By Kin 2007/12/04
    For i = 0 To UBound(varSeqNo)
        strChkSO003 = "Select Count(*) Cnt From " & GetOwner & "SO106  Where CustId=" & gCustId & _
                    " And AccountID='" & strAccNo & "'" & _
                    " And CompCode=" & gCompCode & _
                    " And StopFlag=0" & _
                    " And StopDate is Null" & _
                    IIf(EditMode = giEditModeEdit, " And RowID<>'" & rsSO106("RowID") & "'", "") & _
                    IIf(varSeqNo(i) <> "", " And Instr(','||Citemstr||',' ,','||Chr(39)||" & varSeqNo(i) & "||Chr(39)||',')>0", "")
        If gcnGi.Execute(strChkSO003)("Cnt") = 0 Then

            '#3573 測試不OK,清除SO033資料也要根據收費項目 By Kin 2007/12/06
            If varSeqNo(i) <> "" Then
                strCitemSQL = "Select CitemCode From " & GetOwner & "SO003 Where CustId=" & gCustId & _
                            " And AccountNo='" & strAccNo & "'" & _
                            " And CompCode=" & gCompCode & _
                            " And StopFlag=0" & _
                            " And SeqNo=" & varSeqNo(i) & _
                            IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode)
                
                strCitemCode = GetRsValue(strCitemSQL, gcnGi) & ""

            End If
            
            '#3680 如果該帳號找不到週期性項目,則不停用SO003與SO033 By Kin 2007/12/12
            '#3417 如果停用要連SO003與SO033的InvSeqNo一併清除 By Kin 2007/12/13
            '#5173  SO003一定要強迫停用,而SO033則還是要判斷,所以將If strCitemCode移至下面 By Kin 2009/07/06
'            If strCitemCode <> "" Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                "BANKCODE=NULL" & _
                                ",BANKNAME=NULL" & _
                                ",ACCOUNTNO=NULL" & _
                                ",InvSeqNo=NULL" & _
                                ",CMCODE=" & strCMCode & _
                                ",CMNAME='" & strCMName & "'" & _
                                ",PTCODE=" & strPTCode & _
                                ",PTNAME='" & strPTName & "'" & _
                                " WHERE CUSTID=" & gCustId & _
                                " AND ACCOUNTNO='" & strAccNo & "'" & _
                                " AND COMPCODE=" & gCompCode & _
                                 IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                 IIf(csmCitem1.GetQueryCode = Empty, "", " AND SEQNO IN (" & varSeqNo(i) & ")"), RetVal
                Debug.Print RetVal
            If strCitemCode <> "" Then
                '#3573 測試不OK,多增加CitemCode的條件,避免全部都清除 By Kin 2007/12/06
                '#5623 要將相同的BillNo再清除一次 By Kin 2010/04/02
                strBillNoSQL = "SELECT BILLNO,ServiceType FROM " & GetOwner & "SO033 " & _
                            " WHERE CUSTID=" & gCustId & _
                            " AND ACCOUNTNO='" & strAccNo & "'" & _
                            " AND COMPCODE=" & gCompCode & _
                            IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                            " AND UCCODE > 0 AND CANCELFLAG=0" & _
                            IIf(strCitemCode <> "", " AND CitemCode=" & strCitemCode, "") & _
                            " AND ROWNUM=1"
                If Not GetRS(rsBill, strBillNoSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                If Not rsBill.EOF Then
                    rsBill.MoveFirst
                    Do While Not rsBill.EOF
                        strBillNos = rsBill("BillNo") & ""
                        strServiceType = rsBill("ServiceType") & ""
                        gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                "BANKCODE=NULL" & _
                                ",BANKNAME=NULL" & _
                                ",ACCOUNTNO=NULL" & _
                                ",InvSeqNo=NULL" & _
                                ",CMCODE=" & strCMCode & _
                                ",CMNAME='" & strCMName & "'" & _
                                ",PTCODE=" & strPTCode & _
                                ",PTNAME='" & strPTName & "'" & _
                                " WHERE CUSTID=" & gCustId & _
                                " AND ACCOUNTNO='" & strAccNo & "'" & _
                                " AND COMPCODE=" & gCompCode & _
                                 IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                 " AND UCCODE > 0 AND CANCELFLAG=0" & _
                                 IIf(strCitemCode <> "", " AND CitemCode=" & strCitemCode, ""), RetVal
                                 
                        If strBillNos <> "" Then
                            strUCCodeSQL = "SELECT CODENO,Description FROM " & GetOwner & "CD013 " & _
                                        " WHERE CODENO = (SELECT UCCODE FROM " & GetOwner & "SO044 " & _
                                        " WHERE SERVICETYPE='" & strServiceType & "'" & _
                                        " AND COMPCODE= " & gCompCode & " )" & _
                                        " AND STOPFLAG<>1"
                            If Not GetRS(rsPara, strUCCodeSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                            If Not rsPara.EOF Then
                                strUCCode = rsPara("CODENO") & ""
                                strUCName = rsPara("DESCRIPTION") & ""
                            End If
                            '#5623 要將相同的BillNo再清除一次 By Kin 2010/04/02
                            gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                            "BANKCODE=NULL" & _
                                            ",BANKNAME=NULL" & _
                                            ",ACCOUNTNO=NULL" & _
                                            ",InvSeqNo=NULL" & _
                                            ",CMCODE=" & strCMCode & _
                                            ",CMNAME='" & strCMName & "'" & _
                                            ",PTCODE=" & strPTCode & _
                                            ",PTNAME='" & strPTName & "'" & _
                                            ",UCCODE=" & IIf(strUCCode = "", "UCcode", strUCCode) & _
                                            ",UCNAME='" & IIf(strUCName = "", "UCName", strUCName & "'") & _
                                            " WHERE CUSTID=" & gCustId & _
                                            " AND COMPCODE=" & gCompCode & _
                                            " AND UCCODE > 0 AND CANCELFLAG=0" & _
                                            " AND BILLNO='" & strBillNos & "'", RetVal
                        End If
                        rsBill.MoveNext
                    Loop
                    
                End If
            End If
        End If
    Next i
    Debug.Print RetVal
    StopSO002A strAccNo
    
    On Error Resume Next
    CloseRS rsTmp
    CloseRS rsPara
    CloseRS rsBill
  Exit Sub
ChkErr:
    ErrSub Me.Name, "StopGo"
End Sub
Private Function ChkInvData() As Boolean
  On Error GoTo ChkErr
    Dim bytCnt As Byte
    Dim strAccNo As String
    If Not rsSO106.EOF Then
        If rsSO106("AccountID").OriginalValue & "" = Empty Then
            strAccNo = txtAccountNo
        Else
            If EditMode = giEditModeInsert Then
                strAccNo = txtAccountNo
            Else
                strAccNo = rsSO106("AccountID").OriginalValue & ""
            End If
        End If
    Else
        strAccNo = txtAccountNo
    End If

    bytCnt = Val(RPxx(gcnGi.Execute("Select count(*) from " & GetOwner & "so002a where " & _
                      "AccountNo='" & strAccNo & "' and CustId=" & _
                       gCustId & " and compcode=" & gCompCode & " And StopFlag<>1").GetString))
    
    If EditMode = giEditModeEdit Then
        If (rsSO106("InvSeqNo") & "" <> "") And (Not IsNull(rsSO106("SnactionDate"))) Then ChkInvData = True
    End If
    If bytCnt > 0 Then ChkInvData = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "ChkInvData"
End Function
Private Sub ChkSO002A()
  On Error GoTo ChkErr
'    0 = 轉帳帳號
'    1 = 信用卡號
'    2 = 虛擬帳號
  
    If txtAccountNo = Empty Then Exit Sub
    Dim bytCnt As Byte
    Dim strAccNo As String
    Dim strBankCode As String
    Dim intPayID As Integer
    
    If gilCMCode.GetCodeNo <> Empty Then
        On Error Resume Next
        intCD031RefNo = GetRsValue("SELECT REFNO FROM " & GetOwner & "CD031 WHERE CODENO=" & gilCMCode.GetCodeNo)
        If Err.Number <> 0 Then
            Err.Clear
            intCD031RefNo = 0
        End If
    End If
    '#8724 Add refno=5
    If intCD031RefNo = 2 Or intCD031RefNo = 4 Or intCD031RefNo = 5 Then
        blnVAcc = False
    Else
        blnVAcc = True
    End If
    
    If blnVAcc Then
        intPayID = 2
    Else
        '#8724 add refno=5
        If intCD031RefNo = 4 Or intCD031RefNo = 5 Then
            If gilCardName.GetCodeNo <> Empty Then
                intPayID = 1
            Else
                intPayID = 0
            End If
        Else
            intPayID = 0
        End If
    End If
    
    With frmSO1100BMDI
        If Not rsSO106.EOF Then
            If rsSO106("AccountID").OriginalValue & "" = Empty Then
                strAccNo = txtAccountNo
                strBankCode = gilBankCode.GetCodeNo
            Else
                If EditMode = giEditModeInsert Then
                    strAccNo = txtAccountNo
                    strBankCode = gilBankCode.GetCodeNo
                Else
                    strAccNo = rsSO106("AccountID").OriginalValue & ""
                    strBankCode = rsSO106("BankCode").OriginalValue & ""
                End If
            End If
        Else
            strAccNo = txtAccountNo
        End If
        bytCnt = Val(RPxx(gcnGi.Execute("select count(*) from " & GetOwner & "so002a where " & _
                      "accountNo='" & strAccNo & "' and custid=" & _
                       gCustId & " and compcode=" & gCompCode).GetString))
        
        
'    txtChargeTitle = frmSO1100BMDI.txtCustName
'    txtInvTitle.Text = frmSO1100BMDI.rsSo002.Fields("InvTitle") & ""
'    mskInvNo.Text = frmSO1100BMDI.rsSo002.Fields("InvNo") & ""
'    txtInvAddress.Text = frmSO1100BMDI.rsSo002.Fields("InvAddress") & ""
'    cboInvoiceType = frmSO1100BMDI.ChildForm.cboInvoiceType
        
        If strAccNo = Empty Then strAccNo = txtAccountNo
        
        If bytCnt = 0 Then
            gcnGi.Execute "INSERT INTO " & GetOwner & "SO002A " & _
                            "(CUSTID,COMPCODE,BANKCODE,BANKNAME,ID,ACCOUNTNO," & _
                            "CARDNAME,CARDEXPDATE,CHARGEADDRNO,CHARGEADDRESS," & _
                            "MAILADDRNO,MAILADDRESS,CVC2,NOTE,CHARGETITLE," & _
                            "INVNO,INVTITLE,INVADDRESS,INVOICETYPE," & _
                            "CITEMSTR,CITEMSTR2,ADDCITEMACCOUNT) VALUES (" & _
                             gCustId & "," & gCompCode & "," & gilBankCode.GetCodeNo2 & ",'" & _
                             gilBankCode.GetDescription & "'," & intPayID & ",'" & strAccNo & "','" & _
                             gilCardName.GetDescription & "'," & IIf(Len(mskCardExpDate.Text) = 0, "Null", mskCardExpDate.Text) & "," & _
                            .rsSo001!ChargeAddrNo & ",'" & .rsSo001!ChargeAddress & "'," & _
                            .rsSo001!MailAddrNo & ",'" & .rsSo001!MailAddress & "','" & txtCVC2 & "','" & txtNote & "','" & txtAccountOwner & "','" & _
                             frmSO1100BMDI.rsSo002.Fields("InvNo") & "','" & frmSO1100BMDI.rsSo002.Fields("InvTitle") & "','" & _
                             frmSO1100BMDI.rsSo002.Fields("InvAddress") & "'," & frmSO1100BMDI.rsSo002.Fields("InvoiceType") & ",'" & _
                             csmCitem1.QryCD & "','" & csmCitem2.QryCD & "'," & chkAddAcc.Value & ")"
                             
            '*************************************************************************************************
            '#3436 新增時要連同發票資訊新增過去 By Kin 2007/12/13
            If strInvSeqNo <> "" Then
                If gcnGi.Execute("Select Count(*) From " & GetOwner & "SO002AD Where Compcode=" & gCompCode & _
                                    " And CustId=" & gCustId & " And AccountNo='" & strAccNo & "'" & _
                                    " And InvSeqNo=" & strInvSeqNo)(0) = 0 Then
                                    
                       gcnGi.Execute "INSERT INTO " & GetOwner & "SO002AD " & _
                                    "(CUSTID,AccountNo,CompCode,InvSeqNo) VALUES (" & _
                                    gCustId & ",'" & strAccNo & "'," & gCompCode & "," & strInvSeqNo & ")"
                End If
            End If
            '***************************************************************************************************
        ElseIf bytCnt = 1 Then
            ' MinChen 需求 , 修改時不 Update Address 相關欄位至 SO002A
            gcnGi.Execute "UPDATE " & GetOwner & "SO002A SET " & _
                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                            ",BANKNAME='" & gilBankCode.GetDescription & _
                            "',ID=" & intPayID & _
                            ",ACCOUNTNO='" & strRealAccNo & _
                            "',CARDNAME='" & gilCardName.GetDescription & _
                            "',CARDEXPDATE=" & IIf(Len(mskCardExpDate.Text) = 0, "Null", mskCardExpDate.Text) & _
                            ",CVC2='" & txtCVC2 & _
                            "',NOTE='" & txtNote & "'" & _
                            ",CITEMSTR='" & csmCitem1.QryCD & "'" & _
                            ",CITEMSTR2='" & csmCitem2.QryCD & "'" & _
                            ",ADDCITEMACCOUNT=" & chkAddAcc.Value & _
                            ",STOPFLAG=0,STOPDATE=NULL" & _
                            " WHERE CUSTID=" & gCustId & _
                            " AND ACCOUNTNO='" & strAccNo & _
                            "' AND COMPCODE=" & gCompCode & _
                            IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode)
                            
                            '#3929 同步更新SO002A時將ChargeTitle的條件取消 By Kin 2008/06/02
                            '" AND CHARGETITLE='" & txtAccountOwner & "'" & _

            '*************************************************************************************************
            '#3436 新增時要連同發票資訊新增過去 By Kin 2007/12/13
            If strInvSeqNo <> "" Then
                If gcnGi.Execute("Select Count(*) From " & GetOwner & "SO002AD Where Compcode=" & gCompCode & _
                                    " And CustId=" & gCustId & " And AccountNo='" & strRealAccNo & "'" & _
                                    " And InvSeqNo=" & strInvSeqNo)(0) = 0 Then
                                    
                                    
                    gcnGi.Execute "INSERT INTO " & GetOwner & "SO002AD " & _
                                    "(CUSTID,AccountNo,CompCode,InvSeqNo) VALUES (" & _
                                    gCustId & ",'" & strRealAccNo & "'," & gCompCode & "," & strInvSeqNo & ")"
                End If
            End If
            '***************************************************************************************************

'            gcnGi.Execute "UPDATE " & GetOwner & "SO002A SET " & _
                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                            ",BANKNAME='" & gilBankCode.GetDescription & _
                            "',ID=" & intPayID & _
                            ",ACCOUNTNO='" & strRealAccNo & _
                            "',CARDNAME='" & gilCardName.GetDescription & _
                            "',CARDEXPDATE=" & IIf(Len(mskCardExpDate.Text) = 0, "Null", mskCardExpDate.Text) & _
                            ",CHARGEADDRNO=" & .rsSo001!ChargeAddrNo & _
                            ",CHARGEADDRESS='" & .rsSo001!ChargeAddress & _
                            "',MAILADDRNO=" & .rsSo001!MailAddrNo & _
                            ",MAILADDRESS='" & .rsSo001!MailAddress & _
                            "',CVC2='" & txtCVC2 & _
                            "',NOTE='" & txtNote & "'" & _
                            ",CITEMSTR='" & csmCitem1.QryCD & "'" & _
                            ",CITEMSTR2='" & csmCitem2.QryCD & "'" & _
                            ",ADDCITEMACCOUNT=" & chkAddAcc.Value & _
                            ",STOPFLAG=0,STOPDATE=NULL" & _
                            " WHERE CUSTID=" & gCustId & _
                            " AND ACCOUNTNO='" & strAccNo & _
                            "' AND COMPCODE=" & gCompCode & _
                            " AND CHARGETITLE='" & txtAccountOwner & "'" & _
                            IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode)
        End If
    End With
    
    Dim strChild2A As String
    strChild2A = Get2AChild
    
    If strChild2A <> Empty And strChild2A <> "''" Then
        gcnGi.Execute "UPDATE " & GetOwner & "SO002A SET " & _
                        "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                        ",BANKNAME='" & gilBankCode.GetDescription & _
                        "',ID=" & intPayID & _
                        ",ACCOUNTNO='" & strRealAccNo & _
                        "',CARDNAME='" & gilCardName.GetDescription & _
                        "',CARDEXPDATE=" & IIf(Len(mskCardExpDate.Text) = 0, "Null", mskCardExpDate.Text) & _
                        ",CVC2='" & txtCVC2 & _
                        "',NOTE='" & txtNote & "'" & _
                        ",CITEMSTR='" & csmCitem1.QryCD & "'" & _
                        ",CITEMSTR2='" & csmCitem2.QryCD & "'" & _
                        ",ADDCITEMACCOUNT=" & chkAddAcc.Value & _
                        ",STOPFLAG=0,STOPDATE=NULL" & _
                        " WHERE INHERITKEY IN (" & strChild2A & ")" & _
                        " AND INHERITFLAG=1" & _
                        " AND COMPCODE=" & gCompCode
    End If
    
    '(AccountNo=SO106. AccountID and CustId=SO106.CustId and CompCode=SO106.CompCode).
    'If 沒有該帳號資料 Then Insert into SO002A
    '
    'PS: 新增至SO002A時, 收費地址與郵寄地址部分則由基本資料中帶入當預設值.
    'PS:  不管是否要更新至SO002皆要執行客戶帳號Insert into流程
    '
    '勾選停用部分 : 原機制只做紀錄而無其他處理機制.
    '現調整為要刪除SO002A中之該筆帳號/卡號,
    '以避免User再次誤選到該帳號.
    '
    '原週期性A項之項目有用到該帳號的項目 , 則將它們 update成
    '
    'SO002中之主帳號相關資料並show Message告知User此狀況,
    '
    '提供User判斷是否還要至週期性頁籤中調整資料內容.
    '
    '以上兩個機制串起來後 , SO002A則成為客戶的有效帳戶管理, 以方便多帳戶機制使用
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ChkSO002A"
End Sub

Private Sub ChkSO003()
  On Error GoTo ChkErr
    If txtAccountNo = Empty Then Exit Sub
    Dim strAccNo As String
    Dim strBankCode As String
    Dim strSeqNo As String
    Dim strQryCnt As String
    Dim blnUpdINV As Boolean
    blnUpdINV = True
    With frmSO1100BMDI
        If EditMode = giEditModeInsert Then
            strAccNo = txtAccountNo
            strBankCode = gilBankCode.GetCodeNo
            strSeqNo = csmCitem1.GetQueryCode
        
        Else
            If Not rsSO106.EOF Then
                If rsSO106("AccountID").OriginalValue & "" = Empty Then
                    strAccNo = txtAccountNo
                    strBankCode = gilBankCode.GetCodeNo
    '                strCitemCode = gimCitem.GetQueryCode
                    strSeqNo = csmCitem1.GetQueryCode
                Else
                    strAccNo = rsSO106("AccountID").OriginalValue & ""
                    strBankCode = rsSO106("BankCode").OriginalValue & ""
                    strSeqNo = rsSO106("CitemStr").OriginalValue & ""
                End If
            Else
                strAccNo = txtAccountNo
                strBankCode = gilBankCode.GetCodeNo
    '            strCitemCode = gimCitem.GetQueryCode
                strSeqNo = csmCitem1.GetQueryCode
            End If
        End If
        
        Dim RetVal As Integer
        Dim strChildCustID As String
        '***********************************************************************************************
        '#5045 如果有發票抬頭不一樣的資料詢問User是否回填,選擇否則直接跳離開 By Kin 2009/05/04
        strQryCnt = "Select Count(*) Cnt From " & GetOwner & " SO003 " & _
                    " Where INVSEQNO IS Not Null And INVSEQNO<>" & strInvSeqNo & _
                    " And CustID=" & gCustId & " And CompCode=" & gCompCode & _
                    " And ACCOUNTNO='" & strAccNo & "'" & _
                    " And CMCODE=" & gilCMCode.GetCodeNo & _
                    IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                    IIf(strSeqNo = Empty, "", " AND SEQNO IN (" & strSeqNo & ")") & _
                    " And PTCODE=" & gilPTCode.GetCodeNo
        '***********************************************************************************************
        If Val(gcnGi.Execute(strQryCnt)("Cnt")) > 0 Then
            If MsgBox("本帳戶所指定的發票抬頭與週期性收費項目不符，是否回填？", vbQuestion + vbYesNo, "訊息") = vbNo Then
                blnUpdINV = False
            End If
        End If
        
        strChildCustID = GetChild
        
        If EditMode = giEditModeEdit Then
'            strCitemCode = gimCitem.GetQueryCode
            strSeqNo = csmCitem1.GetQueryCode
'            If gimCitem.GetQueryCode = Empty And rsSO106("CitemStr").OriginalValue <> Empty Then
            If csmCitem1.GetQueryCode = Empty And rsSO106("CitemStr").OriginalValue <> Empty Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                "BANKCODE=NULL" & _
                                ",BANKNAME=NULL" & _
                                ",ACCOUNTNO=NULL" & _
                                IIf(blnUpdINV = True, ",InvSeqNO=NULL", "") & _
                                ",CMCODE=" & frmSO1100BMDI.rsSO044!CMcode & _
                                ",CMNAME='" & GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [收費方式代碼檔] 查無代碼為 " & frmSO1100BMDI.rsSO044!CMcode & " 的資料 !") & "'" & _
                                " WHERE CUSTID=" & gCustId & _
                                " AND ACCOUNTNO='" & strAccNo & "'" & _
                                " AND COMPCODE=" & gCompCode & _
                                 IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                 IIf(strSeqNo = Empty, "", " AND SEQNO IN (" & rsSO106("CitemStr").OriginalValue & ")"), RetVal
                Debug.Print RetVal
                If strChildCustID <> Empty Then
                    gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                    "BANKCODE=NULL" & _
                                    ",BANKNAME=NULL" & _
                                    ",ACCOUNTNO=NULL" & _
                                    IIf(blnUpdINV = True, ",InvSeqNo=NULL", "") & _
                                    ",CMCODE=" & frmSO1100BMDI.rsSO044!CMcode & _
                                    ",CMNAME='" & GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [收費方式代碼檔] 查無代碼為 " & frmSO1100BMDI.rsSO044!CMcode & " 的資料 !") & "'" & _
                                    " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                    " AND ACCOUNTNO='" & strAccNo & "'" & _
                                    " AND COMPCODE=" & gCompCode & _
                                     IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                     IIf(strSeqNo = Empty, "", " AND SEQNO IN (" & rsSO106("CitemStr").OriginalValue & ")"), RetVal
                    Debug.Print RetVal
                End If
            End If
'            If gimCitem.GetQueryCode <> Empty And rsSO106("CitemStr").OriginalValue <> Empty Then
            If csmCitem1.GetQueryCode <> Empty And rsSO106("CitemStr").OriginalValue <> Empty Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                "BANKCODE=NULL" & _
                                ",BANKNAME=NULL" & _
                                ",ACCOUNTNO=NULL" & _
                                IIf(blnUpdINV = True, ",InvSeqNo=NULL", "") & _
                                ",CMCODE=" & frmSO1100BMDI.rsSO044!CMcode & _
                                ",CMNAME='" & GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [收費方式代碼檔] 查無代碼為 " & frmSO1100BMDI.rsSO044!CMcode & " 的資料 !") & "'" & _
                                " WHERE CUSTID=" & gCustId & _
                                " AND ACCOUNTNO='" & strAccNo & "'" & _
                                " AND COMPCODE=" & gCompCode & _
                                 IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                 IIf(strSeqNo = Empty, "", " AND SEQNO IN (" & rsSO106("CitemStr").OriginalValue & "" & ")"), RetVal
                Debug.Print RetVal
                If strChildCustID <> Empty Then
                    gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                    "BANKCODE=NULL" & _
                                    ",BANKNAME=NULL" & _
                                    ",ACCOUNTNO=NULL" & _
                                    IIf(blnUpdINV = True, ",InvSeqNo=NULL", "") & _
                                    ",CMCODE=" & frmSO1100BMDI.rsSO044!CMcode & _
                                    ",CMNAME='" & GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [收費方式代碼檔] 查無代碼為 " & frmSO1100BMDI.rsSO044!CMcode & " 的資料 !") & "'" & _
                                    " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                    " AND ACCOUNTNO='" & strAccNo & "'" & _
                                    " AND COMPCODE=" & gCompCode & _
                                     IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                     IIf(strSeqNo = Empty, "", " AND SEQNO IN (" & rsSO106("CitemStr").OriginalValue & "" & ")"), RetVal
                    Debug.Print RetVal
                End If
            Else
                If rsSO106("BankCode").OriginalValue & "" <> gilBankCode.GetCodeNo & "" Then
                    If rsSO106("AccountID").OriginalValue & "" = strRealAccNo Then
                        gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                        "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                        ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                        ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                        ",CMCODE=" & gilCMCode.GetCodeNo & _
                                        ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                        ",PTCODE=" & gilPTCode.GetCodeNo & _
                                        ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                        IIf(blnUpdINV = True, IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, ""), "") & _
                                        " WHERE CUSTID=" & gCustId & _
                                        " AND COMPCODE=" & gCompCode & _
                                        " AND BANKCODE=" & rsSO106("BankCode").OriginalValue, RetVal
                        Debug.Print RetVal
                        If strChildCustID <> Empty Then
                            gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                            ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                            ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                            ",CMCODE=" & gilCMCode.GetCodeNo & _
                                            ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                            ",PTCODE=" & gilPTCode.GetCodeNo & _
                                            ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                            IIf(blnUpdINV = True, IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, ""), "") & _
                                            " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                            " AND COMPCODE=" & gCompCode & _
                                            " AND BANKCODE=" & rsSO106("BankCode").OriginalValue, RetVal
                            Debug.Print RetVal
                        End If
                    Else
                        gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                        "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                        ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                        ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                        ",CMCODE=" & gilCMCode.GetCodeNo & _
                                        ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                        ",PTCODE=" & gilPTCode.GetCodeNo & _
                                        ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                        IIf(blnUpdINV = True, IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, ""), "") & _
                                        " WHERE CUSTID=" & gCustId & _
                                        " AND COMPCODE=" & gCompCode & _
                                        " AND BANKCODE=" & rsSO106("BankCode").OriginalValue & _
                                        " AND ACCOUNTNO='" & rsSO106("AccountID").OriginalValue & "'", RetVal
                        Debug.Print RetVal
                        If strChildCustID <> Empty Then
                            gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                            ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                            ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                            ",CMCODE=" & gilCMCode.GetCodeNo & _
                                            ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                            ",PTCODE=" & gilPTCode.GetCodeNo & _
                                            ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                           IIf(blnUpdINV = True, IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, ""), "") & _
                                            " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                            " AND COMPCODE=" & gCompCode & _
                                            " AND BANKCODE=" & rsSO106("BankCode").OriginalValue & _
                                            " AND ACCOUNTNO='" & rsSO106("AccountID").OriginalValue & "'", RetVal
                            Debug.Print RetVal
                        End If
                    End If
                Else
                    If rsSO106("AccountID").OriginalValue & "" <> strRealAccNo Then
                        gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                        "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                        ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                        ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                        ",CMCODE=" & gilCMCode.GetCodeNo & _
                                        ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                        ",PTCODE=" & gilPTCode.GetCodeNo & _
                                        ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                        IIf(blnUpdINV = True, IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, ""), "") & _
                                        " WHERE CUSTID=" & gCustId & _
                                        " AND COMPCODE=" & gCompCode & _
                                        " AND ACCOUNTNO='" & rsSO106("AccountID").OriginalValue & "'", RetVal
                        Debug.Print RetVal
                        If strChildCustID <> Empty Then
                            gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                            ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                            ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                            ",CMCODE=" & gilCMCode.GetCodeNo & _
                                            ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                            ",PTCODE=" & gilPTCode.GetCodeNo & _
                                            ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                            IIf(blnUpdINV = True, IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, ""), "") & _
                                            " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                            " AND COMPCODE=" & gCompCode & _
                                            " AND ACCOUNTNO='" & rsSO106("AccountID").OriginalValue & "'", RetVal
                            Debug.Print RetVal
                        End If
                    End If
                End If
            End If
        End If
        
'        If gimCitem.GetQueryCode <> Empty Then
        If csmCitem1.GetQueryCode <> Empty Then
'            strCitemCode = gimCitem.GetQueryCode
            strSeqNo = csmCitem1.GetQueryCode
            gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                            ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                            ",ACCOUNTNO='" & strRealAccNo & "'" & _
                            ",CMCODE=" & gilCMCode.GetCodeNo & _
                            ",CMNAME='" & gilCMCode.GetDescription & "' " & _
                            ",PTCODE=" & gilPTCode.GetCodeNo & _
                            ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                            IIf(blnUpdINV = True, IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, ""), "") & _
                            " WHERE CUSTID=" & gCustId & _
                            " AND COMPCODE=" & gCompCode & _
                             IIf(strSeqNo = Empty, "", " AND SEQNO IN (" & strSeqNo & ")"), RetVal
            Debug.Print RetVal
            If strChildCustID <> Empty Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                ",CMCODE=" & gilCMCode.GetCodeNo & _
                                ",CMNAME='" & gilCMCode.GetDescription & "' " & _
                                ",PTCODE=" & gilPTCode.GetCodeNo & _
                                ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                IIf(blnUpdINV = True, IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, ""), "") & _
                                " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                " AND COMPCODE=" & gCompCode & _
                                 IIf(strSeqNo = Empty, "", " AND SEQNO IN (" & strSeqNo & ")"), RetVal
                Debug.Print RetVal
            End If
        End If
    End With
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ChkSO003"
End Sub

Private Function GetChild() As String
  On Error Resume Next
    If rsSO106("INHERITKEY") & "" <> Empty And rsSO106("INHERITFLAG") = 0 Then
        GetChild = gcnGi.Execute("SELECT DISTINCT CUSTID FROM " & GetOwner & "SO106" & _
                                                 " WHERE INHERITFLAG=1 AND INHERITKEY='" & rsSO106("INHERITKEY") & "'").GetString(2, , "", ",", "")
        If Err Then
            Err.Clear
            GetChild = ""
        Else
            GetChild = Replace(GetChild, ",,", ",", 1)
            If Right(GetChild, 1) = "," Then GetChild = Left(GetChild, Len(GetChild) - 1)
        End If
    End If
End Function

Private Sub ChkSO033()
  On Error GoTo ChkErr
    If txtAccountNo = Empty Then Exit Sub
    Dim strAccNo As String
    Dim strBankCode As String
    Dim strRowId As String

    With frmSO1100BMDI
    
        
    
        If Not rsSO106.EOF Then
            If rsSO106("AccountID").OriginalValue & "" = Empty Then
                strAccNo = txtAccountNo
                strBankCode = gilBankCode.GetCodeNo
                strRowId = csmCitem2.GetQryFld(11)
            Else
                strAccNo = rsSO106("AccountID").OriginalValue & ""
                strBankCode = rsSO106("BankCode").OriginalValue & ""
'               strROWID = rsSO106("CitemStr2").OriginalValue & ""
                strRowId = csmCitem2.GetQryFld(11)
            End If
        Else
            strAccNo = txtAccountNo
            strBankCode = gilBankCode.GetCodeNo
            strRowId = csmCitem2.GetQryFld(11)
'            strROWID = csmCitem2.GetQueryCode
        End If
        
        Dim RetVal As Integer
        Dim strChildCustID As String
        strChildCustID = GetChild
        
        If EditMode = giEditModeEdit Then
            strRowId = csmCitem2.GetQryFld(11)
            If csmCitem2.GetQueryCode = Empty And rsSO106("CitemStr2").OriginalValue <> Empty Then
                '#4023 如果RowId是空值,則不Update SO033,並且將SO106.CitemStr2清除 By Kin 2008/07/25
                If strRowId <> Empty Then
                    gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                    "BANKCODE=NULL" & _
                                    ",BANKNAME=NULL" & _
                                    ",ACCOUNTNO=NULL" & _
                                    ",InvSeqNo=NULL" & _
                                    ",CMCODE=" & frmSO1100BMDI.rsSO044!CMcode & _
                                    ",CMNAME='" & GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [收費方式代碼檔] 查無代碼為 " & frmSO1100BMDI.rsSO044!CMcode & " 的資料 !") & "'" & _
                                    " WHERE CUSTID=" & gCustId & _
                                    " AND ACCOUNTNO='" & strAccNo & "'" & _
                                    " AND COMPCODE=" & gCompCode & _
                                     IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                     IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & ")") & _
                                    IIf(blnUCCodeIsNull, "", " AND UCCODE > 0 ") & " AND CANCELFLAG=0", RetVal
                                     Debug.Print RetVal
                End If
                If strChildCustID <> Empty Then
                    If strRowId <> Empty Then
                        gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                        "BANKCODE=NULL" & _
                                        ",BANKNAME=NULL" & _
                                        ",ACCOUNTNO=NULL" & _
                                        ",InvSeqNo=NULL" & _
                                        ",CMCODE=" & frmSO1100BMDI.rsSO044!CMcode & _
                                        ",CMNAME='" & GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [收費方式代碼檔] 查無代碼為 " & frmSO1100BMDI.rsSO044!CMcode & " 的資料 !") & "'" & _
                                        " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                        " AND ACCOUNTNO='" & strAccNo & "'" & _
                                        " AND COMPCODE=" & gCompCode & _
                                         IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                         IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & ")") & _
                                        IIf(blnUCCodeIsNull, "", " AND UCCODE > 0 ") & " AND CANCELFLAG=0", RetVal
                                         Debug.Print RetVal
                    End If
                End If
            End If
            If csmCitem2.GetQueryCode <> Empty And rsSO106("CitemStr2").OriginalValue <> Empty Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                "BANKCODE=NULL" & _
                                ",BANKNAME=NULL" & _
                                ",ACCOUNTNO=NULL" & _
                                ",InvSeqNo=NULL" & _
                                ",CMCODE=" & frmSO1100BMDI.rsSO044!CMcode & _
                                ",CMNAME='" & GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [收費方式代碼檔] 查無代碼為 " & frmSO1100BMDI.rsSO044!CMcode & " 的資料 !") & "'" & _
                                " WHERE CUSTID=" & gCustId & _
                                " AND ACCOUNTNO='" & strAccNo & "'" & _
                                " AND COMPCODE=" & gCompCode & _
                                 IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                 IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & "" & ")") & _
                                 IIf(blnUCCodeIsNull, "", " AND UCCODE > 0 ") & " AND CANCELFLAG=0", RetVal
                                 Debug.Print RetVal
                If strChildCustID <> Empty Then
                    gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                    "BANKCODE=NULL" & _
                                    ",BANKNAME=NULL" & _
                                    ",ACCOUNTNO=NULL" & _
                                    ",InvSeqNo=NULL" & _
                                    ",CMCODE=" & frmSO1100BMDI.rsSO044!CMcode & _
                                    ",CMNAME='" & GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [收費方式代碼檔] 查無代碼為 " & frmSO1100BMDI.rsSO044!CMcode & " 的資料 !") & "'" & _
                                    " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                    " AND ACCOUNTNO='" & strAccNo & "'" & _
                                    " AND COMPCODE=" & gCompCode & _
                                     IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                     IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & "" & ")") & _
                                     IIf(blnUCCodeIsNull, "", " AND UCCODE > 0 ") & " AND CANCELFLAG=0", RetVal
                                     Debug.Print RetVal
                End If
            Else
                '#5303 如果沒有選收費資料,會誤UPD相關欄位,增加stRowId判斷 By Kin 2009/09/29
                If strRowId <> Empty Then
                    If rsSO106("BankCode").OriginalValue & "" <> gilBankCode.GetCodeNo & "" Then
                        If rsSO106("AccountID").OriginalValue & "" = strRealAccNo Then
                            gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                            ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                            ",ACCOUNTNO='" & txtAccountNo & "'" & _
                                            ",CMCODE=" & gilCMCode.GetCodeNo & _
                                            ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                            ",PTCODE=" & gilPTCode.GetCodeNo & _
                                            ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                            IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, "") & _
                                            " WHERE CUSTID=" & gCustId & _
                                            " AND COMPCODE=" & gCompCode & _
                                            " AND BANKCODE=" & rsSO106("BankCode").OriginalValue & _
                                            IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & "" & ")"), RetVal
                            Debug.Print RetVal
                            If strChildCustID <> Empty Then
                                gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                                "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                                ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                                ",ACCOUNTNO='" & txtAccountNo & "'" & _
                                                ",CMCODE=" & gilCMCode.GetCodeNo & _
                                                ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                                ",PTCODE=" & gilPTCode.GetCodeNo & _
                                                ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                                IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, "") & _
                                                " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                                " AND COMPCODE=" & gCompCode & _
                                                " AND BANKCODE=" & rsSO106("BankCode").OriginalValue & _
                                                IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & "" & ")"), RetVal
                                Debug.Print RetVal
                            End If
                        Else
                            gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                            ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                            ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                            ",CMCODE=" & gilCMCode.GetCodeNo & _
                                            ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                            ",PTCODE=" & gilPTCode.GetCodeNo & _
                                            ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                            IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, "") & _
                                            " WHERE CUSTID=" & gCustId & _
                                            " AND COMPCODE=" & gCompCode & _
                                            " AND BANKCODE=" & rsSO106("BankCode").OriginalValue & _
                                            " AND ACCOUNTNO='" & rsSO106("AccountID").OriginalValue & "'" & _
                                            IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & "" & ")"), RetVal
                            Debug.Print RetVal
                            If strChildCustID <> Empty Then
                                gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                                "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                                ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                                ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                                ",CMCODE=" & gilCMCode.GetCodeNo & _
                                                ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                                ",PTCODE=" & gilPTCode.GetCodeNo & _
                                                ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                                IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, "") & _
                                                " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                                " AND COMPCODE=" & gCompCode & _
                                                " AND BANKCODE=" & rsSO106("BankCode").OriginalValue & _
                                                " AND ACCOUNTNO='" & rsSO106("AccountID").OriginalValue & "'" & _
                                                IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & "" & ")"), RetVal
                                Debug.Print RetVal
                            End If
                        End If
                    Else
                        If rsSO106("AccountID").OriginalValue & "" <> strRealAccNo Then
                            gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                            ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                            ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                            ",CMCODE=" & gilCMCode.GetCodeNo & _
                                            ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                            ",PTCODE=" & gilPTCode.GetCodeNo & _
                                            ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                            IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, "") & _
                                            " WHERE CUSTID=" & gCustId & _
                                            " AND COMPCODE=" & gCompCode & _
                                            " AND ACCOUNTNO='" & rsSO106("AccountID").OriginalValue & "'" & _
                                            IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & "" & ")"), RetVal
                            Debug.Print RetVal
                            If strChildCustID <> Empty Then
                                gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                                "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                                ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                                ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                                ",CMCODE=" & gilCMCode.GetCodeNo & _
                                                ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                                ",PTCODE=" & gilPTCode.GetCodeNo & _
                                                ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                                IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, "") & _
                                                " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                                " AND COMPCODE=" & gCompCode & _
                                                " AND ACCOUNTNO='" & rsSO106("AccountID").OriginalValue & "'" & _
                                                IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & "" & ")"), RetVal
                                Debug.Print RetVal
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        If csmCitem2.GetQueryCode <> Empty Then
            gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                            ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                            ",ACCOUNTNO='" & strRealAccNo & "'" & _
                            ",CMCODE=" & gilCMCode.GetCodeNo & _
                            ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                            ",PTCODE=" & gilPTCode.GetCodeNo & _
                            ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                            IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, "") & _
                            " WHERE CUSTID=" & gCustId & _
                            " AND COMPCODE=" & gCompCode & _
                             IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & ")") & _
                             IIf(blnUCCodeIsNull, "", " AND UCCODE > 0 ") & " AND CANCELFLAG=0", RetVal
                             Debug.Print RetVal
            If strChildCustID <> Empty Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                ",CMCODE=" & gilCMCode.GetCodeNo & _
                                ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                ",PTCODE=" & gilPTCode.GetCodeNo & _
                                ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, "") & _
                                " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                " AND COMPCODE=" & gCompCode & _
                                 IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & ")") & _
                                IIf(blnUCCodeIsNull, "", " AND UCCODE > 0 ") & " AND CANCELFLAG=0", RetVal
                                 Debug.Print RetVal
            End If
        End If
    End With
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ChkSO033"
End Sub

Private Sub StopSO002A(Optional strAccNo As String = "")
  On Error Resume Next
    Dim strChkSameAcc As String
    Dim strStopDate As String
    If strAccNo = Empty Then strAccNo = strRealAccNo
    If strAccNo = Empty Then Exit Sub
    If gdaStopDate.GetValue = Empty Then gdaStopDate.SetValue RightDate
    
    '#3542 新增時會自動填入停用日期，將判斷移至上面，並用變數記錄，使用Update語法時，用變數帶入 By Kin 2007/09/29
    strStopDate = gdaStopDate.GetValue
    If chkStop.Value <> 1 Then gdaStopDate.SetValue ""
    
    '********************************************************************************************************
    '#3324 檢查是否有相同帳號用於其他服務之收費資料中，如果有其它資料，不停用SO002A的資料 By Kin 2007/08/01
    If Not rsSO106.EOF Then
        strChkSameAcc = "Select Count(*) Cnt From " & GetOwner & "SO106 Where " & _
                        "AccountID='" & strAccNo & "' And " & _
                        "CompCode=" & gCompCode & " And " & _
                        "CUSTID=" & gCustId & " And " & _
                        "StopFlag=0 And StopDate Is Null" & _
                        IIf(lngEditMode = giEditModeInsert, "", " And RowID<>'" & rsSO106("RowId") & "'")
        If gcnGi.Execute(strChkSameAcc)(0) > 0 Then Exit Sub
    End If
    '*********************************************************************************************************
'   gcnGi.Execute ("delete from " & GetOwner & "so002a where AccountNo='" & strAccNo & "' and custid=" & gCustId & " and compcode=" & gCompCode)
    gcnGi.Execute ("UPDATE " & GetOwner & "SO002A" & _
                                " SET STOPFLAG=1,STOPDATE=TO_DATE('" & strStopDate & "','YYYY/MM/DD')" & _
                                " WHERE ACCOUNTNO='" & strAccNo & "'" & _
                                " AND CUSTID=" & gCustId & _
                                " AND COMPCODE=" & gCompCode)
    gimCitem.Clear
'    If chkStop.Value <> 1 Then gdaStopDate.SetValue ""
    
    ' 父帳戶如已停用時，則子帳戶需連動停用處理並將父子關聯清除
    ' (即InheritFlag=0, .InheritKey=Null)。
    ' (父子關聯條件: 子帳戶.InheritFlag=1 and 子帳戶.InheritKey=父帳戶.InheritKey)
    If rsSO106("InheritKey") & "" <> Empty And rsSO106("InheritFlag") = 0 Then
    
        gcnGi.Execute "UPDATE " & GetOwner & "SO106" & _
                                " SET INHERITFLAG=0,INHERITKEY=NULL" & _
                                " WHERE ACCOUNTNO='" & strAccNo & "'" & _
                                " AND INHERITKEY='" & rsSO106("InheritKey") & "'" & _
                                " AND INHERITFLAG=1" & _
                                " AND COMPCODE=" & gCompCode
                                
        Dim strChild2A As String
        strChild2A = Get2AChild
        If strChild2A <> Empty Then
            gcnGi.Execute "UPDATE " & GetOwner & "SO002A" & _
                                    " SET INHERITFLAG=0,INHERITKEY=NULL" & _
                                    " WHERE INHERITKEY IN (" & strChild2A & ")" & _
                                    " AND INHERITFLAG=1" & _
                                    " AND COMPCODE=" & gCompCode
        End If
    End If
End Sub

Private Function Get2AChild() As String
  On Error Resume Next
    If rsSO106("INHERITKEY") & "" <> Empty And rsSO106("INHERITFLAG") = 0 Then
        Get2AChild = gcnGi.Execute("SELECT INHERITKEY FROM " & GetOwner & "SO002A" & _
                                                        " WHERE INHERITFLAG=1" & _
                                                        " AND CUSTID=" & gCustId & _
                                                        " AND COMPCODE=" & gCompCode & _
                                                        " AND ACCOUNTNO='" & rsSO106("AccountID") & "'").GetString(2, , "", "','", "")
        If Err Then
            Err.Clear
            Get2AChild = ""
        Else
            If Right(Get2AChild, 2) = ",'" Then Get2AChild = Left(Get2AChild, Len(Get2AChild) - 2)
            If Left(Get2AChild, 2) = "'," Then Get2AChild = Mid(Get2AChild, 3)
            If Left(Get2AChild, 1) <> "'" Then Get2AChild = "'" & Get2AChild
        End If
    End If
End Function

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
    Set ObjActiveForm = Me
    strActFrmName = "frmSO1100F"
    If frmSO114FA.frmShow Then
        frmSO114FA.SetFocus
    End If
    'strInvSeqNo = frmSO114FA.uInvSeqNo
    Screen.MousePointer = vbDefault
    Set ObjActiveForm = Me
    strActFrmName = "frmSO1100G"
    If rsSO106.RecordCount <> frmSO1100BMDI.rsSO106.RecordCount Then OpenData
    
End Sub

Private Sub Form_KeyPress(keyAscii As Integer)
  On Error Resume Next
   If keyAscii = Asc("'") Then keyAscii = 0
End Sub
Public Property Let Set106Inv(ByVal vData As String)
    strInvSeqNo = vData
    lblInvSeqNo = strInvSeqNo
End Property
Private Sub Form_Load()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    Dim lngAbsPos As Long
    Dim blnDefCustStatusFaci As Boolean
    posAchContition = " PRGNAME = 'X'  "
    strInvSeqNo = Empty
    InitializingListOcx
    blnStartPost = False
    If Crm Then
        PolyFrmFunction Me
        Me.Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2)
    Else
        'If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 + 300
        If Not 800600 Then Me.Move (Screen.Width - Me.Width) / 2, frmSO1100BMDI.Top + frmSO1100BMDI.tbrMenu.Height + 700
    End If
    'Me.Move (Screen.Width - Me.Width) / 2, frmSO1100BMDI.Top + frmSO1100BMDI.tbrMenu.Height + 600
    If EditMode <> giEditModeInsert Then
        lngAbsPos = frmSO1100BMDI.rsSO106.AbsolutePosition
        OpenData
        If lngAbsPos > 0 Then frmSO1100BMDI.rsSO106.AbsolutePosition = lngAbsPos
    End If
    
    With mFlds
        .Add "AcceptName", , , , , "受理人員 "
        .Add "AcceptTime", giControlTypeDate, , , , "受理時間 "
        .Add "Proposer", , , , , "申請人   "
        .Add "ID", , , , , "身份證字號"
        .Add "BankName", , , , , "銀行名稱" & Space(18)
        .Add "CardName", , , , , "信用卡別"
        .Add "StopYM", , , "00/0000", , "卡有效期"
        .Add "AccountID", , , , , "帳號/卡號" & Space(12)
        .Add "AccountName", , , , , "帳號所有人"
        .Add "AccountNameID", , , , , "帳號所有人身份証號"
        .Add "PropDate", giControlTypeDate, , , , "申請日期"
        .Add "SendDate", giControlTypeDate, , , , "送件日期"
        .Add "SnactionDate", giControlTypeDate, , , , "核准日期"
        .Add "StopFlag", , , , , "停用"
        .Add "StopDate", giControlTypeDate, , , , "停用日期"
        .Add "MediaName", , , , , "介紹媒介"
        .Add "IntroName", , , , , "介紹人"
        .Add "Note", , , , , "備註"

'        .Add "AccountID", , , , , "信用卡別"
'        .Add "CardExpDate", , , , , "信用卡期限"
    End With
    With ggrData
        .UseCellForeColor = True
        .AllFields = mFlds
        .SetHead
         Set .Recordset = rsSO106
        .Refresh
    End With
    '**********************************************************************************************
    '#4089 增加判斷參數，如果為0要再判斷是否可顯示ACH歷史明細，如果為1則要強迫隱藏 By Kin 2008/09/11
    blnDefCustStatusFaci = Val(GetRsValue("Select Nvl(DefCustStatusFaci,0) From " & GetOwner & _
                            "SO041 Where SysID=" & gCompCode, gcnGi)) - 1
    '**********************************************************************************************
    
    '#3585 判斷ACHCustID是否為1,如果為1,新增資料時，要寫入Ach用戶號碼 By Kin 2007/10/26
    'If blnDefCustStatusFaci Then
        blnACHCustID = Val(GetRsValue("Select Nvl(ACHCustID,0) From " & GetOwner & _
                                "SO041 Where SysID=" & gCompCode, gcnGi))
    'Else
    '    blnACHCustID = False
    'End If
    '#3925 SO041.ACHCUSTID=0 ACH歷程按鈕要Disable By Kin 2008/05/08
    If Val(GetRsValue("Select Nvl(StartPost,0) From " & GetOwner & "SO041 Where SysID = " & gCompCode, gcnGi) & "") = 1 Then
        posAchContition = " PRGNAME like '%POST%' "
        AchTypeCondition = "  In (1, 2 ) "
        blnStartPost = True
    Else
        posAchContition = " PRGNAME = 'X'  "
        AchTypeCondition = " In (1) "
    End If
    cmdACHDetail.Enabled = blnACHCustID And blnDefCustStatusFaci
    cmdACHDetail.Visible = blnACHCustID And blnDefCustStatusFaci
    If Not blnACHCustID Then
        lblState.Move lblApplication.Left
        lblStatus.Move lblApplication.Left + lblState.Width + 50
    End If
    '#5045 增加參數判斷是否可挑選已收費的資料 By Kin 2009/05/05
    blnUCCodeIsNull = GetRsValue("Select Nvl(HItemAutoUpd,0) From " & GetOwner & "SO041", gcnGi) = 1
    '#3454 將SO001.CustId與SO994.DeclarantName 埧入ComboBox By Kin 2007/12/21
    Call ShowProposer
    Call ChangeMode(giEditModeView)
'    個資法
'    If Len(strAccountName) > 0 Then
'        If strAccountName <> ggrData.Recordset.Fields("Proposer") Then
'            cboProposer.Text = "XXXXX"
'            gilCMCode.SetCodeNo "X"
'            gilCMCode.SetDescription "XXX"
'            gilPTCode.SetCodeNo "X"
'            gilPTCode.SetDescription "XXX"
'            txtAccountNo.PasswordChar = "X"
'            txtAccountOwner.PasswordChar = "X"
'            txtHide.PasswordChar = "X"
'        End If
'    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub OpenData()
  On Error GoTo ChkErr
    Set rsSO106 = New ADODB.Recordset
    With rsSO106
         If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open "SELECT ROWID,SO106.* FROM " & GetOwner & "SO106 SO106 WHERE CUSTID= " & gCustId & " AND COMPCODE=" & gCompCode & " ORDER BY ACCEPTTIME DESC", gcnGi, adOpenKeyset, adLockOptimistic
    End With
    With ggrData
         If .Recordset.State = 1 Then
             Set .Recordset = rsSO106
            .Refresh
         End If
    End With
  Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub

Public Sub UserPermissionGo()
  On Error GoTo ChkErr
    Dim blnFlag As Boolean
    If ChkSo1100BMDI Then
        If frmSO1100BMDI.rsSO106.State = adStateOpen Then blnFlag = Not frmSO1100BMDI.rsSO106.EOF
    End If
    ' 根據權限組別, 取相關操作權限
    MenuEnabled GetUserPriv("SO1100G", "(SO1100G1)"), _
                        GetUserPriv("SO1100G", "(SO1100G2)") And blnFlag, _
                        GetUserPriv("SO1100G", "(SO1100G3)") And blnFlag, _
                        blnFlag And EditMode <> giEditModeView, _
                        GetUserPriv("SO1100G", "(SO1100G5)") And blnFlag And False
'    MenuEnabled GetUserPriv("SO1100G", "(SO1100G1)"), _
                        GetUserPriv("SO1100G", "(SO1100G2)") And blnflag, _
                        GetUserPriv("SO1100G", "(SO1100G3)") And blnflag, _
                        GetUserPriv("SO1100G", "(SO1100G4)") And blnflag And EditMode <> giEditModeView, _
                        GetUserPriv("SO1100G", "(SO1100G5)") And blnflag And False
  Exit Sub
ChkErr:
    ErrSub Me.Name, "UserPermissionGo"
End Sub

Private Sub InitializingListOcx()
  On Error GoTo ChkErr
    gilCardName.ListType = OneColumn
    SetLst gilCardName, "CodeNo", "Description", 3, 12, 405, 1620, "CD037"
    SetLst gilBankCode, "CodeNo", "Description", 3, 24, 405, 1850, "CD018", , True
    SetLst gilAcceptName, "EmpNo", "EmpName", 10, 20, , , "CM003"
    SetLst gilMediaCode, "CodeNo", "Description", 3, 12, 405, 1620, "CD009"
    SetLst gilCMCode, "CodeNo", "Description", 3, 12, 405, 1620, "CD031"
    Call SetgiList(gilPTCode, "CodeNo", "Description", "CD032", , , , , 10, 20, "Where Nvl(ServiceType,'" & gServiceType & "' ) ='" & gServiceType & "'", True)
    
'    On Error Resume Next
'    strFilter = gcnGi.Execute("SELECT CITEMSTR FROM SO106 WHERE CUSTID=" & gCustId & " AND STOPFLAG <> 1 AND CITEMSTR IS NOT NULL").GetString(adClipString, , "", ",", "")
'    If Err.Number = 3021 Then
'        Err.Clear
'    Else
'        strFilter = RPxx(strFilter)
'        strFilter = Left(strFilter, Len(strFilter) - 1)
'    End If
    SetMS gimCitem, "代碼", "名稱", "CodeNo", "Description", "CD019", 1, " WHERE PERIODFLAG <> 0 AND ( SERVICETYPE ='" & gServiceType & "' OR SERVICETYPE IS NULL)" '& IIf(strFilter = Empty, "", " AND CODENO NOT IN (" & strFilter & ")")
    gilAcceptName.Filter = " Where CompCode =" & gCompCode & "'"
    gilBankCode.Filter = " Where CompCode =" & gCompCode
    gilMediaCode.Filter = " Where ( ServiceType ='" & gServiceType & "' OR ServiceType IS Null )"
    gilCMCode.Filter = " Where ( ServiceType ='" & gServiceType & "' OR ServiceType IS Null )"
  
'    SetMsQry csmCitem1, "SELECT SEQNO,CITEMCODE,CITEMNAME,PERIOD,AMOUNT,ACCOUNTNO,CMNAME,STARTDATE,STOPDATE FROM " & _
                         GetOwner & "SO003 SO003 WHERE CUSTID = " & _
                         gCustId & " AND COMPCODE=" & gCompCode & _
                         " ORDER BY CLCTDATE,CITEMCODE", _
                         "序號", "代碼", , , True, "收費項目", "期數", "金額", "扣帳帳號", "收費方式", _
                         "起始日", "截止日", , , , 1, 560, 2000, 520, 640, 2000, 900, 900, 900, , , , , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD"
    
'    SetMsQry csmCitem1, "SELECT CITEMCODE,CITEMNAME,PERIOD,AMOUNT,ACCOUNTNO,CMNAME,STARTDATE,STOPDATE,SEQNO FROM " & _
'                         GetOwner & "SO003 SO003 WHERE CUSTID = " & _
'                         gCustId & " AND COMPCODE=" & gCompCode & _
'                         " ORDER BY CLCTDATE,CITEMCODE", _
'                         "代碼", "收費項目", "", , True, "期數", "金額", "扣帳帳號", "收費方式", _
'                         "起始日", "截止日", "序號", , , , 500, 1800, 460, 560, 1800, 850, 850, 850, 770, , , , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD", , , , , 9
    
'    SetMsQry csmCitem2, "SELECT DISTINCT CITEMCODE,CITEMNAME,REALPERIOD,SHOULDAMT,ACCOUNTNO,CMNAME,REALSTARTDATE,REALSTOPDATE,SEQNO,PK,ROWID FROM " & _
'                         "(SELECT CITEMCODE,CITEMNAME,REALPERIOD,SHOULDAMT,ACCOUNTNO,CMNAME,REALSTARTDATE,REALSTOPDATE,SEQNO,BILLNO || ITEM PK,ROWID FROM " & _
'                         GetOwner & "SO033 WHERE CUSTID = " & gCustId & _
'                         " AND COMPCODE=" & gCompCode & _
'                         " AND CANCELFLAG <> 1 AND UCCODE IS NOT NULL AND CHEVEN <> 1 " & _
'                         "ORDER BY REALDATE DESC,BILLNO DESC,SEQNO DESC)", _
'                         "代碼", "收費項目", "", , True, "期數", "金額", "扣帳帳號", "收費方式", _
'                         "起始日", "截止日", "序號", "PK", , , 500, 1800, 460, 560, 1800, 850, 850, 850, 770, 1, 1, , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD", , , , , 10

    '#3454 多串SO003.FaciSNo、SO004.DeclarantName、SO004.FaciName By Kin 2007/12/10
    SetMsQry csmCitem1, "SELECT A.CITEMCODE,A.CITEMNAME,A.PERIOD,A.AMOUNT,A.ACCOUNTNO,A.CMNAME,A.STARTDATE," & _
                        "A.STOPDATE,A.SEQNO,A.FaciSNo,B.DeclarantName,B.FaciName FROM " & _
                         GetOwner & "SO003 A," & GetOwner & "SO004 B  WHERE A.CUSTID = " & _
                         gCustId & " AND A.COMPCODE=" & gCompCode & _
                         " AND A.CustId=B.CustId(+) And A.CompCode=B.CompCode(+)" & _
                         " And A.FaciSeqNo=B.SeqNo(+)" & _
                         " ORDER BY A.CLCTDATE,A.CITEMCODE", _
                         , , "代碼,收費項目,期數,金額,扣帳帳號,收費方式,起始日,截止日,序號,物品序號,申請人姓名,項目名稱", _
                         "500,1800,460,560,1800,850,850,850,770,850,1600,1600", True, , , , , _
                         , , , , , , , , , , , , , , , , , , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD", , , , , 12

    SetMsQry csmCitem2, "SELECT DISTINCT CITEMCODE,CITEMNAME,REALPERIOD,SHOULDAMT,ACCOUNTNO,CMNAME,REALSTARTDATE," & _
                        "REALSTOPDATE,SEQNO,PK,RowID,FaciSNo,DeclarantName,FaciName FROM " & _
                         "(SELECT A.CITEMCODE,A.CITEMNAME,A.REALPERIOD,A.SHOULDAMT,A.ACCOUNTNO,A.CMNAME,A.REALSTARTDATE," & _
                            "A.REALSTOPDATE,A.SEQNO,A.BILLNO || A.ITEM PK,A.ROWID,A.FaciSNo,B.DeclarantName,B.FaciName FROM " & _
                         GetOwner & "SO033 A," & GetOwner & "SO004 B WHERE A.CUSTID = " & gCustId & _
                         " AND A.COMPCODE=" & gCompCode & _
                         " AND A.CANCELFLAG <> 1 AND A.UCCODE IS NOT NULL AND A.CHEVEN <> 1 " & _
                         " AND A.CustId=B.CustId(+) And A.CompCode=B.CompCode(+)" & _
                         " And A.FaciSeqNo=B.SeqNo(+) " & _
                         "ORDER BY A.REALDATE DESC,A.BILLNO DESC,A.SEQNO DESC)", _
                         , , "代碼,收費項目,期數,金額,扣帳帳號,收費方式,起始日,截止日,序號,PK,RowID,物品序號,申請人姓名,項目名稱", _
                         "500,1800,460,560,1800,850,850,850,770,1,1,850,1600,1600", True, , , , , _
                             , , , , , , , , , , , , , , , , , , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD", , , , , 14
    '#3236 ACHTNO代碼沒有唯一,顯示出來都會變成第一個代碼,SO106新增ACHTDESC欄位,SetQueryCode改用ACHDESC欄位 By Kin 2007/05/28
    '#7049 增加BillHeadFmt欄位,可以串到CD068A By Kin 2015/07/08
    SetMsQry csmACH, "SELECT ACHTNO,ACHTDESC,CITEMCODESTR,BillHeadFmt FROM " & GetOwner & "CD068 Where ACHTNO Is Not Null And ACHTDESC is Not NULL And ACHType IN (1,2) ", _
                            "ACH交易代碼", "ACH交易說明", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
                                
    SetMsQry csmTest, "select codeno,description,amount,accidp1 from cd019 where rownum<=10", _
                            "codeno", "description", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
    SetMsQry csmTest, "select codeno,description,amount,accidp1 from cd019 where codeno = 305", _
                            "codeno", "description", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
    csmTest.SetQueryCode "'(代收)CM存入保證金轉出'"
'    SetMsQry csmACH, "SELECT ACHTNO,ACHTDESC,CITEMCODESTR,BillHeadFmt FROM " & GetOwner & "CD068", _
'                                    "ACH交易代碼", "ACH交易說明", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
  Exit Sub
ChkErr:
    ErrSub Me.Name, "InitializingListOcx"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift)
  Exit Sub
ChkErr:
       Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Public Sub AddNewGo()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
    MenuEnabled , , , True, , True
    frmSO1100BMDI.Pic1.Enabled = False
    EditMode = giEditModeInsert
    OpenData
    Me.Show , frmSO1100BMDI
    ChangeMode giEditModeInsert
    lblStatus = ""
    NewRcd
    lblState.Visible = False
    lblStatus.Visible = False
    gdaSendDate.Enabled = True
    gdaAuthStopDate.Enabled = True
    gdaSnactionDate.Enabled = True
    gilBankCode.Enabled = True
    txtAccountNo.Enabled = True
    txtAccOwnerID.Enabled = True
    lblAccountOwner = "帳號所有人"
    txtHide.Visible = False
    txtAccountNo.Enabled = True
    gilCMCode.Enabled = True
    mskCardExpDate.Visible = True
    mskCardExpDate.Enabled = True
'    lblAccountOwner.ForeColor = vbBlack
    gdtAcceptTime.Text = RightNow
    gdaContiDate.Text = RightDate
    gilAcceptName.SetCodeNo garyGi(0)
    gilAcceptName.SetDescription garyGi(1)
    '#3454 申請人改用ComboBox By Kin 2007/12/18
    cboProposer.Text = ""
    Screen.MousePointer = vbDefault
    On Error Resume Next
    'txtProposer.SetFocus
    cboProposer.SetFocus
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "AddNewGo")
End Sub
 
Public Sub EditGo()
  On Error GoTo ChkErr
    '#4094 增加參數判斷ACH銀行時是否能使用停用 By Kin 2009/09/19
    blnEnableStop = Val(GetRsValue("Select Nvl(ACHCancelStop,0) From " & GetOwner & _
                              "SO041 Where SysID=" & gCompCode, gcnGi))
    '**************************************************************************
    '#3728 當停用時,銀行別是ACH時不允許更新資料 By Kin 2008/03/18
    If Screen.ActiveForm.Name <> "frmSO1100G" Then
        If Not blnEnableStop Then
                If Val(frmSO1100BMDI.rsSO106("StopFlag")) = 1 Then
                    If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018" & _
                                            " WHERE CODENO=" & frmSO1100BMDI.rsSO106("BankCode") & _
                                            " AND ( PRGNAME LIKE 'ACH%' Or " & posAchContition & " ) " & _
                                            " AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
                                            
                        MsgBox "該筆資料已停用授權，不允許修改！", vbInformation, "訊息"
                        Exit Sub
                    End If
                End If
        End If
    Else
        If chkStop.Value Then
            '#4094 如果ACHCancelStop=1 要讓User可以進入修改 By Kin 2008/09/19
            If Not blnEnableStop Then
                If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018" & _
                                        " WHERE CODENO=" & gilBankCode.GetCodeNo & _
                                        " AND ( PRGNAME LIKE 'ACH%' Or " & posAchContition & " ) " & _
                                        " AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
                                        
                    MsgBox "該筆資料已停用授權，不允許修改！", vbInformation, "訊息"
                    Exit Sub
                End If
            End If
        End If
    End If
    '**************************************************************************
    Screen.MousePointer = vbHourglass
    MenuEnabled , , , True, , True
    frmSO1100BMDI.Pic1.Enabled = False
    Me.Show , frmSO1100BMDI
    If Screen.ActiveForm.Name <> "frmSO1100G" Then OpenData
    ' 子帳戶不能做異動，只有父帳戶才能做異動，而子帳戶只做連動處理。
    '（即InheritFlag=0的資料方可作異動，InheritFlag=1的資料將被ReadOnly）
    If Val(rsSO106("InheritFlag") & "") = 1 Then
        ViewGo
        MsgBox "子帳戶資料不能單獨做異動 !", vbInformation, "訊息"
        Exit Sub
    End If
    AbsPos
    ChangeMode giEditModeEdit
    PrcACH
    Screen.MousePointer = vbDefault
    On Error Resume Next
    'txtProposer.SetFocus
    cboProposer.SetFocus
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "EditGo")
End Sub

Private Sub AbsPos()
  On Error GoTo ChkErr
    With frmSO1100BMDI.rsSO106
         If .RecordCount > 0 Then
            If Not .EOF Or Not .BOF Then
                If rsSO106.AbsolutePosition <> .AbsolutePosition Then rsSO106.AbsolutePosition = .AbsolutePosition
                RcdToScr
            End If
         End If
    End With
    
  Exit Sub
ChkErr:
    ErrSub Me.Name, "AbsPos"
End Sub
 
Public Sub ViewGo()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
    MenuEnabled , , , True, , True
    frmSO1100BMDI.Pic1.Enabled = False
    Me.Show , frmSO1100BMDI
    If Not Me.Visible Then OpenData
    ChangeMode giEditModeView
    AbsPos
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "ViewGo")
End Sub

Public Sub CancelGo()
   On Error GoTo ChkErr
    gilBankCode_Change
    blnCallSO106A = False
    If EditMode = giEditModeView Then
        Unload Me
        Exit Sub
   Else
      If EditMode = giEditModeEdit Then
         Call ChangeMode(giEditModeView)
         RevertRcd
      Else
         If Not rsSO106.EOF Then
            rsSO106.MoveFirst
            RcdToScr
         Else
            NewRcd
         End If
         Call ChangeMode(giEditModeView)
      End If
'      OpenData
   End If
 Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "CancelGo")
End Sub

Private Sub RevertRcd() '還原資料內容
 On Error GoTo ChkErr
   RcdToScr
 Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "RevertRcd")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ChkErr
        If EditMode <> giEditModeView Then
            If giMsgNotSave = vbNo Then
                Cancel = True: Exit Sub
            Else
                CancelGo
            End If
        End If
        On Error Resume Next
        rsSO106.Close
        Set rsSO106 = Nothing
        DoEvents
        FormQueryUnload
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_QueryUnload"
End Sub

Private Sub gdaSendDate_Change()
  On Error GoTo ChkErr
     If InStr(gdaSendDate.GetOriginalValue, "_") = 0 Then
        If gilBankCode.GetCodeNo <> Empty Then
            If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018  " & _
                            " WHERE CODENO=" & gilBankCode.GetCodeNo & _
                            " AND ( PRGNAME LIKE 'ACH%' Or  " & posAchContition & " )  AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
                If rsSO106.RecordCount > 0 Then
                    If rsSO106("AuthorizeStatus") & "" = 1 Or rsSO106("AuthorizeStatus") & "" = Empty Then
                        gilBankCode.Enabled = False
                        txtAccountNo.Enabled = False
                        blnACHAccountNoCanEdit = False
                        txtAccOwnerID.Enabled = False
                    Else
                        CtlEnable
                    End If
                Else
                    CtlEnable
                End If
            Else
                CtlEnable
            End If
        Else
            CtlEnable
        End If
     Else
        CtlEnable
     End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "gdaSendDate_Change"
End Sub

Private Sub CtlEnable()
  On Error Resume Next
    gilBankCode.Enabled = True
    txtAccountNo.Enabled = True
    blnACHAccountNoCanEdit = True
    txtAccOwnerID.Enabled = True
End Sub

Private Sub gdaSnactionDate_GotFocus()
  On Error GoTo ChkErr
    '#5045 需自動帶入系統日期 By Kin 2009/05/05
    If gdaSnactionDate.GetValue & "" = "" Then
        gdaSnactionDate.SetValue RightNow
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaSnactionDate_GotFocus")
End Sub

Private Sub gdtAcceptTime_Validate(Cancel As Boolean)
  On Error Resume Next
    If Not IsDate(gdtAcceptTime.Text) And gdtAcceptTime.GetValue <> "" Then Cancel = True: Exit Sub
End Sub

Private Sub ggrData_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If UCase(Fld.Name) = "STOPFLAG" Then
        If Value = 1 Then Value = vbRed
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ggrData_ColorCellData"
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If UCase(Fld.Name) = "STOPFLAG" Then Value = IIf(Value = 0, "", "是")
    If UCase(Fld.Name) = "STOPYM" Then
        If Not GetUserPriv("SO1100G", "(SO1100G7)") Then
            Value = "XX/XXXX"
        End If
    End If
    If UCase(Fld.Name) = "ACCOUNTID" Then
        If Not GetUserPriv("SO1100G", "(SO1100G9)") Then
            If Value = Empty Then
                Value = ""
            Else
                Value = Left(Value, 5) & "XXX" & Mid(Value, 9, 2) & "XX" & Right(Value, 4)
            End If
        End If
    End If
    If UCase(Fld.Name) = "ID" Or UCase(Fld.Name) = "ACCOUNTNAMEID" Then
        If Not GetUserPriv("SO1100G", "(SO1100G0)") Then
            If Value = Empty Then
                Value = ""
            Else
                Value = Left(Value, 3) & "XXXX" & Right(Value, 3)
            End If
        End If
    End If
    If Len(strAccountName) > 0 Then
        If strAccountName <> ggrData.Recordset.Fields("Proposer") Then
            Value = "XXXXX"
        End If
    End If
    
'    If Not GetUserPriv("SO1100G", "(SO1100G9)") Then
'        If UCase(Fld.Name) = "ACCOUNTID" Then Value = Left(Value, 5) & "XXX" & Mid(Value, 9, 2) & "XX" & Right(Value, 4)
'    End If
'    If Not GetUserPriv("SO1100G", "(SO1100G0)") Then
'        If UCase(Fld.Name) = "ID" Or UCase(Fld.Name) = "ACCOUNTNAMEID" Then Value = Left(Value, 2) & "XXXX" & Right(Value, 2)
'    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ggrData_ShowCellData"
End Sub

Private Sub gilBankCode_Change()
  On Error Resume Next
  
    Dim intActLen As Integer
    Dim intVNO As Integer
    If EditMode <> giEditModeView Then
        If Not IsNumeric(gilBankCode.GetCodeNo) Then Exit Sub
        '#4148 多加PrcACH Sub,因為如果先選擇銀行,若選到ACH銀行的話,ACH Enabled會失效 By Kin 2008/10/16
        If gilCMCode.GetCodeNo = Empty Then PrcACH: Exit Sub
        '#8724
        If intCD031RefNo = 2 Or intCD031RefNo = 4 Or intCD031RefNo = 5 Then
            blnVAcc = False
            If gilCardName.GetCodeNo <> Empty Then
                On Error Resume Next
                '*********************************************************************************
                '#3395 如果是使用信用卡，使用信用卡的檢核點 By Kin 2007/08/20
                '#8724
                If intCD031RefNo = 4 Or intCD031RefNo = 5 Then
                    If Not ChkCreditCard(gilCardName.GetCodeNo, txtAccountNo, intActLen) Then
                        txtAccountNo.SetFocus
                        txtAccountNo.MaxLength = intActLen
                    End If
                Else
                    intActLen = RPxx(gcnGi.Execute("SELECT CARDNOLEN FROM " & GetOwner & "CD037 WHERE CODENO=" & gilCardName.GetCodeNo).GetString(adClipString, 1, "", "", ""))
                    If intActLen <> Empty Then txtAccountNo.MaxLength = intActLen
                End If
                '*********************************************************************************
            Else
                If Len(gilBankCode.GetCodeNo) > 0 Then
                    txtAccountNo.MaxLength = Val(gcnGi.Execute("SELECT ActLength FROM " & GetOwner & "CD018 WHERE CodeNo = " & gilBankCode.GetCodeNo).GetString)
                Else
                    txtAccountNo.MaxLength = 24
                End If
            End If
        Else
            '#3541改變虛擬帳號規則 By Kin 2007/11/30
            Dim iCustCount As Long
            Dim strVNO As String
            If gilBankCode.GetCodeNo <> Empty Then
                intActLen = GetRsValue("SELECT ACTLENGTH FROM " & GetOwner & "CD018 WHERE CODENO=" & gilBankCode.GetCodeNo)
                txtAccountNo.MaxLength = intActLen
                
                iCustCount = gcnGi.Execute("SELECT COUNT(*) FROM " & GetOwner & _
                                          "SO002A WHERE CUSTID=" & gCustId & _
                                          " AND ID=2")(0)
                
                If iCustCount > 0 Then
                    strVNO = GetRsValue("Select Max(SubStr(AccountNO,1,8)) FROM " & GetOwner & _
                                        "SO002A WHERE CUSTID=" & gCustId & " AND ID=2")
                    strVNO = Format(CLng(strVNO) + 1, "00000000")
                Else
                    strVNO = Format(iCustCount + 1, "00000000")
                End If
                                          
                txtAccountNo = strVNO & String(intActLen - Len(CStr(strVNO)) - Len(CStr(gCustId)), "0") & gCustId
                blnVAcc = True
            End If
        End If
    End If
    '個資法
    If Len(strAccountName) > 0 Then
        If strAccountName <> rsSO106.Fields("Proposer") Then
            Exit Sub
        End If
    End If
    PrcACH
End Sub

Private Sub PrcACH()
  On Error GoTo ChkErr
    '#3728 如果銀行別為ACH則ACH授權交易別為可選,反之為不可選 By Kin 2008/03/07
    '#7537 Add POST
    If gilBankCode.GetCodeNo <> Empty Then
        If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018 " & _
                " WHERE CODENO=" & gilBankCode.GetCodeNo & " AND ( PRGNAME LIKE 'ACH%'  Or " & posAchContition & " )  " & _
                " AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
            If blnStartPost Then
                If GetRsValue("Select count(*) from cd018 Where CD018.PRGNAME like '%POST%' and CodeNo = " & gilBankCode.GetCodeNo, gcnGi) = 0 Then
                    AchTypeCondition = " In ( 1 )"
                Else
                    AchTypeCondition = " In ( 2 )"
                End If
            
                Dim qrySelected As String
                qrySelected = Empty
                If Len(csmACH.GetQueryCode & "") > 0 Then
                    qrySelected = csmACH.GetQueryCode
                End If
                    
                SetMsQry csmACH, "SELECT ACHTNO,ACHTDESC,CITEMCODESTR,BillHeadFmt FROM " & GetOwner & "CD068  " & _
                                            " Where  ACHTNO is Not Null And ACHTDESC is Not NULL And ACHType " & AchTypeCondition, _
                                        "ACH交易代碼", "ACH交易說明", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
                If Len(qrySelected & "") > 0 Then
                    csmACH.SetQueryCode qrySelected
                End If
            End If
            
            lblState.Visible = True
            lblStatus.Visible = True
            FraACH.Enabled = True
            'cmdACHDetail.Enabled = True
            '#3925 SO041.ACHCUSTID=0時 ACH歷程明細按鈕要Disable By Kin 2008/05/08
            cmdACHDetail.Enabled = blnACHCustID
            If EditMode <> giEditModeView Then
                gdaSendDate.Enabled = False
                gdaSnactionDate.Enabled = False
            End If
            '#4094 多增加參數控制是否能使用停用 By Kin 2008/09/19
'            chkStop.Enabled = blnEnableStop
'            gdaStopDate.Enabled = blnEnableStop
        Else
            csmACH.Clear
            FraACH.Enabled = False
            cmdACHDetail.Enabled = False
            lblState.Visible = False
            lblStatus.Visible = False
            '#4094 如果不是ACH的銀行還是要能使用停用 By Kin 2008/09/19
'            chkStop.Enabled = True
'            gdaStopDate.Enabled = True
            If EditMode <> giEditModeView Then
                gdaSendDate.Enabled = True
                gdaSnactionDate.Enabled = True
                chkToACH.Value = 0
                csmACH.Clear
                'txtAchSN = ""   '#5045 任何模式下都要能編輯 By Kin 2009/05/05
            End If
        End If
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "PrcACH"
End Sub

Private Sub gilCardName_Change()
  On Error GoTo ChkErr
    If EditMode <> giEditModeView Then
'        If ServiceCommon.AccountNo_Validate(Me) Then txtAccountNo.SetFocus
        gilBankCode_Change
        If intCD031RefNo = 2 Then
            ServiceCommon.AccountNo_Validate Me
        End If
'        If intCD031RefNo = 4 Then
'            ChkCreditCard gilCardName.GetCodeNo, txtAccountNo
'        End If
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "gilCardName_Change"
End Sub

Private Sub gilCMCode_Change()
  On Error Resume Next
    Dim strCD As Integer
    Dim rsCD032 As New ADODB.Recordset
    strCD = gilCMCode.GetCodeNo
    If strCD <> Empty Then
        intCD031RefNo = GetRsValue("SELECT REFNO  FROM " & GetOwner & "CD031 WHERE CODENO=" & strCD)
        If Err.Number <> 0 Then
            Err.Clear
            intCD031RefNo = 0
        End If
'        Debug.Print intCD; intCD031RefNo
        If intCD031RefNo = 2 Or intCD031RefNo = 4 Or intCD031RefNo = 5 Then
'            lblAccountOwner.ForeColor = vbRed
            lblAccountOwner = "帳號所有人"
'            If txtAccountOwner = Empty Then txtAccountOwner = frmSO1100BMDI.txtCustName
'            lblID.ForeColor = vbRed
            lblAccNameID.ForeColor = vbRed
            If EditMode <> giEditModeView Then
                If Len(txtAccountNo) = 16 Then
'                    If Not InStr(1, txtAccountNo, gCustId, vbTextCompare) Then
'                        txtAccountNo = ""
'                    End If
                Else
                    If Not InStr(1, txtAccountNo, gCustId, vbTextCompare) Then
                        txtAccountNo = ""
                    End If
                End If
            End If
            If intCD031RefNo = 4 Or intCD031RefNo = 5 Then
                '#5045 如果收費方式參考號=4,付款種類也要自動帶出參考號為4的資料 By Kin 2009/05/05
                If Not GetRS(rsCD032, "Select * From " & GetOwner & "CD032 Where RefNO in (4,5) And " & _
                            " StopFlag<>1", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                If Not rsCD032.EOF Then
                    gilPTCode.SetCodeNo rsCD032("CodeNo") & ""
                    gilPTCode.Query_Description
                End If
                lblCardName.ForeColor = vbRed
                lblCardExpDate.ForeColor = vbRed
            Else
                lblCardName.ForeColor = vbBlack
                lblCardExpDate.ForeColor = vbBlack
            End If
        ElseIf intCD031RefNo = 0 Then
            If EditMode <> giEditModeView Then
'                lblID.ForeColor = vbBlack
                lblAccNameID.ForeColor = vbBlack
                lblAccountOwner = "聯絡人"
'                lblAccountOwner.ForeColor = vbBlack
                If Len(txtAccountNo) = 16 Then
                    If Not InStr(1, txtAccountNo, gCustId, vbTextCompare) Then
                        txtAccountNo = ""
                    End If
                End If
                lblCardName.ForeColor = vbBlack
                lblCardExpDate.ForeColor = vbBlack
            End If
        Else
'            lblID.ForeColor = vbBlack
            lblAccNameID.ForeColor = vbBlack
            lblAccountOwner = "聯絡人"
'            lblAccountOwner.ForeColor = vbBlack
            lblAccNameID.ForeColor = vbBlack
            lblCardName.ForeColor = vbBlack
            lblCardExpDate.ForeColor = vbBlack
            If gilBankCode.GetCodeNo <> Empty Then gilBankCode_Change
        End If
        '#5045 直接帶申請人姓名 By Kin 2009/05/05
        '#5121 2009.05.26 by Corey 原意#5045 應該是指 如果帳號所有人是空白的話才需要帶入預設值，如果本身已經有值的話就以本身的值為準
        'If cboProposer.Text <> "" Then
        If cboProposer.Text <> "" And txtAccountOwner.Text = "" Then
            txtAccountOwner.Text = cboProposer.Text
        End If
    End If
    '#8724
    If intCD031RefNo <> 2 And intCD031RefNo <> 4 And intCD031RefNo <> 5 And txtAccountNo = Empty Then
        If gilBankCode.GetCodeNo <> Empty Then gilBankCode_Change
    End If
    '#8724
    If intCD031RefNo <> 4 And intCD031RefNo <> 5 Then
        If gilPTCode.GetCodeNo = Empty Then
            gilPTCode.SetCodeNo 1
            gilPTCode.Query_Description
        End If
    End If
    Call Close3Recordset(rsCD032)
End Sub

Private Sub gilMediaCode_Change()
  On Error GoTo ChkErr
    ServiceCommon.CommonMediaCode_Change Me
  Exit Sub
ChkErr:
    ErrSub Me.Name, "gilMediaCode_Change"
End Sub

Private Sub gilMediaCode_LostFocus()
  On Error GoTo ChkErr
    ServiceCommon.MediaCode_LostFocus Me
  Exit Sub
ChkErr:
    ErrSub Me.Name, "gilMediaCode_LostFocus"
End Sub

Private Sub mskCardExpDate_GotFocus()
  On Error Resume Next
    ObjGotFocus ImeClose
End Sub

Private Sub mskCardExpDate_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    Cancel = ServiceCommon.CardExpireDate(Me)
  Exit Sub
ChkErr:
    ErrSub Me.Name, "mskCardExpDate_Validate"
End Sub

Private Sub rsSO106_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  On Error GoTo ChkErr
    If adReason = 10 Then
        If Not rsSO106.EOF And Not rsSO106.BOF And lngEditMode = giEditModeView Then
            If Not ggrData.Enabled Then ggrData.Enabled = True
            RcdToScr
            On Error Resume Next
            frmSO1100BMDI.rsSO106.AbsolutePosition = rsSO106.AbsolutePosition
        Else
            ggrData.Enabled = False
        End If
        txtAccountNo.PasswordChar = Empty
        txtAccountOwner.PasswordChar = Empty
        txtHide.PasswordChar = Empty
        '個資法
'        If Len(strAccountName) > 0 Then
'            If strAccountName <> rsSO106.Fields("cboProposer") Then
'                cboProposer.Text = "XXXXX"
'                gilCMCode.SetCodeNo "X"
'                gilCMCode.SetDescription "XXX"
'                gilPTCode.SetCodeNo "X"
'                gilPTCode.SetDescription "XXX"
'                txtAccountNo.PasswordChar = "X"
'                txtAccountOwner.PasswordChar = "X"
'                txtHide.PasswordChar = "X"
'            End If
'        End If
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "rsSO106_MoveComplete")
End Sub

Private Sub ReFilter()
  On Error Resume Next
    
    Dim strFilter As String
    strFilter = gcnGi.Execute("SELECT CITEMSTR FROM " & GetOwner & "SO106 WHERE CUSTID=" & gCustId & " AND STOPFLAG <> 1 AND CITEMSTR IS NOT NULL AND ROWID <> '" & rsSO106!RowId & "'").GetString(adClipString, , "", ",", "")
    If Err.Number = 3021 Then
        Err.Clear
    Else
        strFilter = RPxx(strFilter)
        strFilter = Left(strFilter, Len(strFilter) - 1)
    End If
    SetMS gimCitem, "代碼", "名稱", "CodeNo", "Description", "CD019", 1, " WHERE PERIODFLAG <> 0 AND ( SERVICETYPE ='" & gServiceType & "' OR SERVICETYPE IS NULL)" & IIf(strFilter = Empty, "", " AND CODENO NOT IN (" & strFilter & ")")
End Sub
Public Property Let uAccountName(ByVal vData As String)
  On Error Resume Next
    strAccountName = vData
End Property
Public Property Let uInvSeqNo(ByVal vData As String)
  On Error GoTo ChkErr
    strInvSeqNo = vData
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Let uInvSeqNo")
End Property
Public Property Get EditMode() As giEditModeEnu
  On Error GoTo ChkErr
    EditMode = lngEditMode    ' 取目前編輯模式
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Get EditMode")
End Property

Public Property Let EditMode(ByVal vNewValue As giEditModeEnu)
  On Error GoTo ChkErr
    lngEditMode = vNewValue '設定編輯模式
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Let EditMode")
End Property

'MID,PROMPT,LINK
'SO1100G , 信用卡轉帳, SO1100
'SO1100G0 , 是否完全顯示身份証號, SO1100G
'SO1100G1 , 資料新增, SO1100G
'SO1100G2 , 資料修改, SO1100G
'SO1100G3 , 資料刪除, SO1100G
'SO1100G4 , 資料查看, SO1100G
'SO1100G5 , 資料列印, SO1100G
'SO1100G6 , 可以修改信用卡有效期, SO1100G
'SO1100G7 , 可以讀取信用卡有效期, SO1100G
'SO1100G8 , 不可讀取信用卡有效期, SO1100G
'SO1100G9 , 是否完全顯示帳號, SO1100G
'SO1100GA , 可修改轉帳帳號, SO1100G
'SO1100GB , 可修改收費方式, SO1100G

'根據傳來之編輯模式, 設定各物件屬性
Public Sub ChangeMode(ByVal lngMode As giEditModeEnu)
  On Error GoTo ChkErr
    Dim blnFlag As Boolean  ' 記錄是否在資料瀏覽模式，
    lngEditMode = lngMode
    Select Case lngMode
           Case giEditModeInsert
                blnFlag = False
                MenuEnabled , , False, True, False, True
           Case giEditModeEdit
                blnFlag = False
                MenuEnabled , , False, True, False, True
           Case giEditModeView
                blnFlag = True
                MenuEnabled True, rsSO106.RecordCount > 0, rsSO106.RecordCount > 0, , False, True
    End Select
    
    Select Case True
           Case (GetUserPriv("SO1100G", "(SO1100G6)"))
                 txtHide.Visible = False
                 mskCardExpDate.Visible = True
                 mskCardExpDate.Enabled = True
           Case (GetUserPriv("SO1100G", "(SO1100G7)"))
                 txtHide.Visible = False
                 mskCardExpDate.Visible = True
                 mskCardExpDate.Enabled = False
           Case (GetUserPriv("SO1100G", "(SO1100G8)"))
                 txtHide.Visible = True
                 mskCardExpDate.Visible = False
                 mskCardExpDate.Enabled = False
    End Select
    
    If lngMode <> giEditModeView Then
        gdaSendDate_Change
        txtAccountNo.Enabled = (GetUserPriv("SO1100G", "(SO1100GA)")) And blnACHAccountNoCanEdit
        gilCMCode.Enabled = (GetUserPriv("SO1100G", "(SO1100GB)"))
        ReFilter
    End If
    
    MenuEnabled GetUserPriv("SO1100G", "(SO1100G1)") And blnFlag, _
                GetUserPriv("SO1100G", "(SO1100G2)") And blnFlag, _
                GetUserPriv("SO1100G", "(SO1100G3)") And blnFlag, _
                Not blnFlag And EditMode <> giEditModeView, _
                GetUserPriv("SO1100G", "(SO1100G5)") And blnFlag And False, True
    '***************************************************************************
    '#3929 如果核淮日期沒值,則要能重新選擇發票流水號 By Kin 2008/06/09
    If (lngEditMode = giEditModeInsert) Or (lngEditMode = giEditModeEdit) Then
        If Not ChkInvData Then
            cmdInvSeqNo.Enabled = True
            cmdCreateINV.Enabled = True
        Else
            cmdInvSeqNo.Enabled = False
            cmdCreateINV.Enabled = False
            '**************************************************************
            '#3929 要多再判斷一次,因為資料有可能多筆 By Kin 2008/06/09
            If lngEditMode = giEditModeEdit Then
                If IsNull(rsSO106("SnactionDate")) Then
                    cmdInvSeqNo.Enabled = True
                    cmdCreateINV.Enabled = True
                End If
            End If
            '****************************************************************
        End If
    Else
        cmdInvSeqNo.Enabled = False
        cmdCreateINV.Enabled = False
    End If
    '***************************************************************************
    SetStatusBar lngEditMode
    fraData.Enabled = Not blnFlag
    ggrData.Enabled = blnFlag
    '#5556 在瀏灠模式要能捲動捲軸 By Kin 2010/03/01
    txtNote.Locked = blnFlag
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ChangeMode")
End Sub

Private Sub rsSO106_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  
  On Error Resume Next
    blnAchChgClick = False
End Sub

Private Sub txtAccountNo_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    Dim lngAccMax As Integer
    If gilCMCode.GetCodeNo <> Empty Then
        On Error Resume Next
        intCD031RefNo = GetRsValue("SELECT REFNO FROM " & GetOwner & "CD031 WHERE CODENO=" & gilCMCode.GetCodeNo)
        If Err.Number <> 0 Then
            Err.Clear
            intCD031RefNo = 0
        End If
    End If
    On Error GoTo ChkErr
    If EditMode <> giEditModeView Then
        '*************************************************************************************************************************
        '#3395 收費方式為信用卡時(CD031.RefNo=4)，使用信用卡檢核方式，收費方式為銀行時規則不變(CD031.RefNo=2) By Kin 2007/08/20
        If intCD031RefNo = 2 Then  'Or intCD031RefNo = 4 Then
            Cancel = ServiceCommon.AccountNo_Validate(Me)
        End If
        '#8724
        If intCD031RefNo = 4 Or intCD031RefNo = 5 Then
            Cancel = Not ChkCreditCard(gilCardName.GetCodeNo, txtAccountNo.Text, lngAccMax)
        End If
        '**************************************************************************************************************************
    End If
'    If EditMode = giEditModeInsert Then
'        If Not ChkInvData Then
'            cmdInvSeqNo.Enabled = True
'
'        Else
'
'        End If
'    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "txtAccountNo_Validate"
End Sub

Private Sub txtAccountOwner_Change()
  On Error Resume Next
    CML
End Sub

Private Sub txtAccountOwner_KeyPress(keyAscii As Integer)
  On Error Resume Next
    ML keyAscii
End Sub

Private Sub txtAccOwnerID_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If chkFore2.Value Then Exit Sub
    Select Case Len(txtAccOwnerID)
           Case 0
                Exit Sub
           Case 8
                If GetUserPriv("SO1100G", "(SO1100G0)") Then
                    InvNoIsOk txtAccOwnerID.Text, True
                Else
                    InvNoIsOk txtAccOwnerID.Tag, True
                End If
           Case 10
                If GetUserPriv("SO1100G", "(SO1100G0)") Then
                    IDIsOk txtAccOwnerID.Text, True
                Else
                    IDIsOk txtAccOwnerID.Tag, True
                End If
           Case Else
'                MsgBox "本欄位長度應該為 8 或 10 個字元", vbInformation, "訊息"
                MsgBox "請確認本欄位長度 ! " & vbCrLf & _
                        "八碼為 -- 公司之統一編號" & vbCrLf & _
                        "十碼為 -- 個人之身份證字號", vbInformation, "訊息"
                Cancel = True
    End Select
  Exit Sub
ChkErr:
    ErrSub Me.Name, "txtAccOwnerID_Validate"
End Sub

Private Sub txtId_GotFocus()
  On Error Resume Next
    ObjGotFocus
End Sub

Private Sub txtID_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
'    Cancel = Not IDIsOk(txtID, True)
    If chkFore.Value Then Exit Sub
    Select Case Len(txtID)
           Case 0
                Exit Sub
           Case 8
                If GetUserPriv("SO1100G", "(SO1100G0)") Then
                    InvNoIsOk txtID.Text, True
                Else
                    InvNoIsOk txtID.Tag, True
                End If
           Case 10
                If GetUserPriv("SO1100G", "(SO1100G0)") Then
                    IDIsOk txtID.Text, True
                Else
                    IDIsOk txtID.Tag, True
                End If
           Case Else
'                MsgBox "本欄位長度應該為 8 或 10 個字元", vbInformation, "訊息"
                MsgBox "請確認本欄位長度 ! " & vbCrLf & _
                        "八碼為 -- 公司之統一編號" & vbCrLf & _
                        "十碼為 -- 個人之身份證字號", vbInformation, "訊息"
                Cancel = True
    End Select
  Exit Sub
ChkErr:
    ErrSub Me.Name, "txtId_Validate"
End Sub

Private Sub txtIntroId_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = vbKeyF1 Then cmdFind.Value = True
  Exit Sub
ChkErr:
    ErrSub Me.Name, "txtIntroId_KeyDown"
End Sub

Private Sub txtIntroId_LostFocus()
  On Error Resume Next
    If Len(txtIntroId.Text) = 0 Then lblIntroName.Caption = ""
End Sub

Private Sub txtIntroID_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    Cancel = ServiceCommon.IntroId_Validate(Me)
  Exit Sub
ChkErr:
    ErrSub Me.Name, "txtIntroId_Validate"
End Sub

Private Sub txtNote_Change()
  On Error Resume Next
    CML
End Sub

Private Sub txtNote_KeyPress(keyAscii As Integer)
  On Error Resume Next
    ML keyAscii
End Sub

Private Sub txtProposer_Change()
  On Error Resume Next
    CML
End Sub

Private Sub txtProposer_KeyPress(keyAscii As Integer)
  On Error Resume Next
    ML keyAscii
End Sub

'#3454 申請人改用ComboBox,將SO001.CustName與SO004.Declarantname填入到ComboBox By Kin 2007/12/18
Private Sub ShowProposer()
  On Error GoTo ChkErr
    Dim rsProposer As New ADODB.Recordset
    Dim strQry As String
    strQry = "Select A.CustName From " & GetOwner & "SO001 A " & _
             " Where A.CustId=" & gCustId & _
             " And A.CompCode=" & gCompCode & _
             " Union All " & _
             "Select B.DeclarantName CustName From " & GetOwner & "SO004 B " & _
             " Where B.CustId=" & gCustId & _
             " And B.CompCode=" & gCompCode & _
             " And (B.PRDate Is Null OR B.InstDate > B.PRDate)"
    strQry = "Select Distinct(A.CustName) From (" & strQry & ") A"
    If Not GetRS(rsProposer, strQry, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Sub
    rsProposer.MoveFirst
    cboProposer.Clear
    Do While Not rsProposer.EOF
        If rsProposer("CustName") & "" <> "" Then
            cboProposer.AddItem rsProposer("CustName") & ""
        End If
        rsProposer.MoveNext
    Loop
    CloseRecordset rsProposer
    Exit Sub
ChkErr:
    ErrSub Me.Name, "ShowProposer"
End Sub


'#3728 需求
'將SO106的CitemStr字串拆解開來,並取得那些收費項目是屬於某一個ACHTNO,並得知應該是異動、新增授權、取消授權 By Kin 2008/03/04
'程式有點小複雜,大致上是使用串列結構方式,時間複雜度N平方
'回傳一個X*4的陣列
'(0,0) ACH代碼
'(0,1) ACH名稱
'(0,2) 異動別 0:新增授權 1:取消授權 2:異動收費項目
'(0,3) 收費項目代碼
'(0,4) 收費項目名稱
Private Sub GetCitemCode(ByVal strACHTNO As String, _
                        ByVal strACHTDESC As String, _
                        ByVal strCitemCode As String, _
                        ByVal strCitemName As String, _
                        ByVal blnAdd As Boolean, _
                        ByRef aryCitem() As String)
  On Error GoTo ChkErr
    '0:代表要新增一筆授權,1:代表要新增一筆取消授權 2:代表要Update原本的資料
    Dim aryACHTNO() As String
    Dim aryACHTDESC() As String
    Dim aryCitemCode() As String
    Dim aryCitemName() As String
    Dim rs As New ADODB.Recordset
    Dim i As Long, j As Long
    Dim lngACHTNO As Long
    Dim lngUbound As Long
    Dim strTmpACH As String
    '新增資料
    If blnAdd Then
        If strACHTNO = Empty Then ReDim aryCitem(0, 4): GoTo Fin    '如果沒有選擇ACH,不進行比較
        aryACHTDESC = Split(Replace(strACHTDESC, "'", ""), ",")
        lngUbound = UBound(aryACHTDESC)
        ReDim aryCitem(lngUbound, 4)
        ReDim aryACHTNO(UBound(aryACHTDESC), 1)
        For i = LBound(aryACHTDESC) To UBound(aryACHTDESC)
            aryACHTNO(i, 0) = Split(Replace(strACHTNO, "'", ""), ",")(i)
            aryACHTNO(i, 1) = aryACHTDESC(i)
            aryCitem(i, 0) = aryACHTNO(i, 0)
            aryCitem(i, 1) = aryACHTNO(i, 1)
            aryCitem(i, 2) = "0"    '0代表新增
        Next i
        aryCitemCode = Split(Replace(strCitemCode, "'", ""), ",")
        aryCitemName = Split(Replace(strCitemName, "'", ""), ",")
        strTmpACH = strACHTNO
    '編輯資料
    Else
        With rsSO106
            Dim aryOldACHTDESC() As String  'SO106.CitemStr
            Dim aryOldACHTNO() As String
            Dim strQry1 As String
            Dim strQry2 As String
            Dim blnDiff As Boolean
            Dim lngAry As Long
            Dim strOldCitemCode As String
            Dim strOldCitemName As String
            Dim blnUpdCitem As Boolean
            Dim aryUpdCitemCode() As String
            Dim aryUpdCitemName() As String
            Dim blnAddNew As Boolean
            '**************************************************************************
            '編輯模式如果原SO106沒有ACH而畫面也沒有ACH 則不進行比較
            '#3728 測試不OK,.Fields("CitemStr")要改成"ACHTNO" By Kin 2008/04/09
            If strACHTNO = Empty And IsNull(.Fields("ACHTNO").OriginalValue) Then
                ReDim aryCitem(0, 4)
                GoTo Fin
            End If
            '*************************************************************************
            aryACHTDESC = Split(Replace(strACHTDESC, "'", ""), ",") '畫面上所選取的ACH授權交易別
            aryOldACHTDESC = Split(Replace(.Fields("ACHTDESC").OriginalValue & "", "'", ""), ",")
            '#4179 多加一個字串符別,不然會出現Null的錯誤訊息 By Kin 2008/10/29
            strTmpACH = .Fields("ACHTNO").OriginalValue & ""
            strQry1 = "Select CitemCode From " & GetOwner & "SO003 Where SeqNo in(" & _
                        Replace(rsSO106("CitemStr").OriginalValue & "", "'", "") & ")" & _
                        " And Custid=" & gCustId & " Order by ClctDate,CitemCode"
            strQry2 = "Select CitemName From " & GetOwner & "SO003 Where SeqNo in(" & _
                        Replace(rsSO106("CitemStr").OriginalValue & "", "'", "") & ")" & _
                        " And Custid=" & gCustId & " Order by ClctDate,CitemCode"
            'ACH畫面所選取與SO106相同表示只有收費項目有異動
            If UBound(aryOldACHTDESC) = UBound(aryACHTDESC) Then
                lngUbound = UBound(aryACHTDESC)
                ReDim aryCitem(lngUbound, 4)
                ReDim aryACHTNO(UBound(aryACHTDESC), 1)
                aryCitemCode = Split(Replace(strCitemCode, "'", ""), ",")
                aryCitemName = Split(Replace(strCitemName, "'", ""), ",")
                For i = LBound(aryACHTDESC) To UBound(aryACHTDESC)
                    aryACHTNO(i, 0) = Split(Replace(strACHTNO, "'", ""), ",")(i)
                    aryACHTNO(i, 1) = aryACHTDESC(i)
                    aryCitem(i, 0) = aryACHTNO(i, 0)
                    aryCitem(i, 1) = aryACHTNO(i, 1)
                    aryCitem(i, 2) = "2"            '2代表是編輯
                Next i
                strTmpACH = strACHTNO
            Else
                blnUpdCitem = True  '代表ACH有異動,收費項目可能也有異動,所以也要偵測其它的ACH收費項目
                aryUpdCitemCode = Split(Replace(strCitemCode, "'", ""), ",")
                aryUpdCitemName = Split(Replace(strCitemName, "'", ""), ",")
                If Not IsNull(rsSO106("CitemStr").OriginalValue) Then
                    '#4260 如果對應不到SO003.SeqNO程式會出錯,所以增加判斷是否Eof By Kin 2008/12/05
                    Set rs = gcnGi.Execute(strQry1)
                    strOldCitemCode = Empty
                    If Not rs.EOF Then
                        'strOldCitemCode = gcnGi.Execute(strQry1).GetString(, , , ",")
                        strOldCitemCode = rs.GetString(, , , ",")
                    End If
                    Set rs = gcnGi.Execute(strQry2)
                    If Not rs.EOF Then
                        'strOldCitemName = gcnGi.Execute(strQry2).GetString(, , , ",")
                        strOldCitemName = rs.GetString(, , , ",")
                    End If
                End If
                '***********************************************************************************
                '有增加或減少的ACH以最多的收費項目為準,尋找該異動過ACH的收費項目
                If Right(strOldCitemCode, 1) = "," Then
                    strOldCitemCode = Mid(strOldCitemCode, 1, Len(strOldCitemCode) - 1)
                End If
                If Len(strOldCitemCode) > Len(Replace(strCitemCode, "'", "")) Then
                    aryCitemCode = Split(strOldCitemCode, ",")
                    aryCitemName = Split(strOldCitemName, ",")
                Else
                    aryCitemCode = Split(Replace(strCitemCode, "'", ""), ",")
                    aryCitemName = Split(Replace(strCitemName, "'", ""), ",")
                End If
                '**********************************************************************************
                
                'SO106的ACH項目大於畫面上所挑的ACH項目
                If UBound(aryOldACHTDESC) > UBound(aryACHTDESC) Then
                    lngUbound = UBound(aryOldACHTDESC)
                    ReDim aryCitem(lngUbound, 4)
                    ReDim aryACHTNO(lngUbound, 1)
                    For i = LBound(aryOldACHTDESC) To UBound(aryOldACHTDESC)
                        blnDiff = True
                        '填入沒異動過的ACH
                        For j = LBound(aryACHTDESC) To UBound(aryACHTDESC)
                            If aryOldACHTDESC(i) = aryACHTDESC(j) Then
                                aryCitem(lngAry, 0) = Split(Replace(strACHTNO, "'", ""), ",")(j)
                                aryCitem(lngAry, 1) = aryACHTDESC(j)
                                aryCitem(lngAry, 2) = "2"
                                lngAry = lngAry + 1
                                blnDiff = False
                                Exit For
                            End If
                        Next j
                        
                        '填入異動過的ACH
                        If blnDiff Then
                            aryACHTNO(lngAry, 0) = Split(Replace(.Fields("ACHTNO").OriginalValue, "'", ""), ",")(i)
                            aryACHTNO(lngAry, 1) = aryOldACHTDESC(i)
                            aryCitem(lngAry, 0) = aryACHTNO(lngAry, 0)
                            aryCitem(lngAry, 1) = aryACHTNO(lngAry, 1)
                            aryCitem(lngAry, 2) = "1"   '1代表新增一筆取消授權
                            lngAry = lngAry + 1
                        End If
                    Next i
                Else
                    lngUbound = UBound(aryACHTDESC)
                    strTmpACH = strACHTNO
                    ReDim aryCitem(lngUbound, 4)
                    ReDim aryACHTNO(lngUbound, 1)
                    For i = LBound(aryACHTDESC) To UBound(aryACHTDESC)
                        blnDiff = True
                        For j = LBound(aryOldACHTDESC) To UBound(aryOldACHTDESC)
                            If aryACHTDESC(i) = aryOldACHTDESC(j) Then
                                aryCitem(lngAry, 0) = Split(Replace(strACHTNO, "'", ""), ",")(i)
                                aryCitem(lngAry, 1) = aryOldACHTDESC(j)
                                aryCitem(lngAry, 2) = "2"
                                lngAry = lngAry + 1
                                blnDiff = False
                                Exit For
                            End If
                        Next j
                        If blnDiff Then
                            aryACHTNO(lngAry, 0) = Split(Replace(strACHTNO, "'", ""), ",")(i)
                            aryACHTNO(lngAry, 1) = aryACHTDESC(i)
                            aryCitem(lngAry, 0) = aryACHTNO(lngAry, 0)
                            aryCitem(lngAry, 1) = aryACHTNO(lngAry, 1)
                            aryCitem(lngAry, 2) = "0"
                            blnAddNew = True
                            lngAry = lngAry + 1
                        End If
                    Next i
                End If
            End If
        End With
    End If
    '異動過ACH,比對畫面上的ACH與對應的收費項目
    If blnUpdCitem Then
        For i = LBound(aryUpdCitemCode) To UBound(aryUpdCitemCode)
            '#4094 過濾停用要取消 By Kin 2008/09/17
'            Set rs = gcnGi.Execute("Select * From " & GetOwner & "CD068 Where " & _
'                                    "InStr(',' || CitemCodeStr || ',','," & aryUpdCitemCode(i) & ",')>0 And StopFlag<>1")
            '#4106 增加判斷ACHType=1參數 By Kin 2008/09/22
'            Set rs = gcnGi.Execute("Select * From " & GetOwner & "CD068 Where " & _
'                        "InStr(',' || CitemCodeStr || ',','," & aryUpdCitemCode(i) & ",')>0  And ACHTNO='" & aryCitem(j, 0) & "' And ACHType=1")
            
'            Set rs = gcnGi.Execute("Select * From " & GetOwner & "CD068 Where " & _
'                        "InStr(',' || CitemCodeStr || ',','," & aryUpdCitemCode(i) & ",')>0  And ACHTNO In(" & strTmpACH & ") And ACHType=1")
            '#7049測試不OK 改用CD068A去尋找 By Kin 2015/08/06
            If blnStartPost Then
                AchTypeCondition = " In (1,2) "
            Else
                AchTypeCondition = " In (1) "
            End If
            Set rs = gcnGi.Execute("Select CD068.* From " & GetOwner & "CD068A," & GetOwner & "CD068 Where " & _
                                         " Exists (Select * From " & GetOwner & "CD068 Where CD068A.BillHeadFmt = CD068.BillHeadFmt " & _
                                          " And ACHTNo In (" & strTmpACH & ") And ACHType " & AchTypeCondition & ")" & _
                                          " And CitemCode = " & aryCitemCode(i) & _
                                          " And CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                          IIf(strACHTDESC & "" = "", " And 1=1 ", " And CD068.ACHTDESC In (" & strACHTDESC & ")"))

            For j = LBound(aryCitem) To UBound(aryCitem)
                If aryCitem(j, 0) = rs("ACHTNO") And aryCitem(j, 1) = rs("ACHTDESC") Then
                    aryCitem(j, 3) = IIf(aryCitem(j, 3) = Empty, aryUpdCitemCode(i), aryCitem(j, 3) & "," & aryUpdCitemCode(i))
                    aryCitem(j, 4) = IIf(aryCitem(j, 4) = Empty, aryUpdCitemName(i), aryCitem(j, 4) & "," & aryUpdCitemName(i))
                    Exit For
                
                End If
            Next j
        Next i
    
    End If
    '填入ACH所對應的收費項目,如果是減少ACH,則要尋找是"1"的條件,反之要尋找為0的
    For i = LBound(aryCitemCode) To UBound(aryCitemCode)
        '#4094 過濾停用要取消 By Kin 2008/09/17

'        Set rs = gcnGi.Execute("Select * From " & GetOwner & "CD068 Where " & _
'                                    "InStr(',' || CitemCodeStr || ',','," & aryCitemCode(i) & ",')>0 And ACHTNO In(" & strTmpACH & ") And ACHType=1")
        
        '#7049測試不OK 改用CD068A去尋找 By Kin 2015/08/06
            If blnStartPost Then
                AchTypeCondition = " In ( 1,2 ) "
            Else
                AchTypeCondition = " In ( 1 ) "
            End If
         Set rs = gcnGi.Execute("Select CD068.* From " & GetOwner & "CD068A," & GetOwner & "CD068 Where " & _
                                          " Exists (Select * From " & GetOwner & "CD068 Where CD068A.BillHeadFmt = CD068.BillHeadFmt " & _
                                           " And ACHTNo In (" & strTmpACH & ") And ACHType " & AchTypeCondition & ")" & _
                                           " And CitemCode = " & aryCitemCode(i) & _
                                           " And CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                           IIf(strACHTDESC & "" = "", " And 1 = 1", " And CD068.ACHTDESC In (" & strACHTDESC & ")"))


                                                    
                                    '"InStr(',' || CitemCodeStr || ',','," & aryCitemCode(i) & ",')>0 And ACHTNO In(" & strTmpACH & ") And ACHType=1")
        If Not rs.EOF Then
            For j = LBound(aryCitem) To UBound(aryCitem)
                If aryCitem(j, 0) = rs("ACHTNO") And aryCitem(j, 1) = rs("ACHTDESC") Then
                    If blnUpdCitem Then
                        If blnAddNew Then
'                            If aryCitem(j, 2) = "0" Then
'                                aryCitem(j, 3) = IIf(aryCitem(j, 3) = Empty, aryCitemCode(i), aryCitem(j, 3) & "," & aryCitemCode(i))
'                                aryCitem(j, 4) = IIf(aryCitem(j, 4) = Empty, aryCitemName(i), aryCitem(j, 4) & "," & aryCitemName(i))
'                                Exit For
'                            End If
                        Else
                            If aryCitem(j, 2) = "1" Then
                                aryCitem(j, 3) = IIf(aryCitem(j, 3) = Empty, aryCitemCode(i), aryCitem(j, 3) & "," & aryCitemCode(i))
                                aryCitem(j, 4) = IIf(aryCitem(j, 4) = Empty, aryCitemName(i), aryCitem(j, 4) & "," & aryCitemName(i))
                                Exit For
                            End If
                        End If
                    Else
                        aryCitem(j, 3) = IIf(aryCitem(j, 3) = Empty, aryCitemCode(i), aryCitem(j, 3) & "," & aryCitemCode(i))
                        aryCitem(j, 4) = IIf(aryCitem(j, 4) = Empty, aryCitemName(i), aryCitem(j, 4) & "," & aryCitemName(i))
                        Exit For
                    End If
                End If
            Next j

        End If
    Next i
Fin:
    On Error Resume Next
    Call CloseRecordset(rs)
    Erase aryACHTDESC
    Erase aryACHTNO
    Erase aryCitemCode
    Erase aryCitemName
    Erase aryOldACHTDESC
    Erase aryUpdCitemCode
    Erase aryUpdCitemName
  Exit Sub
ChkErr:
  ErrSub Me.Name, "GetCitemCode"
End Sub
'#3946 取得SO106.MasterId的Sequence By Kin 2008/05/30
Private Function Get_MasterID_Seq() As String
  On Error GoTo ChkErr
    Get_MasterID_Seq = RPxx(gcnGi.Execute("SELECT " & GetOwner & "S_SO106_MasterId.NEXTVAL FROM DUAL").GetString & "")
    'Get_Cmd_Seq_No = Format(Get_Cmd_Seq_No, "00000000")
  Exit Function
ChkErr:
    ErrSub "mod_SysLib", "Get_MasterID_Seq"
End Function

