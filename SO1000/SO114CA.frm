VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.17#0"; "csMulti.ocx"
Begin VB.Form frmSO114CA 
   Caption         =   "認證資訊管理 [SO114CA]"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   1395
   ClientWidth     =   11895
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO114CA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraDate 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4755
      Left            =   90
      TabIndex        =   46
      Top             =   -30
      Width           =   11715
      Begin VB.TextBox txtAccountNo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8190
         MaxLength       =   16
         TabIndex        =   39
         Top             =   3825
         Width           =   2850
      End
      Begin VB.ComboBox cboSextual 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "SO114CA.frx":0442
         Left            =   5220
         List            =   "SO114CA.frx":044C
         TabIndex        =   7
         Top             =   1095
         Width           =   1215
      End
      Begin VB.TextBox txtInvAddress 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1680
         MaxLength       =   90
         TabIndex        =   24
         Top             =   4380
         Width           =   4760
      End
      Begin VB.TextBox txtInvTitle 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   23
         Top             =   4095
         Width           =   4760
      End
      Begin VB.ComboBox cboInvoiceType 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "SO114CA.frx":0458
         Left            =   1680
         List            =   "SO114CA.frx":0465
         Style           =   2  '單純下拉式
         TabIndex        =   21
         Top             =   3780
         Width           =   1635
      End
      Begin VB.CommandButton cmdFind 
         Height          =   315
         Left            =   6060
         Picture         =   "SO114CA.frx":0485
         Style           =   1  '圖片外觀
         TabIndex        =   42
         Top             =   150
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.ComboBox cboStatus 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   2  '關閉
         ItemData        =   "SO114CA.frx":0A67
         Left            =   8190
         List            =   "SO114CA.frx":0A71
         TabIndex        =   35
         Top             =   2910
         Width           =   1950
      End
      Begin VB.ComboBox cboRente 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   2  '關閉
         ItemData        =   "SO114CA.frx":0A81
         Left            =   8190
         List            =   "SO114CA.frx":0A91
         TabIndex        =   33
         Top             =   2280
         Width           =   2850
      End
      Begin VB.ComboBox cboJob 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   2  '關閉
         ItemData        =   "SO114CA.frx":0AC3
         Left            =   8190
         List            =   "SO114CA.frx":0ADF
         TabIndex        =   32
         Top             =   1965
         Width           =   2850
      End
      Begin VB.ComboBox cboMemberAge 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   2  '關閉
         ItemData        =   "SO114CA.frx":0B28
         Left            =   8190
         List            =   "SO114CA.frx":0B35
         TabIndex        =   31
         Top             =   1650
         Width           =   2850
      End
      Begin VB.TextBox txtMemberID 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   15
         Top             =   2895
         Width           =   1845
      End
      Begin VB.ComboBox cboEducation 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   2  '關閉
         ItemData        =   "SO114CA.frx":0B4D
         Left            =   8190
         List            =   "SO114CA.frx":0B5D
         TabIndex        =   27
         Top             =   780
         Width           =   2850
      End
      Begin VB.ComboBox cboMarried 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   2  '關閉
         ItemData        =   "SO114CA.frx":0B85
         Left            =   8190
         List            =   "SO114CA.frx":0B8F
         TabIndex        =   29
         Top             =   1335
         Width           =   1095
      End
      Begin VB.CheckBox chkMessage 
         Caption         =   "有線電視服務及頻道節目"
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
         Left            =   8190
         TabIndex        =   28
         Top             =   1125
         Width           =   2355
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3285
         MaxLength       =   90
         TabIndex        =   10
         Top             =   1680
         Width           =   3150
      End
      Begin VB.TextBox txtTown 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4185
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1395
         Width           =   2250
      End
      Begin VB.TextBox txtAuthenticID 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Top             =   180
         Width           =   2295
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtMobile 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4770
         MaxLength       =   20
         TabIndex        =   2
         Top             =   475
         Width           =   1665
      End
      Begin VB.TextBox txtTel 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   3
         Top             =   795
         Width           =   1995
      End
      Begin VB.TextBox txtID 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   2  '關閉
         Left            =   3660
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1095
         Width           =   1125
      End
      Begin VB.TextBox txtEmail 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   14
         Top             =   2595
         Width           =   4760
      End
      Begin VB.TextBox txtCity 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1395
         Width           =   1140
      End
      Begin VB.TextBox txtCustID 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5370
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1065
      End
      Begin VB.TextBox txtFaciSno 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3975
         MaxLength       =   20
         TabIndex        =   41
         Top             =   180
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.TextBox txtPPVOrderPwd 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4590
         MaxLength       =   20
         TabIndex        =   16
         Top             =   2895
         Width           =   1845
      End
      Begin VB.TextBox txtLoginPwd 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4590
         MaxLength       =   20
         TabIndex        =   18
         Top             =   3190
         Width           =   1845
      End
      Begin VB.TextBox txtAccBefore 
         Alignment       =   1  '靠右對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   19
         Top             =   3490
         Width           =   1545
      End
      Begin VB.TextBox txtAccBalance 
         Alignment       =   1  '靠右對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4590
         MaxLength       =   8
         TabIndex        =   20
         Top             =   3490
         Width           =   1845
      End
      Begin VB.TextBox txtTel2 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4770
         MaxLength       =   20
         TabIndex        =   4
         Top             =   795
         Width           =   1665
      End
      Begin VB.TextBox txtHintAnswer 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8190
         MaxLength       =   20
         TabIndex        =   26
         Top             =   481
         Width           =   2850
      End
      Begin VB.TextBox txtBodyCount 
         Alignment       =   1  '靠右對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10170
         MaxLength       =   3
         TabIndex        =   30
         Top             =   1335
         Width           =   855
      End
      Begin VB.TextBox txtClassId 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   34
         Top             =   2610
         Width           =   1950
      End
      Begin VB.TextBox txtHintName 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8190
         MaxLength       =   20
         TabIndex        =   25
         Top             =   180
         Width           =   2850
      End
      Begin VB.TextBox txtLoginID 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   17
         Top             =   3190
         Width           =   1845
      End
      Begin prjGiList.GiList gilCompCode 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   1980
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   503
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth1       =   550
         FldWidth2       =   2000
         F2Corresponding =   -1  'True
      End
      Begin Gi_Date.GiDate gdaBirthday 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   1095
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
      End
      Begin Gi_Date.GiDate gdaCreateDate 
         Height          =   285
         Left            =   8190
         TabIndex        =   36
         Top             =   3240
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
      End
      Begin Gi_Date.GiDate gdaUseDate 
         Height          =   285
         Left            =   10140
         TabIndex        =   37
         Top             =   3240
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
      End
      Begin Gi_Date.GiDate gdaLastLoginDate 
         Height          =   285
         Left            =   8190
         TabIndex        =   40
         Top             =   4125
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
         Enabled         =   0   'False
      End
      Begin MSMask.MaskEdBox mskInvNo 
         Height          =   285
         Left            =   4590
         TabIndex        =   22
         Top             =   3795
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "########"
         PromptChar      =   "_"
      End
      Begin CS_Multi.CSmulti csmFaciSno 
         Height          =   345
         Left            =   150
         TabIndex        =   13
         Top             =   2250
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   609
         ButtonCaption   =   "認證設備序號"
      End
      Begin prjGiList.GiList gilBankCode 
         Height          =   270
         Left            =   8190
         TabIndex        =   38
         Top             =   3540
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   476
         Enabled         =   0   'False
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
      Begin VB.Label lblBankCode 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "銀行"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   7770
         TabIndex        =   89
         Top             =   3585
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "帳號/卡號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   7365
         TabIndex        =   88
         Top             =   3870
         Width           =   765
      End
      Begin VB.Label lblInvAddress 
         AutoSize        =   -1  'True
         Caption         =   "發票地址"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   900
         TabIndex        =   87
         Top             =   4425
         Width           =   720
      End
      Begin VB.Label lblInvTitle 
         AutoSize        =   -1  'True
         Caption         =   "發票抬頭"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   900
         TabIndex        =   86
         Top             =   4155
         Width           =   720
      End
      Begin VB.Label lblInvoiceType 
         AutoSize        =   -1  'True
         Caption         =   "發票種類"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   900
         TabIndex        =   85
         Top             =   3855
         Width           =   720
      End
      Begin VB.Label lblLoginID 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "客服網站登入帳號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   180
         TabIndex        =   84
         Top             =   3255
         Width           =   1440
      End
      Begin VB.Label lblAccBefore 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "結算前交易限額"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   360
         TabIndex        =   83
         Top             =   3555
         Width           =   1260
      End
      Begin VB.Label lblBirthday 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "出生日期"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   900
         TabIndex        =   82
         Top             =   1155
         Width           =   720
      End
      Begin VB.Label lblTel 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "聯絡電話(宅)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   600
         TabIndex        =   81
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label lblName 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "會員姓名"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   900
         TabIndex        =   80
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblCompCode 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "認證公司別"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   720
         TabIndex        =   79
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1215
         TabIndex        =   78
         Top             =   2670
         Width           =   405
      End
      Begin VB.Label lblAuthenticId 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "認證編號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   900
         TabIndex        =   77
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblAddress 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "通訊地址"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   900
         TabIndex        =   76
         Top             =   1470
         Width           =   720
      End
      Begin VB.Label lblMemberID 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "Web會員編號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   570
         TabIndex        =   75
         Top             =   2970
         Width           =   1050
      End
      Begin VB.Label lblInvNo 
         AutoSize        =   -1  'True
         Caption         =   "統一編號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3840
         TabIndex        =   74
         Top             =   3855
         Width           =   720
      End
      Begin VB.Label lblUpdTime 
         AutoSize        =   -1  'True
         Caption         =   "ee/mm/dd hh:mm:ss"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8220
         TabIndex        =   73
         Top             =   4425
         Width           =   1725
      End
      Begin VB.Label lblAddr 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "(路段街巷弄衖號樓)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   1650
         TabIndex        =   72
         Top             =   1740
         Width           =   1560
      End
      Begin VB.Label Label2 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "(鄉鎮市區)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   3330
         TabIndex        =   71
         Top             =   1460
         Width           =   840
      End
      Begin VB.Label lblCity 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "(縣市)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   1650
         TabIndex        =   70
         Top             =   1460
         Width           =   480
      End
      Begin VB.Label lblUpdTimeB 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "最後異動日期"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7050
         TabIndex        =   69
         Top             =   4455
         Width           =   1080
      End
      Begin VB.Label lblSextual 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "性別"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4830
         TabIndex        =   68
         Top             =   1155
         Width           =   360
      End
      Begin VB.Label lblID 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "身份證號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2850
         TabIndex        =   67
         Top             =   1155
         Width           =   720
      End
      Begin VB.Label lblCustId 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "認證客編"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4605
         TabIndex        =   66
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label lblMobile 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "手機號碼"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4005
         TabIndex        =   65
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblPPVOrderPwd 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "PPV訂購密碼"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3555
         TabIndex        =   64
         Top             =   2970
         Width           =   1020
      End
      Begin VB.Label lblLoginPassword 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "登入密碼"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3840
         TabIndex        =   63
         Top             =   3255
         Width           =   720
      End
      Begin VB.Label lblAccBalance 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "結算後交易限額"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3300
         TabIndex        =   62
         Top             =   3555
         Width           =   1260
      End
      Begin VB.Label lblTel2 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "聯絡電話(公)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3720
         TabIndex        =   61
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label lblHintAnswer 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "會員密碼提示答案"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6690
         TabIndex        =   60
         Top             =   538
         Width           =   1440
      End
      Begin VB.Label lblEducation 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "學歷"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7770
         TabIndex        =   59
         Top             =   836
         Width           =   360
      End
      Begin VB.Label lblMessage 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "願意收到相關訊息"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6690
         TabIndex        =   58
         Top             =   1134
         Width           =   1440
      End
      Begin VB.Label lblMarried 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "婚姻狀況"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7410
         TabIndex        =   57
         Top             =   1425
         Width           =   720
      End
      Begin VB.Label lblBodyCount 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "同住人數"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   9390
         TabIndex        =   56
         Top             =   1395
         Width           =   720
      End
      Begin VB.Label lblMemberAge 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "同住成員"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7410
         TabIndex        =   55
         Top             =   1730
         Width           =   720
      End
      Begin VB.Label lblJob 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "職業別"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7590
         TabIndex        =   54
         Top             =   2025
         Width           =   540
      End
      Begin VB.Label lblRente 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "個人年收入"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7230
         TabIndex        =   53
         Top             =   2355
         Width           =   900
      End
      Begin VB.Label lblClassId 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "身份別"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7590
         TabIndex        =   52
         Top             =   2655
         Width           =   540
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "帳號狀態"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7410
         TabIndex        =   51
         Top             =   2985
         Width           =   720
      End
      Begin VB.Label lblCreateDate 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "申請日期時間"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7050
         TabIndex        =   50
         Top             =   3285
         Width           =   1080
      End
      Begin VB.Label lblHintName 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "會員密碼提示問題"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6690
         TabIndex        =   49
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label lblUseDate 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "啟用日期"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   9360
         TabIndex        =   48
         Top             =   3285
         Width           =   720
      End
      Begin VB.Label lblLastLoginDate 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "最後登入日期"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7050
         TabIndex        =   47
         Top             =   4170
         Width           =   1080
      End
   End
   Begin VB.ComboBox cboAccountNo 
      Enabled         =   0   'False
      Height          =   315
      IMEMode         =   2  '關閉
      ItemData        =   "SO114CA.frx":0B9F
      Left            =   -330
      List            =   "SO114CA.frx":0BA9
      TabIndex        =   44
      Top             =   285
      Visible         =   0   'False
      Width           =   390
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   2250
      Left            =   90
      TabIndex        =   43
      Top             =   4785
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   3969
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblAccountNo 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "指定付款帳號"
      Enabled         =   0   'False
      Height          =   165
      Left            =   -1080
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
End
Attribute VB_Name = "frmSO114CA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmer: Hammer
' Last Modify: 2005/05/07

Option Explicit

Private lngEditMode As giEditModeEnu ' 記錄目前在編輯、新增或檢視模式
Private WithEvents rsSO002B As ADODB.Recordset
Attribute rsSO002B.VB_VarHelpID = -1

' 自訂屬性：EditMode
' 記錄目前在編輯、新增或檢視模式
' giEditModeEnu(自訂列舉值，設定於Sys_Lib)
Public Property Get EditMode() As giEditModeEnu
  On Error GoTo ChkErr
    EditMode = lngEditMode    ' 取目前編輯模式
  Exit Property
ChkErr:
    Call ErrSub(Name, "Get EditMode")
End Property

Public Property Let EditMode(ByVal vNewValue As giEditModeEnu)
  On Error GoTo ChkErr
    lngEditMode = vNewValue '設定編輯模式
  Exit Property
ChkErr:
    Call ErrSub(Name, "Let EditMode")
End Property

'根據傳來之編輯模式, 設定各物件屬性
Private Sub ChangeMode(ByVal lngMode As giEditModeEnu)
  On Error GoTo ChkErr
    Dim blnFlag As Boolean  ' 記錄是否在資料瀏覽模式，
    lngEditMode = lngMode
    Select Case lngMode
           Case giEditModeInsert
                blnFlag = False
                MenuEnabled False, False, False, True, False, True, False, False, False, False, False
           Case giEditModeEdit
                blnFlag = False
                MenuEnabled False, False, False, True, False, True, False, False, False, False, False
           Case giEditModeView
                blnFlag = True
                MenuEnabled True, rsSO002B.RecordCount > 0, False, False, False, True, False, True, True, True, True
    End Select
    If lngMode <> giEditModeView Then
        Dim strFilter As String
        On Error Resume Next
        strFilter = ""
        ' " AND CUSTID=" & gCustId &
        If lngMode = giEditModeInsert Then
            strFilter = RPxx(gcnGi.Execute("SELECT FACISNO FROM " & GetOwner & "SO002B" & _
                                                            " WHERE FACISNO IS NOT NULL").GetString(adClipString, , "", ",", ""))
        Else
            strFilter = RPxx(gcnGi.Execute("SELECT FACISNO FROM " & GetOwner & "SO002B" & _
                                                            " WHERE ROWID <> '" & rsSO002B!ROWID & "'" & _
                                                            " AND FACISNO IS NOT NULL").GetString(adClipString, , "", ",", ""))
        End If
        If Err = 3021 Then
            Err.Clear
        Else
            If Len(strFilter) > 0 Then strFilter = " AND SEQNO NOT IN ( " & Left(strFilter, Len(strFilter) - 1) & " )"
        End If
        SetMsQry csmFaciSno, "SELECT SERVICETYPE,SEQNO,FACISNO,SMARTCARDNO,FACINAME FROM " & GetOwner & "SO004" & _
                                            " WHERE CUSTID=" & gCustId & " AND FACISNO IS NOT NULL " & _
                                            " AND ( INSTDATE IS NOT NULL AND PRDATE IS NULL )" & strFilter & GetUnionStr, _
                                            "服務別", "設備流水號", , , True, "設備序號", "智慧卡序號", "設備名稱", , , , , , , , _
                                            900, 1700, 1700, 2120, 2000, , , , , , , , , , , , , , , , , , , , 2
    Else
        NoFilterFaci
    End If
    fraDate.Enabled = Not blnFlag
    ggrData.Enabled = blnFlag
  Exit Sub
ChkErr:
    Call ErrSub(Name, "ChangeMode")
End Sub

'   cable_penny:
'   CSmulti 是顯示 SERVICETYPE AS 服務別,FACISNO AS 設備序號,SMARTCARDNO AS 卡號,
'   FACINAME AS 設備名稱,FACISEQNO AS 設備流程號, 然後回存FACISEQNO
'   服務別,設備流水號,設備序號,卡號,設備名稱

Private Sub NoFilterFaci()
  On Error GoTo ChkErr
    SetMsQry csmFaciSno, "SELECT SERVICETYPE,SEQNO,FACISNO,SMARTCARDNO,FACINAME FROM " & GetOwner & "SO004" & _
                                        " WHERE CUSTID=" & gCustId & " AND FACISNO IS NOT NULL " & _
                                        " AND ( INSTDATE IS NOT NULL AND PRDATE IS NULL )" & GetUnionStr, _
                                        "服務別", "設備流水號", , , True, "設備序號", "智慧卡序號", "設備名稱", , , , , , , , _
                                        900, 1700, 1700, 2120, 2000, , , , , , , , , , , , , , , , , , , , 2

'    SetMsQry csmFaciSno, "SELECT SMARTCARDNO,FACISNO,FACINAME FROM " & GetOwner & "SO004" & _
                                        " WHERE CUSTID=" & gCustId & " AND FACISNO IS NOT NULL " & _
                                        " AND ( INSTDATE IS NOT NULL AND PRDATE IS NULL )", _
                                        "智慧卡序號", "設備序號", , , , "項目名稱", , , , , , , , , , 1650, 1880, 1860, , , , , , , , , , , , , , , , , , , , , , 2
  Exit Sub
ChkErr:
    ErrSub Name, "NoFilterFaci"
End Sub

Private Function GetUnionStr() As String
    GetUnionStr = " UNION ALL SELECT 'C' SERVICETYPE,'" & gCustId & "' SEQNO,NULL FACISNO," & _
                            "NULL SMARTCARDNO,'CATV服務' FACINAME FROM DUAL"
End Function

Private Sub RcdToScr()
    On Error GoTo ChkErr
    '將資料檔讀取之欄位資料轉換成螢幕上物件內容
        With rsSO002B
                txtAuthenticID = .Fields("AuthenticID") & ""
                txtName = .Fields("Name") & ""
                txtTel = .Fields("Tel") & ""
                txtMobile = .Fields("Mobile") & ""
                If IsNull(.Fields("BirthdayYYYY")) Then
                    gdaBirthday.Text = ""
                Else
                    gdaBirthday.Text = .Fields("BirthdayYYYY") & "/" & .Fields("BirthdayMM") & "/" & .Fields("BirthdayDD") & ""
                End If
                txtID = .Fields("ID") & ""
                cboSextual = .Fields("Sextual") & ""
                txtAddress = .Fields("Address") & ""
                gilCompCode.SetCodeNo .Fields("CompCode") & ""
                gilCompCode.Query_Description
                txtCustID = .Fields("CustId") & ""
                txtEmail = .Fields("Email") & ""
'                cboAccountNo.Text = (.Fields("AccountNo") & "")
                
                ' 問題 : 2229 ( Penny 提的 )
                ' 請將PPV訂購密碼及訂購密碼加密顯示,只顯示頭尾1碼數字.
                
                txtPPVOrderPwd.Tag = .Fields("PPVOrderPwd") & ""
                txtPPVOrderPwd.Text = GetPwdFmt(txtPPVOrderPwd.Tag)
                
                txtLoginPwd.Tag = .Fields("LoginPassword") & ""
                txtLoginPwd.Text = GetPwdFmt(txtLoginPwd.Tag)

                txtAccBefore = (.Fields("AccountingBefore") & "")
                txtAccBalance = (.Fields("AccountingBalance") & "")
                
                txtMemberID = (.Fields("MemberID") & "")
                txtHintName = (.Fields("HintName") & "")
                txtTel2 = (.Fields("Tel2") & "")
                txtCity = (.Fields("City") & "")
                txtTown = (.Fields("Town") & "")
                txtLoginID = (.Fields("LoginID") & "")
                txtHintAnswer = (.Fields("HintAnswer") & "")
                cboEducation = (.Fields("Education") & "")
                chkMessage.Value = Val(.Fields("Message") & "")
                cboMarried = (.Fields("Married") & "")
                txtBodyCount = (.Fields("BodyCount") & "")
                cboMemberAge = (.Fields("MemberAge") & "")
                cboJob = (.Fields("Job") & "")
                cboRente = (.Fields("Rente") & "")
                txtClassId = (.Fields("ClassID") & "")
                cboStatus = (.Fields("Status") & "")
                gdaCreateDate.Text = (.Fields("CreateDate") & "")
                gdaUseDate.Text = (.Fields("UseDate") & "")
                gdaLastLoginDate.Text = (.Fields("LastLoginDate") & "")
                lblUpdTime = (.Fields("UpdTime") & "")
                
                ' 2006/04/13 By Hammer
                ' 問題集編號 : 2166
                ' CSV5.2_SO114CA_20060311_Penny_認證帳號設備可複選.doc
'                txtFaciSno = .Fields("FaciSNo") & ""
                If .Fields("FaciSNo") & "" = Empty Then
                    csmFaciSno.Clear
                Else
                    csmFaciSno.SetQueryCode .Fields("FaciSNo") & ""
                End If
                Select Case (.Fields("InvoiceType") & "")
                        Case 0
                                cboInvoiceType.ListIndex = 0
                        Case 2
                                cboInvoiceType.ListIndex = 1
                        Case 3
                                cboInvoiceType.ListIndex = 2
                        Case Else
                                cboInvoiceType.ListIndex = -1
                End Select
                txtInvTitle = .Fields("InvTitle") & ""
                mskInvNo = .Fields("InvNo") & ""
                txtInvAddress = .Fields("InvAddress") & ""
                
                ' 讀取 SO002A 的銀行帳號相關資料
                Read_2A_Data
        End With
    Exit Sub
ChkErr:
    If Err.Number = 0 Then
        Err.Clear
        Resume Next
    Else
        Call ErrSub(Name, "RcdToScr")
    End If
End Sub

Private Sub Read_2A_Data()
  On Error GoTo ChkErr
    ' 於SO002B中加秀銀行代碼、名稱、帳號。
    ' (SELECT BANKCODE,BANKNAME,ACCOUNTID FROM SO002A WHERE AUTHENTICID=?本身認證ID")。
    Dim rs2A As New ADODB.Recordset
    Dim strQry As String
    '#3541 Like多了一個單引號 By Kin 2007/11/30
'    strQry = "SELECT BANKCODE,BANKNAME,ACCOUNTNO FROM " & GetOwner & "SO002A" & _
'                    " WHERE AUTHENTICID LIKE '%''" & rsSO002B!AUTHENTICID & "''%'"
    strQry = "SELECT BANKCODE,BANKNAME,ACCOUNTNO FROM " & GetOwner & "SO002A" & _
                    " WHERE AUTHENTICID LIKE '%" & rsSO002B!AuthenticID & "%'"

    If GetRS(rs2A, strQry, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then
        If Not rs2A.EOF Then
            gilBankCode.SetCodeNo rs2A!BankCode & ""
            gilBankCode.SetDescription rs2A!BankName & ""
            txtAccountNo = rs2A!AccountNo & ""
        End If
    End If
    On Error Resume Next
    CloseRS rs2A
  Exit Sub
ChkErr:
    ErrSub Name, "Read_2A_Data"
End Sub

Private Function GetPwdFmt(strValue As String) As String
  On Error GoTo ChkErr
    If strValue = Empty Then
        GetPwdFmt = ""
    Else
        If Len(strValue) >= 4 Then
            GetPwdFmt = Left(strValue, 2) & String(Len(strValue) - 3, "X") & Right(strValue, 1)
        Else
            Select Case Len(strValue)
                    Case 1
                        GetPwdFmt = "X"
                    Case 2
                        GetPwdFmt = Left(strValue, 1) & "X"
                    Case 3
                        GetPwdFmt = Left(strValue, 1) & "XX"
            End Select
        End If
    End If
  Exit Function
ChkErr:
    ErrSub Name, "GetPwdFmt"
End Function


Private Function ScrToRcd() As Boolean
  On Error GoTo ChkErr
   '將螢幕上物件內容轉換成資料檔對應欄位內容
   
'    認證編號由程式產生為唯一值規則(公司別3碼[右靠左補零]+流水號7碼[右靠左補零])
'    流水號需讀取Seqence Object S_SO002B_AuthenticId產生
'    Ex: 公司別=1, 流水號為123; 則認證編號='0010000123'
    
    With rsSO002B
         If EditMode = giEditModeInsert Then
            .AddNew
            .Fields("AuthenticID") = GetRsValue("SELECT " & GetOwner & "S_SO002B_AUTHENTICID.NEXTVAL AUTHENTICID FROM DUAL", gcnGi)
            .Fields("AuthenticID") = Right("0000000" & .Fields("AuthenticID"), 7)
            .Fields("AuthenticID") = Right("000" & gilCompCode.GetCodeNo, 3) & .Fields("AuthenticID")
         Else
'            ' 2006/01/12 By Hammer ( Review Spec )
''            修改模式存檔後需將該筆異動資料Insert into so002blog, 程序如下 :
''            1. delete from so002blog where AuthenticId=so002b.AuthenticId
''            2. Insert into so002blog (select * from so002b where AuthenticId=[該筆異動資料之AuthenticId變數]
'            gcnGi.Execute "DELETE FROM " & GetOwner & "SO002BLOG WHERE AUTHENTICID='" & txtAuthenticId & "'"
'            gcnGi.Execute "INSERT INTO " & GetOwner & "SO002BLOG ( SELECT * FROM " & GetOwner & "SO002B WHERE AUTHENTICID='" & txtAuthenticId & "' )"
         End If
        .Fields("Name") = SolveNull(txtName)
        .Fields("Tel") = SolveNull(txtTel)
        .Fields("Mobile") = SolveNull(txtMobile)
         If gdaBirthday.Text <> Empty Then
            .Fields("BirthdayYYYY") = Format(gdaBirthday.Text, "YYYY")
            .Fields("BirthdayMM") = Format(gdaBirthday.Text, "MM")
            .Fields("BirthdayDD") = Format(gdaBirthday.Text, "DD")
         Else
            .Fields("BirthdayYYYY") = Null
            .Fields("BirthdayMM") = Null
            .Fields("BirthdayDD") = Null
         End If
        .Fields("ID") = SolveNull(txtID)
        .Fields("Sextual") = cboSextual
        .Fields("Address") = SolveNull(txtAddress)
        .Fields("CompCode") = gilCompCode.GetCodeNo
        .Fields("CustId") = SolveNull(txtCustID)
        .Fields("Email") = SolveNull(txtEmail)
'        .Fields("AccountNo") = SolveNull(cboAccountNo.Text)

        If txtPPVOrderPwd.Tag = Empty And txtPPVOrderPwd.Text <> Empty Then
            txtPPVOrderPwd.Tag = txtPPVOrderPwd.Text
        End If
        If txtPPVOrderPwd.Text <> Empty And txtPPVOrderPwd.Text <> GetPwdFmt(txtPPVOrderPwd.Tag) Then
            txtPPVOrderPwd.Tag = txtPPVOrderPwd.Text
        End If
        .Fields("PPVOrderPwd") = SolveNull(txtPPVOrderPwd.Tag)
        
        If txtLoginPwd.Tag = Empty And txtLoginPwd.Text <> Empty Then
            txtLoginPwd.Tag = txtLoginPwd.Text
        End If
        If txtLoginPwd.Text <> Empty And txtLoginPwd.Text <> GetPwdFmt(txtLoginPwd.Tag) Then
            txtLoginPwd.Tag = txtLoginPwd.Text
        End If
        .Fields("LoginPassword") = SolveNull(txtLoginPwd.Tag)
        
        .Fields("AccountingBefore") = SolveNull(txtAccBefore)
        .Fields("AccountingBalance") = SolveNull(txtAccBalance)
        .Fields("MemberID") = SolveNull(txtMemberID)
        .Fields("HintName") = SolveNull(txtHintName)
        .Fields("Tel2") = SolveNull(txtTel2)
        .Fields("City") = SolveNull(txtCity)
        .Fields("Town") = SolveNull(txtTown)
        .Fields("LoginID") = SolveNull(txtLoginID)
        .Fields("HintAnswer") = SolveNull(txtHintAnswer)
        .Fields("Education") = SolveNull(cboEducation)
        .Fields("Message") = Abs(chkMessage.Value)
        .Fields("Married") = cboMarried
        .Fields("BodyCount") = Val(txtBodyCount)
        .Fields("MemberAge") = SolveNull(cboMemberAge)
        .Fields("Job") = SolveNull(cboJob)
        .Fields("Rente") = SolveNull(cboRente)
        .Fields("ClassID") = SolveNull(txtClassId)
        .Fields("Status") = SolveNull(cboStatus)
        .Fields("CreateDate") = SolveNull(gdaCreateDate.Text)
        .Fields("UseDate") = SolveNull(gdaUseDate.Text)
        .Fields("LastLoginDate") = SolveNull(gdaLastLoginDate.Text)
        .Fields("UpdTime") = GetDTString(RightNow)
        
        ' 2006/04/13 By Hammer
        ' 問題集編號 : 2166
        ' CSV5.2_SO114CA_20060311_Penny_認證帳號設備可複選.doc
        
        Select Case cboInvoiceType.ListIndex
               Case 0
                    .Fields("InvoiceType") = 0
               Case 1
                    .Fields("InvoiceType") = 2
               Case 2
                    .Fields("InvoiceType") = 3
               Case Else
                    .Fields("InvoiceType") = Null
        End Select
        .Fields("InvTitle") = IIf(txtInvTitle = "", Null, txtInvTitle)
        .Fields("InvNo") = IIf(mskInvNo = "", Null, mskInvNo)
        .Fields("InvAddress") = IIf(txtInvAddress = "", Null, txtInvAddress)
'        .Fields("FaciSNo") = SolveNull(txtFaciSno)
        .Fields("FaciSNo") = csmFaciSno.GetQueryCode
        .Update
    End With
    ScrToRcd = True
  Exit Function
ChkErr:
    Call ErrSub(Name, "ScrToRcd")
    On Error Resume Next
    ScrToRcd = False
    gcnGi.RollbackTrans
End Function

' 移動資料錄指標時，在螢幕顯示最新的資料欄資料內含值
Public Sub FunctionKey(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    Select Case KeyCode
           '----------------------------------------------------
           Case vbKeyF2  '   F3:存檔, 相當於按下cmdsave
                    If Not ChkGiList(KeyCode) Then Exit Sub
                    UpdateGo
           Case vbKeyEscape
                    CancelGo
    End Select
End Sub

Private Function IsDataOK()
  On Error GoTo ChkErr
  
   ' 必要欄位: '認證編號', '會員姓名', '認證公司別', '認證客戶編號', '認證設備序號'且認證編號為唯一值
   ' 檢查必要欄位是否有值, 若為必要欄位且無值, 則顯示"欄位必需有值",
   ' 且focus移到該欄位上
    IsDataOK = False
'    If txtAuthenticID = Empty Then
'        MsgBox "怪怪 ! [認證編號] 怎能為空白！", vbExclamation, "訊息"
'        Exit Function
'    End If
    If txtName = Empty Then
        MsgBox "[姓名] 不能為空白！", vbExclamation, "訊息"
        On Error Resume Next
        txtName.SetFocus
        Exit Function
    End If
    If gilCompCode.GetCodeNo = Empty Then
        MsgBox "[認證公司別] 不能為空白！", vbExclamation, "訊息"
        On Error Resume Next
        gilCompCode.SetFocus
        Exit Function
    End If
    If txtCustID = Empty Then
        MsgBox "[認證客戶編號] 不能為空白！", vbExclamation, "訊息"
        On Error Resume Next
        txtCustID.SetFocus
        Exit Function
    End If
'    If txtFaciSno = Empty Then
    If csmFaciSno.GetDispStr = Empty Then
        MsgBox "[認證設備序號] 不能為空白！", vbExclamation, "訊息"
        On Error Resume Next
        csmFaciSno.SetFocus
        Exit Function
    End If
    If Not MustExist(cboSextual, , "性別") Then Exit Function
    
    IsDataOK = True
  Exit Function
66:
    MsgBox " 此為必要欄位,須有值 !! ", vbExclamation, "訊息.."
  Exit Function
ChkErr:
    Call ErrSub(Name, "IsDataOk")
End Function

Private Sub NewRcd()
  On Error Resume Next
    '新增資料時, 清除畫面物件內容, 或取新序號
    '取新序號: 若資料中有依序遞增之流水號, 則可於此處取得該號碼,
    '例: 代碼編號為數字型代碼檔的次一編號
    
    Dim objCtl As Control
    
    For Each objCtl In Me
        If TypeOf objCtl Is TextBox Then objCtl = ""
    Next
    
    gdaBirthday.Text = ""
    gdaCreateDate.Text = ""
    gdaUseDate.Text = ""
    gdaLastLoginDate.Text = ""
'    cboAccountNo = ""
    txtAccBefore = Val(GetSystemParaItem("AccountingBefore", gCompCode, gServiceType, "SO043", , True) & Empty)
    txtAccBalance = Val(GetSystemParaItem("AccountingBalance", gCompCode, gServiceType, "SO043", , True) & Empty)
  
    cboSextual.ListIndex = 1
    cboMarried = ""
    cboEducation = ""
    cboMemberAge = ""
    cboJob = ""
    cboRente = ""
    cboStatus = ""
    
    lblUpdTime = ""
    cboInvoiceType.ListIndex = 0
    mskInvNo = ""
    txtInvTitle = ""
    txtInvAddress = ""
    csmFaciSno.Clear
    
    txtPPVOrderPwd.Tag = ""
    txtLoginPwd.Tag = ""
    
    chkMessage.Value = 0
    
  Exit Sub
ChkErr:
    Call ErrSub(Name, "NewRcd")
End Sub

Public Sub AddNewGo()
  On Error GoTo ChkErr
   '進入新增模式, 清除物件內為初始狀態, 游標停在第一個可編輯物件
    ChangeMode (giEditModeInsert)
    NewRcd
    gilCompCode.SetCodeNo gCompCode
    gilCompCode.Query_Description
    txtCustID = gCustId
'    lblUpdTime = GetDTString(RightNow)
    On Error Resume Next
    txtName.SetFocus
  Exit Sub
ChkErr:
   Call ErrSub(Name, "AddNewGo")
End Sub

Public Sub UpdateGo()
  On Error GoTo ChkErr
   '先檢查必要欄位, 若不過則不可繼續
   '若原為新增模式, 則存檔後繼續新增
   '若原為修改模式, 則存檔後進入顯示模式
    Dim strSQL As String

   ' 檢查資料是否可儲存
    If EditMode = giEditModeView Then Exit Sub
    If IsDataOK() = False Then Exit Sub
    
    On Error Resume Next
    LockWindowUpdate hwnd
    
    gcnGi.BeginTrans
    
    If Err.Number = -2147168237 Then
        Err.Clear
        gcnGi.RollbackTrans
        gcnGi.BeginTrans
    End If
    
    On Error GoTo ChkErr
    
    If Not ScrToRcd Then            ' 把控制項內的值，存到資料庫裡
        On Error Resume Next
        gcnGi.RollbackTrans
        Exit Sub
    End If
    
    gcnGi.CommitTrans
    
    With rsSO002B
            If .State = adStateOpen Then
                 strSQL = .Source
                .Close
                .Open strSQL, gcnGi, adOpenKeyset, adLockOptimistic
                 Set ggrData.Recordset = rsSO002B
            End If
    End With
    
    ggrData.Refresh

    LockWindowUpdate 0
    
'    修改模式存檔後需將該筆異動資料Insert into so002blog, 程序如下 :
'    1. delete from so002blog where AuthenticId=so002b.AuthenticId
'    2. Insert into so002blog (select * from so002b where AuthenticId=[該筆異動資料之AuthenticId變數]

'    cable_liga: 就是阿…so002b有異動的時候，現況會存一筆異動記錄到so002blog ， 現況so002blog會是存「異動前」的資料
'    cable_liga: 如原本姓名是liga , 改為abc, 然後so002blog會存liga
'    cable_liga: 這個正確 應該是存「異動後」的，存為abc才對
    
    gcnGi.Execute "DELETE FROM " & GetOwner & "SO002BLOG WHERE AUTHENTICID='" & txtAuthenticID & "'"
    gcnGi.Execute "INSERT INTO " & GetOwner & "SO002BLOG ( SELECT * FROM " & GetOwner & "SO002B WHERE AUTHENTICID='" & txtAuthenticID & "' )"
    
    On Error GoTo ChkErr
    ' 繼續新增
'    If EditMode = giEditModeInsert Then
'        Call ChangeMode(giEditModeInsert)
'        NewRcd
'       '進入顯示模式
'    Else
'        cmdAdd.SetFocus
'    End If

    Dim blnAdd As Boolean
    blnAdd = EditMode = giEditModeInsert
    
'    If EditMode = giEditModeInsert Then
'        'AddNewGo
'    Else
'        Call ChangeMode(giEditModeView)
'        RcdToScr
'    End If
    
    Call ChangeMode(giEditModeView)
    
    If blnAdd Then
        AddNewSO002A
    End If
  
    RcdToScr
  
  Exit Sub
ChkErr:
    LockWindowUpdate 0
    If Err.Number = -2147168242 Then
        Err.Clear
        Resume Next
    Else
        Call ErrSub(Name, "UpdateGo")
        On Error Resume Next
        gcnGi.RollbackTrans
    End If
End Sub

Private Sub AddNewSO002A()
  On Error GoTo ChkErr
    'Me.Enabled = False
    frmSO1100F.AddNewSO002A Me, txtInvAddress, txtInvTitle, cboInvoiceType.Text, mskInvNo, txtName
    Exit Sub
'    Screen.MousePointer = vbHourglass
'    With frmSO1100F
'        .Show , Me
'        .AddNewGo
'        ' 發票地址、發票抬頭、發票種類、統一編號取由SO002B之資訊。
'        ' SO002B.會員姓名填為SO002A.收件人。
'        If txtInvAddress <> Empty Then .txtInvAddress = txtInvAddress
'        If txtInvTitle <> Empty Then .txtInvTitle = txtInvTitle
'        If cboInvoiceType <> Empty Then .cboInvoiceType = cboInvoiceType
'        If mskInvNo <> Empty Then .mskInvNo = mskInvNo
'        .cboID.ListIndex = 2
'        ' SO002A.ID為 2　，銀行名稱帶入之判斷為
'        .gilBankCode.SetCodeNo GetRsValue("SELECT CODENO FROM " & GetOwner & "CD018" & _
'                                                                    " WHERE PRGNAME IS NULL AND ACTLENGTH=16")
'        .gilBankCode.Query_Description
'        ' (SELECT CODENO,DESCRIPTION FROM CD018 WHERE Prgname is null and Actlength=16)。
'        .txtChargeTitle = txtName ' SO002B.會員姓名填為SO002A.收件人
'        ' 收費地址及郵寄地址取SO002為預設值。
'        .ChangeMode giEditModeInsert
'    End With
'    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
    ErrSub Name, "AddNewSO002A"
End Sub

Private Sub cboEducation_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboJob_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboMarried_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboMemberAge_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboRente_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboSextual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then KeyAscii = 0
End Sub

'Private Sub cboSextual_KeyPress(KeyAscii As Integer)
'    KeyAscii = 0
'End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

'Private Sub cmdFind_Click()
'  On Error GoTo ChkErr
'    Dim strRetVal As String
'    With frmSO114CA1
'            .Show vbModal, Me
'             strRetVal = .ReturnValue
'    End With
'    If strRetVal <> Empty Then txtFaciSNo = strRetVal
'  Exit Sub
'ChkErr:
'    ErrSub Name, "cmdFind_Click"
'End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
    Set ObjActiveForm = Me
    strActFrmName = "frmSO114CA"
    Width = 12015
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error Resume Next
   If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    
    If Crm Then
        PolyFrmFunction Me
        Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Height) / 2)
    Else
        With frmSO1100BMDI.tbrMenu
            If Not800600 Then Move (Screen.Width - Width) / 2, frmSO1100BMDI.Top + .Top + .Height + 666
        End With
    End If
    
    SetLst gilBankCode, "CodeNo", "Description", 3, 24, 405, 2500, "CD018", , True
    SetLst gilCompCode, "CodeNo", "Description", 3, 12, 550, 2000, "CD039"
    NoFilterFaci
5
'    Fill_Cbo_AccountNo
    OpenData
    ChangeMode giEditModeView
    
'    UserPermissionGo
  Exit Sub
ChkErr:
    Call ErrSub(Name, "Form_Load")
End Sub

'Private Sub Fill_Cbo_AccountNo()
'    Dim varRowSet As Variant
'    Dim varRow As Variant
'    On Error Resume Next
'    varRowSet = gcnGi.Execute("SELECT ACCOUNTNO FROM " & GetOwner & _
'                            "SO002A WHERE CUSTID=" & gCustId & _
'                            " AND COMPCODE=" & gCompCode & _
'                            " AND STOPFLAG <> 1").GetRows
'    If Err.Number = 0 Then
'        With cboAccountNo
'                .Clear
'                 For Each varRow In varRowSet
'                    If varRow <> Empty Then .AddItem varRow
'                 Next
''                .AddItem ""
'        End With
'    Else
'        Err.Clear
'    End If
'    varRowSet = vbNullString
'    varRow = vbNullString
'End Sub

Private Sub OpenData()
    
  On Error GoTo ChkErr
    
    Set rsSO002B = New ADODB.Recordset
    
    LockWindowUpdate hwnd
    
    With rsSO002B
            If .State = 1 Then .Close
            .CursorLocation = adUseClient
            .Open "SELECT SO002B.ROWID,SO002B.* FROM " & GetOwner & "SO002B" & _
                        " WHERE CUSTID=" & gCustId & " AND COMPCODE=" & gCompCode, _
                        gcnGi, adOpenKeyset, adLockOptimistic
    End With
    
    Dim mFlds As New GiGridFlds
    With mFlds
            .Add "AUTHENTICID", , , , , "認證編號   ", vbLeftJustify
            .Add "NAME", , , , , "會員姓名  ", vbLeftJustify
            .Add "TEL", , , , , "聯絡電話(宅)", vbLeftJustify
            .Add "MOBILE", , , , , "行動電話  ", vbLeftJustify
            .Add "ID", , , , , "身份證號  ", vbLeftJustify
            .Add "SEXTUAL", , , , , "性別", vbLeftJustify
            .Add "CUSTID", , , , , "認證客戶編號", vbRightJustify
            .Add "FACISNO", , , , , "認證設備序號", vbLeftJustify
            .Add "EMAIL", , , , , "Email Address          ", vbLeftJustify
            .Add "MEMBERID", , , , , "Web會員編號", vbLeftJustify
            .Add "HINTNAME", , , , , "客服網站會員密碼提示問題", vbLeftJustify
'            .Add "PPVORDERPWD", , , , , "PPV訂購密碼", vbLeftJustify
'            .Add "LOGINPASSWORD", , , , , "登入密碼", vbLeftJustify
            .Add "TEL2", , , , , "聯絡電話(公)", vbLeftJustify
            .Add "LOGINID", , , , , "客服網站登入帳號", vbLeftJustify
'            .Add "HINTANSWER", , , , , "客服網站會員密碼提示答案", vbLeftJustify
            .Add "EDUCATION", , , , , "學歷           ", vbLeftJustify
            .Add "JOB", , , , , "職業別             ", vbLeftJustify
            .Add "RENTE", , , , , "個人年收入", vbLeftJustify
            .Add "CLASSID", , , , , "身份別  ", vbLeftJustify
            .Add "STATUS", , , , , "帳號狀態", vbLeftJustify
            .Add "CREATEDATE", giControlTypeDate, , , , "申請日期時間", vbLeftJustify
            .Add "USEDATE", giControlTypeDate, , , , "  啟用日期  ", vbLeftJustify
            .Add "LASTLOGINDATE", giControlTypeDate, , , , "最後登入日期", vbLeftJustify
            .Add "UPDTIME", , , , , "最後異動日期         ", vbLeftJustify
            .Add "ACCOUNTINGBEFORE", , , , , "結算前交易限額", vbRightJustify
            .Add "ACCOUNTINGBALANCE", , , , , "結算後欠費限額", vbRightJustify
    End With
    With ggrData
            .AllFields = mFlds
            Set .Recordset = rsSO002B
            .Refresh
    End With
    
    LockWindowUpdate 0
    On Error Resume Next
    
    If rsSO002B.State > 0 Then
        If rsSO002B.EOF Or rsSO002B.BOF Then
            NewRcd
            txtAccBefore = ""
            txtAccBalance = ""
            cboSextual = ""
        End If
    End If
    
  Exit Sub
ChkErr:
    LockWindowUpdate 0
    ErrSub Name, "OpenData"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift)
  Exit Sub
ChkErr:
       Call ErrSub(Name, "Form_KeyDown")
End Sub
 
Public Sub EditGo()
  On Error GoTo ChkErr
   '進入修改模式, 游標停在第一個可編輯物件
    Call ChangeMode(giEditModeEdit)
    If rsSO002B.Fields("FaciSNo") & "" <> Empty Then
'        csmFaciSno.Clear
'        csmFaciSno.Filter = ""
'        csmFaciSno.QueryString = ""
        csmFaciSno.SetQueryCode rsSO002B.Fields("FaciSNo") & ""
        csmFaciSno.SetQueryCode rsSO002B.Fields("FaciSNo") & ""
    End If
    On Error Resume Next
    txtName.SetFocus
  Exit Sub
ChkErr:
   Call ErrSub(Name, "cmdEdit_Click")
End Sub

Public Sub CancelGo()
   On Error GoTo ChkErr
   '若顯示模式, 則回至前一畫面
   '若修改模式, 則還原資料內容, 進入顯示模式
   '若新增模式, 則清除畫面內容, 讀取第一筆資料, 進入顯示模式
    
   '若顯示模式, 則回至前一畫面
   NoFilterFaci
   If EditMode = giEditModeView Then
        Unload Me
        Exit Sub
   Else
      '若修改模式, 則還原資料內容
      If EditMode = giEditModeEdit Then
         'If Not rsSO002B.EOF Then
         RevertRcd
      Else
        '若新增模式, 則清除畫面內容, 讀取第一筆資料
         If Not rsSO002B.EOF Then
            rsSO002B.MoveFirst
            RcdToScr
         Else
            NewRcd
         End If
      End If
      '進入顯示模式
      Call ChangeMode(giEditModeView)
   End If
 Exit Sub
ChkErr:
   Call ErrSub(Name, "CancelGo")
End Sub

Private Sub RevertRcd()
 On Error GoTo ChkErr
   '還原資料內容
   RcdToScr
 Exit Sub
ChkErr:
   Call ErrSub(Name, "RevertRcd")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ChkErr
        If EditMode <> giEditModeView Then
            If giMsgNotSave = vbNo Then
                Cancel = True
                Exit Sub
            Else
                CancelGo
            End If
        End If
        On Error Resume Next
        If rsSO002B.State = 1 Then rsSO002B.Close
        Set rsSO002B = Nothing
        Call FormQueryUnload
    Exit Sub
ChkErr:
    If Err.Number = 3710 Then
        Resume Next
    Else
        ErrSub Name, "Form_QueryUnload"
    End If
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
'    If LCase(giFld.FieldName) = "sextual" Then
'        If CStr(Value) = "1" Then
'            Value = "男"
'        Else
'            Value = "女"
'        End If
'    End If
  Exit Sub
ChkErr:
    ErrSub Name, "ggrData_ShowCellData"
End Sub

Public Sub PreviousGo()
  On Error GoTo ChkErr
    With rsSO002B
         If .EOF Or .BOF Or .RecordCount <= 0 Then Exit Sub
        .MovePrevious
         If .BOF Then
            .MoveFirst
             MsgBox "已無上一筆資料！", vbInformation, "訊息"
         Else
             RcdToScr
         End If
    End With
  Exit Sub
ChkErr:
    ErrSub Name, "PreviousGo"
End Sub

Public Sub NextGo()
  On Error GoTo ChkErr
    With rsSO002B
         If .EOF Or .BOF Or .RecordCount <= 0 Then Exit Sub
        .MoveNext
         If .EOF Then
            .MoveLast
             MsgBox "已無下一筆資料！", vbInformation, "訊息"
         Else
             RcdToScr
         End If
    End With
  Exit Sub
ChkErr:
    ErrSub Name, "NextGo"
End Sub

Public Sub FirstGo()
  On Error GoTo ChkErr
    With rsSO002B
         If .RecordCount > 0 Then
            .MoveFirst
             RcdToScr
         End If
    End With
  Exit Sub
ChkErr:
    ErrSub Name, "FirstGo"
End Sub

Public Sub LastGo()
  On Error GoTo ChkErr
    With rsSO002B
         If .RecordCount > 0 Then
            .MoveLast
             RcdToScr
         End If
    End With
  Exit Sub
ChkErr:
    ErrSub Name, "LastGo"
End Sub

Private Sub mskInvNo_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If Len(mskInvNo.Text) > 0 Then Cancel = ServiceCommon.InvNo_LostFocus(Me)
  Exit Sub
ChkErr:
    ErrSub Name, "mskInvNo_Validate"
End Sub

Private Sub rsSO002B_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  On Error GoTo ChkErr
    If adReason = 10 Then
        If Not rsSO002B.EOF And Not rsSO002B.BOF And lngEditMode = giEditModeView Then
            If Not ggrData.Enabled Then ggrData.Enabled = True
            RcdToScr
        Else
            ggrData.Enabled = False
        End If
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Name, "rsSO002B_MoveComplete")
End Sub

Private Sub txtAccBalance_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    OnlyNum KeyAscii
End Sub

Private Sub txtAccBefore_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    OnlyNum KeyAscii
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    ChkMaxLen KeyAscii
End Sub

Private Sub txtBodyCount_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    OnlyNum KeyAscii
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
    ChkMaxLen KeyAscii
End Sub

Private Sub txtClassId_KeyPress(KeyAscii As Integer)
    ChkMaxLen KeyAscii
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    ChkMaxLen KeyAscii
End Sub

Private Sub txtEmail_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If Len(txtEmail) > 0 Then
        If InStr(txtEmail, "@") = 0 Then
            MsgBox "輸入的電子郵件地址無效! 請確認!", vbExclamation, "警告"
            Cancel = True
        End If
    End If
  Exit Sub
ChkErr:
    ErrSub Name, "txtEmail_Validate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    ReleaseCOM Me
End Sub

Private Sub txtCustId_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    OnlyNum KeyAscii
End Sub

'Private Sub txtFaciSno_KeyPress(KeyAscii As Integer)
'    ChkMaxLen KeyAscii
'End Sub

Private Sub txtHintAnswer_KeyPress(KeyAscii As Integer)
    ChkMaxLen KeyAscii
End Sub

Private Sub txtHintName_KeyPress(KeyAscii As Integer)
    ChkMaxLen KeyAscii
End Sub

Private Sub txtId_KeyPress(KeyAscii As Integer)
    ChkMaxLen KeyAscii
End Sub

Private Sub txtID_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If Len(txtID) > 0 Then
        If IDIsOk(txtID, True) Then
            If EditMode = giEditModeInsert Then
                On Error Resume Next
                If Mid(txtID, 2, 1) = "1" Then
                    cboSextual = "男"
                Else
                    cboSextual = "女"
                End If
            End If
        End If
    End If
  Exit Sub
ChkErr:
    ErrSub Name, "txtID_Validate"
End Sub

Private Sub txtInvAddress_Change()
  On Error Resume Next
    CML
End Sub

Private Sub txtInvTitle_Change()
  On Error Resume Next
    CML
End Sub

Private Sub txtInvTitle_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    ML KeyAscii
End Sub

Private Sub txtLoginID_KeyPress(KeyAscii As Integer)
    ChkMaxLen KeyAscii
End Sub

Private Sub txtLoginPwd_KeyPress(KeyAscii As Integer)
    ChkMaxLen KeyAscii
End Sub

Private Sub txtMemberID_KeyPress(KeyAscii As Integer)
    ChkMaxLen KeyAscii
End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
    ChkMaxLen KeyAscii
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    ChkMaxLen KeyAscii
End Sub

Private Sub txtPPVOrderPwd_KeyPress(KeyAscii As Integer)
    ChkMaxLen KeyAscii
End Sub

Private Sub txtTel_KeyPress(KeyAscii As Integer)
    ChkMaxLen KeyAscii
End Sub

Private Sub txtTel2_KeyPress(KeyAscii As Integer)
    ChkMaxLen KeyAscii
End Sub

Private Sub txtTown_KeyPress(KeyAscii As Integer)
    ChkMaxLen KeyAscii
End Sub
