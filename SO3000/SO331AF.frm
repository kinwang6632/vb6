VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO331AF 
   BorderStyle     =   1  '單線固定
   Caption         =   "收費資料編輯  [SO331AF]"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   3495
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
   Icon            =   "SO331AF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin prjGiGridR.GiGridR ggrMduRate 
      Height          =   1995
      Left            =   8370
      TabIndex        =   54
      Top             =   510
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3519
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
   Begin prjGiGridR.GiGridR ggrCustRate 
      Height          =   1995
      Left            =   5010
      TabIndex        =   55
      Top             =   510
      Visible         =   0   'False
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   3519
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
   Begin VB.CommandButton cmdShowRate 
      Caption         =   "F7.顯示費率表"
      Height          =   375
      Left            =   7350
      TabIndex        =   27
      Top             =   5370
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "F2. 存檔"
      Height          =   375
      Left            =   270
      TabIndex        =   24
      Top             =   5370
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   8970
      TabIndex        =   26
      Top             =   5370
      Width           =   1215
   End
   Begin VB.Frame fraData 
      Height          =   5325
      Left            =   300
      TabIndex        =   25
      Top             =   30
      Width           =   11475
      Begin VB.CommandButton cmdInvSeqNo 
         Caption         =   "選擇發票抬頭"
         Height          =   315
         Left            =   9120
         TabIndex        =   65
         Top             =   2910
         Width           =   1575
      End
      Begin VB.CommandButton cmdHouse 
         Caption         =   "查詢設備明細"
         Height          =   315
         Left            =   3000
         TabIndex        =   64
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtFaciSNo 
         Height          =   315
         Left            =   1110
         MaxLength       =   20
         TabIndex        =   62
         Top             =   960
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.TextBox txtAuthorizeNo 
         Height          =   315
         Left            =   9120
         MaxLength       =   20
         TabIndex        =   20
         Top             =   3285
         Width           =   2115
      End
      Begin VB.TextBox txtAccountNo 
         Height          =   315
         Left            =   5550
         TabIndex        =   19
         Top             =   3285
         Width           =   2070
      End
      Begin VB.ComboBox cboSTType 
         Height          =   315
         ItemData        =   "SO331AF.frx":0442
         Left            =   3300
         List            =   "SO331AF.frx":044C
         Style           =   2  '單純下拉式
         TabIndex        =   10
         Top             =   2925
         Width           =   1125
      End
      Begin VB.TextBox txtBillNo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1110
         MaxLength       =   15
         TabIndex        =   0
         Top             =   210
         Width           =   1830
      End
      Begin prjNumber.GiNumber gnbQuantity 
         Height          =   315
         Left            =   5550
         TabIndex        =   21
         Top             =   3630
         Visible         =   0   'False
         Width           =   525
         _ExtentX        =   926
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
      End
      Begin prjGiList.GiList gilPTCode 
         Height          =   315
         Left            =   1110
         TabIndex        =   14
         Top             =   4500
         Width           =   3330
         _ExtentX        =   5874
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
         FldWidth2       =   2200
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilClctEn 
         Height          =   315
         Left            =   1110
         TabIndex        =   13
         Top             =   4110
         Width           =   3330
         _ExtentX        =   5874
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
         FldWidth2       =   2200
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilSTCode 
         Height          =   315
         Left            =   1110
         TabIndex        =   12
         Top             =   3720
         Width           =   3330
         _ExtentX        =   5874
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
         FldWidth2       =   2200
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilUCCode 
         Height          =   315
         Left            =   1110
         TabIndex        =   11
         Top             =   3315
         Visible         =   0   'False
         Width           =   3330
         _ExtentX        =   5874
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
         FldWidth2       =   2200
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilCMCode 
         Height          =   315
         Left            =   1110
         TabIndex        =   8
         Top             =   2535
         Width           =   3330
         _ExtentX        =   5874
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
         FldWidth2       =   2200
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilCitemCode 
         Height          =   315
         Left            =   1110
         TabIndex        =   1
         Top             =   600
         Width           =   3330
         _ExtentX        =   5874
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
         FldWidth2       =   2200
         F2Corresponding =   -1  'True
      End
      Begin Gi_Date.GiDate gdaRealStopDate 
         Height          =   315
         Left            =   3300
         TabIndex        =   5
         Top             =   1740
         Width           =   1125
         _ExtentX        =   1984
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
      Begin Gi_Date.GiDate gdaRealDate 
         Height          =   315
         Left            =   1110
         TabIndex        =   9
         Top             =   2925
         Width           =   1125
         _ExtentX        =   1984
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
      Begin Gi_Date.GiDate gdaRealStartDate 
         Height          =   315
         Left            =   1110
         TabIndex        =   4
         Top             =   1755
         Width           =   1125
         _ExtentX        =   1984
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
      Begin VB.TextBox txtManualNo 
         Height          =   315
         Left            =   5550
         MaxLength       =   10
         TabIndex        =   22
         Top             =   3990
         Width           =   1005
      End
      Begin VB.TextBox txtNote 
         Height          =   315
         Left            =   5550
         MaxLength       =   250
         TabIndex        =   23
         Top             =   4335
         Width           =   5685
      End
      Begin prjNumber.GiNumber gnbRealPeriod 
         Height          =   315
         Left            =   3300
         TabIndex        =   3
         Top             =   1350
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
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
         ForeColor       =   0
         MaxLength       =   2
      End
      Begin Gi_Date.GiDate gdaShouldDate 
         Height          =   315
         Left            =   1110
         TabIndex        =   2
         Top             =   1350
         Width           =   1125
         _ExtentX        =   1984
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
      Begin prjNumber.GiNumber gnbShouldAmt 
         Height          =   315
         Left            =   1110
         TabIndex        =   6
         Top             =   2145
         Width           =   1125
         _ExtentX        =   1984
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
         AllowMinus      =   -1  'True
      End
      Begin prjNumber.GiNumber gnbRealAmt 
         Height          =   315
         Left            =   3300
         TabIndex        =   7
         Top             =   2130
         Width           =   1125
         _ExtentX        =   1984
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
         AllowMinus      =   -1  'True
      End
      Begin prjGiList.GiList gilBankCode 
         Height          =   315
         Left            =   5550
         TabIndex        =   17
         Top             =   2940
         Width           =   3135
         _ExtentX        =   5530
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
         FldWidth2       =   2000
         F2Corresponding =   -1  'True
      End
      Begin VB.ComboBox cboAccountNo 
         Height          =   315
         Left            =   5550
         TabIndex        =   18
         Top             =   3285
         Visible         =   0   'False
         Width           =   2325
      End
      Begin prjNumber.GiNumber ginNextAmt 
         Height          =   315
         Left            =   7530
         TabIndex        =   16
         Top             =   2580
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
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
      Begin prjNumber.GiNumber ginNextPeriod 
         Height          =   315
         Left            =   5550
         TabIndex        =   15
         Top             =   2580
         Visible         =   0   'False
         Width           =   465
         _ExtentX        =   820
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
         MaxLength       =   2
         AutoSelect      =   -1  'True
      End
      Begin VB.Label lblFaciSNo 
         AutoSize        =   -1  'True
         Caption         =   "設備序號"
         Height          =   195
         Left            =   210
         TabIndex        =   63
         Top             =   1005
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblNextAmt 
         AutoSize        =   -1  'True
         Caption         =   "下期金額"
         ForeColor       =   &H00C000C0&
         Height          =   195
         Left            =   6600
         TabIndex        =   61
         Top             =   2640
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblNextPeriod 
         AutoSize        =   -1  'True
         Caption         =   "下期期數"
         ForeColor       =   &H00C000C0&
         Height          =   195
         Left            =   4680
         TabIndex        =   60
         Top             =   2640
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblBankCode 
         AutoSize        =   -1  'True
         Caption         =   "銀行名稱"
         Height          =   195
         Left            =   4680
         TabIndex        =   59
         Top             =   3000
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "扣帳帳號"
         Height          =   195
         Left            =   4680
         TabIndex        =   58
         Top             =   3345
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "信用卡授權碼"
         Height          =   195
         Left            =   7890
         TabIndex        =   57
         Top             =   3345
         Width           =   1170
      End
      Begin VB.Label lblSTType 
         AutoSize        =   -1  'True
         Caption         =   "短收類別"
         Height          =   195
         Left            =   2370
         TabIndex        =   56
         Top             =   2985
         Width           =   780
      End
      Begin VB.Label lblQuantity 
         AutoSize        =   -1  'True
         Caption         =   "數量"
         Height          =   195
         Left            =   4680
         TabIndex        =   52
         Top             =   3690
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label lblManualNo 
         AutoSize        =   -1  'True
         Caption         =   "手開單號"
         Height          =   195
         Left            =   4680
         TabIndex        =   51
         Top             =   4050
         Width           =   780
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "備註"
         Height          =   195
         Left            =   4680
         TabIndex        =   50
         Top             =   4395
         Width           =   390
      End
      Begin VB.Label lblCustRate 
         AutoSize        =   -1  'True
         Caption         =   "客戶類別費率表"
         Height          =   195
         Left            =   5010
         TabIndex        =   49
         Top             =   270
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lblMduRate 
         AutoSize        =   -1  'True
         Caption         =   "大樓費率表"
         Height          =   195
         Left            =   8310
         TabIndex        =   48
         Top             =   270
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblCustClass 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   6540
         TabIndex        =   47
         Top             =   270
         Width           =   45
      End
      Begin VB.Label lblMduName 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   9390
         TabIndex        =   46
         Top             =   270
         Width           =   45
      End
      Begin VB.Label lblRealPeriod 
         AutoSize        =   -1  'True
         Caption         =   "期數"
         Height          =   195
         Left            =   2370
         TabIndex        =   45
         Top             =   1380
         Width           =   390
      End
      Begin VB.Label lblRealAmt 
         AutoSize        =   -1  'True
         Caption         =   "實收金額"
         Height          =   195
         Left            =   2370
         TabIndex        =   44
         Top             =   2190
         Width           =   780
      End
      Begin VB.Label lblShouldAmt 
         AutoSize        =   -1  'True
         Caption         =   "出單金額"
         Height          =   195
         Left            =   210
         TabIndex        =   43
         Top             =   2205
         Width           =   780
      End
      Begin VB.Label lblUpdTime 
         Caption         =   " yyy/mm/dd hh:mm:ss"
         Height          =   165
         Left            =   7545
         TabIndex        =   42
         Top             =   4005
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblUpdEn 
         Caption         =   "XXXXXXXX"
         Height          =   195
         Left            =   10275
         TabIndex        =   41
         Top             =   4005
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblUTime 
         Caption         =   "異動時間:"
         Height          =   255
         Left            =   6615
         TabIndex        =   40
         Top             =   4005
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lblUEn 
         Caption         =   "異動人員:"
         Height          =   255
         Left            =   9375
         TabIndex        =   39
         Top             =   4005
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblBillNo 
         AutoSize        =   -1  'True
         Caption         =   "單據編號"
         Height          =   195
         Left            =   210
         TabIndex        =   38
         Top             =   270
         Width           =   780
      End
      Begin VB.Label lblRealDate 
         AutoSize        =   -1  'True
         Caption         =   "入帳日期"
         Height          =   195
         Left            =   210
         TabIndex        =   37
         Top             =   2985
         Width           =   780
      End
      Begin VB.Label lblClctEn 
         AutoSize        =   -1  'True
         Caption         =   "收費人員"
         Height          =   195
         Left            =   210
         TabIndex        =   36
         Top             =   4170
         Width           =   780
      End
      Begin VB.Label lblSTCode 
         AutoSize        =   -1  'True
         Caption         =   "短收原因"
         Height          =   195
         Left            =   210
         TabIndex        =   35
         Top             =   3780
         Width           =   780
      End
      Begin VB.Label lblUCCode 
         AutoSize        =   -1  'True
         Caption         =   "未收原因"
         Height          =   195
         Left            =   210
         TabIndex        =   34
         Top             =   3375
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblCMCode 
         AutoSize        =   -1  'True
         Caption         =   "收費方式"
         Height          =   195
         Left            =   210
         TabIndex        =   33
         Top             =   2595
         Width           =   780
      End
      Begin VB.Label lblRealStopDate 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2370
         TabIndex        =   32
         Top             =   1800
         Width           =   195
      End
      Begin VB.Label lblRealStartDate 
         AutoSize        =   -1  'True
         Caption         =   "有效期限"
         Height          =   195
         Left            =   210
         TabIndex        =   31
         Top             =   1815
         Width           =   780
      End
      Begin VB.Label lblShouldDate 
         AutoSize        =   -1  'True
         Caption         =   "應收日期"
         Height          =   195
         Left            =   210
         TabIndex        =   30
         Top             =   1410
         Width           =   780
      End
      Begin VB.Label lblCitemCode 
         AutoSize        =   -1  'True
         Caption         =   "收費項目"
         Height          =   195
         Left            =   210
         TabIndex        =   29
         Top             =   663
         Width           =   780
      End
      Begin VB.Label lblPTCode 
         AutoSize        =   -1  'True
         Caption         =   "付款種類"
         Height          =   195
         Left            =   210
         TabIndex        =   28
         Top             =   4560
         Width           =   780
      End
   End
   Begin VB.Label lblEditMode 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  '單線固定
      Caption         =   "顯示"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   10530
      TabIndex        =   53
      Top             =   5400
      Width           =   675
   End
End
Attribute VB_Name = "frmSO331AF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2000/01/18
Option Explicit
Private lngEditMode As giEditModeEnu

'起始參數檔
Private Ypara7 As Integer  '有效期限警告日數
Private Ypara10 As Integer '使用費率表
Private Ypara12 As Integer '必需輸入短收原因
Private Ypara14 As Integer '地址依據
Private Ypara5 As Integer '次收日(1999/12/21)
Private Ypara8 As Integer '隨收費資料改變(1999/12/21)
Private Ypara9 As Integer '隨收費資料改變(1999/12/21)
Private YPara37 As Integer
Private YUseManual As Integer

'收費項目，期數，有效起始日，有效截止日、收費日結截止日之異動暫存變數
Private intXCitemCode As Long
Private intXRealPeriod As Integer
Private strXRealStartDate As String
Private strXRealStopDate As String
Private strXLastDate As String
Private intRefNo As Integer  '為收費項目之參考號

'property
Private MstrRealDate As String
Private MstrClctEn As String
Private MstrClctName As String
Private MintPTCode As Integer
Private MstrPTName As String


'宣告一記錄集合供本表單使用
Private rsTmpRec As New ADODB.Recordset

Private rs033 As New ADODB.Recordset

'接收派工單傳來之Connection
Public Conn As New ADODB.Connection

'判斷是否為週期性費用
Private blnPeriodFlag As Boolean

Private lngCustId As Long '由frmSo3311B傳來之客戶編號
Private strCustName As String '由frmso3311B傳來之客戶姓名
Private intBillMode As Integer '由So3311B傳來之單據狀態 1:單據編號
Private strNote As String '由frmso3311B傳來之備註資料
Private strBillNo As String '在新增時由frmso3311b傳來之單據編號
Private strPrtSNo As String '印單序號
Private intItem As Integer '若修改時無法取得RowId即須使用BillNo and Item
Private strCompCode As String

Private strRowId As String '由frmSo3311B傳來之RowId修改時使用
Private strMdbSql As String '記錄在mdb中的So074欄位指令
Private rsMdbRec As New ADODB.Recordset '記錄在修改模式下的RecordSet
Private lngRealPeriod As Integer
Private objParentForm As Object
Private blnLoad As Boolean
Private intDBType As ChangeDB
Private lngOldAmt As Long
Private lngOldPeriod As Long
Private blnFaciSNoChange As Boolean
Dim strInvSeqNo As String
Private Sub cboAccountNo_Click()
    txtAccountNo = GetMaskAccountNo(cboAccountNo)
    
    '#3436 如果帳戶有挑選的話，必需一定要選擇發票 By Kin 2007/12/03
    If cboAccountNo.Text <> "" Then
        cmdInvSeqNo.Enabled = True
    Else
        cmdInvSeqNo.Enabled = False
    End If
End Sub

Private Sub cboSTType_Click()
    On Error Resume Next
    Dim strFilter As String
    Dim strmsg As String
        If gilSTCode.GetCodeNo <> "" And Not blnLoad Then
            If lblSTCode.Caption = "短收原因" And cboSTType.ListIndex = 1 Then
                If MsgBox("您確定將該筆短收改為調改資料?", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then
                    cboSTType.ListIndex = 0
                    Exit Sub
                End If
                gilSTCode.Clear
        ElseIf lblSTCode.Caption = "調改原因" And cboSTType.ListIndex = 0 Then
                If MsgBox("您確定將該筆調改資料改為短收?", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then
                    cboSTType.ListIndex = 1
                    Exit Sub
                End If
                gilSTCode.Clear
            End If
        End If
        strFilter = "Where (ServiceType='" & gServiceType & "' Or ServiceType is null)"
        If cboSTType.ListIndex = 0 Then
            gilSTCode.Filter = strFilter & " And Nvl(RefNo,0) <> 2"
            lblSTCode.Caption = "短收原因"
        Else
            gilSTCode.Filter = strFilter & " And Nvl(RefNo,0) = 2"
            lblSTCode.Caption = "調改原因"
        End If
End Sub

Private Sub cmdCancel_Click()
    If rsTmpRec.State = adStateOpen Then rsTmpRec.Close: Set rsTmpRec = Nothing
    Unload Me
End Sub

Private Sub cmdHouse_Click()
On Error Resume Next
    Dim rsClone As New ADODB.Recordset
    Dim str004FaciSeqNo As String
    Dim blnFaciChange As Boolean
        With frmSO1131E
            .uParentForm = Me
            .uEditMode = lngEditMode
            If txtFaciSNo.Tag <> gCustId Then .uDefFaciSeqNo = txtFaciSNo.Tag
            .uShowClctDate = True
            .uCitemCode = gilCitemCode.GetCodeNo
            .uPRCanClick = True
'            If ZZ <= 1 Then
'                .uInst = True
'                .uInstRsCopy = True
'                .u004Filter = Get004Filter
'                Set .uInstRS = objParentForm.rsSO004
'            End If
            .Show vbModal
            If lngEditMode <> giEditModeView Then
                If .uFaciSeqNo <> "" Then
                    blnFaciChange = txtFaciSNo.Tag <> .uFaciSeqNo And txtFaciSNo.Text = .uFaciSNo
                    txtFaciSNo.Tag = .uFaciSeqNo
                    txtFaciSNo.Text = .uFaciSNo
                    
                    txtFaciSNo.ToolTipText = "設備流水號:" & txtFaciSNo.Tag
                    '2006/08/30 Jacky 修改成當設備序號無值但設備流水號有變動時要做txtFaciSNo Change事件
                    If blnFaciChange Then Call txtFaciSNo_Change
                End If
            End If
        End With
        Call CloseRecordset(rsClone)
End Sub

Private Sub cmdInvSeqNo_Click()
  On Error GoTo chkErr
    Screen.MousePointer = vbHourglass
    With frmSO1131G
        .uAccountNo = txtAccountNo.Text
        .uCompCode = gCompCode
        .uCustId = gCustId
        .uParentForm = Me
        .uInvSeqNo = IIf(strInvSeqNo & "" = "", "", strInvSeqNo)
        .Show 1
    End With
    
    strInvSeqNo = frmSO1131G.uInvSeqNo
    
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdInvSeqNo_Click"

End Sub

Private Sub cmdSave_Click()
On Error GoTo chkErr
    If IsDataOK Then
        If ScrToRcd Then MsgBox "存檔失敗！請重新操作！", vbExclamation, "訊息！"
        cmdCancel.Value = True
    End If
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdSave_Click")
End Sub

Private Sub cmdShowRate_Click()
    On Error GoTo chkErr
        Dim rsCustRate As New ADODB.Recordset
        Dim rsMduRate As New ADODB.Recordset
        Dim MfldsCustRate As New GiGridFlds
        Dim MfldsMduRate As New GiGridFlds
        Dim strCustRateSQL   As String
        Dim strMduRateSQL As String, strMudId As String
        
        ' 設定ggrCustRate 客戶費率檔 顯示收費項，期數，金額
        strCustRateSQL = "Select A.CustName,C.Description,B.Period,B.Amount "
        strCustRateSQL = strCustRateSQL & "From " & GetOwner & "So001 A," & GetOwner & "CD019CD004 B," & GetOwner & "CD019 C Where  A.CustID=" & lngCustId & " and B.ClassCode=A.ClassCode1 and C.CodeNo=B.CitemCode and Nvl(B.CompCode,A.CompCode)=A.CompCode"
        rsCustRate.CursorLocation = adUseClient
        rsCustRate.Open strCustRateSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
        
        If Not rsCustRate.EOF Then
            MfldsCustRate.Add "Description", , , , False, "  收費項目  ", vbLeftJustify
            MfldsCustRate.Add "Period", , , , False, " 期數 ", vbLeftJustify
            MfldsCustRate.Add "Amount", , , , False, "  金額  ", vbLeftJustify
        Else
            If rsCustRate.State = adStateOpen Then rsCustRate.Close
            strCustRateSQL = "Select Description,PeriodFlag,Amount From " & GetOwner & "CD019 Where PeriodFlag =1"
            rsCustRate.Open strCustRateSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
            MfldsCustRate.Add "Description", , , , False, "  收費項目  ", vbLeftJustify
            MfldsCustRate.Add "PeriodFlag", , , , False, " 期數 ", vbLeftJustify
            MfldsCustRate.Add "Amount", , , , False, "  金額  ", vbLeftJustify
        End If
        ggrCustRate.AllFields = MfldsCustRate
        ggrCustRate.SetHead
        lblCustClass.Caption = "(" & GetRsValue("Select CustName From " & GetOwner & "So001 Where Custid=" & lngCustId) & ")"
        Set ggrCustRate.Recordset = rsCustRate
        ggrCustRate.Refresh
        '設定ggrMduRate 大樓費率檔
        MfldsMduRate.Add "Description", , , , False, "  收費項目  ", vbLeftJustify
        MfldsMduRate.Add "Period", , , , False, " 期數 ", vbLeftJustify
        MfldsMduRate.Add "Amount", , , , False, "  金額  ", vbLeftJustify
        ggrMduRate.AllFields = MfldsMduRate
        ggrMduRate.SetHead
        strMudId = GetRsValue("Select Mduid From " & GetOwner & "So001 Where Custid=" & lngCustId) & ""
        If strMudId <> "" Then
            rsMduRate.CursorLocation = adUseClient
            strMduRateSQL = "Select B.Description ,A.Period,A.Amount From " & GetOwner & "CD019SO017 A," & GetOwner & "CD019 B Where A.MduId='" & strMudId & "' and B.CodeNo=A.CitemCode"
            rsMduRate.Open strMduRateSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
            lblMduName = "(" & GetRsValue("Select Name From " & GetOwner & "So017 Where MduId='" & strMudId & "'") & ")"
            Set ggrMduRate.Recordset = rsMduRate
            ggrMduRate.Refresh
        End If
        ggrCustRate.Visible = True
        ggrMduRate.Visible = True
        lblCustRate.Visible = True
        lblMduRate.Visible = True
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdShowrate_Click")
End Sub

Private Sub Form_Activate()
    On Error GoTo chkErr
    '設定初始時的Focus
            Select Case lngEditMode
                Case giEditModeInsert
                    gilCitemCode.SetFocus
                Case giEditModeEdit
                    gnbShouldAmt.SetFocus
                Case giEditModeView
                    cmdCancel.SetFocus
            End Select
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Activeate")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo chkErr
        'F2存檔 F7顯示費率表
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF2 And cmdSave.Enabled Then Call cmdSave_Click: Exit Sub
        If KeyCode = vbKeyF7 And cmdShowRate.Enabled Then Call cmdShowRate_Click: Exit Sub
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
    On Error GoTo chkErr
    Dim strTmpSql As String
        blnLoad = True
        'If Not800600 Then
            'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 + 1200
            Me.Move objParentForm.Left, objParentForm.Top + objParentForm.Height - Me.Height
        'End If
        lblBillNo.Caption = "單據編號"
    '    取得系統參數
    '     Call GetGlobal
    
        strMdbSql = "Select * "
    
        '至最近日結參數記錄檔(So062)取收費日結截止日
        'frmSo3311 在使用
        If lngEditMode <> giEditModeView Then
            strXLastDate = GetRsValue("Select tranDate from " & GetOwner & "So062 Where Type=1 Order By TranDate Desc") & ""
            'rsTmpRec.CursorLocation = adUseClient
            'rsTmpRec.Open strTmpSql, gcnGi, adOpenForwardOnly, adLockReadOnly
            'If Not rsTmpRec.EOF Then strXLastDate = rsTmpRec(0).Value & ""
            'rsTmpRec.Close
        End If
        
        '接收由So3311B傳來之客戶編號和客戶姓名！
        'lngCustID = frmSO3311B.uCustid
        'strCustName = frmSO3311B.uCustName
        'strNote = frmSO3311B.uNote
        
        '至收費參數檔(So043)取各項參數！
        Call GetParameter
    
        '設定GiList控制項
        Call subGil
        If EditMode = giEditModeEdit Then
            txtFaciSNo.Enabled = False
            cmdHouse.Enabled = False
        End If
        '預設為週期性費用，因為收費項目預設值為第一值<1999/12/22>
        blnPeriodFlag = True '<1999/12/22>
    
        '設定異動模式
        Call ChangeMode(lngEditMode)
        If txtAccountNo.Text = "" Then cmdInvSeqNo.Enabled = False
        blnLoad = False
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo chkErr
    '判斷使用者是否在異動的狀態中！
        If cmdCancel.Caption = "取消(&X)" Or lngEditMode = giEditModeView Then
            If rsTmpRec.State = adStateOpen Then rsTmpRec.Close: Set rsTmpRec = Nothing
           ' Set frmSO3311F = Nothing
            Unload Me
            Exit Sub
        End If
            MsgBox "請先存檔或取消後再離開！", vbExclamation, "訊息"
            Cancel = True
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_QueryUnload")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub gdaRealStartDate_Validate(Cancel As Boolean)
On Error GoTo chkErr
Dim lngChk As Long
If Not IsDateOk(gdaRealStartDate.GetValue) Then gdaRealStartDate.SetFocus: Exit Sub
'記錄SFGetAmount傳回之值
Dim strStopDate As String, lngRealAmount As Long, strmsg As String
'記錄RealStartDate ,RealStopDate 之資料
Dim strRealStartDate As String, strRealStopDate As String
strRealStartDate = gdaRealStartDate.GetValue(True) & ""
strRealStopDate = gdaRealStopDate.GetValue(True) & ""

    If strRealStartDate = "" Or strRealStartDate > strRealStopDate Then MsgBox "此欄位需有值且<=截止日期": gdaRealStartDate.SetFocus: Exit Sub
    If strXRealStartDate <> strRealStartDate Then
    '#3243 多增加傳入設備序號參數
    '#3465 會出現錯誤,是因為傳入設備序號沒有取到值,將設備序號參數取消 By Kin 2007/08/22
        lngChk = SFGetAmount(False, 2, lngCustId, gilCitemCode.GetCodeNo, gnbRealPeriod.Text, gdaRealStartDate.GetValue, gServiceType, Val(gCompCode), strStopDate, lngRealAmount, strmsg, lngRealPeriod)
        
        If lngChk < 0 Then
            MsgBox strmsg, vbExclamation, "訊息"
            gdaRealStartDate.SetFocus
        Else
            gnbShouldAmt.Text = lngRealAmount
            Call gdaRealStopDate.SetValue(strStopDate)
            '設定各個變數以供下列程式使用
            strXRealStartDate = gdaRealStartDate.GetValue(True)
            strXRealStopDate = gdaRealStopDate.GetValue(True)
        End If
    End If
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "gdaRealStartDate_VAlidate")
End Sub

Private Sub gdaRealStopDate_Validate(Cancel As Boolean)
On Error GoTo chkErr
Dim lngChk As Long, strVBYesNo As String
If Not IsDateOk(gdaRealStopDate.GetValue) Then gdaRealStopDate.SetFocus: Exit Sub
'記錄RealStartDate ,RealStopDate 之資料
Dim strRealStartDate As String, strRealStopDate As String
strRealStartDate = gdaRealStartDate.GetValue(True) & "" 'StrToDate YYYY/MM/DD
strRealStopDate = gdaRealStopDate.GetValue(True) & ""
'記錄SFGetAmount傳回之值
Dim strStopDate As String, lngRealAmount As Long, strmsg As String

    If strRealStopDate = "" Or strRealStartDate > strRealStopDate Then MsgBox "此欄位需有值且<=截止日期": gdaRealStopDate.SetFocus: Exit Sub
    If strXRealStopDate <> strRealStopDate Then
        '若本欄日期減去有效起始日期大於intYPara7，則顯示確認視窗
        If DateDiff("d", CDate(strRealStartDate), CDate(strRealStopDate)) > Ypara7 Then
                If vbNo = MsgBox("有效期限超過系統預設天數，是否接受此日期！", vbExclamation, "訊息") Then gdaRealStopDate.SetFocus: Exit Sub
        End If
        '#3243 會導致應收金額計算錯誤,現在將此計算拿掉,改為與SO1132C一樣 By Kin 2007/05/23
        'lngChk = SFGetAmount(True, 2, lngCustId, gilCitemCode.GetCodeNo, gnbRealPeriod.Text, gdaRealStartDate.GetValue, gServiceType, Val(gCompCode), Format(strRealStopDate, "yyyymmdd"), lngRealAmount, strMsg, lngRealPeriod)
        If lngChk < 0 Then
            MsgBox strmsg, vbExclamation, "訊息"
            gdaRealStopDate.SetFocus
        Else
            If GetNumberType(lngRealAmount & "") <> 0 Then
                 gnbShouldAmt.Text = Val(lngRealAmount & "")
                '設定各個變數以供下列程式使用
                strXRealStartDate = gdaRealStartDate.GetValue(True) & ""
                strXRealStopDate = gdaRealStopDate.GetValue(True) & ""
            End If
        End If
    End If
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "gdaRealStopDate_VAlidate")
End Sub

Private Sub gdaShouldDate_Validate(Cancel As Boolean)
    If gdaShouldDate.Text = "" Then Exit Sub
    If Not IsDateOk(gdaShouldDate.GetValue) Then MsgBox "日期不合法！": gdaShouldDate.SetFocus: Exit Sub
End Sub

Private Sub ggrCustRate_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
On Error GoTo chkErr
    If Fld.Name = "AMOUNT" Then
        Value = Format(Value, "###,###,###")
    End If
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "ggrCustRate_ShowCellDate")
End Sub

Private Sub ggrMduRate_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
On Error GoTo chkErr
    If Fld.Name = "AMOUNT" Then
        Value = Format(Value, "###,###,###")
    End If
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "ggrMduRate_ShowCellDate")
End Sub

Private Sub gilCitemCode_Change()
  On Error Resume Next
    Call EnableFaciSNo
End Sub

Private Sub gilCitemCode_Validate(Cancel As Boolean)
On Error GoTo chkErr
    Dim strTmpSql As String, intAmount As Long, intPeriod As Integer
    Dim blnPeriodFlagMode As Boolean
    Dim rsTmpCitemCode As New ADODB.Recordset
    '記錄SFGetAmount 傳回的相關資料
    Dim strStopDate As String, lngRealAmount As Long, strmsg As String
    Dim lngChk As Long
        If blnLoad And lngEditMode <> giEditModeInsert Then Exit Sub
        If gilCitemCode.GetCodeNo = "" Or gilCitemCode.GetDescription = "" Then MsgBox gMsgIsDataOK, vbExclamation, "訊息": gilCitemCode.SetFocus: Exit Sub
        '若有值且有改變
        If intXCitemCode <> gilCitemCode.GetCodeNo Then
            '至收費項目代碼檔取（是否週期性費用），（預設金額），(預設期數）<<1999/12/22>>，（參考號）<<1999/12/22>>
            strTmpSql = "Select PeriodFlag,Amount,Period,RefNo From " & GetOwner & "CD019 Where CodeNo=" & gilCitemCode.GetCodeNo
            
            '開啟連線！
            If OpenRecordset(rsTmpCitemCode, strTmpSql, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseServer, False, False) <> giOK Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！"
            
            blnPeriodFlag = Val(rsTmpCitemCode("PeriodFlag").Value & "") = 1
            
            intAmount = GetNumberType(rsTmpCitemCode("Amount").Value & "")
            intPeriod = Val(rsTmpCitemCode("Period").Value & "")      '<<1999/12/22>>預設期數
            intRefNo = Val(rsTmpCitemCode("RefNO").Value & "")        '<<1999/12/22>>參考號
            rsTmpCitemCode.Close
            Set rsTmpCitemCode = Nothing
            '若（是否週期性費費用）=否 則：
            '▲期數=0，有效起始日期/截止日期=無，出單金額=（預設金額）
            '▲Disable 期數，有效起日期/截止日期
            If blnPeriodFlag = False Then
                blnPeriodFlagMode = False
                gnbRealPeriod.Text = ""
                Call gdaRealStopDate.SetValue("")
                Call gdaRealStartDate.SetValue("")
                gnbShouldAmt.Text = intAmount
            Else
                blnPeriodFlagMode = True
                gnbRealPeriod.Text = intPeriod '<<1999/12/22>> Old gnbRealPeriod.Text =1
                
                Call gdaShouldDate.SetValue(ReturnRealStartDate(gilCitemCode.GetCodeNo, lngCustId, , Me)) '<<1999/12/21>>
                gdaRealStartDate.SetValue gdaShouldDate.GetValue
                ''#3243 多增加傳入設備序號參數
                '#3465 會出現錯誤,是因為傳入設備序號沒有取到值,將設備序號參數取消 By Kin 2007/08/22
                lngChk = SFGetAmount(False, 2, lngCustId, gilCitemCode.GetCodeNo, gnbRealPeriod.Text, gdaRealStartDate.GetValue, gServiceType, Val(gCompCode), strStopDate, lngRealAmount, strmsg, lngRealPeriod)
                gnbRealPeriod.Text = lngRealPeriod '<<1999/12/22>> Old gnbRealPeriod.Text =1
                If lngChk < 0 Then
                    MsgBox strmsg, vbExclamation, "訊息"
                    On Error Resume Next
                    gilCitemCode.SetFocus
                    On Error GoTo chkErr
                Else
                    gnbShouldAmt.Text = lngRealAmount
                    Call gdaRealStopDate.SetValue(strStopDate & "")
                End If
            End If
            gnbRealAmt.Text = gnbShouldAmt.Value
            '#5661 如果Para37 要能修改 By Kin 2010/05/27
            gnbRealPeriod.Enabled = blnPeriodFlagMode Or (YPara37 = 1)
            gdaRealStartDate.Enabled = blnPeriodFlagMode Or (YPara37 = 1)
            gdaRealStopDate.Enabled = blnPeriodFlagMode Or (YPara37 = 1)
            intXCitemCode = Val(gilCitemCode.GetCodeNo & "")
            intXRealPeriod = Val(gnbRealPeriod.Text & "")
            strXRealStartDate = gdaRealStartDate.GetValue(True) & ""
            strXRealStopDate = gdaRealStopDate.GetValue(True) & ""
            If lngEditMode = giEditModeInsert Then gilUCCode_Change
         End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "gilCitemCode_Validate")
End Sub

Private Sub gilUCCode_Change()
    On Error GoTo chkErr
        '2.逐筆確認 95/04/07 Jacky 2283 KC
        If blnLoad Then Exit Sub
        EnableNextPeriod
    Exit Sub
chkErr:
    ErrSub Me.Name, "gilUCCode_Change"
End Sub

Private Sub EnableNextPeriod()
    On Error Resume Next
    Dim blnFlag As Boolean
    Dim strFaciSeqNo As String
    
        If Ypara8 = 2 Then
            blnFlag = True
            
            ginNextPeriod.Text = 0
            ginNextAmt.Text = 0
            '客戶指定不顯示
            If lngEditMode = giEditModeEdit Then strFaciSeqNo = rs033("FaciSeqNo") & ""
            If Val(GetSO003Value("CustAllot", Val(gCustId), Val(gCompCode), gServiceType, strFaciSeqNo, gilCitemCode.GetCodeNo)) = 1 Then blnFlag = False
            If blnFlag Then
                ginNextPeriod.Text = gnbRealPeriod.Value
                If Ypara9 = 1 Then
                    If lngEditMode <> giEditModeInsert Then
                        ginNextPeriod.Text = rs033("OldPeriod")
                        ginNextAmt.Text = rs033("OldAmt")
                    Else
                        ginNextAmt.Text = gnbShouldAmt.Value
                    End If
                ElseIf Ypara9 = 2 Then
                     ginNextAmt.Text = gnbShouldAmt.Value
                Else
                     ginNextAmt.Text = gnbRealAmt.Value
                End If
                
            End If
            
            lblNextPeriod.Visible = blnFlag
            ginNextPeriod.Visible = blnFlag
            lblNextAmt.Visible = blnFlag
            ginNextAmt.Visible = blnFlag
        End If
End Sub

Private Sub gnbQuantity_GotFocus()
    With gnbQuantity
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub gnbRealAmt_GotFocus()
    With gnbRealAmt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub gnbRealAmt_Validate(Cancel As Boolean)
On Error GoTo chkErr
Dim intShouldAmt As Long
Dim intRealAmt As Long
intShouldAmt = gnbShouldAmt.Value
intRealAmt = gnbRealAmt.Value
'若有值，則清除未收原因
'若有值且出單金額有值：
'   '、若二者不相等，則警告："實收與出單金額不相等",並填入短收原因初值1及對應名稱
    '、若二者相等，則清除短收原因
    If intRealAmt > 0 Then
        gilUCCode.Clear
    Else
        gilSTCode.Clear
        gilUCCode.ListIndex = 1
    End If

    If intShouldAmt > 0 And intRealAmt > 0 Then
        If intShouldAmt <> intRealAmt And gilSTCode.GetCodeNo = "" Then
            MsgBox "實收與出單金額不相等！", vbExclamation, "訊息"
            If intRealAmt > intShouldAmt Then gnbRealAmt.Text = 0
            gilSTCode.ListIndex = 1
            gilSTCode.SetFocus
        ElseIf intShouldAmt = intRealAmt Then
            gilSTCode.Clear
        End If
    End If
    EnableNextPeriod
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "gnbRealAmt_VailDate")
End Sub

Private Sub gnbRealPeriod_GotFocus()
    With gnbRealPeriod
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub gnbRealPeriod_Validate(Cancel As Boolean)
On Error GoTo chkErr
Dim lngChk As Long, strStopDate As String
Dim lngRealAmount As Long, strmsg As String
    '若期數在Enable的狀態下須有值
If Not gnbRealPeriod.Enabled Then Exit Sub

    '若無值或<=0則錯誤:"此欄位必需為正值！"且離開
    If gnbRealPeriod.Text <= 0 Or str(gnbRealPeriod.Text) = "" Then MsgBox "此欄位必需有值且需為正值", vbExclamation, "訊息": Exit Sub
    
    If intXRealPeriod <> gnbRealPeriod.Text Then
    lngChk = SFGetAmount(False, 2, lngCustId, gilCitemCode.GetCodeNo, gnbRealPeriod.Text, gdaRealStartDate.GetValue, gServiceType, Val(gCompCode), strStopDate, lngRealAmount, strmsg, lngRealPeriod)
        If lngChk < 0 Then
            MsgBox strmsg, vbExclamation, "訊息"
            gnbRealPeriod.SetFocus
        Else
            If lngRealAmount = 0 Then lngChk = SFGetAmount(True, 2, lngCustId, gilCitemCode.GetCodeNo, gnbRealPeriod.Text, gdaRealStartDate.GetValue, gServiceType, Val(gCompCode), gdaRealStopDate.GetValue, lngRealAmount, strmsg, lngRealPeriod)
            If lngChk < 0 Then
                MsgBox strmsg, vbExclamation, "訊息"
                gnbRealPeriod.SetFocus
                Exit Sub
            End If
            gnbShouldAmt.Text = lngRealAmount
            Call gdaRealStopDate.SetValue(strStopDate)
            '設定各個變數以供下列程式使用
            intXRealPeriod = gnbRealPeriod.Text
            strXRealStartDate = gdaRealStartDate.GetValue
            strXRealStopDate = gdaRealStopDate.GetValue
        End If
    End If
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "gnbRealPeriod_Validate")
End Sub

Private Sub gnbShouldAmt_Change()
    gnbRealAmt.Text = gnbShouldAmt.Value
End Sub

Private Sub gnbShouldAmt_GotFocus()
    With gnbShouldAmt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub gnbShouldAmt_Validate(Cancel As Boolean)
On Error GoTo chkErr
Dim intShouldAmt As Long
Dim intRealAmt As Long
intShouldAmt = gnbShouldAmt.Value
intRealAmt = gnbRealAmt.Value
    '出單金額若有值且實收金額有值：
        '若二者不相等，則警告
        '若二者相等，則清除短收原因<1999/12/27>
        If intRealAmt > 0 And intShouldAmt > 0 Then
        If intShouldAmt <> intRealAmt And gilSTCode.GetCodeNo = "" Then
                MsgBox "實收與出單金額不相等"
                If intRealAmt > intShouldAmt Then
                        gnbRealAmt.Text = 0
                ElseIf intRealAmt = 0 Then
                        gilSTCode.Clear
                        gilUCCode.ListIndex = 1
                    Else
                        gilSTCode.ListIndex = 1
                    End If
            ElseIf intRealAmt = intShouldAmt Then
                gilSTCode.Clear
            End If
        End If
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "gnbShouldAmt_Validate")
End Sub

Private Sub txtAccountNo_Change()
  On Error Resume Next
    If txtAccountNo.Text <> "" Then
        cmdInvSeqNo.Enabled = True
    Else
        cmdInvSeqNo.Enabled = False
    End If

End Sub

Private Sub txtFaciSNo_Change()
    '讓收費項目重新取得新的起始截止日 93/11/18
    If txtFaciSNo <> "" Or txtFaciSNo.Tag <> "" Then
        blnFaciSNoChange = True
        intXCitemCode = 0
        gilCitemCode_Change
        blnFaciSNoChange = False
    End If

End Sub

Private Sub txtManualNo_GotFocus()
    With txtManualNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtNote_GotFocus()
    With txtNote
        .IMEMode = 1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub ChangeMode(ByVal lngMode As giEditModeEnu)
On Error GoTo chkErr
    Dim blnFlag As Boolean  ' 記錄是否在資料瀏覽模式，
    Dim blnFlag2 As Boolean
    lngEditMode = lngMode
    
    Select Case lngEditMode
      Case giEditModeInsert
           blnFlag = True
           blnFlag2 = True
           lblEditMode.Caption = "新增"
           cmdCancel.Caption = "取消(&X)"
           Call NewRec
      Case giEditModeEdit
           blnFlag = True
           blnFlag2 = False
           lblEditMode.Caption = "修改"
           cmdCancel.Caption = "取消(&X)"
           Call RcdToScr
      Case giEditModeView
           blnFlag = False
           lblEditMode.Caption = "顯示"
           cmdCancel.Caption = "結束 (&X)"
           Call RcdToScr
    End Select
    '設定各個物件的Enabled/Disabled
    gilCitemCode.Enabled = blnFlag2  '收費項目
    gdaShouldDate.Enabled = blnFlag2 '應收日期
    fraData.Enabled = blnFlag
    cmdSave.Enabled = blnFlag
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "ChangeMode")
End Sub

' 自訂屬性：EditMode
' 記錄目前在編輯、新增或檢視模式
' giEditModeEnu(自訂列舉值，設定於Sys_Lib)
Public Property Get EditMode() As giEditModeEnu
  On Error GoTo chkErr
    ' 取目前編輯模式
    EditMode = lngEditMode
  Exit Property
chkErr:
   Call ErrSub(Me.Name, "Get EditMode")
End Property

Public Property Let EditMode(ByVal vNewValue As giEditModeEnu)
  On Error GoTo chkErr
    '設定編輯模式
    lngEditMode = vNewValue
  Exit Property
chkErr:
    Call ErrSub(Me.Name, "Let EditMode")
End Property

Private Sub NewRec()
On Error GoTo chkErr
    '印單序號
    txtBillNo.Text = strBillNo
'    gilCitemCode
    '收費項目
    Call gilCMCode.SetCodeNo("")
    Call gilCMCode.SetDescription("")
    '入帳日期
    Call gdaRealDate.SetValue(IIf(MstrRealDate <> "", MstrRealDate, ""))
    '收費方式
     gilCitemCode.ListIndex = 1
    'Call gilCitemCode.SetDescription(IIf(frmSO3311B.uCMName <> "", frmSO3311B.uCMName, ""))
    
    '#XXXX 判斷是否有第二層優惠 By Kin 2009/04/01
    If GetSystemParaItem("SecondDiscount", gCompCode, "", "SO041") = 1 Then
        gilCitemCode.Filter = gilCitemCode.Filter & IIf(gilCitemCode.Filter = "", " Where ", " And ") & "Nvl(SecendDiscount,0) = 0 "
    End If

    
    '收費人員
    Call gilClctEn.SetCodeNo(MstrClctEn)
    Call gilClctEn.SetDescription(MstrClctName)
    '付款種類
    Call gilPTCode.SetCodeNo(IIf(MintPTCode > 0, MintPTCode, ""))
    Call gilPTCode.SetDescription(MstrPTName)
    '異動時間
    lblUpdTime.Caption = GetDTString(RightNow)
    '異動人員
    lblUpdEn.Caption = garyGi(1)
    '未收原因
    gilUCCode.Clear
    '短收原因
    gilSTCode.Clear
    '應收日期
    gdaShouldDate.SetValue (MstrRealDate)
    '有效起始日
    gdaRealStartDate.SetValue ("")
    '有效截止日
    gdaRealStopDate.SetValue ("")
    '期數
    gnbRealPeriod.Text = ""
    '實收金額
    gnbRealAmt.Text = ""
    '應收金額
    gnbShouldAmt.Text = ""
    '數量
    gnbQuantity.Text = ""
    '初始值
    cboSTType.ListIndex = 0
    gilBankCode.Clear
    txtAuthorizeNo = ""
    txtAccountNo = ""
    intXCitemCode = 0
    intXRealPeriod = 0
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "NewRec")
End Sub

Private Sub GetParameter()
    On Error GoTo chkErr
        Dim rsTmp As New ADODB.Recordset
        If Val(gCompCode) = 0 Then gCompCode = 1
        gCustId = lngCustId
        '#5661 增加Para37判斷期數是否能修改 By Kin 2010/05/27
        If Not GetSystemPara(rsTmp, gCompCode, gServiceType, "SO043", "Para7,Para10,Para12,Para14,Para5,Para8,Para9,UseManual,Para37") Then Exit Sub

        Ypara5 = rsTmp("Para5").Value 'Update(199/12/21)
        Ypara8 = rsTmp("Para8").Value '(199/12/21)
        Ypara9 = rsTmp("Para9").Value '(199/12/21)
        Ypara7 = rsTmp("Para7").Value
        Ypara10 = rsTmp("Para10").Value
        Ypara12 = rsTmp("Para12").Value
        Ypara14 = rsTmp("Para14").Value
        YPara37 = Val(rsTmp("Para37").Value & "")
        YUseManual = rsTmp("UseManual").Value
        Call Close3Recordset(rsTmp)
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "GetParameter")
End Sub

Private Sub RcdToScr()
On Error GoTo chkErr
Dim strSelSql As String
Dim strTable As String
    If intDBType = giOracle Then
        strTable = "SO074A"
    Else
        strTable = "SO3311"
        If rsMdbRec.State = adStateOpen Then rsMdbRec.Close
    End If
    If strRowId <> "" Then
        strSelSql = strMdbSql & " From " & strTable & " Where " & IIf(intDBType = giAccessDb, "Rowid", "RcdRowId") & "='" & strRowId & "'"
    Else
        strSelSql = strMdbSql & " From " & strTable & " Where Billno='" & strBillNo & "' And Item = " & intItem
    End If
    If OpenRecordset(rsMdbRec, strSelSql, Conn, adOpenKeyset, adLockOptimistic, adUseClient, False) <> giOK Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": Unload Me
        txtBillNo.Text = rsMdbRec("BillNo").Value                               '單據編號
        If Not GetRS(rs033, "Select * From " & GetOwner & "SO033 Where BillNO = '" & rsMdbRec("BillNo") & "' And Item= " & rsMdbRec("Item")) Then Exit Sub
    Call gilCitemCode.SetCodeNo(rsMdbRec("CitemCode").Value & "")           '收費項目代號
    Call gilCitemCode.SetDescription(rsMdbRec("CitemName").Value & "")      '收費項目名稱
    Call gdaShouldDate.SetValue(Format(rsMdbRec("ShouldDate").Value & "", "YYYYMMDD"))  '應收日期
    
     gnbShouldAmt.Text = IIf(GetNumberType(rsMdbRec("ShouldAmt").Value & "") = 0, "", GetNumberType(rsMdbRec("ShouldAmt").Value & ""))              '出單金額
     gnbRealAmt.Text = GetNumberType(rsMdbRec("RealAmt").Value & "")     '實金金額
     '-----------------------------------------------------------------------------------------
    '若為修改模式：若期數 >0 ，則Enable 期數,有效起始／截止日,反之則disable
    If rsMdbRec("RealPeriod").Value > 0 Then
        blnPeriodFlag = True
    Else
        blnPeriodFlag = False
    End If
    gnbRealPeriod.Enabled = blnPeriodFlag
    gdaRealStartDate.Enabled = blnPeriodFlag
    gdaRealStopDate.Enabled = blnPeriodFlag
    'If rsMdbRec("RealPeriod").Value > 0 Then
            gnbRealPeriod.Text = IIf(GetNumberType(rsMdbRec("RealPeriod").Value) = 0, "", GetNumberType(rsMdbRec("RealPeriod").Value))        '期數
            Call gdaRealStartDate.SetValue(Format(rsMdbRec("RealStartDate") & "", "YYYYMMDD"))      '有效起始日
            Call gdaRealStopDate.SetValue(Format(rsMdbRec("RealStopDate") & "", "YYYYMMDD"))     '有效截止日
            '設定暫存變數內容
            intXCitemCode = GetNumberType(rsMdbRec("CitemCode").Value & "")
            intXRealPeriod = rsMdbRec("RealPeriod").Value
            strXRealStartDate = Format(rsMdbRec("RealStartDate").Value & "", "YYYYMMDD")
    'End If
     
'-----------------------------------------------------------------------------------------
     Call gilCMCode.SetCodeNo(rsMdbRec("CMCode").Value & "")              '收費代碼
     Call gilCMCode.SetDescription(rsMdbRec("CMName").Value & "")                   '名稱
     Call gdaRealDate.SetValue(Format((rsMdbRec("RealDate").Value & ""), "YYYYMMDD"))         '實收日期
     'Call gilUCCode.SetCodeNo(rsMdbRec("UCCode").Value & "")                        '未收原因代碼
     'Call gilUCCode.SetDescription(rsMdbRec("UCName").Value & "")                   '未收原因名稱
     Call gilSTCode.SetCodeNo(rsMdbRec("STCode").Value & "")                        '短收原因代碼
     Call gilSTCode.SetDescription(rsMdbRec("STName").Value & "")                   '短收原因名稱
     Call gilClctEn.SetCodeNo(rsMdbRec("ClctEn").Value & "")                        '收費人員代碼
     Call gilClctEn.SetDescription(rsMdbRec("ClctName").Value & "")                 '收費人員名稱
     Call gilPTCode.SetCodeNo(rsMdbRec("PTCode").Value & "")                        '付款種類代碼
     Call gilPTCode.SetDescription(rsMdbRec("PTName").Value & "")                   '付款種類名稱
     '*************************************************************************************
     '#3465 決定設定流水號是否出現讓User選擇 By Kin 2007/08/28
     Call EnableFaciSNo
     '*************************************************************************************
     'gilCitemCode.Filter = gilCitemCode.Filter & IIf(gilCitemCode.Filter = "", "Where ", " And ") & " ByHouse = " & GetByHouse(gilCitemCode.GetCodeNo)
     txtManualNo.Text = rsMdbRec("ManualNo") & ""                        '手開單號
     txtNote.Text = rsMdbRec("Note").Value & ""                                      '備註
     
     gilBankCode.SetCodeNo rsMdbRec("BankCode") & ""
     gilBankCode.SetDescription rsMdbRec("BankName") & ""
     txtAccountNo.Text = rsMdbRec("AccountNo") & ""
     txtAuthorizeNo.Text = rsMdbRec("AuthorizeNo") & ""
     cboSTType.ListIndex = Val(rsMdbRec("AdjustFlag") & "")
     If gilSTCode.GetCodeNo = "" And lngEditMode = giEditModeEdit Then cboSTType.ListIndex = GetDefaultSTType(rsMdbRec("ServiceType") & "")
     
    ginNextPeriod.Text = Val(rsMdbRec("NextPeriod") & "")
    ginNextAmt.Text = Val(rsMdbRec("NextAmt") & "")
    '*********************************************************************************************************************************
    '#3465 取得物品序號與設備流水號 By Kin 2007/08/28
    txtFaciSNo.Text = GetRsValue("Select FaciSNo From " & GetOwner & "SO004 Where SeqNo='" & rsMdbRec("FaciSeqNo") & "'", gcnGi) & ""
    txtFaciSNo.Tag = rsMdbRec("FaciSeqNo") & ""
    '**********************************************************************************************************************************
    
    '******************************************************************
    '#3436 如果有使用帳號，記錄所使用的發票流水號 By Kin 2007/12/17
    strInvSeqNo = rsMdbRec("InvSeqNo") & ""
    '******************************************************************
    
     'strSign = GetRsValue("Select Sign From CD019 Where CodeNo = " & gilCitemCode.GetCodeNo) & ""
    intXCitemCode = Val(gilCitemCode.GetCodeNo & "")
    intXRealPeriod = Val(gnbRealPeriod.Text & "")
    strXRealStartDate = gdaRealStartDate.GetValue(True) & ""
    strXRealStopDate = gdaRealStopDate.GetValue(True) & ""
     
     '判斷如果是顯示模式就將rsTmpRec結束，否則即為編輯模式須再提供存檔時使用
        EnableMode
        EnableNextPeriod
Exit Sub
chkErr:
     Call ErrSub(Me.Name, "RcdToScr")
End Sub
Private Sub EnableMode()
    On Error GoTo chkErr
    Dim intCount As Integer
    Dim strSQL As String
    Dim blnTFNBillNo As Boolean
    
        
        strSQL = "Select Count(*) From " & GetOwner & "SO003 Where CustId = " & gCustId & _
            " And CompCode = " & gCompCode & " And ServiceType = '" & gServiceType & "' And CitemCode = " & gilCitemCode.GetCodeNo & _
            " And ContNo is not null"
        '#3465 測試不OK 程式多判斷rs033筆數是否為0筆 By Kin 2007/08/30
        If (GetByHouse(gilCitemCode.GetCodeNo) = 0) And (Not rs033.EOF) Then
            strSQL = strSQL & " And FaciSeqNo = '" & rs033("FaciSeqNo") & "'"
        End If
        
        intCount = GetRsValue(strSQL)
        If lngEditMode = giEditModeEdit Then
            '#3465測試 不OK多判斷rs033是否為空值，不然會有錯誤訊息 By Kin 2007/08/28
            If Not rs033.EOF Then
                blnTFNBillNo = rs033("TFNBillNo") & "" <> ""
            End If
        End If
        '#5661 如果Para37=1要能修改 By Kin 2010/05/27
        If Not rs033.EOF Then
            If intCount > 0 Or rs033("ContNo") & "" <> "" Or blnTFNBillNo Then
                If Mid(txtBillNo, 7, 1) <> "T" Then gilCitemCode.Enabled = False
                gnbRealPeriod.Enabled = False Or (YPara37 = 1)
                gdaRealStartDate.Enabled = False Or (YPara37 = 1)
                gdaRealStopDate.Enabled = False Or (YPara37 = 1)
                gnbShouldAmt.Enabled = False Or (YPara37 = 1)
                gdaShouldDate.Enabled = False Or (YPara37 = 1)
            End If
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "EnableMode"
End Sub

'新增時取得帳單編號
Private Function GetBillNo(ByVal strDate As Date, ByVal strBillNo As String) As String
On Error Resume Next
Dim str_BillNo As String
    str_BillNo = Year(strDate) & String(2 - Len(Month(strDate)), "0") & Month(strDate)
    str_BillNo = str_BillNo & "T" & String(8 - Len(strBillNo), "0") & strBillNo
    GetBillNo = str_BillNo
End Function

Private Function IsDataOK() As Boolean
    On Error GoTo chkErr
    Dim strCPGroupCode As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngSubAmt As Long
    Dim lngTotalAdjAmt As Long
        IsDataOK = False
        '出單金額不得小於等於0，或空值！
        
        '      If gnbShouldAmt.Value <= 0 Or gnbShouldAmt.Text = "" Then MsgBox "出單金額不得小於等於零！", vbExclamation, "訊息！": Exit Function
         If Not ChkDTok Then Exit Function
        
        '必要欄位：單據編號、項目，收費項目代碼，應收日期，數量，收費方式代碼
        If Not MustExist(gdaShouldDate, 1, "應收日期") Then Exit Function
        If Not MustExist(gilCitemCode, 2, "收費項目") Then Exit Function
        If Not MustExist(gilCMCode, 2, "收費方式") Then Exit Function
        lngSubAmt = gnbShouldAmt.Value - gnbRealAmt.Value
        
        '若期數>0，則有效起始日與截止日需有值，且起始日<=截止日
        If GetNumberType(gnbRealPeriod.Text & "") > 0 Then
            If Not MustExist(gdaRealStartDate, 1, "有效起始日") Then Exit Function
            If Not MustExist(gdaRealStopDate, 1, "有效截止日") Then Exit Function
           
            If (CDate(gdaRealStartDate.GetValue(True)) > CDate(gdaRealStopDate.GetValue(True))) Then
                Call MsgDateRangeX("有效期限")
                Exit Function
             End If
        End If
        
        '若出單金額 <>0，且實收金額 <>0，且出單金額!=實收金額，則短收原因需有值 2000/03/10
        If Val(gnbShouldAmt.Value & "") > 0 And Val(gnbRealAmt.Value & "") > 0 And (Val(gnbRealAmt.Value & "") <> Val(gnbShouldAmt.Value & "")) Then
            On Error Resume Next
            gnbRealAmt.SetFocus
            On Error GoTo chkErr
            If gilSTCode.GetCodeNo = "" Then
                gilSTCode.SetFocus
                MsgBox "出單金額不等於實收金額需有" & lblSTCode.Caption & "！", vbExclamation, "訊息": Exit Function
            End If
            If cboSTType.ListIndex = 1 Then
                If Not GetRS(rsTmp, "Select A.AdjAmt1,A.AdjAmt2 From " & GetOwner & "CD071 A," & GetOwner & "CM003 B Where A.CodeNo(+) = B.WorkClass And B.EmpNo = '" & garyGi(0) & "'") Then Exit Function
                If Not rsTmp.EOF Then
                    If IsNull(rsTmp("AdjAmt1")) And IsNull(rsTmp("AdjAmt2")) Then
                        MsgBox "無權限做帳務調改功能!!", vbCritical, gimsgPrompt
                        Exit Function
                    Else
                        If lngSubAmt < Val(rsTmp("AdjAmt1") & "") Or lngSubAmt > Val(rsTmp("AdjAmt2") & "") Then
                            MsgBox "調改金額" & lngSubAmt & "已超出您的權限範圍，請修正!!", vbCritical, gimsgPrompt
                            Exit Function
                        End If
                    End If
                End If
                If lngEditMode = giEditModeEdit Then
                    strCPGroupCode = GetRsValue("Select CPGroupCode From " & GetOwner & "SO033 Where RowId = '" & rsMdbRec(IIf(intDBType = giAccessDb, "Rowid", "RcdRowId")) & "'") & ""
                    
                    If strCPGroupCode & "" <> "" Then
                        'select sum(ShouldAmt) into v_ ShouldAmt from so033 where BillNo=[目前該筆的BillNo] and CPGroupCode=[目前該筆的CPAdjCitemCode] and Cancelflag=0。(取可調改之總額上限)
                        lngTotalAdjAmt = Val(GetRsValue("Select Sum(ShouldAmt) from " & GetOwner & "SO033 where BillNo = '" & txtBillNo & "' and CPGroupCode = '" & strCPGroupCode & "' and Cancelflag=0") & "")
                        If lngTotalAdjAmt <= 0 Then
                            MsgBox "該筆收費資料為負數，將無法使用調改帳單功能！"
                            Exit Function
                        Else
                            '如果該筆之ShouldAmt >= v_ShouldAmt時，則可調改金額不可大於v_ShouldAmt值。
                            If gnbShouldAmt.Value >= lngTotalAdjAmt And lngSubAmt > lngTotalAdjAmt Then
                                MsgBox "調改金額已超出金額" & lngTotalAdjAmt & "上限，請修正!", vbExclamation, gimsgPrompt
                                Exit Function
                            End If
                            '如果該筆之ShouldAmt < v_ShouldAmt時，則可調改金額不可大於ShouldAmt值。
                            If gnbShouldAmt.Value < lngTotalAdjAmt And lngSubAmt > gnbShouldAmt.Value Then
                                MsgBox "調改金額已超出金額" & gnbShouldAmt.Value & "上限，請修正!", vbExclamation, gimsgPrompt
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
        '*******************************************************************************************
        '#3465 判斷收費項項目是By機或By戶 By Kin 2007/08/28
        If GetByHouse(gilCitemCode.GetCodeNo) = 0 Then
            If txtFaciSNo.Tag = "" Then
                MsgBox "該收費項目為 BY機收費項目,請選擇設備!!", vbCritical, gimsgPrompt
                cmdHouse.Value = True
                Exit Function
            End If
        End If
        '*******************************************************************************************
        '91/09/30 檢查手開單號
        'Call clsChargeValidate.RealDateValidate(gdaRealDate.GetValue, gnbRealAmt.Value, strXLastDate, Cancel)
        'If Cancel Then Exit Function
'        If txtManualNo <> "" And YUseManual = 1 Then
'            If GetRsValue("Select Count(*) From " & GetOwner & "SO127 Where PaperNum = '" & txtManualNo & "' And CompCode = " & gCompCode) = 0 Then
'                MsgBox "無此手開單號,請查證!!", vbExclamation, gimsgPrompt
'                On Error Resume Next
'                txtManualNo.SetFocus
'                Exit Function
'                'If MsgBox("無此手開單號,確定存檔??", vbQuestion + vbYesNo + vbDefaultButton2, gimsgPrompt) = vbNo Then Exit Function
'            End If
'            If GetRsValue("Select Count(*) From " & GetOwner & "SO127 Where PaperNum = '" & txtManualNo & "' And CompCode = " & gCompCode & " And BillNo <> '" & txtBillNo & "'") > 0 Then
'                MsgBox "該手開單已有收費資料,請查證!!", vbExclamation, gimsgPrompt
'                On Error Resume Next
'                txtManualNo.SetFocus
'                Exit Function
'            End If
'        End If

        '******************************************************************
        '#3163 檢查手開單號,以Jacky所寫的Function為準 By Kin 2007/12/20
        If txtManualNo.Text <> "" Then
            If Not ChkManualNoOk(txtManualNo, txtBillNo) Then
              On Error Resume Next
              txtManualNo.SetFocus
              Exit Function
            End If
        End If
        '******************************************************************
        
        '******************************************************************
        '#3436 如果有帳號資料一定要有發票抬頭資料 By Kin 2007/11/27
        If cboAccountNo.Text <> "" Or txtAccountNo.Text <> "" Then
            If strInvSeqNo = Empty Then
                MsgBox "請選擇發票抬頭資料", vbInformation, "訊息"
                cmdInvSeqNo.SetFocus
                Exit Function
            End If
        End If
        '******************************************************************
        
        '#XXXX 判斷是否有第二層優惠 By Kin 2009/04/01
        If gilSTCode.GetCodeNo <> "" Then
            If ChkHave2Discount(gilCitemCode.GetCodeNo) Then MsgBox "有第2層優惠不得做短收!!", vbCritical, gimsgPrompt: gilSTCode.SetFocus: Exit Function
            If Chk2Discount(gilCitemCode.GetCodeNo) Then MsgBox "第2層優惠不得做短收!!", vbCritical, gimsgPrompt: gilSTCode.SetFocus: Exit Function
            If Chk3Discount(gilCitemCode.GetCodeNo) Then MsgBox "第3層優惠不得做短收!!", vbCritical, gimsgPrompt: gilSTCode.SetFocus: Exit Function
        End If


        IsDataOK = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "IsDataOK")
End Function

Private Function ReturnDate(ByVal strDate As String) As Date
On Error GoTo chkErr
    If strDate <> "" Then
        ReturnDate = CDate(strDate) + Ypara5
    End If
Exit Function
chkErr:
        Call ErrSub(Me.Name, "Returndate")
End Function

Private Sub GetPeriod_Amt(ByRef intReturnPeriod As Integer, ByRef intReturnAmt As Long)
On Error GoTo chkErr
    '若隨收費資料改變2（YPara9)為1（原應收額），則：
    '   期數=<該筆原收費期數>，金額=<該筆原應收金額>
    '若YPara9為 2（出單金額），則：
    '   期數=<該筆收費期數>，金額=<該筆出單金額>
Dim rsTmp As New ADODB.Recordset
Dim strSQL As String
    rsTmp.MaxRecords = 1
    strSQL = "Select OldAmt,OldPeriod From " & GetOwner & "So033 Where BillNo='" & strBillNo & "' and Item =" & intItem
    If Ypara9 = 1 Then
            If OpenRecordset(rsTmp, strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False, False) <> giOK Then MsgBox "連接失敗！請重新操作！", vbExclamation, "訊息！"
            intReturnPeriod = rsTmp("Oldperiod").Value
            intReturnAmt = rsTmp("OldAmt").Value
            rsTmp.Close
            Set rsTmp = Nothing
    Else
            intReturnPeriod = Val(gnbRealPeriod.Text)
            intReturnAmt = Val(gnbShouldAmt.Text)
    End If
    
Exit Sub
chkErr:
        Call ErrSub(Me.Name, "GetPeriod_Amt")
End Sub

Private Function ScrToRcd() As Boolean
    On Error GoTo chkErr
    Dim rsScr As New ADODB.Recordset '在ScrtoRcd 使用的Recordset
    Dim strSQL As String
    Dim lngItem As Long '記錄要儲存時的項目編號
        'If Not IsDataOk Then Exit Function
        'strsql = "Select Max(Item)+1 as MaxItem From So033 Where Billno='" & strBillNo & "'"
        'If Not GetRS(rsScr, strsql) Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": Exit Function
        '記錄項目編號
        'If Not (rsScr.EOF Or IsNull(rsScr("MaxItem").Value)) Then
        '    lngItem = rsScr("MaxItem").Value
        'Else
        '    lngItem = 1
        'End If
        'rsScr.Close
        'Set rsScr = Nothing
        
        If rsMdbRec.State = adStateClosed Then
            If Not GetRS(rsMdbRec, "Select * from So3311", Conn, adUseClient, adOpenKeyset, adLockOptimistic) Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": Unload Me: Exit Function
        End If
        
        With rsMdbRec
            If lngEditMode = giEditModeInsert Then
                lngItem = GetMaxItem(rsMdbRec)
                .AddNew
                .Fields("Item").Value = lngItem
                .Fields("BillNo").Value = strBillNo
                .Fields("RowId").Value = lngItem
            Else
                If intDBType = giAccessDb Then .Fields("RowID").Value = strRowId
            End If
            
            .Fields("PrtSno").Value = strPrtSNo
            .Fields("CustName").Value = strCustName
            .Fields("CustId").Value = lngCustId
            .Fields("CItemCode").Value = gilCitemCode.GetCodeNo
            .Fields("CItemName").Value = gilCitemCode.GetDescription
            .Fields("ShouldAmt").Value = NoZero(gnbShouldAmt.Value, True)
            .Fields("ShouldDate").Value = Format(gdaShouldDate.GetValue, "####/##/##")
            .Fields("RealDate").Value = gdaRealDate.GetValue(True)
            If gnbRealPeriod.Value > 0 Then
                .Fields("RealPeriod").Value = NoZero(gnbRealPeriod.Value, True)
                .Fields("RealStartDate").Value = gdaRealStartDate.GetValue(True)
                .Fields("RealStopDate").Value = gdaRealStopDate.GetValue(True)
            Else
                .Fields("RealPeriod").Value = 0
                .Fields("RealStartDate").Value = Null
                .Fields("RealStopDate").Value = Null
            End If
            .Fields("EntryEn").Value = garyGi(0)
            .Fields("Note").Value = NoZero(txtNote)
            .Fields("CancelFlag").Value = 0
            If gilClctEn.GetDescription <> "" Then
                .Fields("ClctEn").Value = gilClctEn.GetCodeNo
                .Fields("ClctName").Value = gilClctEn.GetDescription
            Else
                .Fields("ClctEn").Value = Null
                .Fields("ClctName").Value = Null
            End If
            .Fields("CmCode").Value = gilCMCode.GetCodeNo
            .Fields("CMName").Value = gilCMCode.GetDescription
            If gilPTCode.GetDescription <> "" Then
                .Fields("PTCode").Value = gilPTCode.GetCodeNo
                .Fields("PTName").Value = gilPTCode.GetDescription
            Else
                .Fields("PTCode").Value = Null
                .Fields("PTName").Value = Null
            End If
            'If gnbRealAmt.Value > 0 Then
                .Fields("RealAmt").Value = NoZero(gnbRealAmt.Value, True)
            'Else
                '.Fields("RealAmt").Value = 0
            'End If
            If gilSTCode.GetDescription <> "" Then
                .Fields("STCode").Value = gilSTCode.GetCodeNo
                .Fields("STName").Value = gilSTCode.GetDescription
            Else
                .Fields("STCode").Value = Null
                .Fields("STName").Value = Null
            End If
            .Fields("ManualNo").Value = NoZero(txtManualNo.Text)
            .Fields("ServiceType") = gServiceType
            .Fields("CompCode") = GetRsValue("Select Compcode from " & GetOwner & "so001 where custid=" & lngCustId)
            
            .Fields("BankCode") = gilBankCode.GetCodeNo2
            .Fields("BankName") = gilBankCode.GetDescription2
            .Fields("AccountNo") = NoZero(txtAccountNo)
            .Fields("AuthorizeNo") = NoZero(txtAuthorizeNo)
            .Fields("AdjustFlag") = cboSTType.ListIndex
            
            .Fields("NextPeriod") = ginNextPeriod.Value
            .Fields("NextAmt") = ginNextAmt.Value
            '********************************************************************************************
            '#3465 如果是By戶則更新MDB檔的FaciSeqNo為客編，反之更新所選定的設備流水號 By Kin 2007/08/28
            If GetByHouse(gilCitemCode.GetCodeNo) = 1 Then
                .Fields("FaciSeqNo") = lngCustId
            Else
                .Fields("FaciSeqNo") = txtFaciSNo.Tag
            End If
            '********************************************************************************************
            If strInvSeqNo <> Empty Then
                .Fields("InvSeqNo") = strInvSeqNo
            End If
            .Update
        End With
        rsMdbRec.Close
        Set rsMdbRec = Nothing
Exit Function
chkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Function

Private Function GetNumberType(ByVal strAmount As String) As Long
    GetNumberType = Val(Format(strAmount, "00000000"))
End Function

Private Sub txtNote_Validate(Cancel As Boolean)
    txtNote.IMEMode = 2
End Sub

Private Sub subGil()
    On Error GoTo chkErr
        '#XXXX 判斷是否有第二層優惠 By Kin 2009/04/01
        If lngEditMode = giEditModeInsert Then
            SetgiList gilCitemCode, "CodeNo", "Description", "CD019", , , , , 3, 12, "Where Nvl(ServiceType,'" & gServiceType & "' ) ='" & gServiceType & "' And Nvl(RefNo,0) Not In (19,20) ", True
        Else
            SetgiList gilCitemCode, "CodeNo", "Description", "CD019", , , , , 3, 12, "Where Nvl(ServiceType,'" & gServiceType & "' ) ='" & gServiceType & "' And (StopFlag=0 or RefNo=9)"
        End If

        'SetgiList gilCitemCode, "CodeNo", "Description", "CD019", , , , , 3, 12, "Where Nvl(ServiceType,'" & gServiceType & "' ) ='" & gServiceType & "' And (StopFlag=0 or RefNo=9)"
        gilCitemCode.ListIndex = 1
        SetgiList gilCMCode, "CodeNo", "Description", "CD031", , , , , 3, 12, , True
        SetgiList gilUCCode, "CodeNo", "Description", "CD013", , , , , 3, 12, , True
        SetgiList gilSTCode, "CodeNo", "Description", "CD016", , , , , 3, 12, , True
        SetgiList gilClctEn, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True
        SetgiList gilPTCode, "CodeNo", "Description", "CD032", , , , , 3, 12, , True
        '#5267 原本是取SO002A,現在改取SO106 By Kin 2010/04/26
        SetgiList gilBankCode, "BankCode", "BankName", "SO106", , , , , , , "Where CustId = " & gCustId & " And nvl(StopFlag,0) = 0 Group By BankCode,BankName"
         
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

'接收由SO3311b傳來之ROWID 修改時使用！
Public Property Let uRowId(ByVal vData As String)
    strRowId = vData
End Property

'接收由So3311B傳來之單據編號或印單序號！新增時使用！
Public Property Let uBillNo(ByVal vData As String)
    strBillNo = vData
End Property

'接收由So3311b傳來之印單序號
Public Property Let uPrtSNo(ByVal vData As String)
    strPrtSNo = vData
End Property

'由So3311B傳來之項目編號
Public Property Let uItem(ByVal vData As Integer)
    intItem = vData
End Property

Public Property Let uRealDate(ByVal vData As String)
    MstrRealDate = vData
End Property

Public Property Let uClctEn(ByVal vData As String)
    MstrClctEn = vData
End Property

Public Property Let uClctName(ByVal vData As String)
    MstrClctName = vData
End Property

Public Property Let uPTCode(ByVal vData As Integer)
    MintPTCode = vData
End Property

Public Property Let uPTName(ByVal vData As String)
    MstrPTName = vData
End Property

'接收由So3311B傳來之客戶編號和客戶姓名！
    'lngCustID = frmSO3311B.uCustid
    'strCustName = frmSO3311B.uCustName
    'strNote = frmSO3311B.uNote
Public Property Let uCustId(ByVal vData As Long)
    lngCustId = vData
End Property

Public Property Let uCustName(ByVal vData As String)
    strCustName = vData
End Property

''接收由So331AA 傳來之服務類別
'Public Property Let uServiceType(ByVal vData As String)
'    strServiceType = vData
'End Property

'接收由So3311b傳來之公司別
Public Property Let uCompCode(ByVal vData As String)
    strCompCode = vData
End Property

Public Property Set uParentForm(ByVal vData As Object)
    Set objParentForm = vData
End Property

Public Property Let uDBType(ByVal vData As ChangeDB)
    intDBType = vData
End Property

Private Sub EnableFaciSNo()
    On Error Resume Next
        If GetByHouse(gilCitemCode.GetCodeNo) = 1 Then
            lblFaciSNo.Visible = False
            txtFaciSNo.Visible = False
            cmdHouse.Visible = False
        Else
            lblFaciSNo.Visible = True
            txtFaciSNo.Visible = True
            cmdHouse.Visible = True
        End If

End Sub
