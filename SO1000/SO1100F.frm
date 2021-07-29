VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{99A4DA9D-8D59-11D3-8311-0080C8453DDF}#11.4#0"; "GiAddress1.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO1100F 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶帳戶資料 [SO1100F]"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   1425
   ClientWidth     =   11880
   Icon            =   "SO1100F.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSingle 
      Caption         =   "單筆指定"
      Height          =   375
      Left            =   9390
      TabIndex        =   49
      Top             =   1380
      Width           =   1275
   End
   Begin prjGiGridR.GiGridR ggrData3 
      Height          =   1815
      Left            =   180
      TabIndex        =   47
      Top             =   2340
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   3201
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
   Begin VB.CommandButton cmdCreateINV 
      BackColor       =   &H00C0FFFF&
      Caption         =   "發票抬頭維護"
      Height          =   375
      Left            =   8070
      Style           =   1  '圖片外觀
      TabIndex        =   46
      Top             =   1380
      Width           =   1275
   End
   Begin VB.CommandButton cmdUpdate2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "更新至客戶資料主檔供日後月費扣款"
      Enabled         =   0   'False
      Height          =   525
      Left            =   9930
      Style           =   1  '圖片外觀
      TabIndex        =   44
      Top             =   8340
      Width           =   1695
   End
   Begin VB.CheckBox chkChargeAddress 
      Caption         =   "更新收費地址至客戶主檔的收費地址"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4560
      TabIndex        =   41
      Top             =   8085
      Width           =   3225
   End
   Begin VB.CheckBox chkMailAddress 
      Caption         =   "更新郵寄地址至客戶主檔的收費地址"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2040
      TabIndex        =   34
      Top             =   8655
      Width           =   3225
   End
   Begin VB.CheckBox chkMailAddress2 
      Caption         =   "更新郵寄地址至客戶帳戶的收費地址"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5880
      TabIndex        =   33
      Top             =   8655
      Width           =   3225
   End
   Begin VB.CheckBox chkChargeAddress2 
      Caption         =   "更新收費地址至客戶帳戶的收費地址"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5880
      TabIndex        =   32
      Top             =   8025
      Width           =   3225
   End
   Begin VB.TextBox txtChargeTitle 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1590
      MaxLength       =   50
      TabIndex        =   31
      Top             =   7170
      Width           =   8370
   End
   Begin VB.TextBox txtInvTitle 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3780
      MaxLength       =   50
      TabIndex        =   30
      Top             =   7275
      Width           =   2445
   End
   Begin VB.TextBox txtInvAddress 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6990
      MaxLength       =   90
      TabIndex        =   29
      Top             =   7275
      Width           =   4995
   End
   Begin VB.ComboBox cboInvoiceType 
      Enabled         =   0   'False
      Height          =   300
      ItemData        =   "SO1100F.frx":0442
      Left            =   1590
      List            =   "SO1100F.frx":044F
      Style           =   2  '單純下拉式
      TabIndex        =   28
      Top             =   7275
      Width           =   1395
   End
   Begin VB.CheckBox chkShowPay 
      BackColor       =   &H80000018&
      Caption         =   "顯示帳戶支付費用"
      Height          =   375
      Left            =   2190
      Style           =   1  '圖片外觀
      TabIndex        =   26
      Top             =   1380
      Width           =   1725
   End
   Begin prjGiGridR.GiGridR ggrData2 
      Height          =   1485
      Left            =   3330
      TabIndex        =   25
      Top             =   4350
      Visible         =   0   'False
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   2619
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
      Height          =   3855
      Left            =   150
      TabIndex        =   13
      Top             =   60
      Width           =   11595
      Begin VB.CommandButton cmdDelSeqNo 
         Caption         =   "刪除發票抬頭"
         Height          =   375
         Left            =   6600
         TabIndex        =   50
         Top             =   1320
         Width           =   1275
      End
      Begin VB.CommandButton cmdInvSeqNo 
         Caption         =   "選擇發票抬頭"
         Height          =   375
         Left            =   4230
         TabIndex        =   27
         Top             =   1800
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CheckBox chkStop 
         Caption         =   "停用"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3870
         TabIndex        =   22
         Top             =   1440
         Width           =   675
      End
      Begin VB.ComboBox cboID 
         Height          =   300
         ItemData        =   "SO1100F.frx":046F
         Left            =   5010
         List            =   "SO1100F.frx":047C
         TabIndex        =   1
         Top             =   195
         Width           =   1905
      End
      Begin VB.TextBox txtNote 
         Height          =   390
         Left            =   5820
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   9
         Top             =   900
         Width           =   5625
      End
      Begin VB.TextBox txtCVC2 
         Height          =   315
         Left            =   10920
         MaxLength       =   3
         TabIndex        =   3
         Top             =   210
         Width           =   495
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
         Left            =   6030
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "*******"
         Top             =   555
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox txtAccountNo 
         Height          =   315
         Left            =   7860
         MaxLength       =   16
         TabIndex        =   2
         Top             =   195
         Width           =   2370
      End
      Begin prjGiList.GiList gilBankCode 
         Height          =   315
         Left            =   1050
         TabIndex        =   0
         Top             =   195
         Width           =   3240
         _ExtentX        =   5715
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
         FldWidth2       =   2500
         F2Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin MSMask.MaskEdBox mskCardExpDate 
         Height          =   315
         Left            =   6015
         TabIndex        =   5
         Tag             =   "OK"
         Top             =   555
         Width           =   1170
         _ExtentX        =   2064
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
      Begin prjGiList.GiList gilCardName 
         Height          =   315
         Left            =   1050
         TabIndex        =   4
         Top             =   555
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
      Begin MSComctlLib.Toolbar tbrMenu 
         Height          =   360
         Left            =   30
         TabIndex        =   16
         Top             =   -330
         Visible         =   0   'False
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   635
         ButtonWidth     =   1667
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imlToolbarIcons"
         DisabledImageList=   "imlToolbarIcons3"
         HotImageList    =   "imlToolbarIcons2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "新 增"
               Key             =   "AddNew"
               Object.ToolTipText     =   "(F6)"
               ImageKey        =   "IMG1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修 改"
               Key             =   "Edit"
               Object.ToolTipText     =   "(F11)"
               ImageKey        =   "IMG2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刪除"
               Key             =   "刪除"
               Object.ToolTipText     =   "(F10)"
               ImageKey        =   "IMG3"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "儲 存"
               Key             =   "Update"
               Object.ToolTipText     =   "(F2)"
               ImageKey        =   "IMG4"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "取 消"
               Key             =   "Cancel"
               Object.ToolTipText     =   "(&X)"
               ImageKey        =   "IMG5"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "第一筆"
               Key             =   "First"
               Object.ToolTipText     =   "(&F)"
               ImageKey        =   "IMG8"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "上一筆"
               Key             =   "Previous"
               Object.ToolTipText     =   "(&P)"
               ImageKey        =   "IMG9"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "下一筆"
               Key             =   "Next"
               Object.ToolTipText     =   "(&N)"
               ImageKey        =   "IMG10"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "最後筆"
               Key             =   "Last"
               Object.ToolTipText     =   "(&L)"
               ImageKey        =   "IMG11"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "離 開"
               Key             =   "Exit"
               Object.ToolTipText     =   "(ESC)"
               ImageKey        =   "IMG12"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         BorderStyle     =   1
         MousePointer    =   99
         MouseIcon       =   "SO1100F.frx":04AA
         Begin VB.TextBox txtStatus 
            Appearance      =   0  '平面
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   10140
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Edit Mode"
            Top             =   26
            Width           =   555
         End
      End
      Begin prjGiList.GiList gilCMCode 
         Height          =   315
         Left            =   9090
         TabIndex        =   7
         Top             =   555
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         BackColor       =   14737632
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
         FldWidth1       =   405
         FldWidth2       =   1620
         F2Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin Gi_Date.GiDate gdaStopDate 
         Height          =   285
         Left            =   5460
         TabIndex        =   23
         Top             =   1380
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
         ReplaceText     =   -1  'True
      End
      Begin CS_Multi.CSmulti csmAuthenticID 
         Height          =   360
         Left            =   90
         TabIndex        =   8
         Top             =   930
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   635
         ButtonCaption   =   "認證帳號"
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "[ 發票資料區 ]"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   120
         TabIndex        =   48
         Top             =   2010
         Width           =   1515
      End
      Begin VB.Label lblAuthenticINV 
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   150
         TabIndex        =   45
         Top             =   1350
         Width           =   1785
      End
      Begin VB.Label lblStopDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "停用日期"
         Height          =   180
         Left            =   4620
         TabIndex        =   24
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label lblCMCode 
         Caption         =   "收費方式"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8250
         TabIndex        =   21
         Top             =   615
         Width           =   750
      End
      Begin VB.Label lblRemark 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "備 註"
         Height          =   180
         Left            =   5370
         TabIndex        =   20
         Top             =   1020
         Width           =   405
      End
      Begin VB.Label lblCVC2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "檢查碼"
         Height          =   180
         Left            =   10320
         TabIndex        =   19
         Top             =   270
         Width           =   540
      End
      Begin VB.Label lblID 
         AutoSize        =   -1  'True
         Caption         =   "帳戶別"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   4380
         TabIndex        =   18
         Top             =   270
         Width           =   540
      End
      Begin VB.Label lblCardExpDate 
         Caption         =   "信用卡有效期(西曆MM/YYYY)"
         Height          =   240
         Left            =   3510
         TabIndex        =   12
         Top             =   615
         Width           =   2385
      End
      Begin VB.Label lblAccountNo 
         AutoSize        =   -1  'True
         Caption         =   "帳號/卡號"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   7005
         TabIndex        =   11
         Top             =   270
         Width           =   765
      End
      Begin VB.Label lblBankCode 
         Caption         =   "銀         行"
         Height          =   240
         Left            =   165
         TabIndex        =   15
         Top             =   270
         Width           =   810
      End
      Begin VB.Label lblCardName 
         AutoSize        =   -1  'True
         Caption         =   "信用卡別"
         Height          =   180
         Left            =   165
         TabIndex        =   14
         Top             =   615
         Width           =   720
      End
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   2880
      Left            =   135
      TabIndex        =   10
      Top             =   4170
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   5080
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
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   0
      Top             =   4050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":07C4
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":0ADE
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":0F30
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":108A
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":11E4
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":133E
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":1790
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":18EA
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":1C04
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":1F1E
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":2238
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":2552
            Key             =   "IMG12"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons2 
      Left            =   0
      Top             =   4620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":26AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":29C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":2E18
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":2F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":30CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":3226
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":3678
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":37D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":3AEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":3E06
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":4120
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":443A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons3 
      Left            =   0
      Top             =   5190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":4594
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":48AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":4D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":4E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":4FB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":510E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":5428
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":5742
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":5A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO1100F.frx":5D76
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin PrjGiAddress1.GiAddress1 ga1ChargeAddress 
      Height          =   360
      Left            =   690
      TabIndex        =   35
      Top             =   7665
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AddressCaption  =   "收費地址"
   End
   Begin PrjGiAddress1.GiAddress1 ga1MailAddress 
      Height          =   360
      Left            =   690
      TabIndex        =   36
      Top             =   8325
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AddressCaption  =   "郵寄地址"
      Enabled         =   0   'False
   End
   Begin MSMask.MaskEdBox mskInvNo 
      Height          =   315
      Left            =   11010
      TabIndex        =   37
      Top             =   7170
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   8
      Mask            =   "########"
      PromptChar      =   "_"
   End
   Begin VB.Label lblInvAddress 
      Caption         =   "發票地址"
      Height          =   240
      Left            =   10350
      TabIndex        =   43
      Top             =   8130
      Width           =   720
   End
   Begin VB.Label lblInvTitle 
      Caption         =   "發票抬頭"
      Height          =   240
      Left            =   3120
      TabIndex        =   42
      Top             =   7320
      Width           =   780
   End
   Begin VB.Label lblChargeTitle 
      AutoSize        =   -1  'True
      Caption         =   "收  件  人"
      Height          =   180
      Left            =   705
      TabIndex        =   40
      Top             =   7455
      Width           =   720
   End
   Begin VB.Label lblInvNo 
      Caption         =   "統一編號"
      Height          =   240
      Left            =   10125
      TabIndex        =   39
      Top             =   7455
      Width           =   810
   End
   Begin VB.Label lblInvoiceType 
      Caption         =   "發票種類"
      Height          =   240
      Left            =   705
      TabIndex        =   38
      Top             =   7350
      Width           =   930
   End
End
Attribute VB_Name = "frmSO1100F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Programmer: Power Hammer
' Last Modify: 2002/06/27

Option Explicit
Private lngEditMode As giEditModeEnu ' 記錄目前在編輯、新增或檢視模式
Private WithEvents rsSO002A As ADODB.Recordset
Attribute rsSO002A.VB_VarHelpID = -1
Private WithEvents rsSO138 As ADODB.Recordset
Attribute rsSO138.VB_VarHelpID = -1
Private strInvSeqNo As String
Dim rsSO138Tmp As New ADODB.Recordset
Dim strInvoiceType As String
Dim strInvTitle As String
Dim strInvNo As String
Dim strInvAddress As String
Dim strSO002BInv As String
Private strSingleInv As String
Private strSingleAcc As String
Private blnShow As Boolean
Private objParentForm As Form
Private rsFrom138 As New ADODB.Recordset
Private blnDelete As Boolean
Private blnShowData As Boolean
Private strAccountName As String
Private Sub cboID_Click()
  On Error Resume Next
    
    Dim intActLen As Integer
    Dim intVNO As Long
    Dim strVNO As String
    Dim iCustCount As Long
    If EditMode = giEditModeInsert Then
        If cboID.ListIndex = 2 Then
'            txtAccountNo = frmSO1100BMDI.txtCustId
            If gilBankCode.GetCodeNo <> Empty Then
                '#3541 改變虛擬帳號規則 By Kin 2007/11/30
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
                
'                intVNO = GetRsValue("SELECT COUNT(*)+1 FROM " & GetOwner & _
'                                    "SO002A WHERE CUSTID=" & gCustId & _
'                                    " AND COMPCODE=" & gCompCode & _
'                                    " AND ID=2")
                txtAccountNo = strVNO & String(intActLen - Len(CStr(strVNO)) - Len(CStr(gCustId)), "0") & gCustId
            End If
        Else
            txtAccountNo.MaxLength = 16
        End If
    End If
    
End Sub

Private Sub cboInvoiceType_Click()
  On Error GoTo ChkErr
    If EditMode <> giEditModeView Then
        If cboInvoiceType.ListIndex = 2 Then
            lblInvNo.ForeColor = vbRed
        Else
            lblInvNo.ForeColor = vbBlack
        End If
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "cboInvoiceType_Click"
End Sub
'******************************************************************************************************************************
'#3364 增加一個ChkBox，控制是否顯示帳戶支付費月 By Kin 2007/08/28 Add
Private Sub chkShowPay_Click()
    On Error Resume Next
        If showPay Then
            If chkShowPay = 1 Then
                ggrData.Width = Me.ScaleWidth / 3
                ggrData2.Move ggrData.Width, ggrData.Top, fraData.Width - (Me.ScaleWidth / 3) + 150, ggrData.Height
                ggrData2.Visible = True
                ggrData2.Refresh
            Else
                ggrData.Width = fraData.Width
                ggrData2.Visible = False
            End If
        End If
'        Call showPay
End Sub
'*****************************************************************************************************************************

Private Sub cmdCreateINV_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
    
'    If EditMode = giEditModeView Then
'        ChangeMode giEditModeEdit
'        lngOldEditMode = giEditModeView
'    End If
    With frmSO114FA
        .OldEditMode = Me.EditMode
        .uCustId = gCustId
        .uOwnerForm = Me
        '******************************************
        '#3789 傳入帳號，以便開啟已儲存的發票資訊 By Kin 2008/02/29
        If EditMode = giEditModeEdit Then
            .uAccountNo = Trim(txtAccountNo.Text)
        Else
            If txtAccountNo.Text <> "" Then
                .uAccountNo = IIf(IsNumeric(txtAccountNo.Text), txtAccountNo.Text, txtAccountNo.Tag) '  Trim(txtAccountNo.Text)
            End If
        End If
        .uMutilChoice = True
        If ggrData3.Recordset.RecordCount > 0 Then
            .uShowAbs = ggrData3.Recordset("InvSeqNo")
        Else
            .uShowAbs = Empty
        End If

'        If rsSO138.RecordCount > 0 Then
'            .uShowAbs = rsSO138("InvSeqNo")
'        Else
'            .uShowAbs = Empty
'        End If
        '******************************************
        .Show , Me
        
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
  ErrSub Me.Name, "cmdCreateINV"
End Sub

Private Sub cmdDelSeqNo_Click()
  On Error GoTo ChkErr
    Dim strDelSQL As String
    Dim rsTmp1 As New ADODB.Recordset
    Dim rsTmp2 As New ADODB.Recordset
    Dim strQry1 As String
    Dim strQry2 As String
    Dim strUPDSO003 As String
    Dim strUPDSO033 As String
    Dim intDelTotal As Integer
    blnDelete = False
    '*********************************************************************************************************************
    '#3929 刪除發票資料如果有週期性資料或待收收費目,則要秀出Grid By Kin 2008/06/04
    If ggrData3.Recordset.RecordCount <= 0 Then MsgBox "無任何發票資料！", vbInformation, "訊息": Exit Sub
    strQry1 = "Select * From " & GetOwner & "SO003" & _
              " Where AccountNo='" & IIf(IsNumeric(txtAccountNo.Text), txtAccountNo.Text, txtAccountNo.Tag) & "'" & _
              " And CustId=" & gCustId & _
              " And CompCode=" & gCompCode & _
              " And InvSeqNo=" & ggrData3.Recordset("InvSeqNo")
    strQry2 = "Select * From " & GetOwner & "SO033 " & _
            " Where AccountNo='" & IIf(IsNumeric(txtAccountNo.Text), txtAccountNo.Text, txtAccountNo.Tag) & "'" & _
            " And CustId=" & gCustId & _
            " And CompCode=" & gCompCode & _
            " And UCCode is Not Null" & _
            " And InvSeqNo=" & ggrData3.Recordset("InvSeqNo")
    strUPDSO003 = "Update " & GetOwner & "SO003 Set InvSeqNo=Null,ACcountNo=Null" & _
                " Where AccountNo='" & IIf(IsNumeric(txtAccountNo.Text), txtAccountNo.Text, txtAccountNo.Tag) & "'" & _
                " And CustId=" & gCustId & _
                " And CompCode=" & gCompCode & _
                " And InvSeqNo=" & ggrData3.Recordset("InvSeqNo")
    strUPDSO033 = "Update " & GetOwner & "SO033 Set InvSeqNo=Null,AccountNo=Null" & _
                " Where AccountNo='" & IIf(IsNumeric(txtAccountNo.Text), txtAccountNo.Text, txtAccountNo.Tag) & "'" & _
                " And CustId=" & gCustId & _
                " And CompCode=" & gCompCode & _
                " And InvSeqNo=" & ggrData3.Recordset("InvSeqNo")
    strDelSQL = "Delete From " & GetOwner & "SO002AD" & _
                    " Where CustID=" & gCustId & _
                    " And AccountNo='" & IIf(rsSO002A.EOF, "0", rsSO002A("AccountNo")) & "'" & _
                    " And CompCode=" & gCompCode & _
                    " And InvSeqNo=" & ggrData3.Recordset("InvSeqNo")
                    
    '**********************************************************************************************************************
    If Not GetRS(rsTmp1, strQry1, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If Not GetRS(rsTmp2, strQry2, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    intDelTotal = 0
    '#3929 刪除資料如果有週期性項目或收費資料要秀出瀏覽畫面 By Kin 2008/06/02
    gcnGi.BeginTrans
    If ggrData3.Recordset.RecordCount > 0 Then
        
        If Not rsTmp1.EOF Or Not rsTmp2.EOF Then
            With frmSO1100F1
                 .uRs1 = rsTmp1
                 .uRs2 = rsTmp2
                .Show 1
            End With
            If blnDelete Then
                gcnGi.Execute strDelSQL, intDelTotal
                '#3929測試不OK 要一起將SO003與SO033一併清除
                gcnGi.Execute strUPDSO003
                gcnGi.Execute strUPDSO033
                ggrData3.Recordset.Delete
                DoEvents
                If intDelTotal > 0 Then
                    OpenSO138 False
                End If
                ggrData.Refresh
            End If
        Else
            If MsgBox("是否確定刪除此筆發票資料?", vbQuestion + vbYesNo, "訊息") = vbYes Then
                gcnGi.Execute strDelSQL
                ggrData3.Recordset.Delete
                DoEvents
                If intDelTotal > 0 Then
                    OpenSO138 False
                End If
                ggrData3.Refresh
            End If
        End If
    Else
        MsgBox "無任何發票資料 ！", vbInformation, "訊息"
    End If
    gcnGi.CommitTrans
    On Error Resume Next
    Call Close3Recordset(rsTmp1)
    Call Close3Recordset(rsTmp2)
    Exit Sub
ChkErr:
  gcnGi.RollbackTrans
  ErrSub Me.Name, "cmdDelSeqNo"
End Sub

Private Sub cmdInvSeqNo_Click()
  On Error GoTo ChkErr
    With frmSO1131H
        .uParentForm = Me
        .uMutilChoice = True
        .Show 1
    End With
    Set rsSO138Tmp = frmSO1131H.uRS
    If rsSO138Tmp.State = adStateClosed Then Exit Sub
    rsSO138Tmp.Filter = "Choice='1'"
    If Not AddChoiceInv(rsSO138Tmp) Then Exit Sub
    Set ggrData3.Recordset = rsSO138
    ggrData3.Refresh
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdInvSeqNo"
End Sub

Private Sub cmdSingle_Click()
  On Error Resume Next
  objParentForm.uAccountNo = txtAccountNo.Tag & ""
  objParentForm.uInvSeqNo = rsSO138("InvSeqNo") & ""
  Unload Me
End Sub

'記錄目前使用者權限
'Private blnUserPrivAddNew   As Boolean   '新增
'Private blnUserPrivUpdate   As Boolean  '修改
'加一欄位ChargeTitle收件人VARCHAR2(50)

Private Sub cmdUpdate2_Click()
  On Error GoTo ChkErr
    gcnGi.Execute "UPDATE " & GetOwner & "SO002 SET ACCOUNTNO='" & txtAccountNo.Tag & "'," & _
                    "BANKCODE=" & IIf(gilBankCode.GetCodeNo = Empty, "NULL", gilBankCode.GetCodeNo2) & "," & _
                    "BANKNAME=" & IIf(gilBankCode.GetDescription = Empty, "NULL", "'" & gilBankCode.GetDescription2 & "'") & "," & _
                    "CARDNAME=" & IIf(gilCardName.GetDescription = Empty, "NULL", "'" & gilCardName.GetDescription2 & "'") & "," & _
                    "CARDEXPDATE=" & IIf(mskCardExpDate.Text = Empty, "NULL", mskCardExpDate.Text) & "," & _
                    "CVC2='" & txtCVC2 & "'" & _
                    " WHERE CUSTID=" & gCustId & " AND COMPCODE=" & gCompCode
    MsgBox "更新至客戶資料主檔完成!", vbOKOnly, "訊息"
  Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdUpdate2_Click"
End Sub

Public Sub DeleteGo()
  On Error GoTo ChkErr
    Dim strSQL As String
    Dim AdjFlag As Boolean
    If rsSO002A.EOF Then
        MsgBox "已無資料 !", vbInformation, "訊息"
        GoTo 66
    End If
    If (frmSO1100BMDI.rsSo002!AccountNo & "") = (rsSO002A!AccountNo & "") Then
        If MsgBox("注意 !! 目前客戶正使用此帳號扣款 !! " & vbCrLf & vbCrLf & "您確定刪除 [帳號/卡號] 為 " & rsSO002A!AccountNo & " 的客戶帳戶資料嗎 ?", vbQuestion + vbYesNo, "詢問") = vbYes Then
            AdjFlag = True
            gcnGi.Execute "DELETE FROM " & GetOwner & "SO002A WHERE CUSTID=" & gCustId & " AND ACCOUNTNO='" & rsSO002A!AccountNo & "' AND COMPCODE=" & gCompCode
            
            '#3436 一併刪除發票資料 By Kin 2007/11/28
            gcnGi.Execute "DELETE FROM " & GetOwner & "SO002AD Where CustID=" & gCustId & " And AccountNo='" & rsSO002A("AccountNo") & "' And CompCode=" & gCompCode
        Else
            Exit Sub
        End If
    Else
        If MsgBox("確定刪除 [帳號/卡號] 為 " & rsSO002A!AccountNo & " 的客戶帳戶資料嗎 ?", vbQuestion + vbYesNo, "詢問") = vbYes Then
            AdjFlag = False
            gcnGi.Execute "DELETE FROM " & GetOwner & "SO002A WHERE CUSTID=" & gCustId & " AND ACCOUNTNO='" & rsSO002A!AccountNo & "' AND COMPCODE=" & gCompCode
            
            '#3436 一併刪除發票資料 By Kin 2007/11/28
            gcnGi.Execute "DELETE FROM " & GetOwner & "SO002AD Where CustID=" & gCustId & " And AccountNo='" & rsSO002A("AccountNo") & "' And CompCode=" & gCompCode
        Else
            Exit Sub
        End If
    End If
66:
    With rsSO002A
         If .State = adStateOpen Then
             strSQL = rsSO002A.Source
             .Close
             .Open strSQL, gcnGi, adOpenKeyset, adLockOptimistic
             Set ggrData.Recordset = rsSO002A
         End If
         ggrData.Refresh
         If .RecordCount = 0 Then
            NewRcd
            ga1ChargeAddress.AddrNo = 0
            ga1ChargeAddress.AddrString = ""
            ga1MailAddress.AddrNo = 0
            ga1MailAddress.AddrString = ""
            chkStop.Value = 0
            gdaStopDate.Text = ""
            cboInvoiceType.ListIndex = -1
            InitializingListOcx
            ChangeMode giEditModeView
         End If
         If AdjFlag Then
            Me.Hide
            frmSO1100BMDI.AdjCMBank
            Unload Me
         End If
    End With
  Exit Sub
ChkErr:
    ErrSub Me.Name, "DeleteGo"
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub csmAuthenticID_SelectChange()
  On Error Resume Next
    If csmAuthenticID.GetQueryCode = Empty Then
        lblAuthenticINV.Caption = ""
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    ReleaseCOM Me
End Sub

Private Sub RcdToScr()
  On Error GoTo ChkErr
    If rsSO002A.RecordCount = 0 Then Exit Sub
    With rsSO002A
         gilBankCode.SetCodeNo .Fields("BankCode") & ""
         gilBankCode.SetDescription .Fields("BankName") & ""
         
'         txtID = .Fields("ID") & ""
         cboID.ListIndex = Val(.Fields("ID") & "")
'         txtAccountNo = .Fields("AccountNo") & ""
         If GetUserPriv("SO1100G", "(SO1100G9)") Then
            txtAccountNo.Text = .Fields("AccountNo") & ""
            txtAccountNo.Tag = .Fields("AccountNo") & ""
         Else
            txtAccountNo.Tag = .Fields("AccountNo") & ""
            If txtAccountNo.Tag = Empty Then
                txtAccountNo.Text = ""
            Else
                txtAccountNo.Text = Left(txtAccountNo.Tag, 5) & "XXX" & Mid(txtAccountNo.Tag, 9, 2) & "XX" & Right(txtAccountNo.Tag, 4)
            End If
         End If
         
         
         txtCVC2 = .Fields("CVC2") & ""
         gilCardName.SetDescription .Fields("CardName") & ""
         gilCardName.Query_CodeNo
         mskCardExpDate.Text = .Fields("CardExpDate") & ""
         'ga1ChargeAddress.AddrNo = Val(.Fields("ChargeAddrNo") & "")
         'ga1ChargeAddress.AddrString = .Fields("ChargeAddress") & ""
         'ga1MailAddress.AddrNo = Val(.Fields("MailAddrNo") & "")
         'ga1MailAddress.AddrString = .Fields("MailAddress") & ""
         'txtChargeTitle = .Fields("ChargeTitle") & ""
         If Len(.Fields("CardExpDate")) > 1 Then
            mskCardExpDate.Text = Right("0" & (.Fields("CardExpDate") & ""), 6)
            txtHide.Text = mskCardExpDate.Text
         Else
            mskCardExpDate.Text = ""
            txtHide.Text = ""
         End If
         txtNote = .Fields("Note") & ""
'         txtInvTitle.Text = .Fields("InvTitle") & ""
'         mskInvNo.Text = .Fields("InvNo") & ""
'         txtInvAddress.Text = .Fields("InvAddress") & ""
         chkStop.Value = .Fields("StopFlag")
         gdaStopDate.Text = .Fields("StopDate") & ""
'         Select Case (.Fields("InvoiceType") & "")
'                Case 0
'                     cboInvoiceType.ListIndex = 0
'                Case 2
'                     cboInvoiceType.ListIndex = 1
'                Case 3
'                     cboInvoiceType.ListIndex = 2
'                Case Else
'                     cboInvoiceType.ListIndex = -1
'         End Select
         On Error Resume Next
         gilCMCode.SetCodeNo GetRsValue("SELECT CMCODE FROM " & GetOwner & "SO106 WHERE CUSTID=" & gCustId & " AND ACCOUNTID='" & rsSO002A("AccountNo") & "' AND COMPCODE=" & gCompCode & IIf(gilBankCode.GetCodeNo = Empty, "", " AND BANKCODE=" & gilBankCode.GetCodeNo))
         gilCMCode.Query_Description
        If .Fields("AuthenticId") & "" = Empty Then
            csmAuthenticID.Clear
        Else
            csmAuthenticID.SetQueryCode .Fields("AuthenticId") & ""
        End If
        Call OpenSO138
        
        lblAuthenticINV.Caption = ""
        If csmAuthenticID.GetQueryCode <> "" Then
            If .Fields("AuthenticId") & "" = "" Then
                lblAuthenticINV.Caption = ""
                strSO002BInv = Empty
            Else
                strSO002BInv = gcnGi.Execute("Select InvSeqNo From " & GetOwner & _
                                                "SO002B Where AuthenticId IN(" & .Fields("AuthenticId") & "" & ")")(0) & ""
                If strSO002BInv <> "" Then
                    lblAuthenticINV.Caption = "發票流水編號： " & strSO002BInv
                    strInvoiceType = gcnGi.Execute("Select InvoiceType From " & GetOwner & "SO138 Where InvSeqNO=" & strSO002BInv & "")(0)
                End If
            End If
        End If
    
    End With
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScr")
End Sub

Private Function ScrToRcd() As Boolean
  On Error GoTo ChkErr
    Dim rsUpd2A As New ADODB.Recordset
    gcnGi.Execute "Delete From " & GetOwner & "SO002AD Where CompCode=" & gCompCode & _
                    " And CustId=" & gCustId & _
                    " And AccountNo='" & IIf(lngEditMode = giEditModeEdit, txtAccountNo.Tag, txtAccountNo.Text) & "'"
    If rsSO138.RecordCount > 0 Then
        Dim rsSO002AD As New ADODB.Recordset
        If Not GetRS(rsSO002AD, "Select * From " & GetOwner & "SO002AD Where 1=0", gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
        rsSO138.MoveFirst
        Do While Not rsSO138.EOF
            DoEvents
            With rsSO002AD
                .AddNew
                .Fields("InvSeqNo") = rsSO138.Fields("InvSeqNo") & ""
                .Fields("CompCode") = gCompCode
                .Fields("CustId") = gCustId
                If lngEditMode = giEditModeEdit Then
                    .Fields("AccountNo") = txtAccountNo.Tag & ""
                Else
                    .Fields("AccountNo") = txtAccountNo.Text & ""
                End If
                .Update
            End With
                
            rsSO138.MoveNext
        Loop
    End If
    
    
    With rsSO002A
        
         If EditMode = giEditModeInsert Then .AddNew
         
        .Fields("COMPCODE") = gCompCode & ""
        .Fields("CustID") = gCustId
        .Fields("BankCode") = SolveNull(gilBankCode.GetCodeNo & "")
        .Fields("BankName") = SolveNull(gilBankCode.GetDescription & "")
'         If cboID.Text = Empty Then
'            .Fields("ID") = Null
'         Else
        .Fields("ID") = Left(cboID.Text, 1)
'         End If
        
        If EditMode = giEditModeEdit Then
             If IsNumeric(txtAccountNo) Then
                .Fields("AccountNo") = SolveNull(txtAccountNo & "")
             Else
                .Fields("AccountNo") = txtAccountNo.Tag
             End If
        ElseIf EditMode = giEditModeInsert Then
            .Fields("AccountNo") = SolveNull(txtAccountNo & "")
        End If
         
        '.Fields("ChargeTitle") = txtChargeTitle.Text
        .Fields("CVC2") = txtCVC2.Text
        .Fields("CardName") = SolveNull(gilCardName.GetDescription & "")
        .Fields("CardExpDate") = SolveNull(mskCardExpDate.Text & "")
        '.Fields("ChargeAddrNo") = ga1ChargeAddress.AddrNo
        '.Fields("ChargeAddress") = ga1ChargeAddress.AddrString
        '.Fields("MailAddrNo") = ga1MailAddress.AddrNo
        '.Fields("MailAddress") = ga1MailAddress.AddrString
        .Fields("Note") = txtNote.Text
        '.Fields("InvTitle") = IIf(txtInvTitle.Text = "", Null, txtInvTitle.Text)
        '.Fields("InvNo") = IIf(mskInvNo.Text = "", Null, mskInvNo.Text)
        '.Fields("InvAddress") = IIf(txtInvAddress.Text = "", Null, txtInvAddress.Text)
'         Select Case cboInvoiceType.ListIndex
'                Case 0
'                    .Fields("InvoiceType") = 0
'                Case 1
'                    .Fields("InvoiceType") = 2
'                Case 2
'                    .Fields("InvoiceType") = 3
'                Case Else
'                    .Fields("InvoiceType") = Null
'         End Select
        If csmAuthenticID.GetQueryCode <> Empty Then
            .Fields("AuthenticId") = csmAuthenticID.GetQueryCode(True)
        Else
            .Fields("AuthenticId") = Null
        End If
        .Update
    End With
    ScrToRcd = True
  Exit Function
ChkErr:
    ScrToRcd = False
    gcnGi.RollbackTrans
    Call ErrSub(Me.Name, "ScrToRcd")
End Function

Public Sub FunctionKey(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Select Case KeyCode
           '----------------------------------------------------
           Case vbKeyF2  '   F3:存檔, 相當於按下cmdsave
                If EditMode = giEditModeView Then Exit Sub
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

Private Function IsDataOk()
  On Error GoTo ChkErr
    Dim lngRefNo As Long
    IsDataOk = False
    
    '*****************************************************************************************************************
    '#3228判斷收件人最大長度 By Kin 2007/07/04
    '#3375將SO1100F訊息”收信人輸入長度不得超過50。改為"收信人輸入長度不得超過25個中文字" By Kin 2007/08/01
    If StrLen(txtChargeTitle) > txtChargeTitle.MaxLength Then
        txtChargeTitle.SetFocus
        MsgBox "收信人輸入長度不得超過" & txtChargeTitle.MaxLength / 2 & "個中文字!!", vbInformation, "警告訊息"
        Exit Function
    End If
    '*****************************************************************************************************************
    
    
    '*************************************************************************************************************************
    '#3395 收費方式為信用卡時(CD031.RefNo=4)，使用信用卡檢核方式，收費方式為銀行時規則不變(CD031.RefNo=2) By Kin 2007/08/28
    '#3395 測試不OK,原本要傳入信用卡別,結果傳成收費方式代碼 By Kin 2007/09/18
    If gilCMCode.GetCodeNo & "" <> "" Then
        lngRefNo = GetRsValue("Select nvl(RefNo,0) From " & GetOwner & "CD031 Where CodeNo=" & gilCMCode.GetCodeNo & "")
    End If
    If lngRefNo = 4 Then
        If Not ChkCreditCard(gilCardName.GetCodeNo, txtAccountNo) Then Exit Function
    End If
    '*************************************************************************************************************************
'    If txtID = "" Then txtID.SetFocus: GoTo 66

    If cboID.Text = Empty Then cboID.SetFocus: GoTo 66
    If txtAccountNo = "" Then txtAccountNo.SetFocus: GoTo 66
'    If cboInvoiceType.ListIndex = 2 Then
'        If mskInvNo = "" Then mskInvNo.SetFocus: GoTo 66
'    End If
'    If mskInvNo <> Empty Then
'        If Not InvNoIsOk(mskInvNo.Text, True) Then mskInvNo.SetFocus: Exit Function
'        If cboInvoiceType.ListIndex <> 2 Then
'            MsgBox "[統一編號] 有值 ! 須將 [發票種類] 設成 [ 3.三聯式 ] !", vbInformation, "訊息"
'            Exit Function
'        End If
'    End If
    
    '**************************************************************************
    '#3436 判斷發票資料是否有值 By Kin 2007/11/28
    '#3929 如果沒有選擇發票自動跳出視窗 By Kin 2008/06/02
    If rsSO138.RecordCount = 0 Then
        MsgBox "請選擇發票抬頭資料", vbInformation, "訊息"
        'cmdInvSeqNo.SetFocus
        If cmdCreateINV.Enabled Then cmdCreateINV.Value = True
        Exit Function
    End If
    '**************************************************************************
    
    If csmAuthenticID.GetQueryCode <> Empty Then
        If lblAuthenticINV = "" Then
            MsgBox "請在[ 發票資料區 ]點選一筆發票抬頭資料！！", vbInformation, "訊息"
            Exit Function
        End If
    
    End If
    
    If ServiceCommon.AccountNo_Validate(Me) Then
        txtAccountNo.SetFocus
        Exit Function
    End If
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
    'ggrData.Blank = True
    gilBankCode.Clear
    gilCardName.Clear
'    txtID = 0
    cboID.ListIndex = 0
    txtAccountNo = ""
    txtChargeTitle = ""
    txtCVC2 = ""
    mskCardExpDate.Text = ""
    mskInvNo = ""
    txtInvTitle = ""
    txtInvAddress = ""
    gilCMCode.Clear
    Call OpenSO138(True)
'    With ga1ChargeAddress
'        .AddrNo = frmSO1100BMDI.ChildForm.ga1ChargeAddress.AddrNo
'        .AddrString = frmSO1100BMDI.ChildForm.ga1ChargeAddress.AddrString
'    End With
'    With ga1MailAddress
'        .AddrNo = frmSO1100BMDI.ChildForm.ga1MailAddress.AddrNo
'        .AddrString = frmSO1100BMDI.ChildForm.ga1MailAddress.AddrString
'    End With
    csmAuthenticID.Clear
    On Error Resume Next
    gilBankCode.SetFocus
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "NewRcd")
End Sub

Public Sub AddNewSO002A(objParent As Form, _
                                            strAddr As String, _
                                            strInvTitle As String, _
                                            strInvoiceType As String, _
                                            strInvNo As String, _
                                            strName As String)
  On Error GoTo ChkErr
'    Set ObjActiveForm = Me
'    strActFrmName = "frmSO1100F"
    Show , objParent
    AddNewGo
    PrcAuth
    'lngEditMode = giEditModeInsert
    ' 發票地址、發票抬頭、發票種類、統一編號取由SO002B之資訊。
    ' SO002B.會員姓名填為SO002A.收件人。
    ' SO002A.ID為 2　，銀行名稱帶入之判斷為
    
    'SO002A.ID為 2　，銀行名稱帶入之判斷為
    ' (SELECT CODENO,DESCRIPTION FROM CD018 WHERE Prgname is null and Actlength=16)。
    ' 此判斷需再加強SO002A.ID為 2　，銀行名稱帶入之判斷為
    ' (SELECT CODENO,DESCRIPTION FROM CD018 WHERE Prgname= AGENCSINGBANK and Actlength=16 AND CHCHKATM=1 )。
    
'    gilBankCode.SetCodeNo GetRsValue("SELECT CODENO FROM " & GetOwner & "CD018" & _
                                                                " WHERE PRGNAME IS NULL AND ACTLENGTH=16")
    gilBankCode.SetCodeNo GetRsValue("SELECT CODENO FROM " & GetOwner & "CD018" & _
                                                                " WHERE PRGNAME='AGENCSINGBANK' AND ACTLENGTH=16 AND CHECKATM=1")
    gilBankCode.Query_Description
    ' (SELECT CODENO,DESCRIPTION FROM CD018 WHERE Prgname is null and Actlength=16)。
    If strAddr <> Empty Then txtInvAddress = strAddr
    If strInvTitle <> Empty Then txtInvTitle = strInvTitle
    If strInvNo <> Empty Then mskInvNo = strInvNo
    
    If Not (strAddr = "" And strInvTitle = "" And strInvNo = "") Then
        If strInvoiceType <> Empty Then cboInvoiceType = strInvoiceType
    End If
    
    cboID.ListIndex = 2
    txtChargeTitle = strName ' SO002B.會員姓名填為SO002A.收件人
    ' 收費地址及郵寄地址取SO002為預設值。
  Exit Sub
ChkErr:
    ErrSub Name, "AddNewSO002A"
End Sub

Public Sub AddNewGo()
  On Error GoTo ChkErr
    ChangeMode (giEditModeInsert)
    NewRcd
'    txtChargeTitle = frmSO1100BMDI.txtCustName
'    txtInvTitle.Text = frmSO1100BMDI.rsSo002.Fields("InvTitle") & ""
'    mskInvNo.Text = frmSO1100BMDI.rsSo002.Fields("InvNo") & ""
'    txtInvAddress.Text = frmSO1100BMDI.rsSo002.Fields("InvAddress") & ""
'    cboInvoiceType.ListIndex = frmSO1100BMDI.ChildForm.cboInvoiceType.ListIndex
'    Select Case (frmSO1100BMDI.rsSo002.Fields("InvoiceType") & "")
'           Case 0
'                cboInvoiceType.ListIndex = 0
'           Case 2
'                cboInvoiceType.ListIndex = 1
'           Case 3
'                cboInvoiceType.ListIndex = 2
'           Case Else
'                cboInvoiceType.ListIndex = -1
'    End Select
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "cmdAdd_Click")
End Sub

Public Sub UpdateGo()
  On Error GoTo ChkErr
    Dim strSQL As String
    Dim lngChargeAddrNo As Long
    Dim lngMailAddrNo As Long
    If IsDataOk() = False Then Exit Sub
    gcnGi.BeginTrans
    
'    If EditMode = giEditModeEdit Then
'        lngChargeAddrNo = rsSO002A!ChargeAddrNo
'        lngMailAddrNo = rsSO002A!MailAddrNo
'    End If
'
'    With ga1ChargeAddress
'        If chkChargeAddress.Value Then
'            If .AddrNo <> lngChargeAddrNo Then
'                If frmSO1100BMDI.ChildForm.ga1ChargeAddress.AddrNo = rsSO002A!ChargeAddrNo Then
'                    gcnGi.Execute "UPDATE " & GetOwner & "SO001 SET ChargeAddrNo=" & .AddrNo & ",ChargeAddress='" & .AddrString & "' WHERE CUSTID=" & gCustId & " AND COMPCODE=" & garyGi(9)
'                End If
'            End If
'        End If
'    End With
'
'    With ga1MailAddress
'        If chkMailAddress.Value Then
'            If .AddrNo <> rsSO002A!MailAddrNo Then
'                If frmSO1100BMDI.ChildForm.ga1MailAddress.AddrNo = rsSO002A!MailAddrNo Then
'                    gcnGi.Execute "UPDATE " & GetOwner & "SO001 SET MailAddrNo=" & .AddrNo & ",MailAddress='" & .AddrString & "' WHERE CUSTID=" & gCustId & " AND COMPCODE=" & garyGi(9)
'                End If
'            End If
'        End If
'    End With
    
    ' 2.  當SO002A之 發票地址、發票抬頭、發票種類、統一編號有異動時，
    '       要同步異動到SO002B該AUTHENTICID之發票地址、發票抬頭、發票種類、統一編號。
    Dim blnInvChg As Boolean
    Dim strKey As String
    
    strKey = Trim(csmAuthenticID.GetQueryCode)
    
    If EditMode = giEditModeEdit And strKey <> Empty Then
        blnInvChg = True
'        Select Case True
'            Case Val(rsSO002A("InvoiceType").OriginalValue) <> Choose(cboInvoiceType.ListIndex + 1, 0, 2, 3)
'            Case rsSO002A("InvTitle").OriginalValue & "" <> txtInvTitle
'            Case rsSO002A("InvAddress").OriginalValue & "" <> txtInvAddress
'            Case rsSO002A("InvNo").OriginalValue & "" <> mskInvNo.Text
'            Case Else
'                    blnInvChg = False
'        End Select
    Else
        blnInvChg = True
    End If
    
    If Not ScrToRcd Then Exit Sub           ' 把控制項內的值，存到資料庫裡
    
    
    
    If blnInvChg And strKey <> Empty Then
        Dim strUpd2B As String
        strUpd2B = "UPDATE " & GetOwner & "SO002B" & _
                            " SET INVOICETYPE=" & strInvoiceType & _
                            ",INVTITLE=" & IIf(strInvTitle = Empty, "Null", "'" & strInvTitle & "'") & _
                            ",INVNO='" & strInvNo & "'" & _
                            ",InvSeqNo=" & strSO002BInv & _
                            ",INVADDRESS=" & IIf(strInvAddress = Empty, "Null", "'" & strInvAddress & "'") & _
                            " WHERE AUTHENTICID IN (" & strKey & ")"
        gcnGi.Execute strUpd2B
        ' 處理 Update SO002BLOG
        ' #5937 不做Log By Kin 2011/05/14
'        gcnGi.Execute "DELETE FROM " & GetOwner & "SO002BLOG" & _
'                                " WHERE AUTHENTICID IN (" & strKey & ")"
'        gcnGi.Execute "INSERT INTO " & GetOwner & "SO002BLOG" & _
'                                " ( SELECT * FROM " & GetOwner & "SO002B" & _
'                                " WHERE AUTHENTICID IN (" & strKey & "))"

    
        
'        Dim strUpd2B As String
'        strUpd2B = "UPDATE " & GetOwner & "SO002B" & _
'                            " SET INVOICETYPE=" & Choose(cboInvoiceType.ListIndex + 1, 0, 2, 3) & _
'                            ",INVTITLE='" & txtInvTitle & "'" & _
'                            ",INVNO='" & mskInvNo.Text & "'" & _
'                            ",INVADDRESS='" & txtInvAddress & "'" & _
'                            " WHERE AUTHENTICID IN (" & strKey & ")"
'        gcnGi.Execute strUpd2B
'        ' 處理 Update SO002BLOG
'        gcnGi.Execute "DELETE FROM " & GetOwner & "SO002BLOG" & _
'                                " WHERE AUTHENTICID IN (" & strKey & ")"
'        gcnGi.Execute "INSERT INTO " & GetOwner & "SO002BLOG" & _
'                                " ( SELECT * FROM " & GetOwner & "SO002B" & _
'                                " WHERE AUTHENTICID IN (" & strKey & "))"
    End If
    
    Dim lngAbsPos As Long
    
    If EditMode = giEditModeEdit Then
        lngAbsPos = rsSO002A.AbsolutePosition
'        With ga1ChargeAddress
'            If chkChargeAddress2.Value Then
'                If .AddrNo <> lngChargeAddrNo Then
'                    gcnGi.Execute "UPDATE " & GetOwner & "SO002A SET ChargeAddrNo=" & .AddrNo & ",ChargeAddress='" & .AddrString & "' WHERE CUSTID=" & gCustId & " AND COMPCODE=" & garyGi(9) & " AND ChargeAddrNo=" & lngChargeAddrNo
'                End If
'            End If
'        End With
'
'        With ga1MailAddress
'            If chkMailAddress2.Value Then
'                If .AddrNo <> lngMailAddrNo Then
'                    gcnGi.Execute "UPDATE " & GetOwner & "SO002A SET MailAddrNo=" & .AddrNo & ",MailAddress='" & .AddrString & "' WHERE CUSTID=" & gCustId & " AND COMPCODE=" & garyGi(9) & " AND MailAddrNo=" & lngMailAddrNo
'                End If
'            End If
'        End With
    End If
    
   'rsSO002A.Requery
    If rsSO002A.State = adStateOpen Then
        strSQL = rsSO002A.Source
        rsSO002A.Close
        rsSO002A.Open strSQL, gcnGi, adOpenKeyset, adLockOptimistic
        Set ggrData.Recordset = rsSO002A
    End If
    
    ggrData.Refresh
    gcnGi.CommitTrans
'    If EditMode = giEditModeInsert Then '繼續新增
'        Call ChangeMode(giEditModeInsert)
'        AddNewGo
'    Else '進入顯示模式
        Call ChangeMode(giEditModeView)
        RcdToScr
        If lngAbsPos <> 0 Then rsSO002A.AbsolutePosition = lngAbsPos
'    End If
   Exit Sub
ChkErr:
    gcnGi.RollbackTrans
    Call ErrSub(Me.Name, "UpdateGo")
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    
    Set ObjActiveForm = Me
    strActFrmName = "frmSO1100F"
    If blnShow Then
        cmdSingle.Enabled = True
    Else
        cmdSingle.Enabled = False
    End If
    '********************************************************************
    '#3929 改成複選,所以改讀取SO114FA整個RecordSet By Kin 2008/05/26
    If frmSO114FA.frmShow Then
        frmSO114FA.SetFocus
    Else
        If rsFrom138 Is Nothing Then
            Exit Sub
        End If
        If rsFrom138.State = adStateClosed Then
            Exit Sub
        End If
        '***************************************************************************************
        '如果SO114FA 有選擇資料,而且又是在瀏覽模式,要改成編輯模式,以便可以存檔 By Kin 2008/06/04
        If EditMode = giEditModeView Then ChangeMode giEditModeEdit
        '***************************************************************************************
        rsFrom138.Filter = "Choice='1'"
        If Not AddChoiceInv(rsFrom138) Then Exit Sub
        Set ggrData3.Recordset = rsSO138
        ggrData3.Refresh
    End If
    CloseRecordset rsFrom138
    strInvSeqNo = Empty
    Screen.MousePointer = vbDefault
    '**********************************************************************
End Sub

Private Sub Form_KeyPress(keyAscii As Integer)
  On Error Resume Next
   If keyAscii = Asc("'") Then keyAscii = 0
End Sub
Private Sub SetConcealSO002A()
    On Error GoTo ChkErr
    Dim aSQL As String
    
    
    
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "SetConcealSO002A")
End Sub
Private Sub Form_Load()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    
    InitializingListOcx
    
    If Crm Then
        PolyFrmFunction Me
        Me.Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2)
    Else

        'If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 + 300
        If Not 800600 Then Me.Move (Screen.Width - Me.Width) / 2, frmSO1100BMDI.Top + frmSO1100BMDI.tbrMenu.Height + 700
    End If
'    Me.Move (Screen.Width - Me.Width) / 2, frmSO1100BMDI.Top + frmSO1100BMDI.tbrMenu.Height + 600
    Set rsSO002A = New ADODB.Recordset
    Set rsSO138 = New ADODB.Recordset
    With rsSO002A
         If .State = 1 Then
            .Close
         Else
            .CursorLocation = adUseClient
         End If
         Dim aSQL As String
         Dim aSQL2AD As String
         aSQL2AD = "Select b.accountno,A.chargetitle " & _
                            " FROM  " & GetOwner & "SO138 a," & GetOwner & "so002ad b " & _
                            " WHERE Exists(Select B.InvSeqNo" & _
                                                " FROM " & GetOwner & "SO002AD B " & _
                                                " Where b.CustID = " & gCustId & _
                                                " AND B.CompCode=" & gCompCode & _
                                                " AND A.InvSeqNo=B.InvSeqNo) " & _
                                                " AND Nvl(Stopflag,0)=0 " & _
                                                " and a.InvSeqNo=b.InvSeqNo "
         
         aSQL = "SELECT ROWID,SO002A.* " & _
                    " FROM " & GetOwner & "SO002A SO002A " & _
                    " WHERE CUSTID= " & gCustId & " AND COMPCODE=" & gCompCode
'        aSQL = "SELECT SO002A.ROWID,SO002A.*,B.ChargeTitle " & _
'           " FROM " & GetOwner & "SO002A SO002A, (" & aSQL2AD & ") B " & _
'           " WHERE SO002A.CUSTID= " & gCustId & " AND SO002A.COMPCODE=" & gCompCode & _
'           " AND SO002A.ACCOUNTNO = B.ACCOUNTNO(+) "

        .Open aSQL, gcnGi, adOpenKeyset, adLockOptimistic
    End With
    
    With mFlds
        .Add "BankName", , , , , "銀行名稱" & Space(25)
        .Add "ID", , , , , "帳戶別 " & Space(10)
        .Add "AccountNo", , , , , "帳號/卡號" & Space(20)
        .Add "CVC2", , , , , "檢查碼" & Space(10)
        .Add "CardName", , , , , "信用卡別" & Space(10)
        .Add "CardExpDate", , , "00/0000", , "卡期限  " & Space(10)
        '.Add "ChargeAddress", , , , , "收費地址" & Space(26)
        '.Add "MailAddress", , , , , "郵寄地址" & Space(26)
    End With
    ggrData.AllFields = mFlds
    ggrData.SetHead
    Set ggrData.Recordset = rsSO002A
    ggrData.Refresh
    
    If Not OpenSO138 Then Exit Sub
'    '#6413 判斷是否要顯示完整資訊 By Kin 2013/03/21
'    blnShowData = True
'    If Not rsSO138.EOF Then
'        If Len(strAccountName) > 0 Then
'            blnShow = rsSO138("ChargeTitle") = strAccountName
'        End If
'    End If
    
    Call ChangeMode(giEditModeView)
    UserPermissionGo True
'    Dim strPermis As String
'    Call GetPermission(garyGi(4), "SO1162", strPermis)
'    Call SetPermission(strPermis, blnAdd:=blnUserPrivAddNew, AddKey:="SO11621" _
'                                , blnUpdate:=blnUserPrivUpdate, UpdKey:="SO11622")
                                    
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub InitializingListOcx()
  On Error GoTo ChkErr
    With ga1ChargeAddress
        .SendConn gcnGi
        .SetOwner GetOwner
        .SendEntry garyGi(1) & ""
        .ShowStrtCode
        .SetCompCodeAndServiceType gCompCode & ""
    End With
    With ga1MailAddress
        .SendConn gcnGi
        .SetOwner GetOwner
        .SendEntry garyGi(1) & ""
        .ShowStrtCode
        .SetCompCodeAndServiceType gCompCode & ""
    End With
    gilCardName.ListType = OneColumn
    SetLst gilCardName, "CodeNo", "Description", 3, 12, 405, 1620, "CD037", , True
    SetLst gilBankCode, "CodeNo", "Description", 3, 24, 405, 2500, "CD018", , True
    SetLst gilCMCode, "CodeNo", "Description", 3, 12, 405, 1620, "CD031"
    gilBankCode.Filter = " Where CompCode =" & gCompCode
    SetMsQry csmAuthenticID, "SELECT AUTHENTICID,NAME FROM " & GetOwner & "SO002B", _
                    "認證編號", "會員姓名", , , False
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
 
Public Sub EditGo()
  On Error GoTo ChkErr
    ' 子帳戶不能做異動，只有父帳戶才能做異動，而子帳戶只做連動處理。
    '（即InheritFlag=0的資料方可作異動，InheritFlag=1的資料將被ReadOnly）
    If Val(rsSO002A("InheritFlag") & "") = 1 Then
        MsgBox "子帳戶資料不能單獨做異動 !", vbInformation, "訊息"
        Exit Sub
    End If
    Call ChangeMode(giEditModeEdit)
    If rsSO002A.Fields("AuthenticId") & "" <> Empty Then
        csmAuthenticID.SetQueryCode rsSO002A.Fields("AuthenticId") & ""
        csmAuthenticID.SetQueryCode rsSO002A.Fields("AuthenticId") & ""
    End If
    On Error Resume Next
    gilBankCode.SetFocus
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "cmdEdit_Click")
End Sub

Public Sub CancelGo()
  On Error GoTo ChkErr
    If EditMode = giEditModeView Then
      Unload Me
      Exit Sub
    Else
        '*******************************************************************
        '#3929 如果在編輯模式取消時一定要有發票資料 By Kin 2008/06/04
        If EditMode = giEditModeEdit Then
            If ggrData3.Recordset.RecordCount = 0 Then
                MsgBox "請選擇發票抬頭資料", vbInformation, "訊息"
                If cmdCreateINV.Enabled Then cmdCreateINV.Value = True
                Exit Sub
            End If
        End If
        '******************************************************************
        If EditMode = giEditModeEdit Then
            RevertRcd
        Else
            If Not rsSO002A.EOF Then
                rsSO002A.MoveFirst
                RcdToScr
            Else
                gilBankCode.Clear
                gilCardName.Clear
                '            txtID = 0
                cboID.ListIndex = 0
                txtAccountNo = ""
                txtCVC2 = ""
                txtChargeTitle = ""
                mskCardExpDate.Text = ""
                cboID.Text = ""
                txtInvTitle = ""
                txtInvAddress = ""
                mskInvNo = ""
                gilCMCode.Clear
                cboInvoiceType.ListIndex = -1
                With ga1ChargeAddress
                    .AddrNo = 0
                    .AddrString = ""
                End With
                With ga1MailAddress
                    .AddrNo = 0
                    .AddrString = ""
                End With
            End If
        End If
        Call ChangeMode(giEditModeView)
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
        rsSO002A.Close
        Set rsSO002A = Nothing
        CloseRecordset rsSO138
        CloseRecordset rsSO138Tmp
        CloseRecordset rsFrom138
        
        Call FormQueryUnload
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_QueryUnload"
End Sub

Private Sub ga1ChargeAddress_Click()
  On Error GoTo ChkErr
    ga1ChargeAddress.SetParentPosition Me.Left, Me.Top, Me.Width, Me.Height
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ga1ChargeAddress_Click"
End Sub

Private Sub ga1MailAddress_Click()
  On Error GoTo ChkErr
    ga1MailAddress.SetParentPosition Me.Left, Me.Top, Me.Width, Me.Height
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ga1MailAddress_Click"
End Sub



Private Sub ggrData_MoveRecord()
  On Error GoTo ChkErr:
'    Dim strPaySQL As String
'    Dim rsPay As New ADODB.Recordset
'    If chkShowPay = 0 Then Exit Sub
'    strPaySQL = "SELECT D.CUSTID,B.PROPOSER,B.ACCOUNTID,C.CUSTNAME,D.CITEMNAME,E.FACISNO,E.DECLARANTNAME " & _
'                " FROM ( SELECT ACCOUNTID,PROPOSER FROM " & _
'                     "( SELECT ACCOUNTID,PROPOSER,RANK() OVER (PARTITION BY ACCOUNTID ORDER BY UPDTIME DESC ) SORTFLD " & _
'                      " FROM " & GetOwner & "SO106 WHERE ACCOUNTID='" & ggrData.Recordset("ACcountNo") & "'" & _
'                               " AND COMPCODE=" & gCompCode & ") C" & _
'                      " WHERE C.SORTFLD=1) B , SO001 C , SO003 D , SO004 E" & _
'                " WHERE D.ACCOUNTNO='" & ggrData.Recordset("AccountNo") & "'" & _
'                " AND D.COMPCODE=" & gCompCode & _
'                " AND D.ACCOUNTNO=B.ACCOUNTID(+)" & _
'                " AND D.COMPCODE=C.COMPCODE" & _
'                " AND D.CUSTID=C.CUSTID" & _
'                " AND D.FACISEQNO=E.SEQNO(+)" & _
'                " AND D.CUSTID=E.CUSTID(+)" & _
'                " AND D.ServiceType=E.ServiceType(+)" & _
'                " AND D.COMPCODE=E.COMPCODE(+)"
'    If Not GetRS(rsPay, strPaySQL, gcnGi, adUseClient) Then Exit Sub
'    subGrid2 rsPay
    
Exit Sub
ChkErr:
    ErrSub Me.Name, "ggrData_MoveRecord"
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If UCase(Fld.Name) = UCase("ID") Then Value = IIf(Value = 0, "帳號", "卡號")
    If UCase(Fld.Name) = "ACCOUNTNO" Then
        If Not GetUserPriv("SO1100G", "(SO1100G9)") Then
            If Value = Empty Then
                Value = ""
            Else
                Value = Left(Value, 5) & "XXX" & Mid(Value, 9, 2) & "XX" & Right(Value, 4)
            End If
        End If
       
    End If
     '#6413
    If Len(strAccountName) > 0 Then
        If strAccountName <> ggrData.Recordset("ChargeTitle") & "" Then
            Value = "XXXXXX"
        End If
        
    End If
'    If UCase(Fld.Name) = UCase("CardExpDate") Then Value = Format(Value, "00/0000")
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ggrData_ShowCellData"
End Sub

Private Sub ggrData3_DblClick()
  On Error Resume Next
    If csmAuthenticID.GetQueryCode <> Empty Then
        strInvAddress = ggrData3.Recordset("InvAddress") & ""
        strInvNo = ggrData3.Recordset("InvNo") & ""
        strInvoiceType = ggrData3.Recordset("InvoiceType") & ""
        strInvTitle = ggrData3.Recordset("InvTitle") & ""
        strSO002BInv = ggrData3.Recordset("InvSeqNo")
        lblAuthenticINV.Caption = "發票流水編號： " & ggrData3.Recordset("InvSeqNo")
    End If
End Sub

Private Sub ggrData3_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    If giFld.FieldName = "InvoiceType" Then
        Select Case Value
            Case 0
                Value = "免開"
            Case 2
                Value = "二聯式"
            Case 3
                Value = "三聯式"
        End Select
    End If
    '#6413
    If Len(strAccountName) > 0 Then
        If ggrData3.Recordset("ChargeTitle") & "" <> strAccountName Then
            Value = "XXXXX"
        End If
    End If
        
End Sub

Private Sub mskCardExpDate_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    Cancel = ServiceCommon.CardExpireDate(Me)
  Exit Sub
ChkErr:
    ErrSub Me.Name, "mskCardExpDate_Validate"
End Sub

Private Sub mskInvNo_LostFocus()
'  On Error Resume Next
'    Call InvNoIsOk(mskInvNo.Text, True)
End Sub

'Private Sub mskInvNo_Validate(Cancel As Boolean)
'  On Error Resume Next
'    Cancel = Not InvNoIsOk(mskInvNo.Text, True)
'End Sub

Private Sub rsSO002A_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  On Error GoTo ChkErr
    If adReason = 10 Then
        If Not rsSO002A.EOF And Not rsSO002A.BOF And lngEditMode = giEditModeView Then
            If Not ggrData.Enabled Then ggrData.Enabled = True
            RcdToScr
        Else
            ggrData.Enabled = False
        End If
    End If
    txtAccountNo.MaxLength = 16
    '*****************************************************************************
    '#3364 顯示帳戶支付費用 By Kin 2007/08/28 Add
    Call showPay
    '*****************************************************************************
    txtAccountNo.PasswordChar = ""
    txtCVC2.PasswordChar = ""
    txtHide.PasswordChar = ""
    '#3436 顯示發票資料 By Kin 2007/11/27
    If Not OpenSO138 Then Exit Sub
    '#6413
    If Not rsSO138.EOF Then
        If Len(strAccountName) > 0 Then
            If rsSO138("ChargeTitle") & "" <> strAccountName Then
                If IsNumeric(txtAccountNo.Text) Then
                    txtAccountNo.Tag = txtAccountNo.Text
                End If
                txtAccountNo.PasswordChar = "X"
                txtCVC2.PasswordChar = "X"
                txtHide.PasswordChar = "X"
                cboID.Text = "XXXX"
                gilBankCode.SetCodeNo "X"
                gilBankCode.SetDescription "XXXX"
                gilCardName.SetCodeNo "X"
                gilCardName.SetDescription "XXXX"
                gilCMCode.SetCodeNo "X"
                gilCMCode.SetDescription "XXXX"
            End If
        End If
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "rsSO002A_MoveComplete")
End Sub
Public Property Let uInvSeqNo(ByVal vData As String)
  On Error GoTo ChkErr
    strInvSeqNo = vData
    Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Let Let uInvSeqNo")
End Property
Public Property Get EditMode() As giEditModeEnu
  On Error GoTo ChkErr
    EditMode = lngEditMode    ' 取目前編輯模式
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Get EditMode")
End Property
Public Property Let uShow(ByVal vData As Boolean)
    blnShow = vData
End Property
Public Property Let uParentForm(ByVal vData As Form)
    Set objParentForm = vData
End Property
Public Property Let uAccountName(ByVal vData As String)
    strAccountName = vData
End Property
Public Property Let EditMode(ByVal vNewValue As giEditModeEnu)
  On Error GoTo ChkErr
    lngEditMode = vNewValue '設定編輯模式
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
        Case giEditModeInsert
            blnFlag = False
            MenuEnabled , , False, True, False, True
            
        Case giEditModeEdit
            blnFlag = False
            MenuEnabled , False, False, True, False, True
            
        Case giEditModeView
            blnFlag = True
            MenuEnabled True, rsSO002A.RecordCount > 0, rsSO002A.RecordCount > 0, False, False, True
            
    End Select
    SetStatusBar lngEditMode
    fraData.Enabled = Not blnFlag
    ggrData.Enabled = blnFlag
    'cmdCreateINV.Enabled = blnFlag
    cmdCreateINV.Enabled = True
    'ggrData3.Enabled = blnFlag
    UserPermissionGo blnFlag
    
    PrcAuth
    
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ChangeMode")
End Sub

Private Sub PrcAuth()
  On Error GoTo ChkErr
    If EditMode <> giEditModeView Then
        Dim strFilter As String
        On Error Resume Next
        strFilter = ""
        If EditMode = giEditModeEdit Then
            strFilter = RPxx(gcnGi.Execute("SELECT AUTHENTICID FROM " & GetOwner & "SO002A" & _
                                                            " WHERE ROWID <> '" & rsSO002A!ROWID & "'" & _
                                                            " AND AUTHENTICID IS NOT NULL" & _
                                                            " AND COMPCODE=" & gCompCode).GetString(adClipString, , "", ",", "NULL"))
        Else
            strFilter = RPxx(gcnGi.Execute("SELECT AUTHENTICID FROM " & GetOwner & "SO002A" & _
                                                            " WHERE AUTHENTICID IS NOT NULL" & _
                                                            " AND COMPCODE=" & gCompCode).GetString(adClipString, , "", ",", "NULL"))
        End If
        If Err = 3021 Then
            Err.Clear
        Else
            If Len(strFilter) > 0 Then
                strFilter = " WHERE AUTHENTICID NOT IN ( " & Left(strFilter, Len(strFilter) - 1) & " )" & _
                                " AND CUSTID=" & gCustId
            End If
        End If

        SetMsQry csmAuthenticID, "SELECT AUTHENTICID,NAME FROM " & GetOwner & "SO002B" & strFilter, _
                                            "認證編號", "會員姓名", , , False
    Else
        SetMsQry csmAuthenticID, "SELECT AUTHENTICID,NAME FROM " & GetOwner & "SO002B", _
                                            "認證編號", "會員姓名", , , False
    End If
  Exit Sub
ChkErr:
    ErrSub Name, "PrcAuth"
End Sub

Private Sub UserPermissionGo(blnFlag As Boolean)
  On Error GoTo ChkErr
    MenuEnabled GetUserPriv("SO1100F", "(SO1100F1)") And blnFlag, _
                GetUserPriv("SO1100F", "(SO1100F2)") And rsSO002A.RecordCount > 0 And blnFlag, _
                GetUserPriv("SO1100F", "(SO1100F3)") And rsSO002A.RecordCount > 0 And blnFlag, Not blnFlag, , True
' INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName) VALUES
' ( 'SO1100F6', null, '更新至客戶主檔', '', 'SO1100F', null);
    cmdUpdate2.Enabled = GetUserPriv("SO1100F", "(SO1100F6)")
    chkChargeAddress.Enabled = GetUserPriv("SO1100F", "(SO1100F6)")
    chkChargeAddress2.Enabled = GetUserPriv("SO1100F", "(SO1100F6)")
    chkMailAddress.Enabled = GetUserPriv("SO1100F", "(SO1100F6)")
    chkMailAddress2.Enabled = GetUserPriv("SO1100F", "(SO1100F6)")
    'cable_lawrence: INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
    'VALUES ( 'SO1100F1', null, '資料新增', '', 'SO1100F', null);
    'INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
    'VALUES ( 'SO1100F2', null, '資料修改', '', 'SO1100F', null);
    'INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
    'VALUES ( 'SO1100F3', null, '資料刪除', '', 'SO1100F', null);
    'INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
    'VALUES ( 'SO1100F4', null, '資料查看', '', 'SO1100F', null);
    
    '********************************************************
    '#3929 多增加刪除發票資料權限 By Kin 2008/06/03
    If GetUserPriv("SO1100F", "(SO1100F7)") Then
        cmdDelSeqNo.Enabled = True
    Else
        cmdDelSeqNo.Enabled = False
    End If
    '********************************************************
    
    '******************************************************
    '#3929 判斷是否有選擇發票資料的權限 By Kin 2008/06/02
    'Call PermissionGo114FA
    '******************************************************
  Exit Sub
ChkErr:
    ErrSub Me.Name, "UserPermissionGo"
End Sub
Private Sub PermissionGo114FA()
    cmdCreateINV.Enabled = True
    If GetUserPriv("SO114FA", "(SO114FA1)") Then Exit Sub
    If GetUserPriv("SO114FA", "(SO114FA2)") Then Exit Sub
    If GetUserPriv("SO114FA", "(SO114FA3)") Then Exit Sub
    cmdCreateINV.Enabled = False
End Sub

Private Sub rsSO138_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  On Error Resume Next
'    Dim aSQL As String
'    Dim aPos As Integer
'    Dim mFlds As New GiGridFlds
'    If Len(strAccountName) > 0 Then
'        If strAccountName <> rsSO138("ChargeTitle") & "" Then
'            ggrData.Recordset("BankName2") = "XXXXX"
'            ggrData.Recordset("AccountNo2") = "XXXXXX"
'        End If
'    End If
    
End Sub

Private Sub txtAccountNo_KeyPress(keyAscii As Integer)
  On Error Resume Next
    ML keyAscii
End Sub
'******************************************************************************************************************************
'#3364 顯示帳戶支付費用的Grid By Kin 2007/08/28 Add
Private Sub subGrid2(rs As Recordset)
    On Error Resume Next
    Dim mFlds As New GiGridFlds
    With mFlds
        .Add "Proposer", , , , , "帳戶申請人" & Space(5)
        .Add "CustID", , , , , "客戶編號" & Space(10)
        .Add "CustName", , , , , "客戶名稱          "
        .Add "FaciSNO", , , , , "設備序號" & Space(8)
        .Add "DECLARANTNAME", , , , , "設備申請人   "
        .Add "CitemName", , , , , "收費項目        "
    End With
    ggrData2.AllFields = mFlds
    ggrData2.SetHead
    Set ggrData2.Recordset = rs
    ggrData2.Refresh
End Sub
'*******************************************************************************************************************************
'***********************************************************************************************************************************
''#3364 顯示帳戶支付費月 By Kin 2007/08/28 Add
Private Function showPay() As Boolean
  On Error GoTo ChkErr
    Dim strPaySQL As String
    Dim rsPay As New ADODB.Recordset
    showPay = False
    If EditMode <> giEditModeView Then chkShowPay = 0
    If chkShowPay = 0 Then showPay = True: Exit Function
    If rsSO002A.EOF Then MsgBox "無帳戶資料", vbInformation, "訊息":  chkShowPay.Value = 0: Exit Function
    '#3823 過濾沒有收費項目的客戶,相同客編收費項目只要出現一次就好 By Kin 2008/03/21
    strPaySQL = "SELECT DISTINCT D.CUSTID,B.PROPOSER,B.ACCOUNTID,C.CUSTNAME,D.CITEMNAME,E.FACISNO,E.DECLARANTNAME " & _
                " FROM ( SELECT ACCOUNTID,PROPOSER,BANKCODE FROM " & _
                     "( SELECT ACCOUNTID,PROPOSER,BANKCODE,RANK() OVER (PARTITION BY ACCOUNTID ORDER BY UPDTIME DESC ) SORTFLD " & _
                      " FROM " & GetOwner & "SO106 WHERE ACCOUNTID='" & rsSO002A("ACcountNo") & "'" & _
                               " AND COMPCODE=" & gCompCode & ") C" & _
                      " WHERE C.SORTFLD=1) B , SO001 C , SO003 D , SO004 E" & _
                " WHERE D.ACCOUNTNO='" & rsSO002A("AccountNo") & "'" & _
                " AND D.BANKCODE='" & rsSO002A("BankCode") & "'" & _
                " AND D.COMPCODE=" & gCompCode & _
                " AND D.ACCOUNTNO=B.ACCOUNTID(+)" & _
                " AND D.COMPCODE=C.COMPCODE" & _
                " AND D.CUSTID=C.CUSTID" & _
                " AND D.FACISEQNO=E.SEQNO(+)" & _
                " AND D.CUSTID=E.CUSTID(+)" & _
                " AND D.ServiceType=E.ServiceType(+)" & _
                " AND D.COMPCODE=E.COMPCODE(+)" & _
                " AND D.BANKCODE=B.BANKCODE" & _
                " AND D.CITEMCODE IS NOT NULL"
                
    If Not GetRS(rsPay, strPaySQL, gcnGi, adUseClient) Then Exit Function
    If rsPay.EOF And ggrData2.Visible = False Then MsgBox "無帳戶支付費用", vbInformation, "訊息": chkShowPay.Value = 0: Exit Function
    subGrid2 rsPay
    showPay = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "showPay"
End Function
'************************************************************************************************************************************

'#3436 開啟SO138資料 By Kin 2007/11/27
Public Function OpenSO138(Optional ByVal blnAddNew As Boolean = False) As Boolean
  On Error GoTo ChkErr
    Dim str02AD As String
    Dim strQry As String
    Dim mFlds As New GiGridFlds
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
      '#3899 如果帳號有遮罩字元,所串出的語法也會包含遮罩字元,改為判斷是否全部為字數,如果不是取Tag的值 By Kin 2008/04/25
    If txtAccountNo.Text <> "" Then
        str02AD = "Select B.InvSeqNo From " & GetOwner & "SO002AD B " & _
                "Where B.CustId=" & gCustId & _
                " And B.AccountNo='" & IIf(IsNumeric(txtAccountNo.Text), txtAccountNo.Text, txtAccountNo.Tag) & "' " & _
                " And B.CompCode=" & gCompCode & _
                " And A.InvSeqNo=B.InvSeqNo"
    Else
        str02AD = "Select B.InvSeqNo From " & GetOwner & "SO002AD B " & _
                    "Where B.CustId=" & gCustId & _
                    " And B.AccountNo=' ' " & _
                    " And B.CompCode=" & gCompCode & _
                    " And A.InvSeqNo=B.InvSeqNo"
    End If
    If blnAddNew Then
        strQry = "Select A.* From " & GetOwner & "SO138 A Where 1=0"
    Else
        strQry = "Select A.* From " & GetOwner & "SO138 A Where Exists(" & str02AD & ") And Stopflag=0"
    End If
'    If Not GetRS(rsSO138, strQry, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
    If Not GetRS(rsTmp, strQry, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
    'If rsSO138 Is Nothing Or rsSO138.State = adStateClosed Then
        Set rsSO138 = rsTmp.Clone
    'Else
        'If rsSO138.RecordCount > 0 Then
            'rsTmp.MoveFirst
            'Do While Not rsTmp.EOF
                'rsSO138.MoveFirst
                'rsSO138.Find "InvSeqNo=" & rsTmp("InvSeqNo"), , , 1
                'If rsSO138.EOF Then
                    'rsSO138.AddNew
                    'For i = 0 To rsTmp.Fields.Count - 1
                        'If rsTmp.Fields(i) & "" <> "" Then
                            'rsSO138(rsTmp.Fields(i).Name) = rsTmp.Fields(i).Value
                        'End If
                    'Next i
                    'rsSO138.Update
                'End If
                
                'rsTmp.MoveNext
            'Loop
        'Else
            'Set rsSO138 = rsTmp.Clone
        'End If
    'End If
    Set rsSO138.ActiveConnection = Nothing
   '#3436 新增一個Grid顯示發票資料 By Kin 2007/11/27
   '#3823 增加郵寄地址 By Kin 2008/03/21
    With mFlds
        .Add "InvSeqNo", , , , , "流水編號"
        .Add "ChargeTitle", , , , , "收件人" & Space(10)
        .Add "InvoiceType", , , , , "發票種類"
        .Add "InvNo", , , , , "發票統一編號"
        .Add "InvTitle", , , , , "發票抬頭" & Space(10)
        .Add "ChargeAddress", , , , , "收費地址" & Space(30)
        .Add "MailAddress", , , , , "郵寄地址" & Space(30)
        .Add "InvAddress", , , , , "發票地址" & Space(30)
    End With
'
    ggrData3.AllFields = mFlds
    ggrData3.SetHead
    Set ggrData3.Recordset = rsSO138
    ggrData3.Refresh
    CloseRecordset rsTmp
    OpenSO138 = True
    
    Exit Function
ChkErr:
    CloseRecordset rsTmp
    ErrSub Me.Name, "ChkErr"
End Function

'#3436 將選擇到的發票資料填入SO002AD中 By Kin 2007/11/27
Private Function AddChoiceInv(rs As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    Dim i As Long
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If rsSO138.RecordCount = 0 Then
            Do While Not rs.EOF
                DoEvents
                With rsSO138
                    .AddNew
                    For i = 0 To .Fields.Count - 1
                        If Not IsNull(rs.Fields(.Fields(i).Name)) Then
                            If rs.Fields(.Fields(i).Name) & "" <> "" Then
                                .Fields(i).Value = rs.Fields(.Fields(i).Name).Value & ""
                            End If
                        End If
                    Next i
                End With
                rsSO138.Update
                rs.MoveNext
            Loop
        Else
            Do While Not rs.EOF
                rsSO138.MoveFirst
                rsSO138.Find "InvSeqNo=" & rs("InvSeqNo") & ""
                If rsSO138.EOF Then
                    With rsSO138
                        .AddNew
                        For i = 0 To .Fields.Count - 1
                            If Not IsNull(rs.Fields(.Fields(i).Name)) Then
                                If rs.Fields(.Fields(i).Name) & "" <> "" Then
                                    .Fields(i).Value = rs.Fields(.Fields(i).Name).Value & ""
                                End If
                            End If
                        Next i
                    End With
                    rsSO138.Update
                End If
                rs.MoveNext
            Loop
        End If
    End If
    
    AddChoiceInv = True
  Exit Function
ChkErr:
    ErrSub Me.Name, "ChkErr"
End Function
'#3929 改成複選,所以要讀取SO138整個RecordSet By Kin 2008/05/26
Public Property Let uRS(ByRef vData As Recordset)
    Set rsFrom138 = vData.Clone
End Property

Public Property Let uDelete(ByVal vData As Boolean)
    blnDelete = vData
End Property
