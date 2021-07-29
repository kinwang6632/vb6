VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSO19J0A 
   BorderStyle     =   1  '單線固定
   Caption         =   "CP 帳務資料拋轉作業（for 台灣固網 TFN）[SO19J0A]"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO19J0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   11880
   StartUpPosition =   1  '所屬視窗中央
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   13996
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "帳務項目對應設定"
      TabPicture(0)   =   "SO19J0A.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblEditMode"
      Tab(0).Control(1)=   "ggrData"
      Tab(0).Control(2)=   "cmdSave"
      Tab(0).Control(3)=   "cmdCancel"
      Tab(0).Control(4)=   "cmdAdd"
      Tab(0).Control(5)=   "cmdEdit"
      Tab(0).Control(6)=   "cmdPrint"
      Tab(0).Control(7)=   "FraData"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "資料匯入"
      TabPicture(1)   =   "SO19J0A.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "資料匯出"
      TabPicture(2)   =   "SO19J0A.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab3"
      Tab(2).ControlCount=   1
      Begin VB.Frame FraData 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   98
         Top             =   4320
         Width           =   11415
         Begin VB.TextBox txtTFNHead 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   7440
            MaxLength       =   10
            TabIndex        =   14
            Text            =   "代收"
            Top             =   2280
            Width           =   2160
         End
         Begin VB.TextBox txtCpGroupName 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   8160
            MaxLength       =   50
            TabIndex        =   11
            Top             =   1110
            Width           =   3000
         End
         Begin VB.TextBox txtCpGroupCode 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   7440
            MaxLength       =   10
            TabIndex        =   10
            Top             =   1110
            Width           =   720
         End
         Begin VB.TextBox txtCpCItemName 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   8160
            MaxLength       =   50
            TabIndex        =   8
            Top             =   315
            Width           =   3000
         End
         Begin VB.TextBox txtCpCItemCode 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   7440
            MaxLength       =   10
            TabIndex        =   7
            Top             =   315
            Width           =   720
         End
         Begin VB.TextBox txtCpAdjCItemCode 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   7440
            MaxLength       =   10
            TabIndex        =   9
            Top             =   720
            Width           =   720
         End
         Begin VB.TextBox txtCpItem 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
            Height          =   330
            Left            =   7440
            MaxLength       =   10
            TabIndex        =   12
            Top             =   1515
            Width           =   720
         End
         Begin VB.ComboBox cboCpTax 
            Height          =   300
            ItemData        =   "SO19J0A.frx":0496
            Left            =   7440
            List            =   "SO19J0A.frx":04A3
            TabIndex        =   13
            Top             =   1920
            Width           =   1395
         End
         Begin VB.ComboBox cboCpProperty 
            Height          =   300
            ItemData        =   "SO19J0A.frx":04C3
            Left            =   2160
            List            =   "SO19J0A.frx":04D3
            TabIndex        =   2
            Top             =   315
            Width           =   1400
         End
         Begin prjGiList.GiList gilCItem1 
            Height          =   330
            Left            =   2160
            TabIndex        =   3
            Top             =   720
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FldWidth1       =   640
            FldWidth2       =   2100
            F2Corresponding =   -1  'True
            FilterStop      =   -1  'True
         End
         Begin prjGiList.GiList gilTax1 
            Height          =   330
            Left            =   2160
            TabIndex        =   4
            Top             =   1125
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FldWidth1       =   640
            FldWidth2       =   1000
            F2Corresponding =   -1  'True
            FilterStop      =   -1  'True
         End
         Begin prjGiList.GiList gilCItem2 
            Height          =   330
            Left            =   2160
            TabIndex        =   5
            Top             =   1515
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FldWidth1       =   640
            FldWidth2       =   2100
            F2Corresponding =   -1  'True
            FilterStop      =   -1  'True
         End
         Begin prjGiList.GiList gilTax2 
            Height          =   330
            Left            =   2160
            TabIndex        =   6
            Top             =   1920
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FldWidth1       =   640
            FldWidth2       =   1000
            F2Corresponding =   -1  'True
            FilterStop      =   -1  'True
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "收費項目前導字串"
            Height          =   180
            Left            =   5880
            TabIndex        =   166
            Top             =   2355
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "　　稅率 (借貸 : + )"
            Height          =   180
            Left            =   480
            TabIndex        =   108
            Top             =   1200
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "收費項目 (借貸 : - )"
            Height          =   180
            Left            =   480
            TabIndex        =   107
            Top             =   1590
            Width           =   1485
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "　　稅率 (借貸 : - )"
            Height          =   180
            Left            =   480
            TabIndex        =   106
            Top             =   1995
            Width           =   1485
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "收費項目 (借貸 : + )"
            Height          =   180
            Left            =   480
            TabIndex        =   105
            Top             =   795
            Width           =   1515
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "電信帳單調改項目代碼"
            Height          =   180
            Left            =   5520
            TabIndex        =   104
            Top             =   795
            Width           =   1800
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "電信帳單群組代碼\名稱"
            Height          =   180
            Left            =   5520
            TabIndex        =   103
            Top             =   1185
            Width           =   1845
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "電信列印順序"
            Height          =   180
            Left            =   6240
            TabIndex        =   102
            Top             =   1590
            Width           =   1080
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "電信帳單項目代碼\名稱"
            Height          =   180
            Left            =   5520
            TabIndex        =   101
            Top             =   390
            Width           =   1845
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "電信稅率"
            Height          =   180
            Left            =   6600
            TabIndex        =   100
            Top             =   1995
            Width           =   720
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "電信項目屬性"
            Height          =   180
            Left            =   960
            TabIndex        =   99
            Top             =   385
            Width           =   1080
         End
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "F5. 列印"
         Height          =   375
         Left            =   -72930
         TabIndex        =   16
         Top             =   7320
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "F11.修改"
         Height          =   375
         Left            =   -70470
         TabIndex        =   18
         Top             =   7320
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "F6. 新增"
         Height          =   375
         Left            =   -71700
         TabIndex        =   17
         Top             =   7320
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&X)"
         Height          =   375
         Left            =   -66435
         TabIndex        =   19
         Top             =   7320
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "F2. 存檔"
         Height          =   375
         Left            =   -74160
         TabIndex        =   15
         Top             =   7320
         Width           =   1215
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   7335
         Left            =   -74880
         TabIndex        =   148
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   12938
         _Version        =   393216
         Tabs            =   5
         Tab             =   1
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "CP通信費銷帳資料匯出"
         TabPicture(0)   =   "SO19J0A.frx":0501
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdChooseFile7"
         Tab(0).Control(1)=   "txtLogFile4"
         Tab(0).Control(2)=   "cmdCancel2"
         Tab(0).Control(3)=   "cmdExport1"
         Tab(0).Control(4)=   "cmdSelectPath1"
         Tab(0).Control(5)=   "txtExportPath1"
         Tab(0).Control(6)=   "txtExportFile1"
         Tab(0).Control(7)=   "GiRealDate2"
         Tab(0).Control(8)=   "GiRealDate1"
         Tab(0).Control(9)=   "GiExportDate1"
         Tab(0).Control(10)=   "GiTempDate"
         Tab(0).Control(11)=   "Label49"
         Tab(0).Control(12)=   "lblExportFile1"
         Tab(0).Control(13)=   "Label22"
         Tab(0).Control(14)=   "Label21"
         Tab(0).Control(15)=   "Label20"
         Tab(0).Control(16)=   "Label12"
         Tab(0).Control(17)=   "Label11"
         Tab(0).ControlCount=   18
         TabCaption(1)   =   "CP月租費資料匯出"
         TabPicture(1)   =   "SO19J0A.frx":051D
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label23"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label24"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label25"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label26"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label27"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label28"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label29"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "lblExportFile2"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Label50"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "GiExportDate3"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "GiExportDate2"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "GiRealDate3"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "GiRealDate4"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "chkCanExport1"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "txtExportFile2"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "cmdExport2"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "cmdCancel3"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "txtExportPath2"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "cmdSelectPath2"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "cmdChooseFile8"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "txtLogFile5"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).ControlCount=   21
         TabCaption(2)   =   "CP退租資料匯出"
         TabPicture(2)   =   "SO19J0A.frx":0539
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label31"
         Tab(2).Control(1)=   "Label30"
         Tab(2).Control(2)=   "Label32"
         Tab(2).Control(3)=   "Label33"
         Tab(2).Control(4)=   "Label34"
         Tab(2).Control(5)=   "Label35"
         Tab(2).Control(6)=   "lblExportFile3"
         Tab(2).Control(7)=   "gimReasonCode"
         Tab(2).Control(8)=   "GiExportDate4"
         Tab(2).Control(9)=   "GiExportRefDate1"
         Tab(2).Control(10)=   "GiPrDate1"
         Tab(2).Control(11)=   "GiPrDate2"
         Tab(2).Control(12)=   "gimPrCode1"
         Tab(2).Control(13)=   "txtExportFile3"
         Tab(2).Control(14)=   "cmdExport3"
         Tab(2).Control(15)=   "cmdCancel4"
         Tab(2).Control(16)=   "txtExportPath3"
         Tab(2).Control(17)=   "cmdSelectPath3"
         Tab(2).ControlCount=   18
         TabCaption(3)   =   "CMCP強拆資料匯出"
         TabPicture(3)   =   "SO19J0A.frx":0555
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label36"
         Tab(3).Control(1)=   "Label37"
         Tab(3).Control(2)=   "Label38"
         Tab(3).Control(3)=   "Label39"
         Tab(3).Control(4)=   "Label40"
         Tab(3).Control(5)=   "lblExportFile4"
         Tab(3).Control(6)=   "gimUCCode"
         Tab(3).Control(7)=   "GiExportDate5"
         Tab(3).Control(8)=   "GiPrDate3"
         Tab(3).Control(9)=   "GiPrDate4"
         Tab(3).Control(10)=   "gimPrCode2"
         Tab(3).Control(11)=   "txtExportFile4"
         Tab(3).Control(12)=   "cmdExport4"
         Tab(3).Control(13)=   "cmdCancel5"
         Tab(3).Control(14)=   "txtExportPath4"
         Tab(3).Control(15)=   "cmdSelectPath4"
         Tab(3).ControlCount=   16
         TabCaption(4)   =   "CP通信退費資料匯出"
         TabPicture(4)   =   "SO19J0A.frx":0571
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Label41"
         Tab(4).Control(1)=   "Label42"
         Tab(4).Control(2)=   "Label43"
         Tab(4).Control(3)=   "Label44"
         Tab(4).Control(4)=   "Label45"
         Tab(4).Control(5)=   "lblExportFile5"
         Tab(4).Control(6)=   "Label51"
         Tab(4).Control(7)=   "GiExportDate6"
         Tab(4).Control(8)=   "GiRealDate5"
         Tab(4).Control(9)=   "GiRealDate6"
         Tab(4).Control(10)=   "cmdCancel6"
         Tab(4).Control(11)=   "cmdExport5"
         Tab(4).Control(12)=   "cmdSelectPath5"
         Tab(4).Control(13)=   "txtExportPath5"
         Tab(4).Control(14)=   "txtExportFile5"
         Tab(4).Control(15)=   "txtLogFile6"
         Tab(4).Control(16)=   "cmdChooseFile9"
         Tab(4).ControlCount=   17
         Begin VB.CommandButton cmdChooseFile9 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   89
            Top             =   4560
            Width           =   420
         End
         Begin VB.TextBox txtLogFile6 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   88
            Top             =   4560
            Width           =   5565
         End
         Begin VB.TextBox txtLogFile5 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            TabIndex        =   57
            Top             =   4800
            Width           =   5565
         End
         Begin VB.CommandButton cmdChooseFile8 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9000
            TabIndex        =   58
            Top             =   4800
            Width           =   420
         End
         Begin VB.CommandButton cmdChooseFile7 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   46
            Top             =   4560
            Width           =   420
         End
         Begin VB.TextBox txtLogFile4 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   45
            Top             =   4560
            Width           =   5565
         End
         Begin VB.TextBox txtExportFile5 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   84
            Text            =   "NVW_REPAY_U"
            Top             =   2040
            Width           =   3700
         End
         Begin VB.TextBox txtExportPath5 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   86
            Top             =   3720
            Width           =   5565
         End
         Begin VB.CommandButton cmdSelectPath5 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   87
            Top             =   3720
            Width           =   420
         End
         Begin VB.CommandButton cmdExport5 
            Caption         =   "匯出"
            Height          =   375
            Left            =   -71040
            TabIndex        =   90
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel6 
            Caption         =   "取消"
            Height          =   375
            Left            =   -68565
            TabIndex        =   91
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdSelectPath4 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   79
            Top             =   4580
            Width           =   420
         End
         Begin VB.TextBox txtExportPath4 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   78
            Top             =   4580
            Width           =   5565
         End
         Begin VB.CommandButton cmdCancel5 
            Caption         =   "取消"
            Height          =   375
            Left            =   -68565
            TabIndex        =   81
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdExport4 
            Caption         =   "匯出"
            Height          =   375
            Left            =   -71040
            TabIndex        =   80
            Top             =   5880
            Width           =   1215
         End
         Begin VB.TextBox txtExportFile4 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   76
            Text            =   "NVW_REMOVE_"
            Top             =   3180
            Width           =   3700
         End
         Begin VB.CommandButton cmdSelectPath3 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   69
            Top             =   4680
            Width           =   420
         End
         Begin VB.TextBox txtExportPath3 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   68
            Top             =   4680
            Width           =   5565
         End
         Begin VB.CommandButton cmdCancel4 
            Caption         =   "取消"
            Height          =   375
            Left            =   -68565
            TabIndex        =   71
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdExport3 
            Caption         =   "匯出"
            Height          =   375
            Left            =   -71040
            TabIndex        =   70
            Top             =   5880
            Width           =   1215
         End
         Begin VB.TextBox txtExportFile3 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   66
            Text            =   "NVW_REMOVE_R"
            Top             =   3480
            Width           =   3700
         End
         Begin VB.CommandButton cmdSelectPath2 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9000
            TabIndex        =   56
            Top             =   4200
            Width           =   420
         End
         Begin VB.TextBox txtExportPath2 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            TabIndex        =   55
            Top             =   4200
            Width           =   5565
         End
         Begin VB.CommandButton cmdCancel3 
            Caption         =   "取消"
            Height          =   375
            Left            =   6435
            TabIndex        =   60
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdExport2 
            Caption         =   "匯出"
            Height          =   375
            Left            =   3960
            TabIndex        =   59
            Top             =   5880
            Width           =   1215
         End
         Begin VB.TextBox txtExportFile2 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            TabIndex        =   53
            Text            =   "NVW_PIH_R"
            Top             =   3000
            Width           =   3700
         End
         Begin VB.CheckBox chkCanExport1 
            Height          =   330
            Left            =   3240
            TabIndex        =   52
            Top             =   2400
            Width           =   375
         End
         Begin VB.CommandButton cmdCancel2 
            Caption         =   "取消"
            Height          =   375
            Left            =   -68565
            TabIndex        =   48
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdExport1 
            Caption         =   "匯出"
            Height          =   375
            Left            =   -71040
            TabIndex        =   47
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdSelectPath1 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   44
            Top             =   3720
            Width           =   420
         End
         Begin VB.TextBox txtExportPath1 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   43
            Top             =   3720
            Width           =   5565
         End
         Begin VB.TextBox txtExportFile1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   41
            Text            =   "NVW_PIH_U"
            Top             =   2040
            Width           =   3700
         End
         Begin Gi_Date.GiDate GiRealDate2 
            Height          =   330
            Left            =   -70080
            TabIndex        =   40
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_Date.GiDate GiRealDate1 
            Height          =   330
            Left            =   -71760
            TabIndex        =   39
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_Date.GiDate GiExportDate1 
            Height          =   330
            Left            =   -71760
            TabIndex        =   42
            Top             =   2880
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_Date.GiDate GiRealDate4 
            Height          =   330
            Left            =   4800
            TabIndex        =   50
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_Date.GiDate GiRealDate3 
            Height          =   330
            Left            =   3240
            TabIndex        =   49
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_YM.GiYM GiExportDate2 
            Height          =   330
            Left            =   3240
            TabIndex        =   51
            Top             =   1800
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   582
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
         Begin Gi_Date.GiDate GiExportDate3 
            Height          =   330
            Left            =   3240
            TabIndex        =   54
            Top             =   3600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_Multi.GiMulti gimPrCode1 
            Height          =   330
            Left            =   -73725
            TabIndex        =   61
            Top             =   1080
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   582
            ButtonCaption   =   "拆機類別"
         End
         Begin Gi_Date.GiDate GiPrDate2 
            Height          =   330
            Left            =   -70200
            TabIndex        =   64
            Top             =   2280
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_Date.GiDate GiPrDate1 
            Height          =   330
            Left            =   -71760
            TabIndex        =   63
            Top             =   2280
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_Date.GiDate GiExportRefDate1 
            Height          =   330
            Left            =   -71760
            TabIndex        =   65
            Top             =   2880
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_Date.GiDate GiExportDate4 
            Height          =   330
            Left            =   -71760
            TabIndex        =   67
            Top             =   4080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_Multi.GiMulti gimPrCode2 
            Height          =   330
            Left            =   -73720
            TabIndex        =   72
            Top             =   1080
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   582
            ButtonCaption   =   "拆機類別"
         End
         Begin Gi_Date.GiDate GiPrDate4 
            Height          =   330
            Left            =   -70155
            TabIndex        =   75
            Top             =   2480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_Date.GiDate GiPrDate3 
            Height          =   330
            Left            =   -71760
            TabIndex        =   74
            Top             =   2480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_Date.GiDate GiExportDate5 
            Height          =   330
            Left            =   -71760
            TabIndex        =   77
            Top             =   3880
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_Date.GiDate GiTempDate 
            Height          =   330
            Left            =   -66840
            TabIndex        =   144
            Top             =   5880
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_Date.GiDate GiRealDate6 
            Height          =   330
            Left            =   -70080
            TabIndex        =   83
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_Date.GiDate GiRealDate5 
            Height          =   330
            Left            =   -71760
            TabIndex        =   82
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_Date.GiDate GiExportDate6 
            Height          =   330
            Left            =   -71760
            TabIndex        =   85
            Top             =   2880
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin Gi_Multi.GiMulti gimReasonCode 
            Height          =   375
            Left            =   -73725
            TabIndex        =   62
            Top             =   1680
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   661
            ButtonCaption   =   "停 拆 機 原 因"
         End
         Begin Gi_Multi.GiMulti gimUCCode 
            Height          =   375
            Left            =   -73720
            TabIndex        =   73
            Top             =   1780
            Width           =   7605
            _ExtentX        =   13414
            _ExtentY        =   661
            ButtonCaption   =   "停 拆 機 原 因"
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "問題記錄檔"
            Height          =   180
            Left            =   -72900
            TabIndex        =   160
            Top             =   4635
            Width           =   900
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "問題記錄檔"
            Height          =   180
            Left            =   2100
            TabIndex        =   159
            Top             =   4875
            Width           =   900
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "問題記錄檔"
            Height          =   180
            Left            =   -72960
            TabIndex        =   158
            Top             =   4635
            Width           =   900
         End
         Begin VB.Label lblExportFile5 
            AutoSize        =   -1  'True
            Height          =   180
            Left            =   -71760
            TabIndex        =   154
            Top             =   5040
            Width           =   45
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "實收日期"
            Height          =   180
            Left            =   -72720
            TabIndex        =   153
            Top             =   1275
            Width           =   720
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Left            =   -70440
            TabIndex        =   152
            Top             =   1275
            Width           =   180
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "拋轉檔名"
            Height          =   180
            Left            =   -72720
            TabIndex        =   151
            Top             =   2115
            Width           =   720
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "拋轉日期"
            Height          =   180
            Left            =   -72720
            TabIndex        =   150
            Top             =   2955
            Width           =   720
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "匯出檔案路徑"
            Height          =   180
            Left            =   -73080
            TabIndex        =   149
            Top             =   3795
            Width           =   1080
         End
         Begin VB.Label lblExportFile4 
            AutoSize        =   -1  'True
            Height          =   180
            Left            =   -71760
            TabIndex        =   147
            Top             =   4920
            Width           =   45
         End
         Begin VB.Label lblExportFile3 
            AutoSize        =   -1  'True
            Height          =   180
            Left            =   -71760
            TabIndex        =   146
            Top             =   5160
            Width           =   45
         End
         Begin VB.Label lblExportFile2 
            AutoSize        =   -1  'True
            Height          =   180
            Left            =   3240
            TabIndex        =   145
            Top             =   5280
            Width           =   45
         End
         Begin VB.Label lblExportFile1 
            AutoSize        =   -1  'True
            Height          =   180
            Left            =   -71760
            TabIndex        =   143
            Top             =   5040
            Width           =   45
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "匯出檔案路徑"
            Height          =   180
            Left            =   -73080
            TabIndex        =   142
            Top             =   4655
            Width           =   1080
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "拋轉日期"
            Height          =   180
            Left            =   -72720
            TabIndex        =   141
            Top             =   3955
            Width           =   720
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "拋轉檔名"
            Height          =   180
            Left            =   -72720
            TabIndex        =   140
            Top             =   3255
            Width           =   720
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "拆機完工日(退租日)"
            Height          =   180
            Left            =   -73560
            TabIndex        =   139
            Top             =   2555
            Width           =   1560
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Left            =   -70515
            TabIndex        =   138
            Top             =   2555
            Width           =   180
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "匯出檔案路徑"
            Height          =   180
            Left            =   -73080
            TabIndex        =   137
            Top             =   4755
            Width           =   1080
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "拋轉日期"
            Height          =   180
            Left            =   -72720
            TabIndex        =   136
            Top             =   4155
            Width           =   720
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "拋轉檔名"
            Height          =   180
            Left            =   -72720
            TabIndex        =   135
            Top             =   3555
            Width           =   720
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "帳款拋轉參考日"
            Height          =   180
            Left            =   -73260
            TabIndex        =   134
            Top             =   2955
            Width           =   1260
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "拆機完工日(退租日)"
            Height          =   180
            Left            =   -73560
            TabIndex        =   133
            Top             =   2355
            Width           =   1560
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Left            =   -70560
            TabIndex        =   132
            Top             =   2355
            Width           =   180
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "匯出檔案路徑"
            Height          =   180
            Left            =   1920
            TabIndex        =   131
            Top             =   4275
            Width           =   1080
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "拋轉日期"
            Height          =   180
            Left            =   2280
            TabIndex        =   130
            Top             =   3675
            Width           =   720
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "拋轉檔名"
            Height          =   180
            Left            =   2280
            TabIndex        =   129
            Top             =   3075
            Width           =   720
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "前期預收是否拋轉"
            Height          =   180
            Left            =   1560
            TabIndex        =   128
            Top             =   2475
            Width           =   1440
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "帳款拋轉月份"
            Height          =   180
            Left            =   1920
            TabIndex        =   127
            Top             =   1875
            Width           =   1080
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Left            =   4440
            TabIndex        =   126
            Top             =   1275
            Width           =   180
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "實收日期"
            Height          =   180
            Left            =   2280
            TabIndex        =   125
            Top             =   1275
            Width           =   720
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "匯出檔案路徑"
            Height          =   180
            Left            =   -73080
            TabIndex        =   124
            Top             =   3795
            Width           =   1080
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "拋轉日期"
            Height          =   180
            Left            =   -72720
            TabIndex        =   123
            Top             =   2955
            Width           =   720
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "拋轉檔名"
            Height          =   180
            Left            =   -72720
            TabIndex        =   122
            Top             =   2115
            Width           =   720
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Left            =   -70440
            TabIndex        =   121
            Top             =   1275
            Width           =   180
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "實收日期"
            Height          =   180
            Left            =   -72720
            TabIndex        =   120
            Top             =   1275
            Width           =   720
         End
      End
      Begin prjGiGridR.GiGridR ggrData 
         Height          =   3825
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   6747
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
      Begin TabDlg.SSTab SSTab2 
         Height          =   7335
         Left            =   120
         TabIndex        =   110
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   12938
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         TabCaption(0)   =   "CP通信費資料匯入"
         TabPicture(0)   =   "SO19J0A.frx":058D
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label13"
         Tab(0).Control(1)=   "Label14"
         Tab(0).Control(2)=   "Label15"
         Tab(0).Control(3)=   "Label52"
         Tab(0).Control(4)=   "Label53"
         Tab(0).Control(5)=   "txtImportFile1"
         Tab(0).Control(6)=   "txtLogFile1"
         Tab(0).Control(7)=   "cmdCancel1"
         Tab(0).Control(8)=   "cmdInport1"
         Tab(0).Control(9)=   "cmdChooseFile1"
         Tab(0).Control(10)=   "cmdChooseFile2"
         Tab(0).Control(11)=   "Frame1"
         Tab(0).Control(12)=   "ComdPath"
         Tab(0).Control(13)=   "PB1"
         Tab(0).Control(14)=   "txtShouldDay"
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "CP通話明細資料匯入"
         TabPicture(1)   =   "SO19J0A.frx":05A9
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label19"
         Tab(1).Control(1)=   "Label48"
         Tab(1).Control(2)=   "PB2"
         Tab(1).Control(3)=   "txtImportFile5"
         Tab(1).Control(4)=   "cmdChooseFile5"
         Tab(1).Control(5)=   "Frame2"
         Tab(1).Control(6)=   "cmdInport3"
         Tab(1).Control(7)=   "cmdCancel8"
         Tab(1).Control(8)=   "cmdChooseFile6"
         Tab(1).Control(9)=   "txtLogFile3"
         Tab(1).ControlCount=   10
         TabCaption(2)   =   "基本資料核對匯入"
         TabPicture(2)   =   "SO19J0A.frx":05C5
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Label46"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Label47"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "PB3"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "txtImportFile2"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "txtLogFile2"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "cmdChooseFile3"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "cmdChooseFile4"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "cmdInport2"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "cmdCancel7"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).ControlCount=   9
         Begin VB.TextBox txtShouldDay 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            MaxLength       =   2
            TabIndex        =   22
            Text            =   "1"
            Top             =   2400
            Width           =   400
         End
         Begin MSComctlLib.ProgressBar PB1 
            Height          =   375
            Left            =   -73560
            TabIndex        =   161
            Top             =   4920
            Visible         =   0   'False
            Width           =   7980
            _ExtentX        =   14076
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.TextBox txtLogFile3 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   35
            Top             =   4200
            Width           =   5565
         End
         Begin VB.CommandButton cmdChooseFile6 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   36
            Top             =   4200
            Width           =   420
         End
         Begin VB.CommandButton cmdCancel7 
            Caption         =   "取消"
            Height          =   375
            Left            =   6435
            TabIndex        =   97
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdInport2 
            Caption         =   "匯入"
            Height          =   375
            Left            =   3960
            TabIndex        =   96
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdChooseFile4 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8880
            TabIndex        =   95
            Top             =   3360
            Width           =   420
         End
         Begin VB.CommandButton cmdChooseFile3 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8880
            TabIndex        =   93
            Top             =   2040
            Width           =   420
         End
         Begin VB.TextBox txtLogFile2 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3120
            TabIndex        =   94
            Top             =   3360
            Width           =   5565
         End
         Begin VB.TextBox txtImportFile2 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3120
            TabIndex        =   92
            Top             =   2040
            Width           =   5565
         End
         Begin MSComDlg.CommonDialog ComdPath 
            Left            =   -74280
            Top             =   5880
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton cmdCancel8 
            Caption         =   "取消"
            Height          =   375
            Left            =   -68565
            TabIndex        =   38
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdInport3 
            Caption         =   "匯入"
            Height          =   375
            Left            =   -71040
            TabIndex        =   37
            Top             =   5880
            Width           =   1215
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  '沒有框線
            Height          =   1140
            Left            =   -73200
            TabIndex        =   116
            Top             =   2160
            Width           =   7815
            Begin VB.TextBox txtImportFile4 
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3720
               TabIndex        =   32
               Text            =   "CNS_RATING_M"
               Top             =   675
               Width           =   3700
            End
            Begin VB.OptionButton optDetailDaily 
               Caption         =   "每日"
               Height          =   180
               Left            =   1440
               TabIndex        =   29
               Top             =   75
               Value           =   -1  'True
               Width           =   800
            End
            Begin VB.TextBox txtImportFile3 
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3720
               TabIndex        =   30
               Text            =   "CNS_RATING_D"
               Top             =   0
               Width           =   3700
            End
            Begin VB.OptionButton optDetailMonthly 
               Caption         =   "每月"
               Height          =   180
               Left            =   1440
               TabIndex        =   31
               Top             =   750
               Width           =   800
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "檢核匯入檔名"
               Height          =   180
               Left            =   2400
               TabIndex        =   119
               Top             =   750
               Width           =   1080
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "檢核匯入檔名"
               Height          =   180
               Left            =   2400
               TabIndex        =   118
               Top             =   75
               Width           =   1080
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "通話明細種類"
               Height          =   180
               Left            =   0
               TabIndex        =   117
               Top             =   75
               Width           =   1080
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Left            =   -71760
            TabIndex        =   115
            Top             =   1680
            Width           =   2775
            Begin VB.OptionButton optEmpNo 
               Caption         =   "裝機地址"
               Height          =   250
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Width           =   1035
            End
            Begin VB.OptionButton optEmpName 
               Caption         =   "收費地址"
               Height          =   250
               Left            =   1440
               TabIndex        =   21
               Top             =   0
               Value           =   -1  'True
               Width           =   1035
            End
         End
         Begin VB.CommandButton cmdChooseFile5 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   34
            Top             =   3600
            Width           =   420
         End
         Begin VB.TextBox txtImportFile5 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   33
            Top             =   3600
            Width           =   5565
         End
         Begin VB.CommandButton cmdChooseFile2 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   26
            Top             =   3840
            Width           =   420
         End
         Begin VB.CommandButton cmdChooseFile1 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   24
            Top             =   3120
            Width           =   420
         End
         Begin VB.CommandButton cmdInport1 
            Caption         =   "匯入"
            Height          =   375
            Left            =   -71040
            TabIndex        =   27
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel1 
            Caption         =   "取消"
            Height          =   375
            Left            =   -68565
            TabIndex        =   28
            Top             =   5880
            Width           =   1215
         End
         Begin VB.TextBox txtLogFile1 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   25
            Top             =   3840
            Width           =   5565
         End
         Begin VB.TextBox txtImportFile1 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   23
            Top             =   3120
            Width           =   5565
         End
         Begin MSComctlLib.ProgressBar PB2 
            Height          =   375
            Left            =   -73200
            TabIndex        =   162
            Top             =   5040
            Visible         =   0   'False
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin MSComctlLib.ProgressBar PB3 
            Height          =   375
            Left            =   1740
            TabIndex        =   163
            Top             =   4680
            Visible         =   0   'False
            Width           =   7560
            _ExtentX        =   13335
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "(應收日須介於 1～31 之間)"
            Height          =   180
            Left            =   -71160
            TabIndex        =   165
            Top             =   2475
            Width           =   2100
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "應收日"
            Height          =   180
            Left            =   -72600
            TabIndex        =   164
            Top             =   2475
            Width           =   540
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "問題記錄檔"
            Height          =   180
            Left            =   -72960
            TabIndex        =   157
            Top             =   4275
            Width           =   900
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "問題記錄檔"
            Height          =   180
            Left            =   1920
            TabIndex        =   156
            Top             =   3435
            Width           =   900
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "檢核匯入檔名"
            Height          =   180
            Left            =   1740
            TabIndex        =   155
            Top             =   2115
            Width           =   1080
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "匯入檔案"
            Height          =   180
            Left            =   -72840
            TabIndex        =   114
            Top             =   3720
            Width           =   720
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "問題記錄檔"
            Height          =   180
            Left            =   -72960
            TabIndex        =   113
            Top             =   3915
            Width           =   900
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "匯入檔案"
            Height          =   180
            Left            =   -72780
            TabIndex        =   112
            Top             =   3195
            Width           =   720
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "收費單之地址依據"
            Height          =   180
            Left            =   -73560
            TabIndex        =   111
            Top             =   1710
            Width           =   1440
         End
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  '單線固定
         Caption         =   "顯示"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   -64815
         TabIndex        =   109
         Top             =   7380
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmSO19J0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents rsSo165 As ADODB.Recordset
Attribute rsSo165.VB_VarHelpID = -1
Private WithEvents rsSo168 As ADODB.Recordset
Attribute rsSo168.VB_VarHelpID = -1
Private WithEvents rsSO033 As ADODB.Recordset
Attribute rsSO033.VB_VarHelpID = -1
Private WithEvents rsSO033Cp As ADODB.Recordset
Attribute rsSO033Cp.VB_VarHelpID = -1
Private WithEvents rsTmp012 As ADODB.Recordset
Attribute rsTmp012.VB_VarHelpID = -1
Private FSO As New FileSystemObject
Private strmTextFile1, strmTextFile2 As TextStream
Private lngEditMode As giEditModeEnu

Private Sub cboCpProperty_Click()
On Error GoTo ChkErr
  gilCItem1.Enabled = (cboCpProperty.ListIndex < 2) Or (cboCpProperty.ListIndex = 3)
  gilTax1.Enabled = gilCItem1.Enabled
  Label4.Enabled = gilCItem1.Enabled
  Label1.Enabled = gilCItem1.Enabled
  If Not gilCItem1.Enabled Then
    gilCItem1.Clear
    gilTax1.Clear
  End If
ChkErr:
  Call ErrSub(Me.Name, "cboCpProperty_Click")
End Sub

Private Sub cmdAdd_Click()
On Error GoTo ChkErr
  '新增
  Call ChangeMode(giEditModeInsert)
  Call NewRcd
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdAdd_Click")
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ChkErr
  '若是瀏覽模式，則關閉程式
  '否則回到瀏覽模式
  If lngEditMode = giEditModeView Then
    Unload Me
  Else
    Call ChangeMode(giEditModeView)
    Call RcdToScr
  End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdCancel_Click")
End Sub

Private Sub cmdCancel1_Click()
  Call cmdCancel_Click
End Sub

Private Sub cmdCancel2_Click()
  Call cmdCancel_Click
End Sub

Private Sub cmdCancel3_Click()
  Call cmdCancel_Click
End Sub

Private Sub cmdCancel4_Click()
  Call cmdCancel_Click
End Sub

Private Sub cmdCancel5_Click()
  Call cmdCancel_Click
End Sub

Private Sub cmdCancel6_Click()
  Call cmdCancel_Click
End Sub

Private Sub cmdCancel7_Click()
  Call cmdCancel_Click
End Sub

Private Sub cmdCancel8_Click()
  Call cmdCancel_Click
End Sub

Private Sub cmdChooseFile1_Click()
  Call SelectImportFile(txtImportFile1, txtLogFile1)
End Sub

Private Sub SelectImportFile(ByRef ImportFile As TextBox, ByRef LogFile As TextBox)
On Error GoTo ChkErr
  '選擇匯入的文字檔
  With ComdPath
    .FileName = ImportFile.Text
    .Filter = "文字檔 (*.txt;*.dat)|*.txt;*.dat|所有檔案 (*.*)|*.*"
    .Action = 1
    ImportFile.Text = .FileName
  End With
  
  '自動根據文字檔的檔名設定對應的 LOG 檔檔名
  Call SetLogFileName(ImportFile.Text, LogFile)
  
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "SelectImportFile")
End Sub

Private Sub cmdChooseFile2_Click()
  Call SelectImportLog(txtLogFile1)
End Sub

Private Sub SelectImportLog(ByRef LogFile As TextBox)
On Error GoTo ChkErr
  '選擇 LOG 檔
  With ComdPath
    .FileName = LogFile.Text
    .Filter = "文字檔 (*.txt;*.dat)|*.txt;*.dat|所有檔案 (*.*)|*.*"
    .Action = 1
    LogFile.Text = .FileName
  End With
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "SelectImportLog")
End Sub

Private Sub cmdChooseFile3_Click()
  Call SelectImportFile(txtImportFile2, txtLogFile2)
End Sub

Private Sub cmdChooseFile4_Click()
  Call SelectImportLog(txtLogFile2)
End Sub

Private Sub cmdChooseFile5_Click()
  Call SelectImportFile2(txtImportFile5, IIf(optDetailDaily.Value, txtImportFile3.Text, txtImportFile4.Text), txtLogFile3)
End Sub
Private Sub cmdChooseFile6_Click()
  Call SelectImportLog(txtLogFile3)
End Sub

Private Sub cmdChooseFile7_Click()
  Call SelectImportLog(txtLogFile4)
End Sub

Private Sub cmdChooseFile8_Click()
  Call SelectImportLog(txtLogFile5)
End Sub

Private Sub cmdChooseFile9_Click()
  Call SelectImportLog(txtLogFile6)
End Sub

Private Sub cmdEdit_Click()
On Error GoTo ChkErr
  ChangeMode (giEditModeEdit)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdEdit_Click")
End Sub

Private Sub cmdExport1_Click()
On Error GoTo ChkErr
  'CP通信費銷帳資料匯出
  If GiRealDate1.GetValue = "" Then
    MsgBox "請先指定實收日期(起)。", vbExclamation, "警告"
    GiRealDate1.SetFocus
  ElseIf GiRealDate2.GetValue = "" Then
    MsgBox "請先指定實收日期(迄)。", vbExclamation, "警告"
    GiRealDate2.SetFocus
  ElseIf GiRealDate2.GetValue < GiRealDate1.GetValue Then
    MsgBox "實收日期(迄)必須大於等於實收日期(起)。", vbExclamation, "警告"
    GiRealDate2.SetFocus
  ElseIf txtExportFile1.Text = "" Then
    MsgBox "請先指定拋轉檔名。", vbExclamation, "警告"
    txtExportFile1.SetFocus
  ElseIf GiExportDate1.GetValue = "" Then
    MsgBox "請先指定拋轉日期。", vbExclamation, "警告"
    GiExportDate1.SetFocus
  ElseIf txtExportPath1.Text = "" Then
    MsgBox "請先指定匯出檔案路徑。", vbExclamation, "警告"
    txtExportPath1.SetFocus
  ElseIf txtLogFile4.Text = "" Then
    MsgBox "請先指定問題記錄檔。", vbExclamation, "警告"
    txtLogFile4.SetFocus
  ElseIf Replace(txtExportPath1.Text & "\" & lblExportFile1.Caption, "\\", "\") = txtLogFile4.Text Then
    MsgBox "匯入檔案和問題記錄檔不可以一樣。", vbExclamation, "警告"
    txtLogFile4.SetFocus
  Else
    If FSO.FileExists(txtLogFile4.Text) Then
      FSO.DeleteFile (txtLogFile4.Text)
    End If
    Set strmTextFile2 = FSO.OpenTextFile(txtLogFile4.Text, ForAppending, True)
    strmTextFile2.Write vbCrLf & "################################################################################"
    strmTextFile2.Write vbCrLf & "CP通信費銷帳資料匯出　資料重複匯出記錄　記錄日期：" & Now & vbCrLf & vbCrLf
    
    Set strmTextFile1 = FSO.CreateTextFile(txtExportPath1.Text & "\" & lblExportFile1.Caption, True)
    Call ExportData1
    Set strmTextFile1 = Nothing
    
    Set strmTextFile2 = Nothing
  End If
ChkErr:
  Call ErrSub(Me.Name, "cmdExport1_Click")
End Sub

Private Sub cmdExport2_Click()
On Error GoTo ChkErr
  'CP月租費資料匯出
  If GiRealDate3.GetValue = "" Then
    MsgBox "請先指定實收日期(起)。", vbExclamation, "警告"
    GiRealDate3.SetFocus
  ElseIf GiRealDate4.GetValue = "" Then
    MsgBox "請先指定實收日期(迄)。", vbExclamation, "警告"
    GiRealDate4.SetFocus
  ElseIf GiRealDate4.GetValue < GiRealDate3.GetValue Then
    MsgBox "實收日期(迄)必須大於等於實收日期(起)。", vbExclamation, "警告"
    GiRealDate4.SetFocus
  ElseIf GiExportDate2.GetValue = "" Then
    MsgBox "請先指定帳款拋轉月份。", vbExclamation, "警告"
    GiExportDate2.SetFocus
  ElseIf txtExportFile2.Text = "" Then
    MsgBox "請先指定拋轉檔名。", vbExclamation, "警告"
    txtExportFile2.SetFocus
  ElseIf GiExportDate3.GetValue = "" Then
    MsgBox "請先指定拋轉日期。", vbExclamation, "警告"
    GiExportDate3.SetFocus
  ElseIf txtExportPath2.Text = "" Then
    MsgBox "請先指定匯出檔案路徑。", vbExclamation, "警告"
    txtExportPath2.SetFocus
  ElseIf txtLogFile5.Text = "" Then
    MsgBox "請先指定問題記錄檔。", vbExclamation, "警告"
    txtLogFile5.SetFocus
  ElseIf Replace(txtExportPath2.Text & "\" & lblExportFile2.Caption, "\\", "\") = txtLogFile5.Text Then
    MsgBox "匯入檔案和問題記錄檔不可以一樣。", vbExclamation, "警告"
    txtLogFile5.SetFocus
  Else
    If FSO.FileExists(txtLogFile5.Text) Then
      FSO.DeleteFile (txtLogFile5.Text)
    End If
    Set strmTextFile2 = FSO.OpenTextFile(txtLogFile5.Text, ForAppending, True)
    strmTextFile2.Write vbCrLf & "################################################################################"
    strmTextFile2.Write vbCrLf & "CP月租費資料匯出　資料重複匯出記錄　記錄日期：" & Now & vbCrLf & vbCrLf
    
    Set strmTextFile1 = FSO.CreateTextFile(txtExportPath2.Text & "\" & lblExportFile2.Caption, True)
    Call ExportData2
    Set strmTextFile1 = Nothing
    
    Set strmTextFile2 = Nothing
  End If
ChkErr:
  Call ErrSub(Me.Name, "cmdExport2_Click")
End Sub

Private Sub cmdExport3_Click()
On Error GoTo ChkErr
  'CP退租資料匯出
  '如果拆機類別沒有選取，則預設為全選
  If gimPrCode1.GetQueryCode = "" Then
    gimPrCode1.SelectAll
  End If
  If GiPrDate1.GetValue = "" Then
    MsgBox "請先指定拆機完工日(退租日)(起)。", vbExclamation, "警告"
    GiPrDate1.SetFocus
  ElseIf GiPrDate2.GetValue = "" Then
    MsgBox "請先指定拆機完工日(退租日)(迄)。", vbExclamation, "警告"
    GiPrDate2.SetFocus
  ElseIf GiPrDate2.GetValue < GiPrDate1.GetValue Then
    MsgBox "拆機完工日(退租日)(迄)必須大於等於拆機完工日(退租日)(起)。", vbExclamation, "警告"
    GiPrDate2.SetFocus
  ElseIf GiExportRefDate1.GetValue = "" Then
    MsgBox "請先指定帳款拋轉參考日。", vbExclamation, "警告"
    GiExportRefDate1.SetFocus
  ElseIf txtExportFile3.Text = "" Then
    MsgBox "請先指定拋轉檔名。", vbExclamation, "警告"
    txtExportFile3.SetFocus
  ElseIf GiExportDate4.GetValue = "" Then
    MsgBox "請先指定拋轉日期。", vbExclamation, "警告"
    GiExportDate4.SetFocus
  ElseIf txtExportPath3 = "" Then
    MsgBox "請先指定匯出檔案路徑。", vbExclamation, "警告"
    txtExportPath3.SetFocus
  ElseIf gimPrCode1.GetQueryCode = "" Then
    MsgBox "請先指定拆機類別。", vbExclamation, "警告"
    gimPrCode1.SetFocus
  ElseIf gimReasonCode.GetQueryCode = "" Then
    MsgBox "請先指定停拆機原因。", vbExclamation, "警告"
    gimReasonCode.SetFocus
  Else
    Set strmTextFile1 = FSO.CreateTextFile(txtExportPath3.Text & "\" & lblExportFile3.Caption, True)
    Call ExportData3
    Set strmTextFile1 = Nothing
  End If
ChkErr:
  Call ErrSub(Me.Name, "cmdExport3_Click")
End Sub

Private Sub cmdExport4_Click()
On Error GoTo ChkErr
  'CMCP強拆資料匯出
  '如果拆機類別沒有選取，則預設為全選
  If gimPrCode2.GetQueryCode = "" Then
    gimPrCode2.SelectAll
  End If
  If GiPrDate3.GetValue = "" Then
    MsgBox "請先指定拆機完工日(退租日)(起)。", vbExclamation, "警告"
    GiPrDate3.SetFocus
  ElseIf GiPrDate4.GetValue = "" Then
    MsgBox "請先指定拆機完工日(退租日)(迄)。", vbExclamation, "警告"
    GiPrDate4.SetFocus
  ElseIf GiPrDate4.GetValue < GiPrDate3.GetValue Then
    MsgBox "拆機完工日(退租日)(迄)必須大於等於拆機完工日(退租日)(起)。", vbExclamation, "警告"
    GiPrDate4.SetFocus
  ElseIf txtExportFile4.Text = "" Then
    MsgBox "請先指定拋轉檔名。", vbExclamation, "警告"
    txtExportFile4.SetFocus
  ElseIf GiExportDate5.GetValue = "" Then
    MsgBox "請先指定拋轉日期。", vbExclamation, "警告"
    GiExportDate5.SetFocus
  ElseIf txtExportPath4 = "" Then
    MsgBox "請先指定匯出檔案路徑。", vbExclamation, "警告"
    txtExportPath4.SetFocus
  ElseIf gimPrCode2.GetQueryCode = "" Then
    MsgBox "請先指定拆機類別。", vbExclamation, "警告"
    gimPrCode2.SetFocus
  ElseIf gimUCCode.GetQueryCode = "" Then
    MsgBox "請先指定停拆機原因。", vbExclamation, "警告"
    gimUCCode.SetFocus
  Else
    Set strmTextFile1 = FSO.CreateTextFile(txtExportPath4.Text & "\" & lblExportFile4.Caption, True)
    Call ExportData4
    Set strmTextFile1 = Nothing
  End If
ChkErr:
  Call ErrSub(Me.Name, "cmdExport4_Click")
End Sub

Private Sub cmdExport5_Click()
On Error GoTo ChkErr
  'CP通信退費資料匯出
  If GiRealDate5.GetValue = "" Then
    MsgBox "請先指定實收日期(起)。", vbExclamation, "警告"
    GiRealDate5.SetFocus
  ElseIf GiRealDate6.GetValue = "" Then
    MsgBox "請先指定實收日期(迄)。", vbExclamation, "警告"
    GiRealDate6.SetFocus
  ElseIf GiRealDate6.GetValue < GiRealDate5.GetValue Then
    MsgBox "實收日期(迄)必須大於等於實收日期(起)。", vbExclamation, "警告"
    GiRealDate6.SetFocus
  ElseIf txtExportFile5.Text = "" Then
    MsgBox "請先指定拋轉檔名。", vbExclamation, "警告"
    txtExportFile5.SetFocus
  ElseIf GiExportDate6.GetValue = "" Then
    MsgBox "請先指定拋轉日期。", vbExclamation, "警告"
    GiExportDate6.SetFocus
  ElseIf txtExportPath5 = "" Then
    MsgBox "請先指定匯出檔案路徑。", vbExclamation, "警告"
    txtExportPath5.SetFocus
  ElseIf txtLogFile6.Text = "" Then
    MsgBox "請先指定問題記錄檔。", vbExclamation, "警告"
    txtLogFile6.SetFocus
  ElseIf Replace(txtExportPath5.Text & "\" & lblExportFile5.Caption, "\\", "\") = txtLogFile6.Text Then
    MsgBox "匯入檔案和問題記錄檔不可以一樣。", vbExclamation, "警告"
    txtLogFile6.SetFocus
  Else
    If FSO.FileExists(txtLogFile6.Text) Then
      FSO.DeleteFile (txtLogFile6.Text)
    End If
    Set strmTextFile2 = FSO.OpenTextFile(txtLogFile6.Text, ForAppending, True)
    strmTextFile2.Write vbCrLf & "################################################################################"
    strmTextFile2.Write vbCrLf & "CP通信退費資料匯出　資料重複匯出記錄　記錄日期：" & Now & vbCrLf & vbCrLf
    
    Set strmTextFile1 = FSO.CreateTextFile(txtExportPath5.Text & "\" & lblExportFile5.Caption, True)
    Call ExportData5
    Set strmTextFile1 = Nothing
    
    Set strmTextFile2 = Nothing
  End If
ChkErr:
  Call ErrSub(Me.Name, "cmdExport5_Click")
End Sub

Private Sub cmdInport1_Click()
On Error GoTo ChkErr
  'CP通信費資料匯入
  If txtImportFile1.Text = "" Then
    MsgBox "請先指定匯入檔案。", vbExclamation, "警告"
    txtImportFile1.SetFocus
  ElseIf txtLogFile1.Text = "" Then
    MsgBox "請先指定問題記錄檔。", vbExclamation, "警告"
    txtLogFile1.SetFocus
  ElseIf Not FSO.FileExists(txtImportFile1.Text) Then
    MsgBox "要匯入的檔案不存在。", vbExclamation, "警告"
    txtImportFile1.SetFocus
  ElseIf txtImportFile1.Text = txtLogFile1.Text Then
    MsgBox "匯入檔案和問題記錄檔不可以一樣。", vbExclamation, "警告"
    txtImportFile1.SetFocus
  Else
  
    '開一個 SO033 及 SO033CP 的 Client Dataset
    Set rsSO033 = New ADODB.Recordset
    If Not GetRS(rsSO033, "Select * From " & GetOwner & "So033 Where 1=0", gcnGi, , adOpenKeyset, adLockOptimistic, adUseClient) Then Exit Sub
    Set rsSO033Cp = New ADODB.Recordset
    If Not GetRS(rsSO033Cp, "Select * From " & GetOwner & "So033Cp Where 1=0", gcnGi, , adOpenKeyset, adLockOptimistic, adUseClient) Then Exit Sub
    '新建一個 LOG 檔
    If FSO.FileExists(txtLogFile1.Text) Then
      FSO.DeleteFile (txtLogFile1.Text)
    End If
    
    Set strmTextFile1 = FSO.OpenTextFile(txtLogFile1.Text, ForAppending, True)
    strmTextFile1.Write vbCrLf & "################################################################################"
    strmTextFile1.Write vbCrLf & "CP通信費資料匯入　錯誤記錄　記錄日期：" & Now & vbCrLf & vbCrLf
    
    Call ImportData1
    
    Set strmTextFile1 = Nothing
    Set rsSO033 = Nothing
    Set rsSO033Cp = Nothing
  End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdInport1_Click")
End Sub
Private Sub ImportData1()
On Error GoTo ChkErr
  Dim blnInMaster  As Boolean, blnInDetail  As Boolean, blnGoOn As Boolean
  Dim blnHasHead  As Boolean, blnHasTail As Boolean, blnHasMaster  As Boolean, blnHasDetail  As Boolean, blnHasSum As Boolean
  Dim lngHeadErrorCount As Long
  Dim lngTailErrorCount As Long
  Dim lngDetailErrorCount As Long
  Dim lngSumErrorCount As Long
  Dim lngMasterErrorCount As Long
  Dim lngMasterCount As Long, lngDetailCount As Long, lngSumCount As Long, lngTempLong As Long, lngLineNumber As Long, lngBillItem As Long
  Dim dblShouldAmtTotal As Double, dblShouldAmtWithTax As Double, dblShouldAmtWithoutTax As Double, dblShouldAmtZeroTax As Double
  Dim dblShouldAmt2 As Double, dblShouldAmtNoTax2 As Double, dblTax2 As Double
  Dim strMasterShouldCount As String, strDetailShouldCount As String, strSumShouldCount As String, UserDateTime As String
  Dim strTextData As String, strString1 As String
  Dim strTotalDataCount As String, strMoneyTotal As String, strUCCode As String, strUCName As String
  Dim strCitemCode As String, strTempDate As String, strBillNo As String
  Dim strYear As String, strMonth As String, strString2 As String
  Dim strCustid, strCpArea, strCpTel, strTFNBillNo, strTfnServiceId, str_M_48_55, str_M_56_63 As String
  Dim str_M_64_71, str_M_72_79, str_M_86_135, strClassCode1 As String
  Dim strCustId2, strShouldDate As String
  Dim strShouldAmtNoTax, strTax As String
  Dim strCpShouldYm As String
  Dim strCpCItemCode, strCpCItemName, strCpTax, strSign, strShouldAmt, strNote As String
  Dim strShouldAmt3, strShouldAmtNoTax3, strTax3, str_S_87_92
  Dim str_S_103_152 As String
  Dim strCpAdjCItemCode As String
  Dim vDataArray As Variant, vElement As Variant
  Dim rsTemp As New ADODB.Recordset
  Dim rsTemp2 As New ADODB.Recordset
  Screen.MousePointer = vbHourglass
  Me.Enabled = False
  
  '若匯入失敗，則需整批重匯，已匯入 SO033 的資料需 RollBack 回來
  gcnGi.BeginTrans
  
  '先依 SO041.SysTime 判斷要取主機或客戶端電腦的時間當作匯入的時間
  Set rsTemp = gcnGi.Execute("Select SysTime From " & GetOwner & "So041")
  If rsTemp("SysTime") = 1 Then
    Set rsTemp = Nothing
    Set rsTemp = gcnGi.Execute("Select SysDate As UserDateTime From Dual")
    UserDateTime = rsTemp("UserDateTime")
  Else
    UserDateTime = Now
  End If
  Set rsTemp = Nothing
  
  Set rsTemp = gcnGi.Execute("Select UcCode From " & GetOwner & "So044")
  strUCCode = rsTemp("UcCode") & ""
  Set rsTemp = Nothing
  
  Set rsTemp = gcnGi.Execute("Select Description From " & GetOwner & "Cd013 Where CodeNo=" & strUCCode)
  strUCName = rsTemp("Description")
  Set rsTemp = Nothing
  
  lngMasterErrorCount = 0
  lngDetailErrorCount = 0
  lngSumErrorCount = 0
  lngHeadErrorCount = 0
  lngTailErrorCount = 0
  lngMasterCount = 0
  lngDetailCount = 0
  lngSumCount = 0
  dblShouldAmtTotal = 0
  dblShouldAmtWithTax = 0
  dblShouldAmtWithoutTax = 0
  dblShouldAmtZeroTax = 0
  dblShouldAmt2 = 0
  dblShouldAmtNoTax2 = 0
  dblTax2 = 0
  blnHasHead = False
  blnHasTail = False
  blnHasMaster = False
  blnHasDetail = False
  blnInMaster = False
  blnInDetail = False
  
  '讀入文字檔，並分段
  strTextData = FSO.OpenTextFile(txtImportFile1).ReadAll
  vDataArray = Split(strTextData, vbCrLf)
  
  'Progress Bar
  PB1.Visible = True
  PB1.Value = PB1.Min
  On Error Resume Next
  PB1.Max = UBound(vDataArray)
  
  '開始一行一行剖析
  lngLineNumber = 1
  For Each vElement In vDataArray
    PB1.Value = lngLineNumber - 1
    Select Case Left(vElement, 1)
      Case "H"
        '檔頭
        If blnHasMaster Or blnHasDetail Or blnHasSum Or blnHasTail Then
          Call LogError1(lngHeadErrorCount, "表頭必須在文字檔的第一行", "", lngLineNumber, vElement)
        ElseIf Len(vElement) <> 7 Then
          Call LogError1(lngHeadErrorCount, "表頭長度共需 7 碼", "", lngLineNumber, vElement)
        ElseIf Not IsNumeric(Right(vElement, 6)) Then
          Call LogError1(lngHeadErrorCount, "表頭第 2-7 碼必須為 YYYYMM 格式", "", lngLineNumber, vElement)
        Else
          strYear = Mid(vElement, 2, 4)
          strMonth = Right(vElement, 2)
          If Not IsDate(strYear & "/" & strMonth & "/01") Then
            Call LogError1(lngHeadErrorCount, "表頭年月錯誤", "", lngLineNumber, vElement)
          Else
            blnHasHead = True
          End If
        End If
      Case "M"
        lngBillItem = 0
        'Master 項目
        If Not blnHasHead Then
          Call LogError1(lngMasterErrorCount, "資料不可以出現在表頭之前", "", lngLineNumber, vElement)
        ElseIf blnHasTail Then
          Call LogError1(lngMasterErrorCount, "資料不可以出現在表尾之後", "", lngLineNumber, vElement)
        'ElseIf blnInMaster Then
        '  Call LogError1(lngMasterErrorCount, "前一個 Master 資料必須要有對應的 Detail 項目", "",lngLineNumber,vElement)
        Else
          lngMasterCount = lngMasterCount + 1
          blnInMaster = True
          blnInDetail = False
          blnHasMaster = True
          
          '依固定寬度取得各欄位的值
          strCustid = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 2, 8), vbUnicode))
          strCustId2 = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 2, 10), vbUnicode))
          strCustId2 = Right("0000000" & strCustId2, 7)
          
          '區碼要加入 '-' 符號
          strCpArea = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 12, 3), vbUnicode)) & "-"
          strCpTel = StrConv(MidB(StrConv(vElement, vbFromUnicode), 15, 10), vbUnicode)
          
          strTfnServiceId = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 25, 10), vbUnicode))
          strTFNBillNo = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 35, 13), vbUnicode))
          str_M_48_55 = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 48, 8), vbUnicode))
          str_M_56_63 = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 56, 8), vbUnicode))
          str_M_64_71 = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 64, 8), vbUnicode))
          str_M_72_79 = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 72, 8), vbUnicode))
          strShouldDate = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 80, 6), vbUnicode))
          str_M_86_135 = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 86, 50), vbUnicode))
          
          '用 strBillNo 來判斷不同的小計項目
          strBillNo = ""
          
          blnGoOn = True
          
          '底下進行各式各樣的檢核
          Set rsTemp = gcnGi.Execute("Select CustId, ClassCode1 From " & GetOwner & "So001 Where CustId=" & strCustid)
          If rsTemp.EOF Then
            blnGoOn = False
            Call LogError1(lngMasterErrorCount, "客編不存在：" & strCustid, "", lngLineNumber, vElement)
          Else
            strClassCode1 = rsTemp("ClassCode1") & ""
          End If
          Set rsTemp = Nothing
          
'          If blnGoOn Then
            strString1 = Trim(strCpArea & strCpTel)
'            Set rsTemp = gcnGi.Execute("Select CustId From " & GetOwner & "So004 Where FaciSNo='" & strString1 & "'")
'            If rsTemp.EOF Then
'              blnGoOn = False
'              Call LogError1(lngMasterErrorCount, "客編：" & strCustId & "　門號不存在：" & strString1, "", lngLineNumber, vElement)
'            End If
'            Set rsTemp = Nothing
'          End If
          
          If blnGoOn Then
            Set rsTemp = gcnGi.Execute("Select CustId From " & GetOwner & "So004 Where (CustId=" & strCustid & ") And (FaciSNo='" & strString1 & "')")
            If rsTemp.EOF Then
              blnGoOn = False
              Call LogError1(lngMasterErrorCount, "客編與門號不符或不存在：" & strCustid & "，" & strString1, "", lngLineNumber, vElement)
            End If
            Set rsTemp = Nothing
          End If
        End If
      Case "D"
        'Detail 項目
        If Not blnHasHead Then
          Call LogError1(lngDetailErrorCount, "Detail 項目不可以出現在表頭之前", "", lngLineNumber, vElement)
        ElseIf blnHasTail Then
          Call LogError1(lngDetailErrorCount, "Detail 項目不可以出現在表尾之後", "", lngLineNumber, vElement)
        ElseIf (Not blnInMaster) And (Not blnInDetail) Then
          Call LogError1(lngDetailErrorCount, "Detail 項目沒有對應的 Master 資料", "", lngLineNumber, vElement)
        Else
          lngDetailCount = lngDetailCount + 1
          blnInMaster = False
          'blnInDetail = True
          blnHasDetail = True
          
          '依固定寬度取得各欄位的值
          strCpCItemCode = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 2, 10), vbUnicode))
          strCpCItemName = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 12, 50), vbUnicode))
          strCpTax = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 62, 1), vbUnicode))
          strSign = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 63, 1), vbUnicode))
          strShouldAmt = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 63, 8), vbUnicode))
          strShouldAmtNoTax = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 71, 8), vbUnicode))
          strTax = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 79, 8), vbUnicode))
          dblShouldAmt2 = dblShouldAmt2 + Val(strShouldAmt)
          dblShouldAmtNoTax2 = dblShouldAmtNoTax2 + Val(strShouldAmtNoTax)
          dblTax2 = dblTax2 + Val(strTax)
          strCpShouldYm = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 87, 6), vbUnicode))
          strNote = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 93, 50), vbUnicode))
          dblShouldAmtTotal = dblShouldAmtTotal + Val(strShouldAmt)
          
          Select Case strCpTax
            Case "Y"
              dblShouldAmtWithTax = dblShouldAmtWithTax + Val(strShouldAmt)
            Case "F"
              dblShouldAmtWithoutTax = dblShouldAmtWithoutTax + Val(strShouldAmt)
            Case "N"
              dblShouldAmtZeroTax = dblShouldAmtZeroTax + Val(strShouldAmt)
          End Select
          
          '底下進行各式各樣的檢核
          blnGoOn = True
          Set rsTemp = gcnGi.Execute("Select CItemCode1,CpCItemName, CpTax From " & GetOwner & "So165 Where CpCItemCode='" & strCpCItemCode & "'")
          If rsTemp.EOF Then
            blnGoOn = False
            Call LogError1(lngDetailErrorCount, "客編：" & strCustid & "　TFN帳單項目代碼不存在：" & strCpCItemCode, "", lngLineNumber, vElement)
          Else
            If rsTemp("CpTax") <> strCpTax Then
              blnGoOn = False
              Call LogError1(lngDetailErrorCount, "客編：" & strCustid & "　稅率不符：" & strCpTax, rsTemp("CpTax"), lngLineNumber, vElement)
            End If
'            If rsTemp("CItemCode1") & "" = "" Then
'              blnGoOn = False
'              Call LogError1(lngDetailErrorCount, "客編：" & strCustId & "　TFN帳單項目名稱不存在：" & strCpCItemName, "", lngLineNumber, vElement)
'            End If
          End If
          Set rsTemp = Nothing
                    
          If blnGoOn Then
            If strSign = "-" Then
              strString1 = "CItemCode2"
            Else
              strString1 = "CItemCode1"
            End If
            Set rsTemp = gcnGi.Execute("Select CItemCode1 From " & GetOwner & "So165 Where " & strString1 & " Is Not Null")
            If rsTemp.EOF Then
              blnGoOn = False
              Call LogError1(lngDetailErrorCount, "客編：" & strCustid & "　借貸符號不符：" & strSign, "", lngLineNumber, vElement)
            End If
            Set rsTemp = Nothing
          End If
          
          If blnGoOn Then
            If strSign = "-" Then
              strString1 = "CItemCode2 As CItemCode"
            Else
              strString1 = "CItemCode1 As CItemCode"
            End If
            Set rsTemp = gcnGi.Execute("Select " & strString1 & " From " & GetOwner & "So165 Where (CpCItemCode='" & strCpCItemCode & "') And (CpProperty In (1,2))")
            strCitemCode = NoZero(rsTemp("CItemCode"), True)
            Set rsTemp = Nothing
            strString1 = "Select CustId " & _
                         "From " & GetOwner & "So033 " & _
                         "Where (CustId ='" & strCustid & "') " & _
                         "And (FaciSNo ='" & Trim(strCpArea & strCpTel) & "') " & _
                         "And (TfnBillNo ='" & strTFNBillNo & "') " & _
                         "And (ClctYm ='" & strShouldDate & "') " & _
                         "And (CItemCode ='" & strCitemCode & "') " & _
                         "And (ShouldAmt ='" & strShouldAmt & "') " & _
                         "And (CPShouldYM ='" & strCpShouldYm & "') " & _
                         "Union All " & _
                         "Select CustId " & _
                         "From " & GetOwner & "So034 " & _
                         "Where (CustId ='" & strCustid & "') " & _
                         "And (FaciSNo ='" & Trim(strCpArea & strCpTel) & "') " & _
                         "And (TfnBillNo ='" & strTFNBillNo & "') " & _
                         "And (ClctYm ='" & strShouldDate & "') " & _
                         "And (CItemCode ='" & strCitemCode & "') " & _
                         "And (ShouldAmt ='" & strShouldAmt & "') " & _
                         "And (CPShouldYM ='" & strCpShouldYm & "') "

            Set rsTemp = gcnGi.Execute(strString1)
            If Not rsTemp.EOF Then
              blnGoOn = False
              Call LogError1(lngDetailErrorCount, "客編：" & strCustid & "　資料已存在，可能是電信帳單項目有重複或通信費用資料重複", "", lngLineNumber, vElement)
            End If
            Set rsTemp = Nothing
          End If
          

          '前面的檢核都通過了，準備將匯入的資料寫入 SO033Cp 中
          If blnGoOn Then
            rsSO033Cp.AddNew
            
            rsSO033Cp("CustId") = strCustid
            
            '問題2722 第三點 ----for Taos 加入SO034日結後的ITEM來判斷 2006/09/18
            If strBillNo = "" Then
              strBillNo = strShouldDate & "BP" & strCustId2
              Set rsTemp = gcnGi.Execute("SELECT MAX(MAXITEM) MaxItem FROM (SELECT MAX(ITEM) MAXITEM From " & GetOwner & "So033 Where BillNo='" & strBillNo & "' UNION ALL SELECT MAX(ITEM) MAXITEM From " & GetOwner & "SO034 Where BillNo='" & strBillNo & "')")
              lngBillItem = NoZero(rsTemp("MaxItem"), True) + 1
              Set rsTemp = Nothing
            End If
              
            rsSO033Cp("BillNo") = strBillNo
            rsSO033Cp("Item") = lngBillItem
            
            If strSign = "-" Then
              strString1 = "CItemCode2 As CItemCode, CItemName2 As CItemName"
              strString2 = " And CItemCode2 Is Not Null "
            Else
              strString1 = "CItemCode1 As CItemCode, CItemName1 As CItemName"
              strString2 = " And CItemCode1 Is Not Null "
            End If
            Set rsTemp = gcnGi.Execute("Select " & strString1 & ", CPAdjCitemCode, CPGroupCode, CPItem From " & GetOwner & "So165 Where CpCItemCode='" & strCpCItemCode & "' " & strString2)
            If Not rsTemp.EOF Then
              rsSO033Cp("CitemCode") = NoZero(rsTemp("CItemCode"), True)
              rsSO033Cp("CitemName") = rsTemp("CItemName") & ""
              rsSO033Cp("CPAdjCitemCode") = rsTemp("CPAdjCitemCode") & ""
              rsSO033Cp("CPGroupCode") = rsTemp("CPGroupCode") & ""
              rsSO033Cp("CPItem") = rsTemp("CPItem") & ""
            Else
              rsSO033Cp("CitemCode") = 0
              rsSO033Cp("CitemName") = " "
              rsSO033Cp("CPAdjCitemCode") = 0
              rsSO033Cp("CPGroupCode") = 0
              rsSO033Cp("CPItem") = " "
            End If
            Set rsTemp = Nothing
            
            rsSO033Cp("OldAmt") = strShouldAmt
            rsSO033Cp("OldPeriod") = 0
            
            strString1 = Right("00" & txtShouldDay.Text, 2)
            strTempDate = Left(strShouldDate, 4) & "/" & Right(strShouldDate, 2) & "/" & strString1
            If IsDate(strTempDate) Then
              rsSO033Cp("ShouldDate") = strTempDate
            Else
              Call LogError1(lngDetailErrorCount, "客編：" & strCustid & "　應收日期格式錯誤：" & strTempDate, "", lngLineNumber, vElement)
              rsSO033Cp("ShouldDate") = Date
            End If
            
            rsSO033Cp("ShouldAmt") = strShouldAmt
            rsSO033Cp("RealAmt") = 0
            rsSO033Cp("RealPeriod") = 0
            
            'Set rsTemp = gcnGi.Execute("Select UcCode From " & GetOwner & "So044")
            'If rsTemp.EOF Then
            '  If Not blnInDetail Then
            '    Call LogError1(lngDetailErrorCount, "客編：" & strCustid & "　未收原因代碼未定義", "", lngLineNumber, vElement)
            '  End If
            '  rsSO033Cp("UcCode") = "0"
            'Else
            '  rsSO033Cp("UcCode") = rsTemp("UcCode")
            'End If
            'Set rsTemp = Nothing
            
            rsSO033Cp("UcCode") = strUCCode
            rsSO033Cp("UcName") = strUCName
            
            'Set rsTemp = gcnGi.Execute("Select Description From " & GetOwner & "Cd013 Where CodeNo=" & rsSO033Cp("UcCode"))
            'If rsTemp.EOF Then
            '  Call LogError1(lngDetailErrorCount, "客編：" & strCustid & "　未繳費原因代碼不存在：" & rsSO033Cp("UcCode"), "", lngLineNumber, vElement)
            '  rsSO033Cp("UcName") = "0"
            'Else
            '  rsSO033Cp("UcName") = rsTemp("Description")
            'End If
            'Set rsTemp = Nothing
            
            rsSO033Cp("Note") = strNote
            rsSO033Cp("CreateTime") = UserDateTime
            rsSO033Cp("CreateEn") = garyGi(1)
            rsSO033Cp("CompCode") = garyGi(9)
            rsSO033Cp("CancelFlag") = 0
            rsSO033Cp("PrtCount") = 0
            
            'Set rsTemp = gcnGi.Execute("Select ClassCode1 From " & GetOwner & "So001 Where CustId=" & strCustid)
            'If rsTemp.EOF Then
            '  rsSO033("ClassCode") = "0"
            'Else
            '  rsSO033("ClassCode") = rsTemp("ClassCode1")
            'End If
            'Set rsTemp = Nothing
            
            rsSO033Cp("ClassCode") = strClassCode1
            
            rsSO033Cp("UpdTime") = GetDTString(RightNow)
            rsSO033Cp("UpdEn") = garyGi(1)
            rsSO033Cp("ClctYm") = strShouldDate
            rsSO033Cp("Quantity") = 0
            rsSO033Cp("CustCount") = 0
            rsSO033Cp("ServiceType") = "P"
            rsSO033Cp("ProcessMark") = 0
            rsSO033Cp("ChEven") = 0
            rsSO033Cp("Budget") = 0
            rsSO033Cp("FaciSNo") = Trim(strCpArea & strCpTel)
            rsSO033Cp("Punish") = 0
            rsSO033Cp("MergePrint") = 0
            rsSO033Cp("ApartFlag") = 0
            rsSO033Cp("TfnBillNo") = strTFNBillNo
            rsSO033Cp("TfnServiceId") = strTfnServiceId
            rsSO033Cp("CpShouldYm") = strCpShouldYm
            
'            strString1 = "Select Count(*) As DataCount " & _
'                         "From " & GetOwner & "So003 " & _
'                         "Where (CustId=" & strCustId & ") " & _
'                         "And (ServiceType='P') " & _
'                         "And (FaciSNo='" & Trim(strCpArea & strCpTel) & "') "
'            Set rsTemp = gcnGi.Execute(strString1)
'            lngTempLong = rsTemp("DataCount")

            strString1 = "Select FaciSeqNo, CmCode, CmName, PtCode, " & _
                         "       PtName, AccountNo, BankCode, BankName " & _
                         "From " & GetOwner & "So003 " & _
                         "Where (FaciSNo='" & Trim(strCpArea & strCpTel) & "')" & _
                         "And (CustId=" & strCustid & ") " & _
                         "And (ServiceType='P') "
            Set rsTemp = gcnGi.Execute(strString1)
            
            If rsTemp.RecordCount = 1 Then
                rsSO033Cp("FaciSeqNo") = rsTemp("FaciSeqNo")
                rsSO033Cp("CmCode") = rsTemp("CmCode")
                rsSO033Cp("CmName") = rsTemp("CmName")
                rsSO033Cp("PtCode") = rsTemp("PtCode")
                rsSO033Cp("PtName") = rsTemp("PtName")
                rsSO033Cp("AccountNo") = rsTemp("AccountNo")
                rsSO033Cp("BankCode") = rsTemp("BankCode")
                rsSO033Cp("BankName") = rsTemp("BankName")
'              End If
            Else
            
              strString1 = "Select CodeNo, Description " & _
                           "From " & GetOwner & "Cd031 A," & GetOwner & "So044 B " & _
                           "Where (A.CodeNo=B.CmCode) " & _
                           "And ((A.ServiceType='P') " & _
                           "Or (A.ServiceType Is Null)) "
              Set rsTemp2 = gcnGi.Execute(strString1)
              If rsTemp2.EOF Then
                If Not blnInDetail Then
                  Call LogError1(lngDetailErrorCount, "客編：" & strCustid & "　收費方式參數未定義", "", lngLineNumber, vElement)
                End If
                rsSO033Cp("CmCode") = "0"
                rsSO033Cp("CmName") = "0"
                rsSO033Cp("FaciSeqNo") = "0"
                rsSO033Cp("PtCode") = "0"
                rsSO033Cp("PtName") = "0"
                rsSO033Cp("AccountNo") = "0"
                rsSO033Cp("BankCode") = "0"
                rsSO033Cp("BankName") = "0"
              Else
                rsSO033Cp("CmCode") = rsTemp2("CodeNo")
                rsSO033Cp("CmName") = rsTemp2("Description")
                rsSO033Cp("FaciSeqNo") = Null
                rsSO033Cp("PtCode") = Null
                rsSO033Cp("PtName") = Null
                rsSO033Cp("AccountNo") = Null
                rsSO033Cp("BankCode") = Null
                rsSO033Cp("BankName") = Null
              End If
              Set rsTemp2 = Nothing
            End If
            Set rsTemp = Nothing
            
            '依裝機地址或收費地址，將地址填入到 SO033 的地址資料中
            If optEmpNo.Value Then
              '裝機地址
              strString1 = "Select B.ClctEn, B.ClctName, B.AddrNo, B.StrtCode, " & _
                           "       A.MduId, A.ServCode, A.ClctAreaCode, " & _
                           "       B.AreaCode, B.NodeNo " & _
                           "From " & GetOwner & "So001 A, " & GetOwner & "So014 B " & _
                           "Where (A.InstAddrNo=B.AddrNo) " & _
                           "And (A.CustId=" & strCustid & ") "
              Set rsTemp = gcnGi.Execute(strString1)
            Else
             '收費地址
              strString1 = "Select Count(*) As DataCount " & _
                           "From " & GetOwner & "So003 " & _
                           "Where (FaciSNo='" & Trim(strCpArea & strCpTel) & "') " & _
                           "And (CustId=" & strCustid & ") " & _
                           "And (ServiceType='P') " & _
                           "And (AccountNo Is Not Null) "
              Set rsTemp = gcnGi.Execute(strString1)
              lngTempLong = rsTemp("DataCount")
              Set rsTemp = Nothing
              If lngTempLong = 0 Then
                strString1 = "Select B.ClctEn, B.ClctName, B.AddrNo, B.StrtCode, " & _
                             "       A.MduId, A.ServCode, A.ClctAreaCode, " & _
                             "       B.AreaCode, B.NodeNo " & _
                             "From " & GetOwner & "So001 A, " & GetOwner & "So014 B " & _
                             "Where (A.ChargeAddrNo=B.AddrNo) " & _
                             "And (A.CustId=" & strCustid & ") "
                Set rsTemp = gcnGi.Execute(strString1)
              Else
                strString1 = "Select C.ClctEn, C.ClctName, C.AddrNo, C.StrtCode, " & _
                             "       A.MduId, A.ServCode, A.ClctAreaCode, " & _
                             "       C.AreaCode, C.NodeNo " & _
                             "From " & GetOwner & "So001 A, " & GetOwner & "So002A B, " & GetOwner & "So014 C " & _
                             "Where (B.ChargeAddrNo=C.AddrNo) " & _
                             "And (B.CustId=" & strCustid & ") " & _
                             "And (B.CustId=A.CustId) "
                Set rsTemp = gcnGi.Execute(strString1)
              End If
            End If
            
            '填入到 SO033 的地址資料中
            If rsTemp.EOF Then
              If Not blnInDetail Then
                Call LogError1(lngDetailErrorCount, "客編：" & strCustid & "　地址資料不存在：" & strCustid & "，" & Trim(strCpArea & strCpTel), "", lngLineNumber, vElement)
              End If
              rsSO033Cp("ClctEn") = "0"
              rsSO033Cp("ClctName") = "0"
              rsSO033Cp("AddrNo") = "0"
              rsSO033Cp("StrtCode") = "0"
              rsSO033Cp("MduId") = "0"
              rsSO033Cp("ServCode") = "0"
              rsSO033Cp("ClctAreaCode") = "0"
              rsSO033Cp("OldClctEn") = "0"
              rsSO033Cp("OldClctName") = "0"
              rsSO033Cp("AreaCode") = "0"
              rsSO033Cp("NodeNo") = "0"
            Else
              rsSO033Cp("ClctEn") = rsTemp("ClctEn")
              rsSO033Cp("ClctName") = rsTemp("ClctName")
              rsSO033Cp("AddrNo") = rsTemp("AddrNo")
              rsSO033Cp("StrtCode") = rsTemp("StrtCode")
              rsSO033Cp("MduId") = rsTemp("MduId")
              rsSO033Cp("ServCode") = rsTemp("ServCode")
              rsSO033Cp("ClctAreaCode") = rsTemp("ClctAreaCode")
              rsSO033Cp("OldClctEn") = rsTemp("ClctEn")
              rsSO033Cp("OldClctName") = rsTemp("ClctName")
              rsSO033Cp("AreaCode") = rsTemp("AreaCode")
              rsSO033Cp("NodeNo") = rsTemp("NodeNo")
            End If
            
            Set rsTemp = Nothing
            rsSO033Cp.Update
            
          End If
          blnInDetail = True
        End If
      Case "S"
        '小計項目
        If Not blnHasHead Then
          Call LogError1(lngSumErrorCount, "小計項目不可以出現在表頭之前", "", lngLineNumber, vElement)
        ElseIf blnHasTail Then
          Call LogError1(lngSumErrorCount, "小計項目不可以出現在表尾之後", "", lngLineNumber, vElement)
        ElseIf Not blnInDetail Then
          Call LogError1(lngSumErrorCount, "小計項目沒有對應的 Detail 資料", "", lngLineNumber, vElement)
        Else
          lngSumCount = lngSumCount + 1
          blnInMaster = False
          blnHasSum = True
          
          '依固定寬度取得各欄位的值
          strCpCItemCode = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 2, 10), vbUnicode))
          strCpCItemName = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 12, 50), vbUnicode))
          strCpTax = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 62, 1), vbUnicode))
          strSign = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 63, 1), vbUnicode))
          strShouldAmt = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 63, 8), vbUnicode))
          strShouldAmtNoTax = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 71, 8), vbUnicode))
          strTax = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 79, 8), vbUnicode))
          strCpShouldYm = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 87, 6), vbUnicode))
          strCpAdjCItemCode = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 93, 10), vbUnicode))
          strNote = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 103, 50), vbUnicode))
          
          '底下進行各式各樣的檢核
          Set rsTemp = gcnGi.Execute("Select CpAdjCItemCode,CItemCode1, CpTax,CpCItemName From " & GetOwner & "So165 Where CpCItemCode='" & strCpCItemCode & "' And (CpProperty In (1,2))")
          If rsTemp.EOF Then
            Call LogError1(lngSumErrorCount, "客編：" & strCustid & "　電信帳單項目代碼不存在：" & strCpCItemCode, "", lngLineNumber, vElement)
            
          Else
            If rsTemp("CpAdjCItemCode") & "" <> strCpAdjCItemCode Then
              Call LogError1(lngSumErrorCount, "客編：" & strCustid & "　調改項目不符：" & strCpAdjCItemCode, "", lngLineNumber, vElement)
            End If
            If rsTemp("CpTax") <> strCpTax Then
              blnGoOn = False
              Call LogError1(lngSumErrorCount, "客編：" & strCustid & "　稅率不符：" & strCpTax, rsTemp("CpTax"), lngLineNumber, vElement)
            End If
'            If rsTemp("CItemCode1") & "" = "" Then
'                blnGoOn = False
'                Call LogError1(lngSumErrorCount, "客編：" & strCustId & "　TFN帳單項目代碼或名稱不存在：" & strCpCItemCode, "", lngLineNumber, vElement)
'            End If
          End If
          Set rsTemp = Nothing
          
          blnGoOn = True
          
          If blnGoOn Then
            If strSign = "-" Then
              strString1 = "CItemCode2"
            Else
              strString1 = "CItemCode1"
            End If
            Set rsTemp = gcnGi.Execute("Select CItemCode1 From " & GetOwner & "So165 Where " & strString1 & " Is Not Null")
            If rsTemp.EOF Then
              blnGoOn = False
              Call LogError1(lngSumErrorCount, "客編：" & strCustid & "　借貸符號不符：" & strSign, "", lngLineNumber, vElement)
            End If
            Set rsTemp = Nothing
          End If
          '問題集 2722 第四點沒修改到
          If blnGoOn Then
            If strSign = "-" Then
              strString1 = "CItemCode2 As CItemCode"
            Else
              strString1 = "CItemCode1 As CItemCode"
            End If
            Set rsTemp = gcnGi.Execute("Select " & strString1 & " From " & GetOwner & "So165 Where (CpCItemCode='" & strCpCItemCode & "') And (CpProperty In (1,2))")
            strCitemCode = NoZero(rsTemp("CItemCode"), True)
            If strCitemCode = "" Then
                strString1 = "CItemCode2 As CItemCode"
                Set rsTemp2 = gcnGi.Execute("Select " & strString1 & " From " & GetOwner & "So165 Where (CpCItemCode='" & strCpCItemCode & "') And (CpProperty In (1,2))")
                strCitemCode = NoZero(rsTemp2("CItemCode"), True)
                Set rsTemp2 = Nothing
            End If
            Set rsTemp = Nothing
            strString1 = "Select CustId " & _
                         "From " & GetOwner & "So033 " & _
                         "Where (FaciSNo ='" & Trim(strCpArea & strCpTel) & "') " & _
                         "And (CustId ='" & strCustid & "') " & _
                         "And (TfnBillNo ='" & strTFNBillNo & "') " & _
                         "And (ClctYm =" & strYear & strMonth & ") " & _
                         "And (CItemCode ='" & strCitemCode & "') " & _
                         "And (ShouldAmt ='" & strShouldAmt & "') " & _
                         "And (CPShouldYM ='" & strCpShouldYm & "') " & _
                         "Union All " & _
                         "Select CustId " & _
                         "From " & GetOwner & "So034 " & _
                         "Where (FaciSNo ='" & Trim(strCpArea & strCpTel) & "') " & _
                         "And (CustId ='" & strCustid & "') " & _
                         "And (TfnBillNo ='" & strTFNBillNo & "') " & _
                         "And (ClctYm =" & strYear & strMonth & ") " & _
                         "And (CItemCode ='" & strCitemCode & "') " & _
                         "And (ShouldAmt ='" & strShouldAmt & "') " & _
                         "And (CPShouldYM ='" & strCpShouldYm & "') "
'            Set rsTemp = gcnGi.Execute(strString1)
            If Not GetRS(rsTemp, strString1, gcnGi, , adOpenForwardOnly, adLockOptimistic, adUseClient, False) Then Exit Sub
            If Not rsTemp.EOF Then
              blnGoOn = False
              Call LogError1(lngSumErrorCount, "客編：" & strCustid & "　資料已存在，可能是電信帳單項目有重複或通信費用資料重複", "", lngLineNumber, vElement)
            End If
            Set rsTemp = Nothing
          End If
          
          '前面的檢核都通過了，準備將匯入的資料寫入 SO033 中
          If blnGoOn Then
            rsSO033.AddNew
            
            rsSO033("CustId") = strCustid
            
            rsSO033("BillNo") = strBillNo
            rsSO033("Item") = lngBillItem
            '問題集2722 台固第四點沒修改到
            If strSign = "-" Then
              strString1 = "CItemCode2 As CItemCode, CItemName2 As CItemName"
              strString2 = " And CItemCode2 Is Not Null "
            Else
              strString1 = "CItemCode1 As CItemCode, CItemName1 As CItemName"
              strString2 = " And CItemCode1 Is Not Null "
            End If
            Set rsTemp = gcnGi.Execute("Select " & strString1 & ", CPAdjCitemCode, CPGroupCode From " & GetOwner & "So165 Where CPProperty=1 AND CpCItemCode='" & strCpCItemCode & "' " & strString2)
            If Not rsTemp.EOF Then
              rsSO033("CitemCode") = NoZero(rsTemp("CItemCode"), True)
              rsSO033("CitemName") = rsTemp("CItemName") & " "
              rsSO033("CPAdjCitemCode") = strCpAdjCItemCode
              rsSO033("CPGroupCode") = rsTemp("CPGroupCode") & ""
            Else
              strString1 = "CItemCode2 As CItemCode, CItemName2 As CItemName"
              strString2 = " And CItemCode2 Is Not Null "
              Set rsTemp2 = gcnGi.Execute("Select " & strString1 & ", CPAdjCitemCode, CPGroupCode From " & GetOwner & "So165 Where CPProperty=1 AND CpCItemCode='" & strCpCItemCode & "' " & strString2)
              If Not rsTemp2.EOF Then
                rsSO033("CitemCode") = NoZero(rsTemp2("CItemCode"), True)
                rsSO033("CitemName") = rsTemp2("CItemName") & " "
                rsSO033("CPAdjCitemCode") = strCpAdjCItemCode
                rsSO033("CPGroupCode") = rsTemp2("CPGroupCode") & ""
              Else
                rsSO033("CitemCode") = 0
                rsSO033("CitemName") = " "
                rsSO033("CPAdjCitemCode") = 0
                rsSO033("CPGroupCode") = 0
              End If
              Set rsTemp = Nothing
            End If
            
            Set rsTemp2 = Nothing
            
            rsSO033("OldAmt") = Round(strShouldAmt)
            rsSO033("OldPeriod") = 0
            
            strString1 = Right("00" & txtShouldDay.Text, 2)
            strTempDate = Left(strShouldDate, 4) & "/" & Right(strShouldDate, 2) & "/" & strString1
            If IsDate(strTempDate) Then
              rsSO033("ShouldDate") = strTempDate
            Else
              Call LogError1(lngSumErrorCount, "客編：" & strCustid & "　應收日期格式錯誤：" & strTempDate, "", lngLineNumber, vElement)
              rsSO033("ShouldDate") = Date
            End If
            
            rsSO033("ShouldAmt") = Round(strShouldAmt)
            rsSO033("RealAmt") = 0
            rsSO033("RealPeriod") = 0
            rsSO033("UcCode") = strUCCode
            rsSO033("UcName") = strUCName
            rsSO033("Note") = strNote
            rsSO033("CreateTime") = UserDateTime
            rsSO033("CreateEn") = garyGi(1)
            rsSO033("CompCode") = garyGi(9)
            rsSO033("CancelFlag") = 0
            rsSO033("PrtCount") = 0
            rsSO033("ClassCode") = strClassCode1
            rsSO033("UpdTime") = GetDTString(RightNow)
            rsSO033("UpdEn") = garyGi(1)
            rsSO033("ClctYm") = strShouldDate
            rsSO033("Quantity") = 0
            rsSO033("CustCount") = 0
            rsSO033("ServiceType") = "P"
            rsSO033("ProcessMark") = 0
            rsSO033("ChEven") = 0
            rsSO033("Budget") = 0
            rsSO033("FaciSNo") = Trim(strCpArea & strCpTel)
            rsSO033("Punish") = 0
            rsSO033("MergePrint") = 0
            rsSO033("ApartFlag") = 0
            rsSO033("TfnBillNo") = strTFNBillNo
            rsSO033("TfnServiceId") = strTfnServiceId
            rsSO033("CpShouldYm") = strCpShouldYm
            
            strString1 = "Select FaciSeqNo, CmCode, CmName, PtCode, " & _
                         "       PtName, AccountNo, BankCode, BankName " & _
                         "From " & GetOwner & "So003 " & _
                         "Where (FaciSNo='" & Trim(strCpArea & strCpTel) & "') " & _
                         "And (CustId=" & strCustid & ") " & _
                         "And (ServiceType='P') "
            Set rsTemp = gcnGi.Execute(strString1)
            
            If Not rsTemp.EOF Then
                rsSO033("FaciSeqNo") = rsTemp("FaciSeqNo")
                rsSO033("CmCode") = rsTemp("CmCode")
                rsSO033("CmName") = rsTemp("CmName")
                rsSO033("PtCode") = rsTemp("PtCode")
                rsSO033("PtName") = rsTemp("PtName")
                rsSO033("AccountNo") = rsTemp("AccountNo")
                rsSO033("BankCode") = rsTemp("BankCode")
                rsSO033("BankName") = rsTemp("BankName")
            Else
              strString1 = "Select CodeNo, Description " & _
                           "From " & GetOwner & "Cd031 A," & GetOwner & "So044 B " & _
                           "Where (A.CodeNo=B.CmCode) " & _
                           "And ((A.ServiceType='P') " & _
                           "Or (A.ServiceType Is Null)) "
              Set rsTemp2 = gcnGi.Execute(strString1)
              If rsTemp2.EOF Then
                If Not blnInDetail Then
                  Call LogError1(lngSumErrorCount, "客編：" & strCustid & "　收費方式參數未定義", "", lngLineNumber, vElement)
                End If
                rsSO033("CmCode") = "0"
                rsSO033("CmName") = "0"
                rsSO033("FaciSeqNo") = "0"
                rsSO033("PtCode") = "0"
                rsSO033("PtName") = "0"
                rsSO033("AccountNo") = "0"
                rsSO033("BankCode") = "0"
                rsSO033("BankName") = "0"
              Else
                rsSO033("CmCode") = rsTemp2("CodeNo")
                rsSO033("CmName") = rsTemp2("Description")
                rsSO033("FaciSeqNo") = Null
                rsSO033("PtCode") = Null
                rsSO033("PtName") = Null
                rsSO033("AccountNo") = Null
                rsSO033("BankCode") = Null
                rsSO033("BankName") = Null
              End If
              Set rsTemp2 = Nothing
            End If
            Set rsTemp = Nothing
            '依裝機地址或收費地址，將地址填入到 SO033 的地址資料中
            If optEmpNo.Value Then
              '裝機地址
              strString1 = "Select B.ClctEn, B.ClctName, B.AddrNo, B.StrtCode, " & _
                           "       A.MduId, A.ServCode, A.ClctAreaCode, " & _
                           "       B.AreaCode, B.NodeNo " & _
                           "From " & GetOwner & "So001 A, " & GetOwner & "So014 B " & _
                           "Where (A.InstAddrNo=B.AddrNo) " & _
                           "And (A.CustId=" & strCustid & ") "
              Set rsTemp = gcnGi.Execute(strString1)
            Else
             '收費地址
              strString1 = "Select Count(*) As DataCount " & _
                           "From " & GetOwner & "So003 " & _
                           "Where (FaciSNo='" & Trim(strCpArea & strCpTel) & "') " & _
                           "And (CustId=" & strCustid & ") " & _
                           "And (ServiceType='P') " & _
                           "And (AccountNo Is Not Null) "
              Set rsTemp = gcnGi.Execute(strString1)
              lngTempLong = rsTemp("DataCount")
              Set rsTemp = Nothing
              If lngTempLong = 0 Then
                strString1 = "Select B.ClctEn, B.ClctName, B.AddrNo, B.StrtCode, " & _
                             "       A.MduId, A.ServCode, A.ClctAreaCode, " & _
                             "       B.AreaCode, B.NodeNo " & _
                             "From " & GetOwner & "So001 A, " & GetOwner & "So014 B " & _
                             "Where (A.ChargeAddrNo=B.AddrNo) " & _
                             "And (A.CustId=" & strCustid & ") "
                Set rsTemp = gcnGi.Execute(strString1)
              Else
                strString1 = "Select C.ClctEn, C.ClctName, C.AddrNo, C.StrtCode, " & _
                             "       A.MduId, A.ServCode, A.ClctAreaCode, " & _
                             "       C.AreaCode, C.NodeNo " & _
                             "From " & GetOwner & "So001 A, " & GetOwner & "So002A B, " & GetOwner & "So014 C " & _
                             "Where (B.ChargeAddrNo=C.AddrNo) " & _
                             "And (B.CustId=" & strCustid & ") " & _
                             "And (B.CustId=A.CustId) "
                Set rsTemp = gcnGi.Execute(strString1)
              End If
            End If
            
            '填入到 SO033 的地址資料中
            If rsTemp.EOF Then
              If Not blnInDetail Then
                Call LogError1(lngSumErrorCount, "客編：" & strCustid & "　地址資料不存在：" & strCustid & "，" & Trim(strCpArea & strCpTel), "", lngLineNumber, vElement)
              End If
              rsSO033("ClctEn") = "0"
              rsSO033("ClctName") = "0"
              rsSO033("AddrNo") = "0"
              rsSO033("StrtCode") = "0"
              rsSO033("MduId") = "0"
              rsSO033("ServCode") = "0"
              rsSO033("ClctAreaCode") = "0"
              rsSO033("OldClctEn") = "0"
              rsSO033("OldClctName") = "0"
              rsSO033("AreaCode") = "0"
              rsSO033("NodeNo") = "0"
            Else
              rsSO033("ClctEn") = rsTemp("ClctEn")
              rsSO033("ClctName") = rsTemp("ClctName")
              rsSO033("AddrNo") = rsTemp("AddrNo")
              rsSO033("StrtCode") = rsTemp("StrtCode")
              rsSO033("MduId") = rsTemp("MduId")
              rsSO033("ServCode") = rsTemp("ServCode")
              rsSO033("ClctAreaCode") = rsTemp("ClctAreaCode")
              rsSO033("OldClctEn") = rsTemp("ClctEn")
              rsSO033("OldClctName") = rsTemp("ClctName")
              rsSO033("AreaCode") = rsTemp("AreaCode")
              rsSO033("NodeNo") = rsTemp("NodeNo")
            End If
            
            Set rsTemp = Nothing
            rsSO033.Update
            
          End If
          
          'If (dblShouldAmt2 <> Val(strShouldAmt3)) Or (dblShouldAmtNoTax2 <> Val(strShouldAmtNoTax3)) Or (dblTax2 <> Val(strTax3)) Then
          '  Call LogError1(lngSumErrorCount, "小計金額與Detail總合不同", vElement,lngLineNumber,vElement)
          'End If
          
          dblShouldAmt2 = 0
          dblShouldAmtNoTax2 = 0
          dblTax2 = 0
          
          '用 strBillNo 來判斷不同的小計項目
          strBillNo = ""
          
        End If
      Case "T"
        '檔尾
        If (Not blnHasMaster) Or (Not blnHasDetail) Or (Not blnHasHead) Then
          Call LogError1(lngTailErrorCount, "表尾必須在文字檔的最後一行", "", lngLineNumber, vElement)
        Else
          blnHasTail = True
          
          '依固定寬度取得各欄位的值
          strMasterShouldCount = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 2, 6), vbUnicode))
          strDetailShouldCount = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 8, 6), vbUnicode))
          strSumShouldCount = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 14, 6), vbUnicode))
          strTotalDataCount = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 20, 6), vbUnicode))
          strMoneyTotal = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 26, 8), vbUnicode))
          
          '比對檔尾記載的筆數及實際匯入的筆數
          blnGoOn = True
          If (lngMasterCount + lngDetailCount + lngSumCount) <> Val(strTotalDataCount) Then
            blnGoOn = False
            Call LogError1(lngTailErrorCount, "資料筆數(" & (lngMasterCount + lngDetailCount + lngSumCount) & ")與表尾記載(" & strTotalDataCount & ")的不符", "", lngLineNumber, vElement)
          End If
          
          '比對檔尾記載的金額及實際匯入的金額
          'If blnGoOn Then
            If Val(dblShouldAmtTotal) <> Val(strMoneyTotal) Then
              blnGoOn = False
              Call LogError1(lngTailErrorCount, "資料金額(" & dblShouldAmtTotal & ")與表尾記載(" & CDbl(strMoneyTotal) & ")的不符", "", lngLineNumber, vElement)
            End If
          'End If
        End If
      Case Else
        If vElement <> "" Then
          Call LogError1(lngDetailErrorCount, "資料無效，請檢查第1碼是否標示錯誤", "", lngLineNumber, vElement)
        End If
    End Select
    lngLineNumber = lngLineNumber + 1
    DoEvents
  Next
  If Not blnHasTail Then
    Call LogError1(lngTailErrorCount, "沒有表尾資料", "", lngLineNumber, vElement)
    lngLineNumber = lngLineNumber + 1
  End If
  PB1.Visible = False
  
  If (lngMasterErrorCount > 0) Or (lngDetailErrorCount > 0) Or (lngSumErrorCount > 0) Or (lngHeadErrorCount > 0) Or (lngTailErrorCount > 0) Then
    '有錯誤，已匯入的資料須取消匯入
    gcnGi.RollbackTrans
    strTextData = "預計匯入筆數：" & lngDetailCount & vbCrLf & _
                  "表頭錯誤筆數：" & lngHeadErrorCount & vbCrLf & _
                  "表尾錯誤筆數：" & lngTailErrorCount & vbCrLf & _
                  "Master錯誤筆數：" & lngMasterErrorCount & vbCrLf & _
                  "Detail錯誤筆數：" & lngDetailErrorCount & vbCrLf & _
                  "小計錯誤筆數：" & lngSumErrorCount & vbCrLf & vbCrLf & _
                  "匯入資料中有錯誤筆數，因此將不存檔。" & vbCrLf & _
                  "請參考問題記錄檔。" & vbCrLf
    MsgBox strTextData, vbExclamation, "警告"
    Shell "Notepad " & txtLogFile1.Text, vbNormalFocus
  Else
    gcnGi.CommitTrans
    strTextData = "匯入完成，所有資料皆確認無誤 !!" & vbCrLf & vbCrLf & _
                  "匯入筆數：" & lngMasterCount + lngDetailCount + lngSumCount & vbCrLf & _
                  "帳單筆數：" & lngMasterCount & vbCrLf & _
                  "通信費用總計：" & dblShouldAmtTotal & vbCrLf & _
                  "通信費用應稅金額總計：" & dblShouldAmtWithTax & vbCrLf & _
                  "通信費用免稅金額總計：" & dblShouldAmtWithoutTax & vbCrLf & _
                  "通信費用應稅零稅金額總計：" & dblShouldAmtZeroTax & vbCrLf
    MsgBox strTextData, vbExclamation, "訊息"
  End If
  
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  gcnGi.RollbackTrans
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Call ErrSub(Me.Name, "ImportData1")
End Sub

Private Sub cmdInport2_Click()
On Error GoTo ChkErr
  '基本資料核對匯入
  If txtImportFile2.Text = "" Then
    MsgBox "請先指定匯入檔案。", vbExclamation, "警告"
    txtImportFile2.SetFocus
  ElseIf txtLogFile2.Text = "" Then
    MsgBox "請先指定問題記錄檔。", vbExclamation, "警告"
    txtLogFile2.SetFocus
  ElseIf Not FSO.FileExists(txtImportFile2.Text) Then
    MsgBox "要匯入的檔案不存在。", vbExclamation, "警告"
    txtImportFile2.SetFocus
  ElseIf txtImportFile2.Text = txtLogFile2.Text Then
    MsgBox "匯入檔案和問題記錄檔不可以一樣。", vbExclamation, "警告"
    txtImportFile2.SetFocus
  Else
  
    '須先清空 TMP012
    gcnGi.Execute ("Delete From " & GetOwner & "Tmp012")
    
    '＊＊＊＊＊＊ SO19J0A 台固格式以Oracle TMP012 實體表格匯入比對　＊＊＊＊＊＊＊＊＊
    '＊＊＊＊＊＊ SO19J0A 台固格式以Oracle TMP012 實體表格匯入比對　＊＊＊＊＊＊＊＊＊
    
    '開一個 TMP012 的 Client Dataset
    Set rsTmp012 = New ADODB.Recordset
'    rsTmp012.CursorLocation = adUseClient
'    rsTmp012.Open "Select * From " & GetOwner & "Tmp012 Where 1=0", gcnGi, adOpenKeyset, adLockOptimistic
    If Not GetRS(rsTmp012, "Select * From " & GetOwner & "Tmp012 Where 1=0", gcnGi, , adOpenKeyset, adLockOptimistic, adUseClient) Then Exit Sub
    '新建一個 LOG 檔
    If FSO.FileExists(txtLogFile2.Text) Then
      FSO.DeleteFile (txtLogFile2.Text)
    End If
    
    Set strmTextFile1 = FSO.OpenTextFile(txtLogFile2.Text, ForAppending, True)
    strmTextFile1.Write vbCrLf & "########################################################################################################################################################################################################"
    strmTextFile1.Write vbCrLf & "基本資料核對匯入　錯誤記錄　記錄日期：" & Now & vbCrLf & vbCrLf
    strmTextFile1.Write vbCrLf & "行號      TFN客編   MIS客編   TFN門號             訊息                                    TFN                 MIS                 由TFN門號反查之客編      來源單號            安裝時間       拆除時間"
    
  
    strmTextFile1.Write vbCrLf & "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf
    Call ImportData2
    
    Set strmTextFile1 = Nothing
    Set rsTmp012 = Nothing
  End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdInport1_Click")
End Sub

Private Sub cmdInport3_Click()
On Error GoTo ChkErr
  Dim StrTableName As String
  
  'CP通話明細資料匯入
  If optDetailDaily.Value And (txtImportFile3.Text = "") Then
    MsgBox "請先指定每日檢核匯入檔名。", vbExclamation, "警告"
    txtImportFile3.SetFocus
  ElseIf optDetailMonthly.Value And (txtImportFile4.Text = "") Then
    MsgBox "請先指定每月檢核匯入檔名。", vbExclamation, "警告"
    txtImportFile4.SetFocus
  ElseIf txtImportFile5.Text = "" Then
    MsgBox "請先指定匯入檔案。", vbExclamation, "警告"
    txtImportFile5.SetFocus
  ElseIf txtLogFile3.Text = "" Then
    MsgBox "請先指定問題記錄檔。", vbExclamation, "警告"
    txtLogFile1.SetFocus
  ElseIf Not FSO.FileExists(txtImportFile5.Text) Then
    MsgBox "要匯入的檔案不存在。", vbExclamation, "警告"
    txtImportFile5.SetFocus
  ElseIf txtImportFile5.Text = txtLogFile3.Text Then
    MsgBox "匯入檔案和問題記錄檔不可以一樣。", vbExclamation, "警告"
    txtImportFile5.SetFocus
  Else
  
    '通話明細種類，SO168=每日資料，SO169=每月資料
    StrTableName = IIf(optDetailDaily.Value, "So168", "So169")
    
    '開一個 SO168 的 Client Dataset
    Set rsSo168 = New ADODB.Recordset
    rsSo168.CursorLocation = adUseClient
    
    '新建一個 LOG 檔
    If FSO.FileExists(txtLogFile3.Text) Then
      FSO.DeleteFile (txtLogFile3.Text)
    End If
    
    Set strmTextFile1 = FSO.OpenTextFile(txtLogFile3.Text, ForAppending, True)
    strmTextFile1.Write vbCrLf & "################################################################################"
    strmTextFile1.Write vbCrLf & "CP通話明細資料匯入　錯誤記錄　記錄日期：" & Now & vbCrLf & vbCrLf
    
    Call ImportData3
    
    Set strmTextFile1 = Nothing
    Set rsSo168 = Nothing
  End If
ChkErr:
  Call ErrSub(Me.Name, "cmdInport3_Click")
End Sub

Private Sub cmdPrint_Click()
On Error GoTo ChkErr
  Dim strsql As String
  
  '列印
  strsql = "Select CItemCode1, CItemName1, TaxCode1, TaxName1, CItemCode2, CItemName2, TaxCode2, " & _
           "       TaxName2, CpCItemCode, CpCItemName, CpAdjCItemCode, CpGroupCode, CpGroupName, " & _
           "       CpItem, CpTax, CpProperty " & _
           "From " & GetOwner & "So165"
  Call ReadyGoPrint
  Call PrintRpt2(GetPrinterName(5), RptName("SO19J0"), , "CP 帳務資料拋轉作業（for 台灣固網 TFN）[SO19J0A]", strsql, "全部", "", True, , False, , GiPaperLandscape)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub
Private Sub cmdSave_Click()
On Error GoTo ChkErr
  '存檔
  If IsDataOk() Then
    Call ScrToRcd
    
    '若是新增模式．則斷續新增，否則回瀏覽模式
    If lngEditMode = giEditModeInsert Then
      Call ChangeMode(giEditModeInsert)
    Else
      Call ChangeMode(giEditModeView)
    End If
  End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdSave_Click")
End Sub

Private Sub cmdSelectPath1_Click()
  Call SelectPath(txtExportPath1)
End Sub

Private Sub SelectPath(ByRef ExportPath As TextBox)
On Error GoTo ChkErr
  ExportPath.Text = FolderDialog("請選擇要匯出檔案的路徑")
ChkErr:
  Call ErrSub(Me.Name, "SelectPath")
End Sub

Private Sub cmdSelectPath2_Click()
  Call SelectPath(txtExportPath2)
End Sub

Private Sub Command11_Click()
  Call cmdCancel_Click
End Sub

Private Sub cmdSelectPath3_Click()
  Call SelectPath(txtExportPath3)
End Sub

Private Sub cmdSelectPath4_Click()
  Call SelectPath(txtExportPath4)
End Sub

Private Sub cmdSelectPath5_Click()
  Call SelectPath(txtExportPath5)
End Sub

Private Sub Command1_Click()
  Call SelectImportFile(txtImportFile1, txtLogFile1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ChkErr
  If ChkGiList(KeyCode) Then
    Select Case KeyCode
      Case vbKeyF6
        cmdAdd.Value = cmdAdd.Enabled
      Case vbKeyF2
        cmdSave.Value = cmdSave.Enabled
      Case vbKeyF11
        cmdEdit.Value = cmdEdit.Enabled
      Case vbKeyF5
        cmdPrint.Value = cmdPrint.Enabled
      Case vbKeyEscape
        cmdCancel.Value = True
    End Select
  End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
On Error GoTo ChkErr
  Dim blnGoOn As Boolean
  
  Set rsSo165 = New ADODB.Recordset
  Call LoadCom
  Call LoadGrid
  Call OpenData
  Call RcdToScr
  Call ChangeMode(giEditModeView)
  
  '填入畫面元件的預設值
  GiExportDate1.SetValue (RightDate)
  GiExportDate2.SetValue (RightDate)
  GiExportDate3.SetValue (RightDate)
  gimPrCode1.SelectAll
  GiExportRefDate1.SetValue (RightDate)
  GiExportDate4.SetValue (RightDate)
  gimPrCode2.SelectAll
  GiExportDate5.SetValue (RightDate)
  GiExportDate6.SetValue (RightDate)
  blnGoOn = GetUserPriv("SO19J01", ("SO19J01"))
  txtExportFile1.Enabled = blnGoOn
  txtExportFile2.Enabled = blnGoOn
  txtExportFile3.Enabled = blnGoOn
  txtExportFile4.Enabled = blnGoOn
  txtExportFile5.Enabled = blnGoOn
  
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ChkErr
  If lngEditMode <> giEditModeView Then
    MsgBox "請先確定存檔或取消存檔之後，再關閉表單。", vbExclamation, "警告"
    Cancel = True
  End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_QueryUnload")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ChkErr
  '結束程式
  Set FSO = Nothing
  CloseRecordset rsSo165
  Call ReleaseCOM(Me)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Unload")
End Sub

Private Sub LoadCom()
On Error GoTo ChkErr
  SetLst gilCItem1, "CodeNo", "Description", 5, 40, , , "CD019", " Where (Sign='+') And (StopFlag=0) And (ServiceType='P') ", True
  SetLst gilCItem2, "CodeNo", "Description", 5, 40, , , "CD019", " Where (Sign='-') And (StopFlag=0) And (ServiceType='P') ", True
  SetLst gilTax1, "CodeNo", "Description", 5, 40, , , "CD033", " Where (StopFlag=0) ", True
  SetLst gilTax2, "CodeNo", "Description", 5, 40, , , "CD033", " Where (StopFlag=0) ", True
  'SetgiMulti gimPrCode1, "CodeNo", "Description", "CD007", "拆機類別代碼", "拆機類別名稱", " Where (StopFlag=0) And (ServiceType='P') And (RefNo=2) ", True
  '2006/01/17 討論參考號改為 5
  '2006/01/26 又說要改成 2
  SetgiMulti gimPrCode1, "CodeNo", "Description", "CD007", "拆機類別代碼", "拆機類別名稱", " Where (StopFlag=0) And (ServiceType='P') And (RefNo=2) ", True
  SetgiMulti gimPrCode2, "CodeNo", "Description", "CD007", "拆機類別代碼", "拆機類別名稱", " Where (StopFlag=0) And (ServiceType='P') And (RefNo=2) ", True
  SetgiMulti gimReasonCode, "CodeNo", "Description", "CD014", "停拆機原因代碼", "停拆機原因代碼", " Where (StopFlag=0) And ((ServiceType='P') Or (ServiceType Is Null)) ", True
  SetgiMulti gimUCCode, "CodeNo", "Description", "CD014", "停拆機原因代碼", "停拆機原因代碼", " Where (StopFlag=0) And ((ServiceType='P') Or (ServiceType Is Null)) ", True
 Exit Sub
ChkErr:
  ErrSub Me.Name, "LoadCom"
End Sub

Private Sub LoadGrid()
On Error GoTo ChkErr
  Dim mFlds As New prjGiGridR.GiGridFlds
  
  '畫面的 Grid 元件欄位設定
  mFlds.Add "CItemCode1", , , , , "收費項目代碼(+)"
  mFlds.Add "CItemName1", , , , , "收費項目名稱(+)"
  mFlds.Add "TaxName1", , , , , "稅率(+)"
  mFlds.Add "CItemCode2", , , , , "收費項目代碼(-)"
  mFlds.Add "CItemName2", , , , , "收費項目名稱(-)"
  mFlds.Add "TaxName2", , , , , "稅率(-)"
  mFlds.Add "CpCItemCode", , , , , "電信帳單項目代碼"
  mFlds.Add "CpCItemName", , , , , "電信帳單項目名稱"
  mFlds.Add "CpAdjCItemCode", , , , , "電信帳單調改項目代碼"
  mFlds.Add "CpGroupCode", , , , , "電信帳單群組代碼"
  mFlds.Add "CpGroupName", , , , , "電信帳單群組名稱"
  mFlds.Add "CpItem", , , , , "電信列印順序"
  mFlds.Add "TfnHead", , , , , "收費項目前導字串"
  ggrData.AllFields = mFlds
  ggrData.SetHead
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "LoadGrid")
End Sub

Private Sub ChangeMode(lngMode As giEditModeEnu)
On Error GoTo ChkErr
  Dim blnViewMode As Boolean
  
  '將編輯模式儲存起來
  lngEditMode = lngMode
  
  '切換模式
  Select Case lngMode
    Case giEditModeView
      blnViewMode = True
      lblEditMode.Caption = "顯示"
      cmdCancel.Caption = "結束(&X)"
    Case giEditModeEdit
      blnViewMode = False
      lblEditMode.Caption = "修改"
      cmdCancel.Caption = "取消(&X)"
    Case giEditModeInsert
      blnViewMode = False
      lblEditMode.Caption = "新增"
      cmdCancel.Caption = "取消(&X)"
  End Select
  
  '根據編輯模式來設定畫面的按鈕狀態
  ggrData.Enabled = blnViewMode
  FraData.Enabled = Not blnViewMode
  cmdSave.Enabled = Not blnViewMode
  cmdAdd.Enabled = blnViewMode
  cmdEdit.Enabled = (blnViewMode) And (Not rsSo165.EOF)
  cmdPrint.Enabled = blnViewMode And (Not rsSo165.EOF)
  Call cboCpProperty_Click
  cboCpProperty.SetFocus
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "ChangeMode")
End Sub

Private Sub OpenData()
On Error GoTo ChkErr
  Dim strsql As String
  
  strsql = "Select CItemCode1, CItemName1, TaxCode1, TaxName1, CItemCode2, CItemName2, TaxCode2, " & _
           "       TaxName2, CpCItemCode, CpCItemName, CpAdjCItemCode, CpGroupCode, CpGroupName, " & _
           "       CpItem, CpTax, CpProperty, TfnHead, RowId " & _
           "From " & GetOwner & "So165"
  Call OpenRec2(rsSo165, ggrData, strsql)
  Exit Sub
ChkErr:
   ErrSub Me.Name, "OpenData"
End Sub

Private Sub RcdToScr()
On Error GoTo ChkErr
  Dim strTmpString As String
    
  '讀出資料，填入畫面元件中
  If Not rsSo165.EOF Then
    gilCItem1.SetCodeNo rsSo165("CItemCode1") & ""
    gilCItem1.Query_Description
    gilTax1.SetCodeNo rsSo165("TaxCode1") & ""
    gilTax1.Query_Description
    gilCItem2.SetCodeNo rsSo165("CItemCode2") & ""
    gilCItem2.Query_Description
    gilTax2.SetCodeNo rsSo165("TaxCode2") & ""
    gilTax2.Query_Description
    txtCpCItemCode.Text = rsSo165("CpCItemCode") & ""
    txtCpCItemName.Text = rsSo165("CpCItemName") & ""
    txtCpAdjCItemCode.Text = rsSo165("CpAdjCItemCode") & ""
    txtCpGroupCode.Text = rsSo165("CpGroupCode") & ""
    txtCpGroupName.Text = rsSo165("CpGroupName") & ""
    txtCpItem.Text = rsSo165("CpItem") & ""
    
    strTmpString = rsSo165("CpTax") & ""
    Select Case strTmpString
      Case "Y"
        cboCpTax.ListIndex = 0
      Case "N"
        cboCpTax.ListIndex = 1
      Case "F"
        cboCpTax.ListIndex = 2
      Case Else
        cboCpTax.ListIndex = -1
    End Select
    
    strTmpString = rsSo165("CpProperty") & ""
    If strTmpString <= 4 Then
      cboCpProperty.ListIndex = strTmpString - 1
    Else
      cboCpProperty.ListIndex = -1
    End If
    
    txtTFNHead.Text = rsSo165("TFNHead") & ""
  End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScr")
End Sub

Private Sub GiExportDate1_Change()
On Error GoTo ChkErr
  Call SetExporFileName(txtExportPath1, txtExportFile1, txtLogFile4, lblExportFile1, GiExportDate1)
ChkErr:
  Call ErrSub(Me.Name, "GiExportDate1_Change")
End Sub

Private Sub SetExporFileName(ByRef PathArea, NameArea, LogArea As TextBox, ByRef LabelArea As Label, ByRef DateArea As GiDate)
On Error GoTo ChkErr
  Dim strString1 As String
  
  strString1 = PathArea.Text
  If strString1 <> "" Then
    strString1 = strString1 & "\"
  End If
  strString1 = Replace(strString1, "\\", "\")
  LabelArea.Caption = NameArea.Text & DateArea.GetValue & ".TXT"
  Call SetLogFileName(strString1 & LabelArea.Caption, LogArea)
ChkErr:
  Call ErrSub(Me.Name, "SetExporFileName")
End Sub

Private Sub GiExportDate3_Change()
On Error GoTo ChkErr
  Call SetExporFileName(txtExportPath2, txtExportFile2, txtLogFile5, lblExportFile2, GiExportDate3)
ChkErr:
  Call ErrSub(Me.Name, "GiExportDate3_Change")
End Sub

Private Sub GiExportDate4_Change()
On Error GoTo ChkErr
  lblExportFile3.Caption = txtExportFile3 & GiExportDate4.GetValue & ".TXT"
ChkErr:
  Call ErrSub(Me.Name, "GiExportDate4_Change")
End Sub

Private Sub GiExportDate5_Change()
On Error GoTo ChkErr
  lblExportFile4.Caption = txtExportFile4 & GiExportDate5.GetValue & ".TXT"
ChkErr:
  Call ErrSub(Me.Name, "GiExportDate5_Change")
End Sub

Private Sub GiExportDate6_Change()
On Error GoTo ChkErr
  Call SetExporFileName(txtExportPath5, txtExportFile5, txtLogFile6, lblExportFile5, GiExportDate6)
ChkErr:
  Call ErrSub(Me.Name, "GiExportDate6_Change")
End Sub


Private Sub GiRealDate3_GotFocus()
  On Error Resume Next
  If GiRealDate3.GetValue = "" Then GiRealDate3.SetValue (RightDate)
End Sub
Private Sub GiRealDate4_GotFocus()
  On Error Resume Next
  If GiRealDate4.GetValue = "" Or GiRealDate4.GetValue = "" Then GiRealDate4.SetValue (GiRealDate3.GetValue)
End Sub

Private Sub GiRealDate4_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(GiRealDate3, GiRealDate4)
End Sub
Private Sub RsSo165_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo ChkErr
  If Not rsSo165.EOF Then
    If (lngEditMode = giEditModeView) Or (lngEditMode = giEditModeEdit) Then
      RcdToScr
    End If
  End If
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "RsSo165_MoveComplete")
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo ChkErr
  If (SSTab1.Tab <> 0) And (lngEditMode <> giEditModeView) Then
    SSTab1.Tab = PreviousTab
    MsgBox "請先確定存檔或取消存檔之後，再切換頁面。", vbExclamation, "警告"
  End If
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "SSTab1_Click")
End Sub

Private Sub NewRcd()
On Error GoTo ChkErr
  '新增資料須先將 Client Dataset 的欄位清空
  gilCItem1.SetCodeNo ""
  gilCItem1.Query_Description
  gilTax1.SetCodeNo ""
  gilTax1.Query_Description
  gilCItem2.SetCodeNo ""
  gilCItem2.Query_Description
  gilTax2.SetCodeNo ""
  gilTax2.Query_Description
  txtCpCItemCode.Text = ""
  txtCpCItemName.Text = ""
  txtCpAdjCItemCode.Text = ""
  txtCpGroupCode.Text = ""
  txtCpGroupName.Text = ""
  txtCpItem.Text = ""
  cboCpTax.ListIndex = -1
  cboCpProperty.ListIndex = -1
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "NewRcd")
End Sub

Private Sub Text1_Change()
End Sub

Private Sub txtCpCItemCode_LostFocus()
On Error GoTo ChkErr
  If txtCpAdjCItemCode.Text = "" Then
    txtCpAdjCItemCode.Text = txtCpCItemCode.Text
  End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtCpCItemCode_LostFocus")
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo ChkErr
  Dim blnGoOn As Boolean
  Dim strErrMsg, strsql As String
  Dim rsCD019 As New ADODB.Recordset
  
  blnGoOn = True
  If (blnGoOn) And (gilCItem1.GetCodeNo <> "") And (gilTax1.GetCodeNo = "") Then
    strErrMsg = "稅率 (借貸 : + ) 須有值"
    gilTax1.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And (gilCItem2.GetCodeNo <> "") And (gilTax2.GetCodeNo = "") Then
    strErrMsg = "稅率 (借貸 : - ) 須有值"
    gilTax2.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And (gilCItem1.GetCodeNo = "") And (gilCItem2.GetCodeNo = "") Then
    strErrMsg = "收費項目 (借貸 : + ) 與 收費項目 (借貸 : - ) 至少其中一個須有值"
    gilCItem1.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And (gilCItem1.GetCodeNo <> "") And (gilCItem2.GetCodeNo <> "") And (gilTax1.GetCodeNo <> gilTax2.GetCodeNo) Then
    strErrMsg = "收費項目(借貸:+) 與 收費項目(借貸:- ) 皆有值時,其稅率須相同"
    gilCItem1.SetFocus
    blnGoOn = False
  End If
  
  If blnGoOn Then
    If gilCItem1.GetCodeNo <> "" Then
      Set rsCD019 = gcnGi.Execute("Select PeriodFlag From " & GetOwner & "Cd019 Where CodeNo=" & gilCItem1.GetCodeNo)
      If Not rsCD019.EOF Then
        If ((cboCpProperty.ListIndex = 0) Or (cboCpProperty.ListIndex = 3)) And (rsCD019("PeriodFlag") <> "0") Then
          strErrMsg = "當[電信項目屬性]=1或4時,收費項目 (借貸 : + )須為非週期性項目"
          gilCItem1.SetFocus
          blnGoOn = False
        ElseIf ((cboCpProperty.ListIndex = 1) Or (cboCpProperty.ListIndex = 2)) And (rsCD019("PeriodFlag") <> "1") Then
          strErrMsg = "當[電信項目屬性]=2或3時,收費項目 (借貸 : + )須為週期性項目"
          gilCItem1.SetFocus
          blnGoOn = False
        End If
      Else
        strErrMsg = "找不到收費項目：" & gilCItem1.GetCodeNo
        gilCItem1.SetFocus
        blnGoOn = False
      End If
      Set rsCD019 = Nothing
    End If
    If (blnGoOn) And (gilCItem2.GetCodeNo <> "") Then
      Set rsCD019 = gcnGi.Execute("Select PeriodFlag From " & GetOwner & "Cd019 Where CodeNo=" & gilCItem2.GetCodeNo)
      If Not rsCD019.EOF Then
        If ((cboCpProperty.ListIndex = 0) Or (cboCpProperty.ListIndex = 3)) And (rsCD019("PeriodFlag") <> "0") Then
          strErrMsg = "當[電信項目屬性]=1或4時,收費項目 (借貸 : - )須為非週期性項目"
          gilCItem2.SetFocus
          blnGoOn = False
        ElseIf ((cboCpProperty.ListIndex = 1) Or (cboCpProperty.ListIndex = 2)) And (rsCD019("PeriodFlag") <> "1") Then
          strErrMsg = "當[電信項目屬性]=2或3時,收費項目 (借貸 : - )須為週期性項目"
          gilCItem2.SetFocus
          blnGoOn = False
        End If
      Else
        strErrMsg = "找不到收費項目：" & gilCItem2.GetCodeNo
        gilCItem1.SetFocus
        blnGoOn = False
      End If
      Set rsCD019 = Nothing
    End If
  End If
  
  If (blnGoOn) And (gilCItem1.GetCodeNo <> "") Then
    If Not CheckRepeatData(gilCItem1.GetCodeNo, "CItemCode1") Then
      strErrMsg = "收費項目 (借貸 : + ) 資料重複"
      gilCItem1.SetFocus
      blnGoOn = False
    End If
  End If
  
  If (blnGoOn) And (gilCItem2.GetCodeNo <> "") Then
    If Not CheckRepeatData(gilCItem2.GetCodeNo, "CItemCode2") Then
      strErrMsg = "收費項目 (借貸 : - ) 資料重複"
      gilCItem2.SetFocus
      blnGoOn = False
    End If
  End If
  
  If (blnGoOn) And (txtCpCItemCode.Text = "") Then
    strErrMsg = "電信帳單項目代碼須有值"
    txtCpCItemCode.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And ((cboCpProperty.ListIndex = 0)) And (txtCpCItemCode.Text <> "") Then
    If Not CheckRepeatData("'" & txtCpCItemCode.Text & "'", "CpCItemCode") Then
      strErrMsg = "電信帳單項目代碼資料重複"
      txtCpCItemCode.SetFocus
      blnGoOn = False
    End If
  End If
  
  If (blnGoOn) And (txtCpCItemName.Text = "") Then
    strErrMsg = "電信帳單項目代碼須有值"
    txtCpCItemName.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And ((cboCpProperty.ListIndex = 0) Or (cboCpProperty.ListIndex = 3)) And (txtCpCItemName.Text <> "") Then
    If Not CheckRepeatData("'" & txtCpCItemName.Text & "'", "CpCItemName") Then
      strErrMsg = "電信帳單項目名稱資料重複"
      txtCpCItemName.SetFocus
      blnGoOn = False
    End If
  End If
  
  'If (blnGoOn) And (txtCpAdjCItemCode.Text = "") Then
  '  strErrMsg = "電信帳單調改項目代碼須有值"
  '  txtCpAdjCItemCode.SetFocus
  '  blnGoOn = False
  'End If
  
  If (blnGoOn) And (txtCpGroupCode.Text = "") Then
    strErrMsg = "電信帳單群組代碼須有值"
    txtCpGroupCode.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And (txtCpGroupName.Text = "") Then
    strErrMsg = "電信帳單群組名稱須有值"
    txtCpGroupName.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And (txtCpItem.Text = "") Then
    strErrMsg = "電信列印順序須有值"
    txtCpItem.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And (txtCpItem.Text <> "") Then
    If Not CheckRepeatData("'" & txtCpItem.Text & "'", "CpItem") Then
      strErrMsg = "電信列印順序資料重複"
      txtCpItem.SetFocus
      blnGoOn = False
    End If
  End If
  
  If (blnGoOn) And (cboCpTax.ListIndex < 0) Then
    strErrMsg = "電信稅率須有值或資料不正確"
    cboCpTax.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And (cboCpProperty.ListIndex < 0) Then
    strErrMsg = "電信項目屬性須有值或資料不正確"
    cboCpProperty.SetFocus
    blnGoOn = False
  End If
  
  If Not blnGoOn Then
    MsgBox strErrMsg, vbInformation, "訊息"
  End If
  
  IsDataOk = blnGoOn
  
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub ScrToRcd()
On Error GoTo ChkErr
  '將畫面元件的值填入 Client Dataset 的欄位中
  
  Dim strString1 As String
 
  '新增模式時要先 Add 一筆空的資料列
  If lngEditMode = giEditModeInsert Then
    rsSo165.AddNew
  End If
  
  rsSo165("CItemCode1") = NoZero(gilCItem1.GetCodeNo)
  rsSo165("CItemName1") = NoZero(gilCItem1.GetDescription)
  rsSo165("TaxCode1") = NoZero(gilTax1.GetCodeNo)
  rsSo165("TaxName1") = NoZero(gilTax1.GetDescription)
  rsSo165("CItemCode2") = NoZero(gilCItem2.GetCodeNo)
  rsSo165("CItemName2") = NoZero(gilCItem2.GetDescription)
  rsSo165("TaxCode2") = NoZero(gilTax2.GetCodeNo)
  rsSo165("TaxName2") = NoZero(gilTax2.GetDescription)
  rsSo165("CpCItemCode") = txtCpCItemCode.Text
  rsSo165("CpCItemName") = txtCpCItemName.Text
  rsSo165("CpAdjCItemCode") = txtCpAdjCItemCode.Text
  rsSo165("CpGroupCode") = txtCpGroupCode.Text
  rsSo165("CpGroupName") = txtCpGroupName.Text
  rsSo165("CpItem") = txtCpItem.Text
  rsSo165("CpTax") = Left(cboCpTax.List(cboCpTax.ListIndex), 1)
  rsSo165("CpProperty") = cboCpProperty.ListIndex + 1
  If txtTFNHead.Text = "" Then
    strString1 = "代收"
  Else
    strString1 = txtTFNHead.Text
  End If
  rsSo165("TfnHead") = strString1
  rsSo165.Update
  Call OpenData
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "ScrToRcd")
End Sub

Private Function CheckRepeatData(ByVal CheckData As String, ByVal CheckField As String) As Boolean
On Error GoTo ChkErr
  Dim rsTempSo165 As New ADODB.Recordset
  
  '新增或修改資料，存檔時要先判斷資料是否已存在了
  CheckRepeatData = True
  Set rsTempSo165 = gcnGi.Execute("Select RowId From " & GetOwner & "So165 Where " & CheckField & "=" & CheckData)
  If Not rsTempSo165.EOF Then
    '若資料已存在，且為新增模式，或為編輯模式，但是是不同的資料列，則表示資料重複了
    If (lngEditMode = giEditModeInsert) Or ((lngEditMode = giEditModeEdit) And (rsTempSo165("RowId") <> rsSo165("RowId"))) Then
      CheckRepeatData = False
    End If
  End If
  Set rsTempSo165 = Nothing
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "CheckRepeatData")
End Function

Private Sub SetLogFileName(ByVal SourceFilename As String, ByRef LogFileTextBox As TextBox)
On Error GoTo ChkErr
  Dim strFileName, strExtensionName As String
  
  ''自動根據文字檔的檔名設定對應的 LOG 檔檔名
  strFileName = FSO.GetFileName(SourceFilename)
  strExtensionName = FSO.GetExtensionName(SourceFilename)
  
  If strExtensionName <> "" Then
    strFileName = Left(strFileName, Len(strFileName) - Len(strExtensionName) - 1)
  End If
  strFileName = Replace(SourceFilename, strFileName, strFileName & "_ErrLog")
  LogFileTextBox.Text = strFileName
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "SetLogFileName")
End Sub

Private Sub txtExportFile1_Change()
On Error GoTo ChkErr
  Call SetExporFileName(txtExportPath1, txtExportFile1, txtLogFile4, lblExportFile1, GiExportDate1)
ChkErr:
  Call ErrSub(Me.Name, "txtExportFile1_Change")
End Sub

Private Sub txtExportFile2_Change()
On Error GoTo ChkErr
  Call SetExporFileName(txtExportPath2, txtExportFile2, txtLogFile5, lblExportFile2, GiExportDate3)
ChkErr:
  Call ErrSub(Me.Name, "txtExportFile2_Change")
End Sub

Private Sub txtExportFile3_Change()
On Error GoTo ChkErr
  Call GiExportDate4_Change
ChkErr:
  Call ErrSub(Me.Name, "txtExportFile3_Change")
End Sub

Private Sub txtExportFile4_Change()
On Error GoTo ChkErr
  Call GiExportDate5_Change
ChkErr:
  Call ErrSub(Me.Name, "txtExportFile4_Change")
End Sub

Private Sub txtExportFile5_Change()
On Error GoTo ChkErr
  Call SetExporFileName(txtExportPath5, txtExportFile5, txtLogFile6, lblExportFile5, GiExportDate6)
ChkErr:
  Call ErrSub(Me.Name, "txtExportFile5_Change")
End Sub

Private Sub txtExportPath1_Change()
On Error GoTo ChkErr
  Call SetExporFileName(txtExportPath1, txtExportFile1, txtLogFile4, lblExportFile1, GiExportDate1)
ChkErr:
  Call ErrSub(Me.Name, "txtExportPath1_Change")
End Sub

Private Sub txtExportPath1_LostFocus()
On Error GoTo ChkErr
  If (txtExportPath1.Text <> "") And (Not FSO.FolderExists(txtExportPath1.Text)) Then
    MsgBox txtExportPath1.Text & vbCrLf & vbCrLf & "路徑不存在。", vbExclamation, "警告"
    txtExportPath1.Text = ""
    SSTab3.Tab = 0
    txtExportPath1.SetFocus
  End If
ChkErr:
  Call ErrSub(Me.Name, "txtExportPath1_LostFocus")
End Sub

Private Sub txtExportPath2_Change()
On Error GoTo ChkErr
  Call SetExporFileName(txtExportPath2, txtExportFile2, txtLogFile5, lblExportFile2, GiExportDate3)
ChkErr:
  Call ErrSub(Me.Name, "txtExportPath2_Change")
End Sub

Private Sub txtExportPath2_LostFocus()
On Error GoTo ChkErr
  If (txtExportPath2.Text <> "") And (Not FSO.FolderExists(txtExportPath2.Text)) Then
    MsgBox txtExportPath2.Text & vbCrLf & vbCrLf & "路徑不存在。", vbExclamation, "警告"
    txtExportPath2.Text = ""
    SSTab3.Tab = 1
    txtExportPath2.SetFocus
  End If
ChkErr:
  Call ErrSub(Me.Name, "txtExportPath2_LostFocus")
End Sub

Private Sub txtExportPath3_LostFocus()
On Error GoTo ChkErr
  If (txtExportPath3.Text <> "") And (Not FSO.FolderExists(txtExportPath3.Text)) Then
    MsgBox txtExportPath3.Text & vbCrLf & vbCrLf & "路徑不存在。", vbExclamation, "警告"
    txtExportPath3.Text = ""
    SSTab3.Tab = 2
    txtExportPath3.SetFocus
  End If
ChkErr:
  Call ErrSub(Me.Name, "txtExportPath3_LostFocus")
End Sub

Private Sub txtExportPath4_LostFocus()
On Error GoTo ChkErr
  If (txtExportPath4.Text <> "") And (Not FSO.FolderExists(txtExportPath4.Text)) Then
    MsgBox txtExportPath4.Text & vbCrLf & vbCrLf & "路徑不存在。", vbExclamation, "警告"
    txtExportPath4.Text = ""
    SSTab3.Tab = 3
    txtExportPath4.SetFocus
  End If
ChkErr:
  Call ErrSub(Me.Name, "txtExportPath4_LostFocus")
End Sub

Private Sub txtExportPath5_Change()
On Error GoTo ChkErr
  Call SetExporFileName(txtExportPath5, txtExportFile5, txtLogFile6, lblExportFile5, GiExportDate6)
ChkErr:
  Call ErrSub(Me.Name, "txtExportPath5_Change")
End Sub

Private Sub txtExportPath5_LostFocus()
On Error GoTo ChkErr
  If (txtExportPath5.Text <> "") And (Not FSO.FolderExists(txtExportPath5.Text)) Then
    MsgBox txtExportPath5.Text & vbCrLf & vbCrLf & "路徑不存在。", vbExclamation, "警告"
    txtExportPath5.Text = ""
    SSTab3.Tab = 4
    txtExportPath5.SetFocus
  End If
ChkErr:
  Call ErrSub(Me.Name, "txtExportPath5_LostFocus")
End Sub

Private Sub txtImportFile1_LostFocus()
On Error GoTo ChkErr
  Call SetLogFileName(txtImportFile1.Text, txtLogFile1)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtImportFile1_LostFocus")
End Sub

Private Sub LogError1(ByRef ErrorCounts As Long, ByVal Reason As String, ByVal RightValue As String, ByVal LineNumber, ByVal LineText As String)
On Error GoTo ChkErr
  '記錄錯誤的資料到 LOG 檔中
  Dim strString2 As String
  Dim strString1 As String

  '行號
  strString1 = ""
  If LineNumber > 0 Then
    strString1 = Right("00000" & LineNumber, 5) & "："
  End If
  
  strString2 = ""
  strString1 = strString1 & Reason & Space(90)
  strString1 = StrConv(LeftB(StrConv(strString1, vbFromUnicode), 90), vbUnicode)
  If RightValue <> "" Then
    strString2 = "正確資料：" & RightValue & Space(60)
    strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 60), vbUnicode)
  End If
  strmTextFile1.Write strString1 & strString2 & vbCrLf
  'strmTextFile1.Write "       " & LineText & vbCrLf
  ErrorCounts = ErrorCounts + 1
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "LogError1")
End Sub

Private Sub LogError3(ByRef ErrorCounts As Long, ByVal Reason As String)
On Error GoTo ChkErr
  '記錄錯誤的資料到 LOG 檔中
  Dim strString2 As String
  Dim strString1 As String
  strString1 = "原因：" & Reason
  strmTextFile2.Write strString1 & vbCrLf
  ErrorCounts = ErrorCounts + 1
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "LogError3")
End Sub

Private Sub ExportData1()
On Error GoTo ChkErr
  'CP通信費銷帳資料匯出
  Dim lngLogCount As Long
  Dim lngExportDetailCount As Long, lngExportMasterCount As Long, lngDataCountTail As Long, lngTmp As Long
  Dim dblRealAmtMaster, dblRealAmtWithTaxMaster  As Double, dblRealAmtWithoutTaxMaster  As Double, dblAdjustAmtTail As Double
  Dim dblRealAmtZeroTaxMaster  As Double, dblAdjustAmtMaster  As Double, dblAdjustAmtWithTaxMaster  As Double, dblAmtTail As Double
  Dim dblAdjustAmtZeroTaxMaster  As Double, dblAdjustAmtWithoutTaxMaster  As Double, dblShouldAmtMaster As Double
  Dim dblTotalRealAmtTail  As Double, dblTotalAdjustAmtTail As Double
  Dim strMaster, strDetail, strRealDateMaster, strTail As String
  Dim strString1 As String
  Dim strCustid, strFaciSNo, strTfnServiceId, strBillNo, strTFNBillNo As String
  Dim rsTemp As New ADODB.Recordset
  Dim strLog As String
  
  Screen.MousePointer = vbHourglass
  Me.Enabled = False
  
  'strString1 = "Select * From " & GetOwner & "So166 " & _
               "Where (CompCode=" & garyGi(9) & ") " & _
               "And (Servicetype='P') " & _
               "And (RealDate Between To_Date('" & GiRealDate1.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate2.GetValue & "', 'YYYYMMDD')) " & _
               "And (CpExport=0) " & _
               "Order By CustId, FaciSNo, TfnServiceId, BillNo, TfnBillNo "
  '本來已匯出過的資料不再匯出，現在改了
  strString1 = "Select * From " & GetOwner & "So166 " & _
               "Where (CompCode=" & garyGi(9) & ") " & _
               "And (Servicetype='P') " & _
               "And (RealDate Between To_Date('" & GiRealDate1.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate2.GetValue & "', 'YYYYMMDD')) " & _
               "Order By CustId, FaciSNo, TfnServiceId, BillNo, TfnBillNo "
  Set rsTemp = gcnGi.Execute(strString1)
  
  '寫入檔頭
'  If Not rsTemp.EOF Then
    strmTextFile1.Write "H" & GiExportDate1.GetValue & vbCrLf
'  End If
  
  lngExportMasterCount = 0
  lngExportDetailCount = 0
  lngLogCount = 0
  strCustid = ""
  strFaciSNo = ""
  strTfnServiceId = ""
  strBillNo = ""
  strTFNBillNo = ""
  
  '寫入表身
  While Not rsTemp.EOF
    If (strCustid <> rsTemp("CustId") & "") Or (strFaciSNo <> rsTemp("FaciSNo") & "") Or (strTfnServiceId <> rsTemp("TfnServiceId") & "") Or (strBillNo <> rsTemp("BillNo") & "") Or (strTFNBillNo <> rsTemp("TfnBillNo") & "") Then
      '依 CustId, FaciSNo, TfnServiceId, BillNo, TfnBillNo 分組產生 Master
      If (rsTemp("ShouldAmt") - rsTemp("RealAmt")) >= 0 Then

      If strCustid <> "" Then
        strMaster = "M" & strMaster
        strMaster = strMaster & Left(strCustid & Space(10), 10)
        
        '針對電話號碼處理，因取出的電話區碼和號碼間有 - 符號，匯出時要濾掉
        '- 的位置
        If Len(strFaciSNo) > 0 Then
          lngTmp = InStr(strFaciSNo, "-")
      
          '判斷有沒有區碼
          If lngTmp > 0 Then
            strString1 = Left(strFaciSNo, lngTmp - 1)
            strMaster = strMaster & Left(strString1 & Space(3), 3)
          Else
            '沒有，填空白
            strMaster = strMaster & Space(3)
          End If
      
          '號碼，取 10 碼
          strString1 = Mid(strFaciSNo, lngTmp + 1, 10)
          strMaster = strMaster & Left(strString1 & Space(10), 10)
        Else
          strMaster = strMaster & Space(13)
        End If
        
        'strMaster = strMaster & Left(strFaciSNo & Space(13), 13)
        strMaster = strMaster & Left(strTfnServiceId & Space(10), 10)
        strMaster = strMaster & Left(strBillNo & Space(16), 16)
        strMaster = strMaster & Left(strTFNBillNo & Space(13), 13)
        strMaster = strMaster & MakeMoney(dblRealAmtMaster)
        strMaster = strMaster & MakeMoney(dblRealAmtWithTaxMaster)
        strMaster = strMaster & MakeMoney(dblRealAmtZeroTaxMaster)
        strMaster = strMaster & MakeMoney(dblRealAmtWithoutTaxMaster)
        strMaster = strMaster & MakeMoney(dblAdjustAmtMaster)
        strMaster = strMaster & MakeMoney(dblAdjustAmtWithTaxMaster)
        strMaster = strMaster & MakeMoney(dblAdjustAmtZeroTaxMaster)
        strMaster = strMaster & MakeMoney(dblAdjustAmtWithoutTaxMaster)
        strMaster = strMaster & strRealDateMaster
        strString1 = Space(50)
        If dblRealAmtMaster <> dblShouldAmtMaster Then
          strString1 = "有短收" & strString1
          strString1 = StrConv(LeftB(StrConv(strString1, vbFromUnicode), 50), vbUnicode)
        End If
        
        strMaster = strMaster & strString1
        
        '將 Buffer 中的資料匯出
        strmTextFile1.Write strMaster & vbCrLf
        strmTextFile1.Write strDetail
        
        strMaster = ""
        strDetail = ""
      End If
      '將產生一筆新的 Master 資料，將相關欄位先清空及填入預設值
      dblRealAmtMaster = 0
      dblShouldAmtMaster = 0
      dblRealAmtWithTaxMaster = 0
      dblRealAmtZeroTaxMaster = 0
      dblRealAmtWithoutTaxMaster = 0
      dblAdjustAmtMaster = 0
      dblAdjustAmtWithTaxMaster = 0
      dblAdjustAmtZeroTaxMaster = 0
      dblAdjustAmtWithoutTaxMaster = 0
      strCustid = rsTemp("CustId") & ""
      strFaciSNo = rsTemp("FaciSNo") & ""
      strTfnServiceId = rsTemp("TfnServiceId") & ""
      strBillNo = rsTemp("BillNo") & ""
      strTFNBillNo = rsTemp("TfnBillNo") & ""
      GiTempDate.SetValue (rsTemp("RealDate"))
      strRealDateMaster = Left(GiTempDate.GetValue, 6)
      lngExportMasterCount = lngExportMasterCount + 1
    End If
    End If

    dblRealAmtMaster = dblRealAmtMaster + rsTemp("RealAmt")
    dblShouldAmtMaster = dblShouldAmtMaster + rsTemp("ShouldAmt")
    
    If rsTemp("CpRate") = "Y" Then
      dblRealAmtWithTaxMaster = dblRealAmtWithTaxMaster + rsTemp("RealAmt")
    End If
    
    If rsTemp("CpRate") = "N" Then
      dblRealAmtZeroTaxMaster = dblRealAmtZeroTaxMaster + rsTemp("RealAmt")
    End If
    
    If rsTemp("CpRate") = "F" Then
      dblRealAmtWithoutTaxMaster = dblRealAmtWithoutTaxMaster + rsTemp("RealAmt")
    End If
    
    dblAdjustAmtMaster = dblAdjustAmtMaster + rsTemp("ShouldAmt") - rsTemp("RealAmt")
    
    If rsTemp("CpRate") = "Y" Then
      dblAdjustAmtWithTaxMaster = dblAdjustAmtWithTaxMaster + (rsTemp("ShouldAmt") - rsTemp("RealAmt"))
    End If
    
    If rsTemp("CpRate") = "N" Then
      dblAdjustAmtZeroTaxMaster = dblAdjustAmtZeroTaxMaster + (rsTemp("ShouldAmt") - rsTemp("RealAmt"))
    End If
    
    If rsTemp("CpRate") = "F" Then
      dblAdjustAmtWithoutTaxMaster = dblAdjustAmtWithoutTaxMaster + (rsTemp("ShouldAmt") - rsTemp("RealAmt"))
    End If
    '有調改項目
    If (rsTemp("ShouldAmt") - rsTemp("RealAmt")) >= 0 Then
        If rsTemp("ShouldAmt") <> rsTemp("RealAmt") Then
           strDetail = strDetail & "D" & Left(rsTemp("CpAdjCitemCode") & Space(10), 10)
           strDetail = strDetail & Left(rsTemp("CpRate") & Space(1), 1)
           '13-20 調改費用
           strDetail = strDetail & MakeMoney(rsTemp("ShouldAmt") - rsTemp("RealAmt"))
           
           If rsTemp("CpRate") = "Y" Then
             strDetail = strDetail & MakeMoney(Round((rsTemp("ShouldAmt") - rsTemp("RealAmt")) * 0.95))
           Else
             strDetail = strDetail & MakeMoney(rsTemp("ShouldAmt") - rsTemp("RealAmt"))
           End If
           
           If rsTemp("CpRate") = "Y" Then
             strDetail = strDetail & MakeMoney((rsTemp("ShouldAmt") - rsTemp("RealAmt")) - (Round((rsTemp("ShouldAmt") - rsTemp("RealAmt")) * 0.95)))
           Else
             strDetail = strDetail & MakeMoney(0)
           End If
           
           strDetail = strDetail & Left(rsTemp("CpAdjCode") & Space(4), 4)
           GiTempDate.SetValue (rsTemp("RealDate"))
           strDetail = strDetail & Left(GiTempDate.GetValue, 6)
           strString1 = rsTemp("StName") & Space(50)
           strString1 = StrConv(LeftB(StrConv(strString1, vbFromUnicode), 50), vbUnicode)
           strDetail = strDetail & strString1
           strDetail = strDetail & vbCrLf
           lngExportDetailCount = lngExportDetailCount + 1
         End If
        
         '已匯出過的要 LOG
         If rsTemp("CpExport") <> 0 Then
           strString1 = "客編(" & rsTemp("CustId") & ")，" & _
                        "CP門號(" & rsTemp("FaciSNo") & ")，" & _
                        "TFN服務代碼(" & rsTemp("TfnServiceId") & ")，" & _
                        "單據編號(" & rsTemp("BillNo") & ")，" & _
                        "TFN帳單編號(" & rsTemp("TfnBillNo") & ")，" & _
                        "實收日期(" & rsTemp("RealDate") & ")，" & _
                        "實收金額(" & rsTemp("RealAmt") & ")，" & _
                        "此筆資料已匯出過"
           Call LogError3(lngLogCount, strString1)
         End If
    Else
           strString1 = "客編(" & rsTemp("CustId") & ")，" & _
                        "CP門號(" & rsTemp("FaciSNo") & ")，" & _
                        "TFN服務代碼(" & rsTemp("TfnServiceId") & ")，" & _
                        "單據編號(" & rsTemp("BillNo") & ")，" & _
                        "TFN帳單編號(" & rsTemp("TfnBillNo") & ")，" & _
                        "實收日期(" & rsTemp("RealDate") & ")，" & _
                        "實收金額(" & rsTemp("RealAmt") & ")，" & _
                        "應收金額(" & rsTemp("ShouldAmt") & ")，" & _
                        "此筆應收金額少於實收金額"
        Call LogError3(lngLogCount, strString1)
    End If
    rsTemp.MoveNext
  Wend
  '全部跑完後須將 Buffer 中尚未匯出的 Master 及 Detail 資料也匯出
  If dblAdjustAmtMaster >= 0 Then
  If strCustid <> "" Then
    strMaster = "M" & strMaster
    strMaster = strMaster & Left(strCustid & Space(10), 10)
    
    '針對電話號碼處理，因取出的電話區碼和號碼間有 - 符號，匯出時要濾掉
    '- 的位置
    If Len(strFaciSNo) > 0 Then
      lngTmp = InStr(strFaciSNo, "-")
      
      '判斷有沒有區碼
      If lngTmp > 0 Then
        strString1 = Left(strFaciSNo, lngTmp - 1)
        strMaster = strMaster & Left(strString1 & Space(3), 3)
      Else
        '沒有，填空白
        strMaster = strMaster & Space(3)
      End If
      
      '號碼，取 10 碼
      strString1 = Mid(strFaciSNo, lngTmp + 1, 10)
      strMaster = strMaster & Left(strString1 & Space(10), 10)
    Else
      strMaster = strMaster & Space(13)
    End If
        
    'strMaster = strMaster & Left(strFaciSNo & Space(13), 13)
    strMaster = strMaster & Left(strTfnServiceId & Space(10), 10)
    strMaster = strMaster & Left(strBillNo & Space(16), 16)
    strMaster = strMaster & Left(strTFNBillNo & Space(13), 13)
    strMaster = strMaster & MakeMoney(dblRealAmtMaster)
    strMaster = strMaster & MakeMoney(dblRealAmtWithTaxMaster)
    strMaster = strMaster & MakeMoney(dblRealAmtZeroTaxMaster)
    strMaster = strMaster & MakeMoney(dblRealAmtWithoutTaxMaster)
    strMaster = strMaster & MakeMoney(dblAdjustAmtMaster)
    strMaster = strMaster & MakeMoney(dblAdjustAmtWithTaxMaster)
    strMaster = strMaster & MakeMoney(dblAdjustAmtZeroTaxMaster)
    strMaster = strMaster & MakeMoney(dblAdjustAmtWithoutTaxMaster)
    strMaster = strMaster & strRealDateMaster
    strString1 = Space(50)
    If dblRealAmtMaster <> dblShouldAmtMaster Then
      strString1 = "有短收" & strString1
      strString1 = StrConv(LeftB(StrConv(strString1, vbFromUnicode), 50), vbUnicode)
    End If
    
    strMaster = strMaster & strString1
    
    '將 Buffer 中的資料匯出
    strmTextFile1.Write strMaster & vbCrLf
    strmTextFile1.Write strDetail
    
    strMaster = ""
    strDetail = ""
  End If
  Set rsTemp = Nothing
  End If
  '取得依 TFNBillNo 分組的筆數
  strString1 = "Select Count(Distinct TFnBillNo) As DataCount, " & _
               "       Sum(RealAmt) As TotalRealAmt, " & _
               "       Sum(ShouldAmt)-Sum(RealAmt) As TotalAdjustAmt " & _
               "From " & GetOwner & "So166 " & _
               "Where (CompCode=" & garyGi(9) & ") " & _
               "And (Servicetype='P') " & _
               "And (Realdate Between To_Date('" & GiRealDate1.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate2.GetValue & "', 'YYYYMMDD')) " & _
               " AND (ShouldAmt-RealAmt) >=0"
               '"And (CpExport=0) "
  Set rsTemp = gcnGi.Execute(strString1)
  lngDataCountTail = rsTemp("DataCount")
'  If rsTemp("TotalAdjustAmt") > 0 Then
  '匯入表尾的部分
  dblTotalRealAmtTail = NoZero(rsTemp("TotalRealAmt"), True)
  dblTotalAdjustAmtTail = NoZero(rsTemp("TotalAdjustAmt"), True)
  strTail = "T" & Right("000000" & lngDataCountTail, 6)
  strTail = strTail & Right("000000" & lngExportDetailCount, 6)
  strTail = strTail & Right("000000" & lngDataCountTail + lngExportDetailCount, 6) 'Right("000000" & lngExportMasterCount, 6)
  strTail = strTail & MakeMoney(dblTotalRealAmtTail)
  strTail = strTail & MakeMoney(dblTotalAdjustAmtTail)
  Set rsTemp = Nothing
  strmTextFile1.Write strTail
  '已匯出的資料下次不再匯出
  strString1 = "Update " & GetOwner & "So166 " & _
               "Set CpExport=1 " & _
               "Where (CompCode=" & garyGi(9) & ") " & _
               "And (Servicetype='P') " & _
               "And (RealDate Between To_Date('" & GiRealDate1.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate2.GetValue & "', 'YYYYMMDD')) " & _
               "And (CpExport=0) "
  gcnGi.Execute (strString1)
'  End If
  
  If lngLogCount > 0 Then
    strString1 = "匯出Master筆數：" & lngExportMasterCount & vbCrLf & _
                 "匯出Detail筆數：" & lngExportDetailCount & vbCrLf & vbCrLf & _
                 "資料中有問題的資料存在，筆數：" & lngLogCount & vbCrLf & _
                 "請參考問題記錄檔。" & vbCrLf
    MsgBox strString1, vbExclamation, "訊息"
    Shell "Notepad " & txtLogFile4.Text, vbNormalFocus
'    strmTextFile2
  Else
    strString1 = "匯出完成 !!" & vbCrLf & vbCrLf & _
                 "匯出Master筆數：" & lngExportMasterCount & vbCrLf & _
                 "匯出Detail筆數：" & lngExportDetailCount
    MsgBox strString1, vbExclamation, "訊息"
  End If
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Call ErrSub(Me.Name, "ExportData1")
End Sub

Private Function MakeMoney(ByVal Money As Double) As String
On Error GoTo ChkErr
  Dim strString1 As String
  
  '將金額轉換為 8 位的字串，不足 8 位前補 0
  If Money < 0 Then
    Money = -Money
    strString1 = "-" & Right("00000000" & Money, 7)
  Else
    strString1 = Right("00000000" & Money, 8)
  End If
  MakeMoney = strString1
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "MakeMoney")
End Function
Private Sub ExportData2()
On Error GoTo ChkErr
  'CP月租費資料匯出
  Dim dblRealAmtTotal As Double
  Dim lngDataTotal, lngTmp As Long
  Dim lngLogCount As Long
  Dim strMaster As String
  Dim strString1 As String
  Dim rsTemp As New ADODB.Recordset
  Dim rsTemp1 As New ADODB.Recordset
  Dim rsTemp2 As New ADODB.Recordset
  Dim strsql As String, strSqlIN As String
  Dim tmpCustid As String, tmpFaciSNO As String
  Dim tmpBelongYM As String, tmpTFNServiceID As String
  Screen.MousePointer = vbHourglass
  Me.Enabled = False
  
  '前期預收是否拋轉（是）
  If chkCanExport1.Value Then
    strString1 = "Select * From " & GetOwner & "So167 A, " & GetOwner & "So004 B " & _
                 "Where (A.CustId=B.CustId) " & _
                 "And (A.FaciSno=B.FaciSno) " & _
                 "And (B.InstDate Is Not Null) " & _
                 "And (B.PrDate Is Null) " & _
                 "And (A.CompCode=" & garyGi(9) & ") " & _
                 "And (A.Servicetype='P') " & _
                 "And (A.RealDate<=To_Date('" & GiRealDate4.GetValue & "', 'YYYYMMDD')) " & _
                 "And (A.BelongYm='" & GiExportDate2.GetValue & "') " & _
                 " Order by A.CustID,A.CPStartDate"     '本來已匯出過的資料不再匯出，現在改了
                 '"And (CpReturnRent=0) "   'Dennis 與 lawrence 討論暫不加
  Else
    '前期預收是否拋轉（否）
    strString1 = "Select * From " & GetOwner & "So167 A, " & GetOwner & "So004 B " & _
                 "Where (A.CustId=B.CustId) " & _
                 "And (A.FaciSno=B.FaciSno) " & _
                 "And (B.InstDate Is Not Null) " & _
                 "And (B.PrDate Is Null) " & _
                 "And (A.CompCode=" & garyGi(9) & ") " & _
                 "And (A.Servicetype='P') " & _
                 "And (A.RealDate Between To_Date('" & GiRealDate3.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate4.GetValue & "', 'YYYYMMDD')) " & _
                 "And (A.BelongYm='" & GiExportDate2.GetValue & "') " & _
                 " Order by A.CustID,A.CPStartDate"     '本來已匯出過的資料不再匯出，現在改了
                 '"And (CpReturnRent=0) "   'Dennis 與 lawrence 討論暫不加
  End If
  Set rsTemp = gcnGi.Execute(strString1)
  
  lngDataTotal = 0
  dblRealAmtTotal = 0
  lngLogCount = 0
  
  '寫入表頭
'  If Not rsTemp.EOF Then
    strmTextFile1.Write "H" & GiExportDate3.GetValue & vbCrLf
'  End If
  
  '寫入表身
  Do While Not rsTemp.EOF
'        rsTemp.MovePrevious
      If tmpCustid = rsTemp("Custid") And tmpFaciSNO = rsTemp("FaciSNO") And tmpBelongYM = rsTemp("BelongYM") And tmpTFNServiceID = rsTemp("TFNServiceID") & "" Then
        rsTemp.MoveNext
        If rsTemp.EOF Then Exit Do
      End If
      tmpCustid = rsTemp("Custid") & ""
      tmpFaciSNO = rsTemp("FaciSNO") & ""
      tmpBelongYM = rsTemp("BelongYM") & ""
      tmpTFNServiceID = rsTemp("TFNServiceID") & ""
      strSqlIN = ""
      If tmpTFNServiceID <> "" Then strSqlIN = "AND A.TFNServiceID='" & rsTemp("TFNServiceID") & "" & "' "
      
      If chkCanExport1.Value Then
          strsql = "Select A.CustId, A.FaciSNO, A.BelongYM, A.TFNServiceID,SUM(A.ShouldAmt) ShouldAmt,SUM(RealAmt) RealAmt From " & GetOwner & "So167 A, " & GetOwner & "So004 B " & _
                     "Where (A.CustId=B.CustId) " & _
                     "And (A.FaciSno=B.FaciSno) " & _
                     "And (B.InstDate Is Not Null) " & _
                     "And (B.PrDate Is Null) " & _
                     "And (A.CompCode=" & garyGi(9) & ") " & _
                     "And (A.Servicetype='P') " & _
                     "And (A.RealDate<=To_Date('" & GiRealDate4.GetValue & "', 'YYYYMMDD')) " & _
                     "And (A.BelongYm='" & GiExportDate2.GetValue & "') " & _
                     "AND A.Custid='" & rsTemp("Custid") & "" & "' " & _
                     "AND A.FaciSNO='" & rsTemp("FaciSNO") & "" & "' " & _
                     "AND A.BelongYM='" & rsTemp("BelongYM") & "" & "' " & _
                      strSqlIN & _
                     " GROUP BY A.CustId, A.FaciSNO, A.BelongYM, A.TFNServiceID"
        If Not GetRS(rsTemp1, strsql, gcnGi, , adOpenForwardOnly, adLockOptimistic, adUseClient, False) Then Exit Sub
        '找出
          strsql = "Select A.* From " & GetOwner & "So167 A, " & GetOwner & "So004 B " & _
                     "Where (A.CustId=B.CustId) " & _
                     "And (A.FaciSno=B.FaciSno) " & _
                     "And (B.InstDate Is Not Null) " & _
                     "And (B.PrDate Is Null) " & _
                     "And (A.CompCode=" & garyGi(9) & ") " & _
                     "And (A.Servicetype='P') " & _
                     "And (A.RealDate<=To_Date('" & GiRealDate4.GetValue & "', 'YYYYMMDD')) " & _
                     "And (A.BelongYm='" & GiExportDate2.GetValue & "') " & _
                     "AND A.Custid='" & rsTemp("Custid") & "" & "' " & _
                     "AND A.FaciSNO='" & rsTemp("FaciSNO") & "" & "'" & _
                     "AND A.BelongYM='" & rsTemp("BelongYM") & "" & "' " & _
                      strSqlIN & " Order by CPStartDate"
                     
        If Not GetRS(rsTemp2, strsql, gcnGi, , adOpenForwardOnly, adLockOptimistic, adUseClient, False) Then Exit Sub
        rsTemp2.MoveLast
    Else
          strsql = "Select A.CustId, A.FaciSNO, A.BelongYM, A.TFNServiceID,SUM(A.ShouldAmt) ShouldAmt,SUM(RealAmt) RealAmt From " & GetOwner & "So167 A, " & GetOwner & "So004 B " & _
                     "Where (A.CustId=B.CustId) " & _
                     "And (A.FaciSno=B.FaciSno) " & _
                     "And (B.InstDate Is Not Null) " & _
                     "And (B.PrDate Is Null) " & _
                     "And (A.CompCode=" & garyGi(9) & ") " & _
                     "And (A.Servicetype='P') " & _
                     "And (A.RealDate Between To_Date('" & GiRealDate3.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate4.GetValue & "', 'YYYYMMDD')) " & _
                     "And (A.BelongYm='" & GiExportDate2.GetValue & "') " & _
                     "AND A.Custid='" & rsTemp("Custid") & "" & "' " & _
                     "AND A.FaciSNO='" & rsTemp("FaciSNO") & "" & "' " & _
                     "AND A.BelongYM='" & rsTemp("BelongYM") & "" & "' " & _
                      strSqlIN & _
                     " GROUP BY A.CustId, A.FaciSNO, A.BelongYM, A.TFNServiceID"
        If Not GetRS(rsTemp1, strsql, gcnGi, , adOpenForwardOnly, adLockOptimistic, adUseClient, False) Then Exit Sub
    
          strsql = "Select A.* From " & GetOwner & "So167 A, " & GetOwner & "So004 B " & _
                     "Where (A.CustId=B.CustId) " & _
                     "And (A.FaciSno=B.FaciSno) " & _
                     "And (B.InstDate Is Not Null) " & _
                     "And (B.PrDate Is Null) " & _
                     "And (A.CompCode=" & garyGi(9) & ") " & _
                     "And (A.Servicetype='P') " & _
                     "And (A.RealDate Between To_Date('" & GiRealDate3.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate4.GetValue & "', 'YYYYMMDD')) " & _
                     "And (A.BelongYm='" & GiExportDate2.GetValue & "') " & _
                     "AND A.Custid='" & rsTemp("Custid") & "" & "' " & _
                     "AND A.FaciSNO='" & rsTemp("FaciSNO") & "" & "'" & _
                     "AND A.BelongYM='" & rsTemp("BelongYM") & "" & "' " & _
                      strSqlIN & " Order by CPStartDate"
        If Not GetRS(rsTemp2, strsql, gcnGi, , adOpenForwardOnly, adLockOptimistic, adUseClient, False) Then Exit Sub
        rsTemp2.MoveLast
    End If
    strMaster = "M" & Left(rsTemp("Custid") & Space(10), 10)
    
    '針對電話號碼處理，因取出的電話區碼和號碼間有 - 符號，匯出時要濾掉
    '- 的位置
    If Len(rsTemp("FaciSNo")) > 0 Then
      lngTmp = InStr(rsTemp("FaciSNo"), "-")
      
      '判斷有沒有區碼
      If lngTmp > 0 Then
        strString1 = Left(rsTemp("FaciSNo"), lngTmp - 1)
        strMaster = strMaster & Left(strString1 & Space(3), 3)
      Else
        '沒有，填空白
        strMaster = strMaster & Space(3)
      End If
      
      '號碼，取 10 碼
      strString1 = Mid(rsTemp("FaciSNo"), lngTmp + 1, 10)
      strMaster = strMaster & Left(strString1 & Space(10), 10)
    Else
      strMaster = strMaster & Space(13)
    End If
    
    'strMaster = strMaster & Left(rsTemp("FaciSNo") & Space(13), 13)
    strMaster = strMaster & Left(rsTemp("TfnServiceId") & "" & Space(10), 10)
    strMaster = strMaster & Left(rsTemp("BillNo") & Space(16), 16)
    strMaster = strMaster & MakeMoney(rsTemp1("ShouldAmt"))
    strMaster = strMaster & MakeMoney(rsTemp1("RealAmt"))
    strMaster = strMaster & Right("00" & rsTemp("RealPeriod"), 2)
    strMaster = strMaster & MakeMoney(rsTemp("MonthAmt"))
    strMaster = strMaster & Left(rsTemp("CpCitemCode") & Space(10), 10)
    strMaster = strMaster & rsTemp("CpContractMode")
    strMaster = strMaster & rsTemp("CpStartDate")
    strMaster = strMaster & rsTemp2("CpStopDate")
    strMaster = strMaster & Left(rsTemp("CpAdjCode") & Space(4), 4)
    strMaster = strMaster & MakeMoney(0 & rsTemp("CpAdjAmount"))
    strString1 = rsTemp("CpAdjName") & Space(50)
    strString1 = StrConv(LeftB(StrConv(strString1, vbFromUnicode), 50), vbUnicode)
    strMaster = strMaster & strString1
    
    strmTextFile1.Write strMaster & vbCrLf
    
    lngDataTotal = lngDataTotal + 1
    dblRealAmtTotal = dblRealAmtTotal + rsTemp("RealAmt")
    
    '已匯出過的要 LOG
    If rsTemp("CpExport") <> 0 Then
      strString1 = "客編(" & rsTemp("CustId") & ")，" & _
                   "CP門號(" & rsTemp("FaciSNo") & ")，" & _
                   "TFN服務代碼(" & rsTemp("TfnServiceId") & ")，" & _
                   "單據編號(" & rsTemp("BillNo") & ")，" & _
                   "此筆資料已匯出過"
      Call LogError3(lngLogCount, strString1)
    End If
    
    rsTemp.MoveNext
  Loop
  
  '寫入表尾
  strString1 = "T" & Right("000000" & lngDataTotal, 6)
  strString1 = strString1 & MakeMoney(dblRealAmtTotal)
  strmTextFile1.Write strString1 & vbCrLf
  
  Set rsTemp = Nothing
  
  '已匯出的資料下次不再匯出
  If chkCanExport1.Value Then
    '前期預收是否拋轉（是）
    strString1 = "Update " & GetOwner & "So167 " & _
                 "Set CpReturnRent=1,CpExport=1 " & _
                 "Where " & GetOwner & "So167.RowId In ( " & _
                 "  Select A.RowId From " & GetOwner & "So167 A, " & GetOwner & "So004 B " & _
                 "  Where (A.CustId=B.CustId) " & _
                 "  And (A.FaciSno=B.FaciSno) " & _
                 "  And (B.InstDate Is Not Null) " & _
                 "  And (A.CompCode=" & garyGi(9) & ") " & _
                 "  And (A.Servicetype='P') " & _
                 "  And (A.RealDate<=To_Date('" & GiRealDate4.GetValue & "', 'YYYYMMDD')) " & _
                 "  And (A.BelongYm='" & GiExportDate2.GetValue & "') " & _
                 "  And (A.CpExport=0)) "
    gcnGi.Execute (strString1)
'    strString1 = "Update " & GetOwner & "So167 " & _
'                 "Set CpExport=1 " & _
'                 "Where " & GetOwner & "So167.RowId In ( " & _
'                 "  Select A.RowId From " & GetOwner & "So167 A, " & GetOwner & "So004 B " & _
'                 "  Where (A.CustId=B.CustId) " & _
'                 "  And (A.FaciSno=B.FaciSno) " & _
'                 "  And (B.InstDate Is Not Null) " & _
'                 "  And (B.PrDate Is Null) " & _
'                 "  And (A.CompCode=" & garyGi(9) & ") " & _
'                 "  And (A.Servicetype='P') " & _
'                 "  And (A.RealDate<=To_Date('" & GiRealDate4.GetValue & "', 'YYYYMMDD')) " & _
'                 "  And (A.BelongYm='" & GiExportDate2.GetValue & "') " & _
'                 "  And (A.CpExport=0)) "
'    gcnGi.Execute (strString1)
  Else
    '前期預收是否拋轉（否）
    strString1 = "Update " & GetOwner & "So167 " & _
                 "Set CpReturnRent=1,CpExport=1 " & _
                 "Where " & GetOwner & "So167.RowId In ( " & _
                 "  Select A.RowId From " & GetOwner & "So167 A, " & GetOwner & "So004 B " & _
                 "  Where (A.CustId=B.CustId) " & _
                 "  And (A.FaciSno=B.FaciSno) " & _
                 "  And (B.PrDate Is Not Null) " & _
                 "  And (A.CompCode=" & garyGi(9) & ") " & _
                 "  And (A.Servicetype='P') " & _
                 "  And (A.RealDate Between To_Date('" & GiRealDate3.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate4.GetValue & "', 'YYYYMMDD')) " & _
                 "  And (A.BelongYm='" & GiExportDate2.GetValue & "') " & _
                 "  And (A.CpExport=0)) "
    gcnGi.Execute (strString1)
'    strString1 = "Update " & GetOwner & "So167 " & _
'                 "Set CpExport=1 " & _
'                 "Where " & GetOwner & "So167.RowId In ( " & _
'                 "  Select A.RowId From " & GetOwner & "So167 A, " & GetOwner & "So004 B " & _
'                 "  Where (A.CustId=B.CustId) " & _
'                 "  And (A.FaciSno=B.FaciSno) " & _
'                 "  And (B.InstDate Is Not Null) " & _
'                 "  And (B.PrDate Is Null) " & _
'                 "  And (A.CompCode=" & garyGi(9) & ") " & _
'                 "  And (A.Servicetype='P') " & _
'                 "  And (A.RealDate Between To_Date('" & GiRealDate3.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate4.GetValue & "', 'YYYYMMDD')) " & _
'                 "  And (A.BelongYm='" & GiExportDate2.GetValue & "') " & _
'                 "  And (A.CpExport=0)) "
'    gcnGi.Execute (strString1)
  End If
  
  If lngLogCount > 0 Then
    strString1 = "匯出筆數：" & lngDataTotal & vbCrLf & vbCrLf & _
                 "資料中有已匯出的資料存在，筆數：" & lngLogCount & vbCrLf & _
                 "請參考問題記錄檔。" & vbCrLf
    MsgBox strString1, vbExclamation, "訊息"
    Shell "Notepad " & txtLogFile5.Text, vbNormalFocus
  Else
    strString1 = "匯出完成 !!" & vbCrLf & vbCrLf & _
                 "匯出筆數：" & lngDataTotal & vbCrLf
    MsgBox strString1, vbExclamation, "訊息"
  End If
  
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Call ErrSub(Me.Name, "ExportData2")
End Sub

Private Sub ExportData3()
On Error GoTo ChkErr
  'CP退租資料匯出
  Dim dblRealAmtTotal As Double
  Dim lngDataTotal, lngTmp As Long
  Dim strMaster As String
  Dim rsTemp As New ADODB.Recordset
  Dim rsTemp2 As New ADODB.Recordset
  Dim strString1 As String

  Screen.MousePointer = vbHourglass
  Me.Enabled = False
  
  strString1 = "Select * From " & GetOwner & "So009 A, " & GetOwner & "So034 B, " & GetOwner & "So165 C " & _
               "Where (A.SNo = B.BillNo) " & _
               "And (B.CItemCode In " & _
               "  (Select CItemCode2 From " & GetOwner & "So165 Where CpProperty=3)) " & _
               "And (B.CItemCode=C.CitemCode2) " & _
               "And (A.CompCode=" & garyGi(9) & ") " & _
               "And (A.Servicetype='P') " & _
               "And (A.PrCode In (" & gimPrCode1.GetQueryCode & ")) " & _
               "And (A.ReasonCode In (" & gimReasonCode.GetQueryCode & ")) " & _
               "And (A.FinTime>=To_Date('" & GiPrDate1.GetValue & "', 'YYYYMMDD')) " & _
               "And (A.FinTime<To_Date('" & GiPrDate2.GetValue + 1 & "', 'YYYYMMDD')) " & _
               "And (A.ReturnCode Is Null) " & _
               "And (To_Date('" & GiExportRefDate1.GetValue & "', 'YYYYMMDD') Between B.RealStartDate And B.RealStopDate) " & _
               "And (B.UcCode Is Null) " & _
               "And (B.CancelFlag=0) "
  Set rsTemp = gcnGi.Execute(strString1)
  
  lngDataTotal = 0
  dblRealAmtTotal = 0
  
  '寫入表頭
'  If Not rsTemp.EOF Then
    strmTextFile1.Write "H" & GiExportDate4.GetValue & vbCrLf
'  End If
  
  '寫入表身
  While Not rsTemp.EOF
    strMaster = "M" & Left(rsTemp("Custid") & Space(10), 10)
    
    '針對電話號碼處理，因取出的電話區碼和號碼間有 - 符號，匯出時要濾掉
    '- 的位置
    If Len(rsTemp("FaciSNo")) > 0 Then
      lngTmp = InStr(rsTemp("FaciSNo"), "-")
      
      '判斷有沒有區碼
      If lngTmp > 0 Then
        strString1 = Left(rsTemp("FaciSNo"), lngTmp - 1)
        strMaster = strMaster & Left(strString1 & Space(3), 3)
      Else
        '沒有，填空白
        strMaster = strMaster & Space(3)
      End If
      
      '號碼，取 10 碼
      strString1 = Mid(rsTemp("FaciSNo"), lngTmp + 1, 10)
      strMaster = strMaster & Left(strString1 & Space(10), 10)
    Else
      strMaster = strMaster & Space(13)
    End If
        
    'strMaster = strMaster & Left(rsTemp("FaciSNo") & Space(13), 13)
    strMaster = strMaster & Left(rsTemp("TfnServiceId") & Space(10), 10)
    strMaster = strMaster & Left(rsTemp("BillNo") & Space(16), 16)
    strMaster = strMaster & MakeMoney(rsTemp("ShouldAmt"))
    
    '取得單月牌價
    strString1 = "Select MonthAmt " & _
                 "From " & GetOwner & "Cd078A A, " & GetOwner & "So003 B " & _
                 "Where (B.CustId=" & rsTemp("CustId") & ") " & _
                 "And (B.FaciSno='" & rsTemp("FaciSno") & "') " & _
                 "And (B.BpCode=A.BpCode) "
    Set rsTemp2 = gcnGi.Execute(strString1)
    
    If rsTemp2.EOF Then
      strMaster = strMaster & Space(8)
    Else
      strMaster = strMaster & MakeMoney(rsTemp2("MonthAmt"))
    End If
    Set rsTemp2 = Nothing
    
    strMaster = strMaster & Left(rsTemp("CpCitemCode") & Space(10), 10)
    GiTempDate.SetValue (rsTemp("FinTime"))
    strMaster = strMaster & GiTempDate.GetValue
    strMaster = strMaster & GiExportDate3.GetValue
    strString1 = rsTemp("ReasonName") & Space(50)
    strString1 = StrConv(LeftB(StrConv(strString1, vbFromUnicode), 50), vbUnicode)
    strMaster = strMaster & strString1
    strmTextFile1.Write strMaster & vbCrLf
    lngDataTotal = lngDataTotal + 1
    dblRealAmtTotal = dblRealAmtTotal + rsTemp("ShouldAmt")
    rsTemp.MoveNext
  Wend
  
  '寫入表尾
  strString1 = "T" & Right("000000" & lngDataTotal, 6)
  strString1 = strString1 & MakeMoney(dblRealAmtTotal)
  strmTextFile1.Write strString1 & vbCrLf
  Set rsTemp = Nothing
  
  strString1 = "匯出完成 !!" & vbCrLf & vbCrLf & _
               "匯出筆數：" & lngDataTotal & vbCrLf
  MsgBox strString1, vbExclamation, "訊息"
  
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Call ErrSub(Me.Name, "ExportData3")
End Sub

Private Sub ExportData4()
On Error GoTo ChkErr
  'CMCP強拆資料匯出
  Dim lngDataTotal As Long
  Dim strMaster As String
  Dim strString1 As String
  Dim rsTemp As New ADODB.Recordset
  Dim rsTemp2 As New ADODB.Recordset
  
  Screen.MousePointer = vbHourglass
  Me.Enabled = False
  
  strString1 = "Select * From " & GetOwner & "So009 " & _
               "Where (CompCode=" & garyGi(9) & ") " & _
               "And (Servicetype='P') " & _
               "And (PrCode In (" & gimPrCode2.GetQueryCode & ")) " & _
               "And (ReasonCode In (" & gimUCCode.GetQueryCode & ")) " & _
               "And (FinTime>=To_Date('" & GiPrDate3.GetValue & "', 'YYYYMMDD')) " & _
               "And (FinTime<To_Date('" & GiPrDate4.GetValue + 1 & "', 'YYYYMMDD')) " & _
               "And (ReturnCode Is Null) "
  Set rsTemp = gcnGi.Execute(strString1)
  
  '寫入檔頭
  lngDataTotal = 0
'  If Not rsTemp.EOF Then
    strmTextFile1.Write "H" & GiExportDate5.GetValue & vbCrLf
'  End If
  
  '寫入表身
  While Not rsTemp.EOF
    strMaster = "B" & Left(rsTemp("Custid") & Space(10), 10)
    strString1 = "Select TfnServiceId " & _
                 "From " & GetOwner & "So004 " & _
                 "Where (PrsNo='" & rsTemp("Sno") & "') " & _
                 "And (CustId=" & rsTemp("Custid") & ") "
    Set rsTemp2 = gcnGi.Execute(strString1)
    
    If rsTemp2.EOF Then
      strMaster = strMaster & Space(10)
    Else
      strMaster = strMaster & Left(rsTemp2("TfnServiceId") & Space(10), 10)
    End If
    Set rsTemp2 = Nothing
    
    strMaster = strMaster & "CP"
    GiTempDate.SetValue (rsTemp("FinTime"))
    strMaster = strMaster & GiTempDate.GetValue
    strString1 = rsTemp("ReasonName") & Space(50)
    strString1 = StrConv(LeftB(StrConv(strString1, vbFromUnicode), 50), vbUnicode)
    strMaster = strMaster & strString1
    strmTextFile1.Write strMaster & vbCrLf
    lngDataTotal = lngDataTotal + 1
    rsTemp.MoveNext
  Wend
  
  '寫入表尾
  strString1 = "T" & Right("000000" & lngDataTotal, 6)
  strmTextFile1.Write strString1 & vbCrLf
  Set rsTemp = Nothing
  
  strString1 = "匯出完成 !!" & vbCrLf & vbCrLf & _
               "匯出筆數：" & lngDataTotal & vbCrLf
  MsgBox strString1, vbExclamation, "訊息"
  
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Call ErrSub(Me.Name, "ExportData4")
End Sub

Private Sub ExportData5()
On Error GoTo ChkErr
  'CP通信退費資料匯出
  Dim lngExportDetailCount, lngExportMasterCount, lngDataCountTail, lngSubDetailCount, lngTmp As Long
  Dim lngLogCount As Long
  Dim dblRealAmtMaster, dblRealAmtWithTaxMaster, dblRealAmtWithoutTaxMaster, dblRealAmtZeroTaxMaster As Double
  Dim dblTotalRealAmtTail As Double
  Dim strCustid, strFaciSNo, strTfnServiceId, strBillNo, strTFNBillNo As String
  Dim strString1 As String
  Dim strMaster, strDetail, strTail, strSTName As String
  Dim rsTemp As New ADODB.Recordset
  
  Screen.MousePointer = vbHourglass
  Me.Enabled = False
  
  'strString1 = "Select * From " & GetOwner & "So170 " & _
               "Where (CompCode=" & garyGi(9) & ") " & _
               "And (Servicetype='P') " & _
               "And (RealDate Between To_Date('" & GiRealDate5.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate6.GetValue & "', 'YYYYMMDD')) " & _
               "And (CpExport=0) " & _
               "Order By CustId, FaciSNo, TfnServiceId, BillNo, TfnBillNo "
  
  '本來已匯出過的資料不再匯出，現在改了
  strString1 = "Select * From " & GetOwner & "So170 " & _
               "Where (CompCode=" & garyGi(9) & ") " & _
               "And (Servicetype='P') " & _
               "And (RealDate Between To_Date('" & GiRealDate5.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate6.GetValue & "', 'YYYYMMDD')) " & _
               "Order By CustId, FaciSNo, TfnServiceId, BillNo, TfnBillNo "
  Set rsTemp = gcnGi.Execute(strString1)
  
  '寫入檔頭
'  If Not rsTemp.EOF Then
    strmTextFile1.Write "H" & GiExportDate1.GetValue & vbCrLf
'  End If
  
  lngExportMasterCount = 0
  lngExportDetailCount = 0
  lngLogCount = 0
  strCustid = ""
  strFaciSNo = ""
  strTfnServiceId = ""
  strBillNo = ""
  strTFNBillNo = ""
  
  '寫入表身
  While Not rsTemp.EOF
    '依 CustId, FaciSNo, TfnServiceId, BillNo, TfnBillNo 分組產生 Master
    If (strCustid <> rsTemp("CustId") & "") Or (strFaciSNo <> rsTemp("FaciSNo") & "") Or (strTfnServiceId <> rsTemp("TfnServiceId") & "") Or (strBillNo <> rsTemp("BillNo") & "") Or (strTFNBillNo <> rsTemp("TfnBillNo") & "") Then
      If strCustid <> "" Then
        strMaster = "M" & strMaster
        strMaster = strMaster & Left(strCustid & Space(10), 10)
        
        '針對電話號碼處理，因取出的電話區碼和號碼間有 - 符號，匯出時要濾掉
        '- 的位置
        If Len(strFaciSNo) > 0 Then
          lngTmp = InStr(strFaciSNo, "-")
      
          '判斷有沒有區碼
          If lngTmp > 0 Then
            strString1 = Left(strFaciSNo, lngTmp - 1)
            strMaster = strMaster & Left(strString1 & Space(3), 3)
          Else
            '沒有，填空白
            strMaster = strMaster & Space(3)
          End If
      
          '號碼，取 10 碼
          strString1 = Mid(strFaciSNo, lngTmp + 1, 10)
          strMaster = strMaster & Left(strString1 & Space(10), 10)
        Else
          strMaster = strMaster & Space(13)
        End If
        
        'strMaster = strMaster & Left(strFaciSNo & Space(13), 13)
        strMaster = strMaster & Left(strTfnServiceId & Space(10), 10)
        strMaster = strMaster & Left(strBillNo & Space(16), 16)
        strMaster = strMaster & Left(strTFNBillNo & Space(13), 13)
        strMaster = strMaster & MakeMoney(dblRealAmtMaster)
        strMaster = strMaster & MakeMoney(dblRealAmtWithTaxMaster)
        strMaster = strMaster & MakeMoney(dblRealAmtZeroTaxMaster)
        strMaster = strMaster & MakeMoney(dblRealAmtWithoutTaxMaster)
        strString1 = Space(50)
        If lngSubDetailCount = 1 Then
          strString1 = strSTName & strString1
          strString1 = StrConv(LeftB(StrConv(strString1, vbFromUnicode), 50), vbUnicode)
        End If
        
        strMaster = strMaster & strString1
        
        '將 Buffer 中的資料匯出
        strmTextFile1.Write strMaster & vbCrLf
        strmTextFile1.Write strDetail
        
        strMaster = ""
        strDetail = ""
      End If
      '將產生一筆新的 Master 資料，將相關欄位先清空及填入預設值
      lngSubDetailCount = 0
      dblRealAmtMaster = 0
      dblRealAmtWithTaxMaster = 0
      dblRealAmtZeroTaxMaster = 0
      dblRealAmtWithoutTaxMaster = 0
      strCustid = rsTemp("CustId") & ""
      strFaciSNo = rsTemp("FaciSNo") & ""
      strTfnServiceId = rsTemp("TfnServiceId") & ""
      strBillNo = rsTemp("BillNo") & ""
      strTFNBillNo = rsTemp("TfnBillNo") & ""
      lngExportMasterCount = lngExportMasterCount + 1
    End If
    
    dblRealAmtMaster = dblRealAmtMaster + rsTemp("RealAmt")
    
    If rsTemp("CpRate") = "Y" Then
      dblRealAmtWithTaxMaster = dblRealAmtWithTaxMaster + rsTemp("RealAmt")
    End If
    If rsTemp("CpRate") = "N" Then
      dblRealAmtZeroTaxMaster = dblRealAmtZeroTaxMaster + rsTemp("RealAmt")
    End If
    If rsTemp("CpRate") = "F" Then
      dblRealAmtWithoutTaxMaster = dblRealAmtWithoutTaxMaster + rsTemp("RealAmt")
    End If
    
    '寫入細項
    strDetail = strDetail & "D" & Left(rsTemp("CpAdjCitemCode") & Space(10), 10)
    strDetail = strDetail & Left(rsTemp("CpRate") & Space(1), 1)
    strDetail = strDetail & MakeMoney(rsTemp("RealAmt"))
    
    If rsTemp("CpRate") = "Y" Then
      strDetail = strDetail & MakeMoney(Round(rsTemp("RealAmt") * 0.95))
    Else
      strDetail = strDetail & MakeMoney(rsTemp("RealAmt"))
    End If
    If rsTemp("CpRate") = "Y" Then
      strDetail = strDetail & MakeMoney(rsTemp("RealAmt") - Round(rsTemp("RealAmt") * 0.95))
    Else
      strDetail = strDetail & MakeMoney(0)
    End If
    
    strDetail = strDetail & Left(rsTemp("CpAdjCode") & Space(4), 4)
    GiTempDate.SetValue (rsTemp("RealDate"))
    strDetail = strDetail & Left(GiTempDate.GetValue, 6)
    strSTName = rsTemp("StName") & ""
    strString1 = strSTName & Space(50)
    strString1 = StrConv(LeftB(StrConv(strString1, vbFromUnicode), 50), vbUnicode)
    strDetail = strDetail & strString1
    strDetail = strDetail & vbCrLf
    
    lngExportDetailCount = lngExportDetailCount + 1
    lngSubDetailCount = lngSubDetailCount + 1
    
    '已匯出過的要 LOG
    If rsTemp("CpExport") <> 0 Then
      strString1 = "客編(" & rsTemp("CustId") & ")，" & _
                   "CP門號(" & rsTemp("FaciSNo") & ")，" & _
                   "TFN服務代碼(" & rsTemp("TfnServiceId") & ")，" & _
                   "單據編號(" & rsTemp("BillNo") & ")，" & _
                   "TFN帳單編號(" & rsTemp("TfnBillNo") & ")，" & _
                   "此筆資料已匯出過"
      Call LogError3(lngLogCount, strString1)
    End If
    
    rsTemp.MoveNext
  Wend
  '全部跑完後須將 Buffer 中尚未匯出的 Master 及 Detail 資料也匯出
  If strCustid <> "" Then
    strMaster = "M" & strMaster
    strMaster = strMaster & Left(strCustid & Space(10), 10)
    
    '針對電話號碼處理，因取出的電話區碼和號碼間有 - 符號，匯出時要濾掉
    '- 的位置
    If Len(strFaciSNo) > 0 Then
      lngTmp = InStr(strFaciSNo, "-")
      
      '判斷有沒有區碼
      If lngTmp > 0 Then
        strString1 = Left(strFaciSNo, lngTmp - 1)
        strMaster = strMaster & Left(strString1 & Space(3), 3)
      Else
        '沒有，填空白
        strMaster = strMaster & Space(3)
      End If
      
      '號碼，取 10 碼
      strString1 = Mid(strFaciSNo, lngTmp + 1, 10)
      strMaster = strMaster & Left(strString1 & Space(10), 10)
    Else
      strMaster = strMaster & Space(13)
    End If
    
    'strMaster = strMaster & Left(strFaciSNo & Space(13), 13)
    strMaster = strMaster & Left(strTfnServiceId & Space(10), 10)
    strMaster = strMaster & Left(strBillNo & Space(16), 16)
    strMaster = strMaster & Left(strTFNBillNo & Space(13), 13)
    strMaster = strMaster & MakeMoney(dblRealAmtMaster)
    strMaster = strMaster & MakeMoney(dblRealAmtWithTaxMaster)
    strMaster = strMaster & MakeMoney(dblRealAmtZeroTaxMaster)
    strMaster = strMaster & MakeMoney(dblRealAmtWithoutTaxMaster)
    strString1 = Space(50)
    If lngSubDetailCount = 1 Then
      strString1 = strSTName & strString1
      strString1 = StrConv(LeftB(StrConv(strString1, vbFromUnicode), 50), vbUnicode)
    End If
    strMaster = strMaster & strString1
    
    '將 Buffer 中的資料匯出
    strmTextFile1.Write strMaster & vbCrLf
    strmTextFile1.Write strDetail
    
    strMaster = ""
    strDetail = ""
  End If
  Set rsTemp = Nothing
  
  '取得依 TFNBillNo 分組的筆數
  strString1 = "Select Count(Distinct TFnBillNo) As DataCount, " & _
               "       Sum(RealAmt) As TotalRealAmt " & _
               "From " & GetOwner & "So170 " & _
               "Where (CompCode=" & garyGi(9) & ") " & _
               "And (Servicetype='P') " & _
               "And (RealDate Between To_Date('" & GiRealDate5.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate6.GetValue & "', 'YYYYMMDD')) "
               '"And (CpExport=0) "
  Set rsTemp = gcnGi.Execute(strString1)
  
  lngDataCountTail = rsTemp("DataCount")
  dblTotalRealAmtTail = NoZero(rsTemp("TotalRealAmt"), True)
  
  '寫入表尾
  strTail = "T" & Right("000000" & lngDataCountTail, 6)
  strTail = strTail & MakeMoney(dblTotalRealAmtTail)
  Set rsTemp = Nothing
  strmTextFile1.Write strTail
  
  '已匯出的資料下次不再匯出
  strString1 = "Update " & GetOwner & "So170 " & _
               "Set CpExport=1 " & _
               "Where (CompCode=" & garyGi(9) & ") " & _
               "And (Servicetype='P') " & _
               "And (Realdate Between To_Date('" & GiRealDate5.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate6.GetValue & "', 'YYYYMMDD')) " & _
               "And (CpExport=0) "
  gcnGi.Execute (strString1)
  
  If lngLogCount > 0 Then
    strString1 = "匯出Master筆數：" & lngExportMasterCount & vbCrLf & _
                 "匯出Detail筆數：" & lngExportDetailCount & vbCrLf & vbCrLf & _
                 "資料中有已匯出的資料存在，筆數：" & lngLogCount & vbCrLf & _
                 "請參考問題記錄檔。" & vbCrLf
    MsgBox strString1, vbExclamation, "訊息"
    Shell "Notepad " & txtLogFile6.Text, vbNormalFocus
  Else
    strString1 = "匯出完成 !!" & vbCrLf & vbCrLf & _
                 "匯出Master筆數：" & lngExportMasterCount & vbCrLf & _
                 "匯出Detail筆數：" & lngExportDetailCount
    MsgBox strString1, vbExclamation, "訊息"
  End If
  
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Call ErrSub(Me.Name, "ExportData5")
End Sub

Private Sub ImportData2()
On Error GoTo ChkErr
  '基本資料核對匯入
  Dim blnHasHead, blnHasBody, blnHasTail, blnGoOn, blnHasError As Boolean
  Dim lngHeadErrorCount As Long
  Dim lngTailErrorCount As Long
  Dim lngBodyErrorCount As Long
  Dim lngBodyCount, lngIndex, lngLineNumber As Long
  Dim strTextData, strTel, strDeclarantName, strReason As String
  Dim strString1 As String
  Dim strCustid, strNumber, strTfn, strCs As String
  Dim vDataArray, vElement, vDataArray2 As Variant
  Dim rsTemp As New ADODB.Recordset
  Dim rsTemp2 As New ADODB.Recordset
  
  Dim CsCustId As String
  Dim CsNumber As String
  Dim CsId As String
  Dim CsName As String
  Dim CsInstDate As String
  Dim CsPrDate As String
  Dim TfnCustId As String
  Dim TfnNumber As String
  Dim TfnId As String
  Dim TfnName As String
  Dim TfnInstDate As String
  Dim TfnPrDate As String
  Dim strTfnServiceId As String
  Dim strReplace As String, vRepArry As Variant
  Screen.MousePointer = vbHourglass
  Me.Enabled = False
  
  lngBodyErrorCount = 0
  lngHeadErrorCount = 0
  lngTailErrorCount = 0
  lngBodyCount = 0
  blnHasHead = False
  blnHasTail = False
  blnHasBody = False
  
  '讀入文字檔，並分段
  strTextData = FSO.OpenTextFile(txtImportFile2).ReadAll
  vDataArray = Split(strTextData, vbCrLf)
  
  'Progress Bar
  PB3.Visible = True
  PB3.Value = PB3.Min
  On Error Resume Next
  PB3.Max = UBound(vDataArray)
  
  '開始一行一行剖析
  lngLineNumber = 1
  For Each vElement In vDataArray
    PB3.Value = lngLineNumber - 1
    Select Case Left(vElement, 1)
      Case "H"
        '檔頭
        blnGoOn = True
        If blnHasBody Or blnHasTail Then
          Call LogError1(lngHeadErrorCount, "表頭必須在文字檔的第一行", "", lngLineNumber, vElement)
          blnGoOn = False
        End If
        vDataArray2 = Split(vElement, ",")
        If UBound(vDataArray2) <> 2 Then
          Call LogError1(lngHeadErrorCount, "表頭欄位數目錯誤，欄位只能有三欄", "", lngLineNumber, vElement)
          blnGoOn = False
        End If
        blnHasHead = True
      Case "B"
        '資料項目
        blnGoOn = True
        blnHasBody = True
        If Not blnHasHead Then
          Call LogError1(lngBodyErrorCount, "資料不可以出現在表頭之前", "", lngLineNumber, vElement)
          blnGoOn = False
        End If
        If blnHasTail Then
          Call LogError1(lngBodyErrorCount, "資料不可以出現在表尾之後", "", lngLineNumber, vElement)
          blnGoOn = False
        End If
        vDataArray2 = Split(vElement, ",")
        If UBound(vDataArray2) <> 10 Then
          Call LogError1(lngBodyErrorCount, "資料欄位數目錯誤，欄位必須是十欄", "", lngLineNumber, vElement)
          blnGoOn = False
        End If
        
        '通過檢核，寫入 TMP012
        If blnGoOn Then
          lngBodyCount = lngBodyCount + 1
          strString1 = "Insert Into " & GetOwner & "Tmp012 " & _
                       "Values ( "
          For lngIndex = 1 To UBound(vDataArray2) - 1 '因為文字檔的最後一欄不匯入,所以要減 1
            vDataArray2(lngIndex) = Replace(vDataArray2(lngIndex), "%44", ",")
            
            Select Case lngIndex
              Case UBound(vDataArray2) - 2
                '行號是狀況外，要個別處理
                strString1 = strString1 & lngLineNumber & ", "
                vDataArray2(lngIndex) = "To_Date('" & vDataArray2(lngIndex) & "','YYYYMMDD')" & ", "
              Case UBound(vDataArray2) - 1
                vDataArray2(lngIndex) = "To_Date('" & vDataArray2(lngIndex) & "','YYYYMMDD')"
              Case Else
                If vDataArray2(lngIndex) = "" Then
                  Call LogError1(lngBodyErrorCount, "欄位資料不能是空值", "", lngLineNumber, vElement)
                End If
                vDataArray2(lngIndex) = "'" & vDataArray2(lngIndex) & "'"
                'If lngIndex <> UBound(vDataArray2) - 1 Then
                  vDataArray2(lngIndex) = vDataArray2(lngIndex) & ", "
                'End If
            End Select
            
            strString1 = strString1 & vDataArray2(lngIndex)
          Next
          '記錄該筆資料在文字檔的位置
          strString1 = strString1 & ",'')"
          gcnGi.Execute (strString1)
        End If
      Case "T"
        '表尾
        blnGoOn = True
        If (Not blnHasBody) Or (Not blnHasHead) Then
          Call LogError1(lngTailErrorCount, "表尾必須在文字檔的最後一行", "", lngLineNumber, vElement)
          blnGoOn = False
        End If
        vDataArray2 = Split(vElement, ",")
        If UBound(vDataArray2) <> 1 Then
          Call LogError1(lngTailErrorCount, "表尾欄位數目錯誤，欄位只能有二欄", "", lngLineNumber, vElement)
          blnGoOn = False
        End If
        If lngBodyCount <> CLng(vDataArray2(1)) Then
          blnGoOn = False
          Call LogError1(lngTailErrorCount, "資料筆數(" & (lngBodyCount) & ")與表尾記載(" & vDataArray2(1) & ")的不符", "", lngLineNumber, vElement)
        End If
        blnHasTail = True
      Case Else
        If vElement <> "" Then
          Call LogError1(lngBodyErrorCount, "無效的資料項目，請檢查文字的第 1 碼是否標示錯誤", "", lngLineNumber, vElement)
        End If
    End Select
    lngLineNumber = lngLineNumber + 1
    DoEvents
  Next
  
  '如果沒有表尾也不行
  If Not blnHasTail Then
    Call LogError1(lngTailErrorCount, "沒有表尾資料", "", lngLineNumber, vElement)
    lngLineNumber = lngLineNumber + 1
  End If
  
  '匯入過程中有錯誤
  If (lngBodyErrorCount > 0) Or (lngHeadErrorCount > 0) Or (lngTailErrorCount > 0) Then
    strTextData = "檢核匯入的文字檔有誤，請參考問題記錄檔。" & vbCrLf
    MsgBox strTextData, vbExclamation, "警告"
    Shell "Notepad " & txtLogFile2.Text, vbNormalFocus
  Else
    lngBodyErrorCount = 0
    lngBodyCount = 0
    lngLineNumber = 2
    
    PB3.Value = PB3.Min
    
    On Error Resume Next
    PB3.Max = UBound(vDataArray) - 2
    Dim strSqlDeName, strSqlDeNameA, strSqlDeNameB As String
    Dim i As Integer
    Set rsTemp = gcnGi.Execute("Select * From " & GetOwner & "Tmp012 Order By RecNo")
    While Not rsTemp.EOF
      PB3.Value = lngLineNumber - 2
      '區碼 + 電話號碼
      strTel = Trim(Trim(rsTemp("CpAreaCode")) & "-" & rsTemp("CpNumber"))
      '用戶名稱
      strDeclarantName = StrConv(LeftB(StrConv(rsTemp("CustName"), vbFromUnicode), 50), vbUnicode)
'      strDeclarantName = Replace(Replace(Replace(strDeclarantName, "先生", ""), "小姐", ""), "女士", "")
'      '服務代碼(台固代碼)
'      strTfnServiceId = Trim(rsTemp("TFNServiceID"))
'     strReplace = "": strSqlDeName = "": strSqlDeNameA = "": strSqlDeNameB = ""
'      strReplace = "先生,小姐,女士"
'      vRepArry = Split(strReplace, ",")
'      strSqlDeName = "SO004.DeclarantName"
'      For i = 0 To UBound(vRepArry)
'        strSqlDeNameA = strSqlDeNameA & "(REPLACE("
'        strSqlDeNameB = ",'" & vRepArry(i) & "','')" & strSqlDeNameB
        
'      Next
'      strSqlDeName = strSqlDeNameA & strSqlDeName & strSqlDeNameB
      '判斷要匯入的資料在 SO004 是否已存在了
      '問題集2798 判斷身份證時,先轉換大寫
      strString1 = "Select SO004.SeqNo, SO004.InstDate, SO004.OldInstDate, SO004.PrDate " & _
                   "From " & GetOwner & "So004," & GetOwner & "CD022 " & _
                   "Where SO004.FaciCode=CD022.CodeNo " & _
                   "AND (SO004.CustId=" & rsTemp("CustId") & " ) " & _
                   "And (SO004.CompCode=" & garyGi(9) & ") " & _
                   "And (SO004.Servicetype='P') " & _
                   "And (SO004.InstDate Is Not Null) " & _
                   "And (SO004.PrDate Is Null) " & _
                   "And (SO004.FaciSNo='" & strTel & "' ) " & _
                   "And ((upper(SO004.Id)='" & UCase(rsTemp("Id")) & "' ) " & _
                   "  Or (upper(SO004.Id2)='" & UCase(rsTemp("Id")) & "' )) " & _
                   "And (REPLACE(REPLACE(REPLACE(SO004.DeclarantName,'女士',''),'小姐',''),'先生','')='" & strDeclarantName & "' ) " & _
                   "AND CD022.RefNo=6"
      Set rsTemp2 = gcnGi.Execute(strString1)
      '是，已存在，要 Update 相關欄位
      If Not rsTemp2.EOF Then
        strString1 = "Update " & GetOwner & "So004 " & _
                     "Set TfnServiceId='" & rsTemp("TfnServiceId") & "', " & _
                     "  TfnStatus='" & rsTemp("TfnStatus") & "' " & _
                     "Where SeqNo='" & rsTemp2("SeqNo") & "' "
        gcnGi.Execute (strString1)
        
        lngIndex = lngBodyErrorCount
        '問題集2785 針對服務啟始日期，先去抓SO004.OldInstDate,若為NULL值時，則改抓SO004.InstDate去做比對----Handsome 2006/09/29
        If (Format(rsTemp2("OldInstDate") & "", "YYYY/MM/DD")) <> "" Then
            If (Format(rsTemp("InstDate") & "", "YYYY/MM/DD")) <> (Format(rsTemp2("OldInstDate") & "", "YYYY/MM/DD")) Then
              Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "服務啟始日期不合!", rsTemp("InstDate") & "", rsTemp2("OldInstDate") & "", rsTemp("RecNo") & "")
            End If
        Else
            If (Format(rsTemp("InstDate") & "", "YYYY/MM/DD")) <> (Format(rsTemp2("InstDate") & "", "YYYY/MM/DD")) Then
              Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "服務啟始日期不合!", rsTemp("InstDate") & "", rsTemp2("InstDate") & "", rsTemp("RecNo") & "")
            End If
        End If
        If (Format(rsTemp("PrDate") & "", "YYYY/MM/DD")) <> (Format(rsTemp2("PrDate") & "", "YYYY/MM/DD")) Then
          Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "服務結束日期不合!", rsTemp("PrDate") & "", rsTemp2("PrDate") & "", rsTemp("RecNo") & "")
        End If
        If lngBodyErrorCount <> lngIndex Then
         lngBodyErrorCount = lngIndex + 1
        End If
      Else
        '否，輸出錯誤 LOG
        
        '先判斷若已拆機且服務起始日期和服務中止日期不符合
        Set rsTemp2 = Nothing
        strString1 = "Select SO004.SeqNo, SO004.InstDate, SO004.PrDate " & _
                     "From " & GetOwner & "So004," & GetOwner & "CD022 " & _
                     "Where SO004.FaciCode=CD022.CodeNo " & _
                     "AND (SO004.CustId=" & rsTemp("CustId") & " ) " & _
                     "And (SO004.CompCode=" & garyGi(9) & ") " & _
                     "And (SO004.Servicetype='P') " & _
                     "And (SO004.PrDate Is Not Null) " & _
                     "And (SO004.FaciSNo='" & strTel & "' ) " & _
                     "And ((upper(SO004.Id)='" & UCase(rsTemp("Id")) & "' ) " & _
                     "  Or (upper(SO004.Id2)='" & UCase(rsTemp("Id")) & "' )) " & _
                     "And (REPLACE(REPLACE(REPLACE(SO004.DeclarantName,'先生',''),'小姐',''),'女士','')='" & strDeclarantName & "' ) " & _
                     "AND CD022.RefNo=6"
        Set rsTemp2 = gcnGi.Execute(strString1)
        If Not rsTemp2.EOF Then
          lngIndex = lngBodyErrorCount
          If (Format(rsTemp("InstDate") & "", "YYYY/MM/DD")) <> (Format(rsTemp2("InstDate") & "", "YYYY/MM/DD")) Then
            Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "服務啟始日期不合!", rsTemp("InstDate") & "", rsTemp2("InstDate") & "", rsTemp("RecNo") & "")
          End If
          If (Format(rsTemp("PrDate") & "", "YYYY/MM/DD")) <> (Format(rsTemp2("PrDate") & "", "YYYY/MM/DD")) Then
            Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "服務結束日期不合!", rsTemp("PrDate") & "", rsTemp2("PrDate") & "", rsTemp("RecNo") & "")
          End If
          If lngBodyErrorCount <> lngIndex Then
            lngBodyErrorCount = lngIndex + 1
          End If
          Set rsTemp2 = Nothing
        Else
          '以 CustId, FaciSNo, Id, DeclarantName 個別判斷，以便能區分錯誤的原因
          Set rsTemp2 = Nothing
        
          '判斷 CustId
          
          strString1 = "Select SO004.SeqNo, SO004.FaciSNo, SO004.Id, SO004.Id2, REPLACE(REPLACE(REPLACE(SO004.DeclarantName,'先生',''),'小姐',''),'女士','') DeclarantName " & _
                       "From " & GetOwner & "So004," & GetOwner & "CD022 " & _
                       "Where SO004.FaciCode=CD022.CodeNo " & _
                       "And (SO004.CompCode=" & garyGi(9) & ") " & _
                       "And (SO004.Servicetype='P') " & _
                       "And (SO004.InstDate Is Not Null) " & _
                       "And (SO004.PrDate Is Null) " & _
                       "AND (SO004.CustId=" & rsTemp("CustId") & " ) " & _
                       "AND CD022.RefNo=6 AND SO004.FACISNO='" & strTel & "'"
                       
          'Set rsTEMP2 = New ADODB.Recordset
'          rsTemp2.CursorLocation = adUseClient
'          rsTemp2.Open strString1, gcnGi, adOpenKeyset, adLockOptimistic
'         If Not GetRS(rsTmp012, strString1, gcnGi, , adOpenKeyset, adLockOptimistic, adUseClient) Then Exit Sub
          If Not GetRS(rsTemp2, strString1, gcnGi, , adOpenKeyset, adLockOptimistic, adUseClient) Then Exit Sub
          lngIndex = lngBodyErrorCount
          If rsTemp2.EOF Then
            '突然說要加入以門號及身分證及名稱來判斷
            Set rsTemp2 = Nothing
            strString1 = "Select SO004.CustId,SO004.SNO,SO004.INSTDATE,SO004.PRDATE " & _
                         "From " & GetOwner & "So004," & GetOwner & "CD022 " & _
                         "Where SO004.FaciCode=CD022.CodeNo " & _
                         "AND (SO004.FaciSNo='" & strTel & "') " & _
                         "And (SO004.CompCode=" & garyGi(9) & ") " & _
                         "And (SO004.Servicetype='P') " & _
                         "And (SO004.InstDate Is Not Null) " & _
                         "And (SO004.PrDate Is Null) " & _
                         "AND CD022.RefNo=6"
            
            Set rsTemp2 = gcnGi.Execute(strString1)
            
            If Not rsTemp2.EOF Then
            '門號在，那就是客編不對了
              Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "客戶編號不合!", rsTemp("CustId") & "", rsTemp2("CustId") & "", rsTemp("RecNo") & "" _
                              , rsTemp2("CustID"), rsTemp2("SNO"), IIf(rsTemp2("INSTDATE") <> "", Format(rsTemp2("INSTDATE"), "YYYY/MM/DD"), "") _
                              , IIf(rsTemp2("PRDATE") <> "", Format(rsTemp2("PRDATE"), "YYYY/MM/DD"), ""))
            Else
            '不然就再比對身分證
              Set rsTemp2 = Nothing
              strString1 = "Select SO004.CustId " & _
                       "From " & GetOwner & "So004," & GetOwner & "CD022 " & _
                       "Where SO004.FaciCode=CD022.CodeNo " & _
                       "AND ((upper(SO004.Id)='" & UCase(rsTemp("Id")) & "' ) " & _
                       "  Or (upper(SO004.Id2)='" & UCase(rsTemp("Id")) & "' )) " & _
                       "And (SO004.CompCode=" & garyGi(9) & ") " & _
                       "And (SO004.Servicetype='P') " & _
                       "And (SO004.InstDate Is Not Null) " & _
                       "And (SO004.PrDate Is Null) " & _
                       "AND CD022.RefNo=6 AND (SO004.FaciSNo='" & strTel & "')"
              Set rsTemp2 = gcnGi.Execute(strString1)
              If Not rsTemp2.EOF Then
                '身分證在，那就是客編不對了
                Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "客戶編號不合!", rsTemp("CustId") & "", rsTemp2("CustId") & "", rsTemp("RecNo") & "")
              Else
                Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "此CP門號不存在!", rsTemp("CustId") & "", "", rsTemp("RecNo") & "")
              End If
            End If
            Set rsTemp2 = Nothing
          Else
            lngIndex = lngBodyErrorCount
            While Not rsTemp2.EOF
              If (rsTemp2("Id") & "" <> rsTemp("Id") & "") And (rsTemp2("Id2") & "" <> rsTemp("Id") & "") Then
                '判斷身分證
                If (rsTemp2("Id") & "") <> "" Then
                  strString1 = rsTemp2("Id") & ""
                Else
                  strString1 = rsTemp2("Id2") & ""
                End If
                Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "身份證字號不合!", rsTemp("Id") & "", strString1 & "", rsTemp("RecNo") & "")
                
              End If
            
              If rsTemp2("DeclarantName") & "" <> strDeclarantName Then
                '判斷客戶名稱
                Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "登記姓名不合!", strDeclarantName & "", rsTemp2("DeclarantName") & "", rsTemp("RecNo") & "")
                
              End If
            '問題集2798 針對第三點，如果找不到客編，反查設備
              If rsTemp2("FaciSNo") & "" = strTel Then
'                '判斷 FaciSNo
                strString1 = "Select SO004.CustId,SO004.SNO,SO004.INSTDATE,SO004.PRDATE " & _
                       "From " & GetOwner & "So004," & GetOwner & "CD022 " & _
                       "Where SO004.FaciCode=CD022.CodeNo " & _
                       "AND ((SO004.Id='" & UCase(rsTemp("Id")) & "' ) " & _
                       "  Or (SO004.Id2='" & rsTemp("Id") & "' )) " & _
                       "And (SO004.CompCode=" & garyGi(9) & ") " & _
                       "And (SO004.Servicetype='P') " & _
                       "And (SO004.InstDate Is Not Null) " & _
                       "And (SO004.PrDate Is Null) " & _
                       "AND CD022.RefNo=6 AND SO004.FaciSNO='" & strTel & "'"
'                Dim rsTEMP2 As New ADODB.Recordset
                Set rsTemp2 = Nothing
                Set rsTemp2 = gcnGi.Execute(strString1)
                '找是否有設備存在，如果存在秀出反查客編
                If rsTemp2.EOF Then
                    Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "客戶設備找不到!", strTel & "", rsTemp2("FaciSNo") & "", rsTemp("RecNo") & "")
                Else
'                    Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "客戶設備找不到!", strTel & "", rsTemp("FaciSNo") & "", rsTemp("RecNo") & "", rsTEMP2("CUSTID") & "", rsTEMP2("SNO") & "", _
'                                    IIf(rsTEMP2("INSTDATE") <> "", Format(rsTEMP2("INSTDATE"), "YYYY/MM/DD"), ""), IIf(rsTEMP2("Prdate") <> "", Format(rsTEMP2("PRDATE"), "YYYY/MM/DD"), ""))

                     Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "客戶設備找不到!", rsTemp("CustId") & "", rsTemp2("CustId") & "", rsTemp("RecNo") & "" _
                              , rsTemp2("CustID"), rsTemp2("SNO"), IIf(rsTemp2("INSTDATE") <> "", Format(rsTemp2("INSTDATE"), "YYYY/MM/DD"), "") _
                              , IIf(rsTemp2("PRDATE") <> "", Format(rsTemp2("PRDATE"), "YYYY/MM/DD"), ""))
                 
                End If
                
              End If
              
'              If chkTFNSID.Value = 1 Then
'                If rsTemp2("TfnServiceId") & "" <> Trim(rsTemp("TFNServiceID") & "") Then
'                  '判斷 TFNServiceID
'                  Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "服務代碼(台固代碼)不合!", strTel & "", rsTemp2("FaciSNo") & "", rsTemp("RecNo") & "")
'                End If
'              End If
              rsTemp2.MoveNext
            Wend
          End If
          lngBodyErrorCount = lngIndex + 1
        End If
      End If
      Set rsTemp2 = Nothing
      
      lngBodyCount = lngBodyCount + 1
      lngLineNumber = lngLineNumber + 1
      rsTemp.MoveNext
    Wend
    Set rsTemp = Nothing
    PB3.Visible = False
    
    If lngBodyErrorCount > 0 Then
      strString1 = "匯入核對筆數：" & lngBodyCount & vbCrLf & _
                   "核對錯誤筆數：" & lngBodyErrorCount & vbCrLf & vbCrLf & _
                   "匯入資料中有錯誤筆數，請參考問題記錄檔。" & vbCrLf
      MsgBox strString1, vbExclamation, "警告"
      Shell "Notepad " & txtLogFile2.Text, vbNormalFocus
    Else
      strString1 = "匯入核對筆數：" & lngBodyCount & vbCrLf & vbCrLf & _
                   "匯入完成，所有資料皆符合無誤 !!" & vbCrLf
      MsgBox strString1, vbExclamation, "訊息"
    End If
  End If
  
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Call ErrSub(Me.Name, "ImportData2")
End Sub

Private Sub SelectImportFile2(ByRef ImportFile As TextBox, ByVal Filters As String, ByRef LogFile As TextBox)
On Error GoTo ChkErr
  With ComdPath
    .FileName = Filters & "*.txt"
    .Filter = "文字檔 (*.txt;*.dat)|*.txt;*.dat|所有檔案 (*.*)|*.*"
    .Action = 1
    If .FileName <> Filters & "*.txt" Then
      ImportFile.Text = .InitDir & .FileName
    End If
  End With
  
  '自動根據文字檔的檔名設定對應的 LOG 檔檔名
  Call SetLogFileName(ImportFile.Text, LogFile)
  
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "SelectImportFile2")
End Sub

Private Sub ImportData3()
On Error GoTo ChkErr
  'CP通話明細資料匯入
  Dim blnGoOn, blnGoOn2 As Boolean
  Dim lngYear, lngMonth As Long
  Dim strTextData, StrTableName, strCheckName, strFileName, strErrStr, strErrData As String
  Dim strString1 As String
  Dim strViewName1 As String
  Dim lngLastLine As Long, lngIndex As Long, lngIndex2 As Long, lngInsertCount As Long, lngUpdateCount As Long
  Dim lngErrorCount As Long
  Dim vDataArray, vElement, vDataArray2, vDataArrayA, vDataArrayB As Variant
  Dim rsTemp As New ADODB.Recordset
  Dim rsTemp2 As New ADODB.Recordset

  Screen.MousePointer = vbHourglass
  Me.Enabled = False
  
  lngErrorCount = 0
  
  strFileName = FSO.GetFileName(txtImportFile5.Text)
  StrTableName = GetOwner & IIf(optDetailDaily.Value, "So168", "So169")
  
  '讀入文字檔，並分段
  strTextData = FSO.OpenTextFile(txtImportFile5).ReadAll
  vDataArray = Split(strTextData, vbCrLf)
  
  '算出最後一行的位置
  lngLastLine = 0
  For lngIndex = UBound(vDataArray) To 1 Step -1
    If Left(vDataArray(lngIndex), 1) = "T" Then
      lngLastLine = lngIndex
      lngIndex = 0
    End If
  Next
  
  'Progress Bar
  PB2.Visible = True
  PB2.Value = PB2.Min
  
  On Error Resume Next
  If lngLastLine > 1 Then
  
    PB2.Max = lngLastLine - 1
  Else
    PB2.Max = 1
  End If
  
  '檢核匯入檔名
  strCheckName = IIf(optDetailDaily.Value, txtImportFile3.Text, txtImportFile4.Text)
  
  blnGoOn = True
  If Left(vDataArray(0), 1) <> "H" Then
    Call LogError1(lngErrorCount, "表頭必須在文字檔的第一行", "", 1, vDataArray(0))
    blnGoOn = False
  End If
  If Left(vDataArray(lngLastLine), 1) <> "T" Then
    Call LogError1(lngErrorCount, "表尾必須在文字檔的最後一行", "", 0, vDataArray(lngLastLine))
    blnGoOn = False
  End If
  If Val(Mid(vDataArray(lngLastLine), 2, 7)) <> lngLastLine - 1 Then
    Call LogError1(lngErrorCount, "資料筆數(" & (lngLastLine - 1) & ")與表尾記載(" & Val(Mid(vDataArray(lngLastLine), 2, 7)) & ")的不符", "", lngLastLine, vDataArray(lngLastLine))
    blnGoOn = False
  End If
  If UCase(Left(strFileName, Len(strCheckName))) <> UCase(strCheckName) Then
    Call LogError1(lngErrorCount, "'匯入檔案名稱' 與 " & "'檢核匯入檔名' 不符", "", 0, "")
    blnGoOn = False
  End If
  
  blnGoOn2 = True
  If optDetailMonthly.Value And blnGoOn Then
    '掃描文字檔內容，看每筆資料的出帳年月是否完全一樣
    blnGoOn2 = True
    strString1 = ""
    For lngIndex = 1 To lngLastLine - 1
      PB2.Value = lngIndex - 1
      vDataArray2 = Split(vDataArray(lngIndex), ",")
      If (strString1 <> "") And (strString1 <> vDataArray2(9)) Then
        Call LogError1(lngErrorCount, "客編：" & vDataArray2(0) & "　文字檔中有不同的出帳年月資料：" & vDataArray2(9), strString1, lngIndex + 1, vDataArray(lngIndex))
        blnGoOn2 = False
      ElseIf strString1 = "" Then
        strString1 = vDataArray2(9)
      End If
    Next
    If blnGoOn2 Then
      '若CDR匯入每月，則檢查出帳年月是否已存在了，若是，將資料庫中同出帳年月的資料刪除
      blnGoOn2 = True
      If optDetailMonthly.Value And blnGoOn Then
        vDataArray2 = Split(vDataArray(1), ",")
        Set rsTemp = gcnGi.Execute("Select CustId From " & GetOwner & "So169 Where Billing_Period=" & vDataArray2(9))
        If Not rsTemp.EOF Then
          blnGoOn2 = MsgBox("出帳年月 " & vDataArray2(9) & "：該月份CDR資料已匯入，是否清空該月份的CDR資料，再匯入？", vbYesNo, "警告") = 6
          If blnGoOn2 Then
            gcnGi.Execute ("Delete From " & GetOwner & "So169 Where Billing_Period=" & vDataArray2(9))
          End If
        End If
        Set rsTemp = Nothing
      End If
    End If
  End If
  
  If blnGoOn2 And blnGoOn Then
    '掃描文字檔內容，看看是否有重複的資料（客編，門號，發話起始時間, 受話方號碼）
    '2006.03.03 改用存暫存檔的方式來做，不然資料多時會很慢
    
    strViewName1 = GetTmpViewName
    On Error Resume Next
    gcnGi.Execute "Drop Table " & GetOwner & strViewName1
    strString1 = "Create Table " & GetOwner & strViewName1 & "(CustId Number(8), Calling_Area_Code Varchar2(9), Calling_Nbr Varchar2(20), Start_Time Date, Called_Area_Code Varchar2(9), Called_Nbr Varchar2(20), RecNo Long)"
    gcnGi.Execute strString1
    strString1 = "Create Index I_Tmp On " & GetOwner & strViewName1 & "(CustId, Calling_Area_Code, Calling_Nbr, Start_Time, Called_Area_Code, Called_Nbr)"
    gcnGi.Execute strString1
    
    For lngIndex = 1 To lngLastLine - 1
      PB2.Value = lngIndex - 1
      vDataArrayA = Split(vDataArray(lngIndex), ",")
      strString1 = "Insert Into " & GetOwner & strViewName1 & " " & _
                   "Values ( " & _
                   vDataArrayA(0) & ", '" & _
                   vDataArrayA(2) & "', '" & _
                   vDataArrayA(3) & "', " & _
                   "To_Date('" & vDataArrayA(10) & "', 'YYYY/MM/DD HH24:MI:SS'), '" & _
                   vDataArrayA(7) & "', '" & _
                   vDataArrayA(8) & "', " & _
                   CStr(lngIndex) & ") "
      gcnGi.Execute strString1
      DoEvents
    Next
    
    strString1 = "Select Distinct CustId, Calling_Area_Code, Calling_Nbr, Start_Time, Called_Area_Code, Called_Nbr From " & GetOwner & strViewName1 & " " & _
                 "Group By CustId, Calling_Area_Code, Calling_Nbr, Start_Time, Called_Area_Code, Called_Nbr " & _
                 "Having Count(CustId)>1 "
    Set rsTemp = gcnGi.Execute(strString1)
    While Not rsTemp.EOF
      strString1 = "Select * From " & GetOwner & strViewName1 & " " & _
                   "Where CustId=" & rsTemp("CustId") & "" & _
                   "And Calling_Area_Code='" & rsTemp("Calling_Area_Code") & "" & "' " & _
                   "And Calling_Nbr='" & rsTemp("Calling_Nbr") & "" & "' " & _
                   "And Start_Time=To_Date('" & rsTemp("Start_Time") & "" & "', 'YYYY/MM/DD HH24:MI:SS') " & _
                   "And Called_Area_Code='" & rsTemp("Called_Area_Code") & "" & "' " & _
                   "And Called_Nbr='" & rsTemp("Called_Nbr") & "" & "' "
      Set rsTemp2 = gcnGi.Execute(strString1)
      If Not rsTemp2.EOF Then
        Call LogError1(lngErrorCount, "客編：" & rsTemp("CustId") & "" & "　CP門號：" & rsTemp("Calling_Area_Code") & "" & "-" & rsTemp("Calling_Nbr") & "" & "　匯入檔案有重複資料", "", CInt(rsTemp2("RecNo") & ""), vDataArray(CLng(rsTemp2("RecNo") & "")))
      End If
      While Not rsTemp2.EOF
        strmTextFile1.Write "       " & vDataArray(CLng(rsTemp2("RecNo") & "")) & vbCrLf
        rsTemp2.MoveNext
        blnGoOn2 = False
      Wend
      Set rsTemp2 = Nothing
      rsTemp.MoveNext
      DoEvents
    Wend
    Set rsTemp = Nothing
                 
    On Error Resume Next
    gcnGi.Execute "Drop Table " & GetOwner & strViewName1
    
    
    'strString1 = ""
    'For lngIndex = 1 To lngLastLine - 2
    '  vDataArrayA = Split(vDataArray(lngIndex), ",")
    '  For lngIndex2 = lngIndex + 1 To lngLastLine - 1
    '    vDataArrayB = Split(vDataArray(lngIndex2), ",")
    '    If (vDataArrayA(0) = vDataArrayB(0)) And (vDataArrayA(2) = vDataArrayB(2)) And (vDataArrayA(3) = vDataArrayB(3)) And (vDataArrayA(10) = vDataArrayB(10)) And (vDataArrayA(7) = vDataArrayB(7)) And (vDataArrayA(8) = vDataArrayB(8)) Then
    '      Call LogError1(lngErrorCount, "客編：" & vDataArrayB(0) & "　CP門號：" & vDataArrayB(2) & "-" & vDataArrayB(3) & "　匯入檔案有重複資料", "", lngIndex2 + 1, vDataArray(lngIndex))
    '      strmTextFile1.Write "       " & vDataArray(lngIndex) & vbCrLf
    '      strmTextFile1.Write "       " & vDataArray(lngIndex2) & vbCrLf
    '      blnGoOn2 = False
    '    End If
    '  Next
    'Next
  End If
  
  gcnGi.BeginTrans
  lngInsertCount = 0
  lngUpdateCount = 0
  If blnGoOn2 And blnGoOn Then
    '若匯入失敗，則將已匯入 SO168/SO169 的資料取消
    blnGoOn = True
    
    For lngIndex = 1 To lngLastLine - 1
      PB2.Value = lngIndex - 1
      vDataArray2 = Split(vDataArray(lngIndex), ",")
      
      If Not IsNumeric(vDataArray2(0)) Then
        Call LogError1(lngErrorCount, "客編：" & vDataArray2(0) & "，格式錯誤", "", lngIndex + 1, vDataArray(lngIndex))
        blnGoOn = False
      End If
      If Not IsDate(vDataArray2(10)) Then
      Call LogError1(lngErrorCount, "客編：" & vDataArray2(0) & "　發話起始時間：" & vDataArray2(10) & "，格式錯誤", "", lngIndex + 1, vDataArray(lngIndex))
        blnGoOn = False
      End If
      If Not IsDate(vDataArray2(11)) Then
        Call LogError1(lngErrorCount, "客編：" & vDataArray2(0) & "　發話終止時間：" & vDataArray2(11) & "，格式錯誤", "", lngIndex + 1, vDataArray(lngIndex))
      blnGoOn = False
      End If
      If Not IsNumeric(vDataArray2(12)) Then
        Call LogError1(lngErrorCount, "客編：" & vDataArray2(0) & "　發話方撥打秒數：" & vDataArray2(12) & "，格式錯誤", "", lngIndex + 1, vDataArray(lngIndex))
        blnGoOn = False
      End If
      If Not IsNumeric(vDataArray2(13)) Then
        Call LogError1(lngErrorCount, "客編：" & vDataArray2(0) & "　發話方計費單位總數：" & vDataArray2(13) & "，格式錯誤", "", lngIndex + 1, vDataArray(lngIndex))
        blnGoOn = False
      End If
      If Not IsNumeric(vDataArray2(14)) Then
        Call LogError1(lngErrorCount, "客編：" & vDataArray2(0) & "　金額：" & vDataArray2(14) & "，格式錯誤", "", lngIndex + 1, vDataArray(lngIndex))
        blnGoOn = False
      End If
      If Not IsNumeric(vDataArray2(16)) Then
        Call LogError1(lngErrorCount, "客編：" & vDataArray2(0) & "　TFN用戶代號：" & vDataArray2(16) & "，格式錯誤", "", lngIndex + 1, vDataArray(lngIndex))
        blnGoOn = False
      End If
    
      If IsDate(vDataArray2(10)) And IsDate(vDataArray2(11)) Then
    
        '若撥打時間一樣但結束時間不同時，要 SHOW 出 MESSAGE 提示
        'strString1 = "Select Start_Time From " & StrTableName & " " & _
                     "Where (Start_Time=To_Date('" & vDataArray2(10) & "', 'YYYY/MM/DD HH24:MI:SS')) " & _
                     "And (End_Time<>To_Date('" & vDataArray2(11) & "', 'YYYY/MM/DD HH24:MI:SS')) " & _
                     "And (Calling_Area_Code='" & vDataArray2(2) & "') " & _
                     "And (Calling_Nbr='" & vDataArray2(3) & "') "
        'Set rsTemp = gcnGi.Execute(strString1)
        'If Not rsTemp.EOF Then
        '  Call LogError1(lngErrorCount, "客編：" & vDataArray2(0) & "　資料庫中有相同的發話時間記錄：" & rsTemp("Start_Time"), "", lngIndex + 1, vDataArray(lngIndex))
        '  blnGoOn = False
        'End If
        'Set rsTemp = Nothing
    
        '判斷撥打時間<結束時間
        If DateDiff("S", vDataArray2(10), vDataArray2(11)) < 0 Then
          Call LogError1(lngErrorCount, "客編：" & vDataArray2(0) & "　CP門號：" & vDataArray2(2) & "-" & vDataArray2(3) & "　發話起始時間必須小於發話終止時間", "", lngIndex + 1, vDataArray(lngIndex))
          blnGoOn = False
        End If
      End If
    
      '檢查客編是否存在
      strString1 = "Select CustId From " & GetOwner & "So001 Where CustId=" & vDataArray2(0)
      Set rsTemp = gcnGi.Execute(strString1)
      If rsTemp.EOF Then
        Call LogError1(lngErrorCount, "客編：" & vDataArray2(0) & " 不存在", "", lngIndex + 1, vDataArray(lngIndex))
        blnGoOn = False
      End If
      Set rsTemp = Nothing
    
      If blnGoOn Then
        '2006.02.20 改為不分每日或每月，以account_id,called_area_code,called_nbr,start_time為key，判斷有沒有重覆
        'strString1 = "Select * From " & StrTableName & " " & _
                     "Where (Start_Time=To_Date('" & vDataArray2(10) & "', 'YYYY/MM/DD HH24:MI:SS')) " & _
                     "And (End_Time=To_Date('" & vDataArray2(11) & "', 'YYYY/MM/DD HH24:MI:SS')) " & _
                     "And (Calling_Area_Code='" & vDataArray2(2) & "') " & _
                     "And (Calling_Nbr='" & vDataArray2(3) & "') "
        strString1 = ""
        strString1 = "Select * From " & StrTableName & " " & _
                     "Where (Account_Id='" & vDataArray2(1) & "') " & _
                     "And (Called_Area_Code='" & vDataArray2(7) & "') " & _
                     "And (Called_Nbr='" & vDataArray2(8) & "') " & _
                     "And (Start_Time=To_Date('" & vDataArray2(10) & "', 'YYYY/MM/DD HH24:MI:SS')) "
'        rsSo168.Open strString1, gcnGi, adOpenKeyset, adLockOptimistic
        If Not GetRS(rsSo168, strString1, gcnGi, , adOpenKeyset, adLockOptimistic, adUseClient) Then Exit Sub
        If rsSo168.EOF Then
          rsSo168.AddNew
          lngInsertCount = lngInsertCount + 1
            rsSo168("CustId") = vDataArray2(0)
            rsSo168("Catv_Nbr") = vDataArray2(0)
            rsSo168("Account_Id") = vDataArray2(1)
            rsSo168("Calling_Area_Code") = vDataArray2(2)
            rsSo168("Calling_Nbr") = vDataArray2(3)
            rsSo168("Bill_Item_Name") = vDataArray2(4)
            rsSo168("Area_Name") = vDataArray2(5)
            rsSo168("Access_Code") = vDataArray2(6)
            rsSo168("Called_Area_Code") = vDataArray2(7)
            rsSo168("Called_Nbr") = vDataArray2(8)
            rsSo168("Billing_Period") = vDataArray2(9)
            rsSo168("Start_Time") = vDataArray2(10)
            rsSo168("End_Time") = vDataArray2(11)
            rsSo168("Duration") = vDataArray2(12)
            rsSo168("Counts") = vDataArray2(13)
            rsSo168("Charge") = vDataArray2(14)
            rsSo168("Hour_Type") = vDataArray2(15)
            rsSo168("Subscr_Id") = vDataArray2(16)
            rsSo168("CreatDate") = Mid(vDataArray(0), 2, 4) & "/" & Mid(vDataArray(0), 6, 2) & "/" & Mid(vDataArray(0), 8, 2)
            rsSo168("UpdTime") = GetDTString(RightNow)
            rsSo168("UpdEn") = garyGi(0)
            rsSo168.Update
            rsSo168.Close
            DoEvents
        Else
          lngUpdateCount = lngUpdateCount + 1
          Call LogError1(lngErrorCount, "客編：" & vDataArray2(0) & "　要匯入的資料已存在於資料庫中", "", lngIndex + 1, vDataArray(lngIndex))
          rsSo168.Close
          'lngErrorCount = lngErrorCount - 1
        End If
      End If
      DoEvents
    Next
    PB2.Visible = False
  End If
  gcnGi.CommitTrans
  If lngErrorCount > 0 Then
    '有錯誤，已匯入的資料須取消匯入
'    gcnGi.RollbackTrans
    strString1 = "文字檔筆數：" & lngLastLine - 1 & vbCrLf & _
                 "核對錯誤/問題數：" & lngErrorCount & vbCrLf & vbCrLf & _
                 "匯入資料中有錯誤/問題筆數，請參考問題記錄檔。" & vbCrLf
    MsgBox strString1, vbExclamation, "警告"
    Shell "Notepad " & txtLogFile3.Text, vbNormalFocus
  Else
    '沒錯誤，寫入資料庫中
    strString1 = "匯入完成 !!" & vbCrLf & vbCrLf & _
                 "文字檔筆數：" & lngLastLine - 1 & vbCrLf & _
                 "新增筆數：" & lngInsertCount & vbCrLf & _
                 "修改筆數：" & lngUpdateCount & vbCrLf
    MsgBox strString1, vbExclamation, "訊息"
  End If
  PB2.Visible = False
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  gcnGi.RollbackTrans
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Call ErrSub(Me.Name, "ImportData3")
End Sub

Private Sub LogError2(ByRef ErrorCounts As Long, ByVal CsCustId, TfnCustId, CsNumber, TfnNumber, CsId, TfnId, CsName, TfnName As String, ByVal LineNumber As Long)
On Error GoTo ChkErr
  '記錄錯誤的資料到 LOG 檔中
  Dim HasError As Boolean
  Dim strString2 As String
  Dim strString1 As String

  HasError = False
  
  '行號
  strString1 = Right("00000" & LineNumber, 5) & "："
  
  '客編
  If (CsCustId <> "") And (TfnCustId <> "") Then
    strString1 = strString1 & "客編 (" & CsCustId & " / " & TfnCustId & ")"
    HasError = True
  End If
  
  'CP 門號
  If (CsNumber <> "") And (TfnNumber <> "") Then
    If HasError Then
      strString1 = strString1 & ", "
    End If
    strString1 = strString1 & "CP 門號 (" & CsNumber & " / " & TfnNumber & ")"
    HasError = True
  End If
  
  '身份證字號
  If (CsId <> "") And (TfnId <> "") Then
    If HasError Then
      strString1 = strString1 & ", "
    End If
    strString1 = strString1 & "身份證字號 (" & CsId & " / " & TfnId & ")"
    HasError = True
  End If
  
  '姓名
  If (CsName <> "") And (TfnName <> "") Then
    If HasError Then
      strString1 = strString1 & ", "
    End If
    strString1 = strString1 & "姓名 (" & CsName & " / " & TfnName & ")"
    HasError = True
  End If
  
  strmTextFile1.Write strString1 & strString2 & vbCrLf
  ErrorCounts = ErrorCounts + 1
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "LogError2")
End Sub

Private Sub txtImportFile2_LostFocus()
On Error GoTo ChkErr
  Call SetLogFileName(txtImportFile2.Text, txtLogFile2)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtImportFile2_LostFocus")
End Sub

Private Sub txtImportFile5_LostFocus()
On Error GoTo ChkErr
  Call SetLogFileName(txtImportFile5.Text, txtLogFile3)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtImportFile5_LostFocus")
End Sub

Private Sub txtShouldDay_LostFocus()
On Error GoTo ChkErr
  Dim blnError As Boolean
  
  blnError = Not IsNumeric(txtShouldDay.Text)
  If Not blnError Then
    If (CInt(txtShouldDay.Text) < 1) Or (CInt(txtShouldDay.Text) > 31) Then
      blnError = True
    End If
  End If
  If blnError Then
    MsgBox "應收日須介於 1～31 之間。", vbExclamation, "警告"
    txtShouldDay.SetFocus
  End If
ChkErr:
  Call ErrSub(Me.Name, "txtShouldDay_LostFocus")
End Sub

Private Sub LogError22(ByRef ErrorCounts As Long, ByVal TFNServiceID, CustID, Number, Reason, TfnValue, CsValue As String, _
                       ByVal LineNumber As Long, Optional ByVal ReturnCust As String = "" _
                       , Optional ByVal SourecSno As String = "", Optional ByVal InstDate As String = "" _
                       , Optional ByVal PRDate As String = "")
  '記錄錯誤的資料到 LOG 檔中
  Dim strString2 As String
  Dim strString1 As String

  '行號
  strString1 = LineNumber & Space(10)
  strString1 = StrConv(LeftB(StrConv(strString1, vbFromUnicode), 10), vbUnicode)
  
  'TFN 客編
  strString2 = TFNServiceID & Space(10)
  strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 10), vbUnicode)
  strString1 = strString1 & strString2
  
  '客編
  strString2 = CustID & Space(10)
  strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 10), vbUnicode)
  strString1 = strString1 & strString2
  
  'CP 門號
  strString2 = Number & Space(20)
  strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 20), vbUnicode)
  strString1 = strString1 & strString2
  
  '訊息
  strString2 = Reason & Space(40)
  strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 40), vbUnicode)
  strString1 = strString1 & strString2
  
  'TFN 值
  If TfnValue = "" Then
    TfnValue = "無"
  End If
  strString2 = TfnValue & Space(20)
  strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 20), vbUnicode)
  strString1 = strString1 & strString2
  
  'CS 值
  If CsValue = "" Then
    CsValue = "無"
  End If
  strString2 = CsValue & Space(20)
  strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 20), vbUnicode)
  strString1 = strString1 & strString2
  '由TNF反查的客編
  If ReturnCust = "" Then ReturnCust = "無"
  strString2 = ReturnCust & Space(25)
  strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 25), vbUnicode)
  strString1 = strString1 & strString2
  
  '來源單號
   If SourecSno = "" Then SourecSno = "無"
   strString2 = SourecSno & Space(20)
   strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 20), vbUnicode)
   strString1 = strString1 & strString2
  '安裝日期
   If InstDate = "" Then InstDate = "無"
   strString2 = InstDate & Space(15)
   strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 15), vbUnicode)
   strString1 = strString1 & strString2
  '拆機日期
   If PRDate = "" Then PRDate = "無"
   strString2 = PRDate & Space(20)
   strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 20), vbUnicode)
   strString1 = strString1 & strString2
   strmTextFile1.Write strString1 & vbCrLf

  ErrorCounts = ErrorCounts + 1
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "LogError22")
End Sub

