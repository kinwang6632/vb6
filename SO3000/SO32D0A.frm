VERSION 5.00
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO32D0A 
   BorderStyle     =   1  '單線固定
   Caption         =   "出單檢核 [SO32D0A]"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11565
   Icon            =   "SO32D0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11565
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.Frame fra4 
      Height          =   2085
      Left            =   1650
      TabIndex        =   66
      Top             =   4800
      Width           =   4245
      Begin VB.CheckBox chkDayAmt 
         Caption         =   "收費資料與週期多階的日額是否相符"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   180
         Width           =   3555
      End
      Begin VB.Frame fra2 
         Height          =   885
         Left            =   60
         TabIndex        =   69
         Top             =   450
         Width           =   4035
         Begin VB.CheckBox chkAfter 
            Caption         =   "1.單據合計總金額<=0 "
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   71
            Top             =   180
            Width           =   2025
         End
         Begin VB.CheckBox chkAfter 
            Caption         =   "2.同一客戶有多筆資料,應收日期同,但帳號不同"
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   70
            Top             =   450
            Width           =   3945
         End
      End
      Begin VB.Frame fra3 
         Height          =   585
         Left            =   90
         TabIndex        =   67
         Top             =   1410
         Width           =   4035
         Begin VB.CheckBox chkFinal 
            Caption         =   "1.媒體單號合計總金額<=0"
            Height          =   315
            Left            =   60
            TabIndex        =   68
            Top             =   180
            Width           =   3105
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1110
      Left            =   3390
      ScaleHeight     =   1050
      ScaleWidth      =   5055
      TabIndex        =   62
      Top             =   2520
      Visible         =   0   'False
      Width           =   5115
      Begin MSComctlLib.ProgressBar pgbBar 
         Height          =   375
         Left            =   165
         TabIndex        =   63
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
         TabIndex        =   64
         Top             =   150
         Width           =   1800
      End
   End
   Begin VB.PictureBox pic2 
      Height          =   2835
      Left            =   60
      ScaleHeight     =   2775
      ScaleWidth      =   11325
      TabIndex        =   16
      Top             =   1560
      Width           =   11385
      Begin VB.VScrollBar vsl1 
         Height          =   2595
         LargeChange     =   50
         Left            =   10980
         Max             =   100
         SmallChange     =   20
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   0
         Width           =   315
      End
      Begin VB.Frame fraWhere 
         BorderStyle     =   0  '沒有框線
         Height          =   4035
         Left            =   30
         TabIndex        =   17
         Top             =   -100
         Width           =   10995
         Begin CS_Multi.CSmulti gimCitemCode 
            Height          =   375
            Left            =   60
            TabIndex        =   35
            Top             =   810
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   661
            ButtonCaption   =   "收費項目"
         End
         Begin Gi_Multi.GiMulti gimClassCode 
            Height          =   375
            Left            =   60
            TabIndex        =   36
            Top             =   1155
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   661
            ButtonCaption   =   "客戶類別"
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin Gi_Multi.GiMulti gimClctArea 
            Height          =   375
            Left            =   60
            TabIndex        =   39
            Top             =   2205
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   661
            ButtonCaption   =   "收  費  區"
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin Gi_Multi.GiMulti gimServCode 
            Height          =   375
            Left            =   60
            TabIndex        =   38
            Top             =   1875
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   661
            ButtonCaption   =   "服  務  區"
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin Gi_Multi.GiMulti gimAreaCode 
            Height          =   375
            Left            =   60
            TabIndex        =   37
            Top             =   1515
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   661
            ButtonCaption   =   "行  政  區"
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin CS_Multi.CSmulti gimStrtCode 
            Height          =   375
            Left            =   60
            TabIndex        =   42
            Top             =   3270
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   661
            ButtonCaption   =   "街道名稱"
         End
         Begin CS_Multi.CSmulti gimMduId 
            Height          =   375
            Left            =   60
            TabIndex        =   41
            Top             =   2910
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   661
            ButtonCaption   =   "大樓名稱"
         End
         Begin Gi_Multi.GiMulti gimCMCode 
            Height          =   375
            Left            =   60
            TabIndex        =   40
            Top             =   2550
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   661
            ButtonCaption   =   "收費方式"
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin Gi_Multi.GiMulti gimServiceType 
            Height          =   375
            Left            =   60
            TabIndex        =   33
            Top             =   150
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   661
            ButtonCaption   =   "服務類別"
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin Gi_Multi.GiMulti gimBillType 
            Height          =   345
            Left            =   60
            TabIndex        =   34
            Top             =   480
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   609
            ButtonCaption   =   "單據類別"
            DataType        =   2
            FldCaption1     =   "單據類別代碼"
            FldCaption2     =   "單據類別名稱"
            DIY             =   -1  'True
            Exception       =   -1  'True
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin CS_Multi.CSmulti gimCreateEn 
            Height          =   345
            Left            =   60
            TabIndex        =   43
            Top             =   3630
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   609
            ButtonCaption   =   "過入人員"
         End
      End
   End
   Begin VB.Frame fraAddrType 
      Caption         =   "異常狀態地址依據"
      Enabled         =   0   'False
      Height          =   855
      Left            =   7920
      TabIndex        =   10
      Top             =   30
      Width           =   1875
      Begin VB.OptionButton optChargeAddress 
         Caption         =   "&A.收費地址"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   270
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton optInstAddress 
         Caption         =   "&B.裝機地址"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   570
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "離開(&X)"
      Height          =   375
      Left            =   4350
      TabIndex        =   32
      Top             =   7110
      Width           =   1065
   End
   Begin VB.CommandButton cmdToExcel 
      Caption         =   "匯成Excel"
      Height          =   375
      Left            =   2910
      TabIndex        =   31
      Top             =   7110
      Width           =   1305
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
      Height          =   375
      Left            =   1380
      TabIndex        =   30
      Top             =   7110
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F5:列印"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   7110
      Width           =   1035
   End
   Begin VB.Frame fraOrder 
      Caption         =   "排序方式"
      Height          =   825
      Left            =   9840
      TabIndex        =   13
      Top             =   60
      Width           =   1635
      Begin VB.OptionButton optSort 
         Caption         =   "異常狀態"
         Height          =   285
         Index           =   1
         Left            =   210
         TabIndex        =   15
         Top             =   510
         Width           =   1125
      End
      Begin VB.OptionButton optSort 
         Caption         =   "客編"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   210
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin VB.Frame fraFormat 
      Caption         =   "檔案格式"
      Height          =   735
      Left            =   8070
      TabIndex        =   54
      Top             =   450
      Visible         =   0   'False
      Width           =   1275
      Begin VB.OptionButton optNew 
         Caption         =   "新式"
         Height          =   255
         Left            =   90
         TabIndex        =   57
         Top             =   540
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton optOld 
         Caption         =   "舊式"
         Height          =   255
         Left            =   60
         TabIndex        =   56
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "收費項目對象"
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
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Width           =   1170
      End
   End
   Begin VB.Frame fra1 
      Height          =   2655
      Left            =   6120
      TabIndex        =   53
      Top             =   4800
      Width           =   5325
      Begin VB.CheckBox chkCustAllot 
         Caption         =   "不檢核客戶指定"
         Height          =   195
         Left            =   3600
         TabIndex        =   65
         Top             =   810
         Width           =   1635
      End
      Begin VB.CheckBox chkCancel1 
         Caption         =   "全部取消"
         Height          =   255
         Left            =   3900
         TabIndex        =   20
         Top             =   180
         Width           =   1185
      End
      Begin VB.CheckBox chkAll1 
         Caption         =   "全選"
         Height          =   225
         Left            =   3120
         TabIndex        =   19
         Top             =   180
         Width           =   765
      End
      Begin VB.CheckBox chkBefore 
         Caption         =   "6.週期的設備序號與設備的設備序號是否相符"
         Height          =   315
         Index           =   5
         Left            =   180
         TabIndex        =   28
         Top             =   2310
         Width           =   4485
      End
      Begin VB.CheckBox chkBefore 
         Caption         =   "5.週期收費項目與設備的速率是否相符"
         Height          =   315
         Index           =   4
         Left            =   180
         TabIndex        =   27
         Top             =   2010
         Width           =   3765
      End
      Begin VB.CheckBox chkBefore 
         Caption         =   "4.週期相同設備正/負項收費資料是否相符"
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   26
         Top             =   1710
         Width           =   3765
      End
      Begin VB.TextBox txtCmcode 
         Height          =   315
         Left            =   2550
         TabIndex        =   25
         Top             =   1410
         Width           =   2535
      End
      Begin VB.CheckBox chkBefore 
         Caption         =   "3.查核週期帳戶別為非虛擬帳號類的異常收費方式"
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   24
         Top             =   1110
         Width           =   4335
      End
      Begin prjNumber.GiNumber ginPeriod 
         Height          =   255
         Left            =   2970
         TabIndex        =   23
         Top             =   780
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   450
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
         AllowMinus      =   -1  'True
      End
      Begin VB.CheckBox chkBefore 
         Caption         =   "2.查核週期異常期數  正常期數："
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   22
         Top             =   780
         Width           =   3855
      End
      Begin VB.CheckBox chkBefore 
         Caption         =   "1.週期多階平均單價與收費項目牌價的平均單價是否相符"
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   450
         Width           =   4785
      End
      Begin VB.Label lblCmcode 
         Caption         =   "請輸入正常收費方式(以,隔開)"
         Height          =   225
         Left            =   210
         TabIndex        =   58
         Top             =   1470
         Width           =   2385
      End
   End
   Begin VB.ComboBox cboCmitem 
      Height          =   300
      ItemData        =   "SO32D0A.frx":0442
      Left            =   5220
      List            =   "SO32D0A.frx":044F
      Style           =   2  '單純下拉式
      TabIndex        =   9
      Top             =   810
      Width           =   1815
   End
   Begin VB.ComboBox cboStatus 
      Height          =   300
      ItemData        =   "SO32D0A.frx":046B
      Left            =   5220
      List            =   "SO32D0A.frx":0478
      Style           =   2  '單純下拉式
      TabIndex        =   7
      Top             =   90
      Width           =   1815
   End
   Begin VB.ComboBox cboPeriod 
      Height          =   300
      ItemData        =   "SO32D0A.frx":049A
      Left            =   5220
      List            =   "SO32D0A.frx":04A7
      Style           =   2  '單純下拉式
      TabIndex        =   8
      Top             =   450
      Width           =   1815
   End
   Begin VB.TextBox txtCustId 
      Height          =   285
      Left            =   1650
      TabIndex        =   18
      Top             =   4440
      Width           =   9765
   End
   Begin Gi_Date.GiDate gdaClctDate1 
      Height          =   315
      Left            =   1050
      TabIndex        =   1
      Top             =   420
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1050
      TabIndex        =   0
      Top             =   60
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
      F2Corresponding =   -1  'True
   End
   Begin Gi_Date.GiDate gdaClctDate2 
      Height          =   315
      Left            =   2730
      TabIndex        =   2
      Top             =   420
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Gi_Date.GiDate gdaShouldDate1 
      Height          =   315
      Left            =   1050
      TabIndex        =   3
      Top             =   780
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Left            =   2730
      TabIndex        =   4
      Top             =   780
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Gi_Time.GiTime gdaCreateTime1 
      Height          =   315
      Left            =   1050
      TabIndex        =   5
      Top             =   1170
      Width           =   1875
      _ExtentX        =   3307
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
   Begin Gi_Time.GiTime gdaCreateTime2 
      Height          =   315
      Left            =   3240
      TabIndex        =   6
      Top             =   1170
      Width           =   1875
      _ExtentX        =   3307
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
   Begin VB.Label Label4 
      Caption         =   "~"
      Height          =   285
      Left            =   3000
      TabIndex        =   61
      Top             =   1230
      Width           =   195
   End
   Begin VB.Label lblCreateTime 
      AutoSize        =   -1  'True
      Caption         =   "出單日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   180
      TabIndex        =   60
      Top             =   1230
      Width           =   780
   End
   Begin VB.Label Label2 
      Caption         =   "~"
      Height          =   285
      Left            =   2400
      TabIndex        =   52
      Top             =   870
      Width           =   195
   End
   Begin VB.Label lblShouldDate 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   180
      TabIndex        =   51
      Top             =   840
      Width           =   780
   End
   Begin VB.Label lblCmitem 
      AutoSize        =   -1  'True
      Caption         =   "收費項目對象"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3975
      TabIndex        =   50
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "出單狀態"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4350
      TabIndex        =   49
      Top             =   150
      Width           =   780
   End
   Begin VB.Label lblPeriod 
      AutoSize        =   -1  'True
      Caption         =   "週期性收費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4170
      TabIndex        =   48
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblCustId 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號(以,相隔)"
      Height          =   195
      Left            =   90
      TabIndex        =   47
      Top             =   4500
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "~"
      Height          =   285
      Left            =   2400
      TabIndex        =   46
      Top             =   480
      Width           =   195
   End
   Begin VB.Label lblClctDate 
      AutoSize        =   -1  'True
      Caption         =   "下收日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   180
      TabIndex        =   45
      Top             =   480
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "公司別"
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
      Left            =   360
      TabIndex        =   44
      Top             =   120
      Width           =   585
   End
End
Attribute VB_Name = "frmSO32D0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#4054需求 By Kin 2008/09/18
Option Explicit
Private rsQry As New ADODB.Recordset
Private strQry As String
Private strStatusTable As String
Private strStatusWhere As String
Private blnExcel As Boolean
Private rsExcel As New ADODB.Recordset
Private strSO003AWhere As String
Private Sub subChoose()
  On Error GoTo chkErr
    strChoose = Empty
    Select Case cboStatus.ListIndex
        Case 0
            If gimServiceType.GetQryStr <> "" Then subAnd "SO003.ServiceType " & gimServiceType.GetQryStr
            If gimCitemCode.GetQryStr <> "" Then subAnd "SO003.CitemCode " & gimCitemCode.GetQryStr
            
            If gimClassCode.GetQryStr <> "" Then subAnd "SO001.ClassCode1 " & gimClassCode.GetQryStr
            If gimAreaCode.GetQryStr <> "" Then subAnd "SO014.AreaCode " & gimAreaCode.GetQryStr
            If gimServCode.GetQryStr <> "" Then subAnd "SO001.ServCode " & gimServCode.GetQryStr
            If gimClctArea.GetQryStr <> "" Then subAnd "SO001.ClctAreaCode " & gimClctArea.GetQryStr
            If gimCMCode.GetQryStr <> "" Then subAnd "SO003.CMCode " & gimCMCode.GetQryStr
            If gimMduId.GetQryStr <> "" Then
                subAnd "SO001.MduId " & gimMduId.GetQryStr
            Else
                If gimMduId.GetDispStr <> "" Then subAnd "So001.MduId Is Not Null"
            End If
            If gdaClctDate1.GetValue & "" <> "" Then subAnd GetCompareDate("SO003.ClctDate ", gdaClctDate1.GetValue, True)
            If gdaClctDate1.GetValue & "" <> "" Then subAnd GetCompareDate("SO003.ClctDate ", gdaClctDate2.GetValue, False)
            If gimStrtCode.GetQryStr <> "" Then subAnd "SO014.StrtCode " & gimStrtCode.GetQryStr
            '#4156 也要過濾SO003.StopFlag By Kin 2008/10/21
            Select Case cboPeriod.ListIndex
                Case 0
                    'subAnd "CD019.StopFlag<>1"
                    subAnd "SO003.StopFlag<>1"
                Case 1
                    'subAnd "CD019.StopFlag=1"
                    subAnd "SO003.StopFlag=1"
            End Select
            If txtCustId.Text <> "" Then
                Call TimetxtCustId(txtCustId, "So003")
            End If
            If optInstAddress.Value Then
                subAnd "SO001.InstAddrNo = SO014.AddrNo "
            Else
                subAnd "SO001.ChargeAddrNo = SO014.AddrNo "
            End If
            '#4258 出單前要過濾CD019. NoShowCitem=1不要顯示異常 By Kin 2008/12/18
            subAnd "Nvl(CD019. NoShowCitem,0)<>1"
            'IIf(optInstAddress.Value, subAnd(" And SO001.InstAddrNo=SO014.AddrNo) ", " And SO001.ChargeAddrNo=SO014.AddrNo "))
        Case 1
            If gimServiceType.GetQryStr <> "" Then subAnd "SO032.ServiceType " & gimServiceType.GetQryStr
            If gimCitemCode.GetQryStr <> "" Then subAnd "SO032.CitemCode " & gimCitemCode.GetQryStr
            If txtCustId.Text <> "" Then
                Call TimetxtCustId(txtCustId, "So032")
            End If

        Case 2
            If gimServiceType.GetQryStr <> "" Then subAnd "SO033.ServiceType " & gimServiceType.GetQryStr
            If gimBillType.GetQryStr <> "" Then subAnd "substr(SO033.BillNo,7,1) " & gimBillType.GetQryStr
            If gimCMCode.GetQryStr <> "" Then subAnd "SO033.CMCode " & gimCMCode.GetQryStr
            'If txtCreateEn.Text <> "" Then subAnd "SO033.CreateEn='" & txtCreateEn.Text & "'"
            If gimCreateEn.GetQryStr <> "" Then subAnd "SO033.CreateEn " & gimCreateEn.GetQryStr
            If gdaShouldDate1.GetValue & "" <> "" Then subAnd GetCompareDate("SO033.ShouldDate ", gdaShouldDate1.GetValue, True)
            If gdaShouldDate2.GetValue & "" <> "" Then subAnd GetCompareDate("SO033.ShouldDate ", gdaShouldDate2.GetValue, False)
            If gimCitemCode.GetQryStr <> "" Then subAnd "SO033.CitemCode " & gimCitemCode.GetQryStr
            If gdaCreateTime1.GetValue <> "" Then Call subAnd("SO033.CreateTime >= To_Date('" & gdaCreateTime1.GetValue & "','YYYYMMDDHH24MISS')")
            If gdaCreateTime2.GetValue <> "" Then Call subAnd("SO033.CreateTime <= To_Date('" & gdaCreateTime2.GetValue & "','YYYYMMDDHH24MISS')")
            If txtCustId.Text <> "" Then
                Call TimetxtCustId(txtCustId, "So033")
            End If
    End Select
    Select Case cboCmitem.ListIndex
        Case 0
            subAnd "CD019.Sign='+'"
        Case 1
            subAnd "CD019.Sign='-'"
    End Select
    '#4158 要過濾收費項目參考號<>18 By Kin 2008/10/21
    '#4156 RefNo is Null也要出來 By Kin 2008/10/29
    subAnd "(CD019.RefNo<>18 Or CD019.RefNo is Null)"
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "subChoose")
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
Private Sub BeforeStatus06()
  On Error GoTo chkErr
    strQry = Empty
    Call subChoose
    '#4137 匯出Excel要增加SO003A.BpCode、SO003A.BpName欄位 By Kin 2008/10/09
    strQry = "Select Distinct SO003.CUSTID,SO003.CitemCode,SO003.CitemName,SO003.FaciSeqNo," & _
             "SO003.FaciSNO,SO003.Period,SO003.ClctDate,SO003.Amount,SO003A.BpCode,SO003A.BpName" & _
              " From " & strStatusTable & "," & GetOwner & "CD019," & GetOwner & "SO003A" & _
                " Where " & strStatusWhere & " And Not Exists(Select * From " & GetOwner & "SO004" & _
                " Where SO004.CustId=SO003.CustId And Nvl(FaciSno,'X')=Nvl(SO003.FaciSNo,'X'))" & _
                " And SO003.FaciSNo is Not Null And SO003.CitemCode=CD019.CodeNo" & _
                IIf(strChoose <> Empty, " And " & strChoose, "") & strSO003AWhere
    If Not GetRS(rsQry, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "BeforeStatus06")
End Sub
Private Sub BeforeStatus05()
  On Error GoTo chkErr
    Dim strTmpQry As String
    strQry = Empty
    Call subChoose
    '#4137 匯出Excel要增加SO003A.BpCode、SO003A.BpName欄位 By Kin 2008/10/09
    strQry = "Select Distinct SO003.CUSTID,SO003.CitemCode,SO003.CitemName,SO003.FaciSeqNo,SO003.FaciSNO,SO003.Period," & _
            "SO003.ClctDate,SO003.Amount,CD019.CMCode,CD019.CMName,SO004.CMBaudRate,SO003A.BpCode,SO003A.BpName " & _
            " From " & strStatusTable & "," & GetOwner & "CD019," & GetOwner & "SO004," & GetOwner & "SO003A " & _
            " Where " & strStatusWhere & _
            IIf(strChoose <> Empty, " And " & strChoose, "") & _
            " And CD019.CodeNo=SO003.CitemCode And SO004.SeqNO=SO003.FaciSeqNo" & _
            " And SO004.CMBaudNo<>CD019.CMCode" & _
            strSO003AWhere
    '#5064 增加判斷工單SignDate is Null的客編不要顯示出來 By Kin 2009/05/20
    strTmpQry = "Select SO007.CustId" & _
                " From " & strStatusTable & "," & GetOwner & "CD019," & GetOwner & "SO004," & GetOwner & "SO003A, " & _
                GetOwner & "SO004D," & GetOwner & "SO007," & GetOwner & "CD005 " & _
                " Where " & strStatusWhere & _
                IIf(strChoose <> Empty, " And " & strChoose, "") & _
                " And CD019.CodeNo=SO003.CitemCode And SO004.SeqNO=SO003.FaciSeqNo" & _
                " And SO004.CMBaudNo<>CD019.CMCode" & _
                strSO003AWhere & _
                " And SO004D.SeqNo=SO004.SeqNo And SO007.SNO=SO004D.SNO " & _
                " And SO007.InstCode=CD005.CodeNo And SO007.SignDate Is Null " & _
                " And CD005.REFNO IN(10,11) And SO007.CustId=SO003.CustId " & _
                " And SO004D.CustId=SO003.CustId "
    strQry = strQry & " And SO003.CustId Not In(" & strTmpQry & ")"
    
    If Not GetRS(rsQry, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "BeforeStatus05")
End Sub
Private Sub FinalStatus01()
  On Error GoTo chkErr
    strQry = Empty
    '#5460 把Group ServiceType拿掉 By Kin 2010/08/23
    Call subChoose
    strQry = "Select SO033.CustId,Null ServiceType,SO033.MediabillNo BillNo,Sum(SO033.ShouldAmt) Total " & _
            " From " & strStatusTable & " Where " & strStatusWhere & _
            IIf(strChoose <> Empty, " And " & strChoose, "") & " And SO033.MediaBillNo is not Null " & _
            " Group By SO033.MediaBillNo,SO033.CustId " & _
            " Having Sum(SO033.ShouldAmt)<1"
    If Not GetRS(rsQry, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "FinalStatus01")
End Sub
Private Sub AfterStatus02()
  On Error GoTo chkErr
    strQry = Empty
    Call subChoose
    strQry = "Select B.CustId,B.BillNo,B.ShouldDate,B.AccountNo From (" & _
    "Select SO032.ServiceType,SO032.CustId,SO032.ShouldDate From " & strStatusTable & " Where " & strStatusWhere & _
            IIf(strChoose <> Empty, " And " & strChoose, "") & _
            " Group By SO032.CustId,SO032.ShouldDate,SO032.ServiceType " & _
            " Having Count(SO032.ShouldDate)>1 And Max(Nvl(SO032.AccountNo,'X'))<>Min(Nvl(SO032.AccountNo,'X'))) A," & GetOwner & "SO032 B" & _
            " Where A.CustId=B.CustId And A.ShouldDate=B.ShouldDate And A.ServiceType=B.ServiceType"
    If Not GetRS(rsQry, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "AfterStatus02")
End Sub
Private Sub AfterStatus01()
  On Error GoTo chkErr
    strQry = Empty
    Call subChoose
    strQry = "Select SO032.CustId,SO032.BillNo,Sum(SO032.ShouldAmt) Total From " & strStatusTable & _
            " Where " & strStatusWhere & IIf(strChoose <> Empty, " And " & strChoose, "") & _
            " Group By SO032.CustId,SO032.BillNo Having Sum(SO032.ShouldAmt)<=0"
    If Not GetRS(rsQry, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "AfterStatus01")
End Sub
Private Sub BeforeStatus04()
  On Error GoTo chkErr
    Dim strSubQry1 As String
    Dim strSubQry2 As String
    Dim strSubQry3 As String
    Dim strSubQry4 As String
    Dim strTmpSubQry3 As String
    strQry = Empty
    
    Call subChoose
    '#4137 匯出Excel要增加SO003A.BpCode、SO003A.BpName欄位 By Kin 2008/10/13
    strSubQry1 = "Select nvl(SO003.Bpcode,'X') BpCode,Nvl(SO003.ContNo,'X') ContNo,SO003.FaciSeqNo,SO003.CUSTID,SO003.ServiceType" & _
                    " From " & strStatusTable & "," & GetOwner & "CD019" & _
                    " Where " & strStatusWhere & " And SO003.CitemCode=CD019.CodeNo " & _
                    IIf(strChoose <> Empty, " And " & strChoose, "") & _
                    " Group By SO003.BpCode,SO003.ContNo,SO003.Faciseqno,SO003.CustId,SO003.ServiceType " & _
                    " Having Count(SO003.Faciseqno)>1 and Max(SO003.CMcode)<>Min(SO003.CMcode)"
    strSubQry1 = "Select Distinct SO003.CUSTID,SO003.CitemCode,SO003.CitemName,SO003.FaciSeqNo,SO003.FaciSNO,SO003.Period," & _
                    "SO003.ClctDate,SO003.Amount,SO003.CmCode,SO003.CmName,SO003A.BpCode,SO003A.BpName,0 CMBaudNo,'X' CMBaudRate, 1 Error " & _
                    " From (" & strSubQry1 & ") A," & GetOwner & "SO003," & GetOwner & "SO003A" & _
                    " Where A.CustId=SO003.CustId And Nvl(SO003.BpCode,'X')=A.Bpcode" & _
                    " And Nvl(SO003.ContNo,'X')=A.ContNo And SO003.FaciSeqNo=A.FaciSeqNo" & _
                    " And SO003.ServiceType=A.ServiceType" & _
                    strSO003AWhere
                    
    strSubQry2 = "Select nvl(SO003.Bpcode,'X') BpCode,Nvl(SO003.ContNo,'X') ContNo,SO003.FaciSeqNo,SO003.CUSTID,SO003.ServiceType" & _
                    " From " & strStatusTable & "," & GetOwner & "CD019" & _
                    " Where " & strStatusWhere & " And SO003.CitemCode=CD019.CodeNo " & _
                    IIf(strChoose <> Empty, " And " & strChoose, "") & _
                    " Group By SO003.BpCode,SO003.ContNo,SO003.Faciseqno,SO003.CustId,SO003.ServiceType " & _
                    " Having Count(SO003.Faciseqno)>1 and Max(SO003.Period)<>Min(SO003.Period)"
    strSubQry2 = "Select Distinct SO003.CUSTID,SO003.CitemCode,SO003.CitemName,SO003.FaciSeqNo,SO003.FaciSNO,SO003.Period," & _
                 "SO003.ClctDate,SO003.Amount,SO003.CmCode,SO003.CmName,SO003A.BpCode,SO003A.BpName,0 CMBaudNo,'X' CMBaudRate, 2 Error" & _
                " From (" & strSubQry2 & ") A," & GetOwner & "SO003," & GetOwner & "SO003A" & _
                " Where A.CustId=SO003.CustId And Nvl(SO003.BpCode,'X')=A.Bpcode" & _
                " And Nvl(SO003.ContNo,'X')=A.ContNo And SO003.FaciSeqNo=A.FaciSeqNo" & _
                " And SO003.ServiceType=A.ServiceType" & _
                strSO003AWhere
    strSubQry3 = "Select nvl(SO003.Bpcode,'X') BpCode,Nvl(SO003.ContNo,'X') ContNo,SO003.FaciSeqNo,SO003.CUSTID,SO003.ServiceType" & _
                    " From " & strStatusTable & "," & GetOwner & "CD019" & _
                    " Where " & strStatusWhere & " And SO003.CitemCode=CD019.CodeNo " & _
                    IIf(strChoose <> Empty, " And " & strChoose, "") & _
                    " Group By SO003.BpCode,SO003.ContNo,SO003.Faciseqno,SO003.CustId,SO003.ServiceType " & _
                    " Having Count(SO003.Faciseqno)>1 and Max(SO003.ClctDate)<>Min(SO003.ClctDate)"
    strSubQry3 = "Select Distinct SO003.CUSTID,SO003.CitemCode,SO003.CitemName,SO003.FaciSeqNo,SO003.FaciSNO,SO003.Period," & _
                 "SO003.ClctDate,SO003.Amount,SO003.CmCode,SO003.CmName,SO003A.BpCode,SO003A.BpName,0 CMBaudNo,'X' CMBaudRate, 3 Error" & _
                " From (" & strSubQry3 & ") A," & GetOwner & "SO003," & GetOwner & "SO003A" & _
                " Where A.CustId=SO003.CustId And Nvl(SO003.BpCode,'X')=A.Bpcode" & _
                " And Nvl(SO003.ContNo,'X')=A.ContNo And SO003.FaciSeqNo=A.FaciSeqNo" & _
                " And SO003.ServiceType=A.ServiceType" & _
                " And SO003.Custid In (Select SO003A.CustId From (" & strSubQry3 & ") A, " & _
                GetOwner & "CD019," & GetOwner & "SO003A," & GetOwner & "SO003" & _
                " Where A.CustId=SO003.CustId And Nvl(SO003.BpCode,'X')=A.Bpcode" & _
                " And Nvl(SO003.ContNo,'X')=A.ContNo And A.FaciSeqno=SO003.FaciSeqNo " & _
                " And A.CustId=SO003A.Custid And SO003A.ServiceType=A.ServiceType" & _
                " And SO003A.FaciSeqNO=A.FaciSeqNO And SO003A.CitemCode=SO003.CitemCode" & _
                " And SO003.CitemCode=CD019.CodeNO And CD019.Sign='-' " & _
                " And NOT Exists(Select CustId From " & GetOwner & "SO003 Where" & _
                                 " So003a.expiretype in(3,4) and SO003a.Custid=SO003.Custid " & _
                                 " And SO003a.CitemCode=SO003.Citemcode and so003.Faciseqno=so003.Faciseqno" & _
                                 " And so003.Servicetype=so003a.Servicetype group by custid" & _
                                 " Having Max(so003a.discountdate2) < Max(so003.clctdate)))" & _
                                 strSO003AWhere
    
                
    strSubQry4 = "Select nvl(SO003.Bpcode,'X') BpCode,Nvl(SO003.ContNo,'X') ContNo,SO003.FaciSeqNo,SO003.CUSTID,SO003.ServiceType" & _
                    " From " & strStatusTable & "," & GetOwner & "CD019" & _
                    " Where " & strStatusWhere & " And SO003.CitemCode=CD019.CodeNo " & _
                    IIf(strChoose <> Empty, " And " & strChoose, "") & _
                    " Group By SO003.BpCode,SO003.ContNo,SO003.Faciseqno,SO003.CustId,SO003.ServiceType " & _
                    " Having Count(SO003.Faciseqno)>1 and Max(CD019.CMCODE)<>Min(CD019.CMCODE)"
    strSubQry4 = "Select Distinct SO003.CUSTID,SO003.CitemCode,SO003.CitemName,SO003.FaciSeqNo,SO003.FaciSNO,SO003.Period," & _
                "SO003.ClctDate,SO003.Amount,SO003.CmCode,SO003.CmName,SO003A.BpCode,SO003A.BpName,CD019.CMCode CMBaudNo,CD019.CMNAME CMBaudRate,4 Error" & _
                " From (" & strSubQry4 & ") A," & GetOwner & "SO003," & GetOwner & "CD019," & GetOwner & "SO003A " & _
                " Where A.CustId=SO003.CustId And Nvl(SO003.BpCode,'X')=A.Bpcode" & _
                " And Nvl(SO003.ContNo,'X')=A.ContNo And SO003.FaciSeqNo=A.FaciSeqNo" & _
                " And SO003.ServiceType=A.ServiceType And SO003.CitemCode=CD019.CodeNo" & _
                strSO003AWhere

    strQry = strSubQry1 & " Union All " & strSubQry2 & " Union All " & strSubQry3 & " Union All " & strSubQry4
    If Not GetRS(rsQry, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "BeforeStatus04")
End Sub
Private Sub BeforeStatus03()
  On Error GoTo chkErr
    strQry = Empty
    Call subChoose
    '#4137 匯出Excel要增加SO003A.BpCode、SO003A.BpName欄位 By Kin 2008/10/09
    '#5267 原本是串SO002A,現在改串SO106
'    strQry = "Select Distinct SO003.CUSTID,SO003.CitemCode,SO003.CitemName,SO003.FaciSeqNo," & _
'             "SO003.FaciSNO,SO003.Period,SO003.ClctDate,SO003.Amount,SO003.CmCode,SO003.CmName,SO003A.BpCode,SO003A.BpName" & _
'            " From " & GetOwner & "SO002A," & GetOwner & "CD019," & GetOwner & "SO003A," & strStatusTable & _
'            " Where SO003.AccountNo is Not Null" & _
'            " And SO002A.AccountNo=SO003.AccountNo And SO002A.CompCode=SO003.CompCode" & _
'            " And SO002A.CustId=SO003.CustId And SO002A.ID<>2 And SO003.CmCode not in (" & txtCmcode.Text & ") And " & strStatusWhere & _
'            IIf(strChoose <> Empty, " And " & strChoose, "") & " And SO003.CitemCode=CD019.CodeNo " & _
'            strSO003AWhere
    
    strQry = "Select Distinct SO003.CUSTID,SO003.CitemCode,SO003.CitemName,SO003.FaciSeqNo," & _
             "SO003.FaciSNO,SO003.Period,SO003.ClctDate,SO003.Amount,SO003.CmCode,SO003.CmName,SO003A.BpCode,SO003A.BpName" & _
            " From " & GetOwner & "SO106," & GetOwner & "CD019," & _
            GetOwner & "SO003A," & GetOwner & "CD031," & strStatusTable & _
            " Where SO003.AccountNo is Not Null" & _
            " And SO106.AccountID=SO003.AccountNo And SO106.CompCode=SO003.CompCode" & _
            " And SO106.CustId=SO003.CustId And SO003.CmCode not in (" & txtCmcode.Text & ") And " & strStatusWhere & _
            " And NVL(CD031.REFNO,0) IN(2,4) AND SO106.CMCode=CD031.CodeNo " & _
            IIf(strChoose <> Empty, " And " & strChoose, "") & " And SO003.CitemCode=CD019.CodeNo " & _
            strSO003AWhere
    If Not GetRS(rsQry, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "BeforeStatus03")
End Sub

Private Sub BeforeStatus02()
  On Error GoTo chkErr
    strQry = Empty
    Call subChoose

    '#4137 匯出Excel要增加SO003A.BpCode、SO003A.BpName欄位 By Kin 2008/10/09
    '#4258 客戶指定異常不要列出 By Kin 2008/12/18
    strQry = "Select Distinct SO003.CUSTID,SO003.CitemCode,SO003.CitemName,SO003.FaciSeqNo," & _
             "SO003.FaciSNO,SO003.Period,SO003.ClctDate,SO003.Amount,SO003A.BpCode,SO003A.BpName" & _
            " From " & GetOwner & "CD019," & GetOwner & "SO003A," & strStatusTable & _
            " Where SO003.Period<>" & ginPeriod.Text & " And " & strStatusWhere & _
            IIf(strChoose <> Empty, " And " & strChoose, "") & " And SO003.CitemCode=CD019.CodeNo" & _
            strSO003AWhere & IIf(chkCustAllot.Value = 1, " And Nvl(SO003.CustAllot,0)<>1 ", "")

    If Not GetRS(rsQry, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
            
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "BeforeStatus02")
End Sub
Private Sub BeforeStatus01()
  On Error GoTo chkErr
    strQry = Empty
    Call subChoose
    '#4137 匯出Excel要增加SO003A.BpCode、SO003A.BpName欄位 By Kin 2008/10/09
    strQry = "Select SO003.CustId,SO003.CitemCode,SO003.CitemName,SO003.FaciSeqNo," & _
             "SO003.FaciSNo,SO003.ClctDate,SO003A.Period,SO003A.DiscountAmt," & _
             "SO003A.DiscountDate1,SO003A.DiscountDate2,SO003A.BpCode,SO003A.BpName" & _
             " From " & GetOwner & "CD078A," & GetOwner & "SO003A," & GetOwner & "CD019," & strStatusTable & _
            " Where SO003A.FaciSeqNo=SO003.FaciSeqNo And SO003A.BPCODE is Not Null" & _
            " And SO003A.ServiceType=SO003.ServiceType And SO003A.CompCode=SO003.CompCode" & _
            " And SO003A.CustId=SO003.CustId And SO003.CitemCode=CD019.CODENO And SO003A.CitemCode=SO003.CitemCode" & _
            " AND SO003.ClctDate<=SO003A.DiscountDate2 And SO003.ClctDate>=SO003A.DiscountDate1" & _
            " And CD078A.BpCode=SO003A.BpCode And CD078A.CitemCode=SO003A.CitemCode" & _
            " And SO003A.Period<>0 And CD019.Period<>0" & _
            " AND Round(SO003A.DiscountAmt/SO003A.Period)<>decode(nvl(CD078A.MonthAmt,0),0,Round((CD019.Amount/CD019.Period)),Round(CD078A.MonthAmt)) And " & strStatusWhere & _
            IIf(strChoose <> Empty, " And " & strChoose, "") & " And SO003A.StopFlag<>1"

    strQry = strQry & " Union " & _
            "Select SO003.CustId,SO003.CitemCode,SO003.CitemName,SO003.FaciSeqNo," & _
            "SO003.FaciSNo,SO003.ClctDate,SO003A.Period,SO003A.DiscountAmt," & _
            "SO003A.DiscountDate1,SO003A.DiscountDate2,SO003A.BpCode,SO003A.BpName" & _
             " From " & GetOwner & "CD019," & GetOwner & "SO003A," & strStatusTable & _
            " Where SO003A.FaciSeqNo=SO003.FaciSeqNo And SO003A.BpCode is Null" & _
            " And SO003A.ServiceType=SO003.ServiceType And SO003A.CompCode=SO003.CompCode" & _
            " And SO003A.CustId=SO003.CustId And SO003.CitemCode=CD019.CODENO And SO003A.CitemCode=SO003.CitemCode" & _
            " AND SO003.ClctDate<=SO003A.DiscountDate2 And SO003.ClctDate>=SO003A.DiscountDate1" & _
            " And SO003A.Period<>0 And CD019.Period<>0" & _
            " And Round(SO003A.DiscountAmt/SO003A.Period)<>Round((CD019.Amount/CD019.Period)) And " & strStatusWhere & _
            IIf(strChoose <> Empty, " And " & strChoose, "") & " And SO003A.StopFlag<>1 "
    strQry = "Select * From (" & strQry & ") Group by CustId,CitemCode,CitemName,FaciSeqNo,ClctDate,Period,DiscountAmt,DiscountDate1,DiscountDate2,FaciSNO,BpCode,BpName"
            
    If Not GetRS(rsQry, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
' And Round(SO003A.MonthAmt*SO003A.Period)<>Round((CD019.Amount/CD019.Period)*SO003.Period) And " & strStatusWhere
' AND Round(SO003A.MonthAmt*SO003A.Period)<>decode(nvl(CD078A.MonthAmt,0),0,Round((CD019.Amount/CD019.Period)*SO003.Period),Round(CD078A.MonthAmt*SO003.Period)) And " & strStatusWhere
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "BeforeStatus01")
End Sub
Private Sub CreateTable()
  On Error Resume Next
    cnn.Close
    Set cnn = Nothing
    Set cnn = GetTmpMdbCn
    cnn.Execute "Drop Table SO32D0A"
    
  On Error GoTo chkErr
  
    cnn.Execute "Create Table SO32D0A (Anomaly text(50),CustId Long default 0," & _
                "CitemCode Long default 0,CitemName text(50),FaciSeqNo text(50)," & _
                "FaciSNO text(50),ClctDate text(10),Period Long default 0," & _
                "Amount Long default 0,Other text(50),BillNo text(50),ShouldAmt Long default 0," & _
                "AccountNo text(50),ShouldDate text(50),ServiceType text(10),BpCode text(10),BpName text(100))"
   Exit Sub
chkErr:
   Call ErrSub(Me.Name, "CreateTable")
End Sub
Private Sub RunStatus3()
  On Error GoTo chkErr
    ReadyGoPrint
    Me.Enabled = False
    Call CreateTable
    cnn.BeginTrans
        If chkFinal.Value Then
            Call FinalStatus01
            SendSQL rsQry.Source, True
            If rsQry.RecordCount > 0 Then
                Do While Not rsQry.EOF
                    cnn.Execute "Insert Into SO32D0A (Anomaly,CustId,BillNo,ShouldAmt,ServiceType) Values " & _
                                "('媒體單號合計總金額<=0'," & rsQry("CustId") & ",'" & rsQry("BillNo") & "'," & rsQry("Total") & _
                                ",'" & rsQry("ServiceType") & "')"
                    pgbBar.Value = (rsQry.AbsolutePosition / rsQry.RecordCount) * 100
                    rsQry.MoveNext
                Loop
            End If
        End If
        pgbBar.Value = 100
    cnn.CommitTrans
    Picture1.Visible = False
    If cnn.Execute("Select Count(*) From SO32D0A")(0) > 0 Then
        If optSort(0).Value Then
            strGroupName = "SortA={SO32D0A.CustId};SortB={SO32D0A.BillNo};Type=3"
        Else
            strGroupName = "SortA={SO32D0A.Anomaly};SortB={SO32D0A.BillNo};Type=3"
        End If
        If blnExcel Then
            Call toExcel(3)
            Call UseProperty(rsExcel, "出單檢核表", "第一頁")
            blnExcel = False
        Else
            PrintRpt2 GetPrinterName(5), RptName("SO32D0", 2), "SO32D0", "出單檢核表 [SO32D0A]", , GetQryStr, , True, "Tmp0000.MDB", , strGroupName
        End If
    Else
        SendSQL , , True
        MsgNoRcd
    End If
    On Error Resume Next
    Picture1.Visible = False
    Me.Enabled = True
    Close3Recordset rsExcel
    Exit Sub
chkErr:
    Picture1.Visible = False
    Me.Enabled = True
    cnn.RollbackTrans
  Call ErrSub(Me.Name, "RunStatus3")
End Sub
'#4258 增加日額計算 By Kin 2008/12/22
Private Sub RunDayAmt()
  On Error GoTo chkErr
    Dim strSQL As String
    Dim strTablName As String
    Dim intPara3 As Integer
    Dim rsTmp As New ADODB.Recordset
    ReadyGoPrint
    Me.Enabled = False
    Call subChoose
    If InStr(1, cboStatus.Text, "過入前") > 0 Then
        strTablName = "SO032"
    Else
        strTablName = "SO033"
    End If
    '0=>1 1=>0
    '#4258 跟Penny與Jacy討論後,以後取到的第一筆為準 By Kin 2008/12/22
    If Not GetRS(rsTmp, "Select nvl(Para16,0) as Para16,nvl(Para17,0) as Para17,nvl(Para3,0) as Para3  From " & GetOwner & "SO043 Where CompCode=" & gCompCode, gcnGi, adUseClient) Then Exit Sub
       intPara3 = IIf(Val(rsTmp("Para3")) = 0, 1, 0)
       strSQL = "Select " & strTablName & ".CUSTID," & strTablName & ".CITEMCODE," & strTablName & ".CITEMNAME," & _
                strTablName & ".SHOULDAMT," & strTablName & ".SHOULDDATE," & strTablName & ".FACISEQNO," & _
                strTablName & ".FACISNO," & strTablName & ".BPCODE," & strTablName & ".BPNAME," & _
                "SO003A.Period as Period,SO003A.DiscountAmt,SO003A.DayAmt, " & _
                IIf(rsTmp("Para16") = 1, " Round(" & strTablName & ".ShouldAmt/(sf_d360(" & strTablName & ".RealStartDate," & strTablName & ".RealStopDate+" & intPara3 & "," & rsTmp("Para17") & "," & gCompCode & "," & strTablName & ".ServiceType)),3) OutDayAmt", _
                " Round(" & strTablName & ".ShouldAmt/(" & strTablName & ".RealStopDate-" & strTablName & ".RealStartDate )+1,3) OutDayAmt ") & _
                " From " & GetOwner & strTablName & "," & GetOwner & "SO003A," & GetOwner & "CD019" & _
                " Where " & strTablName & ".CustId=SO003A.Custid" & _
                " And " & strTablName & ".CitemCode=SO003A.CitemCode" & _
                " And " & strTablName & ".ServiceType=SO003A.ServiceType" & _
                " And " & strTablName & ".FaciSeqNo=SO003A.FaciSeqNo " & _
                " And " & strTablName & ".CompCode=SO003A.CompCode" & _
                " And " & strTablName & ".BpCode=SO003A.BpCode" & _
                " And " & strTablName & ".DiscountDate1=SO003A.DiscountDate1" & _
                " And " & strTablName & ".BpCode is not Null" & _
                " And " & strTablName & ".CitemCode=CD019.CodeNo" & _
                " And CD019.PeriodFlag=1" & _
                " And " & strTablName & ".RealStartDate is Not Null" & _
                " And " & strTablName & ".RealStopDate is Not Null" & _
                IIf(rsTmp("Para16") = 1, " And Round(" & strTablName & ".ShouldAmt/(sf_d360(" & strTablName & ".RealStartDate," & strTablName & ".RealStopDate+" & intPara3 & "," & rsTmp("Para17") & "," & gCompCode & "," & strTablName & ".ServiceType)),3)<>SO003A.DayAmt", _
                    " And Round(" & strTablName & ".ShouldAmt/(" & strTablName & ".RealStopDate-" & strTablName & ".RealStartDate )+1,3)<>SO003A.DayAmt") & _
                IIf(strChoose <> "", " And " & strChoose, "")
    If Not GetRS(rsQry, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If rsQry.RecordCount <= 0 Then
        Me.Enabled = True
        SendSQL , , True
        MsgNoRcd
        Call Close3Recordset(rsTmp)
        Call Close3Recordset(rsExcel)
        Exit Sub
    End If
    Call CreateTable
    cnn.BeginTrans
        SendSQL rsQry.Source, True
        If rsQry.RecordCount > 0 Then
            Do While Not rsQry.EOF
                '#4258 測試後調整規格,其它欄位要改成出單日額與多階日額 By Kin 2009/02/02
                cnn.Execute "Insert Into SO32D0A (Anomaly,CustId,CitemCode,CitemName,ShouldAmt,ShouldDate,FaciSeqNo,FaciSNO,Period,Other,BpCode,BpName) Values " & _
                        "('收費資料與週期多階的平均單價不符'," & rsQry("CustId") & "," & rsQry("CitemCode") & ",'" & rsQry("CitemName") & "'," & _
                        rsQry("ShouldAmt") & ",'" & rsQry("ShouldDate") & "'," & _
                        "'" & rsQry("FaciSeqNo") & "','" & rsQry("FaciSNo") & "" & "'," & rsQry("Period") & "," & _
                        "'出單日額=" & rsQry("OutDayAmt") & Chr(13) & "多階日額=" & rsQry("DayAmt") & "'," & _
                        IIf(IsNull(rsQry("BpCode")), "NULL,", "'" & rsQry("BpCode") & "',") & IIf(IsNull(rsQry("BpName")), "NULL", "'" & rsQry("BpName") & "'") & ")"
                pgbBar.Value = (rsQry.AbsolutePosition / rsQry.RecordCount) * 100
                rsQry.MoveNext
                DoEvents
            Loop
        End If
    cnn.CommitTrans
    If blnExcel Then
        Call toExcel(3)
        Call UseProperty(rsExcel, "出單檢核表", "第一頁")
        blnExcel = False
    Else
        PrintRpt2 GetPrinterName(5), RptName("SO32D0", 3), "SO32D0", "出單檢核表 [SO32D0A]", , GetQryStr, , True, "Tmp0000.MDB", , strGroupName
    End If
    On Error Resume Next
    Me.Enabled = True
    SendSQL , , True
    Call Close3Recordset(rsTmp)
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "RunDayAmt")
End Sub
Private Sub RunStatus2()
  On Error GoTo chkErr
    ReadyGoPrint
    Me.Enabled = False
    Call CreateTable
    cnn.BeginTrans
        If chkAfter(0).Value Then
            Call AfterStatus01
            SendSQL rsQry.Source, True
            If rsQry.RecordCount > 0 Then
                Do While Not rsQry.EOF
                    cnn.Execute "Insert Into SO32D0A (Anomaly,CustId,BillNo,ShouldAmt) Values " & _
                                "('單據合計總金額<=0'," & rsQry("CustId") & ",'" & rsQry("BillNo") & "'," & rsQry("Total") & ")"
                    pgbBar.Value = ((rsQry.AbsolutePosition / rsQry.RecordCount) * (100 / (chkAfter.UBound + 1)))
                    rsQry.MoveNext
                    DoEvents
                Loop
            End If
        End If
        pgbBar.Value = (100 / (chkAfter.UBound + 1)) * 1
        If chkAfter(1).Value Then
            Call AfterStatus02
            SendSQL rsQry.Source, True
            If rsQry.RecordCount > 0 Then
                Do While Not rsQry.EOF
                    cnn.Execute "Insert Into SO32D0A (Anomaly,CustId,BillNo,ShouldDate,AccountNo) Values " & _
                                "('同一客戶有多筆資料,應收日期同,但帳號不同'," & rsQry("CustId") & ",'" & rsQry("BillNo") & _
                                "','" & rsQry("ShouldDate") & "','" & rsQry("AccountNo") & "')"
                    pgbBar.Value = ((rsQry.AbsolutePosition / rsQry.RecordCount) * (100 / (chkAfter.UBound + 1))) + ((100 / (chkAfter.UBound + 1)) * 1)
                    rsQry.MoveNext
                    DoEvents
                Loop
            End If
        End If
        pgbBar.Value = (100 / (chkAfter.UBound + 1)) * 2
    cnn.CommitTrans
    Picture1.Visible = False
    If cnn.Execute("Select Count(*) From SO32D0A")(0) > 0 Then
        If optSort(0).Value Then
            strGroupName = "SortA={SO32D0A.CustId};SortB={SO32D0A.BillNo};Type=2"
        Else
            strGroupName = "SortA={SO32D0A.Anomaly};SortB={SO32D0A.BillNo};Type=2"
        End If
        If blnExcel Then
            Call toExcel(2)
            Call UseProperty(rsExcel, "出單檢核表", "第一頁")
            blnExcel = False
        Else
            PrintRpt2 GetPrinterName(5), RptName("SO32D0", 2), "SO32D0", "出單檢核表 [SO32D0A]", , GetQryStr, , True, "Tmp0000.MDB", , strGroupName
        End If
    Else
        SendSQL , , True
        MsgNoRcd
    End If
    On Error Resume Next
    Picture1.Visible = False
    Me.Enabled = True
    Close3Recordset rsExcel
    Me.Enabled = True
    Exit Sub
chkErr:
  Picture1.Visible = False
  Me.Enabled = True
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "RunStatus2")
End Sub
Private Function IsDataOk() As Boolean
  On Error GoTo chkErr
    Dim i As Long
    Dim blnYes As Boolean
    blnYes = False
    Select Case cboStatus.ListIndex
        Case 0
            If chkBefore(1).Value Then
                If ginPeriod.Value = 0 Or ginPeriod.Text = "" Then
                    MsgBox "請輸入期數！", vbExclamation, "訊息"
                    ginPeriod.SetFocus
                    Exit Function
                End If
            End If
            If chkBefore(2).Value Then
                If txtCmcode.Text = "" Then
                    MsgBox "請輸入正常收費方式！", vbExclamation, "訊息"
                    txtCmcode.SetFocus
                    Exit Function
                End If
            End If
            For i = chkBefore.LBound To chkBefore.UBound
                If chkBefore(i).Value = 1 Then blnYes = True: Exit For
            Next i
        '#4258 如果檢查日額其它的不能選 By Kin 2008/12/22
        Case 1
            If chkDayAmt.Value = 1 Then blnYes = True
            For i = chkAfter.LBound To chkAfter.UBound
                If chkAfter(i).Value = 1 Then blnYes = True
            Next i
        Case 2
            If chkDayAmt.Value = 1 Then blnYes = True
            If chkFinal.Value = 1 Then blnYes = True
    End Select
    If Not blnYes Then MsgBox "請選擇要檢查的項目", vbExclamation, "訊息": Exit Function
    IsDataOk = True
    Exit Function
chkErr:
  ErrSub Me.Name, "IsDataOK"
End Function
Private Sub RunStatus1()
  On Error GoTo chkErr
    Dim strErrorStatus As String
    Dim strOtherStatus As String
    Dim strSO003A As String
    Dim strChkSO002A As String
    Dim strBPCode As String
    Dim strBpName As String
    Me.Enabled = False
    ReadyGoPrint
    CreateTable
    cnn.BeginTrans
    '#4137 匯出Excel要增加SO003A.BpCode、SO003A.BpName欄位 By Kin 2008/10/09
    If chkBefore(0).Value Then
        Call BeforeStatus01
        SendSQL rsQry.Source, True
        If rsQry.RecordCount > 0 Then
            Do While Not rsQry.EOF
                cnn.Execute "Insert Into SO32D0A (Anomaly,CustId,CitemCode,CitemName,FaciSeqNo,FaciSNo,ClctDate,Period,Amount,Other,BpCode,BpName) Values " & _
                            "('週期多階平均單價與牌價不符'," & rsQry("CustId") & "," & rsQry("CitemCode") & ",'" & rsQry("CitemName") & "'," & _
                            "'" & rsQry("FaciSeqNo") & "','" & rsQry("FaciSNo") & "" & "','" & rsQry("ClctDate") & "'," & rsQry("Period") & "," & _
                            rsQry("DiscountAmt") & ",'多階優惠起迄範圍 " & rsQry("DiscountDate1") & "~" & rsQry("DiscountDate2") & "'," & _
                            IIf(IsNull(rsQry("BpCode")), "NULL,", "'" & rsQry("BpCode") & "',") & IIf(IsNull(rsQry("BpName")), "NULL", "'" & rsQry("BpName") & "'") & ")"
                pgbBar.Value = (rsQry.AbsolutePosition / rsQry.RecordCount) * (100 / (chkBefore.UBound + 1)) * 1
                rsQry.MoveNext
                DoEvents
            Loop
        End If
    End If
    pgbBar.Value = (100 / (chkBefore.UBound + 1)) * 1
    If chkBefore(1).Value Then
        Call BeforeStatus02
        SendSQL rsQry.Source, True
        If rsQry.RecordCount > 0 Then
            Do While Not rsQry.EOF
                cnn.Execute "Insert Into SO32D0A (Anomaly,CustId,CitemCode,CitemName,FaciSeqNo,FaciSNO,ClctDate,Period,Amount,BpCode,BpName) Values " & _
                            "('週期異常期數'," & rsQry("CustId") & "," & rsQry("CitemCode") & ",'" & rsQry("CitemName") & "'," & _
                            "'" & rsQry("FaciSeqNo") & "','" & rsQry("FaciSNo") & "" & "','" & rsQry("ClctDate") & "'," & rsQry("Period") & "," & _
                            rsQry("Amount") & "," & _
                            IIf(IsNull(rsQry("BpCode")), "NULL,", "'" & rsQry("BpCode") & "',") & IIf(IsNull(rsQry("BpName")), "NULL", "'" & rsQry("BpName") & "'") & ")"
                pgbBar.Value = ((rsQry.AbsolutePosition / rsQry.RecordCount) * (100 / (chkBefore.UBound + 1))) + ((100 / (chkBefore.UBound + 1)) * 1)
                rsQry.MoveNext
                DoEvents
            Loop
        End If
    End If
    pgbBar.Value = (100 / (chkBefore.UBound + 1)) * 2
    If chkBefore(2).Value Then
        Call BeforeStatus03
        SendSQL rsQry.Source, True
        If rsQry.RecordCount > 0 Then
            Do While Not rsQry.EOF
                cnn.Execute "Insert Into SO32D0A (Anomaly,CustId,CitemCode,CitemName,FaciSeqNo,FaciSNO,ClctDate,Period,Amount,Other,BpCode,BpName) Values " & _
                            "('週期帳戶別為非虛擬帳號類的異常收費方式'," & rsQry("CustId") & "," & rsQry("CitemCode") & ",'" & rsQry("CitemName") & "'," & _
                            "'" & rsQry("FaciSeqNo") & "','" & rsQry("FaciSNo") & "" & "','" & rsQry("ClctDate") & "'," & rsQry("Period") & "," & _
                            rsQry("Amount") & ",'收費方式=" & rsQry("CMCode") & "." & rsQry("CMName") & "'," & _
                            IIf(IsNull(rsQry("BpCode")), "NULL,", "'" & rsQry("BpCode") & "',") & IIf(IsNull(rsQry("BpName")), "NULL", "'" & rsQry("BpName") & "'") & ")"
                            
                pgbBar.Value = ((rsQry.AbsolutePosition / rsQry.RecordCount) * (100 / (chkBefore.UBound + 1))) + ((100 / (chkBefore.UBound + 1)) * 2)
                rsQry.MoveNext
                DoEvents
            Loop
        End If
    End If
    pgbBar.Value = (100 / (chkBefore.UBound + 1)) * 3
    If chkBefore(3).Value Then
        Call BeforeStatus04
        SendSQL rsQry.Source, True
        If rsQry.RecordCount > 0 Then
            Do While Not rsQry.EOF
                strErrorStatus = Empty
                Select Case rsQry("Error")
                    Case 1
                        strErrorStatus = "'週期相同設備正/負項收費資料不符(收費方式)'"
                        strOtherStatus = "'收費方式=" & rsQry("CMCODE") & "." & rsQry("CMNAME") & "'"
                    Case 2
                        strErrorStatus = "'週期相同設備正/負項收費資料不符(期數)'"
                        strOtherStatus = "NULL"
                    Case 3
                        strErrorStatus = "'週期相同設備正/負項收費資料不符(下收日)'"
                        strOtherStatus = "NULL"
                    Case 4
                        strErrorStatus = "'週期相同設備正/負項收費資料不符(速率)'"
                        strOtherStatus = "'速率=" & rsQry("CMBaudRate") & "'"
                End Select
                cnn.Execute "Insert Into SO32D0A (Anomaly,CustId,CitemCode,CitemName,FaciSeqNo,FaciSNO,ClctDate,Period,Amount,Other,BpCode,BpName) Values " & _
                            "(" & strErrorStatus & "," & rsQry("CustId") & "," & rsQry("CitemCode") & ",'" & rsQry("CitemName") & "'," & _
                            "'" & rsQry("FaciSeqNo") & "','" & rsQry("FaciSNo") & "" & "','" & rsQry("ClctDate") & "'," & rsQry("Period") & "," & _
                            rsQry("Amount") & "," & strOtherStatus & "," & _
                            IIf(IsNull(rsQry("BpCode")), "NULL,", "'" & rsQry("BpCode") & "',") & IIf(IsNull(rsQry("BpName")), "NULL", "'" & rsQry("BpName") & "'") & ")"
                pgbBar.Value = ((rsQry.AbsolutePosition / rsQry.RecordCount) * (100 / (chkBefore.UBound + 1))) + ((100 / (chkBefore.UBound + 1)) * 3)
                rsQry.MoveNext
                DoEvents
            Loop
        End If
    End If
    pgbBar.Value = (100 / (chkBefore.UBound + 1)) * 4
    If chkBefore(4).Value Then
        Call BeforeStatus05
        SendSQL rsQry.Source, True
        If rsQry.RecordCount > 0 Then
            Do While Not rsQry.EOF
                cnn.Execute "Insert Into SO32D0A (Anomaly,CustId,CitemCode,CitemName,FaciSeqNo,FaciSNO,ClctDate,Period,Amount,Other,BpCode,BpName) Values " & _
                            "('週期收費項目與設備的速率不符'," & rsQry("CustId") & "," & rsQry("CitemCode") & ",'" & rsQry("CitemName") & "'," & _
                            "'" & rsQry("FaciSeqNo") & "','" & rsQry("FaciSNo") & "" & "','" & rsQry("ClctDate") & "'," & rsQry("Period") & "," & _
                            rsQry("Amount") & ",'週期收費速率=" & rsQry("CMName") & Chr(13) & "設備速率=" & rsQry("CMBaudRate") & "'," & _
                            IIf(IsNull(rsQry("BpCode")), "NULL,", "'" & rsQry("BpCode") & "',") & IIf(IsNull(rsQry("BpName")), "NULL", "'" & rsQry("BpName") & "'") & ")"
                pgbBar.Value = ((rsQry.AbsolutePosition / rsQry.RecordCount) * (100 / (chkBefore.UBound + 1))) + ((100 / (chkBefore.UBound + 1)) * 4)
                rsQry.MoveNext
                DoEvents
            Loop
        End If
    End If
    pgbBar.Value = (100 / (chkBefore.UBound + 1)) * 5
    If chkBefore(5).Value Then
        Call BeforeStatus06
        SendSQL rsQry.Source, True
        If rsQry.RecordCount > 0 Then
            Do While Not rsQry.EOF
                cnn.Execute "Insert Into SO32D0A (Anomaly,CustId,CitemCode,CitemName,FaciSeqNo,FaciSNO,ClctDate,Period,Amount,BpCode,BpName) Values " & _
                                "('週期的設備序號與設備的不符'," & rsQry("CustId") & "," & rsQry("CitemCode") & ",'" & rsQry("CitemName") & "'," & _
                                "'" & rsQry("FaciSeqNo") & "','" & rsQry("FaciSNo") & "" & "','" & rsQry("ClctDate") & "'," & rsQry("Period") & "," & _
                                rsQry("Amount") & "," & _
                                IIf(IsNull(rsQry("BpCode")), "NULL,", "'" & rsQry("BpCode") & "',") & IIf(IsNull(rsQry("BpName")), "NULL", "'" & rsQry("BpName") & "'") & ")"
                pgbBar.Value = ((rsQry.AbsolutePosition / rsQry.RecordCount) * (100 / (chkBefore.UBound + 1))) + ((100 / (chkBefore.UBound + 1)) * 5)
                rsQry.MoveNext
                DoEvents
            Loop
        End If
    End If
    pgbBar.Value = (100 / (chkBefore.UBound + 1)) * 6
    cnn.CommitTrans
    Picture1.Visible = False
    If cnn.Execute("Select Count(*) From SO32D0A")(0) > 0 Then
        If optSort(0).Value Then
            strGroupName = "SortA={SO32D0A.CustId};SortB={SO32D0A.FaciSNo};SortC={SO32D0A.CitemCode};Type=1"
        Else
            strGroupName = "SortA={SO32D0A.Anomaly};SortB={SO32D0A.CustId};SortC={SO32D0A.FaciSNo};SortD={SO32D0A.CitemCode};Type=1"
        End If
        If blnExcel Then
            Call toExcel(1)
            '#7074 To Add a choice-condition parameter into excel print function By Kin 2015/09/09
            Call UseProperty(rsExcel, "出單檢核表", "第一頁", _
                        , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , GetQryStr)
            blnExcel = False
        Else
            PrintRpt2 GetPrinterName(5), RptName("SO32D0"), "SO32D0", "出單檢核表 [SO32D0A]", , GetQryStr, , True, "Tmp0000.MDB", , strGroupName
        End If
    Else
        SendSQL , , True
        MsgNoRcd
    End If
    On Error Resume Next
    Picture1.Visible = False
    Me.Enabled = True
    Close3Recordset rsExcel
    Exit Sub
chkErr:
  Me.Enabled = True
  Picture1.Visible = False
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "RunStatus1")
End Sub
Private Sub toExcel(ByVal intType As Integer)
  On Error GoTo chkErr
    Dim strQry  As String
    strQry = "Select * From SO3440A"
'    RptToTxt "SO32D0A.rpt", , strSQL, GetQryStr, , "Tmp0000.MDB", strGroupName, , , Environ("Temp") & "\SO32D0"
'    If Not Get_RS_From_Txt(Environ("Temp"), "SO32D0.txt", rsExcel) Then blnExcel = False: Exit Sub

    If intType = 1 Then
        RptToTxt "SO32D0A.rpt", , strSQL, GetQryStr, , "Tmp0000.MDB", strGroupName, , , gTempPath & "\SO32D0"
        If Not Get_RS_From_Txt(gTempPath, "SO32D0.txt", rsExcel) Then blnExcel = False: Exit Sub
    ElseIf intType = 2 Then
        RptToTxt "SO32D0A2.rpt", , strSQL, GetQryStr, , "Tmp0000.MDB", strGroupName, , , gTempPath & "\SO32D0"
        If Not Get_RS_From_Txt(gTempPath, "SO32D0.txt", rsExcel) Then blnExcel = False: Exit Sub
    Else
        RptToTxt "SO32D0A3.rpt", , strSQL, GetQryStr, , "Tmp0000.MDB", strGroupName, , , gTempPath & "\SO32D0"
        If Not Get_RS_From_Txt(gTempPath, "SO32D0.txt", rsExcel) Then blnExcel = False: Exit Sub
    End If
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "ToExcel")
End Sub
Private Sub DefaultValue()
  On Error GoTo chkErr
    
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
    cboPeriod.ListIndex = 0
    cboStatus.ListIndex = 0
    cboCmitem.ListIndex = 2
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "DefaultValue")
End Sub
Private Sub JoinCondition()
  On Error GoTo chkErr
    strStatusTable = Empty
    strStatusWhere = Empty
    strSO003AWhere = Empty
    
    strSO003AWhere = " And SO003.CustId=SO003A.CustId(+) And SO003.CompCode=SO003A.CompCode(+)" & _
            " And SO003.CitemCode=SO003A.CitemCode(+) And SO003.ServiceType=SO003A.ServiceType(+)" & _
            " And SO003.FaciSeqNo=SO003A.FaciSeqNo(+) And SO003.ContNo=SO003A.ContNo(+)" & _
            " And SO003.ClctDate>=SO003A.DisCountDate1(+) And SO003.ClctDate<=SO003A.DisCountDate2(+)" & _
            " And 0=SO003A.StopFlag(+)"
            
    Select Case cboStatus.ListIndex
        '#4258 增加判斷SO003.下收日>MAX(非停用的多階優惠截止日)SO003.Discountdate2且週期的"優惠到期計價依據"=3.到期不出 By Kin 2008/12/22
        Case 0
            strStatusTable = GetOwner & "SO003," & GetOwner & "SO002," & GetOwner & "SO001," & GetOwner & "SO014"
            strStatusWhere = " SO003.CompCode=" & gilCompCode.GetCodeNo & " And SO002.CustStatusCode In(1,2)" & _
                              " And SO002.CustId=SO003.CustId And SO002.ServiceType=SO003.ServiceType" & _
                              " And SO002.CompCode=SO003.CompCode And SO001.CustId=SO003.CustId" & _
                              " And SO001.CompCode=SO003.CompCode And SO014.CompCode=SO001.CompCode " & _
                              " And SO003.ClctDate<SO003.Discountdate2 And Nvl(SO003.expiretype,0)<>3"
                              
                              
                              
                              
                              
        Case 1
            strStatusTable = GetOwner & "CD019," & GetOwner & "SO032"
            strStatusWhere = "SO032.CitemCode=CD019.CodeNo And SO032.CompCode=" & gilCompCode.GetCodeNo
        Case 2
            strStatusTable = GetOwner & "CD019," & GetOwner & "SO033"
            strStatusWhere = "SO033.UCCode is Not Null And SO033.CitemCode=CD019.CodeNo" & _
                             " And SO033.CompCode=" & gilCompCode.GetCodeNo & " And CancelFlag=0"
    End Select
chkErr:
  Call ErrSub(Me.Name, "JoinCondition")
End Sub
Private Sub CtlEnable()
  On Error GoTo chkErr
    Dim ctl As Control
    For Each ctl In Me.Controls()
        Select Case UCase(TypeName(ctl))
            Case "LABEL"
                ctl.Enabled = False
            Case "GIMULTI"
                ctl.Enabled = True
                ctl.Clear
            Case "CSMULTI"
                ctl.Enabled = True
                ctl.Clear
            Case "COMMANDBUTTON"
                ctl.Enabled = True
            Case "VSCROLLBAR"
                ctl.Enabled = True
            Case Else
                ctl.Enabled = False
        End Select
    Next
    fraWhere.Enabled = True
    pic2.Enabled = True
    gimCreateEn.Enabled = False
    gdaClctDate1.SetValue ""
    gdaClctDate2.SetValue ""
    gdaShouldDate1.SetValue ""
    gdaShouldDate2.SetValue ""
    gdaCreateTime1.SetValue ""
    gdaCreateTime2.SetValue ""
    fra1.Enabled = False
    fra2.Enabled = False
    fra3.Enabled = False
    fraOrder.Enabled = True
    optSort(0).Enabled = True
    optSort(1).Enabled = True
    gilCompCode.Enabled = True
    cboStatus.Enabled = True
    txtCustId.Enabled = True
    lblCustId.Enabled = True
    lblCompCode.Enabled = True
    lblStatus.Enabled = True
    lblPmt.Enabled = True
    For Each ctl In Me.Controls()
        Select Case UCase(TypeName(ctl))
            Case "CHECKBOX"
                ctl.Value = 0
            Case "TEXTBOX"
                ctl.Text = ""
            Case "GINUMBER"
                ctl.Text = ""
        End Select
    Next
    Select Case cboStatus.ListIndex
        Case 0
            fraWhere.Enabled = True
            gimBillType.Enabled = False
            cboPeriod.Enabled = True
            lblPeriod.Enabled = True
            cboCmitem.Enabled = True
            lblCmitem.Enabled = True
            gdaClctDate1.Enabled = True
            gdaClctDate2.Enabled = True
            lblClctDate.Enabled = True
            fra1.Enabled = True
            fraAddrType.Enabled = True
            For Each ctl In Me.Controls
                If UCase(ctl.Container.Name) = "FRA1" Then
                    ctl.Enabled = True
                ElseIf UCase(ctl.Container.Name) = "FRAADDRTYPE" Then
                    ctl.Enabled = True
                End If
            Next
            ginPeriod.Enabled = False
            txtCmcode.Enabled = False
            chkCustAllot.Enabled = False
        Case 1
            '#4258 增加檢查日額計算 By Kin 2008/12/22
            fra4.Enabled = True
            chkDayAmt.Enabled = True
            cboCmitem.Enabled = True
            lblCmitem.Enabled = True
            fra2.Enabled = True
            
            For Each ctl In Me.Controls
                If UCase(TypeName(ctl)) = "CSMULTI" Then
                    ctl.Enabled = False
                ElseIf UCase(TypeName(ctl)) = "GIMULTI" Then
                    ctl.Enabled = False
                End If
                If UCase(ctl.Container.Name) = "FRA2" Then
                    ctl.Enabled = True
                End If
            Next
            gimServiceType.Enabled = True
            gimCitemCode.Enabled = True
        Case 2
            '#4258 增加檢查日額計算 By Kin 2008/12/22
            fra4.Enabled = True
            chkDayAmt.Enabled = True
            cboCmitem.Enabled = True
            lblCmitem.Enabled = True
            gdaShouldDate1.Enabled = True
            gdaShouldDate2.Enabled = True
            lblShouldDate.Enabled = True
            fra3.Enabled = True
            fraWhere.Enabled = True
            For Each ctl In Me.Controls()
                If UCase(TypeName(ctl)) = "CSMULTI" Then
                    ctl.Enabled = False
                ElseIf UCase(TypeName(ctl)) = "GIMULTI" Then
                    ctl.Enabled = False
                End If
                If UCase(ctl.Container.Name) = "FRA3" Then
                    ctl.Enabled = True
                End If
            Next
            gimServiceType.Enabled = True
            gimBillType.Enabled = True
            gimCitemCode.Enabled = True
            gimCMCode.Enabled = True
            gimCreateEn.Enabled = True
            gdaCreateTime1.Enabled = True
            gdaCreateTime2.Enabled = True
            lblCreateTime.Enabled = True
'            txtCreateEn.Enabled = True
        Case 3
'            fraFormat.Enabled = True
'            fra4.Enabled = True
'            optNew.Enabled = True
'            optOld.Enabled = True
'            For Each ctl In Me.Controls
'                If UCase(ctl.Container.Name) = "FRA4" Then
'                    ctl.Enabled = True
'                End If
'            Next
    End Select

    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "CtlEnable")
End Sub
Private Sub subGim()
    On Error GoTo chkErr
    
    SetgiMulti gimServiceType, "CodeNo", "Description", "CD046", "代碼", "名稱"
    '#4156 測試不OK,RefNo is Null 也要出來 By Kin 2008/10/29
    SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱", " Where PeriodFlag = 1 And RefNo<>18 Or RefNo is Null"
    SetgiMulti gimClassCode, "CodeNo", "Description", "CD004", "代碼", "名稱", , True
    SetgiMulti gimAreaCode, "CodeNo", "Description", "CD001", "代碼", "名稱", , True
    SetgiMulti gimServCode, "CodeNo", "Description", "CD002", "代碼", "名稱", , True
    SetgiMulti gimClctArea, "CodeNo", "Description", "CD040", "代碼", "名稱", , True
    SetgiMulti gimMduId, "Mduid", "Name", "SO017", "大樓編號", "大樓名稱"
    SetgiMulti gimCreateEn, "EmpNo", "EmpName", "CM003", "過入人員代碼", "過入人員名稱"
    SetgiMulti gimStrtCode, "CodeNo", "Description", "CD017", "代碼", "名稱", , True
    SetgiMulti gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱", , True
    Call SetgiMultiAddItem(gimBillType, "B,T,I,M,P", "收費單,臨時收費單,裝機收費單,維修收費單,停拆移機收費單", "代碼", "名稱")

Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub
Private Sub subGil()
  On Error GoTo chkErr
    SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 1, 12, GetCompCodeFilter(gilCompCode)
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub cboStatus_Click()
  On Error GoTo chkErr
    Call CtlEnable
    Call subGim
    JoinCondition
    '#5460 把預設值拿掉 By Kin 2010/08/23
'    If cboStatus.ListIndex = 2 Then
'        gimBillType.SetQueryCode ("B,T")
'    End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cboStatus_Click")
End Sub

Private Sub chkAfter_Click(Index As Integer)
  On Error Resume Next
    If chkAfter(Index).Value = 1 Then chkDayAmt.Value = 0
End Sub

Private Sub chkAll1_Click()
  On Error GoTo chkErr
    Dim ctl As Control
    If chkAll1.Value Then
        For Each ctl In Me.Controls()
            If UCase(ctl.Container.Name) = "FRA1" Then
                If UCase(TypeName(ctl)) = "CHECKBOX" Then
                    ctl.Value = 1
                End If
            End If
        Next
        chkCancel1.Value = 0
    End If
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "chkAll1_Click")
End Sub


Private Sub chkBefore_Click(Index As Integer)
  On Error Resume Next
    Dim i As Long
    Dim blnAll As Boolean
    Dim blnCancel As Boolean
    If Index = 1 Then
        If chkBefore(Index).Value = 0 Then
            ginPeriod.Text = ""
            ginPeriod.Enabled = False
            chkCustAllot.Value = 0
            chkCustAllot.Enabled = False
        Else
            ginPeriod.Enabled = True
            chkCustAllot.Enabled = True
        End If
    End If
    If Index = 2 Then
        If chkBefore(Index).Value = 0 Then
            txtCmcode.Text = ""
            txtCmcode.Enabled = False
        Else
            txtCmcode.Enabled = True
        End If
    End If
    If chkBefore(Index).Value = 1 Then
        chkCancel1.Value = 0
    Else
        chkAll1.Value = 0
    End If
'    If chkBefore(Index).Value = 1 Then
'        blnAll = True
'    Else
'        blnCancel = True
'    End If
'    For i = chkBefore.LBound To chkBefore.UBound
'        If blnAll Then
'            If chkBefore(i).Value = 0 Then blnAll = False
'        Else
'            If chkBefore(i).Value = 1 Then blnCancel = False
'        End If
'    Next
'    If blnAll Then
'        If chkAll1.Value = 0 Then
'            chkAll1.Value = 1
'        End If
'    Else
'        If chkAll1.Value = 1 Then
'            chkAll1.Value = 0
'        End If
'        If blnCancel Then
'            If chkCancel1.Value = 0 Then
'                chkCancel1.Value = 1
'            End If
'        Else
'            If chkCancel1.Value = 1 Then
'                chkCancel1.Value = 0
'            End If
'        End If
'    End If
End Sub

Private Sub chkCancel1_Click()
  On Error GoTo chkErr
    Dim ctl As Control
    If chkCancel1.Value Then
        For Each ctl In Me.Controls()
            If UCase(ctl.Container.Name) = "FRA1" Then
                If UCase(TypeName(ctl)) = "CHECKBOX" Then
                    If UCase(ctl.Name) <> "CHKCANCEL1" Then
                        ctl.Value = 0
                    End If
                End If
            End If
        Next
        chkAll1.Value = 0
    End If
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub



Private Sub chkDayAmt_Click()
  '#4258 如果檢查日額有勾選其它的不能選擇 By Kin 2008/12/22
  On Error GoTo chkErr
    Dim i As Integer
    If chkDayAmt.Value Then
        For i = chkAfter.LBound To chkAfter.UBound
            chkAfter(i).Value = 0
'            chkAfter(i).Enabled = False
        Next
        chkFinal.Value = 0
        'chkFinal.Enabled = False
 '       fra2.Enabled = False
 '       fra3.Enabled = False
    Else
        Select Case cboStatus.ListIndex
            Case 1
                For i = chkAfter.LBound To chkAfter.UBound
                    chkAfter(i).Enabled = True
                Next
            Case 2
                chkFinal.Enabled = True
        End Select
        fra2.Enabled = True
        fra3.Enabled = True
    End If
  Exit Sub
chkErr:
  Call ErrSub(Me.Name, "chkDayAmt_Click")
End Sub

Private Sub chkFinal_Click()
  On Error Resume Next
  If chkFinal.Value = 1 Then chkDayAmt.Value = 0
End Sub

Private Sub cmdExit_Click()
  On Error Resume Next
  Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If IsDataOk Then
        GetQryStr
        Picture1.Visible = True
        Select Case cboStatus.ListIndex
            Case 0
                Call RunStatus1
            Case 1
                If chkDayAmt.Value Then
                    Call RunDayAmt
                Else
                    Call RunStatus2
                End If
            Case 2
                If chkDayAmt.Value Then
                    Call RunDayAmt
                Else
                    Call RunStatus3
                End If
        End Select
    End If
    Screen.MousePointer = vbDefault
    Picture1.Visible = False
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo chkErr
    Screen.MousePointer = vbHourglass
    If cboStatus.ListIndex = 0 Then
         Call PreviousRpt(GetPrinterName(5), RptName("SO32D0"), Me.Caption)
    Else
         Call PreviousRpt(GetPrinterName(5), RptName("SO32D0", 2), Me.Caption)
    End If
   
    Screen.MousePointer = vbDefault
  Exit Sub
chkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdToExcel_Click()
  On Error Resume Next
  blnExcel = True
  cmdOk.Value = True
End Sub

Private Sub Form_Activate()
  On Error Resume Next
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo chkErr
    If Picture1.Visible Then Exit Sub
    If KeyCode = vbKeyEscape Then Unload Me
    If KeyCode = vbKeyF5 Then If cmdOk.Enabled Then cmdOk.Value = True
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo chkErr
    Call subGim
    Call subGil
    Call DefaultValue
    Call CtlEnable
    Call CreateTable
    Call JoinCondition
    gdaCreateTime1.ShowSecond
    gdaCreateTime2.ShowSecond
    Exit Sub
chkErr:
  ErrSub Me.Name, "Form_Load"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  Call Close3Recordset(rsQry)
  Call Close3Recordset(rsExcel)
  cnn.Close
  Set cnn = Nothing
  
End Sub

Private Sub gdaClctDate2_Validate(Cancel As Boolean)
  On Error Resume Next
    If gdaClctDate1.GetValue = "" Or gdaClctDate2.GetValue = "" Then Exit Sub
       If Not IsDate(gdaClctDate2.Text) Then Exit Sub
       If DateDiff("d", gdaClctDate1.GetValue(True), gdaClctDate2.GetValue(True)) < 0 Then
          MsgBox "下收截止日不可小於下收起始日!", vbExclamation, "警告"
          gdaClctDate2.SetValue gdaClctDate1.GetValue
          Cancel = True
    End If

End Sub

Private Sub gdaCreateTime2_Validate(Cancel As Boolean)
  On Error Resume Next
    If gdaCreateTime1.GetValue = "" Or gdaCreateTime2.GetValue = "" Then Exit Sub
       If Not IsDate(gdaCreateTime2.Text) Then Exit Sub
       If DateDiff("s", gdaCreateTime1.GetValue(True), gdaCreateTime2.GetValue(True)) < 0 Then
          MsgBox "出單截止日不可小於出單起始日!", vbExclamation, "警告"
          gdaCreateTime2.SetValue gdaCreateTime1.GetValue
          Cancel = True
    End If



End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
  On Error Resume Next
    If gdaShouldDate1.GetValue = "" Or gdaShouldDate2.GetValue = "" Then Exit Sub
       If Not IsDate(gdaShouldDate2.Text) Then Exit Sub
       If DateDiff("d", gdaShouldDate1.GetValue(True), gdaShouldDate2.GetValue(True)) < 0 Then
          MsgBox "應收截止日不可小於應收起始日!", vbExclamation, "警告"
          gdaShouldDate2.SetValue gdaShouldDate1.GetValue
          Cancel = True
    End If

End Sub

Private Sub gilCompCode_Change()
  On Error GoTo chkErr
    If Not ChgComp(gilCompCode, "SO3400", "SO3440") Then Exit Sub
    Call subGim
    Call subGil
    Call DefaultValue
    Call CtlEnable
    Call JoinCondition
    Exit Sub
chkErr:
  ErrSub Me.Name, "gilCompCode_Change"
End Sub
Private Function GetQryStr() As String
    On Error GoTo chkErr
    Dim str1 As String
    Dim i As Long
    strChooseString = ""
    str1 = ""
    str1 = str1 & "出單狀態:" & cboStatus.Text & " ,"
    If cboStatus.ListIndex = 0 Then
        str1 = str1 & "週期性收費:" & cboPeriod.Text & " ,"
    End If
    str1 = str1 & "收費項目對象:" & cboCmitem.Text
    str1 = str1 & ";異常狀態地址依據:" & IIf(optChargeAddress.Value, "收費地址", "裝機地址")
    Select Case cboStatus.ListIndex
        Case 0
            For i = chkBefore.LBound To chkBefore.UBound
                If chkBefore(i).Value Then
                    If i = 1 Then
                        str1 = str1 & ";" & chkBefore(i).Caption & ":" & ginPeriod.Text
                    ElseIf i = 2 Then
                        str1 = str1 & ";" & chkBefore(i).Caption & " 正常收費方式:" & txtCmcode.Text
                    Else
                        str1 = str1 & ";" & chkBefore(i).Caption & ":是"
                    End If
                End If
            Next i
        Case 1
            For i = chkAfter.LBound To chkAfter.UBound
                If chkAfter(i).Value Then
                    str1 = str1 & ";" & chkAfter(i).Caption & ":是"
                End If
            Next i
        Case 2
            str1 = str1 & ";" & chkFinal.Caption & ":是"
    End Select
    
    If chkDayAmt.Value Then str1 = str1 & ";" & chkDayAmt.Caption & ":是"
        
    strChooseString = "公司別:" & subSpace(gilCompCode.GetDescription) & ";" & _
                          "服務類別:" & subSpace(gimServiceType.GetDispStr) & ";" & _
                          "收費方式:" & subSpace(gimCMCode.GetDispStr) & ";" & _
                          "收費項目:" & subSpace(gimCitemCode.GetDispStr) & ";" & _
                          "客戶類別:" & subSpace(gimClassCode.GetDispStr) & ";" & _
                          "服務區　:" & subSpace(gimServCode.GetDispStr) & ";"
    strChooseString = strChooseString & "行政區　:" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                          "收費區:" & subSpace(gimClctArea.GetDispStr) & ";" & _
                          "大樓名稱:" & subSpace(gimMduId.GetDispStr) & ";" & _
                          "街道範圍:" & subSpace(gimStrtCode.GetDispStr) & ";" & _
                          "客戶編號:" & subSpace(txtCustId) & _
                          "單據類別:" & subSpace(gimBillType.GetDispStr) & ";" & _
                          "應收日期:" & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & _
                          "下收日期:" & subSpace(gdaClctDate1.GetValue(True)) & "~" & subSpace(gdaClctDate2.GetValue(True)) & ";" & _
                          "客戶編號:" & subSpace(txtCustId) & ";" & _
                          "出單日期:" & subSpace(gdaCreateTime1.GetValue(True)) & "~" & subSpace(gdaCreateTime2.GetValue(True)) & ";" & _
                          "過入人員:" & subSpace(gimCreateEn.GetDispStr) & ";" & _
                          "排序方式: " & subSpace(IIf(optSort(0).Value, "客編", "異常狀態")) & ";" & str1 _

                              
    GetQryStr = strChooseString
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "Function GetQryStr")
End Function

Private Sub gimServiceType_Change()
  '#4137 收費項目要過濾服務別 By Kin 2008/10/13
  On Error GoTo chkErr
    Dim strFilter As String
    gimCitemCode.Clear
    '#4156 測試不OK,RefNo is Null也要出來 By Kin 2008/10/29
    If gimServiceType.GetDispStr <> "" And InStr(gimServiceType.GetDispStr, "全選") = 0 Then
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱", _
                " Where PeriodFlag = 1 And (ServiceType In(" & gimServiceType.GetQueryCode & ") Or ServiceType is Null) And (RefNo<>18 Or RefNo is Null)"
    Else
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱", " Where PeriodFlag = 1 And (RefNo<>18 Or RefNo is Null)"
    End If
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "gimServiceType_Change")
End Sub

Private Sub txtCmcode_KeyPress(keyAscii As Integer)
  On Error Resume Next
    
    If keyAscii >= 48 And keyAscii <= 57 Or keyAscii = 44 Or keyAscii = 8 Then
        If keyAscii = 44 Then
            If Asc(Right(" " & txtCmcode.Text, 1)) = 44 Or txtCmcode = "" Then keyAscii = 9
        End If
        If Not ChkMaxLengthOK(txtCmcode) Then
           If Not (keyAscii = 44 Or keyAscii = 8) Then keyAscii = 9
        End If
    Else
        keyAscii = 9
    End If
End Sub

Private Sub txtCustId_KeyPress(keyAscii As Integer)
  On Error Resume Next
    If keyAscii >= 48 And keyAscii <= 57 Or keyAscii = 44 Or keyAscii = 8 Or keyAscii = 45 Then
        If keyAscii = 44 Or keyAscii = 45 Then
            If Asc(Right(" " & txtCustId.Text, 1)) = 44 Or Asc(Right(" " & txtCustId.Text, 1)) = 45 Or txtCustId = "" Then keyAscii = 9
        End If
        If Not ChkMaxLengthOK(txtCustId) Then
           If Not (keyAscii = 44 Or keyAscii = 8 Or keyAscii = 45) Then keyAscii = 9
        End If
    Else
        keyAscii = 9
    End If
End Sub

Private Sub vsl1_Change()
    On Error Resume Next
        fraWhere.Top = -100 - fraWhere.Height * (vsl1.Value / 20) / 3
End Sub
