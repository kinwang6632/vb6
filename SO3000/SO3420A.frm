VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO3420A 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶待收結案並產生拆機單 [SO3420A]"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   4140
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3420A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdInExcel2 
      Caption         =   "外部匯入2"
      Height          =   375
      Left            =   8970
      TabIndex        =   90
      Top             =   3270
      Width           =   1065
   End
   Begin VB.CommandButton cmdInExcel 
      Caption         =   "外部匯入"
      Height          =   375
      Left            =   7950
      TabIndex        =   85
      Top             =   3270
      Width           =   975
   End
   Begin VB.TextBox txtSQL 
      Height          =   525
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   80
      Top             =   3180
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   10950
      TabIndex        =   30
      Top             =   3270
      Width           =   825
   End
   Begin VB.CheckBox chkOtherS 
      Caption         =   "檢查其他服務別是否有效戶"
      Height          =   195
      Left            =   8640
      TabIndex        =   75
      Top             =   150
      Width           =   2685
   End
   Begin VB.Frame fraBill 
      Caption         =   "收費單處理方式"
      Height          =   2745
      Left            =   7920
      TabIndex        =   66
      Top             =   480
      Width           =   3825
      Begin VB.OptionButton optRealStopDate 
         Caption         =   "收視截止結算至工單預約日期"
         Enabled         =   0   'False
         Height          =   195
         Left            =   90
         TabIndex        =   89
         Top             =   2400
         Width           =   3075
      End
      Begin VB.OptionButton optNothing 
         Caption         =   "不處理"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   300
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton optClsBill 
         Caption         =   "收費單一併結案"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   570
         Width           =   1665
      End
      Begin VB.OptionButton optUCCode 
         Caption         =   "未收原因"
         Height          =   195
         Left            =   90
         TabIndex        =   23
         Top             =   2070
         Width           =   1155
      End
      Begin VB.Frame fraBillClose 
         Enabled         =   0   'False
         Height          =   1395
         Left            =   150
         TabIndex        =   76
         Top             =   570
         Width           =   3585
         Begin VB.OptionButton optSTCode 
            Caption         =   "短收原因"
            Height          =   195
            Left            =   90
            TabIndex        =   18
            Top             =   330
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optCancel 
            Caption         =   "作廢原因"
            Height          =   195
            Left            =   90
            TabIndex        =   20
            Top             =   720
            Width           =   1095
         End
         Begin Gi_Date.GiDate gdaClsDate 
            Height          =   345
            Left            =   1200
            TabIndex        =   22
            Top             =   1020
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
            Enabled         =   0   'False
         End
         Begin prjGiList.GiList gilSTCode 
            Height          =   315
            Left            =   1200
            TabIndex        =   19
            Top             =   240
            Width           =   2325
            _ExtentX        =   4101
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
            FldWidth2       =   1400
            F2Corresponding =   -1  'True
         End
         Begin prjGiList.GiList gilCancelCode 
            Height          =   315
            Left            =   1200
            TabIndex        =   21
            Top             =   660
            Width           =   2325
            _ExtentX        =   4101
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
            FldWidth2       =   1400
            F2Corresponding =   -1  'True
         End
         Begin VB.Label lblClsDate 
            AutoSize        =   -1  'True
            Caption         =   "結帳日期"
            Height          =   195
            Left            =   330
            TabIndex        =   77
            Top             =   1095
            Width           =   780
         End
      End
      Begin prjGiList.GiList gilUCCode 
         Height          =   315
         Left            =   1350
         TabIndex        =   24
         Top             =   2010
         Width           =   2325
         _ExtentX        =   4101
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
         FldWidth1       =   600
         FldWidth2       =   1400
         F2Corresponding =   -1  'True
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2. 確定"
      Height          =   375
      Left            =   10020
      TabIndex        =   29
      Top             =   3270
      Width           =   885
   End
   Begin VB.Frame fraPR 
      Caption         =   "停機/拆機單處理方式"
      Height          =   2955
      Left            =   120
      TabIndex        =   51
      Top             =   510
      Width           =   7695
      Begin VB.CheckBox chkSetFaci 
         Caption         =   "指定現有設備"
         Height          =   195
         Left            =   150
         TabIndex        =   84
         Top             =   2280
         Width           =   1485
      End
      Begin VB.CheckBox chkPRFaci 
         Caption         =   "檢核待拆設備"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   2610
         Width           =   1485
      End
      Begin Gi_Date.GiDate gdaEBoxPRTime 
         Height          =   315
         Left            =   5400
         TabIndex        =   14
         Top             =   1140
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
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
      Begin VB.CheckBox chkPREBox 
         Caption         =   "&C.申裝STB 客戶一併設定關機"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3870
         TabIndex        =   13
         Top             =   870
         Width           =   3405
      End
      Begin VB.CheckBox chkDelIntro 
         Caption         =   "&B.一併刪除介紹人及介紹媒體欄位"
         Height          =   195
         Left            =   3870
         TabIndex        =   12
         Top             =   570
         Width           =   3255
      End
      Begin VB.CheckBox chkClsWork 
         Caption         =   "&A.派工單一併完工結案"
         Height          =   195
         Left            =   3870
         TabIndex        =   11
         Top             =   270
         Width           =   2235
      End
      Begin prjGiList.GiList gilWorkerEn3 
         Height          =   315
         Left            =   1500
         TabIndex        =   8
         Top             =   2790
         Visible         =   0   'False
         Width           =   2325
         _ExtentX        =   4101
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
         FldWidth2       =   1400
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilWorkerEn2 
         Height          =   315
         Left            =   1350
         TabIndex        =   7
         Top             =   1860
         Width           =   2325
         _ExtentX        =   4101
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
         FldWidth2       =   1400
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilWorkerEn1 
         Height          =   315
         Left            =   1350
         TabIndex        =   6
         Top             =   1500
         Width           =   2325
         _ExtentX        =   4101
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
         FldWidth2       =   1400
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilGroupCode 
         Height          =   285
         Left            =   1350
         TabIndex        =   5
         Top             =   1200
         Width           =   2325
         _ExtentX        =   4101
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
         FldWidth1       =   600
         FldWidth2       =   1400
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilPRCode 
         Height          =   285
         Left            =   1350
         TabIndex        =   4
         Top             =   915
         Width           =   2325
         _ExtentX        =   4101
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
         FldWidth1       =   600
         FldWidth2       =   1400
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilReasonCode 
         Height          =   285
         Left            =   1350
         TabIndex        =   3
         Top             =   615
         Width           =   2325
         _ExtentX        =   4101
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
         FldWidth1       =   600
         FldWidth2       =   1400
         F2Corresponding =   -1  'True
      End
      Begin Gi_Time.GiTime gdtPRTime 
         Height          =   315
         Left            =   1350
         TabIndex        =   2
         Top             =   285
         Width           =   2325
         _ExtentX        =   4101
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
      Begin VB.TextBox txtPRNote 
         Height          =   705
         Left            =   3870
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   15
         Top             =   1770
         Width           =   3675
      End
      Begin Gi_Multi.GiMulti gimFaciCode 
         Height          =   375
         Left            =   1650
         TabIndex        =   10
         Top             =   2550
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   661
         ButtonCaption   =   "設備項目"
         Enabled         =   0   'False
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin VB.Label lblResvPRTime 
         AutoSize        =   -1  'True
         Caption         =   "預約關機時間"
         Height          =   195
         Left            =   4140
         TabIndex        =   69
         Top             =   1200
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lblWorkerEn3 
         AutoSize        =   -1  'True
         Caption         =   "工作人員三"
         Height          =   195
         Left            =   300
         TabIndex        =   62
         Top             =   2850
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblPRNote 
         AutoSize        =   -1  'True
         Caption         =   "派工單備註"
         Height          =   195
         Left            =   3870
         TabIndex        =   60
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label lblWorkerEn2 
         AutoSize        =   -1  'True
         Caption         =   "工作人員二"
         Height          =   195
         Left            =   150
         TabIndex        =   59
         Top             =   1905
         Width           =   975
      End
      Begin VB.Label lblPRTime 
         AutoSize        =   -1  'True
         Caption         =   "預約派工時間"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   150
         TabIndex        =   58
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label lblPRCode 
         AutoSize        =   -1  'True
         Caption         =   "停拆機類別"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   150
         TabIndex        =   57
         Top             =   975
         Width           =   975
      End
      Begin VB.Label lblReasonCode 
         AutoSize        =   -1  'True
         Caption         =   "停拆機原因"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   150
         TabIndex        =   56
         Top             =   675
         Width           =   975
      End
      Begin VB.Label lblWorkerEn1 
         AutoSize        =   -1  'True
         Caption         =   "工作人員一"
         Height          =   195
         Left            =   150
         TabIndex        =   55
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblGroupCode 
         AutoSize        =   -1  'True
         Caption         =   "工程組別"
         Height          =   195
         Left            =   150
         TabIndex        =   54
         Top             =   1260
         Width           =   780
      End
   End
   Begin TabDlg.SSTab stData 
      Height          =   4215
      Left            =   120
      TabIndex        =   25
      Top             =   3570
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7435
      _Version        =   393216
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&1.收費單據編號產生"
      TabPicture(0)   =   "SO3420A.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblBillNo"
      Tab(0).Control(1)=   "lbl1"
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(3)=   "lblCount"
      Tab(0).Control(4)=   "lblClassName"
      Tab(0).Control(5)=   "lbl6"
      Tab(0).Control(6)=   "txtBillNo"
      Tab(0).Control(7)=   "ggrData"
      Tab(0).Control(8)=   "cmdDelete"
      Tab(0).Control(9)=   "chkUseOldBillNo"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "&2.依大樓編號產生"
      TabPicture(1)   =   "SO3420A.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblRange"
      Tab(1).Control(1)=   "gmsClassCode1"
      Tab(1).Control(2)=   "gmsStatusCode"
      Tab(1).Control(3)=   "gmsMduId2"
      Tab(1).Control(4)=   "gmsMduId"
      Tab(1).Control(5)=   "cboMduRange"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "&3.整批派拆"
      TabPicture(2)   =   "SO3420A.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraUCCode"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.ComboBox cboMduRange 
         Height          =   315
         ItemData        =   "SO3420A.frx":0496
         Left            =   -73950
         List            =   "SO3420A.frx":04A0
         Style           =   2  '單純下拉式
         TabIndex        =   32
         Top             =   660
         Width           =   1755
      End
      Begin VB.Frame fraUCCode 
         Height          =   3705
         Left            =   270
         TabIndex        =   70
         Top             =   360
         Width           =   11295
         Begin VB.OptionButton optPayKind 
            Caption         =   "後付制"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   1
            Left            =   2250
            TabIndex        =   88
            Top             =   630
            Width           =   945
         End
         Begin VB.OptionButton optPayKind 
            Caption         =   "預付制"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   0
            Left            =   1200
            TabIndex        =   87
            Top             =   630
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.TextBox txtPrtSno2 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   9690
            MaxLength       =   12
            TabIndex        =   40
            Top             =   180
            Width           =   1425
         End
         Begin VB.TextBox txtPrtSno1 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   8010
            MaxLength       =   12
            TabIndex        =   39
            Top             =   180
            Width           =   1425
         End
         Begin Gi_Multi.GiMulti gimWipCode3 
            Height          =   345
            Left            =   5310
            TabIndex        =   49
            Top             =   2910
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   609
            ButtonCaption   =   "派工狀態3"
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin Gi_Multi.GiMulti gimCMCode 
            Height          =   345
            Left            =   240
            TabIndex        =   45
            Top             =   2250
            Width           =   10665
            _ExtentX        =   18812
            _ExtentY        =   609
            ButtonCaption   =   "收費方式"
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin CS_Multi.CSmulti gimCitemCode 
            Height          =   345
            Left            =   240
            TabIndex        =   42
            Top             =   1245
            Width           =   10665
            _ExtentX        =   18812
            _ExtentY        =   609
            ButtonCaption   =   "收費項目"
         End
         Begin Gi_Date.GiDate gdaShouldDate1 
            Height          =   315
            Left            =   1170
            TabIndex        =   35
            Top             =   180
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            ForeColor       =   255
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
         Begin Gi_Date.GiDate gdaShouldDate2 
            Height          =   315
            Left            =   2550
            TabIndex        =   36
            Top             =   180
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            ForeColor       =   255
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
         Begin Gi_Date.GiDate gdaRealDate1 
            Height          =   315
            Left            =   4620
            TabIndex        =   37
            Top             =   180
            Width           =   1095
            _ExtentX        =   1931
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
         Begin Gi_Date.GiDate gdaRealDate2 
            Height          =   315
            Left            =   6000
            TabIndex        =   38
            Top             =   180
            Width           =   1095
            _ExtentX        =   1931
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
         Begin Gi_Multi.GiMulti gimUCCode 
            Height          =   345
            Left            =   240
            TabIndex        =   43
            Top             =   1575
            Width           =   10665
            _ExtentX        =   18812
            _ExtentY        =   609
            ButtonCaption   =   "未收原因"
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin CS_Multi.CSmulti gimClassCode 
            Height          =   345
            Left            =   240
            TabIndex        =   44
            Top             =   1920
            Width           =   10665
            _ExtentX        =   18812
            _ExtentY        =   609
            ButtonCaption   =   "客戶類別"
         End
         Begin Gi_Multi.GiMulti gimBillType 
            Height          =   345
            Left            =   240
            TabIndex        =   41
            Top             =   930
            Width           =   10665
            _ExtentX        =   18812
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
         Begin CS_Multi.CSmulti gimMduId2 
            Height          =   345
            Left            =   240
            TabIndex        =   50
            Top             =   3270
            Width           =   10665
            _ExtentX        =   18812
            _ExtentY        =   609
            ButtonCaption   =   "大樓名稱"
         End
         Begin Gi_Multi.GiMulti gimCustStatusCode2 
            Height          =   345
            Left            =   240
            TabIndex        =   48
            Top             =   2910
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   609
            ButtonCaption   =   "客戶狀態"
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin Gi_Multi.GiMulti gimAreaCode 
            Height          =   375
            Left            =   5310
            TabIndex        =   47
            Top             =   2565
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   661
            ButtonCaption   =   "行  政  區"
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin Gi_Multi.GiMulti gimServCode 
            Height          =   345
            Left            =   240
            TabIndex        =   46
            Top             =   2565
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   609
            ButtonCaption   =   "服  務  區"
            FontSize        =   9
            FontName        =   "新細明體"
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "繳付類別"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   300
            TabIndex        =   86
            Top             =   630
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "印單序號"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   7170
            TabIndex        =   82
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "至"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   9480
            TabIndex        =   81
            Top             =   255
            Width           =   195
         End
         Begin VB.Label lblRealDay 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "實收日期"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   3750
            TabIndex        =   74
            Top             =   240
            Width           =   780
         End
         Begin VB.Label lblD1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "至"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   5760
            TabIndex        =   73
            Top             =   255
            Width           =   195
         End
         Begin VB.Label lblShouldDay 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "應收日期"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   300
            TabIndex        =   72
            Top             =   240
            Width           =   780
         End
         Begin VB.Label lblD3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "至"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2310
            TabIndex        =   71
            Top             =   255
            Width           =   195
         End
      End
      Begin VB.CheckBox chkUseOldBillNo 
         Caption         =   "是否使用舊版單號"
         Height          =   195
         Left            =   -72090
         TabIndex        =   27
         Top             =   540
         Width           =   1875
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "刪除(&D)"
         Height          =   345
         Left            =   -64740
         TabIndex        =   28
         Top             =   450
         Width           =   1215
      End
      Begin CS_Multi.CSmulti gmsMduId 
         Height          =   465
         Left            =   -74820
         TabIndex        =   33
         Top             =   1200
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   820
         ButtonCaption   =   "大樓名稱"
      End
      Begin Gi_Multi.GiMulti gmsMduId2 
         Height          =   405
         Left            =   -74880
         TabIndex        =   65
         Top             =   3420
         Visible         =   0   'False
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   714
         ButtonCaption   =   "大樓名稱"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Multi.GiMulti gmsStatusCode 
         Height          =   465
         Left            =   -74820
         TabIndex        =   34
         Top             =   1740
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   820
         ButtonCaption   =   "客戶狀態"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin prjGiGridR.GiGridR ggrData 
         Height          =   3075
         Left            =   -74820
         TabIndex        =   31
         Top             =   930
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5424
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
      Begin VB.TextBox txtBillNo 
         Height          =   315
         Left            =   -73800
         MaxLength       =   15
         TabIndex        =   26
         Top             =   480
         Width           =   1575
      End
      Begin CS_Multi.CSmulti gmsClassCode1 
         Height          =   465
         Left            =   -74820
         TabIndex        =   83
         Top             =   2250
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   820
         ButtonCaption   =   "客戶類別一"
      End
      Begin VB.Label lbl6 
         AutoSize        =   -1  'True
         Caption         =   "客戶類別:"
         ForeColor       =   &H00C000C0&
         Height          =   195
         Left            =   -69720
         TabIndex        =   79
         Top             =   540
         Width           =   825
      End
      Begin VB.Label lblClassName 
         AutoSize        =   -1  'True
         ForeColor       =   &H00C000C0&
         Height          =   195
         Left            =   -68760
         TabIndex        =   78
         Top             =   540
         Width           =   45
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         Caption         =   "99999"
         Height          =   195
         Left            =   -65850
         TabIndex        =   68
         Top             =   540
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "筆數 :"
         Height          =   195
         Left            =   -66480
         TabIndex        =   67
         Top             =   540
         Width           =   480
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "單據序,單據編號,客戶編號,姓名,客戶類別,客戶狀態,派工一,派工二,派工三,欠費狀態"
         Height          =   195
         Left            =   -74730
         TabIndex        =   61
         Top             =   840
         Visible         =   0   'False
         Width           =   7035
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "選擇範圍"
         Height          =   195
         Left            =   -74790
         TabIndex        =   53
         Top             =   735
         Width           =   780
      End
      Begin VB.Label lblBillNo 
         AutoSize        =   -1  'True
         Caption         =   "單據編號"
         Height          =   195
         Left            =   -74670
         TabIndex        =   52
         Top             =   540
         Width           =   780
      End
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   5310
      TabIndex        =   1
      Top             =   90
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
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1290
      TabIndex        =   0
      Top             =   90
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
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "公司別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   270
      TabIndex        =   64
      Top             =   150
      Width           =   585
   End
   Begin VB.Label lblServiceType 
      AutoSize        =   -1  'True
      Caption         =   "服務類別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4350
      TabIndex        =   63
      Top             =   150
      Width           =   780
   End
End
Attribute VB_Name = "frmSO3420A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnMDB As New ADODB.Connection
Private rsArray As New ADODB.Recordset
Private rsSo001 As New ADODB.Recordset
Private rsSO014 As New ADODB.Recordset
Private rsCd007 As New ADODB.Recordset
Private intRefNo As Integer '參考號(在停拆移機類別ValiDate 時取得！)
Private strServiceCode As String
Private strServiceName As String
Private strServiceContentCode As String
Private strServiceContentName As String
Private blnIsOK As Boolean
Private blnAddCodSvr As Boolean
Private strBillNos As String
Dim lngEditMode As giEditModeEnu
Dim strUpdTime As String, strDTUpdTime As String
Dim strSNo As String
Dim blnFlag As Boolean
Dim intPara23 As Integer
Dim FSO As New FileSystemObject
Dim file As TextStream
Dim strFileName As String
Dim blnViewOk As Boolean
Private blnPrcSO033 As Boolean
Private importExcelData As Dictionary
'#3474 多增加csGUINo 列舉值 By Kin 2008/05/12
Private Enum csLogError
    cs001StatusError = 0
    cs009StatusError = 1
    cs003RangeError = 2
    csOtherError = 3
    cn004HaveSTB = 4
    csGUINo = 5
End Enum

Private Sub chkPRFaci_Click()
    On Error Resume Next
        If chkPRFaci.Value = 1 Then
            gimFaciCode.Enabled = True
            chkSetFaci.Value = 0
        Else
            gimFaciCode.Clear
            If chkSetFaci.Value = 0 Then
                gimFaciCode.Enabled = False
            End If
        End If
End Sub

Private Sub chkSetFaci_Click()
  On Error Resume Next
    If chkSetFaci.Value = 1 Then
        gimFaciCode.Enabled = True
        chkPRFaci.Value = 0
    Else
        If chkPRFaci.Value = 0 Then
            gimFaciCode.Enabled = False
        End If
    End If
End Sub

Private Sub cmdInExcel_Click()
  On Error Resume Next
    If Not importExcelData Is Nothing Then
        importExcelData.RemoveAll
        Set importExcelData = Nothing
    End If
    frmSO3420D.uExcelType = 0
    frmSO3420D.Show 1
End Sub

Private Sub cmdInExcel2_Click()
  On Error Resume Next
    If Not importExcelData Is Nothing Then
        importExcelData.RemoveAll
        Set importExcelData = Nothing
    End If
    frmSO3420D.uExcelType = 1
    frmSO3420D.Show 1
End Sub

Private Sub gdaRealDate2_Validate(Cancel As Boolean)
On Error Resume Next
    If gdaRealDate1.GetValue & "" <> "" Then
        If CDate(gdaRealDate1.GetValue(True)) > CDate(gdaRealDate2.GetValue(True)) Then
            MsgBox "實收起始日期不得大於實收終止日", vbInformation, "警告訊息"
            gdaRealDate2.SetValue gdaRealDate1.GetValue
            Cancel = True
        End If
    End If
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
 On Error Resume Next
    If gdaShouldDate1.GetValue & "" <> "" Then
        If CDate(gdaShouldDate1.GetValue(True)) > CDate(gdaShouldDate2.GetValue(True)) Then
            MsgBox "應收起始日期不得大於應收終止日", vbInformation, "警告訊息"
            gdaShouldDate2.SetValue gdaShouldDate1.GetValue
            Cancel = True
        End If
    End If
End Sub

Private Sub optClsBill_Click()
    On Error Resume Next
'收費單一併結案：
'若 Checked：則enabled 結案短收原因(gilSTCode)及結帳日期(gdaClsDate)，結帳日期預設為今日
'若Not Checked ：則清除及Disable 結案短收原因，結帳日期
    If optClsBill.Value = True Then
        fraBillClose.Enabled = True
        gilUCCode.Enabled = False
        gilUCCode.Clear
        gdaClsDate.SetValue Format(RightDate, "YYYYMMDD")
    Else
        gilSTCode.Clear
        gilCancelCode.Clear
        gdaClsDate.SetValue ""
    End If
End Sub

Private Sub chkClsWork_Click()
    If chkClsWork.Value = 1 Then
        If gdtPRTime.Text = "" Then gdtPRTime.Text = RightNow
        'chkPREBox = 1
        chkPREBox.Enabled = True
    Else
        chkPREBox = 0
        chkPREBox.Enabled = False
    End If
End Sub

Private Sub chkPREBox_Click()
    On Error Resume Next
        'If chkClsWork = 1 Then chkPREBox = 1
        If chkPREBox = 1 Then
            gdaEBoxPRTime.Enabled = True
            gdaEBoxPRTime.SetValue RightDate
        Else
            gdaEBoxPRTime.Enabled = False
            gdaEBoxPRTime.SetValue ""
        End If
End Sub

'
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
        If rsArray.RecordCount = 0 Then Exit Sub
        rsArray.Delete
        rsArray.MoveNext
        If rsArray.RecordCount > 0 Then rsArray.MoveLast
        lblCount = rsArray.RecordCount
        Set ggrData.Recordset = rsArray
        ggrData.Refresh
End Sub

Private Sub cmdOK_Click()
    On Error GoTo chkErr
    Dim strSQL As String
    Dim rsData As New ADODB.Recordset
    Dim strBillNo As String
    Dim rsSO009 As New ADODB.Recordset
    Dim rsSO004 As New ADODB.Recordset
    Dim rsSO033 As New ADODB.Recordset
    Dim rsSO006 As New ADODB.Recordset
    Dim rsSo004D As New ADODB.Recordset
    Dim rsPRSO004 As New ADODB.Recordset
    Dim rsClone As New ADODB.Recordset
    Dim objCloseBill As Object
    Dim clsAlterWip As Object
    Dim objNagraSTB As Object
    Dim strFrom As String, lngCount As Long
    Dim lngError As Long
    Dim blnFlag As Boolean
    Dim strMsg As String
    Dim blnUpdateFlag As Boolean
    Dim strChoose033 As String
    Dim lngSTB As Long '計算STB關機筆數暫存
    Dim lngSTBCount As Long '計算STB關機實際筆數
    Dim lngClsWorkCount As Long '計算派工單結案筆數
    Dim blnShowView As Boolean
    Set objCloseBill = CreateObject("csCloseService.clsCloseService")
        If IsDataOk = False Then Screen.MousePointer = vbDefault: Exit Sub
        DoEvents
        Me.Enabled = False
        Screen.MousePointer = vbHourglass
        Call OpenLogFile
        '若範圍為1
        'gmsMduid 為 大樓條件
        'GmsStatusCode 為客戶狀態條件
        If Not GetRS(rsCd007, "Select * From " & GetOwner & "Cd007 Where CodeNo = " & gilPRCode.GetCodeNo) Then Screen.MousePointer = vbDefault: Me.Enabled = True: Exit Sub
        '#4161 增加FaciSeqNo,FaciSNo,Item欄位以便新增至SO180 By Kin 2008/10/30
        If stData.Tab = 0 Then
            If Not GetRS(rsData, "Select Distinct CustID, BillNo,CustName,PrtSNo, " & _
                "ServCode,Tel1,CompCode,InstAddrNo,InstAddress,CustStatusCode,CustStatusName, " & _
                "ClassName1,WipName1,WipName2,WipName3,ServiceType,Balance,StartDateA as RealStartDate,StopDateA as RealStopDate,SeqNo," & _
                "CitemCode,CitemName,MediaBillNo,GUINo,FaciSeqNo,FaciSNO,Item From So3420A Order By CustId", cnMDB) Then Screen.MousePointer = vbDefault: Me.Enabled = True: Exit Sub
        ElseIf stData.Tab = 1 Then
            If gilCompCode.GetCodeNo <> "" Then strSQL = strSQL & " And A.CompCode = " & gilCompCode.GetCodeNo
            If gilServiceType.GetCodeNo <> "" Then strSQL = strSQL & " And B.ServiceType = '" & gilServiceType.GetCodeNo & "'"
            '*******************************************************************************
            '#3474 若客戶類別有值,則多串SO001.ClassCode1 By Kin 2008/05/12
            If gmsClassCode1.GetQryStr <> "" Then
                strSQL = strSQL & " And A.ClassCode1 " & gmsClassCode1.GetQryStr
            End If
            '*******************************************************************************
            If cboMduRange.ListIndex = 0 Then
                strSQL = "Select " & GetUseIndexStr("SO001", "MduId") & " A.CustId,B.ServiceType,B.CustStatusCode,B.CustStatusName From " & _
                            GetOwner & "SO001 A," & GetOwner & "SO002 B Where A.CustId = B.CustId" & _
                            IIf(gmsMduId.GetQueryCode <> "", " And A.Mduid In (" & gmsMduId.GetQueryCode & ") ", "") & strSQL
                '若有客戶狀態條件，則Where =Where + And A.CustStatuscode" + 客戶狀態條件
                If gmsStatusCode.GetQryStr <> "" Then
                    strSQL = strSQL & " And B.CustStatusCode in (" & gmsStatusCode.GetQueryCode & ") "
                End If
            ElseIf cboMduRange.ListIndex = 1 Then
                strSQL = " Where A.Custid=B.Custid " & IIf(gmsMduId.GetQryStr <> "", " And A.Mduid " & gmsMduId.GetQryStr, "") & strSQL
                '若有客戶狀態條件，則Where =where + " And A.CustStatusCode" & <客戶狀態條件>
                'Where =where +" And B.Realperiod > 0 and B.UCCode is Not Null"
                If gmsStatusCode.GetQryStr <> "" Then
                    strFrom = "," & GetOwner & "SO002 C"
                    strSQL = strSQL & " And B.CustId = C.CustId And B.ServiceType = C.ServiceType And C.CustStatusCode in (" & gmsStatusCode.GetQueryCode & ") and B.RealPeriod > 0 And B.UCCode is not null"
                End If
                '#3474 多增加發票號碼欄位、應收日期、未收原因、未收原因名稱 By Kin 2008/05/12
                '#4161 增加FaciSeqNo,FaciSNO,Item以便新增至SO180 By Kin 2008/10/30
                strSQL = "Select Distinct A.Custid,B.ServiceType,B.CitemCode,B.CitemName,B.ShouldDate,B.RealStartDate," & _
                        "B.RealStopDate,B.SeqNo,B.BillNo,B.MediaBillNo,B.PrtSNo,B.GUINo,B.UCCode," & _
                        "B.UCName,B.ShouldAmt,B.RealAmt,B.RealDate,B.FaciSeqNo,B.FaciSNo,B.Item " & _
                        "From " & GetOwner & "So001 A," & GetOwner & "So033 B " & strFrom & strSQL
            End If
            strSQL = strSQL & " Order by A.CustId"
            If Not GetRS(rsData, strSQL, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then Screen.MousePointer = vbDefault: Me.Enabled = True: Exit Sub
        ElseIf stData.Tab = 2 Then
            If gilCompCode.GetCodeNo <> "" Then strSQL = strSQL & " And A.CompCode = " & gilCompCode.GetCodeNo
            If gilServiceType.GetCodeNo <> "" Then strSQL = strSQL & " And A.ServiceType = '" & gilServiceType.GetCodeNo & "'"
            If gdaShouldDate1.GetValue <> "" Then strSQL = strSQL & " And A.ShouldDate >= " & GetNullString(gdaShouldDate1.GetValue(True), giDateV)
            If gdaShouldDate2.GetValue <> "" Then strSQL = strSQL & " And A.ShouldDate < " & GetNullString(gdaShouldDate2.GetValue(True), giDateV) & "+1"
            If gdaRealDate1.GetValue <> "" Then strSQL = strSQL & " And A.RealDate >= " & GetNullString(gdaRealDate1.GetValue(True), giDateV)
            If gdaRealDate2.GetValue <> "" Then strSQL = strSQL & " And A.RealDate < " & GetNullString(gdaRealDate2.GetValue(True), giDateV) & "+1"
            '#3144 增加印單序號條條件 By Kin 2007/4/4
            If Trim(txtPrtSNo1) & "" <> "" Then strSQL = strSQL & " And A.PrtSNo>='" & Trim(txtPrtSNo1) & "'"
            If Trim(txtPrtSNo2) & "" <> "" Then strSQL = strSQL & " And A.PrtSNo<='" & Trim(txtPrtSNo2) & "'"
            
            '93/09/23 Jacky By Casper
            If gimBillType.GetQryStr <> "" Then strSQL = strSQL & " And substr(A.BillNo,7,1) " & gimBillType.GetQryStr
            If gimCitemCode.GetQryStr <> "" Then strSQL = strSQL & " And A.CitemCode " & gimCitemCode.GetQryStr
            If gimUCCode.GetDispStr <> "" Then
                If gimUCCode.GetQryStr = "" Then
                    strSQL = strSQL & " And A.UCCode is not null "
                Else
                    strSQL = strSQL & " And A.UCCode " & gimUCCode.GetQryStr
                End If
            End If
            If gimClassCode.GetQryStr <> "" Then strSQL = strSQL & " And A.ClassCode " & gimClassCode.GetQryStr
            If gimCMCode.GetQryStr <> "" Then strSQL = strSQL & " And A.CMCode " & gimCMCode.GetQryStr
            If gimServCode.GetQryStr <> "" Then strSQL = strSQL & " And A.ServCode " & gimServCode.GetQryStr
            If gimAreaCode.GetQryStr <> "" Then strSQL = strSQL & " And A.AreaCode " & gimAreaCode.GetQryStr
            '#5683 增加繳付類別 By Kin 2010/08/30
            If stData.Tab = 2 Then
                If optPayKind(0).Value Then
                    strSQL = strSQL & " AND NVL(A.PAYKIND,0)= 0 "
                Else
                    strSQL = strSQL & " AND NVL(A.PAYKIND,0)= 1 "
                End If
            End If
            '#5324 增加外部匯入資料 By Kin 2010/03/19
            If strBillNos <> "" Then strSQL = strSQL & " And A.BillNo In(" & strBillNos & ") "
            strBillNos = Empty
            '94/08/01 Jacky Jacy 提 1443 客戶狀態,派工狀態3,大樓名稱
            If gimMduId2.GetQryStr <> "" Then
                strSQL = strSQL & " And A.MduId " & gimMduId2.GetQryStr
            Else
                If gimMduId2.GetDispStr <> "" Then strSQL = strSQL & " And A.MduId is Not Null"
            End If
            strChoose033 = strSQL
            If gimCustStatusCode2.GetQryStr <> "" Then strSQL = strSQL & " And B.CustStatusCode " & gimCustStatusCode2.GetQryStr
            If gimWipCode3.GetQryStr <> "" Then strSQL = strSQL & " And B.WipCode3 " & gimWipCode3.GetQryStr
            
            
            If InStr(1, strSQL, "B.") > 0 Then
                strSQL = " From " & GetOwner & "SO033 A," & GetOwner & "SO002 B Where A.CustId = B.CustId And A.CompCode=B.CompCode And A.ServiceType=B.ServiceType " & strSQL
            Else
                strSQL = " From " & GetOwner & "SO033 A Where " & Mid(strSQL, 6)
            End If
            '#3474 多增加發票號碼欄位、應收日期、未收原因、未收原因名稱 By Kin 2008/05/12
            '#4161 增加FaciSeqNo,FaciSNO,Item以便新增至SO180 By Kin 2008/10/30
            strSQL = "Select Distinct A.Custid,A.ServiceType,A.CitemCode,A.CitemName," & _
                     "A.ShouldDate,A.RealStartDate,A.RealStopDate,A.SeqNo,A.BillNo," & _
                     "A.MediaBillNo,A.PrtSNo,A.GUINo,A.UCCode,A.UCName,A.RealAmt,A.FaciSeqNo,A.FaciSNO,A.Item," & _
                     "A.ShouldAmt,A.RealDate,A.RowId " & strSQL & " Order By A.CustId"
            If Not GetRS(rsData, strSQL) Then Screen.MousePointer = vbDefault: Me.Enabled = True: Exit Sub
        End If
        txtSQL = strSQL
        If rsData.EOF Then MsgBox "無符合條件資料！", vbExclamation, gimsgPrompt: Screen.MousePointer = vbDefault: Me.Enabled = True: Exit Sub
        '取得當時時間字串
        strUpdTime = Format(RightNow, "YYYY/MM/DD HH:MM")
        strDTUpdTime = GetDTString(RightNow)
        rsData.MoveFirst
'        Screen.MousePointer = vbHourglass
        Set clsAlterWip = CreateObject("csAlterWip4.clsAlterWip3")
        clsAlterWip.uOwnerName = GetOwner
        On Error GoTo ErrTrans
        '****************************************************
        '#3474 要Show出瀏覽資料 By Kin 2008/05/13
        If (stData.Tab = 1 And InStr(1, cboMduRange.Text, "只有待收戶") > 0) Or (stData.Tab) = 2 Then
            blnViewOk = False
            Set rsClone = rsData.Clone
            frmSO3420C.uRS = rsClone
            frmSO3420C.Show vbModal
            blnShowView = True
        End If
        If Not blnViewOk And blnShowView Then
            MsgBox "作業取消", vbExclamation, "訊息": GoTo Finall
        Else
            DoEvents
            Me.Enabled = False
            Screen.MousePointer = vbHourglass
        End If
        '****************************************************
        lngClsWorkCount = 0
        lngCount = 0
        lngError = 0
        lngSTB = 0
        lngSTBCount = 0
        Dim aBillNo As String
        gcnGi.BeginTrans '開始異動交易！
         Do While Not rsData.EOF
            aBillNo = ""
            strMsg = ""
            blnPrcSO033 = True
            'to check billno is existed to update datafield by excel for Minchen By Kin
            If stData.Tab = 2 Then
                aBillNo = rsData("BillNo")
            End If
            '***************************************************************************************************
            '#3474 檢查收費資料發票號碼是否有值(依大樓編號頁籤只找"只有待收戶") By Kin 2008/05/12
            '#4030 增加有發票號碼一樣要能產生拆機單
            If stData.Tab = 1 Then
                If InStr(1, cboMduRange.Text, "只有待收戶") > 0 Then
                    If Not IsNull(rsData("GUINo")) Then
                        Call InsertToLog(rsData, csGUINo, strMsg)
                        lngError = lngError + 1
                        blnFlag = True
                        blnPrcSO033 = False
                        'GoTo NextLoop
                    End If
                End If
            Else
                If Not IsNull(rsData("GUINo")) Then
                        Call InsertToLog(rsData, csGUINo, strMsg)
                        lngError = lngError + 1
                        blnFlag = True
                        blnPrcSO033 = False
                        'GoTo NextLoop
                End If
            End If
            '****************************************************************************************************
            If ChkDateRangeOK(rsData) Then
                blnFlag = True
                '檢查其他服務的狀態
                If chkOtherS.Value = 1 Then
                    If Not chkOtherService(rsData, strMsg) Then
                        Call InsertToLog(rsData, csOtherError, strMsg)
                        lngError = lngError + 1
                        blnFlag = False
                    End If
                End If
                '檢查是否為移機 93/08/04 Jacky By Bonnie
                If Not Chk009ReInst(rsData, strMsg) Then
                    Call InsertToLog(rsData, cs009StatusError, strMsg)
                    lngError = lngError + 1
                    blnFlag = False
                End If
                
                '檢核待拆設備 93/09/29 Jacky By Jacy
                If chkHaveFaciInst(rsData("CustId")) Then
                    Call InsertToLog(rsData, cn004HaveSTB, strMsg)
                    lngError = lngError + 1
                    '單筆可拆除,整批派拆不可產生拆機單
                    If stData.Tab <> 0 Then blnFlag = False
                End If
                
                '#5411 增加 rsPRSO004與rsSO004D給阿貴的元件，可以將指定的設備拆除掉 By Kin 2009/12/10
                If chkSetFaci.Value = 1 Then
                    If Not GetPrSO004(rsPRSO004, rsData("CustId") & "") Then
                        Call InsertToLog(rsData, csOtherError, "取PRSO004錯誤")
                        lngError = lngError + 1
                        blnFlag = False
                    End If
                    If Not rsPRSO004.EOF Then
                        If Not GetKindSO004D("拆除", rsSo004D, rsPRSO004) Then
                            Call InsertToLog(rsData, csOtherError, "取SO004D錯誤")
                            lngError = lngError + 1
                            blnFlag = False
                        End If
                    
                    End If
                Else
                    If Not GetRS(rsSo004D, "Select RowId,SO004D.* From " & GetOwner & "SO004D Where SeqNo = ''", , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                End If
                rsPRSO004.ActiveConnection = Nothing
                rsSo004D.ActiveConnection = Nothing
                
                If blnFlag Then
                    If Not OpenData2(rsData("CustId"), rsData("ServiceType") & "", blnFlag) Then GoTo ErrTrans
                    If blnFlag Then
                        '僅處理部分客戶狀態=正常、停機、拆機中客戶
                        If InStr("126", rsSo001("CustStatusCode")) > 0 Then
                            blnUpdateFlag = False
                            'Jacky 90/0716
                            '6=拆機中,不再產生拆機單
                            If rsSo001("CustStatusCode") <> 6 Then
                              
                                If Not InsertToSO00x(giEditModeInsert, rsData("CustId"), gilPRCode.GetCodeNo, rsSO009, aBillNo) Then GoTo ErrTrans
                                '#3254 如果"不新增至服務申告"，則維持原做法，傳進一個Nothing的Record，反之傳進一個SO006的Recordc By Kin 2007/07/18
                                Call ReturnSo006(rsSO006)
                                If Not clsAlterWip.Action(lngEditMode, gcnGi, rsSO009, rsSO004, rsSO033, , , , , False, rsSO006, , , , rsPRSO004, , , , , , , , , , , , , , , rsSo004D) Then GoTo ErrTrans
                                '#5683 以下呼叫Miggie的結清元件 By Kin
                                '#5878 如果選擇不處理不要處理結清 By Kin 2010/12/13
                                If (stData.Tab = 2) And (optRealStopDate.Value) Then
                                    If optPayKind(1).Value Then
                                        Set objCloseBill.objgcnGi = gcnGi
                                        objCloseBill.uParentForm = Me.Name
                                        objCloseBill.uOwnerName = GetOwner
                                        objCloseBill.uCompCode = gCompCode
                                        objCloseBill.uErrPath = ReadGICMIS1("ErrLogPath")
                                        objCloseBill.uPutGlobal = PutGlobal
                                        objCloseBill.uSOEXEPATH = ReadGICMIS1("SOEXEPATH")
                                        If rsSO009.RecordCount > 0 Then
                                            Dim rs033 As New ADODB.Recordset
                                            If Not CopySO033Record(rsData, rs033) Then GoTo ErrTrans
                                            If Not objCloseBill.ReceiveData(5, Format(gdtPRTime.GetValue(True), "yyyy/mm/dd"), _
                                                    rs033, Nothing, Nothing, Nothing, Nothing, garyGi(0), garyGi(1), _
                                                False, 1, 2, 0, 0, , , , , , rsSO009("SNO"), rsData("CustId")) Then
                                            
                                                If objCloseBill.uErrMsg <> "" Then
                                                    MsgBox "客編:" & rsData("CustId") & " " & objCloseBill.uErrMsg, vbCritical, "錯誤"
                                                End If
                                                GoTo ErrTrans
                                            End If
                                        End If
                                    End If
                                End If
                                blnUpdateFlag = True
                            End If
                            If chkClsWork.Value = 1 Then
                                '正常、停機:會產生拆機派工單，且將拆機派工單結案。
                                '拆機中:不會產生拆機派工單，且將原來拆機中派工單結案
                                If Not InsertToSO00x(giEditModeEdit, rsData("CustId"), gilPRCode.GetCodeNo, rsSO009, aBillNo) Then GoTo ErrTrans
                                '********************************************************************************************************
                                '#3474 要計算STB關機筆數 By Kin 2008/05/13
                                lngSTB = 0
                                If chkPREBox Then
                                    If Not subPREBox(objNagraSTB, rsData, rsSO009, strUpdTime, lngSTB) Then
                                        GoTo ErrTrans
                                    Else
                                        lngSTBCount = lngSTBCount + lngSTB
                                    End If
                                End If
                                '********************************************************************************************************
                                If Not clsAlterWip.Action(lngEditMode, gcnGi, rsSO009, rsSO004, rsSO033, , , , , False, , , , , rsPRSO004, , , , , , , , , , , , , , , rsSo004D) Then GoTo ErrTrans
                                lngClsWorkCount = lngClsWorkCount + 1
                                blnUpdateFlag = True
                            End If
                            '僅處理部分客戶狀態=正常、停機客戶 93/09/23 Jacky By Casper
                            If InStr("12", rsSo001("CustStatusCode")) > 0 Then
                                '可處理所有客戶狀態的欠費
                                '#4030 如果有發票號碼不處理收費資料部份 By Kin 2008/08/26
                                '#4161 測試不OK,不管收費單是否處理,都要產生關聯 By Kin 2009/01/14
                                If blnPrcSO033 Then
                                    If stData.Tab = 2 Then
                                        If Not importExcelData Is Nothing Then
                                            If importExcelData.Count > 0 Then
                                                If Not subCloseBill(rsData("CustId"), rsData("ServiceType") & "", strChoose033, aBillNo) Then GoTo ErrTrans
                                                blnUpdateFlag = True
                                            Else
                                                 If optClsBill Or optUCCode Then
                                                    If Not subCloseBill(rsData("CustId"), rsData("ServiceType") & "", strChoose033, aBillNo) Then GoTo ErrTrans
                                                    blnUpdateFlag = True
                                                End If
                                            End If
                                        Else
                                            If optClsBill Or optUCCode Then
                                                If Not subCloseBill(rsData("CustId"), rsData("ServiceType") & "", strChoose033, aBillNo) Then GoTo ErrTrans
                                                blnUpdateFlag = True
                                            End If
                                        End If
                                        
                                    Else
                                        If optClsBill Or optUCCode Then
                                            If Not subCloseBill(rsData("CustId"), rsData("ServiceType") & "", strChoose033, aBillNo) Then GoTo ErrTrans
                                            blnUpdateFlag = True
                                        End If
                                    End If
                                    '#4161 新增一筆記錄至SO180,工單與SO033產生關聯 By Kin 2008/10/30
                                    If stData.Tab = 1 And cboMduRange.ListIndex = 0 Then
                                    
                                    Else
                                        If Not InsSO180(rsData, rsSO009("SNO") & "") Then GoTo ErrTrans
                                    End If
                                End If
                            End If
                            If chkDelIntro.Value = 1 Then
                                If Not subDeleteIntro(rsData("CustId")) Then GoTo ErrTrans
                                blnUpdateFlag = True
                            End If
                            '93/10/21 增加判斷是否確定有更新才計數
                            'If blnUpdateFlag Then lngCount = lngCount + 1
                            lngCount = lngCount + 1
                        Else
                            Call InsertToLog(rsData, cs001StatusError)
                            lngError = lngError + 1
                        End If
                    End If
                End If
            Else
                Call InsertToLog(rsData, cs003RangeError)
                lngError = lngError + 1
            End If
NextLoop:
            DoEvents
            rsData.MoveNext
        Loop
        gcnGi.CommitTrans '確認異動交易！
        On Error GoTo chkErr
        '****************************************************************
        '#3474 訊息要更改 By Kin 2008/05/13
        'MsgBox "產生完畢！成功筆數: " & lngCount & vbCrLf & _
        '       "問題筆數:" & lngError, vbInformation, gimsgPrompt
        MsgBox "產生完畢！" & vbCrLf & _
               "成功派拆筆數: " & lngCount & vbCrLf & _
               "問題筆數:" & lngError & vbCrLf & _
               "STB關機筆數:" & lngSTBCount & vbCrLf & _
               "派工結案筆數:" & lngClsWorkCount, vbInformation, gimsgPrompt

        '***************************************************************
Finall:
        On Error Resume Next
        Call Close3Recordset(rsSO009)
        Call Close3Recordset(rsSO004)
        Call Close3Recordset(rsSO033)
        Call Close3Recordset(rsSo001)
        Call Close3Recordset(rsData)
        Call Close3Recordset(rsSO006)
        Call Close3Recordset(rsClone)
        If Not importExcelData Is Nothing Then
            importExcelData.RemoveAll
            Set importExcelData = Nothing
        End If
        
        Set objNagraSTB = Nothing
        Set clsAlterWip = Nothing
        Set objCloseBill = Nothing
        Me.Enabled = True
        Call NewRcd

        Screen.MousePointer = vbDefault
        If lngError > 0 Then Shell "notepad " & strFileName, vbNormalNoFocus
    Exit Sub
ErrTrans:
        gcnGi.RollbackTrans
        MsgBox "產生失敗！請重新操作！", vbCritical, giMsgWarning
         If Not importExcelData Is Nothing Then
            importExcelData.RemoveAll
            Set importExcelData = Nothing
        End If
        Me.Enabled = True
        Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdOk_Click"
End Sub
Private Function CopySO033Record(ByRef rsSource As ADODB.Recordset, ByRef rs033 As ADODB.Recordset) As Boolean
  On Error GoTo chkErr
    Dim aSQL As String
    aSQL = "SELECT A.ROWID,A.* FROM " & GetOwner & "SO033 A WHERE A.ROWID='" & rsSource("RowId") & "'"
    If Not GetRS(rs033, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    Set rs033.ActiveConnection = Nothing
    
    CopySO033Record = True
    Exit Function
chkErr:
    ErrSub Me.Name, "CopySO033Record"
End Function
Private Function GetPrSO004(ByRef rs As ADODB.Recordset, ByVal aCustId As String) As Boolean
  On Error GoTo chkErr
    Dim strQry As String
    Dim strWhere As String
    If gimFaciCode.GetQueryCode <> "" Then
        strWhere = " AND FaciCode IN(" & gimFaciCode.GetQueryCode & ") "
    End If
    strQry = "SELECT A.* FROM " & GetOwner & "SO004 A " & _
                " WHERE CUSTID=" & aCustId & _
                " And PrDate is null and PRSNo is null and GetDate is null " & _
                strWhere
    If Not GetRS(rs, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    GetPrSO004 = True
    Exit Function
chkErr:
  ErrSub Me.Name, "GetPrSO004"
End Function
'#4161 產生拆機單要與SO033產生關聯 By Kin 2008/10/30
Private Function InsSO180(rsSO033 As Recordset, ByVal strSNo As String) As Boolean
  On Error GoTo chkErr
    Dim rsSO180 As New ADODB.Recordset
    If Not GetRS(rsSO180, "Select * From " & GetOwner & "SO180 Where 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'    Do While Not rsSO033.EOF
        With rsSO180
            .AddNew
            .Fields("CustId") = rsSO033("CustId")
            .Fields("MediaBillNo") = IIf(IsNull(rsSO033("MediaBillNo")), Null, rsSO033("MediaBillNo") & "")
            .Fields("SNo") = strSNo
            .Fields("FaciSNo") = IIf(IsNull(rsSO033("FaciSNo")), Null, rsSO033("FaciSNo") & "")
            .Fields("FaciSeqNo") = IIf(IsNull(rsSO033("FaciSeqNo")), Null, rsSO033("FaciSeqNo") & "")
            .Fields("Batch") = 1
            .Fields("UpdEn") = garyGi(1)
            .Fields("UpdTime") = GetDTString(strUpdTime)
            .Fields("BillNo") = rsSO033("BillNo") & ""
            .Fields("Item") = rsSO033("Item")
            .Update
        End With
  '      rsSO033.MoveNext
 '   Loop
    On Error Resume Next
    Call Close3Recordset(rsSO180)
    InsSO180 = True
    Exit Function
chkErr:
  Call ErrSub(Me.Name, "InsSO180")
End Function
Private Function ChkDateRangeOK(rsData As ADODB.Recordset) As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
        '當A項與收費項目期限有Cover 時做Log File
        If InStr(1, UCase(rsData.Source), "CITEMCODE") = 0 Then ChkDateRangeOK = True: Exit Function
        If Not GetRS(rsTmp, "Select RowId,StartDate,StopDate From " & GetOwner & "SO003 Where CustId =" & rsData("Custid") & " And CitemCode =" & rsData("CitemCode")) Then Exit Function
        If rsTmp.RecordCount > 0 Then
            If rsTmp.RecordCount > 1 Then rsTmp.Filter = "SeqNo = " & rsData("SeqNo")
            If rsTmp.RecordCount > 0 Then
                If rsData.Fields("RealStartDate") < rsTmp("StopDate") And rsData("RealStopDate") > rsTmp("StartDate") Then
                    Call CloseRecordset(rsTmp)
                    Exit Function
                End If
            End If
        End If
        Call CloseRecordset(rsTmp)
        ChkDateRangeOK = True
    Exit Function
chkErr:
    If Err.Number = 3265 Then
        ChkDateRangeOK = True
        Exit Function
    Else
        ErrSub Me.Name, "ChkDateRangeOK"
    End If
End Function

Private Function chkOtherService(rsData As ADODB.Recordset, ByRef strMsg As String) As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
        If Not GetRS(rsTmp, "Select CustStatusCode,WipCode1,ServiceType From " & GetOwner & "SO002 Where CustId = " & rsData("CustId") & " And ServiceType <> '" & gilServiceType.GetCodeNo & "'") Then Exit Function
        Do While Not rsTmp.EOF
            '正常戶..不能拆機
            If rsTmp("CustStatusCode") = 1 Then strMsg = "客戶編號:" & rsData("CustId") & ", 服務類別" & rsTmp("ServiceType") & ",客戶狀態為正常,不可派拆機": Exit Function
            '促銷戶..已派工不能拆機
            If rsTmp("CustStatusCode") = 5 And rsTmp("WipCode1") = 1 Then strMsg = "該客服務類別" & rsTmp("ServiceType") & ",客戶狀態為促銷已派工,不可派拆機": Exit Function
            rsTmp.MoveNext
        Loop
        chkOtherService = True
    Exit Function
chkErr:
    ErrSub Me.Name, "chkOtherService"
End Function

'檢查該客戶是否為拆機 93/08/04 Jacky From Bonnie
Private Function Chk009ReInst(rsData As ADODB.Recordset, ByRef strMsg As String) As Boolean
    On Error GoTo chkErr
    Dim lngCount As Long
        lngCount = Val(GetRsValue("Select Count(*) From " & GetOwner & "SO009 Where CustId = " & rsData("CustId") & " And ServiceType='" & rsData("ServiceType") & "' And SignDate Is Null And PRCode in (Select CodeNo From " & GetOwner & "CD007 Where Refno =3 )") & "")
        If lngCount > 0 Then strMsg = "客戶編號:" & rsData("CustId") & ", 服務類別" & rsData("ServiceType") & ",派工狀態3為移機,不可派拆機": Exit Function
        Chk009ReInst = True
    Exit Function
chkErr:
    ErrSub Me.Name, "Chk009ReInst"
End Function

'檢查該客戶是否有裝該設備 93/09/29 Jacky From Jacy
Private Function chkHaveFaciInst(lngCustId As Long) As Boolean
  On Error Resume Next
    Dim strWhere As String
    Dim intCount As Integer
        If chkPRFaci.Value = 0 Then Exit Function
        
        strWhere = "CustId = " & lngCustId
        If gimFaciCode.GetDispStr <> "" Then subAnd2 strWhere, "FaciCode IN (" & gimFaciCode.GetQueryCode & ")"
        intCount = GetRsValue("Select Count(*) From " & GetOwner & "SO004 Where " & strWhere & " And PRDate Is Null And InstDate is not null")
        chkHaveFaciInst = intCount > 0
End Function

Private Sub NewRcd()
    On Error Resume Next
        txtBillNo.Text = ""
        Call OpenData
        Set ggrData.Recordset = rsArray
        ggrData.Refresh
        
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF2 Then
            Call cmdOK_Click
        ElseIf Shift = 2 And KeyCode = vbKeyF Then
            If txtSQL.Visible Then
                txtSQL.Visible = False
            Else
                txtSQL.Move 0, 0, Me.Width, Me.Height / 2
                txtSQL.Visible = True
            End If
        End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
        If Alfa2 Then
            Call GetGlobal
        End If
        Call OpenData
        Call subGil
        Call subGim
        Call subGrd
        Call DefaultValue
End Sub

Private Sub subGrd()
    On Error GoTo chkErr
        Dim mFlds As New prjGiGridR.GiGridFlds
            With mFlds
                .Add "Rowno", , , , , "序號", vbLeftJustify
                .Add "CustId", , , , , "客戶編號", vbLeftJustify
                .Add "CustName", , , , , "客戶姓名", vbLeftJustify
                .Add "Tel1", , , , , "電話1         ", vbLeftJustify
                .Add "Tel2", , , , , "電話2         ", vbLeftJustify
                .Add "Tel3", , , , , "電話3         ", vbLeftJustify
                .Add "StartDateA", giControlTypeDate, , , , "帳單起始日", vbLeftJustify
                .Add "StopDateA", giControlTypeDate, , , , "帳單截止日", vbLeftJustify
                .Add "StartDateB", giControlTypeDate, , , , "下收起始日", vbLeftJustify
                .Add "StopDateB", giControlTypeDate, , , , "下收截止日", vbLeftJustify
                .Add "ClctDate", giControlTypeDate, , , , "次收費日", vbLeftJustify
                '.Add "CompCode", , , , , "公司別代碼", vbLeftJustify
                '.Add "InstAddrNo", , , , , "裝機址地編號", vbLeftJustify
                .Add "CustStatusName", , , , , "客戶狀態", vbLeftJustify
                .Add "WipName3", , , , False, "派工狀態三", vbLeftJustify
                .Add "WipName1", , , , False, "派工狀態一", vbLeftJustify
                .Add "WipName2", , , , False, "派工狀態二", vbLeftJustify
                .Add "InstAddress", , , , , "裝機地址                             ", vbLeftJustify
                .Add "Balance", , , , , "欠費狀態", vbLeftJustify
                .Add "BillNO", , , , , "單據編號" & Space(12), vbLeftJustify
                .Add "PrtSNo", , , , , "印單序號       ", vbLeftJustify
                .Add "MediaBillNo", , , , , "媒體單號       ", vbLeftJustify
            End With
        ggrData.AllFields = mFlds
        Set ggrData.Recordset = rsArray
        ggrData.SetHead
        ggrData.Refresh
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGrd"
End Sub

Private Sub subGil()
    On Error GoTo chkErr
        '預約派工時間：當時時間
        gdtPRTime.Text = RightNow
        '公司別
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 1, 12, GetCompCodeFilter(gilCompCode)
        '服務類別：
        SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
        '停拆機原因：
        SetgiList gilReasonCode, "CodeNo", "Description", "CD014", , , , , 3, 12, , True
        '停拆機類別：
        SetgiList gilPRCode, "CodeNo", "Description", "CD007", , , , , 3, 12, "Where RefNo =2 ", True
        '工程組別：
        SetgiList gilGroupCode, "CodeNo", "Description", "CD003", , , , , 3, 12, , True
        '工作人員1
        SetgiList gilWorkerEn1, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True
        '工作人員2
        SetgiList gilWorkerEn2, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True
        '工作人員3
        SetgiList gilWorkerEn3, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True
        '結案短收原因：
        SetgiList gilSTCode, "CodeNo", "Description", "CD016", , , , , 3, 12, "Where RefNo = 1", True
        '未收原因
        SetgiList gilUCCode, "CodeNo", "Description", "CD013", , , , , 3, 12, , True
        '作廢原因
        SetgiList gilCancelCode, "CodeNo", "Description", "CD051", , , , , 1, 12
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGil"
End Sub

Private Sub subGim()
    On Error GoTo chkErr
        '大樓
        SetgiMulti gmsMduId, "MduId", "Name", "SO017", "編號", "名稱"
        '客戶狀態
        SetgiMulti gmsStatusCode, "CodeNo", "Description", "CD035", "客戶代碼", "客戶狀態"
        SetgiMulti gimUCCode, "CodeNo", "Description", "CD013", "代碼", "名稱"
        SetgiMulti gimClassCode, "CodeNo", "Description", "CD004", "代碼", "名稱"
        SetgiMulti gimCMCode, "CodeNo", "Description", "CD031", "代碼", "名稱"
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱"
        SetgiMulti gimFaciCode, "CodeNo", "Description", "CD022", "代碼", "名稱"
        SetgiMulti gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱"
        SetgiMulti gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱"
        SetgiMultiAddItem gimBillType, "B,I,P,M,T", "收費單,裝機單,停拆移機單,維修單,臨時收費單"
        '********************************************************************************************
        '#3474 增加客戶類別一 By Kin 2008/05/12
        SetgiMulti gmsClassCode1, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱"
        '********************************************************************************************
        SetgiMulti gimMduId2, "MduId", "Name", "SO017", "編號", "名稱"
        '客戶狀態
        SetgiMulti gimCustStatusCode2, "CodeNo", "Description", "CD035", "客戶代碼", "客戶狀態"
        SetgiMulti gimWipCode3, "CodeNo", "Description", "CD036", "派工狀態代碼", "派工狀態名稱", "Where CodeNo >= 11 And CodeNo < 20 "
        gmsStatusCode.SetQueryCode "1"
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGim"
End Sub

Private Sub DefaultValue()
    On Error GoTo chkErr
    Dim intFlowId As Integer
        '公司別
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        '收費單一併結案：Checked
        optNothing = True
        '結案日期：今日
        'gdaClsDate.SetValue Format(Date, "YYYYMMDD")
        '範圍:1
        cboMduRange.ListIndex = 0
        '單據序,單據編號,客戶編號,姓名,客戶類別,客戶狀態,派工一,派工二,派工三,欠費狀態
        stData.Tab = 0
        lblCount = 0
'        intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043") & "")
'        If intPara23 = 1 Then
'            lblBillNo = "印單序號"
'            txtBillNo.MaxLength = 12
'        ElseIf intPara23 = 2 Then
'            lblBillNo = "媒體單號"
'            txtBillNo.MaxLength = 11
'        Else
'            lblBillNo = "單據編號"
'            txtBillNo.MaxLength = 15
'        End If
        intFlowId = GetSystemParaItem("FlowId", gilCompCode.GetCodeNo, "", "SO041")
        chkPREBox.Visible = (intFlowId = 0)
        lblResvPRTime.Visible = (intFlowId = 0)
        gdaEBoxPRTime.Visible = (intFlowId = 0)
        txtBillNo.Text = ""
        '#3254 判斷是否有啟用"是否新增至服務申告" By Kin 2007/07/18
        blnAddCodSvr = gcnGi.Execute("Select AddCodSvr From " & GetOwner & "SO041 Where CompCode=" & gilCompCode.GetCodeNo)(0)
        optPayKind(0).Caption = gcnGi.Execute("Select Description From " & GetOwner & "CD112 Where CodeNo = 0")(0) & ""
        optPayKind(1).Caption = gcnGi.Execute("Select Description From " & GetOwner & "CD112 Where CodeNo = 1")(0) & ""
        If Not GetPaynowFlag Then
            optPayKind(0).Enabled = False
            optPayKind(1).Enabled = False
        End If
        
    Exit Sub
chkErr:
    ErrSub Me.Name, "DefaultValue"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        Call Close3Recordset(rsArray)
        Call Close3Recordset(rsSo001)
        Call Close3Recordset(rsSO014)
        Call Close3Recordset(rsCd007)
        file.Close
        Set file = Nothing
        Set FSO = Nothing
        cnMDB.Close
        Set cnMDB = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub gdaClsDate_Validate(Cancel As Boolean)
    '若有值，則需小於等於今日，否則錯誤："日期不可大於今日"
    If gdaClsDate.Text = "" Then Exit Sub
    If Not IsDate(gdaClsDate.Text) Then MsgBox "日期不合法！", vbExclamation, gimsgPrompt: Cancel = True: Exit Sub
    If CDate(gdaClsDate.GetValue(True)) > RightDate Then MsgBox "日期不可大於今日！", vbExclamation, gimsgPrompt: Cancel = True: Exit Sub

End Sub

Private Sub gdtPRTime_Validate(Cancel As Boolean)
    If gdtPRTime.Text = "" Then MsgBox "此欄位不得為空白！", vbExclamation, gimsgPrompt: Cancel = True: Exit Sub
    If CDate(Format(gdtPRTime.Text, "YYYY/MM/DD")) < RightDate Then MsgBox "派工預約日期不可小於今日！", vbExclamation, gimsgPrompt: Cancel = True: Exit Sub
End Sub

Private Sub gilCancelCode_GotFocus()
    optCancel.Value = True
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3400", "SO3420") Then Exit Sub
        Call subGil
        Call subGim
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.SetCodeNo ""
        gilServiceType.Query_Description
        gilServiceType.ListIndex = 1
        Call GiMultiFilter(gmsMduId, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimServCode, , gilCompCode.GetCodeNo)
        Call DefaultValue
        Call NewRcd
End Sub

Private Sub gilPRCode_Validate(Cancel As Boolean)
    If gilPRCode.GetDescription <> "" Then
        Dim rsTmp As New ADODB.Recordset
        If Not GetRS(rsTmp, "Select RefNo From " & GetOwner & "CD007 Where CodeNo=" & gilPRCode.GetCodeNo, gcnGi, adUseClient) Then Exit Sub
        If Not rsTmp.EOF Then
            intRefNo = rsTmp("RefNO").Value
        End If
        Set rsTmp = Nothing
    End If
End Sub

Private Sub gilServiceType_Change()
    On Error Resume Next
        Call GiListFilter(gilPRCode, gilServiceType.GetCodeNo)
        Call GiListFilter(gilReasonCode, gilServiceType.GetCodeNo)
        Call GiListFilter(gilSTCode, gilServiceType.GetCodeNo)
        Call GiMultiFilter(gimCitemCode, gilServiceType.GetCodeNo)
        Call GiMultiFilter(gimFaciCode, gilServiceType.GetCodeNo)
        gilPRCode.Filter = gilPRCode.Filter & IIf(gilPRCode.Filter = "", "Where ", "And ") & " RefNo =2 "
        gilSTCode.Filter = gilSTCode.Filter & IIf(gilSTCode.Filter = "", "Where ", "And ") & " RefNo =1 "
        '#3254 測試不OK，當服務別改變時，單據類別也要跟著變動 By Kin 2007/09/19
        intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043") & "")
        If intPara23 = 1 Then
            lblBillNo = "印單序號"
            txtBillNo.MaxLength = 12
        ElseIf intPara23 = 2 Then
            lblBillNo = "媒體單號"
            txtBillNo.MaxLength = 11
        Else
            lblBillNo = "單據編號"
            txtBillNo.MaxLength = 15
        End If
        '#5411 增加指定設備,C服務才能選取 By Kin 2009/12/10
        If gilServiceType.GetCodeNo = "C" Then
            chkSetFaci.Enabled = True
        Else
            chkSetFaci.Value = 0
            chkSetFaci.Enabled = False
        End If
        
End Sub

Private Sub gilSTCode_GotFocus()
    optSTCode.Value = True
End Sub

Private Sub optCancel_Click()
    On Error Resume Next
        gilCancelCode.SetFocus
End Sub

Private Sub optNothing_Click()
    On Error Resume Next
        If optNothing Then
            fraBillClose.Enabled = False
            gilUCCode.Enabled = False
        End If
End Sub

Private Sub optPayKind_Click(Index As Integer)
  On Error Resume Next
    If Index = 1 Then
        optRealStopDate.Enabled = True
        optRealStopDate.Value = True
    Else
        optRealStopDate.Enabled = False
        optNothing.Value = True
    End If
End Sub

Private Sub optSTCode_Click()
    On Error Resume Next
        gilSTCode.SetFocus
End Sub

Private Sub chkUseOldBillNo_GotFocus()
    If blnFlag Then txtBillNo.SetFocus
    blnFlag = False
End Sub

Private Sub optUCCode_Click()
    On Error Resume Next
        If optUCCode Then
            fraBillClose.Enabled = False
            gilUCCode.Enabled = True
        Else
            gilUCCode.Clear
        End If
End Sub

Private Sub stData_Click(PreviousTab As Integer)
  On Error Resume Next
    If stData.Tab = 2 Then
        cmdInExcel.Enabled = True
        cmdInExcel2.Enabled = True
        If optPayKind(1).Value Then
            optRealStopDate.Enabled = True
            optRealStopDate.Value = True
        End If
    Else
        cmdInExcel.Enabled = False
        cmdInExcel2.Enabled = False
        If optRealStopDate.Value Then
            optNothing.Value = True
        End If
        optRealStopDate.Enabled = False
        
    End If
End Sub

Private Sub txtBillNo_Change()
    If Len(txtBillNo) = 11 And chkUseOldBillNo.Value = 1 Then
        Call txtBillNo_Validate(False)
    End If
End Sub

Private Sub txtBillNo_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
End Sub

Private Sub txtBillNo_KeyPress(keyAscii As Integer)
    keyAscii = Asc(UCase(Chr(keyAscii)))
End Sub

Private Sub txtBillNo_Validate(Cancel As Boolean)
    On Error GoTo chkErr
    Dim strFields As String
    Dim strCustid As String, strSQL As String
    Dim rsSo3240 As New ADODB.Recordset
    Static intCount As Integer
    Dim rsSO033 As New ADODB.Recordset
    Dim rsSO003 As New ADODB.Recordset
    Dim strCitemCode As String
    Dim intSNoType As Integer
    Dim strBillField As String
    
        '此欄位無值：則無以下動作！
        If txtBillNo.Text = "" Then Exit Sub
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Sub
        If Not MustExist(gilServiceType, 2, "服務類別") Then Exit Sub
        If chkPRFaci.Value = 1 And gimFaciCode.GetQryStr = "" Then gimFaciCode.SetFocus: MsgBox "設備項目為必要欄位!!", vbExclamation, gimsgPrompt: Exit Sub
        
        If Len(txtBillNo) = 11 And chkUseOldBillNo.Value = 1 Then
            If InStr("BT", Mid(txtBillNo, 5, 1)) > 0 Then
                txtBillNo = GetADString(Left(txtBillNo, 4), False) & Mid(txtBillNo, 5, 1) & "C" & Format(Mid(txtBillNo, 6), "0000000")
            End If
        End If
        
        If Not ChkBillOk Then Cancel = True: Exit Sub
        
        '根據單據編號，至應收資料檔檢查是否有未收資料，若無未收資料，則錯誤！顯示："無此單據，或此單據已登錄過，請核對！"，且並無以下動作！
        '#3474 多增加GUINO(發票號碼欄位) By Kin 2008/05/09
        '#4161 增加FaciSeqNo,FaciSNO以便可以新增至SO180 By Kin 2008/10/30
        strSQL = "Select ServiceType,RealStartDate,RealStopDate,CitemCode,CitemName,CustId,BillNo," & _
                 "Item,PrtSNo,MediaBillNo,SeqNo,ClassCode,GUINO,FaciSeqNo,FaciSNo "
        Select Case Len(txtBillNo.Text)
            Case 15
                '#3430 增加判斷參數，避免只判斷輸入長度，而判斷其它種類的單據 By Kin 2007/08/23
                If intPara23 <> 0 Then Exit Sub
                strSQL = strSQL & " From " & GetOwner & "So033 Where BillNo='" & txtBillNo.Text & "' And UCCode is not Null And CompCode = " & gilCompCode.GetCodeNo & " And ServiceType = '" & gilServiceType.GetCodeNo & "'"
                If Not GetRS(rsSO033, strSQL, gcnGi) Then MsgBox "連線失敗！請重新操作！", vbExclamation, gimsgPrompt: Exit Sub
                If rsSO033.RecordCount = 0 Then
                    MsgBox "無此單據，或此單據已登錄過，請核對！", vbExclamation, gimsgPrompt
                    Cancel = True
                    Exit Sub
                End If
                intSNoType = 0
                strBillField = "BillNo"
            Case 12
                '‧  若長度為12碼, 則為印單序號: (格式YYYYMM999999)
                'SQL = "select <…> from SO033 where PrtSNo='<單據編號>' and RealDate is null and UCCode is not null "
                'lngBillNo = Right(txtBillNo.Text, 8)
                '2001/11/12 Janis 說要改的
                '#3430 增加判斷參數，避免只判斷輸入長度，而判斷其它種類的單據 By Kin 2007/08/23
                If intPara23 <> 1 Then Exit Sub
                strSQL = strSQL & " From " & GetOwner & "SO033 where PrtSNo='" & UCase(txtBillNo.Text) & "' And CompCode = " & gilCompCode.GetCodeNo & " and UCCode is not null "
                If Not GetRS(rsSO033, strSQL, gcnGi) Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！"
                If rsSO033.EOF Then
                    If intPara23 = 1 Then
                        MsgBox "無此單據編號或此單據已登錄過，請核對！", vbExclamation, "訊息！"
                        Cancel = True
                        'Call ClearRcd
                    End If
                    Exit Sub
                End If
'                '*****************************************************************
'                '#3474 如果已有發票號碼要Show出訊息 By Kin 2008/05/09
'                If Not IsNull(rsSO033("GUINO")) Then
'                    MsgBox "此單據編號已有發票號碼！", vbExclamation, "訊息"
'                    Cancel = True
'                    Exit Sub
'                End If
'                '*****************************************************************
                intSNoType = 1
                strBillField = "PrtSNo"
            Case 11
                '‧  若長度為11碼, 則為媒體單號: (格式YYMM9999999)
                '#3430 增加判斷參數，避免只判斷輸入長度，而判斷其它種類的單據 By Kin 2007/08/23
                If intPara23 <> 2 Then Exit Sub
                strSQL = strSQL & " From " & GetOwner & "SO033 where MediaBillNo='" & UCase(txtBillNo.Text) & "' And CompCode = " & gilCompCode.GetCodeNo & " and UCCode is not null "
                If Not GetRS(rsSO033, strSQL, gcnGi) Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！"
                If rsSO033.EOF Then
                    If intPara23 = 2 Then
                        MsgBox "無此單據編號或此單據已登錄過，請核對！", vbExclamation, "訊息！"
                        Cancel = True
                    End If
                    Exit Sub
                'Else
    '                '若輸入值非B單格式 (而是臨時單或印單序號格式), 則根據取出之收費資料第一筆的客戶編號, 至客戶基本資料檔取
                '    Call GetCustname(rsSo3311B("Custid").Value, strCustName, intCustStatus, strCustStatusname, rsSo3311B("ServiceType"))
                End If
                intSNoType = 2
                strBillField = "MediaBillNo"
        End Select
        '*****************************************************************
        '#3474 如果已有發票號碼要Show出訊息 By Kin 2008/05/09
'        If Not IsNull(rsSO033("GUINO")) Then
'            MsgBox "此單據編號已有發票號碼！", vbExclamation, "訊息"
'            Cancel = True
'            Exit Sub
'        End If
        '*****************************************************************
        '#3254 檢查Uccode參考號是否為3 By Kin 2007/07/16
        If Not chkUCCode Then Exit Sub
        If IstmpRsDup(rsArray, strBillField, Trim(txtBillNo.Text)) Then blnFlag = True: Exit Sub
        
        strCustid = rsSO033("CustId")
        
        strFields = "Select B.WipName1,B.WipName2,A.ClassName1,A.Balance,B.wipName3,A.ServCode,A.Tel1,A.Tel2,A.Tel3,A.CompCode,A.InstAddrNo,A.InstAddress,B.CustStatusCode,B.CustStatusName,A.CustName,B.ServiceType"
        strSQL = strFields & " From " & GetOwner & "So001 A," & GetOwner & "So002 B Where A.CustId = B.CustId And A.Custid=" & strCustid & " And B.ServiceType = '" & rsSO033("ServiceType") & "'"
        If Not GetRS(rsSo3240, strSQL, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then MsgBox "連線失敗！請重新操作！", vbExclamation, gimsgPrompt: Exit Sub
        
        '若無資料，則錯誤！:顯示"無此單據之客戶資料"並無以下動作
        If rsSo3240.EOF Then
            MsgBox "無此單據之客戶資料", vbExclamation, gimsgPrompt
            Cancel = True
            Exit Sub
        End If
        
        '有資料！若客戶狀態<CustStatusCode> 是已拆機(3)或拆機中(6) ：顯示"該客戶為拆機戶，不再產生拆機單"且並無以下動作
        If rsSo3240("CustStatusCode").Value = 3 Or rsSo3240("CustStatusCode").Value = 6 Then
            MsgBox "該客戶為已拆戶或拆機中，不再產生拆機單！", vbExclamation, gimsgPrompt
            Cancel = True
            Exit Sub
        End If
        If chkHaveFaciInst(rsSO033("CustId")) Then
            If MsgBox("該客戶檢核待拆設備項目未拆,CATV是否要拆機?", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then
                Cancel = True
                Exit Sub
            End If
        End If
        '94/08/01 Jacky Debby 提1696
        lblClassName = GetRsValue("Select Description From " & GetOwner & "CD004 Where CodeNo = " & rsSO033("ClassCode")) & ""
        
        Dim intPrtSNo As Long
        '單據序號加1
        If rsArray.EOF Then
            intPrtSNo = 1
        Else
            If IsNull(rsArray("RowNO").Value) Then
                intPrtSNo = 1
            Else
                intPrtSNo = rsArray("RowNO").Value + 1
            End If
        End If
        '將取出的資料及單據序號PrtSNo(3)顯示於Grid中
        rsSO033.MoveFirst
        While Not rsSO033.EOF
        'If Not rsSo033.EOF Then
            If strCitemCode <> rsSO033("CitemCode") Then
                strSQL = "Select StartDate,StopDate,ClctDate From " & GetOwner & "So003 Where CustId = " & strCustid & " And CitemCode = " & rsSO033("CitemCode")
                If Not GetRS(rsSO003, strSQL, gcnGi) Then MsgBox "連線失敗！請重新操作！", vbExclamation, gimsgPrompt: Exit Sub
                If Not rsSO003.EOF Then
                    If rsSO003("ClctDate") > RightDate Then
                        If MsgBox("該客戶週期性收費之次收費日(" & GetDT(rsSO003("ClctDate"), GiDate) & ") 未到期,是否確定派拆??", vbQuestion + vbYesNo + vbDefaultButton2, gimsgPrompt) = vbNo Then Exit Sub
                    End If
                End If
            End If
            With rsArray
                .AddNew
                .Fields("Custid").Value = Val(strCustid)
                .Fields("Billno").Value = rsSO033("BillNo")
                .Fields("PrtSNO").Value = rsSO033("PrtSNo")
                .Fields("MediaBillno").Value = rsSO033("MediaBillNo")
                .Fields("CitemCode").Value = rsSO033("CitemCode")
                .Fields("CitemName").Value = rsSO033("CitemName")
                .Fields("SeqNo").Value = rsSO033("SeqNo")
                
                .Fields("WipName1").Value = GetFieldValue(rsSo3240, "WipName1")
                .Fields("WipName2").Value = GetFieldValue(rsSo3240, "WipName2")
                .Fields("WipName3").Value = GetFieldValue(rsSo3240, "WipName3")
                .Fields("ClassName1").Value = GetFieldValue(rsSo3240, "ClassName1")
                .Fields("Balance").Value = GetFieldValue(rsSo3240, "Balance")
                .Fields("ServCode").Value = GetFieldValue(rsSo3240, "ServCode")
                .Fields("Tel1").Value = GetFieldValue(rsSo3240, "Tel1")
                .Fields("CompCode").Value = GetFieldValue(rsSo3240, "CompCode")
                .Fields("InstAddrNO").Value = GetFieldValue(rsSo3240, "InstAddrNo")
                .Fields("InstAddress").Value = GetFieldValue(rsSo3240, "InstAddress")
                .Fields("CustStatusCode").Value = GetFieldValue(rsSo3240, "CustStatusCode")
                .Fields("CustStatusName").Value = GetFieldValue(rsSo3240, "CustStatusName")
                .Fields("CustName").Value = GetFieldValue(rsSo3240, "CustName")
                .Fields("ServiceType").Value = NoZero(rsSO033("ServiceType"))
                .Fields("StartDateA").Value = NoZero(rsSO033("RealStartDate"))
                .Fields("StopDateA").Value = NoZero(rsSO033("RealStopDate"))
                .Fields("RealStartDate").Value = NoZero(rsSO033("RealStartDate"))
                .Fields("RealStopDate").Value = NoZero(rsSO033("RealStopDate"))
                If rsSO003.RecordCount > 0 Then
                    .Fields("StartDateB").Value = NoZero(rsSO003("StartDate"))
                    .Fields("StopDateB").Value = NoZero(rsSO003("StopDate"))
                    .Fields("ClctDate").Value = NoZero(rsSO003("ClctDate"))
                End If
                If Not IsNull(rsSO033("GUINo")) Then
                    .Fields("GUINo") = rsSO033("GUINo")
                End If
                '*************************************************************
                '#4161 增加FaciSeqNo,FaciSNO,Item欄位,以便可以填入SO180 By Kin 2008/10/30
                If Not IsNull(rsSO033("FaciSeqNo")) Then
                    .Fields("FaciSeqNo") = rsSO033("FaciSeqNo")
                End If
                If Not IsNull(rsSO033("FaciSNo")) Then
                    .Fields("FaciSNo") = rsSO033("FaciSNo")
                End If
                .Fields("Item") = rsSO033("Item")
                '**************************************************************
                .Fields("RowNO") = intPrtSNo
                .Update
                strCitemCode = rsSO033("CitemCode") & ""
            End With
            rsSO033.MoveNext
        'End If
        Wend
        lblCount = rsArray.RecordCount
        On Error Resume Next
        rsArray.Requery
        On Error GoTo chkErr
        Set ggrData.Recordset = rsArray
        ggrData.Refresh
        blnFlag = True
        Call Close3Recordset(rsSO033)
        Call Close3Recordset(rsSO003)
        Call Close3Recordset(rsSo3240)
         
    Exit Sub
chkErr:
    ErrSub Me.Name, "txtBillNo_Validate"
End Sub

Private Function ChkBillOk() As Boolean
    On Error GoTo chkErr
        '格式檢查：長度需為15碼，且第7碼為"B"，否則錯誤："單據編號錯誤！"，無以下動作！
        If Not ((Len(txtBillNo.Text) = 15 And InStr(1, txtBillNo.Text, "B") = 7) Or Len(txtBillNo.Text) = 12 Or Len(txtBillNo.Text) = 11) Then
            MsgBox "單據編號錯誤！", vbExclamation, gimsgPrompt
            Exit Function
        End If
        ChkBillOk = True
    Exit Function
chkErr:
    ErrSub Me.Name, "ChkBillOk"
End Function


Private Sub OpenLogFile()
    On Error GoTo chkErr
        
        strFileName = ReadGICMIS1("ErrLogPath")
        strFileName = strFileName & IIf(Right(strFileName, 1) = "\", "", "\") & "SO3420ALog.txt"
    On Error Resume Next
        file.Close
        Set file = Nothing
        Kill strFileName
    On Error GoTo chkErr
        Set file = FSO.CreateTextFile(strFileName, True)
    Exit Sub
chkErr:
    ErrSub Me.Name, "OpenLogFile"
End Sub

Private Sub txtPRNote_GotFocus()
    txtPRNote.SelStart = 0
    txtPRNote.SelLength = Len(txtPRNote.Text)
    txtPRNote.IMEMode = 1
End Sub

Private Sub txtPRNote_LostFocus()
    txtPRNote.IMEMode = 2
End Sub

Private Function OpenData() As Boolean
    On Error GoTo chkErr
    Dim strPath As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
        strPath = ReadGICMIS1("TmpMDBPath")
        If cnMDB.State = 1 Then cnMDB.Close
'        cnMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & IIf(Right(strPath, 1) = "\", strPath, strPath & "\") & "Tmp1111.MDB" & ";Persist Security Info=False"
        cnMDB.Open "Provider=" & GetOleDbProvider & ";Data Source=" & IIf(Right(strPath, 1) = "\", strPath, strPath & "\") & "Tmp1111.MDB" & ";Persist Security Info=False"
        On Error Resume Next
        '#4161 要填入SO180欄位,所以多增加FaciSeqNo,FaciSNO,Item欄位 By Kin 2008/10/30
        strSQL = "Select 1 RowNo,B.PrtSNo,A.CustId,D.WipName1,D.WipName2,A.ClassName1,A.Balance,D.wipName3,A.ServCode,A.Tel1,A.Tel2,A.Tel3,A.CompCode,A.InstAddrNo,A.InstAddress,D.CustStatusCode,D.CustStatusName,A.CustName,A.ServiceType," & _
            "B.BillNo,B.RealStartDate StartDateA,B.RealStopDate StopDateA,C.StartDate StartDateB,C.StopDate StopDateB,C.ClctDate,B.MediaBillNo,B.CitemCode,B.CitemName,B.SeqNo,0 as Type,B.RealStartDate,B.RealStopDate,B.GUINo,B.FaciSeqNo,B.FaciSNo,B.Item " & _
            " From " & GetOwner & "So001 A," & GetOwner & "SO033 B," & GetOwner & "So003 C," & GetOwner & "So002 D " & _
            "Where B.BillNo = '' And B.Item = -1 And A.CustId= B.CustId And B.CustId = C.CustId And B.CustId = D.CustId"
        If Not GetRS(rsArray, strSQL) Then Exit Function
        If Not CreateMDBTable(rsArray, "So3420A", cnMDB) Then Exit Function
        
        'If Not GetRS(rsTmp, strSQL) Then Exit Function
        'If Not CreateMDBTable(rsTmp, "So3420ALog", cnMDB) Then Exit Function
        
        If Not GetRS(rsArray, "Select * From So3420A Order by RowNo Desc", cnMDB, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        Call CloseRecordset(rsTmp)
        OpenData = True
    Exit Function
chkErr:
    ErrSub Me.Name, "OpenData"
End Function

Private Sub DelrsArray()
    If rsArray.RecordCount > 0 Then
        rsArray.MoveFirst
        While Not rsArray.EOF
            rsArray.Delete
            rsArray.MoveNext
        Wend
    End If
End Sub

Private Function IsDataOk() As Boolean
    '必要欄位：停拆機原因、停拆機類別、範圍、預約派工時間
    '若收費單一併結案，則結案短收原因，結帳日期為必要
    On Error GoTo chkErr
        IsDataOk = False
        If Not ChkDTok Then Exit Function
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        '就Servicetype
        'If stData.Tab <> 0 Then
            If Not MustExist(gilServiceType, 2, "服務類別") Then Exit Function
        'End If
        
        '就拆機原因
        'do not check the condition when has imported excel by Kin 2016/10/17
        If stData.Tab = 2 Then
             If Not importExcelData Is Nothing Then
                If importExcelData.Count > 0 Then
                
                Else
                    If Not MustExist(gilReasonCode, 2, "拆機原因") Then Exit Function
                End If
             Else
                If Not MustExist(gilReasonCode, 2, "拆機原因") Then Exit Function
             End If
        Else
            If Not MustExist(gilReasonCode, 2, "拆機原因") Then Exit Function
        End If
        
        '停拆機類別
        If Not MustExist(gilPRCode, 2, "停拆機類別") Then Exit Function
        '預約派工時間
        If Not MustExist(gdtPRTime, 1, "預約派工時間") Then Exit Function
        If CDate(gdtPRTime.GetDate(True)) < RightDate Then
            MsgBox "預約日期不得大於今日", vbExclamation, gimsgPrompt
        End If
        
        '若收費單一併結案，則結案短收原因，結帳日期為必要
        If optClsBill.Value = True Then
            If optSTCode.Value Then If Not MustExist(gilSTCode, 2, "短收原因") Then Exit Function
            If optCancel.Value Then If Not MustExist(gilCancelCode, 2, "作廢原因") Then Exit Function
            If Not MustExist(gdaClsDate, 1, "結帳日期") Then Exit Function
            If CDate(gdaClsDate.GetValue(True)) > RightDate Then
                MsgBox "結帳日期不得大於今日", vbExclamation, gimsgPrompt
                Exit Function
            End If
        ElseIf optUCCode Then
            If Not MustExist(gilUCCode, 2, "未收原因") Then Exit Function
        End If
        
        If stData.Tab = 1 Then
            '大樓編號需有值
            If gmsMduId.GetDispStr = "" Then Call MsgMustBe("大樓編號"): gmsMduId.SetFocus: Exit Function
        ElseIf stData.Tab = 2 Then
            If Not MustExist(gdaShouldDate1, 1, "應收日期起始日") Then Exit Function
            If Not MustExist(gdaShouldDate2, 1, "應收日期起始日") Then Exit Function
            If gimCitemCode.GetDispStr = "" Then MsgMustBe "收費項目": gimCitemCode.SetFocus: Exit Function
            If Not importExcelData Is Nothing Then
                If importExcelData.Count > 0 Then
                
                Else
                     If gimUCCode.GetDispStr = "" Then MsgMustBe "未收原因": gimUCCode.SetFocus: Exit Function
                End If
              Else
                If gimUCCode.GetDispStr = "" Then MsgMustBe "未收原因": gimUCCode.SetFocus: Exit Function
             End If
           
        End If
        '#3254 如果有啟動"是否新增至服務申告" 則Show出frmSO3420B讓User選擇申告原因 By Kin 2007/07/18
        If blnAddCodSvr Then
            With frmSO3420B
                .uServiceType = gilServiceType.GetCodeNo
                .Show 1
            End With
            If Not blnIsOK Then
                MsgBox "無設定服務申告!!請重新操作!", vbInformation, "警告訊息"
                IsDataOk = False
                Exit Function
            End If
            DoEvents
            
        End If
        IsDataOk = True
    Exit Function
chkErr:
    ErrSub Me.Name, "IsDataOk"
End Function

Private Function subCloseBill(lngCustId As Long, strServiceType As String, _
    strChoose033 As String, Optional ByVal aBillNo As String = "") As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strWhere As String
    Dim aUCCode As String
    Dim aUCName As String
    Dim aBillWhere As String
        aUCCode = gilUCCode.GetCodeNo
        aUCName = gilUCCode.GetDescription
        aBillWhere = " And 1 = 1 "
        'update uccode datafield According to excel data By Kin 2016/09/29
        If Len(aBillNo) > 0 Then
            If Not importExcelData Is Nothing Then
                If importExcelData.Count > 0 Then
                    If importExcelData.Exists(aBillNo) Then
                        Dim aData As String
                        aData = importExcelData(aBillNo)
                        aUCCode = Split(aData, ",")(4)
                        aUCName = Split(aData, ",")(5)
                        aBillWhere = " And BillNo = '" & aBillNo & "' "
                    End If
                End If
            End If
        End If
        '8.8若<收費單一併結案>，則將應收資料檔(SO033)中，該單據所指的未收資料，一併做短收結案：
        strSQL = "Update " & GetOwner & "So033 A Set Updtime='" & strDTUpdTime & "',UpdEn='" & garyGi(1) & "'"
        If optClsBill Then
            If optSTCode Then
                strSQL = strSQL & ",STCode=" & gilSTCode.GetCodeNo & ",STName='" & gilSTCode.GetDescription & _
                    "',RealDate=To_Date('" & gdaClsDate.GetValue & "','YYYYMMDD'),UCCode=Null,UCName=Null "
            Else
                strSQL = strSQL & ",CancelCode=" & gilCancelCode.GetCodeNo & ",CancelName='" & gilCancelCode.GetDescription & _
                    "',STCode=null,STName=null,CancelFlag=1,RealAmt = 0 " & _
                    ",RealDate=To_Date('" & gdaClsDate.GetValue & "','YYYYMMDD'),UCCode=Null,UCName=Null "
            End If
        Else
            strSQL = strSQL & ",UCCode=" & aUCCode & ",UCName='" & aUCName & "'"
        End If
        
        If stData.Tab = 0 Then
            If Not GetRS(rsTmp, "Select * From So3420A Where CustId = " & lngCustId, cnMDB) Then Exit Function
            Do While Not rsTmp.EOF
                strWhere = " And UCCode Is Not Null And ServiceType = '" & strServiceType & "' "
                If intPara23 = 2 Then
                    strWhere = " Where MediaBillNo = '" & rsTmp("MediaBillNo") & "'" & strWhere
                ElseIf intPara23 = 1 Then
                    strWhere = " Where PrtSNo = '" & rsTmp("PrtSNo") & "'" & strWhere
                Else
                    strWhere = " Where BillNo = '" & rsTmp("BillNo") & "'" & strWhere
                End If
                gcnGi.Execute strSQL & strWhere
                rsTmp.MoveNext
                DoEvents
            Loop
        Else
            gcnGi.Execute strSQL & " Where CustId = " & lngCustId & _
                            " And UCCode Is Not Null And ServiceType = '" & strServiceType & "'" & strChoose033 & aBillWhere
                            
        End If
        subCloseBill = True
    Exit Function
chkErr:
    ErrSub Me.Name, "subCloseBill"
End Function

Private Function subDeleteIntro(lngCustId As Long) As Boolean
    On Error GoTo chkErr
        '8.9 若<一併刪除介紹人及介紹媒體欄位> :則更新該客戶基本資料！
        '介紹人代碼(IntroId)／名稱(IntroName)=介紹人媒介代碼(MediaCode)／名稱(MediaName)=Null
        gcnGi.Execute "Update " & GetOwner & "So001 Set IntroId=Null,IntroName=Null,MediaCode=Null,MediaName=Null Where Custid=" & lngCustId
        subDeleteIntro = True
    Exit Function
chkErr:
    ErrSub Me.Name, "subDeleteIntro"
End Function

Private Function InsertToSO00x(lngMode As giEditModeEnu, lngCustId As Long, _
        intPRCode As Integer, ByRef rsSo00X As ADODB.Recordset, Optional ByVal aBillNo As String = "") As Boolean
    On Error GoTo chkErr
        lngEditMode = lngMode
        
        With rsSo00X
            If lngEditMode = giEditModeInsert Then
                If Not GetRS(rsSo00X, "Select * From " & GetOwner & "So009 Where SNO= ''", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , , 1) Then Exit Function
                Set rsSo00X.ActiveConnection = Nothing
                .AddNew
                .Fields("AcceptTime") = Format(RightNow, "yyyy/mm/dd hh:mm")
                .Fields("AcceptEn") = garyGi(0)
                .Fields("AcceptName") = garyGi(1)
                .Fields("OldAddress") = GetFieldValue(rsSo001, "InstAddress")
                .Fields("OldAddrNo") = GetFieldValue(rsSo001, "InstAddrNO")
                .Fields("Tel1") = GetFieldValue(rsSo001, "Tel1")
                .Fields("CustName") = GetFieldValue(rsSo001, "CustName")
                .Fields("PinCode") = NoZero(rsSo001("PinCode"))
            Else
                If Not GetRS(rsSo00X, "Select * From " & GetOwner & "So009 Where CustId = " & lngCustId & " And PRCode= " & intPRCode & " And SignDate is Null ", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , , 1) Then Exit Function
                Set rsSo00X.ActiveConnection = Nothing
                If rsSo00X.EOF Then InsertToSO00x = True: Exit Function
                .Fields("FinTime") = NoZero(gdtPRTime.GetValue(True))
                .Fields("CallOkTime") = NoZero(.Fields("FinTime"))
                .Fields("SignDate") = NoZero(.Fields("FinTime"))
                .Fields("SignEn") = garyGi(0)
                .Fields("SignName") = garyGi(1)
            End If
            .Fields("CompCode") = GetFieldValue(rsSo001, "CompCode")
            .Fields("ServiceType") = gilServiceType.GetCodeNo
            .Fields("CustID") = lngCustId
            .Fields("GroupCode") = gilGroupCode.GetCodeNo2
            .Fields("GroupName") = gilGroupCode.GetDescription2
            
            .Fields("Note") = NoZero(txtPRNote.Text)
            .Fields("PRCode") = gilPRCode.GetCodeNo2
            .Fields("PRName") = gilPRCode.GetDescription2
            
            .Fields("ReasonCode") = gilReasonCode.GetCodeNo2
            .Fields("ReasonName") = gilReasonCode.GetDescription2
            
            .Fields("ResvTime") = NoZero(gdtPRTime.GetValue(True))
            
            .Fields("ServCode") = GetFieldValue(rsSo001, "ServCode")
            .Fields("StrtCode") = GetFieldValue(rsSO014, "StrtCode")
            .Fields("UpdEn") = garyGi(1)
            .Fields("UpdTime") = GetDTString(strUpdTime)
            .Fields("WorkerEN1") = gilWorkerEn1.GetCodeNo2
            .Fields("WorkerName1") = gilWorkerEn1.GetDescription2
            .Fields("WorkerEN2") = gilWorkerEn2.GetCodeNo2
            .Fields("WorkerName2") = gilWorkerEn2.GetDescription2
            .Fields("WorkerEn3") = gilWorkerEn3.GetCodeNo
            .Fields("WorkerName3") = gilWorkerEn3.GetDescription
            .Fields("WorkUnit") = NoZero(rsCd007("WorkUnit"), True)
            .Fields("NodeNo") = NoZero(rsSO014("NodeNo"))
            .Fields("InstCount") = 0
            If Len(aBillNo) > 0 Then
                If Not importExcelData Is Nothing Then
                    If importExcelData.Count > 0 Then
                        If importExcelData.Exists(aBillNo) Then
                            Dim aData As String
                            aData = importExcelData(aBillNo)
                            .Fields("ReasonCode") = Split(aData, ",")(0)
                            .Fields("ReasonName") = Split(aData, ",")(1)
                            '#7329 CustRunCode is optional,if it's a empty field that don't fill to table By Kin 2016/10/28
                            If Split(aData, ",")(2) <> "X" Then
                                .Fields("CustRunCode") = Split(aData, ",")(2)
                                .Fields("CustRunName") = Split(aData, ",")(3)
                            End If
                        End If
                    End If
                End If
            End If
            If Len(.Fields("FinTime")) > 0 And Val(.Fields("FinUnit") & "") = 0 Then
                .Fields("FinUnit") = NoZero(rsCd007("WorkUnit"), True)
            'Else
            '    .Fields("FinUnit") = NoZero(rsCd007("WorkUnit"), True)
            End If
        End With
        InsertToSO00x = True
    Exit Function
chkErr:
    ErrSub Me.Name, "InsertToSO00x"
End Function

Private Function OpenData2(lngCustId As Long, strServiceType As String, Optional blnUpdate As Boolean = True) As Boolean
    On Error GoTo chkErr
    Dim strSQL As String
        If rsSo001.State = adStateOpen Then
            If rsSo001("CustId") = lngCustId Then blnUpdate = False: OpenData2 = True: Exit Function
        End If
        strSQL = "Select A.CustId,A.CustName,A.MduId,A.ServCode,A.Tel1,A.CompCode,A.InstAddrNo,A.InstAddress,B.CustStatusCode,B.CustStatusName,B.PinCode "
        If Not GetRS(rsSo001, strSQL & " From " & GetOwner & "So001 A," & GetOwner & "So002 B Where A.CustId = B.CustId And B.ServiceType = '" & strServiceType & "' And A.CustId = " & lngCustId) Then Exit Function
        strSQL = "Select StrtCode,NodeNo "
        If Not GetRS(rsSO014, strSQL & " From " & GetOwner & "So014 Where AddrNo =" & rsSo001("InstAddrNo")) Then Exit Function
        OpenData2 = True
    Exit Function
chkErr:
    ErrSub Me.Name, "OpenData2"
End Function

Private Function subPREBox(objNagraSTB As Object, rsData As ADODB.Recordset, _
    rsSO009 As ADODB.Recordset, strUpdTime As String, lngCount As Long) As Boolean
    On Error GoTo chkErr
    Dim rsSTB As New ADODB.Recordset
        If objNagraSTB Is Nothing Then Set objNagraSTB = CreateObject("csNagraSTB3.clsNagraSTB")
        If Not GetRS(rsSTB, "Select A.RowId,A.CompCode,A.FaciSNo,A.SmartCardNo,A.CustId,B.RefNo From " & GetOwner & "SO004 A," & GetOwner & "Cd022 B Where A.FaciCode = B.CodeNo And A.CustId = " & rsData("CustId") & " And A.PRDate is Null And InstDate Is Not Null And B.RefNo In (3,4)") Then Exit Function
        Do While Not rsSTB.EOF
            If rsSTB("RefNo") = 3 Then
                If Not objNagraSTB.CloseSTBProcess(gcnGi, rsSTB("RowId") & "", rsSTB("CompCode") & "", Val(rsData("CustId")), rsSTB("SmartCardNo") & "", rsSTB("FaciSNo"), strUpdTime, garyGi(1), gdaEBoxPRTime.GetValue(True)) Then Exit Function
                '********************************************
                '#3474 計算STB關機筆數 By Kin 2008/05/14
                lngCount = lngCount + 1
                '*******************************************
            End If
            gcnGi.Execute "Update " & GetOwner & "SO004 Set PRDate = " & GetNullString(gdtPRTime.GetDate(True), giDateV) & _
                " ,PREn1 = " & GetNullString(gilWorkerEn1.GetCodeNo) & ",PRName1 =" & GetNullString(gilWorkerEn1.GetDescription) & _
                " ,PREn2 = " & GetNullString(gilWorkerEn2.GetCodeNo) & ",PRName2 =" & GetNullString(gilWorkerEn2.GetDescription) & _
                " ,PRSNo = '" & rsSO009("SNo") & _
                "' Where RowId = '" & rsSTB("RowId") & "'"
            rsSTB.MoveNext
        Loop
        Call CloseRecordset(rsSTB)
        subPREBox = True
    Exit Function
chkErr:
    ErrSub Me.Name, "subPREBox"
End Function

Private Function InsertToLog(rsData As ADODB.Recordset, Optional intErrorCustStatus As csLogError, _
    Optional ByVal strMsg As String) As Boolean
    On Error Resume Next
        If strMsg = "" Then
            Select Case intErrorCustStatus
                Case cs001StatusError
                    strMsg = "該客戶客戶狀態不可派拆機單,客戶編號:" & rsData("CustId") & ", 客戶狀態: " & rsSo001("CustStatusName")
                Case cs009StatusError
                    strMsg = "該客戶派工狀態3為移機不可派拆機單,客戶編號:" & rsData("CustId") & ", 客戶狀態: " & rsSo001("CustStatusName")
                Case cs003RangeError
                    strMsg = "收費資料週期與週期性收費項目重疊;客戶編號" & _
                        ", 單據編號 ,印單序號,媒體單號" & vbCrLf & _
                        rsData("CustId") & ", " & rsData("BillNo") & ", " & rsData("PrtSNo") & ", " & rsData("MediaBillNo")
                Case cn004HaveSTB
                    strMsg = ";客戶編號" & _
                        "該客戶有其他設備未拆,客戶編號:" & rsData("CustId")
                '********************************************************************************************************************
                '#3474 增加發票號碼已有值的Log訊息 By Kin 2008/05/12
                '#4030 增加有發票號碼不處理收費資料但要產生拆機單
                Case csGUINo
                    strMsg = "該客戶待收資料已有發票號碼, 客戶編號:" & rsData("CustId") & ", 單據編號:" & rsData("BillNo") & " 收費資料不處理，但已產生拆機單"
                '********************************************************************************************************************
                Case csOtherError
                    strMsg = "其他錯誤!!" & vbCrLf & _
                        rsData("CustId") & ", " & rsData("BillNo") & ", " & rsData("PrtSNo") & ", " & rsData("MediaBillNo")
            End Select
            
        End If
        file.WriteLine "錯誤原因:" & strMsg
End Function
'#3254 檢查SO033.UcCode是否為CD013.RefNo=3 By Kin 2007/07/16
Private Function chkUCCode() As Boolean
  On Error GoTo chkErr:
    Dim blnReturn As Boolean
    Dim strUccodeSQL As String
    Dim strField As String
    
    blnReturn = True
    If intPara23 = 1 Then
        strField = "PrtSno"
    ElseIf intPara23 = 2 Then
        strField = "MediaBillNo"
    Else
        strField = "BillNo"
    End If
    strUccodeSQL = "Select Count(*) From " & GetOwner & "SO033 Where " & strField & "='" & txtBillNo & "'" & _
                           " And Exists (Select CodeNo From " & GetOwner & "CD013 Where SO033.UcCode=CD013.CodeNo And CD013.RefNo=3)"
    If gcnGi.Execute(strUccodeSQL)(0) > 0 Then
        MsgBox lblBillNo.Caption & "：" & txtBillNo & " 為櫃臺已收，將不產生拆機工單！！", vbInformation, "警告訊息"
        blnReturn = False
    End If
    chkUCCode = blnReturn
    Exit Function
chkErr:
    ErrSub Me.Name, "ChkUccode"
End Function
'#3254 增一筆SO006的Record，並將此Record回傳給clsAlterWip.Action使用，以便增加來電單號 By Kin 2007/07/18
Private Sub ReturnSo006(rs As Recordset)
  On Error GoTo chkErr
    Dim strSeqNo As String
    'If rs.State = adStateOpen Then rs.Close
    
    '#3254 如果"不新增至服務申告"，則將Record設為Nothing，clsAlterWip就不會新增客戶申告記錄檔了 By Kin 2007/07/18
    If Not blnAddCodSvr Then
        Set rs = Nothing
        Exit Sub
    Else
        If Not GetRS(rs, "Select SO006.RowId,SO006.* From SO006 Where SeqNo is null", gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Sub
        rs.ActiveConnection = Nothing
        strSeqNo = GetInvoiceNo2("SO006")
        rs.AddNew
        With rs
            .Fields("SeqNo").Value = strSeqNo
            .Fields("ServiceCode").Value = strServiceCode
            .Fields("ServiceName").Value = strServiceName
            .Fields("ServDescCode").Value = strServiceContentCode
            .Fields("ServDescName").Value = strServiceContentName
            .Fields("CompCode").Value = gilCompCode.GetCodeNo
            .Fields("ServiceType").Value = gilServiceType.GetCodeNo
            .Update
        End With
    End If
    Exit Sub
    'If rsSo006.State = adStateOpen Then rsSo006.Close
    'Set rsSo006 = Nothing
chkErr:
    Call ErrSub(Me.Name, "ReturnSO006")
End Sub
'回傳來電分類代碼
Public Property Let uServiceCode(ByVal vData As String)
    strServiceCode = vData
End Property
'回傳來電分類名稱
Public Property Let uServiceName(ByVal vData As String)
    strServiceName = vData
End Property
'回傳申告內容代碼
Public Property Let uServiceContentCode(ByVal vData As String)
    strServiceContentCode = vData
End Property
'回傳申告內容名稱
Public Property Let uServiceContentName(ByVal vData As String)
    strServiceContentName = vData
End Property
'判斷使用者是否按下確定
Public Property Let uIsOk(ByVal vData As Boolean)
    blnIsOK = vData
End Property
'判斷使用者在瀏覽畫面是按下確定或取消
Public Property Let uViewOK(ByVal vData As Boolean)
    blnViewOk = vData
End Property
Public Property Let uBillNos(ByVal vData As String)
  strBillNos = vData
End Property
Public Property Set uImportExcleData(ByVal vData As Dictionary)
    Set importExcelData = vData
End Property


