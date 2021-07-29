VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{4C60A153-2E5D-4072-B263-967F2D9A9A1B}#2.0#0"; "csIPtxtbox.ocx"
Object = "{CEA96874-3B0B-4FF6-90B8-E48EE8F30715}#2.0#0"; "GiCombo.ocx"
Begin VB.Form frmSO1623A 
   BorderStyle     =   1  '單線固定
   Caption         =   "CM 設定控制台 [SO1171A]"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "SO1623A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11910
   StartUpPosition =   2  '螢幕中央
   Begin VB.PictureBox pic1 
      Appearance      =   0  '平面
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   3500
      ScaleHeight     =   825
      ScaleWidth      =   4995
      TabIndex        =   44
      Top             =   2130
      Visible         =   0   'False
      Width           =   5025
      Begin MSComctlLib.ProgressBar pbr1 
         Height          =   285
         Left            =   150
         TabIndex        =   46
         Top             =   405
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblProcess 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00FFFFFF&
         Caption         =   "開機中 , 請稍侯 !!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   150
         TabIndex        =   45
         Top             =   110
         Width           =   4695
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5310
      ScaleHeight     =   285
      ScaleWidth      =   5025
      TabIndex        =   49
      Top             =   3000
      Width           =   5025
      Begin Gi_Time.GiTime gdtResvTime 
         Height          =   285
         Left            =   900
         TabIndex        =   63
         Top             =   0
         Visible         =   0   'False
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblResv 
         AutoSize        =   -1  'True
         Caption         =   "預約時間"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   180
         Left            =   150
         TabIndex        =   50
         Top             =   60
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Timer Timer1 
      Left            =   60
      Top             =   6780
   End
   Begin VB.CheckBox chkSetAll 
      BackColor       =   &H00C0FFFF&
      Caption         =   "所有CM一起設定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   210
      TabIndex        =   47
      Top             =   390
      Visible         =   0   'False
      Width           =   1695
   End
   Begin prjNumber.GiNumber ginCustId 
      Height          =   285
      Left            =   5130
      TabIndex        =   1
      Top             =   30
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      WithComma       =   0   'False
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
      AllowZero       =   0   'False
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10530
      TabIndex        =   39
      Top             =   30
      Width           =   1245
   End
   Begin VB.TextBox txtCustName 
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
      Left            =   5970
      TabIndex        =   2
      Top             =   30
      Width           =   2985
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   285
      Left            =   1065
      TabIndex        =   0
      Top             =   30
      Width           =   2730
      _ExtentX        =   4815
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
      FldWidth2       =   1800
      F5Corresponding =   -1  'True
   End
   Begin TabDlg.SSTab sstHead 
      Height          =   4155
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   7329
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   591
      ForeColor       =   8388736
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " 設定作業"
      TabPicture(0)   =   "SO1623A.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "sstData"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSet"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "設定記錄"
      TabPicture(1)   =   "SO1623A.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ggrQuery"
      Tab(1).Control(1)=   "cmdQuery"
      Tab(1).Control(2)=   "cmdQuery2"
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmdQuery2 
         Caption         =   "查詢已拆設備設定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -65580
         TabIndex        =   48
         Top             =   3660
         Width           =   1905
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(F3)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66900
         TabIndex        =   38
         Top             =   3660
         Width           =   1125
      End
      Begin prjGiGridR.GiGridR ggrQuery 
         Height          =   3165
         Left            =   -74670
         TabIndex        =   43
         Top             =   450
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   5583
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
      Begin VB.CommandButton cmdSet 
         Caption         =   "F2. 設定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10110
         TabIndex        =   37
         Top             =   400
         Width           =   1245
      End
      Begin TabDlg.SSTab sstData 
         Height          =   3555
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   6271
         _Version        =   393216
         TabsPerRow      =   5
         TabHeight       =   520
         Enabled         =   0   'False
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&1. CM 基本設定"
         TabPicture(0)   =   "SO1623A.frx":047A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraData(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&2. 變更服務"
         TabPicture(1)   =   "SO1623A.frx":0496
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraData(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "&3. 查詢CM資訊"
         TabPicture(2)   =   "SO1623A.frx":04B2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraData(2)"
         Tab(2).ControlCount=   1
         Begin VB.Frame fraData 
            Height          =   2955
            Index           =   2
            Left            =   -74760
            TabIndex        =   52
            Top             =   420
            Width           =   10635
            Begin VB.CheckBox chkPR2 
               Caption         =   "移機中是否取用新地址資訊"
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
               ForeColor       =   &H000000C0&
               Height          =   225
               Left            =   750
               TabIndex        =   29
               Top             =   990
               Width           =   2505
            End
            Begin Gi_Time.GiTime gdtStartTime 
               Height          =   285
               Left            =   4140
               TabIndex        =   34
               Top             =   240
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.OptionButton optQueryInfo 
               Caption         =   "&5. PCIP連線記錄查詢 [26]"
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
               Index           =   4
               Left            =   330
               TabIndex        =   32
               ToolTipText     =   "將STB重新啟用"
               Top             =   2145
               Width           =   2385
            End
            Begin VB.OptionButton optQueryInfo 
               Caption         =   "&4. CM 連線品質查詢 [23]"
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
               Index           =   3
               Left            =   330
               TabIndex        =   31
               ToolTipText     =   "將STB重新啟用"
               Top             =   1740
               Width           =   2385
            End
            Begin VB.OptionButton optQueryInfo 
               Caption         =   "&1. CM 狀態查詢 [13]"
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
               Index           =   0
               Left            =   330
               TabIndex        =   27
               Top             =   300
               Value           =   -1  'True
               Width           =   1965
            End
            Begin VB.OptionButton optQueryInfo 
               Caption         =   "&2. 查詢 HFC 服務類別 [21]"
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
               Index           =   1
               Left            =   330
               TabIndex        =   28
               Top             =   675
               Width           =   2535
            End
            Begin VB.OptionButton optQueryInfo 
               Caption         =   "&3. CM 連線紀錄查詢 [22]"
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
               Index           =   2
               Left            =   330
               TabIndex        =   30
               ToolTipText     =   "將STB 暫停"
               Top             =   1365
               Width           =   2535
            End
            Begin prjGiGridR.GiGridR ggrQueryInfo 
               Height          =   2235
               Left            =   3330
               TabIndex        =   36
               Top             =   600
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   3942
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Gi_Time.GiTime gdtEndTime 
               Height          =   285
               Left            =   6570
               TabIndex        =   35
               Top             =   240
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin IPvalidatePrj.IP_TextBox txtPCIP 
               Height          =   315
               Left            =   1170
               TabIndex        =   33
               Top             =   2490
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   556
               BorderStyle     =   1
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "PCIP"
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
               Left            =   765
               TabIndex        =   55
               Top             =   2580
               Width           =   360
            End
            Begin VB.Label lblTo 
               AutoSize        =   -1  'True
               Caption         =   "至"
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
               Left            =   6330
               TabIndex        =   54
               Top             =   300
               Width           =   180
            End
            Begin VB.Label Label1 
               Caption         =   "查詢期間"
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
               Left            =   3360
               TabIndex        =   53
               Top             =   300
               Width           =   975
            End
         End
         Begin VB.Frame fraData 
            Height          =   2955
            Index           =   1
            Left            =   -74760
            TabIndex        =   51
            Top             =   420
            Width           =   10485
            Begin VB.OptionButton optCMChange 
               Caption         =   "&3. 方案類別更換 [02]"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   8
               Left            =   510
               TabIndex        =   19
               ToolTipText     =   "將STB重新啟用"
               Top             =   1320
               Width           =   3405
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&6. 帳號復權 [24]"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   7
               Left            =   5100
               TabIndex        =   22
               ToolTipText     =   "將STB重新啟用"
               Top             =   825
               Width           =   2505
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&5. 帳號停權 [24]"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   6
               Left            =   5100
               TabIndex        =   21
               ToolTipText     =   "將STB 暫停"
               Top             =   330
               Width           =   2295
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&4. Clear CPE IP [24]"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   5
               Left            =   7470
               TabIndex        =   25
               ToolTipText     =   "將STB重新啟用"
               Top             =   2430
               Visible         =   0   'False
               Width           =   2595
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&3. ACS帳號密碼相關功能 [24]"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   4
               Left            =   7470
               TabIndex        =   24
               ToolTipText     =   "將STB重新啟用"
               Top             =   2235
               Visible         =   0   'False
               Width           =   2655
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&1. Ping CM [20]"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   510
               TabIndex        =   17
               ToolTipText     =   "將STB 暫停"
               Top             =   330
               Width           =   2235
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&2. 變更CM速率 [02]"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   5190
               TabIndex        =   26
               Top             =   2430
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&1. 換CM [02]"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   5100
               TabIndex        =   23
               Top             =   1320
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&2. 釋放動態IP [25]"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   510
               TabIndex        =   18
               ToolTipText     =   "將STB重新啟用"
               Top             =   825
               Width           =   2505
            End
            Begin prjGiList.GiList gilBP 
               Height          =   285
               Left            =   1665
               TabIndex        =   20
               Top             =   1620
               Width           =   2730
               _ExtentX        =   4815
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
               FldWidth2       =   1800
               F5Corresponding =   -1  'True
            End
            Begin VB.Label lblFixIP 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "固定IP : "
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   180
               Left            =   960
               TabIndex        =   59
               Top             =   2520
               Width           =   645
            End
            Begin VB.Label lblDynaIP 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "動態IP : "
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   180
               Left            =   960
               TabIndex        =   58
               Top             =   2220
               Width           =   645
            End
            Begin VB.Label lblBaudRate 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "速率 : "
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   180
               Left            =   960
               TabIndex        =   57
               Top             =   1950
               Width           =   495
            End
            Begin VB.Label lblSln 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "新方案 : "
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   180
               Left            =   960
               TabIndex        =   56
               Top             =   1680
               Width           =   675
            End
         End
         Begin VB.Frame fraData 
            Height          =   2955
            Index           =   0
            Left            =   240
            TabIndex        =   42
            Top             =   420
            Width           =   10485
            Begin VB.PictureBox picCPno 
               Appearance      =   0  '平面
               BorderStyle     =   0  '沒有框線
               ForeColor       =   &H80000008&
               Height          =   405
               Left            =   5100
               ScaleHeight     =   405
               ScaleWidth      =   2895
               TabIndex        =   60
               Top             =   660
               Visible         =   0   'False
               Width           =   2895
               Begin PrjGiCmobo.GiCombo cboCPno 
                  Height          =   300
                  Left            =   690
                  TabIndex        =   61
                  Top             =   30
                  Width           =   2145
                  _ExtentX        =   3784
                  _ExtentY        =   529
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "新細明體"
                     Size            =   9
                     Charset         =   136
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ListIndex       =   -1
                  SelStart        =   0
                  SelLength       =   0
                  ItemData0       =   0
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1028
                     SubFormatType   =   0
                  EndProperty
               End
               Begin VB.Label lblCPno 
                  Caption         =   "CP門號"
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
                  Left            =   30
                  TabIndex        =   62
                  Top             =   90
                  Width           =   570
               End
            End
            Begin VB.CheckBox chkPR 
               Caption         =   "移機中是否取用新地址資訊"
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
               ForeColor       =   &H000000C0&
               Height          =   225
               Left            =   6750
               TabIndex        =   15
               Top             =   390
               Width           =   2535
            End
            Begin VB.CheckBox ChkAccStart 
               Caption         =   "帳號復權"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   210
               Left            =   1920
               TabIndex        =   8
               Top             =   390
               Width           =   1065
            End
            Begin VB.CheckBox chkAccStop 
               Caption         =   "帳號停權"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   210
               Left            =   1920
               TabIndex        =   10
               Top             =   900
               Width           =   1065
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&5. CP 退租 [27]"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   5
               Left            =   4680
               TabIndex        =   16
               ToolTipText     =   $"SO1623A.frx":04CE
               Top             =   1410
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&4. CP 申裝 [27]"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   4
               Left            =   4680
               TabIndex        =   14
               ToolTipText     =   $"SO1623A.frx":04DD
               Top             =   390
               Width           =   1815
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&3. CM Reset [12]"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   3
               Left            =   510
               TabIndex        =   11
               ToolTipText     =   $"SO1623A.frx":04EC
               Top             =   1410
               Width           =   2265
            End
            Begin VB.OptionButton optBasicXX 
               Caption         =   "&4. 復機"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   510
               TabIndex        =   13
               ToolTipText     =   $"SO1623A.frx":04FA
               Top             =   2430
               Visible         =   0   'False
               Width           =   1365
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&1. 開機 [02]"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   510
               TabIndex        =   7
               ToolTipText     =   $"SO1623A.frx":0506
               Top             =   390
               Width           =   1815
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&2. 關機 [02]"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   510
               TabIndex        =   9
               ToolTipText     =   $"SO1623A.frx":0512
               Top             =   900
               Width           =   1815
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&3. 停機 [02]"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   2
               Left            =   510
               TabIndex        =   12
               ToolTipText     =   $"SO1623A.frx":051E
               Top             =   2130
               Visible         =   0   'False
               Width           =   1785
            End
         End
      End
   End
   Begin prjGiGridR.GiGridR ggrChild 
      Height          =   2265
      Left            =   6420
      TabIndex        =   4
      Top             =   630
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   3995
      Enabled         =   0   'False
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
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   2265
      Left            =   180
      TabIndex        =   3
      Top             =   630
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   3995
      Enabled         =   0   'False
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
   Begin VB.Label lblCustId 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號"
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
      Left            =   4260
      TabIndex        =   41
      Top             =   90
      Width           =   720
   End
   Begin VB.Label lblComp 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   240
      TabIndex        =   40
      Top             =   90
      Width           =   720
   End
End
Attribute VB_Name = "frmSO1623A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents rsData As ADODB.Recordset
Attribute rsData.VB_VarHelpID = -1
Dim WithEvents rsChild As ADODB.Recordset
Attribute rsChild.VB_VarHelpID = -1
Dim WithEvents rsQuery As ADODB.Recordset
Attribute rsQuery.VB_VarHelpID = -1

Dim lngCustId As Long
Dim lngInstAddrNo As Long
Dim lngNodeNo As Long
Dim lngCircuitNo As Long
Dim strChStr As String, strStopChStr As String, strOpenChStr As String
Dim strChStopDate As String
Dim rsParent As New ADODB.Recordset
Dim intProcessType As Integer
Dim blnProcessOk As Boolean
Dim strProcessChStr As String
Dim strDoCaption As String
Dim strCaption As String
Dim blnTransation As Boolean
Dim strResvTime As String
Dim blnSetting As Boolean
Dim strDefaultRowId As String
Dim strOrderNo As String
Dim blnProcessLock As Boolean
Dim intByHouse As Integer

Private Sub cboCPno_DropDown()
  On Error Resume Next
    Dim varAry, varElement
    Dim strQry As String
    strQry = "SELECT DISTINCT FACISNO FROM " & GetOwner & "SO004" & _
                    " WHERE REFACISNO='" & rsData("FaciSno") & "'" & _
                    " AND COMPCODE=" & rsData("CompCode") & _
                    " AND CUSTID=" & rsData("CustID") & _
                    " AND SERVICETYPE='P'" & _
                    " AND PRDATE IS NULL" & _
                    " AND FACISNO IS NOT NULL"
    varAry = gcnGi.Execute(strQry).GetRows
    cboCPno.Clear
    For Each varElement In varAry
        If varElement <> Empty Then cboCPno.AddItem varElement & ""
    Next
    If cboCPno.ListCount <= 0 Then MsgBox "此設備無CP門號，無法正常傳送指令 !", vbInformation, "訊息"
End Sub

Private Sub ChkAccStart_Click()
  On Error Resume Next
    chkAccStop.Value = Abs(ChkAccStart.Value) - 1
    If ChkAccStart.Value Then optBasic(0).Value = True
End Sub

Private Sub chkAccStop_Click()
  On Error Resume Next
    ChkAccStart.Value = Abs(chkAccStop.Value) - 1
    If chkAccStop.Value Then optBasic(1).Value = True
End Sub

Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
    Set frmSO1623A = Nothing
End Sub

Private Sub cmdQuery_Click()
  On Error GoTo ChkErr
    Fetch_Log_Data
  Exit Sub
ChkErr:
    ErrSub Name, "cmdQuery_Click"
End Sub

Private Sub Fetch_Log_Data(Optional blnPRflag As Boolean = False, Optional blnEmptyGrid As Boolean = False)
  On Error GoTo ChkErr
    Dim rsLog As New ADODB.Recordset
    Dim lngRcdCnt As Long
    Dim strQry As String
    
    If blnEmptyGrid Then
        strQry = "SELECT * FROM " & GetOwner & "SO311 WHERE 0=1"
    Else
        If rsData Is Nothing Then Exit Sub
        With rsData
                If .State <= 0 Then Exit Sub
                If .EOF Or .BOF Then Exit Sub
                lngRcdCnt = .RecordCount
                If lngRcdCnt = 0 Then Exit Sub
        End With
        strQry = "SELECT * FROM " & GetOwner & "SO311" & _
                        " WHERE CUSSO='" & rsData("CompCode") & "'" & _
                        " AND CUSID='" & rsData("CustId") & "'"
        If lngRcdCnt = 1 Then
            If blnPRflag Then
                strQry = strQry & " AND CMMAC<>'" & rsData("FaciSNo") & "'"
            Else
                strQry = strQry & " AND CMMAC='" & rsData("FaciSNo") & "'"
            End If
        Else
            If blnPRflag Then
                strQry = strQry & " AND CMMAC NOT IN (" & Get_Faci_Rows & ")"
            Else
                strQry = strQry & " AND CMMAC IN (" & Get_Faci_Rows & ")"
            End If
        End If
        strQry = strQry & " ORDER BY EXECDATE"
    End If
    
    If GetRS(rsLog, strQry, gcnGi, adUseClient, adOpenStatic, adLockReadOnly) Then GrdFmt rsLog
  
  Exit Sub
ChkErr:
    ErrSub Name, "Fetch_Log_Data"
End Sub

Private Sub cmdQuery2_Click()
  On Error GoTo ChkErr
    Fetch_Log_Data True
  Exit Sub
ChkErr:
    ErrSub Name, "cmdQuery2_Click"
End Sub

Private Sub GrdFmt(ByRef rsLog As Object)
  On Error GoTo ChkErr
  
    Dim mFlds As New GiGridFlds
    
    With mFlds
            .Add "MSGCODE", , , , , "執行代碼", vbLeftJustify
            .Add "MSGNAME", , , , , "執行名稱 ", vbLeftJustify
            .Add "WORKID", , , , , "員工編號", vbLeftJustify
            .Add "CUSSO", , , , , "系統台別", vbLeftJustify
            .Add "CUSID", , , , , "客戶編號", vbLeftJustify
            .Add "CMMAC", , , , , "CM MAC      ", vbLeftJustify
            .Add "OPERRESULT", , , , , "操作執行結果", vbLeftJustify
            .Add "EXECENTRY", , , , , "執行者      ", vbLeftJustify
            '************************************************************************************
            '#3418 因為執行日期是字串，所以將Format西元年的參數拿掉 By Kin 2007/08/08
            .Add "EXECDATE", , , , , "執行日期" & Space(10), vbLeftJustify
            '************************************************************************************
            .Add "EMTAMAC", , , , , "EMTA MAC", vbLeftJustify
            .Add "EMTAPORT", , , , , "EMTA Port", vbRightJustify
            .Add "CPNO", , , , , "CP電話號碼", vbLeftJustify
            .Add "CLASSID", , , , , "方案類別", vbLeftJustify
            .Add "PCIPNO", , , , , "動態PCIP數", vbLeftJustify
            .Add "PCIP", , , , , "PCIP", vbLeftJustify
            .Add "OPERTYPE", , , , , "功能指令(開關)", vbLeftJustify
            .Add "STARTDATETIME", , , , , "起始時間", vbLeftJustify
            .Add "ENDDATETIME", , , , , "截止時間", vbLeftJustify
            .Add "ZONE", , , , , "鄉鎮市\村里莊路街\街道名稱", vbLeftJustify
            .Add "LIN", , , , , "鄰  ", vbLeftJustify
            .Add "SECTION", , , , , "段  ", vbLeftJustify
            .Add "LANE", , , , , "巷  ", vbLeftJustify
            .Add "ALLEY", , , , , "弄  ", vbLeftJustify
            .Add "SUBALLEY", , , , , "衖  ", vbLeftJustify
            .Add "NO1", , , , , "號一 ", vbLeftJustify
            .Add "NO2", , , , , "號二 ", vbLeftJustify
            .Add "PROCESSINGDATE", giControlTypeTime, , , , "預約時間    ", vbLeftJustify
            .Add "QUERYRESULT", , , , , "查詢執行結果", vbLeftJustify
            .Add "FAULTREASON", , , , , "失敗原因", vbLeftJustify
            .Add "GROUPID", , , , , "Group ID", vbLeftJustify
            .Add "CMSTATUS", , , , , "狀態  ", vbLeftJustify
            .Add "RCMIP", , , , , "CM IP            ", vbLeftJustify
            .Add "RCMUPFREQ", , , , , "上行頻率 ", vbLeftJustify
            .Add "RCMDOWNFREQ", , , , , "下行頻率 ", vbLeftJustify
            .Add "RCMRECPW", , , , , "接收功率 ", vbLeftJustify
            .Add "RCMTRANSPW", , , , , "發射功率 ", vbLeftJustify
            .Add "RCMSNR", , , , , "訊號雜訊比", vbLeftJustify
            .Add "RCMONLINETIMES", , , , , "開機時間 ", vbLeftJustify
            .Add "RCMINOCTETS", , , , , "上行累計封包", vbLeftJustify
            .Add "RCMOUTOCTETS", , , , , "下行累計封包", vbLeftJustify
            .Add "RCMINERRORS", , , , , "上行錯誤封包", vbLeftJustify
            .Add "RCMOUTERRORS", , , , , "下行錯誤封包", vbLeftJustify
            .Add "RPCIP", , , , , "PCIP回覆訊息", vbLeftJustify
            .Add "RCMTSRECPW", , , , , "RecPw    ", vbLeftJustify
            .Add "RCMTSRXSNR", , , , , "RxSNR    ", vbLeftJustify
            .Add "RCMTSUPMOD", , , , , "調變方式 ", vbLeftJustify
            .Add "RHFCTYPE", , , , , "服務種類", vbLeftJustify
            .Add "RPCOFFLINETIME", , , , , "PC上線時間  ", vbLeftJustify
            .Add "RPCONLINETIME", , , , , "PC上線時間  ", vbLeftJustify
            .Add "ROFFLINETIME", , , , , "CM上線時間  ", vbLeftJustify
            .Add "RONLINETIME", , , , , "CM離線時間  ", vbLeftJustify
            .Add "RDETCTTIME", , , , , "偵測時間    ", vbLeftJustify
            .Add "RCMMAC", , , , , "CM MAC回覆訊息", vbLeftJustify
            .Add "RCTIP", , , , , "CTIP             ", vbLeftJustify
            .Add "RNODEID", , , , , "NODEID   ", vbLeftJustify
            .Add "HOSTNAME", , , , , "電腦名稱          ", vbLeftJustify
            .Add "CMDSEQNO", , , , , "指令序號               ", vbLeftJustify
            .Add "DIALACCOUNT", , , , , "撥接帳號  ", vbLeftJustify
            .Add "DIALPASSWORD", , , , , "撥接密碼  ", vbLeftJustify
            .Add "CUSTNAME", , , , , "申裝人姓名 ", vbLeftJustify
            .Add "FINTIME", giControlTypeDate, , , , "完工時間     ", vbLeftJustify
            .Add "TEL1", , , , , "電話一            ", vbLeftJustify
            .Add "TEL2", , , , , "聯絡人電話", vbLeftJustify
            .Add "COMMANDTYPE", , , , , "指令種類 ", vbLeftJustify
            .Add "XMLDATA", , , , , "XML資料            ", vbLeftJustify
            .Add "ID", , , , , "申裝人身份證號", vbLeftJustify
            .Add "STOPDATE", giControlTypeTime, , , , "取消時間" & Space(12), vbLeftJustify
            .Add "CANCELNAME", , , , , "取消原因              ", vbLeftJustify
            'StopFlag
            ' ggrQuery
    End With
    
'Hammer: Hi
'cable_liga: YO
'Hammer: STOPFLAG , STOPDATE, CANCELNAME
'Hammer: 要不要全 Show
'cable_liga: 嗯~~~好嗚！你很貼心！
'cable_liga: 那就多秀 STOPDATE, CANCELNAME 即可
'cable_liga: STOPFLAG , CANCELCODE就不用了
'Hammer: ok
    
    With ggrQuery
            .AllFields = mFlds
            Set .Recordset = rsLog
            .Refresh
    End With
    
  Exit Sub
ChkErr:
    ErrSub Name, "GrdFmt"
End Sub

Private Function Get_Faci_Rows() As String
  On Error GoTo ChkErr
    
    Dim rsClone As New ADODB.Recordset
    
    Dim varAry As Variant
    Dim varElement As Variant
    Dim strValue As String
    
    Set rsClone = rsData.Clone
    
    With rsClone
        If .RecordCount > 0 Then .MoveFirst
        varAry = rsClone.GetRows(, 1, "FaciSNo")
        For Each varElement In varAry
            If varElement & "" <> Empty Then strValue = strValue & "','" & varElement & ""
        Next
    End With
    
    strValue = Mid(strValue, 3) & "'"
    Get_Faci_Rows = strValue
    
    On Error Resume Next
    CloseRecordset rsClone
    
    Erase varAry
    varElement = Empty
  Exit Function
ChkErr:
    ErrSub Name, "Get_Faci_Rows"
End Function

Private Sub cmdSet_Click()
    On Error GoTo ChkErr
    Dim lngBR As Long
    Dim intLoop As Integer
    Dim intIndex As Integer
    Dim strResult As String, strFaultReason As String, strErrMsg As String
    Dim rsResult As New ADODB.Recordset
    Dim blnShowXML As Boolean
    
    blnShowXML = True
    
        If Not IsDataOk Then Exit Sub
        Me.Enabled = False
        strResvTime = gdtResvTime.GetValue(True)
        blnSetting = True
        Call subProgressBar(True)
        
        On Error Resume Next
        With ggrQueryInfo
            If .Recordset.State > 0 Then
                .Blank = True
                .ClearStructure
            End If
        End With
        
        On Error GoTo ErrTran
        If Not blnTransation Then gcnGi.BeginTrans
        lngBR = rsData.AbsolutePosition
            Select Case sstData.Tab
                Case 0
                    Select Case True
                        '開機
                        Case optBasic(0)
                            If Not OpenCM(rsData, GetPCMac, lngInstAddrNo, GetClassID, GetDynaIP, strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                            
                            If ChkAccStart.Value Then
                                lblProcess.Caption = "帳號復權中 , 請稍候 !!"
                                lblProcess.Refresh
                                If Not Cmd_Acc_24(rsData, "Enable", "") Then GoTo ErrTran
                            End If
                            
'                            MsgBox "準備 OpenData"
                            Call OpenData2
                        '關機
                        Case optBasic(1)
                            If Not CloseCM(rsData, GetPCMac, lngInstAddrNo, GetClassID, GetDynaIP, strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                            ' Liga :
                            '   1．關機[02]：關機後少送一個 RESET CM的指令；關機後尚需要有一個「reset CM」的指令，才會真正關機
                            
                            If chkAccStop.Value Then
                                lblProcess.Caption = "帳號停權中 , 請稍候 !!"
                                lblProcess.Refresh
                                If Not Cmd_Acc_24(rsData, "disable", "") Then GoTo ErrTran
                            End If
                            
                            With rsResult
                                If .State > 0 Then
                                    If .RecordCount > 0 Then
                                        .MoveFirst
                                        If .Fields(0) = 1 Then
                                            lblProcess.Caption = "ResetCM中 , 請稍候 !!"
                                            lblProcess.Refresh
                                            If Not ResetCM(rsData, strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                                        End If
                                    End If
                                End If
                            End With
                            
                            Call OpenData2
                            
'                        Case optBasic(2)
'                            If Not StopCM(rsData, GetPCMac, lngInstAddrNo, GetClassID, GetDynaIP, strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                        
                        Case optBasic(3)
                            If Not ResetCM(rsData, strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                        Case optBasic(4)
                            If chkPR.Value Then lngInstAddrNo = GetNewAddr
                            If Not CPInst(rsData, GetPCMac, lngInstAddrNo, strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                        Case optBasic(5)
                            If Not CPPR(rsData, GetPCMac, lngInstAddrNo, strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                    End Select
                Case 1
                    Select Case True
'                        Case optCMChange(0)
'                            If Not ChangeCM(rsData, GetPCMac, lngInstAddrNo, GetClassID, GetDynaIP, strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
'                        Case optCMChange(1)
'                            If Not ChangeCMRate(rsData, GetPCMac, lngInstAddrNo, GetClassID, GetDynaIP, strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                        Case optCMChange(2)
                            If Not PingCM(rsData, strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                        Case optCMChange(3)
'                            If Not ReleaseDyIP(rsData, strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                            If Not ClearCPEIP(rsData, strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
'                        Case optCMChange(4)
'                            If Not ACSAccPWS(rsData, strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
'                        Case optCMChange(5)
'                            If Not ClearCPEIP(rsData, strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                        Case optCMChange(6)
                            blnShowXML = False
                            If Not Cmd_Acc_24(rsData, "disable") Then GoTo ErrTran
                        Case optCMChange(7)
                            blnShowXML = False
                            If Not Cmd_Acc_24(rsData, "Enable") Then GoTo ErrTran
                        Case optCMChange(8)
                            If gilBP.GetCodeNo = Empty Then
                                MsgBox "請選擇新方案類別 !", vbInformation, "訊息"
                            Else
                                If Not OpenCM(rsData, GetPCMac, lngInstAddrNo, GetClassID(gilBP.GetCodeNo), GetDynaIP(gilBP.GetCodeNo), strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                            End If
                    End Select
                Case 2
                    Select Case True
                        Case optQueryInfo(0)
                            If Not QueryCMStatus(rsData, lngInstAddrNo, strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                        Case optQueryInfo(1)
                            If chkPR.Value Then lngInstAddrNo = GetNewAddr
                            If Not QueryHFCServ(rsData, lngInstAddrNo, strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                        Case optQueryInfo(2)
                            If Not QueryCMHistory(rsData, gdtStartTime.GetValue(True), gdtEndTime.GetValue(True), strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                        Case optQueryInfo(3)
                            If Not QueryCMQuality(rsData, gdtStartTime.GetValue(True), gdtEndTime.GetValue(True), strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                        Case optQueryInfo(4)
                            If Not QueryPCIPHistory(rsData, txtPCIP.Text, gdtStartTime.GetValue(True), gdtEndTime.GetValue(True), strResult, strFaultReason, strErrMsg, rsResult) Then GoTo ErrTran
                    End Select
            End Select
'            If chkSetAll = 0 Then Exit Do
'            rsData.MoveNext
'        Loop
66:
        'rsData.AbsolutePosition = lngBR
        blnSetting = False
'        If chkSetAll = 1 Then Call OpenData2
        If Not blnTransation Then gcnGi.CommitTrans
        '#3271執行關機如果IP狀態有改變，則再進行開機，會出現找不到資料列的訊息，所以再乎叫OpenData一次 By Kin 2007/06/05
        'rsData.Requery
        'OpenData
        rsData.AbsolutePosition = lngBR
        
        blnProcessOk = True
        Call subProgressBar(False)
        Me.Enabled = True
        If blnShowXML And gdtResvTime.Text = Empty Then Show_XML_Result rsResult
        If intProcessType > 0 Then Unload Me
    Exit Sub
ErrTran:
'        If Err = 0 Then GoTo 66
'        MsgBox Err.Description, , " (錯誤) " & Err.Number
        On Error Resume Next
        rsData.AbsolutePosition = lngBR
        On Error GoTo ChkErr
'        If chkSetAll = 1 Then Call OpenData2
        Call DoProcessMode
        blnSetting = False
        If Not blnTransation Then gcnGi.RollbackTrans
        '#3271執行關機如果IP狀態有改變，則再進行開機，會出現找不到資料列的訊息，所以再乎叫OpenData一次 By Kin 2007/06/05
        OpenData
        
        rsData.AbsolutePosition = lngBR
        
        pic1.Visible = False
        Me.Enabled = True
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdSet_Click"
End Sub

Private Function GetNewAddr() As Long
  On Error Resume Next
    Dim strQry As String
    strQry = "SELECT REINSTADDRNO FROM " & GetOwner & "SO009" & _
                    " WHERE CUSTID=" & ginCustId.Text & _
                    " AND COMPCODE=" & gilCompCode.GetCodeNo & _
                    " AND PRCODE IN ( SELECT CODENO FROM " & GetOwner & "CD007 WHERE REFNO=3 )" & _
                    " AND FINTIME IS NULL AND RETURNCODE IS NULL"
    GetNewAddr = Val(GetRsValue(strQry) & "")
    If Err.Number <> 0 Then
        Err.Clear
        GetNewAddr = 0
    End If
End Function


Private Sub Show_XML_Result(ByRef rsResult As ADODB.Recordset)
  On Error GoTo ChkErr
    
    Dim mFlds As New GiGridFlds
    Dim intLoop As Integer
    
    ggrQueryInfo.Blank = True
    
    With rsResult
            If .RecordCount = 0 Then Exit Sub
            .MoveFirst
            For intLoop = 0 To .Fields.Count - 1
                SetGrdFld mFlds, .Fields(intLoop).Name
            Next
    End With
    
    With ggrQueryInfo
            .AllFields = mFlds
            Set .Recordset = rsResult
            .Refresh
            .Blank = False
    End With
    
  Exit Sub
ChkErr:
    ErrSub Name, "Show_XML_Result"
End Sub

Private Sub SetGrdFld(mFlds As GiGridFlds, strFldName As String)
  On Error GoTo ChkErr
    With mFlds
            Select Case LCase(strFldName)
                        Case "queryresult"
                                    .Add "QueryResult", , , , , "查詢執行結果", vbLeftJustify
                        Case "operresult"
                                    .Add "OperResult", , , , , "操作執行結果", vbLeftJustify
                        Case "faultreason"
                                    .Add "FaultReason", , , , , "失敗原因                      ", vbLeftJustify
                        Case "groupid"
                                    .Add "GROUPID", , , , , "Group ID ", vbLeftJustify
                        Case "cmstatus"
                                    .Add "CMStatus", , , , , "狀態        ", vbLeftJustify
                        Case "cmip"
                                    .Add "CMIP", , , , , "CM IP            ", vbLeftJustify
                        Case "cmupfreq"
                                    .Add "CMUPFREQ", , , , , "上行頻率 ", vbLeftJustify
                        Case "cmdownfreq"
                                    .Add "CMDOWNFREQ", , , , , "下行頻率 ", vbLeftJustify
                        Case "cmrecpw"
                                    .Add "CMRECPW", , , , , "接收功率 ", vbLeftJustify
                        Case "cmtranspw"
                                    .Add "CMTRANSPW", , , , , "發射功率 ", vbLeftJustify
                        Case "cmsnr"
                                    .Add "CMSNR", , , , , "訊號雜訊比 ", vbLeftJustify
                        Case "cmonlinetimes"
                                    .Add "CMOnlineTimes", , , , , "開機時間 ", vbLeftJustify
                        Case "cminoctets"
                                    .Add "CMInOctets", , , , , "上行累計封包 ", vbLeftJustify
                        Case "cmoutoctets"
                                    .Add "CMOutOctets", , , , , "下行累計封包 ", vbLeftJustify
                        Case "cminerrors"
                                    .Add "CMInErrors", , , , , "上行錯誤封包 ", vbLeftJustify
                        Case "cmouterrors"
                                    .Add "CMOutErrors", , , , , "下行錯誤封包 ", vbLeftJustify
                        Case "pcip"
                                    .Add "PCIP", , , , , "PCIP回覆訊息 ", vbLeftJustify
                        Case "cmtsrecpw"
                                    .Add "CMTSRECPW", , , , , "RecPw ", vbLeftJustify
                        Case "cmtsrxsnr"
                                    .Add "CMTSRXSNR", , , , , "RxSNR ", vbLeftJustify
                        Case "cmtsupmod"
                                    .Add "CMTSUPMOD", , , , , "調變方式 ", vbLeftJustify
                        Case "hfctype"
                                    .Add "HFCType", , , , , "服務種類 ", vbLeftJustify
                        Case "offlinetime"
                                    .Add "OffLineTime", , , , , "CM上線時間      ", vbLeftJustify
                        Case "onlinetime"
                                    .Add "OnLineTime", , , , , "CM離線時間      ", vbLeftJustify
                        Case "detcttime"
                                    .Add "DetctTime", , , , , "偵測時間      ", vbLeftJustify
                        Case "cmmac"
                                    .Add "CMMAC", , , , , "CM MAC回覆訊息", vbLeftJustify
                        Case "ctip"
                                    .Add "CTIP", , , , , "CTIP            ", vbLeftJustify
                        Case "nodeid"
                                    .Add "NODEID", , , , , "NODEID      ", vbLeftJustify
                        Case "pconlinetime" ' ? CMPCLog.xml
                                    .Add "PCOnlineTime", , , , , "PC上線時間      ", vbLeftJustify
                        Case "pcofflinetime"
                                    .Add "PCOfflineTime", , , , , "PC上線時間      ", vbLeftJustify
                        Case Else
                                    Debug.Print strFldName
            End Select
    End With
  Exit Sub
ChkErr:
    ErrSub Name, "SetGrdFld"
End Sub

Private Function GetDynaIP(Optional strBPcode As String = "") As String
    On Error Resume Next
        Dim strDynaIPfld As String
        If rsData("BPcode") & "" = "" Then Exit Function
        strDynaIPfld = IIf(IsCMCP, "CMCPDYNIP", "DYNIPCOUNT")
        If strBPcode = Empty Then
            If rsData("BPcode") & "" = "" Then Exit Function
            GetDynaIP = GetRsValue("SELECT " & strDynaIPfld & " FROM " & GetOwner & "CD078D" & _
                                                    " WHERE BPCODE='" & rsData("BPcode") & "'" & _
                                                    " AND SERVICETYPE='" & rsData("ServiceType") & "'")
        Else
            GetDynaIP = GetRsValue("SELECT " & strDynaIPfld & " FROM " & GetOwner & "CD078D" & _
                                                    " WHERE BPCODE = '" & strBPcode & "'" & _
                                                    " AND SERVICETYPE='" & rsData("ServiceType") & "'")
        End If
        If Err.Number <> 0 Then GetDynaIP = Empty
End Function

Private Function GetClassID(Optional strBPcode As String = "") As String
    On Error Resume Next
        Dim strClassIdFld As String
'        If rsData("CMBaudNo") & "" = "" Then Exit Function
        strClassIdFld = IIf(IsCMCP, "EMCCMCP", "EMCCM")
'        GetClassID = GetRsValue("SELECT " & strClassIdFld & " FROM " & GetOwner & "CD064 WHERE CODENO = " & rsData("CMBaudNo"))
        If strBPcode = Empty Then
            If rsData("BPcode") & "" = "" Then Exit Function
            GetClassID = GetRsValue("SELECT " & strClassIdFld & " FROM " & GetOwner & "CD078D" & _
                                                        " WHERE BPCODE='" & rsData("BPcode") & "'" & _
                                                        " AND SERVICETYPE='" & rsData("ServiceType") & "'")
        Else
            GetClassID = GetRsValue("SELECT " & strClassIdFld & " FROM " & GetOwner & "CD078D" & _
                                                        " WHERE BPCODE = '" & strBPcode & "'" & _
                                                        " AND SERVICETYPE='" & rsData("ServiceType") & "'")
        End If
        If Err.Number <> 0 Then GetClassID = Empty
End Function

Public Function IsCMCP() As Boolean
  On Error GoTo ChkErr
        IsCMCP = (GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                            " WHERE REFACISNO='" & rsData("FaciSno") & "'" & _
                                            " AND COMPCODE=" & rsData("CompCode") & _
                                            " AND CUSTID=" & rsData("CustID") & _
                                            " AND SERVICETYPE='P'" & _
                                            " AND PRDATE IS NULL") > 0)

'                                            " AND INSTDATE IS NOT NULL" & _
'        Dim strWhere As String
'        strWhere = " AND SERVICETYPE='P'" & _
'                            " AND CUSTID=" & rsData("CustID") & _
'                            " AND COMPCODE=" & rsData("CompCode")
'        IsCMCP = Len(GetRsValue("SELECT PRDATE FROM " & GetOwner & "SO004" & _
'                                " WHERE INSTDATE=" & _
'                                " (SELECT MAX(INSTDATE) FROM " & GetOwner & "SO004" & _
'                                " WHERE REFACISNO='" & rsData("FaciSno") & "'" & strWhere & _
'                                " ) " & strWhere) & "") = 0
  Exit Function
ChkErr:
    ErrSub Name, "IsCMCP"
End Function

Private Sub subProgressBar(blnFlag As Boolean)
    On Error Resume Next
    Dim intLoop As Integer
    Dim strDoCaption As String
    Dim objControl As Object
        pbr1.Value = 0
        pic1.Visible = blnFlag
        If Not blnFlag Then Exit Sub
        Select Case sstData.Tab
            Case 0
                Set objControl = optBasic
            Case 1
                Set objControl = optCMChange
            Case 2
                Set objControl = optQueryInfo
        End Select
        For intLoop = 0 To objControl.UBound
            If objControl(intLoop) Then
                strDoCaption = Mid(objControl(intLoop).Caption, 4)
            End If
        Next
        lblProcess.Caption = strDoCaption & " 中 , 請稍候 !!"

End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr

'        Call OpenConnection
        Call subGil
        Call subGrd
        Call DefaultValue
        Call OpenData
        Call DoProcessMode
        Fetch_Log_Data , True

    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

'Private Sub OpenConnection()
'    On Error GoTo ChkErr
'        If gcnGi2.State = adStateClosed Then
'            gcnGi2.Open gcnGi.ConnectionString
'        End If
'        Set objCmd = CreateObject("CS_Web.clsEMCcommand")
'    Exit Sub
'ChkErr:
'    ErrSub Me.Name, "OpenConnection"
'End Sub

Private Sub RsFind(rsTmp As ADODB.Recordset, strField As String, strValue As String)
    On Error GoTo ChkErr
        ggrData.Visible = False
        With rsTmp
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                If rsTmp(strField) = strValue Then Exit Do
                .MoveNext
            Loop
        End With
        ggrData.Visible = True
    Exit Sub
ChkErr:
    ggrData.Visible = True
    ErrSub Me.Name, "RsFind"
End Sub

Private Sub DoProcessMode()
    On Error GoTo ChkErr
    Dim objControl As Object
    Dim blnFlag As Boolean
    Dim intTab As Integer
    Dim intLoop As Integer
    Dim intIndex As Integer
        If intProcessType <= 0 Then Exit Sub
        If strDefaultRowId <> "" Then RsFind rsData, "RowId", strDefaultRowId
        
        For Each objControl In Me
            If TypeOf objControl Is OptionButton Then
                objControl.Enabled = False
            End If
        Next
        
        gilCompCode.Locked = True
        gilCompCode.Enabled = False
        sstData.Enabled = True
        sstHead.Enabled = True
        sstData.TabEnabled(0) = False
        sstData.TabEnabled(1) = False
        sstData.TabEnabled(2) = False
        
        intLoop = Val(Left(Format(intProcessType, "000"), 1))
        sstData.TabEnabled(intLoop - 1) = True
        sstData.Tab = intLoop - 1
        fraData(intLoop - 1).Enabled = True
        
        intIndex = Val(Right(Format(intProcessType, "000"), 2))
        Select Case intLoop
            Case 1
                ChkAccStart.Enabled = False
                chkAccStop.Enabled = False
                Select Case intIndex
                    Case 0
                        'ChkAccStart.Enabled = True
                        ChkAccStart.Enabled = False
                        ChkAccStart.Value = Abs(blnAccChk)
                    Case 1
                        'chkAccStop.Enabled = True
                        chkAccStop.Enabled = False
                        chkAccStop.Value = Abs(blnAccChk)
                End Select
                optBasic(intIndex).Enabled = True
                optBasic(intIndex).Value = False
                optBasic(intIndex).Value = True
            Case 2
                gilBP.Enabled = False
                Select Case intIndex
                    Case 8
                        gilBP.Enabled = True
                End Select
                optCMChange(intIndex).Enabled = True
                optCMChange(intIndex).Value = False
                optCMChange(intIndex).Value = True
            Case 3
                gdtStartTime.Enabled = False
                gdtEndTime.Enabled = False
                txtPCIP.Enabled = False
                Select Case intIndex
                    Case 0
                        gdtStartTime.Enabled = True
                        gdtEndTime.Enabled = True
                    Case 4
                        txtPCIP.Enabled = True
                End Select
                optQueryInfo(intIndex).Enabled = True
                optQueryInfo(intIndex).Value = False
                optQueryInfo(intIndex).Value = True
        End Select
        
        sstHead.Tab = 0
        ggrData.Enabled = False
        
    Exit Sub
ChkErr:
    ErrSub Me.Name, "DoProcessMode"
End Sub

Private Sub subGil()
    On Error GoTo ChkErr
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 20, GetCompCodeFilter(gilCompCode, garyGi(15))
        SetgiList gilBP, "CodeNo", "Description", "CD078", , , , , 3, 20, , True
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub subGim()
    On Error GoTo ChkErr
    Dim strFilter As String
        If strProcessChStr <> "" Then
'            SetgiMulti gimOpenCh, "CodeNo", "Description", "CD024", "頻道代碼", "頻道名稱", , True
        Else
'            SetgiMulti gimOpenCh, "CodeNo", "Description", "CD024", "頻道代碼", "頻道名稱", "Where Nvl(CompCode," & gilCompCode.GetCodeNo & ") = " & gilCompCode.GetCodeNo & strChNotQry, True
        End If
        
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGrd()
    On Error GoTo ChkErr
    Dim mFlds As New prjGiGridR.GiGridFlds
    Dim mFlds2 As New prjGiGridR.GiGridFlds
    Dim mFlds3 As New prjGiGridR.GiGridFlds
      '定義GirdR各個欄位的內容
      'STB序號,ICC序號,裝設位置,狀態
        With mFlds
            .Add "FaciSNo", , , , , "CM 序號             ", vbLeftJustify
            .Add "FaciCode", , , , , "代碼 ", vbLeftJustify
            .Add "FaciName", , , , , "項目名稱 ", vbLeftJustify
            .Add "ModelName", , , , , "型號名稱 ", vbLeftJustify
            .Add "CMBaudNo", , , , , "代碼", vbLeftJustify
            .Add "CMBaudRate", , , , , "CM速率 ", vbLeftJustify
            .Add "DynIPCount", , , , , "動態IP數目", vbRightJustify
            .Add "FixIPCount", , , , , "固定IP數目", vbRightJustify
            .Add "InstDate", giControlTypeDate, , , , "裝機日期", vbLeftJustify
            .Add "CMOpenDate", giControlTypeTime, , , , "CM開機日期  ", vbLeftJustify
            .Add "CMCloseDate", giControlTypeTime, , , , "CM關機日期  ", vbLeftJustify
            .Add "DialAccount", , , , , "撥接帳號 ", vbLeftJustify
            .Add "EMail", , , , , "電子信箱           ", vbLeftJustify
        End With
        ggrData.AllFields = mFlds
        ggrData.SetHead
        

        '頻道編號,頻道名稱,收費起始日期,截止日期
        With mFlds2
            .Add "CustId", , , , False, "客戶編號", vbLeftJustify
            .Add "Citemname", , , , False, "    收費項目    ", vbLeftJustify
            .Add "ShouldAmt", , , , False, "出單金額", vbRightJustify
            .Add "RealPeriod", , , , False, "期數", vbLeftJustify
            .Add "RealAmt", , , , False, "實收金額", vbRightJustify
            .Add "RealStartDate", , , , False, "期限起始日", vbLeftJustify
            .Add "RealStopDate", , , , False, "期限截止日", vbLeftJustify
            .Add "CancelFlag", , , , False, "作廢", vbLeftJustify
            .Add "Note", , , , False, Space(20) & "  備註            " & Space(15), vbLeftJustify
        End With
        ggrChild.AllFields = mFlds2
        ggrChild.SetHead
        
        '解碼器編號,解碼器序號,客戶名稱,頻道名稱,設定種類設定時間,設定人員,收費項目
        With mFlds3
            .Add "STBSNo", , , , , "機上盒序號     ", vbLeftJustify
            .Add "SmartCardNo", , , , , "智慧卡序號     ", vbLeftJustify
            .Add "CustId", , , , , "客戶名稱", vbLeftJustify
            .Add "ChName", , , , , "頻道名稱", vbLeftJustify
            .Add "ModeType", , , , , "設定種類     ", vbLeftJustify
            .Add "UpdTime", , , , , "設定時間       ", vbLeftJustify
            .Add "UpdEn", , , , , "設定人員", vbLeftJustify
            .Add "AuthorStartDate", giControlTypeDate, , , , "授權起始日期", vbLeftJustify
            .Add "AuthorStopDate", giControlTypeDate, , , , "授權截止日期", vbLeftJustify
            .Add "ResvTime", giControlTypeTime, , , , "預約時間", vbLeftJustify
        End With
        ggrQuery.AllFields = mFlds3
        ggrQuery.SetHead
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subGrd"
End Sub

Private Sub DefaultValue()
    On Error Resume Next
    Dim strCompCode As String
        If gCompCode <> "" Then
            strCompCode = gCompCode
        Else
            strCompCode = gDefaultComp
        End If
        If strCompCode = "" Then strCompCode = garyGi(9)
        
        gilCompCode.SetCodeNo strCompCode
        gilCompCode.Query_Description
        
        '如有傳CustId ginCustId 不予許Key
        If lngCustId > 0 Then
            ginCustId.Text = lngCustId
            Call ginCustId_Validate(False)
            ginCustId.Enabled = False
        End If
        sstHead.Tab = 1
        If strCaption <> "" Then Me.Caption = Me.Caption & "-" & strCaption
        blnProcessOk = False
        
'        If intProcessType = 0 Then
'            chkSetAll.Enabled = True
'        Else
'            If intByHouse Then
'                chkSetAll.Enabled = True
'            Else
'                chkSetAll.Enabled = False
'            End If
'        End If
        If gLogInID <> "" Then gLogInID = gLogInID & "."
End Sub

Private Sub OpenData(Optional strWhere As String)
    On Error GoTo ChkErr
    Dim strSQL As String
        Set rsData = New ADODB.Recordset
        
        If rsParent Is Nothing Then Set rsParent = New ADODB.Recordset
        If rsParent.State = adStateOpen Then
            Set rsData = rsParent.Clone
        Else
            If ginCustId.Text <> "" Then
                strSQL = " A.CustId = " & ginCustId.Value & " And A.CompCode = " & gilCompCode.GetCodeNo & " And A.FaciSNo is not null And FaciCode In (Select CodeNo From " & GetOwner & "CD022 Where RefNo in (2,5)) And A.PRDate is null "
                If strWhere <> "" Then strSQL = " And " & strWhere
            Else
                strSQL = " A.SeqNo = ''"
            End If
            If Not GetRS(rsData, "Select A.RowId, A.* From " & GetOwner & "So004 A Where " & strSQL & " Order By A.InstDate", , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        End If
        Call OpenData2
        'rsData.Requery
        Set ggrData.Recordset = rsData
        ggrData.Refresh

    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub

Private Sub OpenData2(Optional ByVal strWhere As String)
    On Error GoTo ChkErr
        If rsData.EOF Or rsData.BOF Then
            sstData.Enabled = False
            cmdSet.Enabled = False
            Exit Sub
        Else
            sstData.Enabled = True
            cmdSet.Enabled = True
        End If
        If rsChild Is Nothing Then Set rsChild = New ADODB.Recordset
        If strWhere = "" Then
            strWhere = " CustId = " & rsData("CustId") & " And FaciSeqNo = '" & rsData("SeqNo") & "' And UCCode is not null"
        End If
'        MsgBox "Select RowId,A.* From " & GetOwner & "SO033 A Where " & strWhere
        If Not GetRS(rsChild, "Select RowId,A.* From " & GetOwner & "SO033 A Where " & strWhere) Then Exit Sub
'        MsgBox Err.Description, , Err.Number & " OpenData2"
        If intProcessType = 0 Then Call EnableMode(True)
'        MsgBox "Recordset Open OK ! " & Err.Description, , Err.Number
        Set ggrChild.Recordset = rsChild
'        MsgBox "Set ggrChild.Recordset = rsChild OK ! " & Err.Description, , Err.Number
        ggrChild.Refresh
    Exit Sub
ChkErr:
'    MsgBox Err.Description, , Err.Number
    ErrSub Me.Name, "OpenData2"
End Sub

Private Sub EnableMode(blnFlag As Boolean)
    On Error Resume Next
        optBasic(0).Enabled = GetPremission("SO11711") ' 1. 開機 [02]
        optBasic(1).Enabled = GetPremission("SO11712") ' 2. 關機 [02]
        optBasic(4).Enabled = GetPremission("SO11713") ' 4. CP 申裝 [27]
        optBasic(5).Enabled = GetPremission("SO11714") ' 5. CP 退租 [27]
        
        optCMChange(2).Enabled = GetPremission("SO11715") ' 1. Ping CM [20]
        optCMChange(3).Enabled = GetPremission("SO11716") ' 2. 釋放動態IP [25]
        optCMChange(6).Enabled = GetPremission("SO11717") ' 5. 帳號停權 [24]
        optCMChange(7).Enabled = GetPremission("SO11718") ' 6. 帳號復權 [24]
        optCMChange(4).Enabled = False
        optCMChange(8).Enabled = True
        
        optQueryInfo(0).Enabled = GetPremission("SO11719") ' 1. CM 狀態查詢 [13]
        optQueryInfo(1).Enabled = GetPremission("SO1171A") ' 2. 查詢 HFC 服務類別 [21]
        optQueryInfo(2).Enabled = GetPremission("SO1171B") ' 3. CM 連線紀錄查詢 [22]
        optQueryInfo(3).Enabled = GetPremission("SO1171C") ' 4. CM 連線品質查詢 [23]
        optQueryInfo(4).Enabled = GetPremission("SO1171D") ' 5. PCIP連線記錄查詢 [26]
        
        optBasic(3).Enabled = GetPremission("SO1171E") ' CM Reset [12]
        optCMChange(2).Enabled = GetPremission("SO1171E") ' CM Reset [12]
        optCMChange(8).Enabled = GetPremission("SO1171F") ' 方案類別更換 [02]
        gilBP.Enabled = GetPremission("SO1171F") ' 方案類別更換 [02]
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
'        gcnGi2.Close
'        Set gcnGi2 = Nothing
'        Set objCmd = Nothing
        Call CloseRecordset(rsData)
        Call CloseRecordset(rsChild)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        ReleaseCOM Me
        Set frmSO1623A = Nothing
End Sub

Private Sub ggrChild_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    If giFld.FieldName = "Status" Then
        If Value = 0 Then
            Value = "關"
        Else
            Value = "開"
        End If
    End If
End Sub

Private Sub ggrQuery_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    Select Case UCase(giFld.FieldName)
        Case "MODETYPE"
            Value = GetModeTypeValue(Value)
        Case "OPERRESULT"
            If Value & "" = "1" Then
                Value = "成功"
            ElseIf Value & "" = "2" Then
                Value = "失敗"
            Else
                Value = ""
            End If
    End Select
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error Resume Next
        If giFld.FieldName = "ReFaciSNo" Then
            If Value = "" Then
                Value = "母機"
            Else
                Value = "子機"
            End If
        End If
End Sub

Private Sub GiCombo1_Click()

End Sub

Private Sub gilBP_Change()
  On Error GoTo ChkErr
    Dim strFld As String
    If gilBP.GetCodeNo = Empty Then Exit Sub
    If IsCMCP Then
        strFld = ",CMCPDYNIP,CMCPFIXIP"
    Else
        strFld = ",DYNIPCOUNT,FIXIPCOUNT"
    End If
    Dim varData As Variant
    On Error Resume Next
    varData = gcnGi.Execute("SELECT CMBAUDRATE" & strFld & " FROM " & GetOwner & "CD078D" & _
                                                " WHERE BPCODE = '" & gilBP.GetCodeNo & "'").GetString(adClipString, 1, ",", "", "")
    If Err.Number = 0 Then
        varData = RPxx(CStr(varData))
        varData = Split(varData, ",")
        lblBaudRate = "速率 : " & varData(0)
        lblDynaIP = "動態IP : " & varData(1)
        lblFixIP = "固定IP : " & varData(2)
    Else
        Err.Clear
        lblBaudRate = "速率 : "
        lblDynaIP = "動態IP : "
        lblFixIP = "固定IP : "
    End If
  Exit Sub
ChkErr:
    ErrSub Name, "gilBP_Change"
End Sub

Private Sub gilBP_GotFocus()
  On Error Resume Next
    optCMChange(8).Value = True
End Sub

'CM Console 的方案類別是抓 CD078D 的 BPcode , BPname
'速率是抓 CMBaudRate
'動態IP 或靜態 IP 數則是抓取 DynIPCount / FixIPCount 或 CMCPDynIP / CMCPFixIP

Private Sub gilCompCode_Change()
    On Error Resume Next
        If gilCompCode.GetCodeNo = "" Then Exit Sub
        garyGi(16) = gilCompCode.GetOwner
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then If cmdCancel.Enabled Then cmdCancel.Value = True
    Select Case sstHead.Tab
        Case 0
            If KeyCode = vbKeyF2 Then If cmdSet.Enabled Then cmdSet.Value = True
        Case 1
            If KeyCode = vbKeyF3 Then If cmdQuery.Enabled Then cmdQuery.Value = True
    End Select
End Sub

'(3)於CM基本設定頁籤，增加一checkbox 「□移機中是否取用新地址資訊」…
'◆當勾選: CP申裝[27] 選項時，須檢核CATV服務之派工狀態3若為移機中時(SO002.wipcode3=11)，
'   該checkbox選項自動Enable，且Default勾選。
'(因為目前的CM移機機制，設備已被帶到新址(安裝中)，所以工單即使未完工仍可在新址的客編送指令；
'   但整包移機(CATV+CM+CP一起移時)，若CATV未完工，裝機址仍為舊址，所以需增加此選項供系統抓取新址資訊。
'◆當有勾選該選項，於下所選之指令選項時，須取CATV服務未完工之移機工單上"新地址"資訊，做CM指令開通。
'   (CM開機時,若該EMTA有裝CP,指令會傳送地址資訊)

Private Sub optBasic_Click(Index As Integer)
  On Error Resume Next
    picCPno.Visible = False
    Select Case Index
                Case 0 ' 1. 開機 [02]
                    If optBasic(0).Value Then
                        chkAccStop.Value = 0
                        ShowResv
                    End If
                    chkPR.Enabled = False
                Case 1 ' 2. 關機 [02]
                    If optBasic(1).Value Then
                        ChkAccStart.Value = 0
                        ShowResv
                    End If
                    chkPR.Enabled = False
                Case 4 ' 4. CP 申裝 [27]
                    chkPR.Value = Abs(ChkWipCode)
                    'chkPR.Value = Abs(GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO002" & _
                                                                        " WHERE COMPCODE=" & gilCompCode.GetCodeNo & _
                                                                        " AND CUSTID=" & ginCustId.Text & _
                                                                        " AND WIPCODE3=11", gcnGi, "") > 0)
                    If chkPR.Value Then chkPR.Enabled = True
                    picCPno.Top = optBasic(4).Top + 270
                    picCPno.Visible = True
                    cboCPno.Text = ""
                    cboCPno_DropDown
                    If strCPno = Empty Then
                        cboCPno.Enabled = True
                    Else
                        cboCPno.Text = strCPno
                        cboCPno.Enabled = False
                    End If
                Case 5
                    picCPno.Top = optBasic(5).Top + 270
                    picCPno.Visible = True
                    cboCPno.Text = ""
                    cboCPno_DropDown
                    If strCPno = Empty Then
                        cboCPno.Enabled = True
                    Else
                        cboCPno.Text = strCPno
                        cboCPno.Enabled = False
                    End If
                Case Else
                    HideResv
                    chkPR.Enabled = False
    End Select
End Sub

Private Sub ShowResv()
  On Error Resume Next
    lblResv.Visible = True
    gdtResvTime.Visible = True
End Sub

Private Sub HideResv()
  On Error Resume Next
    lblResv.Visible = False
    gdtResvTime.Visible = False
    gdtResvTime.Text = ""
End Sub

Private Sub optCMChange_Click(Index As Integer)
    If optCMChange(8).Value Then
        ShowResv
    Else
        HideResv
    End If
End Sub

Private Sub optQueryInfo_Click(Index As Integer)
  On Error Resume Next
    Select Case Index
                Case 1 ' 4. CP 申裝 [27]
                    'chkPR2.Value = Abs(GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO002" & _
                                                                        " WHERE COMPCODE=" & gilCompCode.GetCodeNo & _
                                                                        " AND CUSTID=" & ginCustId.Text & _
                                                                        " AND WIPCODE3=11", gcnGi, "") > 0)
                    chkPR2.Value = Abs(ChkWipCode)
                    If chkPR2.Value Then chkPR2.Enabled = True
                Case Else
                    chkPR2.Enabled = False
    End Select
End Sub

' 1.當勾選: CP申裝[27] 選項時，須檢核CATV服務之派工狀態3若為移機中時(SO002.wipcode3=11)，
'   該checkbox選項自動Enable，且Default勾選。
Private Function ChkWipCode() As Boolean
  On Error GoTo ChkErr
    ChkWipCode = GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO002" & _
                                                " WHERE COMPCODE=" & gilCompCode.GetCodeNo & _
                                                " AND CUSTID=" & ginCustId.Text & _
                                                " AND SERVICETYPE='C'" & _
                                                " AND WIPCODE3=11", gcnGi, "") > 0
  Exit Function
ChkErr:
    ErrSub Name, "ChkWipCode"
End Function

Private Sub rsData_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
        If Not ggrData.Enabled Then Exit Sub
        If Not blnSetting Then
            Me.Enabled = False
            Call OpenData2
            Fetch_Log_Data , True
            Me.Enabled = True
        End If
End Sub

Private Sub ginCustId_Validate(Cancel As Boolean)
    On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
        If ginCustId.Text = "" Then
            txtCustName = ""
            Exit Sub
        Else
            If Not GetRS(rs, "Select CustName,InstAddrNo From " & GetOwner & "So001 Where CustId = " & ginCustId.Value & " And CompCode = " & gilCompCode.GetCodeNo, gcnGi, , , , , , , True) Then Exit Sub
            If rs.EOF Then
                txtCustName = ""
                lngInstAddrNo = 0
                lngNodeNo = 0
                lngCircuitNo = 0
            Else
                txtCustName = rs("CustName") & ""
                lngInstAddrNo = Val(rs("InstAddrNo") & "")
                lngNodeNo = 0
                lngCircuitNo = 0
            End If
        End If

        If Not GetRS(rs, "SELECT NODENO,CIRCUITNO FROM " & GetOwner & "SO014 WHERE ADDRNO = " & lngInstAddrNo, gcnGi, , , , , , , True) Then Exit Sub
        If rs.EOF Then
            lngNodeNo = 0
            lngCircuitNo = 0
        Else
            lngNodeNo = Val(rs("NODENO") & "")
            lngCircuitNo = Val(rs("CIRCUITNO") & "")
        End If

        If txtCustName = "" Then
            If ginCustId.Text <> "" Then MsgBox "查無此客戶", vbCritical, gimsgPrompt
            ggrData.Enabled = False
            ggrChild.Enabled = False
            If rsData.RecordCount > 0 Then Call OpenData("1=0")
            ginCustId.Text = ""
            Cancel = True
        Else
            Call OpenData
            ggrData.Enabled = True
            ggrChild.Enabled = True
        End If
        Call CloseRecordset(rs)
    Exit Sub
ChkErr:
    ErrSub Me.Name, "ginCustId_Validate"
End Sub

Public Property Let uCaption(ByVal vData As String)
    strCaption = vData
End Property

Public Property Let uCustId(ByVal vData As Long)
    lngCustId = vData
End Property

Public Property Set uParentRs(ByVal vData As ADODB.Recordset)
    Set rsParent = vData
End Property

Public Property Let uProcessType(ByVal vData As Integer)
    On Error Resume Next
        intProcessType = vData
End Property

Public Property Let uDefaultRowId(ByVal vData As String)
    On Error Resume Next
        strDefaultRowId = vData
End Property

Public Property Let uProcessChStr(ByVal vData As String)
    On Error Resume Next
        strProcessChStr = vData
End Property

Public Property Let uOrderNo(ByVal vData As String)
    On Error Resume Next
        strOrderNo = vData
End Property

Public Property Get uProcessOk() As Boolean
    On Error Resume Next
        uProcessOk = blnProcessOk
End Property

Public Property Let uTransationMode(ByVal vData As Boolean)
    On Error Resume Next
        blnTransation = vData
End Property

Public Property Let uProcessLock(ByVal vData As Integer)
    On Error Resume Next
        blnProcessLock = vData
End Property

Public Property Let uByHouse(ByVal vData As Integer)
    On Error Resume Next
        intByHouse = vData
End Property

Private Function IsDataOk() As Boolean
    On Error GoTo ChkErr
        If gdtResvTime.GetValue <> "" Then
            If CDate(gdtResvTime.GetValue(True)) <= CDate(RightNow) Then
                MsgBox "預約時間需大於現在時間", vbCritical, gimsgPrompt
                gdtResvTime.SetValue DateAdd("h", 1, RightNow)
                gdtResvTime.SetFocus
                Exit Function
            Else
                '7676
'                If CDate(gdtResvTime.GetValue(True)) > CDate(RightNow) + 1 Then
'                    MsgBox "預約時間需小於明日時間", vbCritical, gimsgPrompt
'                    gdtResvTime.SetValue DateAdd("h", 1, RightNow)
'                    gdtResvTime.SetFocus
'                    Exit Function
'                End If
                If MsgBox("有設定預約時間,程式將不會等待CM Command Gateway 回應,確定使用預約功能??", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then Exit Function
            End If
        End If
        
'       Doc : SO1171A_20070109_CM主控台.doc
'        2.  【V1.0】
'        (1) 東森CM指令開通，會依據SO004.BPCODE取CD078D之相關參數來下指令，
'           現因客戶常常會漏選來電業務活動，以致設備裡之促案內容皆為空值，直至下CM開通指令失敗時，
'           方發現資料有誤；故提規格調整…增加SO1171A於送指令之防呆訊息檢核。
        
'        3.  於下指令前須檢核來源資料是否完整，若否則須秀出訊息告知USER，
'           其下述檢核皆無誤方可新增至SO311與網管做介接指令處理…
'        (1) SO004.BPCODE若為空值
'        ' 提出訊息「此設備無組合產品內容，無法正常傳送指令」
      
        Select Case sstData.Tab
                Case 0
                        Select Case True
                                Case optBasic(0)
                                    If rsData("BPCODE") & "" = Empty Then
                                        MsgBox "此設備無組合產品內容，無法正常傳送指令 !", vbInformation, "訊息"
                                        IsDataOk = False
                                        Exit Function
                                    End If
                                Case optBasic(1)
                                    If rsData("BPCODE") & "" = Empty Then
                                        MsgBox "此設備無組合產品內容，無法正常傳送指令 !", vbInformation, "訊息"
                                        IsDataOk = False
                                        Exit Function
                                    End If
                                Case optBasic(4)
                                    If Trim(cboCPno.Text) = Empty Then
                                        MsgBox "請選擇 CP門號 !", vbInformation, "訊息"
                                        IsDataOk = False
                                        Exit Function
                                    End If
                                Case optBasic(5)
                                    If Trim(cboCPno.Text) = Empty Then
                                        MsgBox "請選擇 CP門號 !", vbInformation, "訊息"
                                        IsDataOk = False
                                        Exit Function
                                    End If
                                Case Else
                        End Select
                Case Else
        End Select
'        底下是後來加的東東

'        (2) 當為EMTA設備時(CD022.REFNO=5)，其MTAMAC必須有值(SO306.MTAMAC)若無則…
'        ' 提出訊息「此EMTA設備無MTAMAC，無法正常傳送指令」
        
        
        
        ' ------------------------------------------------------------------------------------------------------------------
        
        ' Date : 2007/05/04
        
        ' 送指令時，如果該設備不在SO306就不去稽核有沒有MTAMAC，
        ' 請顯示SO306[EMTA資料檔]中查無HFCMAC為[XXXXXX]的資料，後續仍讓系統可繼續送指令，
        ' 如果此設備序號存在SO306中，但MTAMAC無值，就顯示"此EMTA設備無MTAMAC，無法正常傳送指令"。
        
        ' Doc : SO1171A_20070503_CM主控台.doc
        
        If frmSO1623A.IsCMCP Then
        
            If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO306" & _
                                    " WHERE HFCMAC='" & rsData("FaciSNo") & "'", gcnGi) = 0 Then
                MsgBox "SO306 [EMTA資料檔] 中查無 HFCMAC 為 [" & rsData("FaciSNo") & "] 的資料!", vbInformation, "訊息"
            Else
                If Get_EMTA_Mac(rsData("FaciSNo") & "") = Empty Then
                    MsgBox "此EMTA設備無MTAMAC，無法正常傳送指令 !", vbInformation, "訊息"
                    IsDataOk = False
                    Exit Function
                End If
            End If
        
        End If
        
        ' ------------------------------------------------------------------------------------------------------------------

'        Select Case sstData.Tab
'            Case 0
'                Select Case True
'                    Case optOpen
'                End Select
'            Case 1
'                Select Case True
'                    Case optOpenOldCh
'                        If Not MustExist(gimOpenOldCh, 5, "重開已開頻道") Then Exit Function
'                End Select
'            Case 5
'                Select Case True
'                    Case optSendMsg
'                        If Not MustExist(txtMessage, , "訊息內容") Then Exit Function
'                End Select
'        End Select
        IsDataOk = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "IsDataOk"
End Function

Private Sub sstData_Click(PreviousTab As Integer)
  On Error Resume Next
    Select Case sstData.Tab
        Case 2
            chkSetAll.Value = 0
            chkSetAll.Visible = False
    End Select
    HideResv
End Sub

Private Function GetPCMac() As String
    On Error Resume Next
'        GetPCMac = gcnGi.Execute("Select CPEMAC From " & GetOwner & "SO004C Where FaciSNo = '" & rsData("FaciSNo") & "' And CustId = " & rsData("CustId") & " And RowNum=1").GetString(, , , ",")
'        If GetPCMac = "," Then GetPCMac = ""
    GetPCMac = GetRsValue("Select CPEMAC From " & GetOwner & "SO004C Where FaciSNo = '" & rsData("FaciSNo") & "' And CustId = " & rsData("CustId") & " And RowNum=1", gcnGi, "") & ""
End Function

