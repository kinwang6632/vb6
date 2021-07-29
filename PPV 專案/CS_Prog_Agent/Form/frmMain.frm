VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{898CF0BC-AF45-45AC-9831-A27FC475455D}#1.0#0"; "WinXPctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '沒有框線
   Caption         =   " 開博科技 PPV 節目表 Agent "
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10080
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   10080
   StartUpPosition =   2  '螢幕中央
   Tag             =   "7095,10110"
   Visible         =   0   'False
   Begin WinXP_Engine.WinXPctl xpc 
      Left            =   5550
      Top             =   30
      _ExtentX        =   3995
      _ExtentY        =   873
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8700
      Top             =   30
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   6105
      Left            =   150
      TabIndex        =   0
      Top             =   555
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10769
      _Version        =   393216
      TabHeight       =   635
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BackColor       =   16777215
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "監 控 管 理 (&V)"
      TabPicture(0)   =   "frmMain.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl監控管理(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl系統參數(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl權限管理(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgConsole"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraConsole"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraMonitor"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraTimer"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lstXML"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "系 統 參 數 (&P)"
      TabPicture(1)   =   "frmMain.frx":0E4C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl系統參數(0)"
      Tab(1).Control(1)=   "lbl監控管理(1)"
      Tab(1).Control(2)=   "lbl權限管理(2)"
      Tab(1).Control(3)=   "fraSetup"
      Tab(1).Control(4)=   "fraDatabase"
      Tab(1).Control(5)=   "cmdEditSysPara"
      Tab(1).Control(6)=   "cmdUpdateSysPara"
      Tab(1).Control(7)=   "cmdUndoSysPara"
      Tab(1).Control(8)=   "fraDEX"
      Tab(1).Control(9)=   "fraEPG"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "權 限 管 理 (&G)"
      TabPicture(2)   =   "frmMain.frx":13CE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl權限管理(0)"
      Tab(2).Control(1)=   "lbl監控管理(2)"
      Tab(2).Control(2)=   "lbl系統參數(2)"
      Tab(2).Control(3)=   "fraPassword"
      Tab(2).Control(4)=   "cmdUndoPermission"
      Tab(2).Control(5)=   "cmdUpdatePermission"
      Tab(2).Control(6)=   "cmdEditPermission"
      Tab(2).ControlCount=   7
      Begin VB.ListBox lstXML 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   180
         TabIndex        =   91
         Top             =   4555
         Width           =   9375
      End
      Begin VB.Frame fraEPG 
         BackColor       =   &H00FFFFFF&
         Caption         =   " EPG Wrapper 路 徑 設 定 "
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
         ForeColor       =   &H00FF0000&
         Height          =   2430
         Left            =   -70095
         TabIndex        =   66
         Top             =   2760
         Width           =   4575
         Begin VB.TextBox txtEPGlstImpTime 
            Appearance      =   0  '平面
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            IMEMode         =   3  '暫止
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   88
            Top             =   1860
            Width           =   2175
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   ".."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   4210
            Style           =   1  '圖片外觀
            TabIndex        =   75
            ToolTipText     =   "選擇資料夾"
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   ".."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   4210
            Style           =   1  '圖片外觀
            TabIndex        =   74
            ToolTipText     =   "選擇資料夾"
            Top             =   900
            Width           =   255
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   ".."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   4210
            Style           =   1  '圖片外觀
            TabIndex        =   73
            ToolTipText     =   "選擇資料夾"
            Top             =   1380
            Width           =   255
         End
         Begin VB.TextBox txt_EPG_Src_Path 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   840
            TabIndex        =   13
            Top             =   330
            Width           =   3345
         End
         Begin VB.TextBox txt_EPG_Err_Path 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   840
            TabIndex        =   14
            Top             =   840
            Width           =   3345
         End
         Begin VB.TextBox txt_EPG_Dst_Path 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            IMEMode         =   3  '暫止
            Left            =   840
            TabIndex        =   15
            Top             =   1350
            Width           =   3345
         End
         Begin VB.Label lblWrapLstImtTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "Wrap節目表最新匯入時間"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   90
            TabIndex        =   89
            Top             =   1900
            Width           =   2130
         End
         Begin VB.Label lbl_EPG_Err_Path 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "問題位置"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   69
            Top             =   900
            Width           =   720
         End
         Begin VB.Label lbl_EPG_Dst_Path 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "目的位置"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   68
            Top             =   1410
            Width           =   720
         End
         Begin VB.Label lbl_EPG_Src_Path 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "來源位置"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   67
            Top             =   405
            Width           =   720
         End
      End
      Begin VB.Frame fraDEX 
         BackColor       =   &H00FFFFFF&
         Caption         =   " DEX 路 徑 設 定 "
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
         ForeColor       =   &H00FF0000&
         Height          =   2430
         Left            =   -74745
         TabIndex        =   62
         Top             =   2760
         Width           =   4575
         Begin VB.TextBox txtDEXlstImpTime 
            Appearance      =   0  '平面
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            IMEMode         =   3  '暫止
            Left            =   2220
            Locked          =   -1  'True
            TabIndex        =   86
            Top             =   1860
            Width           =   2235
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   ".."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   4210
            Style           =   1  '圖片外觀
            TabIndex        =   72
            ToolTipText     =   "選擇資料夾"
            Top             =   1380
            Width           =   255
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   ".."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   4210
            Style           =   1  '圖片外觀
            TabIndex        =   71
            ToolTipText     =   "選擇資料夾"
            Top             =   900
            Width           =   255
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   ".."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   4210
            Style           =   1  '圖片外觀
            TabIndex        =   70
            ToolTipText     =   "選擇資料夾"
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox txt_DEX_Dst_Path 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            IMEMode         =   3  '暫止
            Left            =   840
            TabIndex        =   12
            Top             =   1350
            Width           =   3345
         End
         Begin VB.TextBox txt_DEX_Err_Path 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   840
            TabIndex        =   11
            Top             =   840
            Width           =   3345
         End
         Begin VB.TextBox txt_DEX_Src_Path 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   840
            TabIndex        =   10
            Top             =   330
            Width           =   3345
         End
         Begin VB.Label lblDexLstImtTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "DEX節目表最新匯入時間"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   90
            TabIndex        =   87
            Top             =   1900
            Width           =   1995
         End
         Begin VB.Label lbl_DEX_Src_Path 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "來源位置"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   65
            Top             =   405
            Width           =   720
         End
         Begin VB.Label lbl_DEX_Dst_Path 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "目的位置"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   64
            Top             =   1410
            Width           =   720
         End
         Begin VB.Label lbl_DEX_Err_Path 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "問題位置"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   63
            Top             =   900
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdEditPermission 
         Caption         =   "修 改 (&E)"
         Height          =   375
         Left            =   -73830
         Style           =   1  '圖片外觀
         TabIndex        =   22
         Top             =   4950
         Width           =   1125
      End
      Begin VB.CommandButton cmdUpdatePermission 
         Caption         =   "存 檔 (&S)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70635
         Style           =   1  '圖片外觀
         TabIndex        =   23
         Top             =   4950
         Width           =   1125
      End
      Begin VB.CommandButton cmdUndoPermission 
         Caption         =   "取 消 (&C)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67530
         Style           =   1  '圖片外觀
         TabIndex        =   24
         Top             =   4950
         Width           =   1125
      End
      Begin VB.CommandButton cmdUndoSysPara 
         Caption         =   "取 消 (&C)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67590
         Style           =   1  '圖片外觀
         TabIndex        =   18
         Top             =   5400
         Width           =   1125
      End
      Begin VB.CommandButton cmdUpdateSysPara 
         Caption         =   "存 檔 (&S)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70695
         Style           =   1  '圖片外觀
         TabIndex        =   17
         Top             =   5400
         Width           =   1125
      End
      Begin VB.CommandButton cmdEditSysPara 
         Caption         =   "修 改 (&E)"
         Height          =   375
         Left            =   -73815
         Style           =   1  '圖片外觀
         TabIndex        =   16
         Top             =   5400
         Width           =   1125
      End
      Begin VB.Frame fraTimer 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 狀 態 "
         ForeColor       =   &H00FF0000&
         Height          =   540
         Left            =   150
         TabIndex        =   41
         Top             =   5460
         Width           =   9435
         Begin VB.Label lblPeriod 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "時 段"
            ForeColor       =   &H00400000&
            Height          =   180
            Left            =   5235
            TabIndex        =   50
            Top             =   210
            Width           =   435
         End
         Begin VB.Label lblTime 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  '單線固定
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   5745
            TabIndex        =   49
            Top             =   195
            Width           =   975
         End
         Begin VB.Label lblAction 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  '單線固定
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   3675
            TabIndex        =   48
            Top             =   195
            Width           =   1395
         End
         Begin VB.Label lblProcess 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "動 作"
            ForeColor       =   &H00400000&
            Height          =   180
            Left            =   3165
            TabIndex        =   47
            Top             =   210
            Width           =   435
         End
         Begin VB.Label lblCountDown 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  '單線固定
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   7395
            TabIndex        =   46
            Top             =   195
            Width           =   1815
         End
         Begin VB.Label lblDuration 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "倒 數"
            ForeColor       =   &H00400000&
            Height          =   180
            Left            =   6885
            TabIndex        =   45
            Top             =   210
            Width           =   435
         End
         Begin VB.Label lblState 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "狀 態"
            ForeColor       =   &H00400000&
            Height          =   180
            Left            =   195
            TabIndex        =   44
            Top             =   210
            Width           =   435
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  '單線固定
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   705
            TabIndex        =   43
            Top             =   195
            Width           =   2295
         End
      End
      Begin VB.Frame fraMonitor 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 即 時 監 控 "
         ForeColor       =   &H00FF0000&
         Height          =   3390
         Left            =   150
         TabIndex        =   40
         Top             =   1110
         Width           =   9435
         Begin VB.ListBox lstLog 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1110
            Left            =   4740
            TabIndex        =   90
            Top             =   1905
            Width           =   4575
         End
         Begin VB.ListBox lstErr 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1110
            Left            =   6540
            TabIndex        =   83
            Top             =   495
            Width           =   2775
         End
         Begin VB.DirListBox dirErr 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1050
            Left            =   4740
            TabIndex        =   82
            Top             =   500
            Width           =   1755
         End
         Begin VB.DirListBox dirDst 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1050
            Left            =   90
            TabIndex        =   81
            Top             =   1890
            Width           =   1755
         End
         Begin VB.DirListBox dirSrc 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1050
            Left            =   90
            TabIndex        =   80
            Top             =   500
            Width           =   1755
         End
         Begin VB.ListBox lstDst 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1110
            Left            =   1890
            TabIndex        =   77
            Top             =   1905
            Width           =   2810
         End
         Begin VB.ListBox lstSrc 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1110
            Left            =   1890
            TabIndex        =   76
            Top             =   495
            Width           =   2810
         End
         Begin MSComctlLib.ProgressBar pgb 
            Height          =   255
            Left            =   90
            TabIndex        =   42
            Top             =   3015
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblPrc 
            Appearance      =   0  '平面
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  '單線固定
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4710
            TabIndex        =   85
            Top             =   1605
            Width           =   4620
         End
         Begin VB.Label lblErr 
            Appearance      =   0  '平面
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  '單線固定
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4710
            TabIndex        =   84
            Top             =   210
            Width           =   4620
         End
         Begin VB.Label lblDst 
            Appearance      =   0  '平面
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  '單線固定
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   90
            TabIndex        =   79
            Top             =   1605
            Width           =   4650
         End
         Begin VB.Label lblSrc 
            Appearance      =   0  '平面
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  '單線固定
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   90
            TabIndex        =   78
            Top             =   210
            Width           =   4650
         End
      End
      Begin VB.Frame fraConsole 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 啟 動 命 令 "
         ForeColor       =   &H00FF0000&
         Height          =   660
         Left            =   150
         TabIndex        =   39
         Top             =   390
         Width           =   9435
         Begin VB.CheckBox chkShowXML 
            BackColor       =   &H00FFFFFF&
            Caption         =   "顯示 XML"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5790
            TabIndex        =   92
            Top             =   200
            Width           =   1125
         End
         Begin VB.CommandButton cmdStop 
            Caption         =   "停 止 (&X)"
            Enabled         =   0   'False
            Height          =   345
            Left            =   7320
            Style           =   1  '圖片外觀
            TabIndex        =   2
            Top             =   195
            Width           =   1125
         End
         Begin VB.CommandButton cmdStart 
            Caption         =   "啟 動 (&R)"
            Height          =   345
            Left            =   990
            Style           =   1  '圖片外觀
            TabIndex        =   1
            Top             =   195
            Width           =   1125
         End
         Begin VB.Label lblStartTime 
            Appearance      =   0  '平面
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   3660
            TabIndex        =   52
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   1845
         End
         Begin VB.Label lblExecTime 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "執行時間 :"
            ForeColor       =   &H00400000&
            Height          =   180
            Left            =   2595
            TabIndex        =   51
            Top             =   270
            Width           =   825
         End
      End
      Begin VB.Frame fraDatabase 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 共 用 區 資 料 庫 連 線 設 定 "
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
         ForeColor       =   &H00FF0000&
         Height          =   1920
         Left            =   -70095
         TabIndex        =   32
         Top             =   675
         Width           =   4575
         Begin VB.TextBox txtDSN 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1650
            TabIndex        =   7
            Top             =   330
            Width           =   2565
         End
         Begin VB.TextBox txtUID 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1650
            TabIndex        =   8
            Top             =   840
            Width           =   2565
         End
         Begin VB.TextBox txtPWD 
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IMEMode         =   3  '暫止
            Left            =   1650
            PasswordChar    =   "*"
            TabIndex        =   9
            Top             =   1350
            Width           =   2565
         End
         Begin VB.Label lblUID 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "使 用 者"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   870
            TabIndex        =   35
            Top             =   900
            Width           =   660
         End
         Begin VB.Label lblPWD 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "密 碼"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1110
            TabIndex        =   34
            Top             =   1410
            Width           =   420
         End
         Begin VB.Label lblDSN 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "資 料 庫 來 源"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   390
            TabIndex        =   33
            Top             =   420
            Width           =   1140
         End
      End
      Begin VB.Frame fraSetup 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 系 統 設 定 "
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
         ForeColor       =   &H00FF0000&
         Height          =   1920
         Left            =   -74745
         TabIndex        =   28
         Top             =   675
         Width           =   4575
         Begin MSMask.MaskEdBox mskPeakTime1 
            Height          =   345
            Left            =   1560
            TabIndex        =   3
            Top             =   240
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtPeakTime 
            Alignment       =   1  '靠右對齊
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1560
            TabIndex        =   5
            Text            =   "0"
            Top             =   720
            Width           =   705
         End
         Begin VB.TextBox txtOffPeak 
            Alignment       =   1  '靠右對齊
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1560
            TabIndex        =   6
            Text            =   "0"
            Top             =   1200
            Width           =   705
         End
         Begin MSMask.MaskEdBox mskPeakTime2 
            Height          =   345
            Left            =   2580
            TabIndex        =   4
            Top             =   240
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblPeakTime 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "尖 峰 時 段 每              秒 讀 取 一 次 資 料"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   360
            TabIndex        =   31
            Top             =   780
            Width           =   3960
         End
         Begin VB.Label lblOffPeak 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "離 峰 時 段 每              秒 讀 取 一 次 資 料"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   360
            TabIndex        =   30
            Top             =   1260
            Width           =   3600
         End
         Begin VB.Label lblPeakTimeDefine 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "尖 峰 時 段 定 義              至              ( 24 小時制 )"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   300
            Width           =   4320
         End
      End
      Begin VB.Frame fraPassword 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 密   碼   設   定 "
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   3195
         Left            =   -73860
         TabIndex        =   25
         Top             =   1020
         Width           =   7485
         Begin VB.TextBox txtPassword 
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            IMEMode         =   3  '暫止
            Left            =   2610
            PasswordChar    =   "*"
            TabIndex        =   20
            ToolTipText     =   " 請輸入密碼 ! "
            Top             =   1470
            Width           =   3225
         End
         Begin VB.TextBox txtCfmPwd 
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            IMEMode         =   3  '暫止
            Left            =   2610
            PasswordChar    =   "*"
            TabIndex        =   21
            ToolTipText     =   " 確認密碼必須與輸入密碼相同 !"
            Top             =   2160
            Width           =   3225
         End
         Begin VB.CheckBox chkUsePwd 
            Appearance      =   0  '平面
            BackColor       =   &H00FFFFFF&
            Caption         =   " 使 用 密 碼 啟 動 系 統"
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   1635
            MaskColor       =   &H00FF0000&
            TabIndex        =   19
            ToolTipText     =   " 打勾表示使用密碼 , 反之 則無"
            Top             =   690
            UseMaskColor    =   -1  'True
            Width           =   2145
         End
         Begin VB.Image imgKey 
            Appearance      =   0  '平面
            Height          =   765
            Left            =   4590
            Picture         =   "frmMain.frx":1950
            Stretch         =   -1  'True
            Top             =   390
            Width           =   675
         End
         Begin VB.Label lblPassword 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "輸 入 密 碼"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1635
            TabIndex        =   27
            Top             =   1590
            Width           =   915
         End
         Begin VB.Label lblCfmPwd 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "確 認 密 碼"
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   1635
            TabIndex        =   26
            Top             =   2280
            Width           =   855
         End
      End
      Begin VB.Image imgConsole 
         Appearance      =   0  '平面
         Height          =   5700
         Left            =   60
         Picture         =   "frmMain.frx":264E
         Stretch         =   -1  'True
         Tag             =   "6880,5700"
         Top             =   375
         Width           =   9645
      End
      Begin VB.Label lbl權限管理 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "權 限 管 理 (&G)"
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   2
         Left            =   -67465
         TabIndex        =   61
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label lbl權限管理 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "權 限 管 理 (&G)"
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   1
         Left            =   7535
         TabIndex        =   60
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label lbl系統參數 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "系 統 參 數 (&P)"
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   2
         Left            =   -70710
         TabIndex        =   59
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label lbl系統參數 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "系 統 參 數 (&P)"
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   1
         Left            =   4290
         TabIndex        =   58
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label lbl監控管理 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "監 控 管 理 (&V)"
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   2
         Left            =   -73955
         TabIndex        =   57
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label lbl監控管理 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "監 控 管 理 (&V)"
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   1
         Left            =   -73955
         TabIndex        =   56
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label lbl監控管理 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "監 控 管 理 (&V)"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   1045
         TabIndex        =   55
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label lbl權限管理 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "權 限 管 理 (&G)"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   -67465
         TabIndex        =   54
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label lbl系統參數 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "系 統 參 數 (&P)"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   -70710
         TabIndex        =   53
         Top             =   120
         Width           =   1125
      End
   End
   Begin VB.Image imgBG 
      Appearance      =   0  '平面
      Height          =   315
      Left            =   5220
      Picture         =   "frmMain.frx":EF78
      Stretch         =   -1  'True
      Top             =   60
      Width           =   90
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  '單線固定
      Caption         =   "[ 監控管理 ] 主要功能為 [啟動] 、 [停止] 系統 ! 並可即時監控執行狀況 .."
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   0
      TabIndex        =   38
      Top             =   6840
      Width           =   10065
   End
   Begin VB.Image imgIcon 
      Height          =   330
      Left            =   90
      Picture         =   "frmMain.frx":8F02A
      Stretch         =   -1  'True
      Top             =   30
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgMaxButton 
      Height          =   285
      Left            =   2100
      Picture         =   "frmMain.frx":8F8F4
      Stretch         =   -1  'True
      Top             =   7470
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgExit 
      Height          =   285
      Left            =   2760
      Picture         =   "frmMain.frx":8FE78
      Stretch         =   -1  'True
      Top             =   7470
      Width           =   285
   End
   Begin VB.Image imgBottom 
      Height          =   45
      Left            =   2400
      Picture         =   "frmMain.frx":903FC
      Stretch         =   -1  'True
      Top             =   7860
      Width           =   630
   End
   Begin VB.Image imgRight 
      Height          =   375
      Left            =   3480
      Picture         =   "frmMain.frx":905C0
      Stretch         =   -1  'True
      Top             =   7500
      Width           =   45
   End
   Begin VB.Image imgLeft 
      Height          =   495
      Left            =   3360
      Picture         =   "frmMain.frx":9085C
      Stretch         =   -1  'True
      Top             =   7500
      Width           =   45
   End
   Begin VB.Image imgLogOff 
      Height          =   360
      Index           =   1
      Left            =   690
      Picture         =   "frmMain.frx":90AF8
      Top             =   8880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgLogOff 
      Height          =   360
      Index           =   0
      Left            =   90
      Picture         =   "frmMain.frx":911FC
      Top             =   8880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgShutDown 
      Height          =   360
      Index           =   1
      Left            =   690
      Picture         =   "frmMain.frx":91900
      Top             =   8520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgShutDown 
      Height          =   360
      Index           =   0
      Left            =   90
      Picture         =   "frmMain.frx":92004
      Top             =   8520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgDimmed 
      Height          =   315
      Index           =   2
      Left            =   810
      Picture         =   "frmMain.frx":92708
      Top             =   8160
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDimmed 
      Height          =   315
      Index           =   1
      Left            =   450
      Picture         =   "frmMain.frx":92C8C
      Top             =   8160
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDimmed 
      Height          =   315
      Index           =   0
      Left            =   90
      Picture         =   "frmMain.frx":93210
      Top             =   8160
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgEnabled 
      Height          =   315
      Index           =   2
      Left            =   810
      Picture         =   "frmMain.frx":93794
      Top             =   7800
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgEnabled 
      Height          =   315
      Index           =   1
      Left            =   450
      Picture         =   "frmMain.frx":93D18
      Top             =   7800
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgEnabled 
      Height          =   315
      Index           =   0
      Left            =   90
      Picture         =   "frmMain.frx":9429C
      Top             =   7800
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDisabled 
      Height          =   315
      Index           =   2
      Left            =   810
      Picture         =   "frmMain.frx":94820
      Top             =   7440
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDisabled 
      Height          =   315
      Index           =   1
      Left            =   450
      Picture         =   "frmMain.frx":94DA4
      Top             =   7440
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDisabled 
      Height          =   315
      Index           =   0
      Left            =   90
      Picture         =   "frmMain.frx":95328
      Top             =   7440
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgMinButton 
      Height          =   285
      Left            =   2430
      Picture         =   "frmMain.frx":958AC
      Stretch         =   -1  'True
      Top             =   7470
      Width           =   285
   End
   Begin VB.Image imgTopRight 
      Height          =   390
      Left            =   9840
      Picture         =   "frmMain.frx":95E30
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   " 開博科技 PPV 節目表 Agent "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   180
      Left            =   480
      TabIndex        =   37
      Top             =   120
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Image imgTopLeft 
      Height          =   390
      Left            =   0
      Picture         =   "frmMain.frx":962FC
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgTop 
      Height          =   390
      Left            =   210
      Picture         =   "frmMain.frx":967C8
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblCaptionShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   " 開博科技 PPV 節目表 Agent "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   2700
      TabIndex        =   36
      Top             =   120
      Visible         =   0   'False
      Width           =   2205
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PR : Hammer
' Start date : 2005/06/18
' Last Modify : 2005/08/02

Option Explicit

' 宣告
Private intMouseX As Integer
Private intMouseY As Integer

Private stm As New ADODB.Stream

Private clsFS As New clsFileSearch ' 檔案搜尋物件類別模組

Private lngTmr As Long
Private sglSecond As Single
Private blnStarting As Boolean
Private blnStopFlag As Boolean
Private blnExecuting As Boolean
Private strTraceMsg As String
Private blnInitialOK As Boolean
Private intFileCount As Integer
'Private cltFile As New Collection ' 集合物件
Private strFile As String

Private rsCOMM As New ADODB.Recordset

'2.  程式需能做到自動偵測：
'(1) 開發欄位可供設定約多久程式會自動去撈取來源資料夾是否有檔案。
'(2) 若於來源位置有檔案存在，則將該路徑下的所有檔案，剪走至問題位置存放做拆解。
'(3) 若已成功拆解完成的檔案，則將其剪走存放於最後的目的位置，若失敗的檔案則仍保留於問題位置。

Private Sub Form_Load() ' 表單載入事件
  On Error GoTo ChkErr
    If Not Visible Then XPstyle
    With xpc
        .TextControl = False
        .ColorScheme = System
        .InitSubClassing
    End With
    Refresh
    blnStarting = False
    blnStopFlag = False
    blnExecuting = False
    If Len(mskPeakTime1) = 0 Then mskPeakTime1 = "0000"
    If Len(mskPeakTime2) = 0 Then mskPeakTime2 = "0000"
    dirSrc.Path = "C:\"
    dirDst.Path = "C:\"
    dirErr.Path = "C:\"
    Open_Virtual_RS
    If Not GetXMLobj(objDOM) Then
'        MsgBox "Can't create M$ DOM object !", vbInformation, "Message"
        MsgBox "無法建立 MS DOM 物件 (文件物件模型) ! 請洽開博科技客服人員 ..", vbInformation, "訊息"
        End
    End If
  Exit Sub
ChkErr:
    ErrHandle Name, "Load Event : Form_Load"
End Sub

Private Sub cmdStart_Click() ' 啟動SPM系統介接 Gateway
  On Error GoTo ChkErr
    lstLog.Clear
    If Not ChkParaOK Then Exit Sub
    cmdStart.Enabled = False
    cmdStop.Enabled = True
    blnStarting = True
    blnStopFlag = False
    lblSrc = " [ 來源位置 ] "
    lblErr = " [ 問題位置 ] "
    lblDst = " [ 目的位置 ] "
    lblPrc = " [ 處理進度 ] "
    System_GoOn
  Exit Sub
ChkErr:
    ErrHandle Name, "Subroutine : cmdStart_Click"
End Sub

Private Function ChkParaOK() As Boolean ' 檢查參數是否 OK
  On Error GoTo ChkErr
    
    lstLog.AddItem "檢核系統參數設定.."
    lstLog.ListIndex = lstLog.ListCount - 1
    Read_Sys_Para
    
    ChkParaOK = False
    
    If Len(mskPeakTime1) <= 0 Then GoTo Period_Err
    If Len(mskPeakTime2) <= 0 Then GoTo Period_Err
    If Len(txtPeakTime) <= 0 Then GoTo Period_Err
    If Len(txtOffPeak) <= 0 Then GoTo Period_Err
    
    If Len(txtDSN) <= 0 Then GoTo DB_SID_Err
    If Len(txtUID) <= 0 Then GoTo DB_SID_Err
    If Len(txtPWD) <= 0 Then GoTo DB_SID_Err
    
    If OnLine Then
        If Not OpenCN Then GoTo DB_SID_Err
    End If
    
    If Len(txt_DEX_Src_Path) <= 0 Then GoTo DEX_Err
    If Len(txt_DEX_Err_Path) <= 0 Then GoTo DEX_Err
    If Len(txt_DEX_Dst_Path) <= 0 Then GoTo DEX_Err
  
    If Len(txt_EPG_Src_Path) <= 0 Then GoTo EPG_Err
    If Len(txt_EPG_Err_Path) <= 0 Then GoTo EPG_Err
    If Len(txt_EPG_Dst_Path) <= 0 Then GoTo EPG_Err
    
    ChkParaOK = True
  
  Exit Function

EPG_Err:
    tabMain.Tab = 1
    MsgBox "[ 系統參數 ] 中 [ EPG 路徑設定 ] 不正確 ! 請確認 ..", vbInformation, "訊息"
  Exit Function

DEX_Err:
    tabMain.Tab = 1
    MsgBox "[ 系統參數 ] 中 [ DEX 路徑設定 ] 不正確 ! 請確認 ..", vbInformation, "訊息"
  Exit Function

DB_SID_Err:
    tabMain.Tab = 1
    MsgBox "[ 系統參數 ] 中 [ 共用區資料庫連線設定 ] 不正確 ! 請確認 ..", vbInformation, "訊息"
  Exit Function

Period_Err:
    tabMain.Tab = 1
    MsgBox "[ 系統參數 ] 中 [ 系統設定 ] 不正確 ! 請確認 ..", vbInformation, "訊息"
  Exit Function
    
ChkErr:
    ErrHandle Name, "ChkParaOK"
End Function

Private Function OpenCN() As Boolean ' 建立共用區的資料庫連線
  On Error GoTo ChkErr
    OpenCN = False
    If Len(Trim(txtDSN)) = 0 Then txtDSN = Decrypt(Read_From_INI("DB_SID", "KeyA"))
    If Len(Trim(txtUID)) = 0 Then txtUID = Decrypt(Read_From_INI("DB_SID", "KeyB"))
    If Len(Trim(txtPWD)) = 0 Then txtPWD = Decrypt(Read_From_INI("DB_SID", "KeyC"))
    If Len(Trim(txtDSN)) = 0 Then MsgBox "請至 [系統參數] 設定 [資料庫來源] !", vbInformation, "訊息": Exit Function
    If Len(Trim(txtUID)) = 0 Then MsgBox "請至 [系統參數] 設定 [使用者] !", vbInformation, "訊息": Exit Function
    If Len(Trim(txtPWD)) = 0 Then MsgBox "請至 [系統參數] 設定 [密碼] !", vbInformation, "訊息": Exit Function
    With cnCOMM
         If .State <= 0 Then
            .CursorLocation = adUseClient
            .Open "Provider=MSDAORA.1;Password=" & txtPWD & ";User ID=" & txtUID & ";Data Source=" & txtDSN & ";Persist Security Info=True"
        End If
        OpenCN = (.State > 0)
    End With
    lstLog.AddItem "共用區資料庫連線建立完成!"
    lstLog.ListIndex = lstLog.ListCount - 1
  Exit Function
ChkErr:
    strTraceMsg = strTraceMsg & Err.Description & " ( " & Err.Number & " )" & vbCrLf
    ErrHandle Name, "Function : OpenCN"
End Function

Private Sub LookInto() ' 供 [計時器] 呼叫 , 準備處理 XML 檔案 ( Timer 呼叫進入點 )
  On Error GoTo ChkErr
    lstLog.AddItem "檢查目錄是否有待處理之 XML 檔案 .."
    lstLog.ListIndex = lstLog.ListCount - 1
'    If Not Process_XML Then Exit Sub ' 處理 DEX XML 檔案
    DoEvents
    Sleep 500
    DoEvents
'    If Not Process_XML(True) Then Exit Sub ' 處理 EPG XML 檔案
    Process_SO160 ' Process SO160
  Exit Sub
ChkErr:
    ErrHandle Name, "Subroutine : LookInto"
End Sub

'7.  從DEX XML資料檔(SO158)、EPG XML資料檔(SO159)撈取出有對應到的資料：
'    (1) 先將XML處理暫存檔(SO160)清除
'    (2) Select SO158、SO159與SO162資料表，取有對應到的資料insert至SO160。
'           SO158.SeqNo is null and SO159.SeqNo is null
'           And SO158.EventID=SO159.EventID
'           And SO159.ChannelID=SO162.ChannelID的資料
'    (3) 請依照【P7 Schema】SO160資料表處理規則。
Private Sub Process_SO160()
  On Error GoTo ChkErr
  
    Dim str3in1 As String
    Dim strInsert160 As String
    Dim lngRcd As Long
    
    cnCOMM.Execute "DELETE FROM SO160"
    
    str3in1 = "SELECT A.DFILENAME,A.DCREATIONDATE,A.PRODUCTID,A.IMPULSIVEFLAG,A.PRODUCTKIND," & _
                    "A.SPECIALPPV,A.PURCHASABLEFLAG,A.SOLDFLAG,A.SUSPENDEDFLAG," & _
                    "A.PRODUCTNAME,A.PRODUCTDESCRIPTION,A.ENGPRODUCTNAME,A.ENGPRODUCTDESCRIPTION," & _
                    "A.SALEBEGINTIME,A.SALEENDTIME,A.VALIDITYPERIODBEGINTIME,A.VALIDITYPERIODENDTIME," & _
                    "A.PRODUCTCATEGORY,A.EPGPRICE,A.MONEYUNIT,A.SUBALLOWPPVEVENTFLAG," & _
                    "A.EVENTPREVIEWTIME,A.DPARENTALRATING," & _
                    "B.EFILENAME,B.ECREATIONDATE,B.CHANNELID,B.EVENTBEGINTIME,B.DURATION,B.EVENTID," & _
                    "B.EVENTTYPE,B.NAME,B.EPGDESC,B.ENGNAME,B.ENGEPGDESC,B.PARENTALRATING," & _
                    "B.CONTENTNIBBLE1,B.CONTENTNIBBLE2,B.USERNIBBLE1,B.USERNIBBLE2," & _
                    "C.CHANNELNO," & _
                    "'" & Format(Now, "ee/mm/dd hh:mm:ss") & "' UPDTIME," & _
                    "'" & strUpdEn & "' UPDEN " & _
                    " FROM SO158 A , SO159 B, SO162 C" & _
                    " WHERE A.SEQNO IS NULL AND B.SEQNO IS NULL " & _
                    " AND A.EVENTID=B.EVENTID " & _
                    " AND B.CHANNELID=C.CHANNELID"
    
    strInsert160 = "INSERT INTO SO160 (" & _
                            "DFILENAME,DCREATIONDATE,PRODUCTID,IMPULSIVEFLAG,PRODUCTKIND," & _
                            "SPECIALPPV,PURCHASABLEFLAG,SOLDFLAG,SUSPENDEDFLAG," & _
                            "PRODUCTNAME,PRODUCTDESCRIPTION,ENGPRODUCTNAME,ENGPRODUCTDESCRIPTION," & _
                            "SALEBEGINTIME,SALEENDTIME,VALIDITYPERIODBEGINTIME,VALIDITYPERIODENDTIME," & _
                            "PRODUCTCATEGORY,EPGPRICE,MONEYUNIT,SUBALLOWPPVEVENTFLAG," & _
                            "EVENTPREVIEWTIME,DPARENTALRATING," & _
                            "EFILENAME,ECREATIONDATE,CHANNELID,EVENTBEGINTIME,DURATION,EVENTID," & _
                            "EVENTTYPE,NAME,EPGDESC,ENGNAME,ENGEPGDESC,PARENTALRATING," & _
                            "CONTENTNIBBLE1,CONTENTNIBBLE2,USERNIBBLE1,USERNIBBLE2," & _
                            "CHANNELNO,UPDTIME,UPDEN) "
    
    cnCOMM.Execute strInsert160 & str3in1, lngRcd
    
  Exit Sub
ChkErr:
    ErrHandle Name, "Process_SO160"
End Sub

'8.  比對影片資料是否存在
'   (1) 將SO160資料group by Name(即影片名稱)後，取eFileName檔案時間最大的資料，與SO154.Name的資料拿來做比對。
'   (2) SO160.Parentalrating於節目表xml檔裡出現的範圍值為0~15之間，需對照CD029節目等級代碼檔的年齡分類，
'       取其符合的節目等級代碼、名稱回填SO154.ParentalRating、ParentalRatingName。
'       Select codeno,description from cd029
'       where SO160.Parentalrating between AGE1 and AGE2
'   (3) 判斷SO160.Name有無存在SO154
'       A.  若不存在，則取SO160裡的部分資料insert至SO154，並由系統產生一FilmNo，編碼規則為yyyymmdd + 7碼流水號；
'           流水號則取Sequence Object S_SO154_FilmNo。
'       B.  若已存在，則取SO160裡的eCreationDate, Name, EpgDesc, EngName, EngEpgDesc ParentalRating, Duration欄位值，
'           完整update至SO154。
'   (4) 請依照【P7 Schema】SO154資料表處理規則。

'SELECT ECREATIONDATE,NAME,EPGDESC,ENGNAME,ENGEPGDESC,
'    PARENTALRATING,PARENTALRATINGNAME,DURATION,
'    CONTENTPROVIDERID , CONTENTPROVIDERNAME, NOTE
' From SO160
'
'SELECT ECREATIONDATE,NAME,EPGDESC,ENGNAME,ENGEPGDESC,
'    PARENTALRATING , PARENTALRATINGNAME, DURATION
' From SO160

Private Sub Clear_ListBox()
  On Error Resume Next
    lstSrc.Clear
    lstDst.Clear
    lstErr.Clear
    lstXML.Clear
End Sub

Private Sub Set_DirLstBox(XML_Type As String)
  On Error Resume Next
    Process_DirListBox dirDst, clsFS, lblDst, lstDst, XML_Type & " 目的位置"
    Process_DirListBox dirErr, clsFS, lblErr, lstErr, XML_Type & " 問題位置"
    Process_DirListBox dirSrc, clsFS, lblSrc, lstSrc, XML_Type & " 來源位置"
End Sub

Private Function Process_XML(Optional IsEPG As Boolean = False) As Boolean ' 處理 DEX XML 檔案
  On Error GoTo ChkErr
    Dim lngLoop As Long
    Dim strXMLfile As String
    Dim lngRcdCnt As Long
    Dim XML_Type As String
    Process_XML = False
    
    If Not SetDirLoc(IsEPG) Then
        MsgBox "[ 系統參數 ] 設定有誤 ! 請確認 ..", vbInformation, "訊息"
        tabMain.Tab = 1
        cmdStop.Value = True
        Exit Function
    End If
    
    XML_Type = IIf(IsEPG, "EPG", "DEX")
    
    Clear_ListBox
    Set_DirLstBox XML_Type
    
    If intFileCount > 0 Then

        ClearVirtualRS
        
        lstLog.AddItem "將 XML 檔案移至 [問題位置] .."
        lstLog.ListIndex = lstLog.ListCount - 1
        
        For lngLoop = 0 To lstSrc.ListCount - 1
            Move_XML_File XML_Type, lstSrc.List(lngLoop), clsFS.File(lngLoop)
        Next
        
        Set_DirLstBox XML_Type
        
        Generati_DTD IsEPG ' 產生 XML 的 DTD
        
        With rsVirtual
            lngRcdCnt = .RecordCount
            If lngRcdCnt > 0 Then
                
                lstLog.AddItem "準備處理 " & XML_Type & " XML .."
                lstLog.AddItem "---------------------------------------------------------------------------"
                lstLog.ListIndex = lstLog.ListCount - 1

                .Sort = "FileDate"
                .MoveFirst
                pgb.Max = lngRcdCnt
                
                While Not .EOF
                    
                    strXMLfile = .Fields("FileName")
                    lstLog.AddItem "分析 " & strXMLfile & " 檔案 .."
                    lstLog.ListIndex = lstLog.ListCount - 1
                    strXMLfile = dirErr.Path & "\" & strXMLfile
                    
                    If Analyze_XML_then_Update_Table(strXMLfile) Then
                        lstLog.AddItem "更新 " & XML_Type & " Last import time .."
                        lstLog.ListIndex = lstLog.ListCount - 1
                        If Upd_Last_Imp_Date(XML_Type, .Fields("FileDate")) Then
                            lstLog.AddItem "處理完成 ! 將 XML 檔案移至 [目的位置] .."
                            lstLog.AddItem "---------------------------------------------------------------------------"
                            lstLog.ListIndex = lstLog.ListCount - 1
                            Name strXMLfile As dirDst.Path & "\" & .Fields("FileName")
                        Else
                            MsgBox "INI 存檔錯誤! 系統將關閉 , 請洽開博客服人員 !", vbInformation, "訊息"
                            End
                        End If
                    End If
                    
                    pgb.Value = .AbsolutePosition
                    .MoveNext
                Wend
                pgb.Value = 0
            Else
                lstLog.AddItem "未處理 XML 檔案 ! 無符合規格之資料 .."
                lstLog.ListIndex = lstLog.ListCount - 1
            End If
        End With
        Set_DirLstBox XML_Type
    Else
        lstLog.AddItem "目錄中無 XML 檔案 .."
    End If
    Process_XML = True
  Exit Function
ChkErr:
    ErrHandle Name, "Process_XML"
End Function

' 產生文件類別定義檔 ( Document Type Definition ) , Fake 版 , 因ㄐ車 Vendor 不肯給 Specification , so ..
Private Sub Generati_DTD(Optional IsEPG As Boolean = False)
  On Error GoTo ChkErr
    Dim lngHfile As Long
    Dim strDTDfile As String
    If IsEPG Then
        strDTDfile = dirErr.Path & "\BroadcastData.dtd"
    Else
        strDTDfile = dirErr.Path & "\ProductManagementData.dtd"
    End If
    If Len(Dir(strDTDfile)) = 0 Then
        lngHfile = CreateFile(strDTDfile, &H40000000, &H2, ByVal 0&, 1, 0, 0)
        CloseHandle lngHfile
    End If
  Exit Sub
ChkErr:
    ErrHandle Name, "Generati_DTD"
End Sub

Private Function Analyze_XML_then_Update_Table(strXMLfile As String) As Boolean ' 解析 XML 並 Update 至所屬 Table.Field 裡
  On Error GoTo ChkErr
    Analyze_XML_then_Update_Table = False
    str_XML_Class = UCase(Left(rsVirtual(0), 5))
    If LoadXML(strXMLfile, False, lstXML) Then
        cnCOMM.BeginTrans
        PrcNode objDOM.documentElement, lstXML
        If Update_XML_Table Then
            cnCOMM.CommitTrans
        Else
            cnCOMM.RollbackTrans
        End If
    End If
    Analyze_XML_then_Update_Table = True
  Exit Function
ChkErr:
    ErrHandle Name, "Analyze_XML_then_Update_Table"
    On Error Resume Next
    cnCOMM.RollbackTrans
End Function

Private Sub ClearVirtualRS() ' 清空虛擬資料庫
  On Error GoTo ChkErr
    With rsVirtual
            .Close
            .CursorLocation = adUseClient
             stm.Position = 0
            .Open stm
    End With
  Exit Sub
ChkErr:
    ErrHandle Name, "ClearVirtualRS"
End Sub

' 開啟虛擬記錄體資料錄並存入 Stream 物件 , 以供後續排序不同種類之 XML 檔案使用
Private Sub Open_Virtual_RS()
  On Error GoTo ChkErr
    With rsVirtual
        .CursorLocation = adUseClient
        .Fields.Append "FileName", adVarWChar, 26
        .Fields.Append "FileDate", adVarWChar, 15, adFldKeyColumn
        .Open
'        .Fields("FileDate").Properties("Optimize") = True
        .Save stm, adPersistADTG
    End With
  Exit Sub
ChkErr:
    ErrHandle Name, "Open_Virtual_RS"
End Sub

 ' 檢核XML檔案屬性及是否獨佔
Private Function Chk_File_Can_Move(strFullPathFile As String) As Boolean
  On Error GoTo ChkErr
    Chk_File_Can_Move = Not Chk_File_Exclusive(strFullPathFile)
    If Chk_File_Can_Move Then
        If File_ReadOnly(strFullPathFile) Then Set_FileRead (strFullPathFile)
    End If
  Exit Function
ChkErr:
    ErrHandle Name, "Chk_File_Can_Move"
End Function

'2.  程式需能做到自動偵測：
'   (1) 開發欄位可供設定約多久程式會自動去撈取來源資料夾是否有檔案。
'   (2) 若於來源位置有檔案存在，則將該路徑下的所有檔案，剪走至問題位置存放做拆解。
'   (3) 若已成功拆解完成的檔案，則將其剪走存放於最後的目的位置，若失敗的檔案則仍保留於問題位置。
  
'3.  節目表XML檔之檔案命名原則
'   (1) 計次付費(PPV)： CAPPV西元年月日時分秒. XML    例：CAPPV20050520083000.XML
'   (2) 付費頻道(SUB)： CASUB西元年月日時分秒.XML    例：CASUB20050520173015.XML
'   (3) EPG Wrapper ： Nagra_西元年月日時分秒. XML   例：Nagra_20050530031019.XML
'
'4.  DEX 、Wrapper XML檔，可能會一次放多個，故於拆解時需考慮到解讀的順序：
'   (1) 於派工參數SO042增加二個欄位，以儲存DEX和Wrapper檔案最後一次解讀的檔案名稱後14碼
'   (2) 於解讀前，需先逐一判斷各檔案名稱後14碼的時間，需大於等於SO042節目表最新匯入時間，方可開始做拆解比對動作，否則不允理會。
'       A.  檔名為CAPPV、CASUB開頭的，則存取SO042.DexLstImtTime
'       B.  檔名為Nagra_開頭的 , 則存取SO042.WrapLstImtTime
'   (3) 其XML檔名時間部分，皆為 UTC 時間(格林威治標準時間)，因此為資料拆解而已，則直接給什麼就存什麼，不需要再轉成 Taiwan Local 時間。

 ' 搬移 XML 檔案 ( 由 [來源位置] 搬至 [問題位置] 以供後續處理 )
Private Sub Move_XML_File(XML_Type As String, _
                                            strFile As String, _
                                            strFullPathFile As String)
  On Error GoTo ChkErr
    Dim strErrFile As String
    Dim strFileDT As String
    If Chk_File_Can_Move(strFullPathFile) Then
        On Error Resume Next
        strErrFile = dirErr.Path & "\" & strFile
        Name strFullPathFile As strErrFile
        If Err.Number = 0 Then
            If Greater_Than_Last_Time(XML_Type, strFile, strFileDT) Then ' 比對最後匯入時間
                rsVirtual.AddNew Array(0, 1), Array(strFile, strFileDT)
'                cltFile.Add strErrFile  ' 增加欲處理之 XML 檔案清單至 [集合物件]
            End If
        Else
            Err.Clear
        End If
    End If
  Exit Sub
ChkErr:
    ErrHandle Name, "Move_XML_File"
End Sub

' 檢核檔名末14碼是否大於 LastImpTime , PS : 調整規格 , 把 [SO042] 的欄位移至 INI 檔
'       A.  檔名為CAPPV、CASUB開頭的，則存取SO042.DexLstImtTime  => INI 檔裡 [LastImp] Section 裡的 Dex
'       B.  檔名為Nagra_開頭的 , 則存取SO042.WrapLstImtTime  => INI 檔裡 [LastImp] Section 裡的 Wrap
Private Function Greater_Than_Last_Time(XML_Type As String, _
                                                                    strFileName As String, _
                                                                    ByRef strFileDT As String) As Boolean
  On Error GoTo ChkErr
    strFileDT = Right(Replace(strFileName, ".XML", "", 1, , vbTextCompare), 14)
    Greater_Than_Last_Time = strFileDT > CStr(Decrypt(Read_From_INI("LastImp", XML_Type)))
  Exit Function
ChkErr:
    ErrHandle Name, "Greater_Than_Last_Time"
End Function

' 更新 Dex / Epg 的 LastImpTime 到 INI 檔裡
Private Function Upd_Last_Imp_Date(XML_Type As String, strLastTime As String) As Boolean
  On Error GoTo ChkErr
    Upd_Last_Imp_Date = Write_2_INI("LastImp", XML_Type, Encrypt(strLastTime)) = 1
  Exit Function
ChkErr:
    ErrHandle Name, "Upd_Last_Imp_Date"
End Function

 ' 處理 3 個 DirListBox 控制項之路徑 , 檢核路徑下是否有 XML 檔案 , 若有則增加檔案清單至 ListBox
Private Sub Process_DirListBox(ByRef dirObj As Control, _
                                                    ByRef clsObj As Object, _
                                                    ByRef lblObj As Control, _
                                                    ByRef lstObj As Control, _
                                                    ByVal strCaption As String)
  On Error GoTo ChkErr
    Dim lngLoop As Long
    intFileCount = 0
    lstObj.Clear
    If Not EmptyDir(dirObj.Path) Then
        With clsObj
                .ResetSearch
                .FindFiles dirObj.Path, "*.XML", True
                intFileCount = .FilesFound
                If intFileCount > 0 Then
                    lblObj.Caption = " 搜尋 [ " & strCaption & " ] 下 XML 檔案 .. "
                    intFileCount = intFileCount - 1
                    For lngLoop = 0 To intFileCount
                        lstObj.AddItem Replace(.File(lngLoop), dirObj.Path & "\", "", 1)
                    Next
                    intFileCount = intFileCount + 1
                    lblObj = " [ " & strCaption & " ] 找到 " & CStr(intFileCount) & " 個 XML 檔案"
'                    lblObj = " [ " & strCaption & " ] 找到 " & CStr(intFileCount) & " 個 XML 檔案 , 處理中 .."
                Else
                    lblObj = " [ " & strCaption & " ] 中無 XML 檔案"
                End If
        End With
    Else
        lblObj = " [ " & strCaption & " ] 中無 XML 檔案"
    End If
    lblPrc = " 處理進度說明"
  Exit Sub
ChkErr:
    ErrHandle Name, "Process_DirListBox"
End Sub

' 設定 DirListBox 路徑 , 並檢核資料夾是否存在
Private Function SetDirPath(ByRef objDir As Control, ByRef ctlText As Control) As Boolean
  On Error GoTo ChkErr
    SetDirPath = True
    With objDir
            .Tag = Trim(ctlText)
             If DirExists(.Tag) Then
                .Path = .Tag
             Else
                SetDirPath = False
             End If
    End With
  Exit Function
ChkErr:
    ErrHandle Name, "SetDirPath"
End Function

' 設定 DEX / EPG 所有 DirListBox 路徑 , 並檢核資料夾是否存在
Private Function SetDirLoc(Optional IsEPG As Boolean = False) As Boolean
  On Error GoTo ChkErr
    If Not IsEPG Then
        SetDirLoc = SetDirPath(dirSrc, txt_DEX_Src_Path)
        SetDirLoc = SetDirPath(dirDst, txt_DEX_Dst_Path)
        SetDirLoc = SetDirPath(dirErr, txt_DEX_Err_Path)
    Else
        SetDirLoc = SetDirPath(dirSrc, txt_EPG_Src_Path)
        SetDirLoc = SetDirPath(dirDst, txt_EPG_Dst_Path)
        SetDirLoc = SetDirPath(dirErr, txt_EPG_Err_Path)
    End If
  Exit Function
ChkErr:
    ErrHandle Name, "SetDirLoc"
End Function

Public Sub vTimer() ' API Timer Event
  On Error GoTo ChkErr
    lblStatus = "處理資料中.."
    lblAction = "執行中.."
    blnExecuting = True
    DoEvents
    TimerStop lngTmr
    
    lstLog.Clear
    LookInto ' 開始檢查目錄內是否有檔案 , 若有將處理善後
    
    lblStatus = "處理完成.."
    lblAction = "執行完畢.."
    blnExecuting = False
    DoEvents
    If Not blnStopFlag Then System_GoOn
  Exit Sub
ChkErr:
    ErrHandle Name, "Subroutine : vTimer (Call Back Event )"
End Sub

Private Sub System_GoOn() ' 檢查執行時段( 尖峰 / 離峰 )
  On Error GoTo ChkErr
    If Not blnStarting Then Exit Sub
    Dim strTime As String
    strTime = Format(Time, "HHMM")
    If strTime > mskPeakTime1 And strTime < mskPeakTime2 Then
        sglSecond = Val(txtPeakTime) '尖峰
        System_Starting True
    Else
        sglSecond = Val(txtOffPeak) '離峰
        System_Starting False
    End If
  Exit Sub
ChkErr:
    ErrHandle Name, "Subroutine : System_GoOn"
End Sub

Private Sub System_Starting(blnPeak As Boolean) ' API Timer 啟動 (系統啟動)
  On Error GoTo ChkErr
    lblCountDown = SecToHMS(sglSecond)
    tmr.Enabled = True
    lblStartTime = Format(DateAdd("s", sglSecond, CStr(Now)), "YYYY/MM/DD HH:MM:SS")
    TimerGo lngTmr, sglSecond
    If blnPeak Then
        lblTime = "尖峰"
    Else
        lblTime = "離峰"
    End If
    lblStatus = "系統啟動.."
    lblAction = "等待中.."
  Exit Sub
ChkErr:
    ErrHandle Name, "Subroutine : System_Starting"
End Sub

Private Sub cmdOpen_Click(index As Integer) ' 選擇資料夾之控制項陣列 CommandButton Click 事件
  On Error GoTo ChkErr
    Dim strFolder As String
    strFolder = BrowseFolder("請選擇資料夾")
    Select Case index
                Case 0
                        txt_DEX_Src_Path = strFolder
                Case 1
                        txt_DEX_Err_Path = strFolder
                Case 2
                        txt_DEX_Dst_Path = strFolder
                Case 3
                        txt_EPG_Src_Path = strFolder
                Case 4
                        txt_EPG_Err_Path = strFolder
                Case 5
                        txt_EPG_Dst_Path = strFolder
    End Select
  Exit Sub
ChkErr:
    ErrHandle Name, "cmdOpen_Click(Index As Integer)"
End Sub

Private Sub cmdStop_Click() ' 系統停止
  On Error GoTo ChkErr
    tmr.Enabled = False
    blnStarting = False
    blnStopFlag = True
    TimerStop lngTmr
    cmdStart.Enabled = True
    cmdStop.Enabled = False
    lblCountDown = ""
    lblTime = ""
    lblStartTime = ""
    dirSrc.Path = "C:\"
    dirDst.Path = "C:\"
    dirErr.Path = "C:\"
    lstSrc.Clear
    lstDst.Clear
    lstErr.Clear
    lstLog.Clear
    lblSrc = ""
    lblDst = ""
    lblErr = ""
    lblPrc = ""
    If blnExecuting Then
        lblStatus = "系統即將停止"
        lblAction = "執行中.."
    Else
        lblStatus = "系統停止.."
        lblAction = "未啟動.."
        pgb.Value = 0
    End If
    DoEvents
  Exit Sub
ChkErr:
    ErrHandle Name, "Subroutine : cmdStop_Click"
End Sub

Private Sub cmdUpdateSysPara_Click()  ' 系統參數頁籤之存檔按紐
  On Error GoTo ChkErr
    Sys_Para_EditMode False
    If Not Update_Sys_Para Then MsgBox "存檔失敗 !" & strCrLf & "請洽開博科技公司 .. 謝謝!", vbInformation, "訊息"
  Exit Sub
ChkErr:
    ErrHandle Name, "Click Event : cmdUpdateSysPara_Click"
End Sub

Private Sub cmdUpdatePermission_Click()  ' 權限管理頁籤之存檔按紐
  On Error GoTo ChkErr
    If IsPermissionDataOK Then
        Permission_EditMode False
        If Not Update_Permission Then MsgBox "存檔失敗 !" & strCrLf & "請洽開博科技公司 .. 謝謝!", vbInformation, "訊息"
    End If
  Exit Sub
ChkErr:
    ErrHandle Name, "Click Event : cmdUpdatePermission_Click"
End Sub

Private Function IsPermissionDataOK() As Boolean  ' 檢核權限管理頁籤之攔位是否正確
  On Error Resume Next
    IsPermissionDataOK = True
    If chkUsePwd.Value = 1 Then
        If Len(txtPassword) = 0 Or Len(txtCfmPwd) = 0 Then
            MsgBox "[輸入密 碼], [確認密碼] 皆為必要欄位, 請確認 !", vbInformation, "訊息"
            IsPermissionDataOK = False
            txtPassword.SetFocus
        Else
            If txtPassword <> txtCfmPwd Then
                MsgBox "[輸入密 碼], [確認密碼] 需相同, 請確認 !", vbInformation, "訊息"
                IsPermissionDataOK = False
                txtPassword.SetFocus
            End If
        End If
    End If
End Function

Private Function Read_Sys_Para() As Boolean  ' 將INI的系統參數讀至系統參數頁籤的相對欄位上
  On Error GoTo ChkErr
    
    mskPeakTime1 = Decrypt(Read_From_INI("SysSetup", "PeakTime"))
    mskPeakTime2 = Decrypt(Read_From_INI("SysSetup", "OffPeak"))
    txtPeakTime = Decrypt(Read_From_INI("SysSetup", "PeakTimeStep"))
    txtOffPeak = Decrypt(Read_From_INI("SysSetup", "OffPeakStep"))
    
    txtDSN = Decrypt(Read_From_INI("DB_SID", "KeyA"))
    txtUID = Decrypt(Read_From_INI("DB_SID", "KeyB"))
    txtPWD = Decrypt(Read_From_INI("DB_SID", "KeyC"))
  
    txt_DEX_Src_Path = Decrypt(Read_From_INI("DEX", "SrcPath"))
    txt_DEX_Err_Path = Decrypt(Read_From_INI("DEX", "ErrPath"))
    txt_DEX_Dst_Path = Decrypt(Read_From_INI("DEX", "DstPath"))
  
    txt_EPG_Src_Path = Decrypt(Read_From_INI("EPG", "SrcPath"))
    txt_EPG_Err_Path = Decrypt(Read_From_INI("EPG", "ErrPath"))
    txt_EPG_Dst_Path = Decrypt(Read_From_INI("EPG", "DstPath"))
    
    txtDEXlstImpTime = Decrypt(Read_From_INI("LastImp", "DEX"))
    txtEPGlstImpTime = Decrypt(Read_From_INI("LastImp", "EPG"))
  
  Exit Function
ChkErr:
    ErrHandle Name, "Function : Update_Sys_Para"
End Function

Private Function Read_Permission() As Boolean  ' 將INI的權限管理讀至權限管理頁籤的相對欄位上
  On Error GoTo ChkErr
    If IsDate(Decrypt(Read_From_INI("Permission", "ChkCode"))) Then
        chkUsePwd.Value = 0
    Else
        chkUsePwd.Value = 1
    End If
    txtPassword = Decrypt(Read_From_INI("Permission", "KeyWord"))
    txtCfmPwd = txtPassword
  Exit Function
ChkErr:
    ErrHandle Name, "Function : Read_Permission"
End Function

Private Function Update_Permission() As Boolean  ' 將INI的權限管理讀至權限管理頁籤的相對欄位上
  On Error GoTo ChkErr
    Update_Permission = True
    If chkUsePwd = 1 Then
        Update_Permission = Write_2_INI("Permission", "ChkCode", Encrypt("ChouYinDer")) = 1
    Else
        Update_Permission = Write_2_INI("Permission", "ChkCode", Encrypt(CStr(Date))) = 1
    End If
    Update_Permission = Update_Permission And (Write_2_INI("Permission", "KeyWord", Encrypt(txtPassword)) = 1)
  Exit Function
ChkErr:
    ErrHandle Name, "Function : Update_Permission"
End Function

Private Function Update_Sys_Para() As Boolean ' 將系統參數頁籤之欄位值寫回INI檔的系統參數
  On Error GoTo ChkErr
    Update_Sys_Para = False
    
    If Not Write_2_INI("SysSetup", "PeakTime", Encrypt(mskPeakTime1)) = 1 Then Exit Function
    If Not Write_2_INI("SysSetup", "OffPeak", Encrypt(mskPeakTime2)) = 1 Then Exit Function
    If Not Write_2_INI("SysSetup", "PeakTimeStep", Encrypt(txtPeakTime)) = 1 Then Exit Function
    If Not Write_2_INI("SysSetup", "OffPeakStep", Encrypt(txtOffPeak)) = 1 Then Exit Function
    
    If Not Write_2_INI("DB_SID", "KeyA", Encrypt(txtDSN)) = 1 Then Exit Function
    If Not Write_2_INI("DB_SID", "KeyB", Encrypt(txtUID)) = 1 Then Exit Function
    If Not Write_2_INI("DB_SID", "KeyC", Encrypt(txtPWD)) = 1 Then Exit Function
    
    If Not Write_2_INI("DEX", "SrcPath", Encrypt(txt_DEX_Src_Path)) = 1 Then Exit Function
    If Not Write_2_INI("DEX", "ErrPath", Encrypt(txt_DEX_Err_Path)) = 1 Then Exit Function
    If Not Write_2_INI("DEX", "DstPath", Encrypt(txt_DEX_Dst_Path)) = 1 Then Exit Function
  
    If Not Write_2_INI("EPG", "SrcPath", Encrypt(txt_EPG_Src_Path)) = 1 Then Exit Function
    If Not Write_2_INI("EPG", "ErrPath", Encrypt(txt_EPG_Err_Path)) = 1 Then Exit Function
    If Not Write_2_INI("EPG", "DstPath", Encrypt(txt_EPG_Dst_Path)) = 1 Then Exit Function
  
    Update_Sys_Para = True
    
  Exit Function
ChkErr:
    ErrHandle Name, "Function : Update_Sys_Para"
End Function

Private Sub cmdUndoSysPara_Click() ' 系統參數頁籤之取消按紐
  On Error GoTo ChkErr
    Revert_Original_Sys_Para
    Sys_Para_EditMode False
  Exit Sub
ChkErr:
    ErrHandle Name, "Click Event : cmdUndoSysPara_Click"
End Sub

Private Sub cmdUndoPermission_Click() ' 權限管理之取消按紐
  On Error GoTo ChkErr
    Revert_Original_Permission
    Permission_EditMode False
  Exit Sub
ChkErr:
    ErrHandle Name, "Click Event : cmdUndoPermission_Click"
End Sub

Private Sub cmdEditSysPara_Click() ' 系統參數頁籤之修改按紐
  On Error GoTo ChkErr
    Keep_Original_Sys_Para
    If Len(mskPeakTime1) = 0 Or mskPeakTime1 = "0000" Then mskPeakTime1 = "0730"
    If Len(mskPeakTime2) = 0 Or mskPeakTime2 = "0000" Then mskPeakTime2 = "1130"
    If Len(txtPeakTime) = 0 Or txtPeakTime = "0" Then txtPeakTime = "300"
    If Len(txtOffPeak) = 0 Or txtOffPeak = "0" Then txtOffPeak = "3600"
    Sys_Para_EditMode True
    On Error Resume Next
    mskPeakTime1.SetFocus
  Exit Sub
ChkErr:
    ErrHandle Name, "Click Event : cmdEditSysPara_Click"
End Sub

Private Sub cmdEditPermission_Click() ' 權限管理頁籤之修改按紐
  On Error GoTo ChkErr
    Keep_Original_Permission
    Permission_EditMode True
    On Error Resume Next
    chkUsePwd.SetFocus
  Exit Sub
ChkErr:
    ErrHandle Name, "Click Event : cmdEditPermission_Click"
End Sub

Private Sub Keep_Original_Sys_Para() ' 保留系統參數頁籤之畫面值
  On Error Resume Next
    mskPeakTime1.Tag = mskPeakTime1
    mskPeakTime2.Tag = mskPeakTime2
    txtPeakTime.Tag = txtPeakTime
    txtOffPeak.Tag = txtOffPeak
End Sub

Private Sub Revert_Original_Sys_Para() ' 還原系統參數頁籤之畫面值
  On Error Resume Next
    mskPeakTime1 = mskPeakTime1.Tag
    mskPeakTime2 = mskPeakTime2.Tag
    txtPeakTime = txtPeakTime.Tag
    txtOffPeak = txtOffPeak.Tag
End Sub

Private Sub Keep_Original_Permission() ' 保留權限管理頁籤之畫面值
  On Error Resume Next
    chkUsePwd.Tag = chkUsePwd.Value
    txtPassword.Tag = txtPassword
End Sub

Private Sub Revert_Original_Permission() ' 還原權限管理頁籤之畫面值
  On Error Resume Next
    chkUsePwd.Value = chkUsePwd.Tag
    txtPassword = txtPassword.Tag
    txtCfmPwd = txtPassword
End Sub

Private Sub Permission_EditMode(blnModify As Boolean) ' 權限管理頁籤之編輯模式畫面控管
  On Error Resume Next
    tabMain.TabEnabled(0) = Not blnModify
    tabMain.TabEnabled(1) = Not blnModify
    fraPassword.Enabled = blnModify
    cmdEditPermission.Enabled = Not blnModify
    cmdUpdatePermission.Enabled = blnModify
    cmdUndoPermission.Enabled = blnModify
End Sub

Private Sub Sys_Para_EditMode(blnModify As Boolean) ' 系統參數頁籤之編輯模式畫面控管
  On Error Resume Next
    tabMain.TabEnabled(0) = Not blnModify
    tabMain.TabEnabled(2) = Not blnModify
    fraSetup.Enabled = blnModify
    fraDatabase.Enabled = blnModify
    fraDEX.Enabled = blnModify
    fraEPG.Enabled = blnModify
    cmdEditSysPara.Enabled = Not blnModify
    cmdUpdateSysPara.Enabled = blnModify
    cmdUndoSysPara.Enabled = blnModify
End Sub

Private Sub Form_Initialize() ' 表單初始化
  On Error GoTo ChkErr
    With imgBG
         Set .Container = tabMain
         .Left = -3388
    End With
  Exit Sub
ChkErr:
    ErrHandle Name, "Event : Form_Initialize"
End Sub

Private Sub lbl系統參數_Click(index As Integer)
  On Error Resume Next
    tabMain.Tab = 1
End Sub

Private Sub lbl監控管理_Click(index As Integer)
  On Error Resume Next
    tabMain.Tab = 0
End Sub

Private Sub lbl權限管理_Click(index As Integer)
  On Error Resume Next
    tabMain.Tab = 2
End Sub

Private Sub tabMain_Click(PreviousTab As Integer) ' Tab 切換時讀取該頁籤資料
  On Error GoTo ChkErr
    LockWindowUpdate hwnd
    Select Case tabMain.Tab
             Case 0
                    lblMsg = " [ 監控管理 ] 主要功能為 [啟動] 、 [停止] 系統 ! 並可即時監控執行狀況 .."
                    SetDirLoc
             Case 1
                    imgBG.Move 15, 375, 9655, 5695
                    Read_Sys_Para
                    lblMsg = " [ 系統設定 ]、[ 共用區資料連線設定 ]、[ DEX 路徑設定]、[ EPG Wrapper 路徑設定 ] 需正確設定 , 系統才能正常運作 !"
             Case 2
                    imgBG.Move 12, 375, 9655, 5680
                    Read_Permission
                    lblMsg = " 當 [ 使用密碼啟動系統 ] 選取 ( 打勾 ) 時 , 則需輸入[ 輸入密碼 ]、[ 確認密碼] , 之後按 [ 存檔 ] 即完成成設定 !"
    End Select
    LockWindowUpdate 0
  Exit Sub
ChkErr:
    ErrHandle Name, "Click Event : tabMain_Click"
    LockWindowUpdate 0
End Sub

Private Sub chkUsePwd_Click() ' 當不使用密碼啟動系統, 則清空密碼欄位
  On Error Resume Next
    If fraPassword.Enabled Then
        If chkUsePwd.Value = 0 Then
            txtPassword = ""
            txtCfmPwd = ""
        Else
            txtPassword.SetFocus
        End If
    End If
End Sub

Private Sub tmr_Timer() ' VB timer 用來處理執行時間的倒數
  On Error GoTo ChkErr
    If sglSecond > 0 Then sglSecond = sglSecond - 1
    lblCountDown = SecToHMS(sglSecond)
  Exit Sub
ChkErr:
    ErrHandle Name, "tmr_Timer"
End Sub

Private Sub mskPeakTime1_Validate(Cancel As Boolean) ' 時分檢核
  On Error GoTo ChkErr
    Cancel = Not Chk_HM_OK(mskPeakTime1)
  Exit Sub
ChkErr:
    ErrHandle Name, "Event : mskPeakTime1_Validate"
End Sub

Private Sub mskPeakTime2_Validate(Cancel As Boolean) ' 時分檢核
  On Error GoTo ChkErr
    Cancel = Not Chk_HM_OK(mskPeakTime2)
  Exit Sub
ChkErr:
    ErrHandle Name, "Event : mskPeakTime2_Validate"
End Sub

'底下的 GotFocus 事件中的程式碼主要用來 Focus 進攔位時 , 全選欄位值
Private Sub mskPeakTime1_GotFocus()
  On Error Resume Next
    objSelAll mskPeakTime1
End Sub

Private Sub mskPeakTime2_GotFocus()
  On Error Resume Next
    objSelAll mskPeakTime2
End Sub

Private Sub txtCfmPwd_GotFocus()
  On Error Resume Next
    objSelAll txtCfmPwd
End Sub

Private Sub txtDSN_GotFocus()
  On Error Resume Next
    objSelAll txtDSN
End Sub

Private Sub txtOffPeak_GotFocus()
  On Error Resume Next
    objSelAll txtOffPeak
End Sub

Private Sub txtPassword_GotFocus()
  On Error Resume Next
    objSelAll txtPassword
End Sub

Private Sub txtPeakTime_GotFocus()
  On Error Resume Next
    objSelAll txtPeakTime
End Sub

Private Sub txtPWD_GotFocus()
  On Error Resume Next
    objSelAll txtPWD
End Sub

Private Sub txtUID_GotFocus()
  On Error Resume Next
    objSelAll txtUID
End Sub

Private Sub Form_Unload(Cancel As Integer) ' Unload form & Relase variant
  On Error Resume Next
    
    If blnStarting Then cmdStop.Value = True

    Rlx intMouseX
    Rlx intMouseY
    
    Rlx strTraceMsg
    Rlx blnInitialOK
    Rlx lngTmr
    Rlx sglSecond
    
    Rlx blnStarting
    Rlx blnStopFlag
    Rlx blnExecuting
    
    Rlx strTraceMsg
    Rlx blnInitialOK
    Rlx intFileCount
    
    Rlx cnCOMM
    Rlx rsCOMM
    Rlx rsVirtual
    
    Rlx rsSO158
    Rlx rsSO163
    Rlx rsSO159
    Rlx rsSO162
    Rlx rsSO161
    
    Rlx strCreationDate
    
    Rlx stm
    Rlx str_XML_Class
'    Rlx cltFile
    Set frmMain = Nothing
    End
    
End Sub

' 以下程式碼主要用來控制 XP 樣式的表單行為
Private Sub XPstyle() ' 設定表單為 XP 樣式外框
  On Error Resume Next
    imgTop.Visible = True
    imgTopRight.Visible = True
    lblCaption.Visible = True
    lblCaptionShadow.Visible = True
    imgIcon.Visible = True
    imgTopLeft.Visible = True
    imgExit.Top = 55
    imgBottom.Left = 0
    imgLeft.Top = 0
    imgRight.Top = 0
    imgTop.Top = 0
    imgLeft.Left = 0
    imgTopRight.Left = Me.ScaleWidth - imgTopRight.Width
    imgTopLeft.Left = 0
    imgTop.Width = Me.ScaleWidth
    imgLeft.Height = Me.ScaleHeight
    imgRight.Height = Me.ScaleHeight
    imgRight.Left = Me.ScaleWidth - imgRight.Width
    imgBottom.Top = Me.ScaleHeight - imgBottom.Height
    imgBottom.Width = Me.ScaleWidth
    imgExit.Left = Me.ScaleWidth - imgExit.Width - imgLeft.Width - 40
    imgMaxButton.Top = imgExit.Top
    imgMinButton.Top = imgExit.Top
    If imgMaxButton.Visible Then
        imgMaxButton.Left = imgExit.Left - imgMinButton.Width - 50
        imgMinButton.Left = imgExit.Left - imgMinButton.Width - imgMaxButton.Width - 100
    Else
        imgMinButton.Left = imgExit.Left - imgMinButton.Width - 50
    End If
    imgLeft.ZOrder
    imgRight.ZOrder
    imgBottom.ZOrder
    imgTopRight.ZOrder
    imgTopLeft.ZOrder
    imgExit.ZOrder
    lblCaptionShadow.Left = lblCaption.Left + 10
    lblCaptionShadow.Top = lblCaption.Top + 10
    lblCaption_Change
    lblCaption.ZOrder
    imgIcon.ZOrder
    Visible = True
End Sub

Private Sub imgExit_Click()
  On Error Resume Next
    Unload Me
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgExit.Picture = imgDisabled(2).Picture
End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgExit.Picture = imgEnabled(2).Picture
End Sub

Private Sub imgIcon_DblClick()
  On Error Resume Next
    If imgExit.Enabled Then imgExit_Click
End Sub

Private Sub imgMinButton_Click()
  On Error Resume Next
    Me.WindowState = 1
End Sub

Private Sub imgMinButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgMinButton.Picture = imgDisabled(1).Picture
End Sub

Private Sub imgMinButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgMinButton.Picture = imgEnabled(1).Picture
End Sub

Private Sub imgTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    If Button = 1 Then
        intMouseX = X
        intMouseY = Y
    End If
End Sub

Private Sub imgTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    If Button = 1 Then
        Me.Left = Me.Left + X - intMouseX
        Me.Top = Me.Top + Y - intMouseY
    End If
End Sub

Private Sub imgTopLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgTop_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgTopLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgTop_MouseMove Button, Shift, X, Y
End Sub

Private Sub imgTopRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgTop_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgTopRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgTop_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblCaption_Change() ' 設定 Form Caption
  On Error Resume Next
    lblCaptionShadow.Caption = lblCaption.Caption
    SetWindowText Me.hwnd, lblCaption.Caption
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgTop_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgTop_MouseMove Button, Shift, X, Y
End Sub


'Private Sub Clear_File_Clt() ' 清空 [集合物件] 中之檔案清單
'  On Error GoTo ChkErr
'    Dim intLoop As Integer
'    With cltFile
'        If .count < 1 Then Exit Sub
'        For intLoop = 1 To .count
'            .Remove 1
'        Next
'    End With
'    Rlx intLoop
'  Exit Sub
'ChkErr:
'    ErrHandle Name, "Clear_File_Clt"
'End Sub

