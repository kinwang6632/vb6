VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{818C3FA5-C15E-4B09-B383-B06546ED97C7}#1.0#0"; "PHbtn.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '沒有框線
   Caption         =   " CS 錄音系統介接 Gateway"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10185
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   10185
   StartUpPosition =   2  '螢幕中央
   Visible         =   0   'False
   Begin VB.Timer tmr 
      Left            =   9930
      Top             =   600
   End
   Begin VB.PictureBox picBackground 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5730
      Left            =   -6900
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   5700
      ScaleWidth      =   6930
      TabIndex        =   38
      Top             =   960
      Visible         =   0   'False
      Width           =   6955
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   6105
      Left            =   2925
      TabIndex        =   0
      Top             =   555
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   10769
      _Version        =   393216
      TabHeight       =   635
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BackColor       =   14737632
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "監 控 管 理 (&C)"
      TabPicture(0)   =   "frmMain.frx":79E14
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraConsole"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraMonitor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraTimer"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "系 統 參 數 (&P)"
      TabPicture(1)   =   "frmMain.frx":79E30
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDatabase"
      Tab(1).Control(1)=   "fraSetup"
      Tab(1).Control(2)=   "fraRcdSys"
      Tab(1).Control(3)=   "cmdUpdateSysPara"
      Tab(1).Control(4)=   "cmdUndoSysPara"
      Tab(1).Control(5)=   "cmdEditSysPara"
      Tab(1).Control(6)=   "imgSysPara"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "權 限 管 理 (&S)"
      TabPicture(2)   =   "frmMain.frx":79E4C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraPassword"
      Tab(2).Control(1)=   "cmdEditPermission"
      Tab(2).Control(2)=   "cmdUpdatePermission"
      Tab(2).Control(3)=   "cmdUndoPermission"
      Tab(2).Control(4)=   "imgPermission"
      Tab(2).ControlCount=   5
      Begin VB.Frame fraTimer 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 執 行 時 間 "
         ForeColor       =   &H00FF0000&
         Height          =   810
         Left            =   900
         TabIndex        =   43
         Top             =   5100
         Width           =   5205
      End
      Begin VB.Frame fraMonitor 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 即 時 監 控 "
         ForeColor       =   &H00FF0000&
         Height          =   3450
         Left            =   900
         TabIndex        =   40
         Top             =   1560
         Width           =   5205
      End
      Begin VB.Frame fraConsole 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 啟 動 命 令 "
         ForeColor       =   &H00FF0000&
         Height          =   870
         Left            =   900
         TabIndex        =   39
         Top             =   600
         Width           =   5205
         Begin PHbtn.PHbutton cmdStart 
            Height          =   375
            Left            =   990
            TabIndex        =   41
            Top             =   300
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "啟 動 (&R)"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   12632256
            FCOL            =   12582912
         End
         Begin PHbtn.PHbutton cmdStop 
            Height          =   375
            Left            =   3090
            TabIndex        =   42
            Top             =   300
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "停 止 (&X)"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   12632256
            FCOL            =   12582912
         End
      End
      Begin VB.Frame fraDatabase 
         BackColor       =   &H00FFFFFF&
         Caption         =   "共 用 區 資 料 庫 連 線 設 定 "
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1470
         Left            =   -73905
         TabIndex        =   28
         Top             =   3720
         Width           =   5205
         Begin VB.TextBox txtDSN 
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
            Left            =   2010
            TabIndex        =   31
            Top             =   240
            Width           =   1965
         End
         Begin VB.TextBox txtUID 
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
            Left            =   2010
            TabIndex        =   30
            Top             =   600
            Width           =   1965
         End
         Begin VB.TextBox txtPWD 
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            IMEMode         =   3  '暫止
            Left            =   2010
            PasswordChar    =   "*"
            TabIndex        =   29
            Top             =   960
            Width           =   1965
         End
         Begin VB.Label lblUID 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "U I D"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1230
            TabIndex        =   34
            Top             =   660
            Width           =   600
         End
         Begin VB.Label lblPWD 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "P W D"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1230
            TabIndex        =   33
            Top             =   1020
            Width           =   600
         End
         Begin VB.Label lblDSN 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "D S N"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1230
            TabIndex        =   32
            Top             =   300
            Width           =   600
         End
      End
      Begin VB.Frame fraSetup 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 系 統 設 定 "
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1470
         Left            =   -73905
         TabIndex        =   17
         Top             =   2130
         Width           =   5205
         Begin MSMask.MaskEdBox mskPeakTime1 
            Height          =   345
            Left            =   2010
            TabIndex        =   23
            Top             =   240
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
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
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2010
            TabIndex        =   19
            Text            =   "0"
            Top             =   600
            Width           =   705
         End
         Begin VB.TextBox txtOffPeak 
            Alignment       =   1  '靠右對齊
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
            Left            =   2010
            TabIndex        =   18
            Text            =   "0"
            Top             =   960
            Width           =   705
         End
         Begin MSMask.MaskEdBox mskPeakTime2 
            Height          =   345
            Left            =   3390
            TabIndex        =   24
            Top             =   240
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
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
            Caption         =   "尖 峰 時 段 每                   秒 讀 取 一 次 資 料"
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
            Left            =   690
            TabIndex        =   22
            Top             =   660
            Width           =   3960
         End
         Begin VB.Label lblOffPeak 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "離 峰 時 段 每                   秒 讀 取 一 次 資 料"
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
            Left            =   690
            TabIndex        =   21
            Top             =   1020
            Width           =   3900
         End
         Begin VB.Label lblPeakTimeDefine 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "尖 峰 時 段 定 義                   至"
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
            Left            =   450
            TabIndex        =   20
            Top             =   300
            Width           =   2700
         End
      End
      Begin VB.Frame fraPassword 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 密   碼   設   定 "
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
         ForeColor       =   &H00FF0000&
         Height          =   3195
         Left            =   -73860
         TabIndex        =   5
         Top             =   1140
         Width           =   5115
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
            Left            =   1440
            PasswordChar    =   "*"
            TabIndex        =   8
            ToolTipText     =   " 請輸入密碼 ! "
            Top             =   1560
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
            Left            =   1440
            PasswordChar    =   "*"
            TabIndex        =   7
            ToolTipText     =   " 確認密碼必須與輸入密碼相同 !"
            Top             =   2160
            Width           =   3225
         End
         Begin VB.CheckBox chkUsePwd 
            Appearance      =   0  '平面
            BackColor       =   &H00FFFFFF&
            Caption         =   " 使 用 密 碼 啟 動 系 統"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   465
            MaskColor       =   &H00FF0000&
            TabIndex        =   6
            ToolTipText     =   " 打勾表示使用密碼 , 反之 則無"
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   3255
         End
         Begin VB.Image imgKey 
            Appearance      =   0  '平面
            Height          =   1365
            Left            =   3330
            Picture         =   "frmMain.frx":79E68
            Stretch         =   -1  'True
            Top             =   180
            Width           =   1320
         End
         Begin VB.Label lblPassword 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "輸 入 密 碼"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   465
            TabIndex        =   10
            Top             =   1680
            Width           =   915
         End
         Begin VB.Label lblCfmPwd 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "確 認 密 碼"
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   465
            TabIndex        =   9
            Top             =   2280
            Width           =   855
         End
      End
      Begin VB.Frame fraRcdSys 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 錄 音 系 統 "
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1470
         Left            =   -73905
         TabIndex        =   4
         Top             =   540
         Width           =   5205
         Begin VB.TextBox txtSrvPort 
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
            Left            =   2010
            TabIndex        =   16
            Top             =   960
            Width           =   1965
         End
         Begin VB.TextBox txtSrvIP 
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
            Left            =   2010
            TabIndex        =   15
            Top             =   600
            Width           =   1965
         End
         Begin VB.TextBox txtServer 
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
            Left            =   2010
            TabIndex        =   14
            Top             =   240
            Width           =   1965
         End
         Begin VB.Label lblServerName 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "主 機 名 稱 "
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
            Left            =   930
            TabIndex        =   13
            Top             =   300
            Width           =   960
         End
         Begin VB.Label lblServerPort 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "主 機 Port"
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
            Left            =   930
            TabIndex        =   12
            Top             =   1020
            Width           =   870
         End
         Begin VB.Label lblServerIP 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  '透明
            Caption         =   "主 機 IP"
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
            Left            =   1155
            TabIndex        =   11
            Top             =   660
            Width           =   675
         End
      End
      Begin PHbtn.PHbutton cmdUpdateSysPara 
         Height          =   375
         Left            =   -71865
         TabIndex        =   1
         Top             =   5340
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "存 檔 (&S)"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         FCOL            =   12582912
      End
      Begin PHbtn.PHbutton cmdUndoSysPara 
         Height          =   375
         Left            =   -69840
         TabIndex        =   2
         Top             =   5340
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "取 消 (&C)"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         FCOL            =   12582912
      End
      Begin PHbtn.PHbutton cmdEditSysPara 
         Height          =   375
         Left            =   -73905
         TabIndex        =   3
         Top             =   5340
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "修 改 (&E)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         FCOL            =   12582912
      End
      Begin PHbtn.PHbutton cmdEditPermission 
         Height          =   375
         Left            =   -73860
         TabIndex        =   25
         Top             =   4830
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "修 改 (&E)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         FCOL            =   12582912
      End
      Begin PHbtn.PHbutton cmdUpdatePermission 
         Height          =   375
         Left            =   -71880
         TabIndex        =   26
         Top             =   4830
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "存 檔 (&S)"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         FCOL            =   12582912
      End
      Begin PHbtn.PHbutton cmdUndoPermission 
         Height          =   375
         Left            =   -69900
         TabIndex        =   27
         Top             =   4830
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "取 消 (&C)"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         FCOL            =   12582912
      End
      Begin VB.Image imgSysPara 
         Appearance      =   0  '平面
         BorderStyle     =   1  '單線固定
         Height          =   5745
         Left            =   -75000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   6960
      End
      Begin VB.Image imgPermission 
         Appearance      =   0  '平面
         BorderStyle     =   1  '單線固定
         Height          =   5745
         Left            =   -75000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   6960
      End
   End
   Begin VB.Image imgCS 
      Height          =   1350
      Left            =   240
      Picture         =   "frmMain.frx":7A3FA
      Stretch         =   -1  'True
      Top             =   525
      Width           =   2535
   End
   Begin VB.Image imgLogo 
      Appearance      =   0  '平面
      Height          =   4785
      Left            =   270
      Picture         =   "frmMain.frx":8529C
      Stretch         =   -1  'True
      Top             =   1860
      Width           =   2490
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   " 訊 息 列"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   0
      TabIndex        =   37
      Top             =   6840
      Width           =   10185
   End
   Begin VB.Image imgIcon 
      Height          =   330
      Left            =   90
      Picture         =   "frmMain.frx":AC1EA
      Stretch         =   -1  'True
      Top             =   30
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgMaxButton 
      Height          =   285
      Left            =   2100
      Picture         =   "frmMain.frx":ACAB4
      Stretch         =   -1  'True
      Top             =   7470
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgExit 
      Height          =   285
      Left            =   2760
      Picture         =   "frmMain.frx":AD038
      Stretch         =   -1  'True
      Top             =   7470
      Width           =   285
   End
   Begin VB.Image imgBottom 
      Height          =   45
      Left            =   2400
      Picture         =   "frmMain.frx":AD5BC
      Stretch         =   -1  'True
      Top             =   7860
      Width           =   630
   End
   Begin VB.Image imgRight 
      Height          =   375
      Left            =   3480
      Picture         =   "frmMain.frx":AD780
      Stretch         =   -1  'True
      Top             =   7500
      Width           =   45
   End
   Begin VB.Image imgLeft 
      Height          =   495
      Left            =   3360
      Picture         =   "frmMain.frx":ADA1C
      Stretch         =   -1  'True
      Top             =   7500
      Width           =   45
   End
   Begin VB.Image imgLogOff 
      Height          =   360
      Index           =   1
      Left            =   690
      Picture         =   "frmMain.frx":ADCB8
      Top             =   8880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgLogOff 
      Height          =   360
      Index           =   0
      Left            =   90
      Picture         =   "frmMain.frx":AE3BC
      Top             =   8880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgShutDown 
      Height          =   360
      Index           =   1
      Left            =   690
      Picture         =   "frmMain.frx":AEAC0
      Top             =   8520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgShutDown 
      Height          =   360
      Index           =   0
      Left            =   90
      Picture         =   "frmMain.frx":AF1C4
      Top             =   8520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgDimmed 
      Height          =   315
      Index           =   2
      Left            =   810
      Picture         =   "frmMain.frx":AF8C8
      Top             =   8160
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDimmed 
      Height          =   315
      Index           =   1
      Left            =   450
      Picture         =   "frmMain.frx":AFE4C
      Top             =   8160
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDimmed 
      Height          =   315
      Index           =   0
      Left            =   90
      Picture         =   "frmMain.frx":B03D0
      Top             =   8160
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgEnabled 
      Height          =   315
      Index           =   2
      Left            =   810
      Picture         =   "frmMain.frx":B0954
      Top             =   7800
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgEnabled 
      Height          =   315
      Index           =   1
      Left            =   450
      Picture         =   "frmMain.frx":B0ED8
      Top             =   7800
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgEnabled 
      Height          =   315
      Index           =   0
      Left            =   90
      Picture         =   "frmMain.frx":B145C
      Top             =   7800
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDisabled 
      Height          =   315
      Index           =   2
      Left            =   810
      Picture         =   "frmMain.frx":B19E0
      Top             =   7440
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDisabled 
      Height          =   315
      Index           =   1
      Left            =   450
      Picture         =   "frmMain.frx":B1F64
      Top             =   7440
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDisabled 
      Height          =   315
      Index           =   0
      Left            =   90
      Picture         =   "frmMain.frx":B24E8
      Top             =   7440
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgMinButton 
      Height          =   285
      Left            =   2430
      Picture         =   "frmMain.frx":B2A6C
      Stretch         =   -1  'True
      Top             =   7470
      Width           =   285
   End
   Begin VB.Image imgTopRight 
      Height          =   390
      Left            =   9960
      Picture         =   "frmMain.frx":B2FF0
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   " CS 錄音系統介接 Gateway"
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
      TabIndex        =   36
      Top             =   120
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Image imgTopLeft 
      Height          =   390
      Left            =   0
      Picture         =   "frmMain.frx":B34BC
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgTop 
      Height          =   390
      Left            =   210
      Picture         =   "frmMain.frx":B3988
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblCaptionShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   " CS 錄音系統介接 Gateway"
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
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   2025
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PR : Hammer
' Start date : 2004/12/21
' Last Modify : 2004/12/22

Option Explicit

'RCDpara.ini 的 Section , Key 的定義

'[RcdSysPara]
'ServerName = ""
'ServerIP = ""
'ServerPort = ""

'[SysSetup]
'PeakTime = ""
'OffPeak = ""
'PeakTimeStep = ""
'OffPeakStep = ""

'[Permission]
'ChkCode = ""
'KeyWord = ""

'[DB_SID]
'KeyA = ""
'KeyB = ""
'KeyC = ""

Dim mouseX, mouseY As Integer

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
    txtServer = Decrypt(Read_From_INI("RcdSysPara", "ServerName"))
    txtSrvIP = Decrypt(Read_From_INI("RcdSysPara", "ServerIP"))
    txtSrvPort = Decrypt(Read_From_INI("RcdSysPara", "ServerPort"))
    mskPeakTime1 = Decrypt(Read_From_INI("SysSetup", "PeakTime"))
    mskPeakTime2 = Decrypt(Read_From_INI("SysSetup", "OffPeak"))
    txtPeakTime = Decrypt(Read_From_INI("SysSetup", "PeakTimeStep"))
    txtOffPeak = Decrypt(Read_From_INI("SysSetup", "OffPeakStep"))
    txtDSN = Decrypt(Read_From_INI("DB_SID", "KeyA"))
    txtUID = Decrypt(Read_From_INI("DB_SID", "KeyB"))
    txtPwd = Decrypt(Read_From_INI("DB_SID", "KeyC"))
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
        Update_Permission = Update_Permission And (Write_2_INI("Permission", "ChkCode", Encrypt("ChouYinDer")) = 1)
    Else
        Update_Permission = Update_Permission And (Write_2_INI("Permission", "ChkCode", Encrypt(CStr(Date))) = 1)
    End If
    Update_Permission = Update_Permission And (Write_2_INI("Permission", "KeyWord", Encrypt(txtPassword)) = 1)
  Exit Function
ChkErr:
    ErrHandle Name, "Function : Update_Permission"
End Function

Private Function Update_Sys_Para() As Boolean ' 將系統參數頁籤之欄位值寫回INI檔的系統參數
  On Error GoTo ChkErr
    Update_Sys_Para = True
    Update_Sys_Para = Update_Sys_Para And (Write_2_INI("RcdSysPara", "ServerName", Encrypt(txtServer)) = 1)
    Update_Sys_Para = Update_Sys_Para And (Write_2_INI("RcdSysPara", "ServerIP", Encrypt(txtSrvIP)) = 1)
    Update_Sys_Para = Update_Sys_Para And (Write_2_INI("RcdSysPara", "ServerPort", Encrypt(txtSrvPort)) = 1)
    Update_Sys_Para = Update_Sys_Para And (Write_2_INI("SysSetup", "PeakTime", Encrypt(mskPeakTime1)) = 1)
    Update_Sys_Para = Update_Sys_Para And (Write_2_INI("SysSetup", "OffPeak", Encrypt(mskPeakTime2)) = 1)
    Update_Sys_Para = Update_Sys_Para And (Write_2_INI("SysSetup", "PeakTimeStep", Encrypt(txtPeakTime)) = 1)
    Update_Sys_Para = Update_Sys_Para And (Write_2_INI("SysSetup", "OffPeakStep", Encrypt(txtOffPeak)) = 1)
    Update_Sys_Para = Update_Sys_Para And (Write_2_INI("DB_SID", "KeyA", Encrypt(txtDSN)) = 1)
    Update_Sys_Para = Update_Sys_Para And (Write_2_INI("DB_SID", "KeyB", Encrypt(txtUID)) = 1)
    Update_Sys_Para = Update_Sys_Para And (Write_2_INI("DB_SID", "KeyC", Encrypt(txtPwd)) = 1)
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
    If Len(txtServer) = 0 Then txtServer = Get_Computer_Name
    If Len(txtSrvIP) = 0 Then txtSrvIP = Get_Local_IP(txtServer)
    If Len(txtSrvPort) = 0 Then txtSrvPort = "940"
    If Len(mskPeakTime1) = 0 Or mskPeakTime1 = "0000" Then mskPeakTime1 = "0730"
    If Len(mskPeakTime2) = 0 Or mskPeakTime2 = "0000" Then mskPeakTime2 = "1130"
    If Len(txtPeakTime) = 0 Or txtPeakTime = "0" Then txtPeakTime = "300"
    If Len(txtOffPeak) = 0 Or txtOffPeak = "0" Then txtOffPeak = "3600"
    Sys_Para_EditMode True
  Exit Sub
ChkErr:
    ErrHandle Name, "Click Event : cmdEditSysPara_Click"
End Sub

Private Sub cmdEditPermission_Click() ' 權限管理頁籤之修改按紐
  On Error GoTo ChkErr
    Keep_Original_Permission
    Permission_EditMode True
  Exit Sub
ChkErr:
    ErrHandle Name, "Click Event : cmdEditPermission_Click"
End Sub

Private Sub Keep_Original_Sys_Para() ' 保留系統參數頁籤之畫面值
  On Error Resume Next
    txtServer.Tag = txtServer
    txtSrvIP.Tag = txtSrvIP
    txtSrvPort.Tag = txtSrvPort
    mskPeakTime1.Tag = mskPeakTime1
    mskPeakTime2.Tag = mskPeakTime2
    txtPeakTime.Tag = txtPeakTime
    txtOffPeak.Tag = txtOffPeak
End Sub

Private Sub Revert_Original_Sys_Para() ' 還原系統參數頁籤之畫面值
  On Error Resume Next
    txtServer = txtServer.Tag
    txtSrvIP = txtSrvIP.Tag
    txtSrvPort = txtSrvPort.Tag
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
End Sub

Private Sub Permission_EditMode(blnModify As Boolean) ' 權限管理頁籤之編輯模式畫面控管
  On Error Resume Next
    fraPassword.Enabled = blnModify
    cmdEditPermission.Enabled = Not blnModify
    cmdUpdatePermission.Enabled = blnModify
    cmdUndoPermission.Enabled = blnModify
End Sub

Private Sub Sys_Para_EditMode(blnModify As Boolean) ' 系統參數頁籤之編輯模式畫面控管
  On Error Resume Next
    fraRcdSys.Enabled = blnModify
    fraSetup.Enabled = blnModify
    fraDatabase.Enabled = blnModify
    cmdEditSysPara.Enabled = Not blnModify
    cmdUpdateSysPara.Enabled = blnModify
    cmdUndoSysPara.Enabled = blnModify
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    If Len(mskPeakTime1) = 0 Then mskPeakTime1 = "0000"
    If Len(mskPeakTime2) = 0 Then mskPeakTime2 = "0000"
    If Not Visible Then XPstyle
    GradientTab
    tabMain_Click 0
    Visible = True
  Exit Sub
ChkErr:
    ErrHandle Name, "Load Event : Form_Load"
End Sub

Private Sub tabMain_Click(PreviousTab As Integer) ' Tab 切換時讀取該頁籤資料
  On Error GoTo ChkErr
    Select Case tabMain.Tab
             Case 0
             Case 1
                    Read_Sys_Para
                    lblMsg = " [ 錄音系統 ]、[ 系統設定 ]、[ 共用區資料庫連線設定 ] 所有欄位值都要正確設定 , 系統才能正常運作 !"
             Case 2
                    Read_Permission
                    lblMsg = " 當 [ 使用密碼啟動系統 ] 選取 ( 打勾 ) 時 , 則需輸入[ 輸入密碼 ]、[ 確認密碼] , 之後按 [ 存檔 ] 即完成成設定 !"
    End Select
  Exit Sub
ChkErr:
    ErrHandle Name, "Click Event : tabMain_Click"
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

Private Sub LoadPic(ByRef objIMG As Image)
  On Error Resume Next
    With objIMG
        If .Picture = Empty Then
            Set .Picture = picBackground.Picture
            .Width = picBackground.Width
            .Height = picBackground.Height
        End If
    End With
End Sub

Private Sub XPstyle()
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
        mouseX = X
        mouseY = Y
    End If
End Sub

Private Sub imgTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    If Button = 1 Then
        Me.Left = Me.Left + X - mouseX
        Me.Top = Me.Top + Y - mouseY
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

Private Sub lblCaption_Change()
  On Error Resume Next
    lblCaptionShadow.Caption = lblCaption.Caption
    SetWindowText Me.hWnd, lblCaption.Caption
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgTop_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgTop_MouseMove Button, Shift, X, Y
End Sub

Private Sub GradientTab()
  On Error GoTo ChkErr
'    PaintTabPicture
    With tabMain
        SetStyle .hWnd, 2
        SetGradientDir .hWnd, 1
        SetGradientColor1 .hWnd, vbWhite
        SetGradientColor2 .hWnd, &HFF8080
        SSTabSubclass .hWnd
        RedrawWindow .hWnd, ByVal 0&, ByVal 0&, &H1
   '     SetWindowLong .hWnd, (-4), oldWndProc
    End With
    LoadPic imgPermission
    LoadPic imgSysPara
  Exit Sub
ChkErr:
    ErrHandle Name, "Subroutine : GradientTab"
End Sub

'Private Sub ScalePicture(ByRef Pic As StdPicture, ByRef intWidth As Integer, ByRef intHeight As Integer)
'  On Error GoTo ChkErr
'    With Pic
'        intWidth = ScaleX(.Width, vbHimetric, vbPixels)
'        intHeight = ScaleY(.Height, vbHimetric, vbPixels)
'    End With
'  Exit Sub
'ChkErr:
'    ErrHandle Name, "ScalePicture"
'End Sub
'
'Private Sub PaintTabPicture()
'  On Error GoTo ChkErr
'    Dim intWidth As Integer
'    Dim intHeight As Integer
'    With tabMain
'        SetStyle .hWnd, cPicture
'        ScalePicture picBackground, intWidth, intHeight
'        SetPicture .hWnd, intWidth, intHeight, picBackground
'    End With
'  Exit Sub
'ChkErr:
'    ErrHandle Name, "PaintTabPicture"
'End Sub

'Private Sub imgMaxButton_Click()
'  On Error Resume Next
'    Select Case Me.WindowState
'           Case 0
'                Me.WindowState = 2
'           Case 2
'                Me.WindowState = 0
'    End Select
'End Sub
'
'Private Sub imgMaxButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  On Error Resume Next
'    imgMaxButton.Picture = imgDisabled(0).Picture
'End Sub
'
'Private Sub imgMaxButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  On Error Resume Next
'    imgMaxButton.Picture = imgEnabled(0).Picture
'End Sub

'Private Sub imgTop_DblClick()
'  On Error GoTo ChkErr
'    If imgMaxButton.Enabled And imgMaxButton.Visible Then
'        imgMaxButton_Click
'    End If
'  Exit Sub
'ChkErr:
'    ErrHandle Me.Name, "imgTop_DblClick"
'End Sub
