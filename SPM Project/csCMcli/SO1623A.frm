VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Object = "{4C60A153-2E5D-4072-B263-967F2D9A9A1B}#2.0#0"; "csIPtxtbox.ocx"
Begin VB.Form frmSO1623A 
   BorderStyle     =   1  '單線固定
   Caption         =   "CM 設定控制台 [SO1171A]"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
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
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   3500
      ScaleHeight     =   825
      ScaleWidth      =   4995
      TabIndex        =   64
      Top             =   2430
      Visible         =   0   'False
      Width           =   5025
      Begin MSComctlLib.ProgressBar pbr1 
         Height          =   285
         Left            =   150
         TabIndex        =   66
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
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   150
         TabIndex        =   65
         Top             =   110
         Width           =   4695
      End
   End
   Begin VB.PictureBox picResv 
      Appearance      =   0  '平面
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7500
      ScaleHeight     =   285
      ScaleWidth      =   2550
      TabIndex        =   68
      Top             =   3480
      Width           =   2550
      Begin Gi_Time.GiTime gdtResvTime 
         Height          =   285
         Left            =   930
         TabIndex        =   61
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         TabIndex        =   69
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.Timer Timer1 
      Left            =   -360
      Top             =   6780
   End
   Begin VB.CheckBox chkSetAll 
      BackColor       =   &H00C0FFFF&
      Caption         =   "所有CM一起設定"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   -1560
      TabIndex        =   67
      Top             =   390
      Visible         =   0   'False
      Width           =   1695
   End
   Begin prjNumber.GiNumber ginCustId 
      Height          =   315
      Left            =   5460
      TabIndex        =   47
      Top             =   30
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      WithComma       =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      AllowZero       =   0   'False
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消 (&X)"
      Height          =   315
      Left            =   10530
      TabIndex        =   58
      Top             =   30
      Width           =   1245
   End
   Begin VB.TextBox txtCustName 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6420
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   30
      Width           =   2985
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   975
      TabIndex        =   46
      Top             =   30
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
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
      Height          =   3795
      Left            =   120
      TabIndex        =   51
      Top             =   3420
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   6694
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   564
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      ForeColor       =   8388736
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "        設 定 作 業          "
      TabPicture(0)   =   "SO1623A.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "sstData"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSet"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "        設 定 記 錄          "
      TabPicture(1)   =   "SO1623A.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdQry"
      Tab(1).Control(1)=   "cmdResend"
      Tab(1).Control(2)=   "cmdQuery2"
      Tab(1).Control(3)=   "cmdQuery"
      Tab(1).Control(4)=   "ggrQuery"
      Tab(1).ControlCount=   5
      Begin VB.CommandButton cmdQry 
         Caption         =   "查詢指令狀態"
         Height          =   315
         Left            =   -66660
         TabIndex        =   56
         Top             =   3405
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdResend 
         Caption         =   "重 送"
         Height          =   315
         Left            =   -64830
         TabIndex        =   57
         Top             =   3405
         Width           =   1125
      End
      Begin VB.CommandButton cmdQuery2 
         Caption         =   "查詢已拆設備設定"
         Height          =   315
         Left            =   -73770
         TabIndex        =   55
         Top             =   3405
         Width           =   1905
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查 詢 (F3)"
         Height          =   315
         Left            =   -74910
         TabIndex        =   54
         Top             =   3405
         Width           =   1125
      End
      Begin prjGiGridR.GiGridR ggrQuery 
         Height          =   3315
         Left            =   -74910
         TabIndex        =   53
         Top             =   75
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5847
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
         Height          =   315
         Left            =   10050
         TabIndex        =   44
         Top             =   60
         Width           =   1245
      End
      Begin TabDlg.SSTab sstData 
         Height          =   3615
         Left            =   30
         TabIndex        =   52
         Top             =   105
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   6376
         _Version        =   393216
         TabsPerRow      =   5
         TabHeight       =   529
         WordWrap        =   0   'False
         ShowFocusRect   =   0   'False
         Enabled         =   0   'False
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
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
         TabCaption(2)   =   "&3. 即時查詢CM資訊"
         TabPicture(2)   =   "SO1623A.frx":04B2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraData(2)"
         Tab(2).ControlCount=   1
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
            Height          =   3195
            Index           =   1
            Left            =   -74880
            TabIndex        =   73
            Top             =   330
            Width           =   11055
            Begin Gi_Date.GiDate gdaProbationStopDate 
               Height          =   330
               Left            =   7665
               TabIndex        =   36
               Top             =   2820
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   582
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&B. 終止試用"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   10
               Left            =   8790
               TabIndex        =   37
               Top             =   2850
               Width           =   1305
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&A. 試用速率"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   9
               Left            =   510
               TabIndex        =   34
               Top             =   2850
               Width           =   1845
            End
            Begin VB.TextBox txtHFCNode 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3480
               TabIndex        =   21
               Top             =   1500
               Width           =   1635
            End
            Begin VB.ComboBox cboCPE 
               Height          =   345
               Left            =   3480
               TabIndex        =   29
               Top             =   2175
               Width           =   2295
            End
            Begin VB.ComboBox cboIP 
               Height          =   345
               Left            =   3480
               TabIndex        =   26
               Top             =   1815
               Width           =   2295
            End
            Begin VB.TextBox txtDelFixIP 
               Enabled         =   0   'False
               Height          =   330
               Left            =   4170
               TabIndex        =   39
               Top             =   3525
               Visible         =   0   'False
               Width           =   2295
            End
            Begin VB.TextBox txtAddFixIP 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   4440
               TabIndex        =   16
               ToolTipText     =   "多組時"
               Top             =   810
               Width           =   2295
            End
            Begin VB.TextBox txtDelCPE 
               Enabled         =   0   'False
               Height          =   330
               Left            =   7395
               TabIndex        =   40
               Top             =   3525
               Visible         =   0   'False
               Width           =   2295
            End
            Begin VB.TextBox txtAddCPE 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7665
               TabIndex        =   17
               Top             =   810
               Width           =   2295
            End
            Begin VB.CommandButton cmdFindIP 
               Height          =   360
               Left            =   7290
               Picture         =   "SO1623A.frx":04CE
               Style           =   1  '圖片外觀
               TabIndex        =   23
               Top             =   1470
               Width           =   360
            End
            Begin VB.TextBox txtCPE 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7665
               MaxLength       =   20
               TabIndex        =   30
               Top             =   2145
               Width           =   2295
            End
            Begin VB.TextBox txtFaciSNo 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7665
               MaxLength       =   20
               TabIndex        =   33
               Top             =   2490
               Width           =   2295
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&2. 申請動態IP"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   1
               Left            =   495
               TabIndex        =   10
               Top             =   540
               Width           =   1845
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&1. 速率升降級"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   0
               Left            =   495
               TabIndex        =   8
               Top             =   210
               Value           =   -1  'True
               Width           =   1845
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&3. 取消動態IP"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   2
               Left            =   4425
               TabIndex        =   12
               Top             =   540
               Width           =   1425
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&4. 申請固定IP"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   3
               Left            =   495
               TabIndex        =   14
               Top             =   870
               Width           =   1845
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&5. 取消固定IP"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   4
               Left            =   495
               TabIndex        =   18
               Top             =   1200
               Width           =   1845
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&6. EMTA IP異動 (移機)"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   5
               Left            =   495
               TabIndex        =   20
               Top             =   1530
               Width           =   2085
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&7. CPE IP 異動"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   6
               Left            =   495
               TabIndex        =   25
               Top             =   1860
               Width           =   1845
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&8. CPE MAC 異動"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   7
               Left            =   495
               TabIndex        =   28
               Top             =   2190
               Width           =   1725
            End
            Begin VB.TextBox txtAddDynaIP 
               Alignment       =   1  '靠右對齊
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3480
               TabIndex        =   11
               Top             =   465
               Width           =   645
            End
            Begin VB.TextBox txtFixIPcnt1 
               Alignment       =   1  '靠右對齊
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3480
               TabIndex        =   15
               Top             =   795
               Width           =   645
            End
            Begin VB.TextBox txtFixIPcnt2 
               Alignment       =   1  '靠右對齊
               Enabled         =   0   'False
               Height          =   330
               Left            =   3210
               TabIndex        =   38
               Top             =   3525
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtDelDynaIP 
               Alignment       =   1  '靠右對齊
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   6720
               TabIndex        =   13
               Top             =   480
               Width           =   705
            End
            Begin VB.TextBox txtIPAddress 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   5610
               TabIndex        =   22
               Top             =   1485
               Width           =   1635
            End
            Begin VB.CheckBox chkPR 
               Caption         =   "移機中是否取用新地址資訊"
               ForeColor       =   &H000000C0&
               Height          =   225
               Left            =   7710
               TabIndex        =   24
               Top             =   1530
               Width           =   2505
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&9. CM/ EMTA 更換"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   8
               Left            =   495
               TabIndex        =   31
               Top             =   2520
               Width           =   1845
            End
            Begin prjGiList.GiList gilBaudRate 
               Height          =   285
               Left            =   3480
               TabIndex        =   9
               Top             =   180
               Width           =   2730
               _ExtentX        =   4815
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FldWidth1       =   600
               FldWidth2       =   1800
               F5Corresponding =   -1  'True
            End
            Begin prjGiList.GiList gilFaciCode 
               Height          =   285
               Left            =   3480
               TabIndex        =   32
               Top             =   2520
               Width           =   3075
               _ExtentX        =   5424
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
               FldWidth1       =   750
               FldWidth2       =   2000
               F2Corresponding =   -1  'True
            End
            Begin CS_Multi.CSmulti csmAddIP 
               Height          =   345
               Left            =   6210
               TabIndex        =   83
               Top             =   150
               Visible         =   0   'False
               Width           =   4515
               _ExtentX        =   7964
               _ExtentY        =   609
               ButtonCaption   =   "新 IP"
            End
            Begin CS_Multi.CSmulti csmDelIP 
               Height          =   375
               Left            =   3450
               TabIndex        =   19
               Top             =   1125
               Width           =   6555
               _ExtentX        =   11562
               _ExtentY        =   661
               ButtonCaption   =   "IP , CPE MAC"
            End
            Begin IPvalidatePrj.IP_TextBox txtCPEIP 
               Height          =   315
               Left            =   7665
               TabIndex        =   27
               Top             =   1815
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   556
               BorderStyle     =   1
            End
            Begin prjGiList.GiList gilNewRate 
               Height          =   285
               Left            =   3480
               TabIndex        =   35
               Top             =   2820
               Width           =   2730
               _ExtentX        =   4815
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FldWidth1       =   600
               FldWidth2       =   1800
               F5Corresponding =   -1  'True
            End
            Begin VB.Label lblProbationStopDate 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "預計試用截止日"
               Height          =   225
               Left            =   6330
               TabIndex        =   93
               Top             =   2850
               Width           =   1260
            End
            Begin VB.Label lblNewRate 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "新速率"
               Height          =   225
               Left            =   2895
               TabIndex        =   92
               Top             =   2850
               Width           =   540
            End
            Begin VB.Label lblNode 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "新NODE"
               Height          =   225
               Left            =   2730
               TabIndex        =   91
               Top             =   1530
               Width           =   705
            End
            Begin VB.Label lblOldCPE 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "舊CPE"
               Height          =   225
               Left            =   2880
               TabIndex        =   90
               Top             =   2220
               Width           =   555
            End
            Begin VB.Label lblOldIP 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "舊IP"
               Height          =   225
               Left            =   3090
               TabIndex        =   89
               Top             =   1875
               Width           =   345
            End
            Begin VB.Label lblCPE 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "新CPE"
               Height          =   225
               Left            =   7050
               TabIndex        =   88
               Top             =   2220
               Width           =   555
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "IP"
               Enabled         =   0   'False
               Height          =   225
               Left            =   3960
               TabIndex        =   87
               Top             =   3540
               Visible         =   0   'False
               Width           =   165
            End
            Begin VB.Label lblAddIP 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "IP"
               Height          =   225
               Left            =   4230
               TabIndex        =   86
               Top             =   870
               Width           =   165
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "CPE MAC"
               Enabled         =   0   'False
               Height          =   225
               Left            =   6540
               TabIndex        =   85
               Top             =   3540
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.Label lblCPEMAC 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "CPE MAC"
               Height          =   225
               Left            =   6720
               TabIndex        =   84
               Top             =   870
               Width           =   795
            End
            Begin VB.Label lblCPEIP 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "新IP"
               Height          =   225
               Left            =   7260
               TabIndex        =   82
               Top             =   1890
               Width           =   345
            End
            Begin VB.Label lblCMFacisno 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "CM 序號"
               Height          =   225
               Left            =   6930
               TabIndex        =   81
               Top             =   2550
               Width           =   675
            End
            Begin VB.Label lblAddFixIP 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "固定IP"
               Height          =   225
               Left            =   2910
               TabIndex        =   80
               Top             =   885
               Width           =   525
            End
            Begin VB.Label lblBaudRate 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "新速率"
               Height          =   225
               Left            =   2895
               TabIndex        =   79
               Top             =   210
               Width           =   540
            End
            Begin VB.Label lblDelFixIP 
               AutoSize        =   -1  'True
               Caption         =   "固定IP"
               Enabled         =   0   'False
               Height          =   225
               Left            =   2640
               TabIndex        =   78
               Top             =   3585
               Visible         =   0   'False
               Width           =   525
            End
            Begin VB.Label lblIPAddress 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "新IP"
               Height          =   225
               Left            =   5220
               TabIndex        =   77
               Top             =   1530
               Width           =   345
            End
            Begin VB.Label lblAddDynIP 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "動態IP"
               Height          =   225
               Left            =   2910
               TabIndex        =   76
               Top             =   540
               Width           =   525
            End
            Begin VB.Label lblDelDynIP 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "動態IP"
               Height          =   225
               Left            =   6135
               TabIndex        =   75
               Top             =   540
               Width           =   525
            End
            Begin VB.Label lblFaci 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "品名"
               Height          =   225
               Left            =   3075
               TabIndex        =   74
               Top             =   2550
               Width           =   360
            End
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
            Height          =   3255
            Index           =   0
            Left            =   120
            TabIndex        =   72
            Top             =   330
            Width           =   11055
            Begin VB.OptionButton optBasic 
               Caption         =   "&8. CP 拆分線"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   7
               Left            =   5055
               TabIndex        =   7
               Top             =   1590
               Width           =   2266
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&7. CP 裝分線"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   6
               Left            =   5055
               TabIndex        =   6
               Top             =   1200
               Width           =   2266
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&4. 軟復"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   3
               Left            =   525
               TabIndex        =   3
               Top             =   1590
               Width           =   2266
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&2. 拆機 ( 含裝機退件 )"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   1
               Left            =   525
               TabIndex        =   1
               Top             =   810
               Width           =   2266
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&1. 裝機"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   0
               Left            =   525
               TabIndex        =   0
               Top             =   420
               Value           =   -1  'True
               Width           =   2266
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&5. CP Only 裝機"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   4
               Left            =   5055
               TabIndex        =   4
               Top             =   420
               Width           =   2266
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&3. 軟拆"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   2
               Left            =   525
               TabIndex        =   2
               Top             =   1200
               Width           =   2266
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&6. CP Only 拆機"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   5
               Left            =   5055
               TabIndex        =   5
               Top             =   810
               Width           =   2266
            End
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
            Height          =   3165
            Index           =   2
            Left            =   -74880
            TabIndex        =   45
            Top             =   360
            Width           =   11025
            Begin VB.OptionButton optQueryInfo 
               Caption         =   "&2. 查詢 CM 參數"
               ForeColor       =   &H00800000&
               Height          =   285
               Index           =   1
               Left            =   105
               TabIndex        =   42
               Top             =   690
               Value           =   -1  'True
               Width           =   1755
            End
            Begin Gi_Time.GiTime gdtStartTime 
               Height          =   285
               Left            =   3390
               TabIndex        =   59
               Top             =   -210
               Visible         =   0   'False
               Width           =   1875
               _ExtentX        =   3307
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
               Caption         =   "&1. 重設 CM Reset"
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   0
               Left            =   105
               TabIndex        =   41
               Top             =   330
               Width           =   1755
            End
            Begin prjGiGridR.GiGridR ggrQueryInfo 
               Height          =   2925
               Left            =   1860
               TabIndex        =   43
               Top             =   180
               Width           =   9015
               _ExtentX        =   15901
               _ExtentY        =   5159
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
               Left            =   5640
               TabIndex        =   60
               Top             =   -210
               Visible         =   0   'False
               Width           =   1875
               _ExtentX        =   3307
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
               Left            =   5370
               TabIndex        =   71
               Top             =   -150
               Visible         =   0   'False
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
               Left            =   2580
               TabIndex        =   70
               Top             =   -150
               Visible         =   0   'False
               Width           =   975
            End
         End
      End
   End
   Begin prjGiGridR.GiGridR ggrChild 
      Height          =   1245
      Left            =   120
      TabIndex        =   50
      Top             =   2190
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   2196
      Enabled         =   0   'False
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
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   1845
      Left            =   120
      TabIndex        =   49
      Top             =   360
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   3254
      Enabled         =   0   'False
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseCellForeColor=   -1  'True
   End
   Begin VB.Label lblCustId 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "客 戶 編 號"
      Height          =   225
      Left            =   4500
      TabIndex        =   63
      Top             =   75
      Width           =   855
   End
   Begin VB.Label lblComp 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   150
      TabIndex        =   62
      Top             =   75
      Width           =   720
   End
End
Attribute VB_Name = "frmSO1623A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents rsData As ADODB.Recordset
Attribute rsData.VB_VarHelpID = -1
Private WithEvents rsChild As ADODB.Recordset
Attribute rsChild.VB_VarHelpID = -1
Private WithEvents rsQuery As ADODB.Recordset
Attribute rsQuery.VB_VarHelpID = -1
Private WithEvents rsLog As ADODB.Recordset
Attribute rsLog.VB_VarHelpID = -1

'Private rsParent As New ADODB.Recordset

Public strNodeNo As String
Public strCircuitNO As String
Public blnGetIP As Boolean

Private lngCustId As Long

Private strChStr As String, strStopChStr As String, strOpenChStr As String
Private strChStopDate As String
Private rsParent As New ADODB.Recordset

Private blnPrcOK As Boolean
Private strProcessChStr As String
Private strDoCaption As String
Private strCaption As String
Private blnTransation As Boolean
Private blnSetting As Boolean
Private strDefaultRowId As String
Private strOrderNo As String
Private blnProcessLock As Boolean
Private intByHouse As Integer
Private inResvTime As String
Public strIPaddr As String ' 舊的 IP Address
 
Private strReturnNumber As String
Private strIPAddress As String
Private strHFC As String

Private strCmd As String

'Private strSno As String
'Private strMediaBillNo As String

'Private Sub cboCPE_DropDown()
'  On Error Resume Next
'    Dim varData As Variant
'    Dim varElement As Variant
'    With cboCPE
'            If rsData.State > 0 Then
'                .Clear
'                varData = gcnGi.Execute("SELECT DISTINCT CPEMAC FROM " & GetOwner & "SO004C" & _
'                                                            " WHERE FACISNO = '" & rsData("FaciSNo") & "'" & _
'                                                            " AND CUSTID = " & rsData("CustID") & _
'                                                            " AND STOPDATE IS NULL").GetRows
'                For Each varElement In varData
'                    If varElement & "" <> Empty Then .AddItem varElement
'                Next
'            End If
'    End With
'End Sub

'Private Sub cboNode_Click()
'  On Error GoTo ChkErr
'    Dim strOri As String
'    If cboNode = Empty Then Exit Sub
'    strOri = IIf(int_EMTA_IP_Type = 0, strCircuitNO, strNodeNo)
'    If cboNode <> strOri Then
'        If blnGetIP And strIPaddr <> Empty Then gcnGi.Execute "UPDATE " & GetOwner & "SO048 SET USEFLAG=0 WHERE IPADDRESS='" & strIPaddr & "'"
'        strIPaddr = GetNewIP
'        txtNewIP = strIPaddr
'    Else
'        txtNewIP = rsData("IPAddress") & ""
'    End If
'  Exit Sub
'ChkErr:
'    ErrSub Name, "cboNode_Click"
'End Sub


'Private Function GetNewIP() As String
'  On Error Resume Next
'    Dim lngRet As Long
'    GetNewIP = GetRsValue("SELECT IPADDRESS FROM " & GetOwner & "SO048" & _
'                                            " WHERE CIRCUITNO = '" & cboNode & "'" & _
'                                            " AND COMPCODE=" & gilCompCode.GetCodeNo & _
'                                            " AND USEFLAG=0") & ""
'    If Err.Number = 0 Then
'        If GetNewIP <> Empty Then
'            gcnGi.Execute "UPDATE " & GetOwner & "SO048 SET USEFLAG=1 WHERE IPADDRESS='" & GetNewIP & "'", lngRet
'            If lngRet > 0 And GetNewIP <> Empty Then
'                blnGetIP = True
'            End If
'        End If
'    Else
'        Err.Clear
'        GetNewIP = ""
'    End If
'  Exit Function
'ChkErr:
'    ErrSub Name, "GetNewIP"
'End Function

Private Function GetHFC(lngAddrNo As Long) As String
  On Error Resume Next
    If lngAddrNo <> 0 Then
        GetHFC = GetRsValue("SELECT " & IIf(int_EMTA_IP_Type = 0, "CIRCUITNO", "NODENO") & _
                                            " FROM " & GetOwner & "SO014" & _
                                            " WHERE ADDRNO = " & lngAddrNo & _
                                            " AND COMPCODE = " & gilCompCode.GetCodeNo)
        If Err.Number <> 0 Then GetHFC = ""
    End If
End Function

'Private Function GetIPAddress(lngAddrNo As Long, strHFC As String, strNewIP As String) As Boolean
'  On Error Resume Next
''    If Is_EMTA_Dev Then
'        strField = IIf(int_EMTA_IP_Type = 0, "CIRCUITNO", "NODENO")
'        GetIPAddress = GetRS("SELECT CIRCUITNO,IPADDRESS FROM " & GetOwner & "SO048" & _
'                                                " WHERE CIRCUITNO = " & _
'                                                "(SELECT " & IIf(int_EMTA_IP_Type = 0, "CIRCUITNO", "NODENO") & _
'                                                " FROM " & GetOwner & "SO014" & _
'                                                " WHERE ADDRNO = " & lngAddrNo & _
'                                                " AND COMPCODE = " & gilCompCode.GetCodeNo & ") " & _
'                                                " AND COMPCODE=" & gilCompCode.GetCodeNo & _
'                                                " AND USEFLAG=0") & ""
'        If Err.Number <> 0 Then GetIPAddress = ""
''    End If
'End Function

'Private Sub cboNode_DropDown()
'  On Error GoTo ChkErr
'    Dim varData As Variant
'    Dim varElement As Variant
'
'    'If Mid(GetRsValue("SELECT CHKADRSMODIFY FROM " & GetOwner & "SO044 WHERE COMPCODE=" & gCompCode & " AND SERVICETYPE='" & gServiceType & "'", gcnGi, "SO044 [預設參數檔] 查無公司別為 " & gCompCode & " , 服務別為 " & gServiceType & " 的資料!! "), 2, 1) = 1 Then
'        '   "SELECT ROWNUM,CIRCUITNO FROM (SELECT DISTINCT CIRCUITNO FROM " & GetOwner & "SO016 SO016 WHERE CIRCUITNO IS NOT NULL)", "流水編號", "網路編號"
'    'Else
'        '   "SELECT ROWNUM,CIRCUITNO FROM (SELECT DISTINCT CIRCUITNO FROM " & GetOwner & "SO014 SO014 WHERE CIRCUITNO IS NOT NULL)", "流水編號", "網路編號"
'    'End If
'
'    With cboNode
'            .Clear
'            ' SO041.EMTAIPTYPE ( 0=網路編號 , 1=Node )
''            If GetSystemParaItem("EMTAIPTYPE", gCompCode, gServiceType, "SO041", , True) = 0 Then
'            If int_EMTA_IP_Type = 0 Then
'                'select distinct nodeno from so016 where nodeno is not null
'                varData = gcnGi.Execute("SELECT DISTINCT CIRCUITNO FROM " & GetOwner & "SO016" & _
'                                                        " WHERE COMPCODE=" & gilCompCode.GetCodeNo & _
'                                                        " AND NODENO IS NOT NULL").GetRows
'                For Each varElement In varData
'                    If varElement & "" <> Empty Then .AddItem varElement
'                Next
'            Else
'                ' 需不需要串 ServiceType
'                varData = gcnGi.Execute("SELECT DISTINCT CODENO FROM " & GetOwner & "CD047" & _
'                                                        " WHERE COMPCODE=" & gilCompCode.GetCodeNo & _
'                                                        " AND STOPFLAG <> 1").GetRows
'                For Each varElement In varData
'                    If varElement & "" <> Empty Then .AddItem varElement
'                Next
'            End If
'    End With
'  Exit Sub
'ChkErr:
'    ErrSub Name, ""
'End Sub

Private Sub cboCPE_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboIP_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub chkPR_Click()
  On Error GoTo ChkErr
    Dim lngNewAddrNo As Long
    If chkPR.Value Then
        lngNewAddrNo = GetNewAddr
        If lngNewAddrNo > 0 Then txtHFCNode = GetHFC(lngNewAddrNo)
'        cboNode_Click
    End If
  Exit Sub
ChkErr:
    ErrSub Name, "chkPR_Click"
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

Public Property Let uReturnValue(ByVal vData As String)
    On Error GoTo ChkErr
        strReturnNumber = vData
    Exit Property
ChkErr:
    ErrSub Me.Name, "Let uReturnValue"
End Property

Public Property Let uIPAddress(ByVal vData As String)
    strIPAddress = vData
End Property

Private Sub cboCPE_DropDown()
  On Error GoTo ChkErr
    Dim varElement, varData
    varData = gcnGi.Execute("SELECT DISTINCT CPEMAC FROM " & GetOwner & "SO004C" & _
                                            " WHERE SEQNO='" & rsData!SeqNo & "'" & _
                                            " AND CUSTID=" & rsData!CustID & _
                                            " AND STOPDATE IS NULL").GetRows
    With cboCPE
        .Clear
        For Each varElement In varData
            'Debug.Print "*" & varElement & "*"
            If varElement <> Empty Then .AddItem varElement & ""
        Next
    End With
'    varData = Get_4C_Info("CPEMAC") & ""
'    If varData = "" Then Exit Sub
'    With cboCPE
'        .Clear
'        For Each varElement In varData
'            If varElement <> Empty Then .AddItem varElement & ""
'        Next
'    End With
  Exit Sub
ChkErr:
    If Err.Number = 3021 Then
        Err.Clear
        MsgBox "SO004C 查無 [客編] 為 " & rsData!CustID & " , [序號] 為 " & rsData!SeqNo & " 的 未停用 資料 !", vbInformation, "訊息"
    Else
        ErrSub Name, "cboCPE_DropDown"
    End If
End Sub

Private Sub cboIP_DropDown()
  On Error GoTo ChkErr
    Dim varElement, varData
    varData = gcnGi.Execute("SELECT DISTINCT IPADDRESS FROM " & GetOwner & "SO004C" & _
                                            " WHERE SEQNO='" & rsData!SeqNo & "'" & _
                                            " AND CUSTID=" & rsData!CustID & _
                                            " AND STOPDATE IS NULL").GetRows
    With cboIP
        .Clear
        For Each varElement In varData
            If varElement <> Empty Then .AddItem varElement & ""
        Next
    End With
'    Dim varElement, varData
'    varData = Get_4C_Info("IPADDRESS") & ""
'    If varData = "" Then Exit Sub
'    With cboIP
'        .Clear
'        If varData Is Nothing Then
'            '.AddItem "無資料!"
'        Else
'            For Each varElement In varData
'                If varElement <> Empty Then .AddItem varElement & ""
'            Next
'        End If
'    End With
  Exit Sub
ChkErr:
    If Err.Number = 3021 Then
        Err.Clear
        MsgBox "SO004C 查無 [客編] 為 " & rsData!CustID & " , [序號] 為 " & rsData!SeqNo & " 的 未停用 資料 !", vbInformation, "訊息"
    Else
        ErrSub Name, "cboCPE_DropDown"
    End If
End Sub

'Private Function Get_4C_Info(strFld As String) As Variant
'  On Error GoTo ChkErr
'    Dim varData
'    varData = gcnGi.Execute("SELECT DISTINCT " & strFld & " FROM " & GetOwner & "SO004C" & _
'                                                    " WHERE SEQNO='" & rsData!SeqNo & "'" & _
'                                                    " AND CUSTID=" & rsData!CustID & _
'                                                    " AND STOPDATE IS NULL").GetRows
'    Get_4C_Info = varData
'  Exit Function
'ChkErr:
'    If Err.Number = 3021 Then
'        Get_4C_Info = ""
'        Err.Clear
'        MsgBox "SO004C 查無 [客編] 為 " & rsData!CustID & " , [序號] 為 " & rsData!SeqNo & " 的 未停用 資料 !", vbInformation, "訊息"
'    Else
'        ErrSub Name, "Get_4C_Info"
'    End If
'End Function

Private Sub cmdFindIP_Click()
  On Error GoTo ChkErr
    Dim lngAddrNo As Long
    If chkPR.Value Then
        lngAddrNo = GetNewAddr
        If lngAddrNo > 0 Then
            strHFC = GetHFC(lngAddrNo)
        Else
            strHFC = GetHFC(lngInstAddrNo)
            lngAddrNo = lngInstAddrNo
        End If
    Else
        strHFC = GetHFC(lngInstAddrNo)
        lngAddrNo = lngInstAddrNo
    End If
    With frmSO1120C
        Set .uParentForm = Me
        .uOldNo = strHFC
        .uType = 1
        .Show vbModal
    End With
    If strReturnNumber <> "" Then
        txtIPAddress = strReturnNumber
        strIPaddr = strReturnNumber
        blnGetIP = True
    End If
  Exit Sub
ChkErr:
    ErrSub Name, "cmdFindIP_Click"
End Sub

' F - Failure , N - New ( Not Processed) , S - Success
' 狀態失敗的可按右鍵作指令重送，產生一筆新的指令紀錄且標記舊的已重送(失敗且標記已重送的不可在重送)。

' 指令狀態為 N 的可按右鍵作指令查詢指令狀態，若回應成功需按指令功能回填相關資料
' 3.  設定紀錄，針對SPM(TBC)指令未成功的，按右鍵可以查詢指令執行結果，若SPM 已成功，需要更新指令執行狀態。

Private Sub cmdQry_Click()
  On Error GoTo ChkErr
    subProgressBar True
    lblProcess = "查詢指令狀態中, 請稍後 .."
'    Chk_Cmd True
    subProgressBar False
  Exit Sub
ChkErr:
    ErrSub Name, "cmdQry_Click"
End Sub

Private Sub cmdResend_Click()
  On Error GoTo ChkErr
    ' 重送指令新的指令流水號 : CMDRESEND
    subProgressBar True
    With rsLog
        lblProcess = "[ " & .Fields("MsgName") & " ] 指令重送中, 請稍後 .."
        Debug.Print ParaClt(rsLog, "", .Fields("CMDSeqNo"))
        RequeryRS
    End With
    subProgressBar False
'    cmdQuery.Value = True
  Exit Sub
ChkErr:
    ErrSub Name, "cmdResend_Click"
End Sub

Private Sub RequeryRS()
  On Error GoTo ChkErr
    Dim lngAbsPos As Long
    With rsLog
        lngAbsPos = .AbsolutePosition
        .Requery
        Set ggrQuery.Recordset = rsLog
        ggrQuery.Refresh
        .AbsolutePosition = lngAbsPos + 1
    End With
    With rsData
        lngAbsPos = .AbsolutePosition
        .Requery
        Set ggrData.Recordset = rsData
        ggrData.Refresh
        .AbsolutePosition = lngAbsPos
    End With
    With rsChild
        lngAbsPos = .AbsolutePosition
        If lngAbsPos > 0 Then
            .Requery
            Set ggrData.Recordset = rsChild
            ggrData.Refresh
            .AbsolutePosition = lngAbsPos
        End If
    End With
  Exit Sub
ChkErr:
    ErrSub Name, "RequeryRS"
End Sub

Private Sub cmdQuery_Click()
  On Error GoTo ChkErr
    Fetch_Log_Data
'    ggrQuery.Recordset.MoveLast
'    ggrQuery.Recordset.MoveFirst
  Exit Sub
ChkErr:
    ErrSub Name, "cmdQuery_Click"
End Sub

Private Sub Fetch_Log_Data(Optional blnPRflag As Boolean = False, Optional blnEmptyGrid As Boolean = False)
  On Error GoTo ChkErr
    Dim lngRcdCnt As Long
    Dim strQry As String
    strCmd = ""
    If blnEmptyGrid Then
        strQry = "SELECT * FROM " & GetOwner & "SO307 WHERE 0=1"
    Else
        If rsData Is Nothing Then Exit Sub
        With rsData
                If .State <= 0 Then Exit Sub
                If .EOF Or .BOF Then Exit Sub
                lngRcdCnt = .RecordCount
'                If lngRcdCnt = 0 Then Exit Sub
        End With
        strQry = "SELECT * FROM " & GetOwner & "SO307" & _
                        " WHERE COMPCODE='" & garyGi(9) & "'" & _
                        " AND CUSTID=" & ginCustId.Text
'        If lngRcdCnt = 1 Then
            If blnPRflag Then
                If rsData.EOF Then
'                    Debug.Print "No SO004 Data !"
                    cmdResend.Enabled = False
                Else
                    strQry = strQry & " AND MODEMMAC <> '" & Mask(rsData("FaciSNo") & "") & "'"
                    cmdResend.Enabled = True
                End If
            Else
                strQry = strQry & " AND MODEMMAC = '" & Mask(rsData("FaciSNo") & "") & "'"
            End If
'        Else
'            If blnPRflag Then
'                strQry = strQry & " AND MODEMMAC NOT IN (" & Get_Faci_Rows & ")"
'            Else
'                strQry = strQry & " AND MODEMMAC IN (" & Get_Faci_Rows & ")"
'            End If
'        End If
        strQry = strQry & " ORDER BY EXECTIME DESC"
    End If
    
    If GetRS(rsLog, strQry, gcnGi, adUseClient, adOpenStatic, adLockReadOnly) Then ShowResult rsLog, ggrQuery
  
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
            If varElement & "" <> Empty Then strValue = strValue & "','" & Mask(varElement & "") & ""
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
    
    Dim lngAbsPos As Long
    Dim intRefno As Long
    
    strResvTime = gdtResvTime.GetValue(True)
    
'    gilFaciCode.SelectRefNo
    If optCMChange(8).Value And Len(gilFaciCode.GetCodeNo) > 0 Then

        intRefno = Val(GetRsValue("SELECT REFNO FROM " & GetOwner & "CD022" & _
                                                        " WHERE CODENO=" & gilFaciCode.GetCodeNo & ""))
                                                        
        Select Case Val(GetRsValue("SELECT REFNO FROM " & GetOwner & "CD022" & _
                                                        " WHERE CODENO=" & rsData("FaciCode") & ""))
                Case 2
                    strCmd = IIf(intRefno = 2, "A7", "A9")
'                    strCmd = IIf(gilFaciCode.GetRefNo = 2, "A7", "A9")
                Case 5
                    strCmd = IIf(intRefno = 5, "A8", "A10")
'                    strCmd = IIf(gilFaciCode.GetRefNo = 5, "A8", "A10")
        End Select
    End If
    
    If Not IsDataOK Then Exit Sub
    
    Me.Enabled = False
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
    'If Not blnTransation Then gcnGi.BeginTrans
    lngAbsPos = rsData.AbsolutePosition
'   If chkSetAll = 1 Then rsData.MoveFirst
'   Do While Not rsData.EOF

    Select Case sstData.Tab
            Case 0
                Select Case True
                    Case optBasic(0) ' 甲、裝機 , 設備參考號 , 若是2 CM 則   C1 , 若是5 EMTACM 則   E1
                            strCmd = IIf(Is_EMTA_Dev, "E1", "C1")
                    Case optBasic(1) ' 乙、拆機 ( 含裝機退件 ) 需判斷設備參考號 , 若是2 CM 則   C2 , 若是5 EMTACM 則送E2
                            strCmd = IIf(Is_EMTA_Dev, "E2", "C2")
                    Case optBasic(2) ' 庚、CM 軟拆 A1_D
                            ' 有開機日 且開機日>(關機日OR 無關機日)則只可選軟關
                            strCmd = "A1_D"
                    Case optBasic(3) ' 辛、CM 軟復 A1_C
                            ' 軟關與軟復兩各選項是互斥，設備若(關機日OR 無關機日) > 開機日 則只可以選軟復
                            strCmd = "A1_C"
                    Case optBasic(4) ' 丙、CP裝機 E3
                            strCmd = "E3"
                    Case optBasic(5) ' 丁、CP拆機 A1_P （新的MediaGateway null）
                            strCmd = "A1_P"
                    Case optBasic(6) ' 戊、CP裝分線 E4
                            strCmd = "E4"
                    Case optBasic(7) ' 己、CP拆分線 E5
                            strCmd = "E5"
                End Select
                blnPrcOK = ParaClt(rsData, strCmd, , , , ginCustId.Text, Val(garyGi(9)))
            Case 1
                Select Case True
                    Case optCMChange(0) ' 1. 速率升降級
                            blnPrcOK = ParaClt(rsData, "A1_BU", , , , ginCustId.Text, Val(garyGi(9)), , , gilBaudRate.GetDescription)
                    
                    Case optCMChange(1) ' 2. 申請動態IP
                            blnPrcOK = ParaClt(rsData, "A1_PIP", , , , ginCustId.Text, Val(garyGi(9)), , , , Abs(Val(txtAddDynaIP)))
                    
                    Case optCMChange(2) ' 3. 取消動態IP
                            blnPrcOK = ParaClt(rsData, "A1_MIP", , , , ginCustId.Text, Val(garyGi(9)), , , , Abs(Val(txtDelDynaIP)) * -1)
                    
                    Case optCMChange(3) ' 4. 申請固定IP
                            blnPrcOK = ParaClt(rsData, "A2", , , , ginCustId.Text, Val(garyGi(9)), , , , , Abs(Val(txtFixIPcnt1)), , txtAddCPE, , txtAddFixIP)
                             
                    Case optCMChange(4) ' 5. 取消固定IP
                            ' 換成 Multi-Select
                            Dim strFix_IP_cnt As String
                            Dim strIP_Address As String
                            Dim strCPE_MAC As String
                            strIP_Address = csmDelIP.GetQryFld(1)
                            strIP_Address = Replace(strIP_Address, "'", "", 1)
                            strIP_Address = Replace(strIP_Address, ",", "#", 1)
                            strCPE_MAC = csmDelIP.GetQryFld(2)
                            strCPE_MAC = Replace(strCPE_MAC, "'", "", 1)
                            strCPE_MAC = Replace(strCPE_MAC, ",", "#", 1)
                            strFix_IP_cnt = csmDelIP.GetQueryCode
                            strFix_IP_cnt = UBound(Split(strFix_IP_cnt, ",")) + 1
                            blnPrcOK = ParaClt(rsData, "A3", , , , ginCustId.Text, Val(garyGi(9)), , , , , Abs(Val(strFix_IP_cnt)) * -1, , , strCPE_MAC, , strIP_Address)
                    
                    Case optCMChange(5) ' 6. EMTA IP 異動
                            blnGetIP = False
                            ' 增加NODE 輸入方塊，
                            ' 若選取設備屬於CM 設備參考號2的設備時，IP 的部份需要DISABLE，
                            ' 指令傳送改送C6 指令，將指令TABLE NewHFCNode 填入畫面輸入的值，
                            ' 其他欄位值與A4指令階相同，因為是參考號2 所以IPADDRESS 部份值應該是NULL。
                            
                            ' (PM_RA)_TBC CLI [10]CM 單筆資料設定作業.doc
                            ' 第10頁 A4 指令調整更換IPCOMMAND 的指令C6 與 E6 的邏輯
                            ' 當勾選移機中取用新地址時如果設備
                            ' 1.  參考號2時傳送C6 指令，舊的NODE 取用裝機地址資訊，新NODE 取用畫面輸入或移裝工單的新地址資訊
                            ' 2.  參考號5時傳送E6 指令，舊的NODE 取用裝機地址資訊，新NODE 取用畫面輸入或移裝工單的新地址資訊
                            ' 3.  未勾選時按照原機制傳送A4 指令
                            
                            If Is_EMTA_Dev Then
                                If chkPR.Value Then
                                    blnPrcOK = ParaClt(rsData, "E6", , , , ginCustId.Text, Val(garyGi(9)), , , , , , txtHFCNode, , , , , txtIPAddress)
                                Else
                                    blnPrcOK = ParaClt(rsData, "A4", , , , ginCustId.Text, Val(garyGi(9)), , , , , , strHFC, , , , , txtIPAddress)
                                End If
                            Else
                                blnPrcOK = ParaClt(rsData, "C6", , , , ginCustId.Text, Val(garyGi(9)), , , , , , txtHFCNode, , , , , "")
                            End If
                            
'                            If txtIPAddress.Enabled Then
'                                blnPrcOK = ParaClt(rsData, "A4", , , , ginCustId.Text, Val(garyGi(9)), , , , , , strHFC, , , , , txtIPAddress)
'                            Else
'                                'strCmd = IIf(Is_EMTA_Dev, "E6", "C6")
'                                If Is_EMTA_Dev Then
'                                    blnPrcOK = ParaClt(rsData, "E6", , , , ginCustId.Text, Val(garyGi(9)), , , , , , txtHFCNode, , , , , "")
'                                Else
'                                    blnPrcOK = ParaClt(rsData, "C6", , , , ginCustId.Text, Val(garyGi(9)), , , , , , txtHFCNode, , , , , "")
'                                End If
'                            End If
                            
                    Case optCMChange(6) ' 7. CPE IP 異動
                            blnPrcOK = ParaClt(rsData, "A5", , , , ginCustId.Text, Val(garyGi(9)), , , , , , , , , txtCPEIP.Text)
                            
                    Case optCMChange(7) ' 8. CPE MAC 異動
                            blnPrcOK = ParaClt(rsData, "A6", , , , ginCustId.Text, Val(garyGi(9)), , , , , , , txtCPE)
                
                    Case optCMChange(8) ' 9. CM/ EMTA 更換
'                            參考號2換新CM A7
'                            參考號2舊CM /換新參考號5 EMTA A9
'                            參考號5換新EMTA A8
'                            參考號5 EMTA /換新參考號2 CM A10
                            blnPrcOK = ParaClt(rsData, strCmd, , , , ginCustId.Text, Val(garyGi(9)), txtFaciSNo, GetModelName(txtFaciSNo))
                            
                    Case optCMChange(9) ' A. 試用速率 "TRIAL"
                            strProbStopDate = gdaProbationStopDate.Text
                            blnPrcOK = ParaClt(rsData, "TRIAL", , , , ginCustId.Text, Val(garyGi(9)), , , gilNewRate.GetDescription)
                
                    Case optCMChange(10) ' B. 終止試用 "STOPTRIAL"
                            blnPrcOK = ParaClt(rsData, "STOPTRIAL", , , , ginCustId.Text, Val(garyGi(9)))
                
                End Select
                
            Case 2
                Select Case True
                    Case optQueryInfo(0) ' 1. 重設 CM Reset
                            strCmd = "Q1"
                    Case optQueryInfo(1) ' 2. 查詢 CM 狀態
                            strCmd = "Q2"
                End Select
                blnPrcOK = ParaClt(rsData, strCmd, , , , ginCustId.Text, Val(garyGi(9)))
        End Select
'       If chkSetAll = 0 Then Exit Do
'       rsData.MoveNext
'   Loop
    If strPrcType <> "" And blnPrcOK Then
        Unload Me
    Else
        If strPrcType = "" Then
            With rsData
                If Not .ActiveConnection Is Nothing Then .Requery
            End With
            With ggrData
                If Not rsData.ActiveConnection Is Nothing Then Set .Recordset = rsData
                .Refresh
                rsData.AbsolutePosition = lngAbsPos
            End With
            If blnPrcOK Then OpenData2
        End If
    '   If chkSetAll = 1 Then Call OpenData2
        'If Not blnTransation Then gcnGi.CommitTrans
        'blnPrcOK = True
        subProgressBar False
        Enabled = True
        blnSetting = False
        If strPrcType = "" And strCmd = "Q2" Then ShowVtcFld frmSO1623A.ggrQueryInfo
        'If intProcessType > 0 Then Unload Me
        'DoProcessMode
    End If
  Exit Sub
ErrTran:
    rsData.AbsolutePosition = lngAbsPos
'   If chkSetAll = 1 Then Call OpenData2
    Call DoProcessMode
    blnSetting = False
'    If Not blnTransation Then gcnGi.RollbackTrans
    pic1.Visible = False
    Me.Enabled = True
  Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdSet_Click"
End Sub

Private Function GetDynaIP() As String
    On Error Resume Next
        Dim strDynaIPfld As String
        If rsData("BPcode") & "" = "" Then Exit Function
        strDynaIPfld = IIf(IsCMCP, "CMCPDYNIP", "DYNIPCOUNT")
        GetDynaIP = GetRsValue("SELECT " & strDynaIPfld & " FROM " & GetOwner & "CD078D WHERE BPCODE = '" & rsData("BPcode") & "'")
        If Err.Number <> 0 Then GetDynaIP = Empty
End Function

Private Function GetClassID() As String
    On Error Resume Next
        Dim strClassIdFld As String
'        If rsData("CMBaudNo") & "" = "" Then Exit Function
        If rsData("BPcode") & "" = "" Then Exit Function
        strClassIdFld = IIf(IsCMCP, "EMCCMCP", "EMCCM")
'        GetClassID = GetRsValue("SELECT " & strClassIdFld & " FROM " & GetOwner & "CD064 WHERE CODENO = " & rsData("CMBaudNo"))
        GetClassID = GetRsValue("SELECT " & strClassIdFld & " FROM " & GetOwner & "CD078D WHERE BPCODE = '" & rsData("BPcode") & "'")
        If Err.Number <> 0 Then GetClassID = Empty
End Function

Public Function IsCMCP() As Boolean
  On Error GoTo ChkErr
    IsCMCP = (GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                            " WHERE REFACISNO='" & rsData("FaciSno") & "'" & _
                            " AND COMPCODE=" & rsData("CompCode") & _
                            " AND CUSTID=" & rsData("CustID") & _
                            " AND SERVICETYPE='P'" & _
                            " AND PRDATE IS NULL") = 1)
        
'    IsCMCP = GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                        " WHERE REFACISNO='" & rsData("FaciSno") & "'" & _
                        " AND COMPCODE=" & rsData("CompCode") & _
                        " AND CUSTID=" & rsData("CustID") & _
                        " AND SERVICETYPE='P'" & _
                        " AND PRDATE IS NULL" & _
                        " And FaciCode In" & _
                        " (Select CodeNo From " & GetOwner & "CD022 Where CPImportMode In (0,1))") = 1

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
                DoEvents
            End If
        Next
        lblProcess.Caption = "[ " & strDoCaption & " ] 中 , 請稍候 !!"
        DoEvents
End Sub

Private Sub csmDelIP_ButtonClick()
  On Error GoTo ChkErr
    SetMsQry csmDelIP, "SELECT IPADDRESS,CPEMAC FROM " & GetOwner & "SO004C" & _
                                        " WHERE CUSTID=" & rsData!CustID & _
                                        " AND SEQNO='" & rsData!SeqNo & "'" & _
                                        " AND STOPDATE IS NULL", , , "IP Address,CPE MAC", "2266,2266"
  Exit Sub
ChkErr:
    ErrSub Name, "csmDelIP_ButtonClick"
End Sub

Private Sub Form_Activate()
  On Error GoTo ChkErr
    If blnShowFaci Then
        blnSendKey = False
        sstHead.Tab = 1
        cmdQuery.Value = cmdQuery.Enabled
    End If
    If blnSendKey Then
        blnSendKey = False
        If cmdSet.Enabled Then cmdSet.Value = True
    End If
    blnSendKey = False
    blnShowFaci = False
  Exit Sub
ChkErr:
    ErrSub Name, "Form_Activate"
End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr
        
        Set rsLog = New ADODB.Recordset

        Set objCmdCN = CreateObject("ADODB.Connection")
        With objCmdCN
            .CursorLocation = 3
            If LinkOdie Then
                .Open "Provider=MSDAORA.1;Password=1234;User ID=odie;Data Source=DogHouse;Persist Security Info=True"
            Else
                .Open gcnGi.ConnectionString
            End If
        End With
        
        GetSysPara
        
'        Call OpenConnection
        Call subGil
        Call subGrd
        Call DefaultValue
        Call OpenData
        Call DoProcessMode
        Fetch_Log_Data , True
        '#6542 如果Option無法使用，則要關掉設定按鈕 By Kin 2013/07/11
        Select Case Me.sstData.Tab
            Dim i As Integer
            Case 0
                For i = optBasic.LBound To optBasic.UBound
                    If (optBasic(i).Value) Then
                        If (Not optBasic(i).Enabled) Then
                            cmdSet.Enabled = False
                        End If
                        Exit For
                    End If
                Next
            Case 1
                For i = optCMChange.LBound To optCMChange.UBound
                    If (optCMChange(i).Value) Then
                        If (Not optCMChange(i).Enabled) Then
                            cmdSet.Enabled = False
                        End If
                        Exit For
                    End If
                Next
        End Select
        
        
        int_EMTA_IP_Type = GetSystemParaItem("EMTAIPTYPE", gCompCode, gServiceType, "SO041", , True)
        If strNewBaudRate <> Empty Then
            gilBaudRate.SetCodeNo strNewBaudRate
            gilBaudRate.Query_Description
        End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

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

Public Function GetIdx(cmd As String) As Integer
  On Error GoTo ChkErr
    GetIdx = Switch(cmd = "C1", 0, cmd = "E1", 0, _
                                cmd = "C2", 1, cmd = "E2", 1, _
                                cmd = "E3", 4, cmd = "A1_P", 5, _
                                cmd = "E4", 6, cmd = "E5", 7, _
                                cmd = "A1_D", 2, cmd = "A1_C", 3, _
                                cmd = "A1_BU", 0, cmd = "A1_PIP", 1, _
                                cmd = "A1_MIP", 2, cmd = "Q1", 0, _
                                cmd = "Q2", 1, cmd = "A7", 8, _
                                cmd = "A8", 8, cmd = "A9", 8, _
                                cmd = "A10", 8, cmd = "A6", 7, _
                                cmd = "A5", 6, cmd = "A4", 5, _
                                cmd = "A3", 4, cmd = "A2", 3, _
                                cmd = "C6", 5, cmd = "E6", 5, _
                                cmd = "RP", 8, cmd = "TRIAL", 9, _
                                cmd = "STOPTRIAL", 10, _
                                cmd = "ACE", 5)
  Exit Function
ChkErr:
    ErrSub Name, "GetIdx"
End Function

Private Sub DoProcessMode()
    On Error GoTo ChkErr
    Dim ctl As Control
    Dim intIndex As Integer
    'Dim blnFlag As Boolean
    'Dim intTab As Integer
    'Dim intLoop As Integer
    
    'If intProcessType <= 0 Then Exit Sub
    If strPrcType = "" Then Exit Sub
    
    intIndex = GetIdx(strPrcType)
    
    If strDefaultRowId <> "" Then RsFind rsData, "RowId", strDefaultRowId
    
    For Each ctl In Me
        If Not TypeOf ctl Is GiGridR Then
            ctl.Enabled = False
        End If
    Next
    
    pic1.Enabled = True
    pbr1.Enabled = True
    lblProcess.Enabled = True
        
    gilCompCode.Locked = True
    sstData.Enabled = True
    sstHead.Enabled = True
    sstData.TabEnabled(0) = False
    sstData.TabEnabled(1) = False
    sstData.TabEnabled(2) = False
            
    Select Case strPrcType
        Case "E1", "C1", "E2", "C2", "A1_D", "A1_C", "E3", "A1_P", "E4", "E5"
            fraData(0).Enabled = True
            sstData.TabEnabled(0) = True
            sstData.Tab = 0
            optBasic(intIndex).Enabled = True
            optBasic(intIndex).Value = True
        Case "A1_BU", "TRIAL", "STOPTRIAL", "ACE", "A1_PIP", "A1_MIP", "A2", "A3", "E6", "A4", "C6", "A5", "A6", "A7", "A8", "A9", "A10", "RP"
            fraData(1).Enabled = True
            sstData.TabEnabled(1) = True
            sstData.Tab = 1
            optCMChange(intIndex).Enabled = True
            optCMChange(intIndex).Value = True
            'If strPrcType = "A3" Then csmDelIP.Enabled = True
            If strPrcType = "ACE" Then
                txtHFCNode.Enabled = True
                txtIPAddress.Enabled = True
                cmdFindIP.Enabled = True
                chkPR.Enabled = True
                lblNode.Enabled = True
                lblIPAddress.Enabled = True
            End If
            If strPrcType = "A5" Then
                cboIP.Enabled = True
                txtCPEIP.Enabled = True
                lblOldIP.Enabled = True
                lblCPEIP.Enabled = True
            End If
            If strPrcType = "A6" Then
                cboCPE.Enabled = True
                txtCPE.Enabled = True
                lblOldCPE.Enabled = True
                lblCPE.Enabled = True
            End If
        Case "Q1", "Q2"
            fraData(2).Enabled = True
            sstData.TabEnabled(2) = True
            sstData.Tab = 2
            optQueryInfo(intIndex).Enabled = True
            optQueryInfo(intIndex).Value = True
        Case Else
    End Select
        
'        For intLoop = 0 To sstData.Tabs - 1
'            If intLoop = Left(intProcessType, 1) Then
'                sstData.TabEnabled(intLoop) = True
'                sstData.Tab = intLoop
'                fraData(intLoop) = Enabled
'                intIndex = Right(intProcessType, 1)
'                Select Case intIndex
'                    Case 1
'                        optBasic(intIndex).Enabled = True
'                        optBasic(intIndex).Value = True
'                    Case 2
'                        optCMChange(intIndex).Enabled = True
'                        optCMChange(intIndex).Value = True
'                    Case 3
'                        optQueryInfo(intIndex).Enabled = True
'                        optQueryInfo(intIndex).Value = True
'                End Select
'                Exit For
'            End If
'        Next
        
    sstHead.Tab = 0
    ggrData.Enabled = False
    cmdSet.Enabled = True
    cmdCancel.Enabled = True
    cmdQuery.Enabled = True
    cmdQuery2.Enabled = True
    cmdQry.Enabled = True
    cmdResend.Enabled = True
    gdtResvTime.Enabled = gdtResvTime.Text = ""
    
    Exit Sub
ChkErr:
    ErrSub Me.Name, "DoProcessMode"
End Sub

Private Sub subGil()
    On Error GoTo ChkErr
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 20, GetCompCodeFilter(gilCompCode, gilCompCode.GetCodeNo)
'        SetgiList gilDevStatus, "CodeNo", "Description", "CD081", , , , , 3, 20, GetCompCodeFilter(gilCompCode, gilCompCode.GetCodeNo), True
        SetgiList gilBaudRate, "CodeNo", "Description", "CD064", , , , , 3, 20, GetCompCodeFilter(gilCompCode, gilCompCode.GetCodeNo), True
        SetgiList gilNewRate, "CodeNo", "Description", "CD064", , , , , 3, 20, GetCompCodeFilter(gilCompCode, gilCompCode.GetCodeNo), True
        SetgiList gilFaciCode, "CodeNo", "Description", "CD022", , , , , 3, 12, " WHERE REFNO IN (2,5)", True
        gilFaciCode.SelectRefNo
        ' GaryGi(15)
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub subGrd()
    On Error GoTo ChkErr
    Dim mFlds As New prjGiGridR.GiGridFlds
    Dim mFlds2 As New prjGiGridR.GiGridFlds
    Dim mFlds3 As New prjGiGridR.GiGridFlds
      '定義GirdR各個欄位的內容
      'STB序號,ICC序號,裝設位置,狀態
        
        With mFlds
            .Add "PRDate", , , , , "設備狀態  "
            .Add "PRDate", , , , , "派工狀態  "
            .Add "PRDate", , , , , "開通狀態  "
            .Add "FaciSNo", , , , , "CM 序號               ", vbLeftJustify
            .Add "FaciCode", , , , , "代碼  ", vbRightJustify
            .Add "FaciName", , , , , "項目名稱  ", vbLeftJustify
            .Add "ModelName", , , , , "型號名稱   ", vbLeftJustify
            .Add "CMBaudNo", , , , , "代碼  ", vbRightJustify
            .Add "CMBaudRate", , , , , "CM速率  ", vbLeftJustify
            .Add "DynIPCount", , , , , "動IP數", vbRightJustify
            .Add "FixIPCount", , , , , "固IP數", vbRightJustify
            .Add "InstDate", giControlTypeDate, , , , "裝機日期  ", vbLeftJustify
            .Add "CMOpenDate", giControlTypeTime, , , , "CM 開機日    ", vbLeftJustify
            .Add "CMCloseDate", giControlTypeTime, , , , "CM 關機日    ", vbLeftJustify
            .Add "DialAccount", , , , , "撥接帳號   ", vbLeftJustify
            .Add "EMail", , , , , "電子信箱             ", vbLeftJustify
            .Add "IPAddress", , , , , "CP IP 位址           ", vbLeftJustify
        End With
        ggrData.AllFields = mFlds
        ggrData.SetHead
        

        '頻道編號,頻道名稱,收費起始日期,截止日期
        With mFlds2
            '.Add "CustId", , , , False, "客戶編號", vbLeftJustify
            .Add "Citemname", , , , False, "收費項目            ", vbLeftJustify
            .Add "ShouldAmt", , , , False, "出單金額  ", vbRightJustify
            '.Add "ShouldDate", , , , False, "下收日", vbRightJustify
            .Add "RealPeriod", , , , False, "期 數  ", vbLeftJustify
            .Add "RealAmt", , , , False, "實收金額   ", vbRightJustify
            .Add "RealStartDate", , , , False, "期限 起始日  ", vbLeftJustify
            .Add "RealStopDate", , , , False, "期限 截止日  ", vbLeftJustify
            .Add "CancelFlag", , , , False, "作 廢   ", vbLeftJustify
            .Add "Note", , , , False, "備 註   " & Space(42), vbLeftJustify
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
'            .Add "HandleNote", , , , , "處理說明", vbLeftJustify
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
        
        If gilCompCode.GetCodeNo <> Empty Then gilCompCode.Enabled = False
        
        '如有傳CustId ginCustId 不予許Key
        If lngCustId > 0 Then
            ginCustId.Text = lngCustId
            Call ginCustId_Validate(False)
            ginCustId.Enabled = False
        End If
        
        'sstHead.Tab = 1
        If strCaption <> "" Then Me.Caption = Me.Caption & "-" & strCaption
        blnPrcOK = False
        
        If gLogInID <> "" Then gLogInID = gLogInID & "."
        
        'If strResvTime <> "" Then gdtResvTime.Text = strResvTime
        If Len(strResvTime & "") > 0 Then
            inResvTime = strResvTime
        End If
        gdtResvTime.Text = strResvTime
        gdtResvTime.Enabled = gdtResvTime.Text = ""
        
        If strBaudRate <> "" Then
            ' strBaudRate
            If UCase(strPrcType) = "TRIAL" Then
                gilNewRate.SetCodeNo strBaudRate
                gilNewRate.Query_Description
            Else
                gilBaudRate.SetCodeNo strBaudRate
                gilBaudRate.Query_Description
            End If
            'gilBaudRate.SetDescription strBaudRate
            'gilBaudRate.Query_CodeNo
        End If
        
        If strFC <> "" Then
            gilFaciCode.SetCodeNo strFC
            If strFN <> "" Then
                gilFaciCode.SetDescription strFN
            Else
                gilFaciCode.Query_Description
            End If
        End If
        
        If strProbStopDate <> "" Then gdaProbationStopDate.Text = strProbStopDate
        
        If strFCsno <> "" Then txtFaciSNo = strFCsno
        
        Select Case strPrcType
                Case "A1_PIP"
                        txtAddDynaIP = strDynaIPcnt
                Case "A1_MIP"
                        txtDelDynaIP = strDynaIPcnt
                Case "A2"
                        txtFixIPcnt1 = strFixIPcnt
                        txtAddFixIP = strIP
                        txtAddCPE = strCPE
                Case "A3"
                        If strIP <> "" And strCPE <> "" Then
                            SetMsQry csmDelIP, "SELECT IPADDRESS,CPEMAC FROM " & GetOwner & "SO004C" & _
                                                                " WHERE CUSTID=" & rsParent!CustID & _
                                                                " AND SEQNO='" & rsParent!SeqNo & "'" & _
                                                                " AND STOPDATE IS NULL", , , "IP Address,CPE MAC", "2266,2266"
                            If InStr(1, strIP, "#", vbTextCompare) Then
                                strIP = "'" & Replace(strIP, "#", "','", 1, , vbTextCompare) & "'"
                            Else
                                strIP = "'" & strIP & "'"
                            End If
                            'strIP = "'111.111.111.111','123.240.189.250'"
                            csmDelIP.SetQueryCode strIP
                        End If
                Case "A4", "C6", "E6"
                        If strNode <> "" Then
                            txtHFCNode = strNode
                        Else
                            chkPR_Click
                            If chkPR.Value Then
                                txtHFCNode = GetHFC(GetNewAddr)
                            Else
                                txtHFCNode = GetHFC(lngInstAddrNo)
                            End If
                        End If
                        If strNewIP <> "" Then
                            txtIPAddress = strNewIP
                        End If
                Case "A5"
                            cboIP.Text = strOldIP
                            txtCPEIP.Text = strIP
                Case "A6"
                            cboCPE.Text = strOldCPE
                            txtCPE.Text = strCPE
        End Select
        
'        If strPrcType = "A1_PIP" Then
'            'txtAddDynaIP = strIPnum
'            txtAddDynaIP = strDynaIPcnt
'        End If
'
'        If strPrcType = "A1_MIP" Then
'            'txtDelDynaIP = strIPnum
'            txtDelDynaIP = strDynaIPcnt
'        End If
'
'        If strPrcType = "A2" Then
'            'txtFixIPcnt1 = strIPnum
'            txtFixIPcnt1 = strFixIPcnt
'            txtAddFixIP = strIP
'            txtAddCPE = strCPE
'        End If
'
'        If strPrcType = "A3" And strIP <> "" And strCPE <> "" Then
'            SetMsQry csmDelIP, "SELECT IPADDRESS,CPEMAC FROM " & GetOwner & "SO004C" & _
'                                                " WHERE CUSTID=" & rsParent!CustID & _
'                                                " AND SEQNO='" & rsParent!SeqNo & "'" & _
'                                                " AND STOPDATE IS NULL", , , "IP Address,CPE MAC", "2266,2266"
'            If InStr(1, strIP, "#", vbTextCompare) Then
'                strIP = "'" & Replace(strIP, "#", "','", 1, , vbTextCompare) & "'"
'            Else
'                strIP = "'" & strIP & "'"
'            End If
'            'strIP = "'111.111.111.111','123.240.189.250'"
'            csmDelIP.SetQueryCode strIP
'
'        End If
'
'        If strPrcType = "A5" Then
'            cboIP.Text = strOldIP
'            txtCPEIP.Text = strIP
'        End If
'
'        If strPrcType = "A6" Then
'            cboCPE.Text = strOldCPE
'            txtCPE.Text = strCPE
'        End If
        
End Sub

Private Sub OpenData(Optional strWhere As String)
    On Error GoTo ChkErr
    Dim strSQL As String
        Set rsData = New ADODB.Recordset
        
        If rsParent Is Nothing Then Set rsParent = New ADODB.Recordset
        If rsParent.State = adStateOpen Then
            Set rsData = rsParent.Clone
            If blnShowFaci Then
                rsData.Filter = " PRDate = Null"
            End If
            strDefaultRowId = rsParent("RowId")
        Else
            If ginCustId.Text <> "" Then
                strSQL = " A.CustId = " & ginCustId.Value & _
                                " And A.CompCode = " & gilCompCode.GetCodeNo & _
                                " And A.FaciSNo is not null" & _
                                " And FaciCode In (Select CodeNo From " & GetOwner & "CD022 Where RefNo in (2,5))"
'                strSQL = " A.CustId = " & ginCustId.value & " And A.CompCode = " & gilCompCode.GetCodeNo & " And A.FaciSNo is not null And FaciCode In (Select CodeNo From " & GetOwner & "CD022 Where RefNo in (2,5)) And A.PRDate is null "
                If strWhere <> "" Then strSQL = " And " & strWhere
            Else
                strSQL = " A.SeqNo = ''"
            End If
            'If Not GetRS(rsData, "Select A.RowId, A.* From " & GetOwner & "So004 A Where " & strSQL & " Order By A.InstDate", , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
            If Not GetRS(rsData, "Select A.RowId, A.* From " & GetOwner & "SO004 A" & _
                                                " Where " & strSQL & _
                                                " Order By A.PRdate desc,A.InstDate", , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        End If
        Call OpenData2
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
        If Not GetRS(rsChild, "Select RowId,A.* From " & GetOwner & "SO033 A Where " & strWhere) Then Exit Sub
        Call EnableMode(True)
        Set ggrChild.Recordset = rsChild
        ggrChild.Refresh
        ' 增加NODE 輸入方塊，若選取設備屬於CM 設備參考號2的設備時，IP 的部份需要DISABLE
        If Val(GetRsValue("SELECT REFNO FROM " & GetOwner & "CD022" & _
                                    " WHERE CODENO=" & rsData("FaciCode") & "") & "") = 2 Then
            txtIPAddress.Enabled = False
            txtIPAddress = ""
        Else
            txtIPAddress.Enabled = True
        End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData2"
End Sub

Private Sub EnableMode(blnFlag As Boolean)
    On Error Resume Next
        
'    裝機 SO11721
'    拆機 SO11722
'    軟拆 SO11723
'    軟復 SO11724
'    CP 裝機 SO11725
'    CP拆機 SO11726
'    CP裝分線 SO11727
'    CP拆分線 SO11728
        
    If strPrcType <> "" Then Exit Sub
        
    optBasic(0).Enabled = GetPermission("SO11721") ' 裝機
    optBasic(1).Enabled = GetPermission("SO11722") ' 拆機
    optBasic(2).Enabled = GetPermission("SO11723") ' 軟拆
    optBasic(3).Enabled = GetPermission("SO11724") ' 軟復
    optBasic(4).Enabled = GetPermission("SO11725") ' CP 裝機
    optBasic(5).Enabled = GetPermission("SO11726") ' CP 拆機
    optBasic(6).Enabled = GetPermission("SO11727") ' CP 裝分線
    optBasic(7).Enabled = GetPermission("SO11728") ' CP 拆分線

'    速率升降級 SO11729
'    申請動態IP SO1172A
'    取消動態IP SO1172B
'    申請固定IP SO1172C
'    取消固定IP SO1172D
'    EMTA IP異動 SO1172E
'    CPE IP 異動 SO1172F
'    CPE MAC 異動 SO1172G
'    CM / EMTA 更換 SO1172H
        
    optCMChange(0).Enabled = GetPermission("SO11729") ' 速率升降級
    'optCMChange(0).Enabled = False
    gilBaudRate.Enabled = optCMChange(0).Enabled        '#6542 速率是否可點選依照optCMChange(0) By Kin 2013/07/11
    optCMChange(1).Enabled = GetPermission("SO1172A") ' 申請動態 IP
    optCMChange(2).Enabled = GetPermission("SO1172B") ' 取消動態 IP
    optCMChange(3).Enabled = GetPermission("SO1172C") ' 申請固定 IP
    optCMChange(4).Enabled = GetPermission("SO1172D") ' 取消固定 IP
    optCMChange(5).Enabled = GetPermission("SO1172E") ' EMTA IP異動
    optCMChange(6).Enabled = GetPermission("SO1172F") ' CPE IP 異動
    optCMChange(7).Enabled = GetPermission("SO1172G") ' CPE MAC 異動
    optCMChange(8).Enabled = GetPermission("SO1172H") ' CM / EMTA 更換
    ' 要加權限
    optCMChange(9).Enabled = GetPermission("SO1172K") ' 試用速率
    optCMChange(10).Enabled = GetPermission("SO1172L") ' 終止試用


'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172K', null, '試用速率', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172L', null, '終止試用', '', 'SO1172', null);
              
              
'    重設 CM Reset SO1172I
'    查詢 CM 狀態 SO1172J
    
    optQueryInfo(0).Enabled = GetPermission("SO1172I") ' 重設 CM Reset
    optQueryInfo(1).Enabled = GetPermission("SO1172J") ' 查詢 CM 狀態


'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172', null, 'CM設定控制台', '', 'SO1100', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11721', null, '裝機', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11722', null, '拆機', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11723', null, '軟拆', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11724', null, '軟復', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11725', null, 'CP裝機', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11726', null, 'CP拆機', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11727', null, 'CP裝分線', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11728', null, 'CP拆分線', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11729', null, '速率升降級', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172A', null, '申請動態IP', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172B', null, '取消動態IP', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172C', null, '申請固定IP', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172D', null, '取消固定IP', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172E', null, 'EMTA IP異動', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172F', null, 'CPE IP異動', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172G', null, 'CPE MAC異動', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172H', null, 'CM/EMTA更換', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172I', null, '重設CM Reset', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172J', null, '查詢CM 狀態', '', 'SO1172', null);


'   底下是 SPM 舊專案的權限
'    optBasic(0).Enabled = GetPermission("SO11721") ' CM入庫
'    optBasic(1).Enabled = GetPermission("SO11722") ' CM出庫
'    optBasic(2).Enabled = GetPermission("SO11723") ' 服務設定
'    optBasic(3).Enabled = GetPermission("SO11724") ' 服務刪除
'    optBasic(4).Enabled = GetPermission("SO11725") ' CP服務增加
'    optBasic(5).Enabled = GetPermission("SO11726") ' CP服務刪除
'    optBasic(6).Enabled = GetPermission("SO11727") ' CM重設
'
'    optCMChange(0).Enabled = GetPermission("SO11728") ' 軟拆
'    optCMChange(1).Enabled = GetPermission("SO11729") ' 軟復
'    optCMChange(2).Enabled = GetPermission("SO1172A") ' 變更服務申請IP
'    optCMChange(3).Enabled = GetPermission("SO1172B") ' 速率升降級
'    optCMChange(4).Enabled = GetPermission("SO1172C") ' 變更CPE MAC
'    optCMChange(5).Enabled = GetPermission("SO1172D") ' 取消IP
'    optCMChange(6).Enabled = GetPermission("SO1172E") ' CP Only拆機
'    optCMChange(7).Enabled = GetPermission("SO1172F") ' 網路編號 / NODE or Voice IP
'
'    optQueryInfo(0).Enabled = GetPermission("SO1172G") ' 查詢 CM 狀況
'    optQueryInfo(1).Enabled = GetPermission("SO1172H") ' 查詢 CM 參數


'   底下是 SPM 舊專案的權限
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172', null, 'SPM設定控制台', '', 'SO1100', null);
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11721', null, 'CM入庫', '', 'SO1172', null);
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11722', null, 'CM出庫', '', 'SO1172', null);
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11723', null, '服務設定', '', 'SO1172', null);
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11724', null, '服務刪除', '', 'SO1172', null);
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11725', null, 'CP服務增加', '', 'SO1172', null);
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11726', null, 'CP服務刪除', '', 'SO1172', null);
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11727', null, 'CM重設', '', 'SO1172', null);
'
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11728', null, '軟拆', '', 'SO1172', null);
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO11729', null, '軟復', '', 'SO1172', null);
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172A', null, '變更服務申請IP', '', 'SO1172', null);
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172B', null, '速率升降級', '', 'SO1172', null);
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172C', null, '變更CPE MAC', '', 'SO1172', null);
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172D', null, '取消IP', '', 'SO1172', null);
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172E', null, 'CP Only拆機', '', 'SO1172', null);
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172F', null, '網路編號 / NODE or Voice IP', '', 'SO1172', null);
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172G', null, '查詢 CM 狀況', '', 'SO1172', null);
'   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
'              VALUES ( 'SO1172H', null, '查詢 CM 參數', '', 'SO1172', null);
              
End Sub

Private Sub ggrQuery_DblClick()
  On Error GoTo ChkErr
    ShowVtcFld ggrQuery
  Exit Sub
ChkErr:
    ErrSub Name, "ggrQuery_DblClick"
End Sub

Private Sub ggrQueryInfo_DblClick()
  On Error GoTo ChkErr
    ShowVtcFld ggrQueryInfo
  Exit Sub
ChkErr:
    ErrSub Name, "ggrQueryInfo_DblClick"
End Sub

Private Sub ggrQueryInfo_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If UCase(giFld.FieldName) = "MSGNAME" Then Value = CmdDisp(Value)
  Exit Sub
ChkErr:
    ErrSub Name, "ggrQueryInfo_ShowCellData"
End Sub


'裝機 (CM)   C1 '裝機(EMTA ) E1
'拆機（CM裝機退件)   C2 '拆機（EMTA裝機退件)     E2
'CP Only拆機 A1_P '軟拆 A1_D
'軟復 A1_C 'CP裝機 E3
'速率升降級 A1_BU '申請動態IP A1_PIP
'取消動態IP A1_MIP '申請固定IP A2
'取消固定IP A3 'emta 加裝CP門號 E4
'emta 拆CP 門號  E5 'EMTA IP Priv IP設定 A4
'CM IP  CPE IP 設定  A5 '變更CPE MAC A6
'CM 拆舊/換新    A7 'EMTA 拆舊/換新  A8
'拆舊CM /換新EMTA    A9 '拆舊EMTA /換新 CM   A10
'CM reset    Q1 '查詢CM狀況 Q2
'CM 移機 C6 'EMTA 移機   E6

' F - Failure , N - New ( Not Processed) , S - Success

' 狀態失敗的可按右鍵作指令重送，產生一筆新的指令紀錄且標記舊的已重送(失敗且標記已重送的不可在重送)。

' 3.  設定紀錄，針對SPM(TBC)指令未成功的，按右鍵可以查詢指令執行結果，若SPM 已成功，需要更新指令執行狀態。
' 指令狀態為 N 的可按右鍵作指令查詢指令狀態，若回應成功需按指令功能回填相關資料

'Private Sub ggrQuery_MoveRecord()
'  On Error GoTo ChkErr
'    With ggrQuery
'        If Not .Recordset Is Nothing Then
'            If .Recordset.EOF Then
'                cmdResend.Enabled = False
'                cmdQry.Enabled = False
'            Else
'                Select Case UCase(.Recordset("StatusCode"))
'                            Case "F"
'                                        If Len(.Recordset("CmdResend") & "") > 0 Then
'                                            cmdResend.Enabled = False
'                                            cmdQry.Enabled = False
'                                        Else
'                                            cmdResend.Enabled = True
'                                            cmdQry.Enabled = False
'                                        End If
'                            Case "N"
'                                        cmdResend.Enabled = False
'                                        cmdQry.Enabled = True
'                            Case Else
'                                        cmdResend.Enabled = False
'                                        cmdQry.Enabled = False
'
'                End Select
'            End If
'        End If
'    End With
'  Exit Sub
'ChkErr:
'    ErrSub Name, "ggrQuery_MoveRecord"
'End Sub

Private Sub rsLog_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  On Error GoTo ChkErr
    With rsLog
            If Not rsLog Is Nothing Then
                If .EOF Then
                    cmdResend.Enabled = False
                    cmdQry.Enabled = False
                Else
                    Select Case UCase(.Fields("CommandStatus"))
                                Case "E", "W", "C", "T", "GT"
                                            If Len(.Fields("CmdResend") & "") > 0 Then
                                                cmdResend.Enabled = False
'                                                cmdQry.Enabled = False
                                            Else
                                                cmdResend.Enabled = True
                                                cmdQry.Enabled = False
                                            End If
'                                Case "W", "C", "T"
'                                            cmdResend.Enabled = False
'                                            cmdQry.Enabled = True
                                Case Else
                                            cmdResend.Enabled = False
'                                            cmdQry.Enabled = False
                                
                    End Select
                End If
            End If
    End With
  Exit Sub
ChkErr:
    ErrSub Name, "rsLog_MoveComplete"
End Sub

'CM Console 的方案類別是抓 CD078D 的 BPcode , BPname
'速率是抓 CMBaudRate
'動態IP 或靜態 IP 數則是抓取 DynIPCount / FixIPCount 或 CMCPDynIP / CMCPFixIP

Private Sub ginCustId_Validate(Cancel As Boolean)
    On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
    'If strPrcType <> "" Then Exit Sub
        If ginCustId.Text = "" Then
            txtCustName = ""
            Exit Sub
        Else
            If Not GetRS(rs, "Select CustName,InstAddrNo From " & GetOwner & "So001 Where CustId = " & ginCustId.Value & " And CompCode = " & gilCompCode.GetCodeNo, gcnGi, , , , , , , True) Then Exit Sub
            If rs.EOF Then
                txtCustName = ""
                lngInstAddrNo = 0
                strNodeNo = 0
                strCircuitNO = 0
            Else
                txtCustName = rs("CustName") & ""
                lngInstAddrNo = Val(rs("InstAddrNo") & "")
                strNodeNo = 0
                strCircuitNO = 0
            End If
        End If

        If Not GetRS(rs, "SELECT NODENO,CIRCUITNO FROM " & GetOwner & "SO014 WHERE ADDRNO = " & lngInstAddrNo, gcnGi, , , , , , , True) Then Exit Sub
        If rs.EOF Then
            strNodeNo = 0
            strCircuitNO = 0
        Else
            strNodeNo = rs("NODENO") & ""
            strCircuitNO = rs("CIRCUITNO") & ""
        End If

        Call CloseRecordset(rs)
        
        If txtCustName = "" Then
            If ginCustId.Text <> "" Then MsgBox "查無此客戶", vbCritical, gimsgPrompt
            ggrData.Enabled = False
            ggrChild.Enabled = False
            If rsData.RecordCount > 0 Then Call OpenData("1=0")
            ginCustId.Text = ""
            Cancel = True
        Else
            If strPrcType <> "" Then Exit Sub
            Call OpenData
            ggrData.Enabled = True
            ggrChild.Enabled = True
        End If
        
    Exit Sub
ChkErr:
    ErrSub Me.Name, "ginCustId_Validate"
End Sub

Private Sub txt_KP(KeyAscii As Integer)
  On Error Resume Next
    Dim strChr As String
    strChr = UCase(Chr(KeyAscii))
    Select Case True
            Case strChr = ":", strChr = "-", strChr = "#", KeyAscii = 8
            Case strChr = "A", strChr = "B", strChr = "C", strChr = "D", strChr = "E", strChr = "F"
            Case strChr = " ", Not IsNumeric(strChr)
                    KeyAscii = 0
    End Select
End Sub

Private Function IsDataOK() As Boolean
    On Error GoTo ChkErr
        
        If gilCompCode.GetCodeNo = Empty Then
            MsgBox "[公司別] 為必要欄位 !", vbCritical, gimsgPrompt
            On Error Resume Next
            gilCompCode.SetFocus
            Exit Function
        End If
        If ginCustId.Text = Empty Then
            MsgBox "[客戶編號] 為必要欄位 !", vbCritical, gimsgPrompt
            On Error Resume Next
            ginCustId.SetFocus
            Exit Function
        End If
        
        If Alfa Then
            IsDataOK = True
            Exit Function
        End If
        
        With rsData
            Select Case sstData.Tab
                Case 0
                    Select Case True
                        Case optBasic(0) ' 1. 裝機
                            ' 設備拆機日期是空白與開機日期 是空白 才可以裝機指令 此功能才可以ENABLE
                            If strPrcType = "" Then
                                If !CMOpenDate & "" <> "" Or !PRdate & "" <> "" Then
                                    MsgBox "設備 [開機日期] 與 [拆機日期] 無值才可裝機 !", vbInformation, "訊息"
                                    Exit Function
                                End If
                            End If
                            
                            Dim blnChkAddrNo As Boolean
                            
                            GetHFCNode !CustID, !CompCode, !ServiceType & "", blnChkAddrNo
                            
                            If blnChkAddrNo Then
                                MsgBox "該客戶移機中，未指定新址，請指定新址後再送開機指令！", vbInformation, "訊息"
                                Exit Function
                            End If
                            
                        Case optBasic(1) ' 2. 拆機 ( 含裝機退件 )
                            ' 設備裝機日期非空白 必須有開機日期才可以作拆機指令 此功能才可以ENABLE
'                            If !InstDate & "" = "" Or !CMOpenDate & "" = "" Then
                            If strPrcType = "" Then
                                If !CMOpenDate & "" = "" Then
                                    MsgBox "設備 [裝機日期] 與 [開機日期] 有值才可拆機 !", vbInformation, "訊息"
                                    Exit Function
                                End If
                            End If
                        Case optBasic(2) ' 3. 軟拆
                            ' 有開機日 且開機日>(關機日OR 無關機日)則只可選軟關
                            ' 正常使用中、安裝中不分EMTA 與CM 該筆資料必須（有開機日期 與 關機日期是空的） 或（開機日期>關機日期）此功能才可以ENABLE
                            If strPrcType = "" Then
                                If !CMOpenDate & "" = "" Or !CMCloseDate & "" <> "" Then
                                    If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                                        If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
                                            MsgBox "設備 [正常使用中] 才可軟拆 !", vbInformation, "訊息"
                                            Exit Function
                                        End If
                                    Else
                                        MsgBox "設備 [正常使用中] 才可軟拆 !", vbInformation, "訊息"
                                        Exit Function
                                    End If
                                End If
                            End If
                        Case optBasic(3) ' 4. 軟復
                            ' 軟關與軟復兩各選項是互斥，設備若(關機日OR 無關機日) > 開機日 則只可以選軟復
                            ' 正常使用中、安裝中不分EMTA 與CM 該筆資料必須（有關機日期） 或（開機日期<關機日期）此功能才可以ENABLE
                            If strPrcType = "" Then
                                If !CMOpenDate & "" <> "" Or !CMCloseDate & "" = "" Then
                                    If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                                        If CDT(!CMCloseDate) < CDT(!CMOpenDate) Then
                                            MsgBox "設備 [正常使用中] 才可軟復 !", vbInformation, "訊息"
                                            Exit Function
                                        End If
                                    Else
                                        MsgBox "設備 [正常使用中] 才可軟復 !", vbInformation, "訊息"
                                        Exit Function
                                    End If
                                End If
                            End If
                        Case optBasic(4) ' 5. CP 裝機
                            If strPrcType = "" Then
                                ' 設備正常使用中、安裝中EMTA 開機日期大於關機日期  且關聯到CP 服務需有CP 對應門號，且該門號為安裝中或正常此功能才可以ENABLE
                                If Not Is_EMTA_Dev Then
                                    MsgBox "此設備非 EMTA , 無法進行 CP 裝機 !", vbInformation, "訊息"
                                    Exit Function
                                End If
                                ' #4302,2009/01/03 By Hammer,(RA)_SO1171A_20081223_調整CLI控制台CP門號判斷.doc,增加 IsCP 判斷式
                                '  1.調整 CLI控製台:警告訊息"無[安裝中]或[正常]的[CP門號]，無法進行CP拆機。"判斷條件，使其僅針對CP門號做判斷。
                                '  2.調整CLI控製台:5.CP ONLY裝機、6. CP ONLY拆機、7.CP裝分線、8.CP拆分線過濾條件，使其僅針對CP門號送指令。
                                If IsCP(rsData) Then
                                    If Not IsCMCP Then
                                        MsgBox "無 [ 安裝中 ] 或 [ 正常 ] 的 [ CP 門號 ] , 無法進行 CP 裝機 !", vbInformation, "訊息"
                                        Exit Function
                                    End If
                                End If
    '                            IsCMCP = (GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                                                    " WHERE REFACISNO='" & rsData("FaciSno") & "'" & _
                                                                    " AND COMPCODE=" & rsData("CompCode") & _
                                                                    " AND CUSTID=" & rsData("CustID") & _
                                                                    " AND SERVICETYPE='P'" & _
                                                                    " AND PRDATE IS NULL") = 1)
                                If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                                    If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
                                        MsgBox "無 [ 安裝中 ] 或 [ 正常 ] 的 [ CP 門號 ] , 無法進行 CP 裝機 !", vbInformation, "訊息"
                                        Exit Function
                                    End If
                                End If
                                ' 請加強資料搜尋時接以該客編相同的資料為主（custid = 設定的 custid），
                                ' gateway 取值時請增加group by Gateway ，
                                ' 如果取出有兩筆（含）以上時需要提出警示訊息（該emta 門號分別屬於不同 mediaGateway 無法設定），不做指令傳送
                                '===============================================================================================================
                                ' #4302,2009/01/03 By Hammer,(RA)_SO1171A_20081223_調整CLI控制台CP門號判斷.doc,
                                '     增加 FaciCode In ( Select CodeNo .. CPImportMode.. ) SubQry
                                ' 1.  可參考CD022. CPImportMode 門號匯入模式做為過濾條件(0-台固,1-亞太,2-速博)。
                                If Val(GetRsValue("SELECT COUNT(GATEWAY) FROM " & GetOwner & "SO004" & _
                                                    " WHERE REFACISNO='" & rsData("FaciSno") & "'" & _
                                                    " AND CUSTID=" & rsData("CustID") & _
                                                    " AND COMPCODE=" & rsData("CompCode") & _
                                                    " AND SERVICETYPE='P'" & _
                                                    " AND PRDATE IS NULL" & _
                                                    " And FaciCode In " & _
                                                    " (Select CodeNo From " & GetOwner & "CD022 Where Nvl(CPImportMode,0) In (0,1))" & _
                                                    " GROUP BY GATEWAY", gcnGi, "") & "") > 1 Then
                                    MsgBox "該 EMTA 門號分別屬於不同 Media Gateway 無法設定 !", vbInformation, "訊息"
                                    Exit Function
                                End If
                            End If
                        Case optBasic(5) ' 6. CP 拆機
                            If strPrcType = "" Then
                                ' 設備正常使用中、安裝中 EMTA 開機日期大於關機日期 且關聯到P 服務 CP 對應門號必須要都是拆機的 此功能才可以ENABLE
                                If Not Is_EMTA_Dev Then
                                    MsgBox "此設備非 EMTA , 無法進行 CP 拆機 !", vbInformation, "訊息"
                                    Exit Function
                                End If
                                ' #4302,2009/01/03 By Hammer,(RA)_SO1171A_20081223_調整CLI控制台CP門號判斷.doc,增加IsCP判斷式
                                If IsCP(rsData) Then
                                    If Not IsCMCP Then
                                        MsgBox "無 [ 安裝中 ] 或 [ 正常 ] 的 [ CP 門號 ] , 無法進行 CP 拆機 !", vbInformation, "訊息"
                                        Exit Function
                                    End If
                                End If
    '                            If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
    '                                If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
    '                                    MsgBox "無 [ 安裝中 ] 或 [ 正常 ] 的 [ CP 門號 ] , 無法進行 CP 裝機 !", vbInformation, "訊息"
    '                                    Exit Function
    '                                End If
    '                            End If
                            End If
                        Case optBasic(6) ' 7. CP 裝分線
                            ' 正常使用中、安裝中 EMTA 開機日期大於關機日期 且關聯到P 服務 CP 對應門號必須要有一筆（含以上）正常，
                            ' 新的一筆安裝中（無裝拆機日期） 此功能才可以ENABLE
                            ' SNO <> "" AND InstDate is Null AND PRdate Is Null
                            If strPrcType = "" Then
                                If Not Is_EMTA_Dev Then
                                    MsgBox "此設備非 EMTA , 無法進行 CP 裝分線 !", vbInformation, "訊息"
                                    Exit Function
                                End If
                                If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                                    If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
                                        MsgBox "無 [ 安裝中 ] 或 [ 正常 ] 之 EMTA 設備 , 無法進行 CP 裝分線 !", vbInformation, "訊息"
                                        Exit Function
                                    End If
                                End If
                                If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                                " WHERE REFACISNO='" & rsData("FaciSno") & "'" & _
                                                " AND COMPCODE=" & rsData("CompCode") & _
                                                " AND CUSTID=" & rsData("CustID") & _
                                                " AND SERVICETYPE='P' AND SNO IS NOT NULL " & _
                                                " AND PRDATE IS NULL AND INSTDATE IS NULL") < 1 Then
                                    MsgBox "此 EMTA 無對應 CP 門號 , 無法安裝 CP 分線 !", vbInformation, "訊息"
    '                                MsgBox "此門號無正常對應之 EMTA 設備 , 無法安裝 CP 分線 !", vbInformation, "訊息"
                                    Exit Function
                                End If
                            End If
                        Case optBasic(7) ' 8. CP 拆分線
                            ' 正常使用中、安裝中 EMTA 開機日期大於關機日期 且 關聯到P 服務 CP 對應門號必須要有一筆（含以上）正常，有一筆拆除中（有裝機日期、無拆機日期） 此功能才可以ENABLE
                            ' PRSNO <> "" and PRdate =""
                            
                            If strPrcType = "" Then
                                If Not Is_EMTA_Dev Then
                                    MsgBox "此設備非 EMTA , 無法進行 CP 拆分線 !", vbInformation, "訊息"
                                    Exit Function
                                End If
                                ' 正常使用中、安裝中 EMTA 開機日期大於關機日期 且關聯到該客戶P 服務的CP 對應門號必須要有一筆（含以上）正常，
                                ' 新的一筆安裝中（無裝拆機日期） 此功能才可以ENABLE
                                ' EMTA 序號關聯到so004.servicetype='P' and prdate is null 的取出 p 服務的 so004.Gateway 取一筆即可與cp 相同取值原則
    '                            If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
    '                                If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
    ''                                    MsgBox "無 [ 安裝中 ] 或 [ 正常 ] 之 EMTA 設備 , 無法進行 CP 拆分線 !", vbInformation, "訊息"
    '                                    MsgBox "設備已關機 , 無法進行 CP 拆分線 !", vbInformation, "訊息"
    '                                    Exit Function
    '                                End If
    '                            End If
                                ' PRSNO  有值, InstDate 有 , PRdate 無
                                '  AND SNO IS NOT NULL
                                If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                                " WHERE REFACISNO='" & rsData("FaciSno") & "'" & _
                                                " AND COMPCODE=" & rsData("CompCode") & _
                                                " AND CUSTID=" & rsData("CustID") & _
                                                " AND SERVICETYPE='P'" & _
                                                " AND PRDATE IS NULL" & _
                                                " AND INSTDATE IS NOT NULL") <= 1 Then
                                    MsgBox "無法拆分機 !", vbInformation, "訊息"
    '                                MsgBox "無正常的 CP 門號 , 無法 CP 拆分線 !", vbInformation, "訊息"
    '                                MsgBox "此門號無正常對應之 EMTA 設備 , 無法安裝 CP 拆分線 !", vbInformation, "訊息"
                                    Exit Function
                                End If
                                ' 停拆移機單號
                                If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                                " WHERE REFACISNO='" & rsData("FaciSno") & "'" & _
                                                " AND COMPCODE=" & rsData("CompCode") & _
                                                " AND CUSTID=" & rsData("CustID") & _
                                                " AND SERVICETYPE='P' AND PRSNO IS NOT NULL" & _
                                                " AND PRDATE IS NULL AND INSTDATE IS NOT NULL") > 1 Then
                                    MsgBox "有 2 張 (以上) 拆機單 , 請下拆機命令 !", vbInformation, "訊息"
                                    Exit Function
                                End If
                            End If
                    End Select
                Case 1
                    Select Case True
                        Case optCMChange(0)
                            If strPrcType = "" Then
                                If gilBaudRate.GetCodeNo = Empty Or gilBaudRate.GetDescription & "" = rsData!CMBaudRate & "" Then
                                    MsgBox "請選擇 [ 新速率 ] 後再執行 [ 設定 ] !", vbInformation, "提示"
                                    On Error Resume Next
                                    gilBaudRate.SetFocus
                                    Exit Function
                                End If
                                ' 安裝中或正常的設備
                                If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                                    If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
                                        MsgBox "無 [ 安裝中 ] 或 [ 正常 ] 之設備 , 無法進行 [速率升降級] !", vbInformation, "訊息"
                                        Exit Function
                                    End If
                                End If
                            End If
                        Case optCMChange(1)
                            If strPrcType = "" Then
                                If Val(txtAddDynaIP) = 0 Or txtAddDynaIP = Empty Then
                                    MsgBox "[ 動態IP ] 不得為空或 0 !", vbInformation, "訊息"
                                    On Error Resume Next
                                    txtAddDynaIP.SetFocus
                                    Exit Function
                                End If
                            End If
'                                If Val(txtAddDynaIP) <= Val(rsData("DynIPCount") & "") Then
'                                    MsgBox "[ 動態IP ] 不得小於原始值 [ " & Val(rsData("DynIPCount") & "") & " ] !", vbInformation, "訊息"
'                                    On Error Resume Next
'                                    txtAddDynaIP.SetFocus
'                                    Exit Function
'                                End If
                        Case optCMChange(2)
                            If strPrcType = "" Then
                                If Val(txtDelDynaIP) = 0 Then
                                    MsgBox "[ 動態IP ] 不得為空或 0 !", vbInformation, "訊息"
                                    On Error Resume Next
                                    txtDelDynaIP.SetFocus
                                    Exit Function
                                End If
                                If Val(rsData("DynIPCount") & "") = 0 Then
                                    MsgBox "已無動態 IP 可取消 !", vbInformation, "訊息"
                                    On Error Resume Next
                                    txtDelDynaIP.SetFocus
                                    Exit Function
                                End If
                                If Val(txtDelDynaIP) > Val(rsData("DynIPCount") & "") Then
                                    MsgBox "[ 動態IP ] 不得大於原始值 [ " & Val(rsData("DynIPCount") & "") & " ] !", vbInformation, "訊息"
                                    On Error Resume Next
                                    txtDelDynaIP.SetFocus
                                    Exit Function
                                End If
                            End If
'                                If Val(txtDelDynaIP) >= Val(rsData("DynIPCount") & "") Then
'                                    MsgBox "[ 動態IP ] 必須小於原始值 [ " & Val(rsData("DynIPCount") & "") & " ] !", vbInformation, "訊息"
'                                    On Error Resume Next
'                                    txtDelDynaIP.SetFocus
'                                    Exit Function
'                                End If
'                                If Val(txtFixIP2) < Val(rsData("FixIPCount") & "") Then
'                                    MsgBox "[ 固定IP ] 不得小於原始值 [ " & Val(rsData("FixIPCount") & "") & " ] !", vbInformation, "訊息"
'                                    On Error Resume Next
'                                    txtFixIP2.SetFocus
'                                    Exit Function
'                                End If
                        Case optCMChange(3)
                            'If strPrcType = "" Then
                                If Val(txtFixIPcnt1) = 0 Then
                                    MsgMustBe "固定IP數"
                                    On Error Resume Next
                                    txtFixIPcnt1.SetFocus
                                    Exit Function
                                End If
                                If txtAddFixIP = "" Then
                                    MsgMustBe "IP"
                                    On Error Resume Next
                                    txtAddFixIP.SetFocus
                                    Exit Function
                                End If
                                If txtAddCPE = "" Then
                                    MsgMustBe "CPE MAC"
                                    On Error Resume Next
                                    txtAddCPE.SetFocus
                                    Exit Function
                                End If
                            'End If
                        Case optCMChange(4)
'                                If Val(txtFixIPcnt2) = 0 Then
                                If csmDelIP.GetQueryCode = Empty Then
                                    MsgMustBe "IP , CPE MAC"
                                    On Error Resume Next
                                    csmDelIP.SetFocus
                                    Exit Function
                                End If
'                                If txtDelFixIP = "" Then
'                                    MsgMustBe "IP"
'                                    On Error Resume Next
'                                    txtDelFixIP.SetFocus
'                                    Exit Function
'                                End If
'                                If txtDelCPE = "" Then
'                                    MsgMustBe "CPE MAC"
'                                    On Error Resume Next
'                                    txtDelCPE.SetFocus
'                                    Exit Function
'                                End If
                        Case optCMChange(6)
                                If Len(txtCPEIP.Text) = 0 Then
                                    MsgMustBe "新 CPE IP"
                                    On Error Resume Next
                                    txtCPEIP.SetFocus
                                    Exit Function
                                End If
                                If Len(cboIP.Text) = 0 Then
                                    MsgMustBe "舊 CPE IP"
                                    On Error Resume Next
                                    cboIP.SetFocus
                                    Exit Function
                                End If
                        Case optCMChange(7)
                                If Len(txtCPE) = 0 Then
                                    MsgMustBe "CPE MAC"
                                    On Error Resume Next
                                    txtCPE.SetFocus
                                    Exit Function
                                End If
                        Case optCMChange(8)
                                If strPrcType = "" Then
                                    If Len(gilFaciCode.GetCodeNo) = 0 Then
                                        MsgMustBe "品名"
                                        On Error Resume Next
                                        gilFaciCode.SetFocus
                                        Exit Function
                                    End If
                                    If Len(Trim(txtFaciSNo)) = 0 Then
                                        MsgMustBe "CM 序號"
                                        On Error Resume Next
                                        txtFaciSNo.SetFocus
                                        Exit Function
                                    Else
                                        ' 2006/06/21 輸入新序號時需要檢核該序號在SO004 中是否有重複(沒有拆機日期的)及算重複不可作設定
                                        Dim strFaciSno As String
                                        strFaciSno = txtFaciSNo
                                        strFaciSno = Replace(strFaciSno, "-", "", 1)
                                        strFaciSno = Replace(strFaciSno, ":", "", 1)
                                        If strFaciSno = rsData("FaciSno") Or strFaciSno = rsData("ReFaciSno") Then
                                            MsgBox "[CM 序號] 與舊序號相同 .. 請重新輸入新序號 !", vbInformation, "訊息"
                                            On Error Resume Next
                                            txtFaciSNo.SetFocus
                                            Exit Function
                                        Else
                                            
                                            ' " OR REFACISNO='" & strFaciSno & "' )"
                                            '#6154 依據SO042.FaciAuthority的參數來判斷欄位 By Kin 2011/11/08
                                            If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                                                    " WHERE FACISNO='" & strFaciSno & "'" & _
                                                                    " AND " & strFaciAuthorityField & " IS NULL") > 0 Then
                                                MsgBox "此序號已被使用 , 請重新輸入 !", vbInformation, "訊息"
                                                On Error Resume Next
                                                txtFaciSNo.SetFocus
                                                Exit Function
                                            End If
                                            
                                            If strCmd = "A8" Or strCmd = "A9" Then
                                                If GetMTAMAC(strFaciSno) = Empty Then
        '                                            MsgBox "此序號 [ " & strFaciSno & " ] 入庫資料找不到 !", vbInformation, "訊息"
                                                    On Error Resume Next
                                                    txtFaciSNo.SetFocus
                                                    Exit Function
                                                End If
                                            End If
                                        End If
                                        
                                        ' EMTA 換 CM , 要檢核 , 如果 ReFaciSno 有等於 , 則 Show Msg
                                        If strCmd = "A10" Then
                                            If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                                                    " WHERE REFACISNO='" & rsData("FaciSno") & "'" & _
                                                                    " AND PRDATE IS NULL" & _
                                                                    " AND SERVICETYPE='P'") > 0 Then
                                                MsgBox "有 CP 門號 , 不可換成 CM ", vbInformation, "訊息"
                                                On Error Resume Next
                                                gilFaciCode.SetFocus
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                        Case optCMChange(9) ' 試用速率
                                ' 安裝中或正常的設備
                                If strPrcType = "" Then
                                    If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                                        If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
                                            MsgBox "無 [ 安裝中 ] 或 [ 正常 ] 之設備 , 無法下 [試用速率] 命令 !", vbInformation, "訊息"
                                            Exit Function
                                        End If
                                    End If
                                    If !ProbationStopDate & "" = Empty Or !ProbationStopFlag = 1 Then
                                    Else
                                        MsgBox "無法下 [試用速率] 命令 !", vbInformation, "訊息"
                                        Exit Function
                                    End If
                                    If gilNewRate.GetCodeNo = Empty Or gilNewRate.GetDescription & "" = rsData!CMBaudRate & "" Then
                                        MsgBox "請選擇 [ 新速率 ] 後再執行 [ 設定 ] !", vbInformation, "提示"
                                        On Error Resume Next
                                        gilNewRate.SetFocus
                                        Exit Function
                                    End If
                                    If gdaProbationStopDate.Text = Empty Then
                                        MsgBox "請輸入 [ 預計試用截止日 ] 後再執行 [ 設定 ] !", vbInformation, "提示"
                                        On Error Resume Next
                                        gdaProbationStopDate.SetFocus
                                        Exit Function
                                    End If
                                    ' 預計試用截止日 : ProbationStopDate
                                    ' 停止試用狀態 : ProbationStopFlag ( 0=否, 1=是 )
                                    ' 且(SO004.ProbationStopDate無值或SO004. ProbationStopFlag=1)才能下此命令。
                                End If
                        Case optCMChange(10) ' 終止試用
                                ' 乙、    安裝中或正常的設備且SO004.ProbationStopDate有值
                                '           且SO004. ProbationStopFlag=0才能下此命令。
                                If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                                    If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
                                        MsgBox "無 [ 安裝中 ] 或 [ 正常 ] 之設備 , 無法下 [終止試用] 命令 !", vbInformation, "訊息"
                                        Exit Function
                                    End If
                                End If
                                If !ProbationStopDate & "" <> Empty And !ProbationStopFlag = 0 Then
                                Else
                                    MsgBox "無法下 [終止試用] 命令 !", vbInformation, "訊息"
                                    Exit Function
                                End If
                End Select
            End Select
        End With
        
        If strResvTime <> "" Then
            If CDate(strResvTime) <= RightNow Then
                MsgBox "[預約時間] 需大於現在時間", vbCritical, gimsgPrompt
'                gdtResvTime.SetValue DateAdd("h", 1, RightNow)
                gdtResvTime.Enabled = True
                gdtResvTime.SetFocus
                Exit Function
            Else
                '7676
                If Len(inResvTime & "") > 0 Then
                    Dim compareDate As String
                    compareDate = inResvTime
                    If IsDate(compareDate) Then
                        compareDate = Format(compareDate, "yyyyMMddHHmmss")
                    End If
                    If CDec(gdtResvTime.GetValue(False) & "") > CDec(compareDate) Then
                        MsgBox "預約時間需小於工單預約日期", vbCritical, gimsgPrompt
                        gdtResvTime.SetValue inResvTime
                        gdtResvTime.SetFocus
                        Exit Function
                    End If
                
                End If
'                If CDate(strResvTime) > CDate(RightNow) + 1 Then
'                    MsgBox "[預約時間] 需小於明日時間", vbCritical, gimsgPrompt
'                    gdtResvTime.Enabled = True
'                    gdtResvTime.SetFocus
'                    Exit Function
'                End If
                If MsgBox("有設定預約時間, 程式將不會等待 Gateway 回應,確定使用預約功能??", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then Exit Function
            End If
        End If
        
        IsDataOK = True
    Exit Function
ChkErr:
    If Err.Number = 5 Then
        Err.Clear
    Else
        ErrSub Me.Name, "IsDataOk"
    End If
End Function

Private Function Is_EMTA_Dev() As Boolean
  On Error GoTo ChkErr
    If rsData.State > 0 Then
        If rsData("FaciCode") & "" = Empty Then
            MsgBox "注意 [客戶設備資料檔].[品名代碼] 為空!", vbInformation, "訊息"
        Else
            Is_EMTA_Dev = Val(GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD022" & _
                                                                " WHERE CODENO=" & rsData("FaciCode") & _
                                                                " AND REFNO=5", gcnGi, "") & "") > 0
        End If
    End If
  Exit Function
ChkErr:
    ErrSub Name, "Is_EMTA_Dev"
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        Call CloseRecordset(rsData)
        Call CloseRecordset(rsChild)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
'    If blnGetIP Then gcnGi.Execute "UPDATE " & GetOwner & "SO048 SET USEFLAG=0 WHERE IPADDRESS='" & strIPaddr & "'"
    If frmSO1623A.strIPaddr <> Empty And blnGetIP Then
        gcnGi.Execute "Update " & GetOwner & "SO048" & _
                                " Set UseFlag=0" & _
                                " Where IPAddress = '" & strIPaddr & "'" & _
                                " And CompCode =" & garyGi(9)
    End If
    ReleaseCOM Me
    Set frmSO1623A = Nothing
End Sub

Private Sub ggrChild_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If giFld.FieldName = "Status" Then
        If Value = 0 Then
            Value = "關"
        Else
            Value = "開"
        End If
    End If
  Exit Sub
ChkErr:
    ErrSub Name, "ggrChild_ShowCellData"
End Sub

Private Sub ggrQuery_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    Select Case UCase(giFld.FieldName)
            Case "QDOWNSTREAMFREQUENCY", "QUPSTREAMFREQUENCY"
                    Value = IIf(Value = "", "", Value & "Hz")
            Case "QDOWNSTREAMPOWER", "QUPSTREAMPOWER"
                    Value = IIf(Value = "", "", Value & "dBmV")
            Case "QDOWNSTREAMSNRATIO"
                    Value = IIf(Value = "", "", Value & "dB")
            Case "QMICROREFLECTIONS"
                    Value = IIf(Value = "", "", Value & "dBc")
            Case "MSGNAME"
                    Value = CmdDisp(Value)
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
        If UCase(giFld.FieldName) = "PRDATE" Then
            Select Case Trim(giFld.HeadName)
                Case "設備狀態"
                    Value = GetFaciStatus(ggrData.Recordset)
                Case "派工狀態"
                    Value = GetFaciWipStatus(ggrData.Recordset)
                Case "開通狀態"
                    Value = GetFaciCommandStatus(ggrData.Recordset)
            End Select
        End If
End Sub

Private Sub ggrData_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    Dim strTmp As String
    If giFld.FieldName = "PRDate" Then
        strTmp = ggrData.TextMatrix(Row, Col)
        If strTmp = "拆除" Or strTmp = "拆機" Then
            If Value <> 0 Then Value = vbRed
        End If
    End If
End Sub

Public Function GetFaciStatus(ByVal rs004 As ADODB.Recordset, Optional ByVal Value) As String
    On Error Resume Next
    Dim FaciWipStatus As String
    Dim rs As New ADODB.Recordset
        With rs004
            If .State = adStateClosed Then Exit Function
            If .EOF Or .BOF Then Exit Function
'            1.正常=>InstDate is not null and PRDate is null
'            2.停機=>InstDate is not null and PRDate is null And FaciStatusCode = 3
'            3.拆機/未取回=>PRDate is not null and GetDate is null
'            4.拆機/取回=>PRDate is not null and GetDate is not null
'            5.無:InstDate is null and PrDate is null
            If .Fields("InstDate") & "" <> "" And .Fields("PRDate") & "" = "" Then
                If .Fields("FaciStatusCode") & "" = 3 Then
                    FaciWipStatus = "停機"
                Else
                    If StrToDate(.Fields("StopTime") & "", True) > StrToDate(.Fields("ReInstTime") & "", True) Then
                        FaciWipStatus = "停機"
                    Else
                        FaciWipStatus = "正常"
                    End If
                End If
            ElseIf .Fields("PRDate") & "" <> "" Then
                FaciWipStatus = "拆機"
                If .Fields("GetDate") & "" = "" Then
                    FaciWipStatus = FaciWipStatus & "/未取回"
                End If
            Else
                FaciWipStatus = "無"
            End If
            If .Fields("GetDate") & "" <> "" Then
                FaciWipStatus = FaciWipStatus & "/取回"
            End If
        End With
        GetFaciStatus = FaciWipStatus
End Function

Public Function GetFaciWipStatus(ByVal rs004 As ADODB.Recordset) As String
  On Error Resume Next
    Dim strStatus As String
    strStatus = gcnGi.Execute("Select Kind From " & GetOwner & "SO004D Where CustId = " & rs004("CustId") & " And SeqNo = '" & rs004("SeqNo") & "' And SignDate is Null Order by UpdTime Desc").GetString(, , , "/") & ""
    GetFaciWipStatus = Mid(strStatus, 1, Len(strStatus) - 1)
End Function

Public Function GetFaciCommandStatus(ByVal rs004 As ADODB.Recordset) As String
    On Error Resume Next
    Dim FaciWipStatus As String
    Dim intFaciRef As Integer
    
        With rs004
            If .State = adStateClosed Then Exit Function
            If .EOF Or .BOF Then Exit Function
'            開通狀態
'            1.未開機=>CMOpenDate is null and CMCloseDate is null
'            2.開機=>CMOpenDate > CMCloseDate or CMOpendate is not null and CMCloseDate is null
'            3.關機=>CMOpenDate < CMCloseDate
'            4.開機/停權=>(CMOpenDate > CMCloseDate or CMOpendate is not null and CMCloseDate is null) and
'            (disableaccount> enableAccount or (disableaccount is not null and enableAccount is null)
'            5.關機/停權=>(CMOpenDate < CMCloseDate) and
'            (disableaccount> enableAccount or (disableaccount is not null and enableAccount is null)
            
            '#5227 2009.08.06 by Corey 因客戶認為ICC設備是跟STB綁在一起，為何STB開通ICC還是未開通，討論後調整 設備參考號 Null,0,4 就不需要秀出開通狀態
            intFaciRef = Val(GetRsValue("Select Refno From " & GetOwner & "CD022 Where StopFlag=0 and CodeNo=" & .Fields("FaciCode")) & "")
            If Not (intFaciRef = 4 Or intFaciRef = 0) Then
                If StrToDate(.Fields("CMOpenDate") & "", True, True) > StrToDate(.Fields("CMCloseDate") & "", True, True) Then
                    FaciWipStatus = "開機"
                ElseIf StrToDate(.Fields("CMOpenDate") & "", True, True) < StrToDate(.Fields("CMCloseDate") & "", True, True) Then
                    FaciWipStatus = "關機"
                ElseIf .Fields("CMOpenDate") & "" = "" And .Fields("CMCloseDate") & "" = "" Then
                    FaciWipStatus = "未開機"
                Else
                    FaciWipStatus = "無"
                End If
                If StrToDate(.Fields("disableaccount") & "", True, True) > StrToDate(.Fields("enableAccount") & "", True, True) Then
                    FaciWipStatus = FaciWipStatus & "/停權"
                End If
            Else
                FaciWipStatus = ""
            End If
        End With
        GetFaciCommandStatus = FaciWipStatus
End Function

'Private Function GetFaciStatus(ByVal rs004 As ADODB.Recordset) As String
'  On Error GoTo ChkErr
'    Dim rs As New ADODB.Recordset
'    With rs004
'        If .State = 0 Or (.EOF Or .BOF) Then Exit Function
'        If Not GetRS(rs, "SELECT * FROM " & GetOwner & "SO004D" & _
'                                    " WHERE SEQNO = '" & .Fields("SeqNo") & "'" & _
'                                    " AND FINTIME IS NULL" & _
'                                    " AND RETURNCODE IS NULL" & _
'                                    " ORDER BY ACCEPTTIME DESC", , 3, 1, 3) Then Exit Function
'        If Not rs.EOF Then
'            GetFaciStatus = rs("Kind") & ""
'        Else
'            If .Fields("PRSNo") & "" <> "" Then
'                If .Fields("PRDate") & "" = "" Then
'                    If Mid(.Fields("PRSNo") & "", 7, 1) = "M" Then
'                        GetFaciStatus = "取回中"
'                    Else
'                        GetFaciStatus = "拆除中"
'                    End If
'                Else
'                    If Mid(.Fields("PRSNo") & "", 7, 1) = "M" Then
'                        GetFaciStatus = "取回"
'                    Else
'                        GetFaciStatus = "拆除"
'                    End If
'                End If
'            ElseIf .Fields("PRSNo") & "" = "" And .Fields("PRDate") & "" = "" And .Fields("InstDate") & "" <> "" Then
'                GetFaciStatus = "正常"
'            ElseIf .Fields("InstDate") & "" = "" Then
'                GetFaciStatus = "安裝中"
'            End If
'        End If
'    End With
'  Exit Function
'ChkErr:
'    ErrSub Name, "GetFaciStatus"
'End Function

Private Sub optBasic_Click(Index As Integer)
  On Error GoTo ChkErr
'    Select Case Index
'            Case 2
'                    If txtDynaIP = Empty Then txtDynaIP = rsData("DynIPCount") & ""
'                    If txtFixIP = Empty Then txtFixIP = rsData("FixIPCount") & ""
'    End Select
     '#6542 如果Option無法使用，則要關掉設定按鈕 By Kin 2013/07/11
    cmdSet.Enabled = True
    If Not optBasic(Index).Enabled Then
        cmdSet.Enabled = False
    End If
  Exit Sub
ChkErr:
    ErrSub Name, "optBasic_Click"
End Sub

Private Sub optCMChange_Click(Index As Integer)
  On Error GoTo ChkErr
    Select Case Index
'                Case 2
'                        If txtDynaIP2 = Empty Then txtDynaIP2 = rsData("DynIPCount") & ""
'                        If txtFixIP2 = Empty Then txtFixIP2 = rsData("FixIPCount") & ""
'                Case 4
'                        If cboCPE = Empty Then cboCPE = Get_CPE
        Case 5 ' 7. CPE IP 異動
            chkPR.Value = Abs(GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO002" & _
                                            " WHERE COMPCODE=" & gilCompCode.GetCodeNo & _
                                            " AND CUSTID=" & ginCustId.Text & _
                                            " AND WIPCODE3=11", gcnGi, "") > 0)
'                Case 7
'                        cboNode = IIf(int_EMTA_IP_Type = 0, strCircuitNo, strNodeNo)
'                        If Is_EMTA_Dev Then txtNewIP = rsData("IPAddress") & ""
'                        chkPR.Value = Abs(GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO002" & _
'                                        " WHERE COMPCODE=" & gilCompCode.GetCodeNo & _
'                                        " AND CUSTID=" & ginCustId.Text & _
'                                        " AND WIPCODE3=11", gcnGi, "") > 0)
    End Select
    '#6542 如果Option無法使用，則要關掉設定按鈕 By Kin 2013/07/11
    cmdSet.Enabled = True
    If Not optCMChange(Index).Enabled Then
        cmdSet.Enabled = False
    End If
  Exit Sub
ChkErr:
    ErrSub Name, "optCMChange_Click"
End Sub

Private Sub sstData_Click(PreviousTab As Integer)
  On Error Resume Next
    Dim i As Integer
    Select Case sstData.Tab
        Case 0
            '#6542 如果Option無法使用，則要關掉設定按鈕 By Kin 2013/07/11
            cmdSet.Enabled = True
            For i = optBasic.LBound To optBasic.UBound
                If optBasic(i).Value Then
                    If Not optBasic(i).Enabled Then
                        cmdSet.Enabled = False
                    End If
                    Exit For
                End If
            Next
        Case 1
            '所有服務變更異動必須該選定的設備（EMTA、CM）必須是正常使用中或裝機中且（關機日期是空的或開機日期大於關機日）
            '才可以作下列指令功能，若不符合此頁籤功能項目都DISABLE 空白處顯示 此CM 非正常使用中或裝機中 無法作指令設定
            If Alfa Then Exit Sub
            '#6542 如果Option無法使用，則要關掉設定按鈕 By Kin 2013/07/11
            cmdSet.Enabled = True
            For i = optCMChange.LBound To optCMChange.UBound
                If optCMChange(i).Value Then
                    If Not optCMChange(i).Enabled Then
                        cmdSet.Enabled = False
                    End If
                    Exit For
                End If
            Next

            
            With rsData
                Select Case True
                    Case Not IsNull(!CMOpenDate) And IsNull(!CMCloseDate)
                        fraData(1).Enabled = True
'                        sstData.TabEnabled(1) = True
                    Case Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate)
                        fraData(1).Enabled = CDT(!CMCloseDate) < CDT(!CMOpenDate)
'                        sstData.TabEnabled(1) = CDT(!CMCloseDate) < CDT(!CMOpenDate)
                    Case Else ' IsNull(!CMOpenDate) And IsNull(!CMCloseDate)
                        fraData(1).Enabled = False
'                        sstData.TabEnabled(1) = False
                End Select
                If rsData("FaciCode") <> "" And strFC = "" Then
                    gilFaciCode.SetCodeNo rsData("FaciCode") & ""
                    gilFaciCode.Query_Description
                End If
            End With
        Case 2
            '#6542 切換頁面權限也要判斷 By Kin 2013/07/25
            cmdSet.Enabled = True
            For i = optQueryInfo.LBound To optQueryInfo.UBound
                If optQueryInfo(i).Value Then
                    If Not optQueryInfo(i).Enabled Then
                        cmdSet.Enabled = False
                    End If
                    Exit For
                End If
            Next

'        chkSetAll.Value = 0
'        chkSetAll.Visible = False
    End Select
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If gilCompCode.GetCodeNo = "" Then Exit Sub
        garyGi(16) = gilCompCode.GetOwner
        garyGi(9) = gilCompCode.GetCodeNo
        int_EMTA_IP_Type = GetSystemParaItem("EMTAIPTYPE", gCompCode, gServiceType, "SO041", , True)
        strFaciAuthorityField = "PRDATE"
        '#6154 依據SO042.FaciAuthority的參數來判斷欄位 By Kin 2011/11/08
        If GetRsValue("SELECT FaciAuthority FROM " & GetOwner & "SO042 WHERE COMPCODE = " & gCompCode & _
                    " AND ServiceType = '" & gServiceType & "'", gcnGi) = 2 Then
            strFaciAuthorityField = "GetDate"
        End If
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

Private Sub rsData_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
        If Not blnSetting Then
            Me.Enabled = False
            Call OpenData2
'            Fetch_Log_Data , True
            Me.Enabled = True
        End If
End Sub

'Public Function GetPCMac() As String
'    On Error Resume Next
'        GetPCMac = gcnGi.Execute("SELECT DISTINCT CPEMAC FROM " & GetOwner & "SO004C" & _
'                                                    " WHERE FACISNO = '" & rsData("FaciSNo") & "'" & _
'                                                    " AND CUSTID = " & rsData("CustID")).GetString(, , , "#")
'End Function

Public Sub ShowResult(ByRef rsObj As Object, ByRef grdObj As Object)
  On Error GoTo ChkErr
'    Dim mFlds As New GiGridFlds
    Set mFlds = New GiGridFlds
    grdObj.Blank = True
    '4.  顯示順序為：處理說明、設定人員、設定時間、執行名稱、執行結果．．．．。
    With mFlds
            .Add "HandleNote", , , , , "處理說明    ", vbLeftJustify
            .Add "EXECENTRY", , , , , "設定人員    ", vbLeftJustify
            .Add "EXECTIME", giControlTypeTime, , , , "設定日期時間    ", vbLeftJustify
            
            If strCmd <> "Q2" Then
                .Add "MSGNAME", , , , , "執行名稱     ", vbLeftJustify
                .Add "COMMANDSTATUS", , , , , "執行狀態", vbLeftJustify
                .Add "CMDSEQNO", , , , , "指令流水序號          ", vbLeftJustify
                .Add "MODEMMAC", , , , , "CM MAC Address", vbLeftJustify
                .Add "NEWMODEMMAC", , , , , "換新CM MAC   ", vbLeftJustify
                .Add "MODELNAME", , , , , "型號MODEL   ", vbLeftJustify
                .Add "NEWMODELNAME", , , , , "型號MODEL(新的)", vbLeftJustify
                .Add "CMBAUDRATE", , , , , "CM速率     ", vbLeftJustify
                .Add "NEWCMBAUDRATE", , , , , "CM速率(新的)", vbLeftJustify
                .Add "DYNIPCOUNT", , , , , "動態IP數目", vbRightJustify
                .Add "NEWDYNIPCOUNT", , , , , "動態IP數目(要增減的)", vbRightJustify
                .Add "FIXIPCOUNT", , , , , "固定IP數目", vbRightJustify
                .Add "NEWFIXIPCOUNT", , , , , "固定IP數目(要增減的)", vbRightJustify
                .Add "MEDIAGATEWAY", , , , , "EMTA CP的Media Gateway", vbLeftJustify
                .Add "CPNUMBER", , , , , "CP電話       ", vbLeftJustify
                .Add "TEL", , , , , "CM申裝人電話", vbLeftJustify
                .Add "HFCNODE", , , , , "網路編號(光投落點)", vbLeftJustify
                .Add "NEWHFCNODE", , , , , "網路編號(光投落點)新的", vbLeftJustify
                .Add "CPEMAC", , , , , "原本網卡MAC1#網卡MAC2", vbLeftJustify
                .Add "ADDCPEMAC", , , , , "新增網卡MAC1#網卡MAC2", vbLeftJustify
                .Add "DELCPEMAC", , , , , "取消網卡MAC    ", vbLeftJustify
                .Add "CPESTATIP", , , , , "原本的ip1#ip2   ", vbLeftJustify
                .Add "ADDCPESTATIP", , , , , "新增的ip1#ip2    ", vbLeftJustify
                .Add "DELCPESTATIP", , , , , "取消的ip            ", vbLeftJustify
                .Add "CPEIPRULE", , , , , "CPE IP Rule", vbLeftJustify
                .Add "IPADDRESS", , , , , "EMTA IP Address", vbLeftJustify
                .Add "NEWIPADDRESS", , , , , "(新的)EMTA IP Address", vbLeftJustify
                .Add "CMIPRULE", , , , , "CM IP Rule", vbLeftJustify
                .Add "MTAMAC", , , , , "MTA MAC    ", vbLeftJustify
                .Add "NEWMTAMAC", , , , , "MTA MAC NEW", vbLeftJustify
                .Add "RETCODE", , , , , "回應代碼攔", vbRightJustify
                .Add "RETMSG", , , , , "回應訊息攔", vbLeftJustify
    '            .Add "CLIFAILURECODE", , , , , "執行結果的指令交易錯誤訊息代碼", vbLeftJustify
                .Add "GATEEXECTIME", giControlTypeDate, , , , "CLI開始執行日期時間", vbLeftJustify
                .Add "GATEFINSHTIME", giControlTypeDate, , , , "CLI執行結束或客服系統timeout日期時間", vbLeftJustify
                .Add "GATEROLLBACKTIME", giControlTypeDate, , , , "gateway執行rollback成功或失敗日期時間", vbLeftJustify
                .Add "SNO", , , , , "工單單號", vbLeftJustify
                .Add "ORDERNO", , , , , "訂單單號", vbLeftJustify
                .Add "RESVTIME", giControlTypeDate, , , , "指令預約日期時間", vbLeftJustify
                .Add "CMDRESEND", , , , , "重送指令新的指令流水號", vbLeftJustify
            End If
            .Add "QCMIPADDRESS", , , , , "CMIPAddress ", vbLeftJustify
            .Add "QCPEIPADDRESS", , , , , "CPEIPAddress ", vbLeftJustify
            .Add "QDESCRIPTION", , , , , "Description  ", vbLeftJustify
            .Add "QUPTIME", , , , , "Uptime  ", vbLeftJustify
            .Add "QSERIAL", , , , , "Serial  ", vbLeftJustify
            .Add "QSTATUS", , , , , "Status  ", vbLeftJustify
            .Add "QDOWNSTREAMFREQUENCY", , , , , "Downstream_Frequency", vbRightJustify
            .Add "QDOWNSTREAMPOWER", , , , , "Downstream Power", vbRightJustify
            .Add "QERRORLESS", , , , , "Errorless  ", vbRightJustify
            .Add "QCORRECTABLE", , , , , "Correctable ", vbRightJustify
            .Add "QUNCORRECTABLE", , , , , "Uncorrectable", vbRightJustify
            .Add "QDOWNSTREAMSNRATIO", , , , , "Downstream S/N Ratio", vbRightJustify
            .Add "QMICROREFLECTIONS", , , , , "Microreflections", vbRightJustify
            .Add "QDOWNSTREAMMODULATION", , , , , "Downstream Modulation", vbLeftJustify
            .Add "QUPSTREAMFREQUENCY", , , , , "Upstream Frequency", vbRightJustify
            .Add "QUPSTREAMPOWER", , , , , "Upstream Power", vbRightJustify
            .Add "QTIMINGOFFSET", , , , , "Timing Offset", vbRightJustify
            .Add "QRESETS", , , , , "Resets", vbRightJustify
            .Add "QLOSTSYNC", , , , , "Lost_Sync", vbRightJustify
            .Add "QINVALIDMAPS", , , , , "Invalid MAPs", vbRightJustify
            .Add "QINVALIDUCDS", , , , , "Invalid Ucds", vbRightJustify
            .Add "QINVALIDRANGING", , , , , "Invalid Ranging", vbRightJustify
            .Add "QINVALIDREGISTRATIO", , , , , "Invalid Registratio", vbRightJustify
            .Add "QT1TIMEOUTS", , , , , "T1 timeouts", vbRightJustify
            .Add "QT2TIMEOUTS", , , , , "T2 timeouts", vbRightJustify
            .Add "QT3TIMEOUTS", , , , , "T3 timeouts", vbRightJustify
            .Add "QT4TIMEOUTS", , , , , "T4 timeouts", vbRightJustify
            .Add "QRANGINGABORTS", , , , , "Ranging Aborts", vbRightJustify
            .Add "QQOSSTATUS", , , , , "QoS Status", vbRightJustify
            .Add "QQOSPRIORITY", , , , , "QoS Priority", vbRightJustify
            .Add "QQOSGUBANDWIDTH", , , , , "QoS Guaranteed Upstream Bandwidth", vbRightJustify
            .Add "QQOSMUBANDWIDTH", , , , , "QoS Maximum Upstream Bandwidth", vbRightJustify
            .Add "QQOSMDBANDWIDTH", , , , , "QoS Maximum Downstream Bandwidth", vbRightJustify
            .Add "QQOSMTBURST", , , , , "QoS Maximum Transmit Burst", vbRightJustify
            .Add "QQOSBPIENABLED", , , , , "QoS BPI Enabled", vbLeftJustify
'            .Add "NEWCPEIPRULE", , , , , "New CPE IP Rule", vbLeftJustify

            .Add "Q_PROTOCOL", , , , , "PROTOCOL", vbLeftJustify
            .Add "Q_RELAY", , , , , "RELAY", vbLeftJustify
            .Add "Q_STATE", , , , , "STATE", vbLeftJustify
            .Add "Q_MACADDR", , , , , "MACADDR", vbLeftJustify
            .Add "Q_HARDWARETYPE", , , , , "HARDWARETYPE", vbLeftJustify
            .Add "Q_HOSTSENT", , , , , "HOSTSENT", vbLeftJustify
            .Add "Q_HOSTRECEIVED", , , , , "HOSTRECEIVED", vbLeftJustify
            .Add "Q_REMOTEID", , , , , "REMOTEID", vbLeftJustify
            .Add "Q_CIRCUITID", , , , , "CIRCUITID", vbLeftJustify
            .Add "Q_CLIENTID", , , , , "CLIENTID", vbLeftJustify
            .Add "Q_VENDORID", , , , , "VENDORID", vbLeftJustify
            .Add "Q_DOCSISCLASS", , , , , "DOCSISCLASS", vbLeftJustify
            .Add "Q_ANTI_ROAMING", , , , , "ANTI_ROAMING", vbLeftJustify
            .Add "Q_IPSHUFFLED", , , , , "IPSHUFFLED", vbLeftJustify
            .Add "Q_LOOKUPKEY", , , , , "LOOKUPKEY", vbLeftJustify
            .Add "Q_ALLOCATORKEY", , , , , "ALLOCATORKEY", vbLeftJustify
            .Add "Q_HWMAPPINGKEY", , , , , "HWMAPPINGKEY", vbLeftJustify
            .Add "Q_INITIALTIME", , , , , "INITIALTIME", vbLeftJustify
            .Add "Q_STARTTIME", , , , , "STARTTIME", vbLeftJustify
            .Add "Q_LEASETIME", , , , , "LEASETIME", vbLeftJustify
            .Add "Q_ALLOCATEDBYPRIMARY", , , , , "ALLOCATEDBYPRIMARY", vbLeftJustify
            .Add "Q_DDNSSERVER", , , , , "DDNSSERVER", vbLeftJustify
            .Add "Q_SUBNET_MASK", , , , , "Subnet Mask", vbLeftJustify
            .Add "Q_TIME_OFFSET", , , , , "Time Offset", vbLeftJustify
            .Add "Q_GATEWAYS", , , , , "Gateways", vbLeftJustify
            .Add "Q_TIME_SERVER", , , , , "Time Server", vbLeftJustify
            .Add "Q_DOMAIN_SERVER1", , , , , "Domain Server1", vbLeftJustify
            .Add "Q_DOMAIN_SERVER2", , , , , "Domain Server2", vbLeftJustify
            .Add "Q_DOMAIN_SERVER3", , , , , "Domain Server3", vbLeftJustify
            .Add "Q_DOMAIN_SERVER4", , , , , "Domain Server4", vbLeftJustify
            .Add "Q_DOMAIN_SERVER5", , , , , "Domain Server5", vbLeftJustify
            .Add "Q_DOMAIN_SERVER6", , , , , "Domain Server6", vbLeftJustify
            .Add "Q_DOMAIN_SERVER7", , , , , "Domain Server7", vbLeftJustify
            .Add "Q_LOG_SERVER", , , , , "Log Server", vbLeftJustify
            .Add "Q_DOMAIN_NAME", , , , , "Domain Name", vbLeftJustify
            .Add "Q_LEASE_TIME", , , , , "Lease Time", vbLeftJustify
            .Add "Q_RENEWAL_TIME", , , , , "Renewal Time", vbLeftJustify
            .Add "Q_REBINDING_TIME", , , , , "Rebinding Time", vbLeftJustify
            .Add "Q_TFTP_SERVER_NAME", , , , , "TFTP Server Name", vbLeftJustify
            .Add "Q_BOOTFILE_NAME", , , , , "Bootfile Name", vbLeftJustify
            .Add "Q_RELAY_AGENT_INFORMATION", , , , , "Relay Agent Information", vbLeftJustify
            .Add "Q_PACKETCABLE_VOIP", , , , , "PacketCable VoIP (draft-2)", vbLeftJustify
    End With
    With grdObj
            .AllFields = mFlds
            .SetHead
            Set .Recordset = rsObj
            .Refresh
            .Blank = False
    End With
  Exit Sub
ChkErr:
    ErrSub Name, "ShowResult"
End Sub

'Public Property Let uSno(ByVal vData As String)
'    strSno = vData
'End Property
'
'Public Property Let uMediaBillNo(ByVal vData As String)
'    strMediaBillNo = vData
'End Property

Public Property Let uCaption(ByVal vData As String)
    strCaption = vData
End Property

Public Property Let uCustId(ByVal vData As Long)
    lngCustId = vData
End Property

Public Property Set uParentSTBRs(ByVal vData As ADODB.Recordset)
    Set rsParent = vData
End Property

Public Property Set uParentRs(ByVal vData As ADODB.Recordset)
    Set rsParent = vData
End Property

'Public Property Set uUpdRs(ByVal vData As ADODB.Recordset)
'    Set rsUpd = vData
'End Property

Public Property Let uProcessType(ByVal vData As String)
    On Error Resume Next
        strPrcType = vData
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
        uProcessOk = blnPrcOK
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

Public Function Qry_Result(blnShowMsg As Boolean) As Boolean
    Qry_Result = mod_ControlPanel.Qry_Result(blnShowMsg)
End Function
'
'Public Function Cmd_14(blnShowMsg As Boolean) As Boolean
'    Cmd_14 = mod_ControlPanel.Cmd_14(blnShowMsg)
'End Function
'
'Public Function Cmd_16(blnShowMsg As Boolean) As Boolean
'    Cmd_16 = mod_ControlPanel.Cmd_16(blnShowMsg)
'End Function

'Public Function Qry_01(blnShowMsg As Boolean) As Boolean
'    Qry_01 = mod_ControlPanel.Qry_01(blnShowMsg)
'End Function
'
'Public Function Qry_02(blnShowMsg As Boolean) As Boolean
'    Qry_02 = mod_ControlPanel.Qry_02(blnShowMsg)
'End Function
'
'Public Function Qry_03(blnShowMsg As Boolean) As Boolean
'    Qry_03 = mod_ControlPanel.Qry_03(blnShowMsg)
'End Function
'
'Public Function Qry_04(blnShowMsg As Boolean) As Boolean
'    Qry_04 = mod_ControlPanel.Qry_04(blnShowMsg)
'End Function
'
'Public Function Qry_05(blnShowMsg As Boolean) As Boolean
'    Qry_05 = mod_ControlPanel.Qry_05(blnShowMsg)
'End Function
'
'Public Function Qry_06(blnShowMsg As Boolean) As Boolean
'    Qry_06 = mod_ControlPanel.Qry_06(blnShowMsg)
'End Function

Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
    Set frmSO1623A = Nothing
End Sub

Private Sub txtDynaIP_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    OnlyNum KeyAscii
End Sub

Private Sub txtDynaIP2_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    OnlyNum KeyAscii
End Sub

Private Sub txtDynaIP3_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    OnlyNum KeyAscii
End Sub

Private Sub txtFixIP_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    OnlyNum KeyAscii
End Sub

Private Sub txtFixIP2_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    OnlyNum KeyAscii
End Sub

Private Sub txtFixIP3_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    OnlyNum KeyAscii
End Sub

Private Sub sstHead_Click(PreviousTab As Integer)
  On Error Resume Next
    If sstHead.Tab = 0 Then
        picResv.Top = 3480
        picResv.Left = 7500
    Else
        picResv.Top = 6830
        picResv.Left = 7500
    End If
End Sub

Private Sub txtAddCPE_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    txt_KP KeyAscii
End Sub

Private Function ChkMAC(strMac As String) As Boolean
  On Error GoTo ChkErr
    Dim varElement, varData
    varData = Split(strMac, "#")
    ChkMAC = True
    For Each varElement In varData
        If varElement & "" <> Empty Then
            If Not MAC_IS_OK(varElement & "") Or InStr(1, varElement, "G", 1) > 0 Then
                ChkMAC = False
                Exit For
            End If
        End If
    Next
  Exit Function
ChkErr:
    ErrSub Name, "ChkMAC"
End Function

Private Function MAC_IS_OK(strMac As String) As Boolean
  On Error GoTo ChkErr
    Dim varElement, varData
    varData = Split(strMac, ":")
    If UBound(varData) = 5 Then
        MAC_IS_OK = True
        For Each varElement In varData
            If varElement = Empty Or Val("&H" & varElement) > 255 Then
                MAC_IS_OK = False
                Exit For
            End If
        Next
    End If
  Exit Function
ChkErr:
    ErrSub Name, "MAC_IS_OK"
End Function

Private Sub txtAddCPE_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If txtAddCPE <> Empty Then
        If Not ChkMAC(txtAddCPE) Then
            MsgBox "CPE MAC 不正確 .. 請確認 !", vbInformation, "訊息"
        End If
    End If
  Exit Sub
ChkErr:
    ErrSub Name, "txtAddCPE_Validate"
End Sub

Private Sub txtAddDynaIP_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    CantSpace KeyAscii
End Sub

Private Sub txtAddFixIP_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    CantSpace KeyAscii
End Sub

Private Sub txtCPE_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    txt_KP KeyAscii
End Sub

Private Sub CantSpace(KeyAscii As Integer)
  On Error Resume Next
    If KeyAscii = 32 Then KeyAscii = 0
End Sub

Private Sub txtDelDynaIP_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    CantSpace KeyAscii
End Sub

Private Sub txtFaciSNo_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    CantSpace KeyAscii
End Sub

Private Sub txtFixIPcnt1_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    CantSpace KeyAscii
End Sub

Private Sub txtHFCNode_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    CantSpace KeyAscii
End Sub

Private Sub txtIPAddress_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    CantSpace KeyAscii
End Sub
