VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO1623A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "VOD �]�w����x [SO1171A]"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   BeginProperty Font 
      Name            =   "�s�ө���"
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
   ScaleHeight     =   7260
   ScaleWidth      =   11895
   StartUpPosition =   2  '�ù�����
   Begin VB.PictureBox pic1 
      Appearance      =   0  '����
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   3500
      ScaleHeight     =   825
      ScaleWidth      =   4995
      TabIndex        =   19
      Top             =   2130
      Visible         =   0   'False
      Width           =   5025
      Begin MSComctlLib.ProgressBar pbr1 
         Height          =   285
         Left            =   150
         TabIndex        =   21
         Top             =   405
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblProcess 
         Alignment       =   2  '�m�����
         BackColor       =   &H00FFFFFF&
         Caption         =   "�}���� , �еy�J !!"
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
         TabIndex        =   20
         Top             =   110
         Width           =   4695
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '����
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4830
      ScaleHeight     =   285
      ScaleWidth      =   2805
      TabIndex        =   24
      Top             =   3030
      Width           =   2805
      Begin Gi_Time.GiTime gdtResvTime 
         Height          =   285
         Left            =   990
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
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
      Begin VB.Label lblResv 
         AutoSize        =   -1  'True
         Caption         =   "�w���ɶ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   180
         Left            =   180
         TabIndex        =   25
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
      Caption         =   "�Ҧ�CM�@�_�]�w"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      TabIndex        =   22
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
         Name            =   "�s�ө���"
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
      Caption         =   "����(&X)"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10530
      TabIndex        =   14
      Top             =   30
      Width           =   1245
   End
   Begin VB.TextBox txtCustName 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
      TabIndex        =   10
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
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " �]�w�@�~"
      TabPicture(0)   =   "SO1623A.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "sstData"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSet"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "�]�w�O��"
      TabPicture(1)   =   "SO1623A.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblReturnCode"
      Tab(1).Control(1)=   "gilReturnCode"
      Tab(1).Control(2)=   "ggrQuery"
      Tab(1).Control(3)=   "cmdQuery"
      Tab(1).Control(4)=   "cmdQuery2"
      Tab(1).Control(5)=   "cmdCancelCMD"
      Tab(1).Control(6)=   "cmdQueryAccInfo"
      Tab(1).ControlCount=   7
      Begin VB.CommandButton cmdQueryAccInfo 
         Caption         =   "�d��VOD�b��(F4)"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -71100
         TabIndex        =   56
         Top             =   3660
         Width           =   1785
      End
      Begin VB.CommandButton cmdCancelCMD 
         Caption         =   "�w�����O����(F10)"
         Height          =   345
         Left            =   -69150
         TabIndex        =   32
         Top             =   3660
         Width           =   1965
      End
      Begin VB.CommandButton cmdQuery2 
         Caption         =   "�d�ߤw��]�Ƴ]�w"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -65580
         TabIndex        =   23
         Top             =   3660
         Width           =   1905
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "�d��(F3)"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66900
         TabIndex        =   13
         Top             =   3660
         Width           =   1125
      End
      Begin prjGiGridR.GiGridR ggrQuery 
         Height          =   3015
         Left            =   -74670
         TabIndex        =   18
         Top             =   450
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   5318
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseCellForeColor=   -1  'True
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "F2. �]�w"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10170
         TabIndex        =   12
         Top             =   420
         Width           =   1245
      End
      Begin TabDlg.SSTab sstData 
         Height          =   3555
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   6271
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   5
         TabHeight       =   520
         Enabled         =   0   'False
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&1. �]�w"
         TabPicture(0)   =   "SO1623A.frx":047A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraData(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&2.SMIT�i���R�O"
         TabPicture(1)   =   "SO1623A.frx":0496
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraData(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "&3.IOS��d���A�d��"
         TabPicture(2)   =   "SO1623A.frx":04B2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraData(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "&4.�]�ƥX�J�w"
         TabPicture(3)   =   "SO1623A.frx":04CE
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "optCMDevice(0)"
         Tab(3).Control(1)=   "optCMDevice(1)"
         Tab(3).ControlCount=   2
         Begin VB.OptionButton optCMDevice 
            Caption         =   "&2.�]�ƥX�w [41]"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   -74640
            TabIndex        =   9
            ToolTipText     =   "�NSTB �Ȱ�"
            Top             =   960
            Width           =   1515
         End
         Begin VB.OptionButton optCMDevice 
            Caption         =   "&1.�]�ƤJ�w [40]"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   -74640
            TabIndex        =   8
            ToolTipText     =   "�NSTB �Ȱ�"
            Top             =   600
            Width           =   1515
         End
         Begin VB.Frame fraData 
            Height          =   2955
            Index           =   2
            Left            =   -74790
            TabIndex        =   27
            Top             =   420
            Width           =   10635
            Begin VB.OptionButton optOTTQuery 
               Caption         =   "&A.IOS��d���A�d��[C2]"
               Height          =   405
               Index           =   0
               Left            =   180
               TabIndex        =   58
               Top             =   330
               Width           =   2355
            End
            Begin prjGiList.GiList gilOTTCitem 
               Height          =   315
               Left            =   2730
               TabIndex        =   57
               Top             =   360
               Width           =   2940
               _ExtentX        =   5186
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.Frame fraData 
            Height          =   2955
            Index           =   1
            Left            =   -74850
            TabIndex        =   26
            Top             =   450
            Width           =   10725
            Begin VB.OptionButton optSMITCommand 
               Caption         =   "&6.����CatchUp [B18]"
               Height          =   375
               Index           =   5
               Left            =   360
               TabIndex        =   51
               Top             =   2340
               Width           =   2445
            End
            Begin VB.OptionButton optSMITCommand 
               Caption         =   "&5.�}��CatchUp [B17]"
               Height          =   375
               Index           =   4
               Left            =   360
               TabIndex        =   50
               Top             =   1890
               Width           =   1935
            End
            Begin VB.OptionButton optSMITCommand 
               Caption         =   "&4.���W�������M�\Prod ID [B16]"
               Height          =   285
               Index           =   3
               Left            =   360
               TabIndex        =   49
               Top             =   1530
               Width           =   2955
            End
            Begin VB.OptionButton optSMITCommand 
               Caption         =   "&3.���W���}�q�M�\Prod ID [B15]"
               Height          =   375
               Index           =   2
               Left            =   360
               TabIndex        =   48
               Top             =   1080
               Width           =   2955
            End
            Begin VB.OptionButton optSMITCommand 
               Caption         =   "&2.�������v�e�q [B14]"
               Height          =   375
               Index           =   1
               Left            =   360
               TabIndex        =   47
               Top             =   660
               Width           =   2145
            End
            Begin VB.OptionButton optSMITCommand 
               Caption         =   "&1.�]�w���v�e�q [B13]"
               Height          =   375
               Index           =   0
               Left            =   360
               TabIndex        =   45
               Top             =   240
               Width           =   2145
            End
            Begin VB.OptionButton optSMITCommand 
               Caption         =   "&7.�]�w���W��Tuner�ƶq [A7]"
               Height          =   375
               Index           =   6
               Left            =   5790
               TabIndex        =   44
               Top             =   240
               Width           =   2655
            End
            Begin prjGiList.GiList gilCitemCode 
               Height          =   285
               Left            =   2640
               TabIndex        =   46
               Top             =   270
               Width           =   2730
               _ExtentX        =   4815
               _ExtentY        =   503
               Enabled         =   0   'False
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
            Begin CS_Multi.CSmulti gimOrdSmitProd 
               Height          =   435
               Left            =   3390
               TabIndex        =   52
               Top             =   1080
               Width           =   7275
               _ExtentX        =   12832
               _ExtentY        =   767
               ButtonCaption   =   "�q���W�D"
               Enabled         =   0   'False
            End
            Begin CS_Multi.CSmulti gimCancelSmitProd 
               Height          =   435
               Left            =   3390
               TabIndex        =   53
               Top             =   1500
               Width           =   7275
               _ExtentX        =   12832
               _ExtentY        =   767
               ButtonCaption   =   "�����W�D"
               Enabled         =   0   'False
            End
            Begin CS_Multi.CSmulti gimOrdCatchUpProd 
               Height          =   435
               Left            =   3390
               TabIndex        =   54
               Top             =   1920
               Visible         =   0   'False
               Width           =   7275
               _ExtentX        =   12832
               _ExtentY        =   767
               ButtonCaption   =   "�q���W�D"
            End
            Begin CS_Multi.CSmulti gimCancelCatchUpProd 
               Height          =   435
               Left            =   3390
               TabIndex        =   55
               Top             =   2340
               Visible         =   0   'False
               Width           =   7275
               _ExtentX        =   12832
               _ExtentY        =   767
               ButtonCaption   =   "�����W�D"
            End
         End
         Begin VB.Frame fraData 
            Height          =   2955
            Index           =   0
            Left            =   180
            TabIndex        =   17
            Top             =   420
            Width           =   10665
            Begin VB.OptionButton optBasic 
               Caption         =   "&10.���m"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   10
               Left            =   3060
               TabIndex        =   43
               ToolTipText     =   "�w��FTTX�˾�"
               Top             =   2280
               Width           =   1995
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&9.���״���(ICC)[A9]"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   9
               Left            =   3060
               TabIndex        =   42
               ToolTipText     =   "�w��FTTX�˾�"
               Top             =   1710
               Width           =   1995
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&8.���״���(STB)[A8]"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   8
               Left            =   3060
               TabIndex        =   41
               ToolTipText     =   "�w��FTTX�˾�"
               Top             =   1140
               Width           =   1995
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&6.���״���(���)[A6]"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   7
               Left            =   3060
               TabIndex        =   40
               ToolTipText     =   "�w��FTTX�˾�"
               Top             =   630
               Width           =   1995
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "8.��SVOD�W�D[B12]"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   6
               Left            =   4200
               TabIndex        =   38
               ToolTipText     =   "�w��FTTX�˾�"
               Top             =   180
               Visible         =   0   'False
               Width           =   1995
            End
            Begin CS_Multi.CSmulti gimOrdSVOD 
               Height          =   435
               Left            =   5130
               TabIndex        =   37
               Top             =   540
               Width           =   5145
               _ExtentX        =   9075
               _ExtentY        =   767
               ButtonCaption   =   "�q��SVOD"
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "7.�}SVOD�W�D[B11]"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   5
               Left            =   2070
               TabIndex        =   36
               ToolTipText     =   "�w��FTTX�˾�"
               Top             =   180
               Visible         =   0   'False
               Width           =   1995
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&5.���]�q�ʱK�X[E1]"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   4
               Left            =   5220
               TabIndex        =   35
               ToolTipText     =   "�w��FTTX�˾�"
               Top             =   1500
               Visible         =   0   'False
               Width           =   1995
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&1.�Ұʭq���v [A1]"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   480
               TabIndex        =   28
               ToolTipText     =   "�}��"
               Top             =   600
               Width           =   1845
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&4.��_�q���v [A4]"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   3
               Left            =   480
               TabIndex        =   7
               ToolTipText     =   $"SO1623A.frx":04EA
               Top             =   2280
               Width           =   1785
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&3.�Ȱ��q���v[A3]"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   2
               Left            =   480
               TabIndex        =   6
               ToolTipText     =   "����"
               Top             =   1740
               Width           =   1755
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&2.�����q���v [A2]"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   480
               TabIndex        =   5
               ToolTipText     =   $"SO1623A.frx":04F4
               Top             =   1170
               Width           =   1935
            End
            Begin prjGiList.GiList gilReason 
               Height          =   285
               Left            =   7380
               TabIndex        =   30
               Top             =   150
               Visible         =   0   'False
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "�s�ө���"
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
            Begin CS_Multi.CSmulti gimCancelSVOD 
               Height          =   435
               Left            =   5130
               TabIndex        =   39
               Top             =   1080
               Width           =   5145
               _ExtentX        =   9075
               _ExtentY        =   767
               ButtonCaption   =   "����SVOD"
            End
            Begin VB.Label lblReason 
               Caption         =   "���ʭ�]"
               Height          =   195
               Left            =   6510
               TabIndex        =   31
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
         End
      End
      Begin prjGiList.GiList gilReturnCode 
         Height          =   285
         Left            =   -66600
         TabIndex        =   34
         Top             =   30
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
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
      Begin VB.Label lblReturnCode 
         AutoSize        =   -1  'True
         Caption         =   "�h���]"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   -67380
         TabIndex        =   33
         Top             =   90
         Width           =   720
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
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseCellForeColor=   -1  'True
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   2265
      Left            =   180
      TabIndex        =   3
      Top             =   600
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   3995
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
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
      Caption         =   "�Ȥ�s��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4260
      TabIndex        =   16
      Top             =   90
      Width           =   720
   End
   Begin VB.Label lblComp 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��  �q  �O"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      TabIndex        =   15
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
    Dim WithEvents rsLog As ADODB.Recordset
Attribute rsLog.VB_VarHelpID = -1
    Dim lngCustId As Long
    Dim lngInstAddrNo As Long
    Dim lngNodeNo As Long
    Dim lngCircuitNo As Long
    Dim strZipCode As String
    'Dim strChStr As String, strStopChStr As String, strOpenChStr As String
    'Dim strChStopDate As String
    Dim rsParent As New ADODB.Recordset
    Dim strProcessType As String
    Dim blnProcessOk As Boolean
    Dim strProcessChStr As String
    Dim strDoCaption As String
    Dim strCaption As String
    Dim blnTransation As Boolean
    'Dim strResvTime As String
    Dim blnSetting As Boolean
    Dim strDefaultRowId As String
    Dim strOrderNo As String
    Dim blnProcessLock As Boolean
    Dim strCustName As String
    Dim strCustTel As String
    Dim strInstAddress As String
    Dim intByHouse As Integer
    Dim blnAutoSend As Boolean
    Dim blnEmergency As Boolean
    Dim blnBpCode As Boolean
    Dim intCMType As Integer
    Dim blnUseCancelCmd As Boolean
    Private blnShowFmt As Boolean
    Private strSetCitemCode As String
    Private strCancelCitemCode As String
    Private IsWaringResvTime As Boolean
    Private IsNSTVDevice As Boolean
Private Sub cmdCancel_Click()
  On Error Resume Next
    If pic1.Visible Then Exit Sub
    Unload Me
    Set frmSO1623A = Nothing
End Sub



Private Sub cmdCancelCMD_Click()
On Error GoTo ChkErr
    '#5891 �W�[�����w�����O By Kin 2011/03/25
    Dim strUpdSO555 As String
    Dim varBookMark As Variant
    Dim objCn As New ADODB.Connection
    
    If gilReturnCode.GetCodeNo = Empty Then
        MsgBox "�h���]��������!!", vbInformation, "�T��"
        gilReturnCode.SetFocus
        Exit Sub
    End If
    varBookMark = rsLog.Bookmark
    strUpdSO555 = "Update " & GetOwner & "SO555 " & _
                  " SET CMDSTATUS = 'E' , " & _
                  " CancelCode = " & gilReturnCode.GetCodeNo & " , " & _
                  " CANCELNAME ='" & gilReturnCode.GetDescription & "' " & _
                  " WHERE SEQNO = " & rsLog("SEQNO")
                
    objCn.ConnectionString = gcnGi.ConnectionString
    objCn.CursorLocation = 3
    objCn.Open
    objCn.BeginTrans
    objCn.Execute (strUpdSO555)
    objCn.CommitTrans
    rsLog.Requery
    Set ggrQuery.Recordset = rsLog
    ggrQuery.Refresh
    rsLog.AbsolutePosition = varBookMark
    On Error Resume Next
    objCn.Close
    Set objCn = Nothing
    Exit Sub
ChkErr:
    objCn.RollbackTrans
    ErrSub Me.Name, "cmdCancelCMD_Click"
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
    
    Dim lngRcdCnt As Long
    Dim strQry As String
    If blnEmptyGrid Then
        strQry = "SELECT A.*,B.ProductId PID,B.ProductName PName,C.CustName  FROM " & _
                GetOwner & "SO555 A," & GetOwner & "SO555A B," & _
                GetOwner & "SO001 C," & GetOwner & "SO004 D " & _
                " WHERE 0=1"
    Else
        '#5482 ���դ�OK,�d�߮ɥu�d�ثe���Ъ��]��,���n�������C�X�� By Kin 2010/01/27
        If Not blnPRflag Then
            If rsData Is Nothing Then Exit Sub
            With rsData
                If .State <= 0 Then Exit Sub
                '#6229 �d����]�Ƥ]�n��d�� By Kin 2012/03/28
                If .EOF Then
                    strQry = "Select A.*,B.ProductId PID,B.ProductName PName,C.CustName From " & _
                                GetOwner & "SO555 A," & GetOwner & "SO555A B," & GetOwner & "SO001 C, " & _
                                GetOwner & "SO004 D" & _
                                " Where A.SeqNo=B.SeqNo And A.AccountNum=" & ginCustId.Text & _
                                " And A.AccountNum=C.CustId And D.CustID=A.AccountNum " & _
                                " And A.SmartCardId=D.SmartCardNo  " & _
                                " And D.PrDate Is Null And A.SerialNumber=D.FaciSNo " & _
                                " Order by REQUESTTIME DESC"
                Else
                
                    strQry = "Select A.*,B.ProductId PID,B.ProductName PName,C.CustName From " & _
                            GetOwner & "SO555 A," & GetOwner & "SO555A B," & GetOwner & "SO001 C, " & _
                            GetOwner & "SO004 D" & _
                            " Where A.SeqNo=B.SeqNo And A.AccountNum=" & rsData("CustID") & _
                            " And A.AccountNum=C.CustId And D.CustID=A.AccountNum " & _
                            " And A.SmartCardId=D.SmartCardNo  " & _
                            " And D.PrDate Is Null And A.SerialNumber=D.FaciSNo " & _
                            " And D.FaciSNO='" & rsData("FaciSNo") & "'" & _
                            " Order by REQUESTTIME DESC"
                End If
                        
            End With
        Else
        
            If rsData.EOF Then
                '#6229 �d����]�Ƥ]�n��d�� By Kin 2012/03/28
                strQry = "Select A.*,B.ProductId PID,B.ProductName PName,C.CustName From " & _
                                GetOwner & "SO555 A," & GetOwner & "SO555A B," & GetOwner & "SO001 C " & _
                                " WHERE A.SEQNO=B.SEQNO AND ACCOUNTNUM=" & ginCustId.Text & _
                                " AND A.ACCOUNTNUM=C.CUSTID " & _
                                " ORDER BY REQUESTTIME DESC"
            Else
                 '#5469 �W�[�d�ߤw��]�� By Kin 2010/01/07
                strQry = "Select A.*,B.ProductId PID,B.ProductName PName,C.CustName From " & _
                            GetOwner & "SO555 A," & GetOwner & "SO555A B," & GetOwner & "SO001 C " & _
                            " WHERE A.SEQNO=B.SEQNO AND ACCOUNTNUM=" & rsData("CUSTID") & _
                            " AND A.ACCOUNTNUM=C.CUSTID AND A.SERIALNUMBER NOT IN(" & Get_Faci_Rows & ")" & _
                            " ORDER BY REQUESTTIME DESC"
            End If
        End If
    End If
    If Not GetRS(rsLog, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    GrdFmt rsLog
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
        .Add "SERIALNUMBER", , , , , "���W���Ǹ�" & Space(5), vbLeftJustify
        .Add "SMARTCARDID", , , , , "���z�d�Ǹ�" & Space(5), vbLeftJustify
        .Add "CUSTNAME", , , , , "�Ȥ�W��" & Space(10), vbLeftJustify
        .Add "PNAME", , , , , "�W�D�W��" & Space(20), vbLeftJustify
        .Add "CMDSTATUS", , , , , "�B�z���A" & Space(10), vbLeftJustify
        .Add "CANCELNAME", , , , , "����" & Space(30), vbLeftJustify
        .Add "CMD", , , , , "�]�w����" & Space(20), vbLeftJustify
        .Add "REQUESTTIME", giControlTypeTime, , , , "�]�w�ɶ�" & Space(15), vbLeftJustify
        .Add "UPDEN", , , , , "�]�w�H��" & Space(5), vbLeftJustify
        .Add "RESVTIME", giControlTypeTime, , , , "�w���ɶ�" & Space(20), vbLeftJustify
    End With
    
    With ggrQuery
            .AllFields = mFlds
            If rsLog.State = adStateOpen Then
                Set .Recordset = rsLog
            End If
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
    rsClone.Filter = "PRDATE=NULL"

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

Private Sub CtlEnabled()
  On Error GoTo ChkErr:
    Dim strSQLRef As String
    Dim ctl As Control

    Exit Sub
ChkErr:
  ErrSub Name, "CtlEnabled"
End Sub
'#7068 �s�W�d�߱b�����O By Kin 2015/08/10
Private Sub cmdQueryAccInfo_Click()
  On Error GoTo ErrTran
   Dim strNowDate  As String
   Dim strResult As String, strFaultReason As String, strErrMsg As String
   Dim rsVirtual As New ADODB.Recordset
   Me.Enabled = False
   Me.MousePointer = vbHourglass
   strNowDate = RightNow
   If Not QueryAccInfo(rsData, "C1", strNowDate, , rsData("CustId"), strCustName, _
                                          , , rsData("FaciSNO") & "", _
                                 , rsData("SmartCardNo") & "", _
                                , rsData("FaciSNO") & "", , , , , _
                                strNowDate, GaryGi(1), , , , , _
                                , , , , , , , , , , , _
                                rsVirtual, strResult, _
                                 strErrMsg, strFaultReason) Then GoTo ErrTran
                                
ErrTran:
        On Error Resume Next
        'rsData.AbsolutePosition = lngBR
        On Error GoTo ChkErr
        Call DoProcessMode
        blnSetting = False
        'If Not blnTransation Then gcnGi.RollbackTrans
        'OpenData
        'rsData.AbsolutePosition = lngBR
        pic1.Visible = False
        Me.Enabled = True
        Me.MousePointer = vbDefault
    Exit Sub
ChkErr:
    Me.Enabled = True
    Me.MousePointer = vbDefault
    ErrSub Me.Name, "cmdQueryAccInfo_Click"
                                
End Sub

Private Sub cmdSet_Click()
    On Error GoTo ChkErr
    Dim lngBR As Long
    Dim intLoop As Integer
    Dim intIndex As Integer
    Dim rsVirtual As New ADODB.Recordset
    Dim rsQrySVodCitem As New ADODB.Recordset
    Dim strResult As String, strFaultReason As String, strErrMsg As String
    Dim strNowDate As String
    Dim blnExeB12 As Boolean
    Dim Notes As String
    Dim rsOld As New ADODB.Recordset
    Dim rsOTT As New ADODB.Recordset
    Dim aCapacity As String
    Dim strProductID As String
    Dim strSO555BCitemCode As String
    blnProcessOk = False
    If pic1.Visible = True Then Exit Sub
    If Not IsDataOk Then Exit Sub
    Me.Enabled = False
    blnExeB12 = True
    blnSetting = True
'    If sstData.Tab = 0 Then
'        If optBasic(8).Value = True Then
'            Call cmdEmergency_Click
'            Exit Sub
'        End If
'    End If
    
    Call subProgressBar(True)
    On Error Resume Next
'    With ggrQueryInfo
'        If .Recordset.State > 0 Then
'            '.Blank = True
'            .ClearStructure
'        End If
'    End With
    strNowDate = RightNow
    On Error GoTo ErrTran
    If Not blnTransation Then gcnGi.BeginTrans
    lngBR = rsData.AbsolutePosition
    IsA6Command = False
    IsRestCommand = False
    OldVodAccountId = Empty
        Select Case sstData.Tab
            Case 0
                Select Case True
                     '�}��
                    Case optBasic(0)
                        '#6752 �W�[TunerCount�� By Kin 2014/04/07
                        '#6874 �O���ª�VODAccountId,���\��n�N�­Ȧ^��s�� By Kin 2014/09/30
                        OldVodAccountId = rsData("VodAccountId") & ""
                        If Not OpenVod(rsData, "A1", strNowDate, , rsData("CustId"), strCustName, _
                                         strInstAddress, strZipCode, rsData("FaciSNO") & "", _
                                "0000", rsData("SmartCardNo") & "", _
                                GetMacAddress(rsData("FaciSNO") & ""), rsData("FaciSNO") & "", , , , , _
                                strNowDate, GaryGi(1), , gdtResvTime.GetValue(True), str7SNO, str3MediaBillNo, _
                                strReasonCode, strReasonName, , , , , , , , GetTuner(rsData("FaciSNo"), rsData("SmartCardNo")), , _
                                rsVirtual, strResult, _
                                strErrMsg, strFaultReason) Then GoTo ErrTran
                        Call OpenData2
                    '���
                    '#5713 �W�[��JICCUID By Kin 2010/07/15
                    Case optBasic(1)
                        If Not OpenVod(rsData, "A2", strNowDate, , rsData("CustId"), , , , , _
                                    , rsData("SmartCardNo") & "", GetMacAddress(rsData("FaciSNO") & ""), _
                                    rsData("FaciSNO") & "" _
                                    , , rsData("VODAccountId") & "", , , strNowDate, GaryGi(1), , _
                                    gdtResvTime.GetValue(True), str7SNO, str3MediaBillNo, _
                                    strReason, strReasonName, , , GetICCUID(rsData("SMARTCARDNO") & "", rsData("COMPCODE")), _
                                    , , , , , , Nothing, strResult, strErrMsg, strFaultReason) Then GoTo ErrTran
                                    
                        Call OpenData2
                    '����
                    Case optBasic(2)
                        If Not OpenVod(rsData, "A3", strNowDate, , rsData("CustId"), , , , rsData("FaciSNO") & "", _
                                , rsData("SmartCardNo") & "", _
                                    , rsData("FaciSNo") & "", , rsData("VODAccountId") & "", , _
                                    , strNowDate, GaryGi(1), , _
                                    gdtResvTime.GetValue(True), str7SNO, str3MediaBillNo, _
                                    strReasonCode, strReasonName, , , , , , , , , , _
                                    Nothing, strResult, strErrMsg, strFaultReason) Then GoTo ErrTran

                        OpenData2
                    '�_��
                    Case optBasic(3)
                        If Not OpenVod(rsData, "A4", strNowDate, , rsData("CustId"), , , , rsData("FaciSNO") & "", _
                                    , rsData("SmartCardNo") & "", _
                                    , rsData("FaciSNO") & "", , rsData("VODAccountId") & "", , , strNowDate, , , _
                                    gdtResvTime.GetValue(True), str7SNO, str3MediaBillNo, _
                                    strReasonCode, strReasonName, , , , , , , , , , _
                                     Nothing, strResult, strErrMsg, strFaultReason) Then GoTo ErrTran
                        OpenData2
                    '���]�q�ʱK�X
                    Case optBasic(4)
                        If Not SetVodCmd(rsData, "E1", , , rsData("CustId"), , , , rsData("FaciSNO") & "", _
                                        "0000", rsData("SmartCardNo") & "" _
                                        , , rsData("FaciSNO") & "", , rsData("VODAccountId") & "", _
                                        , , , , , gdtResvTime.GetValue(True), str7SNO, str3MediaBillNo, _
                                        strReasonCode, strReasonName, , , , , , _
                                        Nothing, strResult, strErrMsg, strFaultReason) Then GoTo ErrTran
                        
                        OpenData2
                    Case optBasic(5)
                        '#5891 ���u�y�{�i�ӡA�p�G�S������q�ʸ��,�����Ǧ��\ By Kin 2011/03/25
                        If UCase(strProcessType) = "B11" Then
                            If Len(gimOrdSVOD.GetQueryCode & "") <= 0 Then
                                strResult = "1"
                                GoTo 66
                            End If
                        End If
                        If Not SetVodCmd(rsData, "B11", strNowDate, , rsData("CustId"), , , , , , _
                                         rsData("SmartCardNo") & "", GetMacAddress(rsData("FaciSNo") & ""), rsData("FaciSNO") & "" _
                                        , , rsData("VODAccountId") & "", GetChannelID(gimOrdSVOD.GetQryFld(1), rsData("CompCode"), rsData("CUSTID"), 0, rsData("VodAccountId") & ""), , _
                                        strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                        str3MediaBillNo, strReasonCode, _
                                        strReasonName, , , , gimOrdSVOD.GetQryFld(1), gimOrdSVOD.GetQryFld(1), _
                                        Nothing, strResult, strErrMsg, strFaultReason) Then GoTo ErrTran
                        gimOrdSVOD.Clear
                        OpenData2
                        
                    Case optBasic(6)
                        '#5891 ���u�y�{�i�ӡA�p�G�S������q�ʸ��,�����Ǧ��\ By Kin 2011/03/25
                        If (UCase(strProcessType) = "B2B1") Or (UCase(strProcessType) = "B12") Then
                            If strCancelCitemCode = "" Then
                                blnExeB12 = False
                            Else
                                If gimCancelSVOD.GetDispStr = "" Then
                                    blnExeB12 = False
                                End If
                                
                            End If
                        End If
                        If blnExeB12 Then
                            If Not SetVodCmd(rsData, "B12", strNowDate, , rsData("CustId"), , , , , , _
                                            rsData("SmartCardNo") & "", _
                                             GetMacAddress(rsData("FaciSNo") & ""), rsData("FaciSNO") & "" _
                                            , , rsData("VODAccountId") & "", _
                                             GetChannelID(gimCancelSVOD.GetQryFld(1), rsData("CompCode"), rsData("CUSTID"), 1, rsData("VodAccountId") & ""), , _
                                            strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                            str3MediaBillNo, strReasonCode, _
                                            strReasonName, , , , gimCancelSVOD.GetQryFld(1), strCancelCitemCode, Nothing, strResult, _
                                            strErrMsg, strFaultReason) Then GoTo ErrTran
                            gimCancelSVOD.Clear
                            OpenData2
                        Else
                            If UCase(strProcessType) = "B12" Then
                                strResult = "1"
                                GoTo 66
                            End If
                        End If
                         
                        If UCase(strProcessType) = "B2B1" Then
                            gimOrdSVOD.SetQueryCode strSetCitemCode
                            
                            If strSetCitemCode = "" Then
                                strResult = "1"
                                'strErrMsg = "�L�����W�D!"
                                'strResult = 0
                                'GoTo ErrTran
                                GoTo 66
                            Else
                                If gimOrdSVOD.GetDispStr = "" Then
                                    strResult = "1"
                                    GoTo 66
                                End If
                            End If
                            optBasic(5).Value = True
                            Call subProgressBar(True)
                            If Not SetVodCmd(rsData, "B11", strNowDate, , rsData("CustId"), , , , , , _
                                         rsData("SmartCardNo") & "", GetMacAddress(rsData("FaciSNo") & ""), rsData("FaciSNO") & "" _
                                        , , rsData("VODAccountId") & "", _
                                        GetChannelID(gimOrdSVOD.GetQryFld(1), rsData("CompCode"), rsData("CUSTID"), 0, rsData("VodaccountId") & ""), , _
                                        strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                        str3MediaBillNo, strReasonCode, _
                                        strReasonName, , , , gimOrdSVOD.GetQryFld(1), gimOrdSVOD.GetQryFld(1), Nothing, strResult, strErrMsg, strFaultReason) Then GoTo ErrTran
                            gimOrdSVOD.Clear
                            OpenData2
                        End If
                    '#5713 �W�[��JICCUID By Kin 2010/07/15
                    Case optBasic(7)
                        '�]��NSTV�M�קP�_�n�eA2��A1��A6 By Kin 2013/07/30
                        Dim OldFaciType As Integer
                        Dim NewFaciType As Integer
                        Set rsOld = rsData.Clone
                        Set rsOld.ActiveConnection = Nothing
                        rsOld.AbsolutePosition = rsData.AbsolutePosition
                        OldFaciType = GetDTVCAType(rsData("FaciSNo") & "", rsData("SmartCardNo") & "")
                        NewFaciType = GetDTVCAType(rs2("FaciSNo") & "", rs2("SmartCardNo") & "")
                        If OldFaciType <> NewFaciType Then
                            '#6573 �����2���R�O By Kin 2013/08/27
                            '#6573 �O���ª�VodAccountId By Kin 2013/09/03
                            IsA6Command = True
                            OldVodAccountId = rsData("VodAccountId")
                            strProductID = GetProductId(rsData("VodAccountId"))
                            strSO555BCitemCode = GetSO555BCitemCode(rsData("VodAccountId"))
                            '#6617 �p�G��SVOD�W�D,�n��������A�} By Kin 2013/10/01
                            If Len(strProductID) > 0 Then
                                If Not SetVodCmd(rsData, "B12", strNowDate, , rsData("CustId"), , , , , , _
                                                rsData("SmartCardNo") & "", _
                                                GetMacAddress(rsData("FaciSNo") & ""), rsData("FaciSNO") & "" _
                                                , , rsData("VODAccountId") & "", _
                                                 strProductID, , _
                                                strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                                str3MediaBillNo, strReasonCode, _
                                                strReasonName, , , , strSO555BCitemCode, strCancelCitemCode, Nothing, strResult, _
                                                strErrMsg, strFaultReason) Then
                                    If blnShowMsg Then
                                         ShowMsg "(1) A6   ���״���(��մ�)", strResult, strFaultReason, strErrMsg
                                    End If
                                    GoTo ErrTran
                                End If
                                gimCancelSVOD.Clear
                            End If
                            If Not OpenVod(rsData, "A2", strNowDate, , rsData("CustId"), , , , , _
                                       , rsData("SmartCardNo") & "", GetMacAddress(rsData("FaciSNO") & ""), _
                                           rsData("FaciSNO") & "" _
                                       , , rsData("VODAccountId") & "", , , strNowDate, GaryGi(1), , _
                                       gdtResvTime.GetValue(True), str7SNO, str3MediaBillNo, _
                                       strReason, strReasonName, , , GetICCUID(rsData("SMARTCARDNO") & "", rsData("COMPCODE")), , , , , , , _
                                       Nothing, strResult, strErrMsg, strFaultReason) Then
                                If blnShowMsg Then
                                    ShowMsg "(1) A6   ���״���(��մ�)", strResult, strFaultReason, strErrMsg
                                End If
                                GoTo ErrTran
                            End If
                            If Not OpenVod(rs2, "A1", strNowDate, , rs2("CustId"), strCustName, _
                                       strInstAddress, strZipCode, rs2("FaciSNO") & "", _
                                      "0000", rs2("SmartCardNo") & "", _
                                      GetMacAddress(rs2("FaciSNO") & ""), rs2("FaciSNO") & "", , , , , _
                                      strNowDate, GaryGi(1), , gdtResvTime.GetValue(True), str7SNO, str3MediaBillNo, _
                                      strReasonCode, strReasonName, , , , , , , , , , rsVirtual, strResult, _
                                      strErrMsg, strFaultReason) Then
                                If blnShowMsg Then
                                    ShowMsg "(1) A6   ���״���(��մ�)", strResult, strFaultReason, strErrMsg
                                End If
                                GoTo ErrTran
                            Else
                                'If Len(strProductID) = 0 Then ShowMsg "(1) A6   ���״���(��մ�)", strResult, strFaultReason, strErrMsg
                            End If
                            If Len(strProductID) > 0 Then
                                If Not SetVodCmd(rs2, "B11", strNowDate, , rs2("CustId"), , , , , , _
                                         rs2("SmartCardNo") & "", GetMacAddress(rs2("FaciSNo") & ""), rs2("FaciSNO") & "" _
                                        , , rs2("VODAccountId") & "", strProductID, , _
                                        strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                        str3MediaBillNo, strReasonCode, _
                                        strReasonName, , , , strSO555BCitemCode, strSO555BCitemCode, Nothing, strResult, strErrMsg, strFaultReason) Then
                                        GoTo ErrTran
                                    If blnShowMsg Then
                                        ShowMsg "(1) A6   ���״���(��մ�)", strResult, strFaultReason, strErrMsg
                                    End If
                                End If
'                                If blnShowMsg Then
'                                    ShowMsg "(1) A6   ���״���(��մ�)", strResult, strFaultReason, strErrMsg
'                                End If
                                gimOrdSVOD.Clear
                            End If
                            '#6752 �W�[Smit�R�O(Nagra��NSTV�]�Ʈ�,�n����B13�BB15�BB17) By Kin 2014/05/13
                            If OldFaciType = 0 And NewFaciType = 1 Then
                                Dim aSQL As String
                                Dim aCitemCode As String
                                '����B13
                                aSQL = GetRecordCitemCodeSQL(rsOld)
                                aCitemCode = GetMaxSizeCodeCitemCode(aSQL)
                                If Len(aCitemCode) > 0 Then
                                     aCapacity = GetRecordCapacity(aCitemCode)
                                     If aCapacity > 0 Then
                                         If Not SetSmitVodCmd(rs2, "B13", strNowDate, , rs2("CustId"), , , , , , _
                                                rs2("SmartCardNo") & "", _
                                                GetMacAddress(rs2("FaciSNo") & ""), rs2("FaciSNO") & "" _
                                                , , rs2("VODAccountId") & "", _
                                                  , , _
                                                strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                                str3MediaBillNo, strReasonCode, _
                                                strReasonName, , , , , , aCapacity, , aCitemCode, Nothing, strResult, _
                                                strErrMsg, strFaultReason) Then
                                            If blnShowMsg Then
                                                 ShowMsg "(1) A6   ���״���(��մ�)", strResult, strFaultReason, strErrMsg
                                            End If
                                            GoTo ErrTran
                                        End If
                                     End If
                                End If
                                '����B15
                                aSQL = GetSmitChargeCitemCodeSQL(rsOld, True, True, strSetCitemCode)
                                Notes = GetSmitProductId(aSQL, rsOld("CustId"), rsOld("FaciSNo"), rsOld("CompCode"), True, True)
                                If Len(Notes) > 0 Then
                                    If Not SetSmitVodCmd(rs2, "B15", strNowDate, , rs2("CustId"), , , , , , _
                                                    rs2("SmartCardNo") & "", _
                                                    GetMacAddress(rs2("FaciSNo") & ""), rs2("FaciSNO") & "" _
                                                    , , rs2("VODAccountId") & "", _
                                                      , , _
                                                    strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                                    str3MediaBillNo, strReasonCode, _
                                                    strReasonName, , , , , Notes, , , aCitemCode, Nothing, strResult, _
                                                    strErrMsg, strFaultReason) Then
                                        If blnShowMsg Then
                                            ShowMsg "(1) A6   ���״���(��մ�)", strResult, strFaultReason, strErrMsg
                                        End If
                                        GoTo ErrTran
                                    End If
                                End If
                                '����B17
                                If IsCatchUp(rsOld) Then
                                    If Not SetSmitVodCmd(rs2, "B17", strNowDate, , rs2("CustId"), , , , , , _
                                            rs2("SmartCardNo") & "", _
                                            GetMacAddress(rs2("FaciSNo") & ""), rs2("FaciSNO") & "" _
                                            , , rs2("VODAccountId") & "", _
                                              , , _
                                            strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                            str3MediaBillNo, strReasonCode, _
                                            strReasonName, , , , , , , , aCitemCode, Nothing, strResult, _
                                            strErrMsg, strFaultReason) Then
                                        If blnShowMsg Then
                                            ShowMsg "(1) A6   ���״���(��մ�)", strResult, strFaultReason, strErrMsg
                                        End If
                                        GoTo ErrTran
                                    End If
                                End If
                                
                            End If
                             strResult = "1"
                            If blnShowMsg Then
                                ShowMsg "(1) A6   ���״���(��մ�)", strResult, strFaultReason, strErrMsg
                            End If
                        Else
                            '"A6"
                            If Not OpenVod(rsData, "A6", strNowDate, , rsData("CustId"), , , , , , _
                                  rsData("SmartCardNo") & "" & "," & rs2("SmartCardNo") & "", _
                                GetMacAddress(rsData("FaciSNO") & "") & "," & GetMacAddress(rs2("FaciSNO") & ""), _
                                  rs2("FaciSNO") & "", _
                                , rsData("VODAccountId") & "", GetProductId(rsData("VODAccountId")) _
                                , , _
                                strNowDate, , , gdtResvTime.GetValue(True), _
                                str7SNO, str3MediaBillNo, strReasonCode, _
                                strReasonName, , , GetICCUID(rsData("SMARTCARDNO") & "", rsData("COMPCODE")), _
                                rsData("FaciSNo"), rsData("SmartCardNo") & "", , , , , rsVirtual, strResult, _
                                strErrMsg, strFaultReason) Then GoTo ErrTran
                                
                        End If
                                            
                        OpenData2
                    Case optBasic(8)
                        If Not OpenVod(rsData, "A8", strNowDate, , rsData("CustId"), , , _
                                , , , rs2("SmartCardNo") & "", _
                                GetMacAddress(rsData("FaciSNO") & "") & "," & GetMacAddress(rs2("FaciSNO") & ""), _
                                rsData("FaciSNo") & "", _
                                , rsData("VODAccountId") & "", , _
                                , strNowDate, , , gdtResvTime.GetValue(True), _
                                str7SNO, str3MediaBillNo, strReasonCode, _
                                strReasonName, , , , , , , , , , Nothing, strResult, _
                                strErrMsg, strFaultReason) Then GoTo ErrTran
                        OpenData2
                    Case optBasic(9)
                        If Not OpenVod(rsData, "A9", strNowDate, , rsData("CustId"), , , _
                                , , , rsData("SmartCardNo") & "", _
                                GetMacAddress(rsData("FaciSNO") & ""), rsData("FaciSNo") & "", _
                                , rsData("VODAccountId") & "", , , strNowDate, _
                                , , gdtResvTime.GetValue(True), str7SNO, str3MediaBillNo, _
                                strReasonCode, strReasonName, , , , , , , , , , _
                                Nothing, strResult, strErrMsg, strFaultReason) Then GoTo ErrTran
                    '#7068 �W�[���m�R�O,���eA2�A�eA1 By Kin 2015/08/06
                    Case optBasic(10)
                        '���m
                        IsRestCommand = True
                        If Len(strCancelCitemCode & "") > 0 Then
                            strProductID = GetChannelID(gimCancelSVOD.GetQryFld(1), rsData("CompCode"), rsData("CUSTID"), 1, rsData("VodAccountId") & "")
                            strSO555BCitemCode = strCancelCitemCode
                        Else
                            strProductID = GetProductId(rsData("VodAccountId"))
                            strSO555BCitemCode = GetSO555BCitemCode(rsData("VodAccountId"))
                        End If
                        If Len(strProductID) > 0 Then
                            If Not SetVodCmd(rsData, "B12", strNowDate, , rsData("CustId"), , , , , , _
                                            rsData("SmartCardNo") & "", _
                                            GetMacAddress(rsData("FaciSNo") & ""), rsData("FaciSNO") & "" _
                                            , , rsData("VODAccountId") & "", _
                                             strProductID, , _
                                            strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                            str3MediaBillNo, strReasonCode, _
                                            strReasonName, , , , strSO555BCitemCode, strSO555BCitemCode, Nothing, strResult, _
                                            strErrMsg, strFaultReason) Then
                                If blnShowMsg Then
                                     ShowMsg "���m", strResult, strFaultReason, strErrMsg
                                End If
                                GoTo ErrTran
                            End If
                            gimCancelSVOD.Clear
                        End If
                        If Not OpenVod(rsData, "A2", strNowDate, , rsData("CustId"), , , , , _
                                    , rsData("SmartCardNo") & "", GetMacAddress(rsData("FaciSNO") & ""), _
                                    rsData("FaciSNO") & "" _
                                    , , rsData("VODAccountId") & "", , , strNowDate, GaryGi(1), , _
                                    gdtResvTime.GetValue(True), str7SNO, str3MediaBillNo, _
                                    strReason, strReasonName, , , GetICCUID(rsData("SMARTCARDNO") & "", rsData("COMPCODE")), _
                                    , , , , , , Nothing, strResult, strErrMsg, strFaultReason) Then
                                    If blnShowMsg Then
                                        ShowMsg "���m", strResult, strFaultReason, strErrMsg
                                    End If
                                    GoTo ErrTran
                        End If
                        OldVodAccountId = rsData("VodAccountId") & ""
                        If Not OpenVod(rsData, "A1", strNowDate, , rsData("CustId"), strCustName, _
                                         strInstAddress, strZipCode, rsData("FaciSNO") & "", _
                                "0000", rsData("SmartCardNo") & "", _
                                GetMacAddress(rsData("FaciSNO") & ""), rsData("FaciSNO") & "", , , , , _
                                strNowDate, GaryGi(1), , gdtResvTime.GetValue(True), str7SNO, str3MediaBillNo, _
                                strReasonCode, strReasonName, , , , , , , , GetTuner(rsData("FaciSNo"), rsData("SmartCardNo")), , _
                                rsVirtual, strResult, _
                                strErrMsg, strFaultReason) Then
                               
                                If blnShowMsg Then
                                    ShowMsg "���m", strResult, strFaultReason, strErrMsg
                                End If
                                GoTo ErrTran
                        End If
                        
                        If Len(strProductID) > 0 Then
                            If Not SetVodCmd(rsData, "B11", strNowDate, , rsData("CustId"), , , , , , _
                                       rsData("SmartCardNo") & "", GetMacAddress(rsData("FaciSNo") & ""), rsData("FaciSNO") & "" _
                                      , , rsData("VODAccountId") & "", _
                                      strProductID, , _
                                      strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                      str3MediaBillNo, strReasonCode, _
                                      strReasonName, , , , strSO555BCitemCode, strSO555BCitemCode, _
                                      Nothing, strResult, strErrMsg, strFaultReason) Then
                                    If blnShowMsg Then
                                        ShowMsg "���m", strResult, strFaultReason, strErrMsg
                                    End If
                                    
                                    GoTo ErrTran
                            End If
                        End If
                         If blnShowMsg Then
                             ShowMsg "���m", strResult, strFaultReason, strErrMsg
                         End If
                        Call OpenData2
                                    
                    Case Else
                        MsgBox "�п�ܩR�O�I", vbInformation, "�T��"
                    
                End Select
            '#6752 �W�[Smit�R�O By Kin 2014/04/07
            Case 1
                If Len(strProcessType) > 0 Then
                    If GetDTVCAType(rsData("FaciSNo") & "", rsData("SmartCardNo") & "") <> 1 Then
                         strResult = "1"
                        GoTo 66
                    End If
                End If
                Select Case True
                    Case optSMITCommand(0)
                       
                        If Len(gilCitemCode.GetCodeNo) > 0 Then
                            aCapacity = GetRecordCapacity(gilCitemCode.GetCodeNo)
                        Else
                            aCapacity = "0"
                        End If
                        If aCapacity = "0" And UCase(strProcessType) = "B13" Then
                            strResult = "1"
                            GoTo 66
                        End If
                                            
                        If Not SetSmitVodCmd(rsData, "B13", strNowDate, , rsData("CustId"), , , , , , _
                                                rsData("SmartCardNo") & "", _
                                                GetMacAddress(rsData("FaciSNo") & ""), rsData("FaciSNO") & "" _
                                                , , rsData("VODAccountId") & "", _
                                                  , , _
                                                strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                                str3MediaBillNo, strReasonCode, _
                                                strReasonName, , , , , , aCapacity, , gilCitemCode.GetCodeNo, Nothing, strResult, _
                                                strErrMsg, strFaultReason) Then
                                    GoTo ErrTran
                        End If
                        OpenData2
                    Case optSMITCommand(1)
                        If (UCase(strProcessType) = "B14") And (Len(strCancelCitemCode) = 0) Then
                            strResult = "1"
                            GoTo 66
                        Else
                            If (UCase(strProcessType) = "B14") Then
                                Dim maxSizeCitemCode As String
                                maxSizeCitemCode = GetMaxSizeCodeCitemCode(strCancelCitemCode)
                                If Len(maxSizeCitemCode) = 0 Then
                                    strResult = "1"
                                    GoTo 66
                                End If
                            End If
                        End If
                        If Not SetSmitVodCmd(rsData, "B14", strNowDate, , rsData("CustId"), , , , , , _
                                                rsData("SmartCardNo") & "", _
                                                GetMacAddress(rsData("FaciSNo") & ""), rsData("FaciSNO") & "" _
                                                , , rsData("VODAccountId") & "", _
                                                  , , _
                                                strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                                str3MediaBillNo, strReasonCode, _
                                                strReasonName, , , , , , 0, , strCancelCitemCode, Nothing, strResult, _
                                                strErrMsg, strFaultReason) Then
                                    GoTo ErrTran
                        End If
                        OpenData2
                    Case optSMITCommand(2)
                        If (strProcessType <> "") And (Len(gimOrdSmitProd.GetQueryCode & "") = 0) Then
                            strResult = "1"
                            GoTo 66
                        End If
                        
                        Notes = GetSmitProductId(gimOrdSmitProd.GetQueryCode, rsData("CustId"), rsData("FaciSNo"), rsData("CompCode"), True, True)
                        If Len(Notes) = 0 Then
                            If UCase(strProcessType) = "B15" Then
                                strResult = "1"
                                GoTo 66
                            End If
                            If blnShowMsg Then
                                MsgBox "�L�W�D�i�}�I", vbInformation, "�T��"
                            End If
                            strResult = "1"
                            GoTo 66
                        End If
                        If Not SetSmitVodCmd(rsData, "B15", strNowDate, , rsData("CustId"), , , , , , _
                                                rsData("SmartCardNo") & "", _
                                                GetMacAddress(rsData("FaciSNo") & ""), rsData("FaciSNO") & "" _
                                                , , rsData("VODAccountId") & "", _
                                                  , , _
                                                strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                                str3MediaBillNo, strReasonCode, _
                                                strReasonName, , , , , Notes, , , gimOrdSmitProd.GetQueryCode, Nothing, strResult, _
                                                strErrMsg, strFaultReason) Then
                                    GoTo ErrTran
                        End If
                        OpenData2
                    Case optSMITCommand(3)
                        If (strProcessType <> "") And (Len(gimCancelSmitProd.GetQueryCode & "") = 0) Then
                            strResult = "1"
                            GoTo 66
                        End If
                        
                        Notes = GetSmitProductId(gimCancelSmitProd.GetQueryCode, rsData("CustId"), rsData("FaciSNo"), rsData("CompCode"), True, True)
                         If Len(Notes) = 0 Then
                            If UCase(strProcessType) = "B16" Then
                                strResult = "1"
                                GoTo 66
                            End If
                            If blnShowMsg Then
                                MsgBox "�L�W�D�i���I", vbInformation, "�T��"
                            End If
                            strResult = "1"
                            GoTo 66
                        End If
                        If Not SetSmitVodCmd(rsData, "B16", strNowDate, , rsData("CustId"), , , , , , _
                                                rsData("SmartCardNo") & "", _
                                                GetMacAddress(rsData("FaciSNo") & ""), rsData("FaciSNO") & "" _
                                                , , rsData("VODAccountId") & "", _
                                                  , , _
                                                strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                                str3MediaBillNo, strReasonCode, _
                                                strReasonName, , , , , Notes, , , gimCancelSmitProd.GetQueryCode, Nothing, strResult, _
                                                strErrMsg, strFaultReason) Then
                                    GoTo ErrTran
                        End If
                        OpenData2
                    Case optSMITCommand(4)
                        
                        If UCase(strProcessType) = "B17" Then
                            If Len(strSetCitemCode) = 0 Then
                                strResult = "1"
                                GoTo 66
                            Else
                                If Not IsCatchUpCitemCode(strSetCitemCode) Then
                                    strResult = "1"
                                    GoTo 66
                                End If
                            End If
                        Else
                            If Not IsCatchUp(rsData) Then
                                strResult = "1"
                                GoTo 66
                            End If
                        End If
                       
                        If Not SetSmitVodCmd(rsData, "B17", strNowDate, , rsData("CustId"), , , , , , _
                                                rsData("SmartCardNo") & "", _
                                                GetMacAddress(rsData("FaciSNo") & ""), rsData("FaciSNO") & "" _
                                                , , rsData("VODAccountId") & "", _
                                                  , , _
                                                strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                                str3MediaBillNo, strReasonCode, _
                                                strReasonName, , , , , , , , strSetCitemCode, Nothing, strResult, _
                                                strErrMsg, strFaultReason) Then
                                    GoTo ErrTran
                        End If
                        OpenData2
                    Case optSMITCommand(5)
                        If UCase(strProcessType) = "B18" Then
                            If Len(strCancelCitemCode) = 0 Then
                                strResult = "1"
                                GoTo 66
                            Else
                                If Not IsCatchUpCitemCode(strCancelCitemCode) Then
                                    strResult = "1"
                                    GoTo 66
                                End If
                            End If
                        Else
                            If Not IsCatchUp(rsData) Then
                                strResult = "1"
                                GoTo 66
                            End If
                        End If
                        If Not SetSmitVodCmd(rsData, "B18", strNowDate, , rsData("CustId"), , , , , , _
                                                rsData("SmartCardNo") & "", _
                                                GetMacAddress(rsData("FaciSNo") & ""), rsData("FaciSNO") & "" _
                                                , , rsData("VODAccountId") & "", _
                                                  , , _
                                                strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                                str3MediaBillNo, strReasonCode, _
                                                strReasonName, , , , , , , , strCancelCitemCode, Nothing, strResult, _
                                                strErrMsg, strFaultReason) Then
                                    GoTo ErrTran
                        End If
                        OpenData2
                     Case optSMITCommand(6)
                        If Not SetSmitVodCmd(rsData, "A7", strNowDate, , rsData("CustId"), , , , , , _
                                                rsData("SmartCardNo") & "", _
                                                GetMacAddress(rsData("FaciSNo") & ""), rsData("FaciSNO") & "" _
                                                , , rsData("VODAccountId") & "", _
                                                  , , _
                                                strNowDate, , , gdtResvTime.GetValue(True), str7SNO, _
                                                str3MediaBillNo, strReasonCode, _
                                                strReasonName, , , , , , , GetTuner(rsData("FaciSNo") & "", rsData("SmartCardNo") & ""), strSetCitemCode, _
                                                Nothing, strResult, _
                                                strErrMsg, strFaultReason) Then
                                    GoTo ErrTran
                        End If
                        OpenData2
                End Select
         
            Case 2
                If optOTTQuery(0).Value Then
                    If Len(gilOTTCitem.GetCodeNo & "") = 0 Then
                        MsgBox "�п�ܦ��O���ءI", vbInformation, "�T��"
                        strResult = "1"
                        GoTo 66
                    End If
                    Dim strOTTUserID  As String
                    
                    Dim aOTTSQL As String
                    aOTTSQL = "Select A.OTTUserID,B.iOSOrderNo,B.FVOrderNo From " & strOTTOwner & "SO220 A," & strOTTOwner & "SO003 B  Where " & _
                                        " A.OTTUserID In (Select OTTUserID From " & strOTTOwner & "SO222 Where  SOCompID = " & rsData("COMPCODE") & _
                                        " And FaciID = '" & rsData("FaciID") & "' )  " & _
                                        " AND A.CUSTID = B.CUSTID " & _
                                        " AND B.FVORDERNO IS NOT NULL " & _
                                        " AND B.CitemCode = " & gilOTTCitem.GetCodeNo
                    
                    GetRS rsOTT, aOTTSQL, gcnGi, adUseClient, adOpenKeyset
                    If rsOTT.RecordCount = 0 Then
                        MsgBox "�䤣�즬�O���ج����q�ʸ��", vbInformation, "�T��"
                        strResult = "1"
                        GoTo 66
                    End If
                    If Len(rsOTT("OTTUserID") & "") = 0 Then
                        MsgBox "�䤣��OTTUserID", vbInformation, "�T��"
                        strResult = "1"
                        GoTo 66
                    End If
                    If Len(rsOTT("iOSOrderNo") & "") = 0 Then
                        MsgBox "�䤣��iOS�q��渹�q�ʸ��", vbInformation, "�T��"
                        strResult = "1"
                        GoTo 66
                    End If
                    If Len(rsOTT("FVOrderNo") & "") = 0 Then
                        MsgBox "�䤣��iFonsView�q��渹", vbInformation, "�T��"
                        strResult = "1"
                        GoTo 66
                    End If
                    If Not QueryOTT(rsData, "C2", , , rsData("CustId"), , , , rsData("FaciSNO") & "", _
                                         , rsData("SmartCardNo") & "" _
                                        , , rsData("FaciSNO") & "", , rsData("VODAccountId") & "", _
                                        , , , , , gdtResvTime.GetValue(True), str7SNO, str3MediaBillNo, _
                                        , , , , , , , _
                                        rsOTT("OTTUserID") & "", rsOTT("iOSOrderNo") & "", rsOTT("FVOrderNo") & "", Nothing, _
                                        strResult, strErrMsg, strFaultReason) Then
                                        GoTo ErrTran
                                                
                                                                        
                    End If
                    
                Else
                    MsgBox "�п�ܩR�O�I", vbInformation, "�T��"
                    strResult = "1"
                    GoTo 66
                End If
                
            Case 3

        End Select
66:
        IsWaringResvTime = True
        blnSetting = False
        If Not blnTransation Then gcnGi.CommitTrans
        rsData.AbsolutePosition = lngBR
        
        blnProcessOk = True
        Call subProgressBar(False)
        Me.Enabled = True
        If (blnReturnCmd) Then
            blnReturnCmd = False
            If optBasic(0).Value Then optBasic(1).Value = True: cmdSet.Value = True: Exit Sub
            If optBasic(1).Value Then optBasic(0).Value = True: cmdSet.Value = True: Exit Sub
            Exit Sub
        End If
        
        Call EnableMode(True)
        'Show_Rs_Result rsVirtual
        On Error Resume Next
        Call CloseRecordset(rsOld)
        Call CloseRecordset(rsQrySVodCitem)
        Call CloseRecordset(rsOTT)
        If Len(strProcessType) > 0 Then Unload Me
    Exit Sub
ErrTran:
        On Error Resume Next
        rsData.AbsolutePosition = lngBR
        On Error GoTo ChkErr
        Call DoProcessMode
        blnSetting = False
        If Not blnTransation Then gcnGi.RollbackTrans
        OpenData
        rsData.AbsolutePosition = lngBR
        pic1.Visible = False
        Me.Enabled = True
        If (blnReturnCmd) Then
            If optBasic(0).Value Then
                optBasic(1).Value = True
                cmdSet.Value = True
                Exit Sub
            End If
            If optBasic(1).Value Then
                optBasic(0).Value = True
                cmdSet.Value = True
                Exit Sub
            End If
            
        End If
        Call CloseRecordset(rsOld)
        Call CloseRecordset(rsOTT)
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdSet_Click"
End Sub
'Private Function RunA6SmitCmd(ByRef rsOld As ADODB.Recordset) As Boolean
'  On Error GoTo ChkErr
'    Dim aSQL As String
'    Dim aCitemCode As String
'    aSQL = GetSmitChargeCitemCodeSQL(rsOld, True, True)
'    aCitemCode = GetMaxSizeCodeCitemCode(aSQL)
'    If Len(aCitemCode) > 0 Then
'
'    End If
'  Exit Function
'ChkErr:
'  Call ErrSub(Me.Name, "RunA6SmitCmd")
'End Function
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

'Private Sub Show_Rs_Result(ByRef rsResult As ADODB.Recordset)
'  On Error GoTo ChkErr
'    Dim mFlds As New GiGridFlds
'    Dim intLoop As Long
'    ggrQueryInfo.Blank = True
'    If rsResult.State = adStateClosed Then Exit Sub
'    With rsResult
'        If .RecordCount = 0 Then Exit Sub
'        .MoveFirst
'        For intLoop = 0 To .Fields.Count - 1
'            SetGrdFld mFlds, .Fields(intLoop).Name
'        Next intLoop
'    End With
'    With ggrQueryInfo
'        .AllFields = mFlds
'        Set .Recordset = rsResult
'        .Refresh
'        .Blank = False
'    End With
'
'  Exit Sub
'ChkErr:
'    ErrSub Name, "Show_Rs_Result"
'End Sub

'Private Sub Show_XML_Result(ByRef rsResult As ADODB.Recordset)
'  On Error GoTo ChkErr
'
'    Dim mFlds As New GiGridFlds
'    Dim intLoop As Integer
'
'    ggrQueryInfo.Blank = True
'    If rsResult.State = adStateClosed Then Exit Sub
'    With rsResult
'            If .RecordCount = 0 Then Exit Sub
'            .MoveFirst
'            For intLoop = 0 To .Fields.Count - 1
'                SetGrdFld mFlds, .Fields(intLoop).Name
'            Next
'    End With
'
'    With ggrQueryInfo
'            .AllFields = mFlds
'            .SetHead
'            Set .Recordset = rsResult
'            .Refresh
'            '.Blank = False
'    End With
'
'  Exit Sub
'ChkErr:
'    ErrSub Name, "Show_XML_Result"
'End Sub
Private Sub SetGrdFld(mFlds As GiGridFlds, strFldName As String)
  On Error GoTo ChkErr
  
    With mFlds
        
        Select Case UCase(strFldName)
            Case "CUSTID", "USERIDNO", "USERNAME", "ACCOUNTID", "CUSTNAME", "CMTS IP", "CM STATUS"
                .Add strFldName, , , , , strFldName & Space(10), vbLeftJustify
            Case "CUSTADDRESS"
                .Add strFldName, , , , , strFldName & Space(20), vbLeftJustify
            Case Else
                .Add strFldName, , , , , strFldName, vbLeftJustify
        End Select

    End With
  
  
  Exit Sub
ChkErr:
    ErrSub Name, "SetGrdFld"
End Sub
Private Sub subProgressBar(blnFlag As Boolean)
    On Error Resume Next
    Dim intLoop As Integer
    Dim strDoCaption As String
    Dim objControl As Object
    Dim intLength As Integer
        intLength = 4
        pbr1.Value = 0
        pic1.Visible = blnFlag
        If Not blnFlag Then Exit Sub
        Select Case sstData.Tab
            Case 0
                Set objControl = optBasic
            Case 1
                Set objControl = optSMITCommand
            Case 2
                Set objControl = optOTTQuery
            Case 3
                Set objControl = optCMDevice
        End Select
        For intLoop = 0 To objControl.UBound
            If objControl(intLoop) Then
                If intLoop = 5 Or intLoop = 6 Then intLength = 3
                strDoCaption = Mid(objControl(intLoop).Caption, intLength)
            End If
        Next
        lblProcess.Caption = strDoCaption & " �� , �еy�� !!"

End Sub

Private Sub Form_Activate()
  On Error Resume Next
'    '#3778 �۰���ܳ]�Ƹ�� By Kin 2008/02/25
'    If blnShowFaci Then
'        blnSendKey = False
'        sstHead.Tab = 1
'        If cmdQuery.Enabled Then cmdQuery.Value = True
'    End If
'    blnShowFmt = True
'  '#3778�P�_�O�_�n�۰ʫ��UF2 By Kin 2008/02/22
    If blnSendKey Then
        If cmdSet.Enabled Then cmdSet.Value = True
    End If
    blnSendKey = False
'    blnShowFaci = False
End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr
        blnShowFmt = False
        gdtResvTime.ShowSecond
        blnExeCommand = True '����R�O���P�_(��UI���@�w�n����A�ҥH�w�]��True
        sstData.Tab = 0
        blnReturnCmd = False
        blnHasRunReturn = False
        intRunRetunCount = 0
        '*******************************************************
        sstData.TabVisible(1) = True
        sstData.TabVisible(2) = False
        sstData.TabVisible(3) = False
        
        'gimOrdSVOD.Enabled = False
        'gimCancelSVOD.Enabled = False
        '*******************************************************
        
        '#7031 �W�[�O�_�n�L�oB15�R�O��� By Kin 2015/04/28
        Call GetClctStopDateType
        Call subGrd
        Call subGil
        Call OpenData
        Call DefaultValue
        
        
        Call DoProcessMode
        
        Fetch_Log_Data , True
        If blnAutoSend Then
            blnShowMsg = False
        Else
            blnShowMsg = True
        End If
        blnRecordProcedure = IsRecordProcedure  '�P�_�O�_�n�O�������{��
        
'        If Not rsData.EOF Then
'            SetgiList gilReason, "CodeNo", "Description", "CD014", , , , , 3, 20, " Where ServiceType is Null Or ServiceType='" & rsData("ServiceType") & "'", True
'
'            SetgiList gilReturnCode, "CodeNo", "Description", "CD015", , , , , 3, 20, " Where (ServiceType is Null Or ServiceType='" & rsData("ServiceType") & "') And RefNo=3", True
'        Else
'            'MsgBox "�L�����ƥi����!!", vbCritical, "�T��"
'        End If
        'blnShowResvTimeWaring = True
        If Len(strResvTime) > 0 Then ShowResv
        gdtResvTime.SetValue strResvTime
        If Not rsData.EOF Then
'            Dim aWhere As String
'            If strProcessType = "" Then
'                aWhere = " AND A.CITEMCODE IN(SELECT CITEMCODE FROM " & GetOwner & "SO003 " & _
'                                    " WHERE FACISEQNO='" & rsData("SEQNO") & "' AND CUSTID=" & rsData("CUSTID") & _
'                                    " UNION  SELECT CITEMCODE FROM " & GetOwner & "SO033 " & _
'                                    " WHERE FACISEQNO='" & rsData("SEQNO") & "' AND CUSTID=" & rsData("CUSTID") & ")"
'
'            Else
'                aWhere = ""
'            End If
'            '�u�nSO033��SO003���Ӧ��O���شN��ܥX��,�p�G�S���N���n��� By Kin 2010/06/15
'            SetMsQry gimOrdSVOD, "Select A.CitemCode,B.Description,C.CodeNo,C.ChannelID From " & GetOwner & " CD019A A" & _
'                                "," & GetOwner & "CD019 B," & GetOwner & "CD024 C Where " & _
'                                " A.CitemCode = B.CodeNo And A.CodeNo = C.CodeNo And C.VODType = 2" & _
'                                " And C.StopFlag<>1 And C.CodeNo Not In(Select Distinct ProductID From " & _
'                                GetOwner & "SO555B Where CustId=" & rsData("CustId") & " AND STATUS = 1) " & _
'                                aWhere, _
'                                , , "���O���إN�X,���O���ئW��", "1200,4200,0,0"
            Call SetgimOrdVOD
'            SetMsQry gimCancelSVOD, "Select A.CitemCode,B.Description,A.ProductID," & _
'                                " C.ChannelID,A.RowId " & _
'                                " From " & GetOwner & "SO555B A" & _
'                                "," & GetOwner & "CD019 B," & GetOwner & "CD024 C" & _
'                                " Where A.CitemCode=B.CodeNo And A.CustId=" & rsData("CustId") & _
'                                " AND A.ProductID=C.CodeNo And C.VodType=2 And C.StopFlag<>1 " & _
'                                " AND C.CODENO NOT IN (SELECT DISTINCT PRODUCTID FROM " & _
'                                GetOwner & "SO555B WHERE CUSTID=" & rsData("CUSTID") & " AND STATUS = 0)", _
'                                , , "���O���إN�X,���O���ئW��", "1200,4200,0,0,0"
            
            Call SetgimCancelVod(strCancelCitemCode)
            If strSetCitemCode <> "" Or strCancelCitemCode <> "" Then
                If optBasic(5).Value Then
                    gimOrdSVOD.SetQueryCode strSetCitemCode
                End If
                If optBasic(6).Value Then
                    gimCancelSVOD.SetQueryCode strCancelCitemCode
                End If
            End If
            
            If strProcessType <> "" Then
                gimCancelSVOD.Enabled = False
                gimOrdSVOD.Enabled = False
                '#6752 �W�[Smit �i���R�O By Kin 2014/04/02
'                gimOrdSmitProd.Enabled = False
'                gimCancelSmitProd.Enabled = False
'                gimOrdCatchUpProd.Enabled = False
'                gimCancelCatchUpProd.Enabled = False
            End If
            
            
            
            '#5891 �W�[�h���] By Kin 2011/03/25
            SetgiList gilReturnCode, "CodeNo", "Description", "CD015", , , , , 3, 20, " Where (ServiceType is Null Or ServiceType='" & rsData("ServiceType") & "') And RefNo=3", True
            '#5891 �q�y�{�i�ӭn�۰ʳ]�w By Kin 2011/03/30
            If (strProcessType = "B11" Or strProcessType = "B12" Or strProcessType = "B2B1") And (Not blnAutoSend) Then
                blnSendKey = True
                'cmdSet.Value = True
            End If
        End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Function SetgimOrdVOD() As String
  On Error GoTo ChkErr
  '#6317 �ư�SO033.ChEven=1�����O���� By Kin 2012/11/13
    If Not rsData.EOF Then
        Dim aWhere As String
        Dim aSQL As String
        If strProcessType = "" Then
            aWhere = " AND A.CITEMCODE IN(SELECT CITEMCODE FROM " & GetOwner & "SO003 " & _
                                " WHERE FACISEQNO='" & rsData("SEQNO") & "' AND CUSTID=" & rsData("CUSTID") & _
                                " UNION  SELECT CITEMCODE FROM " & GetOwner & "SO033 " & _
                                " WHERE FACISEQNO='" & rsData("SEQNO") & "' AND CUSTID=" & rsData("CUSTID") & " AND ChEven <> 1 )"
            
        Else
            aWhere = ""
        End If
        
        '�u�nSO033��SO003���Ӧ��O���شN��ܥX��,�p�G�S���N���n��� By Kin 2010/06/15
        
        aSQL = "Select A.CitemCode,B.Description,C.CodeNo,C.ChannelID From " & GetOwner & " CD019A A" & _
                            "," & GetOwner & "CD019 B," & GetOwner & "CD024 C Where " & _
                            " A.CitemCode = B.CodeNo And A.CodeNo = C.CodeNo And C.VODType = 2" & _
                            " And C.StopFlag<>1 And C.CodeNo Not In(Select Distinct ProductID From " & _
                            GetOwner & "SO555B Where STATUS = 1 " & _
                            " AND SO555B.VODACCOUNTID = " & rsData("VODACCOUNTID") & " ) " & aWhere
        
        'aSQL = "select * from cd019a where rowid in ('AAANmGAAHAAAADFAAA','AAANmGAAHAAAADFAAB')"
        SetMsQry gimOrdSVOD, aSQL, _
                            , , "���O���إN�X,���O���ئW��", "1200,4200,0,0"
    End If
    SetgimOrdVOD = aSQL
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "SetgimOrdVOD")
End Function
Private Function SetgimCancelVod(Optional aCitemCodes As String = Empty) As String
  On Error GoTo ChkErr
    Dim aSQL As String
    aSQL = "Select A.CitemCode,B.Description,A.ProductID," & _
                                " C.ChannelID,A.RowId " & _
                                " From " & GetOwner & "SO555B A" & _
                                "," & GetOwner & "CD019 B," & GetOwner & "CD024 C" & _
                                " Where A.CitemCode=B.CodeNo " & _
                                " AND A.ProductID=C.CodeNo And C.VodType=2 And C.StopFlag<>1 " & _
                                " AND C.CODENO NOT IN (SELECT DISTINCT PRODUCTID FROM " & _
                                GetOwner & "SO555B WHERE " & _
                                "  STATUS = 0  " & _
                                " AND VODACCOUNTID=" & rsData("VODACCOUNTID") & _
                                " ) AND A.VODACCOUNTID = " & rsData("VODACCOUNTID")
   
    
    SetMsQry gimCancelSVOD, aSQL, _
                                , , "���O���إN�X,���O���ئW��", "1200,4200,0,0,0"
    SetgimCancelVod = aSQL
    Exit Function
ChkErr:
 Call ErrSub(Me.Name, "SetgimCancelVod")
End Function
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
                
        If strDefaultRowId <> "" Then RsFind rsData, "RowId", strDefaultRowId
        If Len(strProcessType) <= 0 Then Exit Sub
        For Each objControl In Me
            If TypeOf objControl Is OptionButton Then
                objControl.Enabled = False
            End If
            If TypeOf objControl Is CSmulti Then
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
       
        sstData.TabEnabled(3) = False
        If (InStr(1, UCase(strProcessType), "A") > 0) Or (InStr(1, UCase(strProcessType), "B") > 0) Then
            If strProcessType = "A7" Then
                sstData.Tab = 1
                sstData.TabEnabled(1) = True
            Else
                If (InStr(1, UCase(strProcessType), "B") > 0) _
                    And (IsNumeric(Right(strProcessType, Len(strProcessType) - 1))) _
                    And (Val(Right(strProcessType, Len(strProcessType) - 1)) >= 13) Then
                    sstData.Tab = 1
                    sstData.TabEnabled(1) = True
                Else
                    sstData.Tab = 0
                    sstData.TabEnabled(0) = True
                End If
            End If
    
        End If
         If (InStr(1, UCase(strProcessType), "C") > 0) Then
            sstData.TabVisible(2) = intOTTCompID <> 0
            If sstData.TabVisible(2) Then
                sstData.Tab = 2
                sstData.TabEnabled(2) = True
            End If
         End If
         
      
        Select Case UCase(strProcessType)
            Case "A1"
                optBasic(0).Enabled = True
                optBasic(0).Value = True
            Case "A2"
                optBasic(1).Enabled = True
                optBasic(1).Value = True
            Case "A3"
                optBasic(2).Enabled = True
                optBasic(2).Value = True
            Case "A4"
                optBasic(3).Enabled = True
                optBasic(3).Value = True
            Case "E1"
                optBasic(4).Enabled = True
                optBasic(4).Value = True
            Case "B11"
                optBasic(5).Enabled = True
                optBasic(5).Value = True
                gimOrdSVOD.Enabled = True
                'gimOrdSVOD.SetQueryCode strSetCitemCode
            Case "B12"
                optBasic(6).Enabled = True
                optBasic(6).Value = True
                gimCancelSVOD.Enabled = True
                'gimCancelSVOD.SetQueryCode strSetCitemCode
            Case "A6"
                optBasic(7).Enabled = True
                optBasic(7).Value = True
            Case "A7"
                optSMITCommand(6).Enabled = True
                optSMITCommand(6).Value = True
            Case "A8"
                optBasic(8).Enabled = True
                optBasic(8).Value = True
            Case "A9"
                optBasic(9).Enabled = True
                optBasic(9).Value = True
            Case "B2B1"
                optBasic(6).Enabled = True
                optBasic(6).Value = True
                gimCancelSVOD.Enabled = True
                optBasic(5).Enabled = True
                gimOrdSVOD.Enabled = True
                lblResv.Visible = True
                gdtResvTime.Visible = True
            Case "B13"
                optSMITCommand(0).Value = True
                optSMITCommand(0).Enabled = True
                gilCitemCode.Enabled = True
            Case "B14"
                optSMITCommand(1).Value = True
                optSMITCommand(1).Enabled = True
            Case "B15"
                optSMITCommand(2).Value = True
                optSMITCommand(2).Enabled = True
                gimOrdSmitProd.Enabled = True
            Case "B16"
                optSMITCommand(3).Value = True
                optSMITCommand(3).Enabled = True
                gimCancelSmitProd.Enabled = True
             Case "B17"
                optSMITCommand(4).Value = True
                optSMITCommand(4).Enabled = True
                gimOrdCatchUpProd.Enabled = True
            Case "B18"
                optSMITCommand(5).Value = True
                optSMITCommand(5).Enabled = True
                gimCancelCatchUpProd.Enabled = True
        End Select
  
        
    Exit Sub
ChkErr:
    ErrSub Me.Name, "DoProcessMode"
End Sub

Private Sub subGil()
    On Error GoTo ChkErr
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 20, GetCompCodeFilter(gilCompCode, GaryGi(15))
        
        
'        SetgiList gilBaudRate, "CodeNo", "Description", "CD064", , , , , 3, 20, GetCompCodeFilter(gilCompCode, gilCompCode.GetCodeNo), True
'        SetgiList gilReason, "CodeNo", "Description", "CD014", , , , , 3, 20, , True
'        SetgiList gilVendor, "VendorCode", "VendorName", "CM006", , , , , 3, 20, " Where CompCode=" & gCompCode, True
        
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub subGim()
    On Error GoTo ChkErr
    Dim strFilter As String
        If strProcessChStr <> "" Then
'            SetgiMulti gimOpenCh, "CodeNo", "Description", "CD024", "�W�D�N�X", "�W�D�W��", , True
        Else
'            SetgiMulti gimOpenCh, "CodeNo", "Description", "CD024", "�W�D�N�X", "�W�D�W��", "Where Nvl(CompCode," & gilCompCode.GetCodeNo & ") = " & gilCompCode.GetCodeNo & strChNotQry, True
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
      '�w�qGirdR�U����쪺���e
      'STB�Ǹ�,ICC�Ǹ�,�˳]��m,���A
'        With mFlds
'            .Add "FaciSNo", , , , , "�]�ƧǸ�       ", vbLeftJustify
'            .Add "FaciCode", , , , , "�N�X ", vbLeftJustify
'            .Add "FaciName", , , , , "���ئW�� ", vbLeftJustify
'            .Add "ModelName", , , , , "�����W�� ", vbLeftJustify
'            .Add "CMBaudNo", , , , , "�N�X", vbLeftJustify
'            .Add "CMBaudRate", , , , , "CM�t�v ", vbLeftJustify
'            .Add "DynIPCount", , , , , "�ʺAIP�ƥ�", vbRightJustify
'            .Add "FixIPCount", , , , , "�T�wIP�ƥ�", vbRightJustify
'            .Add "InstDate", giControlTypeDate, , , , "�˾����", vbLeftJustify
'            .Add "CMOpenDate", giControlTypeTime, , , , "�}�����         ", vbLeftJustify
'            .Add "CMCloseDate", giControlTypeTime, , , , "�������         ", vbLeftJustify
'            .Add "DialAccount", , , , , "�����b�� ", vbLeftJustify
'            .Add "EMail", , , , , "�q�l�H�c           ", vbLeftJustify
'        End With
        
        With mFlds
            .Add "VODStatus", , , , False, "VOD���A"
            .Add "PRDate", , , , , "�]��  "
            .Add "PRDate", , , , , "�}�q  "
            .Add "VODAccountId", , , , , "VOD�b��           "
            .Add "FaciSNo", , , , , "STB                           ", vbLeftJustify
            .Add "SmartCardNo", , , , , "SmartCard            ", vbLeftJustify
            .Add "STBSNO", , , , , "HDD S/N                ", vbLeftJustify
            .Add "DVRAuthSizeCode", , , , , "���v�ɶ�           ", vbLeftJustify
            .Add "DVRSizeCode", , , , , "HDD �e�q", vbLeftJustify
            .Add "InstDate", giControlTypeDate, , , , "�˾����", vbLeftJustify
        End With
        ggrData.AllFields = mFlds
        ggrData.SetHead
        

        '�W�D�s��,�W�D�W��,���O�_�l���,�I����
'        With mFlds2
'            .Add "CustId", , , , False, "�Ȥ�s��", vbLeftJustify
'            .Add "Citemname", , , , False, "    ���O����    ", vbLeftJustify
'            .Add "ShouldAmt", , , , False, "�X����B", vbRightJustify
'            .Add "RealPeriod", , , , False, "����", vbLeftJustify
'            .Add "RealAmt", , , , False, "�ꦬ���B", vbRightJustify
'            .Add "RealStartDate", , , , False, "�����_�l��", vbLeftJustify
'            .Add "RealStopDate", , , , False, "�����I���", vbLeftJustify
'            .Add "CancelFlag", , , , False, "�@�o", vbLeftJustify
'            .Add "Note", , , , False, Space(20) & "  �Ƶ�            " & Space(15), vbLeftJustify
'        End With

        With mFlds2
            .Add "Status", , , , False, "���A   ", vbLeftJustify
            .Add "CitemCode", , , , False, "���O���ؽs��", vbLeftJustify
            .Add "Description", , , , False, "���O���ئW��         ", vbLeftJustify
            .Add "OpenTime", , , , False, "�}�W�D��" & Space(20), vbLeftJustify
            .Add "CloseTime", , , , False, "���W�D��" & Space(20), vbLeftJustify
        End With
        ggrChild.AllFields = mFlds2
        ggrChild.SetHead
        
        '�ѽX���s��,�ѽX���Ǹ�,�Ȥ�W��,�W�D�W��,�]�w�����]�w�ɶ�,�]�w�H��,���O����
'        With mFlds3
'            .Add "STBSNo", , , , , "���W���Ǹ�     ", vbLeftJustify
'            .Add "SmartCardNo", , , , , "���z�d�Ǹ�     ", vbLeftJustify
'            .Add "CustId", , , , , "�Ȥ�W��", vbLeftJustify
'            .Add "ChName", , , , , "�W�D�W��", vbLeftJustify
'            .Add "ModeType", , , , , "�]�w����     ", vbLeftJustify
'            .Add "UpdTime", , , , , "�]�w�ɶ�       ", vbLeftJustify
'            .Add "UpdEn", , , , , "�]�w�H��", vbLeftJustify
'            .Add "AuthorStartDate", giControlTypeDate, , , , "���v�_�l���", vbLeftJustify
'            .Add "AuthorStopDate", giControlTypeDate, , , , "���v�I����", vbLeftJustify
'            .Add "ResvTime", giControlTypeTime, , , , "�w���ɶ�", vbLeftJustify
'        End With
'        ggrQuery.AllFields = mFlds3
'        ggrQuery.SetHead
'
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
    If strCompCode = "" Then strCompCode = GaryGi(9)
    
    gilCompCode.SetCodeNo strCompCode
    gilCompCode.Query_Description
    
    '�p����CustId ginCustId �����\Key
    If lngCustId > 0 Then
        ginCustId.Text = lngCustId
        Call ginCustId_Validate(False)
        ginCustId.Enabled = False
    End If
    sstHead.Tab = 0
    If strCaption <> "" Then Me.Caption = Me.Caption & "-" & strCaption
    blnProcessOk = False
    If gLogInID <> "" Then gLogInID = gLogInID & "."
End Sub

Private Sub OpenData(Optional strWhere As String)
    On Error GoTo ChkErr
    Dim strSQL As String
        Set rsData = New ADODB.Recordset
        Set rsLog = New ADODB.Recordset
        If rsParent Is Nothing Then Set rsParent = New ADODB.Recordset
        If rsParent.State = adStateOpen Then
            Set rsData = rsParent.Clone
            If blnShowFaci Then
                rsData.Filter = " PRDate = Null"
            End If
            If strDefaultRowId = "" Then
                strDefaultRowId = rsParent("RowId") & ""
            Else
                If strDefaultRowId <> "" Then RsFind rsData, "RowId", strDefaultRowId
            End If
            'strDefaultRowId = rsParent("RowId") & ""
            'rsData.AbsolutePosition = rsParent.AbsolutePosition
        Else
            If ginCustId.Text <> "" Then
                strSQL = " A.CustId = " & ginCustId.Value & " And A.CompCode = " & gilCompCode.GetCodeNo & _
                        " And A.FaciSNo is not null And FaciCode In (Select CodeNo From " & GetOwner & "CD022 Where RefNo in (3))" & _
                        " And A.PRDate is null "
                If strWhere <> "" Then strSQL = strWhere
            Else
                strSQL = " A.SeqNo = ''"
            End If
            If Not GetRS(rsData, "Select A.RowId, A.* From " & GetOwner & "So004 A Where " & strSQL & " Order By A.InstDate", , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        End If
        Call OpenData2
        
        
        If (Not rsData Is Nothing) Then
            If rsData.RecordCount > 0 Then
                If Len(strDefaultRowId) = 0 Then rsData.MoveFirst
                If GetDTVCAType(rsData("FaciSNo") & "", rsData("SmartCardNo") & "") = 0 Then
                    sstData.Tab = 0
                    sstData.TabEnabled(1) = False
                Else
                    sstData.TabEnabled(1) = True
                    
                End If
                If qryPVRCMDType(rsData("ModelCode") & "") = 1 Then
                    sstData.Tab = 0
                    sstData.TabEnabled(1) = False
                End If
            End If
        End If
        Set ggrData.Recordset = rsData
        ggrData.Refresh
        intOTTCompID = Val(GetRsValue("Select OTTCOMPID From " & GetOwner & "SO041") & "")
        If intOTTCompID <> 0 Then
            strOTTOwner = GetRsValue("Select TableOwner From " & GetOwner & "CD039 Where CodeNo = " & intOTTCompID) & "."
            sstData.TabVisible(2) = intOTTCompID <> 0
            If sstData.TabVisible(2) Then
                sstData.TabEnabled(2) = True
            End If
           
        End If
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
'        If strWhere = "" Then
'            strWhere = " CustId = " & rsData("CustId") & " And FaciSeqNo = '" & rsData("SeqNo") & "' And UCCode is not null"
'        End If
'        If Not GetRS(rsChild, "Select RowId,A.* From " & GetOwner & "SO033 A Where " & strWhere) Then Exit Sub
        If strWhere = "" Then
            strWhere = " A.CitemCode=B.CodeNo "
        End If
        strWhere = strWhere & "  AND C.ROWID = '" & rsData("ROWID") & "'" & _
                " AND A.VODACCOUNTID = C.VODACCOUNTID "
        If Not GetRS(rsChild, "Select A.RowId,A.*,B.Description " & _
                              " From " & GetOwner & "SO555B A, " & _
                        GetOwner & "CD019 B," & GetOwner & "SO004 C " & _
                        " Where " & strWhere, gcnGi, adUseClient, _
                        adOpenKeyset, adLockOptimistic) Then Exit Sub
        
        If Len(strProcessType) = 0 Then Call EnableMode(True)
        Set ggrChild.Recordset = rsChild
        ggrChild.Refresh
    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData2"
End Sub

Private Sub EnableMode(blnFlag As Boolean)
    On Error Resume Next
        Dim i As Long
        If Not blnFlag Then Exit Sub
        optBasic(0).Enabled = GetPremission("SO11241")
        optBasic(1).Enabled = GetPremission("SO11242")
        optBasic(2).Enabled = GetPremission("SO11243")
        optBasic(3).Enabled = GetPremission("SO11244")
        optBasic(4).Enabled = GetPremission("SO11245")
        optBasic(5).Enabled = GetPremission("SO11246")
        
        optBasic(6).Enabled = GetPremission("SO11247")
        optBasic(7).Enabled = GetPremission("SO11248")
        optBasic(8).Enabled = GetPremission("SO1124A")
        optBasic(9).Enabled = GetPremission("SO1124B")
        optBasic(10).Enabled = False
        cmdQueryAccInfo.Enabled = False
        gimOrdSVOD.Enabled = optBasic(5).Enabled
        gimCancelSVOD.Enabled = optBasic(6).Enabled
        
        
        '#5469 ��VODStatus=2��,�u��A2�PA4�R�O�i�H��� By Kin 2010/01/07
        If Not rsData.EOF Then
            If Val(rsData("VODStatus") & "") = 2 Then
                For i = optBasic.LBound To optBasic.UBound
                    If Not (i = 1 Or i = 3) Then
                        optBasic(i).Enabled = False
                    End If
                Next i
            End If
            '#7068 �p�G�w�ҥ�,�h�����\�e�}��,���m�h�ۤ� By Kin 2015/08/06
            If Val(rsData("VodStatus") & "") = 1 Then
                optBasic(0).Enabled = False
                optBasic(1).Value = 0
                optBasic(10).Enabled = True
                '#7260 Check permission to access the function by kin 2016/06/16
                optBasic(10).Enabled = GetPremission("SO1124D")
            End If
            If Val(GetVodType(rsData("ModelCode") & "")) = 2 Then
                cmdQueryAccInfo.Enabled = True
            End If
        End If
        
'        optBasic(4).Enabled = GetPremission("SO11731") = Not blnUseFttx ' 1. �˾�
'        optBasic(0).Enabled = GetPremission("SO11732") ' 2. �}��
'        optBasic(1).Enabled = GetPremission("SO11733") ' 3. ����
'        optBasic(2).Enabled = GetPremission("SO11734") = Not blnUseFttx ' 4. ����
'        optBasic(3).Enabled = GetPremission("SO11735") = Not blnUseFttx ' 5. �_��
'        optBasic(5).Enabled = GetPremission("SO11736")  ' 6. ���
'        optBasic(6).Enabled = GetPremission("SO11737") ' 7. ���״���
'        optBasic(7).Enabled = GetPremission("SO1173I") = Not blnUseFttx ' 8. ���m�]��
'        optBasic(8).Enabled = GetPremission("SO1173H") = Not blnUseFttx ' 9. ���}�q
'
'        '************************************************************************
'        '#4276 �W�[FTTX���O By Kin 2008/12/23
'        optBasic(9).Enabled = GetPremission("SO1173L") = blnUseFttx ' 10.�w��FTTX�˾�
'        optBasic(10).Enabled = GetPremission("SO1173M") = blnUseFttx '11.�����w��FTTX�˾�
'        optBasic(11).Enabled = GetPremission("SO1173N") = blnUseFttx '12�ҥ�PPPOE�b��
'        optBasic(12).Enabled = GetPremission("SO1173R") = blnUseFttx '29 �P�ϲ��J
'        '#4384 �W�[ Release Port By Kin 2009/02/26
'        optBasic(13).Enabled = GetPremission("SO1173S") = blnUseFttx ' 33 Release Port
'        '************************************************************************
'
'        'cmdEmergency.Enabled = GetPremission("SO1173H") ' 8. ���}�q
'        optCMChange(0).Enabled = GetPremission("SO11738") ' 1. �ܧ�򥻸��
'        optCMChange(1).Enabled = GetPremission("SO11739") ' 2. �ܧ�ӽФH
'        optCMChange(2).Enabled = GetPremission("SO1173A") ' 3. �ܧ�t�v
'        optCMChange(3).Enabled = GetPremission("SO1173B") ' 4. �ܧ����
'        optCMChange(4).Enabled = GetPremission("SO1173C") = Not blnUseFttx ' 5. �ܧ�K�X
'        optCMChange(5).Enabled = GetPremission("SO1173J") ' 6. �T�{�ϥΤH
'        '**************************************************************************
'        '#4276 �W�[FTTX���O By Kin 2008/12/23
'        optCMChange(6).Enabled = GetPremission("SO1173O") = blnUseFttx '�ӽЩT�wIP
'        optCMChange(7).Enabled = GetPremission("SO1173P") = blnUseFttx '�����ӽЩT�wIP
'        optCMChange(8).Enabled = GetPremission("SO1173Q") = blnUseFttx '���]PPPOE�K�X
'        '**************************************************************************
'
'        optQueryInfo(0).Enabled = GetPremission("SO1173D") ' 1. �d��CM��T
'        optQueryInfo(1).Enabled = GetPremission("SO1173E") ' 2. �d�� �b����T
'
'        optCMDevice(0).Enabled = GetPremission("SO1173F") ' 1. �]�ƤJ�w
'        optCMDevice(1).Enabled = GetPremission("SO1173G") ' 1. �]�ƥX�w
'        '#3800 �W�[�w�����O�����v�� By Kin 2008/03/12
'        cmdCancelCMD.Enabled = GetPremission("SO1173K")
'        blnUseCancelCmd = True
        '#4276 �n�b�P�_�ϥ�FTTX�Ϊ̬OCM���O By Kin 2008/12/23
        'Call CtlEnabled
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        Call CloseRecordset(rsData)
        Call CloseRecordset(rsChild)
        Call CloseRecordset(rsLog)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        ReleaseCOM Me
        Set frmSO1623A = Nothing
End Sub

Private Sub ggrChild_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    If UCase(giFld.FieldName) = "STATUS" Then
        If Fld.Value = 0 Then
            Value = vbRed
        End If
    End If
End Sub

Private Sub ggrChild_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    If giFld.FieldName = "Status" Then
        If Value = 0 Then
            Value = "��"
        Else
            Value = "�}"
        End If
    End If
End Sub

Private Sub ggrData_MoveRecord()
  On Error GoTo ChkErr
   
    Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "ggrData_MoveRecord")
End Sub

Private Sub ggrQuery_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    If UCase(giFld.FieldName) = "STOPFLAG" Then
        If CStr(Fld.Value) = "1" Then
            Value = vbRed
        End If
    End If
    
End Sub

Private Sub ggrQuery_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error Resume Next
    If UCase(giFld.FieldName) = "CMD" Then
        Select Case UCase(Value)
            Case "A1"
                Value = "�ҥ�VOD"
            Case "A2"
                Value = "����VOD"
            Case "A3"
                Value = "�Ȱ�VOD"
            Case "A4"
                Value = "��_VOD"
            Case "A6"
                Value = "VOD-���״���(���)"
            Case "A8"
                Value = "VOD-���״���(STB)"
            Case "A9"
                Value = "VOD-���״���(ICC)"
            Case "E1"
                Value = "VOD-���m�K�X"
            Case "B11"
                Value = "�q��SVOD"
            Case "B12"
                Value = "����SVOD"
            Case "Z2"
                Value = "�q�ʬd��"
            Case "B13"
                Value = "�]�w���v�e�q"
            Case "B14"
                Value = "�������v�e�q"
            Case "B15"
                Value = "�}�q�M�\Prod ID"
            Case "B16"
                Value = "�����M�\Prod ID"
            Case "B17"
                Value = "�}��CatchUp"
            Case "B18"
                Value = "����CatchUp"
            Case "A7"
                Value = "�]�w���W��Tuner�ƶq"
            Case "C1"
                Value = "�d��VOD�b��"
        End Select
    End If
    If UCase(giFld.FieldName) = "CMDSTATUS" Then
        Select Case UCase(Value)
            Case "S"
                Value = "���\"
            Case "W"
                Value = "�ݳB�z"
            Case "P"
                Value = "�B�z��"
            Case "E"
                Value = "����"
            Case "T"
                Value = "Time Out"
            Case Else
                Value = "�������~"
        End Select
    End If
'    Select Case UCase(giFld.FieldName)
'        Case "MODETYPE"
'            Value = GetModeTypeValue(Value)
'        Case "CMDSTATUS"
'            If UCase(Value) & "" = "W" Then
'                Value = "���B�z"
'            ElseIf UCase(Value) & "" = "P" Then
'                Value = "�B�z��"
'            ElseIf UCase(Value) & "" = "E" Then
'                Value = "�����~"
'            ElseIf UCase(Value) & "" = "C" Then
'                Value = "���T"
'            Else
'                Value = ""
'            End If
'        Case "STOPFLAG"
'            If CStr(Value) = "1" Then
'                Value = "�O"
'            Else
'                Value = "�_"
'            End If
'    End Select
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error GoTo ChkErr
    Dim strDvrShowCellTmp As String
    Dim strSTBFaciSNO As String
    Dim strSTBSeqNo As String
    Dim strDVRAuthSizeCode As String
    strSTBFaciSNO = ggrData.Recordset("FaciSNO") & ""
    strSTBSeqNo = ggrData.Recordset("SEQNO") & ""
    If UCase(giFld.FieldName) = "VODSTATUS" Then
        Select Case Val(Value & "")
            Case 0
                Value = "���ҥ�"
            Case 1
                Value = "�ҥ�"
            Case 2
                Value = "�Ȱ�"
        End Select
    End If
    If UCase(giFld.FieldName) = "PRDATE" Then
        Select Case Trim(giFld.HeadName)
            'Case "�]�ƪ��A"
            Case "�]��"
                Value = GetFaciStatus(ggrData.Recordset)
            'Case "�}�q���A"
            Case "�}�q"
                Value = GetFaciCommandStatus(ggrData.Recordset)
                'Value = ggrData.Recordset("CMOpenDate") & ""
        End Select
    End If
    If UCase(giFld.FieldName) = "STBSNO" Then
        If Len(strSTBFaciSNO) > 0 Then
            Value = GetDVRSno(lngCustId, strSTBFaciSNO, , , strSTBSeqNo)
            strDvrShowCellTmp = Value
        End If
    End If
    If UCase(giFld.FieldName) = UCase("DVRSizeCode") Then '�����w�Юe�q
        Value = GetDVRHDD(lngCustId, strSTBFaciSNO, , strSTBSeqNo, True) & "(GB)"
    End If
    If UCase(giFld.FieldName) = UCase("DVRAuthSizeCode") Then '���v�e�q
        If Len(strSTBFaciSNO) > 0 And Len(strDvrShowCellTmp) > 0 Then
            strDVRAuthSizeCode = GetRsValue("Select DVRAuthSizeCode From " & GetOwner & "SO004 Where FaciSno='" & strDvrShowCellTmp & "' and STBSNO='" & strSTBSeqNo & "' and Custid=" & lngCustId) & ""
            If Len(Trim(strDVRAuthSizeCode)) > 0 Then
                Value = GetRsValue("Select Description From " & GetOwner & "CD102 Where CodeNo=" & strDVRAuthSizeCode) & ""
                'SetgiList gilgilReSetDVR, "CodeNo", "Description", "CD019", , , , , 20, 20, " Where StopFlag=0 Type=1 ", True
            Else
                Value = ""
            End If
        Else
            Value = ""
        End If
    End If
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "ggrData_ShowCellData")
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If gilCompCode.GetCodeNo = "" Then Exit Sub
        GaryGi(16) = gilCompCode.GetOwner
'        '#3731 ��CMType=3�ɤ��P�_�զX���~ By Kin 2008/01/15
'        If GetRsValue("Select NVL(CMType,0) From " & GetOwner & "CD039 Where CodeNo=" & gilCompCode.GetCodeNo) = 3 Then
'            blnBpCode = False
'            intCMType = 3
'        Else
'            blnBpCode = True
'        End If
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If pic1.Visible Then Exit Sub
    If KeyCode = vbKeyEscape Then If cmdCancel.Enabled Then cmdCancel.Value = True
    Select Case sstHead.Tab
        Case 0
            If KeyCode = vbKeyF2 Then If cmdSet.Enabled Then cmdSet.Value = True
        Case 1
            If KeyCode = vbKeyF3 Then If cmdQuery.Enabled Then cmdQuery.Value = True
            If KeyCode = vbKeyF10 Then If cmdCancelCMD.Enabled Then cmdCancelCMD.Value = True
            If KeyCode = vbKeyF4 Then If cmdQueryAccInfo.Enabled Then cmdQueryAccInfo.Value = True
    End Select
End Sub




Private Sub gilOTTCitem_Change()
    optOTTQuery(0).Value = True
End Sub

Private Sub gimCancelSVOD_ButtonClick()
  On Error GoTo ChkErr
    Call CancelOptValue
    optBasic(5).Value = False
    optBasic(6).Value = True
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gimCancelSVOD_ButtonClick")
End Sub

Private Sub gimCancelSVOD_SelectChange()
    On Error Resume Next
    If Len(strResvTime) > 0 Then Exit Sub
    If gimCancelSVOD.GetDispStr <> "" Then
        lblResv.Visible = True
        gdtResvTime.Visible = True
    Else
        lblResv.Visible = False
        gdtResvTime.Visible = False
    End If
End Sub





Private Sub gimOrdSVOD_ButtonClick()
  On Error GoTo ChkErr
    Call CancelOptValue
    optBasic(5).Value = True
    optBasic(6).Value = False
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gimOrdSVOD_ButtonClick")
End Sub
Private Sub CancelOptValue()
  On Error GoTo ChkErr
    Dim i As Integer
    For i = optBasic.LBound To optBasic.UBound
        optBasic(i).Value = False
    Next
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "CancelOptValue")
End Sub

Private Sub gimOrdSVOD_SelectChange()
    On Error Resume Next
    If Len(strResvTime) > 0 Then Exit Sub
    If gimOrdSVOD.GetDispStr <> "" Then
        lblResv.Visible = True
        gdtResvTime.Visible = True
    Else
        lblResv.Visible = False
        gdtResvTime.Visible = False
    End If
End Sub

Private Sub optBasic_Click(Index As Integer)
  On Error Resume Next
    'picCPno.Visible = False
    HideResv
    If Index = 5 Then
        gimOrdSVOD.Enabled = True
    Else
        'gimOrdSVOD.Enabled = False
        gimOrdSVOD.Clear
    End If
    If Index = 6 Then
        gimCancelSVOD.Enabled = True
    Else
        'gimCancelSVOD.Enabled = False
        gimCancelSVOD.Clear
    End If
    If Len(strResvTime) = 0 Then
        lblResv.Visible = False
        gdtResvTime.Visible = False
        If Index = 5 Or Index = 6 Then
            lblResv.Visible = True
            gdtResvTime.Visible = True
        End If
    End If
    
'    Select Case Index
''                Case 0 ' 1. �}�� [02]
''                    If optBasic(0).Value Then
''                        chkAccStop.Value = 0
''                        HideResv
''                        ShowResv
''                    End If
''                    chkPR.Enabled = False
''                Case 1 ' 2. ���� [02]
''                    If optBasic(1).Value Then
''                        ChkAccStart.Value = 0
''                        HideResv
''                        ShowResv
''                    End If
''                    chkPR.Enabled = False
'                '���״���
'                Case 6
'                    If optBasic(6).Value Then
'                        '********************************************************************************************************
'                        '#4276 �p�G�OFTTX�u�����FTTX���]�� By Kin 2009/01/09
'                        If blnUseFttx Then
'                            'SetgiList gilFaciCode, "CodeNo", "Description", "CD022", , , , , 3, 12, " WHERE REFNO IN (8)", True
'                        Else
'                            'SetgiList gilFaciCode, "CodeNo", "Description", "CD022", , , , , 3, 12, " WHERE REFNO IN (2,5)", True
'                        End If
'                        '*******************************************************************************************************
''                        gilFaciCode.SetCodeNo ""
''                        gilFaciCode.SetDescription ""
''                        txtFaciSNo.Text = ""
''                        lblFaci.Visible = True
''                        lblCMFacisno.Visible = True
'                        'gilFaciCode.Visible = True
''                        txtFaciSNo.Visible = True
'
'                    End If
'                    If strFaciCode <> "" Then
''                        gilFaciCode.SetCodeNo strFaciCode
''                        gilFaciCode.Query_Description
''                        gilFaciCode.Enabled = False
'                    End If
'                    If strFaciSNo <> "" Then
''                        txtFaciSNo.Text = strFaciSNo
''                        txtFaciSNo.Enabled = False
'                    End If
'                '***********************************************
'                '#3778 �W�[���ʭ�] By Kin 2008/02/25
'                Case 1, 2, 5
'                    gilReason.Visible = True
'                    lblReason.Visible = True
'                    If strReason <> "" Then
'                        gilReason.SetCodeNo strReason
'                        gilReason.Query_Description
'                        gilReason.Enabled = False
'                    Else
'                        gilReason.SetCodeNo ""
'                        gilReason.SetDescription ""
'                    End If
'                '***********************************************
'
'                Case Else
'                    HideResv
'                    'chkPR.Enabled = False
'    End Select
End Sub

Private Sub ShowResv()
  On Error Resume Next
    lblResv.Visible = True
    gdtResvTime.Visible = True
    'lblResv.Visible = True
    'gdtResvTime.Visible = True
    'gilBaudRate.Visible = True
End Sub

Private Sub HideResv()
  On Error Resume Next
'    lblResv.Visible = False
'    lblFaci.Visible = False
'    lblCMFacisno.Visible = False
'    gilFaciCode.Visible = False
'    gilVendor.Visible = False
'    txtFaciSNo.Visible = False
'    txtPassword.Visible = False
''    gdtResvTime.Visible = False
'    gilBaudRate.Visible = False
'    gilReason.Visible = False
'    lblReason.Visible = False
'    gdtResvTime.Text = ""
End Sub

Private Sub optSMITCommand_Click(Index As Integer)
    If strProcessType <> "" Then Exit Sub
    If Index = 0 Then
        gilCitemCode.Enabled = True
    Else
        gilCitemCode.Clear
        gilCitemCode.Enabled = False
    End If
    If Index = 2 Then
        gimOrdSmitProd.Enabled = True
    Else
        gimOrdSmitProd.Clear
        gimOrdSmitProd.Enabled = False
    End If
    If Index = 3 Then
        gimCancelSmitProd.Enabled = True
    Else
        gimCancelSmitProd.Clear
        gimCancelSmitProd.Enabled = False
    End If
    If Index = 4 Then
        gimOrdCatchUpProd.Enabled = True
    Else
        gimOrdCatchUpProd.Clear
        gimOrdCatchUpProd.Enabled = False
    End If
    If Index = 5 Then
        gimCancelCatchUpProd.Enabled = True
    Else
        gimCancelCatchUpProd.Clear
        gimCancelCatchUpProd.Enabled = False
    End If
End Sub

Private Sub rsData_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
        HideResv
        IsWaringResvTime = True
        If rsData.RecordCount > 0 Then
            Call CtlEnabled
        Else
            Exit Sub
        End If
        If Not ggrData.Enabled Then Exit Sub
        If Not blnSetting Then
            'Me.Enabled = False
            Call OpenData2
            Fetch_Log_Data , True
            Call SetgimOrdVOD
            Call SetgimCancelVod
            'Me.Enabled = True
        End If
        Dim aSQL As String
        If (strOTTOwner & "") <> "" Then
            SetOTTgiList gilOTTCitem, "CitemCode", "CitemName", "SO003", , , , , 3, 20, " Where 1 = 0 "
        End If
        aSQL = "SELECT CODENO,DESCRIPTION FROM " & GetOwner & "CD019 WHERE CODENO = 0 "
        SetMsQry gimOrdSmitProd, aSQL, _
                                , , "���O���إN�X,���O���ئW��", "1200,4200"
        SetMsQry gimCancelSmitProd, aSQL, _
                , , "���O���إN�X,���O���ئW��", "1200,4200"
        If Not rsData Is Nothing Then
            If rsData.RecordCount > 0 Then
                If strProcessType = "" Then
                     If GetDTVCAType(rsData("FaciSNo") & "", rsData("SmartCardNo") & "") = 0 Then
                        sstData.Tab = 0
                        sstData.TabEnabled(1) = False
                     Else
                        optSMITCommand(4).Enabled = IsCatchUp(rsData)
                        sstData.TabEnabled(1) = True
                         SetgiList gilCitemCode, "CodeNo", "Description", "CD019", , , , , 3, 20, " WHERE CODENO IN (" & _
                                            GetRecordCitemCodeSQL(rsData) & ")"
                       
                        aSQL = "SELECT CODENO,DESCRIPTION FROM " & GetOwner & "CD019 WHERE Nvl(StopFlag,0)=0 " & _
                                    " AND CODENO IN (" & GetSmitChargeCitemCodeSQL(rsData, True, True, Empty) & ")"
                        SetMsQry gimOrdSmitProd, aSQL, _
                                , , "���O���إN�X,���O���ئW��", "1200,4200"
                         aSQL = "SELECT CODENO,DESCRIPTION FROM " & GetOwner & "CD019 WHERE Nvl(StopFlag,0)=0 " & _
                                    " AND CODENO IN (" & GetSmitChargeCitemCodeSQL(rsData, False, True, Empty) & ")"
                        SetMsQry gimCancelSmitProd, aSQL, _
                                , , "���O���إN�X,���O���ئW��", "1200,4200"

                    End If
                    If sstData.TabVisible(2) Then
                        Dim aOTTWhere As String
                        aOTTWhere = " Where 1 = 0 "
'                        aOTTWhere = " Where FVORDERNO Is Null And CustId in (Select CustId From " & strOTTOwner & "SO220 Where " & _
'                                        " OTTUserID In (Select OTTUserID From " & strOTTOwner & "SO222 Where  SOCompID = " & rsData("CompID") & _
'                                        " And FaciID = '" & rsData("FaciID") & "' ) ) "
                                        
                        aOTTWhere = " Where FVORDERNO Is Null And CustId in (Select CustId From " & strOTTOwner & "SO220 Where " & _
                                        " OTTUserID In (Select OTTUserID From " & strOTTOwner & "SO222 Where  SOCompID = " & rsData("COMPCODE") & _
                                        " And FaciID = '" & rsData("FaciID") & "' ) ) "
                        aOTTWhere = " Where FVORDERNO Is NOT Null And custid=671"
                        If (strOTTOwner & "") <> "" Then
                            SetOTTgiList gilOTTCitem, "CitemCode", "CitemName", "SO003", , , , , 3, 20, aOTTWhere
                        End If
                     End If
                Else
                
                        
                Dim maxSizeCitemCode As String
                    If (strProcessType = "B15") Or (strProcessType = "B16") Or (strProcessType = "B17") Or (strProcessType = "B18") Or (strProcessType = "B13") Or (strProcessType = "B14") Then
                        Select Case UCase(strProcessType)
                            Case "B13"
                                
                                Dim SetCitemCode As String
                                If Len(strSetCitemCode) > 0 Then
                                    SetCitemCode = strSetCitemCode
                                    SetgiList gilCitemCode, "CodeNo", "Description", "CD019", , , , , 3, 20, " WHERE CODENO In ( " & SetCitemCode & ") "
                                Else
                                    aSQL = GetChargeRecordCitemCodeSQL(rsData)
                                    SetgiList gilCitemCode, "CodeNo", "Description", "CD019", , , , , 3, 20, " WHERE CODENO IN (" & _
                                                aSQL & ")"
                                    SetCitemCode = aSQL
                                    IsWaringResvTime = False
                                End If
                               
                                maxSizeCitemCode = GetMaxSizeCodeCitemCode(SetCitemCode)
                                If Len(maxSizeCitemCode) > 0 Then
                                    gilCitemCode.SetCodeNo maxSizeCitemCode
                                    gilCitemCode.Query_Description
                                Else
                                    IsWaringResvTime = False
                                End If
                                                                  
                            Case "B14"
                                If (Len(strCancelCitemCode) > 0) Then
                                    maxSizeCitemCode = GetMaxSizeCodeCitemCode(strCancelCitemCode)
                                    If Len(maxSizeCitemCode) = 0 Then
                                       IsWaringResvTime = False
                                    End If
                                Else
                                    IsWaringResvTime = False
                                 End If
                            Case "B15"
                                aSQL = "SELECT CODENO,DESCRIPTION FROM " & GetOwner & "CD019 WHERE Nvl(StopFlag,0)=0 " & _
                                            " AND CODENO IN (" & GetSmitChargeCitemCodeSQL(rsData, True, True, strSetCitemCode) & ")  Order By CodeNo "
                             
                                If Val(GetRsValue("Select Count(1) From ( " & aSQL & ")")) > 0 Then
                                    SetMsQry gimOrdSmitProd, aSQL, _
                                                          , , "���O���إN�X,���O���ئW��", "1200,4200"
                                    gimOrdSmitProd.SelectAll
                                Else
                                    SetMsQry gimOrdSmitProd, "SELECT CODENO,DESCRIPTION FROM " & GetOwner & "CD019 WHERE 2=0 ", _
                                                          , , "���O���إN�X,���O���ئW��", "1200,4200"
                                End If
                                
                                If Len(gimOrdSmitProd.GetQueryCode) > 0 Then
                                    If Len(GetSmitProductId(gimOrdSmitProd.GetQueryCode, _
                                        rsData("CustId"), rsData("FaciSNo"), rsData("CompCode"), True, True)) = 0 Then
                                         IsWaringResvTime = False
                                    End If
                                Else
                                   IsWaringResvTime = False
                                End If
                            Case "B16"
                                 
                                If Len(strCancelCitemCode) > 0 Then
                                    aSQL = "SELECT CODENO,DESCRIPTION FROM " & GetOwner & "CD019 WHERE Nvl(StopFlag,0)=0 " & _
                                                    " AND CODENO IN (" & strCancelCitemCode & ")"
                                    SetMsQry gimCancelSmitProd, aSQL, _
                                                        , , "���O���إN�X,���O���ئW��", "1200,4200"
                                    gimCancelSmitProd.SelectAll
                                    If Len(GetSmitProductId(strCancelCitemCode, rsData("CustId"), rsData("FaciSNo"), rsData("CompCode"), True, True)) = 0 Then
                                         IsWaringResvTime = False
                                    End If
                                Else
                                    IsWaringResvTime = False
                                End If
                            Case "B17"

                                If Len(strSetCitemCode) = 0 Then
                                    IsWaringResvTime = False
                                Else
                                    IsWaringResvTime = IsCatchUpCitemCode(strSetCitemCode)
                                End If
                            Case "B18"
                                If Len(strCancelCitemCode) = 0 Then
                                    IsWaringResvTime = False
                                Else
                                    IsWaringResvTime = IsCatchUpCitemCode(strCancelCitemCode)
                                End If
                        End Select
 
                    End If
                End If
                If qryPVRCMDType(rsData("ModelCode") & "") = 1 Then
                    If sstData.Tab <> 2 Then
                        sstData.Tab = 0
                    End If
                    sstData.TabEnabled(1) = False
                End If
                '#7068 �p�G�w�ҥ�,�h�����\�e�}��,���m�h�ۤ� By Kin 2015/08/06
                If Val(rsData("VODStatus") & "") = 1 Then
                    optBasic(0).Enabled = False
                    optBasic(0).Value = 0
                    optBasic(10).Enabled = True
                    '#7260 Check permission to access the function by kin 2016/06/16
                    optBasic(10).Enabled = GetPremission("SO1124D")
                Else
                    optBasic(1).Enabled = True
                    optBasic(10).Enabled = False
                    optBasic(10).Value = 0
                End If
                '#7068 �p�GVODType=2�h�d�߱b�����O�~�i�ϥ� By Kin 2015/08/10
                cmdQueryAccInfo.Enabled = False
                If Val(GetVodType(rsData("ModelCode") & "")) = 2 Then
                    cmdQueryAccInfo.Enabled = True
                End If
            End If
        End If
End Sub

Private Sub ginCustId_Validate(Cancel As Boolean)
    On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
        If ginCustId.Text = "" Then
            txtCustName = ""
            strCustName = ""
            Exit Sub
        Else
            If Not GetRS(rs, "Select CustName,InstAddrNo,InstAddress,TELH1,Tel1 From " & GetOwner & _
                            "So001 Where CustId = " & ginCustId.Value & " And CompCode = " & gilCompCode.GetCodeNo, gcnGi, , , , , , , True) Then Exit Sub
            If rs.EOF Then
                txtCustName = ""
                strCustName = ""
                lngInstAddrNo = 0
                strInstAddress = ""
                lngNodeNo = 0
                lngCircuitNo = 0
                strZipCode = ""
            Else
                txtCustName = rs("CustName") & ""
                strCustName = rs("CustName") & ""
                strInstAddress = rs("InstAddress") & ""
                strCustTel = IIf(rs("TelH1") & "" = "", "", rs("TelH1") & "-") & rs("Tel1") & ""
                lngInstAddrNo = Val(rs("InstAddrNo") & "")
                lngNodeNo = 0
                lngCircuitNo = 0
                strZipCode = GetRsValue("Select ZipCode From " & GetOwner & "SO014 Where AddrNo=" & rs("InstAddrNo"), gcnGi) & ""
            End If
        End If

        If txtCustName = "" Then
            If ginCustId.Text <> "" Then MsgBox "�d�L���Ȥ�", vbCritical, gimsgPrompt
            ggrData.Enabled = False
            ggrChild.Enabled = False
            OpenData ("1=0")
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

Public Property Let uProcessType(ByVal vData As String)
    On Error Resume Next
        strProcessType = vData
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
Public Property Let uOrdChannelID(ByVal vData As String)
  On Error Resume Next
    strSetCitemCode = vData
End Property
Public Property Let uCancelChannelID(ByVal vData As String)
  On Error Resume Next
    strCancelCitemCode = vData
End Property

Public Property Let uByHouse(ByVal vData As Integer)
    On Error Resume Next
        intByHouse = vData
End Property

Private Function IsDataOk() As Boolean
    On Error GoTo ChkErr
        If gdtResvTime.Visible Then
            If gdtResvTime.GetValue <> "" Then
                    If Len(strProcessType & "") > 0 Then
                        If (strProcessType = "B15") Or (strProcessType = "B16") Or _
                            (strProcessType = "B17") Or (strProcessType = "B18") Or _
                            (strProcessType = "B13") Or (strProcessType = "B14") Or (strProcessType = "A7") Then
                            If qryPVRCMDType(rsData("ModelCode") & "") = 1 Then
                                IsWaringResvTime = False
                            End If
                        End If
                    End If
                
                If (CDate(gdtResvTime.GetValue(True)) <= CDate(RightNow)) And (IsWaringResvTime) Then
                    MsgBox "�w���ɶ��ݤj��{�b�ɶ�", vbCritical, gimsgPrompt
                    gdtResvTime.SetValue DateAdd("h", 1, RightNow)
                    gdtResvTime.SetFocus
                    Exit Function
                Else
                    '7676
'                    If (CDate(gdtResvTime.GetValue(True)) > CDate(RightNow) + 1) And (IsWaringResvTime) Then
'                        MsgBox "�w���ɶ��ݤp�����ɶ�", vbCritical, gimsgPrompt
'                        gdtResvTime.SetValue DateAdd("h", 1, RightNow)
'                        gdtResvTime.SetFocus
'                        Exit Function
'                    End If
                    If Len(strResvTime & "") > 0 Then
                        If Len(str7SNO & "") > 0 Then
                            Dim compareDate As String
                            compareDate = strResvTime
                            If IsDate(compareDate) Then
                                compareDate = Format(compareDate, "yyyyMMddHHmmss")
                            End If
                            If Not IsDate(compareDate) Then
                                'compareDate = Format(compareDate, "yyyyMMddHHmmss")
'                                compareDate = Mid(compareDate, 1, 4) & "/" & Mid(compareDate, 5, 2) & "/" & Mid(compareDate, 7, 2) & " " & Mid(compareDate, 9, 2) & ":" & _
'                                                    Mid(compareDate, 11, 2) & ":" & Right(compareDate, 2)
                                compareDate = Mid(compareDate, 1, 4) & "/" & Mid(compareDate, 5, 2) & "/" & Mid(compareDate, 7, 2) & " " & "00" & ":" & _
                                                    "00" & ":" & "00"
                            End If
                            If CDate(gdtResvTime.GetValue(True) & "") > CDate(compareDate) + 1 And (IsWaringResvTime) Then
                                MsgBox "�w���ɶ��ݤp��u��w�����", vbCritical, gimsgPrompt
                                gdtResvTime.SetValue strResvTime
                                gdtResvTime.SetFocus
                                Exit Function
                            End If
                        End If
                    End If
                    
                    If (Not blnEmergency) And (Len(strProcessType) = 0) Then
                        '#3778 ���w���ǰe���T�� By Kin 2008/02/22
                        Dim strShowTime As String
                        Dim strExeName As String
                        Dim objControl As Object
                        Dim i As Long
                        strShowTime = gdtResvTime.GetValue
                        strShowTime = Mid(strShowTime, 1, 4) & "�~:" & Mid(strShowTime, 5, 2) & "��:" & Mid(strShowTime, 7, 2) & "��:" & _
                                    Mid(strShowTime, 9, 2) & "��:" & Mid(strShowTime, 11, 2) & "��:" & Mid(strShowTime, 13, 2) & "��"
                         Select Case sstData.Tab
                            Case 0
                                Set objControl = optBasic
                            Case 1
                                Set objControl = optSMITCommand
                            Case 2
                                'Set objControl = optQueryInfo
                            Case 3
                                Set objControl = optCMDevice
                        End Select
                        For i = 0 To objControl.UBound
                            If objControl(i) Then
                                strExeName = Mid(objControl(i).Caption, 4)
                                Exit For
                            End If
                        Next
    '                    If MsgBox("���]�w�w���ɶ�,�{���N���|����CM Command Gateway �^��,�T�w�ϥιw���\��??", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then Exit Function
                        If MsgBox("�w�p" & strShowTime & "�A�N�ǰe " & strExeName & " ���O�A�T�w�ϥιw���\��?", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then Exit Function
                        '�w�p�~:��:��:��:��:��A�N�ǰeXX���O
                    End If
                End If
            End If
        Else
            gdtResvTime.SetValue ""
        End If
        
       
        Select Case sstData.Tab
            Case 0
                Select Case True
                    Case optBasic(0)
                    Case optBasic(1)
                    Case optBasic(2)
                    Case optBasic(5)
                        If Len(UCase(strProcessType)) <= 0 Then
                            If gimOrdSVOD.GetDispStr & "" = "" Then
                                MsgBox "�ЬD��q�ʪ��W�D !", vbInformation, "�T��"
                                IsDataOk = False
                                Exit Function
                            End If
                        End If
                    Case optBasic(6)
                        If Len(UCase(strProcessType)) <= 0 Then
                            If gimCancelSVOD.GetDispStr = "" Then
                                MsgBox "�ЬD��������W�D !", vbInformation, "�T��"
                                IsDataOk = False
                                Exit Function
                            End If
                        End If
                    Case optBasic(7)
                        If rs2.State = adStateClosed Then
                            MsgBox "�s�ȨS���ǤJ", vbInformation, "�T��"
                            IsDataOk = False
                            Exit Function
                        Else
                            
                            If rs2.EOF Then
                                MsgBox "�s�ȨS���ǤJ", vbInformation, "�T��"
                                IsDataOk = False
                            Exit Function
                            End If
                        End If
                    Case optBasic(8)
                        If rs2.State = adStateClosed Then
                            MsgBox "�s�ȨS���ǤJ", vbInformation, "�T��"
                            IsDataOk = False
                            Exit Function
                        Else
                            If rs2.EOF Then
                                MsgBox "�s�ȨS���ǤJ", vbInformation, "�T��"
                                IsDataOk = False
                            Exit Function
                            End If
                        End If
                    Case optBasic(9)
                        If rs2.State = adStateClosed Then
                            MsgBox "�s�ȨS���ǤJ", vbInformation, "�T��"
                            IsDataOk = False
                            Exit Function
                        Else
                            If rs2.EOF Then
                                MsgBox "�s�ȨS���ǤJ", vbInformation, "�T��"
                                IsDataOk = False
                            Exit Function
                            End If
                        End If
                    Case Else

                End Select
            Case 1
                Select Case True
                    Case optSMITCommand(0)
                        If gilCitemCode.GetCodeNo & "" = "" And Len(strProcessType) = 0 Then
                            MsgBox "�ЬD�怜�O���ءI", vbInformation, "�T��"
                            IsDataOk = False
                            Exit Function
                        End If
                    Case optSMITCommand(2)
                        If strProcessType = "" Then
                            If gimOrdSmitProd.GetQueryCode & "" = "" Then
                                MsgBox "�ЬD��q���W�D�I", vbInformation, "�T��"
                                IsDataOk = False
                                Exit Function
                            End If
                            
                        End If
                    Case optSMITCommand(3)
                        If strProcessType = "" Then
                            If gimCancelSmitProd.GetQueryCode & "" = "" Then
                                MsgBox "�ЬD������W�D�I", vbInformation, "�T��"
                                IsDataOk = False
                                Exit Function
                            End If
                        End If
'                    Case optSMITCommand(4)
'                        If strProcessType = "" Then
'                            If gimOrdCatchUpProd.GetQueryCode & "" = "" Then
'                                MsgBox "�ЬD��q���W�D�I", vbInformation, "�T��"
'                                IsDataOk = False
'                                Exit Function
'                            End If
'                        End If
'                    Case optSMITCommand(5)
'                        If strProcessType = "" Then
'                            If gimCancelCatchUpProd.GetQueryCode & "" = "" Then
'                                MsgBox "�ЬD������W�D�I", vbInformation, "�T��"
'                                IsDataOk = False
'                                Exit Function
'                            End If
'                        End If
                End Select
                
            
        End Select

        
        IsDataOk = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "IsDataOk"
End Function


Private Sub rsLog_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  '#5891 �P�_�w���������s�O�_�i�ϥ� By Kin 2011/03/25
  On Error Resume Next
    Dim strQrySO560 As String
    If sstHead.Tab <> 1 Then Exit Sub
    
    cmdCancelCMD.Enabled = True
    gilReturnCode.Enabled = True
    lblReturnCode.Enabled = True
    
    strQrySO560 = "Select Count(*) From " & GetOwner & "SO555 " & _
                " Where SeqNo=" & rsLog("SeqNo") & _
                " AND CMDSTATUS = 'W'" & _
                " AND ResvTime IS NOT NULL "
                
    If (gcnGi.Execute(strQrySO560)(0) <= 0) Then
        gilReturnCode.Clear
        cmdCancelCMD.Enabled = False
        gilReturnCode.Enabled = False
        lblReturnCode.Enabled = False
    End If
    Exit Sub

End Sub

Private Sub sstData_Click(PreviousTab As Integer)
  On Error Resume Next
'    Select Case sstData.Tab
'        Case 0
'            If optBasic(6) Then Call optBasic_Click(6)
'        Case 1
'            If optCMChange(7) Then Call optCMChange_Click(7)
'            If optCMChange(8) Then Call optCMChange_Click(8)
'        Case Else
'
'    End Select
    EnableMode blnShowFmt
End Sub

Private Function GetPCMac() As String
    On Error Resume Next
    GetPCMac = GetRsValue("Select CPEMAC From " & GetOwner & "SO004C Where FaciSNo = '" & rsData("FaciSNo") & "' And CustId = " & rsData("CustId") & " And RowNum=1", gcnGi, "") & ""
End Function

Property Let uAutoSend(ByVal vData As Boolean)
  On Error Resume Next
    blnAutoSend = vData
End Property

Private Sub sstHead_Click(PreviousTab As Integer)
  On Error Resume Next
    If sstHead.Tab = 0 Then
        'lblResv.Visible = True
        'gdtResvTime.Visible = True
    ElseIf sstHead.Tab = 1 Then
        If ggrQuery.Recordset.RecordCount = 0 Then
            lblReturnCode.Enabled = False
            gilReturnCode.Enabled = False
            cmdCancelCMD.Enabled = False
        End If
    End If
    
End Sub
'Private Function GetFaciStatus(ByVal rs004 As ADODB.Recordset, Optional ByVal Value) As String
'    On Error Resume Next
'    Dim FaciWipStatus As String
'    Dim rs As New ADODB.Recordset
'        With rs004
'            If .State = adStateClosed Then Exit Function
'            If .EOF Or .BOF Then Exit Function
''            1.���`=>InstDate is not null and PRDate is null
''            2.����=>InstDate is not null and PRDate is null And FaciStatusCode = 3
''            3.���/�����^=>PRDate is not null and GetDate is null
''            4.���/���^=>PRDate is not null and GetDate is not null
''            5.�L:InstDate is null and PrDate is null
'            If .Fields("InstDate") & "" <> "" And .Fields("PRDate") & "" = "" Then
'                If .Fields("FaciStatusCode") & "" = 3 Then
'                    FaciWipStatus = "����"
'                Else
'                    If StrToDate(.Fields("StopTime") & "", True) > StrToDate(.Fields("ReInstTime") & "", True) Then
'                        FaciWipStatus = "����"
'                    Else
'                        FaciWipStatus = "���`"
'                    End If
'                End If
'            ElseIf .Fields("PRDate") & "" <> "" Then
'                FaciWipStatus = "���"
'                If .Fields("GetDate") & "" = "" Then
'                    FaciWipStatus = FaciWipStatus & "/�����^"
'                End If
'            Else
'                FaciWipStatus = "�L"
'            End If
'            If .Fields("GetDate") & "" <> "" Then
'                FaciWipStatus = FaciWipStatus & "/���^"
'            End If
'        End With
'        GetFaciStatus = FaciWipStatus
'End Function
'Private Function GetFaciWipStatus(ByVal rs004 As ADODB.Recordset) As String
'  On Error Resume Next
'    Dim strStatus As String
'    strStatus = gcnGi.Execute("Select Kind From " & GetOwner & "SO004D Where CustId = " & rs004("CustId") & " And SeqNo = '" & rs004("SeqNo") & "' And SignDate is Null Order by UpdTime Desc").GetString(, , , "/") & ""
'    GetFaciWipStatus = Mid(strStatus, 1, Len(strStatus) - 1)
'End Function
'Private Function GetFaciCommandStatus(ByRef rs004 As ADODB.Recordset) As String
'    On Error Resume Next
'    Dim FaciWipStatus As String
'    Dim intFaciRef As Integer
'    Dim rs As New ADODB.Recordset
'        With rs004
'            If .State = adStateClosed Then Exit Function
'            If .EOF Or .BOF Then Exit Function
''            �}�q���A
''            1.���}��=>CMOpenDate is null and CMCloseDate is null
''            2.�}��=>CMOpenDate > CMCloseDate or CMOpendate is not null and CMCloseDate is null
''            3.����=>CMOpenDate < CMCloseDate
''            4.�}��/���v=>(CMOpenDate > CMCloseDate or CMOpendate is not null and CMCloseDate is null) and
''            (disableaccount> enableAccount or (disableaccount is not null and enableAccount is null)
''            5.����/���v=>(CMOpenDate < CMCloseDate) and
''            (disableaccount> enableAccount or (disableaccount is not null and enableAccount is null)
'            '#5227 2009.08.06 by Corey �]�Ȥ�{��ICC�]�ƬO��STB�j�b�@�_�A����STB�}�qICC�٬O���}�q�A�Q�׫�վ� �]�ưѦҸ� Null,0,4 �N���ݭn�q�X�}�q���A
'            intFaciRef = Val(GetRsValue("Select Refno From " & GetOwner & "CD022 Where StopFlag=0 and CodeNo=" & .Fields("FaciCode")) & "")
'            If Not (intFaciRef = 4 Or intFaciRef = 0) Then
'                If StrToDate(.Fields("CMOpenDate") & "", True, True) > StrToDate(.Fields("CMCloseDate") & "", True, True) Then
'                    FaciWipStatus = "�}��"
'                ElseIf StrToDate(.Fields("CMOpenDate") & "", True, True) < StrToDate(.Fields("CMCloseDate") & "", True, True) Then
'                    FaciWipStatus = "����"
'                ElseIf .Fields("CMOpenDate") & "" = "" And .Fields("CMCloseDate") & "" = "" Then
'                    FaciWipStatus = "���}��"
'                Else
'                    FaciWipStatus = "�L"
'                End If
'                If StrToDate(.Fields("disableaccount") & "", True, True) > StrToDate(.Fields("enableAccount") & "", True, True) Then
'                    FaciWipStatus = FaciWipStatus & "/���v"
'                End If
'            Else
'                FaciWipStatus = ""
'            End If
'        End With
'        GetFaciCommandStatus = FaciWipStatus
'End Function
Private Function GetDVRSno(ByVal lngDVRCustID As Long, ByVal strSTBSNo As String, _
                          Optional ByRef strDVRFaciCode As String = "", Optional ByRef strDVRHDDCode As String = "", _
                          Optional ByRef strSTBSeqNo As String = "", Optional ByRef blnClearPrSNO As Boolean = False, _
                          Optional ByRef strDVRHDDSizeCode As String = "", Optional ByVal strPRSNO As String = "") As String
On Error GoTo ChkErr
    Dim rsDVR As New ADODB.Recordset
    Dim strDVRSNOfnG As String
    Dim intLoop As Integer
    strDVRFaciCode = ""
    strDVRHDDCode = ""
    strDVRHDDSizeCode = ""
                
    If strSTBSeqNo = "" Then
        strSTBSeqNo = GetRsValue("Select SeqNo From " & GetOwner & "SO004 Where FaciSno='" & strSTBSNo & "' and Custid=" & lngDVRCustID & IIf(blnClearPrSNO, "", " and PrSNO Is Null")) & ""
    End If
    
    If Not GetRS(rsDVR, "Select FaciSNO,FaciCode,DVRAuthSizeCode,DVRSizeCode From " & GetOwner & "SO004 Where STBSNO='" & strSTBSeqNo & "' and Custid=" & lngDVRCustID & IIf(Len(strPRSNO) > 0, " and PRSNo = '" & strPRSNO & "'", IIf(blnClearPrSNO, "", " and PrSNO Is Null"))) Then Exit Function
    If rsDVR.RecordCount > 0 Then rsDVR.MoveFirst
    Do While Not rsDVR.EOF
        strDVRSNOfnG = strDVRSNOfnG & "/" & rsDVR("FaciSNO") & ""
        strDVRFaciCode = strDVRFaciCode & "/" & rsDVR("FaciCode") & ""
        strDVRHDDCode = strDVRHDDCode & "/" & rsDVR("DVRAuthSizeCode") & ""
        strDVRHDDSizeCode = strDVRHDDSizeCode & "/" & rsDVR("DVRSizeCode") & ""
        rsDVR.MoveNext
    Loop
    If Len(strDVRSNOfnG) > 0 Then GetDVRSno = Mid(strDVRSNOfnG, 2)
    If Len(strDVRFaciCode) > 0 Then strDVRFaciCode = Mid(strDVRFaciCode, 2)
    If Len(strDVRHDDCode) > 0 Then strDVRHDDCode = Mid(strDVRHDDCode, 2)
    If Len(strDVRHDDSizeCode) > 0 Then strDVRHDDSizeCode = Mid(strDVRHDDSizeCode, 2)
    
    CloseRecordset rsDVR
    Exit Function
ChkErr:
    ErrSub FormName, "GetDVRSno"
End Function

Private Function GetDVRHDD(ByVal lngDVRCustID As Long, ByVal strSTBSNo As String, _
                          Optional ByRef strDvrHDDData As String, Optional ByVal strSTBSeqNo As String, _
                          Optional ByVal blnAuth As Boolean = False, Optional ByVal blnCancel As Boolean) As String
On Error GoTo ChkErr
    Dim strDVRSNOfnHD As String, intDVRCount As Integer
    Dim strDSNO() As String, strDvrHDDSno As String
    Dim rsDVR As New ADODB.Recordset
    Dim strDvrSizeName As String
    
    If blnAuth Then
        strDvrSizeName = "DVRAuthSizeCode"
    Else
        strDvrSizeName = "DVRSizeCode"
    End If
    
    strDVRSNOfnHD = GetDVRSno(lngDVRCustID, strSTBSNo, , , strSTBSeqNo, blnCancel)
    If Len(strDVRSNOfnHD) > 0 Then
        strDSNO = Split(strDVRSNOfnHD, "/")
        For intDVRCount = 0 To UBound(strDSNO)
            '#5265 2009.08.26 by Corey �{���W�[PRDATE�@IS�@NULL������A�åB�p�G����x�ۦPFACISNO�h��SEQNO�̤j�ȨӨ��oDVR���v�e�q�C
            If Not GetRS(rsDVR, "Select * From " & GetOwner & "CD102 Where CodeNo in " & _
                        "(Select " & strDvrSizeName & " FROM (Select " & strDvrSizeName & ",(RANK() OVER ( ORDER BY SeqNo Desc)) RankX From " & _
                                                      GetOwner & "SO004 Where FaciSno='" & strDSNO(intDVRCount) & "'" & _
                                                      " and Custid=" & lngDVRCustID & " and STBSNO='" & strSTBSeqNo & "' and PRDATE IS NULL)" & _
                        " WHERE RankX = 1) ") Then Exit Function
            If rsDVR.RecordCount > 0 Then
                strDvrHDDSno = strDvrHDDSno & "/" & Trim(str(Val(rsDVR("DVRSize") & "")))
                strDvrHDDData = strDvrHDDData & "/" & Trim(str(Val(rsDVR("DVRSize") & "")))
            End If
        Next
        If Len(strDvrHDDSno) > 0 Then GetDVRHDD = Mid(strDvrHDDSno, 2)
        If Len(strDvrHDDData) > 0 Then strDvrHDDData = Mid(strDvrHDDData, 2)
    End If
    CloseRecordset rsDVR
    Exit Function
ChkErr:
    GetDVRHDD = "0"
    ErrSub FormName, "GetDVRSno"
End Function

