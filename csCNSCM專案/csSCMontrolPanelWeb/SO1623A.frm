VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{CEA96874-3B0B-4FF6-90B8-E48EE8F30715}#2.0#0"; "GiCombo.ocx"
Begin VB.Form frmSO1623A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "CM �]�w����x [SO1171A]"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
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
   ScaleHeight     =   7230
   ScaleWidth      =   11910
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
      TabIndex        =   36
      Top             =   2130
      Visible         =   0   'False
      Width           =   5025
      Begin MSComctlLib.ProgressBar pbr1 
         Height          =   285
         Left            =   150
         TabIndex        =   38
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
         TabIndex        =   37
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
      TabIndex        =   41
      Top             =   3030
      Width           =   2805
      Begin Gi_Time.GiTime gdtResvTime 
         Height          =   285
         Left            =   990
         TabIndex        =   53
         Top             =   0
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
         TabIndex        =   42
         Top             =   60
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
      TabIndex        =   39
      Top             =   390
      Visible         =   0   'False
      Width           =   1695
   End
   Begin prjNumber.GiNumber ginCustId 
      Height          =   285
      Left            =   5130
      TabIndex        =   76
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
      TabIndex        =   31
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
      TabIndex        =   24
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
      Tab(1).Control(0)=   "cmdCancelCMD"
      Tab(1).Control(1)=   "cmdQuery2"
      Tab(1).Control(2)=   "cmdQuery"
      Tab(1).Control(3)=   "ggrQuery"
      Tab(1).Control(4)=   "gilReturnCode"
      Tab(1).Control(5)=   "lblReturnCode"
      Tab(1).ControlCount=   6
      Begin VB.CommandButton cmdCancelCMD 
         Caption         =   "�w�����O����(F10)"
         Height          =   345
         Left            =   -69150
         TabIndex        =   60
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
         TabIndex        =   40
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
         TabIndex        =   30
         Top             =   3660
         Width           =   1125
      End
      Begin prjGiGridR.GiGridR ggrQuery 
         Height          =   3165
         Left            =   -74670
         TabIndex        =   35
         Top             =   450
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   5583
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
         Left            =   10110
         TabIndex        =   1
         Top             =   400
         Width           =   1245
      End
      Begin TabDlg.SSTab sstData 
         Height          =   3555
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   6271
         _Version        =   393216
         Tabs            =   4
         Tab             =   1
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
         TabCaption(0)   =   "&1. �򥻳]�w"
         TabPicture(0)   =   "SO1623A.frx":047A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fraData(0)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&2. �ܧ�A��"
         TabPicture(1)   =   "SO1623A.frx":0496
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "fraData(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "&3. �d�߸�T"
         TabPicture(2)   =   "SO1623A.frx":04B2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraData(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "&4.�]�ƥX�J�w"
         TabPicture(3)   =   "SO1623A.frx":04CE
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "optCMDevice(1)"
         Tab(3).Control(1)=   "optCMDevice(0)"
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
            TabIndex        =   22
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
            TabIndex        =   21
            ToolTipText     =   "�NSTB �Ȱ�"
            Top             =   600
            Width           =   1515
         End
         Begin VB.Frame fraData 
            Height          =   2955
            Index           =   2
            Left            =   -74760
            TabIndex        =   44
            Top             =   420
            Width           =   10635
            Begin VB.OptionButton optQueryInfo 
               Caption         =   "&1. �d�߳]�Ƹ�T [30]"
               BeginProperty Font 
                  Name            =   "�s�ө���"
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
               TabIndex        =   19
               ToolTipText     =   "�d�߳]�Ƹ�T"
               Top             =   300
               Value           =   -1  'True
               Width           =   1965
            End
            Begin VB.OptionButton optQueryInfo 
               Caption         =   "&2. �d�߱b����T [31]"
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
               Index           =   1
               Left            =   330
               TabIndex        =   20
               ToolTipText     =   "�d�߱b����T"
               Top             =   675
               Width           =   1935
            End
            Begin prjGiGridR.GiGridR ggrQueryInfo 
               Height          =   2655
               Left            =   2310
               TabIndex        =   29
               Top             =   180
               Width           =   8175
               _ExtentX        =   14420
               _ExtentY        =   4683
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
         End
         Begin VB.Frame fraData 
            Height          =   2955
            Index           =   1
            Left            =   240
            TabIndex        =   43
            Top             =   420
            Width           =   10485
            Begin VB.OptionButton optCMChange 
               Caption         =   "10.���]PPPOE�K�X [28]"
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
               Index           =   8
               Left            =   510
               TabIndex        =   73
               ToolTipText     =   "���]PPPOE�K�X"
               Top             =   2130
               Width           =   2235
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&9. �����ӽЩT�wIP [27]"
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
               Index           =   7
               Left            =   5100
               TabIndex        =   72
               ToolTipText     =   "�����ӽЩT�wIP"
               Top             =   1740
               Width           =   2235
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&1. ��CM [02]"
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
               Index           =   99
               Left            =   9660
               TabIndex        =   70
               Top             =   2310
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&2. �ܧ�CM�t�v [02]"
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
               Index           =   98
               Left            =   9720
               TabIndex        =   69
               Top             =   2640
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&3. ACS�b���K�X�����\�� [24]"
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
               Index           =   97
               Left            =   9630
               TabIndex        =   68
               ToolTipText     =   "�NSTB���s�ҥ�"
               Top             =   2430
               Visible         =   0   'False
               Width           =   2355
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&4. Clear CPE IP [24]"
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
               Index           =   96
               Left            =   9600
               TabIndex        =   67
               ToolTipText     =   "�NSTB���s�ҥ�"
               Top             =   2160
               Visible         =   0   'False
               Width           =   2295
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&8. �ӽЩT�wIP [26]"
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
               Index           =   6
               Left            =   510
               TabIndex        =   66
               ToolTipText     =   "�ӽЩT�wIP"
               Top             =   1740
               Width           =   2235
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&7. �T�{�ϥΤH [25]"
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
               Index           =   5
               Left            =   5100
               TabIndex        =   57
               ToolTipText     =   "�T�{�ϥΤH"
               Top             =   1320
               Width           =   2505
            End
            Begin VB.TextBox txtPassword 
               Height          =   315
               Left            =   6690
               MaxLength       =   24
               TabIndex        =   18
               Top             =   780
               Visible         =   0   'False
               Width           =   2295
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&3. �ܧ�t�v [22]"
               Enabled         =   0   'False
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
               Index           =   2
               Left            =   510
               TabIndex        =   13
               ToolTipText     =   "�ܧ�t�v"
               Top             =   1320
               Width           =   1545
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&6. �ܧ�K�X [24]"
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
               Index           =   4
               Left            =   5100
               TabIndex        =   17
               ToolTipText     =   "�ܧ�K�X"
               Top             =   825
               Width           =   2505
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&5. �ܧ���� [23]"
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
               Index           =   3
               Left            =   5100
               TabIndex        =   15
               ToolTipText     =   "�ܧ����"
               Top             =   330
               Width           =   1575
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&1. �ܧ�򥻸�� [20]"
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
               Left            =   510
               TabIndex        =   11
               ToolTipText     =   "�ܧ�򥻸��"
               Top             =   330
               Width           =   2235
            End
            Begin VB.OptionButton optCMChange 
               Caption         =   "&2. �ܧ�ӽФH [21]"
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
               Left            =   510
               TabIndex        =   12
               ToolTipText     =   "�ܧ�ӽФH"
               Top             =   825
               Width           =   2505
            End
            Begin prjGiList.GiList gilBaudRate 
               Height          =   285
               Left            =   2100
               TabIndex        =   14
               Top             =   1290
               Visible         =   0   'False
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
            Begin prjGiList.GiList gilVendor 
               Height          =   285
               Left            =   6720
               TabIndex        =   16
               Top             =   270
               Visible         =   0   'False
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
            Begin prjGiList.GiList gilBP 
               Height          =   285
               Left            =   9720
               TabIndex        =   71
               Top             =   1860
               Visible         =   0   'False
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
         End
         Begin VB.Frame fraData 
            Height          =   2955
            Index           =   0
            Left            =   -74640
            TabIndex        =   34
            Top             =   390
            Width           =   10485
            Begin VB.OptionButton optBasic 
               Caption         =   "&13.Release Port"
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
               Index           =   13
               Left            =   5310
               TabIndex        =   75
               ToolTipText     =   "Release Port"
               Top             =   2160
               Width           =   2235
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&12.FTTX�P�ϲ��J[29]"
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
               Index           =   12
               Left            =   5310
               TabIndex        =   74
               ToolTipText     =   "FTTX�P�ϲ��J"
               Top             =   1500
               Width           =   2235
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&11.�ҥ�PPPOE�b��[10]"
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
               Index           =   11
               Left            =   5310
               TabIndex        =   65
               ToolTipText     =   "�ҥ�PPPOE�b��"
               Top             =   1080
               Width           =   2235
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&10.�����w��FTTX�˾�[19]"
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
               Left            =   5310
               TabIndex        =   64
               ToolTipText     =   "�����w��FTTX�˾�"
               Top             =   570
               Width           =   2595
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&9. �w��FTTX�˾�[18]"
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
               Left            =   2970
               TabIndex        =   63
               ToolTipText     =   "�w��FTTX�˾�"
               Top             =   2490
               Width           =   2055
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&7. ���m�]�� [17]"
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
               Left            =   3000
               TabIndex        =   56
               ToolTipText     =   "���m�]��"
               Top             =   2130
               Width           =   2265
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&8. ���}�q"
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
               Left            =   510
               TabIndex        =   55
               ToolTipText     =   $"SO1623A.frx":04EA
               Top             =   2490
               Width           =   1815
            End
            Begin VB.CommandButton cmdEmergency 
               BackColor       =   &H00FFC0C0&
               Caption         =   "���}�q"
               Height          =   315
               Left            =   8850
               Style           =   1  '�Ϥ��~�[
               TabIndex        =   54
               Top             =   2160
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&1. �}�� [11]"
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
               Index           =   0
               Left            =   510
               TabIndex        =   52
               ToolTipText     =   "�}��"
               Top             =   1080
               Width           =   1335
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&5. �פ�A�ȱb��[15]"
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
               Left            =   3000
               TabIndex        =   51
               ToolTipText     =   "�פ�A�ȱb��"
               Top             =   1080
               Width           =   2145
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&4. �_�� [14]"
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
               Left            =   3000
               TabIndex        =   8
               ToolTipText     =   $"SO1623A.frx":04F7
               Top             =   600
               Width           =   1785
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&6. �]�Ƨ� [16]"
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
               Left            =   3000
               TabIndex        =   23
               ToolTipText     =   "�]�Ƨ�"
               Top             =   1500
               Width           =   1815
            End
            Begin VB.TextBox txtFaciSNo 
               Height          =   330
               Left            =   7290
               MaxLength       =   20
               TabIndex        =   10
               Top             =   1740
               Visible         =   0   'False
               Width           =   2295
            End
            Begin VB.PictureBox picCPno 
               Appearance      =   0  '����
               BorderStyle     =   0  '�S���ؽu
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   8610
               ScaleHeight     =   315
               ScaleWidth      =   1665
               TabIndex        =   47
               Top             =   2190
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.CheckBox ChkAccStart 
               Caption         =   "�b���_�v"
               BeginProperty Font 
                  Name            =   "�s�ө���"
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
               TabIndex        =   45
               Top             =   1080
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&0. �إ�CM�A�ȱb�� [10]"
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
               Left            =   510
               TabIndex        =   5
               ToolTipText     =   "�إ�CM�A�ȱb��"
               Top             =   600
               Width           =   2235
            End
            Begin VB.CheckBox chkPR 
               Caption         =   "�������O�_���ηs�a�}��T"
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
               ForeColor       =   &H000000C0&
               Height          =   225
               Left            =   9480
               TabIndex        =   28
               Top             =   1950
               Visible         =   0   'False
               Width           =   2535
            End
            Begin VB.CheckBox chkAccStop 
               Caption         =   "�b�����v"
               BeginProperty Font 
                  Name            =   "�s�ө���"
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
               TabIndex        =   26
               Top             =   1500
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&3. ���� [13]"
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
               Left            =   510
               TabIndex        =   7
               ToolTipText     =   "����"
               Top             =   2100
               Width           =   2265
            End
            Begin VB.OptionButton optBasicXX 
               Caption         =   "&4. �_��"
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
               Left            =   9480
               TabIndex        =   27
               ToolTipText     =   $"SO1623A.frx":0501
               Top             =   2310
               Visible         =   0   'False
               Width           =   1365
            End
            Begin VB.OptionButton optBasic 
               Caption         =   "&2. ���� [12]"
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
               Left            =   510
               TabIndex        =   6
               ToolTipText     =   $"SO1623A.frx":050D
               Top             =   1500
               Width           =   1815
            End
            Begin PrjGiCmobo.GiCombo cboCPno 
               Height          =   300
               Left            =   9840
               TabIndex        =   46
               Top             =   2400
               Visible         =   0   'False
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   529
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "�s�ө���"
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
            Begin prjGiList.GiList gilFaciCode 
               Height          =   285
               Left            =   3405
               TabIndex        =   9
               Top             =   1800
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
            Begin prjGiList.GiList gilReason 
               Height          =   285
               Left            =   7380
               TabIndex        =   58
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
            Begin VB.Label lblReason 
               Caption         =   "���ʭ�]"
               Height          =   195
               Left            =   6510
               TabIndex        =   59
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lblFaci 
               AutoSize        =   -1  'True
               Caption         =   "�~�W"
               Height          =   225
               Left            =   3000
               TabIndex        =   50
               Top             =   1830
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.Label lblCMFacisno 
               AutoSize        =   -1  'True
               Caption         =   "�]�ƧǸ�"
               Height          =   195
               Left            =   6510
               TabIndex        =   49
               Top             =   1800
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.Label lblCPno 
               Caption         =   "CP����"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   9270
               TabIndex        =   48
               Top             =   2490
               Visible         =   0   'False
               Width           =   1320
            End
         End
      End
      Begin prjGiList.GiList gilReturnCode 
         Height          =   285
         Left            =   -66600
         TabIndex        =   62
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
         TabIndex        =   61
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
      TabIndex        =   33
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
      TabIndex        =   32
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
'Dim strChStr As String, strStopChStr As String, strOpenChStr As String
'Dim strChStopDate As String
Dim rsParent As New ADODB.Recordset
Dim intProcessType As Integer
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
Const lngCMD30Top As Long = 300
Const lngCMD31Top As Long = 675
Const strCMD30String As String = "&" & "1. �d�߳]�Ƹ�T [30]"
Const strCMD31String As String = "&" & "2. �d�߱b����T [31]"
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
    
    If cboCPno.ListCount <= 0 Then MsgBox "���]�ƵLCP�����A�L�k���`�ǰe���O !", vbInformation, "�T��"
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
    If pic1.Visible Then Exit Sub
    Unload Me
    Set frmSO1623A = Nothing
End Sub

Private Sub cmdCancelCMD_Click()
  On Error GoTo ChkErr
    Dim strUpdSO314 As String
    Dim varBookMark As Variant
    Dim objCn As New ADODB.Connection
    If intCMType <> 3 Then
        MsgBox "CMTYPE����3�A�����\����өR�O�I", vbInformation, "�T��"
        Exit Sub
    End If
    If gilReturnCode.GetCodeNo = Empty Then
        MsgBox "�h���]��������!!", vbInformation, "�T��"
        gilReturnCode.SetFocus
        Exit Sub
    End If
    varBookMark = rsLog.Bookmark
    strUpdSO314 = "Update " & GetOwner & "SO314 " & _
                " Set CancelCode=" & gilReturnCode.GetCodeNo & "," & _
                "CancelName='" & gilReturnCode.GetDescription & "'," & _
                "StopFlag=1," & _
                "CancelDate=To_Date('" & Format(RightNow, "YYYYMMDDHHMMSS") & "','YYYY/MM/DD HH24:MI:SS')," & _
                "HandleEn='" & GaryGi(0) & "'," & _
                "HandleName='" & GaryGi(1) & "'," & _
                "HandleTime=To_Date('" & Format(RightNow, "YYYYMMDDHHMMSS") & "','YYYY/MM/DD HH24:MI:SS')," & _
                "HandleNote='�ӵ���ƵL������u��!!' Where CmdSeqNo=" & rsLog("CmdSeqNo")
    objCn.ConnectionString = gcnGi.ConnectionString
    objCn.CursorLocation = 3
    objCn.Open
    objCn.BeginTrans
    objCn.Execute (strUpdSO314)
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

Private Sub cmdEmergency_Click()
  On Error GoTo ChkErr
    blnEmergency = True
    Dim lngBR As Long
    If gdtResvTime.GetValue = Empty Then
        Me.Enabled = True
        MsgBox "�ϥκ��}�q������J�����w���ɶ�", vbInformation, "�T��"
        Exit Sub
    End If
'    If Not IsDataOk Then Exit Sub
    Me.Enabled = False
    blnSetting = True
    Call subProgressBar(True)
    On Error GoTo ErrTran
    If Not blnTransation Then gcnGi.BeginTrans
    lngBR = rsData.AbsolutePosition
    lblProcess.Caption = "�}�� �� , �еy�� !!"
    If Not OpenCM(rsData, "11", rsData("FaciSNo") & "") Then GoTo ErrTran
    Call OpenData2
    If Not OpenCM(rsData, "12", rsData("FaciSNo") & "", , , , gdtResvTime.GetValue) Then GoTo ErrTran
    lblProcess.Caption = "���� �� , �еy�� !!"
    Call OpenData2
    blnSetting = False
    If Not blnTransation Then gcnGi.CommitTrans
    rsData.AbsolutePosition = lngBR
    blnProcessOk = True
    blnEmergency = False
    Call subProgressBar(False)
    Me.Enabled = True
    MsgBox "���}�q���\�I", vbInformation, "�T��"
    If intProcessType > 0 Then Unload Me
    Exit Sub
ErrTran:
    On Error Resume Next
    blnEmergency = False
    rsData.AbsolutePosition = lngBR
    On Error GoTo ChkErr
    Call DoProcessMode
    blnSetting = False
    If Not blnTransation Then gcnGi.RollbackTrans
    OpenData
    rsData.AbsolutePosition = lngBR
    pic1.Visible = False
    Me.Enabled = True
    MsgBox "���}�q���ѡI", vbInformation, "�T��"
    Exit Sub
ChkErr:
    Me.Enabled = True
    blnEmergency = False
    ErrSub Me.Name, "cmdEmergency_Click"
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
    Dim strQry2 As String
    If blnEmptyGrid Then
        strQry = "SELECT * FROM " & GetOwner & "SO314 WHERE 0=1"
    Else
        If rsData Is Nothing Then Exit Sub
        With rsData
            If .State <= 0 Then Exit Sub
            If .EOF Or .BOF Then
                '#3800 �p�G�L��Ƭd�ߩҦ�SO314�U���ӫȽs��� By Kin 2008/03/14
                If blnPRflag Then
                    '#7408 adding table of so314log to query data by kin 2017/02/21
                    strQry = "Select * From ( Select * From " & GetOwner & "SO314" & _
                            " Where CompCode=" & gilCompCode.GetCodeNo & _
                            " And CUSTID=" & ginCustId.Text & _
                            " UNION " & _
                            "Select * From " & GetOwner & "SO314LOG" & _
                            " Where CompCode=" & gilCompCode.GetCodeNo & _
                            " And CUSTID=" & ginCustId.Text & _
                            " ) ORDER BY Decode(CMDSTATUS,'W',1,2),UpdTime DESC"
                    If GetRS(rsLog, strQry, gcnGi, adUseClient, adOpenStatic, adLockReadOnly) Then
                        GrdFmt rsLog
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
            End If
            lngRcdCnt = .RecordCount
            If lngRcdCnt = 0 Then Exit Sub
        End With
        strQry = "SELECT * FROM " & GetOwner & "SO314" & _
                        " WHERE CompCode=" & rsData("CompCode") & _
                        " AND CUSTID=" & rsData("CustId")
                
        If lngRcdCnt = 1 Then
            If blnPRflag Then
                '#5165 ���դ�OK,���K�վ㪺�W��,�u�n�OFttx���AccountId�Ӭd By Kin 2009/07/29
                If blnUseFttx Then
                    strQry = strQry & " And AccountId<>'" & rsData("AccountNo") & "'"
                Else
                    strQry = strQry & " AND DeviceSNo1<>'" & rsData("FaciSNo") & "'"
                End If
            Else
                '#5165 ���դ�OK,���K�վ㪺�W��,�u�n�OFttx���AccountId�Ӭd By Kin 2009/07/29
                If blnUseFttx Then
                    strQry = strQry & " And AccountId='" & rsData("DialAccount") & "'"
                Else
                    strQry = strQry & " AND DeviceSNo1='" & rsData("FaciSNo") & "'"
                End If
            End If
        Else
            If blnPRflag Then
                '#5165 ���դ�OK,���K�վ㪺�W��,�u�n�OFttx���AccountId�Ӭd By Kin 2009/07/29
                If blnUseFttx Then
                    strQry = strQry & " And AccountId<>'" & rsData("DialAccount") & "'"
                Else
                    strQry = strQry & " AND DeviceSNo1 NOT IN (" & Get_Faci_Rows & ")"
                End If
            Else
                'strQry = strQry & " AND DeviceSNo1 IN (" & Get_Faci_Rows & ")"
                '#5165 ���դ�OK,���K�վ㪺�W��,�u�n�OFttx���AccountId�Ӭd By Kin 2009/07/29
                If blnUseFttx Then
                    strQry = strQry & " And AccountId='" & rsData("DialAccount") & "'"
                Else
                    strQry = strQry & " AND DeviceSNo1='" & rsData("FaciSNo") & "'"
                End If
            End If
        End If
        
        'strQry = strQry & " ORDER BY Decode(CMDSTATUS,'W',1,2),UpdTime DESC"
        '#7408 adding table of so314log to query data by kin 2017/02/21
        strQry2 = strQry
        strQry2 = Replace(strQry2, "SO314", "SO314LOG")
        strQry = strQry & " UNION " & strQry2
        strQry = "Select * From (" & strQry & " ) ORDER BY CMDSTATUS,UpdTime Desc"
        'strQry = strQry & " ORDER BY CMDSTATUS,UpdTime Desc"
    End If
    If GetRS(rsLog, strQry, gcnGi, adUseClient, adOpenStatic, adLockReadOnly) Then
        GrdFmt rsLog
    End If
    
    
'    '*****************************************************************************************************
'    '#3800 �d�ߤv��]��,�p�G�L��ƫh�d��SO314�̩Ҧ��Ƚs��� By Kin 2008/03/12
'    If GetRS(rsLog, strQry, gcnGi, adUseClient, adOpenStatic, adLockReadOnly) Then
'        If rsLog.RecordCount > 0 Then
'            GrdFmt rsLog
'        Else
'            If blnPRflag Then
'                strQry = "Select * From " & GetOwner & "SO314" & _
'                        " Where CompCode=" & rsData("CompCode") & _
'                        " And CUSTID=" & rsData("CustiId") & _
'                        " ORDER BY Decode(CMDSTATUS,'W',1,2),UpdTime DESC"
'                If GetRS(rsLog, strQry, gcnGi, adUseClient, adOpenStatic, adLockReadOnly) Then
'                    GrdFmt rsLog
'                End If
'            Else
'                GrdFmt rsLog
'            End If
'        End If
'    End If
'    '******************************************************************************************************
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
        '.Add "MSGCODE", , , , , "���O�N�X", vbLeftJustify
        .Add "MSGNAME", , , , , "���O�W�� ", vbLeftJustify
        .Add "OPERATORID", , , , , "�ާ@�H��ID", vbLeftJustify
        .Add "CMDSTATUS", , , , , "�ާ@���浲�G", vbLeftJustify
        .Add "PROCESSINGDATE", giControlTypeTime, , , , "���O�w������ɶ�" & Space(10), vbLeftJustify
        .Add "UPDTIME", giControlTypeTime, , , , "���ʮɶ�" & Space(15), vbLeftJustify
        .Add "STOPFLAG", , , , , "�O�_����", vbLeftJustify
        .Add "ReasonName", , , , , "���ʭ�]" & Space(15), vbLeftJustify
        .Add "ERRCODE", , , , , "���O���~�N�X", vbLeftJustify
        .Add "ERRMSG", , , , , "���O���~�W��" & Space(10), vbLeftJustify
        .Add "COMPCODE", , , , , "���q�O", vbLeftJustify
        '.Add "CUSTID", , , , , "�Ȥ�s��", vbLeftJustify
        '.Add "CUSTNAME", , , , , "�Ȥ�m�W" & Space(10), vbLeftJustify
        '.Add "CUSTTEL", , , , , "�q��" & Space(15), vbLeftJustify
        '.Add "CUSTADDRESS", , , , , "�˾��a�}" & Space(50), vbLeftJustify
        .Add "DEVICESNO1", , , , , "HFCMAC" & Space(15), vbLeftJustify
        .Add "DEVICESNO2", , , , , "MATMAC" & Space(15), vbLeftJustify
        
        
        
        .Add "ACCOUNTID", , , , , "CM �b��" & Space(10), vbLeftJustify
        .Add "ACCOUNTPASSWORD", , , , , "CM �K�X" & Space(10), vbLeftJustify
        .Add "SCHEMACODE1", , , , , "�A�Ȥ�ץN�X", vbLeftJustify
        .Add "SCHEMACODE2", , , , , "���ѥN�X", vbLeftJustify
        .Add "USERIDNO", , , , , "�ӽФH�����Ҹ�", vbLeftJustify
        .Add "USERNAME", , , , , "�ӽФH�m�W" & Space(10), vbLeftJustify
        .Add "USERBIRTHDAY", giControlTypeDate, , , , "�ӽФH�ͤ�" & Space(20), vbLeftJustify
        .Add "DEVICETYPE", , , , , "�]�ƫ���", vbLeftJustify
        .Add "DEVICEMODELNO", , , , , "�����N�X", vbLeftJustify
        .Add "CANCELCODE", , , , , "�h���]�N�X", vbLeftJustify
        .Add "CANCELNAME", , , , , "�h���]����" & Space(15), vbLeftJustify
        .Add "CANCELDATE", giControlTypeTime, , , , "�����ɶ�" & Space(30), vbLeftJustify
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
'#4276 �P�_�O���O�ѦҬ�8��,�p�G�O8���N�O�ݩ�Fttx.�䥦���O�n����
Private Sub CtlEnabled()
  On Error GoTo ChkErr:
    Dim strSQLRef As String
    Dim ctl As Control
    strSQLRef = "Select Count(*) From " & GetOwner & "CD022 Where RefNo=8 And CodeNo=" & rsData("FaciCode")
    If gcnGi.Execute(strSQLRef)(0) > 0 Then
        blnUseFttx = True
    Else
        blnUseFttx = False
    End If
    '#4376 �p�G�OFTTX�]��,OPTIONAL�r�˻P��m�n���վ� By Kin 2009/02/16
    If blnUseFttx Then
        optQueryInfo(0).Visible = False
        optQueryInfo(1).Top = lngCMD30Top
        optQueryInfo(1).Caption = Mid(strCMD31String, InStr(1, strCMD31String, ".") + 1)
    Else
        optQueryInfo(0).Visible = True
        optQueryInfo(0).Top = lngCMD30Top
        optQueryInfo(0).Caption = strCMD30String
        optQueryInfo(1).Top = lngCMD31Top
        optQueryInfo(1).Caption = strCMD31String
    End If
    
    For Each ctl In Me.Controls()
        If TypeOf ctl Is OptionButton Then
            ctl.Value = 0
        End If
    Next
    Exit Sub
ChkErr:
  ErrSub Name, "CtlEnabled"
End Sub
Private Sub cmdSet_Click()
    On Error GoTo ChkErr
    Dim lngBR As Long
    Dim intLoop As Integer
    Dim intIndex As Integer
    Dim rsVirtual As New ADODB.Recordset
    Dim strResult As String, strFaultReason As String, strErrMsg As String
    blnProcessOk = False
    frmProcessOk = False
    If pic1.Visible = True Then Exit Sub
    If Not IsDataOk Then
        Exit Sub
    End If
    Me.Enabled = False
    blnSetting = True
    If sstData.Tab = 0 Then
        If optBasic(8).Value = True Then
            Call cmdEmergency_Click
            Exit Sub
        End If
    End If
    
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
                     '�}��
                    Case optBasic(0)
                        If Not OpenCM(rsData, "11", rsData("FaciSNo") & "", , strSNO, , gdtResvTime.GetValue, , , strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                        Call OpenData2
                        
                    '����
                    Case optBasic(1)
                        If Not OpenCM(rsData, "12", rsData("FaciSNo") & "", , strSNO, , gdtResvTime.GetValue, _
                                        IIf(gilReason.GetCodeNo <> "", gilReason.GetCodeNo, Empty), IIf(gilReason.GetCodeNo <> "", gilReason.GetDescription, Empty), _
                                        strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                        Call OpenData2
                    '����
                    Case optBasic(2)
                        If Not OpenCM(rsData, "13", rsData("FaciSNo") & "", , strSNO, , gdtResvTime.GetValue, _
                                        IIf(gilReason.GetCodeNo <> "", gilReason.GetCodeNo, Empty), IIf(gilReason.GetCodeNo <> "", gilReason.GetDescription, Empty), _
                                        strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                        OpenData2
                    '�_��
                    Case optBasic(3)
                        If Not OpenCM(rsData, "14", rsData("FaciSNo") & "", , strSNO, , gdtResvTime.GetValue, strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                        OpenData2
                    '�˾�
                    Case optBasic(4)
                        If Not InstCM(rsData, "10", strCustName, strCustTel, strInstAddress, _
                                        rsData("CMBaudNo") & "", rsData("VendorCode") & "", _
                                        IIf(IsCMCP(rsData), Get_EMTA_Mac(rsData("FaciSNo")), ""), strSNO _
                                        , , gdtResvTime.GetValue, rsVirtual, strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                        OpenData2
                    '���
                    Case optBasic(5)
                        If Not OpenCM(rsData, "15", rsData("FaciSNo") & "", IIf(IsCMCP(rsData), Get_EMTA_Mac(rsData("FaciSNo") & ""), ""), strSNO, , gdtResvTime.GetValue, _
                                        IIf(gilReason.GetCodeNo <> "", gilReason.GetCodeNo, Empty), IIf(gilReason.GetCodeNo <> "", gilReason.GetDescription, Empty), _
                                        strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                        OpenData2
                    '���״���
                    Case optBasic(6)
                        If Not OpenCM(rsData, "16", txtFaciSNo.Text, IIf(IsEMTA(gilFaciCode.GetCodeNo) = True, Get_EMTA_Mac(txtFaciSNo.Text), ""), strSNO, , gdtResvTime.GetValue, , , strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                        OpenData2
                    
                    '#3752 �s�W�R�O By Kin 2008/02/01
                    '���m�]��
                    Case optBasic(7)
                        If Not OpenCM(rsData, "17", rsData("FaciSNo") & "", , strSNO, , gdtResvTime.GetValue, , , strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                        Call OpenData2
                    '#4276 �w��FTTX�˾� By Kin 2008/12/23
                    Case optBasic(9)
                        If Not RequirePPPOEAccount(rsData, "18", strSNO, , gdtResvTime.GetValue, strDynIpCount, strFixIPCount, rsVirtual, strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                        Call OpenData2
                    '#4276 �����w��FTTX�˾� By Kin 2008/12/23
                    Case optBasic(10)
                        If Not ReleasePPPOEAccount(rsData, "19", rsData("FaciSNo") & "", strSeqNo, , gdtResvTime.GetValue, strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                        Call OpenData2
                    '#4276 �ҥ�PPPOE�b�� By Kin 2008/12/24
                    Case optBasic(11)
                        If Not OpenPPPOEAccount(rsData, "10", strCustName, strCustTel, strInstAddress, gdtResvTime.GetValue, strSNO, , strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                        Call OpenData2
                    '#4276 �W�[�P�ϲ��J By Kin 2008/12/30
                    Case optBasic(12)
                        If Not OpenCM(rsData, "29", rsData("FaciSNo") & "", IIf(IsEMTA(rsData("FaciCode") & "") = True, Get_EMTA_Mac(rsData("FaciSNo") & ""), ""), strSNO, , gdtResvTime.GetValue, , , strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                        OpenData2
                    '#4384 �W�[Release Port�R�O By Kin 2009/02/26
                    Case optBasic(13)
                        If Not OpenCM(rsData, "33", rsData("FaciSNo") & "", IIf(IsCMCP(rsData), Get_EMTA_Mac(rsData("FaciSNo") & ""), ""), strSNO, , gdtResvTime.GetValue, _
                                        IIf(gilReason.GetCodeNo <> "", gilReason.GetCodeNo, Empty), IIf(gilReason.GetCodeNo <> "", gilReason.GetDescription, Empty), _
                                        strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                        OpenData2
                    
                    Case Else
                        MsgBox "�п�ܩR�O�I", vbInformation, "�T��"
                    
                End Select
            Case 1
                Select Case True
                    '�ܧ�򥻸��
                    Case optCMChange(0)
                        
                        If Not ChangeData(rsData, "20", strCustName, strCustTel, strInstAddress, , strSNO, , gdtResvTime.GetValue, strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                    Case optCMChange(1)
                    '�ܧ�ӽФH
                        If Not ChangeCust(rsData, "21", strCustName, strCustTel, strInstAddress, IIf(blnUseFttx, Empty, rsData("CMBaudNo") & ""), IIf(blnUseFttx, Empty, rsData("VendorCode") & ""), _
                                        IIf(blnUseFttx, Empty, IIf(IsCMCP(rsData), Get_EMTA_Mac(rsData("FaciSNo")), "")), strSNO, , gdtResvTime.GetValue, rsData("DialAccount") & "", _
                                        IIf(blnUseFttx, rsData("DialPassword") & "", Empty), rsData("ID") & "", _
                                        rsData("DeclarantName") & "", rsData("Birthday") & "", rsVirtual, strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                                                                
                    '�ܧ�t�v
                    Case optCMChange(2)
                        If Not ChangeRate(rsData, "22", gilBaudRate.GetCodeNo, rsData("VendorCode") & "", , gilBaudRate.GetCodeNo, _
                                            gilBaudRate.GetDescription, strSNO, , gdtResvTime.GetValue, , _
                                            strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                                            
                    '�ܧ����
                    Case optCMChange(3)
                        If Not ChangeRate(rsData, "23", rsData("CMBaudNo") & "", gilVendor.GetCodeNo, gilVendor.GetDescription, _
                        , , strSNO, , gdtResvTime.GetValue, IIf(blnUseFttx, rsVirtual, Nothing), strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                        
                        
                    '�ܧ�K�X
                    Case optCMChange(4)
                        If Not ChangePassword(rsData, "24", txtPassword, strSNO, , gdtResvTime.GetValue, strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                        
                    '#3752 �s�W�R�O By Kin 2008/02/01
                    '�T�{�ϥΤH
                    Case optCMChange(5)
                        If Not ConfirmCust(rsData, "25", strCustName, strCustTel, strInstAddress, strSNO, , gdtResvTime.GetValue, strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                    '#4276 �ӽЩT�wIP By Kin 2008/12/24
                    Case optCMChange(6)
                        If Val(strFixIPCount) <= 0 Then
                            MsgBox "�T�wIP�Ƭ�0", vbCritical, "�T��"
                            GoTo ErrTran
                        Else
                            If Not RequireFixIP(rsData, "26", strSNO, , gdtResvTime.GetValue, strDynIpCount, strFixIPCount, rsVirtual, strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                        End If
                    '#4276 �����ӽЩT�wIP By Kin 2008/12/26
                    Case optCMChange(7)
                        If Not ReleaseFixIP(rsData, "27", strSNO, , gdtResvTime.GetValue, , strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                    '#4276 ���s�]�mPPPOE�K�X By Kin 2008/12/29
                    Case optCMChange(8)
                        If Not ResetPPPOEPassowrd(rsData, "28", strSNO, , gdtResvTime.GetValue, strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                    Case Else
                        MsgBox "�п�ܩR�O�I", vbInformation, "�T��"
                    
                End Select
                Call EnableMode(True)
            Case 2
                Select Case True
                    '�d��CM��T
                    Case optQueryInfo(0)
                        If Not QueryCMStatus(rsData, "30", IIf(blnUseFttx, Empty, Get_EMTA_Mode(rsData("FaciSNo"))) & "", rsData("FaciSNo") & "", , gdtResvTime.GetValue, rsVirtual, strResult, strFaultReason, strErrMsg) Then
                            GoTo ErrTran
                        Else
                            If Not rsVirtual Is Nothing Then
                                Show_Rs_Result rsVirtual
                            End If
                        End If
                    '�d�߱b����T
                    '#4276 �p�G�OFTTX�n�h�ǳ]�Ƭy���� By Kin 2008/12/25
                    Case optQueryInfo(1)
                        If Not QueryCMStatus(rsData, "31", , IIf(blnUseFttx, rsData("FaciSNo") & "", Empty), rsData("DialAccount") & "", _
                                            gdtResvTime.GetValue, rsVirtual, strResult, _
                                            strFaultReason, strErrMsg) Then
                            GoTo ErrTran
                        Else
                            If Not rsVirtual Is Nothing Then
                                Show_Rs_Result rsVirtual
                            End If
                        End If
                    Case Else
                        MsgBox "�п�ܩR�O�I", vbInformation, "�T��"
                End Select
            Case 3
                Select Case True
                    '�]�ƤJ�w
                    Case optCMDevice(0)
                        If Not StockCM(rsData, "40", Get_EMTA_Mode(rsData("FaciSNo")) & "", rsData("ModelCode") & "", rsData("FaciSNo"), _
                                        Get_EMTA_Mac(rsData("FaciSNo")), strSNO, , gdtResvTime.GetValue, _
                                        strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                    '�]�ƥX�w
                    Case optCMDevice(1)
                        If Not StockCM(rsData, "41", , , rsData("FaciSNo") & "", , strSNO, , gdtResvTime.GetValue, _
                                        strResult, strFaultReason, strErrMsg) Then GoTo ErrTran
                    
                    Case Else
                        MsgBox "�п�ܩR�O�I", vbInformation, "�T��"
                End Select
        End Select
66:
        blnSetting = False
        If Not blnTransation Then gcnGi.CommitTrans
        rsData.AbsolutePosition = lngBR
        
        blnProcessOk = True
        frmProcessOk = True
        Call subProgressBar(False)
        Me.Enabled = True
        'Show_Rs_Result rsVirtual
        If intProcessType > 0 Then Unload Me
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

Private Sub Show_Rs_Result(ByRef rsResult As ADODB.Recordset)
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    Dim intLoop As Long
    ggrQueryInfo.Blank = True
    If rsResult.State = adStateClosed Then Exit Sub
    With rsResult
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        For intLoop = 0 To .Fields.Count - 1
            SetGrdFld mFlds, .Fields(intLoop).Name
        Next intLoop
    End With
    With ggrQueryInfo
        .AllFields = mFlds
        Set .Recordset = rsResult
        .Refresh
        .Blank = False
    End With
    
  Exit Sub
ChkErr:
    ErrSub Name, "Show_Rs_Result"
End Sub

Private Sub Show_XML_Result(ByRef rsResult As ADODB.Recordset)
  On Error GoTo ChkErr
    
    Dim mFlds As New GiGridFlds
    Dim intLoop As Integer
    
    ggrQueryInfo.Blank = True
    If rsResult.State = adStateClosed Then Exit Sub
    With rsResult
            If .RecordCount = 0 Then Exit Sub
            .MoveFirst
            For intLoop = 0 To .Fields.Count - 1
                SetGrdFld mFlds, .Fields(intLoop).Name
            Next
    End With
    
    With ggrQueryInfo
            .AllFields = mFlds
            .SetHead
            Set .Recordset = rsResult
            .Refresh
            '.Blank = False
    End With
    
  Exit Sub
ChkErr:
    ErrSub Name, "Show_XML_Result"
End Sub
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
            Case 3
                Set objControl = optCMDevice
        End Select
        For intLoop = 0 To objControl.UBound
            If objControl(intLoop) Then
                strDoCaption = Mid(objControl(intLoop).Caption, 4)
            End If
        Next
        lblProcess.Caption = strDoCaption & " �� , �еy�� !!"

End Sub

Private Sub Form_Activate()
  On Error Resume Next
    '#3778 �۰���ܳ]�Ƹ�� By Kin 2008/02/25
    If blnShowFaci Then
        blnSendKey = False
        sstHead.Tab = 1
        If cmdQuery.Enabled Then cmdQuery.Value = True
    End If
    blnShowFmt = True
  '#3778�P�_�O�_�n�۰ʫ��UF2 By Kin 2008/02/22
    If blnSendKey Then
        If cmdSet.Enabled Then cmdSet.Value = True
    End If
    blnSendKey = False
    blnShowFaci = False
   
End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr
        blnShowFmt = False
        gdtResvTime.ShowSecond
        blnExeCommand = True '����R�O���P�_(��UI���@�w�n����A�ҥH�w�]��True
        sstData.Tab = 0
        '*******************************************************
        '#3611 ���դ�OK,�N�]�ƥX�J�w���� By Kin 2008/01/10
        sstData.TabVisible(3) = False
        '*******************************************************
        
        Call subGil
        Call subGrd
        Call DefaultValue
        Call OpenData
        Call DoProcessMode
        
        'blnBpCode = True
        Fetch_Log_Data , True
        If blnAutoSend Then
            blnShowMsg = False
        Else
            blnShowMsg = True
        End If
        '#3778 �W�[�@�ղ��ʭ�]�A����bsubGil�̡A�]���n�L�oRecordSet��ServiceType By Kin 2008/02/26
        If Not rsData.EOF Then
            SetgiList gilReason, "CodeNo", "Description", "CD014", , , , , 3, 20, " Where ServiceType is Null Or ServiceType='" & rsData("ServiceType") & "'", True
            '#3800 �]�w�O���h�W�[�@�Ӱh���] By Kin 2008/03/12
            SetgiList gilReturnCode, "CodeNo", "Description", "CD015", , , , , 3, 20, " Where (ServiceType is Null Or ServiceType='" & rsData("ServiceType") & "') And RefNo=3", True
        Else
            'MsgBox "�L�����ƥi����!!", vbCritical, "�T��"
        End If
        gdtResvTime.SetValue strResvTime
         '#7346 to fix the record's position move to the wrong position after executed command,because the custid validate triggered the opendata event. by kin
        'cmdSet.SetFocus
        'blnShowFmt = True
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

Private Sub DoProcessMode()
    On Error GoTo ChkErr
    Dim objControl As Object
    Dim blnFlag As Boolean
    Dim intTab As Integer
    Dim intLoop As Integer
    Dim intIndex As Integer
        
        If strDefaultRowId <> "" Then RsFind rsData, "RowId", strDefaultRowId
        If intProcessType <= 0 Then Exit Sub
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
        sstData.TabEnabled(3) = False
        '�ϥΨ��ӫ��O(�̥k���X)

         intIndex = Val(intProcessType)
         If intIndex >= 10 And intIndex < 20 Then
            sstData.Tab = 0
            sstData.TabEnabled(0) = True
         ElseIf intIndex >= 20 And intIndex <= 28 Then
            sstData.Tab = 1
            sstData.TabEnabled(1) = True
         ElseIf intIndex = 29 Then
            sstData.Tab = 0
            sstData.TabEnabled(0) = True
         '#4384 �W�[ Release Port�R�O By Kin 2009/02/26
         ElseIf intIndex = 33 Then
            sstData.Tab = 0
            sstData.TabEnabled(0) = True
         Else
            sstData.Tab = 2
            sstData.TabEnabled(2) = True
         End If
         Select Case intIndex
            Case 10
                If blnUseFttx Then
                    optBasic(11).Enabled = True
                    optBasic(11).Value = True
                Else
                    optBasic(4).Enabled = True
                    optBasic(4).Value = True
                End If
            Case 11
                optBasic(0).Enabled = True
                optBasic(0).Value = True
            Case 12
                optBasic(1).Enabled = True
                optBasic(1).Value = True
            Case 13
                optBasic(2).Enabled = True
                optBasic(2).Value = True
            Case 14
                optBasic(3).Enabled = True
                optBasic(3).Value = True
            Case 15
                optBasic(5).Enabled = True
                optBasic(5).Value = True
            Case 16
                optBasic(6).Enabled = True
                optBasic(6).Value = True
            Case 17
                optBasic(7).Enabled = True
                optBasic(7).Value = True
            Case 18
                optBasic(9).Enabled = True
                optBasic(9).Value = True
            Case 19
                optBasic(10).Enabled = True
                optBasic(10).Value = True
            Case 20
                optCMChange(0).Enabled = True
                optCMChange(0).Value = True
            Case 21
                optCMChange(1).Enabled = True
                optCMChange(1).Value = True
            Case 22
                optCMChange(2).Enabled = True
                optCMChange(2).Value = True
            Case 23
                optCMChange(3).Enabled = True
                optCMChange(3).Value = True
            Case 24
                optCMChange(4).Enabled = True
                optCMChange(4).Value = True
            Case 25
                optCMChange(5).Enabled = True
                optCMChange(5).Value = True
            Case 26
                optCMChange(6).Enabled = True
                optCMChange(6).Value = True
            Case 27
                optCMChange(7).Enabled = True
                optCMChange(7).Value = True
            Case 28
                optCMChange(8).Enabled = True
                optCMChange(8).Value = True
            Case 29
                optBasic(12).Enabled = True
                optBasic(12).Value = True
            Case 30
                optQueryInfo(0).Enabled = True
                optQueryInfo(0).Value = True
            Case 31
                optQueryInfo(1).Enabled = True
                optQueryInfo(1).Value = True
            '#4384 �W�[Release Port �R�O By Kin 2009/02/26
            Case 33
                optBasic(13).Enabled = True
                optBasic(13).Value = True
         End Select
         
'        intLoop = Val(Left(intProcessType, 1))
'        sstData.Tab = intLoop - 1
'        sstData.TabEnabled(intLoop - 1) = True
'        fraData(intLoop - 1).Enabled = True
'        Select Case intLoop
'            Case 1
'                Select Case intIndex
'                    Case 10
'                        If blnUseFttx Then
'                            optBasic(11).Enabled = True
'                            optBasic(11).Value = True
'                        Else
'                            optBasic(4).Enabled = True
'                            optBasic(4).Value = True
'                        End If
'                    Case 15
'                        optBasic(5).Enabled = True
'                        optBasic(5).Value = True
'                    Case 16
'                        optBasic(6).Enabled = True
'                        optBasic(6).Value = True
'                    Case 17
'                        optBasic(7).Enabled = True
'                        optBasic(7).Value = True
'                    Case 19
'                        optBasic(9).Enabled = True
'                        optBasic(9).Value = True
'                    Case Else
'                        optBasic(intIndex - 11).Enabled = True
'                        optBasic(intIndex - 11).Value = True
'                End Select
'            Case 2
'                optCMChange(intIndex - 20).Enabled = True
'                optCMChange(intIndex - 20).Value = True
'            Case 3
'                optQueryInfo(intIndex - 30).Enabled = True
'                optQueryInfo(intIndex - 30).Value = True
'        End Select
        
        sstHead.Tab = 0
        ggrData.Enabled = False
        
    Exit Sub
ChkErr:
    ErrSub Me.Name, "DoProcessMode"
End Sub

Private Sub subGil()
    On Error GoTo ChkErr
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 20, GetCompCodeFilter(gilCompCode, GaryGi(15))
        'SetgiList gilBP, "CodeNo", "Description", "CD078", , , , , 3, 20, , True
        SetgiList gilFaciCode, "CodeNo", "Description", "CD022", , , , , 3, 12, " WHERE REFNO IN (2,5)", True
        SetgiList gilBaudRate, "CodeNo", "Description", "CD064", , , , , 3, 20, GetCompCodeFilter(gilCompCode, gilCompCode.GetCodeNo), True
        SetgiList gilReason, "CodeNo", "Description", "CD014", , , , , 3, 20, , True
        SetgiList gilVendor, "VendorCode", "VendorName", "CM006", , , , , 3, 20, " Where CompCode=" & gCompCode, True
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
        With mFlds
            .Add "FaciSNo", , , , , "�]�ƧǸ�                 ", vbLeftJustify
            .Add "FaciCode", , , , , "�N�X ", vbLeftJustify
            .Add "FaciName", , , , , "���ئW�� ", vbLeftJustify
            .Add "ModelName", , , , , "�����W�� ", vbLeftJustify
            .Add "CMBaudNo", , , , , "�N�X", vbLeftJustify
            .Add "CMBaudRate", , , , , "CM�t�v ", vbLeftJustify
            .Add "DynIPCount", , , , , "�ʺAIP�ƥ�", vbRightJustify
            .Add "FixIPCount", , , , , "�T�wIP�ƥ�", vbRightJustify
            .Add "InstDate", giControlTypeDate, , , , "�˾����", vbLeftJustify
            .Add "CMOpenDate", giControlTypeTime, , , , "�}�����         ", vbLeftJustify
            .Add "CMCloseDate", giControlTypeTime, , , , "�������         ", vbLeftJustify
            .Add "DialAccount", , , , , "�����b�� ", vbLeftJustify
            .Add "EMail", , , , , "�q�l�H�c           ", vbLeftJustify
        End With
        ggrData.AllFields = mFlds
        ggrData.SetHead
        

        '�W�D�s��,�W�D�W��,���O�_�l���,�I����
        With mFlds2
            .Add "CustId", , , , False, "�Ȥ�s��", vbLeftJustify
            .Add "Citemname", , , , False, "    ���O����    ", vbLeftJustify
            .Add "ShouldAmt", , , , False, "�X����B", vbRightJustify
            .Add "RealPeriod", , , , False, "����", vbLeftJustify
            .Add "RealAmt", , , , False, "�ꦬ���B", vbRightJustify
            .Add "RealStartDate", , , , False, "�����_�l��", vbLeftJustify
            .Add "RealStopDate", , , , False, "�����I���", vbLeftJustify
            .Add "CancelFlag", , , , False, "�@�o", vbLeftJustify
            .Add "Note", , , , False, Space(20) & "  �Ƶ�            " & Space(15), vbLeftJustify
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
        frmProcessOk = False
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
        Set rsLog = New ADODB.Recordset
        If rsParent Is Nothing Then Set rsParent = New ADODB.Recordset
        If rsParent.State = adStateOpen Then
            Set rsData = rsParent.Clone
            If blnShowFaci Then
                rsData.Filter = " PRDate = Null"
            End If
            If Len(strDefaultRowId & "") = 0 Then
                strDefaultRowId = rsParent("RowId") & ""
            End If
           
            'rsData.AbsolutePosition = rsParent.AbsolutePosition
        Else
            If ginCustId.Text <> "" Then
                '#4276 �W�[�Ѧ�8,�N��FTTX By Kin 2008/12/23
                strSQL = " A.CustId = " & ginCustId.Value & " And A.CompCode = " & gilCompCode.GetCodeNo & " And A.FaciSNo is not null And FaciCode In (Select CodeNo From " & GetOwner & "CD022 Where RefNo in (2,5,8)) And A.PRDate is null "
                If strWhere <> "" Then strSQL = strWhere
            Else
                strSQL = " A.SeqNo = ''"
            End If
            If Not GetRS(rsData, "Select A.RowId, A.* From " & GetOwner & "So004 A Where " & strSQL & " Order By A.InstDate", , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
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
        If intProcessType = 0 Then Call EnableMode(True)
        Set ggrChild.Recordset = rsChild
        ggrChild.Refresh
    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData2"
End Sub

Private Sub EnableMode(blnFlag As Boolean)
    On Error Resume Next
        If Not blnFlag Then Exit Sub
        optBasic(4).Enabled = GetPremission("SO11731") = Not blnUseFttx ' 1. �˾�
        optBasic(0).Enabled = GetPremission("SO11732") ' 2. �}��
        optBasic(1).Enabled = GetPremission("SO11733") ' 3. ����
        optBasic(2).Enabled = GetPremission("SO11734") = Not blnUseFttx ' 4. ����
        optBasic(3).Enabled = GetPremission("SO11735") = Not blnUseFttx ' 5. �_��
        optBasic(5).Enabled = GetPremission("SO11736")  ' 6. ���
        optBasic(6).Enabled = GetPremission("SO11737") ' 7. ���״���
        optBasic(7).Enabled = GetPremission("SO1173I") = Not blnUseFttx ' 8. ���m�]��
        optBasic(8).Enabled = GetPremission("SO1173H") = Not blnUseFttx ' 9. ���}�q
    
        '************************************************************************
        '#4276 �W�[FTTX���O By Kin 2008/12/23
        optBasic(9).Enabled = GetPremission("SO1173L") = blnUseFttx ' 10.�w��FTTX�˾�
        optBasic(10).Enabled = GetPremission("SO1173M") = blnUseFttx '11.�����w��FTTX�˾�
        optBasic(11).Enabled = GetPremission("SO1173N") = blnUseFttx '12�ҥ�PPPOE�b��
        optBasic(12).Enabled = GetPremission("SO1173R") = blnUseFttx '29 �P�ϲ��J
        '#4384 �W�[ Release Port By Kin 2009/02/26
        optBasic(13).Enabled = GetPremission("SO1173S") = blnUseFttx ' 33 Release Port
        '************************************************************************
        
        'cmdEmergency.Enabled = GetPremission("SO1173H") ' 8. ���}�q
        optCMChange(0).Enabled = GetPremission("SO11738") ' 1. �ܧ�򥻸��
        optCMChange(1).Enabled = GetPremission("SO11739") ' 2. �ܧ�ӽФH
        optCMChange(2).Enabled = GetPremission("SO1173A") ' 3. �ܧ�t�v
        optCMChange(3).Enabled = GetPremission("SO1173B") ' 4. �ܧ����
        optCMChange(4).Enabled = GetPremission("SO1173C") = Not blnUseFttx ' 5. �ܧ�K�X
        optCMChange(5).Enabled = GetPremission("SO1173J") ' 6. �T�{�ϥΤH
        '**************************************************************************
        '#4276 �W�[FTTX���O By Kin 2008/12/23
        optCMChange(6).Enabled = GetPremission("SO1173O") = blnUseFttx '�ӽЩT�wIP
        optCMChange(7).Enabled = GetPremission("SO1173P") = blnUseFttx '�����ӽЩT�wIP
        optCMChange(8).Enabled = GetPremission("SO1173Q") = blnUseFttx '���]PPPOE�K�X
        '**************************************************************************
        
        optQueryInfo(0).Enabled = GetPremission("SO1173D") ' 1. �d��CM��T
        optQueryInfo(1).Enabled = GetPremission("SO1173E") ' 2. �d�� �b����T
        
        optCMDevice(0).Enabled = GetPremission("SO1173F") ' 1. �]�ƤJ�w
        optCMDevice(1).Enabled = GetPremission("SO1173G") ' 1. �]�ƥX�w
        '#3800 �W�[�w�����O�����v�� By Kin 2008/03/12
        cmdCancelCMD.Enabled = GetPremission("SO1173K")
        blnUseCancelCmd = True
        '#4276 �n�b�P�_�ϥ�FTTX�Ϊ̬OCM���O By Kin 2008/12/23
        'Call CtlEnabled
        '#7065 SO004.PRSNO IS NOT NULL �����\�e�ܧ�ӽФH"(so314.msgcode=21)��"�ܧ�򥻸��" (so314.msgcode=20) By Kin 2015/08/06
        If Not rsData Is Nothing Then
            If rsData.RecordCount > 0 Then
                If Len(rsData("PRSNO") & "") > 0 Then
                    optCMChange(0).Enabled = False
                    optCMChange(1).Enabled = False
                    optCMChange(0).Value = 0
                    optCMChange(1).Value = 0
                End If
            End If
        End If
      

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
    Select Case UCase(giFld.FieldName)
        Case "MODETYPE"
            Value = GetModeTypeValue(Value)
        Case "CMDSTATUS"
            If UCase(Value) & "" = "W" Then
                Value = "���B�z"
            ElseIf UCase(Value) & "" = "P" Then
                Value = "�B�z��"
            ElseIf UCase(Value) & "" = "E" Then
                Value = "�����~"
            ElseIf UCase(Value) & "" = "C" Then
                Value = "���T"
            Else
                Value = ""
            End If
        Case "STOPFLAG"
            If CStr(Value) = "1" Then
                Value = "�O"
            Else
                Value = "�_"
            End If
    End Select
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error Resume Next
        If giFld.FieldName = "ReFaciSNo" Then
            If Value = "" Then
                Value = "����"
            Else
                Value = "�l��"
            End If
        End If
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If gilCompCode.GetCodeNo = "" Then Exit Sub
        GaryGi(16) = gilCompCode.GetOwner
        '#3731 ��CMType=3�ɤ��P�_�զX���~ By Kin 2008/01/15
        If GetRsValue("Select NVL(CMType,0) From " & GetOwner & "CD039 Where CodeNo=" & gilCompCode.GetCodeNo) = 3 Then
            blnBpCode = False
            intCMType = 3
        Else
            blnBpCode = True
        End If
        
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
    End Select
End Sub


Private Sub optBasic_Click(Index As Integer)
  On Error Resume Next
    picCPno.Visible = False
    HideResv
    Select Case Index
'                Case 0 ' 1. �}�� [02]
'                    If optBasic(0).Value Then
'                        chkAccStop.Value = 0
'                        HideResv
'                        ShowResv
'                    End If
'                    chkPR.Enabled = False
'                Case 1 ' 2. ���� [02]
'                    If optBasic(1).Value Then
'                        ChkAccStart.Value = 0
'                        HideResv
'                        ShowResv
'                    End If
'                    chkPR.Enabled = False
                '���״���
                Case 6
                    If optBasic(6).Value Then
                        '********************************************************************************************************
                        '#4276 �p�G�OFTTX�u�����FTTX���]�� By Kin 2009/01/09
                        If blnUseFttx Then
                            SetgiList gilFaciCode, "CodeNo", "Description", "CD022", , , , , 3, 12, " WHERE REFNO IN (8)", True
                        Else
                            SetgiList gilFaciCode, "CodeNo", "Description", "CD022", , , , , 3, 12, " WHERE REFNO IN (2,5)", True
                        End If
                        '*******************************************************************************************************
                        gilFaciCode.SetCodeNo ""
                        gilFaciCode.SetDescription ""
                        txtFaciSNo.Text = ""
                        lblFaci.Visible = True
                        lblCMFacisno.Visible = True
                        gilFaciCode.Visible = True
                        txtFaciSNo.Visible = True
                        
                    End If
                    If strFaciCode <> "" Then
                        gilFaciCode.SetCodeNo strFaciCode
                        gilFaciCode.Query_Description
                        gilFaciCode.Enabled = False
                    End If
                    If strFaciSNo <> "" Then
                        txtFaciSNo.Text = strFaciSNo
                        txtFaciSNo.Enabled = False
                    End If
                '***********************************************
                '#3778 �W�[���ʭ�] By Kin 2008/02/25
                Case 1, 2, 5
                    gilReason.Visible = True
                    lblReason.Visible = True
                    If strReason <> "" Then
                        gilReason.SetCodeNo strReason
                        gilReason.Query_Description
                        gilReason.Enabled = False
                    Else
                        gilReason.SetCodeNo ""
                        gilReason.SetDescription ""
                    End If
                '***********************************************
                
                Case Else
                    HideResv
                    chkPR.Enabled = False
    End Select
End Sub

Private Sub ShowResv()
  On Error Resume Next
    'lblResv.Visible = True
    'gdtResvTime.Visible = True
    gilBaudRate.Visible = True
End Sub

Private Sub HideResv()
  On Error Resume Next
'    lblResv.Visible = False
    lblFaci.Visible = False
    lblCMFacisno.Visible = False
    gilFaciCode.Visible = False
    gilVendor.Visible = False
    txtFaciSNo.Visible = False
    txtPassword.Visible = False
'    gdtResvTime.Visible = False
    gilBaudRate.Visible = False
    gilReason.Visible = False
    lblReason.Visible = False
'    gdtResvTime.Text = ""
End Sub

Private Sub optCMChange_Click(Index As Integer)
    HideResv
    gdtResvTime.Enabled = True
    Select Case True
        Case optCMChange(2)
            If strBaudRate <> Empty Then
                gilBaudRate.SetCodeNo strBaudRate
                gilBaudRate.Query_Description
                gilBaudRate.Enabled = False
            Else
                gilBaudRate.SetCodeNo ""
                gilBaudRate.SetDescription ""
            End If
            ShowResv
        Case optCMChange(3)
            gilVendor.Visible = False
            gilVendor.Visible = True
            gilVendor.Enabled = True
            If strVendor <> Empty Then
                gilVendor.SetCodeNo strVendor
                gilVendor.Query_Description
                gilVendor.Enabled = False
            Else
                gilVendor.SetCodeNo ""
                gilVendor.SetDescription ""
            End If
            'ShowResv
        Case optCMChange(4)
            '#3800 ��ܱK�X�ܧ�ɭn�N�w���ɶ��M��,�ä����ާ@ By Kin 2008/03/12
            gdtResvTime.Text = ""
            gdtResvTime.Enabled = False
            txtPassword.Visible = False
            txtPassword.Visible = True
            txtPassword.Enabled = True
            If strPassword <> Empty Then
                txtPassword.Text = strPassword
                txtPassword.Enabled = False
            Else
                txtPassword.Text = ""
            End If
        Case Else
            HideResv
    End Select
End Sub

Private Sub rsData_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
        '#4276 �P�_�O�ϥ�FTTX�٬OCM By Kin 2008/12/23
        '���դ�OK..�n����@�Ǳ������ By Kin 2009/01/13
        Dim strSQLRef As String
        If intProcessType > 0 Then
            If Not rsData.EOF Then
                strSQLRef = "Select Count(*) From " & GetOwner & "CD022 Where RefNo=8 And CodeNo=" & rsData("FaciCode")
                If gcnGi.Execute(strSQLRef)(0) > 0 Then
                    blnUseFttx = True
                Else
                    blnUseFttx = False
                End If
            End If
            Exit Sub
        End If
        HideResv
        If rsData.RecordCount > 0 Then
            Call CtlEnabled
        End If
        If Not ggrData.Enabled Then Exit Sub
        If Not blnSetting Then
            Me.Enabled = False
            Call OpenData2
            Fetch_Log_Data , True
            Me.Enabled = True
        End If
        '#7065 SO004.PRSNO IS NOT NULL �����\�e�ܧ�ӽФH"(so314.msgcode=21)��"�ܧ�򥻸��" (so314.msgcode=20) By Kin 2015/08/06
        If Len(rsData("PRSNO") & "") > 0 Then
            optCMChange(0).Enabled = False
            optCMChange(1).Enabled = False
            optCMChange(0).Value = 0
            optCMChange(1).Value = 0
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
            If Not GetRS(rs, "Select CustName,InstAddrNo,InstAddress,TELH1,Tel1 From " & GetOwner & "So001 Where CustId = " & ginCustId.Value & " And CompCode = " & gilCompCode.GetCodeNo, gcnGi, , , , , , , True) Then Exit Sub
            If rs.EOF Then
                txtCustName = ""
                strCustName = ""
                lngInstAddrNo = 0
                strInstAddress = ""
                lngNodeNo = 0
                lngCircuitNo = 0
            Else
                txtCustName = rs("CustName") & ""
                strCustName = rs("CustName") & ""
                strInstAddress = rs("InstAddress") & ""
                strCustTel = IIf(rs("TelH1") & "" = "", "", rs("TelH1") & "-") & rs("Tel1") & ""
                lngInstAddrNo = Val(rs("InstAddrNo") & "")
                lngNodeNo = 0
                lngCircuitNo = 0
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
                MsgBox "�w���ɶ��ݤj��{�b�ɶ�", vbCritical, gimsgPrompt
                gdtResvTime.SetValue DateAdd("h", 1, RightNow)
                gdtResvTime.SetFocus
                Exit Function
            Else
                '7676
'                If CDate(gdtResvTime.GetValue(True)) > CDate(RightNow) + 1 Then
'                    MsgBox "�w���ɶ��ݤp�����ɶ�", vbCritical, gimsgPrompt
'                    gdtResvTime.SetValue DateAdd("h", 1, RightNow)
'                    gdtResvTime.SetFocus
'                    Exit Function
'                End If
                If Len(strResvTime & "") > 0 Then
                    If Len(strSNO & "") > 0 Then
                        Dim compareDate As String
                        compareDate = strResvTime
                        If IsDate(compareDate) Then
                            compareDate = Format(compareDate, "yyyyMMddHHmmss")
                        End If
                        If Not IsDate(compareDate) Then
                                
'                            compareDate = Mid(compareDate, 1, 4) & "/" & Mid(compareDate, 5, 2) & "/" & Mid(compareDate, 7, 2) & " " & Mid(compareDate, 9, 2) & ":" & _
'                                                Mid(compareDate, 11, 2) & ":" & Right(compareDate, 2)
                            compareDate = Mid(compareDate, 1, 4) & "/" & Mid(compareDate, 5, 2) & "/" & Mid(compareDate, 7, 2) & " " & "00" & ":" & _
                                                    "00" & ":" & "00"
                        End If
                        If CDate(gdtResvTime.GetValue(True) & "") >= CDate(compareDate) + 1 Then
                            MsgBox "�w���ɶ��ݤp��u��w�����", vbCritical, gimsgPrompt
                            gdtResvTime.SetValue strResvTime
                            gdtResvTime.SetFocus
                            Exit Function
                        End If
                    End If
                End If
                If Not blnEmergency Then
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
                            Set objControl = optCMChange
                        Case 2
                            Set objControl = optQueryInfo
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
                    'If MsgBox("�w�p" & strShowTime & "�A�N�ǰe " & strExeName & " ���O�A�T�w�ϥιw���\��?", vbQuestion + vbOK, gimsgPrompt) = vbNo Then Exit Function
                     MsgBox "�w�p" & strShowTime & "�A�N�ǰe " & strExeName & " ���O�A�T�w�ϥιw���\��?", vbInformation, "�T��"
                    '�w�p�~:��:��:��:��:��A�N�ǰeXX���O
                End If
            End If
        End If
        
      
        Select Case sstData.Tab
            Case 0
                Select Case True
                    Case optBasic(0)
'                        If IsNull(rsData("CMCloseDate")) Then
'                            MsgBox "�L��������A�L�k���`�ǰe���O !", vbInformation, "�T��"
'                            IsDataOk = False
'                            Exit Function
'                        Else
'                            If Not IsNull(rsData("CMOpenDate")) Then
'                                If rsData("CMOpenDate") > rsData("CMCloseDate") Then
'                                    MsgBox "�}������j����������A�L�k���`�ǰe���O !", vbInformation, "�T��"
'                                    IsDataOk = False
'                                    Exit Function
'                                End If
'                            End If
'                        End If
                    Case optBasic(1)
'                        If rsData("BPCODE") & "" = Empty And blnBpCode Then
'                            MsgBox "���]�ƵL�զX���~���e�A�L�k���`�ǰe���O !", vbInformation, "�T��"
'                            IsDataOk = False
'                            Exit Function
'                        End If
'                        If Not IsNull(rsData("CMCloseDate")) Then
'                            If Not IsNull(rsData("CMOpenDate")) Then
'                                If rsData("CMCloseDate") > rsData("CMOpenDate") Then
'                                    MsgBox "��������j��}������A�L�k���`�ǰe���O !", vbInformation, "�T��"
'                                    IsDataOk = False
'                                    Exit Function
'                                End If
'                            End If
'
'                        End If
                    Case optBasic(2)
                        If rsData("BPCODE") & "" = Empty And blnBpCode Then
                            MsgBox "���]�ƵL�զX���~���e�A�L�k���`�ǰe���O !", vbInformation, "�T��"
                            IsDataOk = False
                            Exit Function
                        End If
                    Case optBasic(6)
                        If Not blnAutoSend Then
                            If gilFaciCode.GetCodeNo = Empty Then
                                MsgBox "�L�~�W���e�A�L�k���`�ǰe���O !", vbInformation, "�T��"
                                IsDataOk = False
                                Exit Function
                            End If
                            If Trim(txtFaciSNo.Text) = "" Then
                                MsgBox "CM �Ǹ��ťաA�L�k���`�ǰe���O !", vbInformation, "�T��"
                                IsDataOk = False
                                Exit Function
                            Else
                                
                                '#3755 chkUseFlag�h�W���q�O�ѼơA�]��CD039.Linkmss=0�ɤ��P�_UseFlag�@�w�n��1 By Kin 2008/02/01
                                If blnUseFttx Then IsDataOk = True: Exit Function
                                If Not chkUseFlag(txtFaciSNo, gilCompCode.GetCodeNo) Then
                                    MsgBox "�]�ƧǸ������T�Τ��i�ϥ� !", vbInformation, "�T��"
                                    IsDataOk = False
                                    Exit Function
                                End If
                            End If
                        End If
                    Case optBasic(8)
'                        If IsNull(rsData("CMCloseDate")) Then
'                            MsgBox "�L��������A�L�k���`�ǰe���O !", vbInformation, "�T��"
'                            IsDataOk = False
'                            Exit Function
'                        Else
'                            If Not IsNull(rsData("CMOpenDate")) Then
'                                If rsData("CMOpenDate") > rsData("CMCloseDate") Then
'                                    MsgBox "�}������j����������A�L�k���`�ǰe���O !", vbInformation, "�T��"
'                                    IsDataOk = False
'                                    Exit Function
'                                End If
'                            End If
'                        End If
                    Case Else
            
                End Select
            Case 1
                Select Case True
                    '#5401 �P�_21�P22���R�O rsData("DialAccount")�p�G�L�Ȥ����\�e�R�O By Kin 2010/01/25
                    Case optCMChange(1)
                        
                    Case optCMChange(2)
                        If gilBaudRate.GetCodeNo = Empty Then
                            MsgBox "�п�ܷs�t�v !", vbInformation, "�T��"
                            IsDataOk = False
                            Exit Function
                        End If
                        If rsData("DialAccount") & "" = "" Then
                            MsgBox "�]�Ʃ|���}�q���\�A�L�k�ǰe���O !", vbInformation, "�T��"
                            IsDataOk = False
                            Exit Function
                        End If
                    Case optCMChange(4)
                        If Trim(txtPassword.Text) = Empty Then
                            MsgBox "�п�J�K�X !", vbInformation, "�T��"
                            IsDataOk = False
                            Exit Function
                        End If
                    Case optCMChange(3)
                        If gilVendor.GetCodeNo = Empty Then
                            MsgBox "�п�ܸ��Ѽt�� !", vbInformation, "�T��"
                            IsDataOk = False
                            Exit Function
                        End If
                        If rsData("DialAccount") & "" = "" Then
                            MsgBox "�]�Ʃ|���}�q���\�A�L�k�ǰe���O !", vbInformation, "�T��"
                            IsDataOk = False
                            Exit Function
                        End If
                    Case Else
                End Select
            
            Case Else
        End Select
        '#5161 �p�GCD022.REFNO=2,8 �h���ˬdSO306����� By Kin 2009/06/26
        If gcnGi.Execute("Select Count(*) From " & GetOwner & "CD022 Where RefNo In(2,8)" & _
                            " And CodeNo=" & rsData("FaciCode"))(0) <= 0 Then
            If IsCMCP(rsData) Then
        
                If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO306" & _
                                     " WHERE HFCMAC='" & rsData("FaciSNo") & "'", gcnGi) = 0 Then
                    MsgBox "SO306 [EMTA�����] ���d�L HFCMAC �� [" & rsData("FaciSNo") & "] �����!", vbInformation, "�T��"
                Else
                    If Get_EMTA_Mac(rsData("FaciSNo") & "") = Empty Then
                        MsgBox "��EMTA�]�ƵLMTAMAC�A�L�k���`�ǰe���O !", vbInformation, "�T��"
                        IsDataOk = False
                        Exit Function
                    End If
                End If
        
            End If
        End If
        
        IsDataOk = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "IsDataOk"
    
End Function


Private Sub rsLog_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  '#3800 �P�_�w���������s�O�_�i�ϥ� By Kin 2008/03/12
  On Error Resume Next
    Dim strQrySO314 As String
    If sstHead.Tab <> 1 Then Exit Sub
    
    cmdCancelCMD.Enabled = True
    gilReturnCode.Enabled = True
    lblReturnCode.Enabled = True
    
    strQrySO314 = "Select Count(*) From " & GetOwner & "SO314 " & _
                "Where CmdSeqNo=" & rsLog("CmdSeqNo") & _
                " And (CmdStatus in ('C','E')" & _
                " or StopFlag=1 " & _
                " or SNo is not null" & _
                " or ProcessingDate is  Null)"
    If (gcnGi.Execute(strQrySO314)(0) > 0) Or (Not blnUseCancelCmd) Then
        cmdCancelCMD.Enabled = False
        gilReturnCode.Enabled = False
        lblReturnCode.Enabled = False
    End If
    Exit Sub

End Sub

Private Sub sstData_Click(PreviousTab As Integer)
  On Error Resume Next
    Select Case sstData.Tab
        Case 0
            If optBasic(6) Then Call optBasic_Click(6)
        Case 1
            If optCMChange(6) Then Call optCMChange_Click(6)
            If optCMChange(7) Then Call optCMChange_Click(7)
            If optCMChange(8) Then Call optCMChange_Click(8)
        Case Else
        
    End Select
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
        lblResv.Visible = True
        gdtResvTime.Visible = True
    ElseIf sstHead.Tab = 1 Then
        If blnUseCancelCmd Then
            If ggrQuery.Recordset.RecordCount = 0 Then
                lblReturnCode.Enabled = False
                gilReturnCode.Enabled = False
                cmdCancelCMD.Enabled = False
            End If
        Else
            lblReturnCode.Enabled = False
            gilReturnCode.Enabled = False
            cmdCancelCMD.Enabled = False
        End If
        lblResv.Visible = False
        gdtResvTime.Visible = False
    Else
        lblResv.Visible = False
        gdtResvTime.Visible = False
    End If
    
End Sub

