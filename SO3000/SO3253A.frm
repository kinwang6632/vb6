VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Begin VB.Form frmSO3253A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "���O�洫��վ� [SO3253A]"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   2925
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3253A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.CheckBox chkUpdNewClct 
      Caption         =   "��s���O��"
      Height          =   195
      Left            =   6330
      TabIndex        =   10
      Top             =   150
      Value           =   1  '�֨�
      Width           =   1335
   End
   Begin VB.CheckBox chkUpdOldClct 
      Caption         =   "��s�즬�O��"
      Height          =   195
      Left            =   4650
      TabIndex        =   9
      Top             =   150
      Width           =   1575
   End
   Begin VB.CheckBox chkPrtFlag 
      Caption         =   "�C�L�վ����"
      Height          =   195
      Left            =   180
      TabIndex        =   37
      Top             =   7800
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   10230
      TabIndex        =   29
      Top             =   30
      Width           =   1065
   End
   Begin TabDlg.SSTab tabData 
      Height          =   6735
      Left            =   120
      TabIndex        =   30
      ToolTipText     =   "�u�վ㥼������"
      Top             =   960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&A.�̫Ȥ�s���վ�"
      TabPicture(0)   =   "SO3253A.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lbl1"
      Tab(0).Control(1)=   "lblCustCnt"
      Tab(0).Control(2)=   "lblClctYMA"
      Tab(0).Control(3)=   "lblClctYMB"
      Tab(0).Control(4)=   "lblNewClctEn"
      Tab(0).Control(5)=   "lblCustId"
      Tab(0).Control(6)=   "cmdOk1"
      Tab(0).Control(7)=   "gilTab1NewClctEn"
      Tab(0).Control(8)=   "gmTab1YM1"
      Tab(0).Control(9)=   "gmTab1YM2"
      Tab(0).Control(10)=   "ggrData"
      Tab(0).Control(11)=   "cmdQuery1"
      Tab(0).Control(12)=   "txtCustId"
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "&B.�̵�D�վ�"
      TabPicture(1)   =   "SO3253A.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblClctYMC"
      Tab(1).Control(1)=   "lblClctYMD"
      Tab(1).Control(2)=   "lblPrtSNo1"
      Tab(1).Control(3)=   "lblPrtSNo2"
      Tab(1).Control(4)=   "gmTab2Ym1"
      Tab(1).Control(5)=   "gmTab2Ym2"
      Tab(1).Control(6)=   "ggrData2"
      Tab(1).Control(7)=   "cmdOk2"
      Tab(1).Control(8)=   "cmdQuery"
      Tab(1).Control(9)=   "txtPrtSNo2"
      Tab(1).Control(10)=   "txtPrtSNo1"
      Tab(1).Control(11)=   "fraTab2Data"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "&C.�̤j�ӽs���վ�"
      TabPicture(2)   =   "SO3253A.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label7"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "gmTab3Ym1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "gmTab3Ym2"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "ggrData3"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "fraTab3Data"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtPrtSNo1_2"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtPrtSNo2_2"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdQuery3"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cmdOk3"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "&D.�̳�ڽs���վ�"
      TabPicture(3)   =   "SO3253A.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label11"
      Tab(3).Control(1)=   "Label12"
      Tab(3).Control(2)=   "gilNewClctEn_4"
      Tab(3).Control(3)=   "ggrData4"
      Tab(3).Control(4)=   "cmdQuery4"
      Tab(3).Control(5)=   "txtBillNo"
      Tab(3).Control(6)=   "cmdOk4"
      Tab(3).ControlCount=   7
      Begin VB.CommandButton cmdOk4 
         Caption         =   "�}�l"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -65490
         TabIndex        =   67
         Top             =   1230
         Width           =   1335
      End
      Begin VB.TextBox txtBillNo 
         Height          =   345
         Left            =   -73500
         TabIndex        =   66
         Top             =   1350
         Width           =   2385
      End
      Begin VB.CommandButton cmdQuery4 
         Caption         =   "���"
         Height          =   375
         Left            =   -65490
         TabIndex        =   65
         Top             =   690
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdOk3 
         Caption         =   "�}�l"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9930
         TabIndex        =   28
         Top             =   1170
         Width           =   1335
      End
      Begin VB.CommandButton cmdQuery3 
         Caption         =   "���"
         Height          =   375
         Left            =   9930
         TabIndex        =   27
         Top             =   630
         Width           =   1335
      End
      Begin VB.TextBox txtPrtSNo2_2 
         Height          =   345
         Left            =   6510
         MaxLength       =   12
         TabIndex        =   23
         Top             =   570
         Width           =   1305
      End
      Begin VB.TextBox txtPrtSNo1_2 
         Height          =   345
         Left            =   4890
         MaxLength       =   12
         TabIndex        =   22
         Top             =   570
         Width           =   1305
      End
      Begin VB.Frame fraTab3Data 
         Height          =   945
         Left            =   270
         TabIndex        =   50
         Top             =   990
         Visible         =   0   'False
         Width           =   9135
         Begin prjGiList.GiList gilOldClctEn_3 
            Height          =   315
            Left            =   300
            TabIndex        =   24
            Top             =   450
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   556
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FldWidth1       =   700
            FldWidth2       =   1400
         End
         Begin prjGiList.GiList gilNewClctEn_3 
            Height          =   315
            Left            =   6330
            TabIndex        =   26
            Top             =   450
            Width           =   2430
            _ExtentX        =   4286
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
            FldWidth1       =   700
            FldWidth2       =   1400
         End
         Begin prjGiList.GiList gilMduId 
            Height          =   315
            Left            =   2880
            TabIndex        =   25
            Top             =   450
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   556
            Enabled         =   0   'False
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
            FldWidth2       =   2100
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "�s���O��"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   6360
            TabIndex        =   53
            Top             =   210
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "�j�ӽs��"
            Height          =   195
            Left            =   2910
            TabIndex        =   52
            Top             =   210
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "�즬�O��"
            Height          =   195
            Left            =   330
            TabIndex        =   51
            Top             =   210
            Width           =   780
         End
      End
      Begin VB.TextBox txtCustId 
         Height          =   345
         Left            =   -73740
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdQuery1 
         Caption         =   "���"
         Height          =   375
         Left            =   -64710
         TabIndex        =   48
         Top             =   540
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Frame fraTab2Data 
         Height          =   945
         Left            =   -74730
         TabIndex        =   43
         Top             =   990
         Visible         =   0   'False
         Width           =   9135
         Begin prjGiList.GiList gilOldClctEn 
            Height          =   315
            Left            =   300
            TabIndex        =   15
            Top             =   450
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   556
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FldWidth1       =   700
            FldWidth2       =   1400
         End
         Begin prjGiList.GiList gilNewClctEn 
            Height          =   315
            Left            =   6330
            TabIndex        =   17
            Top             =   450
            Width           =   2430
            _ExtentX        =   4286
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
            FldWidth1       =   700
            FldWidth2       =   1400
         End
         Begin prjGiList.GiList gilStrtCode 
            Height          =   315
            Left            =   2880
            TabIndex        =   16
            Top             =   450
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   556
            Enabled         =   0   'False
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
            FldWidth2       =   2100
         End
         Begin VB.Label lblOldClctEn 
            AutoSize        =   -1  'True
            Caption         =   "�즬�O��"
            Height          =   195
            Left            =   330
            TabIndex        =   46
            Top             =   210
            Width           =   780
         End
         Begin VB.Label lblStrtCode 
            AutoSize        =   -1  'True
            Caption         =   "��D"
            Height          =   195
            Left            =   2910
            TabIndex        =   45
            Top             =   210
            Width           =   390
         End
         Begin VB.Label lblNewClctEnB 
            AutoSize        =   -1  'True
            Caption         =   "�s���O��"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   6360
            TabIndex        =   44
            Top             =   210
            Width           =   780
         End
      End
      Begin VB.TextBox txtPrtSNo1 
         Height          =   345
         Left            =   -70110
         MaxLength       =   12
         TabIndex        =   13
         Top             =   570
         Width           =   1305
      End
      Begin VB.TextBox txtPrtSNo2 
         Height          =   345
         Left            =   -68490
         MaxLength       =   12
         TabIndex        =   14
         Top             =   570
         Width           =   1305
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "���"
         Height          =   375
         Left            =   -65070
         TabIndex        =   18
         Top             =   630
         Width           =   1335
      End
      Begin prjGiGridR.GiGridR ggrData 
         Height          =   4875
         Left            =   -74730
         TabIndex        =   8
         Top             =   1680
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8599
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_YM.GiYM gmTab1YM2 
         Height          =   345
         Left            =   -68670
         TabIndex        =   6
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
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
      End
      Begin Gi_YM.GiYM gmTab1YM1 
         Height          =   345
         Left            =   -70050
         TabIndex        =   5
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
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
      End
      Begin prjGiList.GiList gilTab1NewClctEn 
         Height          =   315
         Left            =   -73740
         TabIndex        =   4
         Top             =   600
         Width           =   2625
         _ExtentX        =   4630
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
         FldWidth1       =   700
         FldWidth2       =   1600
      End
      Begin VB.CommandButton cmdOk2 
         Caption         =   "�}�l"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -65070
         TabIndex        =   19
         Top             =   1170
         Width           =   1335
      End
      Begin VB.CommandButton cmdOk1 
         Caption         =   "�}�l"
         Enabled         =   0   'False
         Height          =   345
         Left            =   -64710
         TabIndex        =   31
         Top             =   1020
         Width           =   1125
      End
      Begin prjGiGridR.GiGridR ggrData2 
         Height          =   4485
         Left            =   -74700
         TabIndex        =   38
         Top             =   2040
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   7911
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_YM.GiYM gmTab2Ym2 
         Height          =   345
         Left            =   -72630
         TabIndex        =   12
         Top             =   570
         Width           =   795
         _ExtentX        =   1402
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
      End
      Begin Gi_YM.GiYM gmTab2Ym1 
         Height          =   345
         Left            =   -73770
         TabIndex        =   11
         Top             =   570
         Width           =   795
         _ExtentX        =   1402
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
      End
      Begin prjGiGridR.GiGridR ggrData3 
         Height          =   4485
         Left            =   300
         TabIndex        =   54
         Top             =   2040
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   7911
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_YM.GiYM gmTab3Ym2 
         Height          =   345
         Left            =   2370
         TabIndex        =   21
         Top             =   570
         Width           =   795
         _ExtentX        =   1402
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
      End
      Begin Gi_YM.GiYM gmTab3Ym1 
         Height          =   345
         Left            =   1230
         TabIndex        =   20
         Top             =   570
         Width           =   795
         _ExtentX        =   1402
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
      End
      Begin prjGiGridR.GiGridR ggrData4 
         Height          =   4485
         Left            =   -74700
         TabIndex        =   62
         Top             =   2070
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   7911
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjGiList.GiList gilNewClctEn_4 
         Height          =   315
         Left            =   -73500
         TabIndex        =   63
         Top             =   840
         Width           =   2430
         _ExtentX        =   4286
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
         FldWidth1       =   700
         FldWidth2       =   1400
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "��ڵn��"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -74430
         TabIndex        =   68
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "�s���O��"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -74430
         TabIndex        =   64
         Top             =   900
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "�����~��"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   58
         Top             =   660
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   2100
         TabIndex        =   57
         Top             =   660
         Width           =   195
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "�L��Ǹ�"
         Height          =   195
         Left            =   4050
         TabIndex        =   56
         Top             =   660
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   6240
         TabIndex        =   55
         Top             =   660
         Width           =   195
      End
      Begin VB.Label lblCustId 
         AutoSize        =   -1  'True
         Caption         =   "�Ȥ�s��"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -74700
         TabIndex        =   49
         Top             =   1170
         Width           =   780
      End
      Begin VB.Label lblPrtSNo2 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   -68760
         TabIndex        =   42
         Top             =   660
         Width           =   195
      End
      Begin VB.Label lblPrtSNo1 
         AutoSize        =   -1  'True
         Caption         =   "�L��Ǹ�"
         Height          =   195
         Left            =   -70950
         TabIndex        =   41
         Top             =   660
         Width           =   780
      End
      Begin VB.Label lblClctYMD 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   -72900
         TabIndex        =   40
         Top             =   660
         Width           =   195
      End
      Begin VB.Label lblClctYMC 
         AutoSize        =   -1  'True
         Caption         =   "�����~��"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -74640
         TabIndex        =   39
         Top             =   660
         Width           =   780
      End
      Begin VB.Label lblNewClctEn 
         AutoSize        =   -1  'True
         Caption         =   "�s���O��"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -74700
         TabIndex        =   36
         Top             =   660
         Width           =   780
      End
      Begin VB.Label lblClctYMB 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   -68970
         TabIndex        =   35
         Top             =   690
         Width           =   195
      End
      Begin VB.Label lblClctYMA 
         AutoSize        =   -1  'True
         Caption         =   "�����~��"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -70830
         TabIndex        =   34
         Top             =   690
         Width           =   780
      End
      Begin VB.Label lblCustCnt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Height          =   195
         Left            =   -65550
         TabIndex        =   33
         Top             =   1290
         Width           =   45
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  '�z��
         Caption         =   "��ƦX�p"
         Height          =   195
         Left            =   -66780
         TabIndex        =   32
         Top             =   1290
         Width           =   855
      End
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   90
      Width           =   2730
      _ExtentX        =   4815
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
      FldWidth1       =   600
      FldWidth2       =   1800
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   1170
      TabIndex        =   1
      Top             =   510
      Width           =   2730
      _ExtentX        =   4815
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
      FldWidth1       =   600
      FldWidth2       =   1800
      F5Corresponding =   -1  'True
   End
   Begin Gi_Time.GiTime gitCreateTime1 
      Height          =   345
      Left            =   4920
      TabIndex        =   2
      Top             =   510
      Width           =   1935
      _ExtentX        =   3413
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
   End
   Begin Gi_Time.GiTime gitCreateTime2 
      Height          =   345
      Left            =   7260
      TabIndex        =   3
      Top             =   510
      Width           =   1935
      _ExtentX        =   3413
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
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6990
      TabIndex        =   61
      Top             =   570
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "������"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4110
      TabIndex        =   60
      Top             =   570
      Width           =   780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�A�����O"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   210
      TabIndex        =   59
      Top             =   570
      Width           =   780
   End
   Begin VB.Label lblComp 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��  �q  �O"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   210
      TabIndex        =   47
      Top             =   180
      Width           =   765
   End
End
Attribute VB_Name = "frmSO3253A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsSo3253 As New ADODB.Recordset
'�O���Ȥ�W��
Private strCustName As String
Private ObjOpenDb As New giOpenDBObj.OpenDBObj
Private rstmp1 As New ADODB.Recordset
Private cnMDB As New ADODB.Connection
Private rs3253_mdb As New ADODB.Recordset
Private rsMdb As New ADODB.Recordset
Private rsOrcl As New ADODB.Recordset
Private rsQuery4 As New ADODB.Recordset
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk1_Click()
'�E  �YTab1:
'Loop grid���C�@�����, ����������ɤ�(SO033), �N�ŦX���󪺸��, �����O�H���վ�, SQL:
'"update SO033 set ClctEn= <�s���O�H���N��>, ClctName= <�s���O�H���m�W> where BillNo= <�ӵ���ڽs��> and RealDate is null and ClctYM>= <�_�l�~��> and ClctYM<= <�I��~��>"
On Error GoTo chkErr
    Dim strSQL As String
    Dim strClctSQL As String
    Dim intTotalCount As Long
        If Not IsDataOk Then Exit Sub
        If Not rs3253_mdb.EOF Then
            If gilTab1NewClctEn.GetCodeNo = "" Then MsgBox "�s���H�����o���ťաI", vbExclamation, "�T���I": Exit Sub
            Screen.MousePointer = vbHourglass
            If chkUpdOldClct = 1 Then strClctSQL = ",oldClctEn='" & gilTab1NewClctEn.GetCodeNo & "',OldClctName='" & gilTab1NewClctEn.GetDescription & "'"
            If chkUpdNewClct = 1 Then strClctSQL = strClctSQL & ",ClctEn='" & gilTab1NewClctEn.GetCodeNo & "',ClctName='" & gilTab1NewClctEn.GetDescription & "'"
            strClctSQL = Mid(strClctSQL, 2)
            gcnGi.BeginTrans
            rs3253_mdb.MoveFirst
            While Not rs3253_mdb.EOF
                strSQL = "Update " & GetOwner & "So033 Set " & strClctSQL & " Where RowId='" & rs3253_mdb("RowId").Value & "'"
                If ExecuteSQL(strSQL, gcnGi, , , False) <> giOK Then MsgBox "�s�u���ѡI�Э��s�ʧ@�I", vbExclamation, "�T���I": gcnGi.RollbackTrans: Exit Sub
                intTotalCount = intTotalCount + 1
                rs3253_mdb.MoveNext
                DoEvents
            Wend
            Screen.MousePointer = vbDefault
            cnMDB.Execute "Delete From So3253"
            If OpenRecordset(rs3253_mdb, "Select * from so3253", cnMDB, adOpenForwardOnly, adLockReadOnly, adUseClient, False, False) <> giOK Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I": Exit Sub
            gcnGi.CommitTrans
            MsgBox "�@���� " & intTotalCount & " ��", vbInformation, "�T���I"
            Set ggrData.Recordset = rs3253_mdb
            
            ggrData.Refresh
        End If
    Exit Sub
chkErr:
    gcnGi.RollbackTrans
    Call ErrSub(Me.Name, "cmdOk1_Click")
End Sub

Private Function GetChoose() As String
    On Error GoTo chkErr
    Dim strChoose As String
        strChoose = ""
        If txtPrtSNo1.Text <> "" Then strChoose = subAnd2(strChoose, "  PrtSNo >= '" & txtPrtSNo1.Text & "' ")
        If txtPrtSNo2.Text <> "" Then strChoose = subAnd2(strChoose, "  PrtSNo <= '" & txtPrtSNo2.Text & "' ")
        If gmTab2Ym1.Text <> "" Then strChoose = subAnd2(strChoose, "  ClctYM >= " & gmTab2Ym1.GetValue)
        If gmTab2Ym2.Text <> "" Then strChoose = subAnd2(strChoose, "  ClctYM <= " & gmTab2Ym2.GetValue)
        '���D��2905 �h��A�����O�P���ͤ���d�� 2006/12/28
        If gilServiceType.GetCodeNo <> "" Then strChoose = subAnd2(strChoose, " ServiceType='" & gilServiceType.GetCodeNo & "' ")
        If gitCreateTime1.GetValue <> "" Or gitCreateTime2.GetValue <> "" Then
            If gitCreateTime1.GetValue <> "" And gitCreateTime2.GetValue <> "" Then
                strChoose = subAnd2(strChoose, " CreateTime >= To_Date(" & gitCreateTime1.GetValue & ",'yyyymmddhh24mi')")
                strChoose = subAnd2(strChoose, " CreateTime <= To_Date(" & gitCreateTime2.GetValue & ",'yyyymmddhh24mi')")
            Else
                If gitCreateTime1.GetValue <> "" And gitCreateTime2.GetValue = "" Then
                    strChoose = subAnd2(strChoose, " CreateTime = To_Date(" & gitCreateTime1.GetValue & ",'yyyymmddhh24mi')")
                Else
                    strChoose = subAnd2(strChoose, " CreateTime = To_Date(" & gitCreateTime2.GetValue & ",'yyyymmddhh24mi')")
                End If
            End If
        End If
        Call subAnd2(strChoose, "UCCode Is Not null")
        GetChoose = strChoose
    Exit Function
chkErr:
    ErrSub Me.Name, "GetChoose"
End Function

Private Function GetChoose3() As String
    On Error GoTo chkErr
    Dim strChoose As String
        strChoose = ""
        If txtPrtSNo1_2.Text <> "" Then strChoose = subAnd2(strChoose, "  PrtSNo >= '" & txtPrtSNo1_2.Text & "' ")
        If txtPrtSNo2_2.Text <> "" Then strChoose = subAnd2(strChoose, "  PrtSNo <= '" & txtPrtSNo2_2.Text & "' ")
        If gmTab3Ym1.Text <> "" Then strChoose = subAnd2(strChoose, "  ClctYM >= " & gmTab3Ym1.GetValue)
        If gmTab3Ym2.Text <> "" Then strChoose = subAnd2(strChoose, "  ClctYM <= " & gmTab3Ym2.GetValue)
        Call subAnd2(strChoose, "UCCode Is Not null")
        '���D��2905 �h��A�����O�P���ͤ���d�� 2006/12/28
        If gilServiceType.GetCodeNo <> "" Then strChoose = subAnd2(strChoose, " ServiceType='" & gilServiceType.GetCodeNo & "' ")
        If gitCreateTime1.GetValue <> "" Or gitCreateTime2.GetValue <> "" Then
            If gitCreateTime1.GetValue <> "" And gitCreateTime2.GetValue <> "" Then
                strChoose = subAnd2(strChoose, " CreateTime >= To_Date(" & gitCreateTime1.GetValue & ",'yyyymmddhh24mi')")
                strChoose = subAnd2(strChoose, " CreateTime <= To_Date(" & gitCreateTime2.GetValue & ",'yyyymmddhh24mi')")
            Else
                If gitCreateTime1.GetValue <> "" And gitCreateTime2.GetValue = "" Then
                    strChoose = subAnd2(strChoose, " CreateTime = To_Date(" & gitCreateTime1.GetValue & ",'yyyymmddhh24mi')")
                Else
                    strChoose = subAnd2(strChoose, " CreateTime = To_Date(" & gitCreateTime2.GetValue & ",'yyyymmddhh24mi')")
                End If
            End If
        End If
        GetChoose3 = strChoose
    Exit Function
chkErr:
    ErrSub Me.Name, "GetChoose3"
End Function
'���D��2905 ��ڽs�����Ҹ̪�Where����
Private Function GetChoose4() As String
    On Error GoTo chkErr
    Dim strChoose4 As String
    strChoose4 = ""
    If gilServiceType.GetCodeNo <> "" Then strChoose = subAnd2(strChoose4, " A.ServiceType='" & gilServiceType.GetCodeNo & "' ")
    If gitCreateTime1.GetValue <> "" Or gitCreateTime2.GetValue <> "" Then
        If gitCreateTime1.GetValue <> "" And gitCreateTime2.GetValue <> "" Then
            strChoose = subAnd2(strChoose4, " A.CreateTime >= To_Date(" & gitCreateTime1.GetValue & ",'yyyymmddhh24mi')")
            strChoose = subAnd2(strChoose4, " A.CreateTime <= To_Date(" & gitCreateTime2.GetValue & ",'yyyymmddhh24mi')")
        Else
            If gitCreateTime1.GetValue <> "" And gitCreateTime2.GetValue = "" Then
                strChoose = subAnd2(strChoose4, " A.CreateTime = To_Date(" & gitCreateTime1.GetValue & ",'yyyymmddhh24mi')")
            Else
                strChoose = subAnd2(strChoose4, " A.CreateTime = To_Date(" & gitCreateTime2.GetValue & ",'yyyymmddhh24mi')")
            End If
        End If
    End If
    Call subAnd2(strChoose4, " A.CompCode=" & gilCompCode.GetCodeNo)
    Call subAnd2(strChoose4, "A.UCCode Is Not null")
    Call subAnd2(strChoose4, " A.CustID=B.CustID ")
    GetChoose4 = strChoose4
    Exit Function
chkErr:
    ErrSub Me.Name, "GetChoose4"
End Function

Private Sub cmdOk2_Click()
    On Error GoTo chkErr
    '�E  �ˬdgrid���O�_�����, �Y�L, �h�T��: "�L���", �õL�H�U�ʧ@
    '�E  �ˬd�C�@�����s���O�H���O�_����, �Y�L�h�T��: "����쥲�ݦ���", ��Ц^������, �õL�H�U�ʧ@
    '�E  Loop grid���C�@�����, ������B�z, SQL�p�U:
    ' "update SO032 set ClctEn=<�s���O�H���N��> where StrtCode=<��D�N�X> and ClctEn=<�즬�O�H���N��>"
    '�E  ��:<�즬�O�H��>���i�ରnull��
    Dim strSQL As String, strSQL2 As String
    Dim strClctSQL As String
    Dim intCount As Long
    Dim intTotalCount As Long
    Dim rsTmp As New ADODB.Recordset
        'ggddata.SaveNow
        If MsgBox("�T�w�}�l����?", vbQuestion + vbYesNo + vbDefaultButton2, "����") = vbNo Then Exit Sub
        Screen.MousePointer = vbHourglass
        If Not rsMdb.EOF Then
            rsMdb.MoveFirst
            gcnGi.BeginTrans
            Set rsTmp = rsMdb.Clone
            strSQL2 = GetChoose
            Do While Not rsTmp.EOF
                With rsTmp
                    If IsNull(.Fields("NewElctEn").Value) Then GoTo lblNext
                    strClctSQL = ""
                    If chkUpdOldClct.Value = 1 Then strClctSQL = ",oldClctEn='" & .Fields("NewElctEn").Value & "',OldClctName='" & .Fields("NewEmpName").Value & "'"
                    If chkUpdNewClct = 1 Then strClctSQL = strClctSQL & ",ClctEn='" & .Fields("NewElctEn").Value & "',ClctName='" & .Fields("NewEmpName").Value & "'"
                    strClctSQL = Mid(strClctSQL, 2)
                    
                    strSQL = "Update " & GetOwner & "So033 Set " & strClctSQL
                    strSQL = strSQL & " Where StrtCode='" & .Fields("StrtCode").Value & "' And ClctEn" & IIf((.Fields("ClctEn").Value & "") = "", " Is Null", "='" & .Fields("ClctEn").Value & "'")
                    strSQL = strSQL & " And " & strSQL2
                    gcnGi.Execute strSQL, intCount
                    intTotalCount = intTotalCount + intCount
                End With
lblNext:
                rsTmp.MoveNext
            Loop
            Screen.MousePointer = vbDefault
            MsgBox "�@���� " & intTotalCount & " ��", vbInformation, "�T���I"
            gcnGi.CommitTrans
        Else
            MsgBox "�L��ơI�Ч�s��ܫ�A���ʡI", vbExclamation, "�T���I"
        End If
        rsMdb.Filter = " intCount = -99"
        Set ggrData2.Recordset = rsMdb
        ggrData2.Refresh
        cmdQuery.Enabled = True
        cmdOk2.Enabled = False
        Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
    gcnGi.RollbackTrans
    ErrSub Me.Name, "cmdOK2"
End Sub

Private Sub cmdOk3_Click()
    On Error GoTo chkErr
    '�E  �ˬdgrid���O�_�����, �Y�L, �h�T��: "�L���", �õL�H�U�ʧ@
    '�E  �ˬd�C�@�����s���O�H���O�_����, �Y�L�h�T��: "����쥲�ݦ���", ��Ц^������, �õL�H�U�ʧ@
    '�E  Loop grid���C�@�����, ������B�z, SQL�p�U:
    ' "update SO032 set ClctEn=<�s���O�H���N��> where MduId=<�j�ӽs��> and ClctEn=<�즬�O�H���N��>"
    '�E  ��:<�즬�O�H��>���i�ରnull��
    Dim strSQL As String, strSQL2 As String
    Dim strClctSQL As String
    Dim intCount As Long
    Dim intTotalCount As Long
    Dim rsTmp As New ADODB.Recordset
        'ggddata.SaveNow
        If MsgBox("�T�w�}�l����?", vbQuestion + vbYesNo + vbDefaultButton2, "����") = vbNo Then Exit Sub
        Screen.MousePointer = vbHourglass
        If Not rsMdb.EOF Then
            rsMdb.MoveFirst
            gcnGi.BeginTrans
            Set rsTmp = rsMdb.Clone
            strSQL2 = GetChoose3
            Do While Not rsTmp.EOF
                With rsTmp
                    If IsNull(.Fields("NewElctEn").Value) Then GoTo lblNext
                    strClctSQL = ""
                    If chkUpdOldClct.Value = 1 Then strClctSQL = ",oldClctEn='" & .Fields("NewElctEn").Value & "',OldClctName='" & .Fields("NewEmpName").Value & "'"
                    If chkUpdNewClct = 1 Then strClctSQL = strClctSQL & ",ClctEn='" & .Fields("NewElctEn").Value & "',ClctName='" & .Fields("NewEmpName").Value & "'"
                    strClctSQL = Mid(strClctSQL, 2)
                    strSQL = "Update " & GetOwner & "So033 Set " & strClctSQL
                    
                    strSQL = strSQL & " Where MduId='" & .Fields("MduId").Value & "' And ClctEn" & IIf((.Fields("ClctEn").Value & "") = "", " Is Null", "='" & .Fields("ClctEn").Value & "'")
                    strSQL = strSQL & " And " & strSQL2
                    gcnGi.Execute strSQL, intCount
                    intTotalCount = intTotalCount + intCount
                End With
lblNext:
                rsTmp.MoveNext
            Loop
            Screen.MousePointer = vbDefault
            MsgBox "�@���� " & intTotalCount & " ��", vbInformation, "�T���I"
            gcnGi.CommitTrans
        Else
            MsgBox "�L��ơI�Ч�s��ܫ�A���ʡI", vbExclamation, "�T���I"
        End If
        rsMdb.Filter = " intCount = -99"
        Set ggrData3.Recordset = rsMdb
        ggrData3.Refresh
        cmdQuery3.Enabled = True
        cmdOk3.Enabled = False
        Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
    Screen.MousePointer = vbDefault
    MsgBox "������ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
    gcnGi.RollbackTrans
End Sub
'**************************************************************************************
'���D��2905 �}�l���洫��ʧ@�A�ϥμ�����ƪ�Where����hUpdate���
'�ϥΰj���ˬdtmpSO3253A4 KeyBillNo���Ȫ��~��Update���ʧ@�AKeyBillNo�N���J���渹
'�渹�i��|���\�h�i�A��Update�u�ݤ@���A�קKUpdate��X�Ӫ��`���Ƥ����T By Kin 2007/1/12
'***************************************************************************************
Private Sub cmdOk4_Click()
On Error GoTo chkErr

    Dim strSQL As String, strSQL2 As String
    Dim strClctSQL As String
    Dim intCount As Long
    Dim intTotalCount As Long
    Dim rsTmp As New ADODB.Recordset
        'ggddata.SaveNow
        If MsgBox("�T�w�}�l����?", vbQuestion + vbYesNo + vbDefaultButton2, "����") = vbNo Then Exit Sub
        Screen.MousePointer = vbHourglass
        If Not rsMdb.EOF Then
            rsMdb.MoveFirst
            gcnGi.BeginTrans
            Set rsTmp = rsMdb.Clone
            Do While Not rsTmp.EOF
               If rsTmp("KeyBillNo") & "" <> "" Then
                  With rsTmp
                     strSQL2 = UpdateWhere4(rsTmp)
                     strClctSQL = ""
                     If chkUpdOldClct.Value = 1 Then strClctSQL = ",A.oldClctEn='" & gilNewClctEn_4.GetCodeNo & "',A.OldClctName='" & gilNewClctEn_4.GetDescription & "'"
                     If chkUpdNewClct = 1 Then strClctSQL = strClctSQL & ",A.ClctEn='" & gilNewClctEn_4.GetCodeNo & "',A.ClctName='" & gilNewClctEn_4.GetDescription & "'"
                     strClctSQL = Mid(strClctSQL, 2)
                     strSQL = "Update " & GetOwner & "So033 A Set " & strClctSQL
                     strSQL = strSQL & " Where " & strSQL2
                     gcnGi.Execute strSQL, intCount
                     intTotalCount = intTotalCount + intCount
                  End With
               End If
lblNext:
                rsTmp.MoveNext
            Loop
            Screen.MousePointer = vbDefault
            MsgBox "�@���� " & intTotalCount & " ��", vbInformation, "�T���I"
            gcnGi.CommitTrans
        Else
            MsgBox "�L��ơI�Ч�s��ܫ�A���ʡI", vbExclamation, "�T���I"
        End If
        cnMDB.Execute "Delete * from tmpSO3253A4"
        Set ggrData4.Recordset = rsMdb
        ggrData4.Blank = True
        ggrData4.Refresh
        'cmdQuery4.Enabled = True
        cmdOk4.Enabled = False
        txtBillNo = ""
        gilNewClctEn_4.SetCodeNo ""
        gilNewClctEn_4.SetDescription ""
        gilNewClctEn_4.Refresh
        Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
    Screen.MousePointer = vbDefault
    MsgBox "������ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
    gcnGi.RollbackTrans
End Sub

Private Sub cmdQuery_Click()
On Error GoTo chkErr
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
        'ggddata.SaveNow
        '���M���ثe�s�u
        If Not IsDataOk Then Exit Sub
        strSQL = "Select StrtCode,StrtName,ClctEn,D.EmpName,intCount,'' as NewElctEn,'' as NewEmpName From " & _
                " (Select A.StrtCode,B.Description StrtName,A.ClctEn,A.intCount From " & _
                        "(Select StrtCode,ClctEn,Count(*) as intCount From " & GetOwner & "So033 " & _
                            " Where " & GetChoose & " Group By StrtCode,ClctEn) A," & _
                            GetOwner & "CD017 B Where A.StrtCode=B.CodeNo) C," & GetOwner & "CM003 D " & _
             " Where C.ClctEn=D.EmpNo(+)  Order by ClctEn,StrtCode"
        Screen.MousePointer = vbHourglass
        If Not GetRS(rsOrcl, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If Not CreateMDBTable(rsOrcl, "tmpSO3253A2", cnMDB) Then Exit Sub
        rsMdb.Filter = ""
        Screen.MousePointer = vbDefault
        If rsOrcl.EOF Then
            MsgNoRcd
            Exit Sub
        Else
           Call SaveRecToTmp
        End If
        
        If Not rsMdb.EOF Then
            rsMdb.MoveFirst
            Set ggrData2.Recordset = rsMdb
            ggrData2.Refresh
            ggrData2.Blank = False
            ggrData2.Enabled = True
            cmdQuery.Enabled = False
            cmdOk2.Enabled = True
            ggrData2.SetFocus
        End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdQuery_Click")
End Sub

Private Function DefaultRs() As Boolean
    On Error Resume Next
        'cnMDB.Execute "Drop Table tmpSo3253A2"
    On Error GoTo chkErr
        'cnMDB.Execute "Select * Into tmpSo3253A2 From tmpSo3130 Where 1 = 0"
        If tabData.Tab = 1 Then
            If Not GetRS(rsMdb, "Select * from tmpSo3253A2", cnMDB, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        ElseIf tabData.Tab = 2 Then
            If Not GetRS(rsMdb, "Select * from tmpSo3253A3", cnMDB, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        ElseIf tabData.Tab = 3 Then
            If Not GetRS(rsMdb, "Select * from tmpSo3253A4", cnMDB, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        End If
'        If rsMdb.State = adStateOpen Then rsMdb.Close
'        With rsMdb
'            .CursorLocation = adUseClient
'            .CursorType = adOpenKeyset
'            .LockType = adLockOptimistic
'            .Fields.Append "ClctEn", adBSTR, 10, adFldIsNullable
'            .Fields.Append "EmpName", adBSTR, 20, adFldIsNullable
'            .Fields.Append "StrtCode", adInteger, , adFldIsNullable
'            .Fields.Append "StrtName", adBSTR, 40, adFldIsNullable
'            .Fields.Append "Count", adInteger, , adFldIsNullable
'            .Fields.Append "NewElctEn", adBSTR, 10, adFldIsNullable
'            .Fields.Append "NewEmpName", adBSTR, 20, adFldIsNullable
'            .Open
'        End With
        DefaultRs = True
    Exit Function
chkErr:
    ErrSub Me.Name, "DefaultRS"
End Function

Private Sub SaveRecToTmp()
On Error GoTo chkErr
Dim strSQL As String
    '�O����Rs
    DefaultRs
    While Not rsOrcl.EOF
       With rsMdb
        .AddNew
        .Fields("ClctEn") = GetFieldValue(rsOrcl, "ClctEn")
        If Len(rsOrcl("EmpName") & "") > 0 Then
            .Fields("EmpName") = GetFieldValue(rsOrcl, "EmpName")
        End If
        If tabData.Tab = 1 Then
            .Fields("StrtCode") = NoZero(rsOrcl("StrtCode"))
            .Fields("StrtName") = NoZero(rsOrcl("StrtName"))
        ElseIf tabData.Tab = 2 Then
            .Fields("MduId") = NoZero(rsOrcl("MduId"))
            .Fields("MduName") = NoZero(rsOrcl("MduName"))
        End If
        .Fields("intCount") = GetFieldValue(rsOrcl, "intCount")
        .Update
       End With
       rsOrcl.MoveNext
    Wend
    
    'blnChkrs = True
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "SaveRecToTmp")
End Sub


Private Sub cmdQuery3_Click()
    On Error GoTo chkErr
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
        'ggddata.SaveNow
        '���M���ثe�s�u
        If Not IsDataOk Then Exit Sub
        strSQL = "Select MduId,MduName,ClctEn,D.EmpName,intCount,'' as NewElctEn,'' as NewEmpName  From " & _
                " (Select A.MduId,B.Name MduName,A.ClctEn,A.intCount From " & _
                        "(Select MduId,ClctEn,Count(*) as intCount From " & GetOwner & "So033 " & _
                            " Where " & GetChoose3 & " Group By MduId,ClctEn) A," & _
                            " " & GetOwner & "SO017 B Where A.MduId=B.MduId) C," & GetOwner & "CM003 D " & _
             " Where C.ClctEn=D.EmpNo(+)  Order by ClctEn,MduId"
        Screen.MousePointer = vbHourglass
        If Not GetRS(rsOrcl, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If Not CreateMDBTable(rsOrcl, "tmpSO3253A3", cnMDB) Then Exit Sub
        rsMdb.Filter = ""
        Screen.MousePointer = vbDefault
        If rsOrcl.EOF Then
            MsgNoRcd
            Exit Sub
        Else
           Call SaveRecToTmp
        End If
        
        If Not rsMdb.EOF Then
            rsMdb.MoveFirst
            Set ggrData3.Recordset = rsMdb
            ggrData3.Refresh
            ggrData3.Blank = False
            ggrData3.Enabled = True
            cmdQuery3.Enabled = False
            cmdOk3.Enabled = True
            ggrData3.SetFocus
        End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdQuery3_Click")
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyEscape Then Unload Me: Exit Sub
End Sub

Private Sub Form_Load()
    On Error GoTo chkErr
        Call InitData
        Call subGil
        Call subGrd
        tabData.Tab = 0
        CreateTable
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub InitData()
    On Error GoTo chkErr
    Dim strPath As String
        strPath = ReadGICMIS1("TmpMDBPath")
        With cnMDB
            If .State = adStateClosed Then
                '.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Password=;Data Source=" & strPath & "\Tmp1111.MDB" & ";Persist Security Info=True"
                .ConnectionString = "Provider=" & GetOleDbProvider & ";Password=;Data Source=" & strPath & "\Tmp1111.MDB" & ";Persist Security Info=True"
                .Open
            End If
            If Not GetRS(rs3253_mdb, "Select * from So3253", cnMDB, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
            cnMDB.Execute "Delete * from so3253"
        End With
    Exit Sub
chkErr:
    ErrSub Me.Name, "InitData"
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    If Not rs3253_mdb.EOF Then
            cnMDB.Execute "Delete From So3253"
    End If
    rs3253_mdb.Close
    Set rs3253_mdb = Nothing
    If cnMDB.State = adStateOpen Then cnMDB.Close
    Set cnMDB = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub ggrData_GotFocus()
    On Error Resume Next
        'txtCustId.SetFocus
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error Resume Next
        If UCase(Fld.Name) = "SHOULDAMT" Then
            Value = Format(Value, "###,###,###")
        '�אּ�褸�~ By Kin 2007/06/12
'        ElseIf UCase(Fld.Name) = "SHOULDDATE" Then
'            Value = Format(Value, "ee/MM/dd")
        End If
End Sub

Private Sub ggrData2_DblClick()
    On Error GoTo chkErr
        If ggrData2.Recordset.BOF Or ggrData2.Recordset.EOF Then Exit Sub
        fraTab2Data.Visible = True
        gilNewClctEn.Visible = True
        gilNewClctEn.SetFocus
        With ggrData2.Recordset
            gilOldClctEn.SetCodeNo .Fields("ClctEn") & ""
            gilOldClctEn.SetDescription .Fields("EmpName") & ""
            gilStrtCode.SetCodeNo .Fields("StrtCode") & ""
            gilStrtCode.SetDescription .Fields("StrtName") & ""
            gilNewClctEn.SetCodeNo .Fields("NewElctEn") & ""
            gilNewClctEn.SetDescription .Fields("NewEmpName") & ""
        End With
    Exit Sub
chkErr:
    ErrSub Me.Name, "ggrData2_DblClick"
End Sub

Private Sub ggrData3_DblClick()
    On Error GoTo chkErr
        If ggrData3.Recordset.BOF Or ggrData3.Recordset.EOF Then Exit Sub
        fraTab3Data.Visible = True
        gilNewClctEn_3.Visible = True
        gilNewClctEn_3.SetFocus
        With ggrData3.Recordset
            gilOldClctEn_3.SetCodeNo .Fields("ClctEn") & ""
            gilOldClctEn_3.SetDescription .Fields("EmpName") & ""
            gilMduId.SetCodeNo .Fields("MduId") & ""
            gilMduId.SetDescription .Fields("MduName") & ""
            gilNewClctEn_3.SetCodeNo .Fields("NewElctEn") & ""
            gilNewClctEn_3.SetDescription .Fields("NewEmpName") & ""
        End With
    Exit Sub
chkErr:
    ErrSub Me.Name, "ggrData3_DblClick"
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3250", "SO3253") Then Exit Sub
        Call subGil
        Call subGim
        Call GiListFilter(gilTab1NewClctEn, "", gilCompCode.GetCodeNo)
        Call GiListFilter(gilOldClctEn, "", gilCompCode.GetCodeNo)
        Call GiListFilter(gilStrtCode, "", gilCompCode.GetCodeNo)
        Call GiListFilter(gilNewClctEn, "", gilCompCode.GetCodeNo)
        '���D�� 2905,������O�s���O�H�����L�o����
        Call GiListFilter(gilNewClctEn_4, "", gilCompCode.GetCodeNo)
        Call InitData
        On Error Resume Next
        ggrData.Blank = True
        ggrData2.Blank = True
        ggrData3.Blank = True
        ggrData4.Blank = True
        lblCustCnt = 0
        Call EnableMode
End Sub

Private Sub gilNewClctEn_3_Validate(Cancel As Boolean)
    On Error GoTo chkErr
        With ggrData3.Recordset
            If .Fields("ClctEn") = gilNewClctEn_3.GetCodeNo Then
                MsgBox "�s���O�H�����o�P�즬�O�H���ۦP", vbExclamation, "ĵ�i"
                gilNewClctEn_3.GotoCodeNo
                Cancel = True
                Exit Sub
            Else
                .Fields("NewElctEn") = NoZero(gilNewClctEn_3.GetCodeNo)
                .Fields("NewEmpName") = NoZero(gilNewClctEn_3.GetDescription)
                .Update
            End If
        End With
        gilNewClctEn_3.Visible = False
        fraTab3Data.Visible = False
        ggrData.Refresh
    Exit Sub
chkErr:
    ErrSub Me.Name, "gilNewClctEn_3_Validate"

End Sub

Private Sub gilNewClctEn_Validate(Cancel As Boolean)
    On Error GoTo chkErr
        With ggrData2.Recordset
            If .Fields("ClctEn") = gilNewClctEn.GetCodeNo Then
                MsgBox "�s���O�H�����o�P�즬�O�H���ۦP", vbExclamation, "ĵ�i"
                gilNewClctEn.GotoCodeNo
                Cancel = True
                Exit Sub
            Else
                .Fields("NewElctEn") = NoZero(gilNewClctEn.GetCodeNo)
                .Fields("NewEmpName") = NoZero(gilNewClctEn.GetDescription)
                .Update
            End If
        End With
        gilNewClctEn.Visible = False
        fraTab2Data.Visible = False
        ggrData.Refresh
    Exit Sub
chkErr:
    ErrSub Me.Name, "gilNewClctEn_Validate"
End Sub

Private Sub gitCreateTime1_GotFocus()
    On Error Resume Next
    If gitCreateTime1.GetCDate = "" Then gitCreateTime1.SetValue Date
End Sub

Private Sub gitCreateTime1_Validate(Cancel As Boolean)
    On Error Resume Next
    If gitCreateTime2.GetValue = "" Then gitCreateTime2.SetValue Format(Date, "YYYY/MM/DD" & " 23:59")
End Sub

Private Sub gitCreateTime2_GotFocus()
'    On Error Resume Next
'    If gitCreateTime2.GetValue = "" Then gitCreateTime2.SetValue Format(Date, "YYYY/MM/DD" & " 23:59")
End Sub

Private Sub gitCreateTime2_Validate(Cancel As Boolean)
   On Error Resume Next
     If gitCreateTime1.GetValue = "" Or gitCreateTime2.GetValue = "" Then Exit Sub
     If Not IsDate(gitCreateTime2.Text) Then Exit Sub
     If DateDiff("d", gitCreateTime1.GetValue(True), gitCreateTime2.GetValue(True)) < 0 Then
        MsgBox "���ͮɶ��I��餣�i�p�󲣥ͮɶ��_�l��!", vbExclamation, "ĵ�i"
        gitCreateTime2.SetValue gitCreateTime1.GetValue
        Cancel = True
     End If
End Sub

Private Sub gmTab1YM1_Validate(Cancel As Boolean)
    If gmTab1YM1.Text <> "" Then
        gmTab1YM2.Text = gmTab1YM1.Text
    End If
End Sub

Private Sub gmTab2Ym2_GotFocus()
    gmTab2Ym2.SetValue gmTab2Ym1.GetValue
End Sub

Private Sub gmTab3Ym2_GotFocus()
    On Error Resume Next
        gmTab3Ym2.SetValue gmTab3Ym1.GetValue

End Sub

Private Sub tabData_Click(PreviousTab As Integer)
    On Error GoTo chkErr
        
        If tabData.Tab = 0 Then
            On Error Resume Next
            gilTab1NewClctEn.SetFocus
        ElseIf tabData.Tab = 1 Then
            Call subGrd2
            gmTab2Ym1.SetFocus
        ElseIf tabData.Tab = 2 Then
            Call subGrd3
            gmTab3Ym1.SetFocus
        ElseIf tabData.Tab = 3 Then
            Call subGrd4
            txtBillNo.SetFocus
        End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "tabData_Click")
End Sub

Private Sub subGil()
    On Error GoTo chkErr
        SetgiList gilTab1NewClctEn, "EmpNO", "EmpName", "CM003", , , , , 10, 20, , True
        SetgiList gilOldClctEn, "EmpNO", "EmpName", "CM003", , , , , 10, 20, , True
        SetgiList gilNewClctEn, "EmpNO", "EmpName", "CM003", , , , , 10, 20, , True
        SetgiList gilStrtCode, "CodeNO", "Description", "CD017", , , , , 6, 40, , True
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
        SetgiList gilOldClctEn_3, "EmpNO", "EmpName", "CM003", , , , , 10, 20, , True
        SetgiList gilNewClctEn_3, "EmpNO", "EmpName", "CM003", , , , , 10, 20, , True
        SetgiList gilMduId, "MduId", "Name", "SO017", , , , , 6, 40, , True
        '���D��2905 �s�W������O��s�����O�H���N�X��
        SetgiList gilNewClctEn_4, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        '���D��2905 �h�W�[�P�_�A�ȧO
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.SetCodeNo ""
        gilServiceType.Query_Description
        '*************************************************************
        '���D��2905���դ�OK�A�A�ȧO�w�]�Ȼݳ]���ťաA�{���쥻�]��ListIndex=1 By Kin 2007/3/3
        'gilServiceType.ListIndex = 1
        '*************************************************************
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGil"
End Sub

Private Sub subGim()
    On Error Resume Next
        'SetgiMulti gimMduId, "MduID", "Name", "SO017", "�j�ӽs��", "�j�ӦW��"

End Sub

Private Sub subGrd()
    On Error GoTo chkErr
    Dim mFlds As New prjGiGridR.GiGridFlds
        '�Ȥ�s���B�m�W�B���O�H���B���O�ϡB��ڽs���B�L��Ǹ��B�����`�B�B�������
        With mFlds
            .Add "CustId", , , , False, "�Ȥ�s��", vbLeftJustify
            .Add "CustName", , , , False, "         �m�W        ", vbLeftJustify
            .Add "ClctName", , , , False, "     ���O�H��    ", vbLeftJustify
            .Add "BillNo", , , , False, "       ��ڽs��    ", vbLeftJustify
            .Add "PrtSNo", , , , False, "       �L��Ǹ�    ", vbLeftJustify
            .Add "ShouldAmt", , , , False, "    �����`�B    ", vbRightJustify
            '�אּ�褸�~ By Kin 2007/06/12
            .Add "ShouldDate", giControlTypeDate, , , False, "   �������    ", vbLeftJustify
        End With
        'rstmp1.CursorLocation = adUseClient
        'rstmp1.Open "Select RowID ,ClCtName,BillNo,PrtSNo,CitemName,ShouldAmt,ShouldDate,Custid from so033 where 0=1", gcnGi, adOpenForwardOnly, adLockReadOnly
        ggrData.AllFields = mFlds
        ggrData.SetHead
        Set ggrData.Recordset = rs3253_mdb
        ggrData.Refresh
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGrd")
End Sub

Private Sub subGrd2()
    On Error GoTo chkErr
    Dim mFlds As New prjGiGridR.GiGridFlds
        '(�즬�O��(�N��/�m�W),��D,�i��,�s���O��(�N��/�m�W))
        With mFlds
            .Add "ClctEn", , , , False, "���O�H���N��", vbLeftJustify
            .Add "EmpName", , , , False, "���u�m�W", vbLeftJustify
            .Add "StrtCode", , , , False, "��D�N�X", vbRightJustify
            .Add "StrtName", , , , False, "     ��D�W��    ", vbLeftJustify
            .Add "intCount", , , , False, "�i��", vbRightJustify
            .Add "NewElctEn", , , , True, "�s���O�H���N��", vbLeftJustify, , , , "EmpNO", "EmpName", "CM003", 10, 20
            .Add "NewEmpName", , , , True, "�s���O�H���m�W", vbLeftJustify, , , , "EmpNO", "EmpName", "CM003", 10, 20
        End With
        ggrData2.AllFields = mFlds
        ggrData2.SetHead
        ggrData2.Enabled = False
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGrd2")
End Sub

Private Sub subGrd3()
    On Error GoTo chkErr
    Dim mFlds As New prjGiGridR.GiGridFlds
        '(�즬�O��(�N��/�m�W),��D,�i��,�s���O��(�N��/�m�W))
        With mFlds
            .Add "ClctEn", , , , False, "���O�H���N��", vbLeftJustify
            .Add "EmpName", , , , False, "���u�m�W", vbLeftJustify
            .Add "MduId", , , , False, "�j�ӽs��", vbRightJustify
            .Add "MduName", , , , False, "�j�ӦW��      ", vbLeftJustify
            .Add "intCount", , , , False, "�i��", vbRightJustify
            .Add "NewElctEn", , , , True, "�s���O�H���N��", vbLeftJustify, , , , "EmpNO", "EmpName", "CM003", 10, 20
            .Add "NewEmpName", , , , True, "�s���O�H���m�W", vbLeftJustify, , , , "EmpNO", "EmpName", "CM003", 10, 20
        End With
        ggrData3.AllFields = mFlds
        ggrData3.SetHead
        ggrData3.Enabled = False
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGrd3")
End Sub
Private Sub subGrd4()
   On Error GoTo chkErr
   Dim mFlds As New prjGiGridR.GiGridFlds
      With mFlds
         .Add "CustId", , , , False, "�Ȥ�s��", vbLeftJustify
         .Add "CustName", , , , False, "       �m�W       ", vbLeftJustify
         .Add "ClctName", , , , False, "  ���O�H��   ", vbLeftJustify
         .Add "BillNo", , , , False, "      ��ڽs��    ", vbLeftJustify
         .Add "PrtSNo", , , , False, "    �L��Ǹ�    ", vbLeftJustify
         .Add "MediaBillNo", , , , False, "    �C��渹    ", vbLeftJustify
         .Add "ShouldAmt", , , , False, "    �����`�B    ", vbRightJustify
         .Add "ShouldDate", , , , False, "   �������    ", vbLeftJustify
      End With
      ggrData4.AllFields = mFlds
      ggrData4.SetHead
      ggrData4.Enabled = True
chkErr:
   Call ErrSub(Me.Name, "subGrd4")
End Sub
Private Function IsDataOk() As Boolean
    On Error GoTo chkErr
    '2.  ���n���:
    '�E  �Y��Tab1: �s���O��, �����~��, grid���ݦ���
    '�E  �Y��Tab2: �L��Ǹ��d��, �����~��d��, �H�έ즬�O��/��D/�s���O���ݬ�����@�լҦ���
    IsDataOk = False
        If Not MustExist(gilCompCode, 2, "���q�O") Then Exit Function
        If chkUpdNewClct = 0 And chkUpdOldClct = 0 Then
            MsgBox "����s�즬�O���Φ��O���ֻ̤ݿ�@��!!", vbExclamation, gimsgPrompt
            chkUpdNewClct.Value = 1
            Exit Function
        End If
        If tabData.Tab = 0 Then
            
            If Not MustExist(gilTab1NewClctEn, 2, "�s���O��") Then Exit Function
            If Not MustExist(gmTab1YM1, 1, "�����~��_�l��") Then Exit Function
            If Not MustExist(gmTab1YM2, 1, "�����~��I���") Then Exit Function
            
            If gmTab1YM1.Text > gmTab1YM2.Text Then MsgBox "�_�l�~�뤣�o�j�󵲧��~��I", vbExclamation, "�T���I": gmTab1YM1.SetFocus: Exit Function
            If rs3253_mdb.State = adStateClosed Then MsgNoRcd: txtCustId.SetFocus: Exit Function
            If rs3253_mdb.RecordCount <= 0 Then MsgNoRcd: txtCustId.SetFocus: Exit Function
        ElseIf tabData.Tab = 1 Then
            'If txtPrtSNo1.Text = "" Then MsgBox gMsgIsDataOK, vbExclamation, "�T���I": txtPrtSNo1.SetFocus: Exit Function
            'If txtPrtSNo2.Text = "" Then MsgBox gMsgIsDataOK, vbExclamation, "�T���I": txtPrtSNo2.SetFocus: Exit Function
            'If txtPrtSNo2.Text < txtPrtSNo1.Text Then MsgBox "�L��Ǹ�����ȿ��~�I�_�l�Ǹ����o�j�󵲧��Ǹ��I", vbExclamation, "�T���I": txtPrtSNo2.SetFocus: Exit Function
            If Not MustExist(gmTab2Ym1, 1, "�����~��_�l��") Then Exit Function
            If Not MustExist(gmTab2Ym2, 1, "�����~��I���") Then Exit Function
            If gmTab2Ym1.Text > gmTab2Ym2.Text Then MsgBox "�_�l�~�뤣�o�j�󵲧��~��I", vbExclamation, "�T���I": gmTab1YM2.SetFocus: Exit Function
        ElseIf tabData.Tab = 3 Then
            If Not MustExist(gilNewClctEn_4, 2, "�s���O��") Then Exit Function
        End If
    IsDataOk = True
Exit Function
chkErr:
    Call ErrSub(Me.Name, "IsDataOK")
End Function

Private Sub cmdQuery1_Click()
    On Error GoTo chkErr
        If Not MustExist(gilCompCode, 2, "���q�O") Then Exit Sub
        If Not MustExist(gilTab1NewClctEn, 2, "�s���O��") Then Exit Sub
        If Not MustExist(gmTab1YM1, 1, "�����~��_�l��") Then Exit Sub
        If Not MustExist(gmTab1YM2, 1, "�����~��I���") Then Exit Sub
        If Not MustExist(txtCustId, , "�Ȥ�s��") Then Exit Sub
        If gmTab1YM1.Text > gmTab1YM2.Text Then MsgBox "�_�l�~�뤣�o�j�󵲧��~��I", vbExclamation, "�T���I": gmTab1YM1.SetFocus: Exit Sub
                
        If Not CustIdMode Then Exit Sub
        cmdOk1.Enabled = True
        ggrData.Blank = False

    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdQuery1_Click"
End Sub

Private Function MduIdMode() As Boolean
    On Error GoTo chkErr
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
        strSQL = "Select CustName From " & GetOwner & "So033 Where Custid=" & txtCustId & " And CompCode = " & gilCompCode.GetCodeNo
        
        strSQL = "Select RowID ,ClCtName,BillNo,PrtSNo,CitemName,ShouldAmt,ShouldDate,Custid" & _
                    " From " & GetOwner & "So033 Where  CustID=" & txtCustId & " And ClctYM>=" & gmTab1YM1.Text & " And ClctYm <=" & gmTab1YM2.Text & _
                    " And RealDate Is Null"
        
    
        strSQL = "Select CustName From " & GetOwner & "So001 Where Custid=" & txtCustId & " And CompCode = " & gilCompCode.GetCodeNo
    Exit Function
chkErr:
    ErrSub Me.Name, "MduIdMode"
End Function


Private Function CustIdMode() As Boolean
    On Error GoTo chkErr
    '�ھڿ�J�Ȥ�s���A�ܫȤ�����(So001)���Ȥ�m�W�I�Y�L��ơI�h�T���I:"�L���Ȥ�s���I"
    
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim lngAffectCounts As Long
    Static tmpCounts As Long
    '���D�� 2905 �h�W�[�P�_�A�ȧO�P���ͤ���϶��P�_ 2006/12/28
        strChoose = ""
        If gitCreateTime1.GetValue <> "" Or gitCreateTime2.GetValue <> "" Then
            If gitCreateTime1.GetValue <> "" And gitCreateTime2.GetValue <> "" Then
                strChoose = subAnd2(strChoose, " A.CreateTime >= To_Date(" & gitCreateTime1.GetValue & ",'yyyymmddhh24mi')")
                strChoose = subAnd2(strChoose, " A.CreateTime <= To_Date(" & gitCreateTime2.GetValue & ",'yyyymmddhh24mi')")
            Else
                If gitCreateTime1.GetValue <> "" And gitCreateTime2.GetValue = "" Then
                    strChoose = subAnd2(strChoose, " A.CreateTime = To_Date(" & gitCreateTime1.GetValue & ",'yyyymmddhh24mi')")
                Else
                    strChoose = subAnd2(strChoose, " A.CreateTime = To_Date(" & gitCreateTime2.GetValue & ",'yyyymmddhh24mi')")
                End If
            End If
        End If
        '**************************************************************************************************************************************
        '���D��2905���դ�OK,�A�ȧO���i��w�]�ťաA�ҥH�ݧP�_�A�ȧO�O�_���� By Kin 2007/3/3
        strSQL = "Select CustName From " & GetOwner & "So001 Where Custid=" & txtCustId & " And CompCode = " & gilCompCode.GetCodeNo & _
                 IIf(gilServiceType.GetCodeNo & "" <> "", " And Instr(ServiceType,'" & gilServiceType.GetCodeNo & "')>0", "")
        '*************************************************************************************************************************************
        '��So001���Ȥ�W�ټȦs��strcustname
        If Not GetRS(rs, strSQL, gcnGi) Then
            txtCustId.Text = ""
            MsgBox "�s�����ѡI�Э��s��J�I", vbExclamation, "�T���I"
            Exit Function
        End If
        
        
        If Not rs.EOF Then
            strCustName = rs("CustName").Value & ""
            strSQL = "Select Custid From So3253 Where Custid=" & Trim(txtCustId.Text)
            If Not GetRS(rs, strSQL, cnMDB) Then MsgBox "�s�����ѡI�Э��s�ާ@�I": txtCustId.SetFocus:    Exit Function
            If Not rs.EOF Then MsgBox "������Ƥw�s�b�I�L�k���п�J", vbExclamation, "�T���I": txtCustId.Text = "": Exit Function
            Call CloseRecordset(rs)
            '���D�� 2905 �h�W�[�P�_�������d�� 2006/12/28
            '**************************************************************************************************************************************
            '���D��2905���դ�OK,�A�ȧO���i��w�]�ťաA�ҥH�ݧP�_�A�ȧO�O�_���� By Kin 2007/3/3
            strSQL = "Select A.RowID ,A.ClCtName,A.BillNo,A.PrtSNo,A.CitemName,A.ShouldAmt,A.ShouldDate,A.Custid,B.CustName" & _
                        " From " & GetOwner & "So033 A," & GetOwner & "So001 B Where A.CustId = B.CustId And A.CustID=" & txtCustId & " And A.ClctYM>=" & gmTab1YM1.Text & " And A.ClctYm <=" & gmTab1YM2.Text & _
                        " And A.RealDate Is Null " & IIf(gilServiceType.GetCodeNo & "" <> "", " And A.ServiceType='" & gilServiceType.GetCodeNo & "'", "") & _
                         IIf(strChoose <> "", " And " & strChoose, "")
            '*************************************************************************************************************************************
            If Not InsertToMDB(strSQL) Then txtCustId.SetFocus: txtCustId.Text = "": Exit Function
            txtCustId.SelStart = 0
            txtCustId.SelLength = Len(txtCustId.Text)
        Else
            MsgNoRcd
            txtCustId.Text = ""
            Exit Function
        End If
        CustIdMode = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "CustIdMode")
End Function

Private Function InsertToMDB(strSQL As String) As Boolean
    On Error GoTo chkErr
    Dim strTmpSql  As String
        If Not GetRS(rsSo3253, strSQL, gcnGi) Then MsgBox "�s�����ѡI�Э��s�ާ@�I": txtCustId.SetFocus:  Exit Function
        
        lblCustCnt = Val(lblCustCnt) + rsSo3253.RecordCount
        If Not rsSo3253.EOF Then rsSo3253.MoveFirst
         '�N��ƼȦs��Mdb
        Do While Not rsSo3253.EOF
            strSQL = "Insert Into So3253 (RowId,CustId,CustName,ClctName,BillNo,PrtSNo,CitemName,ShouldAmt,ShouldDate) Values("
            strTmpSql = "'" & rsSo3253("RowId").Value & "'," & _
                        rsSo3253("Custid").Value & ",'" & rsSo3253("CustName").Value & "'," & _
                        IIf(rsSo3253("Clctname").Value <> "", "'" & rsSo3253("ClctName").Value & "'", "Null") & _
                        ",'" & rsSo3253("BillNo").Value & "'," & IIf(rsSo3253("PrtSNo").Value <> "", "'" & rsSo3253("PrtSNo").Value & "'", "Null") & _
                        "," & IIf(rsSo3253("CitemName").Value <> "", "'" & rsSo3253("CitemName").Value & "'", "Null") & _
                        "," & IIf(rsSo3253("ShouldAmt").Value > 0, rsSo3253("ShouldAmt").Value, "") & _
                        "," & IIf(rsSo3253("ShouldDate").Value <> "", "'" & rsSo3253("ShouldDate").Value & "'", "Null") & ")"
            strSQL = strSQL & strTmpSql
             cnMDB.Execute strSQL
                 
             rsSo3253.MoveNext
        Loop
        If Not GetRS(rs3253_mdb, "Select * from So3253", cnMDB) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I": Exit Function
        Set ggrData.Recordset = rs3253_mdb
        ggrData.Refresh
        If rsSo3253.RecordCount <= 0 Then MsgNoRcd: Exit Function
        InsertToMDB = True
    Exit Function
chkErr:
    ErrSub Me.Name, "InsertToMDB"
End Function

Private Sub txtBillNo_Change()
On Error GoTo chkErr
    Dim rsMDBQuery As New ADODB.Recordset
    Dim strQuerySQL4 As String
    Dim strBill As String
    If Not IsDataOk Then Exit Sub
    Select Case Len(Trim(txtBillNo))
      Case 11
         strBill = "MediaBillNo"
      Case 12
         strBill = "PrtSNo"
      Case 15
         If InStr(1, "B,T,I,M,P", Mid(txtBillNo.Text, 7, 1)) > 0 Then
             strBill = "BillNo"
         Else
             strBill = ""
             Exit Sub
         End If
      Case 16
         If Left(txtBillNo.Text, 4) >= "9001" And Left(txtBillNo.Text, 4) <= "9912" And InStr(1, "IMPBT", Mid(txtBillNo.Text, 5, 1)) > 0 Then
            'strBillSQL = "BillNo='" & GetOldBillNo(strBillNo) & "'"
            strBill = "BillNo"
         Else
            strBill = "MediaBillNo"
            'strBillSQL = "MediaBillNo='" & strBillNo & "'"
         End If
      Case Else
         strBill = ""
         'Exit Sub
    End Select
    Screen.MousePointer = vbDefault
    If strBill = "" Then Exit Sub
    '���P�_MDB�O�_�����
    If rsMDBQuery.State = adStateOpen Then
        rsMDBQuery.Close
    End If
     rsMDBQuery.CursorLocation = adUseServer
     rsMDBQuery.Open "Select Count(KeyBillNo) from tmpSO3253A4 where KeyBillNo='" & Trim(txtBillNo) & "'", cnMDB, adOpenForwardOnly, adLockReadOnly
     If rsMDBQuery(0) <> 0 Then
        CloseRecordset rsMDBQuery
        MsgBox "����ڤw��J�L!!", vbInformation, "ĵ�i�T��"
        Exit Sub
     End If
    Screen.MousePointer = vbHourglass
    txtBillNo.Locked = True
    '�P�_�O���O���ŦX�����
    If Len(Trim(txtBillNo)) < 16 Then
      If Not GetRS(rsQuery4, GetSQLString4(strBill, Trim(txtBillNo))) Then
         txtBillNo.Locked = False
         Screen.MousePointer = vbDefault
         CloseRecordset rsQuery4
         Exit Sub
      End If
    Else
      If strBill = "BillNo" Then
         If Not GetRS(rsQuery4, GetSQLString4(strBill, GetOldBillNo(Trim(txtBillNo)))) Then
            txtBillNo.Locked = False
            Screen.MousePointer = vbDefault
            CloseRecordset rsQuery4
            Exit Sub
         End If
      Else
         If Not GetRS(rsQuery4, GetSQLString4(strBill, Trim(txtBillNo))) Then
            txtBillNo.Locked = False
            Screen.MousePointer = vbDefault
            CloseRecordset rsQuery4
            Exit Sub
         End If
      End If
    End If
'    If Not GetRS(rsQuery4, GetSQLString4(strBill, IIf(Len(Trim(txtBillNo)) < 16, Trim(txtBillNo), GetOldBillNo(Trim(txtBillNo))))) Then
'        Screen.MousePointer = vbDefault
'        CloseRecordset rsQuery4
'        Exit Sub
'    End If
    If rsQuery4.RecordCount <= 0 Then
        txtBillNo.Locked = False
        Screen.MousePointer = vbDefault
        CloseRecordset rsQuery4
        Exit Sub
    End If
    Call InsertToMDB4
    txtBillNo.Locked = False
    Screen.MousePointer = vbDefault
        If Not rsMdb.EOF Then
            rsMdb.MoveFirst
            Set ggrData4.Recordset = rsMdb
            ggrData4.Refresh
            ggrData4.Blank = False
            ggrData4.Enabled = True
            cmdQuery4.Enabled = False
            cmdOk4.Enabled = True
            ggrData4.SetFocus
        End If
    CloseRecordset rsQuery4
    CloseRecordset rsMDBQuery
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "txtBillNo_Change")

End Sub

Private Sub txtCustId_GotFocus()
    With txtCustId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtCustId_KeyPress(keyAscii As Integer)
    On Error Resume Next
        If keyAscii < vbKey0 Or keyAscii > vbKey9 Then
            If keyAscii = 8 Then Exit Sub
            keyAscii = 0
        End If
End Sub

Private Sub txtCustId_Validate(Cancel As Boolean)
    If txtCustId = "" Then Exit Sub
    Call cmdQuery1_Click
End Sub

Private Sub EnableMode()
    On Error Resume Next
        Select Case tabData.Tab
            Case 0
            Case 1
                cmdQuery.Enabled = ggrData2.Recordset.RecordCount > 0
                cmdOk2.Enabled = Not cmdQuery.Enabled
            Case 2
                cmdQuery3.Enabled = ggrData3.Recordset.RecordCount > 0
                cmdOk3.Enabled = Not cmdQuery3.Enabled
        End Select
End Sub
'���D��2905 �P�_�ŦX��ڽs����SQL�y�k 2007/01/04
Private Function GetSQLString4(ByVal strBillType As String, ByVal strBillNo As String) As String
On Error GoTo chkErr
    Dim strSQLQuery4 As String
'    strSQLQuery4 = "Select A.Custid,B.CustName,A.ClctEn,A.ClctName,A.BillNo,A.PRTSNO,A.MEDIABILLNO," & _
'                   "SUM(A.ShouldAmt) AS ShouldAmt,A.ShouldDate From " & GetOwner & "SO033 A,SO001 B Where " & _
'                   GetChoose4 & " AND A." & strBillType & "='" & strBillNo & "' " & _
'                   "Group by A.Custid,B.CustName,A.ClctEn,A.ClctName,A.BillNo,A.PRTSNO,A.MEDIABILLNO,A.ShouldDate"
    strSQLQuery4 = "Select A.Custid,B.CustName,A.ClctEn,A.ClctName,A.BillNo,A.PRTSNO,A.MEDIABILLNO," & _
                   "A.ShouldAmt,A.ShouldDate From " & GetOwner & "SO033 A," & GetOwner & "SO001 B Where " & _
                   GetChoose4 & " AND A." & strBillType & "='" & strBillNo & "' "
                   

    GetSQLString4 = strSQLQuery4
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "GetSQLString4")
End Function
'���D��2905 �bFormLoad��,�����إ�Table SO3253A4
Private Sub CreateTable()
  On Error Resume Next
        cnMDB.Execute "Drop Table tmpSO3253A4"
  On Error GoTo chkErr
    cnMDB.Execute "Create Table tmpSO3253A4 (CUSTID Long,CUSTNAME Text(50),CLCTEN Text(50)," & _
                                        "CLCTNAME TEXT(50),BILLNO Text(50),PRTSNO TEXT(50)," & _
                                        "MEDIABILLNO TEXT(50),SHOULDAMT LONG DEFAULT 0,SHOULDDATE DATE,KeyBillNo TEXT(50))"
  Exit Sub
chkErr:
   If Err.Number = -2147217900 Then
      cnMDB.Execute "Delete * from tmpSO3253A4"
      Exit Sub
   End If
    Call ErrSub(Me.Name, "CreateTable")
End Sub
'���D��2905 �N���쪺��Ʀs�btmpSO3253A4.MDB
Private Sub InsertToMDB4()
On Error GoTo chkErr
   Dim blnHaveUpdate As Boolean
   Dim rsFindBill As New ADODB.Recordset
   Dim strFindBillSQL As String
   Dim rsMemory As New ADODB.Recordset
   rsMemory.CursorLocation = adUseClient
   rsMemory.Fields.Append "RecNum", adInteger
   rsMemory.Open
    DefaultRs
    rsQuery4.MoveFirst
    blnHaveUpdate = True
    '���ˬd��ƬO�_�s�b�A�p�G���s�b�h�O���nUpdate����m
    Do While Not rsQuery4.EOF
         strFindBillSQL = "Select Count(*) From tmpSO3253A4 Where BillNo" & IIf(IsNull(rsQuery4("BillNo")), " is null", "='" & rsQuery4("BillNo") & "'") & _
                                                          " And PrtSNO" & IIf(IsNull(rsQuery4("PrtSNO")), " is null", "='" & rsQuery4("PrtSNO") & "'") & _
                                                          " And MediaBillNo" & IIf(IsNull(rsQuery4("MediaBillNo")), " is null", "='" & rsQuery4("MediaBillNo") & "'") & _
                                                          " And ShouldDate" & IIf(IsNull(rsQuery4("ShouldDate")), " is null", "=#" & rsQuery4("ShouldDate") & "#")
         If Not GetRS(rsFindBill, strFindBillSQL, cnMDB, adUseServer, adOpenForwardOnly, adLockReadOnly) Then Exit Sub
         If rsFindBill(0) = 0 Then
            rsMemory.AddNew
            rsMemory("RecNum").Value = rsQuery4.AbsolutePosition
            rsMemory.Update
         End If
         rsQuery4.MoveNext
    Loop
    cnMDB.BeginTrans
    If rsMemory.RecordCount <= 0 Then
      MsgBox "��Ƥw�s�b!!", vbInformation, "ĵ�i�T��"
      CloseRecordset rsMemory
      CloseRecordset rsFindBill
      Exit Sub
    End If
    rsMemory.MoveFirst
    While Not rsMemory.EOF
      With rsMdb
         rsQuery4.AbsolutePosition = rsMemory(0)
         .AddNew
         .Fields("CUSTID") = rsQuery4("CUSTID")
         .Fields("CUSTNAME") = rsQuery4("CUSTNAME")
         .Fields("CLCTEN") = rsQuery4("CLCTEN")
         .Fields("CLCTNAME") = rsQuery4("CLCTNAME")
         .Fields("BILLNO") = rsQuery4("BILLNO")
         .Fields("PRTSNO") = rsQuery4("PRTSNO")
         .Fields("MEDIABILLNO") = rsQuery4("MEDIABILLNO")
         .Fields("SHOULDAMT") = rsQuery4("SHOULDAMT")
         If Not IsNull(rsQuery4("SHOULDDATE")) Then
            .Fields("SHOULDDATE") = rsQuery4("SHOULDDATE")
         End If
         If blnHaveUpdate Then .Fields("KeyBillNo") = txtBillNo
         .Update
         blnHaveUpdate = False
      End With
      rsMemory.MoveNext
    Wend
    cnMDB.CommitTrans
    CloseRecordset rsMemory
    CloseRecordset rsFindBill
Exit Sub
chkErr:
    cnMDB.RollbackTrans
    Call ErrSub(Me.Name, "InsertToMDB4")
End Sub
'���D��2905 Update�ɡA���s��SQL�y�k�A�NA.CustID=B.CustID���󮳱�
Private Function UpdateWhere4(rs As Recordset) As String
  On Error GoTo chkErr
      Dim strBillType As String
      Dim strReturnWhere As String
      Select Case Len(Trim(rs("KeyBillNo")))
         Case 11
            strBillType = "MediaBillNo"
         Case 12
            strBillType = "PrtSNo"
         Case 15
            strBillType = "BillNo"
         Case 16
            If Left(rs("KeyBillNo"), 4) >= "9001" And Left(rs("KeyBillNo"), 4) <= "9912" And InStr(1, "IMPBT", Mid(rs("KeyBillNo"), 5, 1)) > 0 Then
               strBillType = "BillNo"
            Else
               strBillType = "MediaBillNo"
            End If
      End Select
      strReturnWhere = Mid(GetChoose4, 1, InStr(1, GetChoose4, "A.CustID") - 1)
      strReturnWhere = strReturnWhere & strBillType & "='" & rs("KeyBillNo") & "'"
      UpdateWhere4 = strReturnWhere
  Exit Function
chkErr:
   Call ErrSub(Me.Name, "UpdateWhere4")
End Function
Private Function GetOldBillNo(ByVal strBillNo As String) As String
    On Error Resume Next
        GetOldBillNo = GetADString(Left(strBillNo, 4), False) & Mid(strBillNo, 5, 1) & "C" & Format(Mid(strBillNo, 6), "0000000")
End Function

