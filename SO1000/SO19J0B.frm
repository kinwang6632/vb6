VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSO19J0B 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "CP �b�ȸ�Ʃ���@�~�]for �Ȥӡ^[SO19J0B]"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO19J0B.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   11880
   StartUpPosition =   1  '���ݵ�������
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
      TabCaption(0)   =   "�b�ȶ��ع����]�w"
      TabPicture(0)   =   "SO19J0B.frx":0442
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
      TabCaption(1)   =   "��ƶפJ"
      TabPicture(1)   =   "SO19J0B.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "��ƶץX"
      TabPicture(2)   =   "SO19J0B.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab3"
      Tab(2).ControlCount=   1
      Begin VB.Frame FraData 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   103
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
            Text            =   "�N��"
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
            ItemData        =   "SO19J0B.frx":0496
            Left            =   7440
            List            =   "SO19J0B.frx":04A3
            TabIndex        =   13
            Top             =   1920
            Width           =   1395
         End
         Begin VB.ComboBox cboCpProperty 
            Height          =   300
            ItemData        =   "SO19J0B.frx":04C3
            Left            =   2160
            List            =   "SO19J0B.frx":04D3
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
               Name            =   "�s�ө���"
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
               Name            =   "�s�ө���"
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
               Name            =   "�s�ө���"
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
               Name            =   "�s�ө���"
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
            Caption         =   "���O���ثe�ɦr��"
            Height          =   180
            Left            =   5880
            TabIndex        =   171
            Top             =   2355
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "�@�@�|�v (�ɶU : + )"
            Height          =   180
            Left            =   480
            TabIndex        =   113
            Top             =   1200
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "���O���� (�ɶU : - )"
            Height          =   180
            Left            =   480
            TabIndex        =   112
            Top             =   1590
            Width           =   1485
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "�@�@�|�v (�ɶU : - )"
            Height          =   180
            Left            =   480
            TabIndex        =   111
            Top             =   1995
            Width           =   1485
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "���O���� (�ɶU : + )"
            Height          =   180
            Left            =   480
            TabIndex        =   110
            Top             =   795
            Width           =   1515
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "�q�H�b��էﶵ�إN�X"
            Height          =   180
            Left            =   5520
            TabIndex        =   109
            Top             =   795
            Width           =   1800
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "�q�H�b��s�եN�X\�W��"
            Height          =   180
            Left            =   5520
            TabIndex        =   108
            Top             =   1185
            Width           =   1845
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "�q�H�C�L����"
            Height          =   180
            Left            =   6240
            TabIndex        =   107
            Top             =   1590
            Width           =   1080
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "�q�H�b�涵�إN�X\�W��"
            Height          =   180
            Left            =   5520
            TabIndex        =   106
            Top             =   390
            Width           =   1845
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "�q�H�|�v"
            Height          =   180
            Left            =   6600
            TabIndex        =   105
            Top             =   1995
            Width           =   720
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "�q�H�����ݩ�"
            Height          =   180
            Left            =   960
            TabIndex        =   104
            Top             =   385
            Width           =   1080
         End
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "F5. �C�L"
         Height          =   375
         Left            =   -72930
         TabIndex        =   16
         Top             =   7320
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "F11.�ק�"
         Height          =   375
         Left            =   -70470
         TabIndex        =   18
         Top             =   7320
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "F6. �s�W"
         Height          =   375
         Left            =   -71700
         TabIndex        =   17
         Top             =   7320
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "����(&X)"
         Height          =   375
         Left            =   -66435
         TabIndex        =   19
         Top             =   7320
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "F2. �s��"
         Height          =   375
         Left            =   -74160
         TabIndex        =   15
         Top             =   7320
         Width           =   1215
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   7335
         Left            =   -74880
         TabIndex        =   153
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   12938
         _Version        =   393216
         Tabs            =   5
         Tab             =   4
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "CP�q�H�O�P�b��ƶץX"
         TabPicture(0)   =   "SO19J0B.frx":0501
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label11"
         Tab(0).Control(1)=   "Label12"
         Tab(0).Control(2)=   "Label20"
         Tab(0).Control(3)=   "Label21"
         Tab(0).Control(4)=   "Label22"
         Tab(0).Control(5)=   "lblExportFile1"
         Tab(0).Control(6)=   "Label49"
         Tab(0).Control(7)=   "GiTempDate"
         Tab(0).Control(8)=   "GiExportDate1"
         Tab(0).Control(9)=   "GiRealDate1"
         Tab(0).Control(10)=   "GiRealDate2"
         Tab(0).Control(11)=   "txtExportFile1"
         Tab(0).Control(12)=   "txtExportPath1"
         Tab(0).Control(13)=   "cmdSelectPath1"
         Tab(0).Control(14)=   "cmdExport1"
         Tab(0).Control(15)=   "cmdCancel2"
         Tab(0).Control(16)=   "txtLogFile4"
         Tab(0).Control(17)=   "cmdChooseFile7"
         Tab(0).ControlCount=   18
         TabCaption(1)   =   "CP�믲�O��ƶץX"
         TabPicture(1)   =   "SO19J0B.frx":051D
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtLogFile5"
         Tab(1).Control(1)=   "cmdChooseFile8"
         Tab(1).Control(2)=   "cmdSelectPath2"
         Tab(1).Control(3)=   "txtExportPath2"
         Tab(1).Control(4)=   "cmdCancel3"
         Tab(1).Control(5)=   "cmdExport2"
         Tab(1).Control(6)=   "txtExportFile2"
         Tab(1).Control(7)=   "chkCanExport1"
         Tab(1).Control(8)=   "GiRealDate4"
         Tab(1).Control(9)=   "GiRealDate3"
         Tab(1).Control(10)=   "GiExportDate2"
         Tab(1).Control(11)=   "GiExportDate3"
         Tab(1).Control(12)=   "Label50"
         Tab(1).Control(13)=   "lblExportFile2"
         Tab(1).Control(14)=   "Label29"
         Tab(1).Control(15)=   "Label28"
         Tab(1).Control(16)=   "Label27"
         Tab(1).Control(17)=   "Label26"
         Tab(1).Control(18)=   "Label25"
         Tab(1).Control(19)=   "Label24"
         Tab(1).Control(20)=   "Label23"
         Tab(1).ControlCount=   21
         TabCaption(2)   =   "CP�h����ƶץX"
         TabPicture(2)   =   "SO19J0B.frx":0539
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
         TabCaption(3)   =   "CMCP�j���ƶץX"
         TabPicture(3)   =   "SO19J0B.frx":0555
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
         TabCaption(4)   =   "CP�q�H�h�O��ƶץX"
         TabPicture(4)   =   "SO19J0B.frx":0571
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "Label41"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "Label42"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "Label43"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).Control(3)=   "Label44"
         Tab(4).Control(3).Enabled=   0   'False
         Tab(4).Control(4)=   "Label45"
         Tab(4).Control(4).Enabled=   0   'False
         Tab(4).Control(5)=   "lblExportFile5"
         Tab(4).Control(5).Enabled=   0   'False
         Tab(4).Control(6)=   "Label51"
         Tab(4).Control(6).Enabled=   0   'False
         Tab(4).Control(7)=   "GiExportDate6"
         Tab(4).Control(7).Enabled=   0   'False
         Tab(4).Control(8)=   "GiRealDate5"
         Tab(4).Control(8).Enabled=   0   'False
         Tab(4).Control(9)=   "GiRealDate6"
         Tab(4).Control(9).Enabled=   0   'False
         Tab(4).Control(10)=   "cmdCancel6"
         Tab(4).Control(10).Enabled=   0   'False
         Tab(4).Control(11)=   "cmdExport5"
         Tab(4).Control(11).Enabled=   0   'False
         Tab(4).Control(12)=   "cmdSelectPath5"
         Tab(4).Control(12).Enabled=   0   'False
         Tab(4).Control(13)=   "txtExportPath5"
         Tab(4).Control(13).Enabled=   0   'False
         Tab(4).Control(14)=   "txtExportFile5"
         Tab(4).Control(14).Enabled=   0   'False
         Tab(4).Control(15)=   "txtLogFile6"
         Tab(4).Control(15).Enabled=   0   'False
         Tab(4).Control(16)=   "cmdChooseFile9"
         Tab(4).Control(16).Enabled=   0   'False
         Tab(4).ControlCount=   17
         Begin VB.CommandButton cmdChooseFile9 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9000
            TabIndex        =   92
            Top             =   4560
            Width           =   420
         End
         Begin VB.TextBox txtLogFile6 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            TabIndex        =   91
            Top             =   4560
            Width           =   5565
         End
         Begin VB.TextBox txtLogFile5 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   60
            Top             =   4800
            Width           =   5565
         End
         Begin VB.CommandButton cmdChooseFile8 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   61
            Top             =   4800
            Width           =   420
         End
         Begin VB.CommandButton cmdChooseFile7 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   49
            Top             =   4560
            Width           =   420
         End
         Begin VB.TextBox txtLogFile4 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   48
            Top             =   4560
            Width           =   5565
         End
         Begin VB.TextBox txtExportFile5 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            TabIndex        =   87
            Top             =   2040
            Width           =   3700
         End
         Begin VB.TextBox txtExportPath5 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            TabIndex        =   89
            Top             =   3720
            Width           =   5565
         End
         Begin VB.CommandButton cmdSelectPath5 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9000
            TabIndex        =   90
            Top             =   3720
            Width           =   420
         End
         Begin VB.CommandButton cmdExport5 
            Caption         =   "�ץX"
            Height          =   375
            Left            =   3960
            TabIndex        =   93
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel6 
            Caption         =   "����"
            Height          =   375
            Left            =   6435
            TabIndex        =   94
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdSelectPath4 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   82
            Top             =   4580
            Width           =   420
         End
         Begin VB.TextBox txtExportPath4 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   81
            Top             =   4580
            Width           =   5565
         End
         Begin VB.CommandButton cmdCancel5 
            Caption         =   "����"
            Height          =   375
            Left            =   -68565
            TabIndex        =   84
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdExport4 
            Caption         =   "�ץX"
            Height          =   375
            Left            =   -71040
            TabIndex        =   83
            Top             =   5880
            Width           =   1215
         End
         Begin VB.TextBox txtExportFile4 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   79
            Top             =   3180
            Width           =   3700
         End
         Begin VB.CommandButton cmdSelectPath3 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   72
            Top             =   4680
            Width           =   420
         End
         Begin VB.TextBox txtExportPath3 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   71
            Top             =   4680
            Width           =   5565
         End
         Begin VB.CommandButton cmdCancel4 
            Caption         =   "����"
            Height          =   375
            Left            =   -68565
            TabIndex        =   74
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdExport3 
            Caption         =   "�ץX"
            Height          =   375
            Left            =   -71040
            TabIndex        =   73
            Top             =   5880
            Width           =   1215
         End
         Begin VB.TextBox txtExportFile3 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   69
            Top             =   3480
            Width           =   3700
         End
         Begin VB.CommandButton cmdSelectPath2 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   59
            Top             =   4200
            Width           =   420
         End
         Begin VB.TextBox txtExportPath2 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   58
            Top             =   4200
            Width           =   5565
         End
         Begin VB.CommandButton cmdCancel3 
            Caption         =   "����"
            Height          =   375
            Left            =   -68565
            TabIndex        =   63
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdExport2 
            Caption         =   "�ץX"
            Height          =   375
            Left            =   -71040
            TabIndex        =   62
            Top             =   5880
            Width           =   1215
         End
         Begin VB.TextBox txtExportFile2 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   56
            Top             =   3000
            Width           =   3700
         End
         Begin VB.CheckBox chkCanExport1 
            Height          =   330
            Left            =   -71760
            TabIndex        =   55
            Top             =   2400
            Width           =   375
         End
         Begin VB.CommandButton cmdCancel2 
            Caption         =   "����"
            Height          =   375
            Left            =   -68565
            TabIndex        =   51
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdExport1 
            Caption         =   "�ץX"
            Height          =   375
            Left            =   -71040
            TabIndex        =   50
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdSelectPath1 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   47
            Top             =   3720
            Width           =   420
         End
         Begin VB.TextBox txtExportPath1 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   46
            Top             =   3720
            Width           =   5565
         End
         Begin VB.TextBox txtExportFile1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   44
            Top             =   2040
            Width           =   3700
         End
         Begin Gi_Date.GiDate GiRealDate2 
            Height          =   330
            Left            =   -70080
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   45
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
            Left            =   -70200
            TabIndex        =   53
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
            Left            =   -71760
            TabIndex        =   52
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
            Left            =   -71760
            TabIndex        =   54
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
            Left            =   -71760
            TabIndex        =   57
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
            TabIndex        =   64
            Top             =   1080
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   582
            ButtonCaption   =   "������O"
            FontSize        =   9
            FontName        =   "�s�ө���"
         End
         Begin Gi_Date.GiDate GiPrDate2 
            Height          =   330
            Left            =   -70200
            TabIndex        =   67
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
            TabIndex        =   66
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
            TabIndex        =   68
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
            TabIndex        =   70
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
            TabIndex        =   75
            Top             =   1080
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   582
            ButtonCaption   =   "������O"
            FontSize        =   9
            FontName        =   "�s�ө���"
         End
         Begin Gi_Date.GiDate GiPrDate4 
            Height          =   330
            Left            =   -70155
            TabIndex        =   78
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
            TabIndex        =   77
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
            TabIndex        =   80
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
            TabIndex        =   149
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
            Left            =   4920
            TabIndex        =   86
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
            Left            =   3240
            TabIndex        =   85
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
            Left            =   3240
            TabIndex        =   88
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
            TabIndex        =   65
            Top             =   1680
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   661
            ButtonCaption   =   "�� �� �� �� �]"
            FontSize        =   9
            FontName        =   "�s�ө���"
         End
         Begin Gi_Multi.GiMulti gimUCCode 
            Height          =   375
            Left            =   -73720
            TabIndex        =   76
            Top             =   1780
            Width           =   7605
            _ExtentX        =   13414
            _ExtentY        =   661
            ButtonCaption   =   "�� �� �� �� �]"
            FontSize        =   9
            FontName        =   "�s�ө���"
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "���D�O����"
            Height          =   180
            Left            =   2100
            TabIndex        =   165
            Top             =   4635
            Width           =   900
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "���D�O����"
            Height          =   180
            Left            =   -72900
            TabIndex        =   164
            Top             =   4875
            Width           =   900
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "���D�O����"
            Height          =   180
            Left            =   -72960
            TabIndex        =   163
            Top             =   4635
            Width           =   900
         End
         Begin VB.Label lblExportFile5 
            AutoSize        =   -1  'True
            Height          =   180
            Left            =   3240
            TabIndex        =   159
            Top             =   5040
            Width           =   45
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "�ꦬ���"
            Height          =   180
            Left            =   2280
            TabIndex        =   158
            Top             =   1275
            Width           =   720
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   4560
            TabIndex        =   157
            Top             =   1275
            Width           =   180
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "�����ɦW"
            Height          =   180
            Left            =   2280
            TabIndex        =   156
            Top             =   2115
            Width           =   720
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   2280
            TabIndex        =   155
            Top             =   2955
            Width           =   720
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "�ץX�ɮ׸��|"
            Height          =   180
            Left            =   1920
            TabIndex        =   154
            Top             =   3795
            Width           =   1080
         End
         Begin VB.Label lblExportFile4 
            AutoSize        =   -1  'True
            Height          =   180
            Left            =   -71760
            TabIndex        =   152
            Top             =   4920
            Width           =   45
         End
         Begin VB.Label lblExportFile3 
            AutoSize        =   -1  'True
            Height          =   180
            Left            =   -71760
            TabIndex        =   151
            Top             =   5160
            Width           =   45
         End
         Begin VB.Label lblExportFile2 
            AutoSize        =   -1  'True
            Height          =   180
            Left            =   -71760
            TabIndex        =   150
            Top             =   5280
            Width           =   45
         End
         Begin VB.Label lblExportFile1 
            AutoSize        =   -1  'True
            Height          =   180
            Left            =   -71760
            TabIndex        =   148
            Top             =   5040
            Width           =   45
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "�ץX�ɮ׸��|"
            Height          =   180
            Left            =   -73080
            TabIndex        =   147
            Top             =   4655
            Width           =   1080
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   -72720
            TabIndex        =   146
            Top             =   3955
            Width           =   720
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "�����ɦW"
            Height          =   180
            Left            =   -72720
            TabIndex        =   145
            Top             =   3255
            Width           =   720
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "������u��(�h����)"
            Height          =   180
            Left            =   -73560
            TabIndex        =   144
            Top             =   2555
            Width           =   1560
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   -70515
            TabIndex        =   143
            Top             =   2555
            Width           =   180
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "�ץX�ɮ׸��|"
            Height          =   180
            Left            =   -73080
            TabIndex        =   142
            Top             =   4755
            Width           =   1080
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   -72720
            TabIndex        =   141
            Top             =   4155
            Width           =   720
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "�����ɦW"
            Height          =   180
            Left            =   -72720
            TabIndex        =   140
            Top             =   3555
            Width           =   720
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "�b�ک���ѦҤ�"
            Height          =   180
            Left            =   -73260
            TabIndex        =   139
            Top             =   2955
            Width           =   1260
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "������u��(�h����)"
            Height          =   180
            Left            =   -73560
            TabIndex        =   138
            Top             =   2355
            Width           =   1560
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   -70560
            TabIndex        =   137
            Top             =   2355
            Width           =   180
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "�ץX�ɮ׸��|"
            Height          =   180
            Left            =   -73080
            TabIndex        =   136
            Top             =   4275
            Width           =   1080
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   -72720
            TabIndex        =   135
            Top             =   3675
            Width           =   720
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "�����ɦW"
            Height          =   180
            Left            =   -72720
            TabIndex        =   134
            Top             =   3075
            Width           =   720
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "�e���w���O�_����"
            Height          =   180
            Left            =   -73440
            TabIndex        =   133
            Top             =   2475
            Width           =   1440
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "�b�ک�����"
            Height          =   180
            Left            =   -73080
            TabIndex        =   132
            Top             =   1875
            Width           =   1080
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   -70560
            TabIndex        =   131
            Top             =   1275
            Width           =   180
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "�ꦬ���"
            Height          =   180
            Left            =   -72720
            TabIndex        =   130
            Top             =   1275
            Width           =   720
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "�ץX�ɮ׸��|"
            Height          =   180
            Left            =   -73080
            TabIndex        =   129
            Top             =   3795
            Width           =   1080
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   -72720
            TabIndex        =   128
            Top             =   2955
            Width           =   720
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "�����ɦW"
            Height          =   180
            Left            =   -72720
            TabIndex        =   127
            Top             =   2115
            Width           =   720
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   -70440
            TabIndex        =   126
            Top             =   1275
            Width           =   180
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "�ꦬ���"
            Height          =   180
            Left            =   -72720
            TabIndex        =   125
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
            Name            =   "�s�ө���"
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
         TabIndex        =   115
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   12938
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "CP�q�H�O��ƶפJ"
         TabPicture(0)   =   "SO19J0B.frx":058D
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label13"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label14"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label15"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label52"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label53"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtImportFile1"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtLogFile1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cmdCancel1"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "cmdInport1"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "cmdChooseFile1"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "cmdChooseFile2"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Frame1"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "ComdPath"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "PB1"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "txtShouldDay"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "CP�q�ܩ��Ӹ�ƶפJ"
         TabPicture(1)   =   "SO19J0B.frx":05A9
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtNumCount"
         Tab(1).Control(1)=   "txtAreaCode"
         Tab(1).Control(2)=   "txtbillYM"
         Tab(1).Control(3)=   "txtLogFile3"
         Tab(1).Control(4)=   "cmdChooseFile6"
         Tab(1).Control(5)=   "cmdCancel8"
         Tab(1).Control(6)=   "cmdInport3"
         Tab(1).Control(7)=   "Frame2"
         Tab(1).Control(8)=   "cmdChooseFile5"
         Tab(1).Control(9)=   "txtImportFile5"
         Tab(1).Control(10)=   "PB2"
         Tab(1).Control(11)=   "Label58"
         Tab(1).Control(12)=   "Label57"
         Tab(1).Control(13)=   "Label56"
         Tab(1).Control(14)=   "Label48"
         Tab(1).Control(15)=   "Label19"
         Tab(1).ControlCount=   16
         TabCaption(2)   =   "�򥻸�Ʈֹ�פJ"
         TabPicture(2)   =   "SO19J0B.frx":05C5
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label46"
         Tab(2).Control(1)=   "Label47"
         Tab(2).Control(2)=   "Label55"
         Tab(2).Control(3)=   "PB3"
         Tab(2).Control(4)=   "txtImportFile2"
         Tab(2).Control(5)=   "txtLogFile2"
         Tab(2).Control(6)=   "cmdChooseFile3"
         Tab(2).Control(7)=   "cmdChooseFile4"
         Tab(2).Control(8)=   "cmdInport2"
         Tab(2).Control(9)=   "cmdCancel7"
         Tab(2).Control(10)=   "chkTFNSID"
         Tab(2).Control(11)=   "txtfilter"
         Tab(2).ControlCount=   12
         Begin VB.TextBox txtNumCount 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   30
            Top             =   960
            Width           =   465
         End
         Begin VB.TextBox txtAreaCode 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -69480
            TabIndex        =   31
            Top             =   1560
            Width           =   1785
         End
         Begin VB.TextBox txtbillYM 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -69480
            TabIndex        =   29
            Top             =   960
            Width           =   1785
         End
         Begin VB.TextBox txtfilter 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71880
            TabIndex        =   100
            Top             =   3720
            Width           =   3135
         End
         Begin VB.CheckBox chkTFNSID 
            Caption         =   "�}�q�X���s���O�_�����n��"
            Height          =   255
            Left            =   -71880
            TabIndex        =   95
            Top             =   1560
            Value           =   1  '�֨�
            Width           =   2775
         End
         Begin VB.TextBox txtShouldDay 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            MaxLength       =   2
            TabIndex        =   22
            Text            =   "1"
            Top             =   2400
            Width           =   400
         End
         Begin MSComctlLib.ProgressBar PB1 
            Height          =   375
            Left            =   1440
            TabIndex        =   166
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
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   38
            Top             =   4200
            Width           =   5565
         End
         Begin VB.CommandButton cmdChooseFile6 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   39
            Top             =   4200
            Width           =   420
         End
         Begin VB.CommandButton cmdCancel7 
            Caption         =   "����"
            Height          =   375
            Left            =   -68565
            TabIndex        =   102
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdInport2 
            Caption         =   "�פJ"
            Height          =   375
            Left            =   -71040
            TabIndex        =   101
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdChooseFile4 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66120
            TabIndex        =   99
            Top             =   2880
            Width           =   420
         End
         Begin VB.CommandButton cmdChooseFile3 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66120
            TabIndex        =   97
            Top             =   2040
            Width           =   420
         End
         Begin VB.TextBox txtLogFile2 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71880
            TabIndex        =   98
            Top             =   2880
            Width           =   5565
         End
         Begin VB.TextBox txtImportFile2 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71880
            TabIndex        =   96
            Top             =   2040
            Width           =   5565
         End
         Begin MSComDlg.CommonDialog ComdPath 
            Left            =   720
            Top             =   5880
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton cmdCancel8 
            Caption         =   "����"
            Height          =   375
            Left            =   -68565
            TabIndex        =   41
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdInport3 
            Caption         =   "�פJ"
            Height          =   375
            Left            =   -71040
            TabIndex        =   40
            Top             =   5880
            Width           =   1215
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  '�S���ؽu
            Height          =   1140
            Left            =   -73200
            TabIndex        =   121
            Top             =   2160
            Width           =   7815
            Begin VB.TextBox txtImportFile4 
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3720
               TabIndex        =   35
               Text            =   "CNS_RATING_M"
               Top             =   675
               Width           =   3700
            End
            Begin VB.OptionButton optDetailDaily 
               Caption         =   "�C��"
               Height          =   180
               Left            =   1440
               TabIndex        =   32
               Top             =   75
               Value           =   -1  'True
               Width           =   800
            End
            Begin VB.TextBox txtImportFile3 
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3720
               TabIndex        =   33
               Text            =   "CNS_RATING_D"
               Top             =   0
               Width           =   3700
            End
            Begin VB.OptionButton optDetailMonthly 
               Caption         =   "�C��"
               Height          =   180
               Left            =   1440
               TabIndex        =   34
               Top             =   750
               Width           =   800
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "�ˮֶפJ�ɦW"
               Height          =   180
               Left            =   2400
               TabIndex        =   124
               Top             =   750
               Width           =   1080
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "�ˮֶפJ�ɦW"
               Height          =   180
               Left            =   2400
               TabIndex        =   123
               Top             =   75
               Width           =   1080
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "�q�ܩ��Ӻ���"
               Height          =   180
               Left            =   0
               TabIndex        =   122
               Top             =   75
               Width           =   1080
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  '�S���ؽu
            Height          =   300
            Left            =   3240
            TabIndex        =   120
            Top             =   1680
            Width           =   2775
            Begin VB.OptionButton optEmpNo 
               Caption         =   "�˾��a�}"
               Height          =   250
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Width           =   1035
            End
            Begin VB.OptionButton optEmpName 
               Caption         =   "���O�a�}"
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
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -66000
            TabIndex        =   37
            Top             =   3600
            Width           =   420
         End
         Begin VB.TextBox txtImportFile5 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -71760
            TabIndex        =   36
            Top             =   3600
            Width           =   5565
         End
         Begin VB.CommandButton cmdChooseFile2 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9000
            TabIndex        =   26
            Top             =   3840
            Width           =   420
         End
         Begin VB.CommandButton cmdChooseFile1 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9000
            TabIndex        =   24
            Top             =   3120
            Width           =   420
         End
         Begin VB.CommandButton cmdInport1 
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3960
            TabIndex        =   27
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel1 
            Caption         =   "����"
            Height          =   375
            Left            =   6435
            TabIndex        =   28
            Top             =   5880
            Width           =   1215
         End
         Begin VB.TextBox txtLogFile1 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            TabIndex        =   25
            Top             =   3840
            Width           =   5565
         End
         Begin VB.TextBox txtImportFile1 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            TabIndex        =   23
            Top             =   3120
            Width           =   5565
         End
         Begin MSComctlLib.ProgressBar PB2 
            Height          =   375
            Left            =   -73200
            TabIndex        =   167
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
            Left            =   -73260
            TabIndex        =   168
            Top             =   4680
            Visible         =   0   'False
            Width           =   7560
            _ExtentX        =   13335
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            Caption         =   "�o�ܤ�q�ܸ��X��"
            Height          =   180
            Left            =   -67560
            TabIndex        =   175
            Top             =   1080
            Width           =   1440
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "�q�H�O�X�b�~��"
            Height          =   180
            Left            =   -71040
            TabIndex        =   174
            Top             =   1080
            Width           =   1260
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "�o�ܤ�ϰ츹�X"
            Height          =   180
            Left            =   -71040
            TabIndex        =   173
            Top             =   1680
            Width           =   1260
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "�ư��m�W�ٿפ���蠟�r��"
            Height          =   300
            Left            =   -74280
            TabIndex        =   172
            Top             =   3840
            Width           =   2160
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "(�����鶷���� 1��31 ����)"
            Height          =   180
            Left            =   3840
            TabIndex        =   170
            Top             =   2475
            Width           =   2100
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   2400
            TabIndex        =   169
            Top             =   2475
            Width           =   540
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "���D�O����"
            Height          =   180
            Left            =   -72960
            TabIndex        =   162
            Top             =   4275
            Width           =   900
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "���D�O����"
            Height          =   180
            Left            =   -73080
            TabIndex        =   161
            Top             =   2955
            Width           =   900
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "�ˮֶפJ�ɦW"
            Height          =   180
            Left            =   -73260
            TabIndex        =   160
            Top             =   2115
            Width           =   1080
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "�פJ�ɮ�"
            Height          =   180
            Left            =   -72840
            TabIndex        =   119
            Top             =   3720
            Width           =   720
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "���D�O����"
            Height          =   180
            Left            =   2040
            TabIndex        =   118
            Top             =   3915
            Width           =   900
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "�פJ�ɮ�"
            Height          =   180
            Left            =   2220
            TabIndex        =   117
            Top             =   3195
            Width           =   720
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "���O�椧�a�}�̾�"
            Height          =   180
            Left            =   1440
            TabIndex        =   116
            Top             =   1710
            Width           =   1440
         End
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  '�m�����
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  '��u�T�w
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   114
         Top             =   7380
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmSO19J0B"
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
Private Fso As New FileSystemObject
Private strmTextFile1, strmTextFile2 As TextStream
Private lngEditMode As giEditModeEnu
Private strErrPath As Variant
Private ErrorType As Boolean
Private iniPages As New Scripting.Dictionary
Private strIniFile As String

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
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cboCpProperty_Click")
End Sub

Private Sub cmdAdd_Click()
On Error GoTo ChkErr
  '�s�W
  Call ChangeMode(giEditModeInsert)
  Call NewRcd
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdAdd_Click")
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ChkErr
  '�Y�O�s���Ҧ��A�h�����{��
  '�_�h�^���s���Ҧ�
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
  '��ܶפJ����r��
  With ComdPath
    .FileName = ImportFile.Text
    .Filter = "��r�� (*.txt;*.dat)|*.txt;*.dat|�Ҧ��ɮ� (*.*)|*.*"
    .Action = 1
    ImportFile.Text = .FileName
  End With
  
  '�۰ʮھڤ�r�ɪ��ɦW�]�w������ LOG ���ɦW
  Dim ImportFileName As String
  ImportFileName = ImportFile.Text
  Call SetLogFileName(ImportFileName, LogFile)
  
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "SelectImportFile")
End Sub

Private Sub cmdChooseFile2_Click()
  Call SelectImportLog(txtLogFile1)
End Sub

Private Sub SelectImportLog(ByRef LogFile As TextBox)
On Error GoTo ChkErr
  '��� LOG ��
  With ComdPath
    .FileName = LogFile.Text
    .Filter = "��r�� (*.txt;*.dat)|*.txt;*.dat|�Ҧ��ɮ� (*.*)|*.*"
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
  'Call ScrToRcdtxt
  ''#3195 �N���󪺭ȼg�JIni�� By Kin 2007/07/13
  Call WriteSO19J0Ini(iniPages, strIniFile, Me)
  'CP�q�H�O�P�b��ƶץX
  If GiRealDate1.GetValue = "" Then
    MsgBox "�Х����w�ꦬ���(�_)�C", vbExclamation, "ĵ�i"
    GiRealDate1.SetFocus
  ElseIf GiRealDate2.GetValue = "" Then
    MsgBox "�Х����w�ꦬ���(��)�C", vbExclamation, "ĵ�i"
    GiRealDate2.SetFocus
  ElseIf GiRealDate2.GetValue < GiRealDate1.GetValue Then
    MsgBox "�ꦬ���(��)�����j�󵥩�ꦬ���(�_)�C", vbExclamation, "ĵ�i"
    GiRealDate2.SetFocus
  ElseIf txtExportFile1.Text = "" Then
    MsgBox "�Х����w�����ɦW�C", vbExclamation, "ĵ�i"
    If txtExportFile1 = "" Then txtExportFile1 = "NVW_PIH_U"
    txtExportFile1.SetFocus
  ElseIf GiExportDate1.GetValue = "" Then
    MsgBox "�Х����w�������C", vbExclamation, "ĵ�i"
    GiExportDate1.SetFocus
  ElseIf txtExportPath1.Text = "" Then
    MsgBox "�Х����w�ץX�ɮ׸��|�C", vbExclamation, "ĵ�i"
    txtExportPath1.SetFocus
  ElseIf txtLogFile4.Text = "" Then
    MsgBox "�Х����w���D�O���ɡC", vbExclamation, "ĵ�i"
    txtLogFile4.SetFocus
  ElseIf Replace(txtExportPath1.Text & "\" & lblExportFile1.Caption, "\\", "\") = txtLogFile4.Text Then
    MsgBox "�פJ�ɮשM���D�O���ɤ��i�H�@�ˡC", vbExclamation, "ĵ�i"
    txtLogFile4.SetFocus
  Else
    If Fso.FileExists(txtLogFile4.Text) Then
      Fso.DeleteFile (txtLogFile4.Text)
    End If
    Set strmTextFile2 = Fso.OpenTextFile(txtLogFile4.Text, ForAppending, True)
    strmTextFile2.Write vbCrLf & "################################################################################"
    strmTextFile2.Write vbCrLf & "CP�q�H�O�P�b��ƶץX�@��ƭ��ƶץX�O���@�O������G" & Now & vbCrLf & vbCrLf
    
    Set strmTextFile1 = Fso.CreateTextFile(txtExportPath1.Text & "\" & lblExportFile1.Caption, True)
    Call ExportData1
    Set strmTextFile1 = Nothing
    
    Set strmTextFile2 = Nothing
  End If
ChkErr:
  Call ErrSub(Me.Name, "cmdExport1_Click")
End Sub

Private Sub cmdExport2_Click()
On Error GoTo ChkErr
  'Call ScrToRcdtxt
  '#3195 �N���󪺭ȼg�JIni�� By Kin 2007/07/13
  Call WriteSO19J0Ini(iniPages, strIniFile, Me)
  'CP�믲�O��ƶץX
  If GiRealDate3.GetValue = "" Then
    MsgBox "�Х����w�ꦬ���(�_)�C", vbExclamation, "ĵ�i"
    GiRealDate3.SetFocus
  ElseIf GiRealDate4.GetValue = "" Then
    MsgBox "�Х����w�ꦬ���(��)�C", vbExclamation, "ĵ�i"
    GiRealDate4.SetFocus
  ElseIf GiRealDate4.GetValue < GiRealDate3.GetValue Then
    MsgBox "�ꦬ���(��)�����j�󵥩�ꦬ���(�_)�C", vbExclamation, "ĵ�i"
    GiRealDate4.SetFocus
  ElseIf GiExportDate2.GetValue = "" Then
    MsgBox "�Х����w�b�ک������C", vbExclamation, "ĵ�i"
    GiExportDate2.SetFocus
  ElseIf txtExportFile2.Text = "" Then
    MsgBox "�Х����w�����ɦW�C", vbExclamation, "ĵ�i"
    If txtExportFile2 = "" Then txtExportFile2 = "NVW_PIH_R"
    txtExportFile2.SetFocus
  ElseIf GiExportDate3.GetValue = "" Then
    MsgBox "�Х����w�������C", vbExclamation, "ĵ�i"
    GiExportDate3.SetFocus
  ElseIf txtExportPath2.Text = "" Then
    MsgBox "�Х����w�ץX�ɮ׸��|�C", vbExclamation, "ĵ�i"
    txtExportPath2.SetFocus
  ElseIf txtLogFile5.Text = "" Then
    MsgBox "�Х����w���D�O���ɡC", vbExclamation, "ĵ�i"
    txtLogFile5.SetFocus
  ElseIf Replace(txtExportPath2.Text & "\" & lblExportFile2.Caption, "\\", "\") = txtLogFile5.Text Then
    MsgBox "�פJ�ɮשM���D�O���ɤ��i�H�@�ˡC", vbExclamation, "ĵ�i"
    txtLogFile5.SetFocus
  Else
    If Fso.FileExists(txtLogFile5.Text) Then
      Fso.DeleteFile (txtLogFile5.Text)
    End If
    Set strmTextFile2 = Fso.OpenTextFile(txtLogFile5.Text, ForAppending, True)
    strmTextFile2.Write vbCrLf & "################################################################################"
    strmTextFile2.Write vbCrLf & "CP�믲�O��ƶץX�@��ƭ��ƶץX�O���@�O������G" & Now & vbCrLf & vbCrLf
    
    Set strmTextFile1 = Fso.CreateTextFile(txtExportPath2.Text & "\" & lblExportFile2.Caption, True)
    Call ExportData2
    Set strmTextFile1 = Nothing
    
    Set strmTextFile2 = Nothing
  End If
ChkErr:
  Call ErrSub(Me.Name, "cmdExport2_Click")
End Sub

Private Sub cmdExport3_Click()
On Error GoTo ChkErr
  'Call ScrToRcdtxt
  '#3195 �N���󪺭ȼg�JIni�� By Kin 2007/07/13
  Call WriteSO19J0Ini(iniPages, strIniFile, Me)
  'CP�h����ƶץX
  '�p�G������O�S������A�h�w�]������
  If gimPrCode1.GetQueryCode = "" Then
    gimPrCode1.SelectAll
  End If
  If GiPrDate1.GetValue = "" Then
    MsgBox "�Х����w������u��(�h����)(�_)�C", vbExclamation, "ĵ�i"
    GiPrDate1.SetFocus
  ElseIf GiPrDate2.GetValue = "" Then
    MsgBox "�Х����w������u��(�h����)(��)�C", vbExclamation, "ĵ�i"
    GiPrDate2.SetFocus
  ElseIf GiPrDate2.GetValue < GiPrDate1.GetValue Then
    MsgBox "������u��(�h����)(��)�����j�󵥩������u��(�h����)(�_)�C", vbExclamation, "ĵ�i"
    GiPrDate2.SetFocus
  ElseIf GiExportRefDate1.GetValue = "" Then
    MsgBox "�Х����w�b�ک���ѦҤ�C", vbExclamation, "ĵ�i"
    GiExportRefDate1.SetFocus
  ElseIf txtExportFile3.Text = "" Then
    MsgBox "�Х����w�����ɦW�C", vbExclamation, "ĵ�i"
    If txtExportFile3 = "" Then txtExportFile3 = "NVW_REMOVE_R"
    txtExportFile3.SetFocus
  ElseIf GiExportDate4.GetValue = "" Then
    MsgBox "�Х����w�������C", vbExclamation, "ĵ�i"
    GiExportDate4.SetFocus
  ElseIf txtExportPath3 = "" Then
    MsgBox "�Х����w�ץX�ɮ׸��|�C", vbExclamation, "ĵ�i"
    txtExportPath3.SetFocus
  ElseIf gimPrCode1.GetQueryCode = "" Then
    MsgBox "�Х����w������O�C", vbExclamation, "ĵ�i"
    gimPrCode1.SetFocus
  ElseIf gimReasonCode.GetQueryCode = "" Then
    MsgBox "�Х����w�������]�C", vbExclamation, "ĵ�i"
    gimReasonCode.SetFocus
  Else
    Set strmTextFile1 = Fso.CreateTextFile(txtExportPath3.Text & "\" & lblExportFile3.Caption, True)
    Call ExportData3
    Set strmTextFile1 = Nothing
  End If
ChkErr:
  Call ErrSub(Me.Name, "cmdExport3_Click")
End Sub

Private Sub cmdExport4_Click()
On Error GoTo ChkErr
  'Call ScrToRcdtxt
  '#3195 �N���󪺭ȼg�JIni�� By Kin 2007/07/13
  Call WriteSO19J0Ini(iniPages, strIniFile, Me)
  'CMCP�j���ƶץX
  '�p�G������O�S������A�h�w�]������
  If gimPrCode2.GetQueryCode = "" Then
    gimPrCode2.SelectAll
  End If
  If GiPrDate3.GetValue = "" Then
    MsgBox "�Х����w������u��(�h����)(�_)�C", vbExclamation, "ĵ�i"
    GiPrDate3.SetFocus
  ElseIf GiPrDate4.GetValue = "" Then
    MsgBox "�Х����w������u��(�h����)(��)�C", vbExclamation, "ĵ�i"
    GiPrDate4.SetFocus
  ElseIf GiPrDate4.GetValue < GiPrDate3.GetValue Then
    MsgBox "������u��(�h����)(��)�����j�󵥩������u��(�h����)(�_)�C", vbExclamation, "ĵ�i"
    GiPrDate4.SetFocus
  ElseIf txtExportFile4.Text = "" Then
    MsgBox "�Х����w�����ɦW�C", vbExclamation, "ĵ�i"
    If txtExportFile4 = "" Then txtExportFile4 = "NVW_REMOVE_"
    txtExportFile4.SetFocus
  ElseIf GiExportDate5.GetValue = "" Then
    MsgBox "�Х����w�������C", vbExclamation, "ĵ�i"
    GiExportDate5.SetFocus
  ElseIf txtExportPath4 = "" Then
    MsgBox "�Х����w�ץX�ɮ׸��|�C", vbExclamation, "ĵ�i"
    txtExportPath4.SetFocus
  ElseIf gimPrCode2.GetQueryCode = "" Then
    MsgBox "�Х����w������O�C", vbExclamation, "ĵ�i"
    gimPrCode2.SetFocus
  ElseIf gimUCCode.GetQueryCode = "" Then
    MsgBox "�Х����w�������]�C", vbExclamation, "ĵ�i"
    gimUCCode.SetFocus
  Else
    Set strmTextFile1 = Fso.CreateTextFile(txtExportPath4.Text & "\" & lblExportFile4.Caption, True)
    Call ExportData4
    Set strmTextFile1 = Nothing
  End If
ChkErr:
  Call ErrSub(Me.Name, "cmdExport4_Click")
End Sub

Private Sub cmdExport5_Click()
On Error GoTo ChkErr
  'Call ScrToRcdtxt
  '#3195 �N���󪺭ȼg�JIni�� By Kin 2007/07/13
  Call WriteSO19J0Ini(iniPages, strIniFile, Me)
  'CP�q�H�h�O��ƶץX
  If GiRealDate5.GetValue = "" Then
    MsgBox "�Х����w�ꦬ���(�_)�C", vbExclamation, "ĵ�i"
    GiRealDate5.SetFocus
  ElseIf GiRealDate6.GetValue = "" Then
    MsgBox "�Х����w�ꦬ���(��)�C", vbExclamation, "ĵ�i"
    GiRealDate6.SetFocus
  ElseIf GiRealDate6.GetValue < GiRealDate5.GetValue Then
    MsgBox "�ꦬ���(��)�����j�󵥩�ꦬ���(�_)�C", vbExclamation, "ĵ�i"
    GiRealDate6.SetFocus
  ElseIf txtExportFile5.Text = "" Then
    MsgBox "�Х����w�����ɦW�C", vbExclamation, "ĵ�i"
    If txtExportFile5 = "" Then txtExportFile5 = "NVW_REPAY_U"
    txtExportFile5.SetFocus
  ElseIf GiExportDate6.GetValue = "" Then
    MsgBox "�Х����w�������C", vbExclamation, "ĵ�i"
    GiExportDate6.SetFocus
  ElseIf txtExportPath5 = "" Then
    MsgBox "�Х����w�ץX�ɮ׸��|�C", vbExclamation, "ĵ�i"
    txtExportPath5.SetFocus
  ElseIf txtLogFile6.Text = "" Then
    MsgBox "�Х����w���D�O���ɡC", vbExclamation, "ĵ�i"
    txtLogFile6.SetFocus
  ElseIf Replace(txtExportPath5.Text & "\" & lblExportFile5.Caption, "\\", "\") = txtLogFile6.Text Then
    MsgBox "�פJ�ɮשM���D�O���ɤ��i�H�@�ˡC", vbExclamation, "ĵ�i"
    txtLogFile6.SetFocus
  Else
    If Fso.FileExists(txtLogFile6.Text) Then
      Fso.DeleteFile (txtLogFile6.Text)
    End If
    Set strmTextFile2 = Fso.OpenTextFile(txtLogFile6.Text, ForAppending, True)
    strmTextFile2.Write vbCrLf & "################################################################################"
    strmTextFile2.Write vbCrLf & "CP�q�H�h�O��ƶץX�@��ƭ��ƶץX�O���@�O������G" & Now & vbCrLf & vbCrLf
    
    Set strmTextFile1 = Fso.CreateTextFile(txtExportPath5.Text & "\" & lblExportFile5.Caption, True)
    Call ExportData5
    Set strmTextFile1 = Nothing
    
    Set strmTextFile2 = Nothing
  End If
ChkErr:
  Call ErrSub(Me.Name, "cmdExport5_Click")
End Sub

Private Sub cmdInport1_Click()
On Error GoTo ChkErr
 '#3195 �N���󪺭ȼg�JIni�� By Kin 2007/07/13
  Call WriteSO19J0Ini(iniPages, strIniFile, Me)
  'CP�q�H�O��ƶפJ
  If txtImportFile1.Text = "" Then
    MsgBox "�Х����w�פJ�ɮסC", vbExclamation, "ĵ�i"
    txtImportFile1.SetFocus
  ElseIf txtLogFile1.Text = "" Then
    MsgBox "�Х����w���D�O���ɡC", vbExclamation, "ĵ�i"
    txtLogFile1.SetFocus
  ElseIf Not Fso.FileExists(txtImportFile1.Text) Then
    MsgBox "�n�פJ���ɮפ��s�b�C", vbExclamation, "ĵ�i"
    txtImportFile1.SetFocus
  ElseIf txtImportFile1.Text = txtLogFile1.Text Then
    MsgBox "�פJ�ɮשM���D�O���ɤ��i�H�@�ˡC", vbExclamation, "ĵ�i"
    txtImportFile1.SetFocus
  Else
  
    '�}�@�� SO033 �� SO033CP �� Client Dataset
    Set rsSO033 = New ADODB.Recordset
'    rsSO033.CursorLocation = adUseClient
'    rsSO033.Open "Select * From " & GetOwner & "So033 Where 1=0", gcnGi, adOpenKeyset, adLockOptimistic
     If Not GetRS(rsSO033, "Select * From " & GetOwner & "So033 Where 1=0", gcnGi, , adOpenKeyset, adLockOptimistic, adUseClient) Then Exit Sub
    Set rsSO033Cp = New ADODB.Recordset
'    rsSO033Cp.CursorLocation = adUseClient
'    rsSO033Cp.Open "Select * From " & GetOwner & "So033Cp Where 1=0", gcnGi, adOpenKeyset, adLockOptimistic
    If Not GetRS(rsSO033Cp, "Select * From " & GetOwner & "So033Cp Where 1=0", gcnGi, , adOpenKeyset, adLockOptimistic, adUseClient) Then Exit Sub
    '�s�ؤ@�� LOG ��
    If Fso.FileExists(txtLogFile1.Text) Then
      Fso.DeleteFile (txtLogFile1.Text)
    End If
    
    Set strmTextFile1 = Fso.OpenTextFile(txtLogFile1.Text, ForAppending, True)
    strmTextFile1.Write vbCrLf & "################################################################################"
    strmTextFile1.Write vbCrLf & "CP�q�H�O��ƶפJ�@���~�O���@�O������G" & Now & vbCrLf & vbCrLf
    
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
  Dim strCustId, strCpArea, strCpTel, strTFNBillNo, strTfnServiceId, str_M_48_55, str_M_56_63 As String
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
  
  '�Y�פJ���ѡA�h�ݾ�孫�סA�w�פJ SO033 ����ƻ� RollBack �^��
  gcnGi.BeginTrans
  
  '���� SO041.SysTime �P�_�n���D���ΫȤ�ݹq�����ɶ���@�פJ���ɶ�
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
  
  'Ū�J��r�ɡA�ä��q
  strTextData = Fso.OpenTextFile(txtImportFile1).ReadAll
  vDataArray = Split(strTextData, vbCrLf)
  
  'Progress Bar
  PB1.Visible = True
  PB1.Value = PB1.Min
  On Error Resume Next
  PB1.Max = UBound(vDataArray)
  
  '�}�l�@��@���R
  lngLineNumber = 1
  For Each vElement In vDataArray
    PB1.Value = lngLineNumber - 1
    Select Case Left(vElement, 1)
      Case "H"
        '���Y
        If blnHasMaster Or blnHasDetail Or blnHasSum Or blnHasTail Then
          Call LogError1(lngHeadErrorCount, "���Y�����b��r�ɪ��Ĥ@��", "", lngLineNumber, vElement)
        ElseIf Len(vElement) <> 7 Then
          Call LogError1(lngHeadErrorCount, "���Y���צ@�� 7 �X", "", lngLineNumber, vElement)
        ElseIf Not IsNumeric(Right(vElement, 6)) Then
          Call LogError1(lngHeadErrorCount, "���Y�� 2-7 �X������ YYYYMM �榡", "", lngLineNumber, vElement)
        Else
          strYear = Mid(vElement, 2, 4)
          strMonth = Right(vElement, 2)
          If Not IsDate(strYear & "/" & strMonth & "/01") Then
            Call LogError1(lngHeadErrorCount, "���Y�~����~", "", lngLineNumber, vElement)
          Else
            blnHasHead = True
          End If
        End If
      Case "M"
         lngBillItem = 0
        'Master ����
        If Not blnHasHead Then
          Call LogError1(lngMasterErrorCount, "��Ƥ��i�H�X�{�b���Y���e", "", lngLineNumber, vElement)
        ElseIf blnHasTail Then
          Call LogError1(lngMasterErrorCount, "��Ƥ��i�H�X�{�b�������", "", lngLineNumber, vElement)
        'ElseIf blnInMaster Then
        '  Call LogError1(lngMasterErrorCount, "�e�@�� Master ��ƥ����n�������� Detail ����", "",lngLineNumber,vElement)
        Else
          lngMasterCount = lngMasterCount + 1
          blnInMaster = True
          blnInDetail = False
          blnHasMaster = True
          
          '�̩T�w�e�ר��o�U��쪺��
          strCustId = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 2, 8), vbUnicode))
          strCustId2 = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 2, 10), vbUnicode))
          strCustId2 = Right("0000000" & strCustId2, 7)
          
          '�ϽX�n�[�J '-' �Ÿ�
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
          
          '�� strBillNo �ӧP�_���P���p�p����
          strBillNo = ""
          '*******************************************************************************************************
          '#3526 �W�[�P�_�����~��O�_�ŦX����榡 By Kin 2007/09/17
          If IsDate(Left(strShouldDate, 4) & "/" & Mid(strShouldDate, 5, 2) & "/01") Then
            blnGoOn = True
          Else
            Call LogError1(lngMasterErrorCount, "�Ƚs�����~�뤣���T�G" & strCustId, "", lngLineNumber, vElement)
            blnGoOn = False
          End If
          '*******************************************************************************************************
          '���U�i��U���U�˪��ˮ�
          Set rsTemp = gcnGi.Execute("Select CustId, ClassCode1 From " & GetOwner & "So001 Where CustId=" & strCustId)
          If rsTemp.EOF Then
            blnGoOn = False
            Call LogError1(lngMasterErrorCount, "�Ƚs���s�b�G" & strCustId, "", lngLineNumber, vElement)
          Else
            strClassCode1 = rsTemp("ClassCode1") & ""
          End If
          Set rsTemp = Nothing
          
'          If blnGoOn Then
            strString1 = Trim(strCpArea & strCpTel)
'            Set rsTemp = gcnGi.Execute("Select CustId From " & GetOwner & "So004 Where FaciSNo='" & strString1 & "'")
'            If rsTemp.EOF Then
'              blnGoOn = False
'              Call LogError1(lngMasterErrorCount, "�Ƚs�G" & strCustId & "�@�������s�b�G" & strString1, "", lngLineNumber, vElement)
'            End If
'            Set rsTemp = Nothing
'          End If
          
          If blnGoOn Then
          
            '#3303 �W�[�P�_SO004.InstDate�O�_���ȡA�p�G�L�ȭn�g�JLog�� By Kin 2007/10/16
            Set rsTemp = gcnGi.Execute("Select CustId,InstDate From " & GetOwner & "So004 Where (CustId=" & strCustId & ") And (FaciSNo='" & strString1 & "')")
            If rsTemp.EOF Then
                blnGoOn = False
                Call LogError1(lngMasterErrorCount, "�Ƚs�P�������ũΤ��s�b�G" & strCustId & "�A" & strString1, "", lngLineNumber, vElement)
            Else
                If rsTemp("InstDate") & "" = "" Then
                    blnGoOn = False
                    Call LogError1(lngMasterErrorCount, "�����G" & strString1 & " �|���w�˧��u�A�Ч��u��A�A���s�פJ", "", lngLineNumber, vElement)
                End If

            End If
            Set rsTemp = Nothing
          End If
        End If
      Case "D"
        'Detail ����
        If Not blnHasHead Then
          Call LogError1(lngDetailErrorCount, "Detail ���ؤ��i�H�X�{�b���Y���e", "", lngLineNumber, vElement)
        ElseIf blnHasTail Then
          Call LogError1(lngDetailErrorCount, "Detail ���ؤ��i�H�X�{�b�������", "", lngLineNumber, vElement)
        ElseIf (Not blnInMaster) And (Not blnInDetail) Then
          Call LogError1(lngDetailErrorCount, "Detail ���بS�������� Master ���", "", lngLineNumber, vElement)
        Else
          lngDetailCount = lngDetailCount + 1
          blnInMaster = False
          'blnInDetail = True
          blnHasDetail = True
          
          '�̩T�w�e�ר��o�U��쪺��
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
          
          '���U�i��U���U�˪��ˮ�
          blnGoOn = True
          Set rsTemp = gcnGi.Execute("Select CItemCode1,CpCItemName, CpTax From " & GetOwner & "So165 Where CpCItemCode='" & strCpCItemCode & "'")
          If rsTemp.EOF Then
            blnGoOn = False
            Call LogError1(lngDetailErrorCount, "�Ƚs�G" & strCustId & "�@APBT�b�涵�إN�X���s�b�G" & strCpCItemCode, "", lngLineNumber, vElement)
          Else
            If rsTemp("CpTax") <> strCpTax Then
              blnGoOn = False
              Call LogError1(lngDetailErrorCount, "�Ƚs�G" & strCustId & "�@�|�v���šG" & strCpTax, rsTemp("CpTax"), lngLineNumber, vElement)
            End If
'            If rsTemp("CItemCode1") & "" = "" Then
'              blnGoOn = False
'              Call LogError1(lngDetailErrorCount, "�Ƚs�G" & strCustId & "�@APBT�b�涵�ئW�٤��s�b�G" & strCpCItemName, "", lngLineNumber, vElement)
'            End If
          End If
          Set rsTemp = Nothing
          
'          If blnGoOn Then
'            Set rsTemp = gcnGi.Execute("Select CItemCode1 From " & GetOwner & "So165 Where CpCItemName='" & strCpCItemName & "'")
'            If rsTemp.EOF Then
'              blnGoOn = False
'              Call LogError1(lngDetailErrorCount, "�Ƚs�G" & strCustId & "�@APBT�b�涵�ئW�٤��s�b�G" & strCpCItemName, "", lngLineNumber, vElement)
'            End If
'            Set rsTemp = Nothing
'          End If
          
          'If blnGoOn Then
          '  Set rsTemp = gcnGi.Execute("Select CpTax From " & GetOwner & "So165 Where CpCItemCode='" & strCpCItemCode & "'")
          '  If rsTemp("CpTax") <> strCpTax Then
          '    blnGoOn = False
          '    Call LogError1(lngDetailErrorCount, "�Ƚs�G" & strCustid & "�@�|�v���šG" & strCpTax, rsTemp("CpTax"), lngLineNumber, vElement)
          '  End If
          '  Set rsTemp = Nothing
          'End If
          
          If blnGoOn Then
            If strSign = "-" Then
              strString1 = "CItemCode2"
            Else
              strString1 = "CItemCode1"
            End If
            Set rsTemp = gcnGi.Execute("Select CItemCode1 From " & GetOwner & "So165 Where " & strString1 & " Is Not Null")
            If rsTemp.EOF Then
              blnGoOn = False
              Call LogError1(lngDetailErrorCount, "�Ƚs�G" & strCustId & "�@�ɶU�Ÿ����šG" & strSign, "", lngLineNumber, vElement)
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
                         "Where (CustId ='" & strCustId & "') " & _
                         "And (FaciSNo ='" & Trim(strCpArea & strCpTel) & "') " & _
                         "And (TfnBillNo ='" & strTFNBillNo & "') " & _
                         "And (ClctYm ='" & strShouldDate & "') " & _
                         "And (CItemCode ='" & strCitemCode & "') " & _
                         "And (ShouldAmt ='" & strShouldAmt & "') " & _
                         "And (CPShouldYM ='" & strCpShouldYm & "') " & _
                         "Union All " & _
                         "Select CustId " & _
                         "From " & GetOwner & "So034 " & _
                         "Where (CustId ='" & strCustId & "') " & _
                         "And (FaciSNo ='" & Trim(strCpArea & strCpTel) & "') " & _
                         "And (TfnBillNo ='" & strTFNBillNo & "') " & _
                         "And (ClctYm ='" & strShouldDate & "') " & _
                         "And (CItemCode ='" & strCitemCode & "') " & _
                         "And (ShouldAmt ='" & strShouldAmt & "') " & _
                         "And (CPShouldYM ='" & strCpShouldYm & "') "

            Set rsTemp = gcnGi.Execute(strString1)
            If Not rsTemp.EOF Then
              blnGoOn = False
              Call LogError1(lngDetailErrorCount, "�Ƚs�G" & strCustId & "�@��Ƥw�s�b�A�i��O�q�H�b�涵�ئ����Ʃγq�H�O�θ�ƭ���", "", lngLineNumber, vElement)
            End If
            Set rsTemp = Nothing
          End If
          
'          If blnGoOn Then
'            strString1 = "Select CustId " & _
'                         "From " & GetOwner & "So033CP " & _
'                         "Where (FaciSNo ='" & Trim(strCpArea & strCpTel) & "') " & _
'                         "And (CustId ='" & strCustId & "') " & _
'                         "And (TfnBillNo ='" & strTFNBillNo & "') " & _
'                         "And (ClctYm ='" & strYear & strMonth & "') " & _
'                         "And (CItemCode ='" & strCitemCode & "') " & _
'                         "And (ShouldAmt ='" & strShouldAmt & "') " & _
'                         "And (CPShouldYM ='" & strCpShouldYm & "') "
'            Set rsTemp = gcnGi.Execute(strString1)
'            If Not rsTemp.EOF Then
'              blnGoOn = False
'              Call LogError1(lngDetailErrorCount, "�Ƚs�G" & strCustId & "�@��Ƥw�s�b�A�i��O�q�H�b�涵�ئ����Ʃγq�H�O�θ�ƭ���", "", lngLineNumber, vElement)
'            End If
'            Set rsTemp = Nothing
'          End If
          
          '�e�����ˮֳ��q�L�F�A�ǳƱN�פJ����Ƽg�J SO033Cp ��
          If blnGoOn Then
            rsSO033Cp.AddNew
            
            rsSO033Cp("CustId") = strCustId
            '���D2722 �ĤT�I ----for Taos �[�JSO034�鵲�᪺ITEM�ӧP�_ 2006/09/18
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
              '#4035 �䤣���ƴN�XLog,���n��0
              Call LogError1(lngDetailErrorCount, "�Ƚs�G" & strCustId & " �䤣����������O����", "", lngLineNumber, vElement)
'              rsSO033Cp("CitemCode") = 0
'              rsSO033Cp("CitemName") = " "
'              rsSO033Cp("CPAdjCitemCode") = 0
'              rsSO033Cp("CPGroupCode") = 0
'              rsSO033Cp("CPItem") = " "
            End If
            Set rsTemp = Nothing
            
            rsSO033Cp("OldAmt") = strShouldAmt
            rsSO033Cp("OldPeriod") = 0
            
            strString1 = Right("00" & txtShouldDay.Text, 2)
            strTempDate = Left(strShouldDate, 4) & "/" & Right(strShouldDate, 2) & "/" & strString1
            If IsDate(strTempDate) Then
              rsSO033Cp("ShouldDate") = strTempDate
            Else
              Call LogError1(lngDetailErrorCount, "�Ƚs�G" & strCustId & "�@��������榡���~�G" & strTempDate, "", lngLineNumber, vElement)
              rsSO033Cp("ShouldDate") = Date
            End If
            
            rsSO033Cp("ShouldAmt") = strShouldAmt
            rsSO033Cp("RealAmt") = 0
            rsSO033Cp("RealPeriod") = 0
            
            'Set rsTemp = gcnGi.Execute("Select UcCode From " & GetOwner & "So044")
            'If rsTemp.EOF Then
            '  If Not blnInDetail Then
            '    Call LogError1(lngDetailErrorCount, "�Ƚs�G" & strCustid & "�@������]�N�X���w�q", "", lngLineNumber, vElement)
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
            '  Call LogError1(lngDetailErrorCount, "�Ƚs�G" & strCustid & "�@��ú�O��]�N�X���s�b�G" & rsSO033Cp("UcCode"), "", lngLineNumber, vElement)
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
                         "And (CustId=" & strCustId & ") " & _
                         "And (ServiceType='P') "
            Set rsTemp = gcnGi.Execute(strString1)
            
            If Not rsTemp.EOF Then
                rsSO033Cp("FaciSeqNo") = rsTemp("FaciSeqNo")
                rsSO033Cp("CmCode") = rsTemp("CmCode")
                rsSO033Cp("CmName") = rsTemp("CmName")
                rsSO033Cp("PtCode") = rsTemp("PtCode")
                rsSO033Cp("PtName") = rsTemp("PtName")
                rsSO033Cp("AccountNo") = rsTemp("AccountNo")
                rsSO033Cp("BankCode") = rsTemp("BankCode")
                rsSO033Cp("BankName") = rsTemp("BankName")
            Else
            
              strString1 = "Select CodeNo, Description " & _
                           "From " & GetOwner & "Cd031 A," & GetOwner & "So044 B " & _
                           "Where (A.CodeNo=B.CmCode) " & _
                           "And ((A.ServiceType='P') " & _
                           "Or (A.ServiceType Is Null)) "
              Set rsTemp2 = gcnGi.Execute(strString1)
              If rsTemp2.EOF Then
                '#4035 �䤣���N�XLog,���n��0 By Kin 2008/08/14
                'If Not blnInDetail Then
                  Call LogError1(lngDetailErrorCount, "�Ƚs�G" & strCustId & "�@���O�覡�Ѽƥ��w�q", "", lngLineNumber, vElement)
                'End If
'                rsSO033Cp("CmCode") = "0"
'                rsSO033Cp("CmName") = "0"
'                rsSO033Cp("FaciSeqNo") = "0"
'                rsSO033Cp("PtCode") = "0"
'                rsSO033Cp("PtName") = "0"
'                rsSO033Cp("AccountNo") = "0"
'                rsSO033Cp("BankCode") = "0"
'                rsSO033Cp("BankName") = "0"
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
            
            '�̸˾��a�}�Φ��O�a�}�A�N�a�}��J�� SO033 ���a�}��Ƥ�
            '#4035 �H���O�a�}���ﶵ�A�n�HSO138�����O�a�}���D By Kin 2008/08/14
            lngTempLong = 0
            If optEmpNo.Value Then
              '�˾��a�}
              strString1 = "Select B.ClctEn, B.ClctName, B.AddrNo, B.StrtCode, " & _
                           "       A.MduId, A.ServCode, A.ClctAreaCode, " & _
                           "       B.AreaCode, B.NodeNo " & _
                           "From " & GetOwner & "So001 A, " & GetOwner & "So014 B " & _
                           "Where (A.InstAddrNo=B.AddrNo) " & _
                           "And (A.CustId=" & strCustId & ") "
              Set rsTemp = gcnGi.Execute(strString1)
            Else
             '���O�a�}
'              strString1 = "Select Count(*) As DataCount " & _
'                           "From " & GetOwner & "So003 " & _
'                           "Where (FaciSNo='" & Trim(strCpArea & strCpTel) & "') " & _
'                           "And (CustId=" & strCustId & ") " & _
'                           "And (ServiceType='P') " & _
'                           "And (AccountNo Is Not Null) "
                strString1 = "Select InvSeqNo From " & GetOwner & "SO003" & _
                            " Where (FaciSNo='" & Trim(strCpArea & strCpTel) & "') " & _
                            "And (CustId=" & strCustId & ") " & _
                            "And (ServiceType='P') " & _
                            "And (AccountNo Is Not Null) "
              Set rsTemp = gcnGi.Execute(strString1)
              'lngTempLong = rsTemp("DataCount")
              If Not rsTemp.EOF Then
                 lngTempLong = rsTemp("InvSeqNo")
              End If
              Set rsTemp = Nothing
              If lngTempLong = 0 Then
                strString1 = "Select B.ClctEn, B.ClctName, B.AddrNo, B.StrtCode, " & _
                             "       A.MduId, A.ServCode, A.ClctAreaCode, " & _
                             "       B.AreaCode, B.NodeNo " & _
                             "From " & GetOwner & "So001 A, " & GetOwner & "So014 B " & _
                             "Where (A.ChargeAddrNo=B.AddrNo) " & _
                             "And (A.CustId=" & strCustId & ") "
                Set rsTemp = gcnGi.Execute(strString1)
              Else
'                strString1 = "Select C.ClctEn, C.ClctName, C.AddrNo, C.StrtCode, " & _
'                             "       A.MduId, A.ServCode, A.ClctAreaCode, " & _
'                             "       C.AreaCode, C.NodeNo " & _
'                             "From " & GetOwner & "So001 A, " & GetOwner & "So002A B, " & GetOwner & "So014 C " & _
'                             "Where (B.ChargeAddrNo=C.AddrNo) " & _
'                             "And (B.CustId=" & strCustId & ") " & _
'                             "And (B.CustId=A.CustId) "
                strString1 = "Select C.ClctEn,C.ClctName,C.AddrNo,C.StrtCode," & _
                             "A.MduId,A.ServCode,A.ClctAreaCode, " & _
                             "C.AreaCode, C.NodeNo " & _
                             "From " & GetOwner & "So001 A," & GetOwner & "So138 B," & GetOwner & "So014 C " & _
                             "Where (B.ChargeAddrNo=C.AddrNo) " & _
                             "And (A.CustId=" & strCustId & ") " & _
                             "And (B.InvSeqNO=" & lngTempLong & ")"
                Set rsTemp = gcnGi.Execute(strString1)
              End If
            End If
            
            '��J�� SO033 ���a�}��Ƥ�
            '#4035 �䤣���ƴN�XLog,���n��0 By Kin 2008/08/14
            If rsTemp.EOF Then
              'If Not blnInDetail Then
                Call LogError1(lngDetailErrorCount, "�Ƚs�G" & strCustId & "�@�a�}��Ƥ��s�b�G" & strCustId & "�A" & Trim(strCpArea & strCpTel), "", lngLineNumber, vElement)
              'End If
'              rsSO033Cp("ClctEn") = "0"
'              rsSO033Cp("ClctName") = "0"
'              rsSO033Cp("AddrNo") = "0"
'              rsSO033Cp("StrtCode") = "0"
'              rsSO033Cp("MduId") = "0"
'              rsSO033Cp("ServCode") = "0"
'              rsSO033Cp("ClctAreaCode") = "0"
'              rsSO033Cp("OldClctEn") = "0"
'              rsSO033Cp("OldClctName") = "0"
'              rsSO033Cp("AreaCode") = "0"
'              rsSO033Cp("NodeNo") = "0"
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
        '�p�p����
        If Not blnHasHead Then
          Call LogError1(lngSumErrorCount, "�p�p���ؤ��i�H�X�{�b���Y���e", "", lngLineNumber, vElement)
        ElseIf blnHasTail Then
          Call LogError1(lngSumErrorCount, "�p�p���ؤ��i�H�X�{�b�������", "", lngLineNumber, vElement)
        ElseIf Not blnInDetail Then
          Call LogError1(lngSumErrorCount, "�p�p���بS�������� Detail ���", "", lngLineNumber, vElement)
        Else
          lngSumCount = lngSumCount + 1
          blnInMaster = False
          blnHasSum = True
          
          '�̩T�w�e�ר��o�U��쪺��
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
          
          '���U�i��U���U�˪��ˮ�
          Set rsTemp = gcnGi.Execute("Select CpAdjCItemCode,CItemCode1, CpTax,CpCItemName From " & GetOwner & "So165 Where CpCItemCode='" & strCpCItemCode & "' And (CpProperty In (1,2))")
          If rsTemp.EOF Then
            Call LogError1(lngSumErrorCount, "�Ƚs�G" & strCustId & "�@�q�H�b�涵�إN�X���s�b�G" & strCpCItemCode, "", lngLineNumber, vElement)
            
          Else
            If rsTemp("CpAdjCItemCode") & "" <> strCpAdjCItemCode Then
              Call LogError1(lngSumErrorCount, "�Ƚs�G" & strCustId & "�@�էﶵ�ؤ��šG" & strCpAdjCItemCode, "", lngLineNumber, vElement)
            End If
            If rsTemp("CpTax") <> strCpTax Then
              blnGoOn = False
              Call LogError1(lngSumErrorCount, "�Ƚs�G" & strCustId & "�@�|�v���šG" & strCpTax, rsTemp("CpTax"), lngLineNumber, vElement)
            End If
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
              Call LogError1(lngSumErrorCount, "�Ƚs�G" & strCustId & "�@�ɶU�Ÿ����šG" & strSign, "", lngLineNumber, vElement)
            End If
            Set rsTemp = Nothing
          End If
          
          If blnGoOn Then
          '���D2722 �ĥ|�� ----For Taos �t�����B�����t�������O����
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
                         "And (CustId ='" & strCustId & "') " & _
                         "And (TfnBillNo ='" & strTFNBillNo & "') " & _
                         "And (ClctYm =" & strYear & strMonth & ") " & _
                         "And (CItemCode ='" & strCitemCode & "') " & _
                         "And (ShouldAmt ='" & strShouldAmt & "') " & _
                         "And (CPShouldYM ='" & strCpShouldYm & "') " & _
                         "Union All " & _
                         "Select CustId " & _
                         "From " & GetOwner & "So034 " & _
                         "Where (FaciSNo ='" & Trim(strCpArea & strCpTel) & "') " & _
                         "And (CustId ='" & strCustId & "') " & _
                         "And (TfnBillNo ='" & strTFNBillNo & "') " & _
                         "And (ClctYm =" & strYear & strMonth & ") " & _
                         "And (CItemCode ='" & strCitemCode & "') " & _
                         "And (ShouldAmt ='" & strShouldAmt & "') " & _
                         "And (CPShouldYM ='" & strCpShouldYm & "') "
'            Set rsTemp = gcnGi.Execute(strString1)
            If Not GetRS(rsTemp, strString1, gcnGi, , adOpenForwardOnly, adLockOptimistic, adUseClient, False) Then Exit Sub
            If Not rsTemp.EOF Then
              blnGoOn = False
              Call LogError1(lngSumErrorCount, "�Ƚs�G" & strCustId & "�@��Ƥw�s�b�A�i��O�q�H�b�涵�ئ����Ʃγq�H�O�θ�ƭ���", "", lngLineNumber, vElement)
            End If
            Set rsTemp = Nothing
          End If
          
          '�e�����ˮֳ��q�L�F�A�ǳƱN�פJ����Ƽg�J SO033 ��
          If blnGoOn Then
            rsSO033.AddNew
            
            rsSO033("CustId") = strCustId
            
            rsSO033("BillNo") = strBillNo
            rsSO033("Item") = lngBillItem
            
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
                '#4035 �䤣���ƴN�XLog,���n��0
                Call LogError1(lngSumErrorCount, "�Ƚs�G" & strCustId & " �䤣����������O����", "", lngLineNumber, vElement)
'                rsSO033("CitemCode") = 0
'                rsSO033("CitemName") = " "
'                rsSO033("CPAdjCitemCode") = 0
'                rsSO033("CPGroupCode") = 0
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
              Call LogError1(lngSumErrorCount, "�Ƚs�G" & strCustId & "�@��������榡���~�G" & strTempDate, "", lngLineNumber, vElement)
              rsSO033("ShouldDate") = Date
            End If
            
            rsSO033("ShouldAmt") = Round(strShouldAmt)
            rsSO033("RealAmt") = 0
            rsSO033("RealPeriod") = 0
            
            'Set rsTemp = gcnGi.Execute("Select UcCode From " & GetOwner & "So044")
            'If rsTemp.EOF Then
            '  If Not blnInDetail Then
            '    Call LogError1(lngSumErrorCount, "�Ƚs�G" & strCustid & "�@������]�N�X���w�q", "", lngLineNumber, vElement)
            '  End If
            '  rsSO033("UcCode") = "0"
            'Else
            '  rsSO033("UcCode") = rsTemp("UcCode")
            'End If
            'Set rsTemp = Nothing
            
            rsSO033("UcCode") = strUCCode
            rsSO033("UcName") = strUCName
            
            'Set rsTemp = gcnGi.Execute("Select Description From " & GetOwner & "Cd013 Where CodeNo=" & rsSO033("UcCode"))
            'If rsTemp.EOF Then
            '  Call LogError1(lngSumErrorCount, "�Ƚs�G" & strCustid & "�@��ú�O��]�N�X���s�b�G" & rsSO033("UcCode"), "", lngLineNumber, vElement)
            '  rsSO033("UcName") = "0"
            'Else
            '  rsSO033("UcName") = rsTemp("Description")
            'End If
            'Set rsTemp = Nothing
            
            rsSO033("Note") = strNote
            rsSO033("CreateTime") = UserDateTime
            rsSO033("CreateEn") = garyGi(1)
            rsSO033("CompCode") = garyGi(9)
            rsSO033("CancelFlag") = 0
            rsSO033("PrtCount") = 0
            
            'Set rsTemp = gcnGi.Execute("Select ClassCode1 From " & GetOwner & "So001 Where CustId=" & strCustid)
            'If rsTemp.EOF Then
            '  rsSO033("ClassCode") = "0"
            'Else
            '  rsSO033("ClassCode") = rsTemp("ClassCode1")
            'End If
            'Set rsTemp = Nothing
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
            
'            strString1 = "Select Count(*) As DataCount " & _
'                         "From " & GetOwner & "So003 " & _
'                         "Where (FaciSNo='" & Trim(strCpArea & strCpTel) & "') " & _
'                         "And (CustId=" & strCustId & ") " & _
'                         "And (ServiceType='P') "
'            Set rsTemp = gcnGi.Execute(strString1)
'            lngTempLong = rsTemp("DataCount")
'            Set rsTemp = Nothing
              '**************************************************************************
              '#3871 SO033�]�n�@�֧�sInvSeqNO By Kin 2008/04/15
              strString1 = "Select FaciSeqNo, CmCode, CmName, PtCode, " & _
                           "       PtName, AccountNo, BankCode, BankName,InvSeqNo " & _
                           "From " & GetOwner & "So003 " & _
                           "Where (FaciSNo='" & Trim(strCpArea & strCpTel) & "') " & _
                           "And (CustId=" & strCustId & ") " & _
                           "And (ServiceType='P') "
              Set rsTemp = gcnGi.Execute(strString1)
              '***************************************************************************
            If Not rsTemp.EOF Then
                rsSO033("FaciSeqNo") = rsTemp("FaciSeqNo")
                rsSO033("CmCode") = rsTemp("CmCode")
                rsSO033("CmName") = rsTemp("CmName")
                rsSO033("PtCode") = rsTemp("PtCode")
                rsSO033("PtName") = rsTemp("PtName")
                rsSO033("AccountNo") = rsTemp("AccountNo")
                rsSO033("BankCode") = rsTemp("BankCode")
                rsSO033("BankName") = rsTemp("BankName")
                rsSO033("InvSeqNo") = IIf(IsNull(rsTemp("InvSeqNo")), Null, rsTemp("InvSeqNo"))
'              End If
'              Set rsTemp = Nothing
            Else
              strString1 = "Select CodeNo, Description " & _
                           "From " & GetOwner & "Cd031 A," & GetOwner & "So044 B " & _
                           "Where (A.CodeNo=B.CmCode) " & _
                           "And ((A.ServiceType='P') " & _
                           "Or (A.ServiceType Is Null)) "
              Set rsTemp2 = gcnGi.Execute(strString1)
              If rsTemp2.EOF Then
                '#4035 �u�n�����~�N�XLog,���n��0 By Kin 2008/08/14
                'If Not blnInDetail Then
                  Call LogError1(lngSumErrorCount, "�Ƚs�G" & strCustId & "�@���O�覡�Ѽƥ��w�q", "", lngLineNumber, vElement)
                'End If
'                rsSO033("CmCode") = "0"
'                rsSO033("CmName") = "0"
'                rsSO033("FaciSeqNo") = "0"
'                rsSO033("PtCode") = "0"
'                rsSO033("PtName") = "0"
'                rsSO033("AccountNo") = "0"
'                rsSO033("BankCode") = "0"
'                rsSO033("BankName") = "0"
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
            lngTempLong = 0
            '�̸˾��a�}�Φ��O�a�}�A�N�a�}��J�� SO033 ���a�}��Ƥ�
            '#4035 �p�G�H���O�a�}���ﶵ�A�n�HSO138�����O�a�}���D By Kin 2008/08/14
            If optEmpNo.Value Then
              '�˾��a�}
              strString1 = "Select B.ClctEn, B.ClctName, B.AddrNo, B.StrtCode, " & _
                           "       A.MduId, A.ServCode, A.ClctAreaCode, " & _
                           "       B.AreaCode, B.NodeNo " & _
                           "From " & GetOwner & "So001 A, " & GetOwner & "So014 B " & _
                           "Where (A.InstAddrNo=B.AddrNo) " & _
                           "And (A.CustId=" & strCustId & ") "
              Set rsTemp = gcnGi.Execute(strString1)
            Else
             '���O�a�}
'              strString1 = "Select Count(*) As DataCount " & _
'                           "From " & GetOwner & "So003 " & _
'                           "Where (FaciSNo='" & Trim(strCpArea & strCpTel) & "') " & _
'                           "And (CustId=" & strCustId & ") " & _
'                           "And (ServiceType='P') " & _
'                           "And (AccountNo Is Not Null) "
                strString1 = "Select InvSeqNo From " & GetOwner & "SO003" & _
                            " Where (FaciSNo='" & Trim(strCpArea & strCpTel) & "') " & _
                            "And (CustId=" & strCustId & ") " & _
                            "And (ServiceType='P') " & _
                            "And (AccountNo Is Not Null) "
              Set rsTemp = gcnGi.Execute(strString1)
              'lngTempLong = rsTemp("DataCount")
              If Not rsTemp.EOF Then
                 lngTempLong = rsTemp("InvSeqNo")
              End If
              Set rsTemp = Nothing
              If lngTempLong = 0 Then
                strString1 = "Select B.ClctEn, B.ClctName, B.AddrNo, B.StrtCode, " & _
                             "       A.MduId, A.ServCode, A.ClctAreaCode, " & _
                             "       B.AreaCode, B.NodeNo " & _
                             "From " & GetOwner & "So001 A, " & GetOwner & "So014 B " & _
                             "Where (A.ChargeAddrNo=B.AddrNo) " & _
                             "And (A.CustId=" & strCustId & ") "
                Set rsTemp = gcnGi.Execute(strString1)
              Else
'                strString1 = "Select C.ClctEn, C.ClctName, C.AddrNo, C.StrtCode, " & _
'                             "       A.MduId, A.ServCode, A.ClctAreaCode, " & _
'                             "       C.AreaCode, C.NodeNo " & _
'                             "From " & GetOwner & "So001 A, " & GetOwner & "So002A B, " & GetOwner & "So014 C " & _
'                             "Where (B.ChargeAddrNo=C.AddrNo) " & _
'                             "And (B.CustId=" & strCustId & ") " & _
'                             "And (B.CustId=A.CustId) "
                strString1 = "Select C.ClctEn,C.ClctName,C.AddrNo,C.StrtCode," & _
                             "A.MduId,A.ServCode,A.ClctAreaCode, " & _
                             "C.AreaCode, C.NodeNo " & _
                             "From " & GetOwner & "So001 A," & GetOwner & "So138 B," & GetOwner & "So014 C " & _
                             "Where (B.ChargeAddrNo=C.AddrNo) " & _
                             "And (A.CustId=" & strCustId & ") " & _
                             "And (B.InvSeqNO=" & lngTempLong & ")"

                Set rsTemp = gcnGi.Execute(strString1)
              End If
            End If
            
            '��J�� SO033 ���a�}��Ƥ�
            '�䤣���ƴN�XLog�A���n��0 By Kin 2008/08/14
            If rsTemp.EOF Then
              'If Not blnInDetail Then
                Call LogError1(lngSumErrorCount, "�Ƚs�G" & strCustId & "�@�a�}��Ƥ��s�b�G" & strCustId & "�A" & Trim(strCpArea & strCpTel), "", lngLineNumber, vElement)
              'End If
'              rsSO033("ClctEn") = "0"
'              rsSO033("ClctName") = "0"
'              rsSO033("AddrNo") = "0"
'              rsSO033("StrtCode") = "0"
'              rsSO033("MduId") = "0"
'              rsSO033("ServCode") = "0"
'              rsSO033("ClctAreaCode") = "0"
'              rsSO033("OldClctEn") = "0"
'              rsSO033("OldClctName") = "0"
'              rsSO033("AreaCode") = "0"
'              rsSO033("NodeNo") = "0"
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
          '  Call LogError1(lngSumErrorCount, "�p�p���B�PDetail�`�X���P", vElement,lngLineNumber,vElement)
          'End If
          
          dblShouldAmt2 = 0
          dblShouldAmtNoTax2 = 0
          dblTax2 = 0
          
          '�� strBillNo �ӧP�_���P���p�p����
          strBillNo = ""
          
        End If
      Case "T"
        '�ɧ�
        If (Not blnHasMaster) Or (Not blnHasDetail) Or (Not blnHasHead) Then
          Call LogError1(lngTailErrorCount, "��������b��r�ɪ��̫�@��", "", lngLineNumber, vElement)
        Else
          blnHasTail = True
          
          '�̩T�w�e�ר��o�U��쪺��
          strMasterShouldCount = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 2, 6), vbUnicode))
          strDetailShouldCount = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 8, 6), vbUnicode))
          strSumShouldCount = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 14, 6), vbUnicode))
          strTotalDataCount = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 20, 6), vbUnicode))
          strMoneyTotal = Trim(StrConv(MidB(StrConv(vElement, vbFromUnicode), 26, 8), vbUnicode))
          
          '����ɧ��O�������Ƥι�ڶפJ������
          blnGoOn = True
          If (lngMasterCount + lngDetailCount + lngSumCount) <> Val(strTotalDataCount) Then
            blnGoOn = False
            Call LogError1(lngTailErrorCount, "��Ƶ���(" & (lngMasterCount + lngDetailCount + lngSumCount) & ")�P����O��(" & strTotalDataCount & ")������", "", lngLineNumber, vElement)
          End If
          
          '����ɧ��O�������B�ι�ڶפJ�����B
          'If blnGoOn Then
            If Val(dblShouldAmtTotal) <> Val(strMoneyTotal) Then
              blnGoOn = False
              Call LogError1(lngTailErrorCount, "��ƪ��B(" & dblShouldAmtTotal & ")�P����O��(" & CDbl(strMoneyTotal) & ")������", "", lngLineNumber, vElement)
            End If
          'End If
        End If
      Case Else
        If vElement <> "" Then
          Call LogError1(lngDetailErrorCount, "��ƵL�ġA���ˬd��1�X�O�_�Хܿ��~", "", lngLineNumber, vElement)
        End If
    End Select
    lngLineNumber = lngLineNumber + 1
    DoEvents
  Next
  If Not blnHasTail Then
    Call LogError1(lngTailErrorCount, "�S��������", "", lngLineNumber, vElement)
    lngLineNumber = lngLineNumber + 1
  End If
  PB1.Visible = False
  
  If (lngMasterErrorCount > 0) Or (lngDetailErrorCount > 0) Or (lngSumErrorCount > 0) Or (lngHeadErrorCount > 0) Or (lngTailErrorCount > 0) Then
    '�����~�A�w�פJ����ƶ������פJ
    gcnGi.RollbackTrans
    strTextData = "�w�p�פJ���ơG" & lngDetailCount & vbCrLf & _
                  "���Y���~���ơG" & lngHeadErrorCount & vbCrLf & _
                  "������~���ơG" & lngTailErrorCount & vbCrLf & _
                  "Master���~���ơG" & lngMasterErrorCount & vbCrLf & _
                  "Detail���~���ơG" & lngDetailErrorCount & vbCrLf & _
                  "�p�p���~���ơG" & lngSumErrorCount & vbCrLf & vbCrLf & _
                  "�פJ��Ƥ������~���ơA�]���N���s�ɡC" & vbCrLf & _
                  "�аѦҰ��D�O���ɡC" & vbCrLf
    MsgBox strTextData, vbExclamation, "ĵ�i"
    Shell "Notepad " & txtLogFile1.Text, vbNormalFocus
  Else
    gcnGi.CommitTrans
    strTextData = "�פJ�����A�Ҧ���ƬҽT�{�L�~ !!" & vbCrLf & vbCrLf & _
                  "�פJ���ơG" & lngMasterCount + lngDetailCount + lngSumCount & vbCrLf & _
                  "�b�浧�ơG" & lngMasterCount & vbCrLf & _
                  "�q�H�O���`�p�G" & dblShouldAmtTotal & vbCrLf & _
                  "�q�H�O�����|���B�`�p�G" & dblShouldAmtWithTax & vbCrLf & _
                  "�q�H�O�ΧK�|���B�`�p�G" & dblShouldAmtWithoutTax & vbCrLf & _
                  "�q�H�O�����|�s�|���B�`�p�G" & dblShouldAmtZeroTax & vbCrLf
    MsgBox strTextData, vbExclamation, "�T��"
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
  'Call ScrToRcdtxt
  '#3195 �N���󪺭ȼg�JIni�� By Kin 2007/07/13
  Call WriteSO19J0Ini(iniPages, strIniFile, Me)
  '�򥻸�Ʈֹ�פJ
  If txtImportFile2.Text = "" Then
    MsgBox "�Х����w�פJ�ɮסC", vbExclamation, "ĵ�i"
    txtImportFile2.SetFocus
  ElseIf txtLogFile2.Text = "" Then
    MsgBox "�Х����w���D�O���ɡC", vbExclamation, "ĵ�i"
    txtLogFile2.SetFocus
  ElseIf Not Fso.FileExists(txtImportFile2.Text) Then
    MsgBox "�n�פJ���ɮפ��s�b�C", vbExclamation, "ĵ�i"
    txtImportFile2.SetFocus
  ElseIf txtImportFile2.Text = txtLogFile2.Text Then
    MsgBox "�פJ�ɮשM���D�O���ɤ��i�H�@�ˡC", vbExclamation, "ĵ�i"
    txtImportFile2.SetFocus
  Else
  
    '�����M�� TMP012
'    gcnGi.Execute ("Delete From " & GetOwner & "Tmp012")
    '�}�@�� TMP012 �� Client Dataset
'    Set rsTmp012 = New ADODB.Recordset
'    rsTmp012.CursorLocation = adUseClient
'    rsTmp012.Open "Select * From " & GetOwner & "Tmp012 Where 1=0", gcnGi, adOpenKeyset, adLockOptimistic

    '������������ SO19J0B �ȤӮ榡�H Asscess TMP012 �Ȧs���פJ���@������������������
    '������������ SO19J0B �ȤӮ榡�H Asscess TMP012 �Ȧs���פJ���@������������������
    '�إ߼Ȧs��
    CreateTable
    '�s�ؤ@�� LOG ��
    If Fso.FileExists(txtLogFile2.Text) Then
      Fso.DeleteFile (txtLogFile2.Text)
    End If
    
    Set strmTextFile1 = Fso.OpenTextFile(txtLogFile2.Text, ForAppending, True)
    strmTextFile1.Write vbCrLf & "################################################################################"
    strmTextFile1.Write vbCrLf & "�򥻸�Ʈֹ�פJ�@���~�O���@�O������G" & Now & vbCrLf & vbCrLf
    strmTextFile1.Write vbCrLf & "�渹      MIS�Ƚs   APBT����             �T��                                              APBT                 MIS               �ӷ��渹            �w�ˮɶ�            ��ɶ�"
    strmTextFile1.Write vbCrLf & "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf
    
    Call ImportData2
    
    Set strmTextFile1 = Nothing
    Set rsTmp012 = Nothing
  End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdInport1_Click")
End Sub
Private Sub CreateTable()
  On Error Resume Next
    cnn.Execute "Drop Table Tmp012"
  On Error GoTo ChkErr
    cnn.Execute "Create Table Tmp012 (CUSTID text(10),TFNSERVICEID Text(10),CPAREACODE text(10),CPNUMBER Long," & _
                "TFNSTATUS text(20),ID text(20),CUSTNAME text(100),RECNO text(20),INSTDATE text(20),PRDATE text(20)," & _
                "Plan text(30),EBTContNo text(30))"
                '�Ƚs,�A�ȥN�X,�o�ܤ�ϰ츹�X,�o�ܤ�q�ܸ��X,
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "CreateTable")
End Sub

Private Sub ScrToRcdtxt()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
    '���D2727 ----Taos��  1.�ݧ�SO19J0A���e������ѼƧ@�x�s�A�ӫD�b�{���g�T�w �N�Ҧ������ɦW�s��log
        Set LogFile = Fso.CreateTextFile(strErrPath & "\SO19JOB.log", True)
        With LogFile
                .WriteLine (txtbillYM.Text)            '�q�H�O�X�b�~��
                .WriteLine (txtAreaCode.Text)          '�o�ܤ�ϰ츹�X
                .WriteLine (txtNumCount.Text)          '�o�ܤ�q�ܸ��X��
                .WriteLine (txtfilter.Text)            '�ư�����蠟�r��
                .WriteLine (txtExportFile1.Text)       'CP�q�H�O�P�b��ƶץX �����ɦW
                .WriteLine (txtExportFile2.Text)       'CP�믲�O��ƶץX �����ɦW
                .WriteLine (txtExportFile3.Text)       'CP�h����ƶץX �����ɦW
                .WriteLine (txtExportFile4.Text)       'CMCP�j���ƶץX �����ɦW
                .WriteLine (txtExportFile5.Text)       'CP�q�H�h�O��ƶץX �����ɦW
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcdtxt")
End Sub
Private Sub RcdToScrtxt()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
    If Dir(strErrPath & "\SO19JOB.log") = "" Then Exit Sub
            On Error Resume Next
            Set LogFile = Fso.OpenTextFile(strErrPath & "\SO19JOB.log", ForReading)
            With LogFile
                    If Not .AtEndOfStream Then txtbillYM.Text = .ReadLine & ""
                       '�q�H�O�X�b�~��
                    If Not .AtEndOfStream Then txtAreaCode.Text = .ReadLine & ""
                       '�o�ܤ�ϰ츹�X
                    If Not .AtEndOfStream Then txtNumCount.Text = .ReadLine & ""
                       '�o�ܤ�q�ܸ��X��
                    If Not .AtEndOfStream Then txtfilter.Text = .ReadLine & ""
                       '�ư�����蠟�r��
                    If Not .AtEndOfStream Then txtExportFile1.Text = .ReadLine & ""
                       'CP�q�H�O�P�b��ƶץX �����ɦW
                    If Not .AtEndOfStream Then txtExportFile2.Text = .ReadLine & ""
                       'CP�믲�O��ƶץX �����ɦW
                    If Not .AtEndOfStream Then txtExportFile3.Text = .ReadLine & ""
                       'CP�h����ƶץX �����ɦW
                    If Not .AtEndOfStream Then txtExportFile4.Text = .ReadLine & ""
                       'CMCP�j���ƶץX �����ɦW
                    If Not .AtEndOfStream Then txtExportFile5.Text = .ReadLine & ""
                       'CP�q�H�h�O��ƶץX �����ɦW
                       
            End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScrtxt")
End Sub

Private Sub cmdInport3_Click()
On Error GoTo ChkErr
  Dim StrTableName As String
  'Call ScrToRcdtxt
'#3195 �N���󪺭ȼg�JIni�� By Kin 2007/07/13
  Call WriteSO19J0Ini(iniPages, strIniFile, Me)
  'CP�q�ܩ��Ӹ�ƶפJ
  If optDetailDaily.Value And (txtImportFile3.Text = "") Then
    MsgBox "�Х����w�C���ˮֶפJ�ɦW�C", vbExclamation, "ĵ�i"
    txtImportFile3.SetFocus
  ElseIf optDetailMonthly.Value And (txtImportFile4.Text = "") Then
    MsgBox "�Х����w�C���ˮֶפJ�ɦW�C", vbExclamation, "ĵ�i"
    txtImportFile4.SetFocus
  ElseIf txtImportFile5.Text = "" Then
    MsgBox "�Х����w�פJ�ɮסC", vbExclamation, "ĵ�i"
    txtImportFile5.SetFocus
  ElseIf txtLogFile3.Text = "" Then
    MsgBox "�Х����w���D�O���ɡC", vbExclamation, "ĵ�i"
    txtLogFile1.SetFocus
  ElseIf Not Fso.FileExists(txtImportFile5.Text) Then
    MsgBox "�n�פJ���ɮפ��s�b�C", vbExclamation, "ĵ�i"
    txtImportFile5.SetFocus
  ElseIf txtImportFile5.Text = txtLogFile3.Text Then
    MsgBox "�פJ�ɮשM���D�O���ɤ��i�H�@�ˡC", vbExclamation, "ĵ�i"
    txtImportFile5.SetFocus
  Else
  
    '�q�ܩ��Ӻ����ASO168=�C���ơASO169=�C����
    StrTableName = IIf(optDetailDaily.Value, "So168", "So169")
    
    '�}�@�� SO168 �� Client Dataset
'    Set rsSo168 = New ADODB.Recordset
'
'    rsSo168.CursorLocation = adUseClient
      Dim rsTmp As New ADODB.Recordset

    '�s�ؤ@�� LOG ��
    If Fso.FileExists(txtLogFile3.Text) Then
      Fso.DeleteFile (txtLogFile3.Text)
    End If
    
    Set strmTextFile1 = Fso.OpenTextFile(txtLogFile3.Text, ForAppending, True)
    strmTextFile1.Write vbCrLf & "################################################################################"
    strmTextFile1.Write vbCrLf & "CP�q�ܩ��Ӹ�ƶפJ�@���~�O���@�O������G" & Now & vbCrLf & vbCrLf
    
    Call ImportData3
    Set strmTextFile1 = Nothing
  End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdInport3_Click")
End Sub

Private Sub cmdPrint_Click()
On Error GoTo ChkErr
  Dim strsql As String
  
  '�C�L
  strsql = "Select CItemCode1, CItemName1, TaxCode1, TaxName1, CItemCode2, CItemName2, TaxCode2, " & _
           "       TaxName2, CpCItemCode, CpCItemName, CpAdjCItemCode, CpGroupCode, CpGroupName, " & _
           "       CpItem, CpTax, CpProperty " & _
           "From " & GetOwner & "So165"
  Call ReadyGoPrint
  Call PrintRpt2(GetPrinterName(5), RptName("SO19J0"), , "CP �b�ȸ�Ʃ���@�~�]for �ȤӺ� APBT�^[SO19J0B]", strsql, "����", "", True, , False, , GiPaperLandscape)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub
Private Sub cmdSave_Click()
On Error GoTo ChkErr
  '�s��
  If IsDataOk() Then
    Call ScrToRcd
    
    '�Y�O�s�W�Ҧ��D�h�_��s�W�A�_�h�^�s���Ҧ�
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
  ExportPath.Text = FolderDialog("�п�ܭn�ץX�ɮת����|")
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
        CmdCancel.Value = True
    End Select
  End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
On Error GoTo ChkErr
  Dim blnGoOn As Boolean
  strErrPath = ReadGICMIS1("ErrLogPath")
  Set rsSo165 = New ADODB.Recordset
  Call LoadCom
  Call LoadGrid
  Call OpenData
  Call RcdToScr
  Call ChangeMode(giEditModeView)
  Set cnn = GetTmpMdbCn

  '��J�e�����󪺹w�]��
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
'********************************************************************
'#3195 �쥻�q��r�ɪ����e�g�줸��{�b�אּIni�� By Kin 2007/07/13
'  Call RcdToScrtxt
'  If txtExportFile1 = "" Then txtExportFile1 = "NVW_PIH_U"
'  If txtExportFile2 = "" Then txtExportFile2 = "NVW_PIH_R"
'  If txtExportFile3 = "" Then txtExportFile3 = "NVW_REMOVE_R"
'  If txtExportFile4 = "" Then txtExportFile4 = "NVW_REMOVE_"
'  If txtExportFile5 = "" Then txtExportFile5 = "NVW_REPAY_U"
'*******************************************************************

'*********************************************************************
  '#3195 �N���󪺹w�]���IniŪ�J By Kin 2007/07/13
  Call ReadIniPages '�N����W�ٻP�w�]�ȶ�J���󶰦X��
  Call ReadSO19J0Ini(iniPages, strIniFile, Me)  '�NIni��Ū�X
'*********************************************************************
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ChkErr
  If lngEditMode <> giEditModeView Then
    MsgBox "�Х��T�w�s�ɩΨ����s�ɤ���A�A�������C", vbExclamation, "ĵ�i"
    Cancel = True
  End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_QueryUnload")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ChkErr
  '�����{��
  Set Fso = Nothing
  CloseRecordset rsSo165
  Call ReleaseCOM(Me)
  Set iniPages = Nothing
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
  'SetgiMulti gimPrCode1, "CodeNo", "Description", "CD007", "������O�N�X", "������O�W��", " Where (StopFlag=0) And (ServiceType='P') And (RefNo=2) ", True
  '2006/01/17 �Q�װѦҸ��אּ 5
  '2006/01/26 �S���n�令 2
  SetgiMulti gimPrCode1, "CodeNo", "Description", "CD007", "������O�N�X", "������O�W��", " Where (StopFlag=0) And (ServiceType='P') And (RefNo=2) ", True
  SetgiMulti gimPrCode2, "CodeNo", "Description", "CD007", "������O�N�X", "������O�W��", " Where (StopFlag=0) And (ServiceType='P') And (RefNo=2) ", True
  SetgiMulti gimReasonCode, "CodeNo", "Description", "CD014", "�������]�N�X", "�������]�N�X", " Where (StopFlag=0) And ((ServiceType='P') Or (ServiceType Is Null)) ", True
  SetgiMulti gimUCCode, "CodeNo", "Description", "CD014", "�������]�N�X", "�������]�N�X", " Where (StopFlag=0) And ((ServiceType='P') Or (ServiceType Is Null)) ", True
 Exit Sub
ChkErr:
  ErrSub Me.Name, "LoadCom"
End Sub

Private Sub LoadGrid()
On Error GoTo ChkErr
  Dim mFlds As New prjGiGridR.GiGridFlds
  
  '�e���� Grid �������]�w
  mFlds.Add "CItemCode1", , , , , "���O���إN�X(+)"
  mFlds.Add "CItemName1", , , , , "���O���ئW��(+)"
  mFlds.Add "TaxName1", , , , , "�|�v(+)"
  mFlds.Add "CItemCode2", , , , , "���O���إN�X(-)"
  mFlds.Add "CItemName2", , , , , "���O���ئW��(-)"
  mFlds.Add "TaxName2", , , , , "�|�v(-)"
  mFlds.Add "CpCItemCode", , , , , "�q�H�b�涵�إN�X"
  mFlds.Add "CpCItemName", , , , , "�q�H�b�涵�ئW��"
  mFlds.Add "CpAdjCItemCode", , , , , "�q�H�b��էﶵ�إN�X"
  mFlds.Add "CpGroupCode", , , , , "�q�H�b��s�եN�X"
  mFlds.Add "CpGroupName", , , , , "�q�H�b��s�զW��"
  mFlds.Add "CpItem", , , , , "�q�H�C�L����"
  mFlds.Add "TfnHead", , , , , "���O���ثe�ɦr��"
  ggrData.AllFields = mFlds
  ggrData.SetHead
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "LoadGrid")
End Sub

Private Sub ChangeMode(lngMode As giEditModeEnu)
On Error GoTo ChkErr
  Dim blnViewMode As Boolean
  
  '�N�s��Ҧ��x�s�_��
  lngEditMode = lngMode
  
  '�����Ҧ�
  Select Case lngMode
    Case giEditModeView
      blnViewMode = True
      lblEditMode.Caption = "���"
      CmdCancel.Caption = "����(&X)"
    Case giEditModeEdit
      blnViewMode = False
      lblEditMode.Caption = "�ק�"
      CmdCancel.Caption = "����(&X)"
    Case giEditModeInsert
      blnViewMode = False
      lblEditMode.Caption = "�s�W"
      CmdCancel.Caption = "����(&X)"
  End Select
  
  '�ھڽs��Ҧ��ӳ]�w�e�������s���A
  ggrData.Enabled = blnViewMode
  fraData.Enabled = Not blnViewMode
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
    
  'Ū�X��ơA��J�e������
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
  Exit Sub
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
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "SetExporFileName")
End Sub

Private Sub GiExportDate3_Change()
On Error GoTo ChkErr
  Call SetExporFileName(txtExportPath2, txtExportFile2, txtLogFile5, lblExportFile2, GiExportDate3)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "GiExportDate3_Change")
End Sub

Private Sub GiExportDate4_Change()
On Error GoTo ChkErr
  lblExportFile3.Caption = txtExportFile3 & GiExportDate4.GetValue & ".TXT"
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "GiExportDate4_Change")
End Sub

Private Sub GiExportDate5_Change()
On Error GoTo ChkErr
  lblExportFile4.Caption = txtExportFile4 & GiExportDate5.GetValue & ".TXT"
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "GiExportDate5_Change")
End Sub

Private Sub GiExportDate6_Change()
On Error GoTo ChkErr
  Call SetExporFileName(txtExportPath5, txtExportFile5, txtLogFile6, lblExportFile5, GiExportDate6)
  Exit Sub
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
    MsgBox "�Х��T�w�s�ɩΨ����s�ɤ���A�A���������C", vbExclamation, "ĵ�i"
  End If
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "SSTab1_Click")
End Sub

Private Sub NewRcd()
On Error GoTo ChkErr
  '�s�W��ƶ����N Client Dataset �����M��
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
    strErrMsg = "�|�v (�ɶU : + ) ������"
    gilTax1.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And (gilCItem2.GetCodeNo <> "") And (gilTax2.GetCodeNo = "") Then
    strErrMsg = "�|�v (�ɶU : - ) ������"
    gilTax2.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And (gilCItem1.GetCodeNo = "") And (gilCItem2.GetCodeNo = "") Then
    strErrMsg = "���O���� (�ɶU : + ) �P ���O���� (�ɶU : - ) �ܤ֨䤤�@�Ӷ�����"
    gilCItem1.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And (gilCItem1.GetCodeNo <> "") And (gilCItem2.GetCodeNo <> "") And (gilTax1.GetCodeNo <> gilTax2.GetCodeNo) Then
    strErrMsg = "���O����(�ɶU:+) �P ���O����(�ɶU:- ) �Ҧ��Ȯ�,��|�v���ۦP"
    gilCItem1.SetFocus
    blnGoOn = False
  End If
  
  If blnGoOn Then
    If gilCItem1.GetCodeNo <> "" Then
      Set rsCD019 = gcnGi.Execute("Select PeriodFlag From " & GetOwner & "Cd019 Where CodeNo=" & gilCItem1.GetCodeNo)
      If Not rsCD019.EOF Then
        If ((cboCpProperty.ListIndex = 0) Or (cboCpProperty.ListIndex = 3)) And (rsCD019("PeriodFlag") <> "0") Then
          strErrMsg = "��[�q�H�����ݩ�]=1��4��,���O���� (�ɶU : + )�����D�g���ʶ���"
          gilCItem1.SetFocus
          blnGoOn = False
        ElseIf ((cboCpProperty.ListIndex = 1) Or (cboCpProperty.ListIndex = 2)) And (rsCD019("PeriodFlag") <> "1") Then
          strErrMsg = "��[�q�H�����ݩ�]=2��3��,���O���� (�ɶU : + )�����g���ʶ���"
          gilCItem1.SetFocus
          blnGoOn = False
        End If
      Else
        strErrMsg = "�䤣�즬�O���ءG" & gilCItem1.GetCodeNo
        gilCItem1.SetFocus
        blnGoOn = False
      End If
      Set rsCD019 = Nothing
    End If
    If (blnGoOn) And (gilCItem2.GetCodeNo <> "") Then
      Set rsCD019 = gcnGi.Execute("Select PeriodFlag From " & GetOwner & "Cd019 Where CodeNo=" & gilCItem2.GetCodeNo)
      If Not rsCD019.EOF Then
        If ((cboCpProperty.ListIndex = 0) Or (cboCpProperty.ListIndex = 3)) And (rsCD019("PeriodFlag") <> "0") Then
          strErrMsg = "��[�q�H�����ݩ�]=1��4��,���O���� (�ɶU : - )�����D�g���ʶ���"
          gilCItem2.SetFocus
          blnGoOn = False
        ElseIf ((cboCpProperty.ListIndex = 1) Or (cboCpProperty.ListIndex = 2)) And (rsCD019("PeriodFlag") <> "1") Then
          strErrMsg = "��[�q�H�����ݩ�]=2��3��,���O���� (�ɶU : - )�����g���ʶ���"
          gilCItem2.SetFocus
          blnGoOn = False
        End If
      Else
        strErrMsg = "�䤣�즬�O���ءG" & gilCItem2.GetCodeNo
        gilCItem1.SetFocus
        blnGoOn = False
      End If
      Set rsCD019 = Nothing
    End If
  End If
  
  If (blnGoOn) And (gilCItem1.GetCodeNo <> "") Then
    If Not CheckRepeatData(gilCItem1.GetCodeNo, "CItemCode1") Then
      strErrMsg = "���O���� (�ɶU : + ) ��ƭ���"
      gilCItem1.SetFocus
      blnGoOn = False
    End If
  End If
  
  If (blnGoOn) And (gilCItem2.GetCodeNo <> "") Then
    If Not CheckRepeatData(gilCItem2.GetCodeNo, "CItemCode2") Then
      strErrMsg = "���O���� (�ɶU : - ) ��ƭ���"
      gilCItem2.SetFocus
      blnGoOn = False
    End If
  End If
  
  If (blnGoOn) And (txtCpCItemCode.Text = "") Then
    strErrMsg = "�q�H�b�涵�إN�X������"
    txtCpCItemCode.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And ((cboCpProperty.ListIndex = 0)) And (txtCpCItemCode.Text <> "") Then
    If Not CheckRepeatData("'" & txtCpCItemCode.Text & "'", "CpCItemCode") Then
      strErrMsg = "�q�H�b�涵�إN�X��ƭ���"
      txtCpCItemCode.SetFocus
      blnGoOn = False
    End If
  End If
  
  If (blnGoOn) And (txtCpCItemName.Text = "") Then
    strErrMsg = "�q�H�b�涵�إN�X������"
    txtCpCItemName.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And ((cboCpProperty.ListIndex = 0) Or (cboCpProperty.ListIndex = 3)) And (txtCpCItemName.Text <> "") Then
    If Not CheckRepeatData("'" & txtCpCItemName.Text & "'", "CpCItemName") Then
      strErrMsg = "�q�H�b�涵�ئW�ٸ�ƭ���"
      txtCpCItemName.SetFocus
      blnGoOn = False
    End If
  End If
  
  'If (blnGoOn) And (txtCpAdjCItemCode.Text = "") Then
  '  strErrMsg = "�q�H�b��էﶵ�إN�X������"
  '  txtCpAdjCItemCode.SetFocus
  '  blnGoOn = False
  'End If
  
  If (blnGoOn) And (txtCpGroupCode.Text = "") Then
    strErrMsg = "�q�H�b��s�եN�X������"
    txtCpGroupCode.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And (txtCpGroupName.Text = "") Then
    strErrMsg = "�q�H�b��s�զW�ٶ�����"
    txtCpGroupName.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And (txtCpItem.Text = "") Then
    strErrMsg = "�q�H�C�L���Ƕ�����"
    txtCpItem.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And (txtCpItem.Text <> "") Then
    If Not CheckRepeatData("'" & txtCpItem.Text & "'", "CpItem") Then
      strErrMsg = "�q�H�C�L���Ǹ�ƭ���"
      txtCpItem.SetFocus
      blnGoOn = False
    End If
  End If
  
  If (blnGoOn) And (cboCpTax.ListIndex < 0) Then
    strErrMsg = "�q�H�|�v�����ȩθ�Ƥ����T"
    cboCpTax.SetFocus
    blnGoOn = False
  End If
  
  If (blnGoOn) And (cboCpProperty.ListIndex < 0) Then
    strErrMsg = "�q�H�����ݩʶ����ȩθ�Ƥ����T"
    cboCpProperty.SetFocus
    blnGoOn = False
  End If
  
  If Not blnGoOn Then
    MsgBox strErrMsg, vbInformation, "�T��"
  End If
  
  IsDataOk = blnGoOn
  
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub ScrToRcd()
On Error GoTo ChkErr
  '�N�e�����󪺭ȶ�J Client Dataset ����줤
  
  Dim strString1 As String
 
  '�s�W�Ҧ��ɭn�� Add �@���Ū���ƦC
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
    strString1 = "�N��"
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
  
  '�s�W�έק��ơA�s�ɮɭn���P�_��ƬO�_�w�s�b�F
  CheckRepeatData = True
  Set rsTempSo165 = gcnGi.Execute("Select RowId From " & GetOwner & "So165 Where " & CheckField & "=" & CheckData)
  If Not rsTempSo165.EOF Then
    '�Y��Ƥw�s�b�A�B���s�W�Ҧ��A�ά��s��Ҧ��A���O�O���P����ƦC�A�h��ܸ�ƭ��ƤF
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
  
  ''�۰ʮھڤ�r�ɪ��ɦW�]�w������ LOG ���ɦW
  strFileName = Fso.GetFileName(SourceFilename)
  strExtensionName = Fso.GetExtensionName(SourceFilename)
  
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
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtExportFile1_Change")
End Sub

Private Sub txtExportFile2_Change()
On Error GoTo ChkErr
  Call SetExporFileName(txtExportPath2, txtExportFile2, txtLogFile5, lblExportFile2, GiExportDate3)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtExportFile2_Change")
End Sub

Private Sub txtExportFile3_Change()
On Error GoTo ChkErr
  Call GiExportDate4_Change
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtExportFile3_Change")
End Sub

Private Sub txtExportFile4_Change()
On Error GoTo ChkErr
  Call GiExportDate5_Change
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtExportFile4_Change")
End Sub

Private Sub txtExportFile5_Change()
On Error GoTo ChkErr
  Call SetExporFileName(txtExportPath5, txtExportFile5, txtLogFile6, lblExportFile5, GiExportDate6)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtExportFile5_Change")
End Sub

Private Sub txtExportPath1_Change()
On Error GoTo ChkErr
  Call SetExporFileName(txtExportPath1, txtExportFile1, txtLogFile4, lblExportFile1, GiExportDate1)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtExportPath1_Change")
End Sub

Private Sub txtExportPath1_LostFocus()
On Error GoTo ChkErr
  If (txtExportPath1.Text <> "") And (Not Fso.FolderExists(txtExportPath1.Text)) Then
    MsgBox txtExportPath1.Text & vbCrLf & vbCrLf & "���|���s�b�C", vbExclamation, "ĵ�i"
    txtExportPath1.Text = ""
    SSTab3.Tab = 0
    txtExportPath1.SetFocus
  End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtExportPath1_LostFocus")
End Sub

Private Sub txtExportPath2_Change()
On Error GoTo ChkErr
  Call SetExporFileName(txtExportPath2, txtExportFile2, txtLogFile5, lblExportFile2, GiExportDate3)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtExportPath2_Change")
End Sub

Private Sub txtExportPath2_LostFocus()
On Error GoTo ChkErr
  If (txtExportPath2.Text <> "") And (Not Fso.FolderExists(txtExportPath2.Text)) Then
    MsgBox txtExportPath2.Text & vbCrLf & vbCrLf & "���|���s�b�C", vbExclamation, "ĵ�i"
    txtExportPath2.Text = ""
    SSTab3.Tab = 1
    txtExportPath2.SetFocus
  End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtExportPath2_LostFocus")
End Sub

Private Sub txtExportPath3_LostFocus()
On Error GoTo ChkErr
  If (txtExportPath3.Text <> "") And (Not Fso.FolderExists(txtExportPath3.Text)) Then
    MsgBox txtExportPath3.Text & vbCrLf & vbCrLf & "���|���s�b�C", vbExclamation, "ĵ�i"
    txtExportPath3.Text = ""
    SSTab3.Tab = 2
    txtExportPath3.SetFocus
  End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtExportPath3_LostFocus")
End Sub

Private Sub txtExportPath4_LostFocus()
On Error GoTo ChkErr
  If (txtExportPath4.Text <> "") And (Not Fso.FolderExists(txtExportPath4.Text)) Then
    MsgBox txtExportPath4.Text & vbCrLf & vbCrLf & "���|���s�b�C", vbExclamation, "ĵ�i"
    txtExportPath4.Text = ""
    SSTab3.Tab = 3
    txtExportPath4.SetFocus
  End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtExportPath4_LostFocus")
End Sub

Private Sub txtExportPath5_Change()
On Error GoTo ChkErr
  Call SetExporFileName(txtExportPath5, txtExportFile5, txtLogFile6, lblExportFile5, GiExportDate6)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtExportPath5_Change")
End Sub

Private Sub txtExportPath5_LostFocus()
On Error GoTo ChkErr
  If (txtExportPath5.Text <> "") And (Not Fso.FolderExists(txtExportPath5.Text)) Then
    MsgBox txtExportPath5.Text & vbCrLf & vbCrLf & "���|���s�b�C", vbExclamation, "ĵ�i"
    txtExportPath5.Text = ""
    SSTab3.Tab = 4
    txtExportPath5.SetFocus
  End If
  Exit Sub
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
  '�O�����~����ƨ� LOG �ɤ�
  Dim strString2 As String
  Dim strString1 As String
  '�渹
  strString1 = ""
  '#3703���դ�OK,�n�h�W�["�q�l�ɤ���"�r�� By Kin 2008/04/24
  If LineNumber > 0 Then
    strString1 = "�q�l�ɤ��� " & Right("00000" & LineNumber, 5) & " �C�G"
  End If
  
  strString2 = ""
  strString1 = strString1 & Reason & Space(90)
  strString1 = StrConv(LeftB(StrConv(strString1, vbFromUnicode), 90), vbUnicode)
  If RightValue <> "" Then
    strString2 = "���T��ơG" & RightValue & Space(60)
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
  '�O�����~����ƨ� LOG �ɤ�
  Dim strString2 As String
  Dim strString1 As String
  strString1 = "��]�G" & Reason
  strmTextFile2.Write strString1 & vbCrLf
  ErrorCounts = ErrorCounts + 1
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "LogError3")
End Sub

Private Sub ExportData1()
On Error GoTo ChkErr
  'CP�q�H�O�P�b��ƶץX
  Dim lngLogCount As Long
  Dim lngExportDetailCount As Long, lngExportMasterCount As Long, lngDataCountTail As Long, lngTmp As Long
  Dim dblRealAmtMaster, dblRealAmtWithTaxMaster  As Double, dblRealAmtWithoutTaxMaster  As Double, dblAdjustAmtTail As Double
  Dim dblRealAmtZeroTaxMaster  As Double, dblAdjustAmtMaster  As Double, dblAdjustAmtWithTaxMaster  As Double, dblAmtTail As Double
  Dim dblAdjustAmtZeroTaxMaster  As Double, dblAdjustAmtWithoutTaxMaster  As Double, dblShouldAmtMaster As Double
  Dim dblTotalRealAmtTail  As Double, dblTotalAdjustAmtTail As Double
  Dim strMaster, strDetail, strRealDateMaster, strTail As String
  Dim strCustId, strFaciSNo, strTfnServiceId, strBillNo, strTFNBillNo As String
  Dim rsTemp As New ADODB.Recordset
  Dim strLog As String
  Dim strString1 As String
  Screen.MousePointer = vbHourglass
  Me.Enabled = False
  
  'strString1 = "Select * From " & GetOwner & "So166 " & _
               "Where (CompCode=" & garyGi(9) & ") " & _
               "And (Servicetype='P') " & _
               "And (RealDate Between To_Date('" & GiRealDate1.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate2.GetValue & "', 'YYYYMMDD')) " & _
               "And (CpExport=0) " & _
               "Order By CustId, FaciSNo, TfnServiceId, BillNo, TfnBillNo "
  '���Ӥw�ץX�L����Ƥ��A�ץX�A�{�b��F
  strString1 = "Select * From " & GetOwner & "So166 " & _
               "Where (CompCode=" & garyGi(9) & ") " & _
               "And (Servicetype='P') " & _
               "And (RealDate Between To_Date('" & GiRealDate1.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate2.GetValue & "', 'YYYYMMDD')) " & _
               "Order By CustId, FaciSNo, TfnServiceId, BillNo, TfnBillNo "
  Set rsTemp = gcnGi.Execute(strString1)
  
  '�g�J���Y
'  If Not rsTemp.EOF Then
    strmTextFile1.Write "H" & GiExportDate1.GetValue & vbCrLf
'  End If
  
  lngExportMasterCount = 0
  lngExportDetailCount = 0
  lngLogCount = 0
  strCustId = ""
  strFaciSNo = ""
  strTfnServiceId = ""
  strBillNo = ""
  strTFNBillNo = ""
  
  '�g�J��
  While Not rsTemp.EOF
    If (strCustId <> rsTemp("CustId") & "") Or (strFaciSNo <> rsTemp("FaciSNo") & "") Or (strTfnServiceId <> rsTemp("TfnServiceId") & "") Or (strBillNo <> rsTemp("BillNo") & "") Or (strTFNBillNo <> rsTemp("TfnBillNo") & "") Then
      '�� CustId, FaciSNo, TfnServiceId, BillNo, TfnBillNo ���ղ��� Master
      If (rsTemp("ShouldAmt") - rsTemp("RealAmt")) >= 0 Then

      If strCustId <> "" Then
        strMaster = "M" & strMaster
        strMaster = strMaster & Left(strCustId & Space(10), 10)
        
        '�w��q�ܸ��X�B�z�A�]���X���q�ܰϽX�M���X���� - �Ÿ��A�ץX�ɭn�o��
        '- ����m
        If Len(strFaciSNo) > 0 Then
          lngTmp = InStr(strFaciSNo, "-")
      
          '�P�_���S���ϽX
          If lngTmp > 0 Then
            strString1 = Left(strFaciSNo, lngTmp - 1)
            strMaster = strMaster & Left(strString1 & Space(3), 3)
          Else
            '�S���A��ť�
            strMaster = strMaster & Space(3)
          End If
      
          '���X�A�� 10 �X
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
          strString1 = "���u��" & strString1
          strString1 = StrConv(LeftB(StrConv(strString1, vbFromUnicode), 50), vbUnicode)
        End If
        
        strMaster = strMaster & strString1
        
        '�N Buffer ������ƶץX
        strmTextFile1.Write strMaster & vbCrLf
        strmTextFile1.Write strDetail
        
        strMaster = ""
        strDetail = ""
      End If
      '�N���ͤ@���s�� Master ��ơA�N���������M�Ťζ�J�w�]��
      dblRealAmtMaster = 0
      dblShouldAmtMaster = 0
      dblRealAmtWithTaxMaster = 0
      dblRealAmtZeroTaxMaster = 0
      dblRealAmtWithoutTaxMaster = 0
      dblAdjustAmtMaster = 0
      dblAdjustAmtWithTaxMaster = 0
      dblAdjustAmtZeroTaxMaster = 0
      dblAdjustAmtWithoutTaxMaster = 0
      strCustId = rsTemp("CustId") & ""
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
    '���էﶵ��
    If (rsTemp("ShouldAmt") - rsTemp("RealAmt")) >= 0 Then
        If rsTemp("ShouldAmt") <> rsTemp("RealAmt") Then
           strDetail = strDetail & "D" & Left(rsTemp("CpAdjCitemCode") & Space(10), 10)
           strDetail = strDetail & Left(rsTemp("CpRate") & Space(1), 1)
           '13-20 �է�O��
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
        
         '�w�ץX�L���n LOG
         If rsTemp("CpExport") <> 0 Then
           strString1 = "�Ƚs(" & rsTemp("CustId") & ")�A" & _
                        "CP����(" & rsTemp("FaciSNo") & ")�A" & _
                        "APBT�A�ȥN�X(" & rsTemp("TfnServiceId") & ")�A" & _
                        "��ڽs��(" & rsTemp("BillNo") & ")�A" & _
                        "APBT�b��s��(" & rsTemp("TfnBillNo") & ")�A" & _
                        "�ꦬ���(" & rsTemp("RealDate") & ")�A" & _
                        "�ꦬ���B(" & rsTemp("RealAmt") & ")�A" & _
                        "������Ƥw�ץX�L"
           Call LogError3(lngLogCount, strString1)
         End If
    Else
           strString1 = "�Ƚs(" & rsTemp("CustId") & ")�A" & _
                        "CP����(" & rsTemp("FaciSNo") & ")�A" & _
                        "APBT�A�ȥN�X(" & rsTemp("TfnServiceId") & ")�A" & _
                        "��ڽs��(" & rsTemp("BillNo") & ")�A" & _
                        "APBT�b��s��(" & rsTemp("TfnBillNo") & ")�A" & _
                        "�ꦬ���(" & rsTemp("RealDate") & ")�A" & _
                        "�ꦬ���B(" & rsTemp("RealAmt") & ")�A" & _
                        "�������B(" & rsTemp("ShouldAmt") & ")�A" & _
                        "�����������B�֩�ꦬ���B"
        Call LogError3(lngLogCount, strString1)
    End If
    rsTemp.MoveNext
  Wend
  '�����]���ᶷ�N Buffer ���|���ץX�� Master �� Detail ��Ƥ]�ץX
  If dblAdjustAmtMaster >= 0 Then
  If strCustId <> "" Then
    strMaster = "M" & strMaster
    strMaster = strMaster & Left(strCustId & Space(10), 10)
    
    '�w��q�ܸ��X�B�z�A�]���X���q�ܰϽX�M���X���� - �Ÿ��A�ץX�ɭn�o��
    '- ����m
    If Len(strFaciSNo) > 0 Then
      lngTmp = InStr(strFaciSNo, "-")
      
      '�P�_���S���ϽX
      If lngTmp > 0 Then
        strString1 = Left(strFaciSNo, lngTmp - 1)
        strMaster = strMaster & Left(strString1 & Space(3), 3)
      Else
        '�S���A��ť�
        strMaster = strMaster & Space(3)
      End If
      
      '���X�A�� 10 �X
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
      strString1 = "���u��" & strString1
      strString1 = StrConv(LeftB(StrConv(strString1, vbFromUnicode), 50), vbUnicode)
    End If
    
    strMaster = strMaster & strString1
    
    '�N Buffer ������ƶץX
    strmTextFile1.Write strMaster & vbCrLf
    strmTextFile1.Write strDetail
    
    strMaster = ""
    strDetail = ""
  End If
  Set rsTemp = Nothing
  End If
  '���o�� TFNBillNo ���ժ�����
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
  '�פJ���������
  dblTotalRealAmtTail = NoZero(rsTemp("TotalRealAmt"), True)
  dblTotalAdjustAmtTail = NoZero(rsTemp("TotalAdjustAmt"), True)
  strTail = "T" & Right("000000" & lngDataCountTail, 6)
  strTail = strTail & Right("000000" & lngExportDetailCount, 6)
  strTail = strTail & Right("000000" & lngDataCountTail + lngExportDetailCount, 6) 'Right("000000" & lngExportMasterCount, 6)
  strTail = strTail & MakeMoney(dblTotalRealAmtTail)
  strTail = strTail & MakeMoney(dblTotalAdjustAmtTail)
  Set rsTemp = Nothing
  strmTextFile1.Write strTail
  '�w�ץX����ƤU�����A�ץX
  strString1 = "Update " & GetOwner & "So166 " & _
               "Set CpExport=1 " & _
               "Where (CompCode=" & garyGi(9) & ") " & _
               "And (Servicetype='P') " & _
               "And (RealDate Between To_Date('" & GiRealDate1.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate2.GetValue & "', 'YYYYMMDD')) " & _
               "And (CpExport=0) "
  gcnGi.Execute (strString1)
'  End If
  
  If lngLogCount > 0 Then
    strString1 = "�ץXMaster���ơG" & lngExportMasterCount & vbCrLf & _
                 "�ץXDetail���ơG" & lngExportDetailCount & vbCrLf & vbCrLf & _
                 "��Ƥ������D����Ʀs�b�A���ơG" & lngLogCount & vbCrLf & _
                 "�аѦҰ��D�O���ɡC" & vbCrLf
    MsgBox strString1, vbExclamation, "�T��"
    Shell "Notepad " & txtLogFile4.Text, vbNormalFocus
'    strmTextFile2
  Else
    strString1 = "�ץX���� !!" & vbCrLf & vbCrLf & _
                 "�ץXMaster���ơG" & lngExportMasterCount & vbCrLf & _
                 "�ץXDetail���ơG" & lngExportDetailCount
    MsgBox strString1, vbExclamation, "�T��"
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
  
  '�N���B�ഫ�� 8 �쪺�r��A���� 8 ��e�� 0
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
  'CP�믲�O��ƶץX
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
  
  '�e���w���O�_����]�O�^
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
                 " Order by A.CustID,A.CPStartDate"     '���Ӥw�ץX�L����Ƥ��A�ץX�A�{�b��F
                 '"And (CpReturnRent=0) "   'Dennis �P lawrence �Q�׼Ȥ��[
  Else
    '�e���w���O�_����]�_�^
    strString1 = "Select * From " & GetOwner & "So167 A, " & GetOwner & "So004 B " & _
                 "Where (A.CustId=B.CustId) " & _
                 "And (A.FaciSno=B.FaciSno) " & _
                 "And (B.InstDate Is Not Null) " & _
                 "And (B.PrDate Is Null) " & _
                 "And (A.CompCode=" & garyGi(9) & ") " & _
                 "And (A.Servicetype='P') " & _
                 "And (A.RealDate Between To_Date('" & GiRealDate3.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate4.GetValue & "', 'YYYYMMDD')) " & _
                 "And (A.BelongYm='" & GiExportDate2.GetValue & "') " & _
                 " Order by A.CustID,A.CPStartDate"     '���Ӥw�ץX�L����Ƥ��A�ץX�A�{�b��F
                 '"And (CpReturnRent=0) "   'Dennis �P lawrence �Q�׼Ȥ��[
  End If
  Set rsTemp = gcnGi.Execute(strString1)
  
  lngDataTotal = 0
  dblRealAmtTotal = 0
  lngLogCount = 0
  
  '�g�J���Y
'  If Not rsTemp.EOF Then
    strmTextFile1.Write "H" & GiExportDate3.GetValue & vbCrLf
'  End If
  
  '�g�J��
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
      If tmpTFNServiceID <> "" Then strSqlIN = "AND B.TFNServiceID='" & rsTemp("TFNServiceID") & "" & "' "
      
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
        '��X
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
        If Not rsTemp2.EOF Then rsTemp2.MoveLast
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
    
    '�w��q�ܸ��X�B�z�A�]���X���q�ܰϽX�M���X���� - �Ÿ��A�ץX�ɭn�o��
    '- ����m
    If Len(rsTemp("FaciSNo")) > 0 Then
      lngTmp = InStr(rsTemp("FaciSNo"), "-")
      
      '�P�_���S���ϽX
      If lngTmp > 0 Then
        strString1 = Left(rsTemp("FaciSNo"), lngTmp - 1)
        strMaster = strMaster & Left(strString1 & Space(3), 3)
      Else
        '�S���A��ť�
        strMaster = strMaster & Space(3)
      End If
      
      '���X�A�� 10 �X
      strString1 = Mid(rsTemp("FaciSNo"), lngTmp + 1, 10)
      strMaster = strMaster & Left(strString1 & Space(10), 10)
    Else
      strMaster = strMaster & Space(13)
    End If
    
    'strMaster = strMaster & Left(rsTemp("FaciSNo") & Space(13), 13)
    strMaster = strMaster & Left(rsTemp("TfnServiceId") & "" & Space(10), 10)
    strMaster = strMaster & Left(rsTemp("BillNo") & Space(16), 16)
    strMaster = strMaster & MakeMoney(rsTemp1("ShouldAmt") & "")
    strMaster = strMaster & MakeMoney(rsTemp1("RealAmt") & "")
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
    
    '�w�ץX�L���n LOG
    If rsTemp("CpExport") <> 0 Then
      strString1 = "�Ƚs(" & rsTemp("CustId") & ")�A" & _
                   "CP����(" & rsTemp("FaciSNo") & ")�A" & _
                   "APBT�A�ȥN�X(" & rsTemp("TfnServiceId") & ")�A" & _
                   "��ڽs��(" & rsTemp("BillNo") & ")�A" & _
                   "������Ƥw�ץX�L"
      Call LogError3(lngLogCount, strString1)
    End If
    
    rsTemp.MoveNext
  Loop
  
  '�g�J���
  strString1 = "T" & Right("000000" & lngDataTotal, 6)
  strString1 = strString1 & MakeMoney(dblRealAmtTotal)
  strmTextFile1.Write strString1 & vbCrLf
  
  Set rsTemp = Nothing
  
  '�w�ץX����ƤU�����A�ץX
  If chkCanExport1.Value Then
    '�e���w���O�_����]�O�^
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
    '�e���w���O�_����]�_�^
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
    strString1 = "�ץX���ơG" & lngDataTotal & vbCrLf & vbCrLf & _
                 "��Ƥ����w�ץX����Ʀs�b�A���ơG" & lngLogCount & vbCrLf & _
                 "�аѦҰ��D�O���ɡC" & vbCrLf
    MsgBox strString1, vbExclamation, "�T��"
    Shell "Notepad " & txtLogFile5.Text, vbNormalFocus
  Else
    strString1 = "�ץX���� !!" & vbCrLf & vbCrLf & _
                 "�ץX���ơG" & lngDataTotal & vbCrLf
    MsgBox strString1, vbExclamation, "�T��"
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
  'CP�h����ƶץX
  Dim dblRealAmtTotal As Double
  Dim lngDataTotal, lngTmp As Long
  Dim strMaster As String
  Dim strString1 As String
  Dim rsTemp As New ADODB.Recordset
  Dim rsTemp2 As New ADODB.Recordset
  
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
  
  '�g�J���Y
'  If Not rsTemp.EOF Then
    strmTextFile1.Write "H" & GiExportDate4.GetValue & vbCrLf
'  End If
  
  '�g�J��
  While Not rsTemp.EOF
    strMaster = "M" & Left(rsTemp("Custid") & Space(10), 10)
    
    '�w��q�ܸ��X�B�z�A�]���X���q�ܰϽX�M���X���� - �Ÿ��A�ץX�ɭn�o��
    '- ����m
    If Len(rsTemp("FaciSNo")) > 0 Then
      lngTmp = InStr(rsTemp("FaciSNo"), "-")
      
      '�P�_���S���ϽX
      If lngTmp > 0 Then
        strString1 = Left(rsTemp("FaciSNo"), lngTmp - 1)
        strMaster = strMaster & Left(strString1 & Space(3), 3)
      Else
        '�S���A��ť�
        strMaster = strMaster & Space(3)
      End If
      
      '���X�A�� 10 �X
      strString1 = Mid(rsTemp("FaciSNo"), lngTmp + 1, 10)
      strMaster = strMaster & Left(strString1 & Space(10), 10)
    Else
      strMaster = strMaster & Space(13)
    End If
        
    'strMaster = strMaster & Left(rsTemp("FaciSNo") & Space(13), 13)
    strMaster = strMaster & Left(rsTemp("TfnServiceId") & Space(10), 10)
    strMaster = strMaster & Left(rsTemp("BillNo") & Space(16), 16)
    strMaster = strMaster & MakeMoney(rsTemp("ShouldAmt"))
    
    '���o���P��
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
  
  '�g�J���
  strString1 = "T" & Right("000000" & lngDataTotal, 6)
  strString1 = strString1 & MakeMoney(dblRealAmtTotal)
  strmTextFile1.Write strString1 & vbCrLf
  Set rsTemp = Nothing
  
  strString1 = "�ץX���� !!" & vbCrLf & vbCrLf & _
               "�ץX���ơG" & lngDataTotal & vbCrLf
  MsgBox strString1, vbExclamation, "�T��"
  
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
  'CMCP�j���ƶץX
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
  
  '�g�J���Y
  lngDataTotal = 0
'  If Not rsTemp.EOF Then
    strmTextFile1.Write "H" & GiExportDate5.GetValue & vbCrLf
'  End If
  
  '�g�J��
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
  
  '�g�J���
  strString1 = "T" & Right("000000" & lngDataTotal, 6)
  strmTextFile1.Write strString1 & vbCrLf
  Set rsTemp = Nothing
  
  strString1 = "�ץX���� !!" & vbCrLf & vbCrLf & _
               "�ץX���ơG" & lngDataTotal & vbCrLf
  MsgBox strString1, vbExclamation, "�T��"
  
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
  'CP�q�H�h�O��ƶץX
  Dim lngExportDetailCount, lngExportMasterCount, lngDataCountTail, lngSubDetailCount, lngTmp As Long
  Dim lngLogCount As Long
  Dim dblRealAmtMaster, dblRealAmtWithTaxMaster, dblRealAmtWithoutTaxMaster, dblRealAmtZeroTaxMaster As Double
  Dim dblTotalRealAmtTail As Double
  Dim strCustId, strFaciSNo, strTfnServiceId, strBillNo, strTFNBillNo As String
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
  
  '���Ӥw�ץX�L����Ƥ��A�ץX�A�{�b��F
  strString1 = "Select * From " & GetOwner & "So170 " & _
               "Where (CompCode=" & garyGi(9) & ") " & _
               "And (Servicetype='P') " & _
               "And (RealDate Between To_Date('" & GiRealDate5.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate6.GetValue & "', 'YYYYMMDD')) " & _
               "Order By CustId, FaciSNo, TfnServiceId, BillNo, TfnBillNo "
  Set rsTemp = gcnGi.Execute(strString1)
  
  '�g�J���Y
'  If Not rsTemp.EOF Then
    strmTextFile1.Write "H" & GiExportDate1.GetValue & vbCrLf
'  End If
  
  lngExportMasterCount = 0
  lngExportDetailCount = 0
  lngLogCount = 0
  strCustId = ""
  strFaciSNo = ""
  strTfnServiceId = ""
  strBillNo = ""
  strTFNBillNo = ""
  
  '�g�J��
  While Not rsTemp.EOF
    '�� CustId, FaciSNo, TfnServiceId, BillNo, TfnBillNo ���ղ��� Master
    If (strCustId <> rsTemp("CustId") & "") Or (strFaciSNo <> rsTemp("FaciSNo") & "") Or (strTfnServiceId <> rsTemp("TfnServiceId") & "") Or (strBillNo <> rsTemp("BillNo") & "") Or (strTFNBillNo <> rsTemp("TfnBillNo") & "") Then
      If strCustId <> "" Then
        strMaster = "M" & strMaster
        strMaster = strMaster & Left(strCustId & Space(10), 10)
        
        '�w��q�ܸ��X�B�z�A�]���X���q�ܰϽX�M���X���� - �Ÿ��A�ץX�ɭn�o��
        '- ����m
        If Len(strFaciSNo) > 0 Then
          lngTmp = InStr(strFaciSNo, "-")
      
          '�P�_���S���ϽX
          If lngTmp > 0 Then
            strString1 = Left(strFaciSNo, lngTmp - 1)
            strMaster = strMaster & Left(strString1 & Space(3), 3)
          Else
            '�S���A��ť�
            strMaster = strMaster & Space(3)
          End If
      
          '���X�A�� 10 �X
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
        
        '�N Buffer ������ƶץX
        strmTextFile1.Write strMaster & vbCrLf
        strmTextFile1.Write strDetail
        
        strMaster = ""
        strDetail = ""
      End If
      '�N���ͤ@���s�� Master ��ơA�N���������M�Ťζ�J�w�]��
      lngSubDetailCount = 0
      dblRealAmtMaster = 0
      dblRealAmtWithTaxMaster = 0
      dblRealAmtZeroTaxMaster = 0
      dblRealAmtWithoutTaxMaster = 0
      strCustId = rsTemp("CustId") & ""
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
    
    '�g�J�Ӷ�
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
    
    '�w�ץX�L���n LOG
    If rsTemp("CpExport") <> 0 Then
      strString1 = "�Ƚs(" & rsTemp("CustId") & ")�A" & _
                   "CP����(" & rsTemp("FaciSNo") & ")�A" & _
                   "APBT�A�ȥN�X(" & rsTemp("TfnServiceId") & ")�A" & _
                   "��ڽs��(" & rsTemp("BillNo") & ")�A" & _
                   "APBT�b��s��(" & rsTemp("TfnBillNo") & ")�A" & _
                   "������Ƥw�ץX�L"
      Call LogError3(lngLogCount, strString1)
    End If
    
    rsTemp.MoveNext
  Wend
  '�����]���ᶷ�N Buffer ���|���ץX�� Master �� Detail ��Ƥ]�ץX
  If strCustId <> "" Then
    strMaster = "M" & strMaster
    strMaster = strMaster & Left(strCustId & Space(10), 10)
    
    '�w��q�ܸ��X�B�z�A�]���X���q�ܰϽX�M���X���� - �Ÿ��A�ץX�ɭn�o��
    '- ����m
    If Len(strFaciSNo) > 0 Then
      lngTmp = InStr(strFaciSNo, "-")
      
      '�P�_���S���ϽX
      If lngTmp > 0 Then
        strString1 = Left(strFaciSNo, lngTmp - 1)
        strMaster = strMaster & Left(strString1 & Space(3), 3)
      Else
        '�S���A��ť�
        strMaster = strMaster & Space(3)
      End If
      
      '���X�A�� 10 �X
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
    
    '�N Buffer ������ƶץX
    strmTextFile1.Write strMaster & vbCrLf
    strmTextFile1.Write strDetail
    
    strMaster = ""
    strDetail = ""
  End If
  Set rsTemp = Nothing
  
  '���o�� TFNBillNo ���ժ�����
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
  
  '�g�J���
  strTail = "T" & Right("000000" & lngDataCountTail, 6)
  strTail = strTail & MakeMoney(dblTotalRealAmtTail)
  Set rsTemp = Nothing
  strmTextFile1.Write strTail
  
  '�w�ץX����ƤU�����A�ץX
  strString1 = "Update " & GetOwner & "So170 " & _
               "Set CpExport=1 " & _
               "Where (CompCode=" & garyGi(9) & ") " & _
               "And (Servicetype='P') " & _
               "And (Realdate Between To_Date('" & GiRealDate5.GetValue & "', 'YYYYMMDD') And To_Date('" & GiRealDate6.GetValue & "', 'YYYYMMDD')) " & _
               "And (CpExport=0) "
  gcnGi.Execute (strString1)
  
  If lngLogCount > 0 Then
    strString1 = "�ץXMaster���ơG" & lngExportMasterCount & vbCrLf & _
                 "�ץXDetail���ơG" & lngExportDetailCount & vbCrLf & vbCrLf & _
                 "��Ƥ����w�ץX����Ʀs�b�A���ơG" & lngLogCount & vbCrLf & _
                 "�аѦҰ��D�O���ɡC" & vbCrLf
    MsgBox strString1, vbExclamation, "�T��"
    Shell "Notepad " & txtLogFile6.Text, vbNormalFocus
  Else
    strString1 = "�ץX���� !!" & vbCrLf & vbCrLf & _
                 "�ץXMaster���ơG" & lngExportMasterCount & vbCrLf & _
                 "�ץXDetail���ơG" & lngExportDetailCount
    MsgBox strString1, vbExclamation, "�T��"
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
  '�򥻸�Ʈֹ�פJ
  Dim blnHasHead, blnHasBody, blnHasTail, blnGoOn, blnHasError As Boolean
  Dim lngHeadErrorCount As Long
  Dim lngTailErrorCount As Long
  Dim lngBodyErrorCount As Long
  Dim lngBodyCount, lngIndex, lngLineNumber As Long
  Dim strTextData, strTel, strDeclarantName, strReason As String
  Dim strString1 As String
  Dim strCustId, strNumber, strTfn, strCs As String
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
    
  Screen.MousePointer = vbHourglass
  Me.Enabled = False
  
  lngBodyErrorCount = 0
  lngHeadErrorCount = 0
  lngTailErrorCount = 0
  lngBodyCount = 0
  blnHasHead = False
  blnHasTail = False
  blnHasBody = False
  
  'Ū�J��r�ɡA�ä��q
  strTextData = Fso.OpenTextFile(txtImportFile2).ReadAll
  vDataArray = Split(strTextData, vbCrLf)
  
  'Progress Bar
  PB3.Visible = True
  PB3.Value = PB3.Min
  On Error Resume Next
  PB3.Max = UBound(vDataArray)
  
  Dim strReplace, strSqlDeName, strSqlDeNameA, strSqlDeNameB, vRepArry
  Dim i As Integer
  
  '�}�l�@��@���R
  lngLineNumber = 1
  For Each vElement In vDataArray
    PB3.Value = lngLineNumber - 1
    Select Case Left(vElement, 1)
      Case "H"
        '���Y
        blnGoOn = True
        If blnHasBody Or blnHasTail Then
          Call LogError1(lngHeadErrorCount, "���Y�����b��r�ɪ��Ĥ@��", "", lngLineNumber, vElement)
          blnGoOn = False
        End If
        vDataArray2 = Split(vElement, ",")
        If UBound(vDataArray2) <> 2 Then
          Call LogError1(lngHeadErrorCount, "���Y���ƥؿ��~�A���u�঳�T��", "", lngLineNumber, vElement)
          blnGoOn = False
        End If
        blnHasHead = True
      Case "B"
        '��ƶ���
        blnGoOn = True
        blnHasBody = True
        If Not blnHasHead Then
          Call LogError1(lngBodyErrorCount, "��Ƥ��i�H�X�{�b���Y���e", "", lngLineNumber, vElement)
          blnGoOn = False
        End If
        If blnHasTail Then
          Call LogError1(lngBodyErrorCount, "��Ƥ��i�H�X�{�b�������", "", lngLineNumber, vElement)
          blnGoOn = False
        End If
        vDataArray2 = Split(vElement, ",")
        If chkTFNSID.Value = 1 Then
            If Trim(UBound(vDataArray2)) <> 11 Then
              Call LogError1(lngBodyErrorCount, "������ƥؿ��~�A��쥲���O�Q�G��", "", lngLineNumber, vElement)
              blnGoOn = False
            End If
        Else
            If Trim(UBound(vDataArray2)) <> 10 Then
              Call LogError1(lngBodyErrorCount, "������ƥؿ��~�A��쥲���O�Q�@��", "", lngLineNumber, vElement)
              blnGoOn = False
            End If
        
        End If
        
        '�q�L�ˮ֡A�g�J TMP012
        If blnGoOn Then
          lngBodyCount = lngBodyCount + 1
          strString1 = "Insert Into Tmp012 " & _
                       "Values ( "
          For lngIndex = 1 To UBound(vDataArray2)  '�]����r�ɪ��̫�@�椣�פJ,�ҥH�n�� 1
            vDataArray2(lngIndex) = Replace(vDataArray2(lngIndex), "%44", ",")
            
            Select Case lngIndex
              Case UBound(vDataArray2) - 3
                '�渹�O���p�~�A�n�ӧO�B�z
                strString1 = strString1 & "'" & Right("00" & CStr(lngLineNumber), 3) & "', "
                vDataArray2(lngIndex) = "'" & vDataArray2(lngIndex) & "', "
              
              Case UBound(vDataArray2) - 2
                vDataArray2(lngIndex) = "'" & vDataArray2(lngIndex) & "', "
              Case Else
                If vDataArray2(lngIndex) = "" Then
                  Call LogError1(lngBodyErrorCount, "����Ƥ���O�ŭ�", "", lngLineNumber, vElement)
                End If
                vDataArray2(lngIndex) = "'" & vDataArray2(lngIndex) & "'"
                If lngIndex <> UBound(vDataArray2) Then
                  vDataArray2(lngIndex) = vDataArray2(lngIndex) & ", "
                Else
                  vDataArray2(lngIndex) = vDataArray2(lngIndex) & ""
                End If
            End Select
            
            strString1 = strString1 & vDataArray2(lngIndex)
          Next
          '�O���ӵ���Ʀb��r�ɪ���m
'          strString1 = strString1 & "'" & lngBodyCount & "'"
          strString1 = strString1 & ")"
'          MsgBox strString1
          cnn.Execute (strString1)
        End If
      Case "T"
        '���
        blnGoOn = True
        If (Not blnHasBody) Or (Not blnHasHead) Then
          Call LogError1(lngTailErrorCount, "��������b��r�ɪ��̫�@��", "", lngLineNumber, vElement)
          blnGoOn = False
        End If
        vDataArray2 = Split(vElement, ",")
        If UBound(vDataArray2) <> 1 Then
          Call LogError1(lngTailErrorCount, "������ƥؿ��~�A���u�঳�G��", "", lngLineNumber, vElement)
          blnGoOn = False
        End If
        If lngBodyCount <> CLng(vDataArray2(1)) Then
          blnGoOn = False
          Call LogError1(lngTailErrorCount, "��Ƶ���(" & (lngBodyCount) & ")�P����O��(" & vDataArray2(1) & ")������", "", lngLineNumber, vElement)
        End If
        blnHasTail = True
      Case Else
        If vElement <> "" Then
          Call LogError1(lngBodyErrorCount, "�L�Ī���ƶ��ءA���ˬd��r���� 1 �X�O�_�Хܿ��~", "", lngLineNumber, vElement)
        End If
    End Select
    lngLineNumber = lngLineNumber + 1
    DoEvents
  Next
  
  '�p�G�S������]����
  If Not blnHasTail Then
    Call LogError1(lngTailErrorCount, "�S��������", "", lngLineNumber, vElement)
    lngLineNumber = lngLineNumber + 1
  End If
  
  '�פJ�L�{�������~
  If (lngBodyErrorCount > 0) Or (lngHeadErrorCount > 0) Or (lngTailErrorCount > 0) Then
    strTextData = "�ˮֶפJ����r�ɦ��~�A�аѦҰ��D�O���ɡC" & vbCrLf
    MsgBox strTextData, vbExclamation, "ĵ�i"
    Shell "Notepad " & txtLogFile2.Text, vbNormalFocus
  Else
    lngBodyErrorCount = 0
    lngBodyCount = 0
    lngLineNumber = 2
    
    PB3.Value = PB3.Min
    
    On Error Resume Next
    PB3.Max = UBound(vDataArray) - 2
    
    Set rsTemp = cnn.Execute("Select * From Tmp012 Order By RecNo")
    While Not rsTemp.EOF
      PB3.Value = lngLineNumber - 2
      '�ϽX + �q�ܸ��X
      strTel = Trim(Trim(rsTemp("CpAreaCode")) & "-" & rsTemp("CpNumber"))
      '�Τ�W��
      strDeclarantName = StrConv(LeftB(StrConv(rsTemp("CustName"), vbFromUnicode), 50), vbUnicode)
      '�A�ȥN�X(�x�T�N�X)

'      strTfnServiceId = Trim(rsTemp("TFNServiceID"))
      '�L�y���ͤp�j�ɧ�ꪺ�y�k
      strReplace = "": strSqlDeName = "": strSqlDeNameA = "": strSqlDeNameB = ""
      strReplace = txtfilter
      vRepArry = Split(strReplace, ",")
      strSqlDeName = "(SO004.DeclarantName"
      For i = 0 To UBound(vRepArry)
        strSqlDeNameA = strSqlDeNameA & "(REPLACE"
        strSqlDeNameB = ",'" & vRepArry(i) & "','')" & strSqlDeNameB

      Next
      strSqlDeName = strSqlDeNameA & strSqlDeName & strSqlDeNameB
      '�P�_�n�פJ����Ʀb SO004 �O�_�w�s�b�F
      
'      strString1 = "Select SeqNo, InstDate, PrDate " & _
                   "From " & GetOwner & "So004 " & _
                   "Where (SO004.CompCode=" & garyGi(9) & ") " & _
                   "And (SO004.Servicetype='P') " & _
                   "And (SO004.InstDate Is Not Null) " & _
                   "And (SO004.PrDate Is Null) " & _
                   "And (SO004.FaciSNo='" & strTel & "' ) " & _
                   "And ((SO004.Id='" & rsTemp("Id") & "' ) " & _
                   "  Or (SO004.Id2='" & rsTemp("Id") & "' )) " & _
                   "And " & strSqlDeName & "='" & strDeclarantName & "' ) " & _
                   "AND (SO004.EBTContNo='" & rsTemp("EBTContNo") & "') "
                   
      '���D��2722 ----for Taos
      '�H 1.�}�q�X���s���A2.CP���� �G����ӧP�_
      '����,�Y Update TfnServiceId �çP�_��1.�Ƚs,2.�����Ҧr��,3.�ӸˤH�m�W,4.�]�Ʀw�ˤ�� ��
      '�L��,�Y�Xlog
      strString1 = "Select SeqNo,CustId,InstDate ,PrDate, EBTContNo,TFNServiceID,Id,Id2,FaciSNo,SNo, " & strSqlDeName & ") DeclarantName " & _
                   "From " & GetOwner & "So004 " & _
                   "Where (SO004.CompCode=" & garyGi(9) & ") " & _
                   "And (SO004.FaciSNo='" & strTel & "' ) " & _
                   "AND (SO004.EBTContNo='" & rsTemp("EBTContNo") & "') "

      Set rsTemp2 = gcnGi.Execute(strString1)

        lngIndex = lngBodyErrorCount
'      �O�A�w�s�b�A�n Update �������

      If Not rsTemp2.EOF Then
        strString1 = "Update " & GetOwner & "So004 " & _
                     "Set TfnServiceId='" & rsTemp("TfnServiceId") & "', " & _
                     "  TfnStatus='" & rsTemp("TfnStatus") & "' " & _
                     "Where SeqNo='" & rsTemp2("SeqNo") & "' "
        gcnGi.Execute (strString1)
        ErrorType = False
        
        If Trim(rsTemp("CustId") & "") <> rsTemp2("CustId") & "" Then
           Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�Ȥ�s�����X!", rsTemp("CustId") & "", rsTemp2("CustId") & "", rsTemp2("SNo") & "", rsTemp2("InstDate") & "", rsTemp2("PrDate") & "", rsTemp("RecNo") & "")
        End If
            
        If (UCase(rsTemp2("Id")) & "" <> UCase(rsTemp("Id")) & "") And (UCase(rsTemp2("Id2")) & "" <> UCase(rsTemp("Id")) & "") Then
          '�P�_������
          If (rsTemp2("Id") & "") <> "" Then
            strString1 = rsTemp2("Id") & ""
          Else
            strString1 = rsTemp2("Id2") & ""
          End If
          Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�����Ҧr�����X!", rsTemp("Id") & "", strString1 & "", rsTemp2("SNo") & "", rsTemp2("InstDate") & "", rsTemp2("PrDate") & "", rsTemp("RecNo") & "")
        End If
        
        If Trim(rsTemp2("DeclarantName")) & "" <> strDeclarantName Then
          '�P�_�Ȥ�W��
          Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�n�O�m�W���X!", strDeclarantName & "", rsTemp2("DeclarantName") & "", rsTemp2("SNo") & "", rsTemp2("InstDate") & "", rsTemp2("PrDate") & "", rsTemp("RecNo") & "")
        End If
        
        If (rsTemp("InstDate") & "") <> Format(rsTemp2("InstDate") & "", "YYYYMMDD") Then
          Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�A�ȱҩl������X!", rsTemp("InstDate") & "", rsTemp2("InstDate") & "", rsTemp2("SNo") & "", rsTemp2("InstDate") & "", rsTemp2("PrDate") & "", rsTemp("RecNo") & "")
        End If
        
        If (rsTemp("PrDate") & "") <> Format(rsTemp2("PrDate") & "", "YYYYMMDD") Then
          Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�A�ȵ���������X!", rsTemp("PrDate") & "", rsTemp2("PrDate") & "", rsTemp2("SNo") & "", rsTemp2("InstDate") & "", rsTemp2("PrDate") & "", rsTemp("RecNo") & "")
        End If
'        If lngBodyErrorCount <> lngIndex Then
'         lngBodyErrorCount = lngIndex + 1
'        End If
        If ErrorType = True Then
            lngBodyErrorCount = lngIndex + 1
        End If
'      Else
         '���D2722 ----for Taos �G�����������{���X�P�_
'        '�_�A��X���~ LOG
               
'        '���P�_�Y�w����B�A�Ȱ_�l����M�A�Ȥ��������ŦX

'        Set rsTemp2 = Nothing
'        strString1 = "Select CustId,InstDate ,PrDate, EBTContNo,TFNServiceID,Id,Id2,FaciSNo,SNo, " & strSqlDeName & ") DeclarantName " & _
'                     "From " & GetOwner & "So004 " & _
'                     "Where  (SO004.CompCode=" & garyGi(9) & ") " & _
'                     "And (SO004.Servicetype='P') " & _
'                     "And (SO004.PrDate Is Null) " & _
'                     "AND (SO004.EBTContNo='" & rsTemp("EBTContNo") & "') "
'
'        Set rsTemp2 = gcnGi.Execute(strString1)
'
'        ErrorType = False
'        If Not rsTemp2.EOF Then
'            lngIndex = lngBodyErrorCount
'              If rsTemp2("FaciSNo") & "" <> strTel Then
'                    Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "��CP�������s�b!", strTel, "", rsTemp2("SNo") & "", rsTemp2("InstDate") & "", rsTemp2("PrDate") & "", rsTemp("RecNo") & "")
'              End If
'            If Trim(rsTemp("CustId") & "") <> rsTemp2("CustId") & "" Then
'               Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�Ȥ�s�����X!", rsTemp("CustId") & "", rsTemp2("CustId") & "", rsTemp2("SNo") & "", rsTemp2("InstDate") & "", rsTemp2("PrDate") & "", rsTemp("RecNo") & "")
'            End If
'
'            If (UCase(rsTemp2("Id")) & "" <> UCase(rsTemp("Id")) & "") And (UCase(rsTemp2("Id2")) & "" <> UCase(rsTemp("Id")) & "") Then
'              '�P�_������
'              If (rsTemp2("Id") & "") <> "" Then
'                strString1 = rsTemp2("Id") & ""
'              Else
'                strString1 = rsTemp2("Id2") & ""
'              End If
'              Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�����Ҧr�����X!", rsTemp("Id") & "", strString1 & "", rsTemp2("SNo") & "", rsTemp2("InstDate") & "", rsTemp2("PrDate") & "", rsTemp("RecNo") & "")
'            End If
'
'            If Trim(rsTemp2("DeclarantName")) & "" <> strDeclarantName Then
'              '�P�_�Ȥ�W��
'              Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�n�O�m�W���X!", strDeclarantName & "", rsTemp2("DeclarantName") & "", rsTemp2("SNo") & "", rsTemp2("InstDate") & "", rsTemp2("PrDate") & "", rsTemp("RecNo") & "")
'            End If
'
'            If rsTemp2("FaciSNo") & "" <> strTel Then
'              '�P�_ FaciSNo
'              Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�Ȥ�]�Ƨ䤣��!", strTel & "", rsTemp2("FaciSNo") & "", rsTemp2("SNo") & "", rsTemp2("InstDate") & "", rsTemp2("PrDate") & "", rsTemp("RecNo") & "")
'            End If
'
'            If (rsTemp("InstDate") & "") <> Format(rsTemp2("InstDate") & "", "YYYYMMDD") Then
'              Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�A�ȱҩl������X!", rsTemp("InstDate") & "", rsTemp2("InstDate") & "", rsTemp2("SNo") & "", rsTemp2("InstDate") & "", rsTemp2("PrDate") & "", rsTemp("RecNo") & "")
'            End If
'            If (rsTemp("PrDate") & "") <> Format(rsTemp2("PrDate") & "", "YYYYMMDD") Then
'              Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�A�ȵ���������X!", rsTemp("PrDate") & "", rsTemp2("PrDate") & "", rsTemp2("SNo") & "", rsTemp2("InstDate") & "", rsTemp2("PrDate") & "", rsTemp("RecNo") & "")
'            End If
          '���D��2422 Lydia
'            If chkTFNSID.Value = 1 Then
'              If Trim(rsTemp("EBTContNo") & "") <> rsTemp2("EBTContNo") & "" Then
'                '�P�_ TFNServiceID
'                Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "EBT�X���s������!", rsTemp("EBTContNo") & "", rsTemp2("FaciSNo") & "", rsTemp("RecNo") & "")
'              End If
'            End If
        Else
            Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "EBT�X���s������!", rsTemp("EBTContNo") & "", "", "", "", "", rsTemp("RecNo") & "")

        End If
'      End If
'--------------------
        If ErrorType = True Then
            lngBodyErrorCount = lngIndex + 1
        End If
      lngBodyCount = lngBodyCount + 1
      lngLineNumber = lngLineNumber + 1
      rsTemp.MoveNext
    Wend
    Set rsTemp = Nothing
    PB3.Visible = False
'          If lngBodyErrorCount <> lngIndex Then
'            lngBodyErrorCount = lngIndex + 1
''          End If
''          Set rsTemp2 = Nothing
''        Else
''          '�H CustId, FaciSNo, Id, DeclarantName �ӧO�P�_�A�H�K��Ϥ����~����]
''          Set rsTemp2 = Nothing
''
''          '�P�_ CustId
''          strString1 = "Select SO004.SeqNo, SO004.FaciSNo, SO004.Id, SO004.Id2, " & strSqlDeName & ")" & _
''                       "From " & GetOwner & "So004," & GetOwner & "CD022 " & _
''                       "Where SO004.FaciCode=CD022.CodeNo " & _
''                       "AND (SO004.CustId=" & rsTemp("CustId") & " ) " & _
''                       "And (SO004.CompCode=" & garyGi(9) & ") " & _
''                       "And (SO004.Servicetype='P') " & _
''                       "And (SO004.InstDate Is Not Null) " & _
''                       "And (SO004.PrDate Is Null) " & _
''                       "AND CD022.RefNo=6"
''
''          If Not GetRS(rsTemp2, strString1, gcnGi, , adOpenKeyset, adLockOptimistic, adUseClient) Then Exit Sub
''            If chkTFNSID.Value = 1 Then
''              If Trim(rsTemp("EBTContNo") & "") <> rsTemp2("EBTContNo") & "" Then
''                '�P�_ TFNServiceID
''                Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "EBT�X���s������!", rsTemp("EBTContNo") & "", rsTemp2("EBTContNo") & "", rsTemp("RecNo") & "")
''              End If
''            End If
'          lngIndex = lngBodyErrorCount
'          If rsTemp2.EOF Then
'            '��M���n�[�J�H�����Ψ����ҤΦW�٨ӧP�_
'            Set rsTemp2 = Nothing
'            strString1 = "Select SO004.CustId " & _
'                         "From " & GetOwner & "So004," & GetOwner & "CD022 " & _
'                         "Where SO004.FaciCode=CD022.CodeNo " & _
'                         "AND (SO004.FaciSNo='" & strTel & "') " & _
'                         "And (SO004.CompCode=" & garyGi(9) & ") " & _
'                         "And (SO004.Servicetype='P') " & _
'                         "And (SO004.InstDate Is Not Null) " & _
'                         "And (SO004.PrDate Is Null) " & _
'                         "AND CD022.RefNo=6"
'
'            Set rsTemp2 = gcnGi.Execute(strString1)
          
'            If Not rsTemp2.EOF Then
'            '�����b�A���N�O�Ƚs����F
'              Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�Ȥ�s�����X!", rsTemp("CustId") & "", rsTemp2("CustId") & "", rsTemp("RecNo") & "")
'            Else

'             If Trim(rsTemp("CustId") & "") <> rsTemp2("CustId ") & "" Then
'                Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�Ȥ�s�����X!", rsTemp("CustId") & "", rsTemp2("CustId") & "", rsTemp("RecNo") & "")
'             End If
'            '���M�N�A��鶴����
'              Set rsTemp2 = Nothing
'              strString1 = "Select SO004.CustId " & _
'                       "From " & GetOwner & "So004," & GetOwner & "CD022 " & _
'                       "Where SO004.FaciCode=CD022.CodeNo " & _
'                       "AND ((SO004.Id='" & rsTemp("Id") & "' ) " & _
'                       "  Or (SO004.Id2='" & rsTemp("Id") & "' )) " & _
'                       "And (SO004.CompCode=" & garyGi(9) & ") " & _
'                       "And (SO004.Servicetype='P') " & _
'                       "And (SO004.InstDate Is Not Null) " & _
'                       "And (SO004.PrDate Is Null) " & _
'                       "AND CD022.RefNo=6"
'              Set rsTemp2 = gcnGi.Execute(strString1)
'              If Not rsTemp2.EOF Then
'                '�����Ҧb�A���N�O�Ƚs����F
'                Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�Ȥ�s�����X!", rsTemp("CustId") & "", rsTemp2("CustId") & "", rsTemp("RecNo") & "")
'              Else
'                Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "��CP�������s�b!", strTel, "", rsTemp("RecNo") & "")
'              End If
'            End If
'            Set rsTemp2 = Nothing
'          Else
'            lngIndex = lngBodyErrorCount
'            While Not rsTemp2.EOF
'              If (rsTemp2("Id") & "" <> rsTemp("Id") & "") And (rsTemp2("Id2") & "" <> rsTemp("Id") & "") Then
'                �P�_������
'                If (rsTemp2("Id") & "") <> "" Then
'                  strString1 = rsTemp2("Id") & ""
'                Else
'                  strString1 = rsTemp2("Id2") & ""
'                End If
'                Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�����Ҧr�����X!", rsTemp("Id") & "", strString1 & "", rsTemp("RecNo") & "")
'              End If
'
'              If rsTemp2("DeclarantName") & "" <> strDeclarantName Then
'                �P�_�Ȥ�W��
'                Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�n�O�m�W���X!", strDeclarantName & "", rsTemp2("DeclarantName") & "", rsTemp("RecNo") & "")
'              End If
'
'              If rsTemp2("FaciSNo") & "" <> strTel Then
'                �P�_ FaciSNo
'                Call LogError22(lngBodyErrorCount, rsTemp("TFNServiceID") & "", rsTemp("CustId") & "", strTel & "", "�Ȥ�]�Ƨ䤣��!", strTel & "", rsTemp2("FaciSNo") & "", rsTemp("RecNo") & "")
'              End If
'
'              rsTemp2.MoveNext
'            Wend
'          End If
'          lngBodyErrorCount = lngIndex + 1
'        End If
'      End If
'      Set rsTemp2 = Nothing
      
'      lngBodyCount = lngBodyCount + 1
'      lngLineNumber = lngLineNumber + 1
'      rsTemp.MoveNext
'    Wend
'    Set rsTemp = Nothing
'    PB3.Visible = False
'
    If lngBodyErrorCount > 0 Then
      strString1 = "�פJ�ֹﵧ�ơG" & lngBodyCount & vbCrLf & _
                   "�ֹ���~���ơG" & lngBodyErrorCount & vbCrLf & vbCrLf & _
                   "�פJ��Ƥ������~���ơA�аѦҰ��D�O���ɡC" & vbCrLf
      MsgBox strString1, vbExclamation, "ĵ�i"
      Shell "Notepad " & txtLogFile2.Text, vbNormalFocus
    Else
      strString1 = "�פJ�ֹﵧ�ơG" & lngBodyCount & vbCrLf & vbCrLf & _
                   "�פJ�����A�Ҧ���ƬҲŦX�L�~ !!" & vbCrLf
      MsgBox strString1, vbExclamation, "�T��"
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
      .Filter = "��r�� (*.txt;*.dat)|*.txt;*.dat|�Ҧ��ɮ� (*.*)|*.*"
      .Action = 1
      If .FileName <> Filters & "*.txt" Then
        ImportFile.Text = .InitDir & .FileName
      End If
    End With
  
  '�۰ʮھڤ�r�ɪ��ɦW�]�w������ LOG ���ɦW
    Call SetLogFileName(ImportFile.Text, LogFile)
  
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "SelectImportFile2")
End Sub

Private Sub ImportData3()
  On Error GoTo ChkErr
    'CP�q�ܩ��Ӹ�ƶפJ
    Dim blnGoOn, blnGoOn2 As Boolean
    Dim lngYear, lngMonth As Long
    Dim strTextData, StrTableName, strCheckName, strFileName, strErrStr, strErrData As String
    Dim strViewName1 As String
    Dim lngLastLine As Long, lngIndex As Long, lngIndex2 As Long, lngInsertCount As Long, lngUpdateCount As Long
    Dim lngErrorCount As Long
    Dim vDataArray, vElement, vDataArray2, vDataArrayA, vDataArrayB As Variant
    Dim rsTemp As New ADODB.Recordset
    Dim rsTemp2 As New ADODB.Recordset
    Dim strString1 As String
    'Dim rsSo168 As New ADODB.Recordset
    Dim strCustId As String
    Dim blnNoImport As Boolean
    Dim rsSO004 As New ADODB.Recordset
    Set rsSo168 = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    
    blnNoImport = False
    lngErrorCount = 0
    
    strFileName = Fso.GetFileName(txtImportFile5.Text)
    StrTableName = GetOwner & IIf(optDetailDaily.Value, "So168", "So169")
    
    'Ū�J��r�ɡA�ä��q
    strTextData = Fso.OpenTextFile(txtImportFile5).ReadAll
    vDataArray = Split(strTextData, vbCrLf)
    
    '��X�̫�@�檺��m
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
    
    '�ˮֶפJ�ɦW
    strCheckName = IIf(optDetailDaily.Value, txtImportFile3.Text, txtImportFile4.Text)
    
    blnGoOn = True
    If Left(vDataArray(0), 1) <> "H" Then
        Call LogError1(lngErrorCount, "���Y�����b��r�ɪ��Ĥ@��", "", 1, vDataArray(0))
        blnGoOn = False
    End If
    If Left(vDataArray(lngLastLine), 1) <> "T" Then
        Call LogError1(lngErrorCount, "��������b��r�ɪ��̫�@��", "", 0, vDataArray(lngLastLine))
        blnGoOn = False
    End If
    If Val(Mid(vDataArray(lngLastLine), 2, 7)) <> lngLastLine - 1 Then
        Call LogError1(lngErrorCount, "��Ƶ���(" & (lngLastLine - 1) & ")�P����O��(" & Val(Mid(vDataArray(lngLastLine), 2, 7)) & ")������", "", lngLastLine, vDataArray(lngLastLine))
        blnGoOn = False
    End If
    If UCase(Left(strFileName, Len(strCheckName))) <> UCase(strCheckName) Then
        Call LogError1(lngErrorCount, "'�פJ�ɮצW��' �P " & "'�ˮֶפJ�ɦW' ����", "", 0, "")
        blnGoOn = False
    End If
    
    blnGoOn2 = True
    If optDetailMonthly.Value And blnGoOn Then
        '���y��r�ɤ��e�A�ݨC����ƪ��X�b�~��O�_�����@��
        blnGoOn2 = True
        strString1 = ""
        If blnGoOn2 Then
            '�YCDR�פJ�C��A�h�ˬd�X�b�~��O�_�w�s�b�F�A�Y�O�A�N��Ʈw���P�X�b�~�몺��ƧR��
            blnGoOn2 = True
            If optDetailMonthly.Value And blnGoOn Then
                vDataArray2 = Split(vDataArray(1), ",")
                strString1 = "Select CustId From " & GetOwner & "So169 Where " & IIf(Trim(txtbillYM) <> "", " Billing_Period='" & Trim(txtbillYM) & "'", " Billing_Period is Null")
                If Not GetRS(rsTemp, strString1, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                'Set rsTemp = gcnGi.Execute("Select CustId From " & GetOwner & "So169 Where " & IIf(Trim(txtbillYM) <> "", " Billing_Period=" & Trim(txtbillYM), " Billing_Period is Null"))
                If Not rsTemp.EOF Then
                    blnGoOn2 = MsgBox("�X�b�~�� " & Trim(txtbillYM) & "�G�Ӥ��CDR��Ƥw�פJ�A�O�_�M�ŸӤ����CDR��ơA�A�פJ�H", vbYesNo, "ĵ�i") = 6
                    '#3703 ���դ�OK,�p�G��ܧ_ �h�̫ᤣ�X�X�{����ĵ�i�T�� By Kin 2008/04/23
                    blnNoImport = Not blnGoOn2
                    If blnGoOn2 Then
                        gcnGi.Execute ("Delete From " & GetOwner & "So169 Where " & IIf(Trim(txtbillYM) <> "", "Billing_Period ='" & txtbillYM & "'", " Billing_Period is Null"))
                    End If
                End If
                Set rsTemp = Nothing
            End If
        End If
    End If
  'EBTContNo
    strString1 = ""
    If blnGoOn2 And blnGoOn Then
        '���y��r�ɤ��e�A�ݬݬO�_�����ƪ���ơ]�Ƚs�A�����A�o�ܰ_�l�ɶ�, ���ܤ踹�X�^
        '2006.03.03 ��Φs�Ȧs�ɪ��覡�Ӱ��A���M��Ʀh�ɷ|�ܺC
        
        strViewName1 = GetTmpViewName
        On Error Resume Next
        gcnGi.Execute "Drop Table " & GetOwner & strViewName1
        strString1 = "Create Table " & GetOwner & strViewName1 & "(CustId Number(8), Calling_Area_Code Varchar2(9), Calling_Nbr Varchar2(20), Start_Time Date, Called_Area_Code Varchar2(9), Called_Nbr Varchar2(20), RecNo Long)"
        '�Ƚs,�o�ܤ�ϰ츹�X,�o�ܤ�q�ܸ��X,�o�ܰ_�l�ɶ�,���ܤ�ϰ츹�X,���ܤ�q�ܸ��X
        gcnGi.Execute strString1
        strString1 = "Create Index I_Tmp On " & GetOwner & strViewName1 & "(CustId, Calling_Area_Code, Calling_Nbr, Start_Time, Called_Area_Code, Called_Nbr)"
        gcnGi.Execute strString1
        
        For lngIndex = 1 To lngLastLine - 1
            PB2.Value = lngIndex - 1
            vDataArrayA = Split(vDataArray(lngIndex), ",")
            strString1 = "Insert Into " & GetOwner & strViewName1 & " " & _
                         "Values ( '" & _
                         vDataArrayA(0) & "', '" & _
                         vDataArrayA(2) & "', '" & _
                         vDataArrayA(3) & "', " & _
                         "To_Date('" & vDataArrayA(10) & "', 'YYYY/MM/DD HH24:MI:SS'), '" & _
                         vDataArrayA(7) & "', '" & _
                         vDataArrayA(8) & "', " & _
                         CStr(lngIndex) & ") "
            gcnGi.Execute strString1
            DoEvents
        Next
  
        On Error Resume Next
        gcnGi.Execute "Drop Table " & GetOwner & strViewName1
    End If
  
    gcnGi.BeginTrans
    lngInsertCount = 0
    lngUpdateCount = 0
    'If Not GetRS(rsTmp, "Select * From " & GetOwner & StrTableName & " Where 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If blnGoOn2 And blnGoOn Then
      '�Y�פJ���ѡA�h�N�w�פJ SO168/SO169 ����ƨ���
        blnGoOn = True
      
        For lngIndex = 1 To lngLastLine - 1
            PB2.Value = lngIndex - 1
            vDataArray2 = Split(vDataArray(lngIndex), ",")
            If Not IsDate(vDataArray2(10)) Then
                Call LogError1(lngErrorCount, "�Ƚs�G" & vDataArray2(0) & "�@�o�ܰ_�l�ɶ��G" & vDataArray2(10) & "�A�榡���~", "", lngIndex + 1, vDataArray(lngIndex))
                blnGoOn = False
            End If
        
            If Not IsNumeric(vDataArray2(12)) Then
                Call LogError1(lngErrorCount, "�Ƚs�G" & vDataArray2(0) & "�@�o�ܤ輷����ơG" & vDataArray2(12) & "�A�榡���~", "", lngIndex + 1, vDataArray(lngIndex))
                blnGoOn = False
            End If

            If Not IsNumeric(vDataArray2(14)) Then
                Call LogError1(lngErrorCount, "�Ƚs�G" & vDataArray2(0) & "�@���B�G" & vDataArray2(14) & "�A�榡���~", "", lngIndex + 1, vDataArray(lngIndex))
                blnGoOn = False
            End If
            If Not IsNumeric(vDataArray2(16)) Then
                Call LogError1(lngErrorCount, "�Ƚs�G" & vDataArray2(0) & "�@APBT�Τ�N���G" & vDataArray2(16) & "�A�榡���~", "", lngIndex + 1, vDataArray(lngIndex))
                blnGoOn = False
            End If
        
            If blnGoOn Then
                '2006.02.20 �אּ�����C��ΨC��A�Haccount_id,called_area_code,called_nbr,start_time��key�A�P�_���S������
                '#3703 �h��@��Hour_Type By Kin 2008/03/03
                strString1 = ""
'                strString1 = "Select * From " & StrTableName & " " & _
'                             "Where (Account_Id='" & vDataArray2(1) & "') " & _
'                             "And (Called_Nbr='" & vDataArray2(8) & "') " & _
'                             "And (Start_Time=To_Date('" & vDataArray2(10) & "', 'YYYY/MM/DD HH24:MI:SS')) " & _
'                             "AND (Duration='" & vDataArray2(10) & "') " & _
'                             "And (hour_type='" & vDataArray2(15) & "')"

                '#3703 ���դ�OK �쥻��Duration,�N����쮳��,���Calling_Area_Code By Kin 2008/04/23
                strString1 = "Select * From " & StrTableName & " " & _
                             "Where (Account_Id='" & vDataArray2(1) & "') " & _
                             "And (Called_Nbr='" & vDataArray2(8) & "') " & _
                             "And (Start_Time=To_Date('" & vDataArray2(10) & "', 'YYYY/MM/DD HH24:MI:SS')) " & _
                             "And (hour_type='" & vDataArray2(15) & "') " & _
                             "And " & IIf(Trim(txtAreaCode.Text) = "", "Calling_Area_Code is Null", "Calling_Area_Code='" & txtAreaCode.Text & "'")
                             
                On Error GoTo ChkErr
                If Not GetRS(rsSo168, strString1, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                If rsSo168.EOF Then
                    rsSo168.AddNew
                    strCustId = ""
                    If Not GetRS(rsSO004, "SELECT Distinct(Custid) FROM " & GetOwner & "SO004 WHERE TFNSERVICEID= '" & vDataArray2(16) & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                    If Not rsSO004.EOF Then strCustId = rsSO004("CustId") & ""
                    'strCustId = Val(GetRsValue("SELECT Custid FROM " & GetOwner & "SO004 WHERE TFNSERVICEID= '" & vDataArray2(16) & "' ", gcnGi))
                    lngInsertCount = lngInsertCount + 1
                    If strCustId <> "" Then
                        rsSo168("CustId") = strCustId  'vDataArray2(0)
                    End If
                    rsSo168("Catv_Nbr") = vDataArray2(0)
                    rsSo168("Account_Id") = vDataArray2(1)
                    rsSo168("Calling_Area_Code") = IIf(Trim(txtAreaCode) = "", Null, Trim(txtAreaCode)) 'vDataArray2(2)
                    rsSo168("Calling_Nbr") = Right(vDataArray2(3), IIf(Trim(txtNumCount) = "", 0, Val(txtNumCount)))
                    rsSo168("Bill_Item_Name") = vDataArray2(4)
                    rsSo168("Area_Name") = vDataArray2(5)
                    rsSo168("Access_Code") = vDataArray2(6)
                    rsSo168("Called_Area_Code") = "0" 'vDataArray2(7)
                    rsSo168("Called_Nbr") = vDataArray2(8)
                    rsSo168("Billing_Period") = Trim(txtbillYM) 'vDataArray2(9)
                    rsSo168("Start_Time") = vDataArray2(10)
                    rsSo168("End_Time") = DateAdd("s", vDataArray2(12), vDataArray2(10))  'vDataArray2(11)
                    rsSo168("Duration") = vDataArray2(12)
                    rsSo168("Counts") = 0   'vDataArray2(13)
                    rsSo168("Charge") = vDataArray2(14)
                    rsSo168("Hour_Type") = vDataArray2(15)
                    rsSo168("Subscr_Id") = vDataArray2(16)
                    rsSo168("CreatDate") = Mid(vDataArray(0), 2, 4) & "/" & Mid(vDataArray(0), 6, 2) ' & "/" & Mid(vDataArray(0), 8, 2)
                    rsSo168("UpdTime") = GetDTString(RightNow)
                    rsSo168("UpdEn") = garyGi(0)
                    rsSo168.Update
                    'rsSo168.Close
                    'DoEvents
                Else
                    'lngUpdateCount = lngUpdateCount + 1
                    '***********************************************************************************************************************************************************************************
                    '#3703 ���դ�OK,���~�T���S���Ƚs By Kin 2008/04/24
                    strCustId = ""
                    If Not GetRS(rsSO004, "SELECT Distinct(Custid) FROM " & GetOwner & "SO004 WHERE TFNSERVICEID= '" & vDataArray2(16) & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                    If Not rsSO004.EOF Then strCustId = rsSO004("CustId") & ""
                    '**********************************************************************************************************************************************************************************
                    Call LogError1(lngErrorCount, "�Ƚs�G" & strCustId & "�@�n�פJ����Ƥw�s�b���Ʈw��", "", lngIndex + 1, vDataArray(lngIndex))
                    'rsSo168.Close
                End If

            End If
            'DoEvents
        Next
        PB2.Visible = False
    End If
    gcnGi.CommitTrans
    '***************************************************************************************************
    '#3703 �P�_�O�_����ܲM���A���s�פJ,�p�G��ܧ_�h���X����ĵ�i By Kin 2008/04/23
    If Not blnNoImport Then
        If lngErrorCount > 0 Or lngUpdateCount > 0 Then
          '�����~�A�w�פJ����ƶ������פJ
          '#3703 �p�G��Ʀ����ƭnShow�X��r�� By Kin 2008/03/03
        '    gcnGi.RollbackTrans
          strString1 = "��r�ɵ��ơG" & lngLastLine - 1 & vbCrLf & _
                       "�s�W���� : " & lngInsertCount & vbCrLf & _
                       "�ֹ���~/���D�ơG" & lngErrorCount + lngUpdateCount & vbCrLf & vbCrLf & _
                       "�פJ��Ƥ������~/���D���ơA�аѦҰ��D�O���ɡC" & vbCrLf
          MsgBox strString1, vbExclamation, "ĵ�i"
          Shell "Notepad " & txtLogFile3.Text, vbNormalFocus
        Else
            '�S���~�A�g�J��Ʈw��
            strString1 = "�פJ���� !!" & vbCrLf & vbCrLf & _
                         "��r�ɵ��ơG" & lngLastLine - 1 & vbCrLf & _
                         "�s�W���ơG" & lngInsertCount & vbCrLf
    
                         '"�קﵧ�ơG" & lngUpdateCount & vbCrLf
            MsgBox strString1, vbExclamation, "�T��"
        End If
    End If
    '*****************************************************************************************************
    PB2.Visible = False
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    On Error Resume Next
    Call CloseRecordset(rsSO004)
    Call CloseRecordset(rsSo168)
    Call CloseRecordset(rsTemp)
    Call CloseRecordset(rsTemp2)
    'Set rsSo168 = Nothing
    Exit Sub
ChkErr:
  gcnGi.RollbackTrans
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Call ErrSub(Me.Name, "ImportData3")
End Sub

Private Sub LogError2(ByRef ErrorCounts As Long, ByVal CsCustId, TfnCustId, CsNumber, TfnNumber, CsId, TfnId, CsName, TfnName As String, ByVal LineNumber As Long)
  On Error GoTo ChkErr
    '�O�����~����ƨ� LOG �ɤ�
    Dim HasError As Boolean
    Dim strString2 As String
    Dim strString1 As String
  
    HasError = False
  
     '�渹
     strString1 = Right("00000" & LineNumber, 5) & "�G"
  
    '�Ƚs
    If (CsCustId <> "") And (TfnCustId <> "") Then
        strString1 = strString1 & "�Ƚs (" & CsCustId & " / " & TfnCustId & ")"
        HasError = True
    End If
  
    'CP ����
    If (CsNumber <> "") And (TfnNumber <> "") Then
        If HasError Then
            strString1 = strString1 & ", "
        End If
        strString1 = strString1 & "CP ���� (" & CsNumber & " / " & TfnNumber & ")"
        HasError = True
    End If
  
    '�����Ҧr��
    If (CsId <> "") And (TfnId <> "") Then
        If HasError Then
            strString1 = strString1 & ", "
        End If
        strString1 = strString1 & "�����Ҧr�� (" & CsId & " / " & TfnId & ")"
        HasError = True
    End If
  
    '�m�W
    If (CsName <> "") And (TfnName <> "") Then
        If HasError Then
            strString1 = strString1 & ", "
        End If
        strString1 = strString1 & "�m�W (" & CsName & " / " & TfnName & ")"
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
        MsgBox "�����鶷���� 1��31 �����C", vbExclamation, "ĵ�i"
        txtShouldDay.SetFocus
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtShouldDay_LostFocus")
End Sub

Private Sub LogError22(ByRef ErrorCounts As Long, ByVal TFNServiceID, CustID, Number, Reason, TfnValue, CsValue As String, SNO, InstDate, PRDate As String, ByVal LineNumber As Long)
  On Error GoTo ChkErr
    '�O�����~����ƨ� LOG �ɤ�
    Dim strString2 As String
    Dim strString1 As String
    
    '�渹
    strString1 = LineNumber & Space(10)
    strString1 = StrConv(LeftB(StrConv(strString1, vbFromUnicode), 10), vbUnicode)
    
      
    '�Ƚs
    strString2 = CustID & Space(10)
    strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 10), vbUnicode)
    strString1 = strString1 & strString2
      
    'CP ����
    strString2 = Number & Space(20)
    strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 20), vbUnicode)
    strString1 = strString1 & strString2
      
    '�T��
    strString2 = Reason & Space(50)
    strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 50), vbUnicode)
    strString1 = strString1 & strString2
      
    'TFN ��
    If TfnValue = "" Then
        TfnValue = "�L"
    End If
    strString2 = TfnValue & Space(20)
    strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 20), vbUnicode)
    strString1 = strString1 & strString2
      
    'CS ��
    If CsValue = "" Then
        CsValue = "�L"
    End If
    strString2 = CsValue & Space(20)
    strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 20), vbUnicode)
    strString1 = strString1 & strString2
      
    'SNO ��
    If SNO = "" Then
        SNO = "�L"
    End If
    strString2 = SNO & Space(20)
    strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 20), vbUnicode)
    strString1 = strString1 & strString2
      
    'InstDate ��
    If InstDate = "" Then
        InstDate = "�L"
    End If
    strString2 = InstDate & Space(20)
    strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 20), vbUnicode)
    strString1 = strString1 & strString2
      
    'PRDate ��
    If PRDate = "" Then
        PRDate = "�L"
    End If
    strString2 = PRDate & Space(20)
    strString2 = StrConv(LeftB(StrConv(strString2, vbFromUnicode), 20), vbUnicode)
    strString1 = strString1 & strString2
      
    
    strmTextFile1.Write strString1 & vbCrLf
    '  ErrorCounts = ErrorCounts + 1
    ErrorType = True
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "LogError22")
End Sub
'#3195�N�n�s�bSO19J0B.Ini����T���إߤ@��Collect���X
'Collect.Key=Section+Key+(����+[���󪺭�]=KeyValue)
'Collect.Item=���󪺹w�]��
Sub ReadIniPages()
  On Error Resume Next
    With iniPages
        .Add "PAGE0201," & txtImportFile1.Name, ""
        .Add "PAGE0201," & txtLogFile1.Name, ""
        .Add "PAGE0202," & txtbillYM.Name, ""
        .Add "PAGE0202," & txtNumCount.Name, ""
        .Add "PAGE0202," & txtAreaCode.Name, ""
        .Add "PAGE0202," & txtImportFile3.Name, "CNS_RATING_D"
        .Add "PAGE0202," & txtImportFile4.Name, "CNS_RATING_M"
        .Add "PAGE0202," & txtImportFile5.Name, ""
        .Add "PAGE0202," & txtLogFile6.Name, ""
        .Add "PAGE0203," & txtImportFile2.Name, ""
        .Add "PAGE0203," & txtLogFile2.Name, ""
        .Add "PAGE0203," & txtfilter.Name, ""
        .Add "PAGE0301," & txtExportFile1.Name, gcnGi.Execute("Select SOCompCode From " & GetOwner & "CD039 Where CodeNo=" & gCompCode)(0).Value & "_PIH_R"
        .Add "PAGE0301," & txtExportPath1.Name, ""
        .Add "PAGE0301," & txtLogFile4.Name, ""
        .Add "PAGE0302," & txtExportFile2.Name, ""
        .Add "PAGE0302," & txtExportPath2.Name, ""
        .Add "PAGE0302," & txtLogFile5.Name, ""
        .Add "PAGE0303," & gimPrCode1.Name, ""
        .Add "PAGE0303," & gimReasonCode.Name, ""
        .Add "PAGE0303," & txtExportFile3.Name, ""
        .Add "PAGE0303," & txtExportPath3.Name, ""
        .Add "PAGE0304," & gimPrCode2.Name, ""
        .Add "PAGE0304," & gimUCCode.Name, ""
        .Add "PAGE0304," & txtExportFile4.Name, ""
        .Add "PAGE0304," & txtExportPath4.Name, ""
        .Add "PAGE0305," & txtExportFile5.Name, ""
        .Add "PAGE0305," & txtExportPath5.Name, ""
        .Add "PAGE0305," & txtLogFile6.Name, ""
    End With
    If Right(strErrPath, 1) = "\" Then
        strIniFile = strErrPath & "SO19J0B.Ini"
    Else
        strIniFile = strErrPath & "\SO19J0B.Ini"
    End If
End Sub
