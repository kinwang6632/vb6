VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO5280A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "���O����Ȥ�Ʋέp��[ SO5280A ]"
   ClientHeight    =   7005
   ClientLeft      =   1995
   ClientTop       =   1515
   ClientWidth     =   9975
   Icon            =   "SO5280A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form52"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.PictureBox pic2 
      Height          =   3075
      Left            =   90
      ScaleHeight     =   3015
      ScaleWidth      =   9675
      TabIndex        =   54
      Top             =   2040
      Width           =   9735
      Begin VB.VScrollBar vsl1 
         Height          =   2985
         LargeChange     =   100
         Left            =   9360
         Max             =   100
         SmallChange     =   100
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.Frame fraMulti 
         BorderStyle     =   0  '�S���ؽu
         Height          =   4695
         Left            =   30
         TabIndex        =   55
         Top             =   0
         Width           =   9285
         Begin CS_Multi.CSmulti gimCitemCode 
            Height          =   345
            Left            =   0
            TabIndex        =   57
            Top             =   330
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   609
            ButtonCaption   =   "��  �O  ��  ��"
         End
         Begin CS_Multi.CSmulti gimClassCode 
            Height          =   345
            Left            =   0
            TabIndex        =   58
            Top             =   660
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   609
            ButtonCaption   =   "��  ��  ��  �O"
         End
         Begin Gi_Multi.GiMulti gimServiceType 
            Height          =   345
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   609
            ButtonCaption   =   "�A  ��  ��  �O"
            DataType        =   2
         End
         Begin CS_Multi.CSmulti gimStrtCode 
            Height          =   345
            Left            =   0
            TabIndex        =   60
            Top             =   2310
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   609
            ButtonCaption   =   "��  �D  �d  ��"
         End
         Begin Gi_Multi.GiMulti gimAreaCode 
            Height          =   345
            Left            =   0
            TabIndex        =   61
            Top             =   1650
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   609
            ButtonCaption   =   "��     �F     ��"
         End
         Begin Gi_Multi.GiMulti gimServCode 
            Height          =   345
            Left            =   0
            TabIndex        =   62
            Top             =   1320
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   609
            ButtonCaption   =   "�A     ��     ��"
         End
         Begin Gi_Multi.GiMulti gimStatusCode 
            Height          =   345
            Left            =   0
            TabIndex        =   63
            Top             =   990
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   609
            ButtonCaption   =   "��  ��  ��  �A"
         End
         Begin Gi_Multi.GiMulti gimCMCode 
            Height          =   345
            Left            =   0
            TabIndex        =   64
            Top             =   1980
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   609
            ButtonCaption   =   "��  �O  ��  ��"
            ColumnOrder     =   -1  'True
         End
         Begin CS_Multi.CSmulti gimMduId 
            Height          =   345
            Left            =   0
            TabIndex        =   65
            Top             =   2640
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   609
            ButtonCaption   =   "�j  ��  �s  ��"
         End
         Begin Gi_Multi.GiMulti gimBankCode 
            Height          =   345
            Left            =   0
            TabIndex        =   66
            Top             =   2970
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   609
            ButtonCaption   =   "��  ��  ��  �O"
         End
         Begin Gi_Multi.GiMulti gimCardCode 
            Height          =   345
            Left            =   0
            TabIndex        =   67
            Top             =   3300
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   609
            ButtonCaption   =   "�H �� �d �� ��"
         End
         Begin CS_Multi.CSmulti gimClctEn 
            Height          =   375
            Left            =   0
            TabIndex        =   68
            Top             =   3630
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   661
            ButtonCaption   =   "��  �O  �H   ��"
         End
         Begin Gi_Multi.GiMulti gimOrder 
            Height          =   345
            Left            =   0
            TabIndex        =   69
            Top             =   3960
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   609
            ButtonCaption   =   "��  ��  ��  ��"
            DataType        =   2
            ColumnOrder     =   -1  'True
         End
         Begin Gi_Multi.GiMulti gimPayType 
            Height          =   390
            Left            =   0
            TabIndex        =   70
            Top             =   4290
            Width           =   10035
            _ExtentX        =   17701
            _ExtentY        =   688
            ButtonCaption   =   "ú  �I  ��  �O"
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�g���ʦ��O"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   8640
      TabIndex        =   53
      Top             =   960
      Width           =   1170
      Begin VB.OptionButton optStopFlag 
         Caption         =   "�ҥ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton optStopFlag 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   945
      End
      Begin VB.OptionButton optStopFlag 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�a�}�̾�"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1035
      Left            =   7320
      TabIndex        =   52
      Top             =   990
      Width           =   1215
      Begin VB.OptionButton optAddress 
         Caption         =   "�l�H�a�}"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   2
         Left            =   90
         TabIndex        =   24
         Top             =   750
         Width           =   1095
      End
      Begin VB.OptionButton optAddress 
         Caption         =   "���O�a�}"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   23
         Top             =   495
         Width           =   1095
      End
      Begin VB.OptionButton optAddress 
         Caption         =   "�˾��a�}"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame frame 
      Caption         =   "�Ȥ���w��ƺ���"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   90
      TabIndex        =   51
      Top             =   1410
      Width           =   3690
      Begin VB.OptionButton OptAllot 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1725
         TabIndex        =   9
         Top             =   270
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton OptYesAllot 
         Caption         =   "�ȫ��w"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2700
         TabIndex        =   10
         Top             =   270
         Width           =   945
      End
      Begin VB.OptionButton OptNoAllot 
         Caption         =   "�����w"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   555
         TabIndex        =   8
         Top             =   270
         Width           =   945
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "�����s��"
      Height          =   945
      Left            =   3840
      TabIndex        =   50
      Top             =   1080
      Width           =   3375
      Begin VB.OptionButton optCircuitNo 
         Caption         =   "�u�L�ťպ����s��"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   150
         TabIndex        =   20
         Top             =   660
         Width           =   1845
      End
      Begin VB.OptionButton optAll 
         Caption         =   "����"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   2490
         TabIndex        =   21
         Top             =   660
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.TextBox mskCircuitNo 
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   150
         TabIndex        =   19
         Top             =   240
         Width           =   3045
      End
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "�צ�Excel"
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
      Height          =   525
      Left            =   3720
      TabIndex        =   38
      Top             =   6330
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "�W���έp���G"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1860
      TabIndex        =   37
      Top             =   6330
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "����(&X)"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8460
      TabIndex        =   39
      Top             =   6330
      Width           =   1395
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:�C�L"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   180
      TabIndex        =   36
      Top             =   6330
      Width           =   1245
   End
   Begin VB.Frame fraPage 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�����覡"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3840
      TabIndex        =   47
      Top             =   420
      Width           =   3375
      Begin VB.OptionButton optNothing 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�L"
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2760
         TabIndex        =   15
         Top             =   255
         Value           =   -1  'True
         Width           =   525
      End
      Begin VB.OptionButton optAreaCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��F��"
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   990
         TabIndex        =   13
         Top             =   255
         Width           =   885
      End
      Begin VB.OptionButton optServCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�A�Ȱ�"
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   255
         Width           =   915
      End
      Begin VB.OptionButton optClctAreaCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "���O��"
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1860
         TabIndex        =   14
         Top             =   255
         Width           =   885
      End
   End
   Begin VB.ComboBox cmbAmoun 
      Height          =   300
      ItemData        =   "SO5280A.frx":0442
      Left            =   2055
      List            =   "SO5280A.frx":0455
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   6
      Top             =   1020
      Width           =   585
   End
   Begin VB.ComboBox cmbPeriod 
      Height          =   300
      ItemData        =   "SO5280A.frx":046A
      Left            =   585
      List            =   "SO5280A.frx":047D
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   4
      Top             =   1020
      Width           =   585
   End
   Begin VB.Frame fraReportType 
      BackColor       =   &H00E0E0E0&
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   915
      Left            =   60
      TabIndex        =   46
      Top             =   5220
      Width           =   9825
      Begin VB.OptionButton optSummaric7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�U���~��+���O���طJ�`+����+���B�J�`"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   5835
         TabIndex        =   35
         Top             =   600
         Width           =   3825
      End
      Begin VB.OptionButton optSummaric6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�U���~��+���O��+��D�J�`"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   2955
         TabIndex        =   34
         Top             =   600
         Width           =   2955
      End
      Begin VB.OptionButton optSummaric5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�U���~��+���O���طJ�`"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   420
         TabIndex        =   33
         Top             =   600
         Width           =   2805
      End
      Begin VB.OptionButton optSummaric4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�U���~��+��D�J�`"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   6405
         TabIndex        =   31
         Top             =   270
         Width           =   2025
      End
      Begin VB.OptionButton optSummaric1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�U���~��+����+���B�J�`"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   420
         TabIndex        =   28
         Top             =   270
         Value           =   -1  'True
         Width           =   2475
      End
      Begin VB.OptionButton optDetail 
         BackColor       =   &H00E0E0E0&
         Caption         =   "���Ӫ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   8685
         TabIndex        =   32
         Top             =   270
         Width           =   915
      End
      Begin VB.OptionButton optSummaric2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����+���B�J�`"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   2955
         TabIndex        =   29
         Top             =   270
         Width           =   1575
      End
      Begin VB.OptionButton optSummaric3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�U���~��J�`"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   4740
         TabIndex        =   30
         Top             =   270
         Width           =   1485
      End
   End
   Begin VB.CheckBox chkXCitem 
      Caption         =   "�O�_�έp�K���O�Ȥ�"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7350
      TabIndex        =   18
      Top             =   690
      Width           =   2265
   End
   Begin VB.CheckBox chkClassCode 
      Caption         =   "�O�_���Ȥ����O"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   7350
      TabIndex        =   17
      Top             =   390
      Width           =   2265
   End
   Begin VB.CheckBox chkInstDate 
      Caption         =   "�M�䦸������O�_�ť�"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   7350
      TabIndex        =   16
      Top             =   60
      Width           =   2265
   End
   Begin VB.TextBox txtAmoun 
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2655
      MaxLength       =   6
      TabIndex        =   7
      Top             =   990
      Width           =   1125
   End
   Begin VB.TextBox txtPeriod 
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1185
      MaxLength       =   2
      TabIndex        =   5
      Top             =   990
      Width           =   405
   End
   Begin Gi_Date.GiDate gdaInstDate2 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   60
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ForeColor       =   12582912
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
   Begin Gi_Date.GiDate gdaInstDate1 
      Height          =   375
      Left            =   1290
      TabIndex        =   0
      Top             =   60
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ForeColor       =   12582912
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
   Begin Gi_Date.GiDate gdaClctDate2 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   540
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ForeColor       =   12582912
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
   Begin Gi_Date.GiDate gdaClctDate1 
      Height          =   375
      Left            =   1290
      TabIndex        =   2
      Top             =   540
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ForeColor       =   12582912
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
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   4680
      TabIndex        =   11
      Top             =   60
      Width           =   2535
      _ExtentX        =   4471
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
      FldWidth2       =   1600
      F5Corresponding =   -1  'True
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�U�����O"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   90
      TabIndex        =   49
      Top             =   540
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
      Caption         =   "��  �q  �O"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3810
      TabIndex        =   48
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "���B"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1635
      TabIndex        =   45
      Top             =   1080
      Width           =   390
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   135
      TabIndex        =   44
      Top             =   1080
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "����d��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   90
      TabIndex        =   43
      Top             =   750
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "---"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2430
      TabIndex        =   42
      Top             =   630
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�˾�����d��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   90
      TabIndex        =   41
      Top             =   150
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "---"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2430
      TabIndex        =   40
      Top             =   150
      Width           =   180
   End
End
Attribute VB_Name = "frmSO5280A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Table:SO003,SO001(,SO014,SO002)
Option Explicit
Dim strOrder As String
Dim strNo As String
Dim blnExcel As Boolean
Dim strTable As String
Dim strUseIndex As String
Dim strSO138Table As String
Dim strSO138Where As String
Private Sub cmdExcel_Click()
  On Error GoTo ChkErr
    blnExcel = True
    cmdPrint_Click
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExcel_Click")
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      Select Case True
             Case optSummaric1.Value
                  If chkClassCode.Value = 1 Then
                      Call PreviousRpt(GetPrinterName(5), RptName("SO5280", "2"), "���O����Ȥ�Ʋέp�J�`�� [SO5280A]")
                  Else
                      Call PreviousRpt(GetPrinterName(5), RptName("SO5280", "1"), "���O����Ȥ�Ʋέp�J�`�� [SO5280A]")
                  End If
             Case optSummaric2.Value
                  Call PreviousRpt(GetPrinterName(5), RptName("SO5280", "3"), "���O����Ȥ�Ʋέp�J�`�� [SO5280A]")
             Case optSummaric3.Value
                  Call PreviousRpt(GetPrinterName(5), RptName("SO5280", "4"), "���O����Ȥ�Ʋέp�J�`�� [SO5280A]")
             Case optSummaric4.Value
                  Call PreviousRpt(GetPrinterName(5), RptName("SO5280", "6"), "���O����Ȥ�Ʋέp�J�`�� [SO5280A]")
             Case optSummaric5.Value
                  Call PreviousRpt(GetPrinterName(5), RptName("SO5280", "7"), "���O����Ȥ�Ʋέp�J�`�� [SO5280A]")
             Case optSummaric6.Value
                  Call PreviousRpt(GetPrinterName(5), RptName("SO5280", "8"), "���O����Ȥ�Ʋέp�J�`�� [SO5280A]")
             Case optSummaric7.Value
                  Call PreviousRpt(GetPrinterName(5), RptName("SO5280", "9"), "���O����Ȥ�Ʋέp�J�`�� [SO5280A]")
             Case optDetail.Value
                  Call PreviousRpt(GetPrinterName(5), RptName("SO5280", "5"), "���O����Ȥ�Ʋέp���Ӫ� [SO5280A]")
      End Select
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo ChkErr
    If Not ChkDTok Then Exit Sub
    Screen.MousePointer = vbHourglass
      cmdExit.SetFocus
      Me.Enabled = False
        ReadyGoPrint
        Call subChoose
        Call subPrint
      Me.Enabled = True
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Sub subChoose()
  On Error GoTo ChkErr
  Dim strpagetype As String
  Dim strGetGroupName As String
  Dim arrOrder() As String
  Dim varOrder As Variant
  Dim intSort As Integer
  Dim blnSO014 As Boolean
  Dim blnSO002 As Boolean
  Dim blnSO002A As Boolean
  Dim strAddrField As String
  Dim strAddrShow As String
  Dim strAllotShow As String
  Dim strStopFlag As String
    strChoose = ""
    strChooseString = ""
    strGroupName = ""
    strOrder = ""
    strGetGroupName = ""
    strUseIndex = ""
    blnSO014 = False
    blnSO002 = False
    blnSO002A = False
    
  '���
    If chkInstDate.Value = 1 Then
        Call subAnd("SO002.InstTime is null"): blnSO002 = True
    Else
        If gdaInstDate1.GetValue <> "" Then Call subAnd("SO002.InstTime >= To_Date('" & gdaInstDate1.GetValue & "','YYYYMMDD')"): blnSO002 = True
        If gdaInstDate2.GetValue <> "" Then Call subAnd("SO002.InstTime < To_Date('" & gdaInstDate2.GetValue & "','YYYYMMDD')+1"): blnSO002 = True
        If gdaInstDate1.GetValue <> "" Or gdaInstDate2.GetValue <> "" Then strUseIndex = GetUseIndexStr("SO002", "InstTime")
    End If
    If gdaClctDate1.GetValue <> "" And gdaClctDate2.GetValue <> "" Then
        Call subAnd("SO003.ClctDate Between To_Date('" & gdaClctDate1.GetValue & "','YYYYMMDD') And To_Date('" & gdaClctDate2.GetValue & "','YYYYMMDD')")
    Else
        If gdaClctDate1.GetValue <> "" Then Call subAnd("SO003.ClctDate >= To_Date('" & gdaClctDate1.GetValue & "','YYYYMMDD')")
        If gdaClctDate2.GetValue <> "" Then Call subAnd("SO003.ClctDate < To_Date('" & gdaClctDate2.GetValue & "','YYYYMMDD')+1")
    End If
    If gdaClctDate1.GetValue <> "" Or gdaClctDate2.GetValue <> "" Then strUseIndex = strUseIndex & " " & GetUseIndexStr("SO003", "ClctDate")
    
  '����
    If txtPeriod <> "" Then Call subAnd("SO003.Period " & cmbPeriod & txtPeriod)
  '���B
    If txtAmoun <> "" Then Call subAnd("SO003.Amount" & cmbAmoun & txtAmoun)
  '�M��U�����O����O�_�ť�
    If chkInstDate.Value = 1 Then subAnd ("SO003.ClctDate is null")
     
  'GiMulti
    If gimServiceType.GetQryStr <> "" Then Call subAnd("SO003.ServiceType " & gimServiceType.GetQryStr)
    If gimCitemCode.GetQueryCode <> "" Then Call subAnd("SO003.CitemCode IN (" & gimCitemCode.GetQueryCode & ")")
    'If gimCompCode.GetQryStr <> "" Then Call subAnd("SO001.CompCode" & gimCompCode.GetQryStr)
    If gimClassCode.GetQryStr <> "" Then Call subAnd("SO001.ClassCode1 " & gimClassCode.GetQryStr)
    If gimCMCode.GetQryStr <> "" Then Call subAnd("SO003.CMCode " & gimCMCode.GetQryStr)
    If gimStatusCode.GetQryStr <> "" Then Call subAnd("SO002.CustStatusCode " & gimStatusCode.GetQryStr): blnSO002 = True
    If gimServCode.GetQryStr <> "" Then Call subAnd("SO014.ServCode " & gimServCode.GetQryStr): blnSO014 = True
    If gimAreaCode.GetQryStr <> "" Then Call subAnd("SO014.AreaCode " & gimAreaCode.GetQryStr): blnSO014 = True
    If gimStrtCode.GetQryStr <> "" Then Call subAnd("SO014.StrtCode " & gimStrtCode.GetQryStr): blnSO014 = True
    '#5267 �쥻�O��SO002A,���SO106 By Kin 2010/04/22
    If gimCardCode.GetQryStr <> "" Then Call subAnd("SO106.CardName " & gimCardCode.GetDescStr): blnSO002A = True
    If gimClctEn.GetQryStr <> "" Then Call subAnd("SO014.ClctEn " & gimClctEn.GetQryStr): blnSO014 = True
    '#5683 �W�[ú�I���O By Kin 2010/08/16
    If gimPayType.GetQryStr & "" <> "" Then
        Call subAnd("SO003.PayKind " & gimPayType.GetQryStr)
    End If
    
  '�Ȧ����O
    If gimBankCode.GetQryStr <> "" Then
        Call subAnd("SO003.BankCode " & gimBankCode.GetQryStr)
    Else
       If gimBankCode.GetDispStr <> "" Then Call subAnd("SO003.BankCode is not null")
    End If
    
  '�j�ӽs��
    If gimMduId.GetQryStr <> "" Then
        Call subAnd("SO014.MduId " & gimMduId.GetQryStr)
        blnSO014 = True
    Else
        If gimMduId.GetDispStr <> "" Then subAnd ("SO014.MduId is Not Null"): blnSO014 = True
    End If
    
  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO003.CompCode=" & gilCompCode.GetCodeNo)
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO003.ServiceType='" & gilServiceType.GetCodeNo & "'")

    '�����s��
    If mskCircuitNo.Text <> "" Then
        Call subAnd("SO014.CircuitNo='" & mskCircuitNo.Text & "'")
        blnSO014 = True
    Else
       If optCircuitNo.Value Then Call subAnd("SO014.CircuitNo is null "): blnSO014 = True
    End If
    '�Ȥ���w��ƺ���
    strAllotShow = OptAllot.Caption
    If OptNoAllot.Value Then Call subAnd("SO003.CustAllot=0"): strAllotShow = OptNoAllot.Caption
    If OptYesAllot.Value Then Call subAnd("SO003.CustAllot=1"): strAllotShow = OptYesAllot.Caption
    
    '�g���ʦ��O
    If optStopFlag(3).Value Then
        Call subAnd("SO003.StopFlag=0")
        strStopFlag = "�ҥ�"
    ElseIf optStopFlag(0).Value Then
        Call subAnd("SO003.StopFlag=1")
        strStopFlag = "����"
    Else
        strStopFlag = "����"
    End If
  '���Ӫ�����覡
    If optDetail Then
        Select Case True
               Case optAreaCode.Value
                    strGroupName = "ReportType=True;GroupName={V.AreaName};GroupCode={V.AreaCode}"
                    strpagetype = "��F��"
               Case optServCode.Value
                    strGroupName = "ReportType=True;GroupName={V.ServArea};GroupCode={V.ServCode}"
                    strpagetype = "�A�Ȱ�"
               Case optClctAreaCode.Value
                    strGroupName = "ReportType=True;GroupName={V.ClctAreaName};GroupCode={V.ClctAreaCode}"
                    strpagetype = "���O��"
               Case optNothing.Value
                    Select Case Left(gimOrder.GetColumnOrderCode, 1)
                           Case "A"
                                strGroupName = "ReportType=False;GroupCode={V.Custid};GroupName=''"
                           Case "B"
                                strGroupName = "ReportType=False;GroupCode={V.AddrSort};GroupName=''"
                           Case "C"
                                strGroupName = "ReportType=False;GroupCode={V.InstTime};GroupName=''"
                           Case "D"
                                strGroupName = "ReportType=False;GroupCode={V.ClctDate};GroupName=''"
                           Case "E"
                                strGroupName = "ReportType=False;GroupCode={V.ClctEn};GroupName=''"
                           Case Else
                                strGroupName = "ReportType=False;GroupCode={V.Custid};GroupName=''"
                    End Select
                    strpagetype = "�L"
         End Select
    End If

  '�ƧǤ覡
    arrOrder = Split(gimOrder.GetColumnOrderCode, ",")
    intSort = 0
    For Each varOrder In arrOrder
        Select Case arrOrder(intSort)
               Case "A"
                    strGroupName = strGroupName & ";Sort" & intSort & "={V.Custid}"
               Case "B"
                    strGroupName = strGroupName & ";Sort" & intSort & "={V.AddrSort}"
               Case "C"
                    strGroupName = strGroupName & ";Sort" & intSort & "=Date({V.InstTime})"
               Case "D"
                    strGroupName = strGroupName & ";Sort" & intSort & "=Date({V.ClctDate})"
               Case "E"
                    strGroupName = strGroupName & ";Sort" & intSort & "={V.ClctEn}"
        End Select
        intSort = intSort + 1
    Next
    
  '�O�_�έp�K���O�Ȥ�
    Select Case True
        Case optAddress(0)
            strAddrField = "InstAddrNo"
            strAddrShow = optAddress(0).Caption
        Case optAddress(1)
            strAddrField = "ChargeAddrNo"
            strAddrShow = optAddress(1).Caption
        Case optAddress(2)
            strAddrField = "MailAddrNo"
            strAddrShow = optAddress(2).Caption
    End Select
    
    If chkXCitem.Value = 1 Then
        If optDetail.Value Or blnSO002 Then
            strTable = " From SO003,SO001,SO002"
            strChoose = "SO003.CustId(+)=SO001.CustId And SO003.Compcode(+)=SO001.Compcode And SO003.Custid=SO002.Custid(+) And SO003.ServiceType=SO002.ServiceType(+) And SO003.Compcode=SO002.Compcode(+) And " & strChoose
        Else
            strTable = " From SO003,SO001"
            strChoose = "SO003.CustId(+)=SO001.CustId And SO003.Compcode(+)=SO001.Compcode And " & strChoose
        End If
        If blnSO014 Or optDetail.Value Or optSummaric4.Value Or optSummaric6.Value Then
            strTable = strTable & ",SO014 "
'            If optAddress(1).Value Or optAddress(2).Value Then
'                strSO138Table = strTable & ",SO138"
'                If optAddress(1).Value Then
'                    strSO138Table = " WHERE SO014.ADDRNO=SO138." & strAddrField & " And SO014.Compcode(+)=SO001.Compcode And " & strChoose
'                End If
'            End If
            strChoose = " Where SO014.AddrNo(+)=SO001." & strAddrField & " And SO014.Compcode(+)=SO001.Compcode And " & strChoose
            
        Else
            strTable = strTable & " "
            strChoose = " Where " & strChoose
        End If
    Else
        If optDetail.Value Or blnSO002 Then
            strTable = " From SO003,SO001,SO002"
            strChoose = "SO003.CustId=SO001.CustId And SO003.Compcode=SO001.Compcode And SO003.Custid=SO002.Custid And SO003.ServiceType=SO002.ServiceType And SO003.Compcode=SO002.Compcode And " & strChoose
        Else
            strTable = " From SO003,SO001"
            strChoose = "SO003.CustId=SO001.CustId And SO003.Compcode=SO001.Compcode And " & strChoose
        End If
        If blnSO014 Or optDetail.Value Or optSummaric4.Value Or optSummaric6.Value Then
            strTable = strTable & ",SO014 "
            strChoose = " Where SO014.AddrNo=SO001." & strAddrField & " And SO014.Compcode=SO001.Compcode And " & strChoose
        Else
            strTable = strTable & " "
            strChoose = " Where " & strChoose
        End If
    End If
    '#5267 �쥻��SO002A,�{�b���SO106 By Kin 2010/04/22
    If blnSO002A Then
        strTable = strTable & ",SO106 "
        strChoose = strChoose & " And SO003.Custid=SO106.Custid And SO003.AccountNo=SO106.AccountID" & _
                " AND SO106.StopFlag <> 1 AND SO106.SnactionDate is Not Null "
    End If
    '#5627 ���դ�OK,�p�G�O�l�H�a�}���O�a�}�n��SO138���a�} By Kin 2010/05/24
    If Not optAddress(0) Then
        strSO138Table = strTable & ",SO138 "
        strSO138Where = Replace(strChoose, "SO001." & strAddrField, "SO138." & strAddrField)
        strSO138Where = strSO138Where & " AND SO003.INVSEQNO IS NOT NULL  AND SO003.INVSEQNO=SO138.INVSEQNO "
        strChoose = strChoose & " AND SO003.INVSEQNO IS NULL "
        
    End If
    
    strChooseString = "�˾����: " & subSpace(gdaInstDate1.GetValue(True)) & "~" & subSpace(gdaInstDate2.GetValue(True)) & ";" & _
                      "�U�����O���: " & subSpace(gdaClctDate1.GetValue(True)) & "~" & subSpace(gdaClctDate2.GetValue(True)) & ";" & _
                      "���q�O�@: " & subSpace(gilCompCode.GetDescription) & ";" & _
                      "�A�����O: " & subSpace(gimServiceType.GetDispStr) & ";" & _
                      "���O�H��: " & subSpace(gimClctEn.GetDispStr) & ";" & _
                      "����" & cmbPeriod.Text & txtPeriod & " ���B" & cmbAmoun.Text & txtAmoun & ";" & _
                      "���O����: " & subSpace(gimCitemCode.GetDispStr) & ";" & _
                      "�Ȥ����O: " & subSpace(gimClassCode.GetDispStr) & ";�g���ʦ��O: " & strStopFlag & ";" & _
                      "�Ȥ᪬�A: " & subSpace(gimStatusCode.GetDispStr) & ";" & _
                      "�A�Ȱϡ@: " & subSpace(gimServCode.GetDispStr) & ";" & _
                      "��F�ϡ@: " & subSpace(gimAreaCode.GetDispStr) & ";" & _
                      "���O�覡: " & subSpace(gimCMCode.GetDispStr) & ";" & _
                      "��D�d��: " & subSpace(gimStrtCode.GetDispStr) & ";" & _
                      "�j�ӽs��: " & subSpace(gimMduId.GetDispStr) & ";" & _
                      "�Ȧ����O: " & subSpace(gimBankCode.GetDispStr) & ";" & _
                      "�H�Υd����:" & subSpace(gimCardCode.GetDispStr) & ";" & _
                      "�ƧǤ覡: " & subSpace(gimOrder.GetColumnOrderDspStr) & ";" & _
                      "�����s��: " & subSpace(mskCircuitNo.Text) & ";" & _
                      "ú�I���O: " & subSpace(gimPayType.GetDispStr) & ";" & _
                      "�a�}�̾�: " & subSpace(strAddrShow) & "; �Ȥ���w��ƺ���:" & strAllotShow & ";" & _
                      "�����覡: " & strpagetype
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub subPrint()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    If optDetail.Value Then '���Ӫ�
        subCreateView
        Call GetRS(rsTmp, "Select Count(*) as intCount From " & strViewName)
        If rsTmp("intCount") = 0 Then
            MsgNoRcd
            SendSQL , , True
        Else
            strsql = "Select * From " & strViewName & " V"
            If blnExcel Then
                Call toExcel(strsql)
            Else
                Call PrintRpt2(GetPrinterName(5), RptName("SO5280", "5"), , "���O����Ȥ�Ʋέp���Ӫ� [SO5280A]", strsql, strChooseString, , True, , , strGroupName, GiPaperLandscape)
            End If
        End If
    Else
        Select Case True
               Case optSummaric1.Value '�J�`��@
                '�O�_���Ȥ����O
                    If chkClassCode.Value = 1 Then
                        strsql = "Select " & strUseIndex & " To_Char(SO003.ClctDate,'YYYY/MM') as ClctDate,SO003.Period,SO003.Amount,SO001.ClassName1,Count(Distinct SO001.custid) as cunCust,Count(*) as intCount " & strTable & strChoose & _
                                 " Group by To_Char(ClctDate,'YYYY/MM'),SO003.Period,SO003.Amount,SO001.ClassName1 "
                        strNo = 2
                        Call subInsertMDB
                    Else
                        strsql = "Select " & strUseIndex & " To_Char(SO003.ClctDate,'YYYY/MM') as ClctDate,SO003.Period,SO003.Amount,1 as ClassName1" & _
                                 ",Count(Distinct SO001.custid) as cunCust,Count(*) as intCount " & strTable & strChoose & _
                                 " Group by To_Char(ClctDate,'YYYY/MM'),SO003.Period,SO003.Amount "
                        strNo = 1
                        Call subInsertMDB
                    End If
               Case optSummaric2.Value '�J�`��G
                    strsql = "Select " & strUseIndex & " SO003.Period,SO003.Amount,'2000/01/01' as ClctDate,1 as ClassName1,Count(Distinct SO001.custid) as cunCust,Count(*) as intCount" & _
                             strTable & strChoose & _
                             " Group by SO003.Period,SO003.Amount"
                    strNo = 3
                    Call subInsertMDB
               Case optSummaric3.Value '�J�`��T
                    strsql = "Select " & strUseIndex & " TO_CHAR(SO003.ClctDate,'YYYY/MM') ClctDate,1 as Period,1 as Amount,1 as ClassName1,Count(Distinct SO001.custid) as cunCust,Count(*) as intCount" & _
                             strTable & strChoose & _
                             " Group by TO_CHAR(SO003.ClctDate,'YYYY/MM')"
                    strNo = 4
                    Call subInsertMDB
               Case optSummaric4.Value '�J�`��|
                    strsql = "Select " & strUseIndex & " SO003.ClctDate,SO014.StrtCode Period,SO014.StrtName ClassName1,SO014.Section Amount,Count(Distinct SO001.custid) as cunCust,Count(*) as intCount" & _
                             strTable & strChoose & _
                             " Group by SO003.ClctDate,SO014.StrtCode,SO014.StrtName,SO014.Section"
                    strNo = 6
                    Call subInsertMDB
               Case optSummaric5.Value '�J�`��
                    strsql = "Select " & strUseIndex & " TO_CHAR(SO003.ClctDate,'YYYY/MM') ClctDate,SO003.CitemCode Period,SO003.CitemName ClassName1,1 as Amount,Count(Distinct SO001.custid) as cunCust,Count(*) as intCount" & _
                             strTable & strChoose & _
                             " Group by TO_CHAR(SO003.ClctDate,'YYYY/MM'),SO003.CitemCode,SO003.CitemName"
                    strNo = 7
                    Call subInsertMDB
               Case optSummaric6.Value '�J�`��
                    strsql = "Select " & strUseIndex & " TO_CHAR(SO003.ClctDate,'YYYY/MM') ClctDate,SO014.StrtCode Period,SO014.StrtName ClassName1,SO014.Section Amount,SO014.ClctEn,SO014.ClctName,Count(Distinct SO001.custid) as cunCust,Count(*) as intCount" & _
                             strTable & strChoose & _
                             " Group by TO_CHAR(SO003.ClctDate,'YYYY/MM'),SO014.StrtCode,SO014.StrtName,SO014.Section,SO014.ClctEn,SO014.ClctName"
                    strNo = 8
                    Call subInsertMDB
               Case optSummaric7.Value '�J�`��C
                    strsql = "Select " & strUseIndex & " To_Char(SO003.ClctDate,'YYYY/MM') as ClctDate,SO003.CitemName,SO003.Period,SO003.Amount,1 as ClassName1" & _
                             ",Count(Distinct SO001.custid) as cunCust,Count(*) as intCount " & strTable & strChoose & _
                             " Group by To_Char(ClctDate,'YYYY/MM'),SO003.CitemName,SO003.Period,SO003.Amount "
                    strNo = 9
                    Call subInsertMDB
        End Select

        Set rsTmp = cnn.Execute("SELECT Count(*) as intCount FROM SO5280A")
        If rsTmp("intCount") = 0 Then
            MsgNoRcd
            SendSQL , , True
        Else
            strsql = "SELECT * FROM SO5280A"
            Call PrintRpt2(GetPrinterName(5), RptName("SO5280", strNo), , "���O����Ȥ�Ʋέp�J�`�� [SO5280A]", strsql, strChooseString, , True, "Tmp0000.Mdb")
        End If
    End If
    CloseRecordset rsTmp
    blnExcel = False
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subPrint")
End Sub

Private Sub toExcel(ByVal strsql As String)
  On Error GoTo ChkErr
    Dim rsExcel As New ADODB.Recordset
    RptToTxt RptName("SO5280", "5"), , strsql, strChooseString, , , strGroupName, , , Environ("Temp") & "\SO5280"
    If Not Get_RS_From_Txt(Environ("Temp"), "SO5280.txt", rsExcel) Then blnExcel = False: Exit Sub
    Call UseProperty(rsExcel, "���O����Ȥ�Ʋέp���Ӫ�", "�Ĥ@��")
    blnExcel = False
    CloseRecordset rsExcel
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "toExcel")
End Sub


Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGim
    Call subGil
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub cmdExit_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "���O���إN�X", "���O���ئW��")
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "�Ȥ����O�N�X", "�Ȥ����O�W��")
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "���O�覡�N�X", "���O�覡�W��")
    Call SetgiMulti(gimStatusCode, "CodeNo", "Description", "CD035", "�Ȥ᪬�A�N�X", "�Ȥ᪬�A�W��")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "�A�ȰϥN�X", "�A�ȰϦW��")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "��F�ϥN�X", "��F�ϦW��")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "��D�N�X", "��D�W��")
    Call SetgiMulti(gimMduId, "MduId", "Name", "SO017", "�j�ӽs��", "�j�ӦW��")
    Call SetgiMulti(gimBankCode, "CodeNo", "Description", "CD018", "�Ȧ�O�N�X", "�Ȧ�O�W��")
    Call SetgiMulti(gimCardCode, "CodeNo", "Description", "CD037", "�H�Υd�����N�X", "�H�Υd�����W��")
    Call SetgiMulti(Me.gimClctEn, "EmpNo", "EmpName", "CM003", "���O�H���N�X", "���O�H���W��")
    Call SetgiMultiAddItem(gimOrder, "A,B,C,D,E", "�Ȥ�s��,�a�},�˾����,�����O���,���O�H��")
    gimStatusCode.SetQueryCode "1"
    '#5683 �W�[ú�I���O By Kin 2010/08/16
    Call SetgiMulti(gimPayType, "CodeNo", "Description", "CD112", "�N�X", "�W��")
    If GetPaynowFlag Then
        With gimPayType
                '.SetDispStr "�w�I��"
            .SetQueryCode "0"
        End With
    Else
        gimPayType.Clear
        gimPayType.Enabled = False
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    'Call SetgiList(gilClctEn, "EmpNo", "EmpName", "CM003")
    'Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub
  
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5280A)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  gcnGi.Execute "Drop View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5280A
End Sub

Private Sub gdaClctDate1_GotFocus()
  On Error Resume Next
    If gdaClctDate1.Text = "" Then gdaClctDate1.SetValue (RightDate)
End Sub

Private Sub gdaClctDate2_GotFocus()
  On Error Resume Next
    If gdaClctDate1.GetValue = "" Or gdaClctDate2.GetValue = "" Then gdaClctDate2.SetValue (gdaClctDate1.GetValue)
End Sub

Private Sub gdaClctDate2_Validate(Cancel As Boolean)
  On Error Resume Next
    Cancel = ChkDate2(gdaClctDate1, gdaClctDate2)
End Sub

Private Sub gdaInstDate1_GotFocus()
 On Error Resume Next
   If gdaInstDate1.GetValue = "" Then gdaInstDate1.SetValue (RightDate)
End Sub

Private Sub gdaInstDate2_GotFocus()
  On Error Resume Next
    If gdaInstDate1.GetValue = "" Or gdaInstDate2.GetValue = "" Then gdaInstDate2.SetValue (gdaInstDate1.GetValue)
End Sub

Private Sub gdaInstDate2_Validate(Cancel As Boolean)
  On Error Resume Next
    Cancel = ChkDate2(gdaInstDate1, gdaInstDate2)
End Sub

Private Sub subInsertMDB()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
  Dim lngLoop As Long
  cnn.Execute "DELETE * FROM SO5280A"
  If rsTmp.State = 1 Then rsTmp.Close
    Set rsTmp = gcnGi.Execute(strsql)
    SendSQL strsql, True
    cnn.BeginTrans
      While Not rsTmp.EOF
          If strNo = 8 Then
          cnn.Execute "INSERT INTO SO5280A (ClctDate,Period,Amount,ClassName1,cunCust,intCount,ClctName,ClctEn) VALUES (" & _
                      GetNullString(rsTmp("ClctDate"), giDateV, giAccessDb) & "," & _
                      GetNullString(rsTmp("Period"), giLongV) & "," & _
                      GetNullString(rsTmp("Amount"), giStringV) & "," & _
                      GetNullString(rsTmp("ClassName1"), giStringV) & "," & _
                      GetNullString(rsTmp("cunCust"), giLongV) & "," & _
                      GetNullString(rsTmp("intCount"), giLongV) & "," & _
                      GetNullString(rsTmp("ClctName"), giStringV) & "," & _
                      GetNullString(rsTmp("ClctEn"), giStringV) & ")"
          ElseIf strNo = 9 Then
          cnn.Execute "INSERT INTO SO5280A (ClctDate,ClctName,Period,Amount,ClassName1,cunCust,intCount) VALUES (" & _
                      GetNullString(rsTmp("ClctDate"), giDateV, giAccessDb) & "," & _
                      GetNullString(rsTmp("CitemName"), giStringV) & "," & _
                      GetNullString(rsTmp("Period"), giLongV) & "," & _
                      GetNullString(rsTmp("Amount"), giStringV) & "," & _
                      GetNullString(rsTmp("ClassName1"), giStringV) & "," & _
                      GetNullString(rsTmp("cunCust"), giLongV) & "," & _
                      GetNullString(rsTmp("intCount"), giLongV) & ")"
          
          Else
          cnn.Execute "INSERT INTO SO5280A (ClctDate,Period,Amount,ClassName1,cunCust,intCount) VALUES (" & _
                      GetNullString(rsTmp("ClctDate"), giDateV, giAccessDb) & "," & _
                      GetNullString(rsTmp("Period"), giLongV) & "," & _
                      GetNullString(rsTmp("Amount"), giStringV) & "," & _
                      GetNullString(rsTmp("ClassName1"), giStringV) & "," & _
                      GetNullString(rsTmp("cunCust"), giLongV) & "," & _
                      GetNullString(rsTmp("intCount"), giLongV) & ")"
                      
          
          End If
          rsTmp.MoveNext
          
          DoEvents
      Wend
    cnn.CommitTrans
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "subInsertMDB")
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gimServiceType)
    'gilServiceType.ListIndex = 1
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimMduId, , gilCompCode.GetCodeNo
    GiMultiFilter gimClctEn, , gilCompCode.GetCodeNo
    GiMultiFilter gimBankCode, , gilCompCode.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Change")
End Sub

Private Sub gilCompCode_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then
        MsgMustBe ("���q�O")
        Cancel = True
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Sub gimServiceType_Change()
  On Error GoTo ChkErr
    Call GiMultiFilter(gimCMCode, gimServiceType.GetQryStr, , True)
    Call GiMultiFilter(gimCitemCode, gimServiceType.GetQryStr, , True)
    gimCitemCode.Filter = gimCitemCode.Filter & IIf(gimCitemCode.Filter = "", " WHERE ", " AND ") & " PeriodFlag = 1"
    Call GiMultiFilter(gimClassCode, gimServiceType.GetQryStr, , True)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gimServiceType_Change")
End Sub


Private Sub mskCircuitNo_Change()
  On Error GoTo ChkErr
    If mskCircuitNo.Text = "" Then
       optCircuitNo.Enabled = True
       optAll.Enabled = True
    Else
       optCircuitNo.Enabled = False
       optAll.Enabled = False
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "mskCircuitNo_Change")
End Sub

Private Sub optDetail_Click()
  On Error GoTo ChkErr
    defEnabled
    cmdExcel.Enabled = True
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "optDetail_Click")
End Sub

Private Sub optSummaric1_Click()
  On Error GoTo ChkErr
    defEnabled
    cmdExcel.Enabled = False
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "optSummaric1_Click")
End Sub

Private Sub optSummaric2_Click()
  On Error GoTo ChkErr
    defEnabled
    cmdExcel.Enabled = False
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "optSummaric2_Click")
End Sub

Private Sub optSummaric3_Click()
  On Error GoTo ChkErr
    defEnabled
    cmdExcel.Enabled = False
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "optSummaric3_Click")
End Sub

Private Sub optSummaric4_Click()
  On Error GoTo ChkErr
    defEnabled
    cmdExcel.Enabled = False
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "optSummaric4_Click")
End Sub

Private Sub optSummaric5_Click()
  On Error GoTo ChkErr
    defEnabled
    cmdExcel.Enabled = False
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "optSummaric5_Click")
End Sub

Private Sub optSummaric6_Click()
  On Error GoTo ChkErr
    defEnabled
    cmdExcel.Enabled = False
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "optSummaric6_Click")
End Sub

Private Sub optSummaric7_Click()
  On Error GoTo ChkErr
    defEnabled
    cmdExcel.Enabled = False
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "optSummaric7_Click")
End Sub
Private Sub defEnabled()
  On Error GoTo ChkErr
    chkClassCode.Enabled = optSummaric1.Value
    optServCode.Enabled = optDetail.Value
    optAreaCode.Enabled = optDetail.Value
    optServCode.Enabled = optDetail.Value
    optClctAreaCode.Enabled = optDetail.Value
    optNothing.Enabled = optDetail.Value
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "defEnabled")
End Sub

Private Sub subCreateView()
  On Error Resume Next
    gcnGi.Execute "Drop View " & strViewName
  On Error GoTo ChkErr
    Dim strViewSql As String
      strViewName = GetTmpViewName
      '�쥻�O��SO002A,�]�����SO106A,�ҥH�n�h�[Distinct By Kin 2010/04/22
      '#5627 ���դ�OK,�p�G�O�l�H�a�}���O�a�}�n��SO138���a�} By Kin 2010/05/24
      '#5683 �W�[ú�I���O By Kin 2010/08/16
      If Not optAddress(0).Value Then
        strSO138Table = " UNION ALL Select Distinct " & strUseIndex & " SO001.CUSTID," & _
                    "SO001.CUSTNAME,SO001.TEL1,SO001.CLASSNAME1,SO002.INSTTIME,SO003.CLCTDATE," & _
                    "SO003.PERIOD," & _
                   "SO003.AMOUNT,SO003.STARTDATE,SO003.STOPDATE,SO014.ServName," & _
                   "SO014.ClctEn,SO014.ClctName,SO014.MDUId,SO014.MDUName," & _
                   "SO001.CHARGEADDRESS,SO001.CustNote,SO014.AddrSort,SO001.ServCode," & _
                   "SO001.ServArea,SO014.AREANAME,SO014.AREACODE," & _
                   "SO001.CLCTAREACODE,SO001.CLCTAREANAME,SO014.Address," & _
                   "so003.CitemName,SO001.TEL2,SO001.TEL3,SO001.CHARGENOTE,Nvl(SO003.PayKind,0) PayKind " & _
                   strSO138Table & strSO138Where
      End If
      
      strViewSql = "Create View " & strViewName & " as (" & _
                   "Select Distinct " & strUseIndex & " SO001.CUSTID,SO001.CUSTNAME," & _
                   "SO001.TEL1,SO001.CLASSNAME1,SO002.INSTTIME,SO003.CLCTDATE,SO003.PERIOD," & _
                   "SO003.AMOUNT,SO003.STARTDATE,SO003.STOPDATE,SO014.ServName," & _
                   "SO014.ClctEn,SO014.ClctName,SO014.MDUId,SO014.MDUName," & _
                   "SO001.CHARGEADDRESS,SO001.CustNote,SO014.AddrSort,SO001.ServCode," & _
                   "SO001.ServArea,SO014.AREANAME,SO014.AREACODE," & _
                   "SO001.CLCTAREACODE,SO001.CLCTAREANAME,SO014.Address,so003.CitemName," & _
                   "SO001.TEL2,SO001.TEL3,SO001.CHARGENOTE,Nvl(SO003.PayKind,0) PayKind " & _
                   strTable & strChoose & strSO138Table & ")"

      gcnGi.Execute strViewSql
      SendSQL strViewSql, True
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Sub


Private Sub vsl1_Change()
  On Error Resume Next
    If vsl1.Value = 0 Then
        fraMulti.Top = 20
    ElseIf vsl1.Value = 100 Then
        fraMulti.Top = -(pic2.Height) + 480
    End If
End Sub
