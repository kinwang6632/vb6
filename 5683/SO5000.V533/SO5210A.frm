VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO5210A 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�C�를ú�O�Ȥ���Ӫ� [SO5210A]"
   ClientHeight    =   7980
   ClientLeft      =   1950
   ClientTop       =   1530
   ClientWidth     =   11955
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO5210A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame fraMediabillNo 
      Caption         =   "�C��渹"
      Height          =   495
      Left            =   9240
      TabIndex        =   85
      Top             =   1380
      Width           =   2535
      Begin VB.OptionButton optMediaAll 
         Caption         =   "����"
         Height          =   225
         Left            =   1410
         TabIndex        =   88
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optMediaNo 
         Caption         =   "�L"
         Height          =   225
         Left            =   780
         TabIndex        =   87
         Top             =   240
         Width           =   585
      End
      Begin VB.OptionButton optMediaYes 
         Caption         =   "��"
         Height          =   225
         Left            =   150
         TabIndex        =   86
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.CheckBox chkMediabillNo 
      Caption         =   "���C��渹"
      Height          =   225
      Left            =   6390
      TabIndex        =   84
      Top             =   1800
      Visible         =   0   'False
      Width           =   2625
   End
   Begin prjNumber.GiNumber ginRealPeriod1 
      Height          =   345
      Left            =   3840
      TabIndex        =   8
      Top             =   1140
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   609
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
   End
   Begin Gi_YM.GiYM gymClctYM1 
      Height          =   345
      Left            =   990
      TabIndex        =   6
      Top             =   1140
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
   Begin VB.ComboBox cboAddress 
      Height          =   315
      ItemData        =   "SO5210A.frx":0442
      Left            =   6270
      List            =   "SO5210A.frx":044F
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   18
      Top             =   1140
      Width           =   2955
   End
   Begin VB.CheckBox chkFaciUse 
      Caption         =   "�ӫȤ�]�ƥ��`�ϥΤ�"
      Enabled         =   0   'False
      Height          =   195
      Left            =   9240
      TabIndex        =   20
      Top             =   270
      Width           =   2625
   End
   Begin VB.CheckBox chkCancel 
      Caption         =   "�@�o��ڬO�_�@�ֲέp"
      Height          =   225
      Left            =   9240
      TabIndex        =   24
      Top             =   1110
      Width           =   2625
   End
   Begin VB.PictureBox pic2 
      Height          =   4200
      Left            =   30
      ScaleHeight     =   4140
      ScaleWidth      =   11835
      TabIndex        =   68
      Top             =   2160
      Width           =   11895
      Begin VB.VScrollBar vsl1 
         Height          =   4095
         LargeChange     =   100
         Left            =   11520
         Max             =   100
         SmallChange     =   100
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.Frame fraMulti 
         BorderStyle     =   0  '�S���ؽu
         Caption         =   "Frame1"
         Height          =   7305
         Left            =   60
         TabIndex        =   70
         Top             =   0
         Width           =   11505
         Begin CS_Multi.CSmulti gimMduId 
            Height          =   345
            Left            =   -30
            TabIndex        =   38
            Top             =   4470
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "�j  ��  �s  ��"
         End
         Begin CS_Multi.CSmulti gimStrtCode 
            Height          =   345
            Left            =   -30
            TabIndex        =   37
            Top             =   4110
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "��  �D  �s  ��"
         End
         Begin Gi_Multi.GiMulti gimUCCode 
            Height          =   345
            Left            =   -30
            TabIndex        =   40
            Top             =   5145
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "��  ��  ��  �]"
         End
         Begin Gi_Multi.GiMulti gimAreaCode 
            Height          =   345
            Left            =   -30
            TabIndex        =   36
            Top             =   3765
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "��     �F     ��"
         End
         Begin Gi_Multi.GiMulti gimServCode 
            Height          =   345
            Left            =   -30
            TabIndex        =   35
            Top             =   3435
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "�A     ��     ��"
         End
         Begin Gi_Multi.GiMulti gimBillType 
            Height          =   345
            Left            =   -30
            TabIndex        =   29
            Top             =   1335
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "��  ��  ��  �O"
            DataType        =   2
            DIY             =   -1  'True
            Exception       =   -1  'True
         End
         Begin Gi_Multi.GiMulti gimOrder 
            Height          =   345
            Left            =   -30
            TabIndex        =   43
            Top             =   6195
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "��  ��  ��  ��"
            DataType        =   2
            ColumnOrder     =   -1  'True
         End
         Begin Gi_Multi.GiMulti gimCustStatus 
            Height          =   345
            Left            =   -30
            TabIndex        =   34
            Top             =   3090
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "��  ��  ��  �A"
         End
         Begin Gi_Multi.GiMulti gimFaci 
            Height          =   315
            Left            =   -30
            TabIndex        =   26
            Top             =   330
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   556
            ButtonCaption   =   "�]  ��  ��  ��"
         End
         Begin Gi_Multi.GiMulti gimCMCode 
            Height          =   345
            Left            =   -30
            TabIndex        =   28
            Top             =   990
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "��  �O  ��  ��"
         End
         Begin Gi_Multi.GiMulti gimBankCode 
            Height          =   345
            Left            =   -30
            TabIndex        =   41
            Top             =   5490
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "��  ��  ��  �O"
         End
         Begin Gi_Multi.GiMulti gimCancelCode 
            Height          =   345
            Left            =   -30
            TabIndex        =   39
            Top             =   4800
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "�@  �o  ��  �]"
         End
         Begin CS_Multi.CSmulti gimClctEn 
            Height          =   345
            Left            =   -30
            TabIndex        =   30
            Top             =   1680
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "��  �O  �H   ��"
         End
         Begin CS_Multi.CSmulti gimCitemCode 
            Height          =   315
            Left            =   -30
            TabIndex        =   27
            Top             =   660
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   556
            ButtonCaption   =   "��  �O  ��  ��"
         End
         Begin CS_Multi.CSmulti gimClassCode 
            Height          =   345
            Left            =   -30
            TabIndex        =   33
            Top             =   2760
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "��  ��  ��  �O"
         End
         Begin Gi_Multi.GiMulti gimCardCode 
            Height          =   375
            Left            =   -30
            TabIndex        =   42
            Top             =   5850
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   661
            ButtonCaption   =   "�H �� �d �� ��"
         End
         Begin Gi_Multi.GiMulti gimServiceType 
            Height          =   345
            Left            =   -30
            TabIndex        =   25
            Top             =   0
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "�A  ��  ��  �O"
            DataType        =   2
         End
         Begin CS_Multi.CSmulti gimOldClctEn 
            Height          =   345
            Left            =   -30
            TabIndex        =   31
            Top             =   2040
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "�� �� �O �H ��"
         End
         Begin Gi_Multi.GiMulti gimClctArea 
            Height          =   345
            Left            =   -30
            TabIndex        =   32
            Top             =   2400
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "��     �O     ��"
         End
         Begin CS_Multi.CSmulti gimInstall 
            Height          =   345
            Left            =   -30
            TabIndex        =   83
            Top             =   6540
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   609
            ButtonCaption   =   "�˾����u���O"
         End
         Begin Gi_Multi.GiMulti gimPayType 
            Height          =   390
            Left            =   -30
            TabIndex        =   89
            Top             =   6900
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   688
            ButtonCaption   =   "ú  �I  ��  �O"
         End
      End
   End
   Begin VB.Frame fraPage 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�����覡"
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   0
      TabIndex        =   67
      Top             =   6450
      Width           =   5715
      Begin VB.OptionButton optServCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�A�Ȱ�"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   270
         TabIndex        =   44
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optAreaCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��F��"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1275
         TabIndex        =   45
         Top             =   300
         Width           =   975
      End
      Begin VB.OptionButton optClctEn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "���O�H��"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3285
         TabIndex        =   47
         Top             =   300
         Width           =   1125
      End
      Begin VB.OptionButton optNothing 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�L"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4440
         TabIndex        =   48
         Top             =   300
         Width           =   495
      End
      Begin VB.OptionButton optClctAreaCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "���O��"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2280
         TabIndex        =   46
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame fraType 
      BackColor       =   &H00E0E0E0&
      Caption         =   "����榡"
      ForeColor       =   &H00004080&
      Height          =   645
      Left            =   5790
      TabIndex        =   66
      Top             =   6450
      Width           =   2805
      Begin VB.OptionButton optSummaries 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�J�`��"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   1230
         TabIndex        =   50
         Top             =   300
         Width           =   915
      End
      Begin VB.OptionButton optDail 
         BackColor       =   &H00E0E0E0&
         Caption         =   "���Ӫ�"
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   180
         TabIndex        =   49
         Top             =   270
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�L��Ǹ�"
      ForeColor       =   &H00800080&
      Height          =   645
      Left            =   8640
      TabIndex        =   65
      Top             =   6450
      Width           =   3285
      Begin VB.OptionButton optPrtSNoYes 
         Caption         =   "�w��"
         ForeColor       =   &H00800080&
         Height          =   405
         Left            =   210
         TabIndex        =   51
         Top             =   180
         Width           =   705
      End
      Begin VB.OptionButton optPrtSNoNo 
         Caption         =   "����"
         ForeColor       =   &H00800080&
         Height          =   405
         Left            =   975
         TabIndex        =   52
         Top             =   180
         Width           =   705
      End
      Begin VB.OptionButton optPrtSNoAll 
         Caption         =   "����"
         ForeColor       =   &H00800080&
         Height          =   405
         Left            =   1740
         TabIndex        =   53
         Top             =   180
         Value           =   -1  'True
         Width           =   705
      End
   End
   Begin VB.TextBox txtCustid 
      Height          =   345
      Left            =   6270
      TabIndex        =   17
      Top             =   750
      Width           =   2925
   End
   Begin VB.TextBox txtPrtSNo2 
      Height          =   345
      Left            =   7830
      MaxLength       =   12
      TabIndex        =   16
      Top             =   390
      Width           =   1365
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "�צ�Excel"
      Height          =   525
      Left            =   3690
      TabIndex        =   56
      Top             =   7350
      Width           =   1245
   End
   Begin VB.TextBox txtPrtSNo1 
      Height          =   345
      Left            =   6270
      MaxLength       =   12
      TabIndex        =   15
      Top             =   390
      Width           =   1395
   End
   Begin VB.CheckBox chkGUINo 
      Caption         =   "�O�_�w�}�o��"
      Height          =   195
      Left            =   9240
      TabIndex        =   23
      Top             =   900
      Width           =   2625
   End
   Begin VB.CheckBox chkInvoice 
      Caption         =   "�O�_�w����"
      Height          =   195
      Left            =   9240
      TabIndex        =   22
      Top             =   705
      Width           =   2625
   End
   Begin VB.CheckBox chkPreInvoice 
      Caption         =   "�O�_�u�C�L�w�}�o���Ȥ�"
      Height          =   195
      Left            =   9240
      TabIndex        =   21
      Top             =   495
      Width           =   2625
   End
   Begin VB.CheckBox chkFaciName 
      Caption         =   "�O�_�C�L�Ȥ�]�ƶ���"
      Height          =   195
      Left            =   9240
      TabIndex        =   19
      Top             =   60
      Width           =   2625
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:�C�L"
      Default         =   -1  'True
      Height          =   525
      Left            =   270
      TabIndex        =   54
      Top             =   7350
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "����(&X)"
      Height          =   525
      Left            =   10500
      TabIndex        =   57
      Top             =   7350
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "�W���έp���G"
      Height          =   525
      Left            =   1920
      TabIndex        =   55
      Top             =   7350
      Width           =   1395
   End
   Begin Gi_Date.GiDate gdaShouldDate2 
      Height          =   345
      Left            =   2400
      TabIndex        =   1
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
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
   Begin Gi_Date.GiDate gdaShouldDate1 
      Height          =   345
      Left            =   990
      TabIndex        =   0
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
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
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   6270
      TabIndex        =   14
      Top             =   60
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
      F5Corresponding =   -1  'True
   End
   Begin Gi_Time.GiTime gdtUpdTime2 
      Height          =   345
      Left            =   3480
      TabIndex        =   5
      Top             =   780
      Width           =   1875
      _ExtentX        =   3307
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
   Begin Gi_Time.GiTime gdtUpdTime1 
      Height          =   345
      Left            =   990
      TabIndex        =   4
      Top             =   780
      Width           =   1875
      _ExtentX        =   3307
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
   Begin Gi_Time.GiTime gdaCreateTime1 
      Height          =   345
      Left            =   990
      TabIndex        =   2
      Top             =   435
      Width           =   1875
      _ExtentX        =   3307
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
   Begin Gi_Time.GiTime gdaCreateTime2 
      Height          =   345
      Left            =   3480
      TabIndex        =   3
      Top             =   435
      Width           =   1875
      _ExtentX        =   3307
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
   Begin Gi_YM.GiYM gymClctYM2 
      Height          =   345
      Left            =   2070
      TabIndex        =   7
      Top             =   1140
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
   Begin Gi_Date.GiDate gdaClctDate1 
      Height          =   345
      Left            =   990
      TabIndex        =   10
      Top             =   1500
      Width           =   1095
      _ExtentX        =   1931
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
   Begin Gi_Date.GiDate gdaClctDate2 
      Height          =   345
      Left            =   2250
      TabIndex        =   11
      Top             =   1500
      Width           =   1095
      _ExtentX        =   1931
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
   Begin prjNumber.GiNumber ginRealPeriod2 
      Height          =   345
      Left            =   4740
      TabIndex        =   9
      Top             =   1140
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   609
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
   End
   Begin prjNumber.GiNumber ginShouldAmt1 
      Height          =   345
      Left            =   4320
      TabIndex        =   12
      Top             =   1500
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   609
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
      AllowMinus      =   -1  'True
   End
   Begin prjNumber.GiNumber ginShouldAmt2 
      Height          =   345
      Left            =   5400
      TabIndex        =   13
      Top             =   1500
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   609
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
      AllowMinus      =   -1  'True
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5220
      TabIndex        =   82
      Top             =   1620
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4560
      TabIndex        =   81
      Top             =   1260
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "���O���B"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3450
      TabIndex        =   80
      Top             =   1590
      Width           =   780
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "���O����"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3000
      TabIndex        =   79
      Top             =   1230
      Width           =   780
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2130
      TabIndex        =   78
      Top             =   1590
      Width           =   105
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�U����"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   330
      TabIndex        =   77
      Top             =   1590
      Width           =   585
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1890
      TabIndex        =   76
      Top             =   1230
      Width           =   105
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�����~��"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   75
      Top             =   1230
      Width           =   780
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      Caption         =   "�a�}�̾�"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   5430
      TabIndex        =   74
      Top             =   1230
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3090
      TabIndex        =   73
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3090
      TabIndex        =   72
      Top             =   495
      Width           =   105
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�X����"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   71
      Top             =   510
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
      Caption         =   "�Ȥ�s��"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   5430
      TabIndex        =   64
      Top             =   840
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "���ʤ��"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   63
      Top             =   870
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   7710
      TabIndex        =   62
      Top             =   450
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
      Caption         =   "�L��Ǹ�"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   5430
      TabIndex        =   61
      Top             =   495
      Width           =   795
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
      Caption         =   "��  �q  �O"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5430
      TabIndex        =   60
      Top             =   150
      Width           =   795
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2190
      TabIndex        =   59
      Top             =   150
      Width           =   105
   End
   Begin VB.Label lblShouldDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�������"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   58
      Top             =   150
      Width           =   780
   End
End
Attribute VB_Name = "frmSO5210A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table:SO033,SO001,SO014,CD004,SO002,SO002A,SO004,SO035
Option Explicit
Dim strChooseSO004 As String
Dim strChooseSO002 As String
Dim strChooseSO002A As String
Dim strChooseSO003 As String
Dim strTableSO004 As String
Dim strWhereSO004 As String
Dim strTableSO007 As String
Dim strChooseSO007 As String
Dim blnExcel As Boolean
Dim strViewName As String
Dim strViewName1 As String
Dim strChoose35 As String
Dim strSort As String
Dim strCustId As String

Private Sub chkCancel_Click()
  On Error GoTo ChkErr
    If chkCancel.Value = 1 Then
        gimCancelCode.Enabled = True
    Else
        gimCancelCode.Enabled = False
        gimCancelCode.Clear
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "chkCancel_Click")
End Sub

Private Sub chkFaciName_Click()
  On Error Resume Next
  If chkFaciName.Value = 1 Then
      chkFaciUse.Enabled = True
  Else
      chkFaciUse.Enabled = False
  End If
End Sub

Private Sub cmdExcel_Click()
  On Error GoTo ChkErr
    blnExcel = True
    chkFaciName.Value = 0
    cmdPrint_Click
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExcel_Click")
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      If optDail Then
          Call PreviousRpt(GetPrinterName(5), RptName("SO5210"), "�C�를ú�O�Ȥ���Ӫ� [SO5210A]")
      Else
          Call PreviousRpt(GetPrinterName(5), RptName("SO5210", 3), "�C�를ú�O�Ȥ�J�`�� [SO5210A]")
      End If
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo ChkErr
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    Screen.MousePointer = vbHourglass
      cmdExit.SetFocus
      Me.Enabled = False
        ReadyGoPrint
        Call subChoose
        'Call subCreateView
        If Not subCreateTable Then Exit Sub
        Call subPrint
      Me.Enabled = True
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: strErrFile = "���q�O": GoTo 66
    If Len(gimCancelCode.GetDispStr) > 0 Then gimUCCode.Clear
    IsDataOk = True
  Exit Function
66:
  MsgMustBe (strErrFile)
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub subChoose()
  On Error GoTo ChkErr
    Dim strpagetype As String
    Dim intSort As Integer
    Dim arrOrder() As String
    Dim varOrder As Variant
    Dim strPrtSNo As String
    Dim blnPrtSNo As Boolean
    Dim strCustId As String
        strChoose = ""
        strChooseSO004 = ""
        strChooseSO002 = ""
        strChooseSO002A = ""
        strChooseString = ""
        strGroupName = ""
        strChooseSO003 = ""
        intSort = 0
        strPrtSNo = ""
        blnPrtSNo = False
        strTableSO007 = ""
        '���
        If gdaShouldDate1.GetValue <> "" Then Call subAnd("SO033.ShouldDate >= To_Date('" & gdaShouldDate1.GetValue & "','YYYYMMDD')")
        'If gdaShouldDate2.GetValue <> "" Then Call subAnd("SO033.ShouldDate < To_Date('" & gdaShouldDate2.GetValue & "','YYYYMMDD')+1")
        If gdaShouldDate2.GetValue <> "" Then Call subAnd("SO033.ShouldDate < To_Date('" & gdaShouldDate2.GetValue & "','YYYYMMDD')+1")
        
        If gdtUpdTime1.GetValue <> "" Then subAnd ("substr(SO033.UpdTime,1,14) >='" & gdtUpdTime1.GetOriginalValue & "'")
        If gdtUpdTime2.GetValue <> "" Then subAnd ("substr(SO033.UpdTime,1,14) <='" & gdtUpdTime2.GetOriginalValue & "'")
        '      If gdaCreateTime1.GetValue <> "" Then subAnd ("substr(SO033.CreateTime,1,14) >='" & gdaCreateTime1.GetOriginalValue & "'")
        '      If gdaCreateTime2.GetValue <> "" Then subAnd ("substr(SO033.CreateTime,1,14) <='" & gdaCreateTime2.GetOriginalValue & "'")
        
        If gdaCreateTime1.GetValue <> "" Then Call subAnd("SO033.CreateTime >= To_Date('" & gdaCreateTime1.GetValue & "','YYYYMMDDHH24MI')")
        If gdaCreateTime2.GetValue <> "" Then Call subAnd("SO033.CreateTime < To_Date('" & gdaCreateTime2.GetValue & "','YYYYMMDDHH24MI')")
        '���D��2731 �W�[�����~��
        'If gymClctYM1.GetValue <> "" Then Call subAnd("SO033.ClctYM >= To_Number('" & gymClctYM1.GetValue & "')")
        If gymClctYM1.GetValue <> "" Then Call subAnd("SO033.ClctYM >=" & gymClctYM1.GetValue)
        If gymClctYM2.GetValue <> "" Then
            Dim strDateAdd As String
            strDateAdd = Format(DateAdd("M", 1, gymClctYM2.GetValue(True)), "YYYYMM")
            Call subAnd("SO033.ClctYM<" & strDateAdd)
            'Call subAnd("SO033.ClctYM < To_Number('" & x & "')")
        End If
        
'        If gymClctYM2.GetValue <> "" Then Call subAnd("SO033.ClctYM < To_Date('" & gymClctYM2.GetValue & "','YYYYMM')+1")
        '�L��Ǹ�
        If Len(Trim(txtPrtSNo1)) <> 0 Then Call subAnd("SO033.PrtSNo>='" & Trim(txtPrtSNo1) & "'"): blnPrtSNo = True
        If Len(Trim(txtPrtSNo2)) <> 0 Then Call subAnd("SO033.PrtSNo<='" & Trim(txtPrtSNo2) & "'"): blnPrtSNo = True
        Select Case True
               Case optPrtSNoYes.Value '���L��Ǹ��~�L
                    If blnPrtSNo = False Then subAnd ("SO033.PrtSNo IS Not Null")
                    strPrtSNo = "�w��"
               Case optPrtSNoNo.Value  '�L�L��Ǹ��~�L
                    subAnd ("SO033.PrtSNo IS Null")
                    strPrtSNo = "����"
               Case optPrtSNoAll.Value '���ަ��L�L��Ǹ��ҦL
                    strPrtSNo = "����"
        End Select
        '�Ȥ�s��
        If Len(Trim(txtCustId)) > 0 Then
            strCustId = Replace(txtCustId.Text, ",", "','")
            Call subAnd("SO033.Custid IN ('" & Trim(strCustId) & "')")
        End If
        '#2810 �W�[�U�����O�� By Kin 2007/05/10
        If gdaClctDate1.GetValue & "" <> "" Then
            Call subAnd2(strChooseSO003, "SO003.ClctDate >= To_Date('" & gdaClctDate1.GetValue & "','YYYYMMDD')")
        End If
        If gdaClctDate2.GetValue & "" <> "" Then
            Call subAnd2(strChooseSO003, "SO003.ClctDate < To_Date('" & gdaClctDate2.GetValue & "','YYYYMMDD')+1")
        End If
        '#2810 �W�[���ư϶����� By Kin 2007/05/10
        If ginRealPeriod1.Text & "" <> "" Then
            Call subAnd("SO033.RealPeriod>=" & ginRealPeriod1.Value)
        End If
        If ginRealPeriod2.Text & "" <> "" Then
            Call subAnd("SO033.RealPeriod<=" & ginRealPeriod2.Value)
        End If
        '#2810 �W�[�������B�϶����� By Kin 2007/05/10
        If ginShouldAmt1.Text & "" <> "" Then
            Call subAnd("SO033.ShouldAmt>=" & ginShouldAmt1.Value)
        End If
        If ginShouldAmt2.Text & "" <> "" Then
            Call subAnd("SO033.ShouldAmt<=" & ginShouldAmt2.Value)
        End If
        
        
        'GiMulti
        If gimFaci.GetQryStr <> "" Then strChooseSO004 = " And SO004.FaciCode " & gimFaci.GetQryStr
        If gimCitemCode.GetQryStr <> "" Then Call subAnd("SO033.CitemCode " & gimCitemCode.GetQryStr)
        If gimCMCode.GetQryStr <> "" Then Call subAnd("SO033.CMCode " & gimCMCode.GetQryStr)
        If gimBillType.GetQryStr <> "" Then Call subAnd("SubStr(SO033.BillNo,7,1) " & gimBillType.GetQryStr)
        If gimClctEn.GetQryStr <> "" Then Call subAnd("SO033.ClctEn " & gimClctEn.GetQryStr)
        '���D2595 �s�W�즬�O�H��,���O�ϱ��� 2006/7/17
        If gimOldClctEn.GetQryStr <> "" Then Call subAnd("SO033.OldClctEn " & gimOldClctEn.GetQryStr)
        If gimClctArea.GetQryStr <> "" Then Call subAnd("SO033.ClctAreaCode " & gimClctArea.GetQryStr)
        If gimClassCode.GetQryStr <> "" Then Call subAnd("SO033.ClassCode " & gimClassCode.GetQryStr)
        If gimCustStatus.GetQryStr <> "" Then Call subAnd2(strChooseSO002, "SO002.CustStatusCode " & gimCustStatus.GetQryStr)
        If gimAreaCode.GetQryStr <> "" Then Call subAnd("SO033.AreaCode " & gimAreaCode.GetQryStr)
        If gimServCode.GetQryStr <> "" Then Call subAnd("SO033.ServCode " & gimServCode.GetQryStr)
        If gimStrtCode.GetQryStr <> "" Then Call subAnd("SO033.StrtCode " & gimStrtCode.GetQryStr)
        '#5683 �W�[ú�I���O By Kin 2010/08/18
        If gimPayType.GetQryStr <> "" Then Call subAnd("SO033.PayKind " & gimPayType.GetQryStr)
        If gimMduId.GetQryStr <> "" Then
            Call subAnd("SO033.MduId " & gimMduId.GetQryStr)
        Else
            If gimMduId.GetDispStr <> "" Then subAnd ("SO033.MduId is Not Null")
        End If
        If Len(gimCancelCode.GetDispStr) > 0 Then
            If gimCancelCode.GetQryStr <> "" Then
                Call subAnd("SO033.CancelCode " & gimCancelCode.GetQryStr)
            Else
                Call subAnd("SO033.CancelCode is not null")
            End If
        End If
        
        
        If gimBankCode.GetQryStr <> "" Then
            Call subAnd("SO033.BankCode " & gimBankCode.GetQryStr)
        Else
           If gimBankCode.GetDispStr <> "" Then Call subAnd("SO033.BankCode is not null")
        End If
        If gimCardCode.GetQryStr <> "" Then Call subAnd2(strChooseSO002A, "SO002A.CardName " & gimCardCode.GetDescStr)
        If gimServiceType.GetQryStr <> "" Then
            Call subAnd("SO033.ServiceType " & gimServiceType.GetQryStr)
'            strChooseSO004 = strChooseSO004 & " And SO004.ServiceType " & gimServiceType.GetQryStr
        End If
        
        'GiList
        If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO033.CompCode=" & gilCompCode.GetCodeNo)
        If gilCompCode.GetCodeNo <> "" Then Call subAnd2(strChooseSO002, "SO002.CompCode=" & gilCompCode.GetCodeNo)
        
        '#4289 �p�G����ܸ˾����O,�h�n��SO007 By Kin 2008/12/17
        If gimInstall.GetQueryCode <> "" Then
            strTableSO007 = ",SO007"
            Call subAnd("SO033.BillNo=SO007.SNO(+)")
            If gimInstall.GetQryStr & "" <> "" Then
                Call subAnd("SO007.InstCode" & gimInstall.GetQryStr)
            End If
        End If
        
        '      If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO033.ServiceType='" & gilServiceType.GetCodeNo & "'")
        
        'CheckBox
        '�O�_�u�C�L�w�}�o���Ȥ�
        If chkPreInvoice.Value = 1 Then Call subAnd2(strChooseSO002, "SO002.PreInvoice=1")
        
        '      If blnSO002 Then
        '        strPreInvoice = ",SO002 Where SO033.Custid=SO002.Custid And SO033.ServiceType=SO002.ServiceType And "
        '      Else
        '        strPreInvoice = " Where "
        '      End If
        
        '�O�_�w����
        If chkInvoice.Value = 1 Then Call subAnd("SO033.InvoiceTime is not null")
        '�O�_�w�}�o��
        If chkGUINo.Value = 1 Then Call subAnd("SO033.GUINo is not null")
        '�@�o��ڬO�_�@�ֲέp --�אּ����SO035��� ���D��1201 Liga 2004/10/27
        '      If chkCancel.Value = 0 Then subAnd ("SO033.cancelFlag=0")
        '�אּ�L�ocancelFlag<>0 ����� ���D��2531 Lydia 2006/06/13
              If chkCancel.Value = 0 Then subAnd ("SO033.cancelFlag<>1")
        '#2810 �ϥΩ��Ӫ�ɦp�G�ӳ]�Ʀ��ӽФH�m�W�h����̪��m�W�P�q��,�nShow�ӽФH�����,�ҥH�n�NSO004����Ƥ]��i�h By Kin 2007/05/10
        If optDail Then
            Call subAnd("SO033.FaciSeqNo=SO004.SEQNO(+) And SO033.CustID=SO004.CustID(+)")
        End If
        
        '�����覡
        Select Case True
               Case optAreaCode.Value
                    strGroupName = "ReportType=True;GroupName={V.AreaCode};GroupName1={V.AreaName}"
                    strpagetype = "��F��"
               Case optServCode.Value
                    strGroupName = "ReportType=True;GroupName={V.ServCode};GroupName1={V.ServName}"
                    strpagetype = "�A�Ȱ�"
               Case optClctAreaCode.Value
                    strGroupName = "ReportType=True;GroupName={V.ClctAreaCode};GroupName1={V.ClctAreaName}"
                    strpagetype = "���O��"
               Case optClctEn.Value
                    strGroupName = "ReportType=True;GroupName={V.ClctEn};GroupName1={V.ClctName}"
                    strpagetype = "���O��"
               Case optNothing.Value
                    Select Case Mid(gimOrder.GetColumnOrderCode, 1, 1)
                           Case "A"
                                strGroupName = "ReportType=False;GroupName={V.ShouldDate};GroupName1={V.CustName}"
                           Case "B"
                                '951130 #2880 �a�}�L�k�Ƨ�GroupName={V.Address}�אּGroupName={V.AddrSort}
                                strGroupName = "ReportType=False;GroupName={V.AddrSort};GroupName1={V.CustName}"
                           Case "C"
                                strGroupName = "ReportType=False;GroupName={V.CustId};GroupName1={V.CustName}"
                           Case "D"
                                strGroupName = "ReportType=False;GroupName={V.ClctEn};GroupName1={V.CustName}"
                           Case "E"
                                strGroupName = "ReportType=False;GroupName={V.BillNo};GroupName1={V.CustName}"
                           Case "F"
                                strGroupName = "ReportType=False;GroupName={V.ServCode};GroupName1={V.CustName}"
                           Case "G"
                                strGroupName = "ReportType=False;GroupName={V.Description};GroupName1={V.CustName}"
                           Case "H"
                                strGroupName = "ReportType=False;GroupName={V.UpdTime};GroupName1={V.CustName}"
                           Case "I"
                                strGroupName = "ReportType=False;GroupName={V.CitemCode};GroupName1={V.CustName}"
                                
                           Case Else
                                strGroupName = "ReportType=False;GroupName={V.CustId};GroupName1={V.CustName}"
                    End Select
        End Select
        
        '�ƧǤ覡
        Call SetGiMultiOrder
        
        If chkFaciName.Value = 1 Then
            strGroupName = strGroupName & ";PrintFaci=true"
        Else
            strGroupName = strGroupName & ";PrintFaci=false"
        End If
        strChooseString = "�������: " & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & "�����~��: " & subSpace(gymClctYM1.GetValue(True)) & "~" & subSpace(gymClctYM2.GetValue(True)) & _
                          "���ʤ��: " & subSpace(gdtUpdTime1.GetValue(True)) & "~" & subSpace(gdtUpdTime2.GetValue(True)) & ";" & _
                          "�X����: " & subSpace(gdaCreateTime1.GetValue(True)) & "~" & subSpace(gdaCreateTime2.GetValue(True)) & ";" & _
                          "�Ȥ�s��: " & subSpace(Trim(txtCustId.Text)) & ";" & _
                          "�L��Ǹ�(" & strPrtSNo & "): " & subSpace(txtPrtSNo1.Text) & "~" & subSpace(txtPrtSNo2.Text) & ";" & _
                          "���q�O�@: " & subSpace(gilCompCode.GetDescription) & ";" & "�A�����O: " & subSpace(gimServiceType.GetDispStr) & ";" & _
                          "�]�ƶ���: " & subSpace(gimFaci.GetDispStr) & ";" & _
                          "���O����: " & subSpace(gimCitemCode.GetDispStr) & ";" & _
                          "���O�覡: " & subSpace(gimCMCode.GetDispStr) & ";" & _
                          "������O: " & subSpace(gimBillType.GetDispStr) & ";" & _
                          "���O�H��: " & subSpace(gimClctEn.GetDispStr) & "; �즬�O�H��: " & subSpace(gimOldClctEn.GetDispStr) & "; ���O��: " & subSpace(gimClctArea.GetDispStr) & ";" & _
                          "�Ȥ����O: " & subSpace(gimClassCode.GetDispStr) & ";" & _
                          "�Ȥ᪬�A: " & subSpace(gimCustStatus.GetDispStr) & ";" & _
                          "�A�Ȱϡ@: " & subSpace(gimServCode.GetDispStr) & ";" & "��F�ϡ@: " & subSpace(gimAreaCode.GetDispStr) & ";" & _
                          "��D�W��: " & subSpace(gimStrtCode.GetDispStr) & ";" & _
                          "�j�ӦW��: " & subSpace(gimMduId.GetDispStr) & ";" & _
                          "�@�o��]: " & subSpace(gimCancelCode.GetDispStr) & ";" & _
                          "������]: " & subSpace(gimUCCode.GetDispStr) & ";" & _
                          "�Ȧ����O: " & subSpace(gimBankCode.GetDispStr) & ";" & _
                          "�H�Υd����:" & subSpace(gimCardCode.GetDispStr) & ";" & _
                          "�ƧǤ覡: " & subSpace(gimOrder.GetColumnOrderDspStr) & ";"
        strChooseString = strChooseString & "ú�I���O: " & subSpace(gimPayType.GetDispStr) & ";"
        strChooseString = strChooseString & _
                          "�˾����u���O:" & subSpace(gimInstall.GetDispStr) & ";" & _
                          IIf(strpagetype = "", "", ";" & "�����覡: " & strpagetype) & ";" & _
                          "�O�_�[�L�Ȥ�]�ƶ���:" & IIf(chkFaciName.Value = 1, "�O", "�_") & "   �@�o��ڬO�_�@�ֲέp:" & IIf(chkCancel.Value = 1, "�O", "�_") & ";" & _
                          IIf(chkPreInvoice.Value = 1, "�u�C�L�w�}�o���Ȥ�;", "") & IIf(chkInvoice.Value = 1, "�u�C�L�w����;", "") & _
                          IIf(chkGUINo.Value = 1, "�u�C�L�w�}�o��;", "") & ";" & _
                          IIf(chkMediabillNo.Value = 1, "���C��渹;", "") & ";" & _
                          "����榡:" & IIf(optDail.Value = True, "���Ӫ�", "�J�`��")
                         
Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Public Sub SetGiMultiOrder()
  On Error GoTo ChkErr
    Dim intSort As Integer
    Dim arrOrder() As String
    Dim varOrder As Variant
    
    intSort = 0
    arrOrder = Split(gimOrder.GetColumnOrderCode, ",")
      
      For Each varOrder In arrOrder
          Select Case arrOrder(intSort)
                 Case "A"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.ShouldDate}"
                      strSort = "ShouldDate"
                 Case "B"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.AddrSort}"
                      strSort = "AddrSort"
                 Case "C"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.CustId}"
                      strSort = "CustId"
                 Case "D"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.ClctEn}"
                      strSort = "ClctEn"
                 Case "E"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.BillNo}"
                      strSort = "BillNo"
                 Case "F"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.ServCode}"
                      strSort = "ServCode"
                 Case "G"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.Description}"
                      strSort = "Description"
                 Case "H"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.UpdTime}"
                      strSort = "UpdTime"
                 Case "I"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.CitemCode}"
                      strSort = "CitemCode"
          End Select
          intSort = intSort + 1
          strSort = " Order by " & strSort
      Next
    'If Len(strGroupName) > 0 Then strGroupName = Mid(strGroupName, 2)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "SetGiMultiOrder")
End Sub

Private Sub subPrint()
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSubQry(7)  As String
    Dim strSubQry2(7)  As String
    Dim intFor As Integer
    'Dim rsExcel As New ADODB.Recordset
    Dim rsSumCust As New ADODB.Recordset
    Dim strSumCust As String
    Dim Para24 As Integer
    Dim strExcel As String
    If optDail.Value Then
        If Not GetRS(rsTmp, "Select Count(*) as intCount From " & strViewName) Then Exit Sub
    Else
        If Not GetRS(rsTmp, "Select intCount From " & strViewName) Then Exit Sub
    End If
    
    
    strsql = "Select * From " & strViewName & " V"
    If rsTmp("intCount") = 0 Then
        MsgNoRcd
        SendSQL , , True
    Else
        'strsql = "Select * From " & strViewName & " V "
        If blnExcel Then
             '#2810 �ץX��,�ݧP�_�v���O�_�N�b��������ܥX��
'            If GetUserPriv("SO1100G", "(SO1100G9)") Then
'                strPrivAccount = "AccountNo"
'            Else
'                strPrivAccount = "Decode(AccountNo,Null,AccountNo,SubStr(AccountNo,1,5) || 'XXX' || " & _
'                               "SubStr(AccountNo,9,2) || 'XX' || SubStr(AccountNo,Length(AccountNo)-3,4))"
''                    txtAccountNo.Text = Left(txtAccountNo.Tag, 5) & "XXX" & Mid(txtAccountNo.Tag, 9, 2) & "XX" & Right(txtAccountNo.Tag, 4)
'            End If
            '#2810 �ץXExcel�W�[ �b��\�d��,�H�Υd����� By Kin 2007/05/10
             
'            GetRS rsExcel, "Select CustId �Ƚs,BILLNO ��ڽs��,CUSTNAME �Ȥ�m�W,SHOULDDATE �������,TEL1 �q��1,TEL2 �q��2,TEL3 �q��3," & _
'                           "Description �Ȥ����O, CitemCode ���O���إN��,CitemName ���O����,RealPeriod ����,ShouldAmt �������B,RealStartDate ���Ĵ����_�l,RealStopDate ���Ĵ����I��," & _
'                           "GUINo �o�����X,UCNAME ������],CLCTNAME ���O��,PrtSNo �L��Ǹ�,ADDRESS �a�},SERVNAME �A�Ȱ�,AreaName ��F��,CustStatusName �Ȥ᪬�A,MDUName �j�ӦW��," & _
'                           "CardName �H�Υd�O,OldClctName �즬�O�H��," & strPrivAccount & " �b���d��,CardExpDate �H�Υd�����  From " & strViewName & strSort, gcnGi, adUseServer
'            GetRS rsExcel, "Select CustId,BILLNO,CUSTNAME,SHOULDDATE,TEL1,TEL2,TEL3," & _
'                   "Description,CitemCode,CitemName,RealPeriod,ShouldAmt,RealStartDate,RealStopDate," & _
'                   "GUINo,UCNAME,CLCTNAME,PrtSNo,ADDRESS,SERVNAME,AreaName,CustStatusName,MDUName," & _
'                   "CardName,OldClctName," & strPrivAccount & " Account,CardExpDate,ServCode,AddrSort," & _
'                   "UpdTime,ClctAreaCode,AreaCode" & _
'                   " From " & strViewName & strSort, gcnGi, adUseServer
            
           '********************************************************************************************
            '#3429 �ץXExcel�쥻���{�������ץX�A�{�b�אּ��Rpt�ץX By Kin 2007/08/17
            strExcel = "Select * From " & strViewName & " V"
            toExcel strExcel
            'Call UseProperty(rsExcel, "�C�를ú�O�Ȥ���Ӫ�", "�Ĥ@��")
           '*********************************************************************************************
            blnExcel = False
        Else
            If optDail Then
                Call PrintRpt2(GetPrinterName(5), RptName("SO5210"), , "�C�를ú�O�Ȥ���Ӫ� [SO5210A]", strsql, strChooseString, , True, , , strGroupName, GiPaperLandscape)
            Else
                
                If strChooseSO002 <> "" Then strChooseSO002 = ",SO002 Where SO033.Custid=SO002.Custid And SO033.ServiceType=SO002.ServiceType And " & strChooseSO002
                If strChooseSO002A <> "" Then strChooseSO002 = ",SO002A" & IIf(strChooseSO002 = "", " WHERE ", strChooseSO002 & " AND ") & "SO033.CUSTID=SO002A.CUSTID AND SO033.AccountNo=SO002A.AccountNo AND " & strChooseSO002A
                strSumCust = "," & rsTmp("intCount") & " as SumCust "
              '���O����
              '#2810���ծɵo�{�ɪ�Bug,�p�G�]�ƶ��ئ��D����,�|�ɭP�������ƻP���Ӥ��۲�,�]���֦�F�]�ƶ��ظ��,�{�b�[�W�h(�������l������strTableSO004) By Kin 2007/05/15
              '#4289 ��ܸ˾����O�n�h��SO007 By Kin 2008/12/17
                strSubQry(0) = "SELECT SO033.CitemName,Sum(SO033.ShouldAmt) SumShouldAmt,Count(*) Counts,Count(Distinct SO033.Custid) CountCustid,Count(Distinct SO033.BillNo) CountBillNo " & strSumCust & _
                                     " From(" & _
                                       "Select SO033.CitemName,SO033.ShouldAmt,SO033.Custid,SO033.BillNo,SO033.ServiceType,SO033.AccountNo From SO033" & strTableSO007 & " Where " & strChoose & _
                                  ") SO033" & strTableSO004 & strChooseSO002 & strWhereSO004 & " Group by SO033.CitemName"
                strSubQry2(0) = "SELECT SO033.CitemName,Sum(SO033.ShouldAmt) SumShouldAmt,Count(*) Counts,Count(Distinct SO033.Custid) CountCustid,Count(Distinct SO033.BillNo) CountBillNo " & strSumCust & _
                                     " From(" & _
                                       "Select SO033.CitemName,SO033.ShouldAmt,SO033.Custid,SO033.BillNo,SO033.ServiceType,SO033.AccountNo From SO033" & strTableSO007 & " Where " & strChoose & _
                                       " Union all Select SO035.CitemName,SO035.ShouldAmt,SO035.Custid,SO035.BillNo,SO035.ServiceType,SO035.AccountNo From SO035" & strTableSO007 & " Where " & strChoose35 & _
                                  ") SO033" & strTableSO004 & strChooseSO002 & strWhereSO004 & " Group by SO033.CitemName"
               '#4289 ��ܸ˾����O�n�h��SO007 By Kin 2008/12/17
               If chkCancel.Value = 1 Then strSubQry(0) = strSubQry2(0)
              '�A�Ȱ�
                strSubQry(1) = "SELECT SO014.ServName,Sum(SO033.ShouldAmt) SumShouldAmt,Count(*) Counts,Count(Distinct SO033.Custid) CountCustid,Count(Distinct SO033.BillNo) CountBillNo " & strSumCust & _
                                     " From(" & _
                                         "Select SO033.ShouldAmt,SO033.Custid,SO033.BillNo,SO033.ServiceType,SO033.AddrNo,SO033.AccountNo From SO033" & strTableSO007 & " Where " & strChoose & _
                                   ") SO033,SO014" & strTableSO004 & IIf(strChooseSO002 = "", " Where ", strChooseSO002 & " And ") & "SO033.ADDRNO=SO014.ADDRNO " & strWhereSO004 & " Group by SO014.ServName"
                
                strSubQry2(1) = "SELECT SO014.ServName,Sum(SO033.ShouldAmt) SumShouldAmt,Count(*) Counts,Count(Distinct SO033.Custid) CountCustid,Count(Distinct SO033.BillNo) CountBillNo " & strSumCust & _
                                     " From(" & _
                                         "Select SO033.ShouldAmt,SO033.Custid,SO033.BillNo,SO033.ServiceType,SO033.AddrNo,SO033.AccountNo From SO033" & strTableSO007 & " Where " & strChoose & _
                                         " Union all Select SO035.ShouldAmt,SO035.Custid,SO035.BillNo,SO035.ServiceType,SO035.AddrNo,SO035.AccountNo From SO035" & strTableSO007 & " Where " & strChoose35 & _
                                   ") SO033,SO014" & strTableSO004 & IIf(strChooseSO002 = "", " Where ", strChooseSO002 & " And ") & "SO033.ADDRNO=SO014.ADDRNO " & strWhereSO004 & " Group by SO014.ServName"
                If chkCancel.Value = 1 Then strSubQry(1) = strSubQry2(1)
              '���O�H��+���O����
                Para24 = GetRsValue("Select Para24 From SO043 Where CompCode='" & gilCompCode.GetCodeNo & "'")
                '#4289 ��ܸ˾����O�n�h��SO007 By Kin 2008/12/17
                If Para24 = 1 Then
                    strSubQry(2) = "SELECT SO033.CLCTNAME,SO033.CitemName,SO033.ShouldAmt,Count(*) Counts,Sum(SO033.ShouldAmt) SumShouldAmt,Count(Distinct SO033.MEDIABILLNO) CountBillNo " & _
                                   " From(" & _
                                        "Select SO033.CLCTNAME,SO033.CitemName,SO033.ShouldAmt,SO033.Custid,SO033.MEDIABILLNO,SO033.ServiceType,SO033.AccountNo From SO033" & strTableSO007 & " Where " & strChoose & _
                                   ") SO033" & strTableSO004 & strChooseSO002 & strWhereSO004 & " Group by SO033.CLCTNAME,SO033.CitemName,SO033.ShouldAmt"
                    strSubQry2(2) = "SELECT SO033.CLCTNAME,SO033.CitemName,SO033.ShouldAmt,Count(*) Counts,Sum(SO033.ShouldAmt) SumShouldAmt,Count(Distinct SO033.MEDIABILLNO) CountBillNo " & _
                                   " From(" & _
                                        "Select SO033.CLCTNAME,SO033.CitemName,SO033.ShouldAmt,SO033.Custid,SO033.MEDIABILLNO,SO033.ServiceType,SO033.AccountNo From SO033" & strTableSO007 & " Where " & strChoose & _
                                        " Union all Select SO035.CLCTNAME,SO035.CitemName,SO035.ShouldAmt,SO035.Custid,SO035.MEDIABILLNO,SO035.ServiceType,SO035.AccountNo From SO035" & strTableSO007 & " Where " & strChoose35 & _
                                   ") SO033" & strTableSO004 & strChooseSO002 & strWhereSO004 & " Group by SO033.CLCTNAME,SO033.CitemName,SO033.ShouldAmt"
                                   

                Else
                    strSubQry(2) = "SELECT SO033.CLCTNAME,SO033.CitemName,SO033.ShouldAmt,Count(*) Counts,Sum(SO033.ShouldAmt) SumShouldAmt,Count(Distinct SO033.BillNo) CountBillNo " & _
                                         " From(" & _
                                             "Select SO033.CLCTNAME,SO033.CitemName,SO033.ShouldAmt,SO033.Custid,SO033.BillNo,SO033.ServiceType,SO033.AccountNo From SO033" & strTableSO007 & " Where " & strChoose & _
                                     ") SO033" & strTableSO004 & strChooseSO002 & strWhereSO004 & " Group by SO033.CLCTNAME,SO033.CitemName,SO033.ShouldAmt"
                    
                    strSubQry2(2) = "SELECT SO033.CLCTNAME,SO033.CitemName,SO033.ShouldAmt,Count(*) Counts,Sum(SO033.ShouldAmt) SumShouldAmt,Count(Distinct SO033.BillNo) CountBillNo " & _
                                         " From(" & _
                                             "Select SO033.CLCTNAME,SO033.CitemName,SO033.ShouldAmt,SO033.Custid,SO033.BillNo,SO033.ServiceType,SO033.AccountNo From SO033" & strTableSO007 & " Where " & strChoose & _
                                             " Union all Select SO035.CLCTNAME,SO035.CitemName,SO035.ShouldAmt,SO035.Custid,SO035.BillNo,SO035.ServiceType,SO035.AccountNo From SO035" & strTableSO007 & " Where " & strChoose35 & _
                                     ") SO033" & strTableSO004 & strChooseSO002 & strWhereSO004 & " Group by SO033.CLCTNAME,SO033.CitemName,SO033.ShouldAmt"
                End If
                If chkCancel.Value = 1 Then strSubQry(2) = strSubQry2(2)
              '��F��
                strSubQry(3) = "SELECT SO014.AreaName,Sum(SO033.ShouldAmt) SumShouldAmt,Count(*) Counts,Count(Distinct SO033.Custid) CountCustid,Count(Distinct SO033.BillNo) CountBillNo " & strSumCust & _
                                     " From(" & _
                                         "Select SO033.ShouldAmt,SO033.Custid,SO033.BillNo,SO033.ServiceType,SO033.AddrNo,SO033.AccountNo From SO033" & strTableSO007 & " Where " & strChoose & _
                                   ") SO033,SO014" & strTableSO004 & IIf(strChooseSO002 = "", " Where ", strChooseSO002 & " And ") & "SO033.ADDRNO=SO014.ADDRNO " & strWhereSO004 & " Group by SO014.AreaName"
                                   
                strSubQry2(3) = "SELECT SO014.AreaName,Sum(SO033.ShouldAmt) SumShouldAmt,Count(*) Counts,Count(Distinct SO033.Custid) CountCustid,Count(Distinct SO033.BillNo) CountBillNo " & strSumCust & _
                                     " From(" & _
                                         "Select SO033.ShouldAmt,SO033.Custid,SO033.BillNo,SO033.ServiceType,SO033.AddrNo,SO033.AccountNo From SO033" & strTableSO007 & " Where " & strChoose & _
                                         " Union all Select SO035.ShouldAmt,SO035.Custid,SO035.BillNo,SO035.ServiceType,SO035.AddrNo,SO035.AccountNo From SO035" & strTableSO007 & " Where " & strChoose35 & _
                                   ") SO033,SO014" & strTableSO004 & IIf(strChooseSO002 = "", " Where ", strChooseSO002 & " And ") & "SO033.ADDRNO=SO014.ADDRNO " & strWhereSO004 & " Group by SO014.AreaName"
                                   
                If chkCancel.Value = 1 Then strSubQry(3) = strSubQry2(3)
              '������]
                strSubQry(4) = "SELECT SO033.UCName,Sum(SO033.ShouldAmt) SumShouldAmt,Count(*) Counts,Count(Distinct SO033.Custid) CountCustid,Count(Distinct SO033.BillNo) CountBillNo " & strSumCust & _
                                     " From(" & _
                                         "Select SO033.UCName,SO033.ShouldAmt,SO033.Custid,SO033.BillNo,SO033.ServiceType,SO033.AccountNo From SO033" & strTableSO007 & " Where " & strChoose & _
                                   ") SO033" & strTableSO004 & strChooseSO002 & strWhereSO004 & " Group by SO033.UCName"
                
                strSubQry2(4) = "SELECT SO033.UCName,Sum(SO033.ShouldAmt) SumShouldAmt,Count(*) Counts,Count(Distinct SO033.Custid) CountCustid,Count(Distinct SO033.BillNo) CountBillNo " & strSumCust & _
                                     " From(" & _
                                         "Select SO033.UCName,SO033.ShouldAmt,SO033.Custid,SO033.BillNo,SO033.ServiceType,SO033.AccountNo From SO033" & strTableSO007 & " Where " & strChoose & _
                                         " Union all Select SO035.UCName,SO035.ShouldAmt,SO035.Custid,SO035.BillNo,SO035.ServiceType,SO035.AccountNo From SO035" & strTableSO007 & " Where " & strChoose35 & _
                                   ") SO033" & strTableSO004 & strChooseSO002 & strWhereSO004 & " Group by SO033.UCName"
                If chkCancel.Value = 1 Then strSubQry(4) = strSubQry2(4)
              '�H�Υd�d�O
                If strChooseSO002A = "" Then strChooseSO002 = ",SO002A" & IIf(strChooseSO002 = "", " WHERE ", strChooseSO002 & " AND ") & "SO033.CUSTID=SO002A.CUSTID(+) AND SO033.AccountNo=SO002A.AccountNo(+) "
                strSubQry(5) = "SELECT SO002A.CardName,Count(Distinct SO033.Custid) CountCustid,Count(Distinct SO002A.CardName) CountCardName " & strSumCust & _
                                     " From(" & _
                                         "Select SO033.Custid,SO033.AccountNo,SO033.ServiceType From SO033" & strTableSO007 & " Where " & strChoose & _
                                   ") SO033" & strTableSO004 & strChooseSO002 & strWhereSO004 & " Group by SO002A.CardName"
                                   
                strSubQry2(5) = "SELECT SO002A.CardName,Count(Distinct SO033.Custid) CountCustid,Count(Distinct SO002A.CardName) CountCardName " & strSumCust & _
                                     " From(" & _
                                         "Select SO033.Custid,SO033.AccountNo,SO033.ServiceType From SO033" & strTableSO007 & " Where " & strChoose & _
                                         " Union all Select SO035.Custid,SO035.AccountNo,SO035.ServiceType From SO035" & strTableSO007 & " Where " & strChoose35 & _
                                   ") SO033" & strTableSO004 & strChooseSO002 & strWhereSO004 & " Group by SO002A.CardName"
                If chkCancel.Value = 1 Then strSubQry(5) = strSubQry2(5)
                
                '�Ȥ����O
                strSubQry(6) = "SELECT CD004.Description,Sum(SO033.ShouldAmt) SumShouldAmt,Count(*) Counts,Count(Distinct SO033.Custid) CountCustid,Count(Distinct SO033.BillNo) CountBillNo " & _
                                     " From(" & _
                                         "Select SO033.ShouldAmt,SO033.Custid,SO033.BillNo,SO033.ServiceType,SO033.ClassCode,SO033.AccountNo From SO033" & strTableSO007 & " Where " & strChoose & _
                                   ") SO033,CD004" & strTableSO004 & IIf(strChooseSO002 = "", " Where ", strChooseSO002 & " And ") & "SO033.ClassCode=CD004.CodeNo " & strWhereSO004 & " Group by CD004.Description "
                                   
                strSubQry2(6) = "SELECT CD004.Description,Sum(SO033.ShouldAmt) SumShouldAmt,Count(*) Counts,Count(Distinct SO033.Custid) CountCustid,Count(Distinct SO033.BillNo) CountBillNo " & _
                                     " From(" & _
                                         "Select SO033.ShouldAmt,SO033.Custid,SO033.BillNo,SO033.ServiceType,SO033.ClassCode,SO033.AccountNo From SO033" & strTableSO007 & " Where " & strChoose & _
                                         " Union all Select SO035.ShouldAmt,SO035.Custid,SO035.BillNo,SO035.ServiceType,SO035.ClassCode,SO035.AccountNo From SO035" & strTableSO007 & " Where " & strChoose35 & _
                                   ") SO033,CD004" & strTableSO004 & IIf(strChooseSO002 = "", " Where ", strChooseSO002 & " And ") & "SO033.ClassCode=CD004.CodeNo " & strWhereSO004 & " Group by CD004.Description "
                If chkCancel.Value = 1 Then strSubQry(6) = strSubQry2(6)
                '�j�ӽs��
                strSubQry(7) = "SELECT SO017.Name,Sum(SO033.ShouldAmt) SumShouldAmt,Count(*) Counts,Count(Distinct SO033.Custid) CountCustid,Count(Distinct SO033.BillNo) CountBillNo,SO017.MduId  " & _
                                     " From(" & _
                                         "Select SO033.ShouldAmt,SO033.Custid,SO033.BillNo,SO033.ServiceType,SO033.MduId,SO033.AccountNo From SO033" & strTableSO007 & " Where " & strChoose & _
                                   ") SO033,SO017" & strTableSO004 & IIf(strChooseSO002 = "", " Where ", strChooseSO002 & " And ") & "SO033.MduId=SO017.MduId " & strWhereSO004 & " Group by SO017.Name,SO017.MduId  "
                                   
                strSubQry2(7) = "SELECT SO017.Name,Sum(SO033.ShouldAmt) SumShouldAmt,Count(*) Counts,Count(Distinct SO033.Custid) CountCustid,Count(Distinct SO033.BillNo) CountBillNo,SO017.MduId  " & _
                                     " From(" & _
                                         "Select SO033.ShouldAmt,SO033.Custid,SO033.BillNo,SO033.ServiceType,SO033.MduId,SO033.AccountNo From SO033" & strTableSO007 & " Where " & strChoose & _
                                         " Union all Select SO035.ShouldAmt,SO035.Custid,SO035.BillNo,SO035.ServiceType,SO035.MduId,SO035.AccountNo From SO035" & strTableSO007 & " Where " & strChoose35 & _
                                   ") SO033,SO017" & strTableSO004 & IIf(strChooseSO002 = "", " Where ", strChooseSO002 & " And ") & "SO033.MduId=SO017.MduId " & strWhereSO004 & " Group by SO017.Name,SO017.MduId  "
                
                If chkCancel.Value = 1 Then strSubQry(7) = strSubQry2(7)
                
                If CreateSubView2(strSubQry()) = False Then Exit Sub
                For intFor = 0 To UBound(strSubQry)
                    strSubQry(intFor) = "Select * From " & strSubViewName(intFor) & " V"
                Next
                Call PrintRpt2(GetPrinterName(5), RptName("SO5210", 3), , "�C�를ú�O�Ȥ�J�`�� [SO5210A]", strsql, strChooseString, , True, , , , GiPaperLandscape, , strSubQry(0), strSubQry(1), strSubQry(2), strSubQry(3), strSubQry(4), , , , , , , strSubQry(5), strSubQry(6), strSubQry(7))
            End If
        End If
    End If
    CloseRecordset rsTmp
    blnExcel = False
    On Error Resume Next
    gcnGi.Execute "Drop View " & strViewName
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subPrint")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGim
    Call subGil
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
    cboAddress.ListIndex = 0
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
    Call SetgiMulti(gimFaci, "CodeNo", "Description", "CD022", "�]�ƥN�X", "�]�ƦW��")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "��F�ϥN�X", "��F�ϦW��")
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "���O���إN�X", "���O���ئW��")
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "���O�覡�N�X", "���O�覡�W��")
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "�Ȥ����O�N�X", "�Ȥ����O�W��")
    Call SetgiMulti(gimCustStatus, "CodeNo", "Description", "CD035", "�Ȥ᪬�A�N�X", "�Ȥ᪬�A�W��")
    Call SetgiMulti(gimMduId, "MduId", "Name", "SO017", "�j�ӽs��", "�j�ӦW��")
    Call SetgiMulti(gimClctEn, "EmpNo", "EmpName", "CM003", "���O���N��", "���O���W��")
    '���D2595 �s�W�즬�O�H��,���O�ϱ��� 2006/7/17
    Call SetgiMulti(gimOldClctEn, "EmpNo", "EmpName", "CM003", "�즬�O���N��", "�즬�O���W��")
    Call SetgiMulti(gimClctArea, "CodeNo", "Description", "CD040", "���O�ϥN��", "���O�ϦW��")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "�A�ȰϥN�X", "�A�ȰϦW��")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "��D�N�X", "��D�W��")
    Call SetgiMulti(gimCancelCode, "CodeNo", "Description", "CD051", "�@�o��]�N�X", "�@�o��]�W��")
    Call SetgiMulti(gimUCCode, "CodeNo", "Description", "CD013", "������]�N�X", "������]�W��")
    Call SetgiMultiAddItem(gimBillType, "B,I,P,M,T", "���O��,�˾���,�������,���׳�,�{�ɦ��O��")
    Call SetgiMulti(gimBankCode, "CodeNo", "Description", "CD018", "�Ȧ�O�N�X", "�Ȧ�O�W��")
    Call SetgiMulti(gimCardCode, "CodeNo", "Description", "CD037", "�H�Υd�����N�X", "�H�Υd�����W��")
    Call SetgiMultiAddItem(gimOrder, "A,B,C,D,E,F,G,H,I", "�������,�a�},�Ȥ�s��,���O���N��,���O�渹,�A�Ȱ�,�Ȥ����O,���ʮɶ�,���O����")
    '#4289 �W�[�˾����u���O By Kin 2008/12/17
    Call SetgiMulti(gimInstall, "CodeNo", "Description", "CD005", "�˾����u���O�N�X", "�˾����u���O�W��")
    '#5683 �W�[ú�I���O By Kin 2010/08/18
    'Call SetgiMultiAddItem(gimPayType, "0,1", "�w�I��,�{�I��", "�N�X", "�W��")
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
    'Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5210A)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  Call DropTMPVIEW(strViewName, strSubViewName())
  gcnGi.Execute "Drop Table " & strViewName
  gcnGi.Execute "Drop Table " & strViewName1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5210A
End Sub


Private Sub gdaCreateTime1_GotFocus()
  On Error GoTo ChkErr
    If gdaCreateTime1.GetValue = "" Then gdaCreateTime1.SetValue (RightDate)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaCreateTime1_GotFocus")
End Sub

Private Sub gdaCreateTime2_GotFocus()
  On Error GoTo ChkErr
    If gdaCreateTime1.GetValue = "" Or gdaCreateTime2.GetValue = "" Then gdaCreateTime2.SetValue (gdaCreateTime1.GetValue)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaCreateTime2_GotFocus")
End Sub

Private Sub gdaCreateTime2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaCreateTime1, gdaCreateTime2)
  gdaCreateTime2_GotFocus
End Sub

Private Sub gdaShouldDate1_GotFocus()
  On Error GoTo ChkErr
    If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetValue (RightDate)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaShouldDate1_GotFocus")
End Sub

Private Sub gdaShouldDate2_GotFocus()
  On Error GoTo ChkErr
    If gdaShouldDate1.GetValue = "" Or gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue (gdaShouldDate1.GetValue)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaShouldDate2_GotFocus")
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaShouldDate1, gdaShouldDate2)
End Sub

Private Sub gdtUpdTime1_GotFocus()
  On Error Resume Next
  If gdtUpdTime1.GetValue = "" Then gdtUpdTime1.SetValue (RightDate)
End Sub

Private Sub gdtUpdTime2_GotFocus()
  On Error Resume Next
  If gdtUpdTime2.GetValue = "" Or gdtUpdTime1.GetValue = "" Then gdtUpdTime2.SetValue GetDayLastTime(gdtUpdTime1.GetValue(True))
End Sub

Private Sub gdtUpdTime2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdtUpdTime1, gdtUpdTime2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gimServiceType)
'    gilServiceType.ListIndex = 1
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimMduId, , gilCompCode.GetCodeNo
    GiMultiFilter gimClctEn, , gilCompCode.GetCodeNo
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimCancelCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimBankCode, , gilCompCode.GetCodeNo
    
    'GiMultiFilter gimUCCode, , gilCompCode.GetCodeNo
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

Private Function subCreateTable() As Boolean
  Dim rsTmp As New ADODB.Recordset
  Dim rsTmpGroup As New ADODB.Recordset
  Dim intFor As Integer
  Dim strField As String
  Dim strFieldFaci As String '�]�ƪ����
  Dim strView As String
  Dim strWhere As String
  Dim strTable As String
  Dim strWhereSO003 As String
  Dim strField2 As String, strView2 As String, strWhere2 As String
  Dim strCardExpDate As String
  Dim strCustName As String
  Dim strTel1 As String
  Dim strPrivAccount As String
    On Error Resume Next
      gcnGi.Execute "Drop Table " & strViewName
      gcnGi.Execute "Drop Table " & strViewName1
    On Error GoTo ChkErr
    strView = ""
    strFieldFaci = ""
    strWhere = ""
    strTable = ""
    strWhereSO004 = ""
    strTableSO004 = ""

    strViewName = GetTmpViewName
    strViewName1 = GetTmpViewName
    '#2810 �ץXExcel�ɼW�[�H�Υd�����A���N�H�Υd�����榡�� By Kin 2007/05/11
    strCardExpDate = "DeCode(length(SO002A.CardExpDate),5,substr(SO002A.CardExpDate,2,4) || '0' || substr(SO002A.CardExpDate,1,1),substr(SO002A.CardExpDate,3,4) ||  substr(SO002A.CardExpDate,1,2)) As CardExpDate "
    
    If strChooseSO004 <> "" Then
          strTableSO004 = ",SO004"
          '�O�_�C�L�Ȥ�]�ƶ���
          
          If chkFaciName.Value = 1 Then
              '�ӫȤ�]�ƥ��`�ϥΤ�
              If chkFaciUse.Value = 1 Then
                strWhereSO004 = " And SO033.Custid=SO004.Custid And SO033.FaciSeqNo=SO004.SEQNo And SO004.InstDate is Not Null And SO004.PRDate is Null " & strChooseSO004
              Else
              '#2810 ���ծɵo�{��Bug " And SO033.Custid=SO004.Custid SO033.FaciSeqNo=SO004.SEQNo " & strChooseSO004 �����֤F�@��And By Kin 2007/05/11
                strWhereSO004 = " And SO033.Custid=SO004.Custid And SO033.FaciSeqNo=SO004.SEQNo " & strChooseSO004
              End If
          Else
              strWhereSO004 = " And SO033.Custid=SO004.Custid " & strChooseSO004
          End If
    End If
    If strChooseSO002 <> "" Then
        strTable = strTable & ",SO002"
        strWhere = strWhere & " AND SO002.Custid=SO033.Custid And SO002.ServiceType=SO033.ServiceType And SO002.CompCode=SO033.CompCode And " & strChooseSO002
    End If
    If strChooseSO002A <> "" Then
        strTable = strTable & ",SO002A"
        strWhere = strWhere & " And SO002A.AccountNo=SO033.AccountNo And SO002A.Custid=SO033.Custid And SO002A.CompCode=SO033.CompCode AND " & strChooseSO002A
    End If
    
    '#2810 �p�G���U����,�ݭn�[�h��SO003,�ҥHSO033��Where����PTable�����ݧ�� By Kin 2007/05/10
    If strChooseSO003 <> "" Then
        strChooseSO003 = " And Exists (Select * from SO003 Where " & strChooseSO003 & " And SO003.CompCode=SO033.CompCode And SO003.ServiceType=SO033.ServiceType" & _
                                                    " And SO003.CustId=SO033.CustId And SO003.CitemCode=SO033.CitemCode And SO003.FaciSeqNo=SO033.FaciSeqNo )"
        strChoose = strChoose & strChooseSO003
    End If
    
    '���NstrChoose������s�J,�]��SO035�ɤ��ݧP�_UCCode
    strChoose35 = Replace(strChoose, "SO033", "SO035")
    If gimUCCode.GetQryStr <> "" Then
          Call subAnd("SO033.UCCode " & gimUCCode.GetQryStr)
    Else
    '���D��2808 �@�o�o���@�֭p��ɡ@�X�J�`��ɸ�ƻݶ]�ܤ[�~�X�ӡA�ӥB��ƥX�Ӫ������T
        If chkCancel = 0 Then
          Call subAnd("SO033.UCCode is not null")
        Else
            'Call subAnd("SO033.UCCode is not null or (SO033.CancelFlag =1 and SO033.UCCode is null) ")
            Call subAnd("(SO033.UCCode is not null or SO033.CancelFlag =1 and SO033.UCCode is null) ")
        End If
    End If
    '#4289 �W�[�P�_�O�_�n���C�^�渹 By Kin 2008/12/17
    '#4289 �令��optional�P�_ By Kin 2009/02/02
    If optMediaYes.Value Then
        Call subAnd(" SO033.MediaBillNo is Not Null ")
    End If
    If optMediaNo.Value Then
        Call subAnd(" SO033.MediaBillNo is Null ")
    End If
'    If chkMediabillNo.Value = 1 Then
'        Call subAnd("SO033.MediaBillNo is Not Null")
'    End If
    
    
    If optDail.Value Then
        '�ˬd�a�}�̾�
        Select Case cboAddress.ListIndex
            Case 0
                 strTable = strTable & ",SO014"
                strWhere = strWhere & " And So033.AddrNo=SO014.AddrNo And SO033.CompCode=SO014.CompCode "
            Case 1
                strTable = strTable & ",SO014,So001"
                strWhere = strWhere & "  And SO001.CUSTID=SO033.CUSTID And So001.InstAddrNo=SO014.AddrNo And SO001.CompCode=SO014.CompCode "
            Case 2
                strTable = strTable & ",SO014,So001"
                strWhere = strWhere & "  And SO001.CUSTID=SO033.CUSTID And So001.ChargeAddrNo=SO014.AddrNo And SO001.CompCode=SO014.CompCode "
            Case Else
                strTable = strTable & ",SO014"
                strWhere = strWhere & " And So033.AddrNo=SO014.AddrNo And SO033.CompCode=SO014.CompCode "
        End Select
        For intFor = 1 To 10
             strFieldFaci = strFieldFaci & ",'                    ' Faciname" & intFor & _
                                           ",'                    ' BuyName" & intFor & _
                                           ",To_Date(Null) Prdate" & intFor & _
                                           ",'                    ' Facisno" & intFor & " " & _
                                           ",'                    ' CMBaudRate" & intFor & " "
        Next
        '#2810 ���ծɵo�{��Bug,�p�G���Ŀ�O�_�C�L�Ȥ�]�ƶ���,�|��SO033.Faciseqno=SO004.SEQNO,���D�諸���èS��FaciSeqNO By Kin 2007/05/11
        '#2810 �p�G���ӽФH�m�W�nShow�ӽФH�m�W�P�ӽФH�q��,�ҥH�h�[SO004.DeclarantName,SO004.ContTel���,�A��Report�ۦ�P�_�n��ܨ��@�� By Kin 2007/05/10
        '#3417 �h�[SO033.NOTE���
        '#4289 �n�W�[BpCode�PBPName,�p�G�˾����O���D�諸��,�n�W�[SO007 By Kin 2008/12/17
        '#5683 �W�[ú�I���OPayKind By Kin 2010/08/18
        strField = "SO033.CUSTID,SO033.BILLNO,SO033.RealStartDate,SO033.RealStopDate," & _
                    "SO033.SHOULDDATE,SO033.CitemCode,SO033.CitemName,SO033.ShouldAmt," & _
                    "SO033.UCNAME," & _
                    "SO033.CLCTNAME,SO033.AreaCode,SO033.ClctEn,SO033.ServCode," & _
                    "SO033.GUINo,SO033.PrtSNo,SO033.ClctAreaCode,SO033.ServiceType," & _
                    "SO033.AddrNo,SO033.ClassCode," & _
                    "SO033.AccountNo,SO033.CompCode,SO033.MediaBillNo,SO033.RealPeriod," & _
                    "SO033.UpdTime,SO033.OldClctName,SO004.DeclarantName," & _
                    "SO004.ContTel,SO033.FaciSEQNO,SO033.NOTE," & _
                    "SO033.BpCode,SO033.BpName,Nvl(SO033.PayKind,0) PayKind "
        strView = "Select " & strField & " From SO033,SO004" & strTableSO007 & " Where "
        
        '2004/10/27 ���D��1201 �@�u��ڤ@�_�έp,�[�JSO035 -- for Liga
        '#2810 �h�W�[�U��������� By Kin 2007/05/10
        If chkCancel.Value = 1 Then
            strView2 = Replace(strView, "SO033", "SO035")
            strView = strView & strChoose & " Union All " & strView2 & strChoose35
        Else
            strView = strView & strChoose
        End If
        '*****************************************************************************************************************************
        '#3429 �ץX�ɧP�_�v���O�_�i�H������ܱb���A��ӬO�b�ץXExcel�ɧP�_�A�]���ץXExcel�令�n�g�LReport�A�ҥH��ӳo�̧P�_ By Kin 2007/08/17
        If GetUserPriv("SO1100G", "(SO1100G9)") Then
            strPrivAccount = "SO033.AccountNo AccountNo"
        Else
            strPrivAccount = "Decode(SO033.AccountNo,Null,SO033.AccountNo,SubStr(SO033.AccountNo,1,5) || 'XXX' || " & _
                           "SubStr(SO033.AccountNo,9,2) || 'XX' || SubStr(SO033.AccountNo,Length(SO033.AccountNo)-3,4)) AccountNo"
        End If
        '******************************************************************************************************************************
        '#2810 �ץXExcel�n�W�[�b���B�H�Υd�����B�ӽФH�m�W�B�ӽФH�q�� By Kin 2007/05/10
        strCustName = "Decode(SO033.DeclarantName,Null,SO001.CustName,SO033.DeclarantName) CustName "
        strTel1 = "Decode(SO033.DeclarantName,Null,SO001.TEL1,SO033.ContTel) TEL1 "
        '**************************************************************************************************************************************************************************************
        '#3429 �ץXExcel�令�g��Rpt�A�ҥH�վ�View���y�k�A�O�_���v��������ܱb�A�����bView�B�z��(strPrivAccount) By Kin 2007/08/17
        '#3417 �h�W�[SO033.NOTE���
        '#3705 �h�W�[SO014.CircuitNo�����s����� By Kin 2008/02/15
        '#4289 �n�W�[BpCode�BBpName By Kin 2008/12/17
        '#5683 �W�[ú�I���O PayKind By Kin 2010/08/13
        strView = "SELECT SO033.CUSTID,SO033.BILLNO,SO033.RealStartDate,SO033.RealStopDate," & _
                  "SO033.SHOULDDATE,SO033.CitemCode,SO033.CitemName," & _
                  "SO033.ShouldAmt,SO033.UCNAME," & _
                  "SO033.CLCTNAME," & strCustName & "," & strTel1 & ",SO001.TEL2," & _
                  "SO001.TEL3,CD004.Description,SO014.ADDRESS," & _
                  "SO014.SERVNAME,SO014.AreaCode,SO014.AreaName,SO033.ClctEn," & _
                  "SO014.ServCode,SO033.GUINo,SO014.AddrSort,SO033.PrtSNo," & _
                  "SO033.ClctAreaCode,SO014.ClctAreaName,SO002.CustStatusName," & _
                  "SO014.MDUName,SO002A.CardName,SO033.MediaBillNo," & _
                  "SO033.RealPeriod,SO033.UpdTime,SO033.OldClctName," & strPrivAccount & "," & strCardExpDate & ",SO033.DeclarantName," & _
                  "SO033.ContTel,SO033.NOTE,SO014.CircuitNo," & _
                  "SO033.BpCode,SO033.BpName,SO033.PayKind " & strFieldFaci & _
                      "From (" & strView & _
                      ") SO033,SO001,SO014,CD004,SO002,SO002A " & strTableSO004 & _
                  " Where SO002.Custid=SO033.Custid And SO002.ServiceType=SO033.ServiceType And SO002.CompCode=SO033.CompCode And SO001.CUSTID=SO033.CUSTID" & _
                  " And SO001.CompCode=SO033.CompCode " & strWhere & _
                  " And SO033.ClassCode=CD004.CodeNo(+) And SO033.AccountNo=SO002A.AccountNo(+) And SO002A.Custid(+)=SO033.Custid And SO002A.CompCode(+)=SO033.CompCode" & _
                  IIf(strChooseSO002 = "", "", " And " & strChooseSO002) & IIf(strChooseSO002A = "", "", " And " & strChooseSO002A) & strWhereSO004
        '****************************************************************************************************************************************************************************************
        If Len(Trim(txtCustId)) > 0 Then
            strCustId = Replace(txtCustId.Text, ",", "','")
            strView = strView & " AND SO033.Custid IN ('" & Trim(strCustId) & "') "
        End If
        
        gcnGi.Execute "Create Table " & strViewName & " as (" & strView & ")"
        gcnGi.Execute "ALTER TABLE " & strViewName & " NOLOGGING"
        SendSQL strView, True
        
        If chkFaciName.Value = 1 Then
                    
            If chkFaciUse.Value = 1 Then
                strView = "Select distinct SO033.Custid,SO033.BillNo,SO004.FaciName,SO004.BuyName,SO004.PRDate,SO004.FaciSno,SO004.CMBaudRate From SO033,SO004" & strTable & " Where SO033.Custid=SO004.Custid And  SO033.FaciSeqNo=SO004.SEQNo And SO004.InstDate is Not Null And SO004.PRDate is Null And " & strChoose & strWhere & strChooseSO004
            Else
                strView = "Select distinct SO033.Custid,SO033.BillNo,SO004.FaciName,SO004.BuyName,SO004.PRDate,SO004.FaciSno,SO004.CMBaudRate From SO033,SO004" & strTable & " Where SO033.Custid=SO004.Custid And SO033.FaciSeqNo=SO004.SEQNo And " & strChoose & strWhere & strChooseSO004
            End If
            gcnGi.Execute "Create Table " & strViewName1 & " As (" & strView & ")"
            gcnGi.Execute "ALTER TABLE " & strViewName1 & " NOLOGGING"
            
            If Not GetRS(rsTmpGroup, "Select Custid,BillNo,Count(*) as intCount From " & strViewName1 & " Group by Custid,BillNo") Then Exit Function
            
            While Not rsTmpGroup.EOF
                If Not GetRS(rsTmp, "Select * From " & strViewName1 & " Where Billno='" & rsTmpGroup.Fields("BillNo") & "' And Custid=" & rsTmpGroup.Fields("Custid")) Then Exit Function
                With rsTmp
                    For intFor = 1 To .RecordCount
                        gcnGi.Execute "Update " & strViewName & " Set " & _
                                      "FaciName" & intFor & "=" & GetNullString(.Fields("FaciName")) & "," & _
                                      "BuyName" & intFor & "=" & GetNullString(.Fields("BuyName")) & "," & _
                                      "PrDate" & intFor & "=" & GetNullString(.Fields("PrDate"), giDateV) & "," & _
                                      "FaciSno" & intFor & "=" & GetNullString(.Fields("FaciSno")) & "," & _
                                      "CMBaudRate" & intFor & "=" & GetNullString(.Fields("CMBaudRate"), giStringV, giOracle) & _
                                      " Where Billno='" & rsTmpGroup("BillNo") & "' And Custid=" & rsTmpGroup.Fields("Custid")
                        .MoveNext
                    Next
                End With
                rsTmpGroup.MoveNext
            Wend
            
        End If

    Else
        If strWhere <> "" Then
            strWhere2 = Replace(Mid(strWhere, 5), "SO033", "SO035")
            strWhere = " WHERE " & Mid(strWhere, 5) & IIf(strChoose = "", "", " AND " & strChoose)
            strWhere2 = " WHERE " & strWhere2 & IIf(strChoose35 = "", "", " AND " & strChoose35)
        Else
            strWhere = IIf(strChoose = "", "", " WHERE " & strChoose)
            strWhere2 = IIf(strChoose35 = "", "", " WHERE " & strChoose35)
        End If
        If chkCancel.Value = 0 Then
            strView = "SELECT Count(Distinct SO033.Custid) intCount From SO033" & strTable & strTableSO004 & strTableSO007 & strWhere & strWhereSO004
        Else
            strView = "SELECT SO033.Custid From SO033" & strTable & strTableSO004 & strTableSO007 & strWhere & strWhereSO004
            '#2810 ���ծɵo�{��Bug,���Ŀ�@�o��ڬO�_�@�ֲέp��,�S����ܳ]�ƶ���,�b��SQL�y�k��,Union SO035��strWhereSO004��Table�S��L�� By Kin 2007/05/15
            strView2 = "SELECT SO035.Custid From SO035" & strTable & strTableSO004 & strWhere2 & Replace(strWhereSO004, "SO033", "SO035")
            strView = "SELECT COUNT(Distinct CUSTID) intCount from (" & strView & " Union All " & strView2 & ")"
        End If
       '2004/10/27 ���D��1201 �@�u��ڤ@�_�έp,�[�JSO035 -- for Liga
'       If chkCancel.Value = 1 Then strView = "select sum(intCount) intCount from (" & strView & " Union All " & strView2 & ")"
       SendSQL strView, True
        gcnGi.Execute "Create Table " & strViewName & " as(" & strView & ")"
        gcnGi.Execute "ALTER TABLE " & strViewName & " NOLOGGING"
    End If
      CloseRecordset rsTmp
      CloseRecordset rsTmpGroup
      subCreateTable = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateTable")
End Function

Private Sub gimServiceType_Change()
  On Error GoTo ChkErr
    GiMultiFilter gimFaci, gimServiceType.GetQryStr, , True
    GiMultiFilter gimCitemCode, gimServiceType.GetQryStr, , True
    GiMultiFilter gimCMCode, gimServiceType.GetQryStr, , True
    GiMultiFilter gimClassCode, gimServiceType.GetQryStr, , True
    GiMultiFilter gimUCCode, gimServiceType.GetQryStr, , True
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gimServiceType_Change")
End Sub

Private Sub gimServiceType_GotFocus()
  On Error Resume Next
      vsl1.Value = 0
End Sub

Private Sub gimOrder_GotFocus()
  On Error Resume Next
    vsl1.Value = 100
End Sub

Private Sub GiYM1_Change()

End Sub

Private Sub GiYM1_GotFocus()

End Sub

Private Sub gymClctYM1_GotFocus()
  On Error GoTo ChkErr
    If gymClctYM1.GetValue = "" Then gymClctYM1.SetValue (RightDate)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gymClctYM1_GotFocus")
End Sub

Private Sub gymClctYM2_GotFocus()
On Error GoTo ChkErr
    If gymClctYM1.GetValue = "" Or gymClctYM2.GetValue = "" Then gymClctYM2.SetValue (gymClctYM1.GetValue)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gymClctYM2_GotFocus")
End Sub

Private Sub gymClctYM2_Validate(Cancel As Boolean)
On Error Resume Next
  Cancel = ChkDate2(gymClctYM1, gymClctYM2)
End Sub

Private Sub optDail_Click()
  On Error Resume Next
  cmdExcel.Enabled = True
End Sub

Private Sub optPrtSNoNo_Click()
  On Error Resume Next
  txtPrtSNo1.Text = ""
  txtPrtSNo2.Text = ""
End Sub

Private Sub optSummaries_Click()
  On Error Resume Next
  cmdExcel.Enabled = False
End Sub

Private Sub txtPrtSNo1_KeyPress(KeyAscii As Integer)
  On Error GoTo ChkErr
      'MsgBox KeyAscii
    If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtPrtSNo1_KeyPress")
End Sub

Private Sub txtPrtSNo2_GotFocus()
  On Error GoTo ChkErr
    If Len(txtPrtSNo2.Text) = 0 Then txtPrtSNo2.Text = txtPrtSNo1.Text
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtPrtSNo2_GotFocus")
End Sub

Private Sub txtPrtSNo2_KeyPress(KeyAscii As Integer)
  On Error GoTo ChkErr
      'MsgBox KeyAscii
    If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtPrtSNo2_KeyPress")
End Sub

Private Sub vsl1_Change()
  On Error Resume Next
    If vsl1.Value = 0 Then
        fraMulti.Top = 20
    ElseIf vsl1.Value = 100 Then
        fraMulti.Top = -(pic2.Height) + 480
    End If
End Sub
Private Sub subInsertMDB(rsTmp As Recordset)
  On Error GoTo ChkErr
    Dim strField As String
    Dim i As Long
    For i = 0 To rsTmp.Fields.Count - 1
        strField = strField & rsTmp.Fields(i).Name & ","
    Next i
    strField = Left(strField, Len(strField) - 1)
    rsTmp.MoveFirst
    cnn.BeginTrans
    While Not rsTmp.EOF
        cnn.Execute "Insert Into SO5210AE (" & strField & _
                    ") Values(" & rsTmp("CustId") & ",'" & rsTmp("BillNO") & "'," & _
                              "'" & rsTmp("CUSTNAME") & "',#" & rsTmp("SHOULDDATE") & "#," & _
                              "'" & rsTmp("TEL1") & "','" & rsTmp("TEL2") & "','" & rsTmp("TEL3") & "'," & _
                              "'" & rsTmp("Description") & "'," & rsTmp("CitemCode") & "," & _
                              "'" & rsTmp("CitemName") & "'," & rsTmp("RealPeriod") & "," & rsTmp("ShouldAmt") & "," & _
                              "#" & rsTmp("RealStartDate") & "#,#" & rsTmp("RealStopDate") & "#," & _
                              "'" & rsTmp("GUINo") & "','" & rsTmp("UCNAME") & "','" & rsTmp("CLCTNAME") & "'," & _
                              "'" & rsTmp("PrtSNo") & "','" & rsTmp("ADDRESS") & "','" & rsTmp("SERVNAME") & "'," & _
                              "'" & rsTmp("AreaName") & "','" & rsTmp("CustStatusName") & "','" & rsTmp("MDUName") & "'," & _
                              "'" & rsTmp("CardName") & "','" & rsTmp("OldClctName") & "','" & rsTmp("Account") & "'," & _
                              "'" & rsTmp("CardExpDate") & "'," & rsTmp("ServCode") & ",'" & rsTmp("AddrSort") & "'," & _
                              "'" & rsTmp("UpdTime") & "','" & rsTmp("ClctAreaCode") & "'," & rsTmp("AreaCode") & ")"
       rsTmp.MoveNext
    Wend
    cnn.CommitTrans
  Exit Sub

ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "subInsertMDB")
End Sub
Private Sub toExcel(ByVal strsql As String)
  On Error GoTo ChkErr
    Dim rsExcel As New ADODB.Recordset
      RptToTxt RptName("SO5210", "E"), , strsql, , _
               , , strGroupName, , , Environ("Temp") & "\SO5210", False
      If Not Get_RS_From_Txt(Environ("Temp"), "SO5210.txt", rsExcel) Then blnExcel = False: Exit Sub
      Call UseProperty(rsExcel, "�C�를ú�O�Ȥ���Ӫ�", "�Ĥ@��")
      blnExcel = False
      CloseRecordset rsExcel
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "toExcel")
End Sub

