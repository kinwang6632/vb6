VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{99A4DA9D-8D59-11D3-8311-0080C8453DDF}#11.4#0"; "GiAddress1.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{84CE7734-5A37-11D3-8311-0080C8453DDF}#7.5#0"; "GiAddress2.ocx"
Begin VB.Form frmSO114FA 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�o�����Y���@[SO114FA]"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11985
   Icon            =   "SO114FA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.Frame frmSearch 
      Caption         =   "�M��"
      Height          =   1245
      Left            =   60
      TabIndex        =   52
      Top             =   120
      Width           =   11835
      Begin VB.Frame fraAPI 
         Height          =   1545
         Left            =   10950
         TabIndex        =   69
         Top             =   150
         Visible         =   0   'False
         Width           =   120
         Begin VB.TextBox txtUrlResponse 
            Height          =   375
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   71
            Top             =   720
            Width           =   8475
         End
         Begin VB.TextBox txtUrl 
            Height          =   405
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   240
            Width           =   8475
         End
      End
      Begin VB.CommandButton cmdLog 
         Caption         =   "���ʰO��"
         Height          =   315
         Left            =   2400
         TabIndex        =   62
         Top             =   840
         Width           =   1275
      End
      Begin VB.ComboBox cboChargeTitleQ 
         Height          =   300
         ItemData        =   "SO114FA.frx":0442
         Left            =   900
         List            =   "SO114FA.frx":0444
         TabIndex        =   0
         Top             =   210
         Width           =   2325
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "F3.�M��"
         Height          =   315
         Left            =   1350
         TabIndex        =   6
         Top             =   840
         Width           =   1005
      End
      Begin VB.TextBox txtInvSeqNoQ2 
         Alignment       =   1  '�a�k���
         Height          =   315
         Left            =   7680
         MaxLength       =   8
         TabIndex        =   3
         Top             =   180
         Width           =   1125
      End
      Begin VB.TextBox txtInvSeqNoQ1 
         Alignment       =   1  '�a�k���
         Height          =   315
         Left            =   6330
         MaxLength       =   8
         TabIndex        =   2
         Top             =   180
         Width           =   1125
      End
      Begin VB.TextBox txtInvTitleQ 
         Height          =   315
         Left            =   900
         MaxLength       =   50
         TabIndex        =   4
         Top             =   540
         Width           =   7905
      End
      Begin VB.CheckBox chkChargeAddr 
         Caption         =   "���O�a�}"
         Height          =   315
         Left            =   60
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   5
         Top             =   840
         Width           =   1245
      End
      Begin MSMask.MaskEdBox mskInvNoQ 
         Height          =   315
         Left            =   4170
         TabIndex        =   1
         Top             =   210
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "99999999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "~"
         Height          =   195
         Left            =   7530
         TabIndex        =   57
         Top             =   240
         Width           =   195
      End
      Begin VB.Label lblInvSeqNoQ 
         Caption         =   "�y�����d��"
         Height          =   225
         Left            =   5370
         TabIndex        =   56
         Top             =   270
         Width           =   915
      End
      Begin VB.Label lblInvTitleQ 
         Caption         =   "�o�����Y"
         Height          =   225
         Left            =   90
         TabIndex        =   55
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lblInvNoQ 
         Caption         =   "�Τ@�s��"
         Height          =   225
         Left            =   3390
         TabIndex        =   54
         Top             =   270
         Width           =   795
      End
      Begin VB.Label lblChargeTitleQ 
         Caption         =   "��  ��  �H"
         Height          =   225
         Left            =   90
         TabIndex        =   53
         Top             =   300
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdInvData 
      BackColor       =   &H80000004&
      Caption         =   "��w"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9630
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   31
      Top             =   4350
      Width           =   1155
   End
   Begin VB.Frame fraDetail 
      Height          =   2475
      Left            =   120
      TabIndex        =   47
      Top             =   6150
      Width           =   11715
      Begin TabDlg.SSTab sstDetail 
         Height          =   2085
         Left            =   120
         TabIndex        =   48
         Top             =   270
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   3678
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "�o�����Y���"
         TabPicture(0)   =   "SO114FA.frx":0446
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ggrData"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "�ϥΫȤ���"
         TabPicture(1)   =   "SO114FA.frx":0462
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ggrData2"
         Tab(1).ControlCount=   1
         Begin prjGiGridR.GiGridR ggrData 
            Height          =   1635
            Left            =   150
            TabIndex        =   29
            ToolTipText     =   "�Ы�����⦸,����o�����"
            Top             =   360
            Width           =   11115
            _ExtentX        =   19606
            _ExtentY        =   2884
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
         Begin prjGiGridR.GiGridR ggrData2 
            Height          =   1485
            Left            =   -74820
            TabIndex        =   30
            Top             =   420
            Width           =   11085
            _ExtentX        =   19553
            _ExtentY        =   2619
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
      End
   End
   Begin VB.Frame fraQuery 
      Height          =   7635
      Left            =   60
      TabIndex        =   33
      Top             =   1440
      Width           =   11865
      Begin VB.PictureBox picAddress 
         Height          =   3015
         Left            =   390
         ScaleHeight     =   2955
         ScaleWidth      =   11655
         TabIndex        =   49
         Top             =   5880
         Width           =   11715
         Begin PrjGiAddress2.GiAddress2 ga2Address 
            Height          =   2145
            Left            =   90
            TabIndex        =   50
            Top             =   0
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   3784
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TableName       =   ""
         End
      End
      Begin VB.TextBox txtChargeTitleQ 
         Height          =   315
         Left            =   5940
         MaxLength       =   50
         TabIndex        =   32
         Top             =   2460
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Frame fraData 
         Caption         =   "�s��"
         Height          =   4515
         Left            =   90
         TabIndex        =   34
         Top             =   210
         Width           =   11715
         Begin VB.TextBox txtA_CarrierId2 
            Height          =   405
            Left            =   7080
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   28
            Top             =   4020
            Width           =   3975
         End
         Begin VB.TextBox txtA_CarrierId1 
            Height          =   405
            Left            =   7080
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   27
            Top             =   3570
            Width           =   3975
         End
         Begin VB.TextBox txtLoveNum 
            Height          =   315
            Left            =   7320
            MaxLength       =   10
            TabIndex        =   17
            Top             =   1530
            Width           =   2265
         End
         Begin VB.TextBox txtCarrierId2 
            Height          =   405
            Left            =   1890
            MaxLength       =   50
            TabIndex        =   26
            Top             =   4020
            Width           =   3975
         End
         Begin VB.TextBox txtCarrierId1 
            Height          =   405
            Left            =   1890
            MaxLength       =   50
            TabIndex        =   25
            Top             =   3570
            Width           =   3975
         End
         Begin VB.CheckBox chkBillMailKind 
            Caption         =   "�q�l�b��"
            Height          =   285
            Left            =   6660
            TabIndex        =   14
            Top             =   1170
            Width           =   1050
         End
         Begin VB.ComboBox cboDenRec 
            Height          =   300
            ItemData        =   "SO114FA.frx":047E
            Left            =   1300
            List            =   "SO114FA.frx":0488
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   18
            Top             =   1860
            Width           =   2000
         End
         Begin VB.CheckBox chkPreInv 
            Caption         =   "�O�_�w�}�o��"
            Height          =   285
            Left            =   4530
            TabIndex        =   13
            Top             =   1170
            Width           =   1455
         End
         Begin VB.CheckBox chkStop 
            Caption         =   "����"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   2880
            Width           =   705
         End
         Begin PrjGiAddress1.GiAddress1 ga1ChargeAddress 
            Height          =   345
            Left            =   180
            TabIndex        =   20
            Top             =   2190
            Width           =   9075
            _ExtentX        =   16007
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
            AddressCaption  =   "���O�a�}"
         End
         Begin Gi_Date.GiDate gdaStopDate 
            Height          =   315
            Left            =   1950
            TabIndex        =   23
            Top             =   2820
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
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
         Begin VB.TextBox txtInvAddress 
            Height          =   315
            Left            =   3330
            MaxLength       =   50
            TabIndex        =   11
            Top             =   810
            Width           =   6255
         End
         Begin VB.TextBox txtInvTitle 
            Height          =   315
            Left            =   3330
            MaxLength       =   100
            TabIndex        =   9
            Top             =   480
            Width           =   6255
         End
         Begin VB.ComboBox cboInvoiceType 
            Height          =   300
            ItemData        =   "SO114FA.frx":04AA
            Left            =   1300
            List            =   "SO114FA.frx":04B7
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   8
            Top             =   450
            Width           =   1155
         End
         Begin VB.TextBox txtChargeTitle 
            Height          =   315
            Left            =   3330
            MaxLength       =   100
            TabIndex        =   7
            Top             =   150
            Width           =   6255
         End
         Begin MSMask.MaskEdBox mskInvNo 
            Height          =   315
            Left            =   1300
            TabIndex        =   10
            Top             =   780
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "99999999"
            PromptChar      =   "_"
         End
         Begin PrjGiAddress1.GiAddress1 ga1MailAddress 
            Height          =   345
            Left            =   180
            TabIndex        =   21
            Top             =   2520
            Width           =   9075
            _ExtentX        =   16007
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
            AddressCaption  =   "�l�H�a�}"
         End
         Begin prjGiList.GiList gilUseInv 
            Height          =   315
            Left            =   1300
            TabIndex        =   12
            Top             =   1140
            Width           =   3135
            _ExtentX        =   5530
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
            FldWidth1       =   800
            FldWidth2       =   2000
            F2Corresponding =   -1  'True
         End
         Begin Gi_Date.GiDate gdaApplyInvDate 
            Height          =   300
            Left            =   5400
            TabIndex        =   19
            Top             =   1830
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
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
         Begin prjGiList.GiList gilDenRecCode 
            Height          =   315
            Left            =   1305
            TabIndex        =   15
            Top             =   1500
            Width           =   2250
            _ExtentX        =   3969
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
            FldWidth1       =   405
            FldWidth2       =   1520
            F2Corresponding =   -1  'True
            FilterStop      =   -1  'True
         End
         Begin Gi_Date.GiDate gidDenRecDate 
            Height          =   300
            Left            =   5400
            TabIndex        =   16
            Top             =   1500
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin prjGiList.GiList gilCarrierTypeCode 
            Height          =   315
            Left            =   1890
            TabIndex        =   24
            Top             =   3210
            Width           =   3330
            _ExtentX        =   5874
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
            BoxWidth        =   4000
            FldWidth1       =   1000
            FldWidth2       =   2000
            F2Corresponding =   -1  'True
            FilterStop      =   -1  'True
         End
         Begin VB.Label Label8 
            Caption         =   "�|���������X"
            Height          =   225
            Left            =   5940
            TabIndex        =   68
            Top             =   4110
            Width           =   1395
         End
         Begin VB.Label Label7 
            Caption         =   "�|��������X"
            Height          =   225
            Left            =   5940
            TabIndex        =   67
            Top             =   3720
            Width           =   1395
         End
         Begin VB.Label Label6 
            Caption         =   "���ؽX"
            Height          =   180
            Left            =   6690
            TabIndex        =   66
            Top             =   1590
            Width           =   780
         End
         Begin VB.Label Label5 
            Caption         =   "�@�q�ʸ������X"
            Height          =   225
            Left            =   210
            TabIndex        =   65
            Top             =   4140
            Width           =   1395
         End
         Begin VB.Label Label4 
            Caption         =   "�@�q�ʸ�����X"
            Height          =   225
            Left            =   210
            TabIndex        =   64
            Top             =   3720
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "�@�q�ʸ������O���X"
            Height          =   225
            Left            =   180
            TabIndex        =   63
            Top             =   3300
            Width           =   1695
         End
         Begin VB.Label lblDenRecCode 
            Caption         =   "���س��"
            Height          =   180
            Left            =   450
            TabIndex        =   61
            Top             =   1575
            Width           =   780
         End
         Begin VB.Label lblDenRecDate 
            Caption         =   "���ؤ��"
            Height          =   180
            Left            =   4635
            TabIndex        =   60
            Top             =   1590
            Width           =   780
         End
         Begin VB.Label lblApplyInvDate 
            Caption         =   "�ӽйq�l�o���ҩ��p���"
            Height          =   285
            Left            =   3360
            TabIndex        =   59
            Top             =   1920
            Width           =   2010
         End
         Begin VB.Label lblDenRec 
            Caption         =   "�o���}�ߺ���"
            Height          =   225
            Left            =   120
            TabIndex        =   58
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label lblUseInv 
            Caption         =   "�o���γ~"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   420
            TabIndex        =   51
            Top             =   1200
            Width           =   795
         End
         Begin VB.Label lblStopDate 
            Caption         =   "���Τ��"
            Height          =   225
            Left            =   1170
            TabIndex        =   46
            Top             =   2910
            Width           =   795
         End
         Begin VB.Label lblUpdEnValue 
            Caption         =   "96/11/25 �U��13:00"
            Height          =   225
            Left            =   6600
            TabIndex        =   45
            Top             =   2880
            Width           =   1695
         End
         Begin VB.Label lblUpdEn 
            Caption         =   "���ʤH��"
            Height          =   225
            Left            =   5700
            TabIndex        =   44
            Top             =   2880
            Width           =   795
         End
         Begin VB.Label lblUpdTimeValue 
            Caption         =   "96/11/25 �U��13:00"
            Height          =   225
            Left            =   3930
            TabIndex        =   43
            Top             =   2880
            Width           =   1695
         End
         Begin VB.Label lblUpdTime 
            Caption         =   "���ʮɶ�"
            Height          =   225
            Left            =   3150
            TabIndex        =   42
            Top             =   2880
            Width           =   795
         End
         Begin VB.Label lblInvAddress 
            Caption         =   "�o���a�}"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2550
            TabIndex        =   41
            Top             =   840
            Width           =   825
         End
         Begin VB.Label lblInvTitle 
            Caption         =   "�o�����Y"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2550
            TabIndex        =   40
            Top             =   510
            Width           =   825
         End
         Begin VB.Label lblInvNo 
            Caption         =   "�Τ@�s��"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   420
            TabIndex        =   39
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "�o������"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   420
            TabIndex        =   38
            Top             =   510
            Width           =   765
         End
         Begin VB.Label lblChargeTitle 
            Caption         =   "��  ��  �H"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2550
            TabIndex        =   37
            Top             =   210
            Width           =   825
         End
         Begin VB.Label lblInvSeqNoValue 
            Caption         =   "00000000"
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
            Left            =   1300
            TabIndex        =   36
            Top             =   210
            Width           =   825
         End
         Begin VB.Label lblInvSeqNo 
            Caption         =   "�y���s��"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   420
            TabIndex        =   35
            Top             =   210
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frmSO114FA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngEditMode As giEditModeEnu ' �O���ثe�b�s��B�s�W���˵��Ҧ�
Private WithEvents rsSO138 As ADODB.Recordset
Attribute rsSO138.VB_VarHelpID = -1
Private rsSO003 As New ADODB.Recordset
Private lngOldEditMode As giEditModeEnu
Private blnFrmShow As Boolean
Private ObjOwnerForm As Form
Private strAccountNo As String
Private WithEvents rsDefTmp As ADODB.Recordset
Attribute rsDefTmp.VB_VarHelpID = -1
Dim blnBeginTrans As Boolean
Dim lngCustID As Long
Dim strInvSeqNo As String
Dim strShowAbs As String
Private blnAdd As Boolean
Private blnEdit As Boolean
Private blnDele As Boolean
Private blnMutilChoice As Boolean
Private iSelectCount As Integer
Private strAccountOwner As String
Private blnStartEinv As Boolean, intDenRecPre As Integer '�o���}�ߺ��� �e�@�Ӫ��A��
Private intINVOICEKIND As Integer
Private blnNoChgApplyInvDate  As Boolean
Private blnNoChgInvoiceKind As Boolean
Private blnShowDenRec As Boolean
Private blnStartEinvoiceXML As Boolean
Private blnCanModifyCarrierId As Boolean
Private blnStartMobileAPI As Boolean
Private blnStartLovenumAPI As Boolean
Private strMobileAPIURL As String
Private strLovenumAPIURL As String
Private strAPIappID As String
Private Sub cboDenRec_Click()
    On Error Resume Next
'    If blnNoChgInvoiceKind Then
'        blnNoChgInvoiceKind = False
'        Exit Sub
'    End If
    
    If cboDenRec.ListIndex = 0 Then
        gdaApplyInvDate.Enabled = True
        '#5760 �n�a�J�t�Τ�� By Kin 2010/09/16
        '#5885 2011.01.18 by Corey �s�W���s��Ʈ�(�s�WSO002��)�A�o���}�ߺ���(InvoiceKind=0)�ɡA
                                    '�ӽйq�l�p����o�����(ApplyInvDate)�����ӹw�]��J�t�Τ�C
                                    
        '                          �ק�o���}�ߺ�����1�վ㬰0�ɡA�~�w�]�N�ӽйq�l�p����o�����(ApplyInvDate)�w�]���t�Τ�A
                                    '���M�ثe���s�W�N��J�O�����D���A�|��CSR�~�P�C
        'If intDenRecPre > 0 Then gdaApplyInvDate.SetValue Format(Now, "yyyymmdd")
        If EditMode = giEditModeView Then
            Exit Sub
        End If
        If EditMode = giEditModeEdit Then
            If Not blnNoChgApplyInvDate Then
                gdaApplyInvDate.SetValue Format(Now, "yyyymmdd")
            End If
        
        End If
    Else
        gdaApplyInvDate.Enabled = False
        
    End If
    
    
End Sub


Private Sub cboInvoiceType_Click()
  On Error Resume Next
    Dim blnCustomerByFaci As Boolean
    Dim blnEinvoiceE As Boolean
    
    gilDenRecCode.Enabled = True
    gidDenRecDate.Enabled = True
    txtLoveNum.Enabled = True
    '��SO041.�Ȥ�D��-�]�Ƭ[�c�y�{=1�ɡA���i�ק�[���O�覡]�B[�o������]�B[�o���γ~]�B[�o�����Y]�B[�Τ@�s��]�B[�O�_�w�}�o��]�B[�o���a�}]�C
    blnCustomerByFaci = Val(GetRsValue("Select nvl(CustomerByFaci,0) From " & GetOwner & "SO041") & "") = 1
'#5668 2010.08.13 by Corey ���դ��MAIL�A�վ㥼�ҥιq�l�o���ɵo���}�ߺ���,�ӽйq�l�p����o������O�n�ݤ��쪺.�����س��, ���ؤ���O�n�i�H�ݨ쪺�C
    '���ؤ��
'    lblDenRecDate.Visible = blnStartEinv
'    gidDenRecDate.Visible = blnStartEinv
    '���س��
'    lblDenRecCode.Visible = blnStartEinv
'    gilDenRecCode.Visible = blnStartEinv
    '�ӽйq�l�p����o�����
    lblApplyInvDate.Visible = blnStartEinv
    gdaApplyInvDate.Visible = blnStartEinv
    '�o���}�ߺ���
    lblDenRec.Visible = blnStartEinv
    cboDenRec.Visible = blnStartEinv
    If Not blnStartEinv Then
        cboDenRec.ListIndex = 0
    End If
    If Len(gdaApplyInvDate.GetValue) > 0 Then blnEinvoiceE = True
    
  '#3840 �p�G�O�G�p���Τ@�s�����i��J
    If InStr(1, cboInvoiceType.Text, "�G�p��") > 0 Then
        mskInvNo.Text = ""
        mskInvNo.Enabled = False
        cboDenRec.Enabled = True And Not blnEinvoiceE
        '#5885 2011.01.18 �Y�]�w��0: �����{�����վ�(�`�N�H�W��1�I���վ�)�A
                '�Y�]�w��1: �s�WSO002��SO138�ɡA�Y���L�νs�B�o���������G�p���o���A "�o���}�ߺ���"�h�w�]��1
        If cboDenRec.Enabled And intINVOICEKIND = 1 Then cboDenRec.ListIndex = intINVOICEKIND
        
'        If cboDenRec.Enabled Then
'            cboDenRec.ListIndex = 1
'        Else
'            cboDenRec.ListIndex = 0
'        End If
        
    ElseIf InStr(1, cboInvoiceType.Text, "�T�p��") > 0 Then
        blnNoChgApplyInvDate = True
        mskInvNo.Enabled = True And Not blnCustomerByFaci
        cboDenRec.Enabled = False
        cboDenRec.ListIndex = 0
        '#5945 ���դ�OK,�T�p�����঳���س�� By Kin 2011/04/20
        gilDenRecCode.Clear
        gidDenRecDate.SetValue ""
        gilUseInv.Clear
        blnNoChgApplyInvDate = False
        
    Else
        mskInvNo.Enabled = True And Not blnCustomerByFaci
        cboDenRec.Enabled = False
        cboDenRec.ListIndex = 0
    End If
'    Call subGil
End Sub



Private Sub chkChargeAddr_Click()
  On Error Resume Next
    cmdInvData.Visible = True
    If chkChargeAddr.Value Then
        picAddress.Visible = True
        fraData.Caption = ""
        cmdInvData.Visible = False
    Else
        fraData.Caption = "�s��"
        picAddress.Visible = False
        If OldEditMode <> giEditModeView Then cmdInvData.Visible = True
        If ObjOwnerForm.Name = "frmSO1100G" Then cmdInvData.Visible = False
    End If
    ga2Address.Clear
End Sub

Private Sub chkStop_Click()
  On Error Resume Next
    If chkStop.Value = 1 Then
        gdaStopDate.Enabled = True
        If gdaStopDate.GetValue = Empty Then
            gdaStopDate.SetValue RightDate
            gdaStopDate.SetFocus
        End If
    Else
        gdaStopDate.SetValue ""
        gdaStopDate.Enabled = False
    End If
End Sub

Private Sub cmdInvData_Click()
  On Error GoTo ChkErr
    ObjOwnerForm.uRS = rsDefTmp
    'ObjOwnerForm.uInvSeqNo = ggrData.Recordset("INVSEQNO") & ""
    blnFrmShow = False
    'ObjOwnerForm.ChangeMode lngOldEditMode
    Unload Me
    Exit Sub
ChkErr:
  ErrSub Me.Name, "cmdInvData_Click"
End Sub

Private Sub cmdLog_Click()
  On Error Resume Next
    If ggrData.Recordset.RecordCount > 0 Then
        frmSO114FB.uInvSeqNo = ggrData.Recordset("INVSEQNO")
        frmSO114FB.Show vbModal, Me
    Else
        MsgBox "�L����O��!!", vbInformation, "�T��"
    End If
End Sub

Private Sub cmdQuery_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
    cmdInvData.Enabled = False
'    Call NewRcd
'    cboInvoiceType.ListIndex = -1
    cmdInvData.Enabled = False
    If Not SchIsDataOK Then
        MsgBox "�L����M����!!�Цܤֿ�J�@�طj�M����", vbInformation, "�T��"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Call NewRcd
    cboInvoiceType.ListIndex = -1
    If Not OpenData Then Exit Sub
'    Set ggrData.Recordset = rsSO138
    Set ggrData.Recordset = rsDefTmp
    ggrData.Refresh
    If rsDefTmp.EOF Then
        MsgBox "�d�L������ �I", vbInformation, "�T��"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    cmdInvData.Enabled = iSelectCount
    Call ChangeMode(giEditModeView)
    chkChargeAddr.Value = 0
    'cmdInvData.Enabled = False
    RcdToScr
    Exit Sub
ChkErr:
    chkChargeAddr.Value = 0
    Screen.MousePointer = vbDefault
    ErrSub Me.Name, "cmdQurey_Click"
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
    Set ObjActiveForm = Me
    strActFrmName = "frmSO114FA"
    blnFrmShow = True
End Sub
Public Property Get frmShow() As Boolean
  On Error Resume Next
    frmShow = blnFrmShow
End Property
'#3929 �P�_�O�_�n�ƿ�(SO1100F�n�h��,SO01100G����ƿ�) By Kin 2008/06/02
Public Property Let uMutilChoice(ByVal vData As Boolean)
    blnMutilChoice = vData
End Property
Public Property Let uOwnerForm(ByRef vData As Form)
    Set ObjOwnerForm = vData
End Property
Public Property Get uInvSeqNo() As String
  On Error Resume Next
    uInvSeqNo = strInvSeqNo
End Property
Public Property Let uInvSeqNo(ByVal vData As String)
    On Error Resume Next
    strInvSeqNo = vData
End Property
Public Property Get OldEditMode() As giEditModeEnu
  On Error GoTo ChkErr
    OldEditMode = lngOldEditMode
    Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Get OldEditMode")
End Property
Public Property Let OldEditMode(ByVal vNewValue As giEditModeEnu)
  On Error GoTo ChkErr
    lngOldEditMode = vNewValue
    Exit Property
ChkErr:
  Call ErrSub(Me.Name, "Let OldEditMode")
End Property
Public Property Let uAccountOwner(ByVal vData As String)
  On Error GoTo ChkErr
    strAccountOwner = vData
    Exit Property
ChkErr:
  Call ErrSub(Me.Name, "Let uAccountOwner")
End Property
Public Property Let uAccountNo(ByVal vData As String)
  On Error GoTo ChkErr
    strAccountNo = vData
    Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Let uAccountNo")
End Property
Public Property Let uShowAbs(ByVal vNewValue As String)
    On Error Resume Next
    strShowAbs = vNewValue
End Property

Public Property Get EditMode() As giEditModeEnu
  On Error GoTo ChkErr
    EditMode = lngEditMode    ' ���ثe�s��Ҧ�
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Get EditMode")
End Property
'Public Property Get uRS() As Recordset
'  On Error GoTo ChkErr
'    uRS = rsDefTmp
'    Exit Property
'ChkErr:
'    Call ErrSub(Me.Name, "Get uRs")
'End Property
Public Property Let EditMode(ByVal vNewValue As giEditModeEnu)
  On Error GoTo ChkErr
    lngEditMode = vNewValue '�]�w�s��Ҧ�
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Let EditMode")
End Property
Public Property Let uCustId(ByVal vData As Long)
    On Error Resume Next
    lngCustID = vData
End Property
Public Sub ChangeMode(ByVal lngMode As giEditModeEnu)
  On Error GoTo ChkErr
    Dim blnFlag As Boolean  ' �O���O�_�b����s���Ҧ��A
    Dim blnDelFlag As Boolean

    lngEditMode = lngMode
    Select Case lngMode
           Case giEditModeInsert
                blnFlag = False
                blnDelFlag = False
                MenuEnabled , , False, True, False, True
           Case giEditModeEdit
                blnFlag = False
                blnDelFlag = False
                MenuEnabled , False, False, True, False, True
           Case giEditModeView
                blnFlag = True
                blnDelFlag = False
                MenuEnabled blnAdd, blnEdit And rsSO138.RecordCount > 0, blnDele And rsSO138.RecordCount > 0, False, False, True
           Case giEditModeDelete
                blnDelFlag = True
                MenuEnabled False, False, False, True, False, True
    End Select
    '#6793 �W�[�v�����ެO�_�i�s�� By Kin 2014/06/12
    If (lngEditMode <> giEditModeView) And (lngEditMode <> giEditModeDelete) Then
        gilCarrierTypeCode.Enabled = GetUserPriv("SO114FA", "(SO114FA5)")
        txtCarrierId1.Enabled = gilCarrierTypeCode.Enabled
        txtCarrierId2.Enabled = gilCarrierTypeCode.Enabled
    End If
    SetStatusBar lngEditMode
    fraData.Enabled = Not blnFlag
    fraDetail.Enabled = blnFlag
    gdaStopDate.Enabled = blnDelFlag
    If Not blnFlag Then txtChargeTitle.SetFocus
    frmSearch.Enabled = blnFlag
    txtChargeTitleQ.Enabled = blnFlag
    txtInvSeqNoQ1.Enabled = blnFlag
    txtInvSeqNoQ2.Enabled = blnFlag
    txtInvTitleQ.Enabled = blnFlag
    mskInvNoQ.Enabled = blnFlag
    cmdQuery.Enabled = blnFlag
    cmdLog.Enabled = blnFlag
    chkChargeAddr.Enabled = blnFlag
    '#3929 �P�_�O�_���R�����v�� By Kin 2008/06/02
    If Not blnDele Then
        chkStop.Enabled = False
        gdaStopDate.Enabled = False
    Else
        chkStop.Enabled = True
        gdaStopDate.Enabled = True
    End If
    
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ChangeMode")
End Sub
Private Sub InitializingListOcx()
  On Error GoTo ChkErr
    Dim ctlTxt As Variant
    With ga1ChargeAddress
        .SendConn gcnGi
        .SetOwner GetOwner
        .SendEntry garyGi(1) & ""
        .ShowStrtCode
        .SetCompCodeAndServiceType gCompCode & ""
    End With
    With ga1MailAddress
        .SendConn gcnGi
        .SetOwner GetOwner
        .SendEntry garyGi(1) & ""
        .ShowStrtCode
        .SetCompCodeAndServiceType gCompCode & ""
    End With
    For Each ctlTxt In Controls
        If UCase(TypeName(ctlTxt)) = "TEXTBOX" Then
            ctlTxt.Text = ""
        End If
    Next
    lblInvSeqNoValue.Caption = ""
    lblUpdTimeValue.Caption = ""
    lblUpdEnValue.Caption = ""
    cmdInvData.Enabled = False
  Exit Sub
ChkErr:
    ErrSub Me.Name, "InitializingListOcx"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift)
    If (KeyCode = vbKeyH) And (Shift = 2) Then
        If Not fraAPI.Visible Then
            fraAPI.Width = frmSearch.Width
            fraAPI.Height = frmSearch.Height
            fraAPI.Top = frmSearch.Top
            fraAPI.Left = frmSearch.Left
            fraAPI.Visible = True
        Else
            fraAPI.Visible = False
            
        End If
    End If
  Exit Sub
ChkErr:
       Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
  Dim rsSO041 As New ADODB.Recordset
    'strInvSeqNo = Empty
    InitializingListOcx
    '**************************************
    '#3785 �W�[�o���γ~ By Kin 2008/04/03
    Call SubGil
    '**************************************
    If Crm Then
        PolyFrmFunction Me
        Me.Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2)
    Else
        If UCase(ObjOwnerForm.Name) = "FRMSO1100BMDI" Then
            If Not 800600 Then Me.Move (Screen.Width - Me.Width) / 2, frmSO1100BMDI.Top + frmSO1100BMDI.tbrMenu.Height + 700
        Else
            Me.Move (Screen.Width - Me.Width) / 2, ObjOwnerForm.Top
        End If
        'If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 + 300
    End If
    '#5949 �N�U��2�q�P�_���ܳo��,�קK�X�� By Kin 2011/02/23
    'blnStartEinv = (Val(GetRsValue("Select nvl(StartEinvoice,0) From " & GetOwner & "SO041") & "") = 1)
    'intINVOICEKIND = Val(GetRsValue("Select nvl(INVOICEKIND,0) From " & GetOwner & "SO041") & "")
    '#6629 �P�_�ҥιq�l�o��XML�榡���س��ηR�߽X���A�ܤ֭n���@����즳�Ȥ~�i�s�� By Kin 2013/12/05
    'blnStartEinvoiceXML = (Val(GetRsValue("Select Nvl(StartEinvoiceXML,0) From " & GetOwner & "SO041") & "") = 1)
    
     Dim aSQL As String
    '#6793 �W�[�P�_�O�_�n���һP���Һ��} By Kin 2014/06/12
    aSQL = "SELECT StartMobileAPI,StartLovenumAPI,MobileAPIURL,LovenumAPIURL,APIappID, " & _
                    "nvl(StartEinvoice,0) StartEinvoice, nvl(INVOICEKIND,0) INVOICEKIND, " & _
                    "Nvl(StartEinvoiceXML,0) StartEinvoiceXML " & _
                    " FROM " & GetOwner & "SO041"
    If Not GetRS(rsSO041, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    blnStartEinv = Val(rsSO041("StartEinvoice")) = 1
    intINVOICEKIND = Val(rsSO041("INVOICEKIND"))
    blnStartEinvoiceXML = (Val(rsSO041("StartEinvoiceXML") & "") = 1)
    blnStartMobileAPI = Val(rsSO041("StartMobileAPI") & "") = 1
    blnStartLovenumAPI = Val(rsSO041("StartLovenumAPI") & "") = 1
    strMobileAPIURL = rsSO041("MobileAPIURL") & ""
    strLovenumAPIURL = rsSO041("LovenumAPIURL") & ""
    strAPIappID = rsSO041("APIappID") & ""
    
    
    
    
    'Val (rsSO041("INVOICEKIND") & "")
    'If OldEditMode = giEditModeView Then cmdInvData.Visible = True
    '#3929 �p�G���O�ƿ�Ҧ�,��Show��w By Kin 2008/06/02
    If Not blnMutilChoice Then cmdInvData.Visible = False
    '#3929 ����H�n�w�]�a�J By Kin 2008/05/26
    cboChargeTitleQ.Clear
    ShowProposer
    ga2Address.SetCompcode gCompCode
    With ga2Address
        .SetParentForm Me
        .SendConn gcnGi
        .TableName = GetOwner & "SO014"
    End With
    picAddress.Move fraData.Left, fraData.Top + 100
    picAddress.Visible = False
    Set rsSO138 = New ADODB.Recordset
    Set rsDefTmp = New ADODB.Recordset
    Call subGrd
    If Not OpenData(True) Then Exit Sub
    'Set ggrData.Recordset = rsSO138
    Set ggrData.Recordset = rsDefTmp
    ggrData.Refresh
    If Not rsDefTmp.EOF Then
        rsDefTmp.MoveFirst
        If strShowAbs <> Empty Then
            rsDefTmp.Find "InvSeqNo=" & strShowAbs
        End If
    End If
    
'
'    '******************************************************
'    '#3789 ���դ�OK,�n�w�]���b�� By Kin 2008/03/03
'    If Not rsSO138.EOF Then
'        If strShowAbs <> Empty Then
'            rsSO138.Find "InvSeqNo=" & strShowAbs
'        End If
'    End If
'    '******************************************************
    If Not OpenSO003 Then Exit Sub
    Set ggrData2.Recordset = rsSO003
    ggrData2.Refresh
    '***********************************
    '#3929 �v������ By Kin 2008/06/02
    Call UserPermissionGo(True)
    '***********************************
    sstDetail.Tab = 0
    Call ChangeMode(giEditModeView)
    iSelectCount = 0
    If blnMutilChoice Then
        If rsDefTmp.RecordCount = 1 Then
            cmdInvData.Enabled = True
        Else
            cmdInvData.Enabled = False
        End If
    End If
    
   
    'cboChargeTitleQ.SetFocus
    '#5668 2010.07.09 by Corey ���Ӹ�ƥ\��վ�
    'blnStartEinv = (Val(GetRsValue("Select nvl(StartEinvoice,0) From " & GetOwner & "SO041") & "") = 1)
    'If Not GetRS(rsSO041, "") Then GoTo ChkErr
    'blnStartEinv = Val(rsSO041("StartEinvoice") & "") = 1 '#5668 2010.07.09 by Corey ���Ӹ�ƥ\��վ�
    'intINVOICEKIND = Val(rsSO041("INVOICEKIND") & "") '#5885 2011.01.18 by Corey  �s�W"�o���}�ߺ����w�]��"(SO041. INVOICEKIND)�i���0.�q�l�p����o����1.�q�l�o��, �w�]��0
    On Error Resume Next
    Close3Recordset rsSO041
    'Call cboInvoiceType_Click
   
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Load"
End Sub
'#6793 �إ�URL���Ѽ� 0=������� 1=�R�߽X By Kin 2014/06/13
Private Function CreateUrlPara(ByVal APIKind As Integer) As String
  On Error GoTo ChkErr
    Dim Result As String
    Result = "version=1.0&action="
    Select Case APIKind
        Case 0
            Result = Result & "bcv" & "&barCode=" & txtCarrierId1.Text & "&TxID=000&appId=" & strAPIappID
        Case 1
            Result = Result & "preserveCodeCheck" & "&pCode=" & txtLoveNum.Text & "&TxID=000&appId=" & strAPIappID
    End Select
    CreateUrlPara = Result
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "CreateUrlPara")
End Function
Private Function OpenData(Optional ByVal blnLoad As Boolean = False) As Boolean
  On Error GoTo ChkErr
    Dim strSO138 As String
    Dim strAccOwner As String
    Dim i As Long
    '#3929 �n�令��ƿ� By Kin 2008/05/26
     
    Call DefineRs
    If strAccountOwner = Empty Then
        strAccOwner = Empty
    Else
        strAccOwner = " Union All Select * From " & GetOwner & "SO138 C" & _
                    " Where C.ChargeTitle='" & strAccountOwner & "'"
        
    End If
    If blnLoad Then
        '************************************************************************************************************
        '#3789 �w�]�a�J�ӥΤ᪺�o����� By Kin 2008/02/29
        'strSO138 = "Select * From " & GetOwner & "SO138 Where 1=0"
        strSO138 = "Select Distinct * From (Select * From " & GetOwner & "SO138 A " & _
                    "Where A.invSeqNo in (Select InvSeqNo From " & GetOwner & "SO002AD B Where B.Custid=" & lngCustID & _
                    " And B.AccountNo='" & IIf(strAccountNo = Empty, "0", strAccountNo) & "')" & _
                    " Union All " & _
                    "Select * From " & GetOwner & "SO138 B " & _
                    " Where B.invSeqNo in (" & IIf(strInvSeqNo = Empty, "0", strInvSeqNo) & ")" & strAccOwner & ") Order by InvseqNo"

                    

        '************************************************************************************************************
    Else
        strChoose = subChoose
        strSO138 = "Select Distinct * From (Select * From " & GetOwner & "SO138 A " & IIf(strChoose <> "", " Where " & strChoose, "") & ") Order By InvSeqNo"
    End If
    If Not GetRS(rsSO138, strSO138, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
    
    '*****************************************************************************************************
    '#3929 �n�令��ƿ� By Kin 2008/05/26
    If rsSO138.RecordCount > 0 Then
        Do While Not rsSO138.EOF
            rsDefTmp.AddNew
            For i = 0 To rsSO138.Fields.Count - 1
                If Not IsNull(rsSO138.Fields(i).Value) Then
                    rsDefTmp(rsSO138.Fields(i).Name).Value = rsSO138.Fields(i).Value
                End If
            Next i
            rsDefTmp.Update
            rsSO138.MoveNext
        Loop
        '************************************************************
        '#3982 �p�G�u���@����ƭn�۰ʿ�� By Kin 2008/07/16
        If rsSO138.RecordCount = 1 Then
            If blnMutilChoice Then
                rsDefTmp("Choice").Value = "1"
                rsDefTmp.Update
                cmdInvData.Enabled = True
                iSelectCount = 1
            End If
        End If
        '************************************************************
    End If
    '****************************************************************************************************
    
    OpenData = True
  Exit Function
ChkErr:
    ErrSub Me.Name, "OpenData"
End Function
Private Function subChoose() As String
    On Error GoTo ChkErr
        Dim lngADDRNO As Long
        Dim rsSO014 As New ADODB.Recordset
        Dim strSO014 As String
        lngADDRNO = -1
        strChoose = ""
        '#3929 ����H���Combox���� By Kin 2008/05/26
        If cboChargeTitleQ.Text <> "" Then subAnd "UPPER(A.ChargeTitle) Like '" & UCase(cboChargeTitleQ.Text) & "%'"
        'If txtChargeTitleQ.Text <> "" Then subAnd "UPPER(A.ChargeTitle) Like '" & UCase(txtChargeTitleQ.Text) & "%'"
        If mskInvNoQ.Text <> "" Then subAnd "A.InvNo= '" & mskInvNoQ.Text & "'"
        If txtInvTitleQ.Text <> "" Then subAnd "UPPER(A.InvTitle) Like '" & UCase(txtInvTitleQ.Text) & "%'"
        If txtInvSeqNoQ1.Text <> "" Then subAnd "A.INVSEQNO>=" & Trim(txtInvSeqNoQ1.Text)
        If txtInvSeqNoQ2.Text <> "" Then subAnd "A.INVSEQNO<=" & Trim(txtInvSeqNoQ2.Text)
        If chkChargeAddr.Value Then
'            Set rsSO014 = gcnGi.Execute("Select AddrNo From " & GetOwner & "SO014 " & _
'                                    " Where ADDRESS='" & ga2Address.GetAddrString & "' And CompCode=" & gCompCode)
'            If Not rsSO014.EOF Or Not rsSO014.BOF Then
'                lngADDRNO = rsSO014("AddrNo")
'
'            End If
            strSO014 = "Select AddrNo From " & GetOwner & "SO014 Where " & ga2Address.GetQryString
            'subAnd "A.ChargeAddrNo= " & lngADDRNO
            subAnd "A.ChargeAddrNo IN(" & strSO014 & ")"
        End If
        CloseRecordset rsSO014
        subChoose = strChoose
    Exit Function
ChkErr:
    CloseRecordset rsSO014
    ErrSub Me.Name, "subChoose"
End Function

Private Sub subGrd()
    On Error Resume Next
    Dim mFlds As New GiGridFlds
    Dim mFlds2 As New GiGridFlds
    With mFlds
        '#3929 �令�n��ƿ� By Kin 2008/05/26
        If blnMutilChoice Then
            .Add "Choice", , , , , "���"
        End If
        .Add "InvSeqNo", , , , , "�y���s��     "
        .Add "ChargeTitle", , , , , "����H" & Space(40)
        .Add "InvoiceType", , , , , "�o������"
        .Add "InvNo", , , , , "�o���Τ@�s��"
        .Add "InvTitle", , , , , "�o�����Y" & Space(50)
        .Add "InvAddress", , , , , "�o���a�}" & Space(50)
        .Add "ChargeAddress", , , , , "���O�a�}" & Space(50)
        .Add "MailAddress", , , , , "�l�H�a�}" & Space(50)
        .Add "StopFlag", , , , , "����"
        .Add "StopDate", giControlTypeDate, , , , "���Τ��" & Space(5)
    End With
    ggrData.AllFields = mFlds
    ggrData.SetHead
    
    With mFlds2
        .Add "CUSTID", , , , , "�Ȥ�s��    "
        .Add "CUSTNAME", , , , , "�Ȥ�m�W         "
        .Add "CITEMCODE", , , , , "���O���إN�X    "
        .Add "CITEMNAME", , , , , "���O���ئW��" & Space(10)
        .Add "SERVICETYPE", , , , , "�A�����O"
        .Add "ACCOUNTNO", , , , , "���ڱb��" & Space(10)
        .Add "CMCODE", , , , , "���O�覡�N�X"
        .Add "CMNAME", , , , , "���O�覡�W��" & Space(10)
        .Add "BANKCODE", , , , , "�Ȧ�N�X"
        .Add "BANKNAME", , , , , "�Ȧ�W��" & Space(15)
        .Add "StopFlag", , , , , "����"
    End With
    ggrData2.AllFields = mFlds2
    ggrData2.SetHead
End Sub

Public Sub AddNewGo()
  On Error GoTo ChkErr
    ChangeMode (giEditModeInsert)
    NewRcd
    chkChargeAddr.Value = 0
    Call DefaultValue
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "cmdAdd_Click")
End Sub
Private Sub NewRcd()
  On Error Resume Next
    txtChargeTitle.Text = ""
    '#3773 �o�������w�]�ȧאּ2�p�� By Kin 2008/02/18
    cboInvoiceType.ListIndex = 1
    txtInvTitle.Text = ""
    mskInvNo.Text = ""
    txtInvAddress.Text = ""
    gdaStopDate.SetValue ""
    ga1ChargeAddress.AddrNo = 0
    ga1ChargeAddress.AddrString = ""
    ga1MailAddress.AddrNo = 0
    ga1MailAddress.AddrString = ""
    lblInvSeqNoValue.Caption = ""
    lblUpdEnValue.Caption = ""
    lblUpdTimeValue.Caption = ""
    gidDenRecDate.SetValue ""
    gdaApplyInvDate.SetValue ""
    gilDenRecCode.SetCodeNo ""
    '#6705 �s�W�ɧ�����ƲM�� By Kin 2014/02/10
    gilCarrierTypeCode.Clear
    txtCarrierId1.Text = ""
    txtCarrierId2.Text = ""
    txtA_CarrierId1.Text = ""
    txtA_CarrierId2.Text = ""
    cboDenRec.ListIndex = 0
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "NewRcd")
End Sub
Public Sub EditGo()
  On Error GoTo ChkErr
  Call ChangeMode(giEditModeEdit)
  chkChargeAddr.Value = 0
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "cmdEdit_Click")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ChkErr
        If EditMode <> giEditModeView Then
            If giMsgNotSave = vbNo Then
                Cancel = True: Exit Sub
            Else
                CancelGo
            End If
        End If
        CloseRecordset rsSO003
        CloseRecordset rsSO138
        CloseRecordset rsDefTmp
        blnFrmShow = False
        
    Exit Sub
ChkErr:
    ErrSub Me.Name, "QueryUnload"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ReleaseCOM Me
    If ObjOwnerForm.Name = "frmSO1100F" Then
        ObjOwnerForm.OpenSO138
    End If
    strAccountNo = Empty
    lngCustID = 0
    strInvSeqNo = Empty
    ObjOwnerForm.ChangeMode lngOldEditMode
    ObjOwnerForm.SetFocus
    'Call FormQueryUnload
End Sub


Private Sub gdaApplyInvDate_Change()
    On Error Resume Next
    If EditMode = giEditModeView Then
        Exit Sub
    End If
    Call cboInvoiceType_Click

End Sub

Private Sub ggrData_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    '#3929 �p�G���O�ƿ�h���} By Kin 2008/06/02
    If Not blnMutilChoice Then Exit Sub
    If EditMode = giEditModeView Then
        If UCase(giFld.FieldName) = "CHOICE" Then
            If Value = 1 Then Value = vbRed
        End If
    End If
End Sub

Private Sub ggrData_DblClick()
  On Error Resume Next
    '#3929 �p�G�O�ƿ�~�i��Update By Kin 2008/06/02
    If rsDefTmp.RecordCount = 0 Then
        strInvSeqNo = Empty
        Exit Sub
    End If
    '#3929 �n�P�_�O�_���ƿ�γ��(SO1100F�PSO1100G) By Kin 2008/06/02
    If blnMutilChoice Then
        If rsDefTmp("Choice") = "1" Then
            rsDefTmp("Choice") = "0"
            If iSelectCount > 0 Then iSelectCount = iSelectCount - 1
        Else
            rsDefTmp("Choice") = "1"
            iSelectCount = iSelectCount + 1
        End If
        rsDefTmp.Update
        If iSelectCount > 0 Then
            cmdInvData.Enabled = True
        Else
            cmdInvData.Enabled = False
        End If
        
    Else
        strInvSeqNo = rsDefTmp("InvSeqNo")
        If UCase(ObjOwnerForm.Name) = "FRMSO1100G" Then
            If strInvSeqNo <> Empty Then
                frmSO1100G.Set106Inv = strInvSeqNo
            End If
        End If
        Unload Me
    End If
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error Resume Next
    '#3929 �ƿ�~���P�_ By Kin 2008/06/02
    If blnMutilChoice Then
        If giFld.FieldName = "Choice" Then
            If Value = "1" Then
                Value = "V"
            Else
                Value = ""
            End If
        End If
    End If
    If giFld.FieldName = "InvoiceType" Then
        Select Case Value
            Case 0
                Value = "�K�}"
            Case 2
                Value = "�G�p��"
            Case 3
                Value = "�T�p��"
        End Select
    End If
    If giFld.FieldName = "StopFlag" Then
        Select Case Value
            Case 0
                Value = "�_"
            Case 1
                Value = "�O"
            Case Else
                Value = "�_"
        End Select
    End If

End Sub

Private Sub ggrData2_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    If giFld.FieldName = "StopFlag" Then
        Select Case Value
            Case 0
                Value = "�_"
            Case 1
                Value = "�O"
            Case Else
                Value = "�_"
        End Select
    End If
  
End Sub

Private Sub gilUseInv_Change()
'#5668 2010.07.28 by Corey ���ի�ݭn�W�[�\��A�o���γ~���ѦҸ�=1 �~�ݭn��J���س��Ψ��ؤ��
    On Error Resume Next
    blnShowDenRec = False
    If Len(gilUseInv.GetCodeNo & "") <> 0 Then
'        gilDenRecCode.Clear
'        gidDenRecDate.SetValue ""
'        txtLoveNum.Text = ""
        blnShowDenRec = (Val(GetRsValue("Select RefNo from " & GetOwner & "CD095 Where Codeno=" & gilUseInv.GetCodeNo) & "") = 1)
    End If
    '#6600 �W�[�R�߽X By Kin 2013/10/21
    gilDenRecCode.Enabled = blnShowDenRec
    gidDenRecDate.Enabled = blnShowDenRec
    txtLoveNum.Enabled = blnShowDenRec
    If Not blnShowDenRec Then
        gilDenRecCode.Clear
        gidDenRecDate.SetValue ""
        txtLoveNum.Text = ""
    End If
End Sub

Private Sub gilUseInv_Validate(Cancel As Boolean)
'#3920 �T�p���o���ɻ�ĵ�i By Kin 2008/07/31
'  On Error GoTo ChkErr
'    If InStr(1, cboInvoiceType.Text, "�T�p��") > 0 Then
'        If gilUseInv.GetCodeNo & "" <> "" Then
'            MsgBox "�T�p��(��~�H)�o�����i����!", vbInformation, "�T��"
'        End If
'    End If
'    Exit Sub
'ChkErr:
'    ErrSub Me.Name, "gilUseInv_Validate"
End Sub

Private Sub mskInvNo_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If cboInvoiceType.Enabled And mskInvNo.Text <> Empty Then
        If Not InvNoIsOk(mskInvNo.Text, False) Then
            MsgBox "�Τ@�s����J���~�I", vbCritical, "�T��"
            Cancel = True
        End If
    Else
        Cancel = False
    End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "mskInvNo_Validate"
End Sub

Private Sub rsDefTmp_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
    If pRecordset.EOF Then Exit Sub
    If pRecordset.EditMode = adEditAdd Then Exit Sub
    RcdToScr
    Set ggrData2.Recordset = rsSO003
    ggrData2.Refresh
End Sub



Public Sub UpdateGo()
  On Error GoTo ChkErr
    If Not IsDataOk Then Exit Sub
    gcnGi.BeginTrans
    Dim varBookMark As Variant
    blnCanModifyCarrierId = False
'    If EditMode <> giEditModeInsert Then varBookMark = rsSO138.Bookmark
    '#3929 �令�ƿ�,�ҥH�nŪ��rsDefTmp By Kin 2008/05/26
    
    If EditMode <> giEditModeInsert Then varBookMark = rsDefTmp.Bookmark
     '#6705 �o�����Y�������(�t�G����H�B�o�����Y�B�Τ@�s��)�A��ק�o�����ɡA�N���o�����Y(SO138. A_CarrierId1)���|�����㭫���s�� By Kin 2014/02/11
    If EditMode <> giEditModeInsert Then
        If (rsDefTmp("ChargeTitle") & "" <> txtChargeTitle.Text) Or _
           (rsDefTmp("InvTitle") & "" <> txtInvTitle.Text) Or _
           (rsDefTmp("InvNo") & "" <> mskInvNo.Text) Then
            If MsgBox("�ק鶴�����(����H�B�o�����Y�B�Τ@�s��)�|�����|������s��!", vbQuestion + vbOKCancel, "�T��") = vbOK Then
                blnCanModifyCarrierId = True
            Else
                gcnGi.RollbackTrans

'                Set ggrData.Recordset = rsDefTmp
'                ggrData.Refresh
'                If EditMode = giEditModeEdit Then
'                    rsDefTmp.Bookmark = varBookMark
'                End If
'                RcdToScr
'                ChangeMode (giEditModeEdit)
                Exit Sub
            End If
        End If
    End If
   
    If Not ScrToRcd Then Exit Sub
    Set ggrData.Recordset = rsDefTmp
    ggrData.Refresh
    
    gcnGi.CommitTrans
    blnBeginTrans = False
    '#3982 �s�W��Ʈɭn���d�b�ӵ� By Kin 2008/07/16
    If EditMode = giEditModeEdit Then
        'rsSO138.Requery
        rsDefTmp.AbsolutePosition = varBookMark
    Else
        rsDefTmp.MoveLast
    End If
    Call ChangeMode(giEditModeView)
    
    If Not RcdToScr Then Exit Sub
   Exit Sub
ChkErr:
    blnBeginTrans = False
    gcnGi.RollbackTrans
    Call ErrSub(Me.Name, "UpdateGo")
End Sub

Private Function AlertRsDefTmp() As Boolean
  On Error GoTo ChkErr
    Dim i As Long
    If EditMode = giEditModeInsert Then
        rsDefTmp.AddNew
    End If
    For i = 0 To rsSO138.Fields.Count - 1
        If Not IsNull(rsSO138.Fields(i)) Then
            If UCase(rsSO138.Fields(i).Name = "STOPDATE") Then
                If IsNull(rsSO138.Fields(i)) Then
                    rsDefTmp.Fields("StopDate") = ""
                Else
                    rsDefTmp.Fields("StopDate").Value = Format(rsSO138.Fields(i).Value, "yyyy/mm/dd")
                End If
            Else
                rsDefTmp.Fields(rsSO138.Fields(i).Name).Value = rsSO138.Fields(i).Value
            End If
        Else
            rsDefTmp.Fields(rsSO138.Fields(i).Name).Value = ""
        End If
    Next
    rsDefTmp.Update
    AlertRsDefTmp = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "AlertRsDefTmp"
End Function
Private Function ScrToRcd() As Boolean
  On Error GoTo ChkErr
    
    If EditMode = giEditModeInsert Then
        If Not GetRS(rsSO138, "Select * From " & GetOwner & "SO138 Where 1=0", gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
        rsSO138.AddNew
        rsSO138("InvSeqNo") = GetSeqNo
         '#6705 �o�����Y�������(�t�G����H�B�o�����Y�B�Τ@�s��)�A��ק�o�����ɡA�N���o�����Y(SO138. A_CarrierId1)���|�����㭫���s�� By Kin 2014/02/11
        rsSO138("A_CarrierId1") = GetA_CarrierId1
    Else
        If Not GetRS(rsSO138, "Select * From " & GetOwner & "SO138 Where InvSeqNo=" & rsDefTmp("InvSeqNo"), gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
    End If
    '#6705 �o�����Y�������(�t�G����H�B�o�����Y�B�Τ@�s��)�A��ק�o�����ɡA�N���o�����Y(SO138. A_CarrierId1)���|�����㭫���s�� By Kin 2014/02/11
    
    If (rsSO138("ChargeTitle") & "" <> txtChargeTitle.Text) Or _
        (rsSO138("InvTitle") & "" <> txtInvTitle.Text) Or _
        (rsSO138("InvNo") & "" <> mskInvNo.Text) Then
        If blnCanModifyCarrierId Then
            rsSO138("A_CarrierId1") = GetA_CarrierId1
        End If
    End If
    
    rsSO138("ChargeTitle") = IIf(txtChargeTitle.Text <> "", txtChargeTitle.Text, Null)
    rsSO138("InvNo") = IIf(mskInvNo.Text <> "", mskInvNo.Text, Null)
    rsSO138("ChargeAddrNo") = IIf(ga1ChargeAddress.AddrNo <> 0, ga1ChargeAddress.AddrNo, Null)
    rsSO138("ChargeAddress") = IIf(ga1ChargeAddress.AddrString <> "", ga1ChargeAddress.AddrString, Null)
    rsSO138("MailAddrNo") = IIf(ga1MailAddress.AddrNo <> 0, ga1MailAddress.AddrNo, Null)
    rsSO138("MailAddress") = IIf(ga1MailAddress.AddrString <> "", ga1MailAddress.AddrString, Null)
    rsSO138("InvTitle") = IIf(txtInvTitle.Text <> "", txtInvTitle.Text, Null)
    rsSO138("InvAddress") = IIf(txtInvAddress.Text <> "", txtInvAddress.Text, Null)
    If gdaStopDate.GetValue <> Empty Then
        rsSO138("StopDate") = gdaStopDate.GetValue(True)
        rsSO138("StopFlag") = 1
    Else
        rsSO138("StopFlag") = 0
        rsSO138("StopDate") = Null
    End If
    Select Case cboInvoiceType.ListIndex
        Case 0
            rsSO138("InvoiceType") = 0
        Case 1
            rsSO138("InvoiceType") = 2
        Case 2
            rsSO138("InvoiceType") = 3
        Case Else
            rsSO138("InvoiceType") = 0
    End Select
    '*****************************************************************
    '#3785 �W�[�o���γ~ By Kin 2008/04/03
    If gilUseInv.GetCodeNo <> Empty Then
        rsSO138("InvPurposeCode") = gilUseInv.GetCodeNo & ""
        rsSO138("InvPurposeName") = gilUseInv.GetDescription & ""
    Else
        rsSO138("InvPurposeCode") = Null
        rsSO138("InvPurposeName") = Null
    End If
    '*****************************************************************
    '#6371 �Ȥ�������O���X�B�Ȥ������X1�B�Ȥ�������X2 By Kin 2013/02/21
    If gilCarrierTypeCode.GetCodeNo <> Empty Then
        rsSO138("CarrierTypeCode") = gilCarrierTypeCode.GetCodeNo & ""
    Else
        rsSO138("CarrierTypeCode") = Null
    End If
    rsSO138("CarrierId1") = Null
    rsSO138("CarrierId2") = Null
    If txtCarrierId1.Text <> "" Then
        rsSO138("CarrierId1") = txtCarrierId1.Text
    End If
    If txtCarrierId2.Text <> "" Then
        rsSO138("CarrierId2") = txtCarrierId2.Text
    End If
    rsSO138("Updtime") = GetDTString(RightNow)
    rsSO138("UpdEn") = garyGi(1)
    rsSO138("PreInvoice") = chkPreInv.Value
    '#5668 2010.07.19 by Corey
    rsSO138("InvoiceKind") = cboDenRec.ListIndex
    rsSO138("ApplyInvDate") = NoZero(gdaApplyInvDate.GetValue(True))
    rsSO138("DenRecDate") = NoZero(gidDenRecDate.GetValue(True))
    rsSO138("DenRecCode") = NoZero(gilDenRecCode.GetCodeNo)
    rsSO138("DenRecName") = NoZero(gilDenRecCode.GetDescription)
    '#6600 �W�[�R�߽X By Kin 2013/10/21
    rsSO138("BillMailKind") = chkBillMailKind.Value
    If Len(txtLoveNum.Text & "") > 0 Then
        rsSO138("LoveNum") = txtLoveNum.Text
    Else
        rsSO138("LoveNum") = Null
    End If
   
    
    
    rsSO138.Update
    
    '#3929 �P�BUpdate�@����ƦܼȦs�� By Kin 2008/05/26
    If Not AlertRsDefTmp() Then Exit Function
    ScrToRcd = True
  Exit Function
ChkErr:
  ErrSub Me.Name, "ScrToRcd"
End Function
'#3929 �n��ƿ�,�ҥH�n��ŪrsDefTmp By Kin 2008/05/26
Private Function RcdToScr() As Boolean
  On Error GoTo ChkErr
    'cmdInvData.Enabled = False
    If EditMode = giEditModeInsert Then Exit Function
    If rsDefTmp.RecordCount > 0 Then
        lblInvSeqNoValue.Caption = rsDefTmp("INVSEQNO") & ""
        txtChargeTitle.Text = rsDefTmp("ChargeTitle") & ""
        Select Case rsDefTmp("InvoiceType") & ""
            Case "0"
                cboInvoiceType.ListIndex = 0
            Case "2"
                cboInvoiceType.ListIndex = 1
            Case "3"
                cboInvoiceType.ListIndex = 2
            Case Else
                cboInvoiceType.ListIndex = 0
        End Select
        mskInvNo.Text = rsDefTmp("InvNo") & ""
        txtInvTitle.Text = rsDefTmp("InvTitle") & ""
        txtInvAddress.Text = rsDefTmp("InvAddress") & ""
        If rsDefTmp("ChargeAddrNo") & "" <> "" Then
            ga1ChargeAddress.AddrNo = rsDefTmp("ChargeAddrNo")
        Else
            ga1ChargeAddress.AddrNo = 0
        End If
        If rsDefTmp("ChargeAddress") & "" <> "" Then
            ga1ChargeAddress.AddrString = rsDefTmp("ChargeAddress") & ""
        Else
            ga1ChargeAddress.AddrString = ""
        End If
        If rsDefTmp("MailAddrNo") & "" <> "" Then
            ga1MailAddress.AddrNo = rsDefTmp("MailAddrNo")
        Else
            ga1MailAddress.AddrNo = 0
        End If
        If rsDefTmp("MailAddrNo") & "" <> "" Then
            ga1MailAddress.AddrString = rsDefTmp("MailAddress") & ""
        Else
            ga1MailAddress.AddrString = ""
        End If

        If Not IsNull(rsDefTmp("StopDate")) Then
            gdaStopDate.SetValue Format(rsDefTmp("StopDate") & "", "yyyy/mm/dd")
        End If
        '*************************************************************
        '#3785 �W�[�o���γ~ By Kin 2008/04/03
        If Not IsNull(rsDefTmp("InvPurposeCode")) Then
            gilUseInv.SetCodeNo rsDefTmp("InvPurposeCode") & ""
            gilUseInv.Query_Description
        Else
            gilUseInv.Clear
            gilUseInv.SetCodeNo ""
        End If
        '************************************************************
        lblUpdTimeValue.Caption = rsDefTmp("UpdTime") & ""
        lblUpdEnValue.Caption = rsDefTmp("UpdEn") & ""
        If rsDefTmp("StopFlag") = 1 Then
            chkStop.Value = 1
            'cmdInvData.Enabled = False
            'MenuEnabled True, False, False, False, False, True
        Else
            chkStop.Value = 0
            'cmdInvData.Enabled = True
            'MenuEnabled True, True, True, False, False, True
        End If
        '#5534 2010.03.22 by Corey
        chkPreInv.Value = NoZero(rsDefTmp("PreInvoice") & "", True)
        '#5668 2010.07.19 by Corey
        blnNoChgInvoiceKind = True  '����@�w�n�A���M�ƥ�|�@�����yĲ�o By Kin 2011/03/24
        cboDenRec.ListIndex = NoZero(rsDefTmp("InvoiceKind") & "", True)
        If cboDenRec.ListIndex = 0 Then
            cboDenRec.Enabled = False
            gdaApplyInvDate.Enabled = True
            '#6130 �ӽйq�l�p����o����� ��Null�� cboDenRec�n��ק� By Kin 2011/09/21
            If gdaApplyInvDate.GetValue = "" Then
                cboDenRec.Enabled = True
            End If
        Else
            cboDenRec.Enabled = True
            gdaApplyInvDate.Enabled = False
        End If
        blnNoChgInvoiceKind = False '#5945 ���դ�OK,�o��@�w�n�[,���M�n�ĤG������~�|Ĳ�o By Kin 2011/04/20
        If Len(rsDefTmp("ApplyInvDate") & "") > 0 Then
            gdaApplyInvDate.SetValue Format(rsDefTmp("ApplyInvDate") & "", "yyyy/mm/dd")
            cboDenRec.Enabled = False
        Else
            'blnNoChgApplyInvDate = True '����@�w�n�A���M�ƥ�|�@�����yĲ�o By Kin 2011/03/24
            gdaApplyInvDate.SetValue ""
            'gdaApplyInvDate.Enabled = True
        End If
        '2011/09/28 Jacky 6131 Kobe ���T�����ި��ؤ���O�_����,���س�즳�ȴN�n��ܡC
        'If Len(rsDefTmp("DenRecDate") & "") > 0 Then
            gidDenRecDate.SetValue Format(rsDefTmp("DenRecDate") & "", "yyyy/mm/dd")
            gilDenRecCode.SetCodeNo rsDefTmp("DenRecCode") & ""
            gilDenRecCode.Query_Description
        'End If
        chkBillMailKind.Value = NoZero(rsDefTmp("BillMailKind") & "", True)
        If Len(gdaApplyInvDate.GetValue & "") > 0 Then gdaApplyInvDate.Enabled = True
        '#6371 �Ȥ�������O���X�B�Ȥ������X1�B�Ȥ�������X2 By Kin 2013/02/21
        gilCarrierTypeCode.Clear
        txtCarrierId1.Text = Empty
        txtCarrierId2.Text = Empty
        If Len(rsDefTmp("CarrierTypeCode") & "") > 0 Then
            gilCarrierTypeCode.SetCodeNo rsDefTmp("CarrierTypeCode") & ""
            gilCarrierTypeCode.Query_Description
        End If
         If Len(rsDefTmp("CarrierId1") & "") > 0 Then
            txtCarrierId1.Text = rsDefTmp("CarrierId1") & ""
        End If
        If Len(rsDefTmp("CarrierId2") & "") > 0 Then
            txtCarrierId2.Text = rsDefTmp("CarrierId2") & ""
        End If
        '#6600 �W�[�R�߽X By Kin 2013/10/21
        If Len(rsDefTmp("LoveNum") & "") > 0 Then
            txtLoveNum.Text = rsDefTmp("LoveNum") & ""
        End If
        '#6629 �W�[�|��������X�P���X By Kin 2013/11/06
        txtA_CarrierId1.Text = rsDefTmp("A_CarrierId1") & ""
        txtA_CarrierId2.Text = rsDefTmp("A_CarrierId2") & ""
        '#6629 �W�[�o���γ~�O���ؤ~�i�ק�R�߽X�P���س�� By Kin 2013/12/03
        If Len(rsDefTmp("InvPurposeCode") & "") = 0 Then
            gilDenRecCode.Enabled = False
            gidDenRecDate.Enabled = False
            txtLoveNum.Enabled = False
        End If
    End If
    If Not OpenSO003 Then Exit Function
    RcdToScr = True
    Exit Function
ChkErr:
  ErrSub Me.Name, "RcdtoScr"
End Function
Private Function GetA_CarrierId1() As String
  On Error GoTo ChkErr
    GetA_CarrierId1 = GetRsValue("select SF_GetA_CarrierId1(" & gCompCode & " ) from dual") & ""
    Exit Function
ChkErr:
  ErrSub Me.Name, "GetA_CarrierId1"
End Function
Private Function GetSeqNo() As String
  On Error GoTo ChkErr
    GetSeqNo = RPxx(gcnGi.Execute("SELECT " & GetOwner & "S_SO138_InvSeqNo.NEXTVAL FROM DUAL").GetString & "")
  Exit Function
ChkErr:
    ErrSub Me.Name, "GetSeqNo"
End Function
Public Sub CancelGo()
   On Error GoTo ChkErr
    '#3929 �令�ƿ�,�ҥH�nŪ��rsDefTmp By Kin 2008/05/26
    If blnBeginTrans Then
        rsDefTmp.CancelUpdate
    End If
    blnBeginTrans = False
    If EditMode = giEditModeView Then
         Unload Me
         Exit Sub
    Else
        If EditMode = giEditModeEdit Then
            Call ChangeMode(giEditModeView)
            RcdToScr
            Exit Sub
        Else
            If Not rsDefTmp.EOF Then
                rsDefTmp.MoveFirst
                Call ChangeMode(giEditModeView)
                RcdToScr
            Else
                Call NewRcd
                cboInvoiceType.ListIndex = -1
                Call ChangeMode(giEditModeView)
                RcdToScr
            End If
        End If
    End If
 Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "CancelGo")
End Sub

Public Sub DeleteGo()
  On Error GoTo ChkErr
    'Dim varBook As Variant
    '#3929 �令�ƿ�,�ҥH�nŪ��rsDefTmp By Kin 2008/05/26
    If rsDefTmp.EOF Then
        MsgBox "�L������ !", vbInformation, "�T��"
        Exit Sub
    End If
    blnBeginTrans = True
    chkStop.Value = 1
    'varBook = rsSO138.Bookmark
    'rsSO138("StopFlag") = 1
    'rsSO138.AbsolutePosition = varBook
    ChangeMode giEditModeDelete
    gdaStopDate.SetValue RightDate
    gdaStopDate.SetFocus
  Exit Sub
ChkErr:
    ErrSub Me.Name, "DeleteGo"
End Sub
'#3929 �令��ƿ�,�ҥH�n��ŪrsDefTmp By Kin 2008/05/26
Private Function OpenSO003() As Boolean
  On Error GoTo ChkErr
    Dim strSO003 As String
    If rsDefTmp.EOF Then
        strSO003 = "Select A.*,B.CustName From " & GetOwner & "SO003 A," & GetOwner & "SO001 B" & _
                 " Where A.CUSTID=B.CUSTID AND 1=0"
    Else
        strSO003 = "Select A.*,B.CustName From " & GetOwner & "SO003 A," & GetOwner & "SO001 B" & _
                    " Where A.CompCode=B.CompCode And A.CUSTID=B.CUSTID(+)" & _
                    " AND A.INVSEQNO=" & rsDefTmp("INVSEQNO")
    End If
    If Not GetRS(rsSO003, strSO003, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    OpenSO003 = True
    Exit Function
ChkErr:
  ErrSub Me.Name, "OpenSO003"
End Function
Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
  Me.Enabled = False
    If txtChargeTitle.Text = Empty Then
        MsgBox "����H�������ȡI", vbInformation, "�T��"
        txtChargeTitle.SetFocus
        Me.Enabled = True
        Exit Function
    Else
        '***********************************************************************************
        '#3819 �ˬd��J������ By Kin 2008/03/21
'        If LenB(StrConv(txtChargeTitle.Text, vbFromUnicode)) > txtChargeTitle.MaxLength Then
'            MsgBox "����H��J�����׶W�L����I", vbInformation, "�T��"
'            txtChargeTitle.SetFocus
'            Exit Function
'        End If
        If Not CtlLimit(txtChargeTitle.Text, txtChargeTitle.MaxLength, True) Then Me.Enabled = True: txtChargeTitle.SetFocus: Exit Function
        
        '***********************************************************************************
    End If
    '***************************************************************************************
    '#3819 �ˬd��J������ By Kin 2008/03/21
    If txtInvTitle.Text <> Empty Then
'        If LenB(StrConv(txtInvTitle.Text, vbFromUnicode)) > txtInvTitle.MaxLength Then
'            MsgBox "�o�����Y��J�����׶W�L����I", vbInformation, "�T��"
'            txtInvTitle.SetFocus
'            Exit Function
'        End If
        If Not CtlLimit(txtInvTitle.Text, txtInvTitle.MaxLength, True) Then Me.Enabled = True: txtInvTitle.SetFocus: Exit Function
    End If
    If txtInvAddress.Text <> Empty Then
'        If LenB(StrConv(txtInvAddress.Text, vbFromUnicode)) > txtInvAddress.MaxLength Then
'            MsgBox "�o���a�}��J�����׶W�L����I", vbInformation, "�T��"
'            txtInvAddress.SetFocus
'            Exit Function
'        End If
        If Not CtlLimit(txtInvAddress.Text, txtInvAddress.MaxLength, True) Then Me.Enabled = True: txtInvAddress.SetFocus: Exit Function
    End If
    '****************************************************************************************
    
    '/***************************************************************/
    '#3773 �T�p���o�@�w�n���νs�P�o�����Y By Kin 2008/02/18
    If cboInvoiceType.ListIndex = 2 Then
        If mskInvNo.Text = Empty Then
            Me.Enabled = True
            MsgBox "�Τ@�s�������ȡI", vbInformation, "�T��"
            mskInvNo.SetFocus
            Exit Function
        End If
        If Trim(txtInvTitle.Text) = "" Then
            Me.Enabled = True
            MsgBox "�o�����Y�����ȡI", vbInformation, "�T��"
            txtInvTitle.SetFocus
            Exit Function
        End If
    End If
    '/*************************************************************/
    
    If EditMode = giEditModeDelete Or chkStop.Value = 1 Then
        If gdaStopDate.GetValue = Empty Then
            Me.Enabled = True
            MsgBox "�п�J���Τ���I", vbInformation, "�T��"
            Exit Function
        End If
    End If
    If mskInvNo.Text <> "" Then
        If Len(Trim(mskInvNo.Text)) <> 8 Then
            Me.Enabled = True
            MsgBox "�Τ@�s�����פ����T�I", vbInformation, "�T��"
            Exit Function
        End If
    End If
    '******************************************************************************
    '#3840 ���դ�OK,�o��]�n�ˮ� By Kin 2008/04/08
    If InStr(1, cboInvoiceType.Text, "�G�p��") <= 0 And mskInvNo.Text <> Empty Then
        If Not InvNoIsOk(mskInvNo.Text, False) Then
            Me.Enabled = True
            MsgBox "�Τ@�s����J���~�I", vbCritical, "�T��"
            mskInvNo.SetFocus
            Exit Function
        End If
    End If
    '******************************************************************************
    '#3920 �T�p���o���ɻ�ĵ�i By Kin 2008/07/31
'    If InStr(1, cboInvoiceType.Text, "�T�p��") > 0 Then
'        If gilUseInv.GetCodeNo & "" <> "" Then
'            MsgBox "�T�p��(��~�H)�o�����i����!", vbInformation, "�T��"
'            Exit Function
'        End If
'
'    End If
    '******************************************************************************
    '#6600 �p�G���o�����ثh���س��ηR�߽X�ݨ䤤�@���� By Kin 2013/10/21
    If blnShowDenRec And blnStartEinvoiceXML Then
        If (Len(gilDenRecCode.GetCodeNo & "") = 0) And (Len(txtLoveNum.Text & "") = 0) Then
            Me.Enabled = True
            MsgBox "���س��ηR�߽X�ݨ䤤�@�Ӧ��ȡI", vbCritical, "�T��"
            gilDenRecCode.SetFocus
            Exit Function
        End If
    End If
    '#6793 �W�[����P�R�߽X����API By Kin 2014/06/13
    '#6793 ������ҿ��~�A��Ƨ令���n�۰ʲM�� For Jacy By Kin 2014/07/07
    Dim APIMsg As String
    Dim UrlRespone As String
    Dim blnNeedVd As Boolean
    Dim urlPara As String
    If (blnStartMobileAPI) Then
        Select Case EditMode
            Case giEditModeEdit
                If (rsDefTmp("CarrierTypeCode") & "" <> gilCarrierTypeCode.GetCodeNo) Or _
                    (rsDefTmp("CarrierId1") & "" <> txtCarrierId1) Then
                    blnNeedVd = True
                End If
            Case giEditModeInsert
                If (Len(gilCarrierTypeCode.GetCodeNo) > 0) Or _
                    (Len(txtCarrierId1.Text) > 0) Then
                    blnNeedVd = True
                End If
        End Select
        If blnNeedVd Then
            urlPara = CreateUrlPara(0)
            txtUrl.Text = strMobileAPIURL & "?" & urlPara
            If Not VdInvWebAPI(0, strMobileAPIURL, urlPara, APIMsg, UrlRespone) Then
                Me.Enabled = True
                'gilCarrierTypeCode.Clear
                'txtCarrierId1.Text = ""
                'txtCarrierId2.Text = ""
                txtUrlResponse.Text = UrlRespone
                MsgBox APIMsg, vbCritical, "�T��"
                Exit Function
            End If
            txtUrlResponse.Text = UrlRespone
        End If
    End If
     If (blnStartLovenumAPI) Then
        If Len(txtLoveNum.Text) > 0 Then
            Select Case EditMode
                Case giEditModeEdit
                    If rsDefTmp("LoveNum") & "" <> txtLoveNum.Text Then
                        blnNeedVd = True
                    End If
                Case giEditModeInsert
                    blnNeedVd = True
            End Select
            If blnNeedVd Then
                urlPara = CreateUrlPara(1)
                txtUrl.Text = strLovenumAPIURL & "?" & urlPara
                If Not VdInvWebAPI(1, strLovenumAPIURL, urlPara, APIMsg, UrlRespone) Then
                    Me.Enabled = True
                    'txtLoveNum.Text = ""
                    MsgBox APIMsg, vbCritical, "�T��"
                    txtUrlResponse.Text = UrlRespone
                    Exit Function
                End If
                txtUrlResponse.Text = UrlRespone
            End If
        End If
     End If
     Me.Enabled = True
    IsDataOk = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "IsDataOK"
End Function
'#6793 ����API�O�_���T 0=��� 1=�R�߽X By Kin 2014/06/13
'Private Function VdInvWebAPI(ByVal APIKind As Integer, _
'                ByVal strURL As String, _
'                ByVal strPara As String, _
'                ByRef Msg As String, Optional ByRef urlResponse As String) As Boolean
'  On Error GoTo ChkErr
'    Dim objHttp As Object
'    Dim i As Integer
'    Dim aryResult() As String
'    Dim blnResult As Boolean
'    Dim Result As String
'    Dim ReturnCode As String
'    Dim ReTurnMsg As String
'    Dim blnOK As Boolean
'    Dim blnExist As Boolean
'    Set objHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
'    objHttp.Open "POST", strURL & "?" & strPara, False
'    objHttp.Send
'    Result = objHttp.ResponseText
'    urlResponse = Result
'    Result = Replace(Result, "{", "")
'    Result = Replace(Result, "}", "")
'    Result = Replace(Result, """", "")
'    aryResult = Split(Result, ",")
'    For i = LBound(aryResult) To UBound(aryResult)
'        Select Case UCase(Replace(Split(aryResult(i), ":")(0), " ", ""))
'            Case UCase("Code")
'                ReturnCode = Replace(Split(aryResult(i), ":")(1), " ", "")
'                If ReturnCode = "200" Then
'                    blnOK = True
'                Else
'                    blnOK = False
'                End If
'            Case UCase("Msg")
'                ReTurnMsg = Split(aryResult(i), ":")(1)
'            Case UCase("isExist")
'                If UCase(Replace(Split(aryResult(i), ":")(1), " ", "")) = "Y" Then
'                    blnExist = True
'                Else
'                    blnExist = False
'                End If
'        End Select
'    Next
'    If blnOK And blnExist Then
'        Select Case APIKind
'            Case 0
'                Msg = "������XAPI���Ҧ��\�I"
'            Case 1
'                Msg = "�R�߽XAPI���Ҧ��\�I"
'        End Select
'    End If
'    If Not blnExist Then
'        Select Case APIKind
'            Case 0
'                Msg = "������XAPI���ҥ��ѡI"
'            Case 1
'                Msg = "�R�߽XAPI���ҥ��ѡI"
'        End Select
'    End If
'    If Not blnOK Then
'        If Len(Msg) = 0 Then
'            Select Case APIKind
'                Case 0
'                    Msg = "������XAPI���ҥ��ѡI"
'                Case 1
'                    Msg = "�R�߽XAPI���ҥ��ѡI"
'            End Select
'        End If
'        Msg = Msg & vbCrLf & "���~�T���X�G" & ReturnCode & " / " & ReTurnMsg
'    End If
'
'    VdInvWebAPI = blnOK And blnExist
'On Error Resume Next
'    Set objHttp = Nothing
'    Exit Function
'ChkErr:
'    If Not objHttp Is Nothing Then
'        Set objHttp = Nothing
'    End If
'    Msg = Err.Description
'End Function

Private Sub txtInvSeqNoQ1_KeyPress(keyAscii As Integer)
  On Error Resume Next
    If keyAscii = 8 Then Exit Sub
    If keyAscii >= 48 And keyAscii <= 57 Then Exit Sub
    keyAscii = 0

End Sub

Private Sub txtInvSeqNoQ2_KeyPress(keyAscii As Integer)
  On Error Resume Next
    If keyAscii = 8 Then Exit Sub
    If keyAscii >= 48 And keyAscii <= 57 Then Exit Sub
    keyAscii = 0
End Sub
Public Sub FunctionKey(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Select Case KeyCode
           '----------------------------------------------------
           Case vbKeyF2  '   F3:�s��, �۷����Ucmdsave
                If EditMode = giEditModeView Then Exit Sub
                If Not ChkGiList(KeyCode) Then Exit Sub
                UpdateGo
           '----------------------------------------------------
           Case vbKeyEscape
                CancelGo
           Case vbKeyF3
                If frmSearch.Enabled Then
                    If cmdQuery.Enabled Then cmdQuery.Value = True
                End If
    End Select
    
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub
Private Sub DefaultValue()
  On Error Resume Next
    Dim rsSo001 As New ADODB.Recordset
    Dim strSo001Sql As String
    '#3929 �h�W�[�o�������P�Τ@�s�� By Kin 2008/06/02
    '#6438 �۰ʱa�J(1) �o���γ~�B(2) ���س��B(3) ���ؤ���B(4) �ӽйq�l�O����o����� By Kin 2013/03/19

    strSo001Sql = "Select A.CustName,A.InstAddrNo,A.InstAddress,A.ChargeAddrNo,B.InvTitle,B.InvAddress," & _
                    "A.ChargeAddress,A.MailAddrNo,A.MailAddress,B.InvoiceType,B.InvNo, " & _
                    "B.ApplyInvDate,B.InvPurposeCode,B.InvPurposeName,B.DenRecCode,B.DenRecName,B.DenRecDate " & _
                    " From " & GetOwner & "SO001 A," & GetOwner & "SO002 B " & _
                    " Where A.Custid=" & lngCustID & " And A.CompCode=" & gCompCode & _
                    " And B.CustId=A.CustId And B.CompCode=A.CompCode" & _
                    " And B.ServiceType='" & gServiceType & "'"
    Set rsSo001 = gcnGi.Execute(strSo001Sql)
    If rsSo001.EOF Or rsSo001.BOF Then Exit Sub
    txtChargeTitle.Text = rsSo001("CustName") & ""
    
    '#3773 �o�����Y�w�]�P����H�P By Kin 2008/02/18
    '#3929 ���SO002 By Kin 2008/06/02
    If Not IsNull(rsSo001("InvTitle")) Then
        txtInvTitle.Text = rsSo001("InvTitle") & ""
    Else
        txtInvTitle.Text = rsSo001("CustName") & ""
    End If
    
    
    '#3929 �o���a�}�w�]��SO002�����
    txtInvAddress = rsSo001("InvAddress") & ""
    
    If rsSo001("ChargeAddrNo") > 0 And Not IsNull(rsSo001("ChargeAddrNo")) Then
        ga1ChargeAddress.AddrNo = rsSo001("ChargeAddrNo")
        ga1ChargeAddress.AddrString = rsSo001("ChargeAddress") & ""
    End If
    If rsSo001("MailAddrNo") > 0 And Not IsNull(rsSo001("MailAddrNo")) Then
        ga1MailAddress.AddrNo = rsSo001("MailAddrNo")
        ga1MailAddress.AddrString = rsSo001("MailAddress")
    End If
    '********************************************************************
    '#3929 �h�W�[�o�������P�Τ@�s�� By Kin 2008/06/02
    If Not IsNull(rsSo001("InvoiceType")) Then
        Select Case rsSo001("InvoiceType") & ""
            Case "0"
                cboInvoiceType.ListIndex = 0
            Case "2"
                cboInvoiceType.ListIndex = 1
            Case "3"
                cboInvoiceType.ListIndex = 2
            Case Else
                cboInvoiceType.ListIndex = 1
        End Select
    End If
    If Not IsNull(rsSo001("InvNo")) Then
        mskInvNo.Text = rsSo001("InvNo") & ""
    End If
    '#5885 �νsNull And �G�p�� And SO041.invoicekind =1 �n�]���q�l�o�� By Kin 2011/03/11
    If mskInvNo.Text & "" = Empty And cboInvoiceType.ListIndex = 1 And intINVOICEKIND = 1 Then
        cboDenRec.ListIndex = 1
        gdaApplyInvDate.SetValue ""
    Else
        'gdaApplyInvDate.SetValue Now
    End If
    '********************************************************************
     '#6438 �۰ʱa�J(1) �o���γ~�B(2) ���س��B(3) ���ؤ���B(4) �ӽйq�l�O����o����� By Kin 2013/03/19
     If Len(rsSo001("InvPurposeCode") & "") > 0 Then
        gilUseInv.SetCodeNo rsSo001("InvPurposeCode") & ""
        gilUseInv.Query_Description
     End If
     If Len(rsSo001("DenRecCode") & "") > 0 Then
        gilDenRecCode.SetCodeNo rsSo001("DenRecCode") & ""
        gilDenRecCode.Query_Description
     End If
     If Len(rsSo001("DenRecDate") & "") > 0 Then
        gidDenRecDate.SetValue rsSo001("DenRecDate") & ""
     End If
     If Len(rsSo001("ApplyInvDate") & "") > 0 Then
        gdaApplyInvDate.SetValue rsSo001("ApplyInvDate") & ""
        gdaApplyInvDate.Enabled = True
     End If
    
    '#6049 ���ի�վ�,�G�p���n��CustName,�o���γ~�B���س��B���ؤ�����O���ť� By Kin 2011/06/21
    '#6438 ���n�O���ťդF By Kin 2013/04/22
'    If InStr(1, cboInvoiceType.Text, "�G�p��") > 0 Then
'        txtInvTitle.Text = rsSo001("CustName") & ""
'        gilUseInv.Clear
'        gilDenRecCode.Clear
'        gidDenRecDate.SetValue ""
'    End If
    
    Call CloseRecordset(rsSo001)
End Sub
Private Function SchIsDataOK() As Boolean
  On Error GoTo ChkErr
    
    SchIsDataOK = True
    '#3929 ����H���comBox���� By Kin 2008/05/26
    If cboChargeTitleQ.Text <> "" Then Exit Function
    'If txtChargeTitleQ.Text <> "" Then Exit Function
    If txtInvSeqNoQ1.Text <> "" Then Exit Function
    If txtInvSeqNoQ2.Text <> "" Then Exit Function
    If txtInvTitleQ.Text <> "" Then Exit Function
    If mskInvNoQ.Text <> "" Then Exit Function
    If chkChargeAddr.Value Then
        If ga2Address.GetAddrString <> Empty Then Exit Function
    End If
    SchIsDataOK = False
    
    Exit Function
ChkErr:
    
    SchIsDataOK = False
    ErrSub Me.Name, "SchIsDataOK"
End Function
'#3785 �W�[�o���γ~ By Kin 2008/04/03
Private Sub SubGil()
  On Error GoTo ChkErr
    If gilUseInv.GetCodeNo & "" = "" Then
        SetgiList gilUseInv, "CodeNo", "Description", "CD095"
    End If
    '**************************************************************************************************
    '#3785 ���դ�OK,�p�G�O�T�p���o�������ܨ�ѦҸ���1 By Kin 2008/05/07
    If InStr(1, cboInvoiceType.Text, "�T�p��") > 0 Then
        gilUseInv.Filter = " Where (RefNO is NULL And StopFlag<>1) or (RefNo<>999 And RefNo<>1) "
    Else
        gilUseInv.Filter = " Where (RefNO<>999 OR RefNO is NULL) And StopFlag<>1 "
    End If
    '**************************************************************************************************
    SetLst gilDenRecCode, "CodeNo", "Description", 3, 12, 405, 1620, "CD110", , True
    '#6371 �W�[�Ȥ�������O���X By Kin 2013/02/21
    SetLst gilCarrierTypeCode, "CodeNo", "Description", 3, 12, 1000, 2000, "CD122", , True
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subGil"
End Sub
'#3929�n�令��ƿ� By Kin 2008/05/26
Private Sub DefineRs()
  On Error GoTo ChkErr
    Dim rsDef138 As New ADODB.Recordset
    Dim i As Long
    If Not GetRS(rsDef138, "Select * From SO138 Where 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    
    With rsDefTmp
        If .State = adStateOpen Then
            .Close
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        
        '***********************************************************
        '#3929 �P�_�O�_�n�ƿ� By Kin 2008/06/02
        If blnMutilChoice Then
            .Fields.Append ("Choice"), adBSTR, 1, adFldIsNullable
        End If
        '***********************************************************
        
    End With
    For i = 0 To rsDef138.Fields.Count - 1
        
        With rsDefTmp
            If UCase(rsDef138.Fields(i).Name) = "STOPDATE" Then
                .Fields.Append "StopDate", adBSTR, 20
            Else
                .Fields.Append rsDef138.Fields(i).Name, adBSTR, rsDef138.Fields(i).DefinedSize, adFldIsNullable
            End If
        End With
    Next i
    rsDefTmp.Open
    On Error Resume Next
    Call CloseRecordset(rsDef138)
    Exit Sub
ChkErr:
  ErrSub Me.Name, "DefineRs"
End Sub
Private Sub ShowProposer()
  On Error GoTo ChkErr
    Dim rsProposer As New ADODB.Recordset
    Dim strQry As String
    Dim rs106Name As New Recordset
    '#3982 �W�[SO106.AccountOwner By Kin 2008/07/16
    cboChargeTitleQ.Clear
    strQry = "Select A.AccountName CustName From " & GetOwner & "SO106 A" & _
             " Where A.CustId=" & gCustId & _
             " And A.CompCode=" & gCompCode & _
             " And A.AccountID='" & strAccountNo & "'" & _
             " Union All " & _
             "Select B.CustName From " & GetOwner & "SO001 B " & _
             " Where B.CustId=" & gCustId & _
             " And B.CompCode=" & gCompCode & _
             " Union All " & _
             "Select C.DeclarantName CustName From " & GetOwner & "SO004 C " & _
             " Where C.CustId=" & gCustId & _
             " And C.CompCode=" & gCompCode & _
             " And (C.PRDate Is Null OR C.InstDate > C.PRDate)" & _
             " Union All " & _
             " Select D.ChargeTitle CustName From " & GetOwner & "SO138 D " & _
             " Where " & IIf(strAccountOwner = Empty, "ChargeTitle is Null", "ChargeTitle='" & strAccountOwner & "'")

    strQry = "Select Distinct(A.CustName) From (" & strQry & ") A"
    
    If Not GetRS(rsProposer, strQry, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Sub
    rsProposer.MoveFirst
    
    Do While Not rsProposer.EOF
        If rsProposer("CustName") & "" <> "" Then
            cboChargeTitleQ.AddItem rsProposer("CustName") & ""
        End If
        rsProposer.MoveNext
    Loop
    
    If cboChargeTitleQ.ListCount > 0 Then
        strQry = "Select AccountName From " & GetOwner & "SO106" & _
                " Where CustId=" & gCustId & _
                " And CompCode=" & gCompCode & _
                " And AccountId='" & strAccountNo & "'"
        Set rs106Name = gcnGi.Execute(strQry)
        If Not rs106Name.EOF Then
            cboChargeTitleQ.Text = rs106Name(0)
        End If
        '�p�G�b���Ҧ��H����ƭnShow�X�b��Ҧ��H����� By Kin 2008/07/16
        If strAccountOwner <> "" Then
            cboChargeTitleQ.Text = strAccountOwner
        End If
    End If
    On Error Resume Next
    CloseRecordset rsProposer
    CloseRecordset rs106Name
    Exit Sub
ChkErr:
    ErrSub Me.Name, "ShowProposer"
End Sub

'#3929 �W�[�v������ By Kin 2008/06/02
Private Sub UserPermissionGo(blnFlag As Boolean)
  On Error GoTo ChkErr
    blnAdd = GetUserPriv("SO114FA", "(SO114FA1)")
    blnEdit = GetUserPriv("SO114FA", "(SO114FA2)")
    blnDele = GetUserPriv("SO114FA", "(SO114FA3)")
    '#3920 �W�[�v������ By Kin 2008/08/04
    gilUseInv.Enabled = GetUserPriv("SO114FA", "(SO114FA4)")
'    MenuEnabled GetUserPriv("SO114FA", "(SO114FA1)") And blnFlag, _
'                GetUserPriv("SO114FA", "(SO114FA2)") And rsSO138.RecordCount > 0 And blnFlag, _
'                GetUserPriv("SO114FA", "(SO114FA3)") And rsSO138.RecordCount > 0 And blnFlag, Not blnFlag, , True
'
  Exit Sub
ChkErr:
    ErrSub Me.Name, "UserPermissionGo"
End Sub

