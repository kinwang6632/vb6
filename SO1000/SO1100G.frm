VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSO1100G 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�ӽ���b�O�� [SO1100G]"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   1425
   ClientWidth     =   11880
   Icon            =   "SO1100G.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNote 
      Height          =   450
      Left            =   750
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  '�������b
      TabIndex        =   73
      Top             =   3450
      Width           =   8235
   End
   Begin VB.CommandButton cmdACHDetail 
      Caption         =   "ACH���v�Ӷ��d��"
      Height          =   315
      Left            =   7740
      TabIndex        =   68
      Top             =   1800
      Width           =   1905
   End
   Begin VB.Frame fraData 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      Left            =   120
      TabIndex        =   35
      Top             =   -30
      Width           =   11625
      Begin VB.TextBox txtAchSN 
         Height          =   315
         Left            =   8640
         MaxLength       =   20
         TabIndex        =   71
         Top             =   2310
         Width           =   1905
      End
      Begin VB.ComboBox cboProposer 
         Height          =   300
         ItemData        =   "SO1100G.frx":0442
         Left            =   1020
         List            =   "SO1100G.frx":0444
         TabIndex        =   0
         Top             =   150
         Width           =   2355
      End
      Begin VB.CommandButton cmdCreateINV 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�o�����Y���@"
         Enabled         =   0   'False
         Height          =   315
         Left            =   210
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   66
         Top             =   3090
         Width           =   1485
      End
      Begin VB.CommandButton cmdInvSeqNo 
         Caption         =   "��ܵo�����Y"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5610
         TabIndex        =   65
         Top             =   3090
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CheckBox chkDeAut 
         Caption         =   "�������v"
         Height          =   210
         Left            =   3060
         TabIndex        =   27
         Top             =   2460
         Width           =   1065
      End
      Begin VB.Frame FraACH 
         Height          =   1440
         Left            =   7560
         TabIndex        =   60
         Top             =   830
         Width           =   3915
         Begin VB.CheckBox chkToACH 
            Caption         =   "��ACH���v"
            Height          =   210
            Left            =   3210
            TabIndex        =   23
            Top             =   90
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.TextBox txtACHcustID 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1230
            MaxLength       =   20
            TabIndex        =   22
            Top             =   230
            Width           =   1665
         End
         Begin CS_Multi.CSmulti csmACH 
            Height          =   375
            Left            =   60
            TabIndex        =   24
            Top             =   570
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   661
            ButtonCaption   =   "ACH���v����O"
         End
         Begin VB.Label lblACHid 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "ACH�Τḹ�X"
            Height          =   180
            Left            =   60
            TabIndex        =   63
            Top             =   300
            Width           =   1110
         End
         Begin VB.Label lblState 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���A : "
            Height          =   180
            Left            =   2010
            TabIndex        =   62
            Top             =   1110
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�����v"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   2640
            TabIndex        =   61
            Top             =   1110
            Visible         =   0   'False
            Width           =   780
         End
      End
      Begin VB.CheckBox chkAddAcc 
         Caption         =   "�ݦ����s�g��,�U���X�b�H���b��ú�O"
         Height          =   255
         Left            =   8250
         TabIndex        =   31
         Top             =   2745
         Width           =   3225
      End
      Begin Gi_Multi.GiMulti gimCitem 
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   4290
         Visible         =   0   'False
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   661
         ButtonCaption   =   "�� �O �� ��"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin VB.TextBox txtCVC2 
         Height          =   315
         Left            =   10950
         MaxLength       =   3
         TabIndex        =   7
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox chkFore2 
         Caption         =   "�~�y�H�h"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   6480
         TabIndex        =   13
         Top             =   1200
         Width           =   1035
      End
      Begin VB.CheckBox chkFore 
         Caption         =   "�~�y�H�h"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5820
         TabIndex        =   2
         Top             =   210
         Width           =   1065
      End
      Begin VB.TextBox txtAccountNo 
         Height          =   315
         Left            =   7920
         MaxLength       =   16
         TabIndex        =   6
         Top             =   480
         Width           =   2370
      End
      Begin VB.TextBox txtHide 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  '�Ȥ�
         Left            =   6120
         Locked          =   -1  'True
         PasswordChar    =   "*"
         TabIndex        =   10
         Text            =   "*******"
         Top             =   810
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtIntroId 
         Height          =   285
         Left            =   4380
         MaxLength       =   10
         TabIndex        =   15
         Top             =   1470
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Appearance      =   0  '����
         Height          =   285
         Left            =   7230
         Picture         =   "SO1100G.frx":0446
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   16
         Top             =   1470
         Width           =   285
      End
      Begin VB.TextBox txtAccountOwner 
         Height          =   315
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1140
         Width           =   2355
      End
      Begin VB.TextBox txtProposer 
         Height          =   315
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   67
         Top             =   150
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.TextBox txtAccOwnerID 
         Height          =   315
         IMEMode         =   2  '����
         Left            =   5130
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1140
         Width           =   1215
      End
      Begin VB.CheckBox chkStop 
         Caption         =   "����"
         Height          =   210
         Left            =   240
         TabIndex        =   25
         Top             =   2460
         Width           =   675
      End
      Begin VB.TextBox txtID 
         Height          =   315
         IMEMode         =   2  '����
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   1
         Top             =   150
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskCardExpDate 
         Height          =   315
         Left            =   6120
         TabIndex        =   9
         Tag             =   "OK"
         Top             =   810
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin Gi_Time.GiTime gdtAcceptTime 
         Height          =   315
         Left            =   1020
         TabIndex        =   20
         Top             =   2070
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
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
      Begin prjGiList.GiList gilAcceptName 
         Height          =   315
         Left            =   3840
         TabIndex        =   21
         Top             =   2070
         Width           =   3030
         _ExtentX        =   5345
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
         FldWidth1       =   880
         FldWidth2       =   1820
         F2Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin prjGiList.GiList gilMediaCode 
         Height          =   285
         Left            =   1020
         TabIndex        =   14
         Top             =   1470
         Width           =   2355
         _ExtentX        =   4154
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
         FldWidth1       =   405
         FldWidth2       =   1620
         F2Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin Gi_Date.GiDate gdaContiDate 
         Height          =   285
         Left            =   3840
         TabIndex        =   18
         Top             =   1770
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ReplaceText     =   -1  'True
      End
      Begin Gi_Date.GiDate gdaSendDate 
         Height          =   285
         Left            =   1020
         TabIndex        =   17
         Top             =   1770
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ReplaceText     =   -1  'True
      End
      Begin Gi_Date.GiDate gdaSnactionDate 
         Height          =   285
         Left            =   6120
         TabIndex        =   19
         Top             =   1770
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ReplaceText     =   -1  'True
      End
      Begin Gi_Date.GiDate gdaStopDate 
         Height          =   285
         Left            =   1860
         TabIndex        =   26
         Top             =   2400
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ReplaceText     =   -1  'True
      End
      Begin prjGiList.GiList gilBankCode 
         Height          =   315
         Left            =   4290
         TabIndex        =   5
         Top             =   480
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth1       =   405
         FldWidth2       =   1850
         F2Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin prjGiList.GiList gilCardName 
         Height          =   315
         Left            =   1020
         TabIndex        =   8
         Top             =   810
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth1       =   405
         FldWidth2       =   1620
         F2Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin prjGiList.GiList gilCMCode 
         Height          =   315
         Left            =   7920
         TabIndex        =   3
         Top             =   150
         Width           =   2370
         _ExtentX        =   4180
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
         FldWidth1       =   420
         FldWidth2       =   1620
         F2Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin CS_Multi.CSmulti csmCitem2 
         Height          =   375
         Left            =   4200
         TabIndex        =   30
         Top             =   2700
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   661
         ButtonCaption   =   "�ݦ� ���O����"
      End
      Begin CS_Multi.CSmulti csmCitem1 
         Height          =   375
         Left            =   180
         TabIndex        =   29
         Top             =   2700
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   661
         ButtonCaption   =   "�g�� ���O����"
      End
      Begin prjGiList.GiList gilPTCode 
         Height          =   315
         Left            =   1020
         TabIndex        =   4
         Top             =   480
         Width           =   2355
         _ExtentX        =   4154
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
         FldWidth1       =   405
         FldWidth2       =   1620
         TableName       =   "CD019"
         F2Corresponding =   -1  'True
      End
      Begin Gi_Date.GiDate gdaAuthStopDate 
         Height          =   285
         Left            =   5770
         TabIndex        =   28
         Top             =   2400
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ReplaceText     =   -1  'True
      End
      Begin VB.Label lblApplication 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ӽЮѳ渹"
         Height          =   180
         Left            =   7620
         TabIndex        =   72
         Top             =   2370
         Width           =   900
      End
      Begin VB.Label lblInvSeqNo 
         Height          =   195
         Left            =   2820
         TabIndex        =   70
         Top             =   3150
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "�o���y����:"
         Height          =   255
         Left            =   1770
         TabIndex        =   69
         Top             =   3150
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���v�I���"
         Height          =   180
         Left            =   4830
         TabIndex        =   64
         Top             =   2460
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "�I�ں���"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   180
         TabIndex        =   59
         Top             =   540
         Width           =   780
      End
      Begin VB.Label lblCVC2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ˬd�X"
         Height          =   180
         Left            =   10350
         TabIndex        =   58
         Top             =   540
         Width           =   540
      End
      Begin VB.Label lblUpdTime 
         AutoSize        =   -1  'True
         Caption         =   "yyyy/mm/dd hh:mm:ss"
         Height          =   195
         Left            =   9810
         TabIndex        =   57
         Top             =   3510
         Width           =   1605
      End
      Begin VB.Label lblUpdEn 
         AutoSize        =   -1  'True
         Caption         =   "XXXXXXXX"
         Height          =   195
         Left            =   9810
         TabIndex        =   56
         Top             =   3750
         Width           =   990
      End
      Begin VB.Label lblUTime 
         AutoSize        =   -1  'True
         Caption         =   "���ʮɶ�:"
         Height          =   195
         Left            =   8970
         TabIndex        =   55
         Top             =   3510
         Width           =   795
      End
      Begin VB.Label lblUEn 
         AutoSize        =   -1  'True
         Caption         =   "���ʤH��:"
         Height          =   195
         Left            =   8970
         TabIndex        =   54
         Top             =   3750
         Width           =   825
      End
      Begin VB.Label lblCMCode 
         Caption         =   "���O�覡"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   7110
         TabIndex        =   53
         Top             =   210
         Width           =   750
      End
      Begin VB.Label lblAccountNo 
         AutoSize        =   -1  'True
         Caption         =   "�b��/�d��"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   7095
         TabIndex        =   52
         Top             =   540
         Width           =   765
      End
      Begin VB.Label lblCardExpDate 
         Caption         =   "�H�Υd���Ĵ��� (���MM/YYYY)"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3450
         TabIndex        =   51
         Top             =   870
         Width           =   2655
      End
      Begin VB.Label lblAcceptName 
         AutoSize        =   -1  'True
         Caption         =   "���z�H��"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3060
         TabIndex        =   50
         Top             =   2130
         Width           =   795
      End
      Begin VB.Label lblAcceptTime 
         AutoSize        =   -1  'True
         Caption         =   "���z�ɶ�"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   180
         TabIndex        =   49
         Top             =   2130
         Width           =   795
      End
      Begin VB.Label lblMediaCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���дC��"
         Height          =   180
         Left            =   180
         TabIndex        =   48
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lblIntro 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���ФH�N�X"
         Height          =   180
         Left            =   3450
         TabIndex        =   47
         Top             =   1530
         Width           =   900
      End
      Begin VB.Label lblIntroName 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '��u�T�w
         Height          =   285
         Left            =   5475
         TabIndex        =   32
         Top             =   1470
         Width           =   1755
      End
      Begin VB.Label lblAccountOwner 
         Alignment       =   1  '�a�k���
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�b���Ҧ��H"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   90
         TabIndex        =   46
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label lblProposer 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ӽФH"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   330
         TabIndex        =   45
         Top             =   210
         Width           =   555
      End
      Begin VB.Label lblAccNameID 
         Caption         =   "�b���Ҧ��H�����Ҹ�"
         Height          =   195
         Left            =   3450
         TabIndex        =   44
         Top             =   1200
         Width           =   1650
      End
      Begin VB.Label lblDelDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�e����"
         Height          =   180
         Left            =   180
         TabIndex        =   43
         Top             =   1830
         Width           =   720
      End
      Begin VB.Label lblContiDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�Ȧ�֦L���"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   2680
         TabIndex        =   42
         Top             =   1830
         Width           =   1080
      End
      Begin VB.Label lblSnDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�֭���"
         ForeColor       =   &H00C000C0&
         Height          =   180
         Left            =   5310
         TabIndex        =   41
         Top             =   1830
         Width           =   720
      End
      Begin VB.Label lblStopDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���Τ��"
         Height          =   180
         Left            =   1050
         TabIndex        =   40
         Top             =   2460
         Width           =   720
      End
      Begin VB.Label lblID 
         Caption         =   "�����Ҹ�"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3450
         TabIndex        =   39
         Top             =   210
         Width           =   795
      End
      Begin VB.Label lblBankCode 
         AutoSize        =   -1  'True
         Caption         =   "�Ȧ�"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3795
         TabIndex        =   38
         Top             =   540
         Width           =   360
      End
      Begin VB.Label lblCardName 
         AutoSize        =   -1  'True
         Caption         =   "�H�Υd�O"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   870
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�Ƶ�"
         Height          =   180
         Left            =   240
         TabIndex        =   36
         Top             =   3630
         Width           =   375
      End
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   3420
      Left            =   60
      TabIndex        =   33
      Top             =   4080
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   6033
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
   Begin CS_Multi.CSmulti csmTest 
      Height          =   375
      Left            =   150
      TabIndex        =   74
      Top             =   7650
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   661
      ButtonCaption   =   "ACH���v����O"
   End
End
Attribute VB_Name = "frmSO1100G"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmer: Power Hammer
' Last Modify: 2002/06/27

Option Explicit
Private lngEditMode As giEditModeEnu ' �O���ثe�b�s��B�s�W���˵��Ҧ�
Private WithEvents rsSO106 As ADODB.Recordset
Attribute rsSO106.VB_VarHelpID = -1
Private intCD031RefNo As Integer
Private blnVAcc As Boolean
Private blnACHAccountNoCanEdit As Boolean
Private strRealAccNo As String
Private strAchChange As String
Private blnAchChgClick As Boolean
Private blnAddFlag As Boolean
Private blnACHCustID As Boolean
Private strInvSeqNo As String
Private blnCallSO106A As Boolean
Private strACHID As String
Private strACHDesc As String
Private strFirstACHID As String
Private strFirstACHDesc As String
Private blnHaveUpd As Boolean
Private blnNoShowMsg As Boolean
Private blnInsSO138 As Boolean
Private blnEnableStop As Boolean
Private blnUCCodeIsNull As Boolean
Private strAccountName As String
Private posAchContition As String
Private AchTypeCondition As String
Private blnStartPost As Boolean
Private Sub chkFore_Click()
  On Error Resume Next
    If chkFore.Value Then
        txtID.MaxLength = 30
    Else
        txtID.MaxLength = 10
        txtID = Left(txtID, 10)
    End If
End Sub
'#3982 �ˬdSO138�O�_���b���Ҧ��H By Kin 2008/07/16
Private Function ChkSO138() As Boolean
  On Error GoTo ChkErr
    Dim strQry As String
    strQry = Empty
    If txtAccountOwner.Text & "" <> "" Then
        strQry = "Select Count(*) From " & GetOwner & "SO138" & _
                " Where ChargeTitle='" & Trim(txtAccountOwner.Text) & "'"
            
        If gcnGi.Execute(strQry)(0) > 0 Then
        
            ChkSO138 = True
        Else
            ChkSO138 = False
        End If
    
    End If
    Exit Function
ChkErr:
  ErrSub Me.Name, "ChkSO138"
End Function
'#3982 �p�G�L��ƫh�s�W�@����Ʀ�138�æ۰ʥN�JH���� By Kin 2008/07/16
Private Sub InsSO138()
  On Error GoTo ChkErr
    Dim rsSo001 As New ADODB.Recordset
    Dim rsSo002 As New ADODB.Recordset
    
    Dim rsSO138Tmp As New ADODB.Recordset
    Dim strSo001Sql As String
    Dim strSo002Sql As String
    Dim aMsg As String
    Dim blnElectron As Boolean
    blnElectron = False
    '#5945 �N���殳��,���SO002����T
    'blnElectron = SetInvKind(aMsg)
    If aMsg <> Empty Then
        If MsgBox(aMsg, vbQuestion + vbYesNo, "�߰�") = vbYes Then
            blnElectron = True
        End If
    End If
    strSo001Sql = "Select A.CustName,A.InstAddrNo,A.InstAddress,A.ChargeAddrNo,B.InvTitle,B.InvAddress," & _
                    "A.ChargeAddress,A.MailAddrNo,A.MailAddress,B.InvoiceType,B.InvNo, " & _
                    "B.ApplyInvDate,B.InvPurposeCode,B.InvPurposeName,B.DenRecCode,B.DenRecName,B.DenRecDate " & _
                    " From " & GetOwner & "SO001 A," & GetOwner & "SO002 B " & _
                    " Where A.Custid=" & gCustId & " And A.CompCode=" & gCompCode & _
                    " And B.CustId=A.CustId And B.CompCode=A.CompCode" & _
                    " And B.ServiceType='" & gServiceType & "'"
    Set rsSo001 = gcnGi.Execute(strSo001Sql)
    strSo002Sql = "SELECT * FROM " & GetOwner & "SO002 " & _
                " WHERE CUSTID = " & gCustId & _
                " AND COMPCODE = " & gCompCode & _
                " AND SERVICETYPE = '" & gServiceType & "'"
    If Not GetRS(rsSo002, strSo002Sql, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    
    If Not GetRS(rsSO138Tmp, "Select * From " & GetOwner & "SO138 Where 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If rsSo001.EOF Or rsSo001.BOF Then Exit Sub
    With rsSO138Tmp
        .AddNew
        .Fields("InvSeqNo") = Get138SeqNo
        
        .Fields("ChargeTitle") = txtAccountOwner.Text
        .Fields("InvTitle") = txtAccountOwner.Text
        .Fields("InvAddress") = rsSo001("InvAddress") & ""
        '*********************************************************************************
        '#4210 �p�G�O�����b���n�a�JSO001���˾��a�} By Kin 2008/11/06
        If intCD031RefNo = 2 Or intCD031RefNo = 4 Or intCD031RefNo = 5 Then
            If rsSo001("ChargeAddrNo") > 0 And Not IsNull(rsSo001("ChargeAddrNo")) Then
                .Fields("ChargeAddrNo") = rsSo001("ChargeAddrNo")
                .Fields("ChargeAddress") = rsSo001("ChargeAddress") & ""
            End If
            If rsSo001("MailAddrNo") > 0 And Not IsNull(rsSo001("MailAddrNo")) Then
                .Fields("MailAddrNo") = rsSo001("MailAddrNo")
                .Fields("MailAddress") = rsSo001("MailAddress")
            End If
        Else
            If rsSo001("InstAddrNo") > 0 And Not IsNull(rsSo001("InstAddress")) Then
                .Fields("ChargeAddrNo") = rsSo001("InstAddrNo")
                .Fields("ChargeAddress") = rsSo001("InstAddress") & ""
                .Fields("MailAddrNo") = rsSo001("InstAddrNo")
                .Fields("MailAddress") = rsSo001("InstAddress") & ""
            End If
        End If
        '*********************************************************************************
        If Not IsNull(rsSo001("InvoiceType")) Then
            Select Case rsSo001("InvoiceType") & ""
                Case "0"
                    .Fields("InvoiceType") = 0
                Case "2"
                    .Fields("InvoiceType") = 2
                '**************************************************************************************************
                '#4210 �p�G�O3�p���o��,�o�����Y�n�aSO002.InvTitle By Kin 2008/11/06
                Case "3"
                    .Fields("InvoiceType") = 3
                    .Fields("InvTitle") = IIf(IsNull(rsSo001("InvTitle")), txtAccountOwner.Text, rsSo001("InvTitle") & "")
                '***************************************************************************************************
                Case Else
                    .Fields("InvoiceType") = 2
            End Select
        End If
        If Not IsNull(rsSo001("InvNo")) Then
            .Fields("InvNo") = rsSo001("InvNo") & ""
            '�p�G�Τ@�s�����Ȥ@�w�n�]�� �q�l�p����o�� By Kin 2010/08/06
            blnElectron = False
        End If
        .Fields("StopFlag") = 0
        .Fields("Updtime") = GetDTString(RightNow)
        .Fields("NEWUPDTIME") = RightNow
        .Fields("UpdEn") = garyGi(1)
        strInvSeqNo = .Fields("InvSeqNo")
        If blnElectron Then
            .Fields("INVOICEKIND") = 1
        Else
            .Fields("INVOICEKIND") = 0
        End If
        '#5945 �o���γ~�B�o�������B�o���}�ߺ����n��SO002����T By Kin 2011/03/24
        '#5945 ���դ�OK,���س��B���ؤ���]�nCopy By Kin 2011/04/20
        If Not rsSo002.EOF Then
            If Len(rsSo002("InvoiceType") & "") > 0 Then
                .Fields("InvoiceType") = rsSo002("InvoiceType") & ""
            End If
            
            If Len(rsSo002("InvoiceKind") & "") > 0 Then
                .Fields("InvoiceKind") = rsSo002("InvoiceKind") & ""
            End If
            
             '#6049 �o���γ~���n��s By Kin 2011/06/17
'            If Len(rsSo002("InvPurposeCode") & "") > 0 Then
'                .Fields("InvPurposeCode") = rsSo002("InvPurposeCode") & ""
'                .Fields("InvPurposeName") = rsSo002("InvPurposeName") & ""
'            End If
            
            If Len(rsSo002("InvNo") & "") > 0 Then
                .Fields("InvoiceKind") = 0
            End If
            
            '#6049 ���س�줣�n��s By Kin 2011/06/17
'            If Len(rsSo002("DenRecCode") & "") > 0 Then
'                .Fields("DenRecCode") = rsSo002("DenRecCode") & ""
'            End If
'
'            If Len(rsSo002("DenRecName") & "") > 0 Then
'                .Fields("DenRecName") = rsSo002("DenRecName") & ""
'            End If
            '#6049 ���ؤ�����n��s By Kin 2011/06/17
'            If Len(rsSo002("DenRecDate") & "") > 0 Then
'                .Fields("DenRecDate") = rsSo002("DenRecDate") & ""
'            End If
'            #6049 �o�����Y�H�W���e����J�����D By Kin 2011/06/17
'            If Len(rsSo002("InvTitle") & "") > 0 Then
'                .Fields("InvTitle") = rsSo002("InvTitle") & ""
'            End If
        End If
         '#6438 �۰ʱa�J(1) �o���γ~�B(2) ���س��B(3) ���ؤ���B(4) �ӽйq�l�O����o����� By Kin 2013/03/19
         '#6438 �P�_�ӽФH�PSO001�O�_�ۦP,�ۦP�~�a�J��� By Kin 2013/04/22
        If cboProposer.Text = rsSo001("CustName") Then
            If Len(rsSo002("InvPurposeCode") & "") > 0 Then
                .Fields("InvPurposeCode") = rsSo002("InvPurposeCode") & ""
                .Fields("InvPurposeName") = rsSo002("InvPurposeName") & ""
            End If
            If Len(rsSo002("DenRecCode") & "") > 0 Then
               .Fields("DenRecCode") = rsSo002("DenRecCode") & ""
               .Fields("DenRecName") = rsSo002("DenRecName") & ""
            End If
            If Len(rsSo002("DenRecDate") & "") > 0 Then
                .Fields("DenRecDate") = rsSo002("DenRecDate") & ""
            End If
            If Len(rsSo002("ApplyInvDate") & "") > 0 Then
               .Fields("ApplyInvDate") = rsSo002("ApplyInvDate") & ""
            End If
        End If
        
        .Update
    End With
    
    On Error Resume Next
    Call CloseRecordset(rsSo001)
    Call CloseRecordset(rsSO138Tmp)
    Call CloseRecordset(rsSo002)
    Exit Sub
ChkErr:
  ErrSub Me.Name, "InsSO138"
End Sub
'#5668 ��ܶg���ʶ��اP�_�O�_�n�]�w���q�l�o�� By Kin 2010/07/19
Private Function SetInvKind(ByRef aMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim aCitemStr As String
    Dim aryCitemStr() As String
    Dim i As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim aTmp As String
    Dim aSQL As String
    Dim iCntRct As Long
    Dim iType As Integer
    Dim strTmpCitem As String
    Dim blnExitFor As Boolean
    aCitemStr = csmCitem1.GetQryFld(1)
    If aCitemStr = "" Then Exit Function
    aryCitemStr = Split(aCitemStr, ",")
    iCntRct = 0
    strTmpCitem = Empty
    For i = LBound(aryCitemStr) To UBound(aryCitemStr)
        aTmp = aryCitemStr(i)
        iType = GetByHouse(aTmp)
        If iType = 0 Then
            aSQL = "SELECT Contmobile,Email   FROM " & GetOwner & "SO004 " & _
                " WHERE (Contmobile IS NOT NULL OR EMail IS NOT NULL )" & _
                " AND CUSTID = " & gCustId & " AND COMPCODE = " & gCompCode & _
                " AND FACISNO IN ( SELECT FACISNO FROM " & GetOwner & "SO003 " & _
                " WHERE CITEMCODE = " & aTmp & " AND FACISNO IS NOT NULL " & _
                " AND CUSTID = " & gCustId & ")"
            
        Else
            aSQL = " SELECT  Null Contmobile,Email  FROM " & GetOwner & "SO002 " & _
                " WHERE CUSTID = " & gCustId & " AND COMPCODE = " & gCompCode & _
                " AND SERVICETYPE IN ( SELECT SERVICETYPE FROM " & GetOwner & "SO003 " & _
                " WHERE CUSTID = " & gCustId & " AND COMPCODE = " & gCompCode & _
                " AND CITEMCODE = " & aTmp & ") " & _
                " AND EMAIL IS NOT NULL " & _
                " UNION ALL " & _
                " SELECT Tel3 Contmobile,Null Email  FROM " & GetOwner & "SO001 " & _
                " WHERE CUSTID = " & gCustId & " AND COMPCODE = " & gCompCode & _
                " AND Tel3 IS NOT NULL "
            aSQL = "Select Distinct * From ( " & aSQL & ") "
        End If
        If Not GetRS(rsTmp, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        Do While Not rsTmp.EOF
             
            If ChkMobileOk(rsTmp("Contmobile") & "") Or ChkEmailOk(rsTmp("Email") & "") Then
                blnExitFor = True
                Exit Do
            End If
            rsTmp.MoveNext
        Loop
        If blnExitFor Then
            iCntRct = UBound(aryCitemStr) + 1
            Exit For
        End If
'        If Val(GetRsValue(aSQL, gcnGi) & "") > 0 Then
'            iCntRct = iCntRct + 1
'        Else
'            If strTmpCitem = Empty Then
'                strTmpCitem = aTmp
'            Else
'                strTmpCitem = strTmpCitem & "�A" & aTmp
'            End If
'        End If
        
    Next i
    If UBound(aryCitemStr) = 0 Then
        SetInvKind = iCntRct = 1
        aMsg = Empty
    Else
        If UBound(aryCitemStr) + 1 <> iCntRct Then
            aMsg = "���O���ءG" & strTmpCitem & vbCrLf & _
                " �L����B�q�l�l���T�A�O�_�]�w���q�l�o���H"
            SetInvKind = False
        Else
            aMsg = Empty
            SetInvKind = True
        End If
        
    End If
    aMsg = Empty
    On Error Resume Next
    Call Close3Recordset(rsTmp)
    Exit Function
ChkErr:
  ErrSub Me.Name, "SetInvKind"
End Function
Private Function Get138SeqNo() As String
  On Error GoTo ChkErr
    Get138SeqNo = RPxx(gcnGi.Execute("SELECT " & GetOwner & "S_SO138_InvSeqNo.NEXTVAL FROM DUAL").GetString & "")
  Exit Function
ChkErr:
    ErrSub Me.Name, "GetSeqNo"
End Function

Private Sub chkFore2_Click()
  On Error Resume Next
    If chkFore2.Value Then
        txtAccOwnerID.MaxLength = 30
    Else
        txtAccOwnerID.MaxLength = 10
        txtAccOwnerID = Left(txtAccOwnerID, 10)
    End If
End Sub

Private Sub chkStop_Click()
  On Error Resume Next
    lblStopDate.ForeColor = IIf(chkStop.Value = 1, vbRed, vbBlack)
    If chkStop.Value <> 1 Then
        gdaStopDate.Text = ""
    Else
        gdaStopDate.Text = Date
        gdaStopDate.SetFocus
    End If
End Sub

Private Sub cmdACHDetail_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
    '*****************************************************************************************
    '#3946 MasterRowId �אּ MasterId By Kin 2008/05/30
    If EditMode = giEditModeInsert Then
        frmSO1100G1.uMasterRowId = 0
    Else
        frmSO1100G1.uMasterRowId = IIf(IsNull(rsSO106("MasterId")), 0, rsSO106("MasterId"))
    End If
    '*****************************************************************************************
    If blnCallSO106A Then
        frmSO1100G1.EditMode = giEditModeEdit
    Else
        frmSO1100G1.EditMode = giEditModeView
    End If
    frmSO1100G1.uInAchId = strACHID
    frmSO1100G1.uInAchDesc = strACHDesc
    frmSO1100G1.uFirstACHDesc = strFirstACHDesc
    frmSO1100G1.uFirstAchId = strFirstACHID
    frmSO1100G1.Show 1
    strACHID = Empty
    strACHDesc = Empty
    strFirstACHID = Empty
    strFirstACHDesc = Empty
    blnCallSO106A = False
    Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdACHDetail"
End Sub

Private Sub cmdCreateINV_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
    With frmSO114FA
        .OldEditMode = Me.EditMode
        .uCustId = gCustId
        If EditMode = giEditModeEdit Then
            .uAccountNo = Trim(txtAccountNo.Text)
            If Not IsNull(rsSO106("InvSeqNo")) Then .uInvSeqNo = rsSO106("InvSeqNo") & ""
            If Not IsNull(rsSO106("InvSeqNo")) Then .uShowAbs = rsSO106("InvSeqNo") & ""
        Else
            If txtAccountNo.Text <> "" Then
                .uAccountNo = Trim(txtAccountNo.Text)
            End If
            .uInvSeqNo = Empty
        End If
        .uAccountOwner = txtAccountOwner.Text
        .uMutilChoice = False
        .uOwnerForm = Me
        .Show , Me
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
  ErrSub Me.Name, "cmdCreateINV"
End Sub

Private Sub cmdFind_Click()
  On Error GoTo ChkErr
    ServiceCommon.FindIntroduce Me
  Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdFind_Click"
End Sub

Private Sub cmdInvSeqNo_Click()
  On Error GoTo ChkErr
    With frmSO1131H
        .uParentForm = Me
        .uMutilChoice = False
        .Show 1
    End With
    strInvSeqNo = Empty
    strInvSeqNo = frmSO1131H.uNewInvSeqNo
   
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdInvSeqNo"
End Sub

'#3236 ACH���ܮɦp�G���O���ئ��Ȼ�Show�Xĵ�i�T��,�ҥH�n�@���ܼưO����Ӫ��� By Kin 2007/05/29
Private Sub csmACH_ButtonClick()
    On Error Resume Next
    strAchChange = csmACH.GetDispStr
    blnAchChgClick = True
End Sub

Private Sub csmACH_Change()
    On Error GoTo ChkErr
    '#3322 �@�}�l�s�W�ɷ|�X�{ĵ�i�T���A�h�[�@��Flag���P�_�A�p�G�@�i�ӬO�s�W�����P�_�A�������X�A���n���W�]��False�A�]�������٬O�n�P�_ By Kin 2007/07/04
    If blnAddFlag Then
        blnAddFlag = False
        Exit Sub
    End If
    If Not blnNoShowMsg Then
        If (csmACH.GetDispStr & "" = "") And (csmCitem1.GetDispStr & "" <> "") And (EditMode <> giEditModeView) Then
            MsgBox "ACH���v����O�w����!!�Э��s�]�w���O����", vbInformation, "ĵ�i�T��"
            csmCitem1.Clear
        End If
    End If
    blnNoShowMsg = False
    Exit Sub
ChkErr:
     ErrSub Name, "csmACH_Change"
End Sub

'#3236 ACH���ܮɦp�G���O���ئ��Ȼ�Show�Xĵ�i�T��By Kin 2007/05/29
Private Sub csmACH_SelectChange()
    On Error GoTo ChkErr
    '���g�LButtonClick�ƥ�
    If blnAchChgClick Then
        
        If (csmACH.GetDispStr <> strAchChange) And (csmCitem1.GetDispStr <> "") Then
            MsgBox "ACH���v����O�w����!!�Э��s�]�w���O����", vbInformation, "ĵ�i�T��"
            csmCitem1.Clear
        End If
    Else 'User�������UClear���s
'        If lngEditMode <> giEditModeView Then
'            If (csmACH.GetDispStr & "" = "") And (csmCitem1.GetDispStr & "" <> "") And (rsSO106("ACHTNo") & "" <> "") Then
'                MsgBox "ACH���v����O�w����!!�Э��s�]�w���O����", vbInformation, "ĵ�i�T��"
'                csmCitem1.Clear
'            End If
'        End If
    End If
    blnAchChgClick = False
    blnNoShowMsg = False
    Exit Sub
ChkErr:
    ErrSub Name, "csmACH_SelectChange"
End Sub

Private Sub csmCitem1_ButtonClick()
  On Error GoTo ChkErr
    Dim strTmp As String
    strTmp = RPxx(Replace(csmACH.GetQryFld(3), "'", "", 1))
    '#7049 �NCitemCode���CD068A By Kin 2015/07/08
    strTmp = "Select CD068A.CitemCode From " & GetOwner & "CD068A Where BillHeadFmt In (" & csmACH.GetQryFld(4) & ")"
    '#3454 �h��SO003.FaciSNo�BSO004.DeclarantName�BSO004.FaciName By Kin 2007/12/18
    If csmACH.GetDispStr <> Empty Then
        '#6577 �W�[StopFlag By Kin 2013/09/04
        SetMsQry csmCitem1, "SELECT " & _
                            " A.CITEMCODE,A.CITEMNAME,  Decode(Nvl(A.StopFlag,0),0,'�_','�O') StopFlag, " & _
                            " A.PERIOD,A.AMOUNT,A.ACCOUNTNO,A.CMNAME,A.STARTDATE," & _
                            "A.STOPDATE,A.SEQNO,A.FaciSNo,B.DeclarantName,B.FaciName FROM " & _
                             GetOwner & "SO003 A," & GetOwner & "SO004 B  WHERE A.CUSTID = " & _
                             gCustId & " AND A.COMPCODE=" & gCompCode & _
                             " AND A.CITEMCODE IN (" & strTmp & ")" & _
                             " AND A.CustId=B.CustId(+) And A.CompCode=B.CompCode(+)" & _
                             " And A.FaciSeqNo=B.SeqNo(+)" & _
                             " ORDER BY A.CLCTDATE,A.CITEMCODE", _
                             , , "�N�X,���O����,����,����,���B,���b�b��,���O�覡,�_�l��,�I���,�Ǹ�,���~�Ǹ�,�ӽФH�m�W,���ئW��", _
                             "1000,1800,500,460,560,1800,850,850,850,770,850,1600,1600", True, , , , , _
                             , , , , , , , , , , , , , , , , , , , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD", , , , 10
    Else
        '#6577 �W�[StopFlag By Kin 2013/09/04
        SetMsQry csmCitem1, "SELECT " & _
                            "A.CITEMCODE,A.CITEMNAME,Decode(Nvl(A.StopFlag,0),0,'�_','�O') StopFlag," & _
                            " A.PERIOD,A.AMOUNT,A.ACCOUNTNO,A.CMNAME,A.STARTDATE," & _
                            "A.STOPDATE,A.SEQNO,A.FaciSNo,B.DeclarantName,B.FaciName FROM " & _
                             GetOwner & "SO003 A," & GetOwner & "SO004 B  WHERE A.CUSTID = " & _
                             gCustId & " AND A.COMPCODE=" & gCompCode & _
                             " AND A.CustId=B.CustId(+) And A.CompCode=B.CompCode(+)" & _
                             " And A.FaciSeqNo=B.SeqNo(+)" & _
                             " ORDER BY A.CLCTDATE,A.CITEMCODE", _
                             , , "�N�X,���O����,����,����,���B,���b�b��,���O�覡,�_�l��,�I���,�Ǹ�,���~�Ǹ�,�ӽФH�m�W,���ئW��", _
                             "1000,1800,500,460,560,1800,850,850,850,770,850,1600,1600", True, , , , , _
                             , , , , , , , , , , , , , , , , , , , , , , "###,###,###", , "EE/MM/DD", "EE/MM/DD", , , , , 10



    End If
  Exit Sub
ChkErr:
    ErrSub Name, "csmCitem1_ButtonClick"
End Sub

Private Sub csmCitem2_ButtonClick()
  On Error GoTo ChkErr
    Dim strTmp As String
    Dim strQry As String
    Dim strOrder As String
    
    '#3454 �h��SO003.FaciSNo�BSO004.DeclarantName�BSO004.FaciName By Kin 2007/12/10
    '#5045 �ΰѼƧP�_�O�_���D��w���O����� By Kin 2009/05/04
    
    strOrder = " ORDER BY A.REALDATE DESC,A.BILLNO DESC,A.SEQNO DESC "
    strTmp = RPxx(Replace(csmACH.GetQryFld(3), "'", "", 1))
    '#7049 �W�[BillHeadFmt���,�i�H���CD068A By Kin 2015/07/08
    strTmp = "Select CD068A.CitemCode From " & GetOwner & "CD068A Where BillHeadFmt In (" & csmACH.GetQryFld(4) & ")"
    If csmACH.GetDispStr <> Empty Then
        strQry = "SELECT DISTINCT CITEMCODE,CITEMNAME,REALPERIOD,SHOULDAMT,ACCOUNTNO,CMNAME,REALSTARTDATE," & _
                            "REALSTOPDATE,SEQNO,PK,RowID,FaciSNo,DeclarantName,FaciName FROM " & _
                             "(SELECT A.CITEMCODE,A.CITEMNAME,A.REALPERIOD,A.SHOULDAMT,A.ACCOUNTNO,A.CMNAME,A.REALSTARTDATE," & _
                                "A.REALSTOPDATE,A.SEQNO,A.BILLNO || A.ITEM PK,A.ROWID,A.FaciSNo,B.DeclarantName,B.FaciName FROM " & _
                             GetOwner & "SO033 A," & GetOwner & "SO004 B WHERE A.CUSTID = " & gCustId & _
                             " AND A.COMPCODE=" & gCompCode & _
                             " AND A.CANCELFLAG <> 1  AND A.CHEVEN <> 1 " & _
                             " AND A.CITEMCODE IN (" & strTmp & ") " & _
                             " AND A.CustId=B.CustId(+) And A.CompCode=B.CompCode(+)" & _
                             " And A.FaciSeqNo=B.SeqNo(+) "
                             
        '#5915 �W�[�L�o PAYOK=1 �P REFNO IN (3,7) By Kin 2011/03/08
        '#5999 HItemAutoUpd�N��}��,���X�w�������]���n�A�L�oPayOK=1 By Kin 2011/04/28
        strQry = strQry & " AND A.UCCODE NOT IN(SELECT CODENO FROM " & GetOwner & "CD013 " & _
                    " WHERE PAYOK = 1 OR REFNO IN (3,7)) "
        
        strQry = strQry & " And A.UCCode Is Not Null " & strOrder & ")" & _
               IIf(blnUCCodeIsNull, " Union " & strQry & " And " & _
            " A.UCCode Is Null  " & _
            " And A.INVSEQNO IS NULL " & strOrder & " )", "")

        SetMsQry csmCitem2, strQry, _
                             , , "�N�X,���O����,����,���B,���b�b��,���O�覡,�_�l��,�I���,�Ǹ�,PK,RowID,���~�Ǹ�,�ӽФH�m�W,���ئW��", _
                             "500,1800,460,560,1800,850,850,850,770,1,1,850,1600,1600", True, , , , , _
                             , , , , , , , , , , , , , , , , , , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD", , , , , 10

    Else
        strQry = "SELECT DISTINCT CITEMCODE,CITEMNAME,REALPERIOD,SHOULDAMT,ACCOUNTNO,CMNAME,REALSTARTDATE," & _
                            "REALSTOPDATE,SEQNO,PK,RowID,FaciSNo,DeclarantName,FaciName FROM " & _
                             "(SELECT A.CITEMCODE,A.CITEMNAME,A.REALPERIOD,A.SHOULDAMT,A.ACCOUNTNO,A.CMNAME,A.REALSTARTDATE," & _
                                "A.REALSTOPDATE,A.SEQNO,A.BILLNO || A.ITEM PK,A.ROWID,A.FaciSNo,B.DeclarantName,B.FaciName FROM " & _
                             GetOwner & "SO033 A," & GetOwner & "SO004 B " & _
                             " WHERE A.CUSTID = " & gCustId & _
                             " AND A.COMPCODE=" & gCompCode & _
                             " AND A.CANCELFLAG <> 1  AND A.CHEVEN <> 1 " & _
                             " AND A.CustId=B.CustId(+) And A.CompCode=B.CompCode(+)" & _
                             " And A.FaciSeqNo=B.SeqNo(+) "
                             
        '#5915 �W�[�L�o PAYOK=1 �P REFNO IN (3,7) By Kin 2011/03/08
        '#5999 HItemAutoUpd�N��}��,���X�w�������]���n�A�L�oPayOK=1 By Kin 2011/04/28
        Dim strQry2 As String
        strQry2 = strQry & " AND A.UCCODE NOT IN(SELECT CODENO FROM " & GetOwner & "CD013 " & _
                    " WHERE PAYOK = 1 OR REFNO IN (3,7)) "
        
                                             
        strQry = strQry2 & " And A.UCCode Is Not Null " & strOrder & ")" & _
                       IIf(blnUCCodeIsNull, " Union " & strQry & " And " & _
                    " A.UCCode Is Null  " & _
                    " And A.INVSEQNO IS NULL " & strOrder & " )", "")
        
    
        SetMsQry csmCitem2, strQry, _
                             , , "�N�X,���O����,����,���B,���b�b��,���O�覡,�_�l��,�I���,�Ǹ�,PK,RowID,���~�Ǹ�,�ӽФH�m�W,���ئW��", _
                             "500,1800,460,560,1800,850,850,850,770,1,1,850,1600,1600", True, , , , , _
                             , , , , , , , , , , , , , , , , , , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD", , , , , 10
                
    End If
  Exit Sub
ChkErr:
    ErrSub Name, "csmCitem2_ButtonClick"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    ReleaseCOM Me
End Sub

Private Sub RcdToScr()
  On Error GoTo ChkErr
    strACHID = Empty
    strACHDesc = Empty
    With rsSO106
         gdtAcceptTime.SetValue .Fields("AcceptTime") & ""
         '#6215 �ϥ� Query_CodeNo�|�����T�令 Query_Description By Kin 2012/02/22
         gilAcceptName.SetCodeNo .Fields("AcceptEn") & ""
         gilAcceptName.Query_Description
'         gilAcceptName.SetDescription .Fields("AcceptName") & ""
'         gilAcceptName.Query_CodeNo
         'txtProposer.Text = .Fields("Proposer") & ""
         '#3454 �ӽФH���ComboBox By Kin 2007/12/18
         cboProposer.Text = .Fields("Proposer") & ""
         If GetUserPriv("SO1100G", "(SO1100G0)") Then
             txtID.Text = .Fields("ID") & ""
             txtAccOwnerID = .Fields("AccountNameID") & ""
         Else
             txtID.Tag = .Fields("ID") & ""
             If txtID.Tag = Empty Then
                 txtID.Text = ""
             Else
                 txtID.Text = Left(txtID.Tag, 3) & "XXXX" & Right(txtID.Tag, 3)
             End If
             txtAccOwnerID.Tag = .Fields("AccountNameID") & ""
             If txtAccOwnerID.Tag = Empty Then
                 txtAccOwnerID.Text = ""
             Else
                 txtAccOwnerID.Text = Left(txtAccOwnerID.Tag, 3) & "XXXX" & Right(txtAccOwnerID.Tag, 3)
             End If
         End If
         
         gilBankCode.SetCodeNo .Fields("BankCode") & ""
         gilBankCode.SetDescription .Fields("BankName") & ""
         gilCardName.SetDescription .Fields("CardName") & ""
         gilCardName.Query_CodeNo
         mskCardExpDate.Text = .Fields("StopYM") & ""
         If Len(.Fields("StopYM")) > 1 Then
            mskCardExpDate.Text = Right("0" & (.Fields("StopYM") & ""), 6)
            txtHide.Text = mskCardExpDate.Text
         Else
            mskCardExpDate.Text = ""
            txtHide.Text = ""
         End If
         
         If GetUserPriv("SO1100G", "(SO1100G9)") Then
            txtAccountNo.Text = .Fields("AccountID") & ""
         Else
            txtAccountNo.Tag = .Fields("AccountID") & ""
            If txtAccountNo.Tag = Empty Then
                txtAccountNo.Text = ""
            Else
                txtAccountNo.Text = Left(txtAccountNo.Tag, 5) & "XXX" & Mid(txtAccountNo.Tag, 9, 2) & "XX" & Right(txtAccountNo.Tag, 4)
            End If
         End If
          If (Len(.Fields("BankCode") & "") > 0) Then
                  If GetRsValue("Select count(*) from cd018 Where CD018.PRGNAME like '%POST%' and CodeNo = " & .Fields("BankCode"), gcnGi) = 0 Then
                      AchTypeCondition = " In ( 1 ) "
                  Else
                      AchTypeCondition = " In ( 2 ) "
                  End If
         End If
         If Not blnStartPost Then
            AchTypeCondition = " In ( 1 ) "
         End If
            
        '#3236 ACHTNO�N�X�S���ߤ@,��ܥX�ӳ��|�ܦ��Ĥ@�ӥN�X,SO106�s�WACHTDESC���,SetQueryCode���ACHDESC��� By Kin 2007/05/28
         If (Not IsNull(.Fields("ACHTNO"))) And (Not IsNull(.Fields("ACHTDESC"))) Then
             'csmACH.SetQueryCode .Fields("ACHTNO") & ""
             '********************************************************************************************************************************************************
             '#3728 �uShow�X���D�諸ACH�N�X,�p�G���ŭȥ���Show�X By Kin 2008/03/03
             '#3946 �e�������ŭȪ���ACH�N�X�٬O�n���� By Kin 2008/05/23
             '#4030 ���L�o����
             '#4106 �W�[�P�_ACHType�Ѽ� By Kin 2008/09/22
             '#7049 �W�[BillHeadFmt���,�i�H���CD068A By Kin 2015/07/08
           
            If Not IsNull(.Fields("SendDate")) Then
               
                SetMsQry csmACH, "SELECT ACHTNO,ACHTDESC,CITEMCODESTR,BillHeadFmt FROM " & GetOwner & "CD068  " & _
                                            " Where ACHTDESC in(" & .Fields("ACHTDESC") & ") And ACHTDESC is NOT NULL And ACHType " & AchTypeCondition, _
                                       "ACH����N�X", "ACH�������", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
            Else
                SetMsQry csmACH, "SELECT ACHTNO,ACHTDESC,CITEMCODESTR,BillHeadFmt FROM " & GetOwner & "CD068  " & _
                                            " Where  ACHTNO Is Not Null And ACHTDESC is NOT NULL And ACHType  " & AchTypeCondition, _
                                    "ACH����N�X", "ACH�������", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
            End If
            
            csmACH.SetQueryCode .Fields("ACHTDESC") & ""
            
            
         Else
             SetMsQry csmACH, "SELECT ACHTNO,ACHTDESC,CITEMCODESTR,BillHeadFmt FROM " & GetOwner & "CD068  " & _
                                        " Where  ACHTNO is Not Null And ACHTDESC is Not NULL And ACHType " & AchTypeCondition, _
                                    "ACH����N�X", "ACH�������", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
             csmACH.Clear
         End If
             '*********************************************************************************************************************************************************
         txtCVC2 = .Fields("CVC2") & ""
         
         If Not IsNull(.Fields("CITEMSTR")) Then
            csmCitem1_ButtonClick
            '#3236 �N csmCitem1.DispStr�M��,����L��Ʈɷ|���d�W�@�����M�����e�� By Kin 2007/05/28
             csmCitem1.SetDispStr ""
             csmCitem1.SetQueryCode .Fields("CITEMSTR") & ""
             csmCitem1.SetQueryCode .Fields("CITEMSTR") & ""
         Else
             csmCitem1.Clear
         End If
         If Not IsNull(.Fields("CITEMSTR2")) Then
            csmCitem2_ButtonClick
            csmCitem2.SetDispStr ""
            csmCitem2.SetQueryCode .Fields("CITEMSTR2") & ""
            csmCitem2.SetQueryCode .Fields("CITEMSTR2") & ""
         Else
             csmCitem2.Clear
         End If
         
         txtAccountOwner = .Fields("AccountName") & ""
         gdaContiDate.Text = .Fields("PropDate") & ""
         gdaSendDate.Text = .Fields("SendDate") & ""
         
         gdaAuthStopDate.Text = .Fields("AuthorizeStopDate") & ""
         
         gdaSnactionDate.Text = .Fields("SnactionDate") & ""
         chkStop.Value = Val(.Fields("StopFlag") & "")
         gdaStopDate.Text = .Fields("StopDate") & ""
         gilMediaCode.SetCodeNo .Fields("MediaCode") & ""
         gilMediaCode.SetDescription .Fields("MediaName") & ""
         txtIntroId.Text = .Fields("IntroID") & ""
         lblIntroName.Caption = .Fields("IntroName") & ""
         gilCMCode.SetCodeNo .Fields("CMcode") & ""
         gilCMCode.SetDescription .Fields("CMname") & ""
         gilPTCode.SetCodeNo .Fields("PTcode") & ""
         gilPTCode.SetDescription .Fields("PTname") & ""
         lblUpdTime.Caption = .Fields("Updtime").Value & ""
         lblUpdEn.Caption = .Fields("UpdEn").Value & ""
         txtNote.Text = .Fields("Note") & ""
         chkFore.Value = IIf(IsNull(.Fields("Alien")), 0, .Fields("Alien"))
         chkFore2.Value = IIf(IsNull(.Fields("AccountAlien")), 0, .Fields("AccountAlien"))
         chkAddAcc.Value = IIf(IsNull(.Fields("AddCitemAccount")), 0, .Fields("AddCitemAccount"))
         ' ACH ���� 93/03/02 Hammer
         Select Case .Fields("AuthorizeStatus") & ""
                Case 1
                      lblStatus = "�w���v���\"
                Case 2
                      lblStatus = "�w�������v"
                Case Else
                      lblStatus = ""
         End Select
         txtACHcustID = .Fields("ACHCustID") & ""
         chkDeAut.Value = IIf(IsNull(.Fields("DeAuthorize")), 0, .Fields("DeAuthorize"))
         
         chkToACH.Value = IIf(IsNull(.Fields("ToACH")), 0, .Fields("ToACH"))
         txtAchSN = .Fields("ACHSN") & ""
         strInvSeqNo = .Fields("InvSeqNo") & ""
         lblInvSeqNo = strInvSeqNo
         
    End With
    '�Ӹ�k
    If Len(strAccountName) > 0 Then
        If strAccountName <> rsSO106.Fields("Proposer") Then
            cboProposer.Text = "XXXXX"
            gilCMCode.SetCodeNo "X"
            gilCMCode.SetDescription "XXX"
            gilPTCode.SetCodeNo "X"
            gilPTCode.SetDescription "XXX"
            gilBankCode.SetCodeNo "X"
            gilBankCode.SetDescription "XXX"
            txtAccountNo.PasswordChar = "X"
            txtAccountOwner.PasswordChar = "X"
            txtHide.PasswordChar = "X"
            
        End If
    End If

  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScr")
End Sub

Public Sub DeleteGo()
  On Error GoTo ChkErr
    '#3728 �R��SO106�ɤ@�֧R��SO106A��������� By Kin 2008/03/10
    
   
    With frmSO1100BMDI
        If Screen.ActiveForm.Name = "frmSO1100G" Then
            If MsgBox("�T�w�R�� [�b��/�d��] �� " & rsSO106!ACCOUNTID & " ���H�Υd��b��ƶ� ?", vbQuestion + vbYesNo, "�߰�") = vbYes Then
                '#5946 �W�[�O��Log�� By Kin 2011/03/16
                EditMode = giEditModeDelete
                Call InsSO106Log
                EditMode = giEditModeView
                StopGo CStr(rsSO106!ACCOUNTID)
                gcnGi.Execute "DELETE FROM " & GetOwner & "SO106 WHERE ROWID='" & .rsSO106!RowId & "'"
                '#3946 ���MasterId��PK By Kin 2008/05/30
                gcnGi.Execute "DELETE FROM " & GetOwner & "SO106A WHERE MasterId='" & .rsSO106("MasterId") & "'"
                OpenData
               .RefreshFormData
            End If
            If rsSO106.RecordCount = 0 Then NewRcd: UserPermissionGo
        Else
            If MsgBox("�T�w�R�� [�b��/�d��] �� " & .rsSO106!ACCOUNTID & " ���H�Υd��b��ƶ� ?", vbQuestion + vbYesNo, "�߰�") = vbYes Then
                '#5946 �W�[�O��Log�� By Kin 2011/03/16
                EditMode = giEditModeDelete
                Call InsSO106Log
                EditMode = giEditModeView
                StopGo CStr(.rsSO106!ACCOUNTID)
                gcnGi.Execute "DELETE FROM " & GetOwner & "SO106 WHERE ROWID='" & .rsSO106!RowId & "'"
                '#3946 ���MasterId��PK By Kin 2008/05/30
                gcnGi.Execute "DELETE FROM " & GetOwner & "SO106A WHERE MasterId='" & .rsSO106("MasterId") & "'"
               .RefreshFormData
            End If
            If .rsSO106.RecordCount = 0 Then UserPermissionGo
            Unload Me
        End If
    End With
    
  Exit Sub
ChkErr:
    ErrSub Me.Name, "DeleteGo"
End Sub

Public Function SmartAdd() As Boolean
  On Error GoTo ChkErr
  Exit Function
ChkErr:
    ErrSub Name, "SmartAdd"
End Function
'#3585 ���դ�OK(RA�W�歫�s�׭q��) ��Ach�̤j�� By Kin 2007/10/30
'#3861 ���SO041.ACHHeadCode����,���F�קKUser�S�]�w,�ҥH�쥻���Ȫ�����������,�p�GSO041�S�]�w����,�@�˨�ACH���̤j�� By Kin 2008/04/24
'#4006 �s�X��h�n���̤j�� By Kin 2008/07/21
Private Function GetMaxAchNo(ByVal strMax As String) As String
  On Error GoTo ChkErr
    Dim arrVar As Variant
    Dim strAchMax As String
    Dim i As Long
    Dim rsACHCode As New ADODB.Recordset
    If blnACHCustID Then
        If Not GetRS(rsACHCode, "Select ACHHeadCode From " & GetOwner & "SO041 Where SysId=" & gCompCode, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        If (Not IsNull(rsACHCode("ACHHeadCode"))) And (rsACHCode("ACHHeadCode") & "" <> "") Then
            GetMaxAchNo = rsACHCode("ACHHeadCode") & ""
        Else
            If Len(strMax) = 0 Then
                GetMaxAchNo = Empty
            Else
                arrVar = Split(Replace(strMax, "'", ""), ",")
                strAchMax = arrVar(0)
                For i = 1 To UBound(arrVar)
                    If Val(arrVar(i)) > Val(strAchMax) Then strAchMax = arrVar(i)
                Next i
                GetMaxAchNo = strAchMax
            End If
        End If
    Else
        If Len(strMax) = 0 Then
             GetMaxAchNo = Empty
        Else
            arrVar = Split(Replace(strMax, "'", ""), ",")
            strAchMax = arrVar(0)
            For i = 1 To UBound(arrVar)
                If Val(arrVar(i)) > Val(strAchMax) Then strAchMax = arrVar(i)
            Next i
            GetMaxAchNo = strAchMax
             
        End If
    End If
    On Error Resume Next
    Call CloseRecordset(rsACHCode)
  Exit Function
ChkErr:
    ErrSub Name, "GetMaxAchNo"
End Function
Private Function ScrToRcd() As Boolean
  On Error GoTo ChkErr
    Dim blnUpd As Boolean
    Dim strSqlCid As String, strAchCid As String
    Dim rsAchCid As New ADODB.Recordset
    Dim arySO106A() As String
    blnCallSO106A = False
    strACHID = Empty
    strACHDesc = Empty
    strFirstACHDesc = Empty
    strFirstACHID = Empty
    Dim i As Long
    Dim rsSO106A As New ADODB.Recordset
    blnUpd = False
    
    '*******************************************************************
    '#4130 �b���W�[�ˮ�,��InvSeqNo=0�ɥX�{ĵ�i�T�� By Kin 2008/10/21
    If Val(strInvSeqNo) <= 0 Then
        MsgBox "�o���y���������T�I", vbInformation, "ĵ�i"
        Exit Function
    End If
    '*******************************************************************
    With rsSO106
         If EditMode = giEditModeInsert Then
            .AddNew
            .Fields("InheritFlag") = 0
            .Fields("InheritKey") = ""
            
            '#3436 �x�sInvSeqNo��T By Kin 2007/12/13
            '.Fields("InvSeqNo") = strInvSeqNo
         Else
            
            If chkStop.Value = 1 Then
                .Fields("InheritFlag") = 0
                .Fields("InheritKey") = ""
                
                '********************************************************************
                '#3728 �p�G����,�h�N���ͨ������v��� By Kin 2008/03/18
                If blnACHCustID Then
                    If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018" & _
                                    " WHERE CODENO=" & gilBankCode.GetCodeNo & _
                                    " AND ( PRGNAME LIKE 'ACH%' Or " & posAchContition & " ) " & _
                                    " AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
                        blnNoShowMsg = True
                        csmACH.Clear
                        csmCitem1.Clear
                        
                    End If
                End If
                '*********************************************************************
            Else
                If .Fields("InheritFlag") = 0 And .Fields("InheritKey") <> "" Then blnUpd = True
                
                '#3436 �x�sInvSeqNo��T By Kin 2007/12/13
'                If strInvSeqNo <> "" Then
'                    .Fields("InvSeqNo") = strInvSeqNo
'                End If
            End If
         End If
         '*****************************************************
         '#3946 �W�[�����,�n�PSO106A���� By Kin 2008/05/30
         If Len(.Fields("MasterId") & "") = 0 Then
            .Fields("MasterId") = Get_MasterID_Seq
         End If
         '******************************************************
        .Fields("InvSeqNo") = strInvSeqNo
        .Fields("COMPCODE") = gCompCode & ""
        .Fields("CustID") = gCustId
        .Fields("AcceptName") = gilAcceptName.GetDescription2
        .Fields("AcceptEn") = gilAcceptName.GetCodeNo2
        .Fields("AcceptTime") = SolveNull(gdtAcceptTime.Text)
        '.Fields("Proposer") = txtProposer.Text
        '#3454 �ӽФH���ComboBox By Kin 2007/12/18
        .Fields("Proposer") = cboProposer.Text
        .Fields("BankCode") = gilBankCode.GetCodeNo2
        .Fields("BankName") = gilBankCode.GetDescription2
        .Fields("CardCode") = gilCardName.GetCodeNo2
        .Fields("CardName") = gilCardName.GetDescription2
        .Fields("StopYM") = SolveNull(mskCardExpDate.Text & "")
         
         If GetUserPriv("SO1100G", "(SO1100G0)") Then
            .Fields("ID") = UCase(txtID.Text)
            .Fields("AccountNameID") = UCase(txtAccOwnerID.Text)
         Else
            If txtID.Tag = Empty And txtID <> Empty Then txtID.Tag = txtID
            .Fields("ID") = UCase(txtID.Tag)
             If txtAccOwnerID.Tag = Empty And txtAccOwnerID <> Empty Then txtAccOwnerID.Tag = txtAccOwnerID
            .Fields("AccountNameID") = UCase(txtAccOwnerID.Tag)
         End If
         
'         If GetUserPriv("SO1100G", "(SO1100G9)") Then
'            .Fields("AccountID") = SolveNull(txtAccountNo & "")
'         Else
'             If txtAccountNo.Tag = Empty And txtAccountNo <> Empty Then
'                txtAccountNo.Tag = txtAccountNo
'            End If
'            .Fields("AccountID") = SolveNull(txtAccountNo.Tag)
'         End If
        Dim aWhereIn As String
        Dim aMaxACHNo As String
        Dim aACHTotal As String
        Dim isAch As Boolean
        
        If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018" & _
                            " WHERE CODENO=" & gilBankCode.GetCodeNo & _
                            " AND ( PRGNAME LIKE 'ACH%' Or " & posAchContition & " )" & _
                            " AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
                                       
            isAch = True
        End If
                        
       If isAch Then
            aMaxACHNo = GetMaxAchNo(csmACH.GetQryFld(1) & "")
            
            For i = 1 To 99
                If aWhereIn = "" Then
                    aWhereIn = aMaxACHNo & GetString(txtAccountNo & "", 6, giRight, True) & _
                            Right("00" & i, 2)
                    aWhereIn = "'" & aWhereIn & "'"
                Else
                    aWhereIn = aWhereIn & ",'" & aMaxACHNo & GetString(txtAccountNo & "", 6, giRight, True) & _
                            Right("00" & i, 2) & "'"
                End If
                 
            Next
        End If
        
        If EditMode = giEditModeEdit Then
             If IsNumeric(txtAccountNo) Then
                .Fields("AccountID") = SolveNull(txtAccountNo & "")
             Else
                .Fields("AccountID") = strRealAccNo
             End If
             '#3585 ���դ�OK,�s�説�A�]�n�ק�(RA��󭫷s����)
             If .Fields("AchCustId") & "" <> "" And gdaSendDate.GetValue = Empty Then
                '#5354 �ק�ACH��H���A�Y�ӵ���ƨS���e�����A�B�ק�e�B�ק��A��b���ۦP�A�N���n����ACH�Τ��ѧO�X By Kin 2012/09/03
                If .Fields("ACCOUNTID").OriginalValue <> txtAccountNo.Text Then
                    If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018" & _
                                        " WHERE CODENO=" & gilBankCode.GetCodeNo & _
                                        " AND (PRGNAME LIKE 'ACH%' Or " & posAchContition & " ) " & _
                                        " AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
                        '�p��ɻݦ����ۤv����
                        '#5354 �p���,�u�n���b����6�X�N�n SUBSTR(LPAD(ACCOUNTID,30,'0'),25,6) By Kin 2012/09/03
                        '#5354 ACHCUSTID�᭱2�X�q01-99��X�S�Ψ쪺 By Kin 2012/09/05
'                        strSqlCid = "SELECT COUNT(*) AS COUNTS FROM " & GetOwner & " SO106 " & _
'                                        " Where ACCOUNTID='" & txtAccountNo.Text & "'" & _
'                                        " AND SUBSTR(LPAD(ACCOUNTID,30,'0'),25,6)='" & Right(txtAccountNo.Text, 6) & "'" & _
'                                        " AND RowID<>'" & .Fields("RowID") & "" & "'"
                        strSqlCid = "SELECT ACHCUSTID FROM " & GetOwner & " SO106 " & _
                                        " Where SUBSTR(LPAD(ACCOUNTID,30,'0'),25,6)='" & Right(txtAccountNo.Text, 6) & "'" & _
                                        " AND RowID<>'" & .Fields("RowID") & "" & "'" & _
                                        " AND ACHCUSTID IN (" & aWhereIn & ") ORDER BY ACHCUSTID "
                        
                        If Not GetRS(rsAchCid, strSqlCid, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
                        aACHTotal = GetAchNoTotal(rsAchCid)
                        
                        If csmACH.GetQryFld(1) & "" <> "" And isAch Then
                            If blnACHCustID Then
'                                strAchCid = GetMaxAchNo(csmACH.GetQryFld(1) & "") & _
'                                            GetString(txtAccountNo & "", 6, giRight, True) & _
'                                                Right("00" & CStr(Trim(rsAchCid("Counts")) + 1), 2)
                                strAchCid = aMaxACHNo & _
                                            GetString(txtAccountNo & "", 6, giRight, True) & _
                                                Right("00" & aACHTotal, 2)
                                .Fields("ACHCustid") = strAchCid
                            Else
                                '#4006 �®榡�s�X��h�]�n�P�s�榡�@�� By Kin 2008/07/21
                                strAchCid = GetMaxAchNo(csmACH.GetQryFld(1) & "") & Format(gCustId & "", "00000000")
                                .Fields("ACHCustid") = strAchCid
                            End If
                        End If
    
                    End If
                End If
             End If
        ElseIf EditMode = giEditModeInsert Then
            .Fields("AccountID") = SolveNull(txtAccountNo & "")
            
            '****************************************************************************************************************************************************************************************
            '#3585 ACHCustID=1�ɡA�ݭn��JACH�Τḹ�X By Kin 2007/10/26
            '#3585 ���դ�OK,���ӭn�P�_�ַǤ��(gdaSnactionDate),��RA�S�ק令�n�P�_�e����(gdaSendDate) By Kin 2007/10/30
            '#3585 ���դ�OK(RA��󭫷s����) SO041.AchCustId=0�ɤ]�n�۰ʲ��� By Kin 2007/10/30
            If gdaSendDate.GetValue = Empty Then
                If blnACHCustID Then
                    If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018" & _
                                    " WHERE CODENO=" & gilBankCode.GetCodeNo & _
                                    " AND ( PRGNAME LIKE 'ACH%' Or " & posAchContition & " ) " & _
                                    " AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
                        '#5354 �p���,�u�n���b����6�X�N�n SUBSTR(LPAD(ACCOUNTID,30,'0'),25,6) By Kin 2012/09/03
                        '#5354 ACHCUSTID�᭱2�X�q01-99��X�S�Ψ쪺 By Kin 2012/09/05
'                        strSqlCid = "SELECT COUNT(*) AS COUNTS FROM " & GetOwner & " SO106 " & _
'                                                                    " Where SUBSTR(LPAD(ACCOUNTID,30,'0'),25,6)='" & Right(txtAccountNo.Text, 6) & "'" & _
'                                                                    " AND SUBSTR(ACHCUSTID,4,6) ='" & Right(txtAccountNo & "", 6) & "'"
                        strSqlCid = "SELECT ACHCUSTID FROM " & GetOwner & " SO106 " & _
                                                                    " Where SUBSTR(LPAD(ACCOUNTID,30,'0'),25,6)='" & Right(txtAccountNo.Text, 6) & "'" & _
                                                                    " AND SUBSTR(ACHCUSTID,4,6) ='" & Right(txtAccountNo & "", 6) & "'" & _
                                                                    " AND ACHCUSTID IN (" & aWhereIn & ")"
                        If Not GetRS(rsAchCid, strSqlCid, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
                        aACHTotal = GetAchNoTotal(rsAchCid)
    '                    strAchCid = Replace(csmACH.GetQryFld(1), "'", "") & _
    '                                GetString(txtAccountNo & "", 6, giRight, True) & _
    '                                Right("00" & CStr(Trim(rsAchCid("Counts")) + 1), 2)
                        If csmACH.GetQryFld(1) & "" <> "" Then
                            strAchCid = aMaxACHNo & _
                                        GetString(txtAccountNo & "", 6, giRight, True) & _
                                        Right("00" & aACHTotal, 2)
                            .Fields("ACHCustid") = strAchCid
                        End If
                    End If
                Else
                    If csmACH.GetQryFld(1) & "" <> "" And isAch Then
                        strAchCid = GetMaxAchNo(csmACH.GetQryFld(1) & "") & Format(gCustId & "", "00000000")
                        .Fields("ACHCustid") = strAchCid
                    End If
                End If
            End If
            '**********************************************************************************************************************************************************************************************
        End If
        
'        .Fields("CITEMSTR") = gimCitem.GetQueryCode
        .Fields("CITEMSTR") = IIf(csmCitem1.GetQueryCode <> Empty, csmCitem1.GetQueryCode(True), Null)
        
        '*************************************************************************************************
        '#3791 �]��CitemStr2�|�Q�s�����'',�ҥH�|�ɭP���SO033����ƿ��~ By Kin 2008/04/10
        .Fields("CITEMSTR2") = IIf(csmCitem2.GetQueryCode <> Empty, csmCitem2.GetQueryCode(True), Null)
        '*************************************************************************************************
        
        '.Fields("ACHTNO") = csmACH.GetQueryCode
        If csmACH.GetDispStr & "" <> "" Then
            '#3236 �]��SetQueryCode��k����,�x�s�ɧ��GetQryFld��k By Kin 2007/05/28
            .Fields("ACHTNO") = IIf(csmACH.GetQueryCode <> Empty, csmACH.GetQryFld(1), Null)
            '#3236 �h�W�[�x�s"ACHTDESC"��� By Kin 2007/05/28
            .Fields("ACHTDESC") = IIf(csmACH.GetQueryCode <> Empty, csmACH.GetQryFld(2), Null)
        Else
            .Fields("ACHTNO") = Null
            .Fields("ACHTDESC") = Null
        End If
        .Fields("CVC2") = SolveNull(txtCVC2 & "")
        .Fields("AccountName") = txtAccountOwner.Text & ""
        .Fields("PropDate") = SolveNull(gdaContiDate.Text)
        .Fields("SendDate") = SolveNull(gdaSendDate.Text)
        
        .Fields("AuthorizeStopDate") = SolveNull(gdaAuthStopDate.Text)
        
        .Fields("SnactionDate") = SolveNull(gdaSnactionDate.Text)
        .Fields("StopFlag") = chkStop.Value
        .Fields("StopDate") = SolveNull(gdaStopDate.Text)
        .Fields("MediaCode") = gilMediaCode.GetCodeNo2
        .Fields("MediaName") = gilMediaCode.GetDescription2
        .Fields("IntroID") = SolveNull(txtIntroId.Text & "")
        .Fields("IntroName") = SolveNull(lblIntroName.Caption & "")
        .Fields("Updtime") = GetDTString(RightNow)
        .Fields("UpdEn") = garyGi(1)
        .Fields("Note") = SolveNull(txtNote.Text & "")
        .Fields("CMcode") = gilCMCode.GetCodeNo2
        .Fields("CMname") = gilCMCode.GetDescription2
        .Fields("PTcode") = gilPTCode.GetCodeNo2
        .Fields("PTname") = gilPTCode.GetDescription2
        .Fields("Alien") = chkFore.Value
        .Fields("AccountAlien") = chkFore2.Value
        .Fields("AddCitemAccount") = chkAddAcc.Value
        'ACH ���� 93/03/02 Hammer
'         Select Case lblStatus
'                Case "�w���v���\"
'                    .Fields("Status") = 1
'                Case "�w�������v"
'                    .Fields("Status") = 2
'                Case Else
'                    .Fields("Status") = 0
'         End Select
'        .Fields("ACHCustID") = txtACHcustID
        .Fields("DeAuthorize") = chkDeAut.Value
        .Fields("ToACH") = chkToACH.Value
        .Fields("ACHSN") = txtAchSN
        '*********************************************************************************************************************************************************
        '#3728 �P�_ACH���Ǧ����� By Kin 2008/03/05
        If EditMode = giEditModeInsert Then
            
            
            Call GetCitemCode(IIf(csmACH.GetDispStr <> Empty, csmACH.GetQryFld(1), Empty), IIf(csmACH.GetDispStr <> Empty, csmACH.GetQryFld(2), Empty), _
                              IIf(csmCitem1.GetDispStr <> Empty, csmCitem1.GetQryFld(1), Empty), _
                              IIf(csmCitem1.GetDispStr <> Empty, csmCitem1.GetQryFld(2), Empty), True, arySO106A)
        Else
            Call GetCitemCode(IIf(csmACH.GetDispStr <> Empty, csmACH.GetQryFld(1), Empty), _
                              IIf(csmACH.GetDispStr <> Empty, csmACH.GetQryFld(2), Empty), _
                              IIf(csmCitem1.GetDispStr <> Empty, csmCitem1.GetQryFld(1), Empty), _
                              IIf(csmCitem1.GetDispStr <> Empty, csmCitem1.GetQryFld(2), Empty), False, arySO106A)
        End If
        '*********************************************************************************************************************************************************
        '#5946 �W�[�O��Log��
        Call InsSO106Log
        .Update
        
        
        '************************************************************************************************************************************************
        '#3728 �p�GACH���i�沧�ʫh�}�l��� �÷s�W�Χ�sSO106A By Kin 2008/03/10
        Dim strUpdTime As String
        Dim strCreateTime As String
        Dim strStopDate As String
        strUpdTime = GetDTString(RightNow)
        strCreateTime = Format(RightNow, "yyyy/mm/dd hh:mm:ss")
        strStopDate = Format(RightNow, "yyyy/mm/dd hh:mm:ss")
        For i = LBound(arySO106A) To UBound(arySO106A)
        
            Dim strUpdSO106A As String
            strUpdSO106A = Empty
            Select Case arySO106A(i, 2)
                '���ͤ@�����v���
                Case "0"
                    '#3946 �s�W�ɱNAuthorizeStatus�쥻��0�令4,�N��ݴ��X
                    strUpdSO106A = "Insert into " & GetOwner & "SO106A " & _
                                "(MasterRowID,ACHTNO,CitemCodeStr,CitemNameStr," & _
                                "UpdEn,UpdTime,CreateTime,CreateEn,RecordType," & _
                                "AuthorizeStatus,ACHDesc,MasterId)" & _
                                " Values ('" & .Fields("RowId") & "'," & _
                                          "'" & arySO106A(i, 0) & "'," & _
                                          IIf(arySO106A(i, 3) = Empty, "Null,", "'" & arySO106A(i, 3) & "',") & _
                                          IIf(arySO106A(i, 4) = Empty, "Null,", "'" & arySO106A(i, 4) & "',") & _
                                          "'" & garyGi(1) & "'," & _
                                          "'" & strUpdTime & "'," & _
                                          "TO_DATE('" & strCreateTime & "','YYYY/MM/DD HH24:MI:SS')," & _
                                          "'" & garyGi(1) & "'," & _
                                          "0,4,'" & arySO106A(i, 1) & "'," & .Fields("MasterId") & ")"
                    gcnGi.Execute strUpdSO106A
                '���ͤ@���������v���,�p�G�O���v������ơA�h�����\���ͤ@�����
                '#3946 �쥻�O�HMasterRowId������,�{�b�אּ�HMasterId������ By Kin 2008/05/30
                Case "1"
                    '#3946 �NAuthorizeStatus���A�אּ4 By Kin 2008/05/30
                    strUpdSO106A = "Select * From " & GetOwner & "SO106A" & _
                                    " Where MasterId='" & .Fields("MasterId") & "'" & _
                                    " And ACHTNO='" & arySO106A(i, 0) & "'" & _
                                    " And ACHDesc='" & arySO106A(i, 1) & "'" & _
                                    " And " & IIf(arySO106A(i, 3) = "", " CitemCodeStr is Null", " CitemCodeStr='" & arySO106A(i, 3) & "'") & _
                                    " And " & IIf(arySO106A(i, 4) = "", "CitemNameStr is Null", " CitemNameStr='" & arySO106A(i, 4) & "'") & _
                                    " And AuthorizeStatus=4"
                    Set rsSO106A = gcnGi.Execute(strUpdSO106A)
                    '3946 �����ʪ���,����������,�����R�� By Kin 2008/05/23
                    If Not rsSO106A.EOF Then
                        If IsNull(rsSO106A("AuthorizeStatus")) Or Val(rsSO106A("AuthorizeStatus") & "") = 3 Or Val(rsSO106A("AuthorizeStatus") & "") = 4 Then
'                            strUpdSO106A = "Update " & GetOwner & "SO106A" & _
'                                         " Set StopFlag=1," & _
'                                         "StopDate=TO_DATE('" & strStopDate & "','YYYY/MM/DD HH24:MI:SS')," & _
'                                         "Notes=Notes" & IIf(IsNull(rsSO106A("AuthorizeStatus")), " || '���v�������ӵ����'", " ||'���v���Ѩ����ӵ�'") & _
'                                         " Where MasterRowId='" & .Fields("RowId") & "'" & _
'                                         " And ACHTNO='" & arySO106A(i, 0) & "'" & _
'                                         " And ACHDesc='" & arySO106A(i, 1) & "'" & _
'                                         " And " & IIf(arySO106A(i, 3) = "", "CitemCodeStr is Null", "CitemCodeStr='" & arySO106A(i, 3) & "'") & _
'                                         " And " & IIf(arySO106A(i, 4) = "", "CitemNameStr is Null", "CitemNameStr='" & arySO106A(i, 4) & "'") & _
'                                         " And AuthorizeStatus is Null"
                            '#3946 �NAuthorizeStatus���A�אּ4 By Kin 2008/05/30
                            strUpdSO106A = "Delete From " & GetOwner & "SO106A " & _
                                            " Where MasterId='" & .Fields("MasterId") & "'" & _
                                            " And ACHTNO='" & arySO106A(i, 0) & "'" & _
                                            " And ACHDesc='" & arySO106A(i, 1) & "'" & _
                                            " And " & IIf(arySO106A(i, 3) = "", "CitemCodeStr is Null", "CitemCodeStr='" & arySO106A(i, 3) & "'") & _
                                            " And " & IIf(arySO106A(i, 4) = "", "CitemNameStr is Null", "CitemNameStr='" & arySO106A(i, 4) & "'") & _
                                            " And  AuthorizeStatus=4"
                            blnCallSO106A = False
                        Else
                            strUpdSO106A = "Insert into " & GetOwner & "SO106A " & _
                                        "(MasterRowID, ACHTNO, CitemCodeStr, CitemNameStr," & _
                                        " UpdEn,UpdTime,CreateTime,CreateEn,RecordType,StopFlag," & _
                                          "StopDate,AuthorizeStatus,ACHDesc,MasterId) " & _
                                        " Values ('" & .Fields("RowId") & "'," & _
                                                  "'" & arySO106A(i, 0) & "'," & _
                                                  IIf(arySO106A(i, 3) = Empty, "Null,", "'" & arySO106A(i, 3) & "',") & _
                                                  IIf(arySO106A(i, 4) = Empty, "Null,", "'" & arySO106A(i, 4) & "',") & _
                                                  "'" & garyGi(1) & "'," & _
                                                  "'" & strUpdTime & "'," & _
                                                  "TO_DATE('" & strCreateTime & "','YYYY/MM/DD HH24:MI:SS')," & _
                                                  "'" & garyGi(1) & "'," & _
                                                  "1,1," & _
                                                  "TO_DATE('" & strStopDate & "','YYYY/MM/DD HH24:MI:SS')," & _
                                                  "Null,'" & arySO106A(i, 1) & "'," & .Fields("MasterId") & ")"
                            If strFirstACHID = Empty Then strFirstACHID = arySO106A(i, 0)
                            If strFirstACHDesc = Empty Then strFirstACHDesc = arySO106A(i, 1)
                            blnCallSO106A = True
                            strACHID = IIf(strACHID = Empty, "'" & arySO106A(i, 0) & "'", strACHID & ",'" & arySO106A(i, 0) & "'")
                            strACHDesc = IIf(strACHDesc = Empty, "'" & arySO106A(i, 1) & "'", strACHDesc & ",'" & arySO106A(i, 1) & "'")
                        End If
                    Else
                        '#3946 �NAuthorizeStatus���A�אּ4 By Kin 2008/05/30
                        strUpdSO106A = "Insert into " & GetOwner & "SO106A " & _
                                        "(MasterRowID, ACHTNO, CitemCodeStr, CitemNameStr," & _
                                        " UpdEn,UpdTime,CreateTime,CreateEn,RecordType,StopFlag," & _
                                          "StopDate,AuthorizeStatus,ACHDesc,MasterId) " & _
                                        " Values ('" & .Fields("RowId") & "'," & _
                                                  "'" & arySO106A(i, 0) & "'," & _
                                                  IIf(arySO106A(i, 3) = Empty, "Null,", "'" & arySO106A(i, 3) & "',") & _
                                                  IIf(arySO106A(i, 4) = Empty, "Null,", "'" & arySO106A(i, 4) & "',") & _
                                                  "'" & garyGi(1) & "'," & _
                                                  "'" & strUpdTime & "'," & _
                                                  "TO_DATE('" & strCreateTime & "','YYYY/MM/DD HH24:MI:SS')," & _
                                                  "'" & garyGi(1) & "'," & _
                                                  "1,1," & _
                                                  "TO_DATE('" & strStopDate & "','YYYY/MM/DD HH24:MI:SS')," & _
                                                  "4,'" & arySO106A(i, 1) & "'," & .Fields("MasterId") & ")"
                        If strFirstACHID = Empty Then strFirstACHID = arySO106A(i, 0)
                        If strFirstACHDesc = Empty Then strFirstACHDesc = arySO106A(i, 1)
                        blnCallSO106A = True
                        strACHID = IIf(strACHID = Empty, "'" & arySO106A(i, 0) & "'", strACHID & ",'" & arySO106A(i, 0) & "'")
                        strACHDesc = IIf(strACHDesc = Empty, "'" & arySO106A(i, 1) & "'", strACHDesc & ",'" & arySO106A(i, 1) & "'")
                    
                    End If
                    gcnGi.Execute strUpdSO106A
                    '�h�������ɡA�I�sSO1100G1�ΡA�P�_SO1100G1�n���d�b���@��
'                    If strFirstACHID = Empty Then strFirstACHID = arySO106A(i, 0)
'                    If strFirstACHDesc = Empty Then strFirstACHDesc = arySO106A(i, 1)
'                    blnCallSO106A = True
'                    strACHID = IIf(strACHID = Empty, "'" & arySO106A(i, 0) & "'", strACHID & ",'" & arySO106A(i, 0) & "'")
'                    strACHDesc = IIf(strACHDesc = Empty, "'" & arySO106A(i, 1) & "'", strACHDesc & ",'" & arySO106A(i, 1) & "'")
                '��sACH�̪����O����
                Case "2"
                    strUpdSO106A = "Select * From " & GetOwner & "SO106A Where " & _
                                 "ACHTNO='" & arySO106A(i, 0) & "' " & _
                                 "And MasterID='" & .Fields("MasterId") & "' " & _
                                 "And ACHDesc='" & arySO106A(i, 1) & "' " & _
                                 "And StopFlag<>1"
                    Set rsSO106A = gcnGi.Execute(strUpdSO106A)
                    If Not rsSO106A.EOF Then
                        If rsSO106A("CitemCodestr") & "" <> arySO106A(i, 3) Then
                            strUpdSO106A = "Update " & GetOwner & "SO106A Set " & _
                                                    "CitemCodeStr=" & IIf(arySO106A(i, 3) <> Empty, "'" & arySO106A(i, 3) & "',", "NULL,") & _
                                                    "CitemNameStr=" & IIf(arySO106A(i, 3) <> Empty, "'" & arySO106A(i, 4) & "',", "NULL,") & _
                                                    "UpdEn='" & garyGi(1) & "'," & _
                                                    "UpdTime='" & strUpdTime & "' " & _
                                           "Where MasterId='" & .Fields("MasterId") & "'" & _
                                           " And ACHTNO='" & arySO106A(i, 0) & "'" & _
                                           " And ACHDesc='" & arySO106A(i, 1) & "'" & _
                                           " And StopFlag<>1"
                            gcnGi.Execute (strUpdSO106A)
                        End If
                    '#3946 ���դ�OK,�p�G�u���@�����ɭ�,�n����Update By Kin 2008/07/25
                    Else
                        strUpdSO106A = "Update " & GetOwner & "SO106A Set " & _
                                                    "CitemCodeStr=" & IIf(arySO106A(i, 3) <> Empty, "'" & arySO106A(i, 3) & "',", "NULL,") & _
                                                    "CitemNameStr=" & IIf(arySO106A(i, 3) <> Empty, "'" & arySO106A(i, 4) & "',", "NULL,") & _
                                                    "ACHTNO=" & IIf(arySO106A(i, 0) <> Empty, "'" & arySO106A(i, 0) & "',", "NULL") & _
                                                    "ACHDesc=" & IIf(arySO106A(i, 1) <> Empty, "'" & arySO106A(i, 1) & "',", "NULL") & _
                                                    "UpdEn='" & garyGi(1) & "'," & _
                                                    "UpdTime='" & strUpdTime & "' " & _
                                           "Where MasterId='" & .Fields("MasterId") & "'"
                        gcnGi.Execute (strUpdSO106A)
                    End If

            End Select
        Next i
        '***************************************************************************************************************************************************************************************
         If blnUpd Then
            Dim strAllChild As String
            Dim varAry, varElement
            Dim lngAft As Long, lngChildID As Long
            strAllChild = GetChild
            varAry = Split(strAllChild, ",")
            For Each varElement In varAry
                lngChildID = Val(varElement & "")
                If lngChildID > 0 Then
                    .Fields("CustID") = lngChildID
                    .Fields("InheritFlag") = 1
                    .Fields("ACHSN") = GetRsValue("SELECT ACHSN FROM " & GetOwner & "SO106" & _
                                                                        " WHERE CUSTID=" & lngChildID & _
                                                                        " AND INHERITKEY='" & rsSO106("InheritKey") & "'" & _
                                                                        " AND INHERITFLAG=1") & ""
                    Debug.Print InsertToOracle("SO106", rsSO106, _
                                                                " WHERE CUSTID=" & lngChildID & _
                                                                " AND INHERITKEY='" & rsSO106("InheritKey") & "'" & _
                                                                " AND INHERITFLAG=1", lngAft, False)
                    .CancelUpdate
                End If
            Next
         End If
    End With
    
    ScrToRcd = True
    'csmACH.Clear
    On Error Resume Next
    Call CloseRecordset(rsSO106A)
  Exit Function
ChkErr:
    
    ScrToRcd = False
    gcnGi.RollbackTrans
    Call ErrSub(Me.Name, "ScrToRcd")
End Function
'���XACH�᭱2�X�S�Q�Ϊ���
Private Function GetAchNoTotal(ByRef rsAchCid As Recordset) As String
  On Error GoTo ChkErr
    Dim aACHTotal As String
    Dim aFind As Boolean
    Dim i As Integer
    aACHTotal = "01"
    aFind = False
    If rsAchCid.RecordCount > 0 Then
        For i = 1 To 99
            If aFind Then Exit For
            aACHTotal = Right("00" & i, 2)
            rsAchCid.MoveFirst
            Do While Not rsAchCid.EOF
                aFind = True
                If Right(rsAchCid("ACHCUSTID"), 2) = aACHTotal Then
                    aFind = False
                    Exit Do
                End If
                rsAchCid.MoveNext
            Loop
        Next
    End If
    GetAchNoTotal = aACHTotal
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "GetAchNoTotal")
End Function
'#5946 �W�[�O��Log��
Private Sub InsSO106Log()
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim rsLog As New ADODB.Recordset
    Dim blnChg As Boolean
    Dim i As Integer
    aSQL = "SELECT * FROM SO106LOG WHERE 1=0 "
    If Not GetRS(rsLog, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    
    Select Case lngEditMode
        Case giEditModeEdit
            '�P�_���S�����ʸ��
            For i = 0 To rsSO106.Fields.Count - 1
                If UCase(rsSO106.Fields(i).Name) <> "UPDTIME" Then
                    If (rsSO106.Fields(i).Value <> rsSO106.Fields(i).OriginalValue) Or _
                      (Len(rsSO106.Fields(i).Value & "") <> Len(rsSO106.Fields(i).OriginalValue & "")) Then
                        blnChg = True
                        Exit For
                    End If
                End If
            Next i
            '�����ʸ��
            If blnChg Then
                rsLog.AddNew
                rsLog("FuncType") = 1
                For i = 0 To rsSO106.Fields.Count - 1
                    If UCase(rsSO106.Fields(i).Name) <> "ROWID" Then
                        If Len(rsSO106.Fields(i).OriginalValue & "") = 0 Then
                            rsLog.Fields(rsSO106.Fields(i).Name) = Null
                        Else
                            rsLog.Fields(rsSO106.Fields(i).Name) = rsSO106.Fields(i).OriginalValue & ""
                        End If
                    End If
                Next i
                
                For i = 0 To rsSO106.Fields.Count - 1
                    If (rsSO106.Fields(i).Value <> rsSO106.Fields(i).OriginalValue) Or _
                      (Len(rsSO106.Fields(i).Value & "") <> Len(rsSO106.Fields(i).OriginalValue & "")) Then
                        Select Case UCase(rsSO106.Fields(i).Name)
                            Case "ROWID"
                            Case "UPDTIME"
                                rsLog("UPDTIME") = GetDTString(RightNow)
                            Case "UPDEN"
                                rsLog("UPDEN") = garyGi(1)
                            Case "COMPCODE"
                            Case "CUSTID"
                            Case "MASTERID"
                            Case "NEWUPDTIME"
                                rsLog("NEWUPDTIME") = rsSO106.Fields("NEWUPDTIME")
                            Case Else
                                If Len(rsSO106.Fields(i).Value & "") = 0 Then
                                    rsLog.Fields(rsSO106.Fields(i).Name & "B").Value = Null
                                Else
                                    rsLog.Fields(rsSO106.Fields(i).Name & "B").Value = rsSO106.Fields(i).Value
                                End If
                        End Select
                      
                    End If
                Next i
                rsLog.Update
            End If
        Case giEditModeInsert
            rsLog.AddNew
            rsLog("FuncType") = 3
            For i = 0 To rsSO106.Fields.Count - 1
                Select Case UCase(rsSO106.Fields(i).Name)
                    Case "ROWID"
                    Case "UPDTIME"
                        rsLog.Fields("UPDTIME") = GetDTString(RightNow)
                    Case "UPDEN"
                        rsLog.Fields("UPDEN").Value = garyGi(1)
                    Case "COMPCODE"
                    Case "CUSTID"
                    Case "MASTERID"
                    Case "NEWUPDTIME"
                                rsLog("NEWUPDTIME") = rsSO106.Fields("NEWUPDTIME")
                    Case Else
                        If Len(rsSO106.Fields(i).Value & "") = 0 Then
                            rsLog.Fields(rsSO106.Fields(i).Name & "B").Value = Null
                        Else
                            rsLog.Fields(rsSO106.Fields(i).Name & "B").Value = rsSO106.Fields(i).Value
                        End If
                End Select
            Next i
            rsLog.Update
        Case giEditModeDelete
            rsLog.AddNew
            rsLog("FuncType") = 2
            If UCase(Screen.ActiveForm.Name) = "FRMSO1100G" Then
                For i = 0 To rsSO106.Fields.Count - 1
                    If UCase(rsSO106.Fields(i).Name) <> "ROWID" Then
                        If Len(rsSO106.Fields(i).Value & "") = 0 Then
                            rsLog.Fields(rsSO106.Fields(i).Name) = Null
                        Else
                            rsLog.Fields(rsSO106.Fields(i).Name) = rsSO106.Fields(i).Value & ""
                        End If
                    End If
                    
                Next i
            Else
                With frmSO1100BMDI
                    For i = 0 To .rsSO106.Fields.Count - 1
                        If UCase(.rsSO106.Fields(i).Name) <> "ROWID" Then
                            If Len(.rsSO106.Fields(i).Value & "") = 0 Then
                                rsLog.Fields(.rsSO106.Fields(i).Name) = Null
                            Else
                                rsLog.Fields(.rsSO106.Fields(i).Name) = .rsSO106.Fields(i).Value & ""
                            End If
                        End If
                        
                    Next i
                End With
            End If
            rsLog.Fields("UPDTIME") = GetDTString(RightNow)
            rsLog.Fields("UPDEN") = garyGi(1)
            rsLog.Update
    End Select

    On Error Resume Next
    Call CloseRecordset(rsLog)
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "InsSO106Log")
  
End Sub
Public Sub FunctionKey(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Select Case KeyCode
           '----------------------------------------------------
           Case vbKeyF2  '   F3:�s��, �۷����Ucmdsave
           '#3337 �P�_�O�_�b�s���Ҧ���F2�A�p�G�O�s���Ҧ�����s�� By Kin 2007/07/23
                If lngEditMode = giEditModeView Then Exit Sub
                If Not ChkGiList(KeyCode) Then Exit Sub
                UpdateGo
           '----------------------------------------------------
           Case vbKeyEscape
                CancelGo
    End Select
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    IsDataOk = False
    blnInsSO138 = False
   
    
    If txtAchSN <> Empty Then
        If EditMode = giEditModeInsert Then
            If Val(RPxx(gcnGi.Execute("SELECT COUNT(*) FROM " & _
                                                        GetOwner & "SO106 WHERE ACHSN='" & txtAchSN & "'").GetString)) > 0 Then
                MsgBox "[�ӽЮѳ渹] ��ƭ��� ! �нT�{ !", vbInformation, "�T��"
                On Error Resume Next
                txtAchSN.SetFocus
                Exit Function
            End If
        ElseIf EditMode = giEditModeEdit Then
            Dim strRID As String
            On Error Resume Next
            strRID = RPxx(gcnGi.Execute("SELECT ROWID FROM " & GetOwner & "SO106 WHERE ACHSN='" & txtAchSN & "'").GetString & "")
            If Err.Number = 0 Then
                If strRID <> rsSO106("ROWID") Then
                    MsgBox "[�ӽЮѳ渹] ��ƭ��� ! �нT�{ !", vbInformation, "�T��"
                    On Error Resume Next
                    txtAchSN.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    If gdtAcceptTime.Text = "" Then gdtAcceptTime.SetFocus: GoTo 66
    If gilAcceptName.GetCodeNo = "" Then gilAcceptName.SetFocus: GoTo 66
    '#3454 �ӽФH���ComboBox By Kin 2007/12/18
    If cboProposer.Text = "" Then cboProposer.SetFocus: GoTo 66
'    If txtProposer = "" Then txtProposer.SetFocus: GoTo 66
'    If lblID.ForeColor = vbRed Then
'        If txtID = "" Then txtID.SetFocus: GoTo 66
'    End If
    If gdaContiDate.Text = "" Then gdaContiDate.SetFocus: GoTo 66
    If chkStop.Value = 1 Then
        If gdaStopDate.Text = "" Then gdaStopDate.SetFocus: GoTo 66
    End If
    If gilCMCode.GetCodeNo = "" Then gilCMCode.SetFocus: GoTo 66
    If gilPTCode.GetCodeNo = Empty Then gilPTCode.SetFocus: GoTo 66
    If gilBankCode.GetCodeNo = "" Then gilBankCode.SetFocus: GoTo 66
    If lblAccountOwner.ForeColor = vbRed Then
        If txtAccountOwner.Text = Empty Then txtAccountOwner.SetFocus: GoTo 66
    End If
    If lblAccNameID.ForeColor = vbRed Then
        If txtAccOwnerID.Text = Empty Then txtAccOwnerID.SetFocus: GoTo 66
    End If
    '#8724
    If intCD031RefNo = 4 Or intCD031RefNo = 5 Then
        If gilCardName.GetCodeNo = Empty Then gilCardName.SetFocus: GoTo 66
        If mskCardExpDate.Text = Empty Then mskCardExpDate.SetFocus: GoTo 66
    End If
    
    If lblAccountNo.ForeColor = vbRed Then
        If Trim(txtAccountNo) = Empty Then txtAccountNo.SetFocus: GoTo 66
    End If
    '*************************************************************************************************************************
    '#3395 ���O�覡���H�Υd��(CD031.RefNo=4)�A�ϥΫH�Υd�ˮ֤覡�A���O�覡���Ȧ�ɳW�h����(CD031.RefNo=2) By Kin 2007/08/20
    '#8724
    If intCD031RefNo = 2 Or intCD031RefNo = 4 Or intCD031RefNo = 5 Then
        If intCD031RefNo = 2 Then
            If ServiceCommon.AccountNo_Validate(Me) Then
                txtAccountNo.SetFocus
                Exit Function
            End If
        Else
            If Not ChkCreditCard(gilCardName.GetCodeNo, txtAccountNo.Text) Then
                txtAccountNo.SetFocus
                Exit Function
            End If
        End If
    End If
    '**************************************************************************************************************************
    If txtID <> Empty And chkFore.Value = 0 Then
        If GetUserPriv("SO1100G", "(SO1100G0)") Then
            IDIsOk txtID.Text, True
        Else
            IDIsOk txtID.Tag, True
        End If
'        If Not IDIsOk(txtID, True) Then
'    '       On Error Resume Next
'    '       txtID.SetFocus
'    '       Exit Function
'        End If
    End If
    If txtAccOwnerID <> Empty And chkFore2.Value = 0 Then
        If GetUserPriv("SO1100G", "(SO1100G0)") Then
            IDIsOk txtAccOwnerID.Text, True
        Else
            IDIsOk txtAccOwnerID.Tag, True
        End If
'        If Not IDIsOk(txtAccOwnerID, True) Then
'    '       On Error Resume Next
'    '       txtID.SetFocus
'    '       Exit Function
'        End If
    End If
    
    If csmACH.GetDispStr = Empty Then
        
        If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018" & _
                                " WHERE CODENO=" & gilBankCode.GetCodeNo & _
                                " AND (PRGNAME LIKE 'ACH%' Or " & posAchContition & " ) " & _
                                " AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
            If Not blnACHCustID Then
                On Error Resume Next
                csmACH.SetFocus
                GoTo 66
                Exit Function
            Else
                If GetRsValue("Select Count(*) From " & GetOwner & "SO106A" & _
                        " Where MasterId=" & IIf(rsSO106.EOF, 0, rsSO106("MasterId")) & " And AuthorizeStatus = 1") = 0 Then
                    On Error Resume Next
                    csmACH.SetFocus
                    GoTo 66
                End If
            End If
        End If
    End If
    '*******************************************************************************
    '#3436 �P�_�O�_���o����� By Kin 2007/12/13
    
    '#3982 �ˬdSO138�O�_���b���Ҧ��H�����,�p�G�S�������s�WSO138,�p�G����User�ۦ��� By Kin 2008/07/16
    If EditMode = giEditModeInsert Then
'        If Not ChkInvData Then
            '***********************************************************
            '#4210 ���ˬd�O�_�s�b�F,�����אּ�����s�W By Kin 2008/11/06
'            If strInvSeqNo = Empty Or Val(strInvSeqNo) <= 0 Then
'                If ChkSO138 Then
'                    MsgBox "�п�ܤ@���o����� !!", vbInformation, "�T��"
'                    If cmdCreateINV.Enabled Then cmdCreateINV.Value = True
'                    Exit Function
'                Else
'                    blnInsSO138 = True
'                    'Call InsSO138
'                End If
'            End If
            blnInsSO138 = True
            '*************************************************************
 '       End If
    ElseIf EditMode = giEditModeEdit Then
        If chkStop.Value <> 1 And gdaSnactionDate.Text <> Empty Then
            If rsSO106("InvSeqNo") & "" = "" Or Val(rsSO106("InvSeqNO") & "") <= 0 Then
                '***************************************************************
                '#4210 ���ˬd�O���s�b�F,�����אּ�����s�W By Kin 2008/11/06
'                If strInvSeqNo = "" Or Val(strInvSeqNo) <= 0 Then
'                    If ChkSO138 Then
'                        MsgBox "�п�ܤ@���o����� !!", vbInformation, "�T��"
'                        If cmdCreateINV.Enabled Then cmdCreateINV.Value = True
'                        Exit Function
'                    Else
'                        'Call InsSO138
'                        blnInsSO138 = True
'                    End If
'                End If
                blnInsSO138 = True
                '*****************************************************************
            Else
                strInvSeqNo = rsSO106("InvSeqNo") & ""
            End If
        End If

    End If
    '******************************************************************************
    IsDataOk = True
  Exit Function
66:
    MsgBox " �������n���,������ !! ", vbExclamation, "�T��.."
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub NewRcd()
  On Error Resume Next
    Dim ctl As Control
    '#3322 �h�W�@��Flag�A�H�P�O�O�_���s�W���A(�]������Show�XACH���A���ܪ��T��) By Kin 2007/07/04
    blnAddFlag = True
    For Each ctl In Me
        If TypeOf ctl Is TextBox Then ctl = ""
        If TypeOf ctl Is GiDate Then Call ctl.SetValue("")
        If TypeOf ctl Is GiTime Then ctl.Text = ""
        If TypeOf ctl Is GiList Then ctl.Clear
        If TypeOf ctl Is CheckBox Then ctl.Value = 0
        If TypeOf ctl Is MaskEdBox Then ctl.Text = ""
        If TypeOf ctl Is CSmulti Then ctl.Clear
    Next
    chkAddAcc.Value = 1
    lblUpdTime.Caption = GetDTString(RightNow)
    lblUpdEn.Caption = garyGi(1)
    cmdInvSeqNo.Enabled = True
    cmdCreateINV.Enabled = True
    strInvSeqNo = Empty
    lblInvSeqNo = ""
    '�s�W��Ʈɹw�]�����ί�ϥ� By Kin 2008/09/19
    chkStop.Enabled = True
    gdaStopDate.Enabled = True
    '#3728 ���e�S���L�o����,�b�����D���o�{,�N���ιL�o�� By Kin 2008/03/10
    '#3946 �h�W�[ ACHTNO is Not Null ���� By Kin 2008/05/23
    '#4030 ���L�o���� By Kin 2008/08/26
    '#4106 �W�[�P�_ACHType�Ѽ� By Kin 2008/09/22
    '#7049 �W�[BillHeadFmt���,�i�H���CD068A By Kin 2015/07/08
    If blnStartPost Then
        AchTypeCondition = " In (1,2 ) "
    Else
        AchTypeCondition = " In ( 1 ) "
    End If
    
    SetMsQry csmACH, "SELECT ACHTNO,ACHTDESC,CITEMCODESTR,BillHeadFmt FROM " & GetOwner & "CD068 Where ACHTNO Is Not Null And ACHTDESC is Not NULL And ACHType " & AchTypeCondition, _
                           "ACH����N�X", "ACH�������", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
    csmACH.Clear
    csmCitem1.Clear
    csmCitem2.Clear
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "NewRcd")
End Sub

Public Sub UpdateGo()
  On Error GoTo ChkErr
    Dim strSQL As String
    Dim lngChargeAddrNo As Long
    Dim lngMailAddrNo As Long
    Dim blnReturn As Boolean
    Dim blnAdjust As Boolean
    Dim lngPos As Long
    
    If IsDataOk() = False Then Exit Sub
    
    If EditMode = giEditModeEdit Then lngPos = rsSO106.AbsolutePosition
    
    gcnGi.BeginTrans
    '#4068 �N�۰ʷs�WSO138��ӳo��,���MstrInvseqNo�|�O�ŭ�,�y��SO003�S�QUpdate�� By Kin 2008/09/02
    '#4130 �W�[�P�_strInvSeqNo�O�_����,�p�G���ȭn�H���Ȫ����� By Kin 2008/11/10
    If blnInsSO138 Then
        If strInvSeqNo = Empty Then
            Call InsSO138
        Else
            blnInsSO138 = False
        End If
    End If
    If EditMode = giEditModeEdit Then
        If Not IsNull(rsSO106.Fields("ACHTNO").OriginalValue) And csmACH.GetDispStr = Empty Then
            chkStop.Value = 1
        End If
    
    End If
    If GetUserPriv("SO1100G", "(SO1100G9)") Then
       strRealAccNo = SolveNull(txtAccountNo & "")
    Else
        If txtAccountNo.Tag = Empty And txtAccountNo <> Empty Then txtAccountNo.Tag = txtAccountNo
        If IsNumeric(txtAccountNo) Then
            strRealAccNo = SolveNull(txtAccountNo)
        Else
            strRealAccNo = SolveNull(txtAccountNo.Tag)
        End If
    End If
    
    If chkStop.Value <> 1 And gdaSnactionDate.Text <> Empty Then
    Else
        StopSO002A
        If StrComp(strRealAccNo, frmSO1100BMDI.rsSo002!AccountNo & "", vbTextCompare) = 0 Then
            blnReturn = MsgBox("�O�_�M���Ȥ��ƥD�ɪ��Ȥ�򥻸�� (�G) �����ڸ�� ? �ÿ�ܷs�����O�覡 !", vbQuestion + vbYesNo, "�߰�") = vbYes
        End If
    End If
    
    If chkStop.Value <> 1 And gdaSnactionDate.Text <> Empty Then
        ChkSO002A
        If chkAddAcc.Value <> 1 Then StopSO002A
        ChkSO003
        ChkSO033
    End If
    
    If chkStop.Value = 1 And gdaStopDate.Text <> Empty Then StopGo
    '#4130 ���դ�OK,�p�G�y������0,�nRollback,���M�|Trans�G�� By Kin 2008/11/10
    If Not ScrToRcd Then gcnGi.RollbackTrans: Exit Sub            ' �ⱱ������ȡA�s���Ʈw��
    gcnGi.CommitTrans
    If blnCallSO106A And blnACHCustID Then
        cmdACHDetail.Value = True
    End If
    With rsSO106
        
        If .State = adStateOpen Then
    '        rsSO106.Requery
            strSQL = .Source
            .Close
            .CursorLocation = adUseClient
            .Open strSQL, gcnGi, adOpenKeyset, adLockOptimistic
            Set ggrData.Recordset = rsSO106
            
        End If
'        Set frmSO1100BMDI.rsSO106 = rsSO106
    End With
    
    With frmSO1100BMDI
        .rsSO106.Requery
        Set .ggrData.Recordset = .rsSO106
        .ggrData.Refresh
        '***********************************************
        '#3929 �o�{�����D,�p�G�b�s��Ҧ�,�n���s���wfrmSO1100BMDI.rsSo106.AbsolutePosititon���M�b���s��ɷ|����Ĥ@�� By Kin 2008/06/09
        If lngEditMode = giEditModeEdit Then
            .rsSO106.AbsolutePosition = lngPos
        End If
        '***********************************************
    End With
    
    ggrData.Refresh
    
    ggrData.Blank = False
    If blnReturn Then
        ChangeMode giEditModeView
'        Me.Visible = False
'        DoEvents
        Me.Hide
        Unload Me
        frmSO1100BMDI.ProcTrans
        Exit Sub
    End If
'    If blnAdjust Then
'        Me.Hide
'        ChangeMode giEditModeView
'        Unload Me
'        frmSO1100BMDI.AdjustCM
'        Exit Sub
'    End If
'     If EditMode = giEditModeInsert Then '�~��s�W
'        Call ChangeMode(giEditModeInsert)
'        AddNewGo
'    Else '�i�J��ܼҦ�
        If EditMode = giEditModeEdit Then rsSO106.AbsolutePosition = lngPos
        Call ChangeMode(giEditModeView)
        RcdToScr
        If blnInsSO138 Then MsgBox "�t�Τw�۰ʲ��͵o�����Y���", vbInformation, "�T��"
        blnInsSO138 = False
'    End If
   Exit Sub
ChkErr:
    gcnGi.RollbackTrans
    Call ErrSub(Me.Name, "UpdateGo")
End Sub

Private Sub StopGo(Optional strAccNo As String = "")
  On Error GoTo ChkErr
    Dim strBankCode As String
    Dim strCMCode As String, strCMName As String
    Dim strPTCode As String, strPTName As String
    Dim rsTmp As New ADODB.Recordset
    Dim RetVal As Long
    Dim strChkSO003 As String
    Dim varSeqNo() As String
    Dim strCitemCode As String
    Dim strCitemSQL As String
    Dim strBillNos As String
    Dim strBillNoSQL As String
    Dim strUCCode As String
    Dim strUCName As String
    Dim strServiceType As String
    Dim rsPara As New ADODB.Recordset
    Dim rsBill As New ADODB.Recordset
    Dim strUCCodeSQL As String
    Dim i As Long
    strCMCode = frmSO1100BMDI.rsSO044!CMcode
    strCMName = GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [���O�覡�N�X��] �d�L�N�X�� " & frmSO1100BMDI.rsSO044!CMcode & " ����� !")
    
    If strCMName = Empty Then Exit Sub
    
    strBankCode = gilBankCode.GetCodeNo
    If strAccNo = Empty Then strAccNo = strRealAccNo
    '#5623 �쥻�ORowNum=1,�{�b�אּCodeNo=1
    If GetRS(rsTmp, "SELECT CODENO,DESCRIPTION FROM " & GetOwner & "CD032 WHERE CodeNo=1") Then
        If Not rsTmp.EOF Then
            strPTCode = rsTmp("CODENO") & ""
            strPTName = rsTmp("DESCRIPTION") & ""
            If strPTCode = Empty Then strPTCode = "NULL"
        End If
    End If
    
    
    If csmCitem1.GetQueryCode = Empty Then
        ReDim varSeqNo(0)
        varSeqNo(0) = ""
    Else
        varSeqNo = Split(csmCitem1.GetQueryCode, ",")
    End If
    
    '#3573 ���դ�OK,��ΰ}�C,���o�C�@�Ӫ��� By Kin 2007/12/04
    For i = 0 To UBound(varSeqNo)
        strChkSO003 = "Select Count(*) Cnt From " & GetOwner & "SO106  Where CustId=" & gCustId & _
                    " And AccountID='" & strAccNo & "'" & _
                    " And CompCode=" & gCompCode & _
                    " And StopFlag=0" & _
                    " And StopDate is Null" & _
                    IIf(EditMode = giEditModeEdit, " And RowID<>'" & rsSO106("RowID") & "'", "") & _
                    IIf(varSeqNo(i) <> "", " And Instr(','||Citemstr||',' ,','||Chr(39)||" & varSeqNo(i) & "||Chr(39)||',')>0", "")
        If gcnGi.Execute(strChkSO003)("Cnt") = 0 Then

            '#3573 ���դ�OK,�M��SO033��Ƥ]�n�ھڦ��O���� By Kin 2007/12/06
            If varSeqNo(i) <> "" Then
                strCitemSQL = "Select CitemCode From " & GetOwner & "SO003 Where CustId=" & gCustId & _
                            " And AccountNo='" & strAccNo & "'" & _
                            " And CompCode=" & gCompCode & _
                            " And StopFlag=0" & _
                            " And SeqNo=" & varSeqNo(i) & _
                            IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode)
                
                strCitemCode = GetRsValue(strCitemSQL, gcnGi) & ""

            End If
            
            '#3680 �p�G�ӱb���䤣��g���ʶ���,�h������SO003�PSO033 By Kin 2007/12/12
            '#3417 �p�G���έn�sSO003�PSO033��InvSeqNo�@�ֲM�� By Kin 2007/12/13
            '#5173  SO003�@�w�n�j������,��SO033�h�٬O�n�P�_,�ҥH�NIf strCitemCode���ܤU�� By Kin 2009/07/06
'            If strCitemCode <> "" Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                "BANKCODE=NULL" & _
                                ",BANKNAME=NULL" & _
                                ",ACCOUNTNO=NULL" & _
                                ",InvSeqNo=NULL" & _
                                ",CMCODE=" & strCMCode & _
                                ",CMNAME='" & strCMName & "'" & _
                                ",PTCODE=" & strPTCode & _
                                ",PTNAME='" & strPTName & "'" & _
                                " WHERE CUSTID=" & gCustId & _
                                " AND ACCOUNTNO='" & strAccNo & "'" & _
                                " AND COMPCODE=" & gCompCode & _
                                 IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                 IIf(csmCitem1.GetQueryCode = Empty, "", " AND SEQNO IN (" & varSeqNo(i) & ")"), RetVal
                Debug.Print RetVal
            If strCitemCode <> "" Then
                '#3573 ���դ�OK,�h�W�[CitemCode������,�קK�������M�� By Kin 2007/12/06
                '#5623 �n�N�ۦP��BillNo�A�M���@�� By Kin 2010/04/02
                strBillNoSQL = "SELECT BILLNO,ServiceType FROM " & GetOwner & "SO033 " & _
                            " WHERE CUSTID=" & gCustId & _
                            " AND ACCOUNTNO='" & strAccNo & "'" & _
                            " AND COMPCODE=" & gCompCode & _
                            IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                            " AND UCCODE > 0 AND CANCELFLAG=0" & _
                            IIf(strCitemCode <> "", " AND CitemCode=" & strCitemCode, "") & _
                            " AND ROWNUM=1"
                If Not GetRS(rsBill, strBillNoSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                If Not rsBill.EOF Then
                    rsBill.MoveFirst
                    Do While Not rsBill.EOF
                        strBillNos = rsBill("BillNo") & ""
                        strServiceType = rsBill("ServiceType") & ""
                        gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                "BANKCODE=NULL" & _
                                ",BANKNAME=NULL" & _
                                ",ACCOUNTNO=NULL" & _
                                ",InvSeqNo=NULL" & _
                                ",CMCODE=" & strCMCode & _
                                ",CMNAME='" & strCMName & "'" & _
                                ",PTCODE=" & strPTCode & _
                                ",PTNAME='" & strPTName & "'" & _
                                " WHERE CUSTID=" & gCustId & _
                                " AND ACCOUNTNO='" & strAccNo & "'" & _
                                " AND COMPCODE=" & gCompCode & _
                                 IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                 " AND UCCODE > 0 AND CANCELFLAG=0" & _
                                 IIf(strCitemCode <> "", " AND CitemCode=" & strCitemCode, ""), RetVal
                                 
                        If strBillNos <> "" Then
                            strUCCodeSQL = "SELECT CODENO,Description FROM " & GetOwner & "CD013 " & _
                                        " WHERE CODENO = (SELECT UCCODE FROM " & GetOwner & "SO044 " & _
                                        " WHERE SERVICETYPE='" & strServiceType & "'" & _
                                        " AND COMPCODE= " & gCompCode & " )" & _
                                        " AND STOPFLAG<>1"
                            If Not GetRS(rsPara, strUCCodeSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                            If Not rsPara.EOF Then
                                strUCCode = rsPara("CODENO") & ""
                                strUCName = rsPara("DESCRIPTION") & ""
                            End If
                            '#5623 �n�N�ۦP��BillNo�A�M���@�� By Kin 2010/04/02
                            gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                            "BANKCODE=NULL" & _
                                            ",BANKNAME=NULL" & _
                                            ",ACCOUNTNO=NULL" & _
                                            ",InvSeqNo=NULL" & _
                                            ",CMCODE=" & strCMCode & _
                                            ",CMNAME='" & strCMName & "'" & _
                                            ",PTCODE=" & strPTCode & _
                                            ",PTNAME='" & strPTName & "'" & _
                                            ",UCCODE=" & IIf(strUCCode = "", "UCcode", strUCCode) & _
                                            ",UCNAME='" & IIf(strUCName = "", "UCName", strUCName & "'") & _
                                            " WHERE CUSTID=" & gCustId & _
                                            " AND COMPCODE=" & gCompCode & _
                                            " AND UCCODE > 0 AND CANCELFLAG=0" & _
                                            " AND BILLNO='" & strBillNos & "'", RetVal
                        End If
                        rsBill.MoveNext
                    Loop
                    
                End If
            End If
        End If
    Next i
    Debug.Print RetVal
    StopSO002A strAccNo
    
    On Error Resume Next
    CloseRS rsTmp
    CloseRS rsPara
    CloseRS rsBill
  Exit Sub
ChkErr:
    ErrSub Me.Name, "StopGo"
End Sub
Private Function ChkInvData() As Boolean
  On Error GoTo ChkErr
    Dim bytCnt As Byte
    Dim strAccNo As String
    If Not rsSO106.EOF Then
        If rsSO106("AccountID").OriginalValue & "" = Empty Then
            strAccNo = txtAccountNo
        Else
            If EditMode = giEditModeInsert Then
                strAccNo = txtAccountNo
            Else
                strAccNo = rsSO106("AccountID").OriginalValue & ""
            End If
        End If
    Else
        strAccNo = txtAccountNo
    End If

    bytCnt = Val(RPxx(gcnGi.Execute("Select count(*) from " & GetOwner & "so002a where " & _
                      "AccountNo='" & strAccNo & "' and CustId=" & _
                       gCustId & " and compcode=" & gCompCode & " And StopFlag<>1").GetString))
    
    If EditMode = giEditModeEdit Then
        If (rsSO106("InvSeqNo") & "" <> "") And (Not IsNull(rsSO106("SnactionDate"))) Then ChkInvData = True
    End If
    If bytCnt > 0 Then ChkInvData = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "ChkInvData"
End Function
Private Sub ChkSO002A()
  On Error GoTo ChkErr
'    0 = ��b�b��
'    1 = �H�Υd��
'    2 = �����b��
  
    If txtAccountNo = Empty Then Exit Sub
    Dim bytCnt As Byte
    Dim strAccNo As String
    Dim strBankCode As String
    Dim intPayID As Integer
    
    If gilCMCode.GetCodeNo <> Empty Then
        On Error Resume Next
        intCD031RefNo = GetRsValue("SELECT REFNO FROM " & GetOwner & "CD031 WHERE CODENO=" & gilCMCode.GetCodeNo)
        If Err.Number <> 0 Then
            Err.Clear
            intCD031RefNo = 0
        End If
    End If
    '#8724 Add refno=5
    If intCD031RefNo = 2 Or intCD031RefNo = 4 Or intCD031RefNo = 5 Then
        blnVAcc = False
    Else
        blnVAcc = True
    End If
    
    If blnVAcc Then
        intPayID = 2
    Else
        '#8724 add refno=5
        If intCD031RefNo = 4 Or intCD031RefNo = 5 Then
            If gilCardName.GetCodeNo <> Empty Then
                intPayID = 1
            Else
                intPayID = 0
            End If
        Else
            intPayID = 0
        End If
    End If
    
    With frmSO1100BMDI
        If Not rsSO106.EOF Then
            If rsSO106("AccountID").OriginalValue & "" = Empty Then
                strAccNo = txtAccountNo
                strBankCode = gilBankCode.GetCodeNo
            Else
                If EditMode = giEditModeInsert Then
                    strAccNo = txtAccountNo
                    strBankCode = gilBankCode.GetCodeNo
                Else
                    strAccNo = rsSO106("AccountID").OriginalValue & ""
                    strBankCode = rsSO106("BankCode").OriginalValue & ""
                End If
            End If
        Else
            strAccNo = txtAccountNo
        End If
        bytCnt = Val(RPxx(gcnGi.Execute("select count(*) from " & GetOwner & "so002a where " & _
                      "accountNo='" & strAccNo & "' and custid=" & _
                       gCustId & " and compcode=" & gCompCode).GetString))
        
        
'    txtChargeTitle = frmSO1100BMDI.txtCustName
'    txtInvTitle.Text = frmSO1100BMDI.rsSo002.Fields("InvTitle") & ""
'    mskInvNo.Text = frmSO1100BMDI.rsSo002.Fields("InvNo") & ""
'    txtInvAddress.Text = frmSO1100BMDI.rsSo002.Fields("InvAddress") & ""
'    cboInvoiceType = frmSO1100BMDI.ChildForm.cboInvoiceType
        
        If strAccNo = Empty Then strAccNo = txtAccountNo
        
        If bytCnt = 0 Then
            gcnGi.Execute "INSERT INTO " & GetOwner & "SO002A " & _
                            "(CUSTID,COMPCODE,BANKCODE,BANKNAME,ID,ACCOUNTNO," & _
                            "CARDNAME,CARDEXPDATE,CHARGEADDRNO,CHARGEADDRESS," & _
                            "MAILADDRNO,MAILADDRESS,CVC2,NOTE,CHARGETITLE," & _
                            "INVNO,INVTITLE,INVADDRESS,INVOICETYPE," & _
                            "CITEMSTR,CITEMSTR2,ADDCITEMACCOUNT) VALUES (" & _
                             gCustId & "," & gCompCode & "," & gilBankCode.GetCodeNo2 & ",'" & _
                             gilBankCode.GetDescription & "'," & intPayID & ",'" & strAccNo & "','" & _
                             gilCardName.GetDescription & "'," & IIf(Len(mskCardExpDate.Text) = 0, "Null", mskCardExpDate.Text) & "," & _
                            .rsSo001!ChargeAddrNo & ",'" & .rsSo001!ChargeAddress & "'," & _
                            .rsSo001!MailAddrNo & ",'" & .rsSo001!MailAddress & "','" & txtCVC2 & "','" & txtNote & "','" & txtAccountOwner & "','" & _
                             frmSO1100BMDI.rsSo002.Fields("InvNo") & "','" & frmSO1100BMDI.rsSo002.Fields("InvTitle") & "','" & _
                             frmSO1100BMDI.rsSo002.Fields("InvAddress") & "'," & frmSO1100BMDI.rsSo002.Fields("InvoiceType") & ",'" & _
                             csmCitem1.QryCD & "','" & csmCitem2.QryCD & "'," & chkAddAcc.Value & ")"
                             
            '*************************************************************************************************
            '#3436 �s�W�ɭn�s�P�o����T�s�W�L�h By Kin 2007/12/13
            If strInvSeqNo <> "" Then
                If gcnGi.Execute("Select Count(*) From " & GetOwner & "SO002AD Where Compcode=" & gCompCode & _
                                    " And CustId=" & gCustId & " And AccountNo='" & strAccNo & "'" & _
                                    " And InvSeqNo=" & strInvSeqNo)(0) = 0 Then
                                    
                       gcnGi.Execute "INSERT INTO " & GetOwner & "SO002AD " & _
                                    "(CUSTID,AccountNo,CompCode,InvSeqNo) VALUES (" & _
                                    gCustId & ",'" & strAccNo & "'," & gCompCode & "," & strInvSeqNo & ")"
                End If
            End If
            '***************************************************************************************************
        ElseIf bytCnt = 1 Then
            ' MinChen �ݨD , �ק�ɤ� Update Address �������� SO002A
            gcnGi.Execute "UPDATE " & GetOwner & "SO002A SET " & _
                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                            ",BANKNAME='" & gilBankCode.GetDescription & _
                            "',ID=" & intPayID & _
                            ",ACCOUNTNO='" & strRealAccNo & _
                            "',CARDNAME='" & gilCardName.GetDescription & _
                            "',CARDEXPDATE=" & IIf(Len(mskCardExpDate.Text) = 0, "Null", mskCardExpDate.Text) & _
                            ",CVC2='" & txtCVC2 & _
                            "',NOTE='" & txtNote & "'" & _
                            ",CITEMSTR='" & csmCitem1.QryCD & "'" & _
                            ",CITEMSTR2='" & csmCitem2.QryCD & "'" & _
                            ",ADDCITEMACCOUNT=" & chkAddAcc.Value & _
                            ",STOPFLAG=0,STOPDATE=NULL" & _
                            " WHERE CUSTID=" & gCustId & _
                            " AND ACCOUNTNO='" & strAccNo & _
                            "' AND COMPCODE=" & gCompCode & _
                            IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode)
                            
                            '#3929 �P�B��sSO002A�ɱNChargeTitle��������� By Kin 2008/06/02
                            '" AND CHARGETITLE='" & txtAccountOwner & "'" & _

            '*************************************************************************************************
            '#3436 �s�W�ɭn�s�P�o����T�s�W�L�h By Kin 2007/12/13
            If strInvSeqNo <> "" Then
                If gcnGi.Execute("Select Count(*) From " & GetOwner & "SO002AD Where Compcode=" & gCompCode & _
                                    " And CustId=" & gCustId & " And AccountNo='" & strRealAccNo & "'" & _
                                    " And InvSeqNo=" & strInvSeqNo)(0) = 0 Then
                                    
                                    
                    gcnGi.Execute "INSERT INTO " & GetOwner & "SO002AD " & _
                                    "(CUSTID,AccountNo,CompCode,InvSeqNo) VALUES (" & _
                                    gCustId & ",'" & strRealAccNo & "'," & gCompCode & "," & strInvSeqNo & ")"
                End If
            End If
            '***************************************************************************************************

'            gcnGi.Execute "UPDATE " & GetOwner & "SO002A SET " & _
                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                            ",BANKNAME='" & gilBankCode.GetDescription & _
                            "',ID=" & intPayID & _
                            ",ACCOUNTNO='" & strRealAccNo & _
                            "',CARDNAME='" & gilCardName.GetDescription & _
                            "',CARDEXPDATE=" & IIf(Len(mskCardExpDate.Text) = 0, "Null", mskCardExpDate.Text) & _
                            ",CHARGEADDRNO=" & .rsSo001!ChargeAddrNo & _
                            ",CHARGEADDRESS='" & .rsSo001!ChargeAddress & _
                            "',MAILADDRNO=" & .rsSo001!MailAddrNo & _
                            ",MAILADDRESS='" & .rsSo001!MailAddress & _
                            "',CVC2='" & txtCVC2 & _
                            "',NOTE='" & txtNote & "'" & _
                            ",CITEMSTR='" & csmCitem1.QryCD & "'" & _
                            ",CITEMSTR2='" & csmCitem2.QryCD & "'" & _
                            ",ADDCITEMACCOUNT=" & chkAddAcc.Value & _
                            ",STOPFLAG=0,STOPDATE=NULL" & _
                            " WHERE CUSTID=" & gCustId & _
                            " AND ACCOUNTNO='" & strAccNo & _
                            "' AND COMPCODE=" & gCompCode & _
                            " AND CHARGETITLE='" & txtAccountOwner & "'" & _
                            IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode)
        End If
    End With
    
    Dim strChild2A As String
    strChild2A = Get2AChild
    
    If strChild2A <> Empty And strChild2A <> "''" Then
        gcnGi.Execute "UPDATE " & GetOwner & "SO002A SET " & _
                        "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                        ",BANKNAME='" & gilBankCode.GetDescription & _
                        "',ID=" & intPayID & _
                        ",ACCOUNTNO='" & strRealAccNo & _
                        "',CARDNAME='" & gilCardName.GetDescription & _
                        "',CARDEXPDATE=" & IIf(Len(mskCardExpDate.Text) = 0, "Null", mskCardExpDate.Text) & _
                        ",CVC2='" & txtCVC2 & _
                        "',NOTE='" & txtNote & "'" & _
                        ",CITEMSTR='" & csmCitem1.QryCD & "'" & _
                        ",CITEMSTR2='" & csmCitem2.QryCD & "'" & _
                        ",ADDCITEMACCOUNT=" & chkAddAcc.Value & _
                        ",STOPFLAG=0,STOPDATE=NULL" & _
                        " WHERE INHERITKEY IN (" & strChild2A & ")" & _
                        " AND INHERITFLAG=1" & _
                        " AND COMPCODE=" & gCompCode
    End If
    
    '(AccountNo=SO106. AccountID and CustId=SO106.CustId and CompCode=SO106.CompCode).
    'If �S���ӱb����� Then Insert into SO002A
    '
    'PS: �s�W��SO002A��, ���O�a�}�P�l�H�a�}�����h�Ѱ򥻸�Ƥ��a�J��w�]��.
    'PS:  ���ެO�_�n��s��SO002�ҭn����Ȥ�b��Insert into�y�{
    '
    '�Ŀﰱ�γ��� : �����u�������ӵL��L�B�z����.
    '�{�վ㬰�n�R��SO002A�����ӵ��b��/�d��,
    '�H�קKUser�A���~���ӱb��.
    '
    '��g����A�������ئ��Ψ�ӱb�������� , �h�N���� update��
    '
    'SO002�����D�b��������ƨ�show Message�i��User�����p,
    '
    '����User�P�_�O�_�٭n�ܶg���ʭ��Ҥ��վ��Ƥ��e.
    '
    '�H�W��Ӿ����_�ӫ� , SO002A�h�����Ȥ᪺���ıb��޲z, �H��K�h�b�����ϥ�
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ChkSO002A"
End Sub

Private Sub ChkSO003()
  On Error GoTo ChkErr
    If txtAccountNo = Empty Then Exit Sub
    Dim strAccNo As String
    Dim strBankCode As String
    Dim strSeqNo As String
    Dim strQryCnt As String
    Dim blnUpdINV As Boolean
    blnUpdINV = True
    With frmSO1100BMDI
        If EditMode = giEditModeInsert Then
            strAccNo = txtAccountNo
            strBankCode = gilBankCode.GetCodeNo
            strSeqNo = csmCitem1.GetQueryCode
        
        Else
            If Not rsSO106.EOF Then
                If rsSO106("AccountID").OriginalValue & "" = Empty Then
                    strAccNo = txtAccountNo
                    strBankCode = gilBankCode.GetCodeNo
    '                strCitemCode = gimCitem.GetQueryCode
                    strSeqNo = csmCitem1.GetQueryCode
                Else
                    strAccNo = rsSO106("AccountID").OriginalValue & ""
                    strBankCode = rsSO106("BankCode").OriginalValue & ""
                    strSeqNo = rsSO106("CitemStr").OriginalValue & ""
                End If
            Else
                strAccNo = txtAccountNo
                strBankCode = gilBankCode.GetCodeNo
    '            strCitemCode = gimCitem.GetQueryCode
                strSeqNo = csmCitem1.GetQueryCode
            End If
        End If
        
        Dim RetVal As Integer
        Dim strChildCustID As String
        '***********************************************************************************************
        '#5045 �p�G���o�����Y���@�˪���Ƹ߰�User�O�_�^��,��ܧ_�h���������} By Kin 2009/05/04
        strQryCnt = "Select Count(*) Cnt From " & GetOwner & " SO003 " & _
                    " Where INVSEQNO IS Not Null And INVSEQNO<>" & strInvSeqNo & _
                    " And CustID=" & gCustId & " And CompCode=" & gCompCode & _
                    " And ACCOUNTNO='" & strAccNo & "'" & _
                    " And CMCODE=" & gilCMCode.GetCodeNo & _
                    IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                    IIf(strSeqNo = Empty, "", " AND SEQNO IN (" & strSeqNo & ")") & _
                    " And PTCODE=" & gilPTCode.GetCodeNo
        '***********************************************************************************************
        If Val(gcnGi.Execute(strQryCnt)("Cnt")) > 0 Then
            If MsgBox("���b��ҫ��w���o�����Y�P�g���ʦ��O���ؤ��šA�O�_�^��H", vbQuestion + vbYesNo, "�T��") = vbNo Then
                blnUpdINV = False
            End If
        End If
        
        strChildCustID = GetChild
        
        If EditMode = giEditModeEdit Then
'            strCitemCode = gimCitem.GetQueryCode
            strSeqNo = csmCitem1.GetQueryCode
'            If gimCitem.GetQueryCode = Empty And rsSO106("CitemStr").OriginalValue <> Empty Then
            If csmCitem1.GetQueryCode = Empty And rsSO106("CitemStr").OriginalValue <> Empty Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                "BANKCODE=NULL" & _
                                ",BANKNAME=NULL" & _
                                ",ACCOUNTNO=NULL" & _
                                IIf(blnUpdINV = True, ",InvSeqNO=NULL", "") & _
                                ",CMCODE=" & frmSO1100BMDI.rsSO044!CMcode & _
                                ",CMNAME='" & GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [���O�覡�N�X��] �d�L�N�X�� " & frmSO1100BMDI.rsSO044!CMcode & " ����� !") & "'" & _
                                " WHERE CUSTID=" & gCustId & _
                                " AND ACCOUNTNO='" & strAccNo & "'" & _
                                " AND COMPCODE=" & gCompCode & _
                                 IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                 IIf(strSeqNo = Empty, "", " AND SEQNO IN (" & rsSO106("CitemStr").OriginalValue & ")"), RetVal
                Debug.Print RetVal
                If strChildCustID <> Empty Then
                    gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                    "BANKCODE=NULL" & _
                                    ",BANKNAME=NULL" & _
                                    ",ACCOUNTNO=NULL" & _
                                    IIf(blnUpdINV = True, ",InvSeqNo=NULL", "") & _
                                    ",CMCODE=" & frmSO1100BMDI.rsSO044!CMcode & _
                                    ",CMNAME='" & GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [���O�覡�N�X��] �d�L�N�X�� " & frmSO1100BMDI.rsSO044!CMcode & " ����� !") & "'" & _
                                    " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                    " AND ACCOUNTNO='" & strAccNo & "'" & _
                                    " AND COMPCODE=" & gCompCode & _
                                     IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                     IIf(strSeqNo = Empty, "", " AND SEQNO IN (" & rsSO106("CitemStr").OriginalValue & ")"), RetVal
                    Debug.Print RetVal
                End If
            End If
'            If gimCitem.GetQueryCode <> Empty And rsSO106("CitemStr").OriginalValue <> Empty Then
            If csmCitem1.GetQueryCode <> Empty And rsSO106("CitemStr").OriginalValue <> Empty Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                "BANKCODE=NULL" & _
                                ",BANKNAME=NULL" & _
                                ",ACCOUNTNO=NULL" & _
                                IIf(blnUpdINV = True, ",InvSeqNo=NULL", "") & _
                                ",CMCODE=" & frmSO1100BMDI.rsSO044!CMcode & _
                                ",CMNAME='" & GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [���O�覡�N�X��] �d�L�N�X�� " & frmSO1100BMDI.rsSO044!CMcode & " ����� !") & "'" & _
                                " WHERE CUSTID=" & gCustId & _
                                " AND ACCOUNTNO='" & strAccNo & "'" & _
                                " AND COMPCODE=" & gCompCode & _
                                 IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                 IIf(strSeqNo = Empty, "", " AND SEQNO IN (" & rsSO106("CitemStr").OriginalValue & "" & ")"), RetVal
                Debug.Print RetVal
                If strChildCustID <> Empty Then
                    gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                    "BANKCODE=NULL" & _
                                    ",BANKNAME=NULL" & _
                                    ",ACCOUNTNO=NULL" & _
                                    IIf(blnUpdINV = True, ",InvSeqNo=NULL", "") & _
                                    ",CMCODE=" & frmSO1100BMDI.rsSO044!CMcode & _
                                    ",CMNAME='" & GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [���O�覡�N�X��] �d�L�N�X�� " & frmSO1100BMDI.rsSO044!CMcode & " ����� !") & "'" & _
                                    " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                    " AND ACCOUNTNO='" & strAccNo & "'" & _
                                    " AND COMPCODE=" & gCompCode & _
                                     IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                     IIf(strSeqNo = Empty, "", " AND SEQNO IN (" & rsSO106("CitemStr").OriginalValue & "" & ")"), RetVal
                    Debug.Print RetVal
                End If
            Else
                If rsSO106("BankCode").OriginalValue & "" <> gilBankCode.GetCodeNo & "" Then
                    If rsSO106("AccountID").OriginalValue & "" = strRealAccNo Then
                        gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                        "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                        ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                        ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                        ",CMCODE=" & gilCMCode.GetCodeNo & _
                                        ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                        ",PTCODE=" & gilPTCode.GetCodeNo & _
                                        ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                        IIf(blnUpdINV = True, IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, ""), "") & _
                                        " WHERE CUSTID=" & gCustId & _
                                        " AND COMPCODE=" & gCompCode & _
                                        " AND BANKCODE=" & rsSO106("BankCode").OriginalValue, RetVal
                        Debug.Print RetVal
                        If strChildCustID <> Empty Then
                            gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                            ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                            ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                            ",CMCODE=" & gilCMCode.GetCodeNo & _
                                            ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                            ",PTCODE=" & gilPTCode.GetCodeNo & _
                                            ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                            IIf(blnUpdINV = True, IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, ""), "") & _
                                            " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                            " AND COMPCODE=" & gCompCode & _
                                            " AND BANKCODE=" & rsSO106("BankCode").OriginalValue, RetVal
                            Debug.Print RetVal
                        End If
                    Else
                        gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                        "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                        ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                        ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                        ",CMCODE=" & gilCMCode.GetCodeNo & _
                                        ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                        ",PTCODE=" & gilPTCode.GetCodeNo & _
                                        ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                        IIf(blnUpdINV = True, IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, ""), "") & _
                                        " WHERE CUSTID=" & gCustId & _
                                        " AND COMPCODE=" & gCompCode & _
                                        " AND BANKCODE=" & rsSO106("BankCode").OriginalValue & _
                                        " AND ACCOUNTNO='" & rsSO106("AccountID").OriginalValue & "'", RetVal
                        Debug.Print RetVal
                        If strChildCustID <> Empty Then
                            gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                            ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                            ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                            ",CMCODE=" & gilCMCode.GetCodeNo & _
                                            ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                            ",PTCODE=" & gilPTCode.GetCodeNo & _
                                            ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                           IIf(blnUpdINV = True, IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, ""), "") & _
                                            " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                            " AND COMPCODE=" & gCompCode & _
                                            " AND BANKCODE=" & rsSO106("BankCode").OriginalValue & _
                                            " AND ACCOUNTNO='" & rsSO106("AccountID").OriginalValue & "'", RetVal
                            Debug.Print RetVal
                        End If
                    End If
                Else
                    If rsSO106("AccountID").OriginalValue & "" <> strRealAccNo Then
                        gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                        "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                        ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                        ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                        ",CMCODE=" & gilCMCode.GetCodeNo & _
                                        ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                        ",PTCODE=" & gilPTCode.GetCodeNo & _
                                        ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                        IIf(blnUpdINV = True, IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, ""), "") & _
                                        " WHERE CUSTID=" & gCustId & _
                                        " AND COMPCODE=" & gCompCode & _
                                        " AND ACCOUNTNO='" & rsSO106("AccountID").OriginalValue & "'", RetVal
                        Debug.Print RetVal
                        If strChildCustID <> Empty Then
                            gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                            ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                            ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                            ",CMCODE=" & gilCMCode.GetCodeNo & _
                                            ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                            ",PTCODE=" & gilPTCode.GetCodeNo & _
                                            ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                            IIf(blnUpdINV = True, IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, ""), "") & _
                                            " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                            " AND COMPCODE=" & gCompCode & _
                                            " AND ACCOUNTNO='" & rsSO106("AccountID").OriginalValue & "'", RetVal
                            Debug.Print RetVal
                        End If
                    End If
                End If
            End If
        End If
        
'        If gimCitem.GetQueryCode <> Empty Then
        If csmCitem1.GetQueryCode <> Empty Then
'            strCitemCode = gimCitem.GetQueryCode
            strSeqNo = csmCitem1.GetQueryCode
            gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                            ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                            ",ACCOUNTNO='" & strRealAccNo & "'" & _
                            ",CMCODE=" & gilCMCode.GetCodeNo & _
                            ",CMNAME='" & gilCMCode.GetDescription & "' " & _
                            ",PTCODE=" & gilPTCode.GetCodeNo & _
                            ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                            IIf(blnUpdINV = True, IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, ""), "") & _
                            " WHERE CUSTID=" & gCustId & _
                            " AND COMPCODE=" & gCompCode & _
                             IIf(strSeqNo = Empty, "", " AND SEQNO IN (" & strSeqNo & ")"), RetVal
            Debug.Print RetVal
            If strChildCustID <> Empty Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO003 SET " & _
                                "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                ",CMCODE=" & gilCMCode.GetCodeNo & _
                                ",CMNAME='" & gilCMCode.GetDescription & "' " & _
                                ",PTCODE=" & gilPTCode.GetCodeNo & _
                                ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                IIf(blnUpdINV = True, IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, ""), "") & _
                                " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                " AND COMPCODE=" & gCompCode & _
                                 IIf(strSeqNo = Empty, "", " AND SEQNO IN (" & strSeqNo & ")"), RetVal
                Debug.Print RetVal
            End If
        End If
    End With
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ChkSO003"
End Sub

Private Function GetChild() As String
  On Error Resume Next
    If rsSO106("INHERITKEY") & "" <> Empty And rsSO106("INHERITFLAG") = 0 Then
        GetChild = gcnGi.Execute("SELECT DISTINCT CUSTID FROM " & GetOwner & "SO106" & _
                                                 " WHERE INHERITFLAG=1 AND INHERITKEY='" & rsSO106("INHERITKEY") & "'").GetString(2, , "", ",", "")
        If Err Then
            Err.Clear
            GetChild = ""
        Else
            GetChild = Replace(GetChild, ",,", ",", 1)
            If Right(GetChild, 1) = "," Then GetChild = Left(GetChild, Len(GetChild) - 1)
        End If
    End If
End Function

Private Sub ChkSO033()
  On Error GoTo ChkErr
    If txtAccountNo = Empty Then Exit Sub
    Dim strAccNo As String
    Dim strBankCode As String
    Dim strRowId As String

    With frmSO1100BMDI
    
        
    
        If Not rsSO106.EOF Then
            If rsSO106("AccountID").OriginalValue & "" = Empty Then
                strAccNo = txtAccountNo
                strBankCode = gilBankCode.GetCodeNo
                strRowId = csmCitem2.GetQryFld(11)
            Else
                strAccNo = rsSO106("AccountID").OriginalValue & ""
                strBankCode = rsSO106("BankCode").OriginalValue & ""
'               strROWID = rsSO106("CitemStr2").OriginalValue & ""
                strRowId = csmCitem2.GetQryFld(11)
            End If
        Else
            strAccNo = txtAccountNo
            strBankCode = gilBankCode.GetCodeNo
            strRowId = csmCitem2.GetQryFld(11)
'            strROWID = csmCitem2.GetQueryCode
        End If
        
        Dim RetVal As Integer
        Dim strChildCustID As String
        strChildCustID = GetChild
        
        If EditMode = giEditModeEdit Then
            strRowId = csmCitem2.GetQryFld(11)
            If csmCitem2.GetQueryCode = Empty And rsSO106("CitemStr2").OriginalValue <> Empty Then
                '#4023 �p�GRowId�O�ŭ�,�h��Update SO033,�åB�NSO106.CitemStr2�M�� By Kin 2008/07/25
                If strRowId <> Empty Then
                    gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                    "BANKCODE=NULL" & _
                                    ",BANKNAME=NULL" & _
                                    ",ACCOUNTNO=NULL" & _
                                    ",InvSeqNo=NULL" & _
                                    ",CMCODE=" & frmSO1100BMDI.rsSO044!CMcode & _
                                    ",CMNAME='" & GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [���O�覡�N�X��] �d�L�N�X�� " & frmSO1100BMDI.rsSO044!CMcode & " ����� !") & "'" & _
                                    " WHERE CUSTID=" & gCustId & _
                                    " AND ACCOUNTNO='" & strAccNo & "'" & _
                                    " AND COMPCODE=" & gCompCode & _
                                     IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                     IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & ")") & _
                                    IIf(blnUCCodeIsNull, "", " AND UCCODE > 0 ") & " AND CANCELFLAG=0", RetVal
                                     Debug.Print RetVal
                End If
                If strChildCustID <> Empty Then
                    If strRowId <> Empty Then
                        gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                        "BANKCODE=NULL" & _
                                        ",BANKNAME=NULL" & _
                                        ",ACCOUNTNO=NULL" & _
                                        ",InvSeqNo=NULL" & _
                                        ",CMCODE=" & frmSO1100BMDI.rsSO044!CMcode & _
                                        ",CMNAME='" & GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [���O�覡�N�X��] �d�L�N�X�� " & frmSO1100BMDI.rsSO044!CMcode & " ����� !") & "'" & _
                                        " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                        " AND ACCOUNTNO='" & strAccNo & "'" & _
                                        " AND COMPCODE=" & gCompCode & _
                                         IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                         IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & ")") & _
                                        IIf(blnUCCodeIsNull, "", " AND UCCODE > 0 ") & " AND CANCELFLAG=0", RetVal
                                         Debug.Print RetVal
                    End If
                End If
            End If
            If csmCitem2.GetQueryCode <> Empty And rsSO106("CitemStr2").OriginalValue <> Empty Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                "BANKCODE=NULL" & _
                                ",BANKNAME=NULL" & _
                                ",ACCOUNTNO=NULL" & _
                                ",InvSeqNo=NULL" & _
                                ",CMCODE=" & frmSO1100BMDI.rsSO044!CMcode & _
                                ",CMNAME='" & GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [���O�覡�N�X��] �d�L�N�X�� " & frmSO1100BMDI.rsSO044!CMcode & " ����� !") & "'" & _
                                " WHERE CUSTID=" & gCustId & _
                                " AND ACCOUNTNO='" & strAccNo & "'" & _
                                " AND COMPCODE=" & gCompCode & _
                                 IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                 IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & "" & ")") & _
                                 IIf(blnUCCodeIsNull, "", " AND UCCODE > 0 ") & " AND CANCELFLAG=0", RetVal
                                 Debug.Print RetVal
                If strChildCustID <> Empty Then
                    gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                    "BANKCODE=NULL" & _
                                    ",BANKNAME=NULL" & _
                                    ",ACCOUNTNO=NULL" & _
                                    ",InvSeqNo=NULL" & _
                                    ",CMCODE=" & frmSO1100BMDI.rsSO044!CMcode & _
                                    ",CMNAME='" & GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD031 WHERE CODENO=" & frmSO1100BMDI.rsSO044!CMcode, gcnGi, "CD031 [���O�覡�N�X��] �d�L�N�X�� " & frmSO1100BMDI.rsSO044!CMcode & " ����� !") & "'" & _
                                    " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                    " AND ACCOUNTNO='" & strAccNo & "'" & _
                                    " AND COMPCODE=" & gCompCode & _
                                     IIf(strBankCode = Empty, "", " AND BANKCODE=" & strBankCode) & _
                                     IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & "" & ")") & _
                                     IIf(blnUCCodeIsNull, "", " AND UCCODE > 0 ") & " AND CANCELFLAG=0", RetVal
                                     Debug.Print RetVal
                End If
            Else
                '#5303 �p�G�S���怜�O���,�|�~UPD�������,�W�[stRowId�P�_ By Kin 2009/09/29
                If strRowId <> Empty Then
                    If rsSO106("BankCode").OriginalValue & "" <> gilBankCode.GetCodeNo & "" Then
                        If rsSO106("AccountID").OriginalValue & "" = strRealAccNo Then
                            gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                            ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                            ",ACCOUNTNO='" & txtAccountNo & "'" & _
                                            ",CMCODE=" & gilCMCode.GetCodeNo & _
                                            ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                            ",PTCODE=" & gilPTCode.GetCodeNo & _
                                            ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                            IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, "") & _
                                            " WHERE CUSTID=" & gCustId & _
                                            " AND COMPCODE=" & gCompCode & _
                                            " AND BANKCODE=" & rsSO106("BankCode").OriginalValue & _
                                            IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & "" & ")"), RetVal
                            Debug.Print RetVal
                            If strChildCustID <> Empty Then
                                gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                                "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                                ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                                ",ACCOUNTNO='" & txtAccountNo & "'" & _
                                                ",CMCODE=" & gilCMCode.GetCodeNo & _
                                                ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                                ",PTCODE=" & gilPTCode.GetCodeNo & _
                                                ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                                IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, "") & _
                                                " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                                " AND COMPCODE=" & gCompCode & _
                                                " AND BANKCODE=" & rsSO106("BankCode").OriginalValue & _
                                                IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & "" & ")"), RetVal
                                Debug.Print RetVal
                            End If
                        Else
                            gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                            ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                            ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                            ",CMCODE=" & gilCMCode.GetCodeNo & _
                                            ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                            ",PTCODE=" & gilPTCode.GetCodeNo & _
                                            ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                            IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, "") & _
                                            " WHERE CUSTID=" & gCustId & _
                                            " AND COMPCODE=" & gCompCode & _
                                            " AND BANKCODE=" & rsSO106("BankCode").OriginalValue & _
                                            " AND ACCOUNTNO='" & rsSO106("AccountID").OriginalValue & "'" & _
                                            IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & "" & ")"), RetVal
                            Debug.Print RetVal
                            If strChildCustID <> Empty Then
                                gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                                "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                                ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                                ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                                ",CMCODE=" & gilCMCode.GetCodeNo & _
                                                ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                                ",PTCODE=" & gilPTCode.GetCodeNo & _
                                                ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                                IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, "") & _
                                                " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                                " AND COMPCODE=" & gCompCode & _
                                                " AND BANKCODE=" & rsSO106("BankCode").OriginalValue & _
                                                " AND ACCOUNTNO='" & rsSO106("AccountID").OriginalValue & "'" & _
                                                IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & "" & ")"), RetVal
                                Debug.Print RetVal
                            End If
                        End If
                    Else
                        If rsSO106("AccountID").OriginalValue & "" <> strRealAccNo Then
                            gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                            ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                            ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                            ",CMCODE=" & gilCMCode.GetCodeNo & _
                                            ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                            ",PTCODE=" & gilPTCode.GetCodeNo & _
                                            ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                            IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, "") & _
                                            " WHERE CUSTID=" & gCustId & _
                                            " AND COMPCODE=" & gCompCode & _
                                            " AND ACCOUNTNO='" & rsSO106("AccountID").OriginalValue & "'" & _
                                            IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & "" & ")"), RetVal
                            Debug.Print RetVal
                            If strChildCustID <> Empty Then
                                gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                                "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                                ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                                ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                                ",CMCODE=" & gilCMCode.GetCodeNo & _
                                                ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                                ",PTCODE=" & gilPTCode.GetCodeNo & _
                                                ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                                IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, "") & _
                                                " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                                " AND COMPCODE=" & gCompCode & _
                                                " AND ACCOUNTNO='" & rsSO106("AccountID").OriginalValue & "'" & _
                                                IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & "" & ")"), RetVal
                                Debug.Print RetVal
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        If csmCitem2.GetQueryCode <> Empty Then
            gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                            "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                            ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                            ",ACCOUNTNO='" & strRealAccNo & "'" & _
                            ",CMCODE=" & gilCMCode.GetCodeNo & _
                            ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                            ",PTCODE=" & gilPTCode.GetCodeNo & _
                            ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                            IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, "") & _
                            " WHERE CUSTID=" & gCustId & _
                            " AND COMPCODE=" & gCompCode & _
                             IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & ")") & _
                             IIf(blnUCCodeIsNull, "", " AND UCCODE > 0 ") & " AND CANCELFLAG=0", RetVal
                             Debug.Print RetVal
            If strChildCustID <> Empty Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO033 SET " & _
                                "BANKCODE=" & gilBankCode.GetCodeNo2 & _
                                ",BANKNAME='" & gilBankCode.GetDescription & "'" & _
                                ",ACCOUNTNO='" & strRealAccNo & "'" & _
                                ",CMCODE=" & gilCMCode.GetCodeNo & _
                                ",CMNAME='" & gilCMCode.GetDescription & "'" & _
                                ",PTCODE=" & gilPTCode.GetCodeNo & _
                                ",PTNAME='" & gilPTCode.GetDescription & "'" & _
                                IIf(strInvSeqNo <> "", ",InvSeqNo=" & strInvSeqNo, "") & _
                                " WHERE CUSTID IN (" & strChildCustID & ")" & _
                                " AND COMPCODE=" & gCompCode & _
                                 IIf(strRowId = Empty, "", " AND ROWID IN (" & strRowId & ")") & _
                                IIf(blnUCCodeIsNull, "", " AND UCCODE > 0 ") & " AND CANCELFLAG=0", RetVal
                                 Debug.Print RetVal
            End If
        End If
    End With
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ChkSO033"
End Sub

Private Sub StopSO002A(Optional strAccNo As String = "")
  On Error Resume Next
    Dim strChkSameAcc As String
    Dim strStopDate As String
    If strAccNo = Empty Then strAccNo = strRealAccNo
    If strAccNo = Empty Then Exit Sub
    If gdaStopDate.GetValue = Empty Then gdaStopDate.SetValue RightDate
    
    '#3542 �s�W�ɷ|�۰ʶ�J���Τ���A�N�P�_���ܤW���A�å��ܼưO���A�ϥ�Update�y�k�ɡA���ܼƱa�J By Kin 2007/09/29
    strStopDate = gdaStopDate.GetValue
    If chkStop.Value <> 1 Then gdaStopDate.SetValue ""
    
    '********************************************************************************************************
    '#3324 �ˬd�O�_���ۦP�b���Ω��L�A�Ȥ����O��Ƥ��A�p�G���䥦��ơA������SO002A����� By Kin 2007/08/01
    If Not rsSO106.EOF Then
        strChkSameAcc = "Select Count(*) Cnt From " & GetOwner & "SO106 Where " & _
                        "AccountID='" & strAccNo & "' And " & _
                        "CompCode=" & gCompCode & " And " & _
                        "CUSTID=" & gCustId & " And " & _
                        "StopFlag=0 And StopDate Is Null" & _
                        IIf(lngEditMode = giEditModeInsert, "", " And RowID<>'" & rsSO106("RowId") & "'")
        If gcnGi.Execute(strChkSameAcc)(0) > 0 Then Exit Sub
    End If
    '*********************************************************************************************************
'   gcnGi.Execute ("delete from " & GetOwner & "so002a where AccountNo='" & strAccNo & "' and custid=" & gCustId & " and compcode=" & gCompCode)
    gcnGi.Execute ("UPDATE " & GetOwner & "SO002A" & _
                                " SET STOPFLAG=1,STOPDATE=TO_DATE('" & strStopDate & "','YYYY/MM/DD')" & _
                                " WHERE ACCOUNTNO='" & strAccNo & "'" & _
                                " AND CUSTID=" & gCustId & _
                                " AND COMPCODE=" & gCompCode)
    gimCitem.Clear
'    If chkStop.Value <> 1 Then gdaStopDate.SetValue ""
    
    ' ���b��p�w���ήɡA�h�l�b��ݳs�ʰ��γB�z�ñN���l���p�M��
    ' (�YInheritFlag=0, .InheritKey=Null)�C
    ' (���l���p����: �l�b��.InheritFlag=1 and �l�b��.InheritKey=���b��.InheritKey)
    If rsSO106("InheritKey") & "" <> Empty And rsSO106("InheritFlag") = 0 Then
    
        gcnGi.Execute "UPDATE " & GetOwner & "SO106" & _
                                " SET INHERITFLAG=0,INHERITKEY=NULL" & _
                                " WHERE ACCOUNTNO='" & strAccNo & "'" & _
                                " AND INHERITKEY='" & rsSO106("InheritKey") & "'" & _
                                " AND INHERITFLAG=1" & _
                                " AND COMPCODE=" & gCompCode
                                
        Dim strChild2A As String
        strChild2A = Get2AChild
        If strChild2A <> Empty Then
            gcnGi.Execute "UPDATE " & GetOwner & "SO002A" & _
                                    " SET INHERITFLAG=0,INHERITKEY=NULL" & _
                                    " WHERE INHERITKEY IN (" & strChild2A & ")" & _
                                    " AND INHERITFLAG=1" & _
                                    " AND COMPCODE=" & gCompCode
        End If
    End If
End Sub

Private Function Get2AChild() As String
  On Error Resume Next
    If rsSO106("INHERITKEY") & "" <> Empty And rsSO106("INHERITFLAG") = 0 Then
        Get2AChild = gcnGi.Execute("SELECT INHERITKEY FROM " & GetOwner & "SO002A" & _
                                                        " WHERE INHERITFLAG=1" & _
                                                        " AND CUSTID=" & gCustId & _
                                                        " AND COMPCODE=" & gCompCode & _
                                                        " AND ACCOUNTNO='" & rsSO106("AccountID") & "'").GetString(2, , "", "','", "")
        If Err Then
            Err.Clear
            Get2AChild = ""
        Else
            If Right(Get2AChild, 2) = ",'" Then Get2AChild = Left(Get2AChild, Len(Get2AChild) - 2)
            If Left(Get2AChild, 2) = "'," Then Get2AChild = Mid(Get2AChild, 3)
            If Left(Get2AChild, 1) <> "'" Then Get2AChild = "'" & Get2AChild
        End If
    End If
End Function

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
    Set ObjActiveForm = Me
    strActFrmName = "frmSO1100F"
    If frmSO114FA.frmShow Then
        frmSO114FA.SetFocus
    End If
    'strInvSeqNo = frmSO114FA.uInvSeqNo
    Screen.MousePointer = vbDefault
    Set ObjActiveForm = Me
    strActFrmName = "frmSO1100G"
    If rsSO106.RecordCount <> frmSO1100BMDI.rsSO106.RecordCount Then OpenData
    
End Sub

Private Sub Form_KeyPress(keyAscii As Integer)
  On Error Resume Next
   If keyAscii = Asc("'") Then keyAscii = 0
End Sub
Public Property Let Set106Inv(ByVal vData As String)
    strInvSeqNo = vData
    lblInvSeqNo = strInvSeqNo
End Property
Private Sub Form_Load()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    Dim lngAbsPos As Long
    Dim blnDefCustStatusFaci As Boolean
    posAchContition = " PRGNAME = 'X'  "
    strInvSeqNo = Empty
    InitializingListOcx
    blnStartPost = False
    If Crm Then
        PolyFrmFunction Me
        Me.Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2)
    Else
        'If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 + 300
        If Not 800600 Then Me.Move (Screen.Width - Me.Width) / 2, frmSO1100BMDI.Top + frmSO1100BMDI.tbrMenu.Height + 700
    End If
    'Me.Move (Screen.Width - Me.Width) / 2, frmSO1100BMDI.Top + frmSO1100BMDI.tbrMenu.Height + 600
    If EditMode <> giEditModeInsert Then
        lngAbsPos = frmSO1100BMDI.rsSO106.AbsolutePosition
        OpenData
        If lngAbsPos > 0 Then frmSO1100BMDI.rsSO106.AbsolutePosition = lngAbsPos
    End If
    
    With mFlds
        .Add "AcceptName", , , , , "���z�H�� "
        .Add "AcceptTime", giControlTypeDate, , , , "���z�ɶ� "
        .Add "Proposer", , , , , "�ӽФH   "
        .Add "ID", , , , , "�����Ҧr��"
        .Add "BankName", , , , , "�Ȧ�W��" & Space(18)
        .Add "CardName", , , , , "�H�Υd�O"
        .Add "StopYM", , , "00/0000", , "�d���Ĵ�"
        .Add "AccountID", , , , , "�b��/�d��" & Space(12)
        .Add "AccountName", , , , , "�b���Ҧ��H"
        .Add "AccountNameID", , , , , "�b���Ҧ��H��������"
        .Add "PropDate", giControlTypeDate, , , , "�ӽФ��"
        .Add "SendDate", giControlTypeDate, , , , "�e����"
        .Add "SnactionDate", giControlTypeDate, , , , "�֭���"
        .Add "StopFlag", , , , , "����"
        .Add "StopDate", giControlTypeDate, , , , "���Τ��"
        .Add "MediaName", , , , , "���дC��"
        .Add "IntroName", , , , , "���ФH"
        .Add "Note", , , , , "�Ƶ�"

'        .Add "AccountID", , , , , "�H�Υd�O"
'        .Add "CardExpDate", , , , , "�H�Υd����"
    End With
    With ggrData
        .UseCellForeColor = True
        .AllFields = mFlds
        .SetHead
         Set .Recordset = rsSO106
        .Refresh
    End With
    '**********************************************************************************************
    '#4089 �W�[�P�_�ѼơA�p�G��0�n�A�P�_�O�_�i���ACH���v���ӡA�p�G��1�h�n�j������ By Kin 2008/09/11
    blnDefCustStatusFaci = Val(GetRsValue("Select Nvl(DefCustStatusFaci,0) From " & GetOwner & _
                            "SO041 Where SysID=" & gCompCode, gcnGi)) - 1
    '**********************************************************************************************
    
    '#3585 �P�_ACHCustID�O�_��1,�p�G��1,�s�W��ƮɡA�n�g�JAch�Τḹ�X By Kin 2007/10/26
    'If blnDefCustStatusFaci Then
        blnACHCustID = Val(GetRsValue("Select Nvl(ACHCustID,0) From " & GetOwner & _
                                "SO041 Where SysID=" & gCompCode, gcnGi))
    'Else
    '    blnACHCustID = False
    'End If
    '#3925 SO041.ACHCUSTID=0 ACH���{���s�nDisable By Kin 2008/05/08
    If Val(GetRsValue("Select Nvl(StartPost,0) From " & GetOwner & "SO041 Where SysID = " & gCompCode, gcnGi) & "") = 1 Then
        posAchContition = " PRGNAME like '%POST%' "
        AchTypeCondition = "  In (1, 2 ) "
        blnStartPost = True
    Else
        posAchContition = " PRGNAME = 'X'  "
        AchTypeCondition = " In (1) "
    End If
    cmdACHDetail.Enabled = blnACHCustID And blnDefCustStatusFaci
    cmdACHDetail.Visible = blnACHCustID And blnDefCustStatusFaci
    If Not blnACHCustID Then
        lblState.Move lblApplication.Left
        lblStatus.Move lblApplication.Left + lblState.Width + 50
    End If
    '#5045 �W�[�ѼƧP�_�O�_�i�D��w���O����� By Kin 2009/05/05
    blnUCCodeIsNull = GetRsValue("Select Nvl(HItemAutoUpd,0) From " & GetOwner & "SO041", gcnGi) = 1
    '#3454 �NSO001.CustId�PSO994.DeclarantName �ŤJComboBox By Kin 2007/12/21
    Call ShowProposer
    Call ChangeMode(giEditModeView)
'    �Ӹ�k
'    If Len(strAccountName) > 0 Then
'        If strAccountName <> ggrData.Recordset.Fields("Proposer") Then
'            cboProposer.Text = "XXXXX"
'            gilCMCode.SetCodeNo "X"
'            gilCMCode.SetDescription "XXX"
'            gilPTCode.SetCodeNo "X"
'            gilPTCode.SetDescription "XXX"
'            txtAccountNo.PasswordChar = "X"
'            txtAccountOwner.PasswordChar = "X"
'            txtHide.PasswordChar = "X"
'        End If
'    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub OpenData()
  On Error GoTo ChkErr
    Set rsSO106 = New ADODB.Recordset
    With rsSO106
         If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open "SELECT ROWID,SO106.* FROM " & GetOwner & "SO106 SO106 WHERE CUSTID= " & gCustId & " AND COMPCODE=" & gCompCode & " ORDER BY ACCEPTTIME DESC", gcnGi, adOpenKeyset, adLockOptimistic
    End With
    With ggrData
         If .Recordset.State = 1 Then
             Set .Recordset = rsSO106
            .Refresh
         End If
    End With
  Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub

Public Sub UserPermissionGo()
  On Error GoTo ChkErr
    Dim blnFlag As Boolean
    If ChkSo1100BMDI Then
        If frmSO1100BMDI.rsSO106.State = adStateOpen Then blnFlag = Not frmSO1100BMDI.rsSO106.EOF
    End If
    ' �ھ��v���էO, �������ާ@�v��
    MenuEnabled GetUserPriv("SO1100G", "(SO1100G1)"), _
                        GetUserPriv("SO1100G", "(SO1100G2)") And blnFlag, _
                        GetUserPriv("SO1100G", "(SO1100G3)") And blnFlag, _
                        blnFlag And EditMode <> giEditModeView, _
                        GetUserPriv("SO1100G", "(SO1100G5)") And blnFlag And False
'    MenuEnabled GetUserPriv("SO1100G", "(SO1100G1)"), _
                        GetUserPriv("SO1100G", "(SO1100G2)") And blnflag, _
                        GetUserPriv("SO1100G", "(SO1100G3)") And blnflag, _
                        GetUserPriv("SO1100G", "(SO1100G4)") And blnflag And EditMode <> giEditModeView, _
                        GetUserPriv("SO1100G", "(SO1100G5)") And blnflag And False
  Exit Sub
ChkErr:
    ErrSub Me.Name, "UserPermissionGo"
End Sub

Private Sub InitializingListOcx()
  On Error GoTo ChkErr
    gilCardName.ListType = OneColumn
    SetLst gilCardName, "CodeNo", "Description", 3, 12, 405, 1620, "CD037"
    SetLst gilBankCode, "CodeNo", "Description", 3, 24, 405, 1850, "CD018", , True
    SetLst gilAcceptName, "EmpNo", "EmpName", 10, 20, , , "CM003"
    SetLst gilMediaCode, "CodeNo", "Description", 3, 12, 405, 1620, "CD009"
    SetLst gilCMCode, "CodeNo", "Description", 3, 12, 405, 1620, "CD031"
    Call SetgiList(gilPTCode, "CodeNo", "Description", "CD032", , , , , 10, 20, "Where Nvl(ServiceType,'" & gServiceType & "' ) ='" & gServiceType & "'", True)
    
'    On Error Resume Next
'    strFilter = gcnGi.Execute("SELECT CITEMSTR FROM SO106 WHERE CUSTID=" & gCustId & " AND STOPFLAG <> 1 AND CITEMSTR IS NOT NULL").GetString(adClipString, , "", ",", "")
'    If Err.Number = 3021 Then
'        Err.Clear
'    Else
'        strFilter = RPxx(strFilter)
'        strFilter = Left(strFilter, Len(strFilter) - 1)
'    End If
    SetMS gimCitem, "�N�X", "�W��", "CodeNo", "Description", "CD019", 1, " WHERE PERIODFLAG <> 0 AND ( SERVICETYPE ='" & gServiceType & "' OR SERVICETYPE IS NULL)" '& IIf(strFilter = Empty, "", " AND CODENO NOT IN (" & strFilter & ")")
    gilAcceptName.Filter = " Where CompCode =" & gCompCode & "'"
    gilBankCode.Filter = " Where CompCode =" & gCompCode
    gilMediaCode.Filter = " Where ( ServiceType ='" & gServiceType & "' OR ServiceType IS Null )"
    gilCMCode.Filter = " Where ( ServiceType ='" & gServiceType & "' OR ServiceType IS Null )"
  
'    SetMsQry csmCitem1, "SELECT SEQNO,CITEMCODE,CITEMNAME,PERIOD,AMOUNT,ACCOUNTNO,CMNAME,STARTDATE,STOPDATE FROM " & _
                         GetOwner & "SO003 SO003 WHERE CUSTID = " & _
                         gCustId & " AND COMPCODE=" & gCompCode & _
                         " ORDER BY CLCTDATE,CITEMCODE", _
                         "�Ǹ�", "�N�X", , , True, "���O����", "����", "���B", "���b�b��", "���O�覡", _
                         "�_�l��", "�I���", , , , 1, 560, 2000, 520, 640, 2000, 900, 900, 900, , , , , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD"
    
'    SetMsQry csmCitem1, "SELECT CITEMCODE,CITEMNAME,PERIOD,AMOUNT,ACCOUNTNO,CMNAME,STARTDATE,STOPDATE,SEQNO FROM " & _
'                         GetOwner & "SO003 SO003 WHERE CUSTID = " & _
'                         gCustId & " AND COMPCODE=" & gCompCode & _
'                         " ORDER BY CLCTDATE,CITEMCODE", _
'                         "�N�X", "���O����", "", , True, "����", "���B", "���b�b��", "���O�覡", _
'                         "�_�l��", "�I���", "�Ǹ�", , , , 500, 1800, 460, 560, 1800, 850, 850, 850, 770, , , , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD", , , , , 9
    
'    SetMsQry csmCitem2, "SELECT DISTINCT CITEMCODE,CITEMNAME,REALPERIOD,SHOULDAMT,ACCOUNTNO,CMNAME,REALSTARTDATE,REALSTOPDATE,SEQNO,PK,ROWID FROM " & _
'                         "(SELECT CITEMCODE,CITEMNAME,REALPERIOD,SHOULDAMT,ACCOUNTNO,CMNAME,REALSTARTDATE,REALSTOPDATE,SEQNO,BILLNO || ITEM PK,ROWID FROM " & _
'                         GetOwner & "SO033 WHERE CUSTID = " & gCustId & _
'                         " AND COMPCODE=" & gCompCode & _
'                         " AND CANCELFLAG <> 1 AND UCCODE IS NOT NULL AND CHEVEN <> 1 " & _
'                         "ORDER BY REALDATE DESC,BILLNO DESC,SEQNO DESC)", _
'                         "�N�X", "���O����", "", , True, "����", "���B", "���b�b��", "���O�覡", _
'                         "�_�l��", "�I���", "�Ǹ�", "PK", , , 500, 1800, 460, 560, 1800, 850, 850, 850, 770, 1, 1, , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD", , , , , 10

    '#3454 �h��SO003.FaciSNo�BSO004.DeclarantName�BSO004.FaciName By Kin 2007/12/10
    SetMsQry csmCitem1, "SELECT A.CITEMCODE,A.CITEMNAME,A.PERIOD,A.AMOUNT,A.ACCOUNTNO,A.CMNAME,A.STARTDATE," & _
                        "A.STOPDATE,A.SEQNO,A.FaciSNo,B.DeclarantName,B.FaciName FROM " & _
                         GetOwner & "SO003 A," & GetOwner & "SO004 B  WHERE A.CUSTID = " & _
                         gCustId & " AND A.COMPCODE=" & gCompCode & _
                         " AND A.CustId=B.CustId(+) And A.CompCode=B.CompCode(+)" & _
                         " And A.FaciSeqNo=B.SeqNo(+)" & _
                         " ORDER BY A.CLCTDATE,A.CITEMCODE", _
                         , , "�N�X,���O����,����,���B,���b�b��,���O�覡,�_�l��,�I���,�Ǹ�,���~�Ǹ�,�ӽФH�m�W,���ئW��", _
                         "500,1800,460,560,1800,850,850,850,770,850,1600,1600", True, , , , , _
                         , , , , , , , , , , , , , , , , , , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD", , , , , 12

    SetMsQry csmCitem2, "SELECT DISTINCT CITEMCODE,CITEMNAME,REALPERIOD,SHOULDAMT,ACCOUNTNO,CMNAME,REALSTARTDATE," & _
                        "REALSTOPDATE,SEQNO,PK,RowID,FaciSNo,DeclarantName,FaciName FROM " & _
                         "(SELECT A.CITEMCODE,A.CITEMNAME,A.REALPERIOD,A.SHOULDAMT,A.ACCOUNTNO,A.CMNAME,A.REALSTARTDATE," & _
                            "A.REALSTOPDATE,A.SEQNO,A.BILLNO || A.ITEM PK,A.ROWID,A.FaciSNo,B.DeclarantName,B.FaciName FROM " & _
                         GetOwner & "SO033 A," & GetOwner & "SO004 B WHERE A.CUSTID = " & gCustId & _
                         " AND A.COMPCODE=" & gCompCode & _
                         " AND A.CANCELFLAG <> 1 AND A.UCCODE IS NOT NULL AND A.CHEVEN <> 1 " & _
                         " AND A.CustId=B.CustId(+) And A.CompCode=B.CompCode(+)" & _
                         " And A.FaciSeqNo=B.SeqNo(+) " & _
                         "ORDER BY A.REALDATE DESC,A.BILLNO DESC,A.SEQNO DESC)", _
                         , , "�N�X,���O����,����,���B,���b�b��,���O�覡,�_�l��,�I���,�Ǹ�,PK,RowID,���~�Ǹ�,�ӽФH�m�W,���ئW��", _
                         "500,1800,460,560,1800,850,850,850,770,1,1,850,1600,1600", True, , , , , _
                             , , , , , , , , , , , , , , , , , , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD", , , , , 14
    '#3236 ACHTNO�N�X�S���ߤ@,��ܥX�ӳ��|�ܦ��Ĥ@�ӥN�X,SO106�s�WACHTDESC���,SetQueryCode���ACHDESC��� By Kin 2007/05/28
    '#7049 �W�[BillHeadFmt���,�i�H���CD068A By Kin 2015/07/08
    SetMsQry csmACH, "SELECT ACHTNO,ACHTDESC,CITEMCODESTR,BillHeadFmt FROM " & GetOwner & "CD068 Where ACHTNO Is Not Null And ACHTDESC is Not NULL And ACHType IN (1,2) ", _
                            "ACH����N�X", "ACH�������", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
                                
    SetMsQry csmTest, "select codeno,description,amount,accidp1 from cd019 where rownum<=10", _
                            "codeno", "description", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
    SetMsQry csmTest, "select codeno,description,amount,accidp1 from cd019 where codeno = 305", _
                            "codeno", "description", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
    csmTest.SetQueryCode "'(�N��)CM�s�J�O�Ҫ���X'"
'    SetMsQry csmACH, "SELECT ACHTNO,ACHTDESC,CITEMCODESTR,BillHeadFmt FROM " & GetOwner & "CD068", _
'                                    "ACH����N�X", "ACH�������", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
  Exit Sub
ChkErr:
    ErrSub Me.Name, "InitializingListOcx"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift)
  Exit Sub
ChkErr:
       Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Public Sub AddNewGo()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
    MenuEnabled , , , True, , True
    frmSO1100BMDI.Pic1.Enabled = False
    EditMode = giEditModeInsert
    OpenData
    Me.Show , frmSO1100BMDI
    ChangeMode giEditModeInsert
    lblStatus = ""
    NewRcd
    lblState.Visible = False
    lblStatus.Visible = False
    gdaSendDate.Enabled = True
    gdaAuthStopDate.Enabled = True
    gdaSnactionDate.Enabled = True
    gilBankCode.Enabled = True
    txtAccountNo.Enabled = True
    txtAccOwnerID.Enabled = True
    lblAccountOwner = "�b���Ҧ��H"
    txtHide.Visible = False
    txtAccountNo.Enabled = True
    gilCMCode.Enabled = True
    mskCardExpDate.Visible = True
    mskCardExpDate.Enabled = True
'    lblAccountOwner.ForeColor = vbBlack
    gdtAcceptTime.Text = RightNow
    gdaContiDate.Text = RightDate
    gilAcceptName.SetCodeNo garyGi(0)
    gilAcceptName.SetDescription garyGi(1)
    '#3454 �ӽФH���ComboBox By Kin 2007/12/18
    cboProposer.Text = ""
    Screen.MousePointer = vbDefault
    On Error Resume Next
    'txtProposer.SetFocus
    cboProposer.SetFocus
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "AddNewGo")
End Sub
 
Public Sub EditGo()
  On Error GoTo ChkErr
    '#4094 �W�[�ѼƧP�_ACH�Ȧ�ɬO�_��ϥΰ��� By Kin 2009/09/19
    blnEnableStop = Val(GetRsValue("Select Nvl(ACHCancelStop,0) From " & GetOwner & _
                              "SO041 Where SysID=" & gCompCode, gcnGi))
    '**************************************************************************
    '#3728 ���ή�,�Ȧ�O�OACH�ɤ����\��s��� By Kin 2008/03/18
    If Screen.ActiveForm.Name <> "frmSO1100G" Then
        If Not blnEnableStop Then
                If Val(frmSO1100BMDI.rsSO106("StopFlag")) = 1 Then
                    If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018" & _
                                            " WHERE CODENO=" & frmSO1100BMDI.rsSO106("BankCode") & _
                                            " AND ( PRGNAME LIKE 'ACH%' Or " & posAchContition & " ) " & _
                                            " AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
                                            
                        MsgBox "�ӵ���Ƥw���α��v�A�����\�ק�I", vbInformation, "�T��"
                        Exit Sub
                    End If
                End If
        End If
    Else
        If chkStop.Value Then
            '#4094 �p�GACHCancelStop=1 �n��User�i�H�i�J�ק� By Kin 2008/09/19
            If Not blnEnableStop Then
                If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018" & _
                                        " WHERE CODENO=" & gilBankCode.GetCodeNo & _
                                        " AND ( PRGNAME LIKE 'ACH%' Or " & posAchContition & " ) " & _
                                        " AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
                                        
                    MsgBox "�ӵ���Ƥw���α��v�A�����\�ק�I", vbInformation, "�T��"
                    Exit Sub
                End If
            End If
        End If
    End If
    '**************************************************************************
    Screen.MousePointer = vbHourglass
    MenuEnabled , , , True, , True
    frmSO1100BMDI.Pic1.Enabled = False
    Me.Show , frmSO1100BMDI
    If Screen.ActiveForm.Name <> "frmSO1100G" Then OpenData
    ' �l�b�ᤣ�వ���ʡA�u�����b��~�వ���ʡA�Ӥl�b��u���s�ʳB�z�C
    '�]�YInheritFlag=0����Ƥ�i�@���ʡAInheritFlag=1����ƱN�QReadOnly�^
    If Val(rsSO106("InheritFlag") & "") = 1 Then
        ViewGo
        MsgBox "�l�b���Ƥ����W������ !", vbInformation, "�T��"
        Exit Sub
    End If
    AbsPos
    ChangeMode giEditModeEdit
    PrcACH
    Screen.MousePointer = vbDefault
    On Error Resume Next
    'txtProposer.SetFocus
    cboProposer.SetFocus
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "EditGo")
End Sub

Private Sub AbsPos()
  On Error GoTo ChkErr
    With frmSO1100BMDI.rsSO106
         If .RecordCount > 0 Then
            If Not .EOF Or Not .BOF Then
                If rsSO106.AbsolutePosition <> .AbsolutePosition Then rsSO106.AbsolutePosition = .AbsolutePosition
                RcdToScr
            End If
         End If
    End With
    
  Exit Sub
ChkErr:
    ErrSub Me.Name, "AbsPos"
End Sub
 
Public Sub ViewGo()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
    MenuEnabled , , , True, , True
    frmSO1100BMDI.Pic1.Enabled = False
    Me.Show , frmSO1100BMDI
    If Not Me.Visible Then OpenData
    ChangeMode giEditModeView
    AbsPos
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "ViewGo")
End Sub

Public Sub CancelGo()
   On Error GoTo ChkErr
    gilBankCode_Change
    blnCallSO106A = False
    If EditMode = giEditModeView Then
        Unload Me
        Exit Sub
   Else
      If EditMode = giEditModeEdit Then
         Call ChangeMode(giEditModeView)
         RevertRcd
      Else
         If Not rsSO106.EOF Then
            rsSO106.MoveFirst
            RcdToScr
         Else
            NewRcd
         End If
         Call ChangeMode(giEditModeView)
      End If
'      OpenData
   End If
 Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "CancelGo")
End Sub

Private Sub RevertRcd() '�٭��Ƥ��e
 On Error GoTo ChkErr
   RcdToScr
 Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "RevertRcd")
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
        On Error Resume Next
        rsSO106.Close
        Set rsSO106 = Nothing
        DoEvents
        FormQueryUnload
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_QueryUnload"
End Sub

Private Sub gdaSendDate_Change()
  On Error GoTo ChkErr
     If InStr(gdaSendDate.GetOriginalValue, "_") = 0 Then
        If gilBankCode.GetCodeNo <> Empty Then
            If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018  " & _
                            " WHERE CODENO=" & gilBankCode.GetCodeNo & _
                            " AND ( PRGNAME LIKE 'ACH%' Or  " & posAchContition & " )  AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
                If rsSO106.RecordCount > 0 Then
                    If rsSO106("AuthorizeStatus") & "" = 1 Or rsSO106("AuthorizeStatus") & "" = Empty Then
                        gilBankCode.Enabled = False
                        txtAccountNo.Enabled = False
                        blnACHAccountNoCanEdit = False
                        txtAccOwnerID.Enabled = False
                    Else
                        CtlEnable
                    End If
                Else
                    CtlEnable
                End If
            Else
                CtlEnable
            End If
        Else
            CtlEnable
        End If
     Else
        CtlEnable
     End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "gdaSendDate_Change"
End Sub

Private Sub CtlEnable()
  On Error Resume Next
    gilBankCode.Enabled = True
    txtAccountNo.Enabled = True
    blnACHAccountNoCanEdit = True
    txtAccOwnerID.Enabled = True
End Sub

Private Sub gdaSnactionDate_GotFocus()
  On Error GoTo ChkErr
    '#5045 �ݦ۰ʱa�J�t�Τ�� By Kin 2009/05/05
    If gdaSnactionDate.GetValue & "" = "" Then
        gdaSnactionDate.SetValue RightNow
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaSnactionDate_GotFocus")
End Sub

Private Sub gdtAcceptTime_Validate(Cancel As Boolean)
  On Error Resume Next
    If Not IsDate(gdtAcceptTime.Text) And gdtAcceptTime.GetValue <> "" Then Cancel = True: Exit Sub
End Sub

Private Sub ggrData_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If UCase(Fld.Name) = "STOPFLAG" Then
        If Value = 1 Then Value = vbRed
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ggrData_ColorCellData"
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If UCase(Fld.Name) = "STOPFLAG" Then Value = IIf(Value = 0, "", "�O")
    If UCase(Fld.Name) = "STOPYM" Then
        If Not GetUserPriv("SO1100G", "(SO1100G7)") Then
            Value = "XX/XXXX"
        End If
    End If
    If UCase(Fld.Name) = "ACCOUNTID" Then
        If Not GetUserPriv("SO1100G", "(SO1100G9)") Then
            If Value = Empty Then
                Value = ""
            Else
                Value = Left(Value, 5) & "XXX" & Mid(Value, 9, 2) & "XX" & Right(Value, 4)
            End If
        End If
    End If
    If UCase(Fld.Name) = "ID" Or UCase(Fld.Name) = "ACCOUNTNAMEID" Then
        If Not GetUserPriv("SO1100G", "(SO1100G0)") Then
            If Value = Empty Then
                Value = ""
            Else
                Value = Left(Value, 3) & "XXXX" & Right(Value, 3)
            End If
        End If
    End If
    If Len(strAccountName) > 0 Then
        If strAccountName <> ggrData.Recordset.Fields("Proposer") Then
            Value = "XXXXX"
        End If
    End If
    
'    If Not GetUserPriv("SO1100G", "(SO1100G9)") Then
'        If UCase(Fld.Name) = "ACCOUNTID" Then Value = Left(Value, 5) & "XXX" & Mid(Value, 9, 2) & "XX" & Right(Value, 4)
'    End If
'    If Not GetUserPriv("SO1100G", "(SO1100G0)") Then
'        If UCase(Fld.Name) = "ID" Or UCase(Fld.Name) = "ACCOUNTNAMEID" Then Value = Left(Value, 2) & "XXXX" & Right(Value, 2)
'    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ggrData_ShowCellData"
End Sub

Private Sub gilBankCode_Change()
  On Error Resume Next
  
    Dim intActLen As Integer
    Dim intVNO As Integer
    If EditMode <> giEditModeView Then
        If Not IsNumeric(gilBankCode.GetCodeNo) Then Exit Sub
        '#4148 �h�[PrcACH Sub,�]���p�G����ܻȦ�,�Y���ACH�Ȧ檺��,ACH Enabled�|���� By Kin 2008/10/16
        If gilCMCode.GetCodeNo = Empty Then PrcACH: Exit Sub
        '#8724
        If intCD031RefNo = 2 Or intCD031RefNo = 4 Or intCD031RefNo = 5 Then
            blnVAcc = False
            If gilCardName.GetCodeNo <> Empty Then
                On Error Resume Next
                '*********************************************************************************
                '#3395 �p�G�O�ϥΫH�Υd�A�ϥΫH�Υd���ˮ��I By Kin 2007/08/20
                '#8724
                If intCD031RefNo = 4 Or intCD031RefNo = 5 Then
                    If Not ChkCreditCard(gilCardName.GetCodeNo, txtAccountNo, intActLen) Then
                        txtAccountNo.SetFocus
                        txtAccountNo.MaxLength = intActLen
                    End If
                Else
                    intActLen = RPxx(gcnGi.Execute("SELECT CARDNOLEN FROM " & GetOwner & "CD037 WHERE CODENO=" & gilCardName.GetCodeNo).GetString(adClipString, 1, "", "", ""))
                    If intActLen <> Empty Then txtAccountNo.MaxLength = intActLen
                End If
                '*********************************************************************************
            Else
                If Len(gilBankCode.GetCodeNo) > 0 Then
                    txtAccountNo.MaxLength = Val(gcnGi.Execute("SELECT ActLength FROM " & GetOwner & "CD018 WHERE CodeNo = " & gilBankCode.GetCodeNo).GetString)
                Else
                    txtAccountNo.MaxLength = 24
                End If
            End If
        Else
            '#3541���ܵ����b���W�h By Kin 2007/11/30
            Dim iCustCount As Long
            Dim strVNO As String
            If gilBankCode.GetCodeNo <> Empty Then
                intActLen = GetRsValue("SELECT ACTLENGTH FROM " & GetOwner & "CD018 WHERE CODENO=" & gilBankCode.GetCodeNo)
                txtAccountNo.MaxLength = intActLen
                
                iCustCount = gcnGi.Execute("SELECT COUNT(*) FROM " & GetOwner & _
                                          "SO002A WHERE CUSTID=" & gCustId & _
                                          " AND ID=2")(0)
                
                If iCustCount > 0 Then
                    strVNO = GetRsValue("Select Max(SubStr(AccountNO,1,8)) FROM " & GetOwner & _
                                        "SO002A WHERE CUSTID=" & gCustId & " AND ID=2")
                    strVNO = Format(CLng(strVNO) + 1, "00000000")
                Else
                    strVNO = Format(iCustCount + 1, "00000000")
                End If
                                          
                txtAccountNo = strVNO & String(intActLen - Len(CStr(strVNO)) - Len(CStr(gCustId)), "0") & gCustId
                blnVAcc = True
            End If
        End If
    End If
    '�Ӹ�k
    If Len(strAccountName) > 0 Then
        If strAccountName <> rsSO106.Fields("Proposer") Then
            Exit Sub
        End If
    End If
    PrcACH
End Sub

Private Sub PrcACH()
  On Error GoTo ChkErr
    '#3728 �p�G�Ȧ�O��ACH�hACH���v����O���i��,�Ϥ������i�� By Kin 2008/03/07
    '#7537 Add POST
    If gilBankCode.GetCodeNo <> Empty Then
        If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD018 " & _
                " WHERE CODENO=" & gilBankCode.GetCodeNo & " AND ( PRGNAME LIKE 'ACH%'  Or " & posAchContition & " )  " & _
                " AND COMPCODE=" & gCompCode, gcnGi, "") > 0 Then
            If blnStartPost Then
                If GetRsValue("Select count(*) from cd018 Where CD018.PRGNAME like '%POST%' and CodeNo = " & gilBankCode.GetCodeNo, gcnGi) = 0 Then
                    AchTypeCondition = " In ( 1 )"
                Else
                    AchTypeCondition = " In ( 2 )"
                End If
            
                Dim qrySelected As String
                qrySelected = Empty
                If Len(csmACH.GetQueryCode & "") > 0 Then
                    qrySelected = csmACH.GetQueryCode
                End If
                    
                SetMsQry csmACH, "SELECT ACHTNO,ACHTDESC,CITEMCODESTR,BillHeadFmt FROM " & GetOwner & "CD068  " & _
                                            " Where  ACHTNO is Not Null And ACHTDESC is Not NULL And ACHType " & AchTypeCondition, _
                                        "ACH����N�X", "ACH�������", , , , "HideColumn", "HideColumn", , , , , , , , , 2000, 2400, 1, 1, , , , , , , , , , , , , , , , , , , , , 2
                If Len(qrySelected & "") > 0 Then
                    csmACH.SetQueryCode qrySelected
                End If
            End If
            
            lblState.Visible = True
            lblStatus.Visible = True
            FraACH.Enabled = True
            'cmdACHDetail.Enabled = True
            '#3925 SO041.ACHCUSTID=0�� ACH���{���ӫ��s�nDisable By Kin 2008/05/08
            cmdACHDetail.Enabled = blnACHCustID
            If EditMode <> giEditModeView Then
                gdaSendDate.Enabled = False
                gdaSnactionDate.Enabled = False
            End If
            '#4094 �h�W�[�ѼƱ���O�_��ϥΰ��� By Kin 2008/09/19
'            chkStop.Enabled = blnEnableStop
'            gdaStopDate.Enabled = blnEnableStop
        Else
            csmACH.Clear
            FraACH.Enabled = False
            cmdACHDetail.Enabled = False
            lblState.Visible = False
            lblStatus.Visible = False
            '#4094 �p�G���OACH���Ȧ��٬O�n��ϥΰ��� By Kin 2008/09/19
'            chkStop.Enabled = True
'            gdaStopDate.Enabled = True
            If EditMode <> giEditModeView Then
                gdaSendDate.Enabled = True
                gdaSnactionDate.Enabled = True
                chkToACH.Value = 0
                csmACH.Clear
                'txtAchSN = ""   '#5045 ����Ҧ��U���n��s�� By Kin 2009/05/05
            End If
        End If
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "PrcACH"
End Sub

Private Sub gilCardName_Change()
  On Error GoTo ChkErr
    If EditMode <> giEditModeView Then
'        If ServiceCommon.AccountNo_Validate(Me) Then txtAccountNo.SetFocus
        gilBankCode_Change
        If intCD031RefNo = 2 Then
            ServiceCommon.AccountNo_Validate Me
        End If
'        If intCD031RefNo = 4 Then
'            ChkCreditCard gilCardName.GetCodeNo, txtAccountNo
'        End If
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "gilCardName_Change"
End Sub

Private Sub gilCMCode_Change()
  On Error Resume Next
    Dim strCD As Integer
    Dim rsCD032 As New ADODB.Recordset
    strCD = gilCMCode.GetCodeNo
    If strCD <> Empty Then
        intCD031RefNo = GetRsValue("SELECT REFNO  FROM " & GetOwner & "CD031 WHERE CODENO=" & strCD)
        If Err.Number <> 0 Then
            Err.Clear
            intCD031RefNo = 0
        End If
'        Debug.Print intCD; intCD031RefNo
        If intCD031RefNo = 2 Or intCD031RefNo = 4 Or intCD031RefNo = 5 Then
'            lblAccountOwner.ForeColor = vbRed
            lblAccountOwner = "�b���Ҧ��H"
'            If txtAccountOwner = Empty Then txtAccountOwner = frmSO1100BMDI.txtCustName
'            lblID.ForeColor = vbRed
            lblAccNameID.ForeColor = vbRed
            If EditMode <> giEditModeView Then
                If Len(txtAccountNo) = 16 Then
'                    If Not InStr(1, txtAccountNo, gCustId, vbTextCompare) Then
'                        txtAccountNo = ""
'                    End If
                Else
                    If Not InStr(1, txtAccountNo, gCustId, vbTextCompare) Then
                        txtAccountNo = ""
                    End If
                End If
            End If
            If intCD031RefNo = 4 Or intCD031RefNo = 5 Then
                '#5045 �p�G���O�覡�ѦҸ�=4,�I�ں����]�n�۰ʱa�X�ѦҸ���4����� By Kin 2009/05/05
                If Not GetRS(rsCD032, "Select * From " & GetOwner & "CD032 Where RefNO in (4,5) And " & _
                            " StopFlag<>1", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                If Not rsCD032.EOF Then
                    gilPTCode.SetCodeNo rsCD032("CodeNo") & ""
                    gilPTCode.Query_Description
                End If
                lblCardName.ForeColor = vbRed
                lblCardExpDate.ForeColor = vbRed
            Else
                lblCardName.ForeColor = vbBlack
                lblCardExpDate.ForeColor = vbBlack
            End If
        ElseIf intCD031RefNo = 0 Then
            If EditMode <> giEditModeView Then
'                lblID.ForeColor = vbBlack
                lblAccNameID.ForeColor = vbBlack
                lblAccountOwner = "�p���H"
'                lblAccountOwner.ForeColor = vbBlack
                If Len(txtAccountNo) = 16 Then
                    If Not InStr(1, txtAccountNo, gCustId, vbTextCompare) Then
                        txtAccountNo = ""
                    End If
                End If
                lblCardName.ForeColor = vbBlack
                lblCardExpDate.ForeColor = vbBlack
            End If
        Else
'            lblID.ForeColor = vbBlack
            lblAccNameID.ForeColor = vbBlack
            lblAccountOwner = "�p���H"
'            lblAccountOwner.ForeColor = vbBlack
            lblAccNameID.ForeColor = vbBlack
            lblCardName.ForeColor = vbBlack
            lblCardExpDate.ForeColor = vbBlack
            If gilBankCode.GetCodeNo <> Empty Then gilBankCode_Change
        End If
        '#5045 �����a�ӽФH�m�W By Kin 2009/05/05
        '#5121 2009.05.26 by Corey ��N#5045 ���ӬO�� �p�G�b���Ҧ��H�O�ťժ��ܤ~�ݭn�a�J�w�]�ȡA�p�G�����w�g���Ȫ��ܴN�H�������Ȭ���
        'If cboProposer.Text <> "" Then
        If cboProposer.Text <> "" And txtAccountOwner.Text = "" Then
            txtAccountOwner.Text = cboProposer.Text
        End If
    End If
    '#8724
    If intCD031RefNo <> 2 And intCD031RefNo <> 4 And intCD031RefNo <> 5 And txtAccountNo = Empty Then
        If gilBankCode.GetCodeNo <> Empty Then gilBankCode_Change
    End If
    '#8724
    If intCD031RefNo <> 4 And intCD031RefNo <> 5 Then
        If gilPTCode.GetCodeNo = Empty Then
            gilPTCode.SetCodeNo 1
            gilPTCode.Query_Description
        End If
    End If
    Call Close3Recordset(rsCD032)
End Sub

Private Sub gilMediaCode_Change()
  On Error GoTo ChkErr
    ServiceCommon.CommonMediaCode_Change Me
  Exit Sub
ChkErr:
    ErrSub Me.Name, "gilMediaCode_Change"
End Sub

Private Sub gilMediaCode_LostFocus()
  On Error GoTo ChkErr
    ServiceCommon.MediaCode_LostFocus Me
  Exit Sub
ChkErr:
    ErrSub Me.Name, "gilMediaCode_LostFocus"
End Sub

Private Sub mskCardExpDate_GotFocus()
  On Error Resume Next
    ObjGotFocus ImeClose
End Sub

Private Sub mskCardExpDate_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    Cancel = ServiceCommon.CardExpireDate(Me)
  Exit Sub
ChkErr:
    ErrSub Me.Name, "mskCardExpDate_Validate"
End Sub

Private Sub rsSO106_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  On Error GoTo ChkErr
    If adReason = 10 Then
        If Not rsSO106.EOF And Not rsSO106.BOF And lngEditMode = giEditModeView Then
            If Not ggrData.Enabled Then ggrData.Enabled = True
            RcdToScr
            On Error Resume Next
            frmSO1100BMDI.rsSO106.AbsolutePosition = rsSO106.AbsolutePosition
        Else
            ggrData.Enabled = False
        End If
        txtAccountNo.PasswordChar = Empty
        txtAccountOwner.PasswordChar = Empty
        txtHide.PasswordChar = Empty
        '�Ӹ�k
'        If Len(strAccountName) > 0 Then
'            If strAccountName <> rsSO106.Fields("cboProposer") Then
'                cboProposer.Text = "XXXXX"
'                gilCMCode.SetCodeNo "X"
'                gilCMCode.SetDescription "XXX"
'                gilPTCode.SetCodeNo "X"
'                gilPTCode.SetDescription "XXX"
'                txtAccountNo.PasswordChar = "X"
'                txtAccountOwner.PasswordChar = "X"
'                txtHide.PasswordChar = "X"
'            End If
'        End If
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "rsSO106_MoveComplete")
End Sub

Private Sub ReFilter()
  On Error Resume Next
    
    Dim strFilter As String
    strFilter = gcnGi.Execute("SELECT CITEMSTR FROM " & GetOwner & "SO106 WHERE CUSTID=" & gCustId & " AND STOPFLAG <> 1 AND CITEMSTR IS NOT NULL AND ROWID <> '" & rsSO106!RowId & "'").GetString(adClipString, , "", ",", "")
    If Err.Number = 3021 Then
        Err.Clear
    Else
        strFilter = RPxx(strFilter)
        strFilter = Left(strFilter, Len(strFilter) - 1)
    End If
    SetMS gimCitem, "�N�X", "�W��", "CodeNo", "Description", "CD019", 1, " WHERE PERIODFLAG <> 0 AND ( SERVICETYPE ='" & gServiceType & "' OR SERVICETYPE IS NULL)" & IIf(strFilter = Empty, "", " AND CODENO NOT IN (" & strFilter & ")")
End Sub
Public Property Let uAccountName(ByVal vData As String)
  On Error Resume Next
    strAccountName = vData
End Property
Public Property Let uInvSeqNo(ByVal vData As String)
  On Error GoTo ChkErr
    strInvSeqNo = vData
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Let uInvSeqNo")
End Property
Public Property Get EditMode() As giEditModeEnu
  On Error GoTo ChkErr
    EditMode = lngEditMode    ' ���ثe�s��Ҧ�
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Get EditMode")
End Property

Public Property Let EditMode(ByVal vNewValue As giEditModeEnu)
  On Error GoTo ChkErr
    lngEditMode = vNewValue '�]�w�s��Ҧ�
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Let EditMode")
End Property

'MID,PROMPT,LINK
'SO1100G , �H�Υd��b, SO1100
'SO1100G0 , �O�_������ܨ�������, SO1100G
'SO1100G1 , ��Ʒs�W, SO1100G
'SO1100G2 , ��ƭק�, SO1100G
'SO1100G3 , ��ƧR��, SO1100G
'SO1100G4 , ��Ƭd��, SO1100G
'SO1100G5 , ��ƦC�L, SO1100G
'SO1100G6 , �i�H�ק�H�Υd���Ĵ�, SO1100G
'SO1100G7 , �i�HŪ���H�Υd���Ĵ�, SO1100G
'SO1100G8 , ���iŪ���H�Υd���Ĵ�, SO1100G
'SO1100G9 , �O�_������ܱb��, SO1100G
'SO1100GA , �i�ק���b�b��, SO1100G
'SO1100GB , �i�ק怜�O�覡, SO1100G

'�ھڶǨӤ��s��Ҧ�, �]�w�U�����ݩ�
Public Sub ChangeMode(ByVal lngMode As giEditModeEnu)
  On Error GoTo ChkErr
    Dim blnFlag As Boolean  ' �O���O�_�b����s���Ҧ��A
    lngEditMode = lngMode
    Select Case lngMode
           Case giEditModeInsert
                blnFlag = False
                MenuEnabled , , False, True, False, True
           Case giEditModeEdit
                blnFlag = False
                MenuEnabled , , False, True, False, True
           Case giEditModeView
                blnFlag = True
                MenuEnabled True, rsSO106.RecordCount > 0, rsSO106.RecordCount > 0, , False, True
    End Select
    
    Select Case True
           Case (GetUserPriv("SO1100G", "(SO1100G6)"))
                 txtHide.Visible = False
                 mskCardExpDate.Visible = True
                 mskCardExpDate.Enabled = True
           Case (GetUserPriv("SO1100G", "(SO1100G7)"))
                 txtHide.Visible = False
                 mskCardExpDate.Visible = True
                 mskCardExpDate.Enabled = False
           Case (GetUserPriv("SO1100G", "(SO1100G8)"))
                 txtHide.Visible = True
                 mskCardExpDate.Visible = False
                 mskCardExpDate.Enabled = False
    End Select
    
    If lngMode <> giEditModeView Then
        gdaSendDate_Change
        txtAccountNo.Enabled = (GetUserPriv("SO1100G", "(SO1100GA)")) And blnACHAccountNoCanEdit
        gilCMCode.Enabled = (GetUserPriv("SO1100G", "(SO1100GB)"))
        ReFilter
    End If
    
    MenuEnabled GetUserPriv("SO1100G", "(SO1100G1)") And blnFlag, _
                GetUserPriv("SO1100G", "(SO1100G2)") And blnFlag, _
                GetUserPriv("SO1100G", "(SO1100G3)") And blnFlag, _
                Not blnFlag And EditMode <> giEditModeView, _
                GetUserPriv("SO1100G", "(SO1100G5)") And blnFlag And False, True
    '***************************************************************************
    '#3929 �p�G�ֲa����S��,�h�n�୫�s��ܵo���y���� By Kin 2008/06/09
    If (lngEditMode = giEditModeInsert) Or (lngEditMode = giEditModeEdit) Then
        If Not ChkInvData Then
            cmdInvSeqNo.Enabled = True
            cmdCreateINV.Enabled = True
        Else
            cmdInvSeqNo.Enabled = False
            cmdCreateINV.Enabled = False
            '**************************************************************
            '#3929 �n�h�A�P�_�@��,�]����Ʀ��i��h�� By Kin 2008/06/09
            If lngEditMode = giEditModeEdit Then
                If IsNull(rsSO106("SnactionDate")) Then
                    cmdInvSeqNo.Enabled = True
                    cmdCreateINV.Enabled = True
                End If
            End If
            '****************************************************************
        End If
    Else
        cmdInvSeqNo.Enabled = False
        cmdCreateINV.Enabled = False
    End If
    '***************************************************************************
    SetStatusBar lngEditMode
    fraData.Enabled = Not blnFlag
    ggrData.Enabled = blnFlag
    '#5556 �b�s���Ҧ��n�౲�ʱ��b By Kin 2010/03/01
    txtNote.Locked = blnFlag
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ChangeMode")
End Sub

Private Sub rsSO106_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  
  On Error Resume Next
    blnAchChgClick = False
End Sub

Private Sub txtAccountNo_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    Dim lngAccMax As Integer
    If gilCMCode.GetCodeNo <> Empty Then
        On Error Resume Next
        intCD031RefNo = GetRsValue("SELECT REFNO FROM " & GetOwner & "CD031 WHERE CODENO=" & gilCMCode.GetCodeNo)
        If Err.Number <> 0 Then
            Err.Clear
            intCD031RefNo = 0
        End If
    End If
    On Error GoTo ChkErr
    If EditMode <> giEditModeView Then
        '*************************************************************************************************************************
        '#3395 ���O�覡���H�Υd��(CD031.RefNo=4)�A�ϥΫH�Υd�ˮ֤覡�A���O�覡���Ȧ�ɳW�h����(CD031.RefNo=2) By Kin 2007/08/20
        If intCD031RefNo = 2 Then  'Or intCD031RefNo = 4 Then
            Cancel = ServiceCommon.AccountNo_Validate(Me)
        End If
        '#8724
        If intCD031RefNo = 4 Or intCD031RefNo = 5 Then
            Cancel = Not ChkCreditCard(gilCardName.GetCodeNo, txtAccountNo.Text, lngAccMax)
        End If
        '**************************************************************************************************************************
    End If
'    If EditMode = giEditModeInsert Then
'        If Not ChkInvData Then
'            cmdInvSeqNo.Enabled = True
'
'        Else
'
'        End If
'    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "txtAccountNo_Validate"
End Sub

Private Sub txtAccountOwner_Change()
  On Error Resume Next
    CML
End Sub

Private Sub txtAccountOwner_KeyPress(keyAscii As Integer)
  On Error Resume Next
    ML keyAscii
End Sub

Private Sub txtAccOwnerID_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If chkFore2.Value Then Exit Sub
    Select Case Len(txtAccOwnerID)
           Case 0
                Exit Sub
           Case 8
                If GetUserPriv("SO1100G", "(SO1100G0)") Then
                    InvNoIsOk txtAccOwnerID.Text, True
                Else
                    InvNoIsOk txtAccOwnerID.Tag, True
                End If
           Case 10
                If GetUserPriv("SO1100G", "(SO1100G0)") Then
                    IDIsOk txtAccOwnerID.Text, True
                Else
                    IDIsOk txtAccOwnerID.Tag, True
                End If
           Case Else
'                MsgBox "�����������Ӭ� 8 �� 10 �Ӧr��", vbInformation, "�T��"
                MsgBox "�нT�{�������� ! " & vbCrLf & _
                        "�K�X�� -- ���q���Τ@�s��" & vbCrLf & _
                        "�Q�X�� -- �ӤH�������Ҧr��", vbInformation, "�T��"
                Cancel = True
    End Select
  Exit Sub
ChkErr:
    ErrSub Me.Name, "txtAccOwnerID_Validate"
End Sub

Private Sub txtId_GotFocus()
  On Error Resume Next
    ObjGotFocus
End Sub

Private Sub txtID_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
'    Cancel = Not IDIsOk(txtID, True)
    If chkFore.Value Then Exit Sub
    Select Case Len(txtID)
           Case 0
                Exit Sub
           Case 8
                If GetUserPriv("SO1100G", "(SO1100G0)") Then
                    InvNoIsOk txtID.Text, True
                Else
                    InvNoIsOk txtID.Tag, True
                End If
           Case 10
                If GetUserPriv("SO1100G", "(SO1100G0)") Then
                    IDIsOk txtID.Text, True
                Else
                    IDIsOk txtID.Tag, True
                End If
           Case Else
'                MsgBox "�����������Ӭ� 8 �� 10 �Ӧr��", vbInformation, "�T��"
                MsgBox "�нT�{�������� ! " & vbCrLf & _
                        "�K�X�� -- ���q���Τ@�s��" & vbCrLf & _
                        "�Q�X�� -- �ӤH�������Ҧr��", vbInformation, "�T��"
                Cancel = True
    End Select
  Exit Sub
ChkErr:
    ErrSub Me.Name, "txtId_Validate"
End Sub

Private Sub txtIntroId_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = vbKeyF1 Then cmdFind.Value = True
  Exit Sub
ChkErr:
    ErrSub Me.Name, "txtIntroId_KeyDown"
End Sub

Private Sub txtIntroId_LostFocus()
  On Error Resume Next
    If Len(txtIntroId.Text) = 0 Then lblIntroName.Caption = ""
End Sub

Private Sub txtIntroID_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    Cancel = ServiceCommon.IntroId_Validate(Me)
  Exit Sub
ChkErr:
    ErrSub Me.Name, "txtIntroId_Validate"
End Sub

Private Sub txtNote_Change()
  On Error Resume Next
    CML
End Sub

Private Sub txtNote_KeyPress(keyAscii As Integer)
  On Error Resume Next
    ML keyAscii
End Sub

Private Sub txtProposer_Change()
  On Error Resume Next
    CML
End Sub

Private Sub txtProposer_KeyPress(keyAscii As Integer)
  On Error Resume Next
    ML keyAscii
End Sub

'#3454 �ӽФH���ComboBox,�NSO001.CustName�PSO004.Declarantname��J��ComboBox By Kin 2007/12/18
Private Sub ShowProposer()
  On Error GoTo ChkErr
    Dim rsProposer As New ADODB.Recordset
    Dim strQry As String
    strQry = "Select A.CustName From " & GetOwner & "SO001 A " & _
             " Where A.CustId=" & gCustId & _
             " And A.CompCode=" & gCompCode & _
             " Union All " & _
             "Select B.DeclarantName CustName From " & GetOwner & "SO004 B " & _
             " Where B.CustId=" & gCustId & _
             " And B.CompCode=" & gCompCode & _
             " And (B.PRDate Is Null OR B.InstDate > B.PRDate)"
    strQry = "Select Distinct(A.CustName) From (" & strQry & ") A"
    If Not GetRS(rsProposer, strQry, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Sub
    rsProposer.MoveFirst
    cboProposer.Clear
    Do While Not rsProposer.EOF
        If rsProposer("CustName") & "" <> "" Then
            cboProposer.AddItem rsProposer("CustName") & ""
        End If
        rsProposer.MoveNext
    Loop
    CloseRecordset rsProposer
    Exit Sub
ChkErr:
    ErrSub Me.Name, "ShowProposer"
End Sub


'#3728 �ݨD
'�NSO106��CitemStr�r���Ѷ}��,�è��o���Ǧ��O���جO�ݩ�Y�@��ACHTNO,�ño�����ӬO���ʡB�s�W���v�B�������v By Kin 2008/03/04
'�{�����I�p����,�j�P�W�O�ϥΦ�C���c�覡,�ɶ�������N����
'�^�Ǥ@��X*4���}�C
'(0,0) ACH�N�X
'(0,1) ACH�W��
'(0,2) ���ʧO 0:�s�W���v 1:�������v 2:���ʦ��O����
'(0,3) ���O���إN�X
'(0,4) ���O���ئW��
Private Sub GetCitemCode(ByVal strACHTNO As String, _
                        ByVal strACHTDESC As String, _
                        ByVal strCitemCode As String, _
                        ByVal strCitemName As String, _
                        ByVal blnAdd As Boolean, _
                        ByRef aryCitem() As String)
  On Error GoTo ChkErr
    '0:�N��n�s�W�@�����v,1:�N��n�s�W�@���������v 2:�N��nUpdate�쥻�����
    Dim aryACHTNO() As String
    Dim aryACHTDESC() As String
    Dim aryCitemCode() As String
    Dim aryCitemName() As String
    Dim rs As New ADODB.Recordset
    Dim i As Long, j As Long
    Dim lngACHTNO As Long
    Dim lngUbound As Long
    Dim strTmpACH As String
    '�s�W���
    If blnAdd Then
        If strACHTNO = Empty Then ReDim aryCitem(0, 4): GoTo Fin    '�p�G�S�����ACH,���i����
        aryACHTDESC = Split(Replace(strACHTDESC, "'", ""), ",")
        lngUbound = UBound(aryACHTDESC)
        ReDim aryCitem(lngUbound, 4)
        ReDim aryACHTNO(UBound(aryACHTDESC), 1)
        For i = LBound(aryACHTDESC) To UBound(aryACHTDESC)
            aryACHTNO(i, 0) = Split(Replace(strACHTNO, "'", ""), ",")(i)
            aryACHTNO(i, 1) = aryACHTDESC(i)
            aryCitem(i, 0) = aryACHTNO(i, 0)
            aryCitem(i, 1) = aryACHTNO(i, 1)
            aryCitem(i, 2) = "0"    '0�N��s�W
        Next i
        aryCitemCode = Split(Replace(strCitemCode, "'", ""), ",")
        aryCitemName = Split(Replace(strCitemName, "'", ""), ",")
        strTmpACH = strACHTNO
    '�s����
    Else
        With rsSO106
            Dim aryOldACHTDESC() As String  'SO106.CitemStr
            Dim aryOldACHTNO() As String
            Dim strQry1 As String
            Dim strQry2 As String
            Dim blnDiff As Boolean
            Dim lngAry As Long
            Dim strOldCitemCode As String
            Dim strOldCitemName As String
            Dim blnUpdCitem As Boolean
            Dim aryUpdCitemCode() As String
            Dim aryUpdCitemName() As String
            Dim blnAddNew As Boolean
            '**************************************************************************
            '�s��Ҧ��p�G��SO106�S��ACH�ӵe���]�S��ACH �h���i����
            '#3728 ���դ�OK,.Fields("CitemStr")�n�令"ACHTNO" By Kin 2008/04/09
            If strACHTNO = Empty And IsNull(.Fields("ACHTNO").OriginalValue) Then
                ReDim aryCitem(0, 4)
                GoTo Fin
            End If
            '*************************************************************************
            aryACHTDESC = Split(Replace(strACHTDESC, "'", ""), ",") '�e���W�ҿ����ACH���v����O
            aryOldACHTDESC = Split(Replace(.Fields("ACHTDESC").OriginalValue & "", "'", ""), ",")
            '#4179 �h�[�@�Ӧr��ŧO,���M�|�X�{Null�����~�T�� By Kin 2008/10/29
            strTmpACH = .Fields("ACHTNO").OriginalValue & ""
            strQry1 = "Select CitemCode From " & GetOwner & "SO003 Where SeqNo in(" & _
                        Replace(rsSO106("CitemStr").OriginalValue & "", "'", "") & ")" & _
                        " And Custid=" & gCustId & " Order by ClctDate,CitemCode"
            strQry2 = "Select CitemName From " & GetOwner & "SO003 Where SeqNo in(" & _
                        Replace(rsSO106("CitemStr").OriginalValue & "", "'", "") & ")" & _
                        " And Custid=" & gCustId & " Order by ClctDate,CitemCode"
            'ACH�e���ҿ���PSO106�ۦP��ܥu�����O���ئ�����
            If UBound(aryOldACHTDESC) = UBound(aryACHTDESC) Then
                lngUbound = UBound(aryACHTDESC)
                ReDim aryCitem(lngUbound, 4)
                ReDim aryACHTNO(UBound(aryACHTDESC), 1)
                aryCitemCode = Split(Replace(strCitemCode, "'", ""), ",")
                aryCitemName = Split(Replace(strCitemName, "'", ""), ",")
                For i = LBound(aryACHTDESC) To UBound(aryACHTDESC)
                    aryACHTNO(i, 0) = Split(Replace(strACHTNO, "'", ""), ",")(i)
                    aryACHTNO(i, 1) = aryACHTDESC(i)
                    aryCitem(i, 0) = aryACHTNO(i, 0)
                    aryCitem(i, 1) = aryACHTNO(i, 1)
                    aryCitem(i, 2) = "2"            '2�N��O�s��
                Next i
                strTmpACH = strACHTNO
            Else
                blnUpdCitem = True  '�N��ACH������,���O���إi��]������,�ҥH�]�n�����䥦��ACH���O����
                aryUpdCitemCode = Split(Replace(strCitemCode, "'", ""), ",")
                aryUpdCitemName = Split(Replace(strCitemName, "'", ""), ",")
                If Not IsNull(rsSO106("CitemStr").OriginalValue) Then
                    '#4260 �p�G��������SO003.SeqNO�{���|�X��,�ҥH�W�[�P�_�O�_Eof By Kin 2008/12/05
                    Set rs = gcnGi.Execute(strQry1)
                    strOldCitemCode = Empty
                    If Not rs.EOF Then
                        'strOldCitemCode = gcnGi.Execute(strQry1).GetString(, , , ",")
                        strOldCitemCode = rs.GetString(, , , ",")
                    End If
                    Set rs = gcnGi.Execute(strQry2)
                    If Not rs.EOF Then
                        'strOldCitemName = gcnGi.Execute(strQry2).GetString(, , , ",")
                        strOldCitemName = rs.GetString(, , , ",")
                    End If
                End If
                '***********************************************************************************
                '���W�[�δ�֪�ACH�H�̦h�����O���ج���,�M��Ӳ��ʹLACH�����O����
                If Right(strOldCitemCode, 1) = "," Then
                    strOldCitemCode = Mid(strOldCitemCode, 1, Len(strOldCitemCode) - 1)
                End If
                If Len(strOldCitemCode) > Len(Replace(strCitemCode, "'", "")) Then
                    aryCitemCode = Split(strOldCitemCode, ",")
                    aryCitemName = Split(strOldCitemName, ",")
                Else
                    aryCitemCode = Split(Replace(strCitemCode, "'", ""), ",")
                    aryCitemName = Split(Replace(strCitemName, "'", ""), ",")
                End If
                '**********************************************************************************
                
                'SO106��ACH���ؤj��e���W�ҬD��ACH����
                If UBound(aryOldACHTDESC) > UBound(aryACHTDESC) Then
                    lngUbound = UBound(aryOldACHTDESC)
                    ReDim aryCitem(lngUbound, 4)
                    ReDim aryACHTNO(lngUbound, 1)
                    For i = LBound(aryOldACHTDESC) To UBound(aryOldACHTDESC)
                        blnDiff = True
                        '��J�S���ʹL��ACH
                        For j = LBound(aryACHTDESC) To UBound(aryACHTDESC)
                            If aryOldACHTDESC(i) = aryACHTDESC(j) Then
                                aryCitem(lngAry, 0) = Split(Replace(strACHTNO, "'", ""), ",")(j)
                                aryCitem(lngAry, 1) = aryACHTDESC(j)
                                aryCitem(lngAry, 2) = "2"
                                lngAry = lngAry + 1
                                blnDiff = False
                                Exit For
                            End If
                        Next j
                        
                        '��J���ʹL��ACH
                        If blnDiff Then
                            aryACHTNO(lngAry, 0) = Split(Replace(.Fields("ACHTNO").OriginalValue, "'", ""), ",")(i)
                            aryACHTNO(lngAry, 1) = aryOldACHTDESC(i)
                            aryCitem(lngAry, 0) = aryACHTNO(lngAry, 0)
                            aryCitem(lngAry, 1) = aryACHTNO(lngAry, 1)
                            aryCitem(lngAry, 2) = "1"   '1�N��s�W�@���������v
                            lngAry = lngAry + 1
                        End If
                    Next i
                Else
                    lngUbound = UBound(aryACHTDESC)
                    strTmpACH = strACHTNO
                    ReDim aryCitem(lngUbound, 4)
                    ReDim aryACHTNO(lngUbound, 1)
                    For i = LBound(aryACHTDESC) To UBound(aryACHTDESC)
                        blnDiff = True
                        For j = LBound(aryOldACHTDESC) To UBound(aryOldACHTDESC)
                            If aryACHTDESC(i) = aryOldACHTDESC(j) Then
                                aryCitem(lngAry, 0) = Split(Replace(strACHTNO, "'", ""), ",")(i)
                                aryCitem(lngAry, 1) = aryOldACHTDESC(j)
                                aryCitem(lngAry, 2) = "2"
                                lngAry = lngAry + 1
                                blnDiff = False
                                Exit For
                            End If
                        Next j
                        If blnDiff Then
                            aryACHTNO(lngAry, 0) = Split(Replace(strACHTNO, "'", ""), ",")(i)
                            aryACHTNO(lngAry, 1) = aryACHTDESC(i)
                            aryCitem(lngAry, 0) = aryACHTNO(lngAry, 0)
                            aryCitem(lngAry, 1) = aryACHTNO(lngAry, 1)
                            aryCitem(lngAry, 2) = "0"
                            blnAddNew = True
                            lngAry = lngAry + 1
                        End If
                    Next i
                End If
            End If
        End With
    End If
    '���ʹLACH,���e���W��ACH�P���������O����
    If blnUpdCitem Then
        For i = LBound(aryUpdCitemCode) To UBound(aryUpdCitemCode)
            '#4094 �L�o���έn���� By Kin 2008/09/17
'            Set rs = gcnGi.Execute("Select * From " & GetOwner & "CD068 Where " & _
'                                    "InStr(',' || CitemCodeStr || ',','," & aryUpdCitemCode(i) & ",')>0 And StopFlag<>1")
            '#4106 �W�[�P�_ACHType=1�Ѽ� By Kin 2008/09/22
'            Set rs = gcnGi.Execute("Select * From " & GetOwner & "CD068 Where " & _
'                        "InStr(',' || CitemCodeStr || ',','," & aryUpdCitemCode(i) & ",')>0  And ACHTNO='" & aryCitem(j, 0) & "' And ACHType=1")
            
'            Set rs = gcnGi.Execute("Select * From " & GetOwner & "CD068 Where " & _
'                        "InStr(',' || CitemCodeStr || ',','," & aryUpdCitemCode(i) & ",')>0  And ACHTNO In(" & strTmpACH & ") And ACHType=1")
            '#7049���դ�OK ���CD068A�h�M�� By Kin 2015/08/06
            If blnStartPost Then
                AchTypeCondition = " In (1,2) "
            Else
                AchTypeCondition = " In (1) "
            End If
            Set rs = gcnGi.Execute("Select CD068.* From " & GetOwner & "CD068A," & GetOwner & "CD068 Where " & _
                                         " Exists (Select * From " & GetOwner & "CD068 Where CD068A.BillHeadFmt = CD068.BillHeadFmt " & _
                                          " And ACHTNo In (" & strTmpACH & ") And ACHType " & AchTypeCondition & ")" & _
                                          " And CitemCode = " & aryCitemCode(i) & _
                                          " And CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                          IIf(strACHTDESC & "" = "", " And 1=1 ", " And CD068.ACHTDESC In (" & strACHTDESC & ")"))

            For j = LBound(aryCitem) To UBound(aryCitem)
                If aryCitem(j, 0) = rs("ACHTNO") And aryCitem(j, 1) = rs("ACHTDESC") Then
                    aryCitem(j, 3) = IIf(aryCitem(j, 3) = Empty, aryUpdCitemCode(i), aryCitem(j, 3) & "," & aryUpdCitemCode(i))
                    aryCitem(j, 4) = IIf(aryCitem(j, 4) = Empty, aryUpdCitemName(i), aryCitem(j, 4) & "," & aryUpdCitemName(i))
                    Exit For
                
                End If
            Next j
        Next i
    
    End If
    '��JACH�ҹ��������O����,�p�G�O���ACH,�h�n�M��O"1"������,�Ϥ��n�M�䬰0��
    For i = LBound(aryCitemCode) To UBound(aryCitemCode)
        '#4094 �L�o���έn���� By Kin 2008/09/17

'        Set rs = gcnGi.Execute("Select * From " & GetOwner & "CD068 Where " & _
'                                    "InStr(',' || CitemCodeStr || ',','," & aryCitemCode(i) & ",')>0 And ACHTNO In(" & strTmpACH & ") And ACHType=1")
        
        '#7049���դ�OK ���CD068A�h�M�� By Kin 2015/08/06
            If blnStartPost Then
                AchTypeCondition = " In ( 1,2 ) "
            Else
                AchTypeCondition = " In ( 1 ) "
            End If
         Set rs = gcnGi.Execute("Select CD068.* From " & GetOwner & "CD068A," & GetOwner & "CD068 Where " & _
                                          " Exists (Select * From " & GetOwner & "CD068 Where CD068A.BillHeadFmt = CD068.BillHeadFmt " & _
                                           " And ACHTNo In (" & strTmpACH & ") And ACHType " & AchTypeCondition & ")" & _
                                           " And CitemCode = " & aryCitemCode(i) & _
                                           " And CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                           IIf(strACHTDESC & "" = "", " And 1 = 1", " And CD068.ACHTDESC In (" & strACHTDESC & ")"))


                                                    
                                    '"InStr(',' || CitemCodeStr || ',','," & aryCitemCode(i) & ",')>0 And ACHTNO In(" & strTmpACH & ") And ACHType=1")
        If Not rs.EOF Then
            For j = LBound(aryCitem) To UBound(aryCitem)
                If aryCitem(j, 0) = rs("ACHTNO") And aryCitem(j, 1) = rs("ACHTDESC") Then
                    If blnUpdCitem Then
                        If blnAddNew Then
'                            If aryCitem(j, 2) = "0" Then
'                                aryCitem(j, 3) = IIf(aryCitem(j, 3) = Empty, aryCitemCode(i), aryCitem(j, 3) & "," & aryCitemCode(i))
'                                aryCitem(j, 4) = IIf(aryCitem(j, 4) = Empty, aryCitemName(i), aryCitem(j, 4) & "," & aryCitemName(i))
'                                Exit For
'                            End If
                        Else
                            If aryCitem(j, 2) = "1" Then
                                aryCitem(j, 3) = IIf(aryCitem(j, 3) = Empty, aryCitemCode(i), aryCitem(j, 3) & "," & aryCitemCode(i))
                                aryCitem(j, 4) = IIf(aryCitem(j, 4) = Empty, aryCitemName(i), aryCitem(j, 4) & "," & aryCitemName(i))
                                Exit For
                            End If
                        End If
                    Else
                        aryCitem(j, 3) = IIf(aryCitem(j, 3) = Empty, aryCitemCode(i), aryCitem(j, 3) & "," & aryCitemCode(i))
                        aryCitem(j, 4) = IIf(aryCitem(j, 4) = Empty, aryCitemName(i), aryCitem(j, 4) & "," & aryCitemName(i))
                        Exit For
                    End If
                End If
            Next j

        End If
    Next i
Fin:
    On Error Resume Next
    Call CloseRecordset(rs)
    Erase aryACHTDESC
    Erase aryACHTNO
    Erase aryCitemCode
    Erase aryCitemName
    Erase aryOldACHTDESC
    Erase aryUpdCitemCode
    Erase aryUpdCitemName
  Exit Sub
ChkErr:
  ErrSub Me.Name, "GetCitemCode"
End Sub
'#3946 ���oSO106.MasterId��Sequence By Kin 2008/05/30
Private Function Get_MasterID_Seq() As String
  On Error GoTo ChkErr
    Get_MasterID_Seq = RPxx(gcnGi.Execute("SELECT " & GetOwner & "S_SO106_MasterId.NEXTVAL FROM DUAL").GetString & "")
    'Get_Cmd_Seq_No = Format(Get_Cmd_Seq_No, "00000000")
  Exit Function
ChkErr:
    ErrSub "mod_SysLib", "Get_MasterID_Seq"
End Function

