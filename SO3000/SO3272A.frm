VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO3272A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�H�Υd���ڸ�Ʋ��ͧ@�~ [SO3272A]"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   Icon            =   "SO3272A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   10950
   StartUpPosition =   1  '���ݵ�������
   Begin VB.Frame Frame1 
      Caption         =   "�d�߱���"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   8250
      Left            =   180
      TabIndex        =   31
      Top             =   -15
      Width           =   10665
      Begin VB.CheckBox chkFubonIntegrate 
         Caption         =   "�]�����x�榡"
         Height          =   285
         Left            =   420
         TabIndex        =   50
         Top             =   1770
         Value           =   1  '�֨�
         Width           =   1485
      End
      Begin VB.CheckBox chkExcI 
         Caption         =   "�ư�I�椤�b����O�O�ŭ�"
         Height          =   195
         Left            =   3480
         TabIndex        =   48
         Top             =   1500
         Width           =   3135
      End
      Begin VB.CheckBox chkZero 
         Caption         =   "��i�b��X�p�`�B<=0 �O�_�n����"
         Height          =   195
         Left            =   420
         TabIndex        =   47
         Top             =   1500
         Width           =   3135
      End
      Begin VB.TextBox txtCustId 
         Height          =   360
         Left            =   2100
         TabIndex        =   43
         Top             =   6900
         Width           =   5925
      End
      Begin VB.CheckBox chkErr 
         Caption         =   "�ˮֲ��X�q�l��"
         Height          =   315
         Left            =   6810
         TabIndex        =   42
         Top             =   1020
         Width           =   1605
      End
      Begin VB.ComboBox cboBillHeadFmt 
         Height          =   300
         ItemData        =   "SO3272A.frx":0442
         Left            =   2310
         List            =   "SO3272A.frx":0444
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   1
         Top             =   600
         Width           =   3855
      End
      Begin VB.ComboBox cboCreditCardType 
         Height          =   300
         ItemData        =   "SO3272A.frx":0446
         Left            =   5550
         List            =   "SO3272A.frx":0448
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   2
         Top             =   270
         Width           =   2715
      End
      Begin VB.Frame fraTCB 
         Enabled         =   0   'False
         Height          =   795
         Left            =   8490
         TabIndex        =   36
         Top             =   1260
         Width           =   1875
         Begin VB.OptionButton optNotTCB 
            Caption         =   "�L��d"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   165
            TabIndex        =   13
            Top             =   555
            Width           =   1335
         End
         Begin VB.OptionButton optTCB 
            Caption         =   "���ذӻȦۦ�d"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   150
            TabIndex        =   12
            Top             =   270
            Value           =   -1  'True
            Width           =   1590
         End
      End
      Begin VB.CheckBox chkChinaBank 
         Caption         =   "�H���ذӻȮ榡�X��"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   8490
         TabIndex        =   11
         Top             =   1050
         Width           =   1965
      End
      Begin VB.Frame Frame2 
         Caption         =   " �j�Ө̾�"
         Height          =   1005
         Left            =   195
         TabIndex        =   32
         Top             =   7185
         Width           =   10185
         Begin VB.OptionButton optMduYes 
            Caption         =   "&A.�n����"
            Height          =   255
            Left            =   450
            TabIndex        =   27
            Top             =   690
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton optMduNo 
            Caption         =   "&B.������"
            Height          =   285
            Left            =   1830
            TabIndex        =   28
            Top             =   660
            Width           =   1125
         End
         Begin CS_Multi.CSmulti gmdMduid 
            Height          =   405
            Left            =   150
            TabIndex        =   26
            Top             =   240
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   714
            ButtonCaption   =   "�j �� �W ��"
         End
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "�]�t�@���"
         Height          =   225
         Left            =   5565
         TabIndex        =   10
         Top             =   1065
         Value           =   1  '�֨�
         Width           =   2145
      End
      Begin Gi_Multi.GiMulti gmdBillType 
         Height          =   360
         Left            =   165
         TabIndex        =   25
         Top             =   5820
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "�� �� �� �O"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin Gi_Multi.GiMulti gmdClctEn 
         Height          =   360
         Left            =   165
         TabIndex        =   20
         Top             =   4125
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "��    �O    ��"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin Gi_Multi.GiMulti gmdCustClass 
         Height          =   360
         Left            =   165
         TabIndex        =   19
         Top             =   3780
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "�� �� �� �O"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin Gi_Multi.GiMulti gmdServCode 
         Height          =   360
         Left            =   165
         TabIndex        =   17
         Top             =   3120
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "�A    ��    ��"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin Gi_Multi.GiMulti gmdAreaCode 
         Height          =   360
         Left            =   165
         TabIndex        =   16
         Top             =   2775
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "��    �F    ��"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin Gi_Multi.GiMulti gmdCMCode 
         Height          =   360
         Left            =   180
         TabIndex        =   15
         Top             =   2430
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "�� �O �� ��"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin Gi_Multi.GiMulti gimCustStatus 
         Height          =   360
         Left            =   165
         TabIndex        =   18
         Top             =   3435
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "��  �� �� �A"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin Gi_Multi.GiMulti gimCreateEn 
         Height          =   360
         Left            =   165
         TabIndex        =   22
         Top             =   4815
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "�� �� �H ��"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin prjGiList.GiList gilServiceType 
         Height          =   330
         Left            =   6675
         TabIndex        =   9
         Top             =   -30
         Visible         =   0   'False
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F2Corresponding =   -1  'True
      End
      Begin Gi_Date.GiDate gdaShouldDate2 
         Height          =   360
         Left            =   9180
         TabIndex        =   6
         Top             =   600
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   635
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
         Height          =   360
         Left            =   7590
         TabIndex        =   5
         Top             =   600
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   635
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
      Begin Gi_Multi.GiMulti gmBank 
         Height          =   360
         Left            =   165
         TabIndex        =   14
         Top             =   2115
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "�� �� �W ��"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin Gi_Multi.GiMulti gimUCCode 
         Height          =   360
         Left            =   165
         TabIndex        =   23
         Top             =   5145
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "��  ��  ��  �]"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin CS_Multi.CSmulti gimCitemCode 
         Height          =   360
         Left            =   165
         TabIndex        =   24
         Top             =   5475
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "��  �O  ��  ��"
         IsReadOnly      =   -1  'True
      End
      Begin prjGiList.GiList gilCompCode 
         Height          =   330
         Left            =   1275
         TabIndex        =   0
         Top             =   240
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F2Corresponding =   -1  'True
      End
      Begin Gi_Date.GiDate gdaCreateTime2 
         Height          =   330
         Left            =   4500
         TabIndex        =   4
         Top             =   -60
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
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
      Begin Gi_Date.GiDate gdaCreateTime1 
         Height          =   330
         Left            =   3180
         TabIndex        =   3
         Top             =   -60
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
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
      Begin Gi_Multi.GiMulti gmdOldClctEn 
         Height          =   360
         Left            =   165
         TabIndex        =   21
         Top             =   4470
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "�� �� �O  ��"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin Gi_Time.GiTime gitCreateTime1 
         Height          =   345
         Left            =   1290
         TabIndex        =   7
         Top             =   1020
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
         Left            =   3630
         TabIndex        =   8
         Top             =   1020
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
      Begin Gi_Multi.GiMulti gmdPayType 
         Height          =   330
         Left            =   180
         TabIndex        =   46
         Top             =   6150
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   582
         ButtonCaption   =   "ú  �I  ��  �O"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin CS_Multi.CSmulti gimAMduId 
         Height          =   405
         Left            =   180
         TabIndex        =   49
         Top             =   6480
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   714
         ButtonCaption   =   "�M  ��  �s  ��"
      End
      Begin prjGiList.GiList gilFubonCMCode 
         Height          =   330
         Left            =   3480
         TabIndex        =   51
         Top             =   1740
         Visible         =   0   'False
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   582
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F2Corresponding =   -1  'True
      End
      Begin VB.Label lblFubonCMCode 
         AutoSize        =   -1  'True
         Caption         =   "�]�����x���O�覡"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   1890
         TabIndex        =   52
         Top             =   1830
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "�ХΡu,�v���} Exp 67,82,199"
         Height          =   195
         Left            =   8070
         TabIndex        =   45
         Top             =   6960
         Width           =   2535
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "���ͫȽs"
         Height          =   195
         Left            =   1230
         TabIndex        =   44
         Top             =   6930
         Width           =   855
      End
      Begin VB.Label lblBillHeadFmt 
         AutoSize        =   -1  'True
         Caption         =   "�h�b�Უ�ͨ̾ڳ]�w"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   420
         TabIndex        =   41
         Top             =   690
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�ӿ�Ȧ�"
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
         Left            =   4740
         TabIndex        =   40
         Top             =   360
         Width           =   780
      End
      Begin VB.Label lblCreateTime 
         AutoSize        =   -1  'True
         Caption         =   "���ͮɶ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   390
         TabIndex        =   39
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label lblunti2 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   3330
         TabIndex        =   38
         Top             =   1110
         Width           =   195
      End
      Begin VB.Label lblCompCode 
         AutoSize        =   -1  'True
         Caption         =   "���q�O"
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
         Height          =   240
         Left            =   615
         TabIndex        =   37
         Top             =   345
         Width           =   585
      End
      Begin VB.Label lbluntil 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   210
         Left            =   8880
         TabIndex        =   35
         Top             =   705
         Width           =   195
      End
      Begin VB.Label lblReadAmt 
         AutoSize        =   -1  'True
         Caption         =   "��������d��"
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
         Height          =   210
         Left            =   6390
         TabIndex        =   34
         Top             =   690
         Width           =   1170
      End
      Begin VB.Label lblServiceType 
         AutoSize        =   -1  'True
         Caption         =   "�A�����O"
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
         Height          =   210
         Left            =   5820
         TabIndex        =   33
         Top             =   45
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�T�w"
      Default         =   -1  'True
      Height          =   405
      Left            =   600
      TabIndex        =   29
      Top             =   8280
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "����(&X)"
      Height          =   405
      Left            =   9090
      TabIndex        =   30
      Top             =   8280
      Width           =   1275
   End
End
Attribute VB_Name = "frmSO3272A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objAgency As Object
'Private objAgency As New clsInterface
Private objAction As Object
Private rsBankData As New ADODB.Recordset
Private rsDefVal As New ADODB.Recordset
'Private strChoose As String
Private blnUnload As Boolean
Private strINIpath1 As String
Private strErrPath As String
Private strChooseString As String
Dim pChooseMultiAcc As String
Private intPara24 As Integer
Private blnChina As Boolean  ''  2003/10/07   �P�_�O�_���s���i������ӻ�
Private rsBillHeadFmt As New ADODB.Recordset
'#3531 �h�s�W�@�Ӥ���H�U(�s)�y�{ By Kin 2007/10/31
Public Enum CreditCardType     ''  2003/10/27  �P�_�ҥΪ��O���@�w���Ȧ�
      CardType_default = 0  '' �w�]�� �p�X�P����  �F�˥Ϊ�
      CardType_chinatrust = 1   '' ����H�U �ثe�O �s���i
      CardType_hsbc = 2  '' ���׻Ȧ�H�Υd
      CardType_newChina = 3
      CardType_CitiBank = 4
      CardType_NewWeb = 5   '�ŷs���
      CardType_Batch = 6    '�p�X�H�Υd��������
      CardType_Fubon = 7    '�I��
End Enum
Public e_CreditCardType As CreditCardType
Private blnPaynowFlag As Boolean
Private strDefSpcNo As String
Private strfunDefCmCode As String
Private CmRefNo5 As String
Private Sub cboBillHeadFmt_Click()
    On Error Resume Next
    
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
    Dim strCD019SQL As String
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
        '#7049 ���CD068A.CitemCode By Kin 2015/07/09
        strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                                    " And CD068.BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1"
        'If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        If Not GetRS(rsCD019, strCD019SQL) Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        Screen.MousePointer = vbHourglass
        Me.Enabled = False
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        gimCitemCode.Filter = gimCitemCode.Filter
        gimCitemCode.SetQueryCode strCitemCode
        Me.Enabled = True
        Screen.MousePointer = vbDefault
    strDefSpcNo = Empty
    If cboBillHeadFmt.ListIndex > -1 Then
        If Not rsBillHeadFmt.EOF Then
            strDefSpcNo = GetRsValue("SELECT SPCNO FROM " & GetOwner & "CD068 " & _
                            " WHERE ROWID = '" & rsBillHeadFmt("ROWID") & "' ", gcnGi) & ""
                            
        End If
    End If
End Sub

Private Sub cboCreditCardType_Click()
    gilFubonCMCode.Enabled = False
    lblFubonCMCode.ForeColor = vbBlack
    gilFubonCMCode.Clear
    Call SetgiMulti(gmdCMCode, "CodeNo", "Description", "CD031", "�N�X", "�W��", _
                                   " Where Nvl(StopFlag,0) = 0 ")
  Select Case cboCreditCardType.ItemData(cboCreditCardType.ListIndex)
       Case 0
           e_CreditCardType = CardType_default
           chkChinaBank.Visible = True
           chkChinaBank.Value = 0
           fraTCB.Visible = True
           Call SetgiMulti(gmBank, "CodeNo", "Description", "CD018", "�N�X", "�W��", _
                                   " Where UPPER(PrgName) = 'CREDITCARDUNION'  AND  COMPCODE =" & _
                                   gilCompCode.GetCodeNo & " AND " & "(BANKID <> 'TCB' OR BANKID IS NULL)")
       Case 1
           e_CreditCardType = CardType_chinatrust
           chkChinaBank.Visible = False
           chkChinaBank.Value = 0
           fraTCB.Visible = False
            Call SetgiMulti(gmBank, "CodeNo", "Description", "CD018", "�N�X", "�W��", _
                                                  " Where UPPER(PrgName) = '" & "CREDITCARDCHINATRUST" & "' AND  COMPCODE =" & gilCompCode.GetCodeNo)
        Case 2
            e_CreditCardType = CardType_hsbc
            chkChinaBank.Visible = False
            chkChinaBank.Value = 0
            fraTCB.Visible = False
            Call SetgiMulti(gmBank, "CodeNo", "Description", "CD018", "�N�X", "�W��", _
                                                  " Where UPPER(PrgName) = '" & "CREDITCARDHSBC" & "' AND  COMPCODE =" & gilCompCode.GetCodeNo)
        '#3531 �s�W����H�U(�s)�榡
        Case 3
            e_CreditCardType = CardType_newChina
            chkChinaBank.Visible = False
            chkChinaBank.Value = 0
            fraTCB.Visible = False
            Call SetgiMulti(gmBank, "CodeNo", "Description", "CD018", "�N�X", "�W��", _
                                                  " Where UPPER(PrgName) = '" & "CREDITCARDTRUSTNEW" & "' AND  COMPCODE =" & gilCompCode.GetCodeNo)
        '#4002
        Case 4
            e_CreditCardType = CardType_CitiBank
            chkChinaBank.Visible = False
            chkChinaBank.Value = 0
            fraTCB.Visible = False
            Call SetgiMulti(gmBank, "CodeNo", "Description", "CD018", "�N�X", "�W��", _
                                            " Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'")
            'gmBank.Filter = gmBank.Filter & IIf(gmBank.Filter = "", " Where ", " And ") & " UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'"
        '#5292 �W�[�ŷs��ޮ榡 By Kin 2009/09/15
        Case 5
            e_CreditCardType = CardType_NewWeb
            chkChinaBank.Visible = False
            chkChinaBank.Value = 0
            fraTCB.Visible = False
            Call SetgiMulti(gmBank, "CodeNo", "Description", "CD018", "�N�X", "�W��", _
                                            " Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "' And CompCode=" & gilCompCode.GetCodeNo)
        Case 6
            e_CreditCardType = CardType_Batch
            chkChinaBank.Visible = True
            chkChinaBank.Value = 0
            fraTCB.Visible = True
            Call SetgiMulti(gmBank, "CodeNo", "Description", "CD018", "�N�X", "�W��", _
                                   " Where UPPER(PrgName) = 'CREDITCARDUNION1'  AND  COMPCODE =" & _
                                   gilCompCode.GetCodeNo & " AND " & "(BANKID <> 'TCB' OR BANKID IS NULL)")
         '#7765 By Kin 2018/05/08
        Case 7
            e_CreditCardType = CardType_Fubon
            'chkChinaBank.Visible = True
            'chkChinaBank.Value = 0
            'fraTCB.Visible = True
            '#8537 By Kin 2019/12/17
            If chkFubonIntegrate.Value = 1 Then
                gilFubonCMCode.Enabled = True
                lblFubonCMCode.ForeColor = vbRed
                Call SetgiMulti(gmdCMCode, "CodeNo", "Description", "CD031", "�N�X", "�W��", _
                                   " Where Nvl(StopFlag,0) = 0 And RefNo = 5 ")
                gmdCMCode.SelectAll
                
'                If Len(CmRefNo5 & "") > 0 Then
'                    gmdCMCode.SetQueryCode CmRefNo5
'
'                End If
            Else
                Call SetgiMulti(gmdCMCode, "CodeNo", "Description", "CD031", "�N�X", "�W��", _
                                   " Where Nvl(StopFlag,0) = 0 And Nvl(RefNo,0) <> 5 ")
            End If
            Call SetgiMulti(gmBank, "CodeNo", "Description", "CD018", "�N�X", "�W��", _
                                   " Where UPPER(PrgName) = 'CREDITCARDFUBON'  AND  COMPCODE =" & _
                                   gilCompCode.GetCodeNo)
                                               
    End Select
End Sub

Private Sub chkChinaBank_Click()
    If chkChinaBank.Value = 1 Then
       gmBank.Clear
       gmBank.Enabled = 0
       fraTCB.Enabled = True
    Else
       gmBank.Enabled = 1
       fraTCB.Enabled = False
    End If
End Sub

Private Sub chkFubonIntegrate_Click()
    If e_CreditCardType <> CardType_Fubon Then Exit Sub
    If chkFubonIntegrate.Value = 0 Then
        gilFubonCMCode.Clear
        lblFubonCMCode.ForeColor = vbBlack
        Call SetgiMulti(gmdCMCode, "CodeNo", "Description", "CD031", "�N�X", "�W��", _
                                   " Where Nvl(StopFlag,0) = 0 And Nvl(RefNo,0) <> 5 ")
    Else
        If e_CreditCardType = CardType_Fubon Then
            lblFubonCMCode.ForeColor = vbRed
        End If
        '#8537 By Kin 2019/12/17
        Call SetgiMulti(gmdCMCode, "CodeNo", "Description", "CD031", "�N�X", "�W��", _
                                   " Where Nvl(StopFlag,0) = 0 And Nvl(RefNo,0) = 5 ")
        gmdCMCode.SelectAll
        'gmdCMCode.SetQueryCode CmRefNo5
    End If
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo chkErr
    Unload Me
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "GetBankData")
End Sub

Private Function GetBankData() As Boolean
On Error GoTo chkErr
Dim strWhere As String
 
   Dim strSQL  As String
   
   If Len(gmBank.GetQryStr) <> 0 Then
      strWhere = " AND CodeNo " & gmBank.GetQryStr
   End If
   
'   If chkChinaBank.Value = 1 Then
'       strWhere = " AND BANKID = 'TCB' "
'   End If
    '#8528 SPCNO2 By Kin 2019/11/29
   strSQL = "Select CodeNO,Description,PrgName,ActNo,BankId,CorpID,Spcno,SpcNo2 From " & GetOwner & "CD018 where " & " UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'" & strWhere
   Set rsBankData = gcnGi.Execute(strSQL)
   If rsBankData.EOF Or rsBankData.BOF Then MsgBox "�Ȧ��Ʀ��~�I���ˬd�Ȧ�N�X�ɡI", vbExclamation, "�T���I": Exit Function
   
   GetBankData = True
Exit Function
chkErr:
    Call ErrSub(Me.Name, "GetBankData")
End Function
Private Sub CallSO3272B()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim objChkErr As Object
    Dim blnErrNum As Boolean
    Dim strBank As String
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    blnUnload = False
    If Not subChoose Then GoTo lExit
    Set objAgency = CreateObject("CreditCardOut.clsInterface")
    If gmdCustClass.GetQryStr <> "" Then
        strSQL = " From SO033 A,SO001 B Where " & strChoose
    Else
        strSQL = " From SO033 A Where " & strChoose
    End If
    If InStr(1, gmBank.GetQueryCode, ",") > 0 Then
        strBank = Mid(gmBank.GetQueryCode, 1, InStr(1, gmBank.GetQueryCode, ",") - 1)
    Else
        strBank = gmBank.GetQueryCode
    End If
    If Not GetBankData Then Exit Sub
        With objAgency
            .errPath = strErrPath
            .iniPath = strINIpath1
            .ChooseStr = strChoose
            .uUPDEN = garyGi(1)
            .uUpdTime = GetDTString(RightNow)
            '.ActNo = rsBankData("ActNo") & ""
           ' .BankId = rsBankData("CodeNO") & ""
           If Len(gmBank.GetQryStr) <> 0 Then
             .BankId = " A.BANKCODE " & gmBank.GetQryStr
           Else
              .BankId = " A.BANKCODE IN (" & gmBank.GetQueryCode & ")"
           End If
            .BankName = Me.Caption & " ��X�榡 "
            '.CorpID = rsBankData("CorpId") & ""
            
              .pTableOwnerName = GetOwner
            .prgName = rsBankData("PrgName") & ""
            .uSpcNO = rsBankData("SpcNO") & ""
            Set .Connection = gcnGi
            If Len(rsBankData("PrgName") & "") = 0 Then
                MsgBox "�]�w�{���W�ٵL�]�εL�ϥ��v���I�I", vbExclamation, "����"
            Else
                'MsgBox ReadGICMIS1("ErrLogPath"), , "frmSO3273A"
                Set objAction = .InitObject(rsBankData("Prgname") & "")
                
                Dim intPara24 As Integer
                
                 intPara24 = CInt(RPxx(gcnGi.Execute( _
                                                          "SELECT Para24 FROM " & GetOwner & "SO043 " & _
                                                          "WHERE " & _
                                                          "CompCode =" & gilCompCode.GetCodeNo & " AND ServiceType ='" & gilServiceType.GetCodeNo & "'" _
                                                          ).GetString))
               Call objAction.Action(pChooseMultiAcc, intPara24)
               
                Set objAction = Nothing
            End If
      End With
      '�P�_�O�_�����
      If objAgency.unodata Then
        Set objAgency = Nothing
        Screen.MousePointer = vbDefault
        blnUnload = True
        On Error Resume Next
        gilCompCode.SetFocus
        Exit Sub
      Else
      '�P�_�O�_���榨�\
        If objAgency.uupdate Then
            Call subInsert
            If chkErr.Value Then
                Set objChkErr = CreateObject("CheckMediaTEXT.clsExechkMediaText")
                With objChkErr
                    .ugiOpenFileType = 1 'Error��r�ɶ}�Ҥ覡  0=Create  1=Append
                    .uClassName = "CreditCardCitiBank_" & objAgency.uChkType 'Class�W��
                    .uPathFile = objAgency.uProcText '��Ƹ��|�� ���|+�ɦW
                    .uErrPathFile = objAgency.uProcErrText '���~�ɮ� ���|+�ɦW
                         '.uSetPathFile = Text3.Text '�]�w�ɮ� ���|+�W��
                    .uBankCode = strBank
                    .uOwner = GetOwner
                    .uCompCode = gilCompCode.GetCodeNo
                    .ugcnGi = gcnGi
                    .CheckMediaTextShow
                         'blnReturn = .uReturnOK '�^�ǬO�_�w�g���� TRUE=���槹���S�����D  FALSE=���椤�Φ����D(�ݭn����)
                    blnErrNum = .ublnError '�^�ǬO�_��ERROR����� TRUE=�����~���  FALSE=�L���~���
                End With
                If blnErrNum Then
                    Shell "Notepad " & objAgency.uProcErrText, vbNormalFocus
                End If
            End If
            blnUnload = True
        End If
      End If
lExit:
      On Error Resume Next
      blnUnload = True
      Set objAgency = Nothing
      Set objChkErr = Nothing
      DoEvents
      Screen.MousePointer = vbDefault
     
  Exit Sub
chkErr:
    blnUnload = True
    If Err.Number = 20 Then Resume Next
    If Err.Number = 91 Then
       MsgBox "��b�{���W�ٿ��~�Ϊ̸ӻȦ�S����b��Ʋ��ͥ\��!", vbExclamation, "ĵ�i..."
       Exit Sub
    End If
    Call ErrSub(Me.Name, "cmdOk_Click")
End Sub
Private Sub LogToSO054(prgName As String, _
                      EmpNo As String, _
                      RunDateTime As String, _
                      SpendTime As String, _
                      Selection As String, _
                      Optional ProcessTime As String)
  On Error GoTo chkErr
  Dim strSQL As String
    strSQL = "INSERT INTO " & GetOwner & "SO054 " & _
                    "(PRGNAME,EMPNO,RUNDATETIME,SECOND,SELECTION,StopDateTime,ReportType) VALUES (" & _
                    GetNullString(prgName) & "," & GetNullString(EmpNo) & _
                    "," & GetNullString(RunDateTime, giDateV, giOracle, True) & "," & _
                    GetNullString(SpendTime) & "," & GetNullString(Selection) & "," & _
                    GetNullString(ProcessTime, giDateV, giOracle, True) & "," & _
                    "0)"
    strSQL = Replace(strSQL, Chr(0), "")
    gcnGi.Execute strSQL
  Exit Sub
chkErr:
    If Err.Number = -2147217865 Then
        MsgBox "SO054 Table ���s�b !! �L�k�g�J ExportLog", vbInformation, "�T��"
    Else
        Call ErrSub(Me.Name, "LogToSO054")
    End If
End Sub
Private Sub cmdOk_Click()
  On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim blnErrNum As Boolean
    Dim objChkErr As Object
    Dim strOtherBank As String
    Dim strBank As String
    Dim LastTime As Single
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    ReadyGoPrint
    blnUnload = False
    If Not subChoose Then GoTo lExit
    '#4002
    If e_CreditCardType = CardType_CitiBank Then
        Call CallSO3272B
        Exit Sub
    Else
        Set objAgency = CreateObject("CreditCardOutEmc.clsInterface")
        If gmdCustClass.GetQryStr <> "" Then
            strSQL = " From " & GetOwner & "SO033 A," & GetOwner & "SO001 B Where " & strChoose
        Else
            strSQL = " From " & GetOwner & "SO033 A Where " & strChoose
        End If
           
          '' 2003/10/08  �p�G�O����ӻȼȮɤ��ޥL�� PTCODE
           
        If e_CreditCardType = CardType_default Then
            If chkChinaBank.Value = 1 Then
                If optTCB.Value = True Then
                     strSQL = strSQL & "  AND A.PTCode = 3 "
                Else
                     strSQL = strSQL & "  AND A.PTCode = 5"
                End If
            Else
                strSQL = strSQL & "  AND (A.PTCode <>  3 AND A.PTCODE <> 5  OR A.PTCode IS  NULL) "
            End If
        End If
          
        If Not GetBankData Then Exit Sub
        If InStr(1, gmBank.GetQueryCode, ",") > 0 Then
            strBank = Mid(gmBank.GetQueryCode, 1, InStr(1, gmBank.GetQueryCode, ",") - 1)
        Else
            strBank = gmBank.GetQueryCode
        End If

        With objAgency
              .errPath = strErrPath
              .iniPath = strINIpath1
              .ChooseStr = strChoose
              .uUPDEN = garyGi(1)
              .uUpdTime = GetDTString(RightNow)
              '#6441 ��i�b��X�p�`�B<=0�O�_���� By Kin 2013/05/13
              .uOutZero = chkZero.Value = 1
              .uExcI = chkExcI.Value = 1
              .uFubonIntegrate = chkFubonIntegrate.Value = 1
              .uSpcNO2 = rsBankData("SPCNo2") & ""
              
              .uFuBonCmCode = gilFubonCMCode.GetCodeNo & ""
              .uFuBonCmName = gilFubonCMCode.GetDescription & ""
            If Len(gmBank.GetQryStr) <> 0 Then
                     '' �H�U�אּ�H SO033 ��BANKCODE �@�P�_ �Ӥ��O SO002 .BankId = " C.BANKCODE " & gmBank.GetQryStr
                     '' 2003/10/08  �p�G�O����ӻȼȮɤ��ޥL�� PTCODE
                Select Case e_CreditCardType
                    Case 0
                        .BankId = " A.BANKCODE " & gmBank.GetQryStr & " AND (A.PTCODE NOT IN (3,5)   OR A.PTCode IS  NULL) "
                    Case 1
                        .BankId = " A.BANKCODE " & gmBank.GetQryStr
                    Case 2
                        .BankId = " A.BANKCODE " & gmBank.GetQryStr
                    '#3531 �W�[����H�U�s�榡 By Kin 2007/11/02
                    Case 3
                        .BankId = " A.BANKCODE " & gmBank.GetQryStr
                        .Bankstr = gmBank.GetQryStr
                    '#5292 �s�W�ŷs��� By Kin 2009/09/15
                    Case 5
                        .BankId = " A.BANKCODE " & gmBank.GetQryStr
                End Select
            Else
                If chkChinaBank.Value = 1 Then
                    If optTCB.Value = True Then
                        .BankId = " A.PTCode = 3  "
                    Else
                        .BankId = " A.PTCode = 5  "
                    End If
                Else
                    Select Case e_CreditCardType
                        Case 0
                            .BankId = " A.BANKCODE IN (" & gmBank.GetQueryCode & ")  AND (A.PTCODE NOT IN (3,5)  OR A.PTCODE IS NULL) "
                        Case Else
                            .BankId = " A.BANKCODE IN(" & gmBank.GetQueryCode & ")"
                    End Select
                End If
            End If
           ''  2003/10/08  �Y�O����ӻȫh  e_CreditCardType = CardType_chinatrust
            strOtherBank = ""
            Select Case e_CreditCardType
                Case 0
                    If chkChinaBank.Value = 0 Then
                        .BankName = Me.Caption & "  �p�X�H�Υd "
                        .blnIsTCB = False
                        .blnIsTCBCard = False
                        strOtherBank = "_0"
                    Else
                        If optTCB.Value = True Then
                            .BankName = Me.Caption & "  ���ذӻȦۦ�d  "
                            .blnIsTCB = True
                            .blnIsTCBCard = True
                            strOtherBank = "_1"
                        Else
                            .BankName = Me.Caption & "  ���ذӻȥL��d "
                            .blnIsTCB = True
                            .blnIsTCBCard = False
                            strOtherBank = "_2"
                        End If
                     End If
                Case 1
                    .BankName = Me.Caption & "  ����H�U�H�Υd "
                    .blnIsTCB = False
                    .blnIsTCBCard = False
                Case 2
                    .BankName = Me.Caption & "  ���׻Ȧ�H�Υd "
                '#3531 �W�[����H�U�s�榡
                Case 3
                    .BankName = Me.Caption & "  ����H�U�H�Υd "
                    .blnIsTCB = False
                    .blnIsTCBCard = False
                Case 5
                    .BankName = Me.Caption & "  �ŷs��ީw���w�B���� "
                    .blnIsTCB = False
                    .blnIsTCBCard = False
                Case 7
                    .BankName = Me.Caption & "  �I���Ȧ� "
            End Select
            .pTableOwnerName = GetOwner
            .prgName = rsBankData("PrgName") & ""
            .uSpcNO = rsBankData("SpcNO") & ""
            .uDefSpcNo = strDefSpcNo
            Set .Connection = gcnGi
            If Len(rsBankData("PrgName") & "") = 0 Then
                MsgBox "�]�w�{���W�ٵL�ĩεL�ϥ��v���I�I", vbExclamation, "����"
            Else
                '' 2003/10/29  �ѩ�ثe�令���O������CD018 ��PRGNAME�@����W�١A�ҥH�u�������R�W
                '' Set objAction = .InitObject(rsBankData("Prgname") & "")
                 Set objAction = .InitObject("CREDITCARDUNION")
                 Call objAction.Action(.blnIsTCB, pChooseMultiAcc, .blnIsTCBCard, CLng(e_CreditCardType))
                 Set objAction = Nothing
            End If
        
        End With
        '�P�_�O�_�����
          If objAgency.unodata Then
            Set objAgency = Nothing
            '7207 Add Data into SO054 Table By Kin 2016/03/08
            LastTime = timeGetTime / 1000
            LogToSO054 Replace(UCase(Screen.ActiveForm.Name), "FRM", ""), _
                        garyGi(0), Format(NowTime, "YYYYMMDD HHMMSS"), SecToHMS(LastTime - startTime), strChooseString, RightNow
            Screen.MousePointer = vbDefault
            blnUnload = True
            On Error Resume Next
            gilCompCode.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
          Else
              '�P�_�O�_���榨�\
              If objAgency.uupdate Then
                    Call subInsert
                    If chkErr.Value And (Not objAgency.unodata) Then
                        Set objChkErr = CreateObject("CheckMediaTEXT.clsExechkMediaText")
                        With objChkErr
                            .ugiOpenFileType = 1 'Error��r�ɶ}�Ҥ覡  0=Create  1=Append
                            .uClassName = "CreditCardUnion_" & e_CreditCardType & strOtherBank  'Class�W��
                            .uPathFile = objAgency.uProcText '��Ƹ��|�� ���|+�ɦW
                            .uErrPathFile = objAgency.uProcErrText '���~�ɮ� ���|+�ɦW
                                 '.uSetPathFile = Text3.Text '�]�w�ɮ� ���|+�W��
                            .uBankCode = strBank
                            .uOwner = GetOwner
                            .uCompCode = gilCompCode.GetCodeNo
                            .ugcnGi = gcnGi
                            .CheckMediaTextShow
                                 'blnReturn = .uReturnOK '�^�ǬO�_�w�g���� TRUE=���槹���S�����D  FALSE=���椤�Φ����D(�ݭn����)
                            blnErrNum = .ublnError '�^�ǬO�_��ERROR����� TRUE=�����~���  FALSE=�L���~���
                        End With
                        If blnErrNum Then
                            Shell "Notepad " & objAgency.uProcErrText, vbNormalFocus
                        End If
                    End If
                  blnUnload = True
              End If
          End If
    End If
      '7207 Add Data into SO054 Table By Kin 2016/03/08
        LastTime = timeGetTime / 1000
        LogToSO054 Replace(UCase(Screen.ActiveForm.Name), "FRM", ""), _
                    garyGi(0), Format(NowTime, "YYYYMMDD HHMMSS"), SecToHMS(LastTime - startTime), strChooseString, RightNow
lExit:
      On Error Resume Next
      blnUnload = True
      Set objAgency = Nothing
      Set objChkErr = Nothing
      DoEvents
      Screen.MousePointer = vbDefault
     
  Exit Sub
chkErr:
    blnUnload = True
    If Err.Number = 20 Then Resume Next
    If Err.Number = 91 Then
       MsgBox "��b�{���W�ٿ��~�Ϊ̸ӻȦ�S����b��Ʋ��ͥ\��!", vbExclamation, "ĵ�i..."
       Exit Sub
    End If
    Call ErrSub(Me.Name, "cmdOk_Click")
End Sub

Private Function subChoose() As Boolean
  On Error GoTo chkErr
  Dim strAddr As String

   subChoose = False
   strChoose = ""
    If gilCompCode.GetCodeNo <> "" Then subAnd ("A.CompCode=" & gilCompCode.GetCodeNo)
    '#3728
    'If gilServiceType.GetCodeNo <> "" Then subAnd ("A.ServiceType='" & gilServiceType.GetCodeNo & "'")
    If gdaShouldDate1.GetValue <> "" Then subAnd ("A.ShouldDate >= To_Date(" & gdaShouldDate1.GetValue & ",'YYYYMMDD')")
    If gdaShouldDate2.GetValue <> "" Then subAnd ("A.ShouldDate < To_Date(" & gdaShouldDate2.GetValue & ",'YYYYMMDD') +1")
    '*************************************************************************************************************************
    '#2672 ���ͮɶ��אּ��ɤ��� By Kin 2008/05/12
'    If gdaCreateTime1.GetValue <> "" Then subAnd ("A.CreateTime >= To_Date(" & gdaCreateTime1.GetValue & ",'YYYYMMDD')")
'    If gdaCreateTime2.GetValue <> "" Then subAnd ("A.CreateTime < To_Date(" & gdaCreateTime2.GetValue & ",'YYYYMMDD') +1")
    If gitCreateTime1.GetValue <> "" Then subAnd ("A.CreateTime>=To_Date(" & gitCreateTime1.GetValue & ",'YYYYMMDDHH24MISS')")
    If gitCreateTime2.GetValue <> "" Then subAnd ("A.CreateTime<To_Date(" & gitCreateTime2.GetValue & ",'YYYYMMDDHH24MISS')+1")
    '*************************************************************************************************************************
    If Len(gmdCMCode.GetQueryCode & "") > 0 Then
        subAnd ("A.CMCode In (" & gmdCMCode.GetQueryCode & ")")
    End If
    
    'If gmdCMCode.GetQryStr <> "" Then subAnd ("A.CMCode " & gmdCMCode.GetQueryCode)
    
    
    If gmdAreaCode.GetQryStr <> "" Then subAnd ("A.AreaCode " & gmdAreaCode.GetQryStr)
    If gmdServCode.GetQryStr <> "" Then subAnd ("A.ServCode " & gmdServCode.GetQryStr)
    '#1659  950517��"���O�H��"�אּ"�즬�O�H��"�A�t�s�W�@"���O�H��"
    If gmdClctEn.GetQryStr <> "" Then subAnd ("A.ClctEn " & gmdClctEn.GetQryStr)
    If gmdOldClctEn.GetQryStr <> "" Then subAnd ("A.OldClctEn " & gmdOldClctEn.GetQryStr)
    '#5683 �W�[ú�I���O���� By Kin 2010/08/06
    If gmdPayType.GetQryStr <> "" Then subAnd ("A.PayKind " & gmdPayType.GetQryStr)
    '#4388 �W�[�Ƚs���� By Kin 2009/04/30
    If txtCustId.Text <> "" Then
        Call TimetxtCustId(txtCustId)
    End If
    
    If gmdBillType.GetQryStr <> "" Then
       subAnd ("SubStr(A.BillNo,7,1) In (" & gmdBillType.GetQueryCode & ")")
    Else
       subAnd ("SubStr(A.BillNo,7,1) In ('B','T','I','M','P')")
    End If
    If gimCreateEn.GetQryStr <> "" Then subAnd ("A.CreateEn " & gimCreateEn.GetQryStr)
    
    '���ͤj��
    If optMduYes Then
       '���ͤ@���
       If chkOther.Value = 1 Then
          If gmdMduid.GetQryStr <> "" Then subAnd ("(A.MduId " & gmdMduid.GetQryStr & " Or A.MduId Is Null) ")
       Else
       '�����ͤ@���
          If gmdMduid.GetQryStr <> "" Then
             subAnd ("A.MduId " & gmdMduid.GetQryStr)
          Else
             subAnd ("A.MduId Is Not Null")
          End If
       End If
    '�����ͤj��
    Else
       '���ͤ@���
       If chkOther.Value = 1 Then
          If gmdMduid.GetQryStr <> "" Then
             subAnd ("(Not A.MduId " & gmdMduid.GetQryStr & " Or A.MduId Is Null) ")
          Else
             subAnd ("A.MduId Is Null")
          End If
       Else
       '�����ͤ@���
          If gmdMduid.GetQryStr <> "" Then
             subAnd ("Not A.MduId " & gmdMduid.GetQryStr)
          Else
             MsgBox "������d�L���!!", vbExclamation, "�T��"
             Exit Function
          End If
       End If
    End If
    
    If gimUCCode.GetQryStr <> "" Then subAnd ("A.UCCODE " & gimUCCode.GetQryStr)
    '#7049 ���CD068.CitemCode By Kin 2015/07/13
    'If gimCitemCode.GetQryStr <> "" Then subAnd ("A.CitemCode " & gimCitemCode.GetQryStr)
    If Len(gimCitemCode.GetQryStr & "") > 0 Then
        subAnd ("A.CitemCode In (Select CitemCode From " & GetOwner & "CD068A " & _
                                                                    " Where BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "')")
    End If
     '#6441 �W�[�O�_�n�P�_���B�p��0 By Kin 2013/03/19
    'If chkZero.Value = 0 Then subAnd "A.ShouldAmt > 0"
    
    pChooseMultiAcc = strChoose   '' ��ثe��� SO033 �������
    '#4002
    If gimCustStatus.GetQryStr <> "" Then subAnd ("D.CustStatusCode " & gimCustStatus.GetQryStr)
    'If gimCustStatus.GetQryStr <> "" Then subAnd ("C.CustStatusCode " & gimCustStatus.GetQryStr)
    If gmdCustClass.GetQryStr <> "" Then subAnd ("B.ClassCode1 " & gmdCustClass.GetQryStr)
     '#7760 By Kin 2018/05/10
    If gimAMduId.GetQueryCode <> "" Then subAnd ("B.AMduId In (" & gimAMduId.GetQueryCode & ")")
    If chkFubonIntegrate.Value = 1 And e_CreditCardType = CardType_Fubon Then
'        subAnd (" ( substr(A.AccountNo,1,6) in (Select CodeNo From " & GetOwner & "CD143 Where Length(CodeNo)= 6 )" & _
'                             " or substr(A.AccountNo,1,7) in (Select CodeNo From " & GetOwner & "CD143 Where Length(CodeNo)= 7 ))")
    
    End If
    If e_CreditCardType = CardType_Fubon And chkFubonIntegrate.Value = 0 Then
        '#8537 By Kin 2019/12/17
        If gmdCMCode.GetQueryCode = "" Then
            subAnd (" A.CMCode  In (Select CodeNo From " & GetOwner & "CD031 Where Nvl(StopFlag,0) = 0 And Nvl(RefNo,0) <> 5) ")
        End If
    End If
'    If optChargeAddr.Value Then
'       strAddr = "���O�a�}"
'    Else
'       strAddr = "�˾��a�}"
'    End If
    subChoose = True
    strChooseString = "���q�O�@    :" & subSpace(gilCompCode.GetDescription) & ";" & _
                      "�A�����O    :" & subSpace(gilServiceType.GetDescription) & ";" & _
                      "�ӿ�Ȧ�    :" & subSpace(gmBank.GetDispStr) & ";" & _
                      "���O����@�@:" & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & _
                      "���O�覡    :" & subSpace(gmdCMCode.GetDispStr) & ";" & _
                      "��F��      :" & subSpace(gmdAreaCode.GetDispStr) & ";" & _
                      "�A�Ȱ�      :" & subSpace(gmdServCode.GetDispStr) & ";" & _
                      "�Ȥ᪬�A    :" & subSpace(gimCustStatus.GetDispStr) & ";" & _
                      "�Ȥ����O�@�@:" & subSpace(gmdCustClass.GetDispStr) & ";" & _
                      "���O��      :" & subSpace(gmdClctEn.GetDispStr) & ";" & _
                      "�즬�O��    :" & subSpace(gmdOldClctEn.GetDispStr) & ";" & _
                      "������O    :" & subSpace(gmdBillType.GetDispStr) & ";" & _
                      "���ͤH��    :" & subSpace(gimCreateEn.GetDispStr) & ";" & _
                      "�j�ӦW��    :" & subSpace(gmdMduid.GetDispStr)
  
  Exit Function
chkErr:
    Call ErrSub(Me.Name, "subChoose")
End Function

Private Sub subAnd(strCH As String)
  On Error GoTo chkErr
    '�O�ťե[And
    If strChoose <> "" Then
       strChoose = strChoose & " And " & strCH
    Else
       strChoose = strCH
    End If
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subAnd")
End Sub

Private Function AddChar(ByVal vStr As String, ByVal vChar As String, Optional OverWrite As Boolean = True) As String
On Error GoTo chkErr
    Static strChar As String
        If OverWrite Then strChar = ""
        If InStr(1, vStr, vChar) > 0 Then
            strChar = strChar & IIf(strChar <> "", ",'", "'") & vChar & "'"
        End If
        AddChar = strChar
Exit Function
chkErr:
    Call ErrSub(Me.Name, "ADDChar")
End Function
Private Sub Form_Activate()
On Error GoTo chkErr
  Dim strSQL As String
  Dim rsTmp As New ADODB.Recordset
  
    strSQL = "Select PrgName From " & GetOwner & "CD018 Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'"
    If rsTmp.State = adStateOpen Then rsTmp.Close
'    rsTmp.CursorLocation = adUseClient
'    rsTmp.Open strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
    If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then Exit Sub
    If rsTmp.EOF Then
       MsgBox "�г]�w�Ȧ�N�X����b�{���W��!", vbExclamation, "ĵ�i"
       Unload Me
    End If
    rsTmp.Close
    Set rsTmp = Nothing
    Screen.MousePointer = vbDefault
  Exit Sub
chkErr:
  ErrSub Me.Name, "Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Not ChkGiList(KeyCode) Then Exit Sub
    If KeyCode = vbKeyF2 Then cmdOk.Value = True: KeyCode = 0
    If KeyCode = vbKeyEscape Then Unload Me: KeyCode = 0
End Sub

Private Sub Form_Load()
    On Error GoTo chkErr
    
    '' 2003/10/07 �o�@�q�Ȯɧ@���ե� �]��true�A�H��ݦp�G�P�_�O�s���i�ҭn����ӻ�
        gitCreateTime1.ShowSecond
        gitCreateTime2.ShowSecond
        If Alfa2 Then
            Call GetGlobal
        End If
        blnUnload = True
        strINIpath1 = GetIniPath1
        strErrPath = ReadGICMIS1("ErrLogPath")
        blnPaynowFlag = GetPaynowFlag
        CmRefNo5 = GetRsValue("Select CodeNo From " & GetOwner & "CD031 Where Nvl(StopFlag,0) = 0 And RefNo = 5 ")
        Call subGil
        Call subGim
        Call getDefault
        chkChinaBank.Enabled = GetUserPriv("SO3272", "SO32721")
        '*************************************************************************************************
        '#3728 �j���榡�H�h�C�^�X�b�覡 By Kin 2008/03/25
        'intPara24 = GetSystemParaItem("Para24", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043")
        cboBillHeadFmt.Enabled = True
        
        
        CboAddItem
        '*************************************************************************************************
        chkErr.Enabled = gcnGi.Execute("Select CheckText From " & GetOwner & "SO041  WHERE CompCode = " & garyGi(9))(0)
        
        If chkErr.Enabled Then chkErr.Value = 1 Else chkErr.Value = 0
   Exit Sub
chkErr:
   Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub subGil()
  On Error GoTo chkErr
  
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12
        SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
        SetgiList gilFubonCMCode, "CodeNo", "Description", "CD031", , , , , , , " Where RefNo = 5", True
        
        'SetgiList gilBank, "CodeNo", "Description", "CD018", , , , , 3, 24, " Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'"
   Exit Sub
chkErr:
   ErrSub Me.Name, "subGil"
End Sub

Private Sub subGim()

Dim strWhereBank As String

  On Error GoTo chkErr
  
  ''2003/10/29  �H�U���Χ@ gmBank ���]�w
  
'    Select Case e_CreditCardType
'        Case 0  '' �w�]
'             strWhereBank = "CREDITCARD"
'        Case 1  '' ����H�U
'             strWhereBank = "CREDITCARDChiatrust"
'     End Select
'    Call SetgiMulti(gmBank, "CodeNo", "Description", "CD018", "�N�X", "�W��", " Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'")
    
    Call SetgiMulti(gmdCMCode, "CodeNo", "Description", "CD031", "�N�X", "�W��")
    Call SetgiMulti(gmdAreaCode, "CodeNo", "Description", "CD001", "�N�X", "�W��")
    Call SetgiMulti(gmdServCode, "CodeNo", "Description", "CD002", "�N�X", "�W��")
    Call SetgiMulti(gimCustStatus, "CodeNo", "Description", "CD035", "�N�X", "�W��")
    Call SetgiMulti(gmdCustClass, "CodeNo", "Description", "CD004", "�N�X", "�W��")
    Call SetgiMulti(gmdClctEn, "EmpNo", "EmpName", "CM003", "�N�X", "�W��")
    Call SetgiMulti(gmdOldClctEn, "EmpNo", "EmpName", "CM003", "�N�X", "�W��")
    Call SetgiMultiAddItem(gmdBillType, "B,T,I,M,P", "���O��,�{�ɦ��O��,�˾���,���׳�,�������", "�N�X", "�W��")
    '#5683 �W�[ú�I���O By Kin 2010/08/06
    Call SetgiMulti(gmdPayType, "CodeNo", "Description", "CD112", "�N�X", "�W��")
    'Call SetgiMultiAddItem(gmdPayType, "0,1", "�w�I��,�{�I��", "�N�X", "�W��")
    '#4388 �L�o�w���λP�ѦҸ�=3 By Kin 2009/04/30
    '#5218 �n��Nvl���覡 By Kin 2009/08/05
    Call SetgiMulti(gimUCCode, "CodeNo", "Description", "CD013", "�N�X", "�W��", " Where Nvl(REFNO,0) NOT IN(3,7) AND NVL(PAYOK,0)=0 ", True)  '' ������]
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "�N�X", "�W��")  ' '���O����

    With gmdBillType
          .SetDispStr "���O��,�{�ɦ��O��"
          .SetQueryCode "B,T"
'          .SetDispStr "���O��,�{�ɦ��O��,�˾���,���׳�,�������"
'          .SetQueryCode "B,T,I,M,P"
    End With
    '#5925 ú�O���O�W�[�w�]�� By Kin 2011/04/12
    If blnPaynowFlag Then
        With gmdPayType
            Select Case Val(GetRsValue("SELECT NVL(PayKindDefault,0) FROM " & GetOwner & "SO041", gcnGi) & "")
                Case 0
                    .SelectAll
                Case 1
                    .SetQueryCode "0"
                Case 2
                    .SetQueryCode "1"
            End Select
        End With
    Else
        gmdPayType.Clear
        gmdPayType.Enabled = False
    End If
    Call SetgiMulti(gimCreateEn, "EmpNo", "EmpName", "CM003", "�N�X", "�W��")
    Call SetgiMulti(gmdMduid, "Mduid", "Name", "SO017", "�N�X", "�W��")
    '#7760 By Kin 2018/05/10
    Call SetgiMulti(gimAMduId, "Mduid", "Name", "SO202", "�N�X", "�W��")
    gimCustStatus.SetQueryCode ("1")
   ''   gmdBillType.SelectAll
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub getDefault()
  On Error GoTo chkErr
     gilCompCode.SetCodeNo garyGi(9)
     gilCompCode.Query_Description
  Exit Sub
chkErr:
  ErrSub Me.Name, "getDefault"
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo chkErr
Dim strErrMsg As String


    IsDataOk = False
        If Len(gilCompCode.GetCodeNo & "") = 0 Then strErrMsg = "���q�O": gilCompCode.GetCodeNo: GoTo Warning
        '#3728 �j���h�C�^�X�b,�ҥH�����A�ȧO�ﶵ By Kin 2008/03/25
        'If Len(gilServiceType.GetCodeNo & "") = 0 Then strErrMsg = "�A�����O": gilServiceType.GetCodeNo: GoTo Warning
     ''   If gilBank.GetDescription = "" Then strErrMsg = "�ӿ�Ȧ�": gilBank.SetFocus: GoTo Warning
        If chkChinaBank.Value = 0 Then
             If gmBank.GetQueryCode = "" Then strErrMsg = "�ӿ�Ȧ�": gmBank.SetFocus: GoTo Warning
        End If
        If gdaShouldDate1.GetValue = "" Then strErrMsg = "���O����_�l��": gdaShouldDate1.SetFocus: GoTo Warning
        If gdaShouldDate2.GetValue = "" Then strErrMsg = "���O����I���": gdaShouldDate2.SetFocus: GoTo Warning
        If DateDiff("d", gdaShouldDate1.Text, gdaShouldDate2.Text) < 0 Then MsgBox "���O�I��餣�o�p�󦬶O�_�l��I", vbExclamation, "�T���I": Exit Function
        '#3728 �j���H�h�C�^�X�b By Kin 2008/03/25
        If cboBillHeadFmt.Text = "" Then strErrMsg = "�h�b�Უ�ͨ̾ڳ]�w": cboBillHeadFmt.SetFocus: GoTo Warning
'        If chkFubonIntegrate.Value = 1 Then
'            If e_CreditCardType = CardType_Fubon Then
'                If Len(gilFubonCMCode.GetCodeNo & "") = 0 Then strErrMsg = "�]�����x���O�覡": gilFubonCMCode.SetFocus: GoTo Warning
'
'            End If
'        End If
        If e_CreditCardType = CardType_Fubon And chkFubonIntegrate.Value = 1 Then
            If gmdCMCode.GetQueryCode = "" Then strErrMsg = "���O�覡": gmdCMCode.SetFocus: GoTo Warning
        End If
    IsDataOk = True
    Exit Function
Warning:
    Call MsgMustBe(strErrMsg)
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo chkErr
    If blnUnload Then
       Set objAgency = Nothing
    Else
       Cancel = True
    End If
    Call CloseRecordset(rsBankData)
    Call CloseRecordset(rsBillHeadFmt)
    Screen.MousePointer = vbDefault
    
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_QueryUnload")
End Sub

Private Sub Frame4_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub gdaCreateTime1_GotFocus()
  On Error Resume Next
    If gdaCreateTime1.GetValue <> "" Then Exit Sub
    gdaCreateTime1.SetValue Date
End Sub

Private Sub gdaCreateTime1_Validate(Cancel As Boolean)
On Error GoTo chkErr
   If Not IsDate(gdaCreateTime1.Text) Then Exit Sub
   If gdaCreateTime1.GetValue <> "" Then
      If gdaCreateTime2.GetValue = "" Then gdaCreateTime2.SetValue GetLastDate(gdaCreateTime1.GetValue(True))
   End If
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "gdaCreateTime1_Validate")
End Sub

Private Sub gdaCreateTime2_GotFocus()
  On Error Resume Next
    If gdaCreateTime2.GetValue <> "" Then Exit Sub
    gdaCreateTime2.SetValue GetLastDate(Date)
    
End Sub

Private Sub gdaCreateTime2_Validate(Cancel As Boolean)
   On Error Resume Next
     If gdaCreateTime1.GetValue = "" Or gdaCreateTime2.GetValue = "" Then Exit Sub
     If Not IsDate(gdaCreateTime2.Text) Then Exit Sub
     If DateDiff("d", gdaCreateTime1.GetValue(True), gdaCreateTime2.GetValue(True)) < 0 Then
        MsgBox "���ͮɶ��I��餣�i�p�󲣥ͮɶ��_�l��!", vbExclamation, "ĵ�i"
        gdaCreateTime2.SetValue gdaCreateTime1.GetValue
        Cancel = True
     End If
End Sub

Private Sub gdaShouldDate1_GotFocus()
  On Error Resume Next
    If gdaShouldDate1.GetValue <> "" Then Exit Sub
    gdaShouldDate1.SetValue Date
End Sub

Private Sub gdaShouldDate1_Validate(Cancel As Boolean)
On Error GoTo chkErr
     
   If Not IsDate(gdaShouldDate1.Text) Then Exit Sub
   If gdaShouldDate1.GetValue <> "" Then
      If gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue GetLastDate(gdaShouldDate1.GetValue(True))
   End If
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "gdaShouldDate1_Validate")
End Sub

Private Sub gdaShouldDate2_GotFocus()
  On Error Resume Next
    If gdaShouldDate2.GetValue <> "" Then Exit Sub
    gdaShouldDate2.SetValue GetLastDate(Date)
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
   On Error Resume Next
     If gdaShouldDate1.GetValue = "" Or gdaShouldDate2.GetValue = "" Then Exit Sub
     If Not IsDate(gdaShouldDate2.Text) Then Exit Sub
     If DateDiff("d", gdaShouldDate1.GetValue(True), gdaShouldDate2.GetValue(True)) < 0 Then
        MsgBox "�����I��餣�i�p�������_�l��!", vbExclamation, "ĵ�i"
        gdaShouldDate2.SetValue gdaShouldDate1.GetValue
        Cancel = True
     End If
End Sub

Private Sub subInsert()
On Error GoTo chkErr
  Dim strSQL As String
  
    strSQL = "INSERT INTO " & GetOwner & "SO097 (EMPNO,PRGNAME,RUNDATETIME,SECOND,SELECTION) " & _
                   "VALUES ('" & garyGi(0) & "','SO3273A',sysdate,'" & _
                   objAgency.utime & "��','" & Replace(strChooseString, "'", "") & "')"
    gcnGi.Execute (strSQL)
   Exit Sub
chkErr:
   ErrSub Me.Name, "subInsert"
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
    
If ChgComp(gilCompCode, "SO3270", "SO3272") = False Then Exit Sub
    
    Call subGil
    Call subGim
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
'    gilServiceType.ListIndex = 1
    
    Call GiMultiFilter(gmBank, , gilCompCode.GetCodeNo)
''    gmBank.Filter = gmBank.Filter & IIf(gmBank.Filter = "", " Where ", " And ") & " UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'  AND (BANKID  <>   '" & "TCB" & "'   OR BANKID IS NULL) "

    Call GiMultiFilter(gmdAreaCode, , gilCompCode.GetCodeNo)
    Call GiMultiFilter(gmdMduid, , gilCompCode.GetCodeNo)
    Call GiMultiFilter(gmdServCode, , gilCompCode.GetCodeNo)
    
   If SetCreditCardTypeCbo = False Then cmdOk.Enabled = False

End Sub

Private Sub gilServiceType_Change()
    On Error Resume Next
        Call GiMultiFilter(gmdCMCode, gilServiceType.GetCodeNo)
        Call GiMultiFilter(gmdCustClass, gilServiceType.GetCodeNo)
End Sub

Private Function SetCreditCardTypeCbo() As Boolean

  Dim strSQL As String
  Dim rsTmp As New ADODB.Recordset
  Dim lbn
  Dim lngIndex As Long
  On Error GoTo chkErr
  
  For lngIndex = 0 To cboCreditCardType.ListCount - 1
       cboCreditCardType.RemoveItem (0)
  Next
    lngIndex = 0
    strSQL = "SELECT DISTINCT PrgName FROM " & GetOwner & "CD018 WHERE UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'  AND COMPCODE =" & gilCompCode.GetCodeNo & "  AND STOPFLAG = 0 "
    Call GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenDynamic)
    Dim blnAdd As Boolean   ''  �P�_�p�G���[�J�H�Υd���زM�檺�ܡA�h�N��]��flase �A�o�˩��U��lngIndex�~�i�H�[ 1  ���M�U����ƪ��j��|�X�{ ���޶W�X�}�C�d��
    blnAdd = False
  
    Do While Not (rsTmp.EOF Or rsTmp.BOF)
        Select Case UCase(rsTmp("PrgName"))
            Case "CREDITCARDUNION"
                 cboCreditCardType.AddItem (" �p�X/���ذӻ� �H�Υd")
                 cboCreditCardType.ItemData(lngIndex) = 0
                 blnAdd = True
            Case "CREDITCARDCHINATRUST"
                 cboCreditCardType.AddItem (" ����H�U �H�Υd")
                 cboCreditCardType.ItemData(lngIndex) = 1
                blnAdd = True
            Case "CREDITCARDHSBC"
                 cboCreditCardType.AddItem (" ���׻Ȧ� �H�Υd")
                 cboCreditCardType.ItemData(lngIndex) = 2
                 blnAdd = True
            '#3531 �h�s�W�@�Ӥ���H�U(�s) By Kin 2007/10/31
            Case "CREDITCARDTRUSTNEW"
                cboCreditCardType.AddItem ("����H�U �H�Υd(�s)")
                cboCreditCardType.ItemData(lngIndex) = 3
                blnAdd = True
            '#4002
            Case "CREDITCARDCITIBANK"
                cboCreditCardType.AddItem ("��X�H�Υd�榡")
                cboCreditCardType.ItemData(lngIndex) = 4
                blnAdd = True
            '#5494 �s�W�ŷs��� By Kin 2009/09/15
            Case "CREDITCARDNEWEB"
                cboCreditCardType.AddItem ("�ŷs��ީw���w�B����")
                cboCreditCardType.ItemData(lngIndex) = 5
                blnAdd = True
            '#5609 �p�X�H�Υd�妸 By Kin 2010/06/03
            Case "CREDITCARDUNION1"
                cboCreditCardType.AddItem (" �p�X�H�Υd�妸�榡")
                cboCreditCardType.ItemData(lngIndex) = 6
                blnAdd = True
            '#7765 By Kin 2018/05/08
            Case "CREDITCARDFUBON"
                cboCreditCardType.AddItem (" �I���H�Υd�榡")
                cboCreditCardType.ItemData(lngIndex) = 7
                blnAdd = True
          End Select
          If blnAdd = True Then lngIndex = lngIndex + 1
          blnAdd = False
         rsTmp.MoveNext
     Loop
    rsTmp.Close
    Set rsTmp = Nothing
      If lngIndex = 0 Then
         MsgBox "�Ȧ�O��ƥ��]�A�г]�w�����Ȧ�O�N�X����A���s����  !!"
         SetCreditCardTypeCbo = False
    Else
         cboCreditCardType.ListIndex = 0
          SetCreditCardTypeCbo = True
    End If
    Screen.MousePointer = vbDefault
   
Exit Function
chkErr:
  ErrSub Me.Name, "SetCreditCardTypeCbo"
End Function
'#2672 ���ͮɶ��אּ��ɤ��� By Kin 2008/05/12
Private Sub gitCreateTime1_GotFocus()
  On Error Resume Next
    If gitCreateTime1.GetValue <> "" Then Exit Sub
    gitCreateTime1.SetValue Now
    
End Sub
'#2672 ���ͮɶ��אּ��ɤ��� By Kin 2008/05/12
Private Sub gitCreateTime1_Validate(Cancel As Boolean)
  On Error GoTo chkErr
   If Not IsDate(gitCreateTime1.Text) Then Exit Sub
   If gitCreateTime1.GetValue <> "" Then
      If gitCreateTime2.GetValue = "" Then gitCreateTime2.SetValue GetLastDate(gitCreateTime1.GetValue(True)) & "235959"
   End If
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "gitCreateTime1_Validate")
End Sub

'#2672 ���ͮɶ��אּ��ɤ��� By Kin 2008/05/12
Private Sub gitCreateTime2_GotFocus()
  On Error Resume Next
    If gitCreateTime2.GetValue <> "" Then Exit Sub
    gitCreateTime2.SetValue GetLastDate(Date) & "235959"
    
End Sub
'#2672 ���ͮɶ��אּ��ɤ��� By Kin 2008/05/12
Private Sub gitCreateTime2_Validate(Cancel As Boolean)
  On Error Resume Next
    If gitCreateTime1.GetValue = "" Or gitCreateTime2.GetValue = "" Then Exit Sub
    If Not IsDate(gitCreateTime2.Text) Then Exit Sub
    If DateDiff("s", gitCreateTime1.GetValue(True), gitCreateTime2.GetValue(True)) < 0 Then
        MsgBox "���ͮɶ��I��餣�i�p�󲣥ͮɶ��_�l��!", vbExclamation, "ĵ�i"
        Cancel = True
        
    End If
End Sub

Private Sub gmBank_Change()
Call SetgiMulti(gmdCMCode, "CodeNo", "Description", "CD031", "�N�X", "�W��", "")

End Sub

Private Sub CboAddItem()
    On Error Resume Next
        'If cboBillHeadFmt.ListCount > 0 Then Exit Sub
        '#5919 �W�[RowId,�H�K���D�O���@��CD068.SPCNO By Kin 2011/06/03
        If Not GetRS(rsBillHeadFmt, "Select RowId, BillHeadFmt,CitemCodeStr From " & GetOwner & "CD068") Then Exit Sub
        Do While Not rsBillHeadFmt.EOF
            cboBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
            rsBillHeadFmt.MoveNext
        Loop
End Sub

Private Sub txtCustId_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 44 Or KeyAscii = 8 Or KeyAscii = 45 Then
        If KeyAscii = 44 Or KeyAscii = 45 Then
            If Asc(Right(" " & txtCustId.Text, 1)) = 44 Or Asc(Right(" " & txtCustId.Text, 1)) = 45 Or txtCustId = "" Then KeyAscii = 9
        End If
        If Not ChkMaxLengthOK(txtCustId) Then
           If Not (KeyAscii = 44 Or KeyAscii = 8 Or KeyAscii = 45) Then KeyAscii = 9
        End If
    Else
        KeyAscii = 9
    End If
End Sub
