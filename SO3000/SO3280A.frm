VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO3280A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "���O�渹�վ� [SO3280A]"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3280A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame gmdPayType 
      Caption         =   "�@�α���"
      ForeColor       =   &H00FF0000&
      Height          =   675
      Left            =   150
      TabIndex        =   44
      Top             =   6120
      Width           =   7065
      Begin Gi_Multi.GiMulti gimPayType 
         Height          =   375
         Left            =   90
         TabIndex        =   18
         Top             =   240
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   661
         ButtonCaption   =   "ú  �I  ��  �O"
         DataType        =   2
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ڦX�ֽվ㶵��"
      ForeColor       =   &H00FF0000&
      Height          =   885
      Left            =   150
      TabIndex        =   42
      Top             =   6810
      Width           =   7065
      Begin VB.CheckBox chkUpdPrtSNo 
         Caption         =   "�վ�L��Ǹ�"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3900
         TabIndex        =   51
         Top             =   540
         Width           =   1905
      End
      Begin VB.CheckBox chkUpdBillCloseDate 
         Caption         =   "�վ�ú�O�I����"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1920
         TabIndex        =   50
         Top             =   540
         Width           =   1905
      End
      Begin VB.CheckBox chkUpdUCcode 
         Caption         =   "�վ㥼����]"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   150
         TabIndex        =   49
         Top             =   540
         Width           =   1485
      End
      Begin VB.CheckBox chkUpdCreateEn 
         Caption         =   "����H��"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5580
         TabIndex        =   22
         Top             =   240
         Width           =   1545
      End
      Begin VB.CheckBox chkCmCode 
         Caption         =   "���O�覡"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3570
         TabIndex        =   21
         Top             =   240
         Width           =   1185
      End
      Begin VB.CheckBox chkCreateTime 
         Caption         =   "���ͤ��"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1920
         TabIndex        =   20
         Top             =   240
         Width           =   1185
      End
      Begin VB.CheckBox chkShouldDate 
         Caption         =   "�������"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   150
         TabIndex        =   19
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "�W���έp���G"
      Height          =   375
      Left            =   2130
      TabIndex        =   24
      Top             =   7770
      Width           =   1365
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2. �T�w"
      Default         =   -1  'True
      Height          =   375
      Left            =   300
      TabIndex        =   23
      Top             =   7800
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   5850
      TabIndex        =   25
      Top             =   7770
      Width           =   1365
   End
   Begin VB.Frame fraData 
      Height          =   6045
      Left            =   120
      TabIndex        =   27
      ToolTipText     =   "���\��u�w�良����T�氵�X��"
      Top             =   30
      Width           =   7155
      Begin VB.ComboBox cboTBillHeadFmt 
         Height          =   315
         Left            =   2040
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   47
         Top             =   3720
         Width           =   4995
      End
      Begin CS_Multi.CSmulti gimTCitemCode 
         Height          =   375
         Left            =   180
         TabIndex        =   46
         Top             =   4500
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   661
         ButtonCaption   =   "��  �O  ��  ��"
         IsReadOnly      =   -1  'True
      End
      Begin CS_Multi.CSmulti gimBCitemCode 
         Height          =   375
         Left            =   180
         TabIndex        =   45
         Top             =   2640
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   661
         ButtonCaption   =   "B�榬�O����"
         IsReadOnly      =   -1  'True
      End
      Begin VB.CheckBox chkShouldDateMerge 
         Caption         =   "�ۦP��������~�ֳ�"
         Enabled         =   0   'False
         Height          =   225
         Left            =   180
         TabIndex        =   4
         Top             =   1170
         Width           =   2235
      End
      Begin VB.TextBox txtCustId 
         Height          =   315
         Left            =   4380
         MaxLength       =   200
         TabIndex        =   5
         Top             =   1500
         Width           =   2610
      End
      Begin Gi_Time.GiTime gdtCreateTime2 
         Height          =   345
         Left            =   3270
         TabIndex        =   11
         Top             =   2250
         Width           =   1845
         _ExtentX        =   3254
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
      Begin Gi_Time.GiTime gdtCreateTime1 
         Height          =   315
         Left            =   1050
         TabIndex        =   10
         Top             =   2280
         Width           =   1845
         _ExtentX        =   3254
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
      Begin CS_Multi.CSmulti gimBCitemCode1 
         Height          =   345
         Left            =   6150
         TabIndex        =   26
         Top             =   5640
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         ButtonCaption   =   "B�榬�O����"
      End
      Begin VB.ComboBox cboBillHeadFmt 
         Height          =   315
         Left            =   1980
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   4995
      End
      Begin VB.CheckBox chkCombine 
         Caption         =   "�������즬�O�檺��ƬO�_�j���ഫ"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   5700
         Value           =   1  '�֨�
         Width           =   3495
      End
      Begin Gi_Date.GiDate gdaProcessDate 
         Height          =   315
         Left            =   3285
         TabIndex        =   3
         Top             =   1095
         Width           =   1185
         _ExtentX        =   2090
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
      Begin prjGiList.GiList gilServiceType 
         Height          =   315
         Left            =   1065
         TabIndex        =   1
         Top             =   660
         Width           =   3570
         _ExtentX        =   6297
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
         FldWidth2       =   2500
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilCompCode 
         Height          =   315
         Left            =   1065
         TabIndex        =   0
         Top             =   240
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
         F2Corresponding =   -1  'True
      End
      Begin Gi_YM.GiYM gymClctYM1 
         Height          =   330
         Left            =   1065
         TabIndex        =   6
         Top             =   1905
         Width           =   900
         _ExtentX        =   1588
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
      Begin Gi_YM.GiYM gymClctYM2 
         Height          =   330
         Left            =   2385
         TabIndex        =   7
         Top             =   1905
         Width           =   900
         _ExtentX        =   1588
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
      Begin CS_Multi.CSmulti gimBUCCode 
         Height          =   345
         Left            =   180
         TabIndex        =   12
         Top             =   3000
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   609
         ButtonCaption   =   "�� �� �O �� �]"
      End
      Begin Gi_Date.GiDate gdaShouldDate2 
         Height          =   345
         Left            =   2730
         TabIndex        =   14
         Top             =   4140
         Width           =   1125
         _ExtentX        =   1984
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
         Left            =   1140
         TabIndex        =   13
         Top             =   4140
         Width           =   1125
         _ExtentX        =   1984
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
      Begin CS_Multi.CSmulti gimTUCCode 
         Height          =   345
         Left            =   180
         TabIndex        =   15
         Top             =   4890
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   609
         ButtonCaption   =   "�� �� �O �� �]"
      End
      Begin Gi_Multi.GiMulti gimBillType 
         Height          =   375
         Left            =   180
         TabIndex        =   16
         Top             =   5280
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   661
         ButtonCaption   =   "��  ��  ��  �O"
         DataType        =   2
         DIY             =   -1  'True
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin Gi_Date.GiDate gdaOweDate1 
         Height          =   345
         Left            =   4380
         TabIndex        =   8
         Top             =   1890
         Width           =   1125
         _ExtentX        =   1984
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
      Begin Gi_Date.GiDate gdaOweDate2 
         Height          =   345
         Left            =   5880
         TabIndex        =   9
         Top             =   1890
         Width           =   1125
         _ExtentX        =   1984
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "�h�b�Უ�ͨ̾ڳ]�w"
         Height          =   195
         Left            =   210
         TabIndex        =   48
         Top             =   3780
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "�Ȥ�s��(�H,�۹j�ΥH-���d��)"
         Height          =   195
         Left            =   1770
         TabIndex        =   43
         Top             =   1590
         Width           =   2565
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   3000
         TabIndex        =   41
         Top             =   2340
         Width           =   195
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "���ͮɶ�"
         Height          =   195
         Left            =   180
         TabIndex        =   40
         Top             =   2340
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   5610
         TabIndex        =   39
         Top             =   1950
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   195
         Left            =   3510
         TabIndex        =   38
         Top             =   1980
         Width           =   780
      End
      Begin VB.Label lblBillHeadFmt 
         AutoSize        =   -1  'True
         Caption         =   "�h�b�Უ�ͨ̾ڳ]�w"
         Height          =   195
         Left            =   150
         TabIndex        =   37
         Top             =   780
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblShouldDate 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   4215
         Width           =   780
      End
      Begin VB.Label lblunti2 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   2400
         TabIndex        =   35
         Top             =   4215
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�Q�X�ֳ�ڱ���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   34
         Top             =   3420
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�X�ֳ�ڱ���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   33
         Top             =   1560
         Width           =   1440
      End
      Begin VB.Label lblServiceType 
         AutoSize        =   -1  'True
         Caption         =   "�A�����O"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   720
         Width           =   780
      End
      Begin VB.Label lblCompCode 
         AutoSize        =   -1  'True
         Caption         =   "���q�O"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   330
         Width           =   585
      End
      Begin VB.Label lblNote2B 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   2055
         TabIndex        =   30
         Top             =   1965
         Width           =   195
      End
      Begin VB.Label lblNote2A 
         AutoSize        =   -1  'True
         Caption         =   "�����~��"
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   1980
         Width           =   780
      End
      Begin VB.Label lblProcessDate 
         AutoSize        =   -1  'True
         Caption         =   "�B�z���"
         Height          =   195
         Left            =   2490
         TabIndex        =   28
         Top             =   1170
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmSO3280A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmer:  Jacky Chen
' Last Modify: 2003/6/3
Option Explicit
Dim intPara24 As Integer
Dim rsBillHeadFmt As New ADODB.Recordset
Dim rsTBillHeadFmt As New ADODB.Recordset
Dim strPBillType As String   '�O�_�ҥθ�A��
Dim strPCitemStr As String  '�D�g���ʶ��ػPP�A��
Dim blnIFlag As Boolean     '�Ұʸ�A�ȫ�A���O���ظ̬O�_���]�tI�A��
Dim strCustIdStr As String
Private blnGetPara As Boolean
Private blnIsOK As Boolean
Public rsM As New ADODB.Recordset
Public rsD As New ADODB.Recordset
Private blnAutoExec As Boolean
Private strReturnMsg As String

Public Property Let uAutoExec(vData As Boolean)
    blnAutoExec = vData
End Property

Public Property Let uGetPara(vData As Boolean)
    blnGetPara = vData
End Property

Public Property Get uIsOk() As Boolean
    uIsOk = blnIsOK
End Property

Public Property Get uReturnMsg() As String
    uReturnMsg = strReturnMsg
End Property

Private Function GetPara() As Boolean
    If Not BatchCharge.GetBatchChargePara(Me, rsM, rsD) Then Exit Function
    blnIsOK = True
    GetPara = True
    Unload Me
End Function

Private Function SetPara() As Boolean
    If Not BatchCharge.SetBatchChargePara(Me, rsM, rsD) Then Exit Function
    SetPara = True
End Function

'Private blnShouldDateMerge As Boolean
Private Sub cboBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
    Dim rsCD019Mutli As New ADODB.Recordset
    Dim strCD019SQL As String
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
        cboTBillHeadFmt.Text = cboBillHeadFmt.Text
        rsTBillHeadFmt.AbsolutePosition = cboTBillHeadFmt.ListIndex + 1
        
        '#7049 ���CD068A.CitemCode By Kin 2015/07/14
        strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                                    " And CD068.BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1"
'        If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        If Not GetRS(rsCD019, strCD019SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        '***************************************************************************************
'        gimBCitemCode.Filter = "Where PeriodFlag = 1"
'        gimTCitemCode.Filter = ""
'        If strCitemCode <> "" Then
'            gimBCitemCode.Filter = "Where CodeNo In (" & strCitemCode & ") And PeriodFlag = 1"
'            gimTCitemCode.Filter = "Where CodeNo In (" & strCitemCode & ")"
'        End If
        '****************************************************************************************
        '#6801 T�榬�O���إ�Filter�|���@�ǨS�Q���A�令��SQL�y�k���s���w By Kin 2014/05/28
        gimBCitemCode.Filter = ""
        gimTCitemCode.Filter = ""
        Dim bCitemSQL As String
        Dim tCitemSQL As String
        bCitemSQL = "Select CodeNo ,Description From " & GetOwner & "CD019 Where 1 = 1  " & _
                        " And StopFlag <> 1 And PeriodFlag = 1  Order  By CodeNo "
        tCitemSQL = "Select CodeNo ,Description From " & GetOwner & "CD019 Where 1= 1" & _
                     " And StopFlag <> 1 Order By CodeNo  "
        'strCitemCode = "Select CitemCode From " & GetOwner & "CD068A Where  CD068A.BillHeadFmt = '" & cboBillHeadFmt.Text & "'"
        '#7049 ���ܵe�������O���� By Kin 2015/07/15
        Dim strCD068ACitem As String
        strCD068ACitem = "Select CitemCode From " & GetOwner & "CD068A Where  CD068A.BillHeadFmt = '" & cboBillHeadFmt.Text & "'"
        If Len(strCitemCode) > 0 Then
            

            bCitemSQL = "Select CodeNo ,Description From " & GetOwner & "CD019 Where CodeNo In (" & strCD068ACitem & ") " & _
                        " And Nvl(StopFlag,0)  =  0 And PeriodFlag = 1  Order  By CodeNo "
            
            tCitemSQL = "Select CodeNo ,Description From " & GetOwner & "CD019 Where CodeNo In (" & strCD068ACitem & ") " & _
                       " And Nvl(StopFlag,0) = 0 Order By CodeNo  "
                        
            SetMsQry gimBCitemCode, bCitemSQL, "���O���إN�X,���O���ئW��"
            SetMsQry gimTCitemCode, tCitemSQL, "���O���إN�X,���O���ئW��"
        Else
            SetMsQry gimBCitemCode, bCitemSQL, "���O���إN�X,���O���ئW��"
            SetMsQry gimTCitemCode, tCitemSQL, "���O���إN�X,���O���ئW��"
        End If
        '#2955 �P�_�ҿ諸�h�C�^�b��O�_���ҥθ�A�ȡA�p�G���n�L�oB��u��P�A�ȻP�D�g���ʪ����� By Kin 2007/04/19 For Jackie
        If rsBillHeadFmt("MutiServiceUnion") = 1 Then '�Ұ�
            '#2955 �p�G���Ұ�P�A���ˬd���O���ظ̭��O�_��I�A�� By Kin 2007/04/24
             'blnIFlag = GetRsValue("Select Count(*) From " & GetOwner & "CD019 Where CodeNo In (" & strCitemCode & ") And ServiceType='I'", gcnGi)
             '#7049 ���CD068A.CitemCode By Kin 2015/07/14
             blnIFlag = GetRsValue("Select Count(*) From " & GetOwner & "CD019 Where CodeNo In (" & strCD068ACitem & ") And ServiceType='I'", gcnGi)
             
            '#2955 �u�dP�A�ȻP�D�g���ʪ����O����
            '#7049 ���CD068A.CitemCode By Kin 2015/07/14
'            If Not GetRS(rsCD019Mutli, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & strCitemCode & ") And PeriodFlag=0 And ServiceType='P'", , adUseClient, adOpenKeyset) Then Exit Sub
            If Not GetRS(rsCD019Mutli, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & strCD068ACitem & ") And PeriodFlag=0 And ServiceType='P'", , adUseClient, adOpenKeyset) Then Exit Sub
            If rsCD019Mutli.RecordCount > 0 Then
                strPBillType = "'B'"
                strPCitemStr = rsCD019Mutli.GetString(, , , ",")
                strPCitemStr = " IN (" & Left(strPCitemStr, Len(strPCitemStr) - 1) & ")"
            Else
                '���M���ҰʡA���䤣�����P�A�Ȫ��D�g���ʶ���
                strPBillType = Empty
                strPCitemStr = Empty
                blnIFlag = False
            End If
        Else
            strPBillType = Empty
            strPCitemStr = Empty
            blnIFlag = False
        End If
        gimBCitemCode.SetQueryCode strCitemCode
        gimTCitemCode.SetQueryCode strCitemCode
        CloseRecordset rsCD019
        CloseRecordset rsCD019Mutli
End Sub



Private Sub cboTBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
    Dim strCD019SQL As String
        If rsTBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsTBillHeadFmt.AbsolutePosition = cboTBillHeadFmt.ListIndex + 1
        strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                                    " And CD068.BillHeadFmt = '" & rsTBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1"
        If Not GetRS(rsCD019, strCD019SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        gimTCitemCode.Filter = ""
        Dim tCitemSQL As String
        tCitemSQL = "Select CodeNo ,Description From " & GetOwner & "CD019 Where 1 = 1" & _
                     " And StopFlag <> 1 Order By CodeNo  "
 
        Dim strCD068ACitem As String
        strCD068ACitem = "Select CitemCode From " & GetOwner & "CD068A Where  CD068A.BillHeadFmt = '" & cboTBillHeadFmt.Text & "'"
        If Len(strCitemCode) > 0 Then
            tCitemSQL = "Select CodeNo ,Description From " & GetOwner & "CD019 Where CodeNo In (" & strCD068ACitem & ") " & _
                       " And Nvl(StopFlag,0) = 0 Order By CodeNo  "
                                    
            SetMsQry gimTCitemCode, tCitemSQL, "���O���إN�X,���O���ئW��"
        Else
            SetMsQry gimTCitemCode, tCitemSQL, "���O���إN�X,���O���ئW��"
        End If
             
        gimTCitemCode.SetQueryCode strCitemCode
        CloseRecordset rsCD019
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo chkErr
        Unload Me
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdCancel_Click"
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo chkErr
    PreviousRpt GetPrinterName(5), "SO3280A.rpt", "���O�渹�վ� [SO3280A]"
    PreviousRpt GetPrinterName(5), "SO3280A1.rpt", "���O�渹�վ� [SO3280A]"
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub Form_Activate()
    On Error GoTo chkErr
        SetPara
        Screen.MousePointer = vbDefault
        If blnAutoExec Then
            cmdOk.Value = True
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo chkErr
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF2 Then
            If cmdOk.Enabled Then Call cmdOK_Click
            KeyCode = 0
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
    On Error GoTo ER
        subGim
        subGil
        Call CboAddItem
        DefaultValue
        chkShouldDateMerge.Value = GetRsValue("Select nvl(ShouldDateMerge,0) From " & GetOwner & "SO041 SO041 where SysID=" & gilCompCode.GetCodeNo, gcnGi)
        'chkShouldDateMerge.Value = blnShouldDateMerge * -1
        If chkShouldDateMerge Then
            gdaOweDate1.SetValue ""
            gdaOweDate2.SetValue ""
            gdaOweDate1.Enabled = False
            gdaOweDate2.Enabled = False
        End If
    Exit Sub
ER:
    If ErrHandle(, True, , "Form_Load") Then Resume 0
End Sub

Private Sub subGil()
    On Error GoTo chkErr
        SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGil"
End Sub

Private Sub CboAddItem()
    On Error Resume Next
        'If cboBillHeadFmt.ListCount > 0 Then Exit Sub
        If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr,MutiServiceUnion From " & GetOwner & "CD068") Then Exit Sub
        Set rsTBillHeadFmt = rsBillHeadFmt.Clone
        Do While Not rsBillHeadFmt.EOF
            cboBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
            cboTBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
            rsBillHeadFmt.MoveNext
        Loop
End Sub

Private Sub DefaultValue()
    On Error GoTo chkErr
    Dim intPara23 As Integer
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        
        gimBillType.SetDispStr "�{�ɦ��O��"
        gimBillType.SetQueryCode "T"
        
        '�B�z���
        gdaProcessDate.SetValue RightDate
        ' �����~��d�� = ��餧�~��ܷ�餧�~��
        gymClctYM1.SetValue (Format(RightNow, "YYYYMM"))
        gymClctYM2.SetValue (Format(RightNow, "YYYYMM"))
        intPara24 = GetSystemParaItem("Para24", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043")
        intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043") & "")
        If intPara24 = 1 Then
            gilServiceType.Clear
            gilServiceType.Visible = False
            lblServiceType.Visible = False
            lblBillHeadFmt.Visible = True
            cboBillHeadFmt.Visible = True
            cboTBillHeadFmt.Visible = True
            cboBillHeadFmt.ListIndex = 0
            cboTBillHeadFmt.ListIndex = 0
            'gimBCitemCode.Enabled = False
            'gimTCitemCode.Enabled = False
            '  �δC��J�b�~Lock
            'If intPara23 = 2 Then
            '    gimBCitemCode.Enabled = False
            '    gimTCitemCode.Enabled = False
            'End If
        Else
            gilServiceType.ListIndex = 1
            gilServiceType.Visible = True
            lblServiceType.Visible = True
            lblBillHeadFmt.Visible = False
            cboBillHeadFmt.Visible = False
            cboTBillHeadFmt.Visible = False
        End If
        
    Exit Sub
chkErr:
    ErrSub Me.Name, "DefalutValue"
End Sub

Private Sub subGim()
    On Error GoTo chkErr
        SetgiMulti gimBCitemCode, "CodeNo", "Description", "CD019", "���O���إN�X", "���O���ئW��", , True
        SetgiMulti gimTCitemCode, "CodeNo", "Description", "CD019", "���O���إN�X", "���O���ئW��", , True
        SetgiMulti gimBUCCode, "CodeNo", "Description", "CD013", "�����O��]�N�X", "�����O��]�W��", " Where Nvl(REFNO,0) NOT IN(3,7) AND NVL(PayOK,0)=0 ", True
        SetgiMulti gimTUCCode, "CodeNo", "Description", "CD013", "�����O��]�N�X", "�����O��]�W��", " Where Nvl(REFNO,0) NOT IN(3,7) AND NVL(PayOK,0)=0 ", True
        SetgiMulti gimPayType, "CodeNo", "Description", "CD112", "�N�X", "�W��"
        Call SetgiMultiAddItem(gimBillType, "I,M,P,T", "�˾���,���׳�,�����,�{�ɦ��O��", "��ڥN��", "��ڦW��")
        If GetPaynowFlag Then
'            gimPayType.SetQueryCode "0"
            '#5925 ú�O���O�W�[�w�]�� By Kin 2011/04/12
            With gimPayType
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
            gimPayType.Clear
            gimPayType.Enabled = False
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGim"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        
        Call ReleaseCOM(Me)
End Sub

Private Sub gdaOweDate1_GotFocus()
    On Error Resume Next
    If gdaOweDate1.GetValue = "" Then gdaOweDate1.SetValue (RightDate)
End Sub

Private Sub gdaOweDate2_GotFocus()
    On Error Resume Next
    If gdaOweDate1.GetValue <> "" Then gdaOweDate2.SetValue gdaOweDate1.GetValue
End Sub

Private Sub gdaShouldDate1_GotFocus()
    On Error Resume Next
        If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetValue RightDate
End Sub

Private Sub gdaShouldDate2_GotFocus()
    On Error Resume Next
    '#2955���դ�OK,�n�a�PgdaShouldDate1�@�˪���,���ORightDate By Kin 2007/06/15 For Penny
        If gdaShouldDate1.GetValue <> "" Then gdaShouldDate2.SetValue gdaShouldDate1.GetValue
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
    On Error Resume Next
        If gdaShouldDate1.GetValue = "" Or gdaShouldDate2.GetValue = "" Then Exit Sub
        If gdaShouldDate2.GetValue < gdaShouldDate1.GetValue Then
            MsgDateRangeX
            Cancel = True
        End If
End Sub

Private Sub gdtCreateTime1_GotFocus()
    On Error Resume Next
    If gdtCreateTime1.GetValue = "" Then gdtCreateTime1.SetValue (RightDate)
End Sub

Private Sub gdtCreateTime2_GotFocus()
  On Error Resume Next
    If gdtCreateTime1.GetValue <> "" Then gdtCreateTime2.SetValue GetDayLastTime(gdtCreateTime1.GetValue(True))
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
    Dim rsTmp As New ADODB.Recordset
    Dim strTmpSql As String
        If Not ChgComp(gilCompCode, "SO3310", "SO3311") Then Exit Sub
        Call subGil
        Call subGim
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.SetCodeNo ""
        gilServiceType.Query_Description
        gilServiceType.ListIndex = 1
        '���D�� 2761 ���դ�ok�A�]�����W�榳���ʻݥ[�J��ڦX�ֽվ㶵�� 2006/12/13
        '#4077 �W�[UpdCreateEn�Ѽ� By Kin 2008/09/18
        '#7397 Add 3 datafield on the  screen conditions that are unable to be modified by kin 2017/04/05
        strTmpSql = "Select nvl(UpdShoulddate,0) as UpdShoulddate,nvl(UpdCreatetime,0) as UpdCreatetime," & _
                    "nvl(UpdCMCode,0) as UpdCMCode,nvl(UpdCreateEn,0) as UpdCreateEn,   " & _
                    "nvl(UpdUCcode,0) as UpdUCcode,nvl(UpdBillCloseDate,0) as UpdBillCloseDate,nvl(UpdPrtSNo,0) as UpdPrtSNo " & _
                    " from " & GetOwner & _
                    "SO041 where SysID=" & gilCompCode.GetCodeNo
        If Not GetRS(rsTmp, strTmpSql) Then Exit Sub
        chkShouldDate = Val(rsTmp("UpdShoulddate") & "")
        chkCreateTime = Val(rsTmp("UpdCreatetime") & "")
        chkCmCode = Val(rsTmp("UpdCMCode") & "")
        chkUpdCreateEn = Val(rsTmp("UpdCreateEn") & "")
        chkUpdUCcode = Val(rsTmp("UpdUCcode") & "")
        chkUpdBillCloseDate = Val(rsTmp("UpdBillCloseDate") & "")
        chkUpdPrtSNo = Val(rsTmp("UpdPrtSNo") & "")
        
        CloseRecordset rsTmp
End Sub

Private Sub gilServiceType_Change()
    On Error Resume Next
        Call GiMultiFilter(gimBCitemCode, gilServiceType.GetCodeNo)
        gimBCitemCode.Filter = gimBCitemCode.Filter & IIf(Trim(gimBCitemCode.Filter) = "", " Where ", " And ") & " PeriodFlag = 1"
        Call GiMultiFilter(gimBUCCode, gilServiceType.GetCodeNo, , , " Nvl(REFNO,0) NOT IN(3,7) AND NVL(PayOK,0)=0 And Nvl(StopFlag,0)=0 ")
        Call GiMultiFilter(gimTCitemCode, gilServiceType.GetCodeNo)
        Call GiMultiFilter(gimTUCCode, gilServiceType.GetCodeNo, , , " Nvl(REFNO,0) NOT IN(3,7) AND NVL(PayOK,0)=0 And Nvl(StopFlag,0)=0 ")
        
End Sub





Private Sub gymClctYM1_Validate(Cancel As Boolean)
    On Error GoTo chkErr
        Call CheckYM
    Exit Sub
chkErr:
    ErrSub Me.Name, "gymClctYM1_Validate"
End Sub

Private Sub gymClctYM2_Validate(Cancel As Boolean)
    On Error GoTo chkErr
        Call CheckYM
    Exit Sub
chkErr:
    ErrSub Me.Name, "gymClctYM2_Validate"
End Sub

Private Sub CheckYM()
    On Error GoTo chkErr
        If gymClctYM1.GetValue <> "" And gymClctYM2.GetValue <> "" Then
            If gymClctYM2.GetValue < gymClctYM1.GetValue Then
                gymClctYM2.SetValue (gymClctYM1.GetValue)
            End If
        End If
        If gymClctYM1.GetValue <> "" And gymClctYM2.GetValue = "" Then
            gymClctYM2.SetValue (gymClctYM1.GetValue)
        End If
        If gymClctYM2.GetValue <> "" And gymClctYM1.GetValue = "" Then
            gymClctYM1.SetValue (gymClctYM2.GetValue)
        End If
        '�X�ֳ���������
        If gdaOweDate2.GetValue <> "" And gdaOweDate1.GetValue = "" Then
            gdaOweDate1.SetValue (gdaOweDate2.GetValue)
        End If
        If gdaOweDate1.GetValue <> "" And gdaOweDate2.GetValue = "" Then
            gdaOweDate2.SetValue (gdaOweDate1.GetValue)
        End If
        '�X�ֳ�ڲ��ͤ��
        If gdtCreateTime2.GetValue <> "" And gdtCreateTime1.GetValue = "" Then
            gdtCreateTime1.SetValue (gdtCreateTime2.GetValue)
        End If
        If gdtCreateTime1.GetValue <> "" And gdtCreateTime2.GetValue = "" Then
            gdtCreateTime2.SetValue (gdtCreateTime1.GetValue)
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "CheckYM"
End Sub

Private Function IsDataOk()
    On Error GoTo chkErr
        IsDataOk = False
        
        If Not ChkDTok Then Exit Function
        If Not MustExist(gilCompCode, 2, "���q�O") Then Exit Function
        If intPara24 = 0 Then If Not MustExist(gilServiceType, 2, "�A�����O") Then Exit Function
        If Not MustExist(gdaProcessDate, 1, "�B�z���") Then Exit Function
        If Not MustExist(gymClctYM1, 1, "�����~��_�l��") Then Exit Function
        If Not MustExist(gymClctYM2, 1, "�����~��I���") Then Exit Function
        If gimBCitemCode.GetQueryCode = "" Then MsgMustBe ("B�榬�O����"): gimBCitemCode.SetFocus: Exit Function
        If Not MustExist(gdaShouldDate1, 1, "�����_�l���") Then Exit Function
        If Not MustExist(gdaShouldDate2, 1, "�����I����") Then Exit Function
        If gimTCitemCode.GetQueryCode = "" Then MsgMustBe ("T�榬�O����"): gimTCitemCode.SetFocus: Exit Function
        If gdaOweDate1.GetValue > gdaOweDate2.GetValue Then MsgBox "�����_�l������o�j��פ��!!", vbInformation, "ĵ�i�T��": gdaOweDate2.SetFocus: Exit Function
        If gdtCreateTime1.GetValue > gdtCreateTime2.GetValue Then MsgBox "���Ͱ_�l�ɶ����o�j��פ�ɶ�!!", vbInformation, "ĵ�i�T��": gdtCreateTime2.SetFocus: Exit Function
        IsDataOk = True
    Exit Function
chkErr:
    ErrSub Me.Name, "IsDataOk"
End Function
' 3.  �T�w���F2���U:
Private Sub cmdOK_Click()
    On Error GoTo chkErr
    Dim strMsg As String
    Dim rsTmp As New ADODB.Recordset
    Dim intCount As Long
    Dim intError As Integer
    Dim aryCustId() As String
    Dim blnShwCombItem As Boolean
    Dim strPayKinds As String
    blnShwCombItem = False
    
        Call CheckYM
        '�E  ���n����ˬd, �Y���L���i�~��C ���F�ƿ露��~, �Ҧ����Ҭ����n
        strCustIdStr = Empty
        '#3478 �W�[�Ȥ�s������ By Kin 2008/02/21
        If Not IsDataOk Then Exit Sub
        If Trim(txtCustId) <> "" Then
            Call ParseWord(txtCustId, aryCustId)
            strCustIdStr = Join(aryCustId, ",")
                   
        End If
        '�I�s���Ѽ� 2011/11/30 Jacky 6153 MinChen
        If blnGetPara Then
            Call GetPara
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        ReadyGoPrint
        
        Call subChoose
        strPayKinds = ""
        '#5756 �W�[ú�I���O�Ѽ� By Kin 2010/08/30
        If (gimPayType.GetQryStr <> "") And (InStr(1, gimPayType.GetQryStr, "����") <= 0) Then
            strPayKinds = gimPayType.GetQryStr
        End If
        If SF_CombineBT(garyGi(1), gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, _
                        gdaProcessDate.GetValue, gymClctYM1.GetValue, gymClctYM2.GetValue, "IN (" & gimBCitemCode.GetQueryCode & ")", _
                        gimBUCCode.GetQryStr, gdaShouldDate1.GetValue, gdaShouldDate2.GetValue, "IN (" & gimTCitemCode.GetQueryCode & ")", _
                        gimTUCCode.GetQryStr, gimBillType.GetQueryCode, chkCombine.Value, gdaOweDate1.GetValue, gdaOweDate2.GetValue, _
                        gdtCreateTime1.GetValue, gdtCreateTime2.GetValue, strPBillType, strPCitemStr, blnIFlag, strCustIdStr, strPayKinds, _
                        cboBillHeadFmt.Text, cboTBillHeadFmt.Text, _
                        strMsg) <= 0 Then
             
            MsgBox strMsg, vbCritical, "ĵ�i"
        Else
            If Not blnAutoExec Then
                MsgBox strMsg, vbInformation, gimsgPrompt
                strSQL = "Select * From " & GetOwner & "SO109 Where CombineId = 0 "
                intError = 1
                intCount = GetRsValue("Select count(*) As intCount From " & GetOwner & "SO109 SO109 Where CombineId = 0 ")
                intError = 2
                
                If intCount = 0 Then
                    MsgBox "�d�L�b��X�ָ��!!", vbInformation, "�T��"
                Else
                    blnShwCombItem = True
                    Call PrintRpt(GetPrinterName(5), "SO3280A.rpt", "SO3280", Me.Caption, strSQL, strChooseString, _
                                        , True, , , "Para24 = " & intPara24, GiPaperPortrait)
                End If
    
                If chkCombine.Value = 0 Then
                    Dim S As String
                    
                    strSQL = "Select * From " & GetOwner & "SO109 Where CombineId = 1 "
                    If rsTmp.State = adStateOpen Then rsTmp.Close
                    intError = 3
                    
                    intCount = GetRsValue("Select count(*) As intCount From " & GetOwner & "SO109 SO109 Where CombineId = 1 ")
                    intError = 4
                    
                    If intCount = 0 Then
                     MsgBox "�d�L�b�楼�X�ָ��!!", vbInformation, "�T��"
                    Else
                     Call PrintRpt(GetPrinterName(5), "SO3280A1.rpt", "SO3280", Me.Caption, strSQL, strChooseString, , True, , , , GiPaperPortrait)
                    End If
    
                End If
                '#5439 �p�G���ҰʦX�ֱb��nShow�X�����T�� By Kin 2010/01/06
                If blnShwCombItem Then
                    If Val(GetRsValue("SELECT NVL(STARTCOMBCITEM,0) FROM " & GetOwner & "SO041", gcnGi) & "") = 1 Then
                        MsgBox "�аO�o���s���榬�O���ئX�֧@�~!!", vbInformation, "�T��"
                    End If
                End If
            Else
                blnIsOK = True
                strReturnMsg = strMsg
                Unload Me
            End If
        End If
        If rsTmp.State = adStateOpen Then rsTmp.Close
        Set rsTmp = Nothing
        Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdOk_Click: ErrorNo:" & intError
End Sub
Public Function SF_CombineBT(ByVal p_UserId As String, ByVal p_CompCode As String, _
    ByVal p_ServiceType As String, ByVal p_ProcDate As Long, ByVal p_YM1 As String, _
    ByVal p_YM2 As String, ByVal p_BCitemStr As String, ByVal p_BUCCode As String, _
    ByVal p_ShouldDate1 As String, ByVal p_ShouldDate2 As String, ByVal p_TCitemStr As String, _
    ByVal p_TUCCode As String, ByVal p_TBillType As String, ByVal p_IMPTCombine As String, _
    ByRef p_BShouldDate1 As String, ByRef p_BShouldDate2 As String, _
    ByRef p_BCreateTime1 As String, ByRef p_BCreateTime2 As String, _
    ByRef p_PBillType As String, ByRef p_PCitemStr As String, ByVal p_iServiceFlag As Boolean, _
    ByVal p_CustStr As String, ByVal p_PayKindSql As String, ByVal p_BillHeadFmt As String, _
    ByVal p_BillHeadFmtT As String, _
    ByRef P_RETMSG As String) As Long
    '#5756 �W�[p_PayKindSql�Ѽ� By Kin 2010/08/30
    On Error GoTo chkErr
        Dim ComSF_CombineBT As New ADODB.Command

        With ComSF_CombineBT
                .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
                .Parameters.Append .CreateParameter("p_UserId", adVarChar, adParamInput, 2000, p_UserId)
                .Parameters.Append .CreateParameter("p_CompCode", adVarNumeric, adParamInput, 2000, p_CompCode)
                .Parameters.Append .CreateParameter("p_ServiceType", adVarChar, adParamInput, 2000, p_ServiceType)
                .Parameters.Append .CreateParameter("p_ProcDate", adVarChar, adParamInput, 2000, p_ProcDate)
                .Parameters.Append .CreateParameter("P_YM1", adVarNumeric, adParamInput, 2000, p_YM1)
                .Parameters.Append .CreateParameter("P_YM2", adVarNumeric, adParamInput, 2000, p_YM2)
                '.Parameters.Append .CreateParameter("p_BCitemStr", adVarChar, adParamInput, 20000, p_BCitemStr)
                '#7049 ���դ�OKp_BCitemStr �H�K���ȴN�n
                .Parameters.Append .CreateParameter("p_BCitemStr", adVarChar, adParamInput, 4000, "0")
                .Parameters.Append .CreateParameter("p_BUCCode", adVarChar, adParamInput, 2000, IIf(p_BUCCode = "", Null, p_BUCCode))
                .Parameters.Append .CreateParameter("p_ShouldDate1", adVarChar, adParamInput, 2000, p_ShouldDate1)
                .Parameters.Append .CreateParameter("p_ShouldDate2", adVarChar, adParamInput, 2000, p_ShouldDate2)
'                .Parameters.Append .CreateParameter("p_TCitemStr", adVarChar, adParamInput, 20000, p_TCitemStr)
                .Parameters.Append .CreateParameter("p_TCitemStr", adVarChar, adParamInput, 2000, "0")
                .Parameters.Append .CreateParameter("p_TUCCode", adVarChar, adParamInput, 2000, IIf(p_TUCCode = "", Null, p_TUCCode))
                .Parameters.Append .CreateParameter("p_TBillType", adVarChar, adParamInput, 2000, IIf(p_TBillType = "", Null, p_TBillType))
                .Parameters.Append .CreateParameter("p_IMPTCombine", adVarChar, adParamInput, 2000, p_IMPTCombine)
                .Parameters.Append .CreateParameter("p_BShouldDate1", adVarChar, adParamInput, 2000, IIf(p_BShouldDate1 = "", Null, p_BShouldDate1))
                .Parameters.Append .CreateParameter("p_BShouldDate2", adVarChar, adParamInput, 2000, IIf(p_BShouldDate2 = "", Null, p_BShouldDate2))
                .Parameters.Append .CreateParameter("p_BCreateTime1", adVarChar, adParamInput, 2000, IIf(p_BCreateTime1 = "", Null, p_BCreateTime1))
                .Parameters.Append .CreateParameter("p_BCreateTime2", adVarChar, adParamInput, 2000, IIf(p_BCreateTime2 = "", Null, p_BCreateTime2))
                '#2955 �W�[�ǤJ���3�ӰѼ� 1.��A�ȧO��������� 2.��A�ȧO�����O���� 3.���O���جO�_���]�tI�A�� By Kin 2007/04/24 For Jackie
                .Parameters.Append .CreateParameter("p_PBillType", adVarChar, adParamInput, 2000, IIf(strPBillType & "" = "", Null, strPBillType))
                .Parameters.Append .CreateParameter("p_PCitemStr", adVarChar, adParamInput, 4000, IIf(strPCitemStr & "" = "", Null, strPCitemStr))
                .Parameters.Append .CreateParameter("p_iServiceFlag", adVarNumeric, adParamInput, 2000, blnIFlag * -1)
                '#3478 �W�[�Ȥ�s���Ѽ� By Kin 2008/02/21
                .Parameters.Append .CreateParameter("p_CustStr", adVarChar, adParamInput, 4000, IIf(strCustIdStr = Empty, Null, strCustIdStr))
                .Parameters.Append .CreateParameter("p_PayKindSql", adVarChar, adParamInput, 2000, p_PayKindSql)
                '#7049 �W�[ �h�b�Უ�ͨ̾� By Kin 2015/07/14
                .Parameters.Append .CreateParameter("p_BillHeadFmt", adVarChar, adParamInput, 2000, p_BillHeadFmt)
                .Parameters.Append .CreateParameter("p_BillHeadFmtT", adVarChar, adParamInput, 2000, p_BillHeadFmtT)
                .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, P_RETMSG)
                
                Set .ActiveConnection = gcnGi
                .CommandText = "SF_CombineBT"
                .CommandType = adCmdStoredProc
                .Execute
                P_RETMSG = .Parameters("P_RETMSG").Value & ""
                SF_CombineBT = Val(.Parameters("Return_value").Value & "")
        End With
   
        Exit Function
chkErr:
        SF_CombineBT = -99: P_RETMSG = "�{������:��L���~"
        ErrSub Me.Name, "SF_CombineBT"
End Function

Private Sub subChoose()
On Error Resume Next

    strChooseString = "���q�O  : " & gilCompCode.GetDescription & ";" & _
                      "�A�����O: " & gilServiceType.GetDescription & ";" & _
                      "�����~��: " & gymClctYM1.GetValue(True) & "~" & gymClctYM2.GetValue(True) & ";" & _
                      "B�榬�O����: " & gimBCitemCode.GetDispStr & ";" & _
                      "B�楼���O��]: " & gimBUCCode.GetDispStr & ";" & _
                      "�������  : " & gdaShouldDate1.GetValue(True) & "~" & gdaShouldDate2.GetValue(True) & ";" & _
                      "T�榬�O����  : " & gimTCitemCode.GetDispStr & ";" & _
                      "T�楼���O��]: " & gimTCitemCode.GetDispStr
End Sub

Private Sub txtCustId_KeyPress(keyAscii As Integer)
  On Error Resume Next
    If keyAscii >= 48 And keyAscii <= 57 Or keyAscii = 44 Or keyAscii = 8 Or keyAscii = 45 Then
        If keyAscii = 44 Or keyAscii = 45 Then
            If Asc(Right(" " & txtCustId.Text, 1)) = 44 Or Asc(Right(" " & txtCustId.Text, 1)) = 45 Or txtCustId = "" Then keyAscii = 9
        End If
        If Not ChkMaxLengthOK(txtCustId) Then
           If Not (keyAscii = 44 Or keyAscii = 8 Or keyAscii = 45) Then keyAscii = 9
        End If
    Else
        keyAscii = 9
    End If
End Sub
