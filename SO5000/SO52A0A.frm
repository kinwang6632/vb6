VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.17#0"; "csMulti.ocx"
Begin VB.Form frmSO52A0A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�C��C��������ڪ��B�έp[ SO52A0A ]"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "SO52A0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame Frame4 
      Caption         =   "�����s��"
      Height          =   735
      Left            =   30
      TabIndex        =   30
      Top             =   4770
      Width           =   7515
      Begin VB.TextBox mskCircuitNo 
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   90
         TabIndex        =   15
         Top             =   240
         Width           =   4605
      End
      Begin VB.OptionButton optAll 
         Caption         =   "����"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6660
         TabIndex        =   17
         Top             =   270
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton optCircuitNo 
         Caption         =   "�u�L�ťպ����s��"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   4830
         TabIndex        =   16
         Top             =   240
         Width           =   1845
      End
   End
   Begin CS_Multi.CSmulti gimStrtCode 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   3990
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   661
      ButtonCaption   =   "��  �D  �d  ��"
   End
   Begin VB.CheckBox chkYN 
      Caption         =   "�O�_�����O������"
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
      Left            =   90
      TabIndex        =   2
      Top             =   540
      Width           =   2205
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
      Left            =   30
      TabIndex        =   22
      Top             =   6960
      Width           =   1245
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
      Left            =   6180
      TabIndex        =   24
      Top             =   6960
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
      Left            =   1710
      TabIndex        =   23
      Top             =   6960
      Width           =   1395
   End
   Begin VB.Frame fraPage 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�p�p�̾�"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   30
      TabIndex        =   25
      Top             =   5550
      Width           =   7515
      Begin VB.OptionButton optServCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�A�Ȱ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   210
         TabIndex        =   18
         Top             =   330
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optAreaCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��F��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1620
         TabIndex        =   19
         Top             =   330
         Width           =   915
      End
      Begin VB.OptionButton optClctEn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "���O�H��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   4740
         TabIndex        =   21
         Top             =   330
         Width           =   1245
      End
      Begin VB.OptionButton optStrtCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��D�s��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   3030
         TabIndex        =   20
         Top             =   330
         Width           =   1245
      End
   End
   Begin Gi_Multi.GiMulti gimBillType 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1260
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   661
      ButtonCaption   =   "��  ��  ��  �O"
      DataType        =   2
   End
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   3600
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   661
      ButtonCaption   =   "��     �F     ��"
   End
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3210
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   661
      ButtonCaption   =   "�A     ��     ��"
   End
   Begin Gi_Multi.GiMulti gimCMCode 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2820
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   661
      ButtonCaption   =   "��  �O  ��  ��"
   End
   Begin Gi_Multi.GiMulti gimStatusCode 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   661
      ButtonCaption   =   "��  ��  ��  �A"
   End
   Begin Gi_Multi.GiMulti gimUCCode 
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   4380
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   661
      ButtonCaption   =   "�� ú �O �� �]"
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   4530
      TabIndex        =   4
      Top             =   450
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
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   4530
      TabIndex        =   3
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
   Begin Gi_Date.GiDate gdaShouldDate2 
      Height          =   345
      Left            =   2340
      TabIndex        =   1
      Top             =   60
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   255
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
      Top             =   60
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   255
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
   Begin CS_Multi.CSmulti gimClctEn 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   2430
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   661
      ButtonCaption   =   "��  �O  �H   ��"
   End
   Begin CS_Multi.CSmulti gimClassCode 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1650
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "��  ��  ��  �O"
   End
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   900
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "��  �O  ��  ��"
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2190
      TabIndex        =   29
      Top             =   180
      Width           =   105
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
      Left            =   3690
      TabIndex        =   28
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3660
      TabIndex        =   27
      Top             =   510
      Width           =   780
   End
   Begin VB.Label lblShouldDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�������"
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
      Left            =   90
      TabIndex        =   26
      Top             =   150
      Width           =   780
   End
End
Attribute VB_Name = "frmSO52A0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table=SO033orSO034 A,(SO014 B,SO002 C)
Option Explicit
Dim strFrom As String
Dim strFieldName As String
Dim TableFieldName As String
Dim strCitemName As String

Private Sub cmdExit_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      Call PreviousRpt(GetPrinterName(5), RptName("SO52A0"), Me.Caption)
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
      Screen.MousePointer = vbHourglass
        cmdExit.SetFocus
        Me.Enabled = False
          ReadyGoPrint
          Call subChoose
          subCreateView
          CreateTable
          Set rsTmp = cnn.Execute("SELECT Count(*) as intCount From SO52A0A2")
          If rsTmp("intCount") = 0 Then
              MsgNoRcd
              SendSQL , , True
          Else
              strsql = "Select * From SO52A0A2"
              Call PrintRpt(GetPrinterName(5), RptName("SO52A0"), , Me.Caption, strsql, strChooseString, , True, "Tmp0000.MDB", , , GiPaperLandscape)
          End If
        Me.Enabled = True
      Screen.MousePointer = vbDefault
      CloseRecordset rsTmp
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: strErrFile = "���q�O": GoTo 66
    If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetFocus: strErrFile = "��������_�l��": GoTo 66
    If gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetFocus: strErrFile = "��������I���": GoTo 66
    If gdaShouldDate2.GetValue(True) >= Format(DateAdd("m", 1, gdaShouldDate1.Text), "YYYY/MM/DD") Then
        MsgBox "���O��������W�L�@�Ӥ�", , "���~�T��"
        gdaShouldDate2_GotFocus
        Exit Function
    End If
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
  Dim strGroupString As String
  Dim blnSO002 As Boolean
  Dim blnSO014 As Boolean
    strChoose = "A.cancelFlag=0"
    strChooseString = ""
    strFrom = "A"
    blnSO002 = False
    blnSO014 = False

  '���
    If gdaShouldDate1.GetValue <> "" Then Call subAnd("A.ShouldDate >= To_Date('" & gdaShouldDate1.GetValue & "','YYYYMMDD')")
    If gdaShouldDate2.GetValue <> "" Then Call subAnd("A.ShouldDate < To_Date('" & gdaShouldDate2.GetValue & "','YYYYMMDD')+1")
    
  'GiMulti
    'If gimCompCode.GetQryStr <> "" Then Call subAnd("C.CompCode" & gimCompCode.GetQryStr)
    If gimCitemCode.GetQryStr <> "" Then Call subAnd("A.CitemCode " & gimCitemCode.GetQryStr)
    If gimBillType.GetQryStr <> "" Then Call subAnd("SubStr(A.BillNo,7,1) " & gimBillType.GetQryStr)
    If gimClassCode.GetQryStr <> "" Then Call subAnd("A.ClassCode " & gimClassCode.GetQryStr)
    If gimStatusCode.GetQryStr <> "" Then Call subAnd("C.CustStatusCode " & gimStatusCode.GetQryStr): blnSO002 = True
    If gimClctEn.GetQryStr <> "" Then Call subAnd("A.ClctEn " & gimClctEn.GetQryStr)
    If gimCMCode.GetQryStr <> "" Then Call subAnd("A.CMCode " & gimCMCode.GetQryStr)
    If gimServCode.GetQryStr <> "" Then Call subAnd("A.ServCode " & gimServCode.GetQryStr)
    If gimAreaCode.GetQryStr <> "" Then Call subAnd("A.AreaCode " & gimAreaCode.GetQryStr)
    If gimStrtCode.GetQryStr <> "" Then Call subAnd("A.StrtCode " & gimStrtCode.GetQryStr)
    If gimUCCode.GetQryStr <> "" Then Call subAnd("A.UCCode " & gimUCCode.GetQryStr)
    
  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("A.CompCode=" & gilCompCode.GetCodeNo)
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("instr(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "')>0")
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("A.ServiceType='" & gilServiceType.GetCodeNo & "'")
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("Instr(Nvl(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "'),'" & gilServiceType.GetCodeNo & "') > 0")
    
  '�����s��
    If mskCircuitNo.Text <> "" Then
       Call subAnd("B.CircuitNo='" & mskCircuitNo.Text & "'")
       blnSO014 = True
    Else
      If optCircuitNo.Value Then Call subAnd("B.CircuitNo is null "): blnSO014 = True
    End If

  '��SO002
    If blnSO002 Then
      strFrom = strFrom & ",SO002 C"
      strChoose = "A.Custid=C.Custid And A.ServiceType=C.ServiceType And " & strChoose
    End If
    
  '��SO014
    If blnSO014 Then
      strFrom = strFrom & ",SO014 B"
      strChoose = "A.AddrNo=B.AddrNo And " & strChoose
    End If

  '�p�p�̾�
    Select Case True
           Case optServCode.Value
                strFrom = strFrom & ",CD002 D"
                strChoose = "A.ServCode=D.CodeNo(+) And " & strChoose
                strGroupName = "D.Description"
                strpagetype = "�A�Ȱ�"
           Case optAreaCode.Value
                strFrom = strFrom & ",CD001 D"
                strChoose = "A.AreaCode=D.CodeNo(+) And " & strChoose
                strGroupName = "D.Description"
                strpagetype = "��F��"
           Case optStrtCode.Value
                strFrom = strFrom & ",CD017 D"
                strChoose = "A.StrtCode=D.CodeNo(+) And " & strChoose
                strGroupName = "D.Description"
                strpagetype = "��D�s��"
           Case optClctEn.Value
                strGroupName = "A.ClctName"
                strpagetype = "���O�H��"
    End Select
    
    If chkYN.Value Then
       strCitemName = "CitemName"
    Else
       strCitemName = "''"
    End If
    
    strChooseString = "�������:" & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & _
                      "�O�_�����O���ؤ���:" & IIf(chkYN.Visible = True, " �O", " �_") & ";" & _
                      "���q�O�@:" & subSpace(gilCompCode.GetDescription) & ";" & _
                      "�A�����O: " & subSpace(gilServiceType.GetDescription) & ";" & _
                      "���O����:" & subSpace(gimCitemCode.GetDispStr) & ";" & _
                      "��ڽs��:" & subSpace(gimBillType.GetDispStr) & ";" & _
                      "�Ȥ����O:" & subSpace(gimClassCode.GetDispStr) & ";" & _
                      "�Ȥ᪬�A:" & subSpace(gimStatusCode.GetDispStr) & ";" & _
                      "���O�H��:" & subSpace(gimClctEn.GetDispStr) & ";" & _
                      "���O�覡:" & subSpace(gimCMCode.GetDispStr) & ";" & _
                      "�A�Ȱϡ@:" & subSpace(gimServCode.GetDispStr) & ";" & _
                      "��F�ϡ@:" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                      "��D�d��:" & subSpace(gimStrtCode.GetDispStr) & ";" & _
                      "��ú�O��]:" & subSpace(gimUCCode.GetDispStr) & ";" & _
                      "�����s��: " & subSpace(mskCircuitNo.Text) & ";" & _
                      "�p�p�̾�: " & strpagetype
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO52A0A)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    If Alfa2 Then Call GetGlobal
    Call subGim
    Call subGil
'    gilCompCode.ListIndex = 1
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "���O���إN�X", "���O���ئW��")
    Call SetgiMultiAddItem(gimBillType, "B,I,P,M,T", "���O��,�˾���,�������,���׳�,�{�ɦ��O��")
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "�Ȥ����O�N�X", "�Ȥ����O�W��")
    Call SetgiMulti(gimStatusCode, "CodeNo", "Description", "CD035", "�Ȥ᪬�A�N�X", "�Ȥ᪬�A�W��")
    Call SetgiMulti(gimClctEn, "EmpNo", "EmpName", "CM003", "���O�H���N�X", "���O�H���W��")
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "���O�覡�N�X", "���O�覡�W��")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "�A�ȰϥN�X", "�A�ȰϦW��")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "��F�ϥN�X", "��F�ϦW��")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "��D�N�X", "��D�W��")
    Call SetgiMulti(gimUCCode, "CodeNo", "Description", "CD013", "��ú�O��]�N�X", "��ú�O��]�W��")
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

Private Function subCreateView() As Boolean
  Dim strView As String
    On Error Resume Next
      gcnGi.Execute "Drop View " & strViewName
    On Error GoTo ChkErr
        strViewName = GetTmpViewName
        strView = "Create View " & strViewName & " as (" & _
                      "SELECT " & strGroupName & " GroupName,A.ShouldDate,A.CitemName,A.RealAmt,A.ShouldAmt,A.STCode,A.UCCode,A.RealAmt-A.ShouldAmt as StAmt,RealDate " & _
                          "From SO033 " & strFrom & " Where " & strChoose & _
                  " Union All " & _
                      "SELECT " & strGroupName & " GroupName,A.ShouldDate,A.CitemName,A.RealAmt,A.ShouldAmt,A.STCode,A.UCCode,A.RealAmt-A.ShouldAmt as StAmt,RealDate " & _
                          "From SO034 " & strFrom & " Where " & strChoose & ")"
        gcnGi.Execute strView
        SendSQL strView
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  gcnGi.Execute "Drop View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO52A0A
End Sub

Private Sub gdaShouldDate1_GotFocus()
  On Error Resume Next
  If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetValue (Format(RightDate, "YYYYMM") & "01")
End Sub

Private Sub gdaShouldDate2_GotFocus()
  On Error Resume Next
    If gdaShouldDate1.GetValue <> "" Then
        gdaShouldDate2.SetValue GetLastDate(gdaShouldDate1.GetValue(True))
    Else
        gdaShouldDate2.SetValue ""
    End If
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaShouldDate1, gdaShouldDate2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimClctEn, , gilCompCode.GetCodeNo
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

Private Sub gilServiceType_Change()
  On Error GoTo ChkErr
    GiMultiFilter gimCitemCode, gilServiceType.GetCodeNo
    GiMultiFilter gimClassCode, gilServiceType.GetCodeNo
    GiMultiFilter gimCMCode, gilServiceType.GetCodeNo
    GiMultiFilter gimUCCode, gilServiceType.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Sub mskCircuitNo_Validate(Cancel As Boolean)
  On Error Resume Next
  If mskCircuitNo.Text = "" Then
    optCircuitNo.Enabled = True
    optAll.Enabled = True
  Else
    optCircuitNo.Enabled = False
    optAll.Enabled = False
  End If
End Sub

Private Sub CreateTable()
  On Error Resume Next
    cnn.Execute "Drop Table SO52A0A1"
    cnn.Execute "Drop Table So52A0A2"
  On Error GoTo ChkErr
    FieldName
    cnn.Execute "Create Table SO52A0A1 (GroupName Text(40),CitemName Text(40),TypeName Text(6)" & TableFieldName & ")"
    '����
    Call subInsertMDB("Select GroupName," & strCitemName & " as CitemName,To_Char(ShouldDate,'DD') as ShouldDay,Sum(ShouldAmt) as Amt From " & strViewName & _
                      " Group by GroupName," & strCitemName & ",To_Char(ShouldDate,'DD') ", "1����")
    '�ꦬ
    Call subInsertMDB("Select GroupName," & strCitemName & " as CitemName,To_Char(ShouldDate,'DD') as ShouldDay,Sum(RealAmt) as Amt From " & strViewName & _
                      " Where RealDate is not null Group by GroupName," & strCitemName & ",To_Char(ShouldDate,'DD') ", "2�ꦬ")
    '�u��
    Call subInsertMDB("Select GroupName," & strCitemName & " as CitemName,To_Char(ShouldDate,'DD') as ShouldDay,Sum(StAmt) as Amt From " & strViewName & _
                      " Where UCCode is Null and ShouldAmt > RealAmt Group by GroupName," & strCitemName & ",To_Char(ShouldDate,'DD') ", "3�u��")
    
    '����
    Call subInsertMDB("Select GroupName," & strCitemName & " as CitemName,To_Char(ShouldDate,'DD') as ShouldDay,Sum(StAmt) as Amt From " & strViewName & _
                      " Where UCCode is Null and ShouldAmt < RealAmt Group by GroupName," & strCitemName & ",To_Char(ShouldDate,'DD') ", "4����")
    '����
    Call subInsertMDB("Select GroupName," & strCitemName & " as CitemName,To_Char(ShouldDate,'DD') as ShouldDay,Sum(ShouldAmt) as Amt From " & strViewName & _
                      " Where UCCode not in (Select CodeNo From CD013 Where Refno=1) Group by GroupName," & strCitemName & ",To_Char(ShouldDate,'DD') ", "5�u��")
    '�ݦ�
    Call subInsertMDB("Select GroupName," & strCitemName & " as CitemName,To_Char(ShouldDate,'DD') as ShouldDay,Sum(ShouldAmt) as Amt From " & strViewName & _
                      " Where UCCode in (Select CodeNo From CD013 Where RefNo=1) Group by GroupName," & strCitemName & ",To_Char(ShouldDate,'DD') ", "6�ݦ�")
    
    InsertTypeName
    SumMDB
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "CreateTabel")
End Sub

Private Sub InsertTypeName()
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim intFor As Integer
    Dim arrType As Variant
    arrType = Array("1����", "2�ꦬ", "3�u��", "4����", "5����", "6�ݦ�")
    Set rsTmp = cnn.Execute("Select GroupName,CitemName From SO52A0A1 Group by GroupName,CitemName")
    While Not rsTmp.EOF
        For intFor = 0 To 5
            cnn.Execute "Insert into SO52A0A1(GroupName,CitemName,TypeName) Values(" & _
                        GetNullString(rsTmp("GroupName")) & "," & _
                        GetNullString(rsTmp("CitemName")) & ",'" & arrType(intFor) & "')"
        Next
        rsTmp.MoveNext
    Wend
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "InsertTypeName")
End Sub

Private Sub subInsertMDB(strsql As String, strType As String)
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
    Set rsTmp = gcnGi.Execute(strsql)
    SendSQL strsql
    cnn.BeginTrans
    While Not rsTmp.EOF
       cnn.Execute "Insert Into SO52A0A1 (GroupName,CitemName,TypeName,Amt" & Format(rsTmp("ShouldDay"), "00") & _
                   ") Values(" & _
                   GetNullString(rsTmp("GroupName")) & "," & _
                   GetNullString(rsTmp("CitemName")) & ",'" & _
                   strType & "'," & _
                   GetNullString(rsTmp("Amt")) & ")"
      rsTmp.MoveNext
    Wend
    cnn.CommitTrans
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "subInsertMDB")
End Sub

Private Sub SumMDB()
  On Error GoTo ChkErr
    cnn.BeginTrans
    cnn.Execute "Select GroupName,CitemName,TypeName" & strFieldName & _
                " Into SO52A0A2 From SO52A0A1 Group by " & _
                " GroupName,CitemName,TypeName"
    cnn.CommitTrans
    SendSQL "Select GroupName,CitemName,TypeName" & strFieldName & _
            " Into SO52A0A2 From SO52A0A1 Group by " & _
            " GroupName,CitemName,TypeName", True
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "SumMDB")
End Sub

Private Sub FieldName()
  On Error GoTo ChkErr
    Dim intFor As Integer
    TableFieldName = ""
    strFieldName = ""
    For intFor = 1 To 31
      TableFieldName = TableFieldName & ",Amt" & Format(intFor, "00") & " Long"
      strFieldName = strFieldName & ",iif(isnull(sum(Amt" & Format(intFor, "00") & ")),0,sum(Amt" & Format(intFor, "00") & ")) as Amt" & Format(intFor, "00")
    Next
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "FieldName")
End Sub
