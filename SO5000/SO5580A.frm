VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO5580A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�j�Ӥ�W��Ʋέp���� [SO5580A]"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   Icon            =   "SO5580A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame Frame4 
      Caption         =   "�����s��"
      Height          =   735
      Left            =   0
      TabIndex        =   24
      Top             =   2970
      Width           =   7665
      Begin VB.TextBox mskCircuitNo 
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   90
         TabIndex        =   9
         Top             =   240
         Width           =   4605
      End
      Begin VB.OptionButton optAll 
         Caption         =   "����"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6660
         TabIndex        =   11
         Top             =   270
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton optCircuitNo 
         Caption         =   "�u�L�ťպ����s��"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   4830
         TabIndex        =   10
         Top             =   240
         Width           =   1845
      End
   End
   Begin CS_Multi.CSmulti gimStrtCode 
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   2400
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   714
      ButtonCaption   =   "��  �D  �d  ��"
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "�צ�Excel"
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
      Left            =   3660
      TabIndex        =   14
      Top             =   4680
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:�C�L"
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
      TabIndex        =   12
      Top             =   4680
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
      Left            =   6240
      TabIndex        =   15
      Top             =   4680
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
      TabIndex        =   13
      Top             =   4680
      Width           =   1395
   End
   Begin Gi_Date.GiDate gdaTime2 
      Height          =   345
      Left            =   2310
      TabIndex        =   1
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
   Begin Gi_Date.GiDate gdaTime1 
      Height          =   345
      Left            =   960
      TabIndex        =   0
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   405
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   714
      ButtonCaption   =   "�A     ��     ��"
      FontSize        =   9
      FontName        =   "�s�ө���"
   End
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   405
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   714
      ButtonCaption   =   "��     �F     ��"
      FontSize        =   9
      FontName        =   "�s�ө���"
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   4590
      TabIndex        =   3
      Top             =   510
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
      Left            =   4590
      TabIndex        =   2
      Top             =   90
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
   Begin CS_Multi.CSmulti gimMduId 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   661
      ButtonCaption   =   "�j  ��  �s  ��"
   End
   Begin CS_Multi.CSmulti gimClassCode1 
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   714
      ButtonCaption   =   "��  ��  ��  �O"
   End
   Begin VB.Label lblHelp 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "����:"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   150
      TabIndex        =   23
      Top             =   3930
      Width           =   510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�˾���-->�˾��ɶ�     �����-->����ɶ�"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   1620
      TabIndex        =   22
      Top             =   3960
      Width           =   3165
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "������-->�����ɶ�     ���P��-->���P�ɶ�"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   1620
      TabIndex        =   21
      Top             =   4170
      Width           =   3165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "����d��:"
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
      Height          =   195
      Left            =   720
      TabIndex        =   20
      Top             =   3960
      Width           =   825
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
      Left            =   3750
      TabIndex        =   19
      Top             =   150
      Width           =   765
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3750
      TabIndex        =   18
      Top             =   555
      Width           =   780
   End
   Begin VB.Label lblInstTime 
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
      Left            =   30
      TabIndex        =   17
      Top             =   180
      Width           =   780
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2160
      TabIndex        =   16
      Top             =   180
      Width           =   165
   End
End
Attribute VB_Name = "frmSO5580A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table=SO001,SO002,SO014
Option Explicit
Dim MdbChoose As String
Dim blnExcel As Boolean

Private Sub cmdExcel_Click()
  On Error GoTo ChkErr
    blnExcel = True
    cmdPrint_Click
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExcel_Click")
End Sub

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
    Call PreviousRpt(GetPrinterName(5), RptName("SO5580"), Me.Caption)
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
        Call SQLView
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
    If gilServiceType.GetCodeNo = "" Then gilServiceType.SetFocus: strErrFile = "�A�����O": GoTo 66
    IsDataOk = True
    Exit Function
66:
    MsgMustBe (strErrFile)
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub subPrint()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
    Set rsTmp = gcnGi.Execute("SELECT Count(*) as intCount From " & strViewName)
      If rsTmp("intCount") = 0 Then
         MsgNoRcd
         SendSQL , , True
      Else
         strSQL = "Select * From " & strViewName & " V"
         If blnExcel Then
            Call toExcel(strSQL)
         Else
            Call PrintRpt(GetPrinterName(5), RptName("SO5580"), , Me.Caption, strSQL, strChooseString, , True, , , , GiPaperLandscape)
         End If
      End If
    CloseRecordset rsTmp
    blnExcel = False
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subPrint")
End Sub

Private Sub toExcel(ByVal strSQL As String)
  On Error GoTo ChkErr
    Dim rsExcel As New ADODB.Recordset
      RptToTxt RptName("SO5580"), , strSQL, strChooseString, , , , , , garyGi(19) & "\SO5580"
      If Not Get_RS_From_Txt(garyGi(19), "SO5580.txt", rsExcel) Then blnExcel = False: Exit Sub
       '#7074 To Add a choice-condition parameter into excel print function by Kin 2015/09/09
      Call UseProperty(rsExcel, "�j�Ӥ�W��Ʋέp����", "�Ĥ@��", _
            , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , strChooseString)
      blnExcel = False
      CloseRecordset rsExcel
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "toExcel")
End Sub

Private Sub subChoose()
  On Error GoTo ChkErr
  Dim blnFlag As Boolean
    'strChoose = " From SO001,SO014 Where SO001.InstAddrNo=SO014.AddrNo And SO001.MduId is Not Null "
    strChoose = ""
    strChooseString = ""
    blnFlag = False
    
  'GiMulti
    'If gimCompCode.GetQryStr <> "" Then Call subAnd("SO001.CompCode " & gimCompCode.GetQryStr)
    '���D�� 2584�h��@�ӫȤ����O
    If gimClassCode1.GetQryStr <> "" And gimClassCode1.GetQryStr <> "����" Then
        Call subAnd("SO001.ClassCode1 " & gimClassCode1.GetQryStr)
    End If
    If gimServCode.GetQryStr <> "" Then Call subAnd("SO001.ServCode " & gimServCode.GetQryStr)
    If gimAreaCode.GetQryStr <> "" Then blnFlag = True: Call subAnd("SO014.AreaCode " & gimAreaCode.GetQryStr)
    If gimStrtCode.GetQryStr <> "" Then blnFlag = True: Call subAnd("SO014.StrtCode " & gimStrtCode.GetQryStr)
    If gimMduId.GetQryStr <> "" Then Call subAnd("SO001.MduId " & gimMduId.GetQryStr)
      
  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO001.CompCode=" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("instr(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "')>0")
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO002.ServiceType='" & gilServiceType.GetCodeNo & "'")
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("Instr(Nvl(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "'),'" & gilServiceType.GetCodeNo & "') > 0")
    
  '�����s��
    If mskCircuitNo.Text <> "" Then
       blnFlag = True
       Call subAnd("SO014.CircuitNo='" & mskCircuitNo.Text & "'")
    Else
      blnFlag = True:
      If optCircuitNo.Value Then Call subAnd("SO014.CircuitNo is null ")
    End If
    
    If blnFlag Then
      strChoose = " From SO001,SO002,SO017,SO014 Where SO001.MDUId=SO017.MDUId And SO001.Custid=SO002.Custid And SO001.InstAddrNo=SO014.AddrNo And SO001.MduId is Not Null " & IIf(strChoose = "", "", " And " & strChoose)
    Else
      strChoose = " From SO001,SO002,SO017 Where SO001.MDUId=SO017.MDUId And SO001.Custid=SO002.Custid And SO001.MduId is Not Null " & IIf(strChoose = "", "", " And " & strChoose)
    End If

   strChooseString = "����d��:" & subSpace(gdaTime1.GetValue(True)) & "~" & subSpace(gdaTime2.GetValue(True)) & ";" & _
                     "���q�O�@:" & subSpace(gilCompCode.GetDescription) & ";" & _
                     "�A�����O: " & subSpace(gilServiceType.GetDescription) & ";" & _
                     "�j�ӦW��:" & subSpace(gimMduId.GetDispStr) & ";" & _
                     "�A�Ȱϡ@:" & subSpace(gimServCode.GetDispStr) & ";" & _
                     "��F�ϡ@:" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                     "��D�d��:" & subSpace(gimStrtCode.GetDispStr) & ";" & _
                     "�����s��: " & subSpace(mskCircuitNo.Text)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub MdbAnd(strCH As String)
  On Error GoTo ChkErr
    If MdbChoose <> "" Then
       MdbChoose = MdbChoose & " And " & strCH
    Else
       MdbChoose = strCH
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "MdbAnd")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5580A)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGim
    Call subGil
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
    '���D��2584 �Ȥ����O�w�]�a������
    gimClassCode1.SelectAll
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    'Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimMduId, "MduId", "Name", "SO017", "�j�ӥN��", "�j�ӦW��")
    '���D��2584 �s�W�@�ӫȤ����O
    Call SetgiMulti(gimClassCode1, "CodeNo", "Description", "CD004", "�Ȥ����O�N�X", "�Ȥ����O�W��")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "�A�ȰϥN�X", "�A�ȰϦW��")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "��F�ϥN�X", "��F�ϦW��")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "��D�N�X", "��D�W��")
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  gcnGi.Execute "Drop View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5580A
End Sub

Private Sub gdaTime1_GotFocus()
  On Error Resume Next
  If gdaTime1.GetValue = "" Then gdaTime1.SetValue (RightDate)
End Sub

Private Sub gdaTime2_GotFocus()
  On Error Resume Next
  If gdaTime1.GetValue = "" Or gdaTime2.GetValue = "" Then gdaTime2.SetValue (gdaTime1.GetValue)
End Sub

Private Sub gdaTime2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaTime1, gdaTime2)
End Sub

Private Sub SQLView()
  On Error GoTo ChkErr
  Dim strSQL As String
    MdbChoose = ""
      '���`
        MdbChoose = " And SO002.CustStatusCode=1 "
        strSQL = " SELECT SO017.MDUId,SO017.Name,Count(Distinct SO001.Custid) as NormalCount,0 as InstCount,0 as StopCount," & _
                 "0 as PRCount,0 as DelCount,0 as PromCount,0 as PRingCount " & strChoose & MdbChoose & " Group by SO017.MDUId,SO017.Name"
        
      '�˾�
        MdbChoose = " And SO002.CustStatusCode=1 "
        If gdaTime1.GetValue <> "" Then Call MdbAnd("SO002.InstTime >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
        If gdaTime2.GetValue <> "" Then Call MdbAnd("SO002.InstTime < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
        strSQL = strSQL & " Union All SELECT SO017.MDUId,SO017.Name,0 as NormalCount,Count(Distinct SO001.Custid) as InstCount,0 as StopCount," & _
                 "0 as PRCount,0 as DelCount,0 as PromCount,0 as PRingCount " & strChoose & MdbChoose & " Group by SO017.MDUId,SO017.Name"
        
      '����
        MdbChoose = " And SO002.CustStatusCode=2 "
        If gdaTime1.GetValue <> "" Then Call MdbAnd("SO002.StopTime >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
        If gdaTime2.GetValue <> "" Then Call MdbAnd("SO002.StopTime < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
        strSQL = strSQL & " Union All SELECT SO017.MDUId,SO017.Name,0 as NormalCount,0 InstCount,Count(Distinct SO001.Custid) as StopCount," & _
                 "0 as PRCount,0 as DelCount,0 as PromCount,0 as PRingCount " & strChoose & MdbChoose & " Group by SO017.MDUId,SO017.Name"
        
      '���
        MdbChoose = " And SO002.CustStatusCode=3 "
        If gdaTime1.GetValue <> "" Then Call MdbAnd("SO002.PRTime >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
        If gdaTime2.GetValue <> "" Then Call MdbAnd("SO002.PRTime < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
        strSQL = strSQL & " Union All SELECT SO017.MDUId,SO017.Name,0 as NormalCount,0 as InstCount,0 as StopCount," & _
                    "Count(Distinct SO001.Custid) as PRCount,0 as DelCount,0 as PromCount,0 as PRingCount " & strChoose & MdbChoose & " Group by SO017.MDUId,SO017.Name"
                    
      '���P
        MdbChoose = " And SO002.CustStatusCode=4 "
        If gdaTime1.GetValue <> "" Then Call MdbAnd("SO002.DelDate >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
        If gdaTime2.GetValue <> "" Then Call MdbAnd("SO002.DelDate < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
        strSQL = strSQL & " Union All SELECT SO017.MDUId,SO017.Name,0 as NormalCount,0 as InstCount,0 as StopCount," & _
                    "0 as PRCount,Count(Distinct SO001.Custid) as DelCount,0 as PromCount,0 as PRingCount " & strChoose & MdbChoose & " Group by SO017.MDUId,SO017.Name"
                    
      '�P�P,�����
       ' MdbChoose = " And SO002.CustStatusCode=5 "
        strSQL = strSQL & " Union All SELECT SO017.MDUId,SO017.Name,0 as NormalCount,0 as InstCount,0 as StopCount,0 as PRCount,0 as DelCount," & _
                    "Count(Distinct DeCode(CustStatusCode,5,SO001.Custid,Null)) as PromCount,Count(Distinct DeCode(CustStatusCode,6,SO001.Custid,Null)) as PRingCount " & _
                    strChoose & " Group by SO017.MDUId,SO017.Name"
        Call subCreateView(strSQL)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "SQLView")
End Sub

Private Sub subCreateView(strSQL As String)
  On Error Resume Next
    gcnGi.Execute "Drop View " & strViewName
  On Error GoTo ChkErr
    strViewName = GetTmpViewName
  
    strSQL = "Create View " & strViewName & " as (" & _
                  "Select A.MDUId,A.Name MDUName,Sum(NormalCount) as NormalCount,Sum(A.InstCount) as InstCount,Sum(A.StopCount) as StopCount," & _
                  "Sum(A.PRCount) as PRCount,Sum(A.DelCount) as DelCount,Sum(A.PromCount) as PromCount,Sum(A.PRingCount) as PRingCount From (" & strSQL & _
                  ") A Group by A.MDUId,A.Name) "
    gcnGi.Execute strSQL
    SendSQL strSQL, True
  Exit Sub
ChkErr:
    ErrSub Me.Name, "subCreateView"
End Sub

'Private Sub SQLMDB()
'  On Error GoTo ChkErr
'  Dim I As Integer
'    cnn.Execute "DELETE * FROM SO5580A"
'    MdbChoose = ""
'    '���
'    blnflag = False
'    If gdaTime1.GetValue <> "" Or gdaTime2.GetValue <> "" Then
'      For I = 1 To 4
'        Select Case I
'          Case 1 '�˾�
'            'MdbChoose = " SO001.CustStatusCode in (1,6) "
'            MdbChoose = ""
'            If gdaTime1.GetValue <> "" Then Call MdbAnd("SO001.InstTime >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
'            If gdaTime2.GetValue <> "" Then Call MdbAnd("SO001.InstTime < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
'
'          Case 2 '����
'            'MdbChoose = " SO001.CustStatusCode=2 "
'            MdbChoose = ""
'            If gdaTime1.GetValue <> "" Then Call MdbAnd("SO001.StopTime >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
'            If gdaTime2.GetValue <> "" Then Call MdbAnd("SO001.StopTime < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
'
'          Case 3 '���
'            'MdbChoose = " SO001.CustStatusCode=3 "
'            MdbChoose = ""
'            If gdaTime1.GetValue <> "" Then Call MdbAnd("SO001.PRTime >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
'            If gdaTime2.GetValue <> "" Then Call MdbAnd("SO001.PRTime < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
'
'          Case 4 '���P
'            'MdbChoose = " SO001.CustStatusCode=4 "
'            MdbChoose = ""
'            If gdaTime1.GetValue <> "" Then Call MdbAnd("SO001.DelDate >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
'            If gdaTime2.GetValue <> "" Then Call MdbAnd("SO001.DelDate < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
'
'        End Select
'
'
'
'        If strChoose <> "" Then
'          MdbChoose = " And " & MdbChoose & " And " & strChoose
'        Else
'          MdbChoose = " And " & MdbChoose
'        End If
'        If I = 4 Then blnflag = True
'        subInsertMDB
'
'      Next
'    Else
'      If strChoose <> "" Then
'        MdbChoose = " And " & strChoose
'      End If
'      blnflag = True
'      subInsertMDB
'    End If
'
'  Exit Sub
'ChkErr:
'    Call ErrSub(Me.Name, "SQLMDB")
'End Sub

'Private Sub subInsertMDB()
'    On Error GoTo ChkErr
'    Dim rsTmp As New adodb.Recordset
'    Dim strMDB As String
'
'      If cnn.State = adStateClosed Then Set cnn = GetTmpMdbCn
'      If rsTmp.State = 1 Then rsTmp.Close
'        strMDB = "SELECT SO001.Custid,SO001.CustStatusCode,SO014.MDUId,SO014.MDUName" & _
'                    " From SO001,SO014 Where SO001.InstAddrNo=SO014.AddrNo " & _
'                    MdbChoose & ""
'        Set rsTmp = gcnGi.Execute(strMDB)
'        If blnflag Then
'          SendSQL strMDB, True
'        Else
'          SendSQL strMDB
'        End If
'        cnn.BeginTrans
'        '(Custid,CustStatusCode,AreaCode,AreaName,ServCode,ServName)
'          While Not rsTmp.EOF
'            cnn.Execute "INSERT INTO SO5580A " & _
'                    "VALUES (" & _
'                    GetNullString(rsTmp("Custid")) & "," & _
'                    GetNullString(rsTmp("CustStatusCode")) & "," & _
'                    GetNullString(rsTmp("MDUId")) & "," & _
'                    GetNullString(rsTmp("MDUName")) & ")"
'            rsTmp.MoveNext
'            DoEvents
'
'          Wend
'        cnn.CommitTrans
'        Set rsTmp = Nothing
'
'    Exit Sub
'ChkErr:
'  cnn.RollbackTrans
'    Call ErrSub(Me.Name, "subInsertMDB")
'End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimMduId, , gilCompCode.GetCodeNo
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

Private Sub mskCircuitNo_Validate(Cancel As Boolean)
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
    Call ErrSub(Me.Name, "mskCircuitNo_Validate")
End Sub
