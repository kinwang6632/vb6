VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.17#0"; "csMulti.ocx"
Begin VB.Form frmSO5540A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "���P����Ȥ��Ƥ@���� [SO5540A]"
   ClientHeight    =   5970
   ClientLeft      =   1950
   ClientTop       =   1530
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO5540A.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin CS_Multi.CSmulti gimStrtCode 
      Height          =   405
      Left            =   0
      TabIndex        =   12
      Top             =   2970
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   714
      ButtonCaption   =   "��  �D  �d  ��"
   End
   Begin MSMask.MaskEdBox mskCircuitNo 
      Height          =   345
      Left            =   5010
      TabIndex        =   7
      Top             =   540
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   609
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "�W���έp���G"
      Height          =   525
      Left            =   1770
      TabIndex        =   20
      Top             =   5370
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "����(&X)"
      Height          =   525
      Left            =   6330
      TabIndex        =   21
      Top             =   5370
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:�C�L"
      Default         =   -1  'True
      Height          =   525
      Left            =   135
      TabIndex        =   19
      Top             =   5370
      Width           =   1245
   End
   Begin VB.Frame fraPage 
      Caption         =   "�����覡"
      ForeColor       =   &H00FF0000&
      Height          =   675
      Left            =   90
      TabIndex        =   26
      Top             =   4200
      Width           =   7545
      Begin VB.OptionButton optClctAreaCode 
         Caption         =   "���O��"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3315
         TabIndex        =   17
         Top             =   330
         Width           =   915
      End
      Begin VB.OptionButton optServCode 
         Caption         =   "�A�Ȱ�"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   630
         TabIndex        =   15
         Top             =   300
         Width           =   915
      End
      Begin VB.OptionButton optAreaCode 
         Caption         =   "��F��"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1980
         TabIndex        =   16
         Top             =   300
         Width           =   915
      End
      Begin VB.OptionButton optNothing 
         Caption         =   "�L"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4665
         TabIndex        =   18
         Top             =   315
         Value           =   -1  'True
         Width           =   585
      End
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   945
      TabIndex        =   3
      Top             =   570
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
   Begin VB.ComboBox cboDateType 
      ForeColor       =   &H00C000C0&
      Height          =   315
      ItemData        =   "SO5540A.frx":0442
      Left            =   60
      List            =   "SO5540A.frx":044F
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   0
      Top             =   120
      Width           =   1485
   End
   Begin Gi_Date.GiDate gdaDate2 
      Height          =   345
      Left            =   2940
      TabIndex        =   2
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   12583104
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
   Begin Gi_Date.GiDate gdaDate1 
      Height          =   345
      Left            =   1620
      TabIndex        =   1
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   12583104
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
   Begin Gi_Date.GiDate gdaShouldDate2 
      Height          =   345
      Left            =   6360
      TabIndex        =   6
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
   Begin Gi_Date.GiDate gdaShouldDate1 
      Height          =   345
      Left            =   5010
      TabIndex        =   5
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
   Begin Gi_Multi.GiMulti gimStatusCode 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   1410
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "��  ��  ��  �A"
   End
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   2580
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "��     �F     ��"
   End
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2190
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "�A     ��     ��"
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   945
      TabIndex        =   4
      Top             =   975
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
   Begin Gi_Multi.GiMulti gimOrder 
      Height          =   405
      Left            =   0
      TabIndex        =   14
      Top             =   3750
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   714
      ButtonCaption   =   "��  ��  ��  ��"
      DataType        =   2
      ColumnOrder     =   -1  'True
   End
   Begin CS_Multi.CSmulti gimMduId 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   3360
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   661
      ButtonCaption   =   "�j  ��  �s  ��"
   End
   Begin CS_Multi.CSmulti gimClassCode 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   661
      ButtonCaption   =   "��  ��  ��  �O"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
      Caption         =   "�A�����O"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   90
      TabIndex        =   28
      Top             =   1035
      Width           =   780
   End
   Begin VB.Label lblCircuitNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�����s��"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4200
      TabIndex        =   27
      Top             =   630
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�� �q �O"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   90
      TabIndex        =   25
      Top             =   630
      Width           =   675
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   6210
      TabIndex        =   24
      Top             =   180
      Width           =   105
   End
   Begin VB.Label lblShouldDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�������"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4200
      TabIndex        =   23
      Top             =   180
      Width           =   780
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   2790
      TabIndex        =   22
      Top             =   180
      Width           =   105
   End
End
Attribute VB_Name = "frmSO5540A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table:SO001,SO002(,SO014),SO017
Option Explicit
Dim strOrder As String
Dim strUseTable As String
Dim strAddrSort As String

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
    Call PreviousRpt(GetPrinterName(5), RptName("SO5540"), "���P����Ȥ��Ƥ@���� [SO5540A]")
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
          Call subCreateView
          Set rsTmp = gcnGi.Execute("SELECT Count(*) as intCount From " & strViewName)
          If rsTmp("intCount") = 0 Then
             MsgNoRcd
             SendSQL , , True
          Else
             'strSql = "SELECT SO001.CUSTID,SO001.CUSTNAME,SO001.TEL1,SO001.INSTADDRESS,SO002.CUSTSTATUSCODE," & _
                     "SO002.CUSTSTATUSNAME,SO001.CLASSNAME1,SO002.INSTTIME,SO002.STOPTIME,SO002.STOPNAME,SO002.PRTIME," & _
                     "SO002.PRNAME,SO002.DELDATE,SO002.DELNAME,SO017.Name,SO014.AddrSort," & _
                     "SO014.AreaCode,SO014.AreaName,SO001.ServCode,SO001.ServArea " & strChoose
             strsql = "Select * From " & strViewName & " V"
             Call PrintRpt(GetPrinterName(5), RptName("SO5540"), , Me.Caption, strsql, strChooseString, , True, , , strGroupName, GiPaperLandscape)
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
    If gilServiceType.GetCodeNo = "" Then gilServiceType.SetFocus: strErrFile = "�A�����O": GoTo 66
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
    Dim strPage As String
    Dim arrOrder() As String
    Dim varOrder As Variant
    Dim intSort As Integer
    Dim blnSO014 As Boolean
    
    'strChoose = " From SO001,SO002,SO014,SO017 Where SO001.CustId=SO002.CustId And SO001.InstAddrNo=SO014.AddrNo And SO001.MduId=SO017.MduId(+) "
    strUseTable = "SO001,SO002,SO017"
    strChoose = "SO001.CustId=SO002.CustId And SO001.MduId=SO017.MduId(+)"
    strChooseString = ""
    strGroupName = ""
    strAddrSort = "''"
    blnSO014 = False

  '���
    If gdaDate1.GetValue <> "" Then Call subAnd(subDateType(">= To_Date('" & gdaDate1.GetValue & "','YYYYMMDD')"))
    If gdaDate2.GetValue <> "" Then Call subAnd(subDateType("< To_Date('" & gdaDate2.GetValue & "','YYYYMMDD')+1"))
    If gdaShouldDate1.GetValue <> "" Then Call subAnd("SO001.NextBCDate >= To_Date('" & gdaShouldDate1.GetValue & "','YYYYMMDD')")
    If gdaShouldDate2.GetValue <> "" Then Call subAnd("SO001.NextBCDate < To_Date('" & gdaShouldDate2.GetValue & "','YYYYMMDD')+1")
    
  'GiList
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("Instr(Nvl(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "'),'" & gilServiceType.GetCodeNo & "') > 0")
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO002.ServiceType='" & gilServiceType.GetCodeNo & "'")
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("instr(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "')>0")
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO001.CompCode=" & gilCompCode.GetCodeNo)
    
  '�����s��
    If mskCircuitNo.Text <> "" Then Call subAnd("SO014.CircuitNo=" & mskCircuitNo.Text): blnSO014 = True
    
  'GiMulti
    If gimAreaCode.GetQryStr <> "" Then Call subAnd("SO014.AreaCode " & gimAreaCode.GetQryStr): blnSO014 = True
    If gimClassCode.GetQryStr <> "" Then Call subAnd("SO001.ClassCode1 " & gimClassCode.GetQryStr)
    If gimStatusCode.GetQryStr <> "" Then
        Call subAnd("SO002.CustStatusCode " & gimStatusCode.GetQryStr)
    Else
        Call subAnd("SO002.CustStatusCode in (2,3,4,6)")
    End If
    If gimServCode.GetQryStr <> "" Then Call subAnd("SO001.ServCode " & gimServCode.GetQryStr)
    If gimStrtCode.GetQryStr <> "" Then Call subAnd("SO014.StrtCode " & gimStrtCode.GetQryStr): blnSO014 = True
    If gimMduId.GetQryStr <> "" Then
       Call subAnd("SO001.MduId " & gimMduId.GetQryStr)
    Else
       If gimMduId.GetDispStr <> "" Then subAnd ("SO001.MduId is Not Null")
    End If
            
  '�����覡
    Select Case True
           Case optAreaCode.Value '��F��
                strGroupName = "GroupName={V.AreaCode};GroupName2={V.AreaName}"
                strPage = "��F��"
                blnSO014 = True
           Case optServCode.Value '�A�Ȱ�
                strGroupName = "GroupName={V.ServCode};GroupName2={V.ServArea}"
               strPage = "�A�Ȱ�"
               
           Case optClctAreaCode.Value
                strGroupName = "GroupName={V.ClctAreaCode};GroupName2={V.ClctAreaName}"
               strPage = "���O��"
                
           Case optNothing.Value '�L
                strGroupName = "GroupName={V.AreaCode};GroupName2={V.AreaName}"
                strPage = "�L"
    End Select
    
  '�ƧǤ覡
    If gimOrder.GetColumnOrderCode <> "" Then
      arrOrder = Split(gimOrder.GetColumnOrderCode, ",")
      intSort = 0
      For Each varOrder In arrOrder
        Select Case arrOrder(intSort)
          Case "A"
            strGroupName = strGroupName & ";Sort" & intSort & "={V.Custid}"
          Case "B"
            strGroupName = strGroupName & ";Sort" & intSort & "={V.AddrSort}"
            blnSO014 = True
            strAddrSort = "SO014.AddrSort"
          Case "C"
            strGroupName = strGroupName & ";Sort" & intSort & "={V.CustStatusCode}"
          Case "D"
            strGroupName = strGroupName & ";Sort" & intSort & "={V.ClassCode1}"
          Case "E"
            strGroupName = strGroupName & ";Sort" & intSort & "=Date({V.InstTime})"
          Case "F"
            strGroupName = strGroupName & ";Sort" & intSort & "={V.ServCode}"
        End Select
        intSort = intSort + 1
      Next
    End If
    
    If blnSO014 Then
        strUseTable = strUseTable & ",SO014"
        strChoose = "SO001.InstAddrNo=SO014.AddrNo And " & strChoose
    End If

    strChooseString = cboDateType.Text & ":" & subSpace(gdaDate1.GetValue(True)) & "~" & subSpace(gdaDate2.GetValue(True)) & ";" & _
                      "�������:" & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & _
                      "���q�O�@:" & subSpace(gilCompCode.GetDescription) & ";" & _
                      "�A�����O:" & subSpace(gilServiceType.GetDescription) & ";" & _
                      "�����s��:" & subSpace(mskCircuitNo.Text) & ";" & _
                      "�Ȥ᪬�A:" & subSpace(gimStatusCode.GetDispStr) & ";" & _
                      "�Ȥ����O:" & subSpace(gimClassCode.GetDispStr) & ";" & _
                      "�A�Ȱϡ@:" & subSpace(gimServCode.GetDispStr) & ";" & _
                      "��F�ϡ@:" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                      "��D�d��:" & subSpace(gimStrtCode.GetDispStr) & ";" & _
                      "�j�ӦW��:" & subSpace(gimMduId.GetDispStr) & ";" & _
                      "�ƧǤ覡:" & subSpace(gimOrder.GetColumnOrderDspStr) & ";" & _
                      "�����覡:" & subSpace(strPage)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Private Function subAndStr(strBody As String, strAnd As String) As String
  On Error GoTo ChkErr
    If strBody = "" Then
       subAndStr = strAnd
    Else
       subAndStr = strBody & " And " & strAnd
    End If
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Function

Private Function subDateType(strGetValue As String) As String
  On Error GoTo ChkErr
    Select Case cboDateType.ListIndex
           Case 1 '�������
                subDateType = "SO002.StopTime " & strGetValue
           Case 2 '���P���
                subDateType = "SO002.DelDate " & strGetValue
           Case Else '������
                subDateType = "SO002.PRTime " & strGetValue
    End Select
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subDateType")
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5540A)
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
    cboDateType.Text = "������"
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "�A�ȰϥN�X", "�A�ȰϦW��")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "��D�N�X", "��D�W��")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "��F�ϥN�X", "��F�ϦW��")
    Call SetgiMulti(gimMduId, "MduId", "Name", "SO017", "�j�ӥN��", "�j�ӦW��")
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "�Ȥ����O�N�X", "�Ȥ����O�W��")
    Call SetgiMultiAddItem(gimStatusCode, "2,3,4,6", "����,���,���P,�����")
    Call SetgiMultiAddItem(gimOrder, "A,B,C,D,E,F", "�Ȥ�s��,�a�},�Ȥ᪬�A,�Ȥ����O,�˾����,�A�Ȱ�")
    gimStatusCode.SelectAll
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  gcnGi.Execute "Drop View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5540A
End Sub

Private Sub gdaDate1_GotFocus()
  On Error Resume Next
  If gdaDate1.GetValue = "" Then gdaDate1.SetValue (RightDate)
End Sub

Private Sub gdaDate2_GotFocus()
  On Error Resume Next
  If gdaDate1.GetValue = "" Or gdaDate2.GetValue = "" Then gdaDate2.SetValue (gdaDate1.GetValue)
End Sub

Private Sub gdaDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaDate1, gdaDate2)
End Sub

Private Sub gdaShouldDate1_GotFocus()
  On Error Resume Next
  If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetValue (RightDate)
End Sub

Private Sub gdaShouldDate2_GotFocus()
  On Error Resume Next
  If gdaShouldDate1.GetValue = "" Or gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue (gdaShouldDate1.GetValue)
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

Private Sub gilServiceType_Change()
  On Error GoTo ChkErr
    GiMultiFilter gimClassCode, gilServiceType.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Function subCreateView() As Boolean
  On Error Resume Next
    gcnGi.Execute "Drop View " & strViewName
  On Error GoTo ChkErr
    Dim strViewSql As String
    Dim strAreaCode As String
    Dim strAreaName As String
        If optAreaCode.Value Then
            strAreaCode = "SO014.AreaCode"
            strAreaName = "SO014.AreaName"
        Else
            strAreaCode = "''"
            strAreaName = "''"
        End If
        strViewName = GetTmpViewName
        
        strViewSql = "Create View " & strViewName & " as (" & _
                        "SELECT SO001.CUSTID,SO001.CUSTNAME,SO001.TEL1,SO002.INSTTIME,SO002.CUSTSTATUSNAME,SO002.STOPTIME,SO002.PRTIME,SO002.DELDATE,SO002.STOPNAME,SO002.PRNAME,SO002.DELNAME,SO001.CLASSNAME1,SO017.Name,SO001.INSTADDRESS,SO017.Note," & strAddrSort & " AddrSort,SO002.CUSTSTATUSCODE,SO001.ClassCode1," & _
                        strAreaCode & " AreaCode," & strAreaName & " AreaName,SO001.ServCode,SO001.ServArea,SO001.ClctAreaCode,SO001.ClctAreaName" & _
                        " From " & strUseTable & " Where " & strChoose & ")"

        gcnGi.Execute strViewSql
        SendSQL strViewSql, True
    Exit Function
ChkErr:
    ErrSub Me.Name, "subCreateView"
End Function
