VERSION 5.00
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO55C0A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�Ʀ�A�ȭq��Ƴ��� [SO55C0A]"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "SO55C0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   8010
   StartUpPosition =   1  '���ݵ�������
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:�C�L"
      Default         =   -1  'True
      Height          =   525
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "����(&X)"
      Height          =   525
      Left            =   6000
      TabIndex        =   5
      Top             =   1800
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "�W���έp���G"
      Height          =   525
      Left            =   2880
      TabIndex        =   4
      Top             =   1800
      Width           =   1395
   End
   Begin Gi_YM.GiYM GiDataYm 
      Height          =   330
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   885
      _ExtentX        =   1561
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
   Begin CS_Multi.CSmulti gimChCode 
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   609
      ButtonCaption   =   "�@  ��  �W  �D"
   End
   Begin CS_Multi.CSmulti gimChCodeT 
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   609
      ButtonCaption   =   "��  ��  �W  �D"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "(ex: 95/01)"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2640
      TabIndex        =   7
      Top             =   195
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�έp�~��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   720
      TabIndex        =   6
      Top             =   195
      Width           =   780
   End
End
Attribute VB_Name = "frmSO55C0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strSqlString1, strSqlString2 As String
Dim strGim As String

Private Sub cmdExit_Click()
On Error GoTo ChkErr
  Unload Me
ChkErr:
  Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdPrevRpt_Click()
On Error GoTo ChkErr
  Screen.MousePointer = vbHourglass
  Call PreviousRpt(GetPrinterName(5), RptName("SO55C0"), Me.Caption)
  Screen.MousePointer = vbDefault
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
On Error GoTo ChkErr
  Dim blnGoOn As Boolean
  Dim strSQL1, strSQL2 As String
  
  blnGoOn = ChkDTok And IsDataOk
  If blnGoOn Then
  
    '�����Ȯ� disabled�A�H�K USer ���ƫ�
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    
    Call ReadyGoPrint
    Call subChoose
    
    If subCreateView Then
         '#7074 To Add a choice-condition parameter into  print function by Kin 2015/09/09
        Call PrintRpt2(GetPrinterName(5), RptName("SO55C0"), , "�Ʀ�A�ȭq��Ƴ��� [SO55C0A]", , strChooseString, , True, "Tmp0000.Mdb", , , GiPaperPortrait)
    End If
    
    '������_ enabled
    Me.Enabled = True
    Screen.MousePointer = vbDefault
  End If
ChkErr:
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub
Private Sub subGim() 'GiMulti
  On Error GoTo ChkErr

    SetMsQry gimChCode, "SELECT CodeNo,Description , ChannelID FROM CD024 Where StopFlag=0 ORDER BY Codeno,Description", _
                         "�W�D�s���N�X", "�W�D�s���W��", "", , True, "nagra�W�Did", , , , , , , , , , 1200, 2000, 1500, , , , , , , , , , , , , , , , , , , , , , 1
    SetMsQry gimChCodeT, "SELECT DISTINCT CodeNo,Description  FROM CD024B ORDER BY To_Number(Codeno),Description", _
                         "�����W�D�N�X", "�����W�D�W��", "", , True, , , , , , , , , , , 1200, 1800, 1500, , , , , , , , , , , , , , , , , , , , , , 1
    
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ChkErr
  If ChkGiList(KeyCode) Then
    Select Case KeyCode
      Case vbKeyF5
        cmdPrint.Value = cmdPrint.Enabled
      Case vbKeyEscape
        cmdExit.Value = True
    End Select
  End If
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
On Error GoTo ChkErr
  Call subGim
  GiDataYm.SetValue (RightDate)
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo ChkErr
  IsDataOk = True
  If GiDataYm.Text = "" Then
    IsDataOk = False
    MsgBox "�Х����w�έp�~��C", vbExclamation, "ĵ�i"
    GiDataYm.SetFocus
  End If
ChkErr:
  Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub subChoose()
  On Error GoTo ChkErr
    Dim lngYear, lngMonth As Long
    Dim strSQL1, strSQL2, strSQL3, strSQL4, strRptYm, StrTableNameA As String
    Dim StrTableNameB, StrTableNameC, strSoDescription, strSoName As String
    Dim rsTmp As New ADODB.Recordset
    strChoose = ""
    strRptYm = GiDataYm.GetValue
    lngYear = strRptYm \ 100
    lngMonth = strRptYm Mod 100
    
    If gimChCode.GetQueryCode <> "" Then Call subAnd("Codeno in (" & gimChCode.GetQueryCode & ")")
    If gimChCodeT.GetQryStr <> "" Then Call subAnd("Ch_code " & gimChCodeT.GetQryStr & "")

    '�̤��q�O�N����X���q�W��
    Set rsTmp = gcnGi.Execute("Select Description, CompName From So507 Where CodeNo='" & garyGi(9) & "'")
    strSoDescription = rsTmp("Description")
    strSoName = rsTmp("CompName") & "���u�q���ѥ��������q�Ʀ�A�ȭq��ƻ�����"
    Set rsTmp = Nothing
    
    StrTableNameA = strSoDescription & "_SO004_" & strRptYm
    StrTableNameB = strSoDescription & "_SO001ALL_" & strRptYm
    StrTableNameC = strSoDescription & "_SO005_" & strRptYm
    
    
    '�`�������
    strSQL1 = "Select Count(*) As CustCount "
    strSQL1 = strSQL1 & "From ( "
    strSQL1 = strSQL1 & "  Select Case When (�Ȥ����O='�E�Ӥ�����') Then '����' "
    strSQL1 = strSQL1 & "              When (�U���~��<" & strRptYm & ") Then '�O��' "
    strSQL1 = strSQL1 & "              When (�Ȥ����O<>'�j�ӲΦ�') AND (�U���~��=" & strRptYm & ") Then '�O��' "
    strSQL1 = strSQL1 & "         Else '���`' "
    strSQL1 = strSQL1 & "         End �P�w���� "
    strSQL1 = strSQL1 & "  From ( "
    strSQL1 = strSQL1 & "    Select Case When ClassCode1 Between 11000 And 11999 Then '�@���' "
    strSQL1 = strSQL1 & "                When ClassCode1 Between 12000 And 12999 Then '�u�ݤ�' "
    strSQL1 = strSQL1 & "                When ClassCode1 Between 13000 And 13999 Then '�K�O��' "
    strSQL1 = strSQL1 & "                When ClassCode1 Between 14000 And 14999 Then '�j�ӲΦ�' "
    strSQL1 = strSQL1 & "                When ClassCode1 Between 15000 And 15999 Then '�j�ӭӦ�' "
    strSQL1 = strSQL1 & "                When ClassCode1=91010 Then '�E�Ӥ�����' "
    strSQL1 = strSQL1 & "           End �Ȥ����O, "
    strSQL1 = strSQL1 & "           To_Char(ClctDate,'yyyymm') �U���~�� "
    strSQL1 = strSQL1 & "    From " & StrTableNameB & " "
    strSQL1 = strSQL1 & "    Where (CustStatusCode=1) "
    strSQL1 = strSQL1 & "    And (ClassCode1 Between 11010 And 91010) "
    strSQL1 = strSQL1 & "    And (ClassCode1 Not In (14990,15990)))) "
    strSQL1 = strSQL1 & "Where �P�w���� Not In ('����','�O��') "

    '�`������ơ]���F�^
    'strSQL1 = "Select CustCount " & _
              "From SO511A " & _
              "Where (RptYm=" & strRptYm & ") " & _
              "And DataType=-1 "
              
    '�Ʀ���W���˾����
    strSQL2 = "Select Count(Distinct A.CustId) " & _
              "From " & StrTableNameA & " A, " & StrTableNameB & " B " & _
              "Where (A.CustId = B.CustId) " & _
              "And (B.CustStatusCode = 1) " & _
              "And (A.FaciCode In " & _
              "  (Select CodeNo From Cd022 Where RefNo=3)) " & _
              "And (A.InstDate Is Not Null) " & _
              "And (A.PrDate Is Null) "
              
    '�@��2�x�H�W(�t)���
    strSQL3 = "Select Count(*) " & _
              "from ( " & _
              "  Select A.CustId " & _
              "  From " & StrTableNameA & " A, " & StrTableNameB & " B " & _
              "  Where (A.CustId = B.CustId) " & _
              "  And (B.CustStatusCode = 1) " & _
              "  And (A.FaciCode In " & _
              "    (Select CodeNo From Cd022 Where RefNo=3)) " & _
              "  And (A.InstDate Is Not Null) " & _
              "  And (A.PrDate Is Null) " & _
              "  Group By A.CustId " & _
              "  Having Count(*)>1) "
              
    '�I�O�W�D�q���
'    strSQL4 = "Select Sum(SumCount) From " & StrTableNameC
    strSQL4 = "Select Count(distinct custid) from " & StrTableNameC & _
              IIf(strChoose = "", "", " Where ") & strChoose
              
    '�W�楼�}���A���i�঳������ơA����w�]�� //Dennis
    strSqlString1 = "Select '" & strSoName & "' SoName, " & _
                    " '" & lngYear & "/" & lngMonth & "' DataYm, " & _
                   "(" & strSQL1 & ") CustTotal, " & _
                   "(" & strSQL2 & ") StbCustTotal, " & _
                   "(" & strSQL3 & ") HasTwoStbCustTotal, " & _
                   "(" & strSQL4 & ") PayCustCount, " & _
                   "5 G, " & _
                   "6 H " & _
                   "From Dual "
                   
    '�U�I�O�W�D�q��ơ]�Ӷ��^�A�ثe��쥼�w�q���� //Dennis
'    strSqlString2 = "Select Description, Price, Sum(SumCount) As SumCount From " & StrTableNameC & " Group By Description, Price"
    strSqlString2 = "Select CH_Desc,Price Price1,Count(*) Count1 From " & StrTableNameC & _
                    IIf(strChoose = "", "", " Where ") & strChoose & " Group by CH_Desc,Price"

  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Function subCreateView() As Boolean
On Error Resume Next
  'gcnGi.Execute "Drop Table " & strViewName1
  'gcnGi.Execute "Drop Table " & strViewName2
On Error GoTo ChkErr
  'Dim strView As String
  Dim rsTemp As ADODB.Recordset
  
  Set rsTemp = gcnGi.Execute(strSqlString1)
  Call CreateMDBTable(rsTemp, "SO55C0A1", cnn)
  Call InsertIntoMDB(rsTemp, "SO55C0A1")
  Set rsTemp = Nothing
  
  Set rsTemp = gcnGi.Execute(strSqlString2)
  Call CreateMDBTable(rsTemp, "SO55C0A2", cnn)
  Call InsertIntoMDB(rsTemp, "SO55C0A2")
  Set rsTemp = Nothing
  
  'strViewName1 = GetTmpViewName
  'strView = "Create Table " & strViewName1 & " As (" & strSqlString1 & ")"
  'gcnGi.Execute strView
  'SendSQL strView, True
  'strViewName2 = GetTmpViewName
  'strView = "Create Table " & strViewName2 & " As (" & strSqlString2 & ")"
  'gcnGi.Execute strView
  'SendSQL strView, True
  subCreateView = True
  Exit Function
ChkErr:
  MsgBox "�䤣����ε��[��A�нT�w�έp�~�몺��Ʀ���J���T", , "���~�T��"
  subCreateView = False
'  Call ErrSub(Me.Name, "subCreateView")
End Function

Private Sub InsertIntoMDB(ByVal rs As ADODB.Recordset, TableName As String)
On Error GoTo ChkErr
  Dim strString1, strString2 As String
  Dim lngIndex As Integer

  While Not rs.EOF
    strString1 = "Insert Into " & TableName & " ( "
    For lngIndex = 0 To rs.Fields.Count - 1
    
      strString1 = strString1 & rs.Fields(lngIndex).Name
      If lngIndex <> rs.Fields.Count - 1 Then
        strString1 = strString1 & ", "
      End If
    Next
    strString1 = strString1 & ")  Values ( "
    For lngIndex = 0 To rs.Fields.Count - 1
      strString2 = rs(rs.Fields(lngIndex).Name) & ""
      Select Case rs.Fields(lngIndex).Type
        Case adBigInt, adVarNumeric, adNumeric, adDecimal, adSingle, adDouble, adInteger '"Double"
          strString2 = NoZero(strString2, True)
        End Select
      strString1 = strString1 & "'" & strString2 & "'"
      If lngIndex <> rs.Fields.Count - 1 Then
        strString1 = strString1 & ", "
      End If
    Next
    strString1 = strString1 & ") "
    cnn.Execute (strString1)
    rs.MoveNext
  Wend
  cnn.Close
  cnn.Open
ChkErr:
  Call ErrSub(Me.Name, "InsertIntoMdb")
End Sub

Private Sub gimChCode_LostFocus()
    If gimChCode.GetQueryCode <> "" Then strGim = " WHERE CHANNELID IN (" & gimChCode.GetQryFld(3) & ") " Else strGim = ""
    SetMsQry gimChCodeT, "SELECT DISTINCT  CodeNo,Description  FROM CD024B " & strGim & " ORDER BY To_Number(Codeno),Description", _
                 "�����W�D�N�X", "�����W�D�W��", "", , True, , , , , , , , , , , 1200, 1800, 1500, , , , , , , , , , , , , , , , , , , , , , 1

End Sub
