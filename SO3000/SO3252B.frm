VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO3252B 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "����s��"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11730
   Icon            =   "SO3252B.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   11730
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "����(&X)"
      Height          =   405
      Left            =   7740
      TabIndex        =   2
      Top             =   5910
      Width           =   1425
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�T�w"
      Height          =   405
      Left            =   2400
      TabIndex        =   1
      Top             =   5910
      Width           =   1425
   End
   Begin prjGiGridR.GiGridR ggrSelectData 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   9763
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
End
Attribute VB_Name = "frmSO3252B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�ھ�#3042 �ݨD �W�[��Form�A���U�T�w���s��}�l����@�o By Kin 2007/03/22
Option Explicit
Private strCompCode As String
Private strSQLWhere As String
Private strReasno As String
Private strCancelCode As String
Private strCancelName As String
Private rsUpdate As ADODB.Recordset
Private strUpdTime As String
Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
End Sub
Private Sub cmdOk_Click()
  On Error GoTo ChkErr
    Dim lngAffectCounts As Long
    Dim lngUpdateCount As Long
    Dim lngErrCount As Long
    gcnGi.BeginTrans
        '��s���
        Screen.MousePointer = vbHourglass
        If ExecuteSQL("Update " & GetOwner & "So033 Set UCCode=Null,UCName=Null,CancelFlag=1,RealDate=To_Date('" & Format(RightDate, "YYYY/MM/DD") & "','YYYY/MM/DD')," & _
                      "Note = Trim(Nvl(Note,'')) || ';�@�o��]:" & strReasno & "; ������]:' || UCName,CancelCode = " & strCancelCode & "," & _
                      "CancelName = '" & strCancelName & "',UpdEn = '" & garyGi(1) & "',UpdTime ='" & strUpdTime & "'  Where " & strSQLWhere, gcnGi, lngAffectCounts, False, False) <> giOK Then
            MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I": GoTo ChkErr
        End If
        strSQLWhere = SetAddComma(strSQLWhere)
        If ExecuteSQL("Insert into " & GetOwner & "So070 (UpdTime,UpdName,Reason,Para,RcdCount) values('" & strUpdTime & "','" & garyGi(1) & "'," & IIf(Trim(strCancelName) <> "", "'" & Trim(strCancelName) & "'", Null) & ",'" & Replace(strSQLWhere, "'", "") & "'," & IIf(lngUpdateCount < 0, 0, lngUpdateCount) & ")", gcnGi, , , False) <> giOK Then
            MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
        End If
    gcnGi.CommitTrans
    MsgBox "�@�o�����A����=" & IIf(lngAffectCounts < 0, 0, lngAffectCounts), vbInformation, "�T���I"
    Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub
ChkErr:
    gcnGi.RollbackTrans
    ErrSub Me.Name, "cmdOk_Click"
End Sub

Private Sub Form_Activate()
On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF2 And cmdOK.Enabled Then Call cmdOk_Click: Exit Sub
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call SetGrd '�]�wGrid��Caption
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

'���q�O
Public Property Let uCompCode(ByVal vData As String)
    strCompCode = vData
End Property
'�e���Ҧ�X�Ӫ�����
Public Property Let uSQLWhere(ByVal vData As String)
    strSQLWhere = vData
End Property
'�e�ݬd�ߥX�Ӫ����
Public Property Let uRS(ByRef vData As ADODB.Recordset)
    Set rsUpdate = vData
End Property
'�Ƶ�
Public Property Let uReasno(ByVal vData As String)
    strReasno = vData
End Property
'�@�o��]�N�X
Public Property Let uCancelCode(ByVal vData As String)
    strCancelCode = vData
End Property
'�@�o��]
Public Property Let uCancelName(ByVal vData As String)
    strCancelName = vData
End Property
'��s�ɶ�
Public Property Let uUpdTime(ByVal vData As String)
    strUpdTime = vData
End Property

Private Sub SetGrd()
   On Error GoTo ChkErr
   Dim mflds As New prjGiGridR.GiGridFlds
      With mflds
         '.Add "Chose", , , , False, "���", vbCenter
'         .Add "CustId", , , , False, "�Ȥ�s��", vbLeftJustify
         .Add "BillNo", , , , False, "       ��ڽs��     ", vbLeftJustify
         .Add "CITEMNAME", , , , False, "     ���O���ئW��    ", vbLeftJustify
         .Add "ShouldAmt", , , , False, "  �X����B  ", vbRightJustify
         .Add "RealPeriod", giControlTypeComboBox, , , False, " ���� ", vbRightJustify
         .Add "ShouldDate", giControlTypeDate, , , False, "  �������  ", vbLeftJustify
         .Add "RealStartDate", giControlTypeDate, , , False, "   �_�l��   ", vbLeftJustify
         .Add "RealStopDate", giControlTypeDate, , , False, "   �I���   ", vbLeftJustify
         .Add "UCName", , , , False, "             ������]                ", vbLeftJustify
'         .Add "GUINO", , , , False, "    �o�����X    "
      End With
      ggrSelectData.AllFields = mflds
      ggrSelectData.SetHead
      ggrSelectData.Enabled = True
      Set ggrSelectData.Recordset = rsUpdate
      ggrSelectData.Blank = False
      ggrSelectData.GotoRow 1
      ggrSelectData.Refresh
    Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "subGrd4")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CloseRecordset rsUpdate
End Sub
Private Function SetAddComma(ByVal strSQL As String) As String
    On Error GoTo ChkErr
        Dim strNewSQl As String
        Dim strTmpSql  As String
        Dim intSitus As Integer
        Do While True
            intSitus = InStr(1, strSQL, "'", vbTextCompare)
            If intSitus <= 0 Then Exit Do
            strTmpSql = strTmpSql & Left(strSQL, intSitus) & "'"
            strSQL = Mid(strSQL, intSitus + 1)
        Loop
        SetAddComma = strTmpSql
    Exit Function
ChkErr:
    ErrSub Me.Name, "SetAddComma"
End Function

