VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1100M 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�������Ӹ�� [SO1100M]"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   1530
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO1100M.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   4080
      Left            =   -15
      TabIndex        =   0
      ToolTipText     =   "�Ы�����⦸,����Ȥ���"
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   7197
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
End
Attribute VB_Name = "frmSO1100M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private rsInstallment As New ADODB.Recordset

'�]�Ƥ����I�ڬd�߾���
'SO1100B���]�ƭ��ҥ[�@[�������Ӹ��]Button, ��User�I��]��Gird����STB�]�ƫ�
'�AClick��Button��, �YShow��STB��������Ʃ��ө�@New Form��Gird��,
'��Gird���榡���e�p���O��ƭ��Ҥ�Gird���e.������Ƴ���
'�ݥ[�L�o����SO033 or SO034�� FaciSeqNo=SO004.SeqNo and
'SNO=SO004.SNO and Budget = 1(�]���ڭ̥u�n���STB������������)
'Default�HSNO, ShouldDate��Order By

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = 27 Then Unload Me
End Sub

Public Sub QuerySO033SO034(ByVal FaciSeqNo As String, _
                            ByVal SNO As String)
  On Error GoTo ChkErr
    Dim strQrySO033SO034 As String
    Dim strCDT As String
    Dim strField As String
    strField = ChargeField
    
    strCDT = "CUSTID=" & gCustId & _
                " AND SERVICETYPE='" & gServiceType & "'" & _
                " AND COMPCODE=" & gCompCode & _
                " AND Budget = 1" & _
                " AND FaciSeqNo='" & FaciSeqNo & "'" & _
                " AND SNO='" & SNO & "'"

    strQrySO033SO034 = "SELECT 0 AS TYPE,ROWID," & strField & _
                        " FROM " & GetOwner & "SO033 SO033 WHERE UCCODE IS NOT NULL AND " & _
                         strCDT & _
                        " UNION  ALL " & _
                        "SELECT 1 AS TYPE,ROWID," & strField & _
                        " FROM " & GetOwner & "SO033 SO033 WHERE UCCODE IS NULL AND " & _
                         strCDT & _
                        " UNION ALL " & _
                        "SELECT 2 AS TYPE,ROWID," & strField & _
                        " FROM " & GetOwner & "SO034 SO034 WHERE " & _
                         strCDT & _
                        " ORDER BY SNO,SHOULDDATE"
    GetRS rsInstallment, strQrySO033SO034, gcnGi, adUseClient, adOpenForwardOnly, adLockPessimistic, "(SO033 UNION SO034)", "QuerySO033SO034"
    GrdFmt
  Exit Sub
ChkErr:
    ErrSub Me.Name, "QuerySO033SO034"
End Sub


Private Sub Form_Load()
  On Error GoTo ChkErr
    If Crm Then
        PolyFrmFunction Me
        Me.Move frmSO1100A.Left + ((frmSO1100A.Width - Me.Width) / 2), frmSO1100A.Top + ((frmSO1100A.Height - Me.Height) / 2)
    Else
'        If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 - 200
        If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, ((Screen.Height - Me.Height) / 2) + 1800
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Public Sub GrdFmt()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    With mFlds
        .Add "Type", , , , , " ����  ", vbLeftJustify
        .Add "BillNo", , , , , "      ��ڽs��    ", vbLeftJustify
        .Add "CitemName", , , , , "���O����  ", vbLeftJustify
        .Add "ShouldAmt", , , , , "�X��$", vbRightJustify
        .Add "RealAmt", , , , , "�ꦬ$", vbRightJustify
        .Add "Realperiod", , , , , " �� ", vbRightJustify
        .Add "ShouldDate", giControlTypeDate, , , , "�������", vbLeftJustify
        .Add "CancelFlag", , , , , "�@�o", vbLeftJustify
        .Add "RealDate", giControlTypeDate, , , , "�ꦬ���", vbLeftJustify
        .Add "RealStartDate", giControlTypeDate, , , , "�_�l��", vbLeftJustify
        .Add "RealStopDate", giControlTypeDate, , , , "�I���", vbLeftJustify
        .Add "UCName", , , , , "������]", vbLeftJustify
        .Add "STName", , , , , " �u����]", vbLeftJustify
        .Add "ClctName", , , , , "���O�H��", vbLeftJustify
        .Add "COMPCODE", , , , , "���q�O", vbLeftJustify
        .Add "CMName", , , , , "���O�覡", vbLeftJustify
        .Add "CitibankATM", , , , , "�����b��    " & Space(6), vbLeftJustify
        .Add "GUIno", , , , , "�o�����X  ", vbLeftJustify
        .Add "PreInvoice", , , , , "�O�_�w�}�o��", vbLeftJustify
        .Add "CancelName", , , , , "�@�o��]", vbLeftJustify & Space(2)
        .Add "Note", , , , , "�Ƶ�" & Space(30), , , vbLeftJustify
        .Add "OldAmt", , , , , "���������B", vbLeftJustify
        .Add "PrtSNo", , , , , "�L��Ǹ�" & Space(12), vbLeftJustify
        .Add "UpdTime", , , , , "���ʮɶ�         ", vbLeftJustify
        .Add "UpdEn", , , , , "���ʤH��", vbLeftJustify
    End With
    With ggrData
        .AllFields = mFlds
        .SetHead
         Set .Recordset = rsInstallment
        .Refresh
    End With
    Me.Show 1
  Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    rsInstallment.Close
    Set rsInstallment = Nothing
    ReleaseCOM Me
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If ggrData.Recordset.EOF Then Exit Sub
    With giFld
        If .FieldName = "Type" Then Value = IIf(Value = "2", "�w�鵲", "���鵲")
        If .FieldName = "CancelFlag" Then Value = IIf(Value = 0, "", "�O")
            If .FieldName = "ShouldAmt" Or .FieldName = "OldAmt" Or .FieldName = "RealAmt" Then
                Value = IIf(Value <> 0, Format(Value, "###,###,###"), "")
            End If
            If .FieldName = "PreInvoice" Then
                If Value <> "" Then
                    Select Case Value
                           Case 0
                                Value = "��}"
                           Case 1
                                Value = "�w�}"
                           Case 2
                                Value = "�{��"
                    End Select
                End If
            End If
    End With
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ggrData_ShowCellData"
End Sub

