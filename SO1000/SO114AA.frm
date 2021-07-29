VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO114AA 
   BorderStyle     =   1  '��u�T�w
   Caption         =   $"SO114AA.frx":0000
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
   Icon            =   "SO114AA.frx":001F
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
Attribute VB_Name = "frmSO114AA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'CM�PCT��\�໡��:
'��ϥΪ�Click CM Button or CT Button��, �h�t�}�@New Form���so152����Ƥ��e.�p�� :
'Gird������ܪ����e�Y��so152���������e(���ެOClick [CM] or [CT] button), ��Ū����ƪ��y�k�p�U :
'select * from so152 where EmcCustID=so001.custid and EmcCompCode=so001.CompCode order by UpdTime desc
'Gird��줤��Caption���峡��, �аѦ�EcSchema.mdb����so152���c���

Private rsSO152 As New ADODB.Recordset

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    If Crm Then
        PolyFrmFunction Me
        Me.Move frmSO1100A.Left + ((frmSO1100A.Width - Me.Width) / 2), frmSO1100A.Top + ((frmSO1100A.Height - Me.Height) / 2)
    Else
        If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, ((Screen.Height - Me.Height) / 2) + 1800
    End If
    GetRS rsSO152, "select * from " & GetOwner & "so152 where emccustid=" & gCustId & " and emccompcode=" & gCompCode & " order by updtime desc", gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly
    GrdFmt
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Public Sub GrdFmt()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    With mFlds
        .Add "EmcCompCode", , , , , "EMC���q�O", vbLeftJustify
        .Add "EmcCustID", , , , , "EMC�Ȥ�s��", vbLeftJustify
        .Add "EbtContractNo", , , , , "�X���s��         ", vbLeftJustify
        .Add "EbtCustID", , , , , "EBT�Ȥ�s��", vbLeftJustify
        .Add "EbtContractBDate", giControlTypeDate, , , , " �X���ͮĤ� ", vbLeftJustify
        .Add "EbtContractEDate", giControlTypeDate, , , , " �X���I��� ", vbLeftJustify
        .Add "EbtCustCName", , , , , "�Ȥᤤ��W��  ", vbLeftJustify
        .Add "EbtCustContactPhone", , , , , "�Ȥ��p���q��", vbLeftJustify
        .Add "EbtCustContactMobile", , , , , "�Ȥ��ʹq��", vbLeftJustify
        .Add "EbtCompOwnerName", , , , , "���q�t�d�H����m�W", vbLeftJustify
        .Add "EbtContactPhone", , , , , "(���q)�p���H�p���q��", vbLeftJustify
        .Add "EbtItContactName", , , , , "�޳N�p���H����m�W", vbLeftJustify
        .Add "EbtItContactPhone", , , , , "�޳N�p���H�q��", vbLeftJustify
        .Add "EbtItContactMobile", , , , , "�޳N�p���H��ʹq��", vbLeftJustify
        .Add "EbtItEMail", , , , , "�޳N�p���H�q�l�l��  ", vbLeftJustify
        .Add "EbtInstAddr", , , , , "�˾��a�}                      ", vbLeftJustify
        .Add "EbtCustAddr", , , , , "���y�a�}                      ", vbLeftJustify
        .Add "EbtBillAddr", , , , , "�b��a�}                      ", vbLeftJustify
        .Add "EbtContractStatusCode", , , , , "�X�����A�N�X", vbLeftJustify
        .Add "EbtContractStatusDesc", , , , , "�X�����A����", vbLeftJustify
        .Add "EbtFeePeriodCode", , , , , "EBTú�O�g���N�X", vbLeftJustify
        .Add "EbtFeePeriodDesc", , , , , "EBTú�O�g������", vbLeftJustify
        .Add "EbtServiceType", , , , , "�A�����O", vbLeftJustify
        .Add "EbtAgentName", , , , , "�N�z�H�m�W ", vbLeftJustify
        .Add "EbtAgentPhone", , , , , "�N�z�H�q�� ", vbLeftJustify
        .Add "EbtAgentID", , , , , "�N�z�H�����Ҹ�", vbLeftJustify
        .Add "EbtAgentAddress", , , , , "�N�z�H�a�}                      ", vbLeftJustify
        .Add "EbtIdCardId", , , , , "�ӤH�����Ҹ�", vbLeftJustify
        .Add "EbtCompanyOwnerId", , , , , "�t�d�H�����Ҹ�", vbLeftJustify
        .Add "EbtNonProfitCompanyId", , , , , "�D��Q�k�H�Ҹ�", vbLeftJustify
        .Add "EbtCompanyId", , , , , "���q�ҷӸ��X", vbLeftJustify
        .Add "IfMoveToOtherSo", , , , , "�O�_�w�g����", vbLeftJustify
        .Add "EbtNotes", , , , , "EBT�Ƶ�" & Space(24), vbLeftJustify
        .Add "UpdTime", giControlTypeTime, , , , "���ʮɶ�            ", vbLeftJustify
        .Add "UpdEn", , , , , "���ʤH��   ", vbLeftJustify
'        .Add "EbtAgentZipCode", , , , , "�N�z�H�l���ϸ�", vbLeftJustify
'        .Add "EbtAgentCity", , , , , "�N�z�H����", vbLeftJustify
'        .Add "EbtAgentTown", , , , , "�N�z�H�m��", vbLeftJustify
    End With
    With ggrData
        .AllFields = mFlds
        .SetHead
         Set .Recordset = rsSO152
        .Refresh
    End With
    Me.Show 1
  Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    rsSO152.Close
    Set rsSO152 = Nothing
    ReleaseCOM Me
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If giFld.FieldName = "IfMoveToOtherSo" Then Value = IIf(Val(Value & "") = 0, "", "�O")
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ggrData_ShowCellData"
End Sub
