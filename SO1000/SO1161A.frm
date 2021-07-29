VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.3#0"; "GiGridR.ocx"
Begin VB.Form frmSO1161A 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�����O���s�� [SO1161A]"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   3870
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
   Icon            =   "SO1161A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   4455
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   7858
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
Attribute VB_Name = "frmSO1161A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�o�OBruce ���W�g����I
'�̫�ק����G89/03/04
'�����W�h�u���ǻ� uCustId �Ȥ�s��

Option Explicit
' �O���ثe�b�s��B�s�W���˵��Ҧ�
Private lngEditMode As giEditModeEnu
Private strParam0 As String
Private rsSo1161A As New ADODB.Recordset
Attribute rsSo1161A.VB_VarHelpID = -1
Private m_CustID As Long
Private m_CustStatus As Integer

'�O���ثe�ϥΪ��v��
Private bleUserPrivAddNew   As Boolean   '�s�W
Private bleUserPrivUpdate   As Boolean  '�ק�
Private bleUserPrivPrint As Boolean   '�C�L

Public Property Let uCustId(ByVal vData As Long)
  On Error GoTo ChkErr
    m_CustID = vData
  Exit Property
ChkErr:
    ErrSub Me.Name, "Let uCustID"
End Property

Public Property Let uCustStatus(ByVal vData As Integer)
  On Error GoTo ChkErr
    m_CustStatus = vData
  Exit Property
ChkErr:
    ErrSub Me.Name, "Let uCustStatus"
End Property

Public Sub CancelGo()
  On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
    Set ObjActiveForm = Me
    strActFrmName = "frmSO1161A"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ChkErr
        If KeyCode = vbKeyEscape Then Unload Me
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
  '��global�ܼƤ��e,����Ƥ��e
 On Error GoTo ChkErr
   Dim mFlds As New GiGridFlds
'   If Alfa2 = True Then
'      GetGlobal
'      m_CustID = 501
'   End If
    If Crm Then
        PolyFrmFunction Me
        Me.Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2 + 1234)
    Else
        If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, ((Screen.Height - Me.Height) / 2) + 1543
    End If
   
   strParam0 = "select A.InDate,A.OutDate,A.PINCODE,A.PRFlag,A.PrPinCode,A.PrSNo,B.Address FROM " & GetOwner & "SO015 A," & GetOwner & "SO014 B WHERE custid=" & m_CustID & " and A.AddrNo=B.AddrNO  order by Indate"
   'for get nextcode
   With rsSo1161A
        .CursorLocation = adUseClient
        .Open strParam0, gcnGi, adOpenForwardOnly, adLockReadOnly
   End With
   With mFlds
        .Add "InDate", giControlTypeDate, , , , "�w�ˤ��  "
        .Add "OutDate", giControlTypeDate, , , , " �E�X���  "
        .Add "PINCODE", , , , , "�Q�P���X  "
        .Add "PRFlag", giControlTypeTextBox, , , , "���A      "
        .Add "PrPinCode", , , , , "����Q�P���X "
        .Add "PrSNo", , , , , "����渹        "
        .Add "Address", , , , , "�¦a�}" & Space(33)
   End With
  ' Print adoData.Recordset.RecordCount
   With ggrData
       .AllFields = mFlds
       .SetHead
        Set .Recordset = rsSo1161A
       .Refresh
   End With
   MenuEnabled , , , , , True
   'ggrData.Blank = True
 Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error GoTo ChkErr
    If rsSo1161A.State = adStateOpen Then rsSo1161A.Close
    Set rsSo1161A = Nothing
    Call FormQueryUnload
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_QueryUnload"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    ReleaseCOM Me
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If Fld.Name = "PRFLAG" And Value > 0 Then
        Select Case Value
            Case 2
                Value = "���"
            Case 3
                Value = "����"
            Case Else
                Value = ""
        End Select
    End If
    '2=���,3=����
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ggrData_ShowCelData")
End Sub

