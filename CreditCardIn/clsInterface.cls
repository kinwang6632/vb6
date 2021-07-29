VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'INI�ɮ׸��|
Public Property Let ExportFormat(ByVal vData As Long)
    objStorePara.ExportFormat = vData
End Property
Public Property Let FilePath(ByVal vData As String)
    objStorePara.FilePath = vData
End Property
'�J�b���
Public Property Let RealDate(ByVal vData As Variant)
    objStorePara.RealDate = vData
End Property
'��l�ɮ׸��|
Public Property Let SourcePath(ByVal vData As Variant)
    objStorePara.SourcePath = vData
End Property
'���ʤH��
Public Property Let UpdEn(ByVal vData As Variant)
    objStorePara.UpdEn = vData
End Property
'���ʤH��
Public Property Let UpdName(ByVal vData As Variant)
    objStorePara.UpdName = vData
End Property
'��b�{���W��
Public Property Let PrgName(ByVal vData As Variant)
    objStorePara.PrgName = vData
End Property
'�W�h��Connection
Public Property Set ugcnGi(ByVal vData As ADODB.Connection)
    Set objStorePara.ugcnGi = vData
End Property
'�Ȧ�W��
Public Property Let BankName(ByVal vData As Variant)
    objStorePara.BankName = vData
End Property
'�I�ں���
Public Property Let PTCode(ByVal vData As Variant)
    objStorePara.PTCode = vData
End Property
'�I�ں����W��
Public Property Let PTName(ByVal vData As Variant)
    objStorePara.PTName = vData
End Property
'�A�Ⱥ���
Public Property Let ServiceType(ByVal vData As Variant)
    objStorePara.ServiceType = vData
End Property
'���q�O
Public Property Let CompCode(ByVal vData As Variant)
    objStorePara.CompCode = vData
End Property
'�O���O�_��b���\
Public Property Get UpdState() As Boolean
     UpdState = objStorePara.UpdState
End Property
'�Ȧ�s��
Public Property Let BankId(ByVal vData As String)
      objStorePara.BankId = vData
End Property
'���O����
Public Property Let uCitemQty(ByVal vData As String)
    objStorePara.uCitemQry = vData
End Property
'���O�H��
Public Property Let uClctEn(ByVal vData As Variant)
    objStorePara.uClctEn = vData
End Property
Public Property Let uClctName(ByVal vData As Variant)
    objStorePara.uClctName = vData
End Property
Public Property Let pOwnerName(ByVal vData As String)
      frmUnion.TableOwnerName = vData
      GetOwner = vData
      strOwnerName = vData
End Property

'�I�O�覡
Public Property Let uCMCode(ByVal vData As Variant)
    objStorePara.uCMCode = vData
End Property
Public Property Let uCMName(ByVal vData As Variant)
    objStorePara.uCMName = vData
End Property
Public Property Let uFubonIntegrate(ByVal vData As Boolean)
    objStorePara.uFubonIntegrate = vData
End Property
Public Function InitObject(ByVal clsName As String) As Object
    On Error GoTo ChkErr
        Dim objCls As Object
        Set objCls = CreateObject("CreditCardIn." & clsName)
        Set InitObject = objCls
    Exit Function
ChkErr:
    Err.Description = "�����l�ƥ��ѡI���ˬd�{���W�١I"
End Function

Private Sub Class_Initialize()
    On Error GoTo ChkErr
        Set objStorePara = New clsStoreParameter
    Exit Sub
ChkErr:
    ErrSub "clsInterface", "Class_Initialize"
End Sub