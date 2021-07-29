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

Public Enum giOpenFileType
    giCreateTEXT = 0
    giOpenTEXT = 1
End Enum

Public Enum GiDTtpye
    GiDate = 1
    GiTime = 2
    GiYM = 3
    GiYear = 4
    GiMonth = 5
    GiDay = 6
    GiHour = 7
    GiMinute = 8
    GiSecond = 9
    GiMD = 10
End Enum

Public Enum giEditModeEnu
    giEditModeView = 0      ' 0=��ܼҦ�
    giEditModeEdit = 1      ' 1.�s��Ҧ�(EditMode)
    giEditModeInsert = 2    ' 2=�s�W�Ҧ�
    giEditModeDelete = 3
End Enum

Public Enum giVType
    giLongV = 0
    giStringV = 1
    giDateV = 2
End Enum

Public Enum ChangeDB
    giAccessDb = 0 'Access Table
    giOracle = 1    'Oracle Table
End Enum

Enum giArrow
    giLeft = 0
    giRight = 1
    giMiddle = 2
End Enum

'INI�ɮ׸��|
Public Property Let FilePath(ByVal vData As String)
    objStorePara.FilePath = vData
End Property
'�J�b���
Public Property Let RealDate(ByVal vData As String)
    objStorePara.RealDate = vData
End Property
'��l�ɮ׸��|
Public Property Let SourcePath(ByVal vData As String)
    objStorePara.SourcePath = vData
End Property
'���ʤH��
Public Property Let UpdEn(ByVal vData As String)
    objStorePara.UpdEn = vData
End Property
'��b�{���W��
Public Property Let PrgName(ByVal vData As String)
    objStorePara.PrgName = vData
End Property
'�W�h��Connection
Public Property Set ugcnGi(ByVal vData As ADODB.Connection)
    Set objStorePara.ugcnGi = vData
End Property
'�Ȧ�W��
Public Property Let BankName(ByVal vData As String)
    objStorePara.BankName = vData
End Property
'�Ȧ�N�X
Property Let BankId(ByVal vData As String)
    objStorePara.BankId = vData
End Property
'�I�ں���
Public Property Let uPTCode(ByVal vData As String)
    objStorePara.uPTCode = vData
End Property
'�I�ں����W��
Public Property Let uPTName(ByVal vData As String)
    objStorePara.uPTName = vData
End Property
'�A�Ⱥ���
Public Property Let ServiceType(ByVal vData As String)
    objStorePara.ServiceType = vData
End Property

'���q�O
Public Property Let CompCode(ByVal vData As String)
    objStorePara.CompCode = vData
End Property

'�O���O�_��b���\
Public Property Get UpdState() As Boolean
     UpdState = objStorePara.UpdState
End Property

Public Function InitObject(ByVal clsName As String) As Object
    On Error GoTo ChkErr
        Dim objCls As Object
        Set objCls = CreateObject("MediaIn4." & clsName)
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
'���O�H��
Public Property Let uClctEn(ByVal vData As String)
    objStorePara.uClctEn = vData
End Property
'���O�H���W��
Public Property Let uClctName(ByVal vData As String)
    objStorePara.uClctName = vData
End Property

'���O�覡
Public Property Let uCMCode(ByVal vData As String)
    objStorePara.uCMCode = vData
End Property
'���O�覡�W��
Public Property Let uCMName(ByVal vData As String)
    objStorePara.uCMName = vData
End Property
'OwnerName
Public Property Let uGetOwner(ByVal vData As String)
    objStorePara.uGetOwner = vData
End Property

'���O����
Public Property Let uCitemQty(ByVal vData As String)
    objStorePara.uCitemQry = vData
End Property

'�y�{����
Public Property Let FlowId(ByVal vData As String)
    objStorePara.FlowId = vData
End Property

Public Property Let uRefNo(ByVal vData As Integer)
    objStorePara.uRefNo = vData
End Property
Public Property Let uUPDTime(ByVal vData As String)
    objStorePara.uUPDTime = vData
End Property