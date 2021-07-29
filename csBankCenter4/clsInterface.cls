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

Public Enum ChangeDB
    giAccessDb = 0 'Access Table
    giOracle = 1    'Oracle Table
End Enum

Public Enum giVType
    giLongV = 0
    giStringV = 1
    giDateV = 2
End Enum

Public Enum giArrow
    giLeft = 0
    giRight = 1
    giMiddle = 2
End Enum

Public Enum giOpenFileType
    giCreateTEXT = 0
    giOpenTEXT = 1
End Enum

'�ۭq�C�|��
Public Enum giEditModeEnu
    giEditModeView = 0      ' 0=��ܼҦ�
    giEditModeEdit = 1      ' 1.�s��Ҧ�(EditMode)
    giEditModeInsert = 2    ' 2=�s�W�Ҧ�
    giEditModeDelete = 3    ' 3=�R���Ҧ�
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

'Choose SQL
Public Property Let uChooseSQL(ByVal vData As Variant)
    clsStoreParameter.uChooseSQL = vData
End Property

'Choose SQL
Public Property Let uChoose33(ByVal vData As Variant)
    clsStoreParameter.uChoose33 = vData
End Property

'INI�ɮ׸��|
Public Property Let uErrPath(ByVal vData As Variant)
    clsStoreParameter.uErrPath = vData
End Property


'�ɮ׸��|
Public Property Let uFilePath(ByVal vData As Variant)
    clsStoreParameter.uFilePath = vData
End Property


'�J�b���
Public Property Let uRealDate(ByVal vData As Variant)
    clsStoreParameter.uRealDate = vData
End Property


'���ʤH���N�X
Public Property Let uUpdEn(ByVal vData As Variant)
    clsStoreParameter.uUpdEn = vData
End Property


'���ʤH���W��
Public Property Let uUpdName(ByVal vData As Variant)
    clsStoreParameter.uUpdName = vData
End Property


'�W�h��Connection
Public Property Set ugcnGi(ByVal vData As ADODB.Connection)
    Set clsStoreParameter.ugcnGi = vData
End Property


'�W�h��Connection
Public Property Set uCnn(ByVal vData As ADODB.Connection)
    Set clsStoreParameter.uCnn = vData
End Property
Public Property Let uBankId(ByVal vData As String)
    clsStoreParameter.uBankId = vData
End Property


'�Ȧ�N�XSQL
Public Property Let uBankSQL(ByVal vData As Variant)
    clsStoreParameter.uBankSQL = vData
End Property


'�I�ں���
Public Property Let uPTCode(ByVal vData As Variant)
    clsStoreParameter.uPTCode = vData
End Property
Public Property Let uPTName(ByVal vData As Variant)
    clsStoreParameter.uPTName = vData
End Property


'�I�O�覡
Public Property Let uCMCode(ByVal vData As Variant)
    clsStoreParameter.uCMCode = vData
End Property
Public Property Let uCMName(ByVal vData As Variant)
    clsStoreParameter.uCMName = vData
End Property



'�A�Ⱥ���
Public Property Let uServiceType(ByVal vData As Variant)
    clsStoreParameter.uServiceType = vData
End Property


'���q�O
Public Property Let uCompCode(ByVal vData As Variant)
    clsStoreParameter.uCompCode = vData
End Property


'���O�H��
Public Property Let uClctEn(ByVal vData As Variant)
    clsStoreParameter.uClctEn = vData
End Property
Public Property Let uClctName(ByVal vData As Variant)
    clsStoreParameter.uClctName = vData
End Property

'�X���ɮ�
Public Property Let uBeOne(ByVal vData As Boolean)
    clsStoreParameter.uBeOne = vData
End Property

'�X���ɮ�--�n��榡
Public Property Let uNT(ByVal vData As Boolean)
    clsStoreParameter.uNT = vData
End Property

'�D�n�Ȧ�
Public Property Let uBankHand(ByVal vData As String)
    clsStoreParameter.uBankHand = vData
End Property
'OwnerName
Public Property Let uGetOwner(ByVal vData As String)
    clsStoreParameter.uGetOwner = vData
End Property
'���O����
Public Property Let uCitemQty(ByVal vData As String)
    clsStoreParameter.uCitemQry = vData
End Property

'�y�{����
Public Property Let FlowId(ByVal vData As String)
    clsStoreParameter.FlowId = vData
End Property

'�n��M��
Public Property Let uChkNT(ByVal vData As Integer)
    clsStoreParameter.uChkNT = vData
End Property

'�t�����B�O�_����
Public Property Let uCkZero(ByVal vData As Boolean)
    clsStoreParameter.uCkZero = vData
End Property

'�ϥή榡
Public Property Let uStyle(ByVal vData As Integer)
    clsStoreParameter.uStyle = vData
End Property

'�帹
Public Property Let uBatch(ByVal vData As String)
    clsStoreParameter.uBatch = vData
End Property

Public Function InitObject(ByVal clsName As String) As Object
    On Error GoTo chkErr
        Dim objCls As Object
        Set objCls = CreateObject(clsName)
        Set InitObject = objCls
    Exit Function
chkErr:
    Err.Description = "�����l�ƥ��ѡI���ˬd�{���W�١I"
End Function


Private Sub Class_Initialize()
    On Error GoTo chkErr
        Set clsStoreParameter = New clsStoreParameter
    Exit Sub
chkErr:
    ErrSub "clsInterface", "Class_Initialize"
End Sub

Public Property Let uChkProc(ByVal vData As Boolean)
    clsStoreParameter.uChkProc = vData
End Property
