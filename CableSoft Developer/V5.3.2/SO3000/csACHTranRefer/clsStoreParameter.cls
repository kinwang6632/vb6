VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStoreParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private strErrPath As Variant
Private strChooseSQL As Variant     'INI�ɮ׸��|
Private strFilePath As Variant      '�ɮ׸��|
Private strRealDate As Variant      '�J�b���
Private strUpdEn As Variant         '�O�����ʤH��
Private strUpdName As Variant       '�O�����ʤH��
Private strPTCode As Variant        '�I�ں���
Private strPTName As Variant
Private strServiceType As Variant   '�A�����O
Private strBankSQL As Variant       '�Ȧ�OSQL
Private strCompCode As Variant      '���q�O
Private strClctEn As Variant        '���O�H��
Private strClctName As Variant
Private strCMCode As Variant        '���O�覡
Private strCMName As Variant
'Private strType As Variant         '�ʽ�O 1.�q�� 2.���G
Private blnBeOne As Boolean         '�X��
Private strBankHand As String
Private strChoose33 As String
Private strGetOwner As String       'OwnerName
Private strCitemQry As String       '���O����
Private intFlowId As Integer         '�y�{����
'Private strOwnerName As String
Private blnCkZero As Integer

'INI�ɮ׸��|
Public Property Let uChooseSQL(ByVal vData As Variant)
    strChooseSQL = vData
End Property
Public Property Get uChooseSQL() As Variant
    uChooseSQL = strChooseSQL
End Property

'INI�ɮ׸��|
Public Property Let uChoose33(ByVal vData As Variant)
    strChoose33 = vData
End Property
Public Property Get uChoose33() As Variant
    uChoose33 = strChoose33
End Property


'�ɮ׸��|
Public Property Let uFilePath(ByVal vData As Variant)
    strFilePath = vData
End Property
Public Property Get uFilePath() As Variant
    uFilePath = strFilePath
End Property


'INI�ɮ׸��|
Public Property Let uErrPath(ByVal vData As Variant)
    strErrPath = vData
End Property
Public Property Get uErrPath() As Variant
    uErrPath = strErrPath
End Property



'�J�b���
Public Property Let uRealDate(ByVal vData As Variant)
    strRealDate = vData
End Property
Public Property Get uRealDate() As Variant
    uRealDate = strRealDate
End Property


'���ʤH���N�X
Public Property Let uUpdEn(ByVal vData As Variant)
    strUpdEn = vData
End Property
Public Property Get uUpdEn() As Variant
    uUpdEn = strUpdEn
End Property


'���ʤH���W��
Public Property Let uUpdName(ByVal vData As Variant)
    strUpdName = vData
End Property
Public Property Get uUpdName() As Variant
    uUpdName = strUpdName
End Property


'�W�h��Connection
Public Property Set ugcnGi(ByVal vData As ADODB.Connection)
    Set gcnGi = vData
End Property

'�W�h��Connection
Public Property Set uCnn(ByVal vData As ADODB.Connection)
    Set cnn = vData
End Property


'�Ȧ�N�XSQL
Public Property Let uBankSQL(ByVal vData As Variant)
    strBankSQL = vData
End Property
Public Property Get uBankSQL() As Variant
    uBankSQL = strBankSQL
End Property


'�I�ں���
Public Property Let uPTCode(ByVal vData As Variant)
    strPTCode = vData
End Property
Public Property Get uPTCode() As Variant
    uPTCode = strPTCode
End Property
Public Property Let uPTName(ByVal vData As Variant)
    strPTName = vData
End Property
Public Property Get uPTName() As Variant
    uPTName = strPTName
End Property

'�I�O�覡
Public Property Let uCMCode(ByVal vData As Variant)
    strCMCode = vData
End Property
Public Property Get uCMCode() As Variant
    uCMCode = strCMCode
End Property
Public Property Let uCMName(ByVal vData As Variant)
    strCMName = vData
End Property
Public Property Get uCMName() As Variant
    uCMName = strCMName
End Property



'�A�Ⱥ���
Public Property Let uServiceType(ByVal vData As Variant)
    strServiceType = vData
End Property
Public Property Get uServiceType() As Variant
    uServiceType = strServiceType
End Property


'���q�O
Public Property Let uCompCode(ByVal vData As Variant)
    strCompCode = vData
End Property
Public Property Get uCompCode() As Variant
    uCompCode = strCompCode
End Property


'���O�H��
Public Property Let uClctEn(ByVal vData As Variant)
    strClctEn = vData
End Property
Public Property Get uClctEn() As Variant
    uClctEn = strClctEn
End Property
Public Property Let uClctName(ByVal vData As Variant)
    strClctName = vData
End Property
Public Property Get uClctName() As Variant
    uClctName = strClctName
End Property

'�X���ɮ�
Public Property Let uBeOne(ByVal vData As Boolean)
    blnBeOne = vData
End Property
Public Property Get uBeOne() As Boolean
    uBeOne = blnBeOne
End Property

'�D�n�Ȧ�
Public Property Let uBankHand(ByVal vData As String)
    strBankHand = vData
End Property
Public Property Get uBankHand() As String
    uBankHand = strBankHand
End Property
'OwnerName
Public Property Let uGetOwner(ByVal vData As String)
    strGetOwner = vData
    strOwnerName = vData
End Property
Public Property Get uGetOwner() As String
    uGetOwner = strGetOwner
End Property
'���O����
Public Property Let uCitemQry(ByVal vData As String)
    strCitemQry = vData
End Property

Public Property Get uCitemQry() As String
    uCitemQry = strCitemQry
End Property

'�y�{����
Public Property Let FlowId(ByVal vData As Integer)
    intFlowId = vData
End Property

Public Property Get FlowId() As Integer
    FlowId = intFlowId
End Property


'
'
''�ʽ�O 1.�q�� 2.���G
'Public Property Let uType(ByVal vData As Variant)
'    strType = vData
'End Property
'Public Property Get uType() As Variant
'    uType = strType
'End Property
'
'�t�����B�O�_����
Public Property Let uCkZero(ByVal vData As Boolean)
    blnCkZero = vData
End Property

Public Property Get uCkZero() As Boolean
    uCkZero = blnCkZero
End Property
