VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStoreParameter"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private cn As New ADODB.Connection
'�J�b�b��
Private strActNo As String
'�Ʒ~���N��
Private strBankId As String
'������N��
Private strBankName As String
'������W��
Private strCorpId As String
'Where ����
Private strChoose As String
'���O�a�}�P�_
Private strPrgName As String
'�{���W��
Private blnChargeAddr As Boolean
'�ɮ׸��|
Private strIniPath As String
Private strErrPath As String
'��s����
Private blnUpdate As Boolean
'����ɶ�
Private strTime As String
'�O�_�����
Private blnNodata As Boolean
'��������I���
Private strStopDate As String
'�a�}�̾�
Private strAddrType As String
'�a�}�̾�2
Private strAddrType2 As String
'�H�Υd���b�S���N�X
Private strSpcNo As String
Private strProcText As String
Private strProcErrText As String
Private intChkType As Integer
Private strUPDTime As String
Private strUPDEN As String
Private blnOutZero As Boolean
Friend Property Let uOutZero(ByVal uData As Boolean)
    blnOutZero = uData
End Property

Friend Property Get uOutZero() As Boolean
    uOutZero = blnOutZero
End Property
Friend Property Get ChargeAddr() As Boolean
    ChargeAddr = blnChargeAddr
End Property

Friend Property Let ChargeAddr(ByVal vCharge As Boolean)
    blnChargeAddr = vCharge
End Property

Friend Property Get ChooseStr() As String
    ChooseStr = strChoose
End Property

Friend Property Let ChooseStr(ByVal vChoose As String)
    strChoose = vChoose
End Property

Friend Property Get BankName() As String
    BankName = strBankName
End Property

Friend Property Let BankName(ByVal vBankName As String)
    strBankName = vBankName
End Property

Friend Property Get Inipath() As String
    Inipath = strIniPath
End Property

Friend Property Let Inipath(ByVal vIniPath As String)
    strIniPath = vIniPath
End Property

Public Property Get ErrPath() As Variant
    ErrPath = strErrPath
End Property

Public Property Let ErrPath(ByVal vErrPath As Variant)
    strErrPath = vErrPath
End Property

Public Property Get Connection() As ADODB.Connection
    Set Connection = cn
End Property

Public Property Set Connection(ByVal vData As ADODB.Connection)
    Set cn = vData
End Property

'�J�b�b��
Friend Property Let ActNo(ByVal vData As String)
    strActNo = vData
End Property

Friend Property Get ActNo() As String
    ActNo = strActNo
End Property

'�Ʒ~���N��
Friend Property Let BankId(ByVal vData As String)
    strBankId = vData
End Property

Friend Property Get BankId() As String
    BankId = strBankId
End Property

'�{���W��
Friend Property Let PrgName(ByVal vData As String)
    strPrgName = vData
End Property

Friend Property Get PrgName() As String
    PrgName = strPrgName
End Property

'������N��
Friend Property Let CorpID(ByVal vData As String)
    strCorpId = vData
End Property

Friend Property Get CorpID() As String
    CorpID = strCorpId
End Property

'���楢��
Public Property Get uUpdate() As Boolean
     uUpdate = blnUpdate
End Property

Public Property Let uUpdate(ByVal vData As Boolean)
     blnUpdate = vData
End Property
'����ɶ�
Public Property Get uTime() As Long
    uTime = Timer
End Property

Public Property Let uTime(ByVal vData As Long)
     strTime = vData
End Property
'�O�_�����
Public Property Get uNoData() As Boolean
     uNoData = blnNodata
End Property

Public Property Let uNoData(ByVal vData As Boolean)
     blnNodata = vData
End Property

'��������I���
Public Property Get uStopDate() As String
   uStopDate = strStopDate
End Property

Public Property Let uStopDate(ByVal vData As String)
   strStopDate = vData
End Property

Public Property Let uSpcNo(ByVal vData As String)
   strSpcNo = vData
End Property
Public Property Get uSpcNo() As String
   uSpcNo = strSpcNo
End Property

Public Property Let uProcText(ByVal vData As String)
    strProcText = vData
End Property
Public Property Get uProcText() As String
    uProcText = strProcText
End Property
Public Property Let uProcErrText(ByVal vData As String)
    strProcErrText = vData
End Property
Public Property Get uProcErrText() As String
    uProcErrText = strProcErrText
End Property
Public Property Let uChkType(ByVal vData As Integer)
    intChkType = vData
End Property
Public Property Get uChkType() As Integer
    uChkType = intChkType
End Property
Public Property Let uUPDEN(ByVal vData As String)
    strUPDEN = vData
End Property
Public Property Get uUPDEN() As String
    uUPDEN = strUPDEN
End Property
Public Property Let uUPDTime(ByVal vData As String)
    strUPDTime = vData
End Property
Public Property Get uUPDTime() As String
    uUPDTime = strUPDTime
End Property
