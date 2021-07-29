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
Private cn As New ADODB.Connection
Private MDBcn As New ADODB.Connection
Private strProcText As String
Private strProcErrText As String
'�J�b�b��
Private strActNo As String
'�Ʒ~���N��
Private strBankId As String
'������N��
Private strBankName As String
'������W��
Private strCorpID As String
'Where ����
Private strChoose As String
'���O�a�}�P�_
Private strPrgName As String
'�{���W��
Private blnChargeAddr As Boolean
'ini�ɮ׸��|
Private strIniPath As String
'Errlog�ɮ׸��|
Private strErrPath As String
'���q�O
Private strCompCode As String
'�A�����O
Private strServiceType As String
'��s����
Private blnUpdate As Boolean
'����ɶ�
Private strTime As String
'�O�_�����
Private blnNodata As Boolean
'��������I���
Private strStopDate As String
'�C��h�b��B�z (0=�_ , 1=�O)
Private strChoose33 As String
'OwnerName
Private strGetOwner As String
'�y�{����
Private strFlowId As String
Private strFlow As String
'�h�b�Უ�ͨ̾ڳ]�w PostUnit �l���o����
Private strBillHeadFmt As String
'#3527 �s�W�l���榡�A�H�ѦҸ��P�O(�ϥ�Pos2) By Kin 2007/10/04
Private intRefNo As Integer
'#3968 �W�[���~�ѧO�X
Private strBID As String
Private strUpdEn As String
Private intPay As Integer   '#5683 �W�[�s�����O
'#3968 ��~�ѧO�X
Friend Property Get uBID() As String
    uBID = strBID
        
End Property
'#5683 �W�[�s�����O
Friend Property Let uPayType(ByVal uData As Integer)
    intPay = Val(uData)
End Property
Friend Property Get uPayType() As Integer
    uPayType = Val(intPay)
End Property
Friend Property Let uBID(ByVal uData As String)
  strBID = uData
End Property
Friend Property Get uRefNo() As Integer
    uRefNo = intRefNo
End Property
Friend Property Let uRefNo(ByVal vRefNo As Integer)
    intRefNo = vRefNo
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

Public Property Get ErrPath() As String
    ErrPath = strErrPath
End Property

Public Property Let ErrPath(ByVal vErrPath As String)
    strErrPath = vErrPath
End Property

Friend Property Get Connection() As ADODB.Connection
    Set Connection = cn
End Property

Friend Property Set Connection(ByVal vData As ADODB.Connection)
    Set cn = vData
End Property
'Open MDB ��Connection
Friend Property Get MDBConn() As ADODB.Connection
    Set MDBConn = MDBcn
End Property

Friend Property Set MDBConn(ByVal vData As ADODB.Connection)
    Set MDBcn = vData
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
    strCorpID = vData
End Property

Friend Property Get CorpID() As String
    CorpID = strCorpID
End Property
'���q�O
Friend Property Get uCompCode() As String
    uCompCode = strCompCode
End Property

Friend Property Let uCompCode(ByVal vData As String)
    strCompCode = vData
End Property
'�A�����O
Friend Property Get uServiceType() As String
    uServiceType = strServiceType
End Property

Friend Property Let uServiceType(ByVal vData As String)
    strServiceType = vData
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
     uTime = strTime
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
'SO033������
Public Property Get uChoose33() As String
   uChoose33 = strChoose33
End Property

Public Property Let uChoose33(ByVal vData As String)
   strChoose33 = vData
End Property

'���oOwnerName
Public Property Let uGetOwner(ByVal vData As String)
    strGetOwner = vData
    strOwnerName = vData
End Property

Friend Property Get uGetOwner() As String
    uGetOwner = strGetOwner
End Property
'�y�{����
Public Property Let FlowId(ByVal vData As String)
    strFlowId = vData
    strFlow = vData
End Property

Friend Property Get FlowId() As String
    FlowId = strFlowId
End Property

'�h�b�Უ�ͨ̾ڳ]�w PostUnit �l���o����
Public Property Let uBillHeadFmt(ByVal vData As String)
    strBillHeadFmt = vData
End Property

Friend Property Get uBillHeadFmt() As String
    uBillHeadFmt = strBillHeadFmt
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
'#4388 �W�[���ʤH�� By Kin 2009/04/29
Public Property Let uUpdEn(ByVal vData As String)
    strUpdEn = vData
End Property
Public Property Get uUpdEn() As String
    uUpdEn = strUpdEn
End Property