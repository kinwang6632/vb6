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
Private strFilePath As String       'INI�ɮ׸��|
Private strRealDate As String       '�J�b���
Private strSourcePath As String     '��l�ɮ׸��|
Private strUpdEn As String          '�O�����ʤH��
Private strPrgName As String        '��b�{���W��
Private strPTCode As String         '�I�ں���
Private strPTName As String         '�I�ں����W��
Private strServiceType As String    '�A�����O
Private strBankName As String       '�Ȧ�W��
Private strBankId As String
Private blnUpdate As Boolean        '��s�{���O�_���\
Private strCompCode As String       '���q�O
Private blnCancel As Boolean        '�O�_���U������
Private strClctEn As String         '���O�H��
Private strClctName As String
Private strCMCode As String         '���O�覡
Private strCMName As String
Private strGetOwner As String       'OwnerName
Private strCitemQry As String      '���O����
Private strFlowId As String         '�y�{����
Private intRefNo As Integer         '#3527 �P�_CD018���ѦҸ� By Kin 2007/10/04
Private strUPDTime As String        '#4388 �W�[�O�����ʮɶ� By Kin 2009/05/04
Friend Property Let uRefNo(ByVal vData As Integer)
    intRefNo = vData
End Property
Friend Property Get uRefNo() As Integer
    uRefNo = intRefNo
End Property
'INI�ɮ׸��|
Public Property Let FilePath(ByVal vData As String)
    strFilePath = vData
End Property

Public Property Get FilePath() As String
     FilePath = strFilePath
End Property
'�J�b���
Public Property Let RealDate(ByVal vData As String)
    strRealDate = vData
End Property

Public Property Get RealDate() As String
     RealDate = strRealDate
End Property
'��l�ɮ׸��|
Public Property Let SourcePath(ByVal vData As String)
    strSourcePath = vData
End Property

Public Property Get SourcePath() As String
     SourcePath = strSourcePath
End Property
'�O�����ʤH��
Public Property Let UpdEn(ByVal vData As String)
    strUpdEn = vData
End Property

Public Property Get UpdEn() As String
     UpdEn = strUpdEn
End Property
'��b�{���W��
Public Property Let PrgName(ByVal vData As String)
    strPrgName = vData
End Property

Public Property Get PrgName() As String
     PrgName = strPrgName
End Property
'�W�h��Connection
Public Property Set ugcnGi(ByVal vData As ADODB.Connection)
    Set gcnGi = vData
End Property
'�I�ں���
Public Property Let uPTCode(ByVal vData As String)
    strPTCode = vData
End Property

Public Property Get uPTCode() As String
     uPTCode = strPTCode
End Property
'�I�ں����W��
Public Property Let uPTName(ByVal vData As String)
    strPTName = vData
End Property

Public Property Get uPTName() As String
     uPTName = strPTName
End Property
'�A�Ⱥ���
Public Property Let ServiceType(ByVal vData As String)
    strServiceType = vData
End Property

Public Property Get ServiceType() As String
     ServiceType = strServiceType
End Property
'�Ȧ�W��
Public Property Let BankName(ByVal vData As String)
     strBankName = vData
End Property
'�Ȧ�N�X
Public Property Let BankId(ByVal vData As String)
    strBankId = vData
End Property
Public Property Get BankId() As String
    BankId = strBankId
End Property
Public Property Get BankName() As String
     BankName = strBankName
End Property
'�O���O�_��b���\
Public Property Get UpdState() As Boolean
     UpdState = blnUpdate
End Property

Public Property Let UpdState(ByVal vData As Boolean)
     blnUpdate = vData
End Property
'���q�O
Public Property Get CompCode() As String
    CompCode = strCompCode
End Property

Public Property Let CompCode(ByVal vData As String)
    strCompCode = vData
End Property

'���O�H��
Public Property Get uClctEn() As String
    uClctEn = strClctEn
End Property

Public Property Let uClctEn(ByVal vData As String)
    strClctEn = vData
End Property
'���O�H���W��
Public Property Get uClctName() As String
    uClctName = strClctName
End Property

Public Property Let uClctName(ByVal vData As String)
    strClctName = vData
End Property

'���O�覡
Public Property Get uCMCode() As String
    uCMCode = strCMCode
End Property

Public Property Let uCMCode(ByVal vData As String)
    strCMCode = vData
End Property
'���O�覡�W��
Public Property Get uCMName() As String
    uCMName = strCMName
End Property

Public Property Let uCMName(ByVal vData As String)
    strCMName = vData
End Property
'���oOwnerName
Public Property Let uGetOwner(ByVal vData As String)
    strGetOwner = vData
    strOwnerName = vData
End Property

Friend Property Get uGetOwner() As String
    uGetOwner = strGetOwner
End Property

'���O����
Public Property Get uCitemQry() As String
    uCitemQry = strCitemQry
End Property

Public Property Let uCitemQry(ByVal vData As String)
    strCitemQry = vData
End Property

'�y�{����
Public Property Get FlowId() As String
     FlowId = strFlowId
End Property

Public Property Let FlowId(ByVal vData As String)
     strFlowId = vData
End Property
Public Property Let uUPDTime(ByVal vData As String)
    strUPDTime = vData
End Property
Public Property Get uUPDTime() As String
    uUPDTime = strUPDTime
End Property
