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
Private intBankId As Integer
Private blnUpdate As Boolean        '��s�{���O�_���\
Private strCompCode As String       '���q�O
Private strGetOwner As String       'OwnerName
Private blnCancel As Boolean        '�O�_���U������

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
Public Property Let PTCode(ByVal vData As String)
    strPTCode = vData
End Property

Public Property Get PTCode() As String
     PTCode = strPTCode
End Property
'�I�ں����W��
Public Property Let PTName(ByVal vData As String)
    strPTName = vData
End Property

Public Property Get PTName() As String
     PTName = strPTName
End Property
'�A�Ⱥ���
Public Property Let ServiceType(ByVal vData As String)
    strServiceType = vData
End Property

Public Property Get ServiceType() As String
     ServiceType = strServiceType
End Property
'�Ȧ�W��
Public Property Let BankID(ByVal vData As Integer)
     intBankId = vData
End Property

Public Property Get BankID() As Integer
     BankID = intBankId
End Property
'�Ȧ�W��
Public Property Let BankName(ByVal vData As String)
     strBankName = vData
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
'���oOwnerName
Public Property Let uGetOwner(ByVal vData As String)
    strGetOwner = vData
    strOwnerName = vData
     garyGi(16) = vData
End Property

Friend Property Get uGetOwner() As String
    uGetOwner = strGetOwner
End Property

