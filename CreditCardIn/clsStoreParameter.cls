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
Private strUPDName As String        '���ʤH��
Private strPrgName As String        '��b�{���W��
Private strPTCode As String         '�I�ں���
Private strPTName As String         '�I�ں����W��
Private strServiceType As String    '�A�����O
Private strBankName As String       '�Ȧ�W��
Private blnUpdate As Boolean        '��s�{���O�_���\
Private strCompCode As String       '���q�O
Private blnCancel As Boolean        '�O�_���U������
Private strBankId As String         '�Ȧ�s��
Private strCitemQry As String       '���O����
Private strClctEn As String         '���O�H��
Private strClctName As String
Private strCMCode As Variant        '���O�覡
Private strCMName As Variant
Private lngExportFormat As Long
Private blnFuBonIntegrate  As Boolean
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
'�O�����ʤH��
Public Property Let UpdName(ByVal vData As String)
    strUPDName = vData
End Property

Public Property Get UpdName() As String
     UpdName = strUPDName
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
'�Ȧ�s��
Public Property Get BankId() As String
    BankId = strBankId
End Property
'���O����
Public Property Let uCitemQry(ByVal vData As String)
    strCitemQry = vData
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
Public Property Get uCitemQry() As String
    uCitemQry = strCitemQry
End Property

Public Property Let BankId(ByVal vData As String)
     strBankId = vData
End Property
Public Property Let ExportFormat(ByVal vData As Long)
    lngExportFormat = vData
End Property
Public Property Get ExportFormat() As Long
    ExportFormat = lngExportFormat
End Property
Public Property Let uFubonIntegrate(ByVal vData As Boolean)
    blnFuBonIntegrate = vData
End Property
Public Property Get uFubonIntegrate() As Boolean
    uFubonIntegrate = blnFuBonIntegrate
End Property
