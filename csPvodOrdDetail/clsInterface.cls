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

Public Property Set uCnCom(ByRef vData As ADODB.Connection)
    Set clsStoreParameter.uCnCom = vData
End Property
Public Property Let uOrderEn(ByVal vData As String)
    clsStoreParameter.uOrderEn = vData
End Property
Public Property Let ufOrderWayCode(ByVal vData As String)
    clsStoreParameter.ufOrderWayCode = vData
End Property
Public Property Let uBillNo(ByVal vData As String)
  clsStoreParameter.uBillNo = vData
End Property
Public Property Let uBillItem(ByVal vData As String)
  clsStoreParameter.uBillItem = vData
End Property
Public Property Let uBillNoAndItem(ByVal vData As String)
    clsStoreParameter.uBillNoAndItem = vData
End Property
Public Property Let uProductType(ByVal vData As Integer)
    clsStoreParameter.uProductType = vData
End Property
Public Property Let uOrderDate1(ByVal vData As String)
    clsStoreParameter.uOrderDate1 = vData
End Property
Public Property Let uOrderDate2(ByVal vData As String)
    clsStoreParameter.uOrderDate2 = vData
End Property
'�O�_�i���
Public Property Let uCanSelect(ByVal vData As Boolean)
    clsStoreParameter.uCanSelect = vData
End Property
'INI�ɮ׸��|
Public Property Let uErrPath(ByVal vData As Variant)
    clsStoreParameter.uErrPath = vData
End Property


'�ɮ׸��|
Public Property Let uFilePath(ByVal vData As Variant)
    clsStoreParameter.uFilePath = vData
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
Public Property Set ugcnGi(ByRef vData As ADODB.Connection)
    Set clsStoreParameter.ugcnGi = vData
End Property


'�W�h��Connection
Public Property Set uCnn(ByRef vData As ADODB.Connection)
    Set clsStoreParameter.uCnn = vData
End Property

'�A�Ⱥ���
Public Property Let uServiceType(ByVal vData As Variant)
    clsStoreParameter.uServiceType = vData
End Property


'���q�O
Public Property Let uCompCode(ByVal vData As Variant)
    clsStoreParameter.uCompCode = vData
End Property

'OwnerName
Public Property Let uGetOwner(ByVal vData As String)
    clsStoreParameter.uGetOwner = vData
End Property
Public Sub Action()
  On Error GoTo ChkErr
    clsStoreParameter.Action

    Exit Sub
ChkErr:
  Err.Description = "�����l�ƥ��ѡI"
End Sub
Private Sub Class_Initialize()
    On Error GoTo ChkErr
        Set clsStoreParameter = New clsStoreParameter
    Exit Sub
ChkErr:
    ErrSub "clsInterface", "Class_Initialize"
End Sub
Public Sub RunBrowData()
  On Error GoTo ChkErr
    
    Dim rsTmp As ADODB.Recordset
    Call InitiComCn
    
    
    Set rsTmp = GetQryData
    Call ShowDataView
    
'    If rsTmp.RecordCount > 0 Then
'        Call ShowDataView
'    Else
'        MsgBox "�L����q�ʸ�Ʃ���", vbInformation, "�T��"
'    End If
  On Error Resume Next
  Call CloseRecordset(rsTmp)
  
  Exit Sub
ChkErr:
  Err.Description = "�s����ƪ����l�ƥ���"
End Sub
Public Property Let uCitemCode(ByVal vData As String)
  On Error Resume Next
    clsStoreParameter.uCitemCode = vData
End Property
Public Property Let uFaciSeqNos(ByVal vData As String)
  On Error Resume Next
    clsStoreParameter.uFaciSeqNos = vData
End Property

Public Property Let uCustId(ByVal vData As String)
    clsStoreParameter.uCustId = vData
End Property

'STB�]�ƧǸ�
Public Property Let uSrNO(ByVal vData As String)
    clsStoreParameter.uSrNO = vData
End Property
Public Property Let uFaciSeqNo(ByVal vData As String)
  On Error Resume Next
    clsStoreParameter.uFaciSeqNo = vData
End Property
Public Property Let uFaciSNo(ByVal vData As String)
  On Error Resume Next
    clsStoreParameter.uFaciSNo = vData
End Property
'�^�ǰ��檺Msg
Public Property Get uMsg() As Variant
    uMsg = clsStoreParameter.uMsg
End Property
'�^�ǩҦ��s�WSO033�����
Public Property Get uRs033() As ADODB.Recordset
  On Error Resume Next
    'Set uRs033 = rs033Clone
    Set uRs033 = clsStoreParameter.uRs033
End Property
'�^�ǩҦ���sSO033VOD�����
Public Property Get uRs033PVod() As ADODB.Recordset
  On Error Resume Next
    Set uRs033PVod = clsStoreParameter.uRs033PVod
End Property

Public Property Set uRs033(ByRef rs As ADODB.Recordset)
    
    Set clsStoreParameter.uRs033 = rs
End Property
Public Property Set uRs033Vod(ByRef rs As ADODB.Recordset)
    Set clsStoreParameter.uRs033Vod = rs
End Property
Public Property Let ugaryGi(ByVal vData As String)
    clsStoreParameter.ugaryGi = vData
End Property


Private Sub Class_Terminate()
  On Error Resume Next
    fBillNo = Empty
    fBillNoItem = Empty
    fBillNoAndItem = Empty
    fCanSelect = False
    fCustId = Empty
    fOrderDate1 = Empty
    fOrderDate2 = Empty
    fOrderEn = Empty
    fOrderWayCode = Empty
    fSTBSeqNos = Empty
    fProductType = 0
    fSTBSeqNo = Empty
    fCitemCode = Empty
    'CloseRecordset rs033Clone
    'CloseRecordset rs033PVodClone
    CloseConnection cnCom
End Sub