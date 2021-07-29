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
Public Property Set ugcnGi(ByVal vData As ADODB.Connection)
    Set clsStoreParameter.ugcnGi = vData
End Property


'�W�h��Connection
Public Property Set uCnn(ByVal vData As ADODB.Connection)
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
Public Property Let uFromBat(ByVal vData As Boolean)
    'blnGive = vData
    clsStoreParameter.uFormBat = vData
End Property

Public Property Let uCustId(ByVal vData As String)
    'strCustId = vData
    'txtCustid.Text = vData
    'blnGive = True
    clsStoreParameter.uCustId = vData
End Property

'STB�]�ƧǸ�
Public Property Let uSrNO(ByVal vData As String)
    'strSrNO = vData
    'txtSrNO.Text = vData
    'txtSrNO.Text = vData
    clsStoreParameter.uSrNO = vData
End Property
Public Property Let uFaciSeqNo(ByVal vData As String)
  On Error Resume Next
    'strFaciSeqNo = vData
    'txtSrNO.Tag = strFaciSeqNo
    clsStoreParameter.uFaciSeqNo = vData
End Property
Public Property Let uFaciSNo(ByVal vData As String)
  On Error Resume Next
    'strFaciSNO = vData
    'txtSrNO.Text = strFaciSNO
    clsStoreParameter.uFaciSNo = vData
End Property

'����������
Public Property Let uEndDate(ByVal vData As String)
    'strEndDate = vData
    clsStoreParameter.uEndDate = vData
End Property

'PPV����@�~���B���B
Public Property Let uPara35(ByVal vData As Integer)
    'strPara35In = vData
    clsStoreParameter.uPara35 = vData
End Property

'PPV����@�~��ƭ��B
Public Property Let uPara36(ByVal vData As Integer)
    'strPara36In = vData
    clsStoreParameter.uPara36 = vData
End Property
'�^�ǰ��檺Msg
Public Property Get uMsg() As Variant
    'uMsg = strMsg
    uMsg = clsStoreParameter.uMsg
End Property
'�^�ǩҦ��s�WSO033�����
Public Property Get uRs033() As ADODB.Recordset
  On Error Resume Next
    'Set uRs033 = rs033Clone
    Set uRs033 = clsStoreParameter.uRs033
End Property
'�^�ǩҦ���sSO033VOD�����
Public Property Get uRs033Vod() As ADODB.Recordset
  On Error Resume Next
    'Set uRs033Vod = rs033VodClone
    Set uRs033Vod = clsStoreParameter.uRs033Vod
End Property
'�^�ǩҦ���sSO182�����
Public Property Get uRs182() As ADODB.Recordset
  On Error Resume Next
    'Set uRs182 = rs182Clone
    Set uRs182 = clsStoreParameter.uRs182
End Property
Public Property Set uRs182(ByRef rs As ADODB.Recordset)
    'Set rs182Clone = rs
    Set clsStoreParameter.uRs182 = rs
End Property
Public Property Set uRs033(ByRef rs As ADODB.Recordset)
    'Set rs033Clone = rs
    Set clsStoreParameter.uRs033 = rs
End Property
Public Property Set uRs033Vod(ByRef rs As ADODB.Recordset)
    'Set rs033VodClone = rs
    Set clsStoreParameter.uRs033Vod = rs
End Property

'��Miggie���M�L�Ӫ�,�����\������Update,��̫�nRollbackTrans
Public Property Let uNoUpdate(ByVal vData As Boolean)
  On Error Resume Next
    'blnNOUpdate = vData
    clsStoreParameter.uNoUpdate = vData
End Property
'�O�_���`���槹�{��
Public Property Get uDo() As Boolean
    'uDo = blnDO
    uDo = clsStoreParameter.uDo
End Property
'�O�_�����檺���
Public Property Get uHaveData() As Boolean
  On Error Resume Next
    'uHaveData = blnNoData
    uHaveData = clsStoreParameter.uHaveData
End Property
Public Property Let uAutoExe(ByVal vData As Boolean)
  On Error Resume Next
    clsStoreParameter.uAutoExe = vData
    'blnAutoEXE = vData
End Property
Public Property Let uforceAmount(ByVal vData As Boolean)
    clsStoreParameter.uforceAmount = vData
End Property
Public Property Let uTrans(ByVal vData As Boolean)
  On Error Resume Next
    clsStoreParameter.uTrans = vData
End Property
Public Property Let ugaryGi(ByVal vData As Variant)
  On Error Resume Next
    clsStoreParameter.ugaryGi = vData
End Property