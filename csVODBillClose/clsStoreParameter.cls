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
Private strServiceType As Variant   '�A�����O
Private strCompCode As Variant      '���q�O
Private strGetOwner As String       'OwnerName
Private blnTrans As Boolean
Private blnAutoEXE As Boolean
Private blnforceAmount As Boolean
Private blnNoData As Boolean
Private blnDO As Boolean
Private blnNOUpdate As Boolean
Private rs033VodClone As ADODB.Recordset
Private rs033Clone As ADODB.Recordset
Private rs182Clone As ADODB.Recordset
Private strPara36In As String
Private strPara35In As String
Private strMsg As Variant
Private strEndDate As String
Private strFaciSNO As String
Private strSrNO As String
Private strCustId As String
Private blnGive As Boolean
Private strFaciSeqNo As String
Public Sub Action()
    
    With frmSO1193A
        .uAutoExe = blnAutoEXE
        .uCompCode = gCompCode
        .uCustId = strCustId
        .uEndDate = strEndDate
        .uFaciSeqNo = strFaciSeqNo
        .uFaciSNo = strFaciSNO
        .uPara35 = Val(strPara35In & "")
        .uPara36 = Val(strPara36In & "")
        .uFromBat = blnGive
        .uNoUpdate = blnNOUpdate
        .uforceAmount = blnforceAmount
        .uTrans = blnTrans
        Set .uRs033 = rs033Clone
        Set .uRs033Vod = rs033VodClone
        Set .uRs182 = rs182Clone
        .Show 1
        blnDO = .uDo
        blnNoData = .uHaveData
        strMsg = .uMsg
    End With
End Sub


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




'�A�Ⱥ���
Public Property Let uServiceType(ByVal vData As Variant)
    strServiceType = vData
    strServiceType = "D"
End Property
Public Property Get uServiceType() As Variant
    uServiceType = strServiceType
End Property


'���q�O
Public Property Let uCompCode(ByVal vData As Variant)
    strCompCode = vData
    gCompCode = strCompCode
End Property
Public Property Get uCompCode() As Variant
    uCompCode = strCompCode
End Property
'OwnerName
Public Property Let uGetOwner(ByVal vData As String)
    strGetOwner = vData
    strOwnerName = vData
End Property
Public Property Get uGetOwner() As String
    uGetOwner = strGetOwner
End Property
Public Property Let uFormBat(ByVal vData As Boolean)
    blnGive = vData
    
End Property



Public Property Let uCustId(ByVal vData As String)
    strCustId = vData
End Property

'STB�]�ƧǸ�
Public Property Let uSrNO(ByVal vData As String)
    strSrNO = vData
'    txtSrNO.Text = vData
    
End Property
Public Property Let uFaciSeqNo(ByVal vData As String)
  On Error Resume Next
    strFaciSeqNo = vData
    
End Property
Public Property Let uFaciSNo(ByVal vData As String)
  On Error Resume Next
    strFaciSNO = vData
    
End Property

'����������
Public Property Let uEndDate(ByVal vData As String)
    strEndDate = vData
End Property

'PPV����@�~���B���B
Public Property Let uPara35(ByVal vData As Integer)
    strPara35In = vData
End Property

'PPV����@�~��ƭ��B
Public Property Let uPara36(ByVal vData As Integer)
    strPara36In = vData
End Property
'�^�ǰ��檺Msg
Public Property Get uMsg() As Variant
    uMsg = strMsg
End Property
'�^�ǩҦ��s�WSO033�����
Public Property Get uRs033() As ADODB.Recordset
  On Error Resume Next
    Set uRs033 = rs033Clone
End Property
'�^�ǩҦ���sSO033VOD�����
Public Property Get uRs033Vod() As ADODB.Recordset
  On Error Resume Next
    Set uRs033Vod = rs033VodClone
End Property
'�^�ǩҦ���sSO182�����
Public Property Get uRs182() As ADODB.Recordset
  On Error Resume Next
    Set uRs182 = rs182Clone
End Property
Public Property Set uRs182(ByRef rs As ADODB.Recordset)
    Set rs182Clone = rs
End Property
Public Property Set uRs033(ByRef rs As ADODB.Recordset)
    Set rs033Clone = rs
End Property
Public Property Set uRs033Vod(ByRef rs As ADODB.Recordset)
    Set rs033VodClone = rs
End Property

'��Miggie���M�L�Ӫ�,�����\������Update,��̫�nRollbackTrans
Public Property Let uNoUpdate(ByVal vData As Boolean)
  On Error Resume Next
    blnNOUpdate = vData
End Property
'�O�_���`���槹�{��
Public Property Get uDo() As Boolean
    uDo = blnDO
End Property
'�O�_�����檺���
Public Property Get uHaveData() As Boolean
  On Error Resume Next
    uHaveData = blnNoData
End Property
Public Property Let uAutoExe(ByVal vData As Boolean)
  On Error Resume Next
    blnAutoEXE = vData
End Property

Public Property Let uTrans(ByVal vData As Boolean)
  On Error Resume Next
    blnTrans = vData
End Property
Public Property Let ugaryGi(ByVal vData As String)
    Call GetGlobal(CStr(vData))
End Property
Public Property Let uforceAmount(ByVal vData As Boolean)
    blnforceAmount = vData
End Property
