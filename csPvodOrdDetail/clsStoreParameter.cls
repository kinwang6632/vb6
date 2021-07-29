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
Private strUpdEn As Variant         '�O�����ʤH��
Private strUpdName As Variant       '�O�����ʤH��
Private strServiceType As Variant   '�A�����O
Private strCompCode As Variant      '���q�O
Private strGetOwner As String       'OwnerName
Private blnTrans As Boolean

Private blnNoData As Boolean
Private blnDO As Boolean
Private blnNOUpdate As Boolean
Private rs033PVodClone As ADODB.Recordset
Private rs033Clone As ADODB.Recordset
Private strMsg As Variant
Private strFaciSNo As String
Private strSrNO As String
Private strCustId As String
Private strFaciSeqNo As String

Private Const frmName As String = "clsStoreParameter"
Public Sub Action()
  On Error GoTo ChkErr
'    If cnCom Is Nothing Then
'        Dim ComConString As String
'        Dim pstrDataSource As String
'        Dim pstrUserId As String
'        Dim pstrPassword As String
'        Dim SysPath As String
'
'        SysPath = garyGi(11)
'        pstrDataSource = ReadFROMINI(SysPath, strComSectionName, "1", True)
'        pstrUserId = ReadFROMINI(SysPath, strComSectionName, "2", True)
'        pstrPassword = ReadFROMINI(SysPath, strComSectionName, "3", True)
'        ComConString = "Provider=MSDAORA.1;" & _
'                    "Password=" & pstrPassword & ";" & _
'                    "User ID=" & pstrUserId & ";" & _
'                    "Data Source=" & pstrDataSource & ";" & _
'                    "Persist Security Info=True"
'
'        Set cnCom = New ADODB.Connection
'        cnCom.ConnectionString = ComConString
'        cnCom.CursorLocation = adUseClient
'        cnCom.Open
'
'    End If
    Call InitiComCn
    With frmSO11A1A
        .Show vbModal
    End With
    
    Exit Sub
ChkErr:
  Call ErrSub(frmName, "Action")
End Sub
Public Property Let uBillNo(ByVal vData As String)
  fBillNo = vData
End Property
Public Property Let uBillItem(ByVal vData As String)
  fBillNoItem = vData
End Property
Public Property Let uProductType(ByVal vData As Integer)
    fProductType = vData
End Property
Public Property Let uBillNoAndItem(ByVal vData As String)
  fBillNoAndItem = vData
End Property
Public Property Set uCnCom(ByRef vData As ADODB.Connection)
    Set cnCom = vData
End Property
Public Property Let uOrderEn(ByVal vData As String)
    fOrderEn = vData
    
End Property
Public Property Let ufOrderWayCode(ByVal vData As String)
    fOrderWayCode = vData
End Property

Public Property Let uOrderDate1(ByVal vData As String)
    fOrderDate1 = Mid(Trim(Replace(vData, "/", "")), 1, 8)
End Property
Public Property Let uOrderDate2(ByVal vData As String)
    fOrderDate2 = Mid(Trim(Replace(vData, "/", "")), 1, 8)
End Property

Public Property Let uCanSelect(ByVal vData As Boolean)
    fCanSelect = vData
End Property

'�ɮ׸��|
Public Property Let uFilePath(ByVal vData As Variant)
    strFilePath = vData
End Property
Public Property Get uFilePath() As Variant
    uFilePath = strFilePath
End Property


'INI�ɮ׸��|(���~�ɮ׸��|)
Public Property Let uErrPath(ByVal vData As Variant)
    strErrPath = vData
    gErrLogPath = strErrPath
End Property
Public Property Get uErrPath() As Variant
    uErrPath = strErrPath
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
Public Property Set ugcnGi(ByRef vData As ADODB.Connection)
    Set gcnGi = vData
End Property

'�W�h��Connection
Public Property Set uCnn(ByRef vData As ADODB.Connection)
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

Public Property Let uCustId(ByVal vData As String)
    strCustId = vData
    gCustId = strCustId
End Property

'STB�]�ƧǸ�
Public Property Let uSrNO(ByVal vData As String)
    strSrNO = vData
End Property
Public Property Let uFaciSeqNo(ByVal vData As String)
  On Error Resume Next
    fSTBSeqNo = vData
End Property

Public Property Let uFaciSNo(ByVal vData As String)
  On Error Resume Next
    strFaciSNo = vData
End Property
Public Property Let uCitemCode(ByVal vData As String)
  On Error Resume Next
    fCitemCode = vData
End Property
Public Property Let uFaciSeqNos(ByVal vData As String)
  On Error Resume Next
    fSTBSeqNos = vData
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
Public Property Get uRs033PVod() As ADODB.Recordset
  On Error Resume Next
    Set uRs033PVod = rs033PVodClone
End Property

Public Property Set uRs033(ByRef rs As ADODB.Recordset)
    Set rs033Clone = rs
End Property
Public Property Set uRs033PVod(ByRef rs As ADODB.Recordset)
    Set rs033PVodClone = rs
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


Public Property Let uTrans(ByVal vData As Boolean)
  On Error Resume Next
    blnTrans = vData
End Property
Public Property Let ugaryGi(ByVal vData As String)
    Call GetGlobal(CStr(vData))
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
    fSTBSeqNo = Empty
    fCitemCode = Empty
    fProductType = 0
    CloseRecordset rs033Clone
    CloseRecordset rs033PVodClone
    CloseConnection cnCom
End Sub