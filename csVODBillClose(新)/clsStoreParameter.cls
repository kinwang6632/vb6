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
Private strChooseSQL As Variant     'INI檔案路徑
Private strFilePath As Variant      '檔案路徑
Private strRealDate As Variant      '入帳日期
Private strUpdEn As Variant         '記錄異動人員
Private strUpdName As Variant       '記錄異動人員
Private strServiceType As Variant   '服務類別
Private strCompCode As Variant      '公司別
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


'檔案路徑
Public Property Let uFilePath(ByVal vData As Variant)
    strFilePath = vData
End Property
Public Property Get uFilePath() As Variant
    uFilePath = strFilePath
End Property


'INI檔案路徑
Public Property Let uErrPath(ByVal vData As Variant)
    strErrPath = vData
End Property
Public Property Get uErrPath() As Variant
    uErrPath = strErrPath
End Property



'入帳日期
Public Property Let uRealDate(ByVal vData As Variant)
    strRealDate = vData
End Property
Public Property Get uRealDate() As Variant
    uRealDate = strRealDate
End Property


'異動人員代碼
Public Property Let uUpdEn(ByVal vData As Variant)
    strUpdEn = vData
End Property
Public Property Get uUpdEn() As Variant
    uUpdEn = strUpdEn
End Property


'異動人員名稱
Public Property Let uUpdName(ByVal vData As Variant)
    strUpdName = vData
End Property
Public Property Get uUpdName() As Variant
    uUpdName = strUpdName
End Property


'上層的Connection
Public Property Set ugcnGi(ByVal vData As ADODB.Connection)
    Set gcnGi = vData
End Property

'上層的Connection
Public Property Set uCnn(ByVal vData As ADODB.Connection)
    Set cnn = vData
End Property




'服務種類
Public Property Let uServiceType(ByVal vData As Variant)
    strServiceType = vData
    strServiceType = "D"
End Property
Public Property Get uServiceType() As Variant
    uServiceType = strServiceType
End Property


'公司別
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

'STB設備序號
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

'結算日期期限
Public Property Let uEndDate(ByVal vData As String)
    strEndDate = vData
End Property

'PPV結算作業金額限額
Public Property Let uPara35(ByVal vData As Integer)
    strPara35In = vData
End Property

'PPV結算作業月數限額
Public Property Let uPara36(ByVal vData As Integer)
    strPara36In = vData
End Property
'回傳執行的Msg
Public Property Get uMsg() As Variant
    uMsg = strMsg
End Property
'回傳所有新增SO033的資料
Public Property Get uRs033() As ADODB.Recordset
  On Error Resume Next
    Set uRs033 = rs033Clone
End Property
'回傳所有更新SO033VOD的資料
Public Property Get uRs033Vod() As ADODB.Recordset
  On Error Resume Next
    Set uRs033Vod = rs033VodClone
End Property
'回傳所有更新SO182的資料
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

'由Miggie結清過來的,不允許任何資料Update,到最後要RollbackTrans
Public Property Let uNoUpdate(ByVal vData As Boolean)
  On Error Resume Next
    blnNOUpdate = vData
End Property
'是否正常執行完程式
Public Property Get uDo() As Boolean
    uDo = blnDO
End Property
'是否有執行的資料
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

