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

'自訂列舉值
Public Enum giEditModeEnu
    giEditModeView = 0      ' 0=顯示模式
    giEditModeEdit = 1      ' 1.編輯模式(EditMode)
    giEditModeInsert = 2    ' 2=新增模式
    giEditModeDelete = 3    ' 3=刪除模式
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
'是否可選擇
Public Property Let uCanSelect(ByVal vData As Boolean)
    clsStoreParameter.uCanSelect = vData
End Property
'INI檔案路徑
Public Property Let uErrPath(ByVal vData As Variant)
    clsStoreParameter.uErrPath = vData
End Property


'檔案路徑
Public Property Let uFilePath(ByVal vData As Variant)
    clsStoreParameter.uFilePath = vData
End Property



'異動人員代碼
Public Property Let uUpdEn(ByVal vData As Variant)
    clsStoreParameter.uUpdEn = vData
End Property


'異動人員名稱
Public Property Let uUpdName(ByVal vData As Variant)
    clsStoreParameter.uUpdName = vData
End Property


'上層的Connection
Public Property Set ugcnGi(ByRef vData As ADODB.Connection)
    Set clsStoreParameter.ugcnGi = vData
End Property


'上層的Connection
Public Property Set uCnn(ByRef vData As ADODB.Connection)
    Set clsStoreParameter.uCnn = vData
End Property

'服務種類
Public Property Let uServiceType(ByVal vData As Variant)
    clsStoreParameter.uServiceType = vData
End Property


'公司別
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
  Err.Description = "物件初始化失敗！"
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
'        MsgBox "無任何訂購資料明細", vbInformation, "訊息"
'    End If
  On Error Resume Next
  Call CloseRecordset(rsTmp)
  
  Exit Sub
ChkErr:
  Err.Description = "瀏覽資料物件初始化失敗"
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

'STB設備序號
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
'回傳執行的Msg
Public Property Get uMsg() As Variant
    uMsg = clsStoreParameter.uMsg
End Property
'回傳所有新增SO033的資料
Public Property Get uRs033() As ADODB.Recordset
  On Error Resume Next
    'Set uRs033 = rs033Clone
    Set uRs033 = clsStoreParameter.uRs033
End Property
'回傳所有更新SO033VOD的資料
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
