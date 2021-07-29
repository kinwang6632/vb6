VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStoreParameter"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private cn As New ADODB.Connection
'入帳帳號
Private strActNo As String
'事業單位代號
Private strBankId As String
'收件單位代號
Private strBankName As String
'收件單位名稱
Private strCorpId As String
'Where 條件
Private strChoose As String
'收費地址與否
Private strPrgName As String
'程式名稱
Private blnChargeAddr As Boolean
'檔案路徑
Private strIniPath As String
Private strErrPath As String
'更新失敗
Private blnUpdate As Boolean
'執行時間
Private strTime As String
'是否有資料
Private blnNodata As Boolean
'應收日期截止日
Private strStopDate As String
'地址依據
Private strAddrType As String
'地址依據2
Private strAddrType2 As String
'信用卡扣帳特店代碼
Private strSpcNo As String
Private strProcText As String
Private strProcErrText As String
Private intChkType As Integer
Private strUPDTime As String
Private strUPDEN As String
Private blnOutZero As Boolean
Friend Property Let uOutZero(ByVal uData As Boolean)
    blnOutZero = uData
End Property

Friend Property Get uOutZero() As Boolean
    uOutZero = blnOutZero
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

Public Property Get ErrPath() As Variant
    ErrPath = strErrPath
End Property

Public Property Let ErrPath(ByVal vErrPath As Variant)
    strErrPath = vErrPath
End Property

Public Property Get Connection() As ADODB.Connection
    Set Connection = cn
End Property

Public Property Set Connection(ByVal vData As ADODB.Connection)
    Set cn = vData
End Property

'入帳帳號
Friend Property Let ActNo(ByVal vData As String)
    strActNo = vData
End Property

Friend Property Get ActNo() As String
    ActNo = strActNo
End Property

'事業單位代號
Friend Property Let BankId(ByVal vData As String)
    strBankId = vData
End Property

Friend Property Get BankId() As String
    BankId = strBankId
End Property

'程式名稱
Friend Property Let PrgName(ByVal vData As String)
    strPrgName = vData
End Property

Friend Property Get PrgName() As String
    PrgName = strPrgName
End Property

'收件單位代號
Friend Property Let CorpID(ByVal vData As String)
    strCorpId = vData
End Property

Friend Property Get CorpID() As String
    CorpID = strCorpId
End Property

'執行失敗
Public Property Get uUpdate() As Boolean
     uUpdate = blnUpdate
End Property

Public Property Let uUpdate(ByVal vData As Boolean)
     blnUpdate = vData
End Property
'執行時間
Public Property Get uTime() As Long
    uTime = Timer
End Property

Public Property Let uTime(ByVal vData As Long)
     strTime = vData
End Property
'是否有資料
Public Property Get uNoData() As Boolean
     uNoData = blnNodata
End Property

Public Property Let uNoData(ByVal vData As Boolean)
     blnNodata = vData
End Property

'應收日期截止日
Public Property Get uStopDate() As String
   uStopDate = strStopDate
End Property

Public Property Let uStopDate(ByVal vData As String)
   strStopDate = vData
End Property

Public Property Let uSpcNo(ByVal vData As String)
   strSpcNo = vData
End Property
Public Property Get uSpcNo() As String
   uSpcNo = strSpcNo
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
Public Property Let uChkType(ByVal vData As Integer)
    intChkType = vData
End Property
Public Property Get uChkType() As Integer
    uChkType = intChkType
End Property
Public Property Let uUPDEN(ByVal vData As String)
    strUPDEN = vData
End Property
Public Property Get uUPDEN() As String
    uUPDEN = strUPDEN
End Property
Public Property Let uUPDTime(ByVal vData As String)
    strUPDTime = vData
End Property
Public Property Get uUPDTime() As String
    uUPDTime = strUPDTime
End Property

