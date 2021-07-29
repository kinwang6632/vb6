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
Private cn As New ADODB.Connection
Private MDBcn As New ADODB.Connection
Private strProcText As String
Private strProcErrText As String
'入帳帳號
Private strActNo As String
'事業單位代號
Private strBankId As String
'收件單位代號
Private strBankName As String
'收件單位名稱
Private strCorpID As String
'Where 條件
Private strChoose As String
'收費地址與否
Private strPrgName As String
'程式名稱
Private blnChargeAddr As Boolean
'ini檔案路徑
Private strIniPath As String
'Errlog檔案路徑
Private strErrPath As String
'公司別
Private strCompCode As String
'服務類別
Private strServiceType As String
'更新失敗
Private blnUpdate As Boolean
'執行時間
Private strTime As String
'是否有資料
Private blnNodata As Boolean
'應收日期截止日
Private strStopDate As String
'媒體多帳戶處理 (0=否 , 1=是)
Private strChoose33 As String
'OwnerName
Private strGetOwner As String
'流程控制
Private strFlowId As String
Private strFlow As String
'多帳戶產生依據設定 PostUnit 郵局發件單位
Private strBillHeadFmt As String
'#3527 新增郵局格式，以參考號判別(使用Pos2) By Kin 2007/10/04
Private intRefNo As Integer
'#3968 增加企業識別碼
Private strBID As String
Private strUpdEn As String
Private intPay As Integer   '#5683 增加編款類別
Private blnOutZero As Boolean
Friend Property Let uOutZero(ByVal uData As Boolean)
    blnOutZero = uData
End Property

Friend Property Get uOutZero() As Boolean
    uOutZero = blnOutZero
End Property

'#3968 止業識別碼
Friend Property Get uBID() As String
    uBID = strBID
        
End Property
'#5683 增加編款類別
Friend Property Let uPayType(ByVal uData As Integer)
    intPay = Val(uData)
End Property
Friend Property Get uPayType() As Integer
    uPayType = Val(intPay)
End Property
Friend Property Let uBID(ByVal uData As String)
  strBID = uData
End Property
Friend Property Get uRefNo() As Integer
    uRefNo = intRefNo
End Property
Friend Property Let uRefNo(ByVal vRefNo As Integer)
    intRefNo = vRefNo
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

Public Property Get ErrPath() As String
    ErrPath = strErrPath
End Property

Public Property Let ErrPath(ByVal vErrPath As String)
    strErrPath = vErrPath
End Property

Friend Property Get Connection() As ADODB.Connection
    Set Connection = cn
End Property

Friend Property Set Connection(ByVal vData As ADODB.Connection)
    Set cn = vData
End Property
'Open MDB 的Connection
Friend Property Get MDBConn() As ADODB.Connection
    Set MDBConn = MDBcn
End Property

Friend Property Set MDBConn(ByVal vData As ADODB.Connection)
    Set MDBcn = vData
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
    strCorpID = vData
End Property

Friend Property Get CorpID() As String
    CorpID = strCorpID
End Property
'公司別
Friend Property Get uCompCode() As String
    uCompCode = strCompCode
End Property

Friend Property Let uCompCode(ByVal vData As String)
    strCompCode = vData
End Property
'服務類別
Friend Property Get uServiceType() As String
    uServiceType = strServiceType
End Property

Friend Property Let uServiceType(ByVal vData As String)
    strServiceType = vData
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
     uTime = strTime
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
'SO033的條件
Public Property Get uChoose33() As String
   uChoose33 = strChoose33
End Property

Public Property Let uChoose33(ByVal vData As String)
   strChoose33 = vData
End Property

'取得OwnerName
Public Property Let uGetOwner(ByVal vData As String)
    strGetOwner = vData
    strOwnerName = vData
End Property

Friend Property Get uGetOwner() As String
    uGetOwner = strGetOwner
End Property
'流程控制
Public Property Let FlowId(ByVal vData As String)
    strFlowId = vData
    strFlow = vData
End Property

Friend Property Get FlowId() As String
    FlowId = strFlowId
End Property


'多帳戶產生依據設定 PostUnit 郵局發件單位
Public Property Let uBillHeadFmt(ByVal vData As String)
    strBillHeadFmt = vData
End Property

Friend Property Get uBillHeadFmt() As String
    uBillHeadFmt = strBillHeadFmt
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
'#4388 增加異動人員 By Kin 2009/04/29
Public Property Let uUpdEn(ByVal vData As String)
    strUpdEn = vData
End Property
Public Property Get uUpdEn() As String
    uUpdEn = strUpdEn
End Property
