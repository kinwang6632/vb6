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
Private strProcErrText As String
'入帳帳號

Private strChoose As String
'收費地址與否

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
Private strGetOwner As String
Private strUpdEn As String

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

'#4388 增加異動人員 By Kin 2009/04/29
Public Property Let uUpdEn(ByVal vData As String)
    strUpdEn = vData
End Property
Public Property Get uUpdEn() As String
    uUpdEn = strUpdEn
End Property
