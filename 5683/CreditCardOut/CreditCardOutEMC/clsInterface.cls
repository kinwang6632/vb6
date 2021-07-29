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
Public blnIsTCB As Boolean  '' 中華商銀格式出單為 true
Public blnIsTCBCard As Boolean  ''  中華商銀自行卡則為 true

Public Property Let ChargeAddr(ByVal vCharge As Boolean)
    objStorePara.ChargeAddr = vCharge
End Property

Public Property Let ChooseStr(ByVal vChoose As String)
    objStorePara.ChooseStr = vChoose
End Property

Public Property Let BankName(ByVal vBankName As String)
    objStorePara.BankName = vBankName
End Property

Public Property Let Inipath(ByVal vIniPath As String)
    objStorePara.Inipath = vIniPath
End Property

Public Property Let ErrPath(ByVal vErrPath As Variant)
    objStorePara.ErrPath = vErrPath
End Property

Public Property Set Connection(ByVal vData As ADODB.Connection)
    Set objStorePara.Connection = vData
End Property

'入帳帳號
Public Property Let ActNo(ByVal vData As String)
    objStorePara.ActNo = vData
End Property

'事業單位代號
Public Property Let BankId(ByVal vData As String)
    objStorePara.BankId = vData
End Property
'#3531 多增加銀行的條件(因為之前的人腦殘,一開始就將銀行條件寫死) By Kin 2007/11/02
Public Property Let BankStr(ByVal vData As String)
    objStorePara.BankStr = vData
End Property

'收件單位代號
Public Property Let CorpID(ByVal vData As String)
    objStorePara.CorpID = vData
End Property

'程式名稱
Public Property Let PrgName(ByVal vData As String)
    objStorePara.PrgName = vData
End Property

'使用者名稱
Public Property Let UserName(ByVal vData As String)
    strUserName = vData
End Property
'執行失敗
Public Property Get uUpdate() As Boolean
    uUpdate = objStorePara.uUpdate
End Property
'執行時間
Public Property Get uTime() As Long
    uTime = objStorePara.uTime
End Property
'是否有資料
Public Property Get uNoData() As Boolean
    uNoData = objStorePara.uNoData
End Property
'應收日期截止日
Public Property Let uStopDate(ByVal vData As String)
    objStorePara.uStopDate = vData
End Property

Public Property Let uSpcNO(ByVal vData As String)
    objStorePara.uSpcNO = vData
End Property

Public Function InitObject(ByVal clsName As String) As Object
    On Error GoTo ChkErr
        Dim objCls As Object
        Set objCls = CreateObject("CreditCardOutEmc." & clsName)
        Set InitObject = objCls
    Exit Function
ChkErr:
    Err.Description = "物件初始化失敗！請檢查程式名稱！"
End Function

Private Sub Class_Initialize()
    On Error GoTo ChkErr
        Set objStorePara = New clsStoreParameter
    Exit Sub
ChkErr:
    ErrSub "clsInterface", "Class_Initialize"
End Sub


Public Property Let pTableOwnerName(ByVal vNewValue As String)
    TableOwnerName = vNewValue
End Property
Public Property Get uProcText() As String
    uProcText = objStorePara.uProcText
End Property
Public Property Get uProcErrText() As String
    uProcErrText = objStorePara.uProcErrText
End Property
Public Property Let uUPDEN(ByVal vData As String)
  On Error GoTo ChkErr
    objStorePara.uUPDEN = vData
  Exit Property
ChkErr:
  Call ErrSub("clsInterface", "uUPDEN")
End Property
Public Property Let uUPDTime(ByVal vData As String)
  On Error Resume Next
    objStorePara.uUPDTime = vData
End Property
