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

'發送訊息的網址
Public Property Let uUrl(ByVal vData As String)
  On Error Resume Next
    strSMSUrl = vData
End Property
Public Property Get uUrl() As String
   On Error Resume Next
   uUrl = strSMSUrl
End Property
'發送訊息的Type
Public Property Let uSMSType(ByVal vData As Integer)
    intSMSType = vData
End Property
Public Property Get uSMSType() As Integer
    uSMSType = intSMSType
End Property
'傳送的內容
Public Property Let uSMSMsg(ByVal vData As String)
  On Error Resume Next
  strMsg = vData
End Property
Public Property Get uSMSMsg() As String
  On Error Resume Next
  uSMSMsg = strMsg
End Property
'傳送的電話
Public Property Let uSMSPhone(ByVal vData As String)
  On Error Resume Next
  strPhone = vData
End Property
Public Property Get uSMSPhone() As String
  On Error Resume Next
  uSMSPhone = strPhone
End Property
'UserName
Public Property Let uUserName(ByVal vData As String)
  On Error Resume Next
  strUser = vData
End Property
Public Property Get uUserName() As String
  On Error Resume Next
  uUserName = strUser
End Property
'Password
Public Property Let uPassword(ByVal vData As String)
  On Error Resume Next
  strPwd = vData
End Property
Public Property Get uPassword() As String
    On Error Resume Next
    uPassword = strPwd
End Property
'ID
Public Property Let uID(ByVal vData As String)
  On Error Resume Next
  strID = vData
End Property
Public Property Get uID() As String
  On Error Resume Next
  uID = strID
End Property
