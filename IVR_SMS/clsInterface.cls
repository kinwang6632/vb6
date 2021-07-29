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

Private intType As Integer
Private strUrl As String
'傳送的網址
Public Property Let uUrl(ByVal vData As String)
  On Error Resume Next
    clsStoreParameter.uUrl = vData
    Exit Property
End Property
'傳送的型態
Public Property Let uSMSType(ByVal vData As Integer)
  On Error Resume Next
  clsStoreParameter.uSMSType = vData
End Property
'傳送的內容
Public Property Let uSMSMsg(ByVal vData As String)
  On Error Resume Next
  clsStoreParameter.uSMSMsg = vData
End Property
'傳送的電話
Public Property Let uSMSPhone(ByVal vData As String)
  On Error Resume Next
  clsStoreParameter.uSMSPhone = vData
End Property
'UserName
Public Property Let uUserName(ByVal vData As String)
  On Error Resume Next
  clsStoreParameter.uUserName = vData
End Property
'Password
Public Property Let uPassword(ByVal vData As String)
  On Error Resume Next
  clsStoreParameter.uPassword = vData
End Property
Public Property Let uID(ByVal vData As String)
  On Error Resume Next
  clsStoreParameter.uID = vData
End Property
Private Sub Class_Initialize()
    On Error Resume Next
    Set clsStoreParameter = New clsStoreParameter
End Sub
Public Property Get uRetOK() As Boolean
  On Error Resume Next
    uRetOK = blnRetOK
End Property
Public Property Get uRetMsg() As String
  On Error Resume Next
    uRetMsg = strRetMsg
End Property

Public Sub RunSMS()
    Select Case intSMSType
        Case 2
            
        Case 3
            Set objSMS = New clsSMSType3
            objSMS.SendSMS
        Case Else
    End Select

End Sub
