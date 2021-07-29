Attribute VB_Name = "clsModule"
Option Explicit

Public clsStoreParameter As Object
Public strSMSUrl As String
Public intSMSType As Integer
Public objSMS As Object
Public objxmlHttp As Object
'Public Const strRequestHeader = "application/x-www-form-urlencoded"
Public strUser As String
Public strPwd As String
Public strMsg As String
Public strPhone As String
Public strID As String
Public blnRetOK As Boolean
Public strRetMsg As String
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Const RequestHeader As String = "application/x-www-form-urlencoded"
