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


Dim strErr As Variant

Public Property Get uErrPath() As Variant
    On Error Resume Next
        uErrPath = strErrPath
End Property

Public Property Let uErrPath(ByVal vErrPath As Variant)
    On Error Resume Next
        strErrPath = vErrPath
End Property

Public Property Set uConn(ByVal vConn As ADODB.Connection)
    On Error Resume Next
        Set gcnGi = vConn
End Property

Public Property Get uConn() As ADODB.Connection
    On Error Resume Next
        Set uConn = gcnGi
End Property

Public Property Get uErr() As Variant
    On Error Resume Next
        uErr = strErr
End Property

Public Property Let ugaryGi(ByVal vData As Variant)
    On Error Resume Next
        Call GetGlobal(CStr(vData))
End Property

Public Property Let uErr(ByVal vErr As Variant)
    On Error Resume Next
        strErr = vErr
End Property

Public Property Let uOwnerName(ByVal vErrPath As Variant)
    On Error Resume Next
        gOwnerName = vErrPath
        GaryGi(16) = gOwnerName
        
End Property

