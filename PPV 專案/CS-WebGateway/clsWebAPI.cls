VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWebAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private objWebPool As Object

Private Sub Class_Initialize()
  On Error GoTo ChkErr
    Set objWebPool = CreateObject("WebPool.PoolManager")
  Exit Sub
ChkErr:
    ErrHandle "clsWebAPI", "Class_Initialize"
End Sub

Public Function Func1(ByRef strPara As String) As String
  On Error GoTo ChkErr
    Func1 = mod_Func1.JustDoIt(objWebPool, strPara)
  Exit Function
ChkErr:
    ErrHandle "clsWebAPI", "Func1"
    Func1 = "-99," & strErr
End Function

Private Sub Class_Terminate()
  On Error Resume Next
    Rlx objWebPool
End Sub
