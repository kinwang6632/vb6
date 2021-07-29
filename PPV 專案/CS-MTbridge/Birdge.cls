VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Birdge"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const strCrLf = vbCrLf & vbCrLf

Public Function StartUp(ByVal strPara As String) As String
  On Error GoTo ChkErr
    StartUp = CreateObject("CS_Web.GateWay").Executing(strPara)
  Exit Function
ChkErr:
    ErrHandle , "StartUp"
End Function

'Public Function StartUp(ByVal strPara As String, Optional ByVal strProgID As String = "") As String
'  On Error GoTo ChkErr
'    If strProgID = Empty Then
'        StartUp = CreateObject("CS_Web.GateWay").Executing(strPara)
'    Else
'        StartUp = CreateObject(strProgID).Executing(strPara)
'    End If
'  Exit Function
'ChkErr:
'    ErrHandle , "StartUp"
'End Function

Public Function Linking(ByVal strPara As String, ByVal strProgID As String) As String
  On Error GoTo ChkErr
    Linking = CreateObject(strProgID).Executing(strPara)
  Exit Function
ChkErr:
    ErrHandle , "Linking"
End Function

Private Sub ErrHandle(Optional FormName As String = "CS_MT.Bridge", Optional ProcedureName As String = "")
    Dim strErr As String
    Dim ErrNum As Variant, ErrDesc As Variant, ErrSrc As Variant
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSrc = Err.Source
    On Error Resume Next
    Screen.MousePointer = vbDefault
    DoEvents
    strErr = "發生錯誤時間 : " & CStr(Now) & strCrLf & _
                "發生錯誤專案 : " & CStr(ErrSrc) & strCrLf & _
                "發生錯誤表單 : " & FormName & strCrLf & _
                "發生錯誤程序 : " & ProcedureName & strCrLf & _
                "錯誤代碼 : " & CStr(ErrNum) & strCrLf & _
                "錯誤原因 : " & CStr(ErrDesc)
    MsgBox strErr, vbOKOnly + vbInformation, "系統執行錯誤!"
End Sub


