VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNoExeCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function NoExeCM(ByRef cn As ADODB.Connection, ByRef rs As ADODB.Recordset, _
                             ByVal strCMD As String, ByVal strOwnerName As String, _
                             ByVal strCMownerName As String, _
                             ByVal strCmdSeq As String, _
                             ByRef GaryGi() As String, _
                             Optional ByVal blnBeginTrans As Boolean = True) As Boolean
    On Error Resume Next
    Dim rsMemory As New ADODB.Recordset
    Dim rsArray As String
    If rs Is Nothing Then Exit Function
    strCMowner = strCMownerName
    strOwner = strOwnerName
    intCompCode = rs("CompCode")
    strCmdSeqNo = strCmdSeqNo
    strArry = Join(GaryGi, Chr(9))
    Set gcnGi = cn
    Call GetGlobal(strArry)
    If Not blnBeginTrans Then gcnGi.BeginTrans
    Select Case strCMD
        Case "10"
            GetScmData cn, rsMemory
            If Not rsResult Is Nothing Then
                rsData("ServiceType") = rsMemory("ServiceType")
                rsData("DialAccount") = rsMemory("AccountId")
                rsData("DialPassWord") = rsMemory("AccountPassword")
                rsData("InstDate") = Format(RightNow, "YYYY/MM/DD HH:MM:SS")
                rsData("UpdEn") = GaryGi(1)
                rsData("UpdTime") = GetDTString(RightNow)
                rsData.Update
            End If
        Case "11"
            rsData("CMOpenDate") = Format(RightNow, "YYYY/MM/DD HH:MM:SS")
            rsData("UpdEn") = GaryGi(1)
            rsData("UpdTime") = GetDTString(RightNow)
            rsData.Update
        Case "12"
            rsData("CMCloseDate") = Format(RightNow, "YYYY/MM/DD HH:MM:SS")
            rsData("UpdEn") = GaryGi(1)
            rsData("UpdTime") = GetDTString(RightNow)
            rsData.Update
        Case "13"
            rsData("FaciStatusCode") = 3
            rsData.Update
            'strCmdName = "°±¾÷"
        Case "14"
            rsData("FaciStatusCode") = 4
            rsData.Update
            'strCmdName = "´_¾÷"
        Case "15"
            rsData("PRDate") = Format(RightNow, "YYYY/MM/DD HH:MM:SS")
            rsData("UpdEn") = GaryGi(1)
            rsData("UpdTime") = GetDTString(RightNow)
            rsData.Update
        Case "16"
    End Select
    Call CloseRecordset(rsMemory)
    If Not blnBeginTrans Then gcnGi.CommitTrans
    NoExeCM = True
End Function


'    Select Case strID_Code
'        Case "11"
'            rsData("CMOpenDate") = Format(RightNow, "YYYY/MM/DD HH:MM:SS")
'            rsData("UpdEn") = GaryGi(1)
'            rsData("UpdTime") = GetDTString(RightNow)
'            rsData.Update
'        Case "12"
'            rsData("CMCloseDate") = Format(RightNow, "YYYY/MM/DD HH:MM:SS")
'            rsData("UpdEn") = GaryGi(1)
'            rsData("UpdTime") = GetDTString(RightNow)
'            rsData.Update
'        Case "13"
'            rsData("FaciStatusCode") = 3
'            rsData.Update
'            'strCmdName = "°±¾÷"
'        Case "14"
'            rsData("FaciStatusCode") = 4
'            rsData.Update
'            'strCmdName = "´_¾÷"
'        Case "15"
'            rsData("PRDate") = Format(RightNow, "YYYY/MM/DD HH:MM:SS")
'            rsData("UpdEn") = GaryGi(1)
'            rsData("UpdTime") = GetDTString(RightNow)
'            rsData.Update
'        Case "16"
'            'strCmdName = "ºû­×´«©î"
'    End Select
'    OpenCM = True
'    strResult = "1"
'    GoTo 66
  

