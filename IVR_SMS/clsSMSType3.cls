VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSMSType3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const clsName As String = "clsSMSType3"
Public Sub SendSMS()
  On Error GoTo ChkErr
    Dim strPhoneNo As String
    Dim strUrl As String
    Dim p As Variant
    Dim RetMsg As Variant
    Dim intIdx As Long
    Dim intTickCount As Long
    Dim lngTimer As Long
    If GetXMLobj(objxmlHttp) Then
        strUrl = strSMSUrl & "?id=" & strUser & "&passwd=" & strPwd
        For Each p In Split(strPhone, ",")
            If p <> "" Then
                strUrl = strUrl & "&tel" & intIdx & "=" & p
                intIdx = intIdx + 1
            End If
        Next
        strUrl = strUrl & "&tran_type=now&message=" & MidMbcs(MsgToWide(strMsg), 1, 65)
        objxmlHttp.open "GET", strUrl, False
        objxmlHttp.setRequestHeader "Content-type", RequestHeader
        objxmlHttp.send
        strRetMsg = GetPureInfo(StrConv(objxmlHttp.responseBody, vbUnicode))
        blnRetOK = True
    Else
        strRetMsg = "¸ü¤JXMLHTTPª«¥ó±Ñ"
        blnRetOK = False
    End If
  Exit Sub
ChkErr:
    ErrHandle clsName, "SendSMS"
End Sub


Private Function GetPureInfo(strData As String) As String
  On Error GoTo ChkErr
    GetPureInfo = strData
    If InStr(1, GetPureInfo, "</SCRIPT>") Then GetPureInfo = Split(GetPureInfo, "</SCRIPT>")(1)
    If Left(GetPureInfo, 2) = vbCrLf Then GetPureInfo = Mid(GetPureInfo, 3)
    GetPureInfo = Replace(GetPureInfo, "<br>", vbCrLf)
    Exit Function
ChkErr:
  ErrHandle clsName, "GetPureInfo"
End Function

