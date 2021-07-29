Attribute VB_Name = "mod_Common"
Option Explicit

Public Const strCrLf As String = vbCrLf & vbCrLf
Public strErrMsg As String

'Public Function GetNetObj(ByRef objNet As Object, Optional strErrMsg As String = "") As Boolean ' 建立 ITC 物件
'  On Error Resume Next
'    GetNetObj = True
'    Set objNet = CreateObject("InetCtls.Inet")
'    If Err Then
'        Err.Clear
'        strErrMsg = "無法建立 M$ Internet Transfer Control 物件 !"
'        GetNetObj = False
'    End If
''    objNet.RequestTimeout = 60
''    RetVal = objNet.OpenUrl(strURL)
'End Function
'
'Public Function GetXHobj(ByRef objXML As Object, Optional strErrMsg As String = "") As Boolean ' 建立 XML http 物件
'
'  On Error Resume Next
'
'    Dim varAry, varElement
'
'    varAry = Array("XMLHTTP.7.0", "XMLHTTP.6.0", "XMLHTTP.5.0", "XMLHTTP.4.0", _
'                            "XMLHTTP.3.0", "XMLHTTP.2.6", "ServerXMLHTTP", "XMLHTTP")
'
'    GetXHobj = True
'
'    For Each varElement In varAry
'        Set objXML = CreateObject("Msxml2." & varElement)
'        If Err = 0 Then Exit For
'        Err.Clear
'    Next
'
'    If objXML Is Nothing Then
'        Set objXML = CreateObject("Microsoft.XMLHTTP2")
'        If Err Then
'            Err.Clear
'            GetXHobj = False
'            strErrMsg = "無法建立 M$ XML HTTP 物件 !"
'        End If
'    End If
'
'End Function

' 錯誤掌控處理
Public Sub ErrHandle(FormName As String, ProcedureName As String)
    
    Dim ErrNum, ErrDesc, ErrSrc
    
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSrc = Err.Source
    
    On Error Resume Next
    
    Screen.MousePointer = vbDefault
    DoEvents
    
    strErrMsg = "發生錯誤時間 : " & CStr(Now) & strCrLf & _
                        "發生錯誤專案 : " & CStr(ErrSrc) & strCrLf & _
                        "發生錯誤表單 : " & FormName & strCrLf & _
                        "發生錯誤程序 : " & ProcedureName & strCrLf & _
                        "錯誤代碼 : " & CStr(ErrNum) & strCrLf & _
                        "錯誤原因 : " & CStr(ErrDesc)
    
    Debug.Print strErrMsg

End Sub


