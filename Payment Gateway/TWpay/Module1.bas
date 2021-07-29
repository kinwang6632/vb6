Attribute VB_Name = "Module1"
Option Explicit

Public strErrMsg As String
Private strCrLf As String

Public Function GetNetObj(ByRef objNet As Object, Optional strErrMsg As String = "") As Boolean ' �إ� ITC ����
  On Error Resume Next
    GetNetObj = True
    Set objNet = CreateObject("InetCtls.Inet")
    If Err Then
        Err.Clear
        strErrMsg = "�L�k�إ� M$ Internet Transfer Control ���� !"
        GetNetObj = False
    End If
'    objNet.RequestTimeout = 60
'    RetVal = objNet.OpenUrl(strURL)
End Function

Public Function GetXHobj(ByRef objXML As Object, Optional strErrMsg As String = "") As Boolean ' �إ� XML http ����
  
  On Error Resume Next
    
    Dim varAry, varElement
    
    varAry = Array("XMLHTTP.7.0", "XMLHTTP.6.0", "XMLHTTP.5.0", "XMLHTTP.4.0", _
                            "XMLHTTP.3.0", "XMLHTTP.2.6", "ServerXMLHTTP", "XMLHTTP")
    
    GetXHobj = True
    
    For Each varElement In varAry
        Set objXML = CreateObject("Msxml2." & varElement)
        If Err = 0 Then Exit For
        Err.Clear
    Next
    
    If objXML Is Nothing Then
        Set objXML = CreateObject("Microsoft.XMLHTTP2")
        If Err Then
            Err.Clear
            GetXHobj = False
            strErrMsg = "�L�k�إ� M$ XML HTTP ���� !"
        End If
    End If

End Function

Public Sub ErrHandle(FormName As String, ProcedureName As String)     ' ���~�x���B�z
    Dim ErrNum As Variant, ErrDesc As Variant, ErrSrc As Variant
    Dim strErr As String
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSrc = Err.Source
    On Error Resume Next
    Screen.MousePointer = vbDefault
    DoEvents
    strErr = "�o�Ϳ��~�ɶ� : " & CStr(Now) & strCrLf & _
                "�o�Ϳ��~�M�� : " & CStr(ErrSrc) & strCrLf & _
                "�o�Ϳ��~��� : " & FormName & strCrLf & _
                "�o�Ϳ��~�{�� : " & ProcedureName & strCrLf & _
                "���~�N�X : " & CStr(ErrNum) & strCrLf & _
                "���~��] : " & CStr(ErrDesc)
    Debug.Print strErr
    strErrMsg = strErr
End Sub
