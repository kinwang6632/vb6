Attribute VB_Name = "mod_CrtObj"
Option Explicit

Public Function GetXMLobj(ByRef objXML As Object, Optional strErrMsg As String = "") As Boolean ' 建立 XML http 物件
  On Error Resume Next
    GetXMLobj = True
    Set objXML = CreateObject("Msxml2.XMLHTTP.5.0")
    If Err.Number <> 0 Then
        Err.Clear
        Set objXML = CreateObject("Msxml2.XMLHTTP.4.0")
        If Err.Number <> 0 Then
            Err.Clear
            Set objXML = CreateObject("Msxml2.XMLHTTP.3.0")
            If Err.Number <> 0 Then
                Err.Clear
                Set objXML = CreateObject("Msxml2.XMLHTTP.2.6")
                If Err.Number <> 0 Then
                    Err.Clear
                    Set objXML = CreateObject("MSXML2.ServerXMLHTTP")
                    If Err.Number <> 0 Then
                        Err.Clear
                        Set objXML = CreateObject("MSXML2.XMLHTTP")
                        If Err.Number <> 0 Then
                            Err.Clear
                            Set objXML = CreateObject("Microsoft.XMLHTTP")
                            If Err.Number <> 0 Then
                                Err.Clear
                                GetXMLobj = False
                                strErrMsg = "無法建立 M$ XML HTTP 物件 !"
                                strErrorMessage = strErrMsg
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

' 建立 XML DOM Document 物件
Public Function GetDOMdocObj(ByRef objDOMdoc As Object, Optional strErrMsg As String = "") As Boolean
  On Error Resume Next
    GetDOMdocObj = True
    Set objDOMdoc = CreateObject("Msxml2.DOMDocument.5.0")
    If Err.Number <> 0 Then
        Err.Clear
        Set objDOMdoc = CreateObject("Msxml2.DOMDocument.4.0")
        If Err.Number <> 0 Then
            Err.Clear
            Set objDOMdoc = CreateObject("Msxml2.DOMDocument.3.0")
            If Err.Number <> 0 Then
                Err.Clear
                Set objDOMdoc = CreateObject("Msxml2.DOMDocument.2.6")
                If Err.Number <> 0 Then
                    Err.Clear
                    Set objDOMdoc = CreateObject("MSXML2.DOMDocument")
                    If Err.Number <> 0 Then
                        Err.Clear
                        Set objDOMdoc = CreateObject("Microsoft.XMLDOM")
                        If Err.Number <> 0 Then
                            Err.Clear
                            objDOMdoc = False
                            strErrMsg = "無法建立 M$ XML DOM Document 物件 !"
                            strErrorMessage = strErrMsg
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Public Function GetNetObj(ByRef objNet As Object, Optional strErrMsg As String = "") As Boolean ' 建立 ITC 物件
  On Error Resume Next
    GetNetObj = True
    Set objNet = CreateObject("InetCtls.Inet")
    If Err.Number <> 0 Then
        Err.Clear
        strErrMsg = "無法建立 M$ Internet Transfer Control 物件 !"
        strErrorMessage = strErrMsg
        GetNetObj = False
    End If
'    objNet.RequestTimeout = 60
'    RetVal = objNet.OpenUrl(strURL)
End Function

