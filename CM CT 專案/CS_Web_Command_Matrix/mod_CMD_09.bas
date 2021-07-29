Attribute VB_Name = "mod_CMD_09"
Option Explicit

'(9) CM 連線紀錄查詢

'CMConnLog

Public Function JustDoIt(ParamArray ParaAry() As Variant) As Boolean
  On Error GoTo ChkErr
    
    strPostData = Empty
    strCMD = "/CMConnLog"
    'strWork_ID , strCust_SO, lngCustID, strCM_MAC, strStart_Date_Time, strEnd_Date_Time
    
    If IsMissing(ParaAry(0)) Or IsMissing(ParaAry(1)) Or IsMissing(ParaAry(2)) Or _
        IsMissing(ParaAry(3)) Or IsMissing(ParaAry(4)) Or IsMissing(ParaAry(5)) Then
        
        JustDoIt = False
        strErrorMessage = "參數不足 !"
'        MsgBox "參數不足 !", vbInformation, "訊息"
        
    Else
    
'        ParaAry(0) = "WORKID=" & ParaAry(0)
'        ParaAry(1) = "CUSSO=" & ParaAry(1)
'        ParaAry(2) = "CUSID=" & ParaAry(2)
'        ParaAry(3) = "CMMAC=" & ParaAry(3)
'        ParaAry(4) = "StartDateTime=" & Format(ParaAry(4), "YYYY-MM-DD HH:MM:SS")
'        ParaAry(5) = "EndDateTime=" & Format(ParaAry(5), "YYYY-MM-DD HH:MM:SS")
'
'        strPostData = Join(ParaAry, "&")
        JustDoIt = True
    
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_CMD_09", "JustDoIt"
End Function

'SOAP
'下列是 SOAP 要求與回應的範例。替代符號顯示的位置必須代入實際的值。
'
'POST /CMWebServiceGWDemo/EMCWebServiceGW.asmx HTTP/1.1
'Host: 210.202.146.97
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'SOAPAction: "http://tempuri.org/CMWebServiceGW/EMCWebServiceGW/CMConnLog"
'
'<?xml version="1.0" encoding="utf-8"?>
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
'  <soap:Body>
'    <CMConnLog xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW">
'      <WORKID>string</WORKID>
'      <CUSSO>string</CUSSO>
'      <CUSID>string</CUSID>
'      <CMMAC>string</CMMAC>
'      <StartDateTime>string</StartDateTime>
'      <EndDateTime>string</EndDateTime>
'    </CMConnLog>
'  </soap:Body>
'</soap:Envelope>
'HTTP/1.1 200 OK
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'
'<?xml version="1.0" encoding="utf-8"?>
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
'  <soap:Body>
'    <CMConnLogResponse xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW">
'      <CMConnLogResult>
'        <xsd:schema>schema</xsd:schema>xml</CMConnLogResult>
'    </CMConnLogResponse>
'  </soap:Body>
'</soap:Envelope>
'HTTP GET
'下列是 HTTP GET 要求與回應的範例。替代符號顯示的位置必須代入實際的值。
'
'GET /CMWebServiceGWDemo/EMCWebServiceGW.asmx/CMConnLog?WORKID=string&CUSSO=string&CUSID=string&CMMAC=string&StartDateTime=string&EndDateTime=string HTTP/1.1
'Host: 210.202.146.97
'
'HTTP/1.1 200 OK
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'
'<?xml version="1.0" encoding="utf-8"?>
'<DataSet xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW">
'  <schema xmlns="http://www.w3.org/2001/XMLSchema">schema</schema>xml</DataSet>
'HTTP POST
'下列是 HTTP POST 要求與回應的範例。替代符號顯示的位置必須代入實際的值。
'
'POST /CMWebServiceGWDemo/EMCWebServiceGW.asmx/CMConnLog HTTP/1.1
'Host: 210.202.146.97
'Content-Type: application/x-www-form-urlencoded
'Content -length: length
'
'WORKID=string&CUSSO=string&CUSID=string&CMMAC=string&StartDateTime=string&EndDateTime=string
'HTTP/1.1 200 OK
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'
'<?xml version="1.0" encoding="utf-8"?>
'<DataSet xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW">
'  <schema xmlns="http://www.w3.org/2001/XMLSchema">schema</schema>xml</DataSet>
'
