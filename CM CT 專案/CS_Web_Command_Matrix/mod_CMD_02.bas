Attribute VB_Name = "mod_CMD_02"
Option Explicit

'(2) CP IP 申請及退租
'
'CPIPReg
'
Public Function JustDoIt(ParamArray ParaAry() As Variant) As Boolean
  On Error GoTo ChkErr
    strPostData = Empty
    strCMD = "/CPIPReg"
    
    If IsMissing(ParaAry(0)) Or IsMissing(ParaAry(1)) Or IsMissing(ParaAry(2)) Or _
        IsMissing(ParaAry(3)) Or IsMissing(ParaAry(4)) Or IsMissing(ParaAry(5)) Or _
        IsMissing(ParaAry(6)) Or IsMissing(ParaAry(7)) Or IsMissing(ParaAry(8)) Or _
        IsMissing(ParaAry(9)) Or IsMissing(ParaAry(10)) Or IsMissing(ParaAry(11)) Or _
        IsMissing(ParaAry(12)) Or IsMissing(ParaAry(13)) Or IsMissing(ParaAry(14)) Then
        
        JustDoIt = False
        strErrorMessage = "參數不足 !"
'        MsgBox "參數不足 !", vbInformation, "訊息"
        
    Else
        
'        ParaAry(0) = "WORKID=" & ParaAry(0)
'        ParaAry(1) = "CUSSO=" & ParaAry(1)
'        ParaAry(2) = "CUSID=" & ParaAry(2)
'        ParaAry(3) = "CMMAC=" & ParaAry(3)
'        ParaAry(4) = "EMTAMAC=" & ParaAry(4)
'        ParaAry(5) = "EMTAPort=" & ParaAry(5)
'        ParaAry(6) = "CPNO=" & ParaAry(6)
'        ParaAry(7) = "Zone=" & ParaAry(7)
'        ParaAry(8) = "Lin=" & ParaAry(8)
'        ParaAry(9) = "Section=" & ParaAry(9)
'        ParaAry(10) = "Lane=" & ParaAry(10)
'        ParaAry(11) = "Alley=" & ParaAry(11)
'        ParaAry(12) = "SubAlley=" & ParaAry(12)
'        ParaAry(13) = "NO1=" & ParaAry(13)
'        ParaAry(14) = "NO2=" & ParaAry(14)
'        ParaAry(15) = "OPERTYPE=" & ParaAry(15)
'
'        strPostData = Join(ParaAry, "&")
        JustDoIt = True
    
    End If
  
  Exit Function
ChkErr:
    ErrHandle "mod_CMD_02", "JustDoIt"
End Function

'SOAP
'下列是 SOAP 要求與回應的範例。替代符號顯示的位置必須代入實際的值。
'
'POST /CMWebServiceGWDemo/EMCWebServiceGW.asmx HTTP/1.1
'Host: 210.202.146.97
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'SOAPAction: "http://tempuri.org/CMWebServiceGW/EMCWebServiceGW/CPIPReg"
'
'<?xml version="1.0" encoding="utf-8"?>
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
'  <soap:Body>
'    <CPIPReg xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW">
'      <WORKID>string</WORKID>
'      <CUSSO>string</CUSSO>
'      <CUSID>string</CUSID>
'      <CMMAC>string</CMMAC>
'      <EMTAMAC>string</EMTAMAC>
'      <EMTAPort>string</EMTAPort>
'      <CPNO>string</CPNO>
'      <Zone>string</Zone>
'      <Lin>string</Lin>
'      <Section>string</Section>
'      <Lane>string</Lane>
'      <Alley>string</Alley>
'      <SubAlley>string</SubAlley>
'      <NO1>string</NO1>
'      <NO2>string</NO2>
'      <OPERTYPE>string</OPERTYPE>
'    </CPIPReg>
'  </soap:Body>
'</soap:Envelope>
'HTTP/1.1 200 OK
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'
'<?xml version="1.0" encoding="utf-8"?>
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
'  <soap:Body>
'    <CPIPRegResponse xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW">
'      <CPIPRegResult>
'        <xsd:schema>schema</xsd:schema>xml</CPIPRegResult>
'    </CPIPRegResponse>
'  </soap:Body>
'</soap:Envelope>
'HTTP GET
'下列是 HTTP GET 要求與回應的範例。替代符號顯示的位置必須代入實際的值。
'
'GET /CMWebServiceGWDemo/EMCWebServiceGW.asmx/CPIPReg?WORKID=string&CUSSO=string&CUSID=string&CMMAC=string&EMTAMAC=string&EMTAPort=string&CPNO=string&Zone=string&Lin=string&Section=string&Lane=string&Alley=string&SubAlley=string&NO1=string&NO2=string&OPERTYPE=string HTTP/1.1
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
'POST /CMWebServiceGWDemo/EMCWebServiceGW.asmx/CPIPReg HTTP/1.1
'Host: 210.202.146.97
'Content-Type: application/x-www-form-urlencoded
'Content -length: length
'
'WORKID=string&CUSSO=string&CUSID=string&CMMAC=string&EMTAMAC=string&EMTAPort=string&CPNO=string&Zone=string&Lin=string&Section=string&Lane=string&Alley=string&SubAlley=string&NO1=string&NO2=string&OPERTYPE=string
'HTTP/1.1 200 OK
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'
'<?xml version="1.0" encoding="utf-8"?>
'<DataSet xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW">
'  <schema xmlns="http://www.w3.org/2001/XMLSchema">schema</schema>xml</DataSet>
'
