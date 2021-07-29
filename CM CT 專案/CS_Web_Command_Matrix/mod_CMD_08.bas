Attribute VB_Name = "mod_CMD_08"
Option Explicit

'(8) 查詢 HFC 服務類別

'HFCTypeQuery

Public Function JustDoIt(ParamArray ParaAry() As Variant) As Boolean
  On Error GoTo ChkErr
    
    strPostData = Empty
    strCMD = "/HFCTypeQuery"
    
    If IsMissing(ParaAry(0)) Or IsMissing(ParaAry(1)) Or IsMissing(ParaAry(2)) Or _
        IsMissing(ParaAry(3)) Or IsMissing(ParaAry(4)) Or IsMissing(ParaAry(5)) Or _
        IsMissing(ParaAry(6)) Or IsMissing(ParaAry(7)) Or IsMissing(ParaAry(8)) Or _
        IsMissing(ParaAry(9)) Then
        
        JustDoIt = False
        strErrorMessage = "參數不足 !"
'        MsgBox "參數不足 !", vbInformation, "訊息"

    Else
    
'        ParaAry(0) = "WORKID=" & ParaAry(0)
'        ParaAry(1) = "CUSSO=" & ParaAry(1)
'        ParaAry(2) = "Zone=" & ParaAry(2)
'        ParaAry(3) = "Lin=" & ParaAry(3)
'        ParaAry(4) = "Section=" & ParaAry(4)
'        ParaAry(5) = "Lane=" & ParaAry(5)
'        ParaAry(6) = "Alley=" & ParaAry(6)
'        ParaAry(7) = "SubAlley=" & ParaAry(7)
'        ParaAry(8) = "NO1=" & ParaAry(8)
'        ParaAry(9) = "NO2=" & ParaAry(9)
'
'        strPostData = Join(ParaAry, "&")
        JustDoIt = True
        
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_CMD_08", "JustDoIt"
End Function

'SOAP
'下列是 SOAP 要求與回應的範例。替代符號顯示的位置必須代入實際的值。
'
'POST /CMWebServiceGWDemo/EMCWebServiceGW.asmx HTTP/1.1
'Host: 210.202.146.97
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'SOAPAction: "http://tempuri.org/CMWebServiceGW/EMCWebServiceGW/HFCTypeQuery"
'
'<?xml version="1.0" encoding="utf-8"?>
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
'  <soap:Body>
'    <HFCTypeQuery xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW">
'      <WORKID>string</WORKID>
'      <CUSSO>string</CUSSO>
'      <Zone>string</Zone>
'      <Lin>string</Lin>
'      <Section>string</Section>
'      <Lane>string</Lane>
'      <Alley>string</Alley>
'      <SubAlley>string</SubAlley>
'      <NO1>string</NO1>
'      <NO2>string</NO2>
'    </HFCTypeQuery>
'  </soap:Body>
'</soap:Envelope>
'HTTP/1.1 200 OK
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'
'<?xml version="1.0" encoding="utf-8"?>
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
'  <soap:Body>
'    <HFCTypeQueryResponse xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW">
'      <HFCTypeQueryResult>
'        <xsd:schema>schema</xsd:schema>xml</HFCTypeQueryResult>
'    </HFCTypeQueryResponse>
'  </soap:Body>
'</soap:Envelope>
'HTTP GET
'下列是 HTTP GET 要求與回應的範例。替代符號顯示的位置必須代入實際的值。
'
'GET /CMWebServiceGWDemo/EMCWebServiceGW.asmx/HFCTypeQuery?WORKID=string&CUSSO=string&Zone=string&Lin=string&Section=string&Lane=string&Alley=string&SubAlley=string&NO1=string&NO2=string HTTP/1.1
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
'POST /CMWebServiceGWDemo/EMCWebServiceGW.asmx/HFCTypeQuery HTTP/1.1
'Host: 210.202.146.97
'Content-Type: application/x-www-form-urlencoded
'Content -length: length
'
'WORKID=string&CUSSO=string&Zone=string&Lin=string&Section=string&Lane=string&Alley=string&SubAlley=string&NO1=string&NO2=string
'HTTP/1.1 200 OK
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'
'<?xml version="1.0" encoding="utf-8"?>
'<DataSet xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW">
'  <schema xmlns="http://www.w3.org/2001/XMLSchema">schema</schema>xml</DataSet>
'
