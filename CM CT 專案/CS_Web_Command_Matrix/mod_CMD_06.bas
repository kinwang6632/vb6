Attribute VB_Name = "mod_CMD_06"
Option Explicit

'(6) CM ���A�d��
'
'CMStatusQuery
'
Public Function JustDoIt(ParamArray ParaAry() As Variant) As Boolean
  On Error GoTo ChkErr
    strPostData = Empty
    strCMD = "/CMStatusQuery"
    'strWork_ID, strCust_SO, lngCustID, strCM_MAC
    If IsMissing(ParaAry(0)) Or IsMissing(ParaAry(1)) Or _
        IsMissing(ParaAry(2)) Or IsMissing(ParaAry(3)) Then
    
        JustDoIt = False
        strErrorMessage = "�ѼƤ��� !"
'        MsgBox "�ѼƤ��� !", vbInformation, "�T��"
    
    Else
    
'        ParaAry(0) = "WORKID=" & ParaAry(0)
'        ParaAry(1) = "CUSSO=" & ParaAry(1)
'        ParaAry(2) = "CUSID=" & ParaAry(2)
'        ParaAry(3) = "CMMAC=" & ParaAry(3)
'
'        strPostData = Join(ParaAry, "&")
        JustDoIt = True
    
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_CMD_06", "JustDoIt"
End Function

'SOAP
'�U�C�O SOAP �n�D�P�^�����d�ҡC���N�Ÿ���ܪ���m�����N�J��ڪ��ȡC
'
'POST /CMWebServiceGWDemo/EMCWebServiceGW.asmx HTTP/1.1
'Host: 210.202.146.97
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'SOAPAction: "http://tempuri.org/CMWebServiceGW/EMCWebServiceGW/CMStatusQuery"
'
'<?xml version="1.0" encoding="utf-8"?>
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
'  <soap:Body>
'    <CMStatusQuery xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW">
'      <WORKID>string</WORKID>
'      <CUSSO>string</CUSSO>
'      <CUSID>string</CUSID>
'      <CMMAC>string</CMMAC>
'    </CMStatusQuery>
'  </soap:Body>
'</soap:Envelope>
'HTTP/1.1 200 OK
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'
'<?xml version="1.0" encoding="utf-8"?>
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
'  <soap:Body>
'    <CMStatusQueryResponse xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW">
'      <CMStatusQueryResult>
'        <xsd:schema>schema</xsd:schema>xml</CMStatusQueryResult>
'    </CMStatusQueryResponse>
'  </soap:Body>
'</soap:Envelope>
'HTTP GET
'�U�C�O HTTP GET �n�D�P�^�����d�ҡC���N�Ÿ���ܪ���m�����N�J��ڪ��ȡC
'
'GET /CMWebServiceGWDemo/EMCWebServiceGW.asmx/CMStatusQuery?WORKID=string&CUSSO=string&CUSID=string&CMMAC=string HTTP/1.1
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
'�U�C�O HTTP POST �n�D�P�^�����d�ҡC���N�Ÿ���ܪ���m�����N�J��ڪ��ȡC
'
'POST /CMWebServiceGWDemo/EMCWebServiceGW.asmx/CMStatusQuery HTTP/1.1
'Host: 210.202.146.97
'Content-Type: application/x-www-form-urlencoded
'Content -length: length
'
'WORKID=string&CUSSO=string&CUSID=string&CMMAC=string
'HTTP/1.1 200 OK
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'
'<?xml version="1.0" encoding="utf-8"?>
'<DataSet xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW">
'  <schema xmlns="http://www.w3.org/2001/XMLSchema">schema</schema>xml</DataSet>
'
