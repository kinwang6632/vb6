Attribute VB_Name = "mod_CMD_10"
Option Explicit

'(10) CM �s�u�~��d��

'CMQALog

Public Function JustDoIt(ParamArray ParaAry() As Variant) As Boolean
  On Error GoTo ChkErr
    
    strPostData = Empty
    strCMD = "/CMQALog"
    'strWork_ID , strCust_SO, lngCustID, strCM_MAC, strStart_Date_Time, strEnd_Date_Time
    
    If IsMissing(ParaAry(0)) Or IsMissing(ParaAry(1)) Or IsMissing(ParaAry(2)) Or _
        IsMissing(ParaAry(3)) Or IsMissing(ParaAry(4)) Or IsMissing(ParaAry(5)) Then
        
        JustDoIt = False
        strErrorMessage = "�ѼƤ��� !"
'        MsgBox "�ѼƤ��� !", vbInformation, "�T��"
    
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
    ErrHandle "mod_CMD_10", "JustDoIt"
End Function

'SOAP
'�U�C�O SOAP �n�D�P�^�����d�ҡC���N�Ÿ���ܪ���m�����N�J��ڪ��ȡC
'
'POST /CMWebServiceGWDemo/EMCWebServiceGW.asmx HTTP/1.1
'Host: 210.202.146.97
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'SOAPAction: "http://tempuri.org/CMWebServiceGW/EMCWebServiceGW/CMQALog"
'
'<?xml version="1.0" encoding="utf-8"?>
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
'  <soap:Body>
'    <CMQALog xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW">
'      <WORKID>string</WORKID>
'      <CUSSO>string</CUSSO>
'      <CUSID>string</CUSID>
'      <CMMAC>string</CMMAC>
'      <StartDateTime>string</StartDateTime>
'      <EndDateTime>string</EndDateTime>
'    </CMQALog>
'  </soap:Body>
'</soap:Envelope>
'HTTP/1.1 200 OK
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'
'<?xml version="1.0" encoding="utf-8"?>
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
'  <soap:Body>
'    <CMQALogResponse xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW">
'      <CMQALogResult>
'        <xsd:schema>schema</xsd:schema>xml</CMQALogResult>
'    </CMQALogResponse>
'  </soap:Body>
'</soap:Envelope>
'HTTP GET
'�U�C�O HTTP GET �n�D�P�^�����d�ҡC���N�Ÿ���ܪ���m�����N�J��ڪ��ȡC
'
'GET /CMWebServiceGWDemo/EMCWebServiceGW.asmx/CMQALog?WORKID=string&CUSSO=string&CUSID=string&CMMAC=string&StartDateTime=string&EndDateTime=string HTTP/1.1
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
'POST /CMWebServiceGWDemo/EMCWebServiceGW.asmx/CMQALog HTTP/1.1
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
