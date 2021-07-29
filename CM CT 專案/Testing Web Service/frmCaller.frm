VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmWScaller 
   Caption         =   "Call Web Service Sample .."
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCaller.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4290
   ScaleWidth      =   8445
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command6 
      Caption         =   "Use M$ SOAP Client"
      Height          =   465
      Left            =   240
      TabIndex        =   6
      Top             =   3540
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Use Web Browser Control"
      Height          =   465
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   2775
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   3765
      Left            =   3210
      TabIndex        =   4
      Top             =   240
      Width           =   4995
      ExtentX         =   8811
      ExtentY         =   6641
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Use Internet Control Dll"
      Height          =   465
      Left            =   240
      TabIndex        =   3
      Top             =   2220
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Use Inet Control"
      Height          =   465
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Call WinInet API"
      Height          =   465
      Left            =   240
      TabIndex        =   1
      Top             =   900
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Use XMLhttp"
      Height          =   465
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmWScaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "Kernel32" () As Long

Private lngST As Long
Private Const strURL = "http://127.0.0.1/ws1/Service1.asmx/GetCustInfo"
'Private Const strURL = "http://210.202.146.97:9090/CMWebServiceGWDemo/EMCWebServiceGW.asmx/HFCTypeQuery"
Private Const strRequestHeader = "application/x-www-form-urlencoded"
Private Const strMethod = "POST"
Private Const strPostData = "strTableName=SO1100A&strCustID=12"
'Private Const strPostData = "WORKID=A&CUSSO=2&Zone=竹田鄉中正路&Lin=&Section=&Lane=&Alley=&SubAlley=&NO1=12&NO2="
'"http://PowerHammerURI.org/ws1/ServiceEveryone/GetCustInfo"

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

Private Sub Command6_Click()
    Dim objSC As New SoapClient
    With objSC
            .mssoapinit "http://PowerHammerURI.org/ws1/ServiceEveryone/GetCustInfo", "", "", ""
    End With
End Sub

Private Sub Command1_Click()
    Dim XMLobj As Object
    Dim DOMobj As Object
    lngST = GetTickCount
    Call GetXMLobj(XMLobj)
    With XMLobj
        .Open "POST", strURL, False, "administrator", "Chou Yin Der is powerhammer"
'        .Open "POST", strURL & "?" & strPostData, False
        .SetRequestHeader "Content-type", strRequestHeader
'        .Send
        .Send strPostData
    End With
    Debug.Print XMLobj.ResponseText
    MsgBox "Spend Time : " & GetTickCount - lngST
'    Call GetDOMdocObj(DOMobj)
'    Set DOMobj = XMLobj.ResponseXML
'    Debug.Print DOMobj.XML
End Sub

Private Sub Command2_Click()
    lngST = GetTickCount
    sCodePage = CP_UTF8
'    Debug.Print mod_Inet_API.HttpPost(strURL, strPostData)
'    Debug.Print CharSetConvert(StrConv(mod_Inet_API.HttpPost(strURL, strPostData), vbFromUnicode), "UTF-8")
'    Debug.Print Utf8Decode(mod_Inet_API.HttpPost(strURL, strPostData),"65001")
    Debug.Print UTF8_Decode(HttpPost(strURL, strPostData))
    MsgBox "Spend Time : " & GetTickCount - lngST
End Sub

Private Sub Command3_Click()
    Dim InetObj As Object
    Dim strData As String
    Dim bytAry() As Byte
    lngST = GetTickCount
    Call GetNetObj(InetObj)
    With InetObj
            Call .Execute(strURL, strMethod, strPostData, "Content-Type: " & strRequestHeader)
            Do Until Not .StillExecuting
                DoEvents
            Loop
'            bytAry() = .GetChunk(999999, 1)
'            strData = StrConv(bytAry, vbUnicode)
            Debug.Print CharSetConvert(StrConv(.GetChunk(999999, 0), vbFromUnicode), "UTF-8")
'            Debug.Print CharSetConvert(StrConv(strData, vbFromUnicode), "UTF-8")
    End With
    Set InetObj = Nothing
    MsgBox "Spend Time : " & GetTickCount - lngST
End Sub

Private Sub Command4_Click()
    Dim IEobj As Object
    Dim bytAry() As Byte
    lngST = GetTickCount
    Set IEobj = CreateObject("InternetExplorer.Application")
    bytAry = StrConv(strPostData, vbFromUnicode)
    With IEobj
            .Visible = True
            .Navigate2 strURL, 0, "", bytAry, "Content-Type: " & strRequestHeader
            Do Until Not .Busy
                DoEvents
            Loop
             On Error Resume Next
            .Quit
    End With
    Set IEobj = Nothing
    MsgBox "Spend Time : " & GetTickCount - lngST
End Sub

Private Sub Command5_Click()
    Dim bytAry() As Byte
    lngST = GetTickCount
    bytAry = StrConv(strPostData, vbFromUnicode)
    With wb
            .Navigate strURL, 0, "", bytAry, "Content-Type: " & strRequestHeader
            Do Until Not .Busy
                DoEvents
            Loop
    End With
    MsgBox "Spend Time : " & GetTickCount - lngST
End Sub

Private Sub Form_Load()
    wb.Navigate "C:\"
End Sub

'SOAP
'下列是 SOAP 要求與回應的範例。替代符號顯示的位置必須代入實際的值。
'
'POST /ws1/Service1.asmx HTTP/1.1
'Host: localhost
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'SOAPAction: "http://tempuri.org/ws1/Service1/GetCustInfo"
'
'<?xml version="1.0" encoding="utf-8"?>
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
'  <soap:Body>
'    <GetCustInfo xmlns="http://tempuri.org/ws1/Service1">
'      <strCustID>string</strCustID>
'    </GetCustInfo>
'  </soap:Body>
'</soap:Envelope>
'HTTP/1.1 200 OK
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'
'<?xml version="1.0" encoding="utf-8"?>
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
'  <soap:Body>
'    <GetCustInfoResponse xmlns="http://tempuri.org/ws1/Service1">
'      <GetCustInfoResult>
'        <xsd:schema>schema</xsd:schema>xml</GetCustInfoResult>
'    </GetCustInfoResponse>
'  </soap:Body>
'</soap:Envelope>
'HTTP POST
'下列是 HTTP POST 要求與回應的範例。替代符號顯示的位置必須代入實際的值。
'
'POST /ws1/Service1.asmx/GetCustInfo HTTP/1.1
'Host: localhost
'Content-Type: application/x-www-form-urlencoded
'Content -length: length
'
'strCustID=string
'HTTP/1.1 200 OK
'Content-Type: text/xml; charset=utf-8
'Content -length: length
'
'<?xml version="1.0" encoding="utf-8"?>
'<DataSet xmlns="http://tempuri.org/ws1/Service1">
'  <schema xmlns="http://www.w3.org/2001/XMLSchema">schema</schema>xml</DataSet>
'
