VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2070
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   525
      Left            =   150
      TabIndex        =   1
      Top             =   2460
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   3030
      TabIndex        =   0
      Top             =   2460
      Width           =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    'http://w3.nvwtv.com.tw/nvwtvsm/msgsend.asp?id=10576&passwd=nvwtv1234&tel0=9999999999&tel1=9999999999&tel2=9999999999&tran_type=now&message=訊息內容
  Dim obj As Object
  Set obj = CreateObject("CS_IVR_SMS.clsInterface")
  With obj
    .uUrl = "http://w3.nvwtv.com.tw/nvwtvsm/msgsend.asp"
    .uSMSType = 3
    .uSMSMsg = "訊息內容1,2,@,#,'"
    .uSMSPhone = "0921255093,0921255093,0921255093,0921255093"
    .uUserName = "99999"
    .uPassword = "root123"
    .RunSMS
  End With
  MsgBox obj.uRetMsg '回傳訊息
  MsgBox obj.uRetOk  '回傳結果(True,false)
  Set obj = Nothing
End Sub

Private Sub Command2_Click()
  Dim strurl As String
  'strurl = "http://w3.nvwtv.com.tw/nvwtvsm/msgsend.asp?id=10576&passwd=nvwtv1234&tel0=9999999999&tel1=9999999999&tel2=9999999999&tran_type=now&message=訊息內容"
  strurl = "http://219.87.9.11:8080/BUM/SendSMS?username=cns7do&password=cns7do&message=簡訊內容1,2;@&target1=9999999"
  Dim msg As String
  msg = Inet1.OpenURL(strurl)
  
  MsgBox msg
End Sub
