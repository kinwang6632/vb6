VERSION 5.00
Begin VB.Form frmEMC 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   1718
      TabIndex        =   0
      Top             =   1335
      Width           =   1245
   End
End
Attribute VB_Name = "frmEMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objWG As New CS_Web.clsEMCcommand
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
    Const strStatus = "CMStatusQuery.xml"
    Const strLog = "CMPCLog.xml"
    Dim strXMLfile As String
    strXMLfile = App.Path & "\XML Data\" & strStatus
    With objWG
            Set rs = .Return_XML_RS(strXMLfile)
    End With
End Sub

Private Sub Form_Load()

End Sub
