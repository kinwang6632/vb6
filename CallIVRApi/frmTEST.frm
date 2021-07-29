VERSION 5.00
Begin VB.Form frmTEST 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdCall 
      Caption         =   "呼叫"
      Height          =   525
      Left            =   1740
      TabIndex        =   0
      Top             =   1350
      Width           =   1245
   End
End
Attribute VB_Name = "frmTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCall_Click()
    Dim x As New CallIVRApi.IVRApi
    x.Executing "test"
End Sub
