VERSION 5.00
Begin VB.Form frmWellcome 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  '沒有框線
   Caption         =   "錄音系統介接 Gateway"
   ClientHeight    =   5265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7860
   Icon            =   "frmWellcome.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.Label lblRcdSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "錄音系統介接Gateway V1.0"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3840
      TabIndex        =   0
      Top             =   4380
      Width           =   3870
   End
   Begin VB.Image imgWellcome 
      Height          =   5205
      Left            =   30
      Picture         =   "frmWellcome.frx":0442
      Top             =   30
      Width           =   7800
   End
End
Attribute VB_Name = "frmWellcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PR : Hammer
' Start date : 2004/12/21
' Last Modify : 2004/12/21

Option Explicit

Private Sub Form_Initialize()
  On Error Resume Next
    With Me
        .Move (Screen.Width - .Width) / 2, (Screen.Height - Me.Height) / 2
         imgWellcome.Refresh
    End With
End Sub

Private Sub Form_Load()
  On Error Resume Next
    Form_On_Top Me
End Sub
