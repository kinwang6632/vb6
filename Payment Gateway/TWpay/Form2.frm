VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "函式使用流程﹕"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   7185
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Height          =   435
      Left            =   3330
      TabIndex        =   0
      Top             =   6660
      Width           =   1125
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Label1_Click()
    Unload Me
End Sub
