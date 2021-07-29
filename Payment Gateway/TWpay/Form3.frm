VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "流程示意圖"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   8130
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

