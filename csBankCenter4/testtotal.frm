VERSION 5.00
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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   525
      Left            =   1650
      TabIndex        =   1
      Top             =   2070
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   1680
      TabIndex        =   0
      Top             =   540
      Width           =   1245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim FSO As New FileSystemObject
    Dim FILE As TextStream
    Dim STR1 As String
    Dim lngtotal As Long
    Dim lngcount As Long
    Set FILE = FSO.OpenTextFile("C:\TEMP\TBMAUTO(910102).txt")
    Do While Not FILE.AtEndOfStream
        STR1 = FILE.ReadLine
        If Left(STR1, 1) = 2 Then
            lngtotal = lngtotal + Mid(STR1, 41, 12)
            lngcount = lngcount + 1
        End If
    Loop
    Debug.Print lngtotal
    Debug.Print lngcount
    FILE.Close
    Set FILE = Nothing
    Set FSO = Nothing
End Sub

Private Sub Command2_Click()
    Dim FSO As New FileSystemObject
    Dim FILE As TextStream
    Dim STR1 As String
    Dim strSQl As String
    Dim lngcount As Long
    Set FILE = FSO.OpenTextFile("C:\TEMP\TBMAUTO(910102).txt")
    Do While Not FILE.AtEndOfStream
        STR1 = FILE.ReadLine
        If Left(STR1, 1) = 2 Then
            strSQl = strSQl & ",'" & (Mid(STR1, 75, 4) + 191100) & Mid(STR1, 79, 1) & "C" & Format(Mid(STR1, 80, 6), "0000000") & "'"
        End If
    Loop
    Debug.Print Mid(strSQl, 2)
    FILE.Close
    Set FILE = Nothing
    Set FSO = Nothing
End Sub
