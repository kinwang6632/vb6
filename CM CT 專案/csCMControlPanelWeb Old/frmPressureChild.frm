VERSION 5.00
Begin VB.Form frmPressureChild 
   Caption         =   "機上盒作業"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmPressureChild.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "frmPressureChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intProcessItem As Integer
Dim intCount As Integer
Dim objControl As Object
Dim strcmdStatus As String, strErrorCode As String, strErrorMsg As String
Dim strRowId As String

Private Sub Form_Load()
    Select Case intProcessItem
        Case 11
            Call OpenProcess
        Case 12
            Call CloseProcess
        Case 13
            Call StopProcess
        Case 14
            Call ReOpenProcess
        Case 15
            Call DeleteProcess
        Case 21
            Call OpenChProcess
        Case 22
            Call CloseChProcess
        Case 23
            Call StopChProcess
        Case 24
            Call ReOpenChProcess
        Case 25
            Call ExtendRangeProcess
        Case 26
            Call ChPrivCopyProcess
        Case 31
            Call ClearPinCodeProcess
        Case 32
            Call SendMsgProcess
    End Select
    objControl = objControl + 1
End Sub

Public Property Let uProcessItem(ByVal vData As Integer)
    intProcessItem = vData
End Property

Public Property Set uControl(ByVal vData As Object)
    Set objControl = vData
End Property

Private Sub OpenProcess()
    Dim strZipCode As String
    Dim strRecallDate As String
    Dim strRecallTime As String
        If Not SendToInterface(strRowId, 1, "A1", "1234567890", "1234567890abcd", , , , strcmdStatus, , , strErrorCode, strErrorMsg, , "12345", , , 0, , , , _
            "1234567", "192.168.010.123", 80, Format(RightNow + 5, "yyyymmdd"), Format(RightNow + 5, "hhmmss"), "3m", Format(RightDate, "yyyy/mm/dd"), "1234567", "1234567", "1234567", , , , , False) Then Exit Sub
            
End Sub

Private Sub CloseProcess()
        If Not SendToInterface(strRowId, 1, "A2", "1234567890", , , , , strcmdStatus, , , strErrorCode, strErrorMsg, , , , , , , , , , , , , , , , , , , , , , , False) Then Exit Sub
        
End Sub

Private Sub StopProcess()
        If Not SendToInterface(strRowId, 1, "A3", "1234567890", , , , , strcmdStatus, , , strErrorCode, strErrorMsg, , , , , , , , , , , , , , , , , , , , , , , False) Then Exit Sub

End Sub

Private Sub ReOpenProcess()
        If Not SendToInterface(strRowId, 1, "A4", "1234567890", , , , , strcmdStatus, , , strErrorCode, strErrorMsg, , , , , , , , , , , , , , , , , , , , , , , False) Then Exit Sub

End Sub

Private Sub DeleteProcess()
        If Not SendToInterface(strRowId, 1, "A5", "1234567890", , , , , strcmdStatus, , , strErrorCode, strErrorMsg, , , , , , , , , , , , , , , , , , , , , , , False) Then Exit Sub

End Sub

Private Sub OpenChProcess()
        If Not SendToInterface(strRowId, 1, "B1", "1234567890", , Format(RightDate, "yyyy/mm/dd"), Format(RightDate + 1000, "yyyy/mm/dd"), "A~000000000012", strcmdStatus, , , strErrorCode, strErrorMsg, , , , , , , , , , , , , , , , , , , , , , , False) Then Exit Sub

End Sub

Private Sub CloseChProcess()
        If Not SendToInterface(strRowId, 1, "B2", "1234567890", , , , "000000000012", strcmdStatus, , , strErrorCode, strErrorMsg, , , , , , , , , , , , , , , , , , , , , , , False) Then Exit Sub

End Sub

Private Sub StopChProcess()
        If Not SendToInterface(strRowId, 1, "B4", "1234567890", , , , "000000000012", strcmdStatus, , , strErrorCode, strErrorMsg, , , , , , , , , , , , , , , , , , , , , , , False) Then Exit Sub

End Sub

Private Sub ReOpenChProcess()
        If Not SendToInterface(strRowId, 1, "B5", "1234567890", , , , "000000000012", strcmdStatus, , , strErrorCode, strErrorMsg, , , , , , , , , , , , , , , , , , , , , , , False) Then Exit Sub

End Sub

Private Sub ExtendRangeProcess()
        If Not SendToInterface(strRowId, 1, "B6", "1234567890", , , Format(RightDate, "yyyy/mm/dd"), "A~000000000012", strcmdStatus, , , strErrorCode, strErrorMsg, , , , , , , , , , , , , , , , , , , , , , , False) Then Exit Sub

End Sub

Private Sub ChPrivCopyProcess()
        If Not SendToInterface(strRowId, 1, "B7", "1234567890", , , , "000000000012^20020901^20030901,000000000013^20020901^20030901~1234567891", strcmdStatus, , , strErrorCode, strErrorMsg, , , , , , , , , , , , , , , , , , , , , , , False) Then Exit Sub

End Sub

Private Sub ClearPinCodeProcess()
        If Not SendToInterface(strRowId, 1, "E1", "1234567890", , , , , strcmdStatus, , , strErrorCode, strErrorMsg, , , , , , , , , , , , , , , , , , , , , , , False) Then Exit Sub

End Sub

Private Sub SendMsgProcess()
        If Not SendToInterface(strRowId, 1, "E2", "1234567890", , , , "斷訊斷訊明天將斷訊,請客戶等待", strcmdStatus, , , strErrorCode, strErrorMsg, , , , , , , , , , , , , , , , , , , , , , , False) Then Exit Sub

End Sub
