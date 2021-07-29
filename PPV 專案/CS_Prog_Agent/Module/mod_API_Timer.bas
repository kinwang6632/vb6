Attribute VB_Name = "mod_API_Timer"
' PR : Hammer
' Start date : 2004/12/25
' Last Modify : 2004/12/25

Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private bytThreadID As Byte

' API 觸發事件
Private Sub APItimer()
  On Error Resume Next
    frmMain.vTimer
End Sub

' 啟動 Timer
Public Sub TimerGo(ByRef tmr As Long, Optional ByVal Interval As Single = 0, Optional ID As Byte = 0)
  On Error GoTo ChkErr
    bytThreadID = ID
    If tmr <> 0 Then tmr = KillTimer(0, tmr)
    Interval = Val(Interval) * Val(1000)
    tmr = SetTimer(0, 0, Interval, AddressOf APItimer)
  Exit Sub
ChkErr:
    ErrHandle "mod_API_Timer", "Subroutine : TimerGo"
End Sub

' 停止 Timer
Public Sub TimerStop(ByVal tmr As Long)
  On Error GoTo ChkErr
    tmr = KillTimer(0, tmr)
  Exit Sub
ChkErr:
    ErrHandle "mod_API_Timer", "Subroutine : TimerStop"
End Sub
