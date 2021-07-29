Attribute VB_Name = "mod_ComDialog"
'' PR : Hammer
'' Start date : 2004/12/24
'' Last Modify : 2004/12/25
'
'Option Explicit
'
'Public lWndProc As Long
Public hHook As Long, lHookWndProc As Long
'Private Declare Function GetUpdateRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
'Private ICurrentHover As Long
'Private GlobalPos     As Integer
'Public DlgNameIsColor As Boolean
'
'Dim yN As Long
'Dim xN As Long
'
'Public Function AppHook(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Dim CWP As CWPSTRUCT
'    CopyMemory CWP, ByVal lParam, Len(CWP)
'    Select Case CWP.message
'            Case WM_INITDIALOG
'                    lWndProc = SetWindowLong(CWP.hwnd, GWL_WNDPROC, AddressOf Dlg_WndProc)
'                    AppHook = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
'                    UnhookWindowsHookEx hHook
'                    hHook = 0
'                    Exit Function
'    End Select
'    AppHook = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
'End Function
'
'Public Function Dlg_WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Dim TempNum As Long
'    Dim iHw As Integer, iLW As Integer
'    Select Case Msg
'            Case WM_INITDIALOG
'                    DlgNameIsColor = IIf(Left$(GetObjectText(hwnd), InStr(GetObjectText(hwnd), vbNullChar) - 1) = "Color", True, False)
'                    ICurrentHover = 41
'                    GlobalPos = 41
'                    EnumChildWindows hwnd, AddressOf EnumChildProc, ByVal 0&
'            Case WM_PAINT
'                    If DlgNameIsColor Then CleanColorPanel hwnd
'            Case WM_MOUSEMOVE
'                    If DlgNameIsColor Then
'                        TempNum = CheckArea(hwnd)
'                        If TempNum <> ICurrentHover Then
'                            ICurrentHover = TempNum
'                            CleanColorPanel hwnd
'                        End If
'                    End If
'            Case WM_LBUTTONDOWN
'                    If DlgNameIsColor Then
'                        GlobalPos = CheckArea(hwnd)
'                        ICurrentHover = GlobalPos
'                        RedrawWindow hwnd, ByVal 0&, ByVal 0&, &H2 '//---(invoke a Paint-event)
'                    End If
'            Case WM_LBUTTONUP
'                    If DlgNameIsColor Then CleanColorPanel hwnd
'            Case WM_KEYDOWN
'                    If DlgNameIsColor Then
'                        LongInt2Int wParam, iHw, iLW
'                        Select Case (iLW)
'                                Case vbKeyDown
'                                        GlobalPos = GlobalPos + 8
'                                        If GlobalPos > 48 Then
'                                            GlobalPos = GlobalPos - 8
'                                        Else
'                                            ICurrentHover = GlobalPos
'                                            CleanColorPanel hwnd
'                                        End If
'                                Case vbKeyUp
'                                        GlobalPos = GlobalPos - 8
'                                        If GlobalPos < 0 Then
'                                            GlobalPos = GlobalPos + 8
'                                        Else
'                                            ICurrentHover = GlobalPos
'                                            CleanColorPanel hwnd
'                                        End If
'                                Case vbKeyRight
'                                        GlobalPos = GlobalPos + 1
'                                        If GlobalPos = 9 Or GlobalPos = 17 Or GlobalPos = 25 Or GlobalPos = 33 Or GlobalPos = 41 Or GlobalPos = 49 Then
'                                            GlobalPos = GlobalPos - 1
'                                        Else
'                                            ICurrentHover = GlobalPos
'                                            CleanColorPanel hwnd
'                                        End If
'                                Case vbKeyLeft
'                                        GlobalPos = GlobalPos - 1
'                                        If GlobalPos = 0 Or GlobalPos = 8 Or GlobalPos = 16 Or GlobalPos = 24 Or GlobalPos = 32 Or GlobalPos = 40 Then
'                                            GlobalPos = GlobalPos + 1
'                                        Else
'                                            ICurrentHover = GlobalPos
'                                            CleanColorPanel hwnd
'                                        End If
'                                Case vbKeySpace
'                                        ICurrentHover = GlobalPos
'                                        RedrawWindow hwnd, ByVal 0&, ByVal 0&, &H2 '//---(invoke a Paint-event)
'                        End Select
'                    End If
'    End Select
'    Dlg_WndProc = CallWindowProc(lWndProc, hwnd, Msg, wParam, lParam)
'End Function
'
'Private Sub CleanColorPanel(ByVal hwnd As Long)
'    Dim TempRect As RECT
'    Dim cHdc As Long
'    Dim cI As Long
'    cHdc = GetDC(hwnd)
'    With TempRect
'        .Left = 6: .Top = 23: .Right = 208: .Bottom = 28
'        For cI = 0 To 5
'            DrawFillRectangle TempRect, GetLngColor(vbButtonFace), cHdc
'            .Top = .Top + 18: .Bottom = .Bottom + 18
'            DrawFillRectangle TempRect, GetLngColor(vbButtonFace), cHdc
'            .Top = .Bottom - 1: .Bottom = .Top + 5
'        Next cI
'        .Left = 6: .Top = 23: .Right = 11: .Bottom = 155
'        For cI = 0 To 7
'            DrawFillRectangle TempRect, GetLngColor(vbButtonFace), cHdc
'            .Left = .Left + 21: .Right = .Right + 21
'            DrawFillRectangle TempRect, GetLngColor(vbButtonFace), cHdc
'            .Left = .Right - 1: .Right = .Left + 5
'        Next cI
'    End With
'    ReleaseDC hwnd, cHdc
'    If ICurrentHover > 0 Then DrawXPRectangle hwnd, ICurrentHover
'    If GlobalPos > 0 Then DrawXPRectangle hwnd, GlobalPos
'End Sub
'
'Private Sub DrawXPRectangle(ByVal hwnd As Long, ByVal iNumber As Long)
'    Dim WinItem As RECT
'    Dim cHdc    As Long
'    Dim zColor  As Long
'    Dim zColor2 As Long
'    zColor = ShiftColorOXP(GetLngColor(XPBlue_Highlight), 190)
'    zColor2 = GetLngColor(XPBlue_Highlight)
'    cHdc = GetDC(hwnd)
'    GetXY_Rectangle iNumber
'    With WinItem
'            .Top = 26 + (yN * 22)
'            .Left = 9 + (xN * 25)
'            .Right = .Left + 2: .Bottom = .Top + 17
'            DrawFillRectangle WinItem, zColor, cHdc
'            .Left = .Left + 18: .Right = .Left + 2: .Bottom = .Top + 17
'            DrawFillRectangle WinItem, zColor, cHdc
'            .Left = 9 + (xN * 25): .Right = .Left + 20: .Top = 26 + (yN * 22): .Bottom = .Top + 2
'            DrawFillRectangle WinItem, zColor, cHdc
'            .Left = 9 + (xN * 25): .Right = .Left + 20: .Top = 26 + (yN * 22) + 15: .Bottom = .Top + 2
'            DrawFillRectangle WinItem, zColor, cHdc
'            .Left = 8 + (xN * 25): .Top = 25 + (yN * 22): .Right = .Left + 22: .Bottom = .Top + 19
'            DrawRectangle WinItem, zColor2, cHdc
'    End With
'    ReleaseDC hwnd, cHdc
'End Sub
'
'Private Sub GetXY_Rectangle(ByVal cNumber As Long)
'    Dim i As Long
'    Dim II As Long
'    For II = 0 To 5
'        For i = 1 To 8
'            If (II * 8) + i = cNumber Then yN = II
'        Next i
'    Next II
'    For II = 0 To 7
'        For i = 1 + II To 48 Step 8
'            If i = cNumber Then xN = II
'        Next i
'    Next II
'End Sub
'
'Private Function CheckArea(ByVal hwnd As Long) As Long
'    Dim WinItem As RECT
'    Dim IParent As RECT
'    Dim Point   As POINTAPI
'    Dim hRgn    As Long
'    Dim count   As Long
'    GetCursorPos Point
'    GetWindowRect hwnd, IParent
'    IParent.Left = IParent.Left + GetSystemMetrics(SM_CXDLGFRAME)
'    IParent.Top = IParent.Top + GetSystemMetrics(SM_CYDLGFRAME) + GetSystemMetrics(SM_CYCAPTION)
'    For count = 1 To 48
'        GetXY_Rectangle count
'        With WinItem
'            .Left = IParent.Left + 8 + (xN * 25): .Top = IParent.Top + 25 + (yN * 22): .Right = .Left + 22: .Bottom = .Top + 19
'        End With
'        hRgn = CreateRectRgnIndirect(WinItem)
'        If PtInRegion(hRgn, CLng(Point.X), CLng(Point.Y)) Then
'            DeleteObject hRgn
'            CheckArea = count
'            Exit Function
'        Else
'            DeleteObject hRgn
'        End If
'    Next count
'End Function
