Attribute VB_Name = "mod_Func"
' PR : Hammer
' Start date : 2004/12/24
' Last Modify : 2004/12/25

Option Explicit

Public Function LongInt2Int(ByVal lLongInt As Long, ByRef iHiWord As Integer, ByRef iLowWord As Integer) As Boolean
    Dim tmpHW As Integer, tmpLW As Integer
    CopyMemory tmpLW, lLongInt, Len(tmpLW)
    tmpHW = (lLongInt / TwoPower16)
    iHiWord = tmpHW
    iLowWord = tmpLW
End Function

Public Function GetLngColor(Color As Long) As Long
    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
End Function

Public Function GetRGBColors(Color As Long) As RGB
    Dim HexColor As String
    HexColor = String(6 - Len(Hex(Color)), "0") & Hex(Color)
    GetRGBColors.R = "&H" & Mid(HexColor, 5, 2) & "00"
    GetRGBColors.G = "&H" & Mid(HexColor, 3, 2) & "00"
    GetRGBColors.B = "&H" & Mid(HexColor, 1, 2) & "00"
End Function

Public Function ShiftColor(ByVal Color As Long, ByVal Value As Long, Optional isXP As Boolean = False) As Long
    Dim R As Long
    Dim G As Long
    Dim B As Long
    R = (Color And &HFF) + Value
    G = ((Color \ &H100) Mod &H100) + Value
    B = ((Color \ &H10000) Mod &H100)
    B = B + ((B * Value) \ &HC0)
    If Value > 0 Then
        If R > 255 Then R = 255
        If G > 255 Then G = 255
        If B > 255 Then B = 255
    ElseIf Value < 0 Then
        If R < 0 Then R = 0
        If G < 0 Then G = 0
        If B < 0 Then B = 0
    End If
    ShiftColor = R + 256& * G + 65536 * B
End Function

Public Sub SelectFont(cHdc As Long, Size As Integer, Italic As Boolean, FontName As String, Underline As Boolean)
    Dim MyFont As LOGFONT
    Dim NewFont As Long
    With MyFont
        .lfHeight = (Size * -20) / Screen.TwipsPerPixelY
        .lfCharSet = 1
        .lfItalic = Italic
        .lfUnderline = Underline
        .lfFaceName = FontName & Chr$(0)
    End With
    NewFont = CreateFontIndirect(MyFont)
    SelectObject cHdc, NewFont
    DeleteObject NewFont
End Sub

Public Sub DrawLine(X, Y, Width, Height, cHdc, Color As Long)
    Dim Pen1 As Long, Pen2 As Long, Outline As Long
    Dim POS As POINTAPI
    Pen1 = CreatePen(0, 1, GetLngColor(Color))
    Pen2 = SelectObject(cHdc, Pen1)
    MoveToEx cHdc, X, Y, POS
    LineTo cHdc, Width, Height
    SelectObject cHdc, Pen2
    DeleteObject Pen2
    DeleteObject Pen1
End Sub

Public Sub DrawGradientMenu(cHdc As Long, X As Long, Y As Long, X2 As Long, Y2 As Long, Color1 As RGB, Color2 As RGB, Optional Direction = 1)
    Dim Vert(1) As TRIVERTEX   '2 Colors
    Dim gRect As GRADIENT_RECT
    With Vert(0)
        .X = X
        .Y = Y
        .Red = Color1.R
        .Green = Color1.G
        .Blue = Color1.B
        .Alpha = 0&
    End With
    With Vert(1)
        .X = Vert(0).X + X2
        .Y = Vert(0).Y + Y2
        .Red = Color2.R
        .Green = Color2.G
        .Blue = Color2.B
        .Alpha = 0&
    End With
    gRect.UPPERLEFT = 1
    gRect.LOWERRIGHT = 0
    GradientFillRect cHdc, Vert(0), 2, gRect, 1, Direction
End Sub

Public Sub DrawRectangle(ByRef BRect As RECT, ByVal Color As Long, ByVal hdc As Long)
    Dim hBrush As Long
    hBrush = CreateSolidBrush(Color)
    FrameRect hdc, BRect, hBrush
    DeleteObject hBrush
End Sub

Public Sub DrawFillRectangle(ByRef hRect As RECT, ByVal Color As Long, ByVal MyHdc As Long)
    Dim hBrush As Long
    hBrush = CreateSolidBrush(GetLngColor(Color))
    FillRect MyHdc, hRect, hBrush
    DeleteObject hBrush
End Sub

Public Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
    Dim Red As Long, Blue As Long, Green As Long
    Dim Delta As Long
    Blue = ((theColor \ &H10000) Mod &H100)
    Green = ((theColor \ &H100) Mod &H100)
    Red = (theColor And &HFF)
    Delta = &HFF - Base
    Blue = Base + Blue * Delta \ &HFF
    Green = Base + Green * Delta \ &HFF
    Red = Base + Red * Delta \ &HFF
    If Red > 255 Then Red = 255
    If Green > 255 Then Green = 255
    If Blue > 255 Then Blue = 255
    ShiftColorOXP = Red + 256& * Green + 65536 * Blue
End Function

Public Sub MakeRegion(ByRef RcItem As RECT, ByVal m_hWnd As Long)
    Dim rgn1 As Long, rgn2 As Long, rgnNorm As Long
    rgnNorm = CreateRectRgn(0, 0, RcItem.Right, RcItem.Bottom)
    rgn2 = CreateRectRgn(0, 0, 0, 0)
    rgn1 = CreateRectRgn(0, 0, 2, 1)
    CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
    DeleteObject rgn1
    rgn1 = CreateRectRgn(0, RcItem.Bottom, 2, RcItem.Bottom - 1)
    CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
    DeleteObject rgn1
    rgn1 = CreateRectRgn(RcItem.Right, 0, RcItem.Right - 2, 1)
    CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
    DeleteObject rgn1
    rgn1 = CreateRectRgn(RcItem.Right, RcItem.Bottom, RcItem.Right - 2, RcItem.Bottom - 1)
    CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
    DeleteObject rgn1
    rgn1 = CreateRectRgn(0, 1, 1, 2)
    CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
    DeleteObject rgn1
    rgn1 = CreateRectRgn(0, RcItem.Bottom - 1, 1, RcItem.Bottom - 2)
    CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
    DeleteObject rgn1
    rgn1 = CreateRectRgn(RcItem.Right, 1, RcItem.Right - 1, 2)
    CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
    DeleteObject rgn1
    rgn1 = CreateRectRgn(RcItem.Right, RcItem.Bottom - 1, RcItem.Right - 1, RcItem.Bottom - 2)
    CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
    DeleteObject rgn1
    DeleteObject rgn2
    SetWindowRgn m_hWnd, rgnNorm, True
    DeleteObject rgnNorm
End Sub

Public Function GetObjectText(ByVal m_hWnd As Long) As String
    GetObjectText = String(GetWindowTextLength(m_hWnd) + 1, Chr$(0))
    GetWindowText m_hWnd, GetObjectText, Len(GetObjectText)
End Function

Public Sub SetCurrentObjectFont(ByVal m_hWnd As Long, ByVal m_Hdc As Long)
    Dim Current As Long
    Current = SendMessageLong(m_hWnd, WM_GETFONT, 0&, 0&)
    Current = SelectObject(m_Hdc, Current)
    DeleteObject Current
End Sub
 
Public Function ThisWindowClassName(ByVal m_hWnd As Long) As String
    Dim RetVal As Long, lpClassName As String
    lpClassName = Space(255)
    RetVal = GetClassName(m_hWnd, lpClassName, 255)
    ThisWindowClassName = Left$(lpClassName, RetVal)
End Function

'Public Function GetTopLevel(ByVal hChild As Long) As Long
'   Dim hwnd As Long
'   ' Read parent chain up to highest visible.
'   hwnd = hChild
'   Do While IsWindowVisible(GetParent(hwnd))
'      hwnd = GetParent(hChild)
'      hChild = hwnd
'   Loop
'   GetTopLevel = hwnd
'End Function

Public Function CleanOptionButtomArray(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Select Case ThisWindowClassName(hwnd)
            Case "ThunderOptionButton"
                    RedrawWindow hwnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event)
    End Select
    CleanOptionButtomArray = 1
End Function

Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim SubclassThis As Boolean
    SubclassThis = False
    Select Case ThisWindowClassName(hwnd)
            Case "SysListView32"
                    SubclassThis = True
            Case "ComboBox"
                    SubclassThis = True
            Case "Edit"
                    SubclassThis = True
            Case "Button"
                    SubclassThis = True
            Case Else
                    'Nothing
    End Select
    Dim SchemeVal As Byte
    SchemeVal = GetProp(GetParent(GetParent(hwnd)), "ColorScheme")
    If (SubclassThis) Then
        Icount = Icount + 1
        ReDim Preserve MsgBoxIClass(1 To Icount) As clsWinXPengine
        Set MsgBoxIClass(Icount) = New clsWinXPengine
        MsgBoxIClass(Icount).SchemeColor = IIf(SchemeVal <> 0, SchemeVal, WindowsXP_Blue) 'SystemColors
        MsgBoxIClass(Icount).AfterAttachMessageBox hwnd
    End If
    EnumChildProc = 1
End Function

Public Function WindowAlignment(ByVal hwnd As Long) As Byte
    Dim lStyle As Long
    lStyle = GetWindowLong(hwnd, GWL_STYLE)
    WindowAlignment = IIf(lStyle And BS_LEFTTEXT, 1, 0)
End Function

Public Function InsideArea(cHandle As Long) As Boolean
    Dim POS As POINTAPI
    GetCursorPos POS
    If (WindowFromPoint(POS.X, POS.Y) <> cHandle) Then
        InsideArea = False
    Else 'NOT (WINDOWFROMPOINT(POS.X,...
        InsideArea = True
    End If
End Function

Public Sub MakeWindowFlat(ByVal hwnd As Long)
    Dim TFlat As Long
    TFlat = GetWindowLong(hwnd, GWL_STYLE)
    TFlat = TFlat And Not WS_BORDER
    SetWindowLong hwnd, GWL_STYLE, TFlat
    TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
    TFlat = TFlat And Not WS_EX_CLIENTEDGE 'Or WS_EX_STATICEDGE
    SetWindowLong hwnd, GWL_EXSTYLE, TFlat
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub


