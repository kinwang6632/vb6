Attribute VB_Name = "mod_DrawTab"
' PR : Hammer
' Start date : 2004/12/23
' Last Modify : 2004/12/23

Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type GRADIENT_RECT
    UPPERLEFT  As Long
    LOWERRIGHT As Long
End Type

Private Type TRIVERTEX
    X        As Long
    Y        As Long
    Red     As Integer
    Green  As Integer
    Blue    As Integer
    Alpha  As Integer
End Type

Private Type RGB
    R As Integer
    G As Integer
    B As Integer
End Type

Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
'Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ValidateRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As Long) As Long

Private Const GWL_WNDPROC As Long = (-4)
Private Const WM_PAINT    As Long = &HF
Private Const WM_DESTROY  As Long = &H2
Private Const WM_TIMER    As Long = &H113
Private Const ID_TIMER    As Long = &HCBABE
                         
Public Enum TabStyle
    cSolidColor = 0
    cPicture = 1
    cGradient = 2
    cAnimatedGradient = 3
End Enum

Public Enum Direction
    cHorizontal = 0
    cVertical = 1
End Enum

Private DestDC As Long
Private MaskDC As Long
Private MemDC As Long
Private OrigDC As Long
Private MaskPic As Long
Private MemPic As Long
Private TempPic As Long
Private OrigPic As Long
Private TempDC As Long

Private origBrush As Long
Private TempBrush As Long
Private origColor As Long
Private gColor1 As Long
Private gColor2 As Long
Private gDir As Long
Private gTime As Long
Private gFadeFlag As Boolean
Private ImageWidth  As Long
Private ImageHeight As Long
Public oldWndProc As Long

Private Function GetLngColor(lngColor As Long) As Long
  On Error GoTo ChkErr
    If (lngColor And &H80000000) Then
        GetLngColor = GetSysColor(lngColor And &H7FFFFFFF)
    Else
        GetLngColor = lngColor
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_DrawTab", "Function : GetLngColor"
End Function

Private Function GetRGBColors(Color As Long) As RGB
  On Error GoTo ChkErr
    Dim strHexColor As String
    strHexColor = String(6 - Len(Hex(Color)), "0") & Hex(Color)
    GetRGBColors.R = "&H" & Mid(strHexColor, 5, 2) & "00"
    GetRGBColors.G = "&H" & Mid(strHexColor, 3, 2) & "00"
    GetRGBColors.B = "&H" & Mid(strHexColor, 1, 2) & "00"
  Exit Function
ChkErr:
    ErrHandle "mod_DrawTab", "Function : GetRGBColors"
End Function

Public Sub SetStyle(ByVal hWnd As Long, ByRef Style As TabStyle)
  On Error GoTo ChkErr
    SetProp hWnd, "MyStyle", Style
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : SetStyle"
End Sub

Public Sub SetFadeTime(ByVal hWnd As Long, ByVal cTime As Long)
  On Error GoTo ChkErr
    If cTime > 10 Then cTime = 10
    If cTime < 1 Then cTime = 1
    SetProp hWnd, "MyFadeTime", cTime
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : SetFadeTime"
End Sub

Private Function GetFadeTime(ByVal hWnd As Long) As Long
  On Error GoTo ChkErr
    GetFadeTime = GetProp(hWnd, "MyFadeTime")
  Exit Function
ChkErr:
    ErrHandle "mod_DrawTab", "Function : SetFadeTime"
End Function

Private Function GetStyleParams(ByVal hWnd As Long) As TabStyle
  On Error GoTo ChkErr
    GetStyleParams = GetProp(hWnd, "MyStyle")
  Exit Function
ChkErr:
    ErrHandle "mod_DrawTab", "Function : GetStyleParams"
End Function

Public Sub SetGradientDir(ByVal hWnd As Long, ByRef Style As Direction)
  On Error GoTo ChkErr
    SetProp hWnd, "MyGradientDir", Style
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Function : SetGradientDir"
End Sub

Private Sub GetGradientDir(ByVal hWnd As Long)
  On Error GoTo ChkErr
    gDir = GetProp(hWnd, "MyGradientDir")
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : GetGradientDir"
End Sub

Private Sub SetHookInstance(ByVal hWnd As Long)
  On Error GoTo ChkErr
    SetProp hWnd, "Hooked", True
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : SetHookInstance"
End Sub

Private Function CheckHookInstance(ByVal hWnd As Long) As Boolean
  On Error GoTo ChkErr
    CheckHookInstance = GetProp(hWnd, "Hooked")
  Exit Function
ChkErr:
    ErrHandle "mod_DrawTab", "Function : CheckHookInstance"
End Function

Public Sub SetSolidColor(ByVal hWnd As Long, ByVal Color As Long)
  On Error GoTo ChkErr
    SetProp hWnd, "MySolidColor", GetLngColor(Color)
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : SetSolidColor"
End Sub

Public Sub SetGradientColor1(ByVal hWnd As Long, ByVal Color As Long)
  On Error GoTo ChkErr
    SetProp hWnd, "MyGradientColor1", GetLngColor(Color)
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : SetGradientColor1"
End Sub

Public Sub SetGradientColor2(ByVal hWnd As Long, ByVal Color As Long)
  On Error GoTo ChkErr
    SetProp hWnd, "MyGradientColor2", GetLngColor(Color)
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : SetGradientColor2"
End Sub

Private Sub GetSolidColor(ByVal hWnd As Long)
  On Error GoTo ChkErr
     TempBrush = CreateSolidBrush(GetProp(hWnd, "MySolidColor"))
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : GetSolidColor"
End Sub

Private Sub GetGradientColor1(ByVal hWnd As Long)
  On Error GoTo ChkErr
    gColor1 = GetProp(hWnd, "MyGradientColor1")
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : GetGradientColor1"
End Sub

Private Sub GetGradientColor2(ByVal hWnd As Long)
  On Error GoTo ChkErr
     gColor2 = GetProp(hWnd, "MyGradientColor2")
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : GetGradientColor2"
End Sub

Public Sub SetPicture(ByVal hWnd As Long, ByVal Width As Long, ByVal Height As Long, ByRef cPicture As StdPicture)
  On Error GoTo ChkErr
    SetProp hWnd, "MyPicture", cPicture.Handle
    SetProp hWnd, "MyPictureWidth", Width
    SetProp hWnd, "MyPictureHeight", Height
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : SetPicture"
End Sub

Private Sub GetPictureParams(ByVal hWnd As Long)
  On Error GoTo ChkErr
    TempBrush = CreatePatternBrush(GetProp(hWnd, "MyPicture"))
    ImageWidth = GetProp(hWnd, "MyPictureWidth")
    ImageHeight = GetProp(hWnd, "MyPictureHeight")
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : GetPictureParams"
End Sub

Public Sub SSTabSubclass(ByVal hWnd As Long)
  On Error GoTo ChkErr
    If Not CheckHookInstance(hWnd) Then
        SetHookInstance hWnd
        oldWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf oldSSTabProc)
    End If
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : SSTabSubclass"
End Sub

Public Function oldSSTabProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  On Error GoTo ChkErr
    If GetStyleParams(hWnd) = cAnimatedGradient Then
       KillTimer hWnd, 0
       SetTimer hWnd, ID_TIMER, 1, 0
    End If
    oldSSTabProc = NewSSTabProc(hWnd, uMsg, wParam, lParam)
  Exit Function
ChkErr:
    ErrHandle "mod_DrawTab", "Function : oldSSTabProc"
End Function

Private Function NewSSTabProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  On Error Resume Next
    Dim m_ItemRect As RECT
    Dim m_Width    As Long
    Dim m_Height   As Long
    If wMsg = WM_PAINT Then
        DestDC = GetDC(hWnd)
        GetWindowRect hWnd, m_ItemRect
        m_Width = m_ItemRect.Right - m_ItemRect.Left
        m_Height = m_ItemRect.Bottom - m_ItemRect.Top
        Select Case GetStyleParams(hWnd)
                Case cPicture
                        GetPictureParams hWnd
                Case cSolidColor
                        GetSolidColor hWnd
                Case cGradient
                        GetGradientColor1 hWnd
                        GetGradientColor2 hWnd
                        GetGradientDir hWnd
                Case cAnimatedGradient
                        GetGradientDir hWnd
                Case Else
                        Debug.Print "Invalid Style"
        End Select
        CreateNewDCWorkArea m_Width, m_Height
        Call SelectBitmap
        CallWindowProc oldWndProc, hWnd, wMsg, OrigDC, lParam
        Call CreateBackMask(m_Width, m_Height)
        origBrush = SelectObject(TempDC, TempBrush)
        If GetStyleParams(hWnd) = cGradient Or GetStyleParams(hWnd) = cAnimatedGradient Then
            DrawGradient TempDC, 0, 0, m_Width, m_Height, GetRGBColors(gColor1), GetRGBColors(gColor2), gDir
        Else
            PatBlt TempDC, 0, 0, m_Width, m_Height, vbPatCopy
        End If
        SelectObject TempDC, origBrush
        Call DOBitBlt(m_Width, m_Height)
        Call CleanDCs
        SetBkColor DestDC, origColor
        ReleaseDC hWnd, DestDC
        ValidateRect hWnd, 0
    ElseIf wMsg = WM_TIMER Then
        If GetStyleParams(hWnd) <> cAnimatedGradient Then
            KillTimer hWnd, 0
            Exit Function
        End If
        If gFadeFlag Then
            gTime = gTime - GetFadeTime(hWnd)
        Else
            gTime = gTime + GetFadeTime(hWnd)
        End If
        If gTime > 255 Then
           gTime = 255
           gFadeFlag = Not gFadeFlag
        ElseIf gTime < 0 Then
           gTime = 0
           gFadeFlag = Not gFadeFlag
        End If
        GetGradientColor1 hWnd
        GetGradientColor2 hWnd
        gColor1 = ShiftColor(gColor1, gTime)
        gColor2 = ShiftColor(gColor2, gTime)
        RedrawWindow hWnd, ByVal 0&, ByVal 0&, &H1
'        Debug.Print gTime
    ElseIf wMsg = WM_DESTROY Then
        KillTimer hWnd, 0
        DeleteObject TempBrush
        SetWindowLong hWnd, GWL_WNDPROC, oldWndProc
        NewSSTabProc = CallWindowProc(oldWndProc, hWnd, wMsg, wParam, lParam)
    Else
        NewSSTabProc = CallWindowProc(oldWndProc, hWnd, wMsg, wParam, lParam)
    End If
End Function

Private Sub SelectBitmap()
  On Error GoTo ChkErr
    Dim cHandle As Long
    cHandle = SelectObject(MaskDC, MaskPic)
    DeleteObject cHandle
    cHandle = SelectObject(MemDC, MemPic)
    DeleteObject cHandle
    cHandle = SelectObject(TempDC, TempPic)
    DeleteObject cHandle
    cHandle = SelectObject(OrigDC, OrigPic)
    DeleteObject cHandle
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : SelectBitmap"
End Sub

Private Sub CreateBackMask(ByVal m_Width As Long, ByVal m_Height As Long)
  On Error GoTo ChkErr
    origColor = SetBkColor(DestDC, GetSysColor(15))
    SetBkColor OrigDC, GetSysColor(15)
    BitBlt MaskDC, 0, 0, m_Width, m_Height, OrigDC, 0, 0, vbSrcCopy
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : CreateBackMask"
End Sub

Private Sub CreateNewDCWorkArea(ByVal m_Width As Long, ByVal m_Height As Long)
  On Error GoTo ChkErr
    MaskDC = CreateCompatibleDC(DestDC)
    MaskPic = CreateBitmap(m_Width, m_Height, 1, 1, ByVal 0&)
    MemDC = CreateCompatibleDC(DestDC)
    MemPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)
    TempDC = CreateCompatibleDC(DestDC)
    TempPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)
    OrigDC = CreateCompatibleDC(DestDC)
    OrigPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : CreateNewDCWorkArea"
End Sub

Private Sub DOBitBlt(ByVal m_Width As Long, ByVal m_Height As Long)
  On Error GoTo ChkErr
    BitBlt MemDC, 0, 0, m_Width, m_Height, MaskDC, 0, 0, vbSrcCopy
    BitBlt MemDC, 0, 0, m_Width, m_Height, OrigDC, 0, 0, vbSrcPaint
    BitBlt TempDC, 0, 0, m_Width, m_Height, MaskDC, 0, 0, vbMergePaint
    BitBlt TempDC, 0, 0, m_Width, m_Height, MemDC, 0, 0, vbSrcAnd
    BitBlt DestDC, 0, 0, m_Width, m_Height, TempDC, 0, 0, vbSrcCopy
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : DOBitBlt"
End Sub

Private Sub CleanDCs()
  On Error GoTo ChkErr
    DeleteDC TempDC
    DeleteObject TempPic
    DeleteDC MaskDC
    DeleteObject MaskPic
    DeleteDC MemDC
    DeleteObject MemPic
    DeleteDC OrigDC
    DeleteObject OrigPic
    DeleteObject TempBrush
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : CleanDCs"
End Sub

Private Sub DrawGradient(cHdc As Long, X As Long, Y As Long, X2 As Long, Y2 As Long, Color1 As RGB, Color2 As RGB, Optional Direction = 1)
  On Error GoTo ChkErr
    Dim Vert(1) As TRIVERTEX
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
    gRect.UPPERLEFT = 0
    gRect.LOWERRIGHT = 1
    GradientFillRect cHdc, Vert(0), 2, gRect, 1, Direction
  Exit Sub
ChkErr:
    ErrHandle "mod_DrawTab", "Subroutine : DrawGradient"
End Sub

Private Function ShiftColor(ByVal lngColor As Long, ByVal lngValue As Long) As Long
  On Error GoTo ChkErr
    Dim lngR As Long
    Dim lngG As Long
    Dim lngB As Long
    lngR = (lngColor And &HFF) + lngValue
    lngG = ((lngColor \ &H100) Mod &H100) + lngValue
    lngB = ((lngColor \ &H10000) Mod &H100)
    lngB = lngB + ((lngB * lngValue) \ &HC0)
    If lngValue > 0 Then
        If lngR > 255 Then lngR = 255
        If lngG > 255 Then lngG = 255
        If lngB > 255 Then lngB = 255
    ElseIf lngValue < 0 Then
        If lngR < 0 Then lngR = 0
        If lngG < 0 Then lngG = 0
        If lngB < 0 Then lngB = 0
    End If
    ShiftColor = lngR + 256& * lngG + 65536 * lngB
  Exit Function
ChkErr:
    ErrHandle "mod_DrawTab", "Function : ShiftColor"
End Function




