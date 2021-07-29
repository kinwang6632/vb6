Attribute VB_Name = "mod_API_Declare"
' PR : Hammer
' Start date : 2004/12/24
' Last Modify : 2004/12/25

Option Explicit

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ChildWindowFromPoint Lib "user32.dll" (ByVal hwnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Public Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Public Declare Function ImageList_Draw Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Public Declare Function ImageList_GetIconSize Lib "COMCTL32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
'Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Icount As Long
Public MsgBoxIClass() As clsWinXPengine

Public Type RECT
    Left      As Long
    Top       As Long
    Right     As Long
    Bottom    As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hwnd As Long
End Type

Public Type DRAWITEMSTRUCT
   CtlType     As Long
   CtlID       As Long       '     //---The DRAWITEMSTRUCT structure provides information the
   itemID      As Long       '     //---owner window must have to determine how to paint an
   itemAction  As Long       '     //---owner-drawn control or menu item.
   itemState   As Long       '     //---The owner window of the owner-drawn control or menu item
   hWndItem    As Long       '     //---receives a pointer to this structure as the lParam parameter
   hdc         As Long       '     //--- of the WM_DRAWITEM message.
   RcItem      As RECT
   itemData    As Long
End Type

Public Type RGBTRIPLE
    rgbBlue  As Byte
    rgbGreen As Byte
    rgbRed   As Byte
End Type

Public Type HDHITTESTINFO
    pt    As POINTAPI        '    //--Contains information about a hit test.
    flags As Long            '    //--This structure is used with the HDM_HITTEST message
    iItem As Long                                 '(HEADERS)
End Type

Public Type TCHITTESTINFO
    pt     As POINTAPI       '    //Contains information about a hit test.
    flags  As Long                                 ' (TABS)
End Type
 
Public Type HDITEM
    mask        As Long
    cxy         As Long
    pszText     As String
    hbm         As Long       '   //--Contains information about an item in a header control.
    cchTextMax  As Long
    fmt         As Long
    lParam      As Long
    iImage      As Long
    iOrder      As Long
End Type

 Public Type TCITEM
    mask        As Long
    dwState     As Long
    dwStateMask As Long       '   //--Specifies or receives the attributes of a tab item.
    pszText     As String     '   //--It is used with the TCM_INSERTITEM, TCM_GETITEM, and TCM_SETITEM messages.
    cchTextMax  As Long
    iImage      As Long
    lParam      As Long
 End Type

Public Type TLoHiLong
    Lo As Integer
    Hi As Integer
End Type

Public Type TAllLong
    All As Long
End Type

Public Type NMHDR
    hwndFrom As Long          ' Window handle of control sending message
    idFrom   As Long          ' Identifier of control sending message
    code    As Long           ' Specifies the notification code
End Type

'Public Type NMCUSTOMDRAWINFO
'    hdr As NMHDR
'    dwDrawStage As Long
'    hdc As Long
'    rc As RECT
'    dwItemSpec As Long       'Code not Ready
'    iItemState As Long
'    lItemLParam As Long
'End Type
'
'Public Type NMLVCUSTOMDRAW
'    nmcmd           As NMCUSTOMDRAWINFO
'    clrText As Long
'    clrTextBk As Long
'    'iSubItem As Integer          'Code not Ready
'End Type

Public Type LOGFONT
    lfHeight          As Long
    lfWidth           As Long
    lfEscapement      As Long
    lfOrientation     As Long
    lfWeight          As Long
    lfItalic          As Byte
    lfUnderline       As Byte     '   //--The LOGFONT structure defines the attributes of a font.
    lfStrikeOut       As Byte
    lfCharSet         As Byte
    lfOutPrecision    As Byte
    lfClipPrecision   As Byte
    lfQuality         As Byte
    lfPitchAndFamily  As Byte
    lfFaceName        As String * 32
End Type

Public Type TRIVERTEX
    X       As Long
    Y       As Long
    Red     As Integer       '   //--The TRIVERTEX structure contains color information and position information.
    Green   As Integer
    Blue    As Integer
    Alpha   As Integer
End Type

Public Type GRADIENT_RECT
    UPPERLEFT  As Long       '  //--The GRADIENT_RECT structure specifies the index of two vertices in the pVertex array.
    LOWERRIGHT As Long       '      These two vertices form the upper-left and lower-right boundaries of a rectangle.
End Type

Public Type RGB
    R As Integer
    G As Integer             '  //--Selects a red, green, blue (RGB) color based on the arguments supplied
    B As Integer
End Type

Public Enum ControlState
    C_Normal = 0
    C_Over = 1
    C_Focus = 2
    C_Down = 3
    C_Disabled = 4
    C_Up = 5
End Enum

Public Enum CWindowColors
    SystemColors = 1
    WindowsXP_Blue = 2
    WindowsXP_OliveGreen = 3   '//-- Soon
    WindowsXP_Silver = 4
End Enum

