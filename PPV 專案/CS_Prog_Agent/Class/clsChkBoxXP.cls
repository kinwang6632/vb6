VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChkBoxXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_hWnd As Long
Private m_Hdc As Long
Private m_Value As Byte
Private m_Alignment As Byte
Private m_Enabled As Boolean
Private m_Down As Boolean
Private m_Over As Boolean
Private RcItem As RECT

Public Sub DrawCheckBox()

   
    Call GetOriginalRect
 
        If Not m_Enabled Then
             DrawGradientMenu m_Hdc, RcItem.Left, RcItem.Top, RcItem.Right, RcItem.Bottom - RcItem.Top, GetRGBColors(vbWhite), GetRGBColors(vbWhite), GRADIENT_HORIZONTAL
             DrawRectangle RcItem, GetLngColor(&HBBC8CA), m_Hdc
             If m_Value = 2 Then CenterRectangle RcItem, GetLngColor(&HBBC8CA)
             Exit Sub
        End If
    
        If m_Down Then
             DrawGradientMenu m_Hdc, RcItem.Left, RcItem.Top, RcItem.Right, RcItem.Bottom - RcItem.Top, GetRGBColors(GetLngColor(&HA7B0B0)), GetRGBColors(&HDFEFF1), GRADIENT_HORIZONTAL
             DrawRectangle RcItem, GetLngColor(&H80511C), m_Hdc
             If m_Value = 1 Then DrawArrow RcItem, GetLngColor(&H1A8A1C)
             If m_Value = 2 Then CenterRectangle RcItem, GetLngColor(&H1A8A1C)
             Exit Sub
        End If
    
        If m_Over Then
             DrawGradientMenu m_Hdc, RcItem.Left, RcItem.Top, RcItem.Right, RcItem.Bottom - RcItem.Top, GetRGBColors(GetLngColor(&H84D6FD)), GetRGBColors(&H30B3F8), GRADIENT_HORIZONTAL
             DrawRectangle RcItem, GetLngColor(&H80511C), m_Hdc
             CenterRectangle RcItem, GetLngColor(&HE3E7E7)
             Call GetOriginalRect
             If m_Value = 1 Then DrawArrow RcItem, GetLngColor(&H21A121)
             If m_Value = 2 Then CenterRectangle RcItem, GetLngColor(&H21A121)
             Exit Sub
        End If
        
        
             DrawGradientMenu m_Hdc, RcItem.Left, RcItem.Top, RcItem.Right, RcItem.Bottom - RcItem.Top, GetRGBColors(GetLngColor(&HD7DCDC)), GetRGBColors(vbWhite), GRADIENT_HORIZONTAL
             DrawRectangle RcItem, GetLngColor(&H80511C), m_Hdc
             If m_Value = 1 Then DrawArrow RcItem, GetLngColor(&H21A121)
             If m_Value = 2 Then CenterRectangle RcItem, GetLngColor(&H72C173)
        
        
              
 



End Sub

Private Sub GetOriginalRect()
  
  GetClientRect m_hWnd, RcItem
  RcItem.Top = (RcItem.Bottom - 13) / 2
  RcItem.Bottom = RcItem.Top + 13
  
  If m_Alignment = 0 Then
     RcItem.Right = RcItem.Left + 13
  Else
    RcItem.Left = RcItem.Right - 13
  End If
  
End Sub

Private Sub CenterRectangle(ByRef BRect As RECT, ByVal Color As Long)
Dim hBrush As Long
    
    hBrush = CreateSolidBrush(Color)
    InflateRect BRect, -3, -3
    FillRect m_Hdc, BRect, hBrush
    DeleteObject hBrush

End Sub

Private Sub DrawArrow(ByRef BRect As RECT, ByVal cColor As Long)

Dim J1 As Integer
Dim Pen1 As Long, Pen2 As Long
Dim POS As POINTAPI
       
    Pen1 = CreatePen(0, 1, cColor)
    Pen2 = SelectObject(m_Hdc, Pen1)
    InflateRect BRect, -3, -3
    
    For J1 = 0 To 2
       MoveToEx m_Hdc, BRect.Left + J1, BRect.Top + J1 + 2, POS
       LineTo m_Hdc, BRect.Left + J1, BRect.Top + J1 + 5
    Next J1
            
    For J1 = 3 To 6
       MoveToEx m_Hdc, BRect.Left + J1, BRect.Top + Abs(-6 + J1), POS
       LineTo m_Hdc, BRect.Left + J1, BRect.Top + 9 - J1
    Next J1
            
    SelectObject m_Hdc, Pen2
    DeleteObject Pen2
    DeleteObject Pen1
  
End Sub

Public Property Let Over(ByVal cOver As Boolean)
   m_Over = cOver
End Property

Public Property Let Down(ByVal cDown As Boolean)
   m_Down = cDown
End Property

Public Property Let Enabled(ByVal cEnabled As Boolean)
   m_Enabled = cEnabled
End Property

Public Property Let Alignment(ByVal cAlignment As Byte)
   m_Alignment = cAlignment
End Property

Public Property Let Value(ByVal cValue As Byte)
   m_Value = cValue
End Property

Public Property Let hwnd(ByVal cHwnd As Long)
   m_hWnd = cHwnd
End Property

Public Property Let hdc(ByVal cHdc As Long)
   m_Hdc = cHdc
End Property
