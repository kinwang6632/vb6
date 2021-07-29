VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComboXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_hWnd        As Long
Private m_Hdc         As Long
Private m_State       As ControlState
Private m_ColorScheme As CWindowColors
Private RcItem        As RECT

Private ActualHighlight As Long
Private ActualGrayText As Long

'===============================================================================================
'Theme Support Colors...
Private TCol1 As Long, TCol2 As Long, TCol3 As Long, TCol4 As Long, TCol5 As Long, TCol6 As Long


Public Sub DrawComboBox()
Dim xColor As Long
Dim ycolor As Long
Dim FColor As Long '//--Frame  Color
Dim BColor As Long '//--Border Color
Dim DColor As Long '//--Arrow  Color
Dim lhWnd As Long
Dim Half As Long
    
 Call SchemeControl '//--Select Colors
 
 
lhWnd = FindWindowEx(m_hWnd, 0&, "Edit", ByVal 0&) '//--- Find Edit Inside control
GetClientRect m_hWnd, RcItem

If lhWnd <> 0 Then
    MoveWindow lhWnd, RcItem.Left + 2, RcItem.Top + 2, RcItem.Right - 19, RcItem.Bottom - 4, 0
Else
    DrawLine RcItem.Right - 19, 1, RcItem.Right - 19, RcItem.Bottom - 2, m_Hdc, vbWhite
    DrawLine RcItem.Right - 18, 1, RcItem.Right - 18, RcItem.Bottom - 2, m_Hdc, vbWhite
End If
                     
                  
    Select Case m_State
      
        Case C_Normal, C_Focus, C_Up
                                 
                                 xColor = ShiftColorOXP(ActualHighlight, 195)
                                 ycolor = ShiftColorOXP(ActualHighlight, 165)
                                 FColor = ShiftColorOXP(ActualHighlight, 150)
                                 BColor = ShiftColorOXP(ActualHighlight, 80)
                                 DColor = IIf(m_ColorScheme = WindowsXP_OliveGreen, vbWhite, IIf(m_ColorScheme = WindowsXP_Silver, vbBlack, &H85614D))
                                 If m_ColorScheme = WindowsXP_OliveGreen Then TCol1 = &HBAD7CB: TCol2 = &HA2C7B6: TCol3 = &H99BBAB: TCol4 = &HA1C4B6: TCol5 = &H8DB3A2: TCol6 = &H7CAD9B
                                 If m_ColorScheme = WindowsXP_Silver Then TCol1 = vbWhite: TCol2 = &HDBCDCC: TCol3 = &HDACCCB: TCol4 = &HE4DAD9
        Case C_Over
                                 xColor = ShiftColorOXP(ActualHighlight, 225)
                                 ycolor = ShiftColorOXP(ActualHighlight, 195)
                                 FColor = ShiftColorOXP(ActualHighlight, 120)
                                 BColor = ShiftColorOXP(ActualHighlight, 80)
                                 DColor = IIf(m_ColorScheme = WindowsXP_OliveGreen, vbWhite, IIf(m_ColorScheme = WindowsXP_Silver, vbBlack, &H85614D))
                                 If m_ColorScheme = WindowsXP_OliveGreen Then TCol1 = &HB9E8DA: TCol2 = TCol1: TCol3 = &HAEE0D1: TCol4 = &HAAD5C8: TCol5 = TCol4: TCol6 = &H9BD3C6
                                 If m_ColorScheme = WindowsXP_Silver Then TCol1 = vbWhite: TCol2 = TCol1: TCol3 = &HDACCCB: TCol4 = &HE4DAD9

        Case C_Down
                                 xColor = ShiftColorOXP(ActualHighlight, 100)
                                 ycolor = ShiftColorOXP(ActualHighlight, 190)
                                 FColor = ShiftColorOXP(ActualHighlight, 170)
                                 BColor = ShiftColorOXP(ActualHighlight, 80)
                                 DColor = IIf(m_ColorScheme = WindowsXP_OliveGreen, vbWhite, IIf(m_ColorScheme = WindowsXP_Silver, vbBlack, &H85614D))
                                 If m_ColorScheme = WindowsXP_OliveGreen Then TCol1 = &H80AA98: TCol2 = TCol1: TCol3 = TCol1: TCol4 = TCol1: TCol5 = &H82AD9B: TCol6 = &H72AA95
                                 If m_ColorScheme = WindowsXP_Silver Then TCol1 = &HDBC2BF: TCol2 = vbWhite: TCol3 = vbWhite: TCol4 = &HCDB5A0

        Case C_Disabled
                                 xColor = &HECF1F1
                                 ycolor = &HDEE7E7
                                 FColor = ShiftColorOXP(ActualGrayText)
                                 BColor = ShiftColorOXP(ActualGrayText, 80)
                                 DColor = GetLngColor(&HC2C9C9)
       
       Case Else
                'Exit Sub

     End Select
                             
 
 If m_ColorScheme = WindowsXP_Silver Then BColor = GetLngColor(XPSilver_Highlight)
 
 Call DrawMenuRectangle(1, 1, RcItem.Right - 1, RcItem.Bottom - 1, vbWhite, vbWhite, True)
 Call DrawMenuRectangle(0, 0, RcItem.Right, RcItem.Bottom, GetLngColor(BColor), GetLngColor(BColor), True)
 
 If m_ColorScheme = SystemColors Or m_ColorScheme = WindowsXP_Blue Or m_State = C_Disabled Then
'============================================================================================================
'XP BLUE SCHEME..
    Call DrawGradientMenu(m_Hdc, RcItem.Right - 17, 2, 15, RcItem.Bottom - 4, GetRGBColors(GetLngColor(xColor)), GetRGBColors(GetLngColor(ycolor)), GRADIENT_HORIZONTAL)
 ElseIf m_ColorScheme = WindowsXP_OliveGreen Then
'============================================================================================================
'XP OLIVE GREEN SCHEME..
    Half = (RcItem.Bottom - 12) / 2
    
    FColor = &H7D998E
    DrawLine RcItem.Right - 17, 3, RcItem.Right - 2, 3, m_Hdc, TCol1
    DrawLine RcItem.Right - 16, 4, RcItem.Right - 16, Half + 4, m_Hdc, TCol1
    DrawLine RcItem.Right - 16, Half + 4, RcItem.Right - 16, Half + 6 + (Half / 2), m_Hdc, TCol2
    DrawLine RcItem.Right - 16, Half + 6 + (Half / 2), RcItem.Right - 16, RcItem.Bottom - 3, m_Hdc, TCol3
    DrawLine RcItem.Right - 4, 4, RcItem.Right - 4, Half + 4, m_Hdc, TCol1
    DrawLine RcItem.Right - 4, Half + 4, RcItem.Right - 4, Half + 6 + (Half / 2), m_Hdc, TCol2
    DrawLine RcItem.Right - 4, Half + 6 + (Half / 2), RcItem.Right - 4, RcItem.Bottom - 3, m_Hdc, TCol3
    DrawLine RcItem.Right - 15, 4, RcItem.Right - 4, 4, m_Hdc, TCol4
    Call DrawGradientMenu(m_Hdc, RcItem.Right - 15, 5, 11, RcItem.Bottom - 7, GetRGBColors(GetLngColor(TCol5)), GetRGBColors(GetLngColor(TCol6)), GRADIENT_VERTICAL)
'============================================================================================================
ElseIf m_ColorScheme = WindowsXP_Silver Then
'============================================================================================================
'XP SILVER SCHEME..
    Half = (RcItem.Bottom - 3) / 2
    Call DrawGradientMenu(m_Hdc, RcItem.Right - 17, 2, 15, Half, GetRGBColors(TCol1), GetRGBColors(GetLngColor(TCol2)), GRADIENT_VERTICAL)
    Call DrawGradientMenu(m_Hdc, RcItem.Right - 17, Half + 2, 15, Half - 3, GetRGBColors(GetLngColor(TCol3)), GetRGBColors(GetLngColor(TCol3)), GRADIENT_VERTICAL)
    DrawLine RcItem.Right - 17, 3, RcItem.Right - 2, 3, m_Hdc, vbWhite
    DrawLine RcItem.Right - 16, 4, RcItem.Right - 16, RcItem.Bottom - 4, m_Hdc, vbWhite
    DrawLine RcItem.Right - 4, 4, RcItem.Right - 4, RcItem.Bottom - 4, m_Hdc, vbWhite
    DrawLine RcItem.Right - 17, RcItem.Bottom - 4, RcItem.Right - 3, RcItem.Bottom - 4, m_Hdc, TCol4

End If
  
 Call DrawMenuRectangle(RcItem.Right - 17, 2, 15, RcItem.Bottom - 4, GetLngColor(FColor), GetLngColor(FColor), True)
 Call SetPixelV(m_Hdc, RcItem.Right - 3, 2, vbWhite)
 Call SetPixelV(m_Hdc, RcItem.Right - 3, RcItem.Bottom - 3, vbWhite)
 Call SetPixelV(m_Hdc, RcItem.Right - 17, 2, vbWhite)
 Call SetPixelV(m_Hdc, RcItem.Right - 17, RcItem.Bottom - 3, vbWhite)
 Call DrawArrow(RcItem.Right - 1, (RcItem.Top + RcItem.Bottom) / 2, GetLngColor(DColor))

End Sub

Private Sub DrawMenuRectangle(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, ByVal Color2 As Long, Optional BorderOnly As Boolean = False)

Dim Fill As Long
Dim Outline As Long
Dim Rec As RECT

    Rec.Left = X
    Rec.Top = Y
    Rec.Right = X + Width
    Rec.Bottom = Y + Height

    Fill = CreateSolidBrush(Color)
    Outline = CreateSolidBrush(Color2)

    If Not BorderOnly Then
        FillRect m_Hdc, Rec, Fill
    Else
        FrameRect m_Hdc, Rec, Outline
    End If

DeleteObject Fill
DeleteObject Outline

End Sub

Private Sub DrawArrow(X As Long, Y As Long, Optional cColor As Long = vbBlack)

Dim J1 As Integer, J2 As Integer
Dim Pen1 As Long, Pen2 As Long
Dim POS As POINTAPI

    Pen1 = CreatePen(0, 1, cColor)
    Pen2 = SelectObject(m_Hdc, Pen1)
          
            MoveToEx m_Hdc, X - 13, Y - 1, POS
            LineTo m_Hdc, X - 8, Y + 4
            MoveToEx m_Hdc, X - 5, Y - 1, POS
            LineTo m_Hdc, X - 10, Y + 4
            
            For J2 = 1 To 2
                For J1 = 0 To 1
                    MoveToEx m_Hdc, X - 12 / J2, Y - J1 - 1, POS
                    LineTo m_Hdc, X - 9 - (J2 - 2), Y + 3 - J1
                Next J1
            Next J2

    SelectObject m_Hdc, Pen2
    DeleteObject Pen2
    DeleteObject Pen1
  
End Sub


Private Sub SchemeControl()

    Select Case m_ColorScheme
              
        Case SystemColors
             ActualHighlight = GetLngColor(vbHighlight)
             ActualGrayText = GetLngColor(vbGrayText)
        Case WindowsXP_Blue
             ActualHighlight = GetLngColor(XPBlue_Highlight)
             ActualGrayText = GetLngColor(XPBlue_GrayText)
        Case WindowsXP_OliveGreen
             ActualHighlight = GetLngColor(XPGreen_Highlight)
             ActualGrayText = GetLngColor(XPBlue_GrayText)
        Case WindowsXP_Silver
             'ActualHighlight = GetLngColor(XPSilver_Highlight)
             ActualGrayText = GetLngColor(XPBlue_GrayText)

             
    End Select
    
    
End Sub


Public Property Let ColorScheme(ByRef cColorScheme As CWindowColors)
   m_ColorScheme = cColorScheme
End Property

Public Property Let State(ByVal cState As ControlState)
   m_State = cState
End Property

Public Property Let hwnd(ByVal cHwnd As Long)
   m_hWnd = cHwnd
End Property

Public Property Let hdc(ByVal cHdc As Long)
   m_Hdc = cHdc
End Property

