VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgressXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_hWnd As Long
Private m_Hdc As Long
Private m_Min As Integer
Private m_Max As Integer
Private m_Value As Integer
Private fPercent As Double
Private m_Scrolling As Byte
Private m_Orientation As Byte
Private TR As RECT, tBR As RECT, tSR As RECT
Private lSegmentWidth As Long, lSegmentSpacing As Long
Private m_ColorScheme As CWindowColors

Private ActualBarColor As Long

'==========================================================
'/---Draw ALL ProgressXP Bar  !!!!PUBLIC CALL!!!
'==========================================================

Public Sub DrawProgressBar()

  GetClientRect m_hWnd, TR                '//--- Reference = Control Client Area
  
    
            SchemeControl                 '//-- Get Scheme Colors
            
            CalcBarSize                   '//--- Calculate Progress and Percent Values
  
            PBarDraw                      '//--- Draw Scolling Bar (Inside Bar)
          
            If m_Scrolling = 0 Then DrawDivisions  '//--- Draw SegmentSpacing (This Will Generate the Blocks Effect)
  
    pDrawBorder                  '//--- Draw The XP Look Border

End Sub


'==========================================================
'/---Calculate Division Bars & Percent Values
'==========================================================

Private Sub CalcBarSize()

   lSegmentWidth = 8   '/-- Windows Default
   lSegmentSpacing = 2 '/-- Windows Default
         
   LSet tBR = TR

   fPercent = (m_Value - m_Min) / (m_Max - m_Min)
   If fPercent > 1# Then fPercent = 1#              '/--  0 < Percent < 100
   If fPercent < 0# Then fPercent = 0#
   
      If m_Orientation = 0 Then
      
      '=======================================================================================
      '                                 Calc Horizontal ProgressBar
      '---------------------------------------------------------------------------------------
         tBR.Right = TR.Left + (TR.Right - TR.Left) * fPercent
         tBR.Right = tBR.Right - ((tBR.Right - tBR.Left) Mod (lSegmentWidth + lSegmentSpacing))
         If tBR.Right < TR.Left Then
            tBR.Right = TR.Left
         End If
         If tBR.Right < TR.Left Then tBR.Right = TR.Left
         
      Else
      
      '=======================================================================================
      '                                 Calc Vertical ProgressBar
      '---------------------------------------------------------------------------------------
         fPercent = 1# - fPercent - 0.03
         tBR.Top = TR.Top + (TR.Bottom - TR.Top) * fPercent
         tBR.Top = tBR.Top - ((tBR.Top - tBR.Bottom) Mod (lSegmentWidth + lSegmentSpacing))
         If tBR.Top > TR.Bottom Then tBR.Top = TR.Bottom
    
         
      
      End If

End Sub

'==========================================================
'/---Draw Division Bars
'==========================================================

Private Sub DrawDivisions()
 Dim i As Long
 Dim hBR As Long
  
  hBR = CreateSolidBrush(vbWhite)
  
      LSet tSR = TR
      
      If m_Orientation = 0 Then
      
      '=======================================================================================
      '                                 Draw Horizontal ProgressBar
      '---------------------------------------------------------------------------------------
         For i = tBR.Left + lSegmentWidth To tBR.Right Step lSegmentWidth + lSegmentSpacing
            tSR.Left = i + 2
            tSR.Right = i + 2 + lSegmentSpacing
            FillRect m_Hdc, tSR, hBR
         Next i
      '---------------------------------------------------------------------------------------
      
      Else
      
      '=======================================================================================
      '                                  Draw Vertical ProgressBar
      '---------------------------------------------------------------------------------------
         For i = tBR.Bottom To tBR.Top + lSegmentWidth Step -(lSegmentWidth + lSegmentSpacing)
            tSR.Top = i - 2
            tSR.Bottom = i - 2 + lSegmentSpacing
            FillRect m_Hdc, tSR, hBR
         Next i
       '---------------------------------------------------------------------------------------
      
      End If
      
      DeleteObject hBR
     
End Sub


'==========================================================
'/---Draw The ProgressXP Bar Border  ;)
'==========================================================

Private Sub pDrawBorder()
Dim RTemp As RECT
 
 Let RTemp = TR
  
 RTemp.Left = TR.Left + 1: RTemp.Top = TR.Top + 1
 DrawRectangle RTemp, GetLngColor(&HBEBEBE), m_Hdc
 RTemp.Left = TR.Left + 1: RTemp.Top = TR.Top + 2: RTemp.Right = TR.Right - 1: RTemp.Bottom = TR.Bottom - 1
 DrawRectangle RTemp, GetLngColor(&HEFEFEF), m_Hdc
 DrawRectangle TR, GetLngColor(&H686868), m_Hdc
 
 Call MakeRegion(TR, m_hWnd)
 Call SetPixelV(m_Hdc, 1, 1, GetLngColor(&H686868))
 Call SetPixelV(m_Hdc, TR.Right - 2, 1, GetLngColor(&H686868))
 Call SetPixelV(m_Hdc, 1, TR.Bottom - 2, GetLngColor(&H686868))
 Call SetPixelV(m_Hdc, TR.Right - 2, TR.Bottom - 2, GetLngColor(&H686868))  '//--Clean Up Corners

End Sub


'==========================================================
'/---Draw The ProgressXP Bar ;)
'==========================================================

Private Sub PBarDraw()
Dim TempRect As RECT

If m_Orientation = 0 Then

    TempRect.Left = 2
    TempRect.Right = tBR.Right
    TempRect.Top = 8
    TempRect.Bottom = TR.Bottom - 6


    '=======================================================================================
    '                                 Draw Horizontal ProgressBar
    '---------------------------------------------------------------------------------------
   
    DrawGradientMenu m_Hdc, 2, 3, tBR.Right - 2, 6, GetRGBColors(ShiftColorOXP(ActualBarColor, 150)), GetRGBColors(ActualBarColor)
    DrawFillRectangle TempRect, ActualBarColor, m_Hdc
    DrawGradientMenu m_Hdc, 2, TempRect.Bottom - 2, tBR.Right - 2, 6, GetRGBColors(ActualBarColor), GetRGBColors(ShiftColorOXP(ActualBarColor, 150))
   
Else
    
    TempRect.Left = 7
    TempRect.Right = TR.Right - 8
    TempRect.Top = tBR.Top
    TempRect.Bottom = TR.Bottom
    
    '=======================================================================================
    '                                 Draw Vertical ProgressBar
    '---------------------------------------------------------------------------------------
   
    DrawGradientMenu m_Hdc, 2, tBR.Top, 6, TR.Bottom, GetRGBColors(ShiftColorOXP(ActualBarColor, 150)), GetRGBColors(ActualBarColor), 0
    DrawFillRectangle TempRect, ActualBarColor, m_Hdc
    DrawGradientMenu m_Hdc, TR.Right - 8, tBR.Top, 6, TR.Bottom, GetRGBColors(ActualBarColor), GetRGBColors(ShiftColorOXP(ActualBarColor, 150)), 0
    
   
    '--------------------   <-------- Gradient Color From (- to +)
    '||||||||||||||||||||   <-------- Fill Color
    '--------------------   <-------- Gradient Color From (+ to -)

End If

End Sub


Private Sub SchemeControl()

    Select Case m_ColorScheme
              
        Case SystemColors
             ActualBarColor = GetLngColor(vbHighlight)
             
        Case WindowsXP_Blue
             ActualBarColor = GetLngColor(XPBlue_ProgressBar)
             
        Case WindowsXP_OliveGreen
             ActualBarColor = GetLngColor(XPGreen_ProgressBar)
        
        Case WindowsXP_Silver
             ActualBarColor = GetLngColor(XPSilver_ProgressBar)
      
         
    End Select
    
    
End Sub

Public Property Let ColorScheme(ByRef cColorScheme As CWindowColors)
   m_ColorScheme = cColorScheme
End Property

Public Property Let hwnd(ByVal cHwnd As Long)
   m_hWnd = cHwnd
End Property

Public Property Let hdc(ByVal cHdc As Long)
   m_Hdc = cHdc
End Property

Public Property Let Min(ByVal cMin As Integer)
   m_Min = cMin
End Property

Public Property Let Max(ByVal cMax As Integer)
   m_Max = cMax
End Property

Public Property Let Scrolling(ByVal cScrolling As Byte)
   m_Scrolling = cScrolling
End Property

Public Property Let Orientation(ByVal cOrientation As Byte)
   m_Orientation = cOrientation
End Property

Public Property Let Value(ByVal cValue As Integer)
   m_Value = cValue
End Property
