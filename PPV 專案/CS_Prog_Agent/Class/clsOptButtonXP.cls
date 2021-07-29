VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOptButtonXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_hWnd As Long
Private m_Hdc As Long
Private cHdc As Long
Private m_Value As Boolean
Private m_Alignment As Byte
Private m_Enabled As Boolean
Private m_Down As Boolean
Private m_Over As Boolean
Private RcItem As RECT

Private Type PxC
    xPos As Long
    Col As Long
End Type


Public Sub DrawOptionButton()

    
    Call GetOriginalRect '//--- Select the Specific Area where the Paint is going to take Action!.
 
        
        
        If m_Down Then
             DrawGradientMenu m_Hdc, RcItem.Left, RcItem.Top, RcItem.Right, RcItem.Bottom - RcItem.Top, GetRGBColors(GetLngColor(&HA7B0B0)), GetRGBColors(&HDFEFF1), GRADIENT_HORIZONTAL
        ElseIf m_Over Then
             DrawGradientMenu m_Hdc, RcItem.Left, RcItem.Top, RcItem.Right, RcItem.Bottom - RcItem.Top, GetRGBColors(GetLngColor(&H84D6FD)), GetRGBColors(&H30B3F8), GRADIENT_HORIZONTAL
        Else
             DrawGradientMenu m_Hdc, RcItem.Left, RcItem.Top, RcItem.Right, RcItem.Bottom - RcItem.Top, GetRGBColors(GetLngColor(&HD7DCDC)), GetRGBColors(vbWhite), GRADIENT_HORIZONTAL
        End If
    
    
            DrawBody          '//--- Draw The Circular Corners ...  :)
            CenterBullet      '//--- Draw The Center Bullet...
            CleanCircularArea '//--- Clean The Corners of the Rectangle --> gives a Option Button look ;)
      

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

Private Sub DrawBody()
Dim Z As Long  '//---For next
Dim C1 As Long, C2 As Long, C3 As Long, C4 As Long    '//---Border Colors


If m_Enabled Then
    C1 = &HB6B09E
    C2 = &H987C57
    C3 = &H825623
    C4 = &HCED9D5
Else               '//--Assing The Gray Disabled Color if Disabled or Blue Color if Enabled
    C1 = &HCCDCDF
    C2 = &HC2D0D3
    C3 = &HBBC9CB
    C4 = &HD4E5E7
End If


Dim MP(0 To 12) As PxC
          
For Z = 0 To 1            '//--Draw The Circle .... <<   0= UpperBody    1= BelowBody  >>>
          
   '//-- Line 0
  GoSub ResetPoints
  MP(3).xPos = 1: MP(4).xPos = 1: MP(5).xPos = 1: MP(6).xPos = 1: MP(7).xPos = 1: MP(8).xPos = 1: MP(9).xPos = 1
  MP(3).Col = C4: MP(4).Col = C1: MP(5).Col = C2: MP(6).Col = C3: MP(7).Col = C2: MP(8).Col = C1: MP(9).Col = C4
  DrCl 0 + (12 * Z), MP
  '//-- Line 1
  GoSub ResetPoints
  MP(2).xPos = 1: MP(3).xPos = 1: MP(4).xPos = 1: MP(5).xPos = 1: MP(6).xPos = 1: MP(7).xPos = 1: MP(8).xPos = 1: MP(9).xPos = 1: MP(10).xPos = 1
  MP(2).Col = C1: MP(3).Col = C3: MP(4).Col = C2: MP(5).Col = C1: MP(6).Col = C4: MP(7).Col = C1: MP(8).Col = C3: MP(9).Col = C2: MP(10).Col = C1
  DrCl 1 + (10 * Z), MP
  '//-- Line 2
  GoSub ResetPoints
  MP(1).xPos = 1: MP(2).xPos = 1: MP(3).xPos = 1: MP(9).xPos = 1: MP(10).xPos = 1: MP(11).xPos = 1
  MP(1).Col = C1: MP(2).Col = C2: MP(3).Col = C1: MP(9).Col = C1: MP(10).Col = C2: MP(11).Col = C1
  DrCl 2 + (8 * Z), MP
  '//-- Line 3
  GoSub ResetPoints
  MP(0).xPos = 1: MP(1).xPos = 1: MP(2).xPos = 1: MP(10).xPos = 1: MP(11).xPos = 1: MP(12).xPos = 1
  MP(0).Col = C4: MP(1).Col = C3: MP(2).Col = C1: MP(10).Col = C1: MP(11).Col = C3: MP(12).Col = C4
  DrCl 3 + (6 * Z), MP
  '//-- Line 4
  GoSub ResetPoints
  MP(0).xPos = 1: MP(1).xPos = 1: MP(11).xPos = 1: MP(12).xPos = 1
  MP(0).Col = C1: MP(1).Col = C2: MP(11).Col = C2: MP(12).Col = C1
  DrCl 4 + (4 * Z), MP
  '//-- Line 5
  GoSub ResetPoints
  MP(0).xPos = 1: MP(1).xPos = 1: MP(11).xPos = 1: MP(12).xPos = 1
  MP(0).Col = C2: MP(1).Col = C1: MP(11).Col = C1: MP(12).Col = C2
  DrCl 5 + (2 * Z), MP
  
  Next Z
  
  '//-- Line 6
  GoSub ResetPoints
  MP(0).xPos = 1: MP(12).xPos = 1    '//--Middle Point of Circle
  MP(0).Col = C3: MP(12).Col = C3
  DrCl 6, MP

 
Exit Sub

ResetPoints:
  MP(0).xPos = 0: MP(1).xPos = 0: MP(2).xPos = 0: MP(3).xPos = 0: MP(4).xPos = 0: MP(5).xPos = 0: MP(6).xPos = 0:
  MP(7).xPos = 0: MP(8).xPos = 0: MP(9).xPos = 0: MP(10).xPos = 0: MP(11).xPos = 0: MP(12).xPos = 0
Return

End Sub


Private Sub DrCl(ByVal Yx As Long, ByRef Pointz() As PxC)
Dim Z As Long '//--For Next

For Z = 0 To 12
    If Pointz(Z).xPos = 1 Then Call SetPixelV(m_Hdc, Z + RcItem.Left, Yx + RcItem.Top, GetLngColor(Pointz(Z).Col))
Next Z

End Sub


Private Sub CleanCircularArea()
Dim Cleanx As Long '//--For Next
Dim Cleany As Long '//--For Next


 For Cleany = 0 To 2                '//--Clean Upper Area L
        For Cleanx = 0 To 2
            SetPixelV m_Hdc, RcItem.Left + Cleanx - Cleany, RcItem.Top + Cleany, GetLngColor(vbButtonFace)
        Next Cleanx
 Next Cleany

 
 For Cleany = 0 To 2                '//--Clean Upper Area R
        For Cleanx = 2 To 0 Step -1
            SetPixelV m_Hdc, RcItem.Left + Cleanx + Cleany + 10, RcItem.Top + Cleany, GetLngColor(vbButtonFace)
        Next Cleanx
 Next Cleany

 
 For Cleany = 0 To 2                '//--Clean Lower Area R
        For Cleanx = 0 To 2
            SetPixelV m_Hdc, RcItem.Left + Cleanx - Cleany, RcItem.Top + Abs(Cleany - 2) + 10, GetLngColor(vbButtonFace)
        Next Cleanx
 Next Cleany

 
 For Cleany = 0 To 2                '//--Clean Lower Area L
        For Cleanx = 2 To 0 Step -1
            SetPixelV m_Hdc, RcItem.Left + Cleanx + Cleany + 10, RcItem.Top + Abs(Cleany - 2) + 10, GetLngColor(vbButtonFace)
        Next Cleanx
 Next Cleany




End Sub

Private Sub CenterBullet()
Dim MD(0 To 4) As PxC
Dim TempRect As RECT
Dim Px As Long
Dim Py As Long


 TempRect.Left = RcItem.Left + 4
 TempRect.Right = RcItem.Right - 4
 TempRect.Top = RcItem.Top + 4
 TempRect.Bottom = RcItem.Bottom - 4

    
 If m_Over And Not m_Down Then DrawFillRectangle TempRect, vbWhite, m_Hdc
 
 If Not m_Down Then
 
 For Py = -1 To 5 Step 6
     For Px = 1 To 3 Step 2
         SetPixelV m_Hdc, TempRect.Left + Px, TempRect.Top + Py, IIf(m_Enabled, GetLngColor(&HB3E3F7), vbWhite)
         SetPixelV m_Hdc, TempRect.Left + Py, TempRect.Top + Px, IIf(m_Enabled, GetLngColor(&HB3E3F7), vbWhite)
         SetPixelV m_Hdc, TempRect.Left + 2, TempRect.Top + Py, vbWhite
         SetPixelV m_Hdc, TempRect.Left + Py, TempRect.Top + 2, vbWhite
    Next Px
 Next Py
    
 End If
 
 If m_Value = False Then Exit Sub  '//---If Need To draw Selected State Go on ;)
 
 
 
 '====================================================================================
 '                           Draw Inside Bullet Area (Green Dot) ...   1-2
 '====================================================================================
 
   '//-- Line 0
  MD(0).xPos = 0:  MD(1).xPos = 1: MD(2).xPos = 1: MD(3).xPos = 1: MD(4).xPos = 0
  MD(1).Col = &HA0DEAC: MD(2).Col = &H48BF4D: MD(3).Col = &H94D1A0
  
  CenterDot 0, MD
  '//-- Line 1
  MD(0).xPos = 1: MD(1).xPos = 1: MD(2).xPos = 1: MD(3).xPos = 1: MD(4).xPos = 1
  MD(0).Col = &HA0DEAC: MD(1).Col = &H51D555: MD(2).Col = &H3FC343: MD(3).Col = &H26A829: MD(4).Col = &H82C898
  CenterDot 1, MD
  '//-- Line 2
  MD(0).xPos = 1: MD(1).xPos = 1: MD(2).xPos = 1: MD(3).xPos = 1: MD(4).xPos = 1
  MD(0).Col = &H48BF4D: MD(1).Col = &H3FC342: MD(2).Col = &H35B938: MD(3).Col = &H21A121: MD(4).Col = &H209525
  CenterDot 2, MD
  '//-- Line 3
  MD(0).xPos = 1: MD(1).xPos = 1: MD(2).xPos = 1: MD(3).xPos = 1: MD(4).xPos = 1
  MD(0).Col = &H94D1A0: MD(1).Col = &H27A82A: MD(2).Col = &H20A222: MD(3).Col = &H109213: MD(4).Col = &H88C293
  CenterDot 3, MD
  '//-- Line 4
  MD(0).xPos = 0:  MD(1).xPos = 1: MD(2).xPos = 1: MD(3).xPos = 1: MD(4).xPos = 0
  MD(1).Col = &H8CC998: MD(2).Col = &H219523: MD(3).Col = &H88C393
  CenterDot 4, MD


End Sub

 '====================================================================================
 '                           Draw Inside Bullet Area (Green Dot) ...   2-2
 '====================================================================================
 
Private Sub CenterDot(ByVal Yx As Long, ByRef Pointz() As PxC)
Dim Z As Long '//--For Next


For Z = 0 To 4
    If m_Enabled = False Then Pointz(Z).Col = GetLngColor(&HB6C0C1)
    If Pointz(Z).xPos = 1 Then Call SetPixelV(m_Hdc, RcItem.Left + 4 + Z, RcItem.Top + 4 + Yx, GetLngColor(Pointz(Z).Col))
Next Z

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

Public Property Let Value(ByVal cValue As Boolean)
   m_Value = cValue
End Property

Public Property Let hwnd(ByVal cHwnd As Long)
   m_hWnd = cHwnd
End Property

Public Property Let hdc(ByVal cHdc As Long)
   m_Hdc = cHdc
End Property

