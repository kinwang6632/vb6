VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWinXPengine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Implements clsISubClass

Private m_bSubclass        As Boolean
Private m_Hwnd             As Long

Private m_ActiveScaleMode As ScaleModeConstants
Private m_SchemeColor As CWindowColors
Private m_IdeSubClass As Boolean
Private m_CBug As Boolean

Private m_Button      As Boolean
Private m_CheckBox    As Boolean
Private m_Combo       As Boolean '+ DriveList
Private m_FileListBox As Boolean '+ DirListBox
Private m_Frame       As Boolean
Private m_ICombo      As Boolean
Private m_ListBox     As Boolean
Private m_MsgBox      As Boolean
Private m_CommonD     As Boolean
Private m_Option      As Boolean
Private m_Picture     As Boolean
Private m_ProgressBar As Boolean
Private m_Text        As Boolean
Private m_Static      As Boolean

Private MyClassObject As Object         '//--- Object Sender  (Contains Subclassed Control Information) VB 6

Private ButtonState As ControlState
Private MovementFlag As ControlState

Private isFocused As Boolean                   '//---  Controls General Properties
Private bOver As Boolean
Private bDown As Boolean

Private m_CurrentButton As Integer          '//---  Adoedc Control Variables

Private m_CurrentItem As Long
    
Private DestDC      As Long
Private MaskDC      As Long
Private MemDC       As Long
Private OrigDC      As Long
Private MaskPic     As Long              'Temporary DC
Private MemPic      As Long
Private TempPic     As Long
Private OrigPic     As Long
Private TempDC      As Long

Private origBrush As Long
Private TempBrush As Long
Private origColor As Long 'BackColor
Dim m_ItemRect As RECT
Dim m_Width    As Long
Dim m_Height   As Long

Private m_PreDraw As Boolean
Private iHw As Integer, iLW As Integer

Public Sub BeforeAttachMessageBox(ByVal lhWnd As Long)
    If lhWnd = 0 Then Exit Sub
    pRelease
    m_MsgBox = True      ' PRE-DEFINE "m_MsgBox" VALUE BEFORE ENGINE IS STARTED
    pAttach lhWnd
End Sub
        
Public Sub BeforeAttachCommonDialog(ByVal lhWnd As Long)
    If lhWnd = 0 Then Exit Sub
    pRelease
    m_CommonD = True
    pAttach lhWnd
End Sub

Public Sub AfterAttachMessageBox(ByVal lhWnd As Long)
    If lhWnd = 0 Then Exit Sub
    pRelease
    Select Case ThisWindowClassName(lhWnd)
            Case "Edit"
                    If InStr(ThisWindowClassName(GetParent(lhWnd)), "ComboBox") = 0 Then
                        MakeWindowFlat lhWnd
                        m_Text = True
                    End If            '//---MessagesBoxes and InputBoxes only have this kind of controls
            Case "Button"
                    If Left$(GetObjectText(lhWnd), InStr(GetObjectText(lhWnd), vbNullChar) - 1) <> "Open as &read-only" Then
                        m_Button = True
                    Else
                        m_CheckBox = True
                    End If
       Case "ComboBox"
                    m_Combo = True
       Case Else
                    Exit Sub '(Nothing) UnSupported Control
     End Select
    pAttach lhWnd
End Sub

Public Sub Attach(ByVal ObjThis As Object)
    Dim lhWnd As Long
    pRelease '//--- First UnSubclass All Controls.. ;)
    On Error Resume Next
    lhWnd = ObjThis.hwnd '//---Short Reference
    If Err.Number <> 0 Then Exit Sub  'Something happen :(
    Set MyClassObject = ObjThis '//--- Control's Properties Container
    Select Case ThisWindowClassName(lhWnd)
            Case "ThunderCheckBox", "ThunderRT6CheckBox"
                    m_CheckBox = True
            Case "ThunderComboBox", "ThunderRT6ComboBox", "ThunderDriveListBox", "ThunderRT6DriveListBox"
                    m_Combo = True
            Case "ThunderCommandButton", "ThunderRT6CommandButton", "Button"
                    m_Button = True
            Case "ThunderFrame", "ThunderRT6Frame"
                    If MyClassObject.BorderStyle = 0 Then Exit Sub        '//--- If Frame <> Fixed Single Then Don't Sublcass ;)
                     m_Frame = True
            Case "ComboBox", "ImageCombo20WndClass"
                    lhWnd = FindWindowEx(lhWnd, 0&, "ComboBox", ByVal 0&) '//--- Find Combo Inside control
                    m_ICombo = True
            Case "ThunderOptionButton", "ThunderRT6OptionButton"
                    m_Option = True
            Case "ProgressBar20WndClass"
                    MakeWindowFlat lhWnd
                    SendMessageLong lhWnd, CCM_SETBKCOLOR, 0&, ByVal &HFFFFFF  '//--Give The XP Look BackGroundColor = White
                    SendMessageLong lhWnd, PBM_SETBARCOLOR, 0&, ByVal &HFFFFFF             '//--Give The XP Look BackGroundColor = White
                    m_ProgressBar = True
            Case "ThunderListBox", "ThunderRT6ListBox", "ThunderDirListBox", "ThunderRT6DirListBox"
                    MakeWindowFlat lhWnd
                    m_ListBox = True
            Case "ThunderFileListBox", "ThunderRT6FileListBox"
                    MakeWindowFlat lhWnd
                    m_FileListBox = True
            Case "ThunderPictureBox", "ThunderPictureBoxDC", "ThunderRT6PictureBoxDC"
                    MakeWindowFlat lhWnd
                    m_Picture = True
            Case Else
                    Exit Sub '(Nothing) UnSupported Control <---Actually This Line is never Fired.
    End Select
   pAttach lhWnd '//--- Finally Subclass
End Sub

Private Sub pAttach(ByRef hWndA As Long)
    m_Hwnd = hWndA '//-- Reference Used in All Class Subs & Functions Point's Control hWnd
    SubclassMessage m_Hwnd, True '//--- Subclass
    m_bSubclass = True '//--- Success Control is Subclassed !!  ;)
End Sub

Private Sub pRelease()
    If m_bSubclass = False Then Exit Sub '//--- No Subclass before...Don't Unsubclass (Exit)
    SubclassMessage m_Hwnd, False  '//--- Unsubclass
    RedrawWindow m_Hwnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) turn Normal ;)
    KillTimer m_Hwnd, 1         '//--  Should never happen that we have a timer left over (but just in case)
    Set MyClassObject = Nothing '//--  Cleaning ClassObject just in case
End Sub

Private Sub SubclassMessage(ByVal Sender As Long, ByVal Action As Boolean)
    If Action = True Then
        If Not m_MsgBox Or Not m_CommonD Then AttachMessage Me, Sender, WM_PAINT
        If m_MsgBox Or m_CommonD Then AttachMessage Me, Sender, WM_ACTIVATE
        If m_Button Then AttachMessage Me, Sender, WM_SETFOCUS
        If m_Button Then AttachMessage Me, Sender, WM_KILLFOCUS
        If m_Button Or m_CheckBox Or m_Option Then AttachMessage Me, Sender, WM_KEYDOWN
        If m_Button Or m_CheckBox Or m_Option Then AttachMessage Me, Sender, WM_KEYUP
        If m_Button Or m_CheckBox Or m_Combo Or m_ICombo Or m_Option Then AttachMessage Me, Sender, WM_MOUSEMOVE
        If m_Button Or m_CheckBox Or m_Combo Or m_ICombo Or m_Option Then AttachMessage Me, Sender, WM_TIMER
        If m_Button Or m_CheckBox Or m_Combo Or m_ICombo Or m_Option Then AttachMessage Me, Sender, WM_LBUTTONDOWN
        If m_Button Or m_CheckBox Or m_Combo Or m_ICombo Or m_Option Then AttachMessage Me, Sender, WM_LBUTTONUP
        If m_Picture Or m_Button Or m_CheckBox Or m_ICombo Or m_Option Then AttachMessage Me, Sender, WM_ENABLE
        If m_Combo Then AttachMessage Me, GetParent(Sender), WM_COMMAND
        If m_ICombo Then AttachMessage Me, GetWindow(Sender, GW_CHILD), WM_COMMAND
    Else
        If Not m_MsgBox Or Not Not m_CommonD Then DetachMessage Me, Sender, WM_PAINT
        If m_MsgBox Or m_CommonD Then DetachMessage Me, Sender, WM_ACTIVATE
        If m_Button Then DetachMessage Me, Sender, WM_SETFOCUS
        If m_Button Then DetachMessage Me, Sender, WM_KILLFOCUS
        If m_Button Or m_CheckBox Or m_Option Then DetachMessage Me, Sender, WM_KEYDOWN
        If m_Button Or m_CheckBox Or m_Option Then DetachMessage Me, Sender, WM_KEYUP
        If m_Button Or m_CheckBox Or m_Combo Or m_ICombo Or m_Option Then DetachMessage Me, Sender, WM_MOUSEMOVE
        If m_Button Or m_CheckBox Or m_Combo Or m_ICombo Or m_Option Then DetachMessage Me, Sender, WM_TIMER
        If m_Button Or m_CheckBox Or m_Combo Or m_ICombo Or m_Option Then DetachMessage Me, Sender, WM_LBUTTONDOWN
        If m_Button Or m_CheckBox Or m_Combo Or m_ICombo Or m_Option Then DetachMessage Me, Sender, WM_LBUTTONUP
        If m_Picture Or m_Button Or m_CheckBox Or m_ICombo Or m_Option Then DetachMessage Me, Sender, WM_ENABLE
        If m_Combo Then DetachMessage Me, GetParent(Sender), WM_COMMAND
        If m_ICombo Then DetachMessage Me, GetWindow(Sender, GW_CHILD), WM_COMMAND
    End If
End Sub

Private Sub DrawControlClass()
    If m_Hwnd = 0 Then Exit Sub
    If m_Combo Or m_ICombo Then
        DrawCombo
    ElseIf m_Button Then
        DrawButton
    ElseIf m_ProgressBar Then
        DrawProgressBar
    ElseIf m_CheckBox Then
        DrawCheckBox
    ElseIf m_Option Then
        DrawOptionButton
    ElseIf m_Frame Then
        DrawFrame
    Else
    End If
End Sub

Private Sub DrawButton()
    Dim NewButton As clsButtonXP
    Dim hdc As Long
    Dim lState As Long
    lState = SendMessageLong(m_Hwnd, BM_GETSTATE, 0, 0)
    hdc = GetDC(m_Hwnd)
    Set NewButton = New clsButtonXP
    With NewButton
        .hdc = IIf(IsWindowEnabled(m_Hwnd) = False, hdc, IIf(m_PreDraw = True, TempDC, hdc))
        .hwnd = m_Hwnd
        .Enabled = IsWindowEnabled(m_Hwnd)
        .Focused = isFocused
        .Over = bOver
        .Down = ((lState And BST_CHECKED) = BST_CHECKED) Or ((lState And BST_PUSHED) = BST_PUSHED)
        .ColorScheme = m_SchemeColor
        .PreDraw = m_PreDraw
        .DrawButtonXP
    End With
    Set NewButton = Nothing
    ReleaseDC m_Hwnd, hdc
End Sub

Private Sub DrawFrame()
    Dim NewFrame As clsFrameXP
    Dim hdc As Long
    hdc = GetDC(m_Hwnd)
    Set NewFrame = New clsFrameXP
    With NewFrame
        .hwnd = m_Hwnd
        .hdc = hdc
        .FontItalic = MyClassObject.FontItalic
        .FontSize = MyClassObject.FontSize
        .FontUnderline = MyClassObject.FontUnderline
        .FrameCaption = MyClassObject.Caption
        .FrameFont = MyClassObject.FontName
        .Backcolor = MyClassObject.Backcolor
        .Forecolor = MyClassObject.Forecolor
        .DrawFrame
    End With
    Set NewFrame = Nothing
    ReleaseDC m_Hwnd, hdc
End Sub

Private Sub DrawProgressBar()
    Dim NewProgressBarXP As clsProgressXP
    Dim hdc As Long
    hdc = GetDC(m_Hwnd)
    Set NewProgressBarXP = New clsProgressXP
    With NewProgressBarXP
        .hdc = hdc
        .hwnd = m_Hwnd
        .Max = MyClassObject.Max
        .Min = MyClassObject.Min
        .Value = MyClassObject.Value
        .Orientation = MyClassObject.Orientation
        .Scrolling = MyClassObject.Scrolling
        .ColorScheme = m_SchemeColor
        .DrawProgressBar
    End With
    Set NewProgressBarXP = Nothing
    ReleaseDC m_Hwnd, hdc
End Sub

Private Sub DrawCheckBox()
    Dim NewCheckBoxXP As clsChkBoxXP
    Dim hdc As Long
    Dim lState As Long
    lState = SendMessageLong(m_Hwnd, BM_GETSTATE, 0, 0)
    hdc = GetDC(m_Hwnd)
    Set NewCheckBoxXP = New clsChkBoxXP
    With NewCheckBoxXP
        .hdc = hdc
        .hwnd = m_Hwnd
        .Enabled = IsWindowEnabled(m_Hwnd)
        .Over = bOver
        .Down = ((lState And BST_PUSHED) = BST_PUSHED)
        .Value = SendMessageLong(m_Hwnd, BM_GETCHECK, 0&, 0&)
        .Alignment = WindowAlignment(m_Hwnd)
        .DrawCheckBox
    End With
    Set NewCheckBoxXP = Nothing
    ReleaseDC m_Hwnd, hdc
End Sub

Private Sub DrawOptionButton()
    Dim NewOptionButtonXP As clsOptButtonXP
    Dim hdc As Long
    Dim lState As Long
    lState = SendMessageLong(m_Hwnd, BM_GETSTATE, 0, 0)
    hdc = GetDC(m_Hwnd)
    Set NewOptionButtonXP = New clsOptButtonXP
    With NewOptionButtonXP
        .hdc = hdc
        .hwnd = m_Hwnd
        .Enabled = IsWindowEnabled(m_Hwnd)
        .Over = bOver
        .Down = ((lState And BST_PUSHED) = BST_PUSHED)
        .Value = SendMessageLong(m_Hwnd, BM_GETCHECK, 0&, 0&)
        .Alignment = WindowAlignment(m_Hwnd)
        .DrawOptionButton
    End With
    Set NewOptionButtonXP = Nothing
    ReleaseDC m_Hwnd, hdc
End Sub

Private Sub DrawCombo()
    Dim NewComboXP As clsComboXP
    Dim hdc As Long
    hdc = GetDC(m_Hwnd)
    Set NewComboXP = New clsComboXP
    With NewComboXP
        .hdc = hdc
        .hwnd = m_Hwnd
        .State = IIf(IsWindowEnabled(m_Hwnd) = 0, C_Disabled, ButtonState)
        .ColorScheme = m_SchemeColor
        .DrawComboBox
    End With
    Set NewComboXP = Nothing
    ReleaseDC m_Hwnd, hdc
End Sub

Private Function DroppedDown() As Boolean
    DroppedDown = (SendMessageLong(m_Hwnd, CB_GETDROPPEDSTATE, 0, 0) <> 0)
End Function

Private Sub Class_Terminate()
    pRelease
'    If hHook Then UnhookWindowsHookEx hHook
End Sub

Public Sub UnloadEngine()
    Class_Terminate
End Sub

Private Property Let clsISubClass_MsgResponse(ByVal RHS As EMsgResponse)
   ' intentionally blank
End Property

Private Property Get clsISubClass_MsgResponse() As EMsgResponse
    If (CurrentMessage = WM_PAINT) Then
        clsISubClass_MsgResponse = emrPreprocess
    Else
        clsISubClass_MsgResponse = emrPostProcess
    End If
End Property

Private Function clsISubClass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If m_ProgressBar Then DrawControlClass
    Select Case iMsg   '//--- SELECT THE SUBCLASSED WINDOW MESSAGE.... CLEAR ENOUGH?
            Case WM_COMMAND
                    If (m_Hwnd = lParam) Then ' lParam handle to Control (combo box)
                        Select Case wParam \ &H10000  'WParam identifier of control (combo box)
                                Case CBN_CLOSEUP
                                    ButtonState = C_Normal
                                    MovementFlag = C_Disabled
                                    bOver = False
                                    DrawControlClass
                                    KillTimer m_Hwnd, 1
                        Case Else
                                    'Nothing
                        End Select
                    End If
            Case WM_PRINTCLIENT
                    PrintTabClientColor
            Case WM_ACTIVATE
                    If m_MsgBox And InStr(ThisWindowClassName(lParam), "#32770") > 0 Then
                        Icount = 0 'm_ActiveClassesNumber
                        EnumChildWindows lParam, AddressOf EnumChildProc, ByVal 0&
                    End If
'                    If m_CommonD And InStr(ThisWindowClassName(lParam), "NewHelpClass") > 0 Then
'                        If hHook Then UnhookWindowsHookEx hHook
'                        hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
'                    End If
            Case WM_PAINT
                    If m_Button Then
                        If IsWindowEnabled(m_Hwnd) Then
                            PrePaintControl
                            clsISubClass_WindowProc = CallOldWindowProc(hwnd, iMsg, OrigDC, lParam)
                        End If
                        PostPaintControl
                    ElseIf Not m_ProgressBar Then
                        DrawControlClass
                    End If
            Case WM_LBUTTONDOWN
                    m_CBug = True
                    bDown = True
                    ButtonState = C_Down
                    RedrawWindow m_Hwnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
            Case WM_LBUTTONUP
                    m_CBug = True
                    bDown = False
                    ButtonState = C_Up
                    If m_Option Then
                        If InsideArea(m_Hwnd) = True Then
                            EnumChildWindows GetParent(m_Hwnd), AddressOf CleanOptionButtomArray, ByVal 0&
                        Else
                            DrawControlClass        '//--This Works for None-Graphic & None-OwnedDraw OptionButtons
                        End If
                    Else
                        RedrawWindow m_Hwnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
                    End If
                    m_CurrentItem = -1
            Case WM_ENABLE
                    If IsWindowEnabled(m_Hwnd) = 0 Then
                        ButtonState = C_Disabled
                    Else
                        ButtonState = C_Normal
                    End If
                    RedrawWindow m_Hwnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
            Case WM_SETFOCUS
                    isFocused = True
                    If m_Button Then
                        PostPaintControl
                    Else
                        RedrawWindow m_Hwnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
                    End If
            Case WM_KILLFOCUS
                    isFocused = False
                    PostPaintControl
            Case WM_KEYDOWN
                    If Not bDown Then
                        bDown = True
                        LongInt2Int wParam, iHw, iLW
                        Select Case (iLW)
                                Case vbKeySpace
                                        RedrawWindow m_Hwnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
                        End Select
                    End If
            Case WM_KEYUP
                    If bDown Then
                        bDown = False
                        LongInt2Int wParam, iHw, iLW
                        Select Case (iLW)
                                Case vbKeySpace
                                        RedrawWindow m_Hwnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
                        End Select
                    End If
            Case WM_MOUSEMOVE
                    If m_Button Or m_CheckBox Or m_Option Then
                        If InsideArea(m_Hwnd) = True Then
                            If bOver = False Then
                                bOver = True
                                RedrawWindow m_Hwnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
                                SetTimer m_Hwnd, 1, 1, 0
                            End If
                        Else
                            If bOver Then
                                RedrawWindow m_Hwnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
                                bOver = False
                            End If
                        End If
                    ElseIf m_Combo Or m_ICombo Then
                        If InsideArea(m_Hwnd) = True Then
                            If bOver = False Then
                                bOver = True
                                ButtonState = C_Over
                                DrawControlClass
                                SetTimer m_Hwnd, 1, 1, 0
                            End If
                        Else
                            ButtonState = C_Down
                        End If
                    End If
    End Select
End Function

Private Sub PrePaintControl()
    DestDC = GetDC(m_Hwnd)
    GetWindowRect m_Hwnd, m_ItemRect
    m_Width = m_ItemRect.Right - m_ItemRect.Left
    m_Height = m_ItemRect.Bottom - m_ItemRect.Top
    CreateNewDCWorkArea m_Width, m_Height
    SelectBitmap
    m_PreDraw = True
    DrawControlClass
End Sub

Private Sub PostPaintControl()
    CreateBackMask m_Width, m_Height
    origBrush = SelectObject(TempDC, TempBrush)
    SelectObject TempDC, origBrush
    '------------------------------------------------------------------------------------------------------
    DOBitBlt m_Width, m_Height '//--- Do RasterOperations
    CleanDCs                    '//--- Free Memory <--Prevent Leaks
    m_PreDraw = False
    DrawControlClass
End Sub

Private Sub CreateBackMask(ByVal m_Width As Long, ByVal m_Height As Long)
    origColor = SetBkColor(DestDC, GetSysColor(15))
    SetBkColor OrigDC, GetSysColor(15)
    BitBlt MaskDC, 0, 0, m_Width, m_Height, OrigDC, 0, 0, vbSrcCopy
End Sub

Private Sub CreateNewDCWorkArea(ByVal m_Width As Long, ByVal m_Height As Long)
    MaskDC = CreateCompatibleDC(DestDC)
    MaskPic = CreateBitmap(m_Width, m_Height, 1, 1, ByVal 0&)
    MemDC = CreateCompatibleDC(DestDC)
    MemPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)
    TempDC = CreateCompatibleDC(DestDC)
    TempPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)
    OrigDC = CreateCompatibleDC(DestDC)
    OrigPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)
End Sub

Private Sub DOBitBlt(ByVal m_Width As Long, ByVal m_Height As Long)
    BitBlt MemDC, 0, 0, m_Width, m_Height, MaskDC, 0, 0, vbSrcCopy
    BitBlt MemDC, 0, 0, m_Width, m_Height, OrigDC, 0, 0, vbSrcPaint
    BitBlt TempDC, 0, 0, m_Width, m_Height, MaskDC, 0, 0, vbMergePaint
    BitBlt TempDC, 0, 0, m_Width, m_Height, MemDC, 0, 0, vbSrcAnd
    BitBlt DestDC, 0, 0, m_Width, m_Height, TempDC, 0, 0, vbSrcCopy
End Sub

Private Sub CleanDCs()
    DeleteDC TempDC
    DeleteObject TempPic
    DeleteDC MaskDC
    DeleteObject MaskPic
    DeleteDC MemDC
    DeleteObject MemPic
    DeleteDC OrigDC
    DeleteObject OrigPic
    DeleteObject TempBrush
End Sub

Private Sub SelectBitmap()
    Dim cHandle As Long
    cHandle = SelectObject(MaskDC, MaskPic)
    DeleteObject cHandle
    cHandle = SelectObject(MemDC, MemPic)
    DeleteObject cHandle
    cHandle = SelectObject(TempDC, TempPic)
    DeleteObject cHandle
    cHandle = SelectObject(OrigDC, OrigPic)
    DeleteObject cHandle
End Sub

Private Sub PrintTabClientColor()
    Dim TRec As RECT
    Dim cHdc As Long
    Dim Valh As Long
    cHdc = GetDC(m_Hwnd)
    GetClientRect m_Hwnd, TRec
    Select Case m_ActiveScaleMode
            Case vbTwips
                    TRec.Top = Abs((MyClassObject.ClientTop - MyClassObject.Top) / Screen.TwipsPerPixelY)
            Case vbPixels
                    TRec.Top = Abs(MyClassObject.ClientTop - MyClassObject.Top)
            Case vbPoints
                    TRec.Top = (Abs(MyClassObject.ClientTop - MyClassObject.Top) * 100) / 75
    End Select
    DrawFillRectangle TRec, vbWhite, cHdc
    ReleaseDC m_Hwnd, cHdc
End Sub

Public Property Let ActiveScaleMode(ByRef cActiveScaleMode As ScaleModeConstants)
   m_ActiveScaleMode = cActiveScaleMode
End Property

Public Property Let IdeSubClass(ByVal cIdeSubClass As Boolean)
   m_IdeSubClass = cIdeSubClass
End Property

Public Property Get SchemeColor() As CWindowColors
   SchemeColor = m_SchemeColor
End Property

Public Property Let SchemeColor(ByRef cSchemeColor As CWindowColors)
   m_SchemeColor = cSchemeColor
End Property
