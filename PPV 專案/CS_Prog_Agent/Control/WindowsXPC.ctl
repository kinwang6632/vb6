VERSION 5.00
Begin VB.UserControl WinXPctl 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2265
   InvisibleAtRuntime=   -1  'True
   LockControls    =   -1  'True
   MaskColor       =   &H00FFFFFF&
   PropertyPages   =   "WindowsXPC.ctx":0000
   ScaleHeight     =   495
   ScaleWidth      =   2265
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "XP Engine 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   495
      TabIndex        =   0
      Top             =   142
      Width           =   1275
   End
   Begin VB.Shape shp 
      BackColor       =   &H00800000&
      BackStyle       =   1  '¤£³z©ú
      BorderColor     =   &H00FFFFFF&
      Height          =   500
      Left            =   0
      Top             =   0
      Width           =   2266
   End
End
Attribute VB_Name = "WinXPctl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_iCount As Long
Private m_SubClassedItem() As clsWinXPengine

Private m_IDECount As Long
Private m_IDEItem() As clsWinXPengine

Public Enum SchemeWindowColors
       System = 1
       XP_Blue = 2
       XP_OliveGreen = 3
       XP_Silver = 4
End Enum

Private m_EngineStarted As Boolean
Private m_IDEStarted As Boolean


Private m_IDE As Boolean
Private StartedEngine As Boolean
Private RunningApp As Boolean
Private i As Long

Private m_ColorScheme         As CWindowColors

Private m_FileListBoxControl  As Boolean
Private m_DirListBoxControl   As Boolean
Private m_ListBoxControl      As Boolean
Private m_ImageComboControl   As Boolean
Private m_ProgressBarControl  As Boolean
Private m_DriveListBoxControl As Boolean
Private m_ComboBoxControl     As Boolean
Private m_OptionControl       As Boolean
Private m_CheckControl        As Boolean
Private m_ButtonControl       As Boolean
Private m_FrameControl        As Boolean
Private m_PictureControl      As Boolean
Private m_MsgBox_InputBox     As Boolean
Private m_CommonDialog        As Boolean

Private Sub Timer1_Timer()
    If UserControl.Ambient.UserMode And Me.EngineStarted = False Then InitSubClassing
End Sub

Private Sub UserControl_InitProperties()
    m_ColorScheme = WindowsXP_Blue '//--DefaultColors
    m_ListBoxControl = True
    m_DirListBoxControl = True
    m_FileListBoxControl = True
    m_ImageComboControl = True
    m_ProgressBarControl = True
    m_DriveListBoxControl = True
    m_ComboBoxControl = True
    m_OptionControl = True
    m_CheckControl = True
    m_ButtonControl = True
    m_FrameControl = True
    m_PictureControl = True
    m_MsgBox_InputBox = True
    m_CommonDialog = False
End Sub

Public Sub InitSubClassing()
    Dim SubclassThis As Boolean
    Dim aControl As Control
    'If Not UserControl.Ambient.UserMode Then Exit Sub
    For Each aControl In Parent.Controls
        SubclassThis = False
        On Error Resume Next
        If Err.Number = 0 Then
            Select Case ThisWindowClassName(aControl.hwnd)
                    Case "ThunderFrame", "ThunderRT6Frame"
                            If m_FrameControl Then SubclassThis = True
                    Case "ThunderListBox", "ThunderRT6ListBox"
                            If m_ListBoxControl Then SubclassThis = True
                    Case "ThunderCommandButton", "ThunderRT6CommandButton", "Button"
                            If m_ButtonControl Then SubclassThis = True
                    Case "ThunderOptionButton", "ThunderRT6OptionButton"
                            If m_OptionControl Then SubclassThis = True
                    Case "ThunderCheckBox", "ThunderRT6CheckBox"
                            If m_CheckControl Then SubclassThis = True
                    Case "ThunderComboBox", "ThunderRT6ComboBox", "ComboBox"
                            If m_ComboBoxControl Then SubclassThis = True
                    Case "ImageCombo20WndClass"
                            If m_ImageComboControl Then SubclassThis = True
                    Case "ProgressBar20WndClass"
                            If m_ProgressBarControl Then SubclassThis = True
                    Case "ThunderDirListBox", "ThunderRT6DirListBox"
                            If m_DirListBoxControl Then SubclassThis = True
                    Case "ThunderDriveListBox", "ThunderRT6DriveListBox"
                            If m_DriveListBoxControl Then SubclassThis = True
                    Case "ThunderFileListBox", "ThunderRT6FileListBox"
                            If m_FileListBoxControl Then SubclassThis = True
                    Case "ThunderPictureBox", "ThunderPictureBoxDC", "ThunderRT6PictureBoxDC"
                            If m_PictureControl Then SubclassThis = True
                    Case Else
                            'Nothing
'                            Debug.Print ThisWindowClassName(aControl.hwnd)
            End Select
            If (SubclassThis) Then
               m_EngineStarted = True
               m_iCount = m_iCount + 1
               ReDim Preserve m_SubClassedItem(1 To m_iCount) As clsWinXPengine
               Set m_SubClassedItem(m_iCount) = New clsWinXPengine
               m_SubClassedItem(m_iCount).SchemeColor = m_ColorScheme
               m_SubClassedItem(m_iCount).ActiveScaleMode = UserControl.Parent.ScaleMode
               m_SubClassedItem(m_iCount).IdeSubClass = m_IDE
               m_SubClassedItem(m_iCount).Attach aControl
            End If
        End If
    Next aControl '//-- Each Control
    If m_MsgBox_InputBox Then
        m_EngineStarted = True
        m_iCount = m_iCount + 1
        ReDim Preserve m_SubClassedItem(1 To m_iCount) As clsWinXPengine
        Set m_SubClassedItem(m_iCount) = New clsWinXPengine
        m_SubClassedItem(m_iCount).SchemeColor = m_ColorScheme
        m_SubClassedItem(m_iCount).IdeSubClass = m_IDE
        m_SubClassedItem(m_iCount).BeforeAttachMessageBox UserControl.Parent.hwnd
    End If
    If m_CommonDialog Then
        m_EngineStarted = True
        m_iCount = m_iCount + 1
        ReDim Preserve m_SubClassedItem(1 To m_iCount) As clsWinXPengine
        Set m_SubClassedItem(m_iCount) = New clsWinXPengine
        m_SubClassedItem(m_iCount).SchemeColor = m_ColorScheme
        m_SubClassedItem(m_iCount).IdeSubClass = m_IDE
        m_SubClassedItem(m_iCount).BeforeAttachCommonDialog UserControl.Parent.hwnd
    End If
    SetProp UserControl.Parent.hwnd, "ColorScheme", m_ColorScheme
End Sub

Public Sub InitIDESubClassing()
    Dim SubclassThis As Boolean
    Dim aControl As Control
    For Each aControl In UserControl.Parent.Controls
        SubclassThis = False
        On Error Resume Next
        If Err.Number = 0 Then
            Select Case ThisWindowClassName(aControl.hwnd)
                    Case "ThunderListBox"
                            SubclassThis = True
                    Case "ThunderCommandButton", "Button"
                            SubclassThis = True
                    Case "ThunderFrame"
                            SubclassThis = True
                    Case "ThunderOptionButton"
                            SubclassThis = True
                    Case "ThunderCheckBox"
                            SubclassThis = True
                    Case "ThunderComboBox", "ComboBox"
                            SubclassThis = True
                    Case "ImageCombo20WndClass"
                            SubclassThis = True
                    Case "ProgressBar20WndClass"
                            SubclassThis = True
                    Case "ThunderDirListBox"
                            SubclassThis = True
                    Case "ThunderDriveListBox"
                            SubclassThis = True
                    Case "ThunderFileListBox"
                            SubclassThis = True
                    Case "ThunderPictureBox", "ThunderPictureBoxDC"
                            SubclassThis = True
                    Case Else
            'Nothing
            End Select
            If (SubclassThis) Then
                m_IDEStarted = True
                m_IDECount = m_IDECount + 1
                ReDim Preserve m_IDEItem(1 To m_IDECount) As clsWinXPengine
                Set m_IDEItem(m_IDECount) = New clsWinXPengine
                m_IDEItem(m_IDECount).SchemeColor = m_ColorScheme
                m_IDEItem(m_IDECount).Attach aControl
            End If
        End If
    Next aControl '//-- Each Control
End Sub

Public Sub EndWinXPCSubClassing()
    If m_EngineStarted = True Then
        For i = 1 To m_iCount
            m_SubClassedItem(i).UnloadEngine
            Set m_SubClassedItem(i) = Nothing
        Next i
        m_iCount = 0
        EngineStarted = False
    End If
End Sub

Public Property Get ColorScheme() As SchemeWindowColors
    ColorScheme = m_ColorScheme
End Property

Public Property Let ColorScheme(ByVal cColorScheme As SchemeWindowColors)
    m_ColorScheme = cColorScheme
    PropertyChanged "ColorScheme"
End Property

Public Property Get DirListBoxControl() As Boolean
    DirListBoxControl = m_DirListBoxControl
End Property

Public Property Let DirListBoxControl(ByVal cDirListBoxControl As Boolean)
    m_DirListBoxControl = cDirListBoxControl
    PropertyChanged "DirListBoxControl"
End Property

Public Property Get FileListBoxControl() As Boolean
    FileListBoxControl = m_FileListBoxControl
End Property

Public Property Let FileListBoxControl(ByVal cFileListBoxControl As Boolean)
    m_FileListBoxControl = cFileListBoxControl
    PropertyChanged "FileListBoxControl"
End Property

Public Property Get ListBoxControl() As Boolean
    ListBoxControl = m_ListBoxControl
End Property

Public Property Let ListBoxControl(ByVal cListBoxControl As Boolean)
    m_ListBoxControl = cListBoxControl
    PropertyChanged "ListBoxControl"
End Property

Public Property Get ImageComboControl() As Boolean
    ImageComboControl = m_ImageComboControl
End Property

Public Property Let ImageComboControl(ByVal cImageComboControl As Boolean)
    m_ImageComboControl = cImageComboControl
    PropertyChanged "ImageComboControl"
End Property

Public Property Get ProgressBarControl() As Boolean
    ProgressBarControl = m_ProgressBarControl
End Property

Public Property Let ProgressBarControl(ByVal cProgressBarControl As Boolean)
    m_ProgressBarControl = cProgressBarControl
    PropertyChanged "ProgressBarControl"
End Property

Public Property Get DriveListBoxControl() As Boolean
    DriveListBoxControl = m_DriveListBoxControl
End Property

Public Property Let DriveListBoxControl(ByVal cDriveListBoxControl As Boolean)
    m_DriveListBoxControl = cDriveListBoxControl
    PropertyChanged "DriveListBoxControl"
End Property

Public Property Get ComboBoxControl() As Boolean
    ComboBoxControl = m_ComboBoxControl
End Property

Public Property Let ComboBoxControl(ByVal cComboBoxControl As Boolean)
    m_ComboBoxControl = cComboBoxControl
    PropertyChanged "ComboBoxControl"
End Property

Public Property Get OptionControl() As Boolean
    OptionControl = m_OptionControl
End Property

Public Property Let OptionControl(ByVal cOptionControl As Boolean)
    m_OptionControl = cOptionControl
    PropertyChanged "OptionControl"
End Property

Public Property Get CheckControl() As Boolean
    CheckControl = m_CheckControl
End Property

Public Property Let CheckControl(ByVal cCheckControl As Boolean)
    m_CheckControl = cCheckControl
    PropertyChanged "CheckControl"
End Property

Public Property Get ButtonControl() As Boolean
    ButtonControl = m_ButtonControl
End Property

Public Property Let ButtonControl(ByVal cButtonControl As Boolean)
    m_ButtonControl = cButtonControl
    PropertyChanged "ButtonControl"
End Property

Public Property Get FrameControl() As Boolean
    FrameControl = m_FrameControl
End Property

Public Property Let FrameControl(ByVal cFrameControl As Boolean)
    m_FrameControl = cFrameControl
    PropertyChanged "FrameControl"
End Property

Public Property Get PictureControl() As Boolean
    PictureControl = m_PictureControl
End Property

Public Property Let PictureControl(ByVal cPictureControl As Boolean)
    m_PictureControl = cPictureControl
    PropertyChanged "PictureControl"
End Property

Public Property Get Common_Dialog() As Boolean
    Common_Dialog = m_CommonDialog
End Property

Public Property Let Common_Dialog(ByVal cCommon_Dialog As Boolean)
    m_CommonDialog = cCommon_Dialog
    PropertyChanged "Common_Dialog"
End Property

Public Property Get MsgBox_InputBox() As Boolean
    MsgBox_InputBox = m_MsgBox_InputBox
End Property

Public Property Let MsgBox_InputBox(ByVal cMsgBox_InputBox As Boolean)
    m_MsgBox_InputBox = cMsgBox_InputBox
    PropertyChanged "MsgBox_InputBox"
End Property

Public Property Get EngineStarted() As Boolean
    EngineStarted = m_EngineStarted
End Property

Public Property Let EngineStarted(ByVal cEngineStarted As Boolean)
    m_EngineStarted = cEngineStarted
    PropertyChanged "EngineStarted"
End Property

Public Property Get IDE() As Boolean
    IDE = m_IDE
End Property

Public Property Let IDE(ByVal cIDE As Boolean)
    m_IDE = cIDE
    PropertyChanged "IDE"
End Property

Private Sub UserControl_Paint()
    If IDE Then
        If Not m_IDEStarted Then InitIDESubClassing
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        ColorScheme = .ReadProperty("ColorScheme", WindowsXP_Blue)
        IDE = .ReadProperty("IDE", False)
        MsgBox_InputBox = .ReadProperty("MsgBox_InputBox", True)
        Common_Dialog = .ReadProperty("Common_Dialog", True)
        PictureControl = .ReadProperty("PictureControl", True)
        FrameControl = .ReadProperty("FrameControl", True)
        ButtonControl = .ReadProperty("ButtonControl", True)
        CheckControl = .ReadProperty("CheckControl", True)
        OptionControl = .ReadProperty("OptionControl", True)
        ComboBoxControl = .ReadProperty("ComboBoxControl", True)
        DriveListBoxControl = .ReadProperty("DriveListBoxControl", True)
        ProgressBarControl = .ReadProperty("ProgressBarControl", True)
        ImageComboControl = .ReadProperty("ImageComboControl", True)
        ListBoxControl = .ReadProperty("ListBoxControl", True)
        DirListBoxControl = .ReadProperty("DirListBoxControl", True)
        FileListBoxControl = .ReadProperty("FileListBoxControl", True)
        EngineStarted = .ReadProperty("EngineStarted", False)
    End With
End Sub

Private Sub UserControl_Resize()
    Width = 2266
    Height = 500
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("ColorScheme", m_ColorScheme, False)
        Call .WriteProperty("IDE", m_IDE, False)
        Call .WriteProperty("EngineStarted", m_EngineStarted, False)
        Call .WriteProperty("MsgBox_InputBox", m_MsgBox_InputBox, True)
        Call .WriteProperty("Common_Dialog", m_CommonDialog, True)
        Call .WriteProperty("ListBoxControl", m_ListBoxControl, True)
        Call .WriteProperty("PictureControl", m_PictureControl, True)
        Call .WriteProperty("FrameControl", m_FrameControl, True)
        Call .WriteProperty("ButtonControl", m_ButtonControl, True)
        Call .WriteProperty("CheckControl", m_CheckControl, True)
        Call .WriteProperty("OptionControl", m_OptionControl, True)
        Call .WriteProperty("ComboBoxControl", m_ComboBoxControl, True)
        Call .WriteProperty("DriveListBoxControl", m_DriveListBoxControl, True)
        Call .WriteProperty("ProgressBarControl", m_ProgressBarControl, True)
        Call .WriteProperty("ImageComboControl", m_ImageComboControl, True)
        Call .WriteProperty("FileListBoxControl", m_FileListBoxControl, True)
        Call .WriteProperty("DirListBoxControl", m_DirListBoxControl, True)
    End With
    If m_IDE = True Then
        Parent.Hide
        Parent.Visible = True
        Parent.Show
        Parent.Refresh
    End If
End Sub

