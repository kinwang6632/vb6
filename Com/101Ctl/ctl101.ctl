VERSION 5.00
Begin VB.UserControl ctl101 
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   21495
   LockControls    =   -1  'True
   ScaleHeight     =   825
   ScaleWidth      =   21495
   Begin VB.Label lblChargeAddressB 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   14790
      TabIndex        =   25
      Top             =   525
      Width           =   3315
   End
   Begin VB.Label lblChargeAddress 
      Appearance      =   0  '平面
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   14790
      TabIndex        =   24
      Top             =   300
      Width           =   3315
   End
   Begin VB.Label lblChargeAddrNoB 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   14070
      TabIndex        =   23
      Top             =   525
      Width           =   735
   End
   Begin VB.Label lblChargeAddrNo 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   14070
      TabIndex        =   22
      Top             =   300
      Width           =   735
   End
   Begin VB.Label lblAddrNo 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6030
      TabIndex        =   21
      Top             =   300
      Width           =   735
   End
   Begin VB.Label lblClass 
      Appearance      =   0  '平面
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   18090
      TabIndex        =   20
      Top             =   300
      Width           =   1545
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   $"ctl101.ctx":0000
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   30
      TabIndex        =   19
      Top             =   60
      Width           =   21000
   End
   Begin VB.Label lblMailAddrNoB 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   10050
      TabIndex        =   18
      Top             =   525
      Width           =   735
   End
   Begin VB.Label lblMailAddrNo 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   10050
      TabIndex        =   17
      Top             =   300
      Width           =   735
   End
   Begin VB.Label lblMailAddressB 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   10770
      TabIndex        =   16
      Top             =   525
      Width           =   3315
   End
   Begin VB.Label lblMailAddress 
      Appearance      =   0  '平面
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   10770
      TabIndex        =   15
      Top             =   300
      Width           =   3315
   End
   Begin VB.Label lblClassB 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   18090
      TabIndex        =   14
      Top             =   525
      Width           =   1545
   End
   Begin VB.Label lblTime 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "UpdTime"
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   19740
      TabIndex        =   13
      Top             =   570
      Width           =   1485
   End
   Begin VB.Label lblEn 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "UpdEn"
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   20040
      TabIndex        =   12
      Top             =   330
      Width           =   705
   End
   Begin VB.Label lblMode 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "裝機地址"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   30
      TabIndex        =   11
      Top             =   450
      Width           =   855
   End
   Begin VB.Label lblName 
      Appearance      =   0  '平面
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   900
      TabIndex        =   10
      Top             =   300
      Width           =   1635
   End
   Begin VB.Label lblNameB 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   900
      TabIndex        =   9
      Top             =   525
      Width           =   1635
   End
   Begin VB.Label lblAddress 
      Appearance      =   0  '平面
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6750
      TabIndex        =   8
      Top             =   300
      Width           =   3315
   End
   Begin VB.Label lblAddressB 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   6750
      TabIndex        =   7
      Top             =   525
      Width           =   3315
   End
   Begin VB.Label lblAddrNoB 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   6030
      TabIndex        =   6
      Top             =   525
      Width           =   735
   End
   Begin VB.Label lblTel3 
      Appearance      =   0  '平面
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4830
      TabIndex        =   5
      Top             =   300
      Width           =   1215
   End
   Begin VB.Label lblTel3B 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   4830
      TabIndex        =   4
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label lblTel2 
      Appearance      =   0  '平面
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3660
      TabIndex        =   3
      Top             =   300
      Width           =   1185
   End
   Begin VB.Label lblTel2B 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   3660
      TabIndex        =   2
      Top             =   525
      Width           =   1185
   End
   Begin VB.Label lblTel1B 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2520
      TabIndex        =   1
      ToolTipText     =   "300"
      Top             =   525
      Width           =   1185
   End
   Begin VB.Label lblTel1 
      Appearance      =   0  '平面
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2520
      TabIndex        =   0
      Top             =   300
      Width           =   1185
   End
End
Attribute VB_Name = "ctl101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Default Property Values:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0

'Property Variables:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer

'Event Declarations:
Event Click()
Attribute Click.VB_Description = "發生於使用者在某物件上方按下並放開滑鼠鍵時。"
Event DblClick()
Attribute DblClick.VB_Description = "發生於使用者在同一個物件上，連續兩次按下並放開滑鼠鍵。"
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "發生於使用者在物件具有駐點 (focus) 時，按下一鍵。"
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "發生於使用者按下並放開一個 ANSI 鍵時。"
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "發生於使用者在物件具有駐點 (focus) 時，放開一鍵。"
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "發生於使用者在物件具有駐點 (focus) 時，按下滑鼠鍵。"
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "發生於使用者移動滑鼠時。"
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "發生於使用者在物件具有駐點 (focus) 時，放開滑鼠鍵。"

Public Property Get FldName() As String
Attribute FldName.VB_ProcData.VB_Invoke_Property = ";資料"
Attribute FldName.VB_MemberFlags = "1c"
    FldName = lblName.Caption
End Property

Public Property Let FldName(ByVal New_FldName As String)
    lblName.Caption() = New_FldName
    PropertyChanged "FldName"
End Property

Public Property Get FldNameB() As String
Attribute FldNameB.VB_ProcData.VB_Invoke_Property = ";資料"
Attribute FldNameB.VB_MemberFlags = "1c"
    FldNameB = lblNameB.Caption
End Property

Public Property Let FldNameB(ByVal New_FldNameB As String)
    lblNameB.Caption() = New_FldNameB
    PropertyChanged "FldNameB"
End Property

Public Property Get FldMode() As String
Attribute FldMode.VB_ProcData.VB_Invoke_Property = ";資料"
Attribute FldMode.VB_MemberFlags = "14"
    FldMode = lblMode.Caption
End Property

Public Property Let FldMode(ByVal New_FldMode As String)
    lblMode.Caption() = New_FldMode
    PropertyChanged "FldMode"
End Property

Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "傳回或設定在物件中，顯示文字與圖形的背景色彩。"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "傳回或設定在物件中，顯示文字與圖形的前景色彩。"
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "傳回或設定一值，決定物件是否能回應使用者產生的事件。"
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "傳回一個 Font 物件"
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "用來決定 Label 本身或 Shape 的背景是否為透明的。"
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "傳回或設定物件的框線樣式。"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "強制物件進行完整的重繪。"
     
End Sub


Public Property Get FldTel1() As String
Attribute FldTel1.VB_Description = "傳回或設定顯示於物件標題列或圖示下方的文字。"
Attribute FldTel1.VB_ProcData.VB_Invoke_Property = ";資料"
Attribute FldTel1.VB_MemberFlags = "1c"
    FldTel1 = lblTel1.Caption
End Property

Public Property Let FldTel1(ByVal New_FldTel1 As String)
    lblTel1.Caption() = New_FldTel1
    PropertyChanged "FldTel1"
End Property

Public Property Get FldTel1B() As String
Attribute FldTel1B.VB_Description = "傳回或設定顯示於物件標題列或圖示下方的文字。"
Attribute FldTel1B.VB_ProcData.VB_Invoke_Property = ";資料"
Attribute FldTel1B.VB_MemberFlags = "1c"
    FldTel1B = lblTel1B.Caption
End Property

Public Property Let FldTel1B(ByVal New_FldTel1B As String)
    lblTel1B.Caption() = New_FldTel1B
    PropertyChanged "FldTel1B"
End Property

Public Property Get FldTel2() As String
Attribute FldTel2.VB_Description = "傳回或設定顯示於物件標題列或圖示下方的文字。"
Attribute FldTel2.VB_ProcData.VB_Invoke_Property = ";資料"
Attribute FldTel2.VB_MemberFlags = "34"
    FldTel2 = lblTel2.Caption
End Property

Public Property Let FldTel2(ByVal New_FldTel2 As String)
    lblTel2.Caption() = New_FldTel2
    PropertyChanged "FldTel2"
End Property

Public Property Get FldTel2B() As String
Attribute FldTel2B.VB_Description = "傳回或設定顯示於物件標題列或圖示下方的文字。"
Attribute FldTel2B.VB_MemberFlags = "1c"
    FldTel2B = lblTel2B.Caption
End Property

Public Property Let FldTel2B(ByVal New_FldTel2B As String)
    lblTel2B.Caption() = New_FldTel2B
    PropertyChanged "FldTel2B"
End Property

Public Property Get FldTel3() As String
Attribute FldTel3.VB_Description = "傳回或設定顯示於物件標題列或圖示下方的文字。"
Attribute FldTel3.VB_ProcData.VB_Invoke_Property = ";資料"
Attribute FldTel3.VB_MemberFlags = "1c"
    FldTel3 = lblTel3.Caption
End Property

Public Property Let FldTel3(ByVal New_FldTel3 As String)
    lblTel3.Caption() = New_FldTel3
    PropertyChanged "FldTel3"
End Property

Public Property Get FldTel3B() As String
Attribute FldTel3B.VB_Description = "傳回或設定顯示於物件標題列或圖示下方的文字。"
Attribute FldTel3B.VB_ProcData.VB_Invoke_Property = ";資料"
Attribute FldTel3B.VB_MemberFlags = "1c"
    FldTel3B = lblTel3B.Caption
End Property

Public Property Let FldTel3B(ByVal New_FldTel3B As String)
    lblTel3B.Caption() = New_FldTel3B
    PropertyChanged "FldTel3B"
End Property

Public Property Get FldAddrNo() As String
Attribute FldAddrNo.VB_Description = "傳回或設定顯示於物件標題列或圖示下方的文字。"
Attribute FldAddrNo.VB_ProcData.VB_Invoke_Property = ";資料"
Attribute FldAddrNo.VB_MemberFlags = "1c"
    FldAddrNo = lblAddrNo.Caption
End Property

Public Property Let FldAddrNo(ByVal New_FldAddrNo As String)
    lblAddrNo.Caption() = New_FldAddrNo
    PropertyChanged "FldAddrNo"
End Property

Public Property Get FldMailAddrNo() As String
    FldMailAddrNo = lblMailAddrNo.Caption
End Property

Public Property Let FldMailAddrNo(ByVal New_FldMailAddrNo As String)
    lblMailAddrNo.Caption() = New_FldMailAddrNo
    PropertyChanged "FldMailAddrNo"
End Property


Public Property Get FldChargeAddrNo() As String
    FldChargeAddrNo = lblChargeAddrNo.Caption
End Property

Public Property Let FldChargeAddrNo(ByVal New_FldChargeAddrNo As String)
    lblChargeAddrNo.Caption() = New_FldChargeAddrNo
    PropertyChanged "FldChargeAddrNo"
End Property

Public Property Get FldAddrNoB() As String
Attribute FldAddrNoB.VB_Description = "傳回或設定顯示於物件標題列或圖示下方的文字。"
Attribute FldAddrNoB.VB_ProcData.VB_Invoke_Property = ";資料"
Attribute FldAddrNoB.VB_MemberFlags = "1c"
    FldAddrNoB = lblAddrNoB.Caption
End Property

Public Property Let FldAddrNoB(ByVal New_FldAddrNoB As String)
    lblAddrNoB.Caption() = New_FldAddrNoB
    PropertyChanged "FldAddrNoB"
End Property

Public Property Get FldMailAddrNoB() As String
    FldMailAddrNoB = lblMailAddrNoB.Caption
End Property

Public Property Let FldMailAddrNoB(ByVal New_FldMailAddrNoB As String)
    lblMailAddrNoB.Caption() = New_FldMailAddrNoB
    PropertyChanged "FldMailAddrNoB"
End Property





Public Property Get FldChargeAddrNoB() As String
    FldChargeAddrNoB = lblChargeAddrNoB.Caption
End Property

Public Property Let FldChargeAddrNoB(ByVal New_FldChargeAddrNoB As String)
    lblChargeAddrNoB.Caption() = New_FldChargeAddrNoB
    PropertyChanged "FldChargeAddrNoB"
End Property


Public Property Get FldAddress() As String
Attribute FldAddress.VB_Description = "傳回或設定顯示於物件標題列或圖示下方的文字。"
Attribute FldAddress.VB_ProcData.VB_Invoke_Property = ";資料"
Attribute FldAddress.VB_MemberFlags = "1c"
    FldAddress = lblAddress.Caption
End Property

Public Property Let FldAddress(ByVal New_FldAddress As String)
    lblAddress.Caption() = New_FldAddress
    PropertyChanged "FldAddress"
End Property

Public Property Get FldMailAddress() As String
    FldMailAddress = lblMailAddress.Caption
End Property

Public Property Let FldMailAddress(ByVal New_FldMailAddress As String)
    lblMailAddress.Caption() = New_FldMailAddress
    PropertyChanged "FldMailAddress"
End Property

Public Property Get FldChargeAddress() As String
  FldChargeAddress = lblChargeAddress.Caption
End Property
Public Property Let FldChargeAddress(ByVal New_FldChargeAddress As String)
  lblChargeAddress.Caption = New_FldChargeAddress
  PropertyChanged "FldChargeAddress"
End Property




Public Property Get FldAddressB() As String
Attribute FldAddressB.VB_Description = "傳回或設定顯示於物件標題列或圖示下方的文字。"
Attribute FldAddressB.VB_ProcData.VB_Invoke_Property = ";資料"
Attribute FldAddressB.VB_MemberFlags = "21c"
    FldAddressB = lblAddressB.Caption
End Property

Public Property Let FldAddressB(ByVal New_FldAddressB As String)
    lblAddressB.Caption() = New_FldAddressB
    PropertyChanged "FldAddressB"
End Property

Public Property Get FldMailAddressB() As String
    FldMailAddressB = lblMailAddressB.Caption
End Property

Public Property Let FldMailAddressB(ByVal New_FldMailAddressB As String)
    lblMailAddressB.Caption() = New_FldMailAddressB
    PropertyChanged "FldMailAddressB"
End Property


Public Property Get FldChargeAddressB() As String
  FldChargeAddressB = lblChargeAddressB.Caption
End Property
Public Property Let FldChargeAddressB(ByVal New_FldChargeAddressB As String)
  lblChargeAddressB.Caption = New_FldChargeAddressB
  PropertyChanged "FldChargeAddressB"
End Property


Public Property Get FldClass() As String
    FldClass = lblClass.Caption
End Property

Public Property Let FldClass(ByVal New_FldClass As String)
    lblClass.Caption() = New_FldClass
    PropertyChanged "FldClass"
End Property

Public Property Get FldClassB() As String
    FldClassB = lblClassB.Caption
End Property

Public Property Let FldClassB(ByVal New_FldClassB As String)
    lblClassB.Caption() = New_FldClassB
    PropertyChanged "FldClassB"
End Property

Public Property Get FldEn() As String
Attribute FldEn.VB_ProcData.VB_Invoke_Property = ";資料"
Attribute FldEn.VB_MemberFlags = "1c"
    FldEn = lblEn.Caption
End Property

Public Property Let FldEn(ByVal New_FldEn As String)
    lblEn.Caption() = New_FldEn
    PropertyChanged "FldEn"
End Property

Public Property Get FldTime() As String
Attribute FldTime.VB_ProcData.VB_Invoke_Property = ";資料"
Attribute FldTime.VB_MemberFlags = "1c"
    FldTime = lblTime.Caption
End Property

Public Property Let FldTime(ByVal New_FldTime As String)
    lblTime.Caption() = New_FldTime
    PropertyChanged "FldTime"
End Property



'1=名稱,2=電話,3=裝機地址,4=客戶註銷,5=家長密碼,6=客戶類別,7=郵寄地址

Private Sub lblMode_Change()
    Select Case lblMode
           Case "1"
                lblMode = "名稱"
           Case "2"
                lblMode = "電話"
           Case "3"
                lblMode = "裝機地址"
           Case "4"
                lblMode = "客戶註銷"
           Case "5"
                lblMode = "家長密碼"
           Case "6"
                lblMode = "客戶類別"
           Case "7"
                lblMode = "郵寄地址"
           Case "8"
                lblMode = "收費地址"
           Case Else
'               lblMode = lblMode
    End Select
End Sub

'初始化使用者控制項的屬性
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
End Sub

'由儲存區載入屬性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        lblName.Caption = .ReadProperty("FldName", "")
        lblNameB.Caption = .ReadProperty("FldNameB", "")
        lblMode.Caption = .ReadProperty("FldMode", "")
        m_BackColor = .ReadProperty("BackColor", m_def_BackColor)
        m_ForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
        m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
        Set m_Font = .ReadProperty("Font", Ambient.Font)
        m_BackStyle = .ReadProperty("BackStyle", m_def_BackStyle)
        m_BorderStyle = .ReadProperty("BorderStyle", m_def_BorderStyle)
        lblTel1.Caption = .ReadProperty("FldTel1", "")
        lblTel1B.Caption = .ReadProperty("FldTel1B", "")
        lblTel2.Caption = .ReadProperty("FldTel2", "")
        lblTel2B.Caption = .ReadProperty("FldTel2B", "")
        lblTel3.Caption = .ReadProperty("FldTel3", "")
        lblTel3B.Caption = .ReadProperty("FldTel3B", "")
        lblAddrNo.Caption = .ReadProperty("FldAddrNo", "")
        lblAddrNoB.Caption = .ReadProperty("FldAddrNoB", "")
        lblAddress.Caption = .ReadProperty("FldAddress", "")
        lblAddressB.Caption = .ReadProperty("FldAddressB", "")
        lblMailAddrNo.Caption = .ReadProperty("FldMailAddrNo", "")
        lblMailAddrNoB.Caption = .ReadProperty("FldMailAddrNoB", "")
        lblMailAddress.Caption = .ReadProperty("FldMailAddress", "")
        lblMailAddressB.Caption = .ReadProperty("FldMailAddressB", "")
        '#5937 增加收費地址 By Kin 2011/05/06
        lblChargeAddrNo.Caption = .ReadProperty("FldChargeAddrNo", "")
        lblChargeAddrNoB.Caption = .ReadProperty("FldChargeAddrNoB", "")
        lblChargeAddress.Caption = .ReadProperty("FldChargeAddress", "")
        lblChargeAddressB.Caption = .ReadProperty("FldChargeAddressB", "")
        lblEn.Caption = .ReadProperty("FldEn", "")
        lblTime.Caption = .ReadProperty("FldTime", "")
        lblClass.Caption = .ReadProperty("FldClass", "")
        lblClassB.Caption = .ReadProperty("FldClassB", "")
    End With
End Sub

'將屬性值寫回儲存區
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("FldName", lblName.Caption, "")
        Call .WriteProperty("FldNameB", lblNameB.Caption, "")
        Call .WriteProperty("FldMode", lblMode.Caption, "")
        Call .WriteProperty("BackColor", m_BackColor, m_def_BackColor)
        Call .WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
        Call .WriteProperty("Enabled", m_Enabled, m_def_Enabled)
        Call .WriteProperty("Font", m_Font, Ambient.Font)
        Call .WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
        Call .WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
        Call .WriteProperty("FldTel1", lblTel1.Caption, "")
        Call .WriteProperty("FldTel1B", lblTel1B.Caption, "")
        Call .WriteProperty("FldTel2", lblTel2.Caption, "")
        Call .WriteProperty("FldTel2B", lblTel2B.Caption, "")
        Call .WriteProperty("FldTel3", lblTel3.Caption, "")
        Call .WriteProperty("FldTel3B", lblTel3B.Caption, "")
        Call .WriteProperty("FldAddrNo", lblAddrNo.Caption, "")
        Call .WriteProperty("FldAddrNoB", lblAddrNoB.Caption, "")
        Call .WriteProperty("FldAddress", lblAddress.Caption, "")
        Call .WriteProperty("FldAddressB", lblAddressB.Caption, "")
        Call .WriteProperty("FldMailAddrNo", lblMailAddrNo.Caption, "")
        Call .WriteProperty("FldMailAddrNoB", lblMailAddrNoB.Caption, "")
        Call .WriteProperty("FldMailAddress", lblMailAddress.Caption, "")
        Call .WriteProperty("FldMailAddressB", lblMailAddressB.Caption, "")
        '#5937 增加收費地址 By Kin 2011/05/06
        Call .WriteProperty("FldChargeAddrNo", lblChargeAddrNo.Caption, "")
        Call .WriteProperty("FldChargeAddrNoB", lblChargeAddrNoB.Caption, "")
        Call .WriteProperty("FldChargeAddress", lblChargeAddress.Caption, "")
        Call .WriteProperty("FldChargeAddressB", lblChargeAddressB.Caption, "")
        
        Call .WriteProperty("FldEn", lblEn.Caption, "")
        Call .WriteProperty("FldTime", lblTime.Caption, "")
        Call .WriteProperty("FldClass", lblClass.Caption, "")
        Call .WriteProperty("FldClassB", lblClassB.Caption, "")
    End With
End Sub
